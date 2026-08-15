#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
mhlw_price_converter_v6_2.py

目的:
    厚労省のExcel群を読み込み、薬価アプリ向けのJSON群を出力する。

出力:
    output/master_bundle.json
    output/price_items.json
    output/brand_items.json
    output/receipt_code_maps.json
    output/generic_name_maps.json
    output/flags.json
    output/join_report.json

主要テーブル:
    - price_items
    - brand_items
    - receipt_code_maps  ※現時点では入力ソース未指定のため空配列
    - generic_name_maps
    - flags

v6.2 の主な変更:
    - schema_version を 1.2.0 に更新
      ※v6.1で price を JSON 上 string Decimal にしたため、スキーマ変更として明示。
    - input_files.json の検証を強化
      ※base_files 空配列、空文字、重複をエラーにする。
    - join_report に input_file_config と _06 未処理メモを出力
    - aux05 の薬価基準収載医薬品コード重複レポートを出力
      ※load時は従来どおり code 単位で先頭1件に dedupe するが、捨てた可能性を可視化。

依存:
    pip install pandas openpyxl

使い方:
    python mhlw_price_converter_v6_2.py --input-dir . --output-dir output

任意:
    python mhlw_price_converter_v6_2.py --input-dir . --output-dir output --manual-aliases manual_aliases.json
    python mhlw_price_converter_v6_2.py --input-dir . --output-dir output --config input_files.json

input_files.json 例:
    {
      "base_files": [
        "tp20260520-01_01.xlsx",
        "tp20260520-01_02.xlsx",
        "tp20260520-01_03.xlsx",
        "tp20260401-01_04.xlsx"
      ],
      "aux05_file": "tp20260520-01_05.xlsx",
      "kiso_file": "tp2026_kiso.xlsx",
      "temporary_exception_file": "tp20260520-01_06.xlsx"
    }

注意:
    temporary_exception_file は v6.2 では処理対象ではない。
    join_report に未処理ファイルとして記録するための任意メタ情報。
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
import unicodedata
from dataclasses import dataclass, field, asdict, is_dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Optional

import pandas as pd

try:
    from zoneinfo import ZoneInfo
except ImportError:  # pragma: no cover
    ZoneInfo = None  # type: ignore


# =========================================================
# 設定
# =========================================================

SCHEMA_VERSION = "1.2.0"
GENERATOR_NAME = "mhlw_price_converter_v6_2"

DEFAULT_BASE_FILE_NAMES = [
    "tp20260520-01_01.xlsx",
    "tp20260520-01_02.xlsx",
    "tp20260520-01_03.xlsx",
    "tp20260401-01_04.xlsx",
]
DEFAULT_AUX05_FILE_NAME = "tp20260520-01_05.xlsx"
DEFAULT_KISO_FILE_NAME = "tp2026_kiso.xlsx"
DEFAULT_TEMPORARY_EXCEPTION_FILE_NAME = "tp20260520-01_06.xlsx"

BASE_REQUIRED_COLUMNS = {
    "category": ["区分"],
    "code": ["薬価基準収載医薬品コード", "コード"],
    "ingredient": ["成分名"],
    "spec": ["規格"],
    "name": ["品名"],
    "manufacturer": ["メーカー名"],
    "price": ["薬価"],
}

BASE_OPTIONAL_COLUMNS = {
    "mark_e": ["E", "品名頭記号1", "局区分", "記号1"],
    "mark_f": ["F", "品名頭記号2", "麻区分", "記号2"],
    "mark_g": ["G", "品名頭記号3", "その他記号", "記号3"],
    "raw_j": ["J", "後発医薬品"],
    "raw_k": ["K", "先発医薬品"],
    "raw_l": ["L", "後発医薬品のある先発医薬品"],
    "raw_n": ["N", "経過措置"],
    "raw_o": ["O", "備考"],
}

AUX05_COLUMNS = {
    "code": ["薬価基準収載医薬品コード", "コード"],
    "d": ["D", "区分記号", "種別"],
    "e": ["E", "収載年月日"],
    "f": ["F", "経過措置期限"],
    "g": ["G", "備考"],
    "h": ["H", "除外区分", "除外コード"],
}

KISO_COLUMNS = {
    "category": ["区分"],
    "name": ["品名"],
    "ingredient": ["成分", "成分名"],
    "spec": ["規格単位", "規格"],
    "manufacturer": ["メーカー名"],
    "price": ["薬価"],
}

BASE_POSITIONAL_FALLBACKS = {
    "category": 0,
    "code": 1,
    "ingredient": 2,
    "spec": 3,
    "mark_e": 4,
    "mark_f": 5,
    "mark_g": 6,
    "name": 7,
    "manufacturer": 8,
    "raw_j": 9,
    "raw_k": 10,
    "raw_l": 11,
    "price": 12,
    "raw_n": 13,
    "raw_o": 14,
}

AUX05_POSITIONAL_FALLBACKS = {
    "code": 0,
    "d": 3,
    "e": 4,
    "f": 5,
    "g": 6,
    "h": 7,
}

KISO_POSITIONAL_FALLBACKS = {
    "category": 0,
    "name": 1,
    "ingredient": 2,
    "spec": 3,
    "manufacturer": 4,
    "price": 5,
}

SPACE_RE = re.compile(r"\s+")
NON_NUMERIC_RE = re.compile(r"[^\d\.\-]+")

THERAPEUTIC_CLASS_NAME_MAP: dict[str, str] = {}
ROUTE_NAME_MAP: dict[str, str] = {}
DOSAGE_FORM_NAME_MAP: dict[str, str] = {}

MANUFACTURER_ALIASES = {
    "光製薬ヒカリセイヤク": "光製薬",
}

# 全フィールド共通。意味を変える置換は入れない。
COMMON_TOKEN_ALIASES = {
    "㎎": "mg",
    "μg": "mcg",
    "㎍": "mcg",
    "ｍｇ": "mg",
    "ｍＬ": "mL",
    "ＭＬ": "mL",
    "ml": "mL",
    "　": " ",
}

# 成分名専用。Na は品名・メーカー名では無条件置換しない。
INGREDIENT_TOKEN_ALIASES = {
    **COMMON_TOKEN_ALIASES,
    "Na": "ナトリウム",
}

DOSAGE_PATTERNS = [
    (r"CR", "CR"),
    (r"OD", "OD"),
    (r"口腔内崩壊", "OD"),
    (r"徐放", "徐放"),
    (r"カプセル", "カプセル"),
    (r"細粒", "細粒"),
    (r"散", "散"),
    (r"シロップ", "シロップ"),
    (r"内用液", "内用液"),
    (r"注", "注射"),
    (r"軟膏", "軟膏"),
]


# =========================================================
# dataclass
# =========================================================

@dataclass
class InputFileConfig:
    base_files: list[str] = field(default_factory=lambda: list(DEFAULT_BASE_FILE_NAMES))
    aux05_file: str = DEFAULT_AUX05_FILE_NAME
    kiso_file: str = DEFAULT_KISO_FILE_NAME

    # v6.2: 処理はしない。join_report に未処理として記録するためのメタ情報。
    temporary_exception_file: Optional[str] = DEFAULT_TEMPORARY_EXCEPTION_FILE_NAME


@dataclass
class SourceSummary:
    base_files: list[str] = field(default_factory=list)
    aux05_present: bool = False
    kiso_match_rule: Optional[str] = None
    kiso_match_confidence: Optional[float] = None


@dataclass
class MetaInfo:
    schema_version: str
    generated_at: str
    generator: str
    source_files: list[str] = field(default_factory=list)


@dataclass
class PriceItem:
    price_item_id: str
    mhlw_code: str

    category: str
    category_normalized: str
    ingredient: str
    ingredient_normalized: str
    spec: str
    spec_normalized: str

    # v6.1以降: 金額はJSON上では文字列。Swift側で Decimal(string:) に渡す前提。
    price: Optional[str]
    currency: str = "JPY"

    therapeutic_class_code: Optional[str] = None
    therapeutic_class_name: Optional[str] = None
    route_code: Optional[str] = None
    route_name: Optional[str] = None
    dosage_form_code: Optional[str] = None
    dosage_form_name: Optional[str] = None

    display_group_key: str = ""
    group_label: str = ""
    search_text: str = ""

    transitional_deadline_text: Optional[str] = None
    remarks_text: Optional[str] = None

    brand_item_ids: list[str] = field(default_factory=list)
    receipt_code_ids: list[str] = field(default_factory=list)
    generic_name_map_ids: list[str] = field(default_factory=list)
    flag_ids: list[str] = field(default_factory=list)

    source_summary: SourceSummary = field(default_factory=SourceSummary)


@dataclass
class BrandSortKeys:
    price_desc: float = -1.0
    originator_first: int = 1
    generic_equivalent_first: int = 1
    name: str = ""
    manufacturer: str = ""


@dataclass
class SourceRowInfo:
    source_file: str
    source_code: str


@dataclass
class BrandItem:
    brand_item_id: str
    price_item_id: str
    mhlw_code: str

    # TODO: YJ/HOTコードは別ソース導入後に実装する。
    yj_code: Optional[str]
    hot_code: Optional[str]
    name: str
    name_normalized: str
    manufacturer: str
    manufacturer_normalized: str
    brand_label: Optional[str]
    display_name: str
    search_text: str

    sort_keys: BrandSortKeys = field(default_factory=BrandSortKeys)
    flag_ids: list[str] = field(default_factory=list)
    source_row: Optional[SourceRowInfo] = None


@dataclass
class MappingSource:
    source_name: str
    source_file: Optional[str] = None
    source_row_hint: Optional[str] = None


@dataclass
class ReceiptCodeMap:
    receipt_code_map_id: str
    receipt_code: str

    price_item_id: str
    mhlw_code: str

    brand_item_id: Optional[str] = None
    yj_code: Optional[str] = None

    mapping_type: str = "exact_price_item"
    mapping_status: str = "active"
    mapping_confidence: float = 1.0

    source: MappingSource = field(default_factory=lambda: MappingSource(source_name="unknown"))
    notes: list[str] = field(default_factory=list)


@dataclass
class GenericNameMap:
    generic_name_map_id: str
    generic_code: str

    price_item_id: str
    mhlw_code: str

    generic_display_name: str
    ingredient: str
    spec: str

    mapping_type: str = "derived_from_first9_plus_ZZZ"
    mapping_status: str = "active"
    mapping_confidence: float = 1.0

    source: MappingSource = field(default_factory=lambda: MappingSource(source_name="unknown"))
    notes: list[str] = field(default_factory=list)


@dataclass
class RawSource:
    columns: list[str] = field(default_factory=list)
    values: list[str] = field(default_factory=list)


@dataclass
class FlagItem:
    flag_id: str
    owner_type: str
    owner_id: str
    flag_type: str

    value_type: str
    value: Any

    raw_source: RawSource = field(default_factory=RawSource)
    derived_by: str = ""
    confidence: Optional[float] = None
    ui_badge: Optional[str] = None
    notes: list[str] = field(default_factory=list)


@dataclass
class ReportBundle:
    join_report: dict[str, Any] = field(default_factory=dict)


@dataclass
class MasterBundle:
    meta: MetaInfo
    price_items: list[PriceItem] = field(default_factory=list)
    brand_items: list[BrandItem] = field(default_factory=list)
    receipt_code_maps: list[ReceiptCodeMap] = field(default_factory=list)
    generic_name_maps: list[GenericNameMap] = field(default_factory=list)
    flags: list[FlagItem] = field(default_factory=list)
    reports: ReportBundle = field(default_factory=ReportBundle)


# =========================================================
# 欠損・JSON安全化
# =========================================================

def is_missing(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    try:
        result = pd.isna(value)
        if isinstance(result, bool):
            return result
        return False
    except Exception:
        return False


def value_is_true(value: Any) -> bool:
    if is_missing(value):
        return False
    try:
        return bool(value) is True
    except Exception:
        return False


def value_is_false(value: Any) -> bool:
    if is_missing(value):
        return False
    try:
        return bool(value) is False
    except Exception:
        return False


def to_optional_float(value: Any) -> Optional[float]:
    if is_missing(value):
        return None
    try:
        f = float(value)
    except Exception:
        return None
    if math.isnan(f) or math.isinf(f):
        return None
    return f


def decimal_to_json_price(value: Optional[Decimal]) -> Optional[str]:
    if value is None or is_missing(value):
        return None
    return str(value)


def decimal_key(value: Any) -> str:
    if is_missing(value):
        return ""
    try:
        d = value if isinstance(value, Decimal) else Decimal(str(value))
    except Exception:
        return safe_str(value)
    return format(d.normalize(), "f")


def clean_json_value(value: Any) -> Any:
    """json.dump(..., allow_nan=False) に渡せる値へ再帰的に変換する。"""
    if is_dataclass(value):
        return clean_json_value(asdict(value))
    if isinstance(value, dict):
        return {str(k): clean_json_value(v) for k, v in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [clean_json_value(v) for v in value]
    if isinstance(value, Decimal):
        return decimal_to_json_price(value)
    if isinstance(value, Path):
        return str(value)
    if is_missing(value):
        return None
    if hasattr(value, "item"):
        try:
            return clean_json_value(value.item())
        except Exception:
            pass
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return None
        return value
    return value


# =========================================================
# 正規化
# =========================================================

def nfkc(text: Any) -> str:
    if text is None:
        return ""
    if is_missing(text):
        return ""
    return unicodedata.normalize("NFKC", str(text)).strip()


def squash_spaces(text: Any) -> str:
    return SPACE_RE.sub(" ", nfkc(text)).strip()


def apply_token_aliases(text: str, aliases: dict[str, str]) -> str:
    s = text
    for k, v in aliases.items():
        s = s.replace(k, v)
    return s


def normalize_text(text: Any, aliases: Optional[dict[str, str]] = None) -> str:
    s = squash_spaces(text)
    s = apply_token_aliases(s, aliases or COMMON_TOKEN_ALIASES)
    return squash_spaces(s)


def normalize_name(text: Any) -> str:
    s = normalize_text(text, COMMON_TOKEN_ALIASES)
    s = s.replace("\n", " ")
    return squash_spaces(s)


def normalize_ingredient(text: Any) -> str:
    return normalize_text(text, INGREDIENT_TOKEN_ALIASES)


def normalize_spec(text: Any) -> str:
    s = normalize_text(text, COMMON_TOKEN_ALIASES)
    s = s.replace("1 錠", "1錠").replace("1 g", "1g").replace("1 ml", "1mL").replace("1 ML", "1mL")
    return squash_spaces(s)


def normalize_category(text: Any) -> str:
    s = normalize_text(text, COMMON_TOKEN_ALIASES)
    if s == "歯科用薬":
        return "歯科用薬剤"
    return s


def normalize_manufacturer(text: Any) -> str:
    s = normalize_text(text, COMMON_TOKEN_ALIASES).replace(" ", "")
    return MANUFACTURER_ALIASES.get(s, s)


def parse_price(value: Any) -> Optional[Decimal]:
    if value is None or is_missing(value):
        return None
    s = normalize_text(value, COMMON_TOKEN_ALIASES).replace(",", "").replace("円", "")
    s = NON_NUMERIC_RE.sub("", s)
    if s in {"", "-", "—", ".", "-."}:
        return None
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def price_to_sort_float(value: Any) -> float:
    if is_missing(value):
        return -1.0
    try:
        return float(value)
    except Exception:
        return -1.0


def safe_str(value: Any) -> str:
    if value is None or is_missing(value):
        return ""
    return str(value)


def series_has_nonempty(series: pd.Series) -> bool:
    return any(normalize_text(v) != "" for v in series.tolist())


# =========================================================
# ID helper
# =========================================================

def make_price_item_id(mhlw_code: str) -> str:
    return mhlw_code


def make_brand_item_id(mhlw_code: str, serial: int) -> str:
    return f"brand:{mhlw_code}:{serial}"


def make_receipt_code_map_id(receipt_code: str) -> str:
    return f"rcpt:{receipt_code}"


def make_generic_name_map_id(generic_code: str, mhlw_code: str) -> str:
    return f"gmap:{generic_code}:{mhlw_code}"


def make_flag_id(owner_type: str, owner_id: str, flag_type: str) -> str:
    return f"flag:{owner_type}:{owner_id}:{flag_type}"


# =========================================================
# 設定ファイル
# =========================================================

def _normalize_file_name(value: Any, field_name: str, allow_none: bool = False) -> Optional[str]:
    if value is None and allow_none:
        return None
    if not isinstance(value, str):
        raise ValueError(f"config.{field_name} は文字列である必要があります")
    s = value.strip()
    if not s:
        raise ValueError(f"config.{field_name} は空文字にできません")
    return s


def load_input_file_config(config_path: Optional[Path]) -> InputFileConfig:
    if config_path is None:
        return InputFileConfig()
    if not config_path.exists():
        raise FileNotFoundError(f"config file not found: {config_path}")

    with config_path.open("r", encoding="utf-8") as f:
        data = json.load(f)

    if not isinstance(data, dict):
        raise ValueError("config JSON はオブジェクトである必要があります")

    raw_base_files = data.get("base_files", DEFAULT_BASE_FILE_NAMES)
    raw_aux05_file = data.get("aux05_file", DEFAULT_AUX05_FILE_NAME)
    raw_kiso_file = data.get("kiso_file", DEFAULT_KISO_FILE_NAME)
    raw_temporary_exception_file = data.get("temporary_exception_file", DEFAULT_TEMPORARY_EXCEPTION_FILE_NAME)

    if not isinstance(raw_base_files, list) or not all(isinstance(x, str) for x in raw_base_files):
        raise ValueError("config.base_files は文字列配列である必要があります")

    base_files = [x.strip() for x in raw_base_files]
    if not base_files:
        raise ValueError("config.base_files は空配列にできません")
    if any(x == "" for x in base_files):
        raise ValueError("config.base_files に空文字は指定できません")

    duplicated_base_files = sorted({x for x in base_files if base_files.count(x) > 1})
    if duplicated_base_files:
        raise ValueError("config.base_files に重複があります: " + ", ".join(duplicated_base_files))

    aux05_file = _normalize_file_name(raw_aux05_file, "aux05_file")
    kiso_file = _normalize_file_name(raw_kiso_file, "kiso_file")
    temporary_exception_file = _normalize_file_name(
        raw_temporary_exception_file,
        "temporary_exception_file",
        allow_none=True,
    )

    reserved = [*base_files, aux05_file, kiso_file]
    duplicated_required_files = sorted({x for x in reserved if reserved.count(x) > 1})
    if duplicated_required_files:
        raise ValueError("入力ファイル指定に重複があります: " + ", ".join(duplicated_required_files))

    return InputFileConfig(
        base_files=base_files,
        aux05_file=aux05_file or DEFAULT_AUX05_FILE_NAME,
        kiso_file=kiso_file or DEFAULT_KISO_FILE_NAME,
        temporary_exception_file=temporary_exception_file,
    )


# =========================================================
# Excel 読み込み
# =========================================================

def read_excel_auto_header(path: Path, max_scan_rows: int = 10) -> pd.DataFrame:
    raw = pd.read_excel(path, dtype=str, header=None)
    raw = raw.dropna(how="all")
    if raw.empty:
        raise ValueError(f"空のExcelです: {path}")

    best_row = 0
    best_score = -1
    keywords = {
        "区分", "薬価基準収載医薬品コード", "コード", "成分名",
        "規格", "品名", "メーカー名", "薬価", "成分", "規格単位"
    }

    scan_rows = min(max_scan_rows, len(raw))
    for i in range(scan_rows):
        row_values = [nfkc(v) for v in raw.iloc[i].tolist()]
        score = sum(1 for v in row_values if v in keywords)
        if score > best_score:
            best_score = score
            best_row = i

    df = pd.read_excel(path, dtype=str, header=best_row)
    df = df.dropna(how="all")
    df.columns = [nfkc(c) if nfkc(c) else f"__col_{idx}" for idx, c in enumerate(df.columns)]
    return df


def resolve_columns(
    df: pd.DataFrame,
    mapping: dict[str, list[str]],
    positional_fallbacks: Optional[dict[str, int]] = None,
    required: bool = True,
) -> dict[str, Optional[str]]:
    resolved: dict[str, Optional[str]] = {}
    cols = list(df.columns)
    col_set = set(cols)

    for logical_name, candidates in mapping.items():
        found: Optional[str] = None

        for c in candidates:
            if c in col_set:
                found = c
                break

        if found is None and positional_fallbacks and logical_name in positional_fallbacks:
            idx = positional_fallbacks[logical_name]
            if 0 <= idx < len(cols):
                found = cols[idx]

        if found is None and required:
            raise ValueError(
                f"列が見つかりません: logical_name={logical_name}, candidates={candidates}, actual={cols}"
            )

        resolved[logical_name] = found

    return resolved


def get_series(df: pd.DataFrame, col_name: Optional[str], length: Optional[int] = None) -> pd.Series:
    if col_name is None:
        if length is None:
            length = len(df)
        return pd.Series([None] * length)
    return df[col_name]


# =========================================================
# バリデーション
# =========================================================

def validate_base_df(df: pd.DataFrame) -> None:
    if df.empty:
        raise ValueError("base df is empty")
    if df["code"].isna().all():
        raise ValueError("code column is empty")

    required_cols = ["code", "category", "ingredient", "spec", "price"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"base df に必要列がありません: {missing}")

    grouped = df.groupby("code", dropna=True, sort=False)
    bad_empty_groups: dict[str, dict[str, int]] = {}

    for code, group in grouped:
        empties = {
            "category_empty": int((group["category"].fillna("") == "").all()),
            "ingredient_empty": int((group["ingredient"].fillna("") == "").all()),
            "spec_empty": int((group["spec"].fillna("") == "").all()),
        }
        if any(v == 1 for v in empties.values()):
            bad_empty_groups[str(code)] = empties

    if bad_empty_groups:
        sample = dict(list(bad_empty_groups.items())[:20])
        raise ValueError(f"親属性が空のコード群があります: {sample}")


def validate_aux05_df(df: pd.DataFrame) -> None:
    if df.empty:
        raise ValueError("aux05 df is empty")


def validate_kiso_df(df: pd.DataFrame) -> None:
    if df.empty:
        raise ValueError("kiso df is empty")


# =========================================================
# ローダ
# =========================================================

def load_base_master(paths: list[Path]) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []

    for path in paths:
        raw = read_excel_auto_header(path)

        req = resolve_columns(
            raw,
            BASE_REQUIRED_COLUMNS,
            positional_fallbacks=BASE_POSITIONAL_FALLBACKS,
            required=True,
        )
        opt = resolve_columns(
            raw,
            BASE_OPTIONAL_COLUMNS,
            positional_fallbacks=BASE_POSITIONAL_FALLBACKS,
            required=False,
        )

        length = len(raw)

        df = pd.DataFrame({
            "source_file": str(path.name),
            "category": get_series(raw, req["category"], length),
            "code": get_series(raw, req["code"], length),
            "ingredient": get_series(raw, req["ingredient"], length),
            "spec": get_series(raw, req["spec"], length),
            "name": get_series(raw, req["name"], length),
            "manufacturer": get_series(raw, req["manufacturer"], length),
            "price_raw": get_series(raw, req["price"], length),
            "mark_e": get_series(raw, opt.get("mark_e"), length),
            "mark_f": get_series(raw, opt.get("mark_f"), length),
            "mark_g": get_series(raw, opt.get("mark_g"), length),
            "raw_j": get_series(raw, opt.get("raw_j"), length),
            "raw_k": get_series(raw, opt.get("raw_k"), length),
            "raw_l": get_series(raw, opt.get("raw_l"), length),
            "raw_n": get_series(raw, opt.get("raw_n"), length),
            "raw_o": get_series(raw, opt.get("raw_o"), length),
        })

        frames.append(df)

    base = pd.concat(frames, ignore_index=True)

    base["code"] = base["code"].map(lambda x: normalize_text(x, COMMON_TOKEN_ALIASES))
    base["category"] = base["category"].map(normalize_category)
    base["ingredient"] = base["ingredient"].map(normalize_ingredient)
    base["spec"] = base["spec"].map(normalize_spec)
    base["name"] = base["name"].map(normalize_name)
    base["manufacturer"] = base["manufacturer"].map(normalize_manufacturer)
    base["price"] = base["price_raw"].map(parse_price)

    base = base[
        (base["code"] != "")
        & (base["name"] != "")
        & (base["ingredient"] != "")
    ].copy()

    validate_base_df(base)
    return base


def build_aux05_duplicate_report(df: pd.DataFrame, sample_limit: int = 20) -> dict[str, Any]:
    if df.empty or "code" not in df.columns:
        return {
            "input_rows_after_code_filter": int(len(df)),
            "unique_code_count": 0,
            "duplicate_code_count": 0,
            "duplicate_rows_to_drop": 0,
            "conflicting_duplicate_code_count": 0,
            "duplicate_code_samples": [],
        }

    counts = df["code"].value_counts(dropna=False)
    duplicate_codes = [str(code) for code, count in counts.items() if str(code) != "" and int(count) > 1]
    value_cols = ["aux05_d", "aux05_e", "aux05_f", "aux05_g", "aux05_h"]

    samples: list[dict[str, Any]] = []
    conflicting_count = 0

    for code in duplicate_codes[:sample_limit]:
        group = df[df["code"] == code]
        unique_payload_count = int(group[value_cols].drop_duplicates().shape[0])
        if unique_payload_count > 1:
            conflicting_count += 1

        row_samples = []
        for _, r in group.head(5).iterrows():
            row_samples.append({c: safe_str(r.get(c)) for c in ["code", *value_cols]})

        samples.append({
            "code": code,
            "row_count": int(len(group)),
            "unique_payload_count": unique_payload_count,
            "has_conflict": unique_payload_count > 1,
            "rows_sample": row_samples,
        })

    # sample外にも conflict がないか数える。
    for code in duplicate_codes[sample_limit:]:
        group = df[df["code"] == code]
        if int(group[value_cols].drop_duplicates().shape[0]) > 1:
            conflicting_count += 1

    return {
        "input_rows_after_code_filter": int(len(df)),
        "unique_code_count": int(df["code"].nunique(dropna=True)),
        "duplicate_code_count": int(len(duplicate_codes)),
        "duplicate_rows_to_drop": int(len(df) - df["code"].nunique(dropna=True)),
        "conflicting_duplicate_code_count": int(conflicting_count),
        "duplicate_code_samples": samples,
    }


def load_aux05(path: Path) -> tuple[pd.DataFrame, dict[str, Any]]:
    raw = read_excel_auto_header(path)

    cols = resolve_columns(
        raw,
        AUX05_COLUMNS,
        positional_fallbacks=AUX05_POSITIONAL_FALLBACKS,
        required=True,
    )

    length = len(raw)
    df = pd.DataFrame({
        "code": get_series(raw, cols["code"], length).map(lambda x: normalize_text(x, COMMON_TOKEN_ALIASES)),
        "aux05_d": get_series(raw, cols["d"], length).map(lambda x: normalize_text(x, COMMON_TOKEN_ALIASES)),
        "aux05_e": get_series(raw, cols["e"], length).map(lambda x: normalize_text(x, COMMON_TOKEN_ALIASES)),
        "aux05_f": get_series(raw, cols["f"], length).map(lambda x: normalize_text(x, COMMON_TOKEN_ALIASES)),
        "aux05_g": get_series(raw, cols["g"], length).map(lambda x: normalize_text(x, COMMON_TOKEN_ALIASES)),
        "aux05_h": get_series(raw, cols["h"], length).map(lambda x: normalize_text(x, COMMON_TOKEN_ALIASES)),
    })

    df = df[df["code"] != ""].copy()
    duplicate_report = build_aux05_duplicate_report(df)

    # 現行のJOINは code 単位。重複は先頭を残す。
    # 重複の有無・内容差は duplicate_report に残す。
    df = df.drop_duplicates(subset=["code"], keep="first").copy()
    validate_aux05_df(df)
    return df, duplicate_report


def load_kiso(path: Path) -> pd.DataFrame:
    raw = read_excel_auto_header(path)

    cols = resolve_columns(
        raw,
        KISO_COLUMNS,
        positional_fallbacks=KISO_POSITIONAL_FALLBACKS,
        required=True,
    )

    length = len(raw)
    df = pd.DataFrame({
        "category": get_series(raw, cols["category"], length).map(normalize_category),
        "name": get_series(raw, cols["name"], length).map(normalize_name),
        "ingredient": get_series(raw, cols["ingredient"], length).map(normalize_ingredient),
        "spec": get_series(raw, cols["spec"], length).map(normalize_spec),
        "manufacturer": get_series(raw, cols["manufacturer"], length).map(normalize_manufacturer),
        "price": get_series(raw, cols["price"], length).map(parse_price),
    })

    df = df[
        (df["name"] != "")
        & (df["ingredient"] != "")
        & (df["spec"] != "")
    ].copy()

    validate_kiso_df(df)
    return df


def load_manual_aliases(path: Optional[Path]) -> list[dict[str, Any]]:
    if path is None or not path.exists():
        return []
    with path.open("r", encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, list):
        raise ValueError("manual_aliases.json は配列である必要があります")
    return data


def load_all_sources(
    input_dir: Path,
    manual_aliases_path: Optional[Path],
    file_config: InputFileConfig,
) -> dict[str, Any]:
    base_paths = [input_dir / name for name in file_config.base_files]
    aux05_path = input_dir / file_config.aux05_file
    kiso_path = input_dir / file_config.kiso_file

    base = load_base_master(base_paths)
    aux05, aux05_duplicate_report = load_aux05(aux05_path)
    kiso = load_kiso(kiso_path)
    manual_aliases = load_manual_aliases(manual_aliases_path)

    temporary_exception_path = (
        input_dir / file_config.temporary_exception_file
        if file_config.temporary_exception_file
        else None
    )

    return {
        "base": base,
        "aux05": aux05,
        "kiso": kiso,
        "manual_aliases": manual_aliases,
        "source_files": [p.name for p in base_paths] + [aux05_path.name, kiso_path.name],
        "input_file_config": asdict(file_config),
        "aux05_duplicate_report": aux05_duplicate_report,
        "unprocessed_optional_files": {
            "temporary_exception_file": {
                "file_name": file_config.temporary_exception_file,
                "exists": bool(temporary_exception_path and temporary_exception_path.exists()),
                "processed": False,
                "reason": "v6.2では tp20260520-01_06.xlsx 等の臨時的取扱い表は未処理。必要時は別テーブル化する。",
            }
        },
    }


# =========================================================
# join / flags / search
# =========================================================

def join_aux05(base: pd.DataFrame, aux05: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, int]]:
    joined = base.merge(aux05, how="left", on="code")

    base_codes = set(base["code"].dropna().tolist())
    aux_codes = set(aux05["code"].dropna().tolist())

    report = {
        "base_has_no_aux05": len(base_codes - aux_codes),
        "aux05_has_no_base": len(aux_codes - base_codes),
        "base_and_aux05_matched": len(base_codes & aux_codes),
    }
    return joined, report


def collect_mark_other(row: pd.Series) -> list[str]:
    values = [normalize_text(row.get("mark_e")), normalize_text(row.get("mark_f")), normalize_text(row.get("mark_g"))]
    cleaned = []
    for v in values:
        if v and v not in {"局", "麻"}:
            cleaned.append(v)

    seen = set()
    result = []
    for v in cleaned:
        if v not in seen:
            seen.add(v)
            result.append(v)
    return result


def derive_flags(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    e = df["mark_e"].fillna("").map(lambda x: normalize_text(x, COMMON_TOKEN_ALIASES))
    f = df["mark_f"].fillna("").map(lambda x: normalize_text(x, COMMON_TOKEN_ALIASES))
    g = df["mark_g"].fillna("").map(lambda x: normalize_text(x, COMMON_TOKEN_ALIASES))

    df["mark_jp"] = e.eq("局") | f.eq("局") | g.eq("局")
    df["mark_narcotic"] = e.eq("麻") | f.eq("麻") | g.eq("麻")
    df["mark_other"] = df.apply(collect_mark_other, axis=1)

    raw_j = df["raw_j"].fillna("").map(lambda x: normalize_text(x, COMMON_TOKEN_ALIASES))
    raw_k = df["raw_k"].fillna("").map(lambda x: normalize_text(x, COMMON_TOKEN_ALIASES))
    raw_l = df["raw_l"].fillna("").map(lambda x: normalize_text(x, COMMON_TOKEN_ALIASES))
    raw_n = df["raw_n"].fillna("").map(lambda x: normalize_text(x, COMMON_TOKEN_ALIASES))
    aux05_f = df["aux05_f"].fillna("").map(lambda x: normalize_text(x, COMMON_TOKEN_ALIASES))
    aux05_h = df["aux05_h"].fillna("").map(lambda x: normalize_text(x, COMMON_TOKEN_ALIASES))

    df["is_generic_designated"] = raw_j.eq("後発品")
    df["is_originator_designated"] = raw_k.isin(["先発品", "準先発品"])
    df["has_generic_equivalent"] = raw_l.isin(["○", "有"])

    df["generic_price_marker"] = raw_j.where(raw_j.eq("★"), "")
    df["originator_price_marker"] = raw_j.where(raw_j.eq("☆"), "")

    df["generic_info_available"] = raw_j.ne("") | raw_k.ne("") | raw_l.ne("")

    df["is_transitional"] = raw_n.ne("") | aux05_f.ne("")
    df["excluded_from_addon"] = aux05_h.ne("")

    return df


def infer_dosage_form(name: str) -> str:
    for pat, label in DOSAGE_PATTERNS:
        if re.search(pat, name, re.IGNORECASE):
            return label
    if "錠" in name:
        return "錠"
    return "その他"


def add_search_fields(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["dosage_form"] = df["name"].map(infer_dosage_form)

    df["display_group_key"] = (
        df["ingredient"].fillna("")
        + "|"
        + df["spec"].fillna("")
        + "|"
        + df["dosage_form"].fillna("")
    )

    df["search_text"] = (
        df["name"].fillna("")
        + " "
        + df["ingredient"].fillna("")
        + " "
        + df["manufacturer"].fillna("")
    ).str.strip()

    return df


def build_exact6_key(df: pd.DataFrame) -> pd.Series:
    return (
        df["category"].fillna("").astype(str)
        + "||"
        + df["name"].fillna("").astype(str)
        + "||"
        + df["ingredient"].fillna("").astype(str)
        + "||"
        + df["spec"].fillna("").astype(str)
        + "||"
        + df["manufacturer"].fillna("").astype(str)
        + "||"
        + df["price"].map(decimal_key)
    )


def build_key(df: pd.DataFrame, cols: list[str]) -> pd.Series:
    def part(col: str) -> pd.Series:
        if col == "price":
            return df[col].map(decimal_key)
        return df[col].fillna("").astype(str)

    s = part(cols[0])
    for c in cols[1:]:
        s = s + "||" + part(c)
    return s


def prices_equal_or_missing(left: Any, right: Any) -> bool:
    if is_missing(left) or is_missing(right):
        return True
    try:
        return Decimal(str(left)) == Decimal(str(right))
    except Exception:
        return safe_str(left) == safe_str(right)


def join_kiso(base: pd.DataFrame, kiso: pd.DataFrame, manual_aliases: list[dict[str, Any]]) -> tuple[pd.DataFrame, dict[str, int]]:
    base = base.copy()

    report = {
        "exact6": 0,
        "unique_ing_spec_mfr": 0,
        "unique_ing_spec_price": 0,
        "manual_alias": 0,
        "no_match": 0,
        "unresolved_ambiguous": 0,
        "ambiguous_events": 0,
    }

    base["is_basic_medicine"] = None
    base["kiso_match_rule"] = None
    base["kiso_match_confidence"] = None
    base["kiso_match_status"] = "unresolved"

    kiso_exact_keys = set(build_exact6_key(kiso).tolist())
    base_exact_keys = build_exact6_key(base)
    exact_mask = base_exact_keys.isin(kiso_exact_keys)

    base.loc[exact_mask, "is_basic_medicine"] = True
    base.loc[exact_mask, "kiso_match_rule"] = "exact6"
    base.loc[exact_mask, "kiso_match_confidence"] = 1.00
    base.loc[exact_mask, "kiso_match_status"] = "matched"
    report["exact6"] = int(exact_mask.sum())

    unresolved_mask = base["is_basic_medicine"].isna()

    cols2 = ["category", "ingredient", "spec", "manufacturer"]
    k2 = build_key(kiso, cols2)
    counts2 = k2.value_counts().to_dict()

    for idx in base[unresolved_mask].index:
        row = base.loc[idx]
        key = "||".join([safe_str(row[c]) for c in cols2])
        cnt = counts2.get(key, 0)

        if cnt == 1:
            hit = kiso[k2 == key].iloc[0]
            if prices_equal_or_missing(row["price"], hit["price"]):
                base.at[idx, "is_basic_medicine"] = True
                base.at[idx, "kiso_match_rule"] = "unique_ing_spec_mfr"
                base.at[idx, "kiso_match_confidence"] = 0.95
                base.at[idx, "kiso_match_status"] = "matched"
                report["unique_ing_spec_mfr"] += 1
        elif cnt > 1:
            base.at[idx, "kiso_match_status"] = "ambiguous"
            report["ambiguous_events"] += 1

    unresolved_mask = base["is_basic_medicine"].isna()

    cols3 = ["category", "ingredient", "spec", "price"]
    k3 = build_key(kiso, cols3)
    counts3 = k3.value_counts().to_dict()

    for idx in base[unresolved_mask].index:
        row = base.loc[idx]
        key_parts = []
        for c in cols3:
            key_parts.append(decimal_key(row[c]) if c == "price" else safe_str(row[c]))
        key = "||".join(key_parts)
        cnt = counts3.get(key, 0)

        if cnt == 1:
            base.at[idx, "is_basic_medicine"] = True
            base.at[idx, "kiso_match_rule"] = "unique_ing_spec_price"
            base.at[idx, "kiso_match_confidence"] = 0.90
            base.at[idx, "kiso_match_status"] = "matched"
            report["unique_ing_spec_price"] += 1
        elif cnt > 1:
            base.at[idx, "kiso_match_status"] = "ambiguous"
            report["ambiguous_events"] += 1

    unresolved_mask = base["is_basic_medicine"].isna()

    for idx in base[unresolved_mask].index:
        row = base.loc[idx]
        for alias in manual_aliases:
            if (
                row["name"] == normalize_name(alias.get("kiso_name"))
                and row["manufacturer"] == normalize_manufacturer(alias.get("kiso_manufacturer"))
                and row["ingredient"] == normalize_ingredient(alias.get("ingredient"))
                and row["spec"] == normalize_spec(alias.get("spec"))
            ):
                base.at[idx, "is_basic_medicine"] = True
                base.at[idx, "kiso_match_rule"] = "manual_alias"
                base.at[idx, "kiso_match_confidence"] = 0.85
                base.at[idx, "kiso_match_status"] = "matched"
                report["manual_alias"] += 1
                break

    unresolved_mask = base["is_basic_medicine"].isna()
    ambiguous_mask = unresolved_mask & base["kiso_match_status"].eq("ambiguous")
    no_match_mask = unresolved_mask & ~base["kiso_match_status"].eq("ambiguous")

    base.loc[no_match_mask, "is_basic_medicine"] = False
    base.loc[no_match_mask, "kiso_match_rule"] = "no_match"
    base.loc[no_match_mask, "kiso_match_confidence"] = None
    base.loc[no_match_mask, "kiso_match_status"] = "no_match"

    # ambiguous は None のまま残す。False と区別するため。
    report["no_match"] = int(no_match_mask.sum())
    report["unresolved_ambiguous"] = int(ambiguous_mask.sum())

    return base, report


def sort_dataframe_for_output(df: pd.DataFrame) -> pd.DataFrame:
    sort_df = df.copy()

    sort_df["price_sort"] = sort_df["price"].map(price_to_sort_float)
    sort_df["originator_sort"] = sort_df["is_originator_designated"].map(lambda x: 0 if bool(x) else 1)
    sort_df["generic_equiv_sort"] = sort_df["has_generic_equivalent"].map(lambda x: 0 if bool(x) else 1)

    sort_df = sort_df.sort_values(
        by=[
            "display_group_key",
            "price_sort",
            "originator_sort",
            "generic_equiv_sort",
            "name",
            "manufacturer",
        ],
        ascending=[True, False, True, True, True, True],
        kind="stable",
    ).drop(columns=["price_sort", "originator_sort", "generic_equiv_sort"])

    return sort_df.reset_index(drop=True)


def build_flag_distribution(df: pd.DataFrame) -> dict[str, Any]:
    basic_true = int(df["is_basic_medicine"].map(value_is_true).sum())
    basic_false = int(df["is_basic_medicine"].map(value_is_false).sum())
    basic_null = int(df["is_basic_medicine"].map(lambda x: x is None or is_missing(x)).sum())

    return {
        "is_generic_designated_true": int(df["is_generic_designated"].sum()),
        "is_originator_designated_true": int(df["is_originator_designated"].sum()),
        "has_generic_equivalent_true": int(df["has_generic_equivalent"].sum()),
        "generic_info_available_true": int(df["generic_info_available"].sum()),
        "is_transitional_true": int(df["is_transitional"].sum()),
        "excluded_from_addon_true": int(df["excluded_from_addon"].sum()),
        "is_basic_medicine_true": basic_true,
        "is_basic_medicine_false": basic_false,
        "is_basic_medicine_null": basic_null,
    }


def build_sample_hits(df: pd.DataFrame, keywords: list[str], limit_per_keyword: int = 10) -> dict[str, list[dict[str, Any]]]:
    result: dict[str, list[dict[str, Any]]] = {}

    for kw in keywords:
        q = normalize_text(kw)
        hit = df[df["search_text"].str.contains(re.escape(q), na=False, regex=True)].head(limit_per_keyword)

        rows = []
        for _, r in hit.iterrows():
            rows.append({
                "code": safe_str(r["code"]),
                "name": safe_str(r["name"]),
                "ingredient": safe_str(r["ingredient"]),
                "manufacturer": safe_str(r["manufacturer"]),
                "price": decimal_to_json_price(r["price"]),
                "is_generic_designated": bool(r["is_generic_designated"]),
                "is_originator_designated": bool(r["is_originator_designated"]),
                "has_generic_equivalent": bool(r["has_generic_equivalent"]),
                "generic_price_marker": safe_str(r.get("generic_price_marker")) or None,
                "originator_price_marker": safe_str(r.get("originator_price_marker")) or None,
                "generic_info_available": bool(r["generic_info_available"]),
                "is_basic_medicine": r["is_basic_medicine"],
                "kiso_match_status": safe_str(r.get("kiso_match_status")) or None,
                "is_transitional": bool(r["is_transitional"]),
                "aux05_f": safe_str(r.get("aux05_f")) or None,
                "aux05_g": safe_str(r.get("aux05_g")) or None,
                "raw_o": safe_str(r.get("raw_o")) or None,
            })

        result[kw] = rows

    return result


def merge_source_tables(loaded: dict[str, Any]) -> tuple[pd.DataFrame, dict[str, Any]]:
    base = loaded["base"]
    aux05 = loaded["aux05"]
    kiso = loaded["kiso"]
    manual_aliases = loaded["manual_aliases"]

    merged, aux05_report = join_aux05(base, aux05)
    merged = derive_flags(merged)
    merged = add_search_fields(merged)
    merged, kiso_report = join_kiso(merged, kiso, manual_aliases)
    merged = sort_dataframe_for_output(merged)

    join_report = {
        "schema_version": SCHEMA_VERSION,
        "generator": GENERATOR_NAME,
        "input_file_config": loaded.get("input_file_config", {}),
        "source_files_processed": loaded.get("source_files", []),
        "unprocessed_optional_files": loaded.get("unprocessed_optional_files", {}),
        "base_rows": int(len(base)),
        "aux05_rows": int(len(aux05)),
        "kiso_rows": int(len(kiso)),
        "joined_rows": int(len(merged)),
        "aux05_report": aux05_report,
        "aux05_duplicate_report": loaded.get("aux05_duplicate_report", {}),
        "kiso_report": kiso_report,
        "flag_distribution": build_flag_distribution(merged),
        "sample_hits": build_sample_hits(
            merged,
            keywords=["ニフェジピン", "アダラート", "サワイ", "トーワ"],
            limit_per_keyword=10,
        ),
    }

    return merged, join_report


# =========================================================
# Builder helper
# =========================================================

def extract_brand_label(name: str) -> Optional[str]:
    m = re.search(r"「(.+?)」", name)
    if m:
        return m.group(1)
    return None


def derive_generic_code_from_mhlw_code(mhlw_code: str) -> Optional[str]:
    if len(mhlw_code) != 12:
        return None
    return mhlw_code[:9] + "ZZZ"


def make_jst_now_iso() -> str:
    if ZoneInfo is not None:
        return datetime.now(ZoneInfo("Asia/Tokyo")).isoformat(timespec="seconds")
    return datetime.now().isoformat(timespec="seconds")


def build_meta_info(source_files: list[str]) -> MetaInfo:
    return MetaInfo(
        schema_version=SCHEMA_VERSION,
        generated_at=make_jst_now_iso(),
        generator=GENERATOR_NAME,
        source_files=source_files,
    )


def build_flag_item(
    owner_type: str,
    owner_id: str,
    flag_type: str,
    value: Any,
    value_type: str,
    columns: list[str],
    values: list[str],
    derived_by: str,
    confidence: Optional[float] = 1.0,
    ui_badge: Optional[str] = None,
) -> FlagItem:
    return FlagItem(
        flag_id=make_flag_id(owner_type, owner_id, flag_type),
        owner_type=owner_type,
        owner_id=owner_id,
        flag_type=flag_type,
        value_type=value_type,
        value=value,
        raw_source=RawSource(columns=columns, values=values),
        derived_by=derived_by,
        confidence=confidence,
        ui_badge=ui_badge,
    )


def pick_representative_value(series: pd.Series) -> str:
    values = [safe_str(v) for v in series.tolist() if safe_str(v) != ""]
    if not values:
        return ""
    vc = pd.Series(values).value_counts()
    return str(vc.index[0])


def pick_representative_decimal(series: pd.Series) -> Optional[Decimal]:
    values = [v for v in series.tolist() if not is_missing(v)]
    if not values:
        return None
    vc = pd.Series([str(v) for v in values]).value_counts()
    top = str(vc.index[0])
    try:
        return Decimal(top)
    except Exception:
        return values[0]


def pick_optional_float(series: pd.Series) -> Optional[float]:
    for value in series.tolist():
        f = to_optional_float(value)
        if f is not None:
            return f
    return None


def collect_price_item_conflicts(code: str, group_df: pd.DataFrame) -> dict[str, list[str]]:
    _ = code
    result: dict[str, list[str]] = {}

    for col in ["category", "ingredient", "spec"]:
        vals = sorted({safe_str(v) for v in group_df[col].tolist() if safe_str(v) != ""})
        if len(vals) > 1:
            result[col] = vals

    price_vals = sorted({decimal_key(v) for v in group_df["price"].tolist() if decimal_key(v) != ""})
    if len(price_vals) > 1:
        result["price"] = price_vals

    return result


def merge_text_candidates(values: list[str]) -> Optional[str]:
    cleaned = [normalize_text(v) for v in values if normalize_text(v) != ""]
    if not cleaned:
        return None

    unique_values = list(dict.fromkeys(cleaned))
    return " / ".join(unique_values)


def group_any_true(group_df: pd.DataFrame, col: str) -> bool:
    return bool(group_df[col].map(value_is_true).any())


def first_row_where_true(group_df: pd.DataFrame, col: str) -> pd.Series:
    mask = group_df[col].map(value_is_true)
    if bool(mask.any()):
        return group_df[mask].iloc[0]
    return group_df.iloc[0]


def build_price_item_from_group(code: str, group_df: pd.DataFrame) -> PriceItem:
    row0 = group_df.iloc[0]

    category = pick_representative_value(group_df["category"])
    ingredient = pick_representative_value(group_df["ingredient"])
    spec = pick_representative_value(group_df["spec"])
    price_decimal = pick_representative_decimal(group_df["price"])
    price = decimal_to_json_price(price_decimal)

    conflicts = collect_price_item_conflicts(code, group_df)

    therapeutic_class_code = code[:3] if len(code) >= 3 else None
    route_code = code[3] if len(code) >= 4 else None
    dosage_form_code = code[7] if len(code) >= 8 else None

    dosage_form_name = DOSAGE_FORM_NAME_MAP.get(
        dosage_form_code or "",
        safe_str(row0.get("dosage_form")) or None,
    )
    therapeutic_class_name = THERAPEUTIC_CLASS_NAME_MAP.get(therapeutic_class_code or "")
    route_name = ROUTE_NAME_MAP.get(route_code or "")

    dosage_form_fallback = dosage_form_name or safe_str(row0.get("dosage_form"))
    display_group_key = f"{ingredient}|{spec}|{dosage_form_fallback}".strip("|")
    group_label = " ".join(
        x for x in [ingredient, spec, dosage_form_fallback] if x
    ).strip()
    search_text = " ".join(
        filter(None, [ingredient, spec, therapeutic_class_name or "", dosage_form_fallback or ""])
    ).strip()

    transitional_deadline_text_raw = pick_representative_value(group_df["aux05_f"])
    transitional_deadline_text = transitional_deadline_text_raw if transitional_deadline_text_raw != "" else None

    remarks_text = merge_text_candidates([
        pick_representative_value(group_df["aux05_g"]),
        pick_representative_value(group_df["raw_o"]),
    ])

    base_files = sorted(set(group_df["source_file"].dropna().astype(str).tolist()))
    if conflicts:
        base_files = base_files + [f"__conflict__:{json.dumps(conflicts, ensure_ascii=False)}"]

    source_summary = SourceSummary(
        base_files=base_files,
        aux05_present=series_has_nonempty(group_df["aux05_d"]) or series_has_nonempty(group_df["aux05_e"]),
        kiso_match_rule=pick_representative_value(group_df["kiso_match_rule"]) or None,
        kiso_match_confidence=pick_optional_float(group_df["kiso_match_confidence"]),
    )

    return PriceItem(
        price_item_id=make_price_item_id(code),
        mhlw_code=code,
        category=category,
        category_normalized=category,
        ingredient=ingredient,
        ingredient_normalized=ingredient,
        spec=spec,
        spec_normalized=spec,
        price=price,
        therapeutic_class_code=therapeutic_class_code,
        therapeutic_class_name=therapeutic_class_name or None,
        route_code=route_code,
        route_name=route_name or None,
        dosage_form_code=dosage_form_code,
        dosage_form_name=dosage_form_name,
        display_group_key=display_group_key,
        group_label=group_label,
        search_text=search_text,
        transitional_deadline_text=transitional_deadline_text,
        remarks_text=remarks_text,
        source_summary=source_summary,
    )


def build_brand_item(row: pd.Series, price_item_id: str, serial: int) -> BrandItem:
    mhlw_code = safe_str(row["code"])
    name = safe_str(row["name"])
    manufacturer = safe_str(row["manufacturer"])

    price_value = price_to_sort_float(row["price"])
    is_originator_designated = bool(row.get("is_originator_designated"))
    has_generic_equivalent = bool(row.get("has_generic_equivalent"))

    normalized_name = normalize_name(name)
    normalized_manufacturer = normalize_manufacturer(manufacturer)

    return BrandItem(
        brand_item_id=make_brand_item_id(mhlw_code, serial),
        price_item_id=price_item_id,
        mhlw_code=mhlw_code,
        yj_code=None,
        hot_code=None,
        name=name,
        name_normalized=normalized_name,
        manufacturer=manufacturer,
        manufacturer_normalized=normalized_manufacturer,
        brand_label=extract_brand_label(name),
        display_name=name,
        search_text=" ".join(filter(None, [name, manufacturer, safe_str(row.get("ingredient"))])).strip(),
        sort_keys=BrandSortKeys(
            price_desc=price_value,
            originator_first=0 if is_originator_designated else 1,
            generic_equivalent_first=0 if has_generic_equivalent else 1,
            name=name,
            manufacturer=manufacturer,
        ),
        source_row=SourceRowInfo(
            source_file=safe_str(row.get("source_file")),
            source_code=mhlw_code,
        ),
    )


def build_generic_name_map(price_item: PriceItem) -> Optional[GenericNameMap]:
    generic_code = derive_generic_code_from_mhlw_code(price_item.mhlw_code)
    if not generic_code:
        return None

    generic_display_name = f"【般】{price_item.ingredient}{price_item.spec}"

    return GenericNameMap(
        generic_name_map_id=make_generic_name_map_id(generic_code, price_item.mhlw_code),
        generic_code=generic_code,
        price_item_id=price_item.price_item_id,
        mhlw_code=price_item.mhlw_code,
        generic_display_name=generic_display_name,
        ingredient=price_item.ingredient,
        spec=price_item.spec,
        mapping_type="derived_from_first9_plus_ZZZ",
        mapping_status="active",
        mapping_confidence=1.0,
        source=MappingSource(
            source_name="derived",
            source_file=None,
            source_row_hint=None,
        ),
    )


def dedupe_flags(flags: list[FlagItem]) -> tuple[list[FlagItem], dict[str, Any]]:
    seen: dict[str, FlagItem] = {}
    duplicate_count = 0
    duplicate_ids: list[str] = []
    conflicting_duplicates: list[dict[str, Any]] = []

    for flag in flags:
        existing = seen.get(flag.flag_id)
        if existing is None:
            seen[flag.flag_id] = flag
            continue

        duplicate_count += 1
        duplicate_ids.append(flag.flag_id)

        if (
            existing.value != flag.value
            or existing.value_type != flag.value_type
            or existing.owner_type != flag.owner_type
            or existing.owner_id != flag.owner_id
            or existing.flag_type != flag.flag_type
        ):
            conflicting_duplicates.append({
                "flag_id": flag.flag_id,
                "kept": {
                    "owner_type": existing.owner_type,
                    "owner_id": existing.owner_id,
                    "flag_type": existing.flag_type,
                    "value_type": existing.value_type,
                    "value": existing.value,
                },
                "dropped": {
                    "owner_type": flag.owner_type,
                    "owner_id": flag.owner_id,
                    "flag_type": flag.flag_type,
                    "value_type": flag.value_type,
                    "value": flag.value,
                },
            })

    deduped = list(seen.values())
    report = {
        "input_count": len(flags),
        "output_count": len(deduped),
        "duplicate_count": duplicate_count,
        "duplicate_id_samples": duplicate_ids[:20],
        "conflicting_duplicate_samples": conflicting_duplicates[:20],
    }
    return deduped, report


# =========================================================
# Bundle builder
# =========================================================

def build_price_items(merged_df: pd.DataFrame) -> list[PriceItem]:
    price_items: list[PriceItem] = []

    grouped = merged_df.groupby("code", sort=False)
    for code, group in grouped:
        price_items.append(build_price_item_from_group(str(code), group))

    return price_items


def build_brand_items(merged_df: pd.DataFrame) -> list[BrandItem]:
    brand_items: list[BrandItem] = []

    grouped = merged_df.groupby("code", sort=False)
    for code, group in grouped:
        price_item_id = make_price_item_id(str(code))
        seen_brand_keys: set[tuple[str, str, str]] = set()
        serial = 1

        for _, row in group.iterrows():
            normalized_name = normalize_name(row["name"])
            normalized_manufacturer = normalize_manufacturer(row["manufacturer"])

            brand_key = (
                safe_str(row["code"]),
                normalized_name,
                normalized_manufacturer,
            )
            if brand_key in seen_brand_keys:
                continue
            seen_brand_keys.add(brand_key)

            brand_item = build_brand_item(row, price_item_id, serial)
            brand_item.name_normalized = normalized_name
            brand_item.manufacturer_normalized = normalized_manufacturer

            brand_items.append(brand_item)
            serial += 1

    return brand_items


def build_flags(merged_df: pd.DataFrame, brand_items: list[BrandItem]) -> list[FlagItem]:
    flags: list[FlagItem] = []

    brand_lookup: dict[tuple[str, str, str], BrandItem] = {
        (b.mhlw_code, b.name_normalized, b.manufacturer_normalized): b
        for b in brand_items
    }

    for _, row in merged_df.iterrows():
        key = (
            safe_str(row["code"]),
            normalize_name(row["name"]),
            normalize_manufacturer(row["manufacturer"]),
        )
        brand_item = brand_lookup.get(key)
        if not brand_item:
            continue

        if bool(row.get("is_generic_designated")):
            flags.append(
                build_flag_item(
                    owner_type="brand_item",
                    owner_id=brand_item.brand_item_id,
                    flag_type="is_generic_designated",
                    value=True,
                    value_type="boolean",
                    columns=["raw_j"],
                    values=[safe_str(row.get("raw_j"))],
                    derived_by="rule:raw_j == 後発品",
                    ui_badge="後発指定",
                )
            )

        if bool(row.get("is_originator_designated")):
            flags.append(
                build_flag_item(
                    owner_type="brand_item",
                    owner_id=brand_item.brand_item_id,
                    flag_type="is_originator_designated",
                    value=True,
                    value_type="boolean",
                    columns=["raw_k"],
                    values=[safe_str(row.get("raw_k"))],
                    derived_by="rule:raw_k in [先発品, 準先発品]",
                    ui_badge="先発指定",
                )
            )

        if bool(row.get("has_generic_equivalent")):
            flags.append(
                build_flag_item(
                    owner_type="brand_item",
                    owner_id=brand_item.brand_item_id,
                    flag_type="has_generic_equivalent",
                    value=True,
                    value_type="boolean",
                    columns=["raw_l"],
                    values=[safe_str(row.get("raw_l"))],
                    derived_by="rule:raw_l in [○, 有]",
                    ui_badge="同規格後発あり",
                )
            )

        generic_price_marker = safe_str(row.get("generic_price_marker"))
        if generic_price_marker:
            flags.append(
                build_flag_item(
                    owner_type="brand_item",
                    owner_id=brand_item.brand_item_id,
                    flag_type="generic_price_marker",
                    value=generic_price_marker,
                    value_type="string",
                    columns=["raw_j"],
                    values=[safe_str(row.get("raw_j"))],
                    derived_by="rule:raw_j == ★",
                    ui_badge=generic_price_marker,
                )
            )

        originator_price_marker = safe_str(row.get("originator_price_marker"))
        if originator_price_marker:
            flags.append(
                build_flag_item(
                    owner_type="brand_item",
                    owner_id=brand_item.brand_item_id,
                    flag_type="originator_price_marker",
                    value=originator_price_marker,
                    value_type="string",
                    columns=["raw_j"],
                    values=[safe_str(row.get("raw_j"))],
                    derived_by="rule:raw_j == ☆",
                    ui_badge=originator_price_marker,
                )
            )

        flags.append(
            build_flag_item(
                owner_type="brand_item",
                owner_id=brand_item.brand_item_id,
                flag_type="generic_info_available",
                value=bool(row.get("generic_info_available")),
                value_type="boolean",
                columns=["raw_j", "raw_k", "raw_l"],
                values=[
                    safe_str(row.get("raw_j")),
                    safe_str(row.get("raw_k")),
                    safe_str(row.get("raw_l")),
                ],
                derived_by="rule:raw_j != '' or raw_k != '' or raw_l != ''",
                ui_badge=None,
            )
        )

    grouped = merged_df.groupby("code", sort=False)
    for code, group in grouped:
        price_item_id = make_price_item_id(str(code))

        if group_any_true(group, "mark_jp"):
            row = first_row_where_true(group, "mark_jp")
            flags.append(
                build_flag_item(
                    owner_type="price_item",
                    owner_id=price_item_id,
                    flag_type="mark_jp",
                    value=True,
                    value_type="boolean",
                    columns=["mark_e", "mark_f", "mark_g"],
                    values=[
                        safe_str(row.get("mark_e")),
                        safe_str(row.get("mark_f")),
                        safe_str(row.get("mark_g")),
                    ],
                    derived_by="rule:any(mark_e,mark_f,mark_g)==局",
                    ui_badge="局",
                )
            )

        if group_any_true(group, "mark_narcotic"):
            row = first_row_where_true(group, "mark_narcotic")
            flags.append(
                build_flag_item(
                    owner_type="price_item",
                    owner_id=price_item_id,
                    flag_type="mark_narcotic",
                    value=True,
                    value_type="boolean",
                    columns=["mark_e", "mark_f", "mark_g"],
                    values=[
                        safe_str(row.get("mark_e")),
                        safe_str(row.get("mark_f")),
                        safe_str(row.get("mark_g")),
                    ],
                    derived_by="rule:any(mark_e,mark_f,mark_g)==麻",
                    ui_badge="麻",
                )
            )

        mark_other_values: list[str] = []
        for v in group["mark_other"].tolist():
            if isinstance(v, list):
                mark_other_values.extend(v)
        mark_other = list(dict.fromkeys([v for v in mark_other_values if v]))
        if mark_other:
            flags.append(
                build_flag_item(
                    owner_type="price_item",
                    owner_id=price_item_id,
                    flag_type="mark_other",
                    value=mark_other,
                    value_type="string_list",
                    columns=["mark_e", "mark_f", "mark_g"],
                    values=[
                        pick_representative_value(group["mark_e"]),
                        pick_representative_value(group["mark_f"]),
                        pick_representative_value(group["mark_g"]),
                    ],
                    derived_by="rule:collect_mark_other",
                    ui_badge=None,
                )
            )

        if group_any_true(group, "is_transitional"):
            row = first_row_where_true(group, "is_transitional")
            flags.append(
                build_flag_item(
                    owner_type="price_item",
                    owner_id=price_item_id,
                    flag_type="is_transitional",
                    value=True,
                    value_type="boolean",
                    columns=["raw_n", "aux05_f", "aux05_g", "raw_o"],
                    values=[
                        safe_str(row.get("raw_n")),
                        safe_str(row.get("aux05_f")),
                        safe_str(row.get("aux05_g")),
                        safe_str(row.get("raw_o")),
                    ],
                    derived_by="rule:any(raw_n != '' or aux05_f != '') in code group",
                    ui_badge="経過措置",
                )
            )

        if group_any_true(group, "excluded_from_addon"):
            row = first_row_where_true(group, "excluded_from_addon")
            flags.append(
                build_flag_item(
                    owner_type="price_item",
                    owner_id=price_item_id,
                    flag_type="excluded_from_addon",
                    value=True,
                    value_type="boolean",
                    columns=["aux05_h"],
                    values=[safe_str(row.get("aux05_h"))],
                    derived_by="rule:any(aux05_h != '') in code group",
                    ui_badge="除外",
                )
            )

        basic_mask = group["is_basic_medicine"].map(value_is_true)
        if bool(basic_mask.any()):
            row = group[basic_mask].iloc[0]
            conf = to_optional_float(row.get("kiso_match_confidence"))
            flags.append(
                build_flag_item(
                    owner_type="price_item",
                    owner_id=price_item_id,
                    flag_type="is_basic_medicine",
                    value=True,
                    value_type="boolean",
                    columns=["category", "ingredient", "spec", "manufacturer", "price"],
                    values=[
                        safe_str(row.get("category")),
                        safe_str(row.get("ingredient")),
                        safe_str(row.get("spec")),
                        safe_str(row.get("manufacturer")),
                        decimal_key(row.get("price")),
                    ],
                    derived_by=f"rule:{safe_str(row.get('kiso_match_rule')) or 'unknown'}",
                    confidence=conf,
                    ui_badge="基礎的",
                )
            )

    return flags


def build_receipt_code_maps(merged_df: pd.DataFrame) -> list[ReceiptCodeMap]:
    _ = merged_df
    # TODO:
    #   現時点の入力Excel群には、レセプト電算コードへ確実に接続できる列を指定していない。
    #   不確実な推定マップを作るより、空配列として明示する。
    #   導入時は source_name / mapping_confidence / mapping_status を必ず付与する。
    return []


def build_generic_name_maps(price_items: list[PriceItem]) -> list[GenericNameMap]:
    result: list[GenericNameMap] = []
    for price_item in price_items:
        item = build_generic_name_map(price_item)
        if item:
            result.append(item)
    return result


def attach_relations(
    price_items: list[PriceItem],
    brand_items: list[BrandItem],
    receipt_code_maps: list[ReceiptCodeMap],
    generic_name_maps: list[GenericNameMap],
    flags: list[FlagItem],
) -> None:
    price_map = {p.price_item_id: p for p in price_items}
    brand_map = {b.brand_item_id: b for b in brand_items}

    for b in brand_items:
        if b.price_item_id in price_map:
            if b.brand_item_id not in price_map[b.price_item_id].brand_item_ids:
                price_map[b.price_item_id].brand_item_ids.append(b.brand_item_id)

    for r in receipt_code_maps:
        if r.price_item_id in price_map:
            if r.receipt_code_map_id not in price_map[r.price_item_id].receipt_code_ids:
                price_map[r.price_item_id].receipt_code_ids.append(r.receipt_code_map_id)

    for g in generic_name_maps:
        if g.price_item_id in price_map:
            if g.generic_name_map_id not in price_map[g.price_item_id].generic_name_map_ids:
                price_map[g.price_item_id].generic_name_map_ids.append(g.generic_name_map_id)

    for f in flags:
        if f.owner_type == "price_item" and f.owner_id in price_map:
            if f.flag_id not in price_map[f.owner_id].flag_ids:
                price_map[f.owner_id].flag_ids.append(f.flag_id)
        elif f.owner_type == "brand_item" and f.owner_id in brand_map:
            if f.flag_id not in brand_map[f.owner_id].flag_ids:
                brand_map[f.owner_id].flag_ids.append(f.flag_id)


def build_master_bundle(merged_df: pd.DataFrame, join_report: dict[str, Any]) -> MasterBundle:
    price_items = build_price_items(merged_df)
    brand_items = build_brand_items(merged_df)

    raw_flags = build_flags(merged_df, brand_items)
    flags, flag_dedupe_report = dedupe_flags(raw_flags)

    receipt_code_maps = build_receipt_code_maps(merged_df)
    generic_name_maps = build_generic_name_maps(price_items)

    join_report = dict(join_report)
    join_report["flag_dedupe_report"] = flag_dedupe_report
    join_report["implementation_notes"] = {
        "receipt_code_maps": "未実装。入力ソース未指定のため空配列。",
        "price_json_type": "string Decimal。Swift側では Decimal(string:) 推奨。",
        "is_basic_medicine": "true=一致、false=明確な非一致、null=曖昧で未解決。",
        "temporary_exception_06": "tp20260520-01_06.xlsx 等の臨時的取扱い表はv6.2では未処理。必要時は temporary_exception_items 等として別テーブル化する。",
        "aux05_d": "aux05_d は読込済みだが、1/2/3/☆/★ の正式フラグ化は次版以降。現時点の後発品関連判定は主に base 表の raw_j/raw_k/raw_l に基づく。",
        "aux05_dedupe": "aux05 は code 単位で先頭1件に dedupe。重複状況は aux05_duplicate_report を参照。",
    }

    attach_relations(
        price_items=price_items,
        brand_items=brand_items,
        receipt_code_maps=receipt_code_maps,
        generic_name_maps=generic_name_maps,
        flags=flags,
    )

    return MasterBundle(
        meta=MetaInfo(
            schema_version=SCHEMA_VERSION,
            generated_at="",
            generator=GENERATOR_NAME,
            source_files=[],
        ),
        price_items=price_items,
        brand_items=brand_items,
        receipt_code_maps=receipt_code_maps,
        generic_name_maps=generic_name_maps,
        flags=flags,
        reports=ReportBundle(join_report=join_report),
    )


def build_master_bundle_from_merged(
    merged_df: pd.DataFrame,
    source_files: list[str],
    join_report: dict[str, Any],
) -> MasterBundle:
    bundle = build_master_bundle(merged_df=merged_df, join_report=join_report)
    bundle.meta = build_meta_info(source_files)
    return bundle


# =========================================================
# Writer
# =========================================================

def write_json_safe(data: Any, out_path: Path) -> None:
    clean_data = clean_json_value(data)
    with out_path.open("w", encoding="utf-8") as f:
        json.dump(clean_data, f, ensure_ascii=False, indent=2, allow_nan=False)


def write_master_bundle_json(bundle: MasterBundle, out_path: Path) -> None:
    write_json_safe(asdict(bundle), out_path)


def write_table_json(records: list[Any], out_path: Path) -> None:
    write_json_safe([asdict(x) for x in records], out_path)


def write_report_json(report: dict[str, Any], out_path: Path) -> None:
    write_json_safe(report, out_path)


def write_bundle_outputs(bundle: MasterBundle, output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    write_master_bundle_json(bundle, output_dir / "master_bundle.json")
    write_table_json(bundle.price_items, output_dir / "price_items.json")
    write_table_json(bundle.brand_items, output_dir / "brand_items.json")
    write_table_json(bundle.receipt_code_maps, output_dir / "receipt_code_maps.json")
    write_table_json(bundle.generic_name_maps, output_dir / "generic_name_maps.json")
    write_table_json(bundle.flags, output_dir / "flags.json")
    write_report_json(bundle.reports.join_report, output_dir / "join_report.json")


# =========================================================
# CLI
# =========================================================

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="MHLW Excel -> JSON converter")
    parser.add_argument("--input-dir", type=Path, default=Path("."), help="Excelファイルのあるディレクトリ")
    parser.add_argument("--output-dir", type=Path, default=Path("output"), help="出力ディレクトリ")
    parser.add_argument("--manual-aliases", type=Path, default=None, help="manual_aliases.json のパス")
    parser.add_argument("--config", type=Path, default=None, help="入力ファイル名設定JSONのパス")
    return parser.parse_args()


def ensure_input_files(input_dir: Path, file_config: InputFileConfig) -> None:
    required = list(file_config.base_files) + [file_config.aux05_file, file_config.kiso_file]
    missing = [name for name in required if not (input_dir / name).exists()]
    if missing:
        raise FileNotFoundError(
            "必要ファイルが見つかりません:\n" + "\n".join(f" - {name}" for name in missing)
        )


# =========================================================
# Runner
# =========================================================

def run_converter(args: argparse.Namespace) -> None:
    file_config = load_input_file_config(args.config)
    ensure_input_files(args.input_dir, file_config)

    loaded = load_all_sources(
        input_dir=args.input_dir,
        manual_aliases_path=args.manual_aliases,
        file_config=file_config,
    )

    merged_df, join_report = merge_source_tables(loaded)

    bundle = build_master_bundle_from_merged(
        merged_df=merged_df,
        source_files=loaded["source_files"],
        join_report=join_report,
    )

    write_bundle_outputs(bundle, args.output_dir)

    print("done")
    print(json.dumps(clean_json_value(bundle.reports.join_report), ensure_ascii=False, indent=2, allow_nan=False))


def main() -> int:
    args = parse_args()

    try:
        run_converter(args)
        return 0
    except Exception as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
