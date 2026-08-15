# yakka_to_machine-language
Convert the publicly available primary source data for Japanese medical drug prices into JSON format.


## Project Background

I am a pharmacist working at a dispensing pharmacy in Japan. 
In recent years, I have needed to check drug prices more often than before.
One reason is that I sometimes need to confirm drug prices when handling prescription changes based on dispensing rules. 
Also, in daily work, I do not always have enough time to sit in front of a computer for a long time. 
Because of this, I started to want a simple and practical way to prepare drug price data for easier use.
Drug prices in Japan are revised every year. 
There can also be revisions during the fiscal year, and new drugs are listed as well.
So I felt that checking and organizing drug price data has become something I need to do continuously in daily work.

## Why This Converter Exists

In Japan, prescription drugs have several kinds of codes, such as the MHLW code, the YJ code, and the receipt-computer code.
The list of MHLW codes is publicly available. 
YJ codes can also be obtained, but in that case, the purpose of use and the place of use must be clearly stated.
Because of this situation,
I decided to work mainly with primary source data published by the Ministry of Health, Labour and Welfare. 
I wanted to prepare a data source that could be processed locally and used as the basis for a standalone workflow.
This repository is for the Excel-to-JSON converter that organizes official spreadsheet data into a format that is easier to handle in software.

## Scope of This Repository

This public repository contains the data converter only.
The app itself is being developed separately in a private repository. 
I chose to keep the application code private for now, while making the converter public because the conversion process itself may still be useful to others who want to work with the official Excel files.

## Data Source Policy

This project uses primary source data published by the Ministry of Health, Labour and Welfare.
Since medical DX is being promoted by the MHLW in Japan, 
I first thought that official data provided by the government would already be organized in a way that would be easier to use.
In reality, I felt there was still room to reorganize the data into a more practical form for day-to-day work.
So this project started from a simple idea: 
if the official data exists, I want to convert it into a form that is easier to reuse and handle.
This converter is not intended to replace official information. 
It is only a small tool to help process official data more easily.

### Official MHLW Source

The converter uses the official drug price list and related data published by the Ministry of Health, Labour and Welfare (MHLW).

The MHLW source page is updated over time as drug price information changes. Therefore, each converter version in this repository identifies its source data by the Excel filenames configured in the script rather than by assuming that the current contents of the MHLW page correspond to that converter version.

Official source:
https://www.mhlw.go.jp/topics/2026/04/tp20260401-01.html

See the Versions section below for the default input set used by each converter version.

## Versions

This repository preserves several versions of the converter to document how the conversion logic evolved over time.

| Version | Default input set | Notes                                                                                                                      |
| ------- | ----------------- | -------------------------------------------------------------------------------------------------------------------------- |
| v3      | April 1, 2026     | Early implementation using the April 1 MHLW files                                                                          |
| v4      | April 1, 2026     | Intermediate development version based on the same April 1 source set                                                      |
| v5      | April 15, 2026    | Updated for the April 15 revision; `_04` continues to use the April 1 file                                                 |
| v6.2    | May 20, 2026      | Current version; `_04` continues to use the April 1 file. `_06` is recorded as an optional source but is not yet processed |

All versions use the 2026 basic-medicine (`kiso`) source file where applicable.

Older versions are intentionally retained as part of the development history.

For new use, `mhlw_price_converter_v6_2.py` is the recommended version.

## Requirements

- Python 3
- pandas
- openpyxl

Install dependencies with:

```bash
pip install -r requirements.txt
