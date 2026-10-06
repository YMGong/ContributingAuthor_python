# CA Registration Sync Tool

`Sync_CA_registration.py` is a Python tool for processing Contributing Author (CA) registration data collected through Microsoft Forms and transferring it into a standardised Excel workbook.

The script was developed for a workflow in which the Microsoft Forms response workbook and the destination workbook are stored in different Microsoft 365 / SharePoint environments. Rather than connecting directly to SharePoint, the script uses a simple **download → process → upload** workflow.

## Overview

The script:

- imports selected fields from a Microsoft Forms Excel response file;
- transfers them to the `Data` worksheet of an existing Excel template;
- standardises selected fields;
- identifies multiple submissions from the same person;
- reconciles information across duplicate submissions;
- applies formatting rules;
- generates chapter-level gender and regional statistics;
- creates pie charts in an external-CA dashboard; and
- saves the processed data as a new Excel workbook.

The original source and template workbooks are not modified.

## Requirements

- Python 3.9 or later
- `openpyxl`

Install the required Python package with:

```bash
pip install openpyxl
```

The script also uses `tkinter` for file-selection dialogs. `tkinter` is included with most standard Python installations.

## Running the tool

Run:

```bash
python Sync_CA_registration.py
```

You will be asked to select:

1. the downloaded Microsoft Forms response workbook;
2. the destination/template workbook; and
3. the location where the processed workbook should be saved.

The default output filename is:

```text
02. WGI CA list.xlsx
```

## Workflow

```text
Microsoft Forms
      │
      ▼
Download response workbook
      │
      ▼
Run Sync_CA_registration.py
      │
      ├── Select Forms response workbook
      │
      ├── Select destination/template workbook
      │
      └── Select output location
      │
      ▼
02. WGI CA list.xlsx
      │
      ▼
Review output
      │
      ▼
Upload/replace workbook in SharePoint
```

## Excel column mapping

The script transfers the following columns from the Forms response workbook to the destination `Data` worksheet:

| Source | Destination |
|:------:|:-----------:|
| A | A |
| H | B |
| I | C |
| J | D |
| O | E |
| Q | F |
| P | G |
| W | H |
| K | I |
| L | J |
| M | K |
| N | L |
| R | M |
| S | N |
| T | O |
| U | P |
| V | Q |

Source data are read from row 2 onwards. The header row of the destination worksheet is preserved.

## Data standardisation

### Gender

Values from source column J are converted to:

| Forms value | Destination value |
|---|---|
| Male | M |
| Female | F |

### Yes/No field

Values from source column K are converted to:

| Forms value | Destination value |
|---|---:|
| Yes | 1 |
| No | 0 |

### Chapter

Chapter information from source column W is converted from text to a chapter number.

For example:

```text
Chapter 1  →  1
Chapter 5  →  5
Chapter 10 → 10
```

The conversion is tolerant of differences in capitalisation and spacing.

## Family-name standardisation

Family names are standardised before duplicate submissions are identified.

The script:

- removes leading and trailing spaces;
- collapses repeated spaces;
- converts Latin names to uppercase; and
- removes Latin accents/diacritics.

For example:

```text
García  → GARCIA
Müller  → MULLER
Dvořák  → DVORAK
```

Non-Latin scripts are preserved rather than transliterated into Latin characters.

## Handling multiple submissions

A person is identified by the combination of:

```text
Family name + Given name
```

If the same person submitted the registration form more than once, selected information is reconciled across **all of their submissions**.

### CLA/LA-related fields

For each person:

- if source column K contains `No` in **any** submission, destination column I is set to `0` for all of that person's records;
- if source column L is empty in **any** submission, destination column J is empty for all of that person's records;
- if source column M is empty in **any** submission, destination column K is empty for all of that person's records.

Reconciliation is completed before dashboard statistics are calculated.

## Formatting

Family names in destination column B are displayed in red when the person's final reconciled value in destination column I is:

```text
1
```

Other existing formatting from the template is retained where applicable.

The workbook can also be configured to use **Arial Narrow** as its standard font.

## Dashboard

The script generates/updates:

```text
Dashboard_externalCAs
```

The dashboard is restricted to:

> Authors that are not a AR7 WGI CLA or LA.

This filter is based on the final reconciled value of destination column K.

## Avoiding double counting

Repeated submissions should not cause the same author to be counted multiple times for the same chapter.

The counting unit used by the dashboard is:

```text
Person + Chapter
```

For example, if one person submits the form three times for Chapter 4, they are counted **once** for Chapter 4.

If that person is associated with both Chapter 4 and Chapter 7, they can be counted once in each chapter.

If multiple submissions from the same person for the same chapter contain conflicting gender or regional information, the **latest submission**, according to destination column A, is used.

## Gender statistics

Gender statistics are calculated using destination column D.

A gender pie chart can be generated for each chapter from Chapter 1 to Chapter 10.

Categories with a count of zero are excluded from both the pie chart and its legend.

## Regional statistics

Regional statistics are calculated using destination column G.

The six regions are:

1. Africa
2. Asia
3. Europe
4. North America, Central America and the Caribbean
5. South America
6. South-West Pacific

Regions with a count of zero are excluded from both the pie chart and its legend.

## Pie charts

The dashboard contains chapter-level gender and regional pie charts.

Data labels use the format:

```text
5 (25%)
```

where:

- `5` is the number of authors; and
- `25%` is their percentage of the relevant chapter total.

The charts are configured with:

- data labels without coloured legend-key squares;
- leader lines for labels placed outside the pie;
- legends below the pie charts;
- horizontal legend layout where sufficient space is available; and
- 11-point legend and data-label text.

Excel ultimately determines whether a long legend needs to wrap depending on the available chart width.

## Helper worksheet

A hidden worksheet called:

```text
ChartData
```

stores the data used to construct the dashboard charts.

This worksheet is generated automatically and should not normally require manual editing.

## Expected workbook structure

The destination/template workbook is expected to contain:

```text
Data
Dashboard_externalCAs
```

The script creates or recreates:

```text
ChartData
```

`ChartData` is hidden in the final workbook.

## Important limitations

The script does **not**:

- connect directly to Microsoft Forms;
- authenticate to Microsoft 365;
- automatically download the source workbook;
- automatically upload the output to SharePoint;
- require Power Automate; or
- require Office Scripts.

The user remains responsible for downloading the current Forms response workbook and uploading the processed workbook to the appropriate SharePoint location.

## Data protection

The registration workbook may contain personal and confidential information.

**Do not commit response workbooks, processed workbooks, or other files containing registration data to a public GitHub repository.**

It is strongly recommended to include a `.gitignore` file such as:

```gitignore
# Excel data files
*.xlsx
*.xls
*.xlsm
*.xlsb

# Excel temporary files
~$*

# Python cache
__pycache__/
*.py[cod]

# Local environments
.venv/
venv/
env/

# OS files
.DS_Store
Thumbs.db
```

The GitHub repository should normally contain only the script, documentation, and other non-confidential supporting material.

## Repository structure

A simple repository structure is:

```text
CA-registration-sync/
│
├── Sync_CA_registration.py
├── README.md
├── .gitignore
└── requirements.txt
```

The `requirements.txt` file only needs:

```text
openpyxl
```

## License

Add an appropriate licence before distributing or reusing the tool outside its intended organisational context.
