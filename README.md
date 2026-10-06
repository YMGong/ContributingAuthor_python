# CA Registration Sync Tool

`Sync_CA_registration.py` is a Python tool for processing Contributing Author (CA) registration data collected through Microsoft Forms and transferring it into a standardised Excel workbook.

The tool is designed for a workflow in which the Microsoft Forms response workbook and the destination workbook may be stored in different Microsoft 365 / SharePoint environments. It therefore uses a simple **download → process → upload** workflow rather than connecting directly to SharePoint.

## Features

The script:

- imports selected fields from a Microsoft Forms response workbook;
- transfers the data to the `Data` worksheet of an existing Excel template;
- standardises selected fields and family names;
- identifies multiple submissions from the same person;
- reconciles information across duplicate submissions;
- preserves template formatting where applicable;
- applies **Arial Narrow** throughout the generated workbook;
- creates separate dashboards for external CAs and all CAs;
- calculates chapter-level gender and regional statistics;
- avoids double-counting repeated submissions;
- generates gender and regional pie charts for Chapters 1–10; and
- saves the processed data as a new Excel workbook.

The original source and template workbooks are not modified.

## Requirements

- Python 3.9 or later
- `openpyxl`

Install the required package with:

```bash
pip install openpyxl
```

The script also uses `tkinter` for file-selection dialogs. `tkinter` is included with most standard Python installations.

## Running the tool

Run:

```bash
python Sync_CA_registration.py
```

The script will ask you to select:

1. the downloaded Microsoft Forms response workbook;
2. the destination/template workbook; and
3. the location where the processed workbook should be saved.

The suggested output filename is:

```text
02. WGI CA list.xlsx
```

## Workflow

```text
Microsoft Forms
      │
      ▼
Download response workbook
[IPCC WGI AR7 Contributing Author Registration Form.xlsx](https://upsud-my.sharepoint.com/:x:/r/personal/yongmei_gong_universite-paris-saclay_fr/Documents/IPCC%20WGI%20AR7%20Contributing%20Author%20Registration%20Form.xlsx?d=w0e7da531d8184f7eb4e8cb7dbfcdc232&csf=1&web=1&e=OYAcE4)
      │
      ▼
Download [02. WGI CA list_template.xlsx](templates/02.%20WGI%20CA%20list_template.xlsx) from this repository 
      │
      ▼
Run Sync_CA_registration.py
      │
      ├── Select Forms response workbook
      ├── Select destination/template workbook
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

## Source-to-destination mapping

The following fields are copied from the Forms response workbook to the destination `Data` worksheet:

| Source column | Destination column |
|:---:|:---:|
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

Source data are read from row 2 onwards.

Row 1 of the destination `Data` worksheet is preserved as the header row.

## Data standardisation

### Gender

Values from source column J are converted as follows:

| Source | Destination |
|---|---|
| `Male` | `M` |
| `Female` | `F` |

### Yes/No field

Values from source column K are converted as follows:

| Source | Destination |
|---|---:|
| `Yes` | `1` |
| `No` | `0` |

### Chapter

Chapter information from source column W is converted from text into a numeric chapter value.

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

- removes leading and trailing whitespace;
- collapses repeated spaces;
- converts Latin names to uppercase; and
- removes Latin accents/diacritics.

Examples:

```text
García  → GARCIA
Müller  → MULLER
Dvořák  → DVORAK
```

Non-Latin scripts are preserved rather than transliterated into Latin characters.

## Identifying the same person

A person is identified using the combination of:

```text
Family name + Given name
```

The comparison is case-insensitive and ignores unnecessary whitespace.

Family-name standardisation is performed before this comparison.

## Handling multiple submissions

If the same person submitted the form multiple times, selected information is reconciled across **all of their submissions**.

### Destination column I

If source column K contains `No` in **any** submission, destination column I is set to:

```text
0
```

for all submissions from that person.

### Destination column J

If source column L is empty in **any** submission, destination column J is empty for all submissions from that person.

### Destination column K

If source column M is empty in **any** submission, destination column K is empty for all submissions from that person.

Reconciliation is performed before dashboard statistics are calculated.

## Formatting

The output workbook uses:

```text
Arial Narrow
```

as its standard font.

The font is applied to worksheet cells and chart text while preserving other formatting properties where applicable, including:

- font size;
- bold/italic formatting; and
- font colour.

### Red family names

Family names in destination column B are displayed in red when the person's final reconciled value in destination column I is:

```text
1
```

Otherwise, the normal template formatting is retained.

# Dashboards

The script generates two dashboards:

```text
Dashboard_externalCAs
Dashboard_allCAs
```

Both dashboards use the same chapter-level counting and chart-generation logic but include different populations.

## Dashboard_externalCAs

`Dashboard_externalCAs` includes only authors whose final reconciled value in destination column K is empty.

The dashboard displays:

```text
Total external CAs:
```

This is the number of **unique external CAs**, not the number of form submissions.

Multiple submissions from the same person therefore do not inflate the total.

## Dashboard_allCAs

`Dashboard_allCAs` includes **all CAs**, regardless of the value in destination column K.

The dashboard displays:

```text
Total CAs:
```

This is the number of **unique CAs** in the dataset.

Again, multiple submissions from the same person are counted only once in this total.

## Chapter-level counting

For chapter-level statistics, the counting unit is:

```text
Person + Chapter
```

This prevents repeated submissions from causing an author to be counted more than once within the same chapter.

For example, if one author submits the form three times for Chapter 4:

```text
Author A + Chapter 4
Author A + Chapter 4
Author A + Chapter 4
```

the author is counted **once** for Chapter 4.

However, if the same author is associated with two chapters:

```text
Author A + Chapter 4
Author A + Chapter 7
```

the author is counted once in Chapter 4 and once in Chapter 7.

## Conflicting duplicate submissions

If multiple submissions from the same person for the same chapter contain conflicting gender or regional information, the **latest submission** is used.

Submission order is determined using destination column A.

## Gender statistics

Gender statistics are calculated from destination column D.

Statistics are generated separately for Chapters 1–10.

Categories with a count of zero are omitted from both the pie chart and its legend.

## Regional statistics

Regional statistics are calculated from destination column G.

The following six regions are used:

1. Africa
2. Asia
3. Europe
4. North America, Central America and the Caribbean
5. South America
6. South-West Pacific

Regions with a count of zero are omitted from both the pie chart and its legend.

# Pie charts

Both dashboards contain chapter-level gender and regional pie charts.

Data labels use the format:

```text
5 (25%)
```

where:

- `5` is the number of authors; and
- `25%` is the percentage of the relevant chapter population.

## Chart formatting

The charts are configured with:

- width: `20`;
- height: `7`;
- Arial Narrow text;
- 11-point legend and data-label text;
- legends positioned below the pie chart;
- horizontal legend layout where sufficient width is available;
- no coloured legend-key squares beside the `5 (25%)` data labels;
- labels positioned outside the pie when necessary; and
- leader lines for labels positioned outside the pie.

The coloured markers remain in the legend so that categories can be matched to their corresponding pie slices.

Excel ultimately determines whether a particularly long legend needs to wrap depending on the available chart width.

# Helper worksheets

The dashboard calculations use hidden helper worksheets.

The external-CA dashboard uses:

```text
ChartData
```

The all-CA dashboard uses:

```text
ChartData_allCAs
```

These worksheets contain the values used to construct the pie charts and are hidden in the final workbook.

They should not normally require manual editing.

## Expected workbook structure

The destination/template workbook should contain:

```text
Data
Dashboard_externalCAs
```

The script creates or recreates the dashboard/helper worksheets required for the output, including:

```text
Dashboard_externalCAs
Dashboard_allCAs
ChartData
ChartData_allCAs
```

# Important limitations

The script does **not**:

- connect directly to Microsoft Forms;
- authenticate to Microsoft 365;
- automatically download the Forms response workbook;
- automatically upload the resulting workbook to SharePoint;
- require Power Automate; or
- require Office Scripts.

The user is responsible for downloading the current Forms response workbook and uploading/replacing the processed workbook in the appropriate SharePoint location.

# Data protection

The registration workbook may contain personal and confidential information.

**Do not commit Forms response workbooks, processed CA workbooks, or other files containing registration data to a public GitHub repository.**

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

# Virtual environments
.venv/
venv/
env/

# Operating-system files
.DS_Store
Thumbs.db
```

Only the Python source code, documentation, and other non-confidential supporting files should normally be committed.

# Suggested repository structure

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

# License

Add an appropriate licence before distributing or reusing the tool outside its intended organisational context.
