# Flux Analysis Template Specification Reference

## Overview
This document provides exact cell-level formatting specifications for recreating the February 2026 flux analysis workbook using openpyxl. Use this as the source of truth for all template formatting details.

**Template File:** 2026-02 Flux Analysis v2.xlsx
**Analysis Date:** 2026-04-06

---

## Global Formatting Standards

### Font Defaults
- **Font Name:** Arial (all cells)
- **Default Font Size:** 10pt
- **Title Font Size:** 12pt (company name rows)
- **Section Title Font Size:** 14pt

### Number Formats
- **Currency:** `"$"#,##0.00_);\("$"#,##0.00\)` (standard accounting format with parentheses for negatives)
- **Percentage:** `0.0%` (percentage columns, one decimal place)
- **Flag Column:** Yellow fill (`FFFFFF00`) when value = "Y"

### Color Codes
- **Header Fill (Gray):** `FFD0D0D0` (light gray)
- **Yellow Highlight:** `FFFFFF00` (for flagged/notable items)

---

## Income Statement Tab

### Sheet Dimensions
- **Rows:** 135
- **Columns:** 14 (A through N)

### Column Widths
```
A: 47.0       (Description/GL account)
B: 14.109375  (Dec 2025 Amount)
C: 13.44140625 (Jan 2026 Amount)
D: 13.0       (Feb 2026 Amount)
E: 12.44140625 ($ Variance)
F: 9.109375   (% Variance)
G: 2.0        (Spacer)
H: 8.109375   (Comment)
I: 13.0       (Unused)
J: 13.0       (Unused)
K: 15.109375  (Per Flux Meeting - Feb 2026)
L: 14.5546875 (Post Meeting Adjs - Feb 2026)
M: 12.0       (Changes)
N: 13.0       (Comments)
```

### Merged Cells
```
A1:D1  (Company name)
A2:D2  (Holdings subtitle)
A3:D3  (Statement title)
A4:D4  (Period range)
A5:D5  (Spacer)
A6:D6  (Options line)
```

### Freeze Panes
- **Position:** B9 (freeze columns A and rows 1-8)

### Header Rows (Rows 1-8)

#### Row 1: Company Name
- **Cell A1:** "Acme, Inc."
  - Font: Arial, 12pt, bold
  - Merged: A1:D1

#### Row 2: Holdings Subtitle
- **Cell A2:** "Acme Holdings (Consolidated)"
  - Font: Arial, 12pt, bold
  - Merged: A2:D2

#### Row 3: Statement Title
- **Cell A3:** "Acme Income Statement"
  - Font: Arial, 14pt, bold
  - Merged: A3:D3

#### Row 4: Period Label
- **Cell A4:** "Dec 2025, Jan 2026, Feb 2026"
  - Font: Arial, 14pt, bold
  - Merged: A4:D4

#### Row 5: Spacer Row
- **Cell A5:** "" (empty)
  - Font: Arial, 14pt, bold
  - Merged: A5:D5

#### Row 6: Options and Flux Meeting Headers
- **Cell A6:** "Options: Activity Only"
  - Font: Arial, 14pt, bold
  - Merged: A6:D6
- **Cell K6:** "Per Flux Meeting"
  - Font: Arial, 8pt, not bold
- **Cell L6:** "Post meeting adjs"
  - Font: Arial, 8pt, not bold

#### Row 7: Column Headers (Part 1)
All cells use Arial 7pt bold with gray fill (`FFD0D0D0`)
```
A7: "Financial Row"
B7: "Dec 2025"
C7: "Jan 2026"
D7: "Feb 2026"
K7: "Feb 2026"
L7: "Feb 2026"
```

#### Row 8: Column Headers (Part 2)
All cells use Arial 7pt bold with gray fill (`FFD0D0D0`)
```
A8: "\xa0" (non-breaking space)
B8: "Amount"
C8: "Amount"
D8: "Amount"
E8: "$ Variance"
F8: "% Variance" (number format: 0.0%)
H8: "Comment"
K8: "Amount"
L8: "Amount"
M8: "Changes"
N8: "Comments"
```

### Data Rows (9+)

#### Row 9 Onwards Structure
- **Column A:** GL account descriptions (e.g., "411000 - Revenue")
  - Font: Arial, 10pt, normal
- **Columns B, C, D:** Monthly amounts
  - Number Format: `"$"#,##0.00_);\("$"#,##0.00\)`
  - Font: Arial, 10pt
- **Column E:** Variance (Feb 2026 - Jan 2026, typically)
  - Number Format: `"$"#,##0.00_);\("$"#,##0.00\)`
  - Font: Arial, 10pt
- **Column F:** % Variance
  - Number Format: `0.0%`
  - Font: Arial, 10pt
- **Column H:** Comment field
  - Font: Arial, 10pt
- **Columns K, L:** Flux meeting amounts
  - Number Format: `"$"#,##0.00_);\("$"#,##0.00\)`
- **Column M:** Changes flag/indicator
  - Number Format: `"$"#,##0.00_);\("$"#,##0.00\)`
- **Column N:** Comments
  - Font: Arial, 10pt

#### Group Headers
GL account group rows (e.g., "Ordinary Income/Expense") are typically:
- **Bold** Arial 10pt
- Left-aligned
- Not indented

---

## Balance Sheet Tab

### Sheet Dimensions
- **Rows:** 240
- **Columns:** 12 (A through L)

### Column Widths
```
A: 52.44140625  (Description/GL account)
B: 15.0         (Amount - FY 2025)
C: 13.0         (Jan 2026 Amount)
D: 13.0         (Feb 2026 Amount)
E: 14.0         ($ Variance)
F: 10.6640625   (% Variance)
G: 2.0          (Spacer)
H: 33.0         (Comment/Notes)
I: 13.77734375  (Per Flux Meeting)
J: 13.0         (Post Meeting Adjs)
K: 11.77734375  (Changes)
L: 13.0         (Comments)
```

### Merged Cells
```
A1:D1  (Company name)
A2:D2  (Holdings subtitle)
A3:D3  (Statement title)
A4:D4  (Period label)
A5:D5  (Spacer)
A6:D6  (Options line)
C7:D7  (Column header span)
B7:B9  (Vertical span for FY label)
```

### Freeze Panes
- **Position:** B10 (freeze columns A and rows 1-9)

### Header Rows (Rows 1-10)

#### Rows 1-6: Title Section
Same structure as Income Statement:
- Row 1: Company name (Arial 12pt bold)
- Row 2: Holdings subtitle (Arial 12pt bold)
- Row 3: "*ACS Balance Sheet - Fiscal Year by Period" (Arial 14pt bold)
- Row 4: "End of Feb 2026" (Arial 14pt bold)
- Row 5: Empty spacer (Arial 14pt bold)
- Row 6: "Options: Activity Only" (Arial 14pt bold)

#### Row 7: Main Column Headers
All cells use Arial 7pt bold with gray fill (`FFD0D0D0`)
```
A7: "Financial Row"
B7: "Amount (As of FY 2025)"
C7: "Amount"
I7: "Per Flux Meeting"
J7: "Post meeting adjs"
Note: C7:D7 is merged
Note: B7:B9 is merged
```

#### Row 8: Subheaders
All cells use Arial 7pt bold with gray fill (`FFD0D0D0`)
```
A8: "\xa0" (non-breaking space)
C8: "As of Jan 2026"
D8: "As of Feb 2026"
I8: "Feb 2026"
J8: "Feb 2026"
```

#### Row 9: Variance Headers
All cells use Arial 7pt bold with gray fill (`FFD0D0D0`)
```
A9: "\xa0"
C9: "\xa0"
D9: "\xa0"
E9: "$ Variance"
F9: "% Variance"
I9: "Amount"
J9: "Amount"
K9: "Changes"
L9: "Comments"
```

#### Row 10: Section Header
```
A10: "ASSETS"
Font: Arial, 8pt, bold
```

### Data Rows (11+)
- **Column A:** GL account descriptions
  - Font: Arial, 10pt
  - Section headers (ASSETS, LIABILITIES, EQUITY, etc.) are bold
- **Columns B, C, D:** Amounts
  - Number Format: `"$"#,##0.00_);\("$"#,##0.00\)`
- **Column E:** Variance
  - Number Format: `"$"#,##0.00_);\("$"#,##0.00\)`
- **Column F:** % Variance
  - Number Format: `0.0%`
- **Columns I, J:** Flux meeting amounts
  - Number Format: `"$"#,##0.00_);\("$"#,##0.00\)`
- **Columns K, L:** Changes and comments
  - Font: Arial, 10pt

---

## COGS Tab

### Sheet Dimensions
- **Rows:** 28
- **Columns:** 6 (A through F)

### Column Widths
```
A: 40.0         (Vendor/GL account)
B: 18.0         (Month name label)
C: 11.77734375  (Dec 2025)
D: 13.0         (Jan 2026)
E: 12.44140625  (Feb 2026)
F: 9.77734375   (Variance)
```

### Merged Cells
- None

### Freeze Panes
- None

### Header Structure

#### Row 3: Pivot Table Labels
```
A3: "Sum of Amount"
   Font: Arial, 10pt
B3: "Column Labels"
   Font: Arial, 10pt
```

#### Row 4: Column Headers
All cells use Arial 7pt bold (no special fill)
```
A4: "Row Labels"
B4: "Dec 2025"
C4: "Jan 2026"
D4: "Feb 2026"
E4: "Variance"
F4: "Note"
```

### GL Account Header Rows
When GL accounts are grouped (e.g., "511400 - COGS - Hosting"):
- **Font:** Arial, 10pt, **bold**
- Row contains the GL account code and description in column A
- Columns C, D, E may contain subtotals
- Column F may contain notes

### Data Rows (5+)
Vendor detail rows under each GL account:
- **Column A:** Vendor name
  - Font: Arial, 10pt, normal
- **Columns C, D, E:** Monthly amounts
  - Number Format: `"$"#,##0.00_);\("$"#,##0.00\)` or `_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)`
  - Font: Arial, 10pt
- **Column F:** Notes or variance
  - Font: Arial, 10pt

### Variance Calculation
The rightmost numeric column (F) typically contains:
- Absolute variance between Feb 2026 and Jan 2026
- Or variance notation/flag
- Number Format: `"$"#,##0.00_);\("$"#,##0.00\)`

---

## Contractors Tab

### Sheet Dimensions
- **Rows:** 18
- **Columns:** 7 (A through G)

### Column Widths
```
A: 70.5546875   (Contractor name - wider than COGS)
B: 16.6640625   (Month name label)
C: 11.44140625  (Dec 2025)
D: 13.0         (Jan 2026)
E: 13.0         (Feb 2026)
F: 13.0         (Variance)
G: 10.0         (Note/Flag)
```

### Merged Cells
- None

### Freeze Panes
- None

### Header Structure

#### Row 3: Pivot Table Labels
```
A3: "Sum of Net"
   Font: Arial, 10pt, normal
B3: "Column Labels"
   Font: Arial, 10pt, normal
```

#### Row 4: Column Headers
All cells use Arial 7pt bold
```
A4: "Row Labels"
B4: "Dec 2025"
C4: "Jan 2026"
D4: "Feb 2026"
E4: "Variance"
F4: "Note"
(Note: F4 may be blank or contain label)
```

### Data Rows (5+)
Contractor detail rows:
- **Column A:** Contractor name (wider column allows full names)
  - Font: Arial, 10pt, normal
- **Columns C, D, E:** Monthly amounts
  - Number Format: `_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)` or `"$"#,##0.00_);\("$"#,##0.00\)`
  - Font: Arial, 10pt
- **Column F:** Variance
  - Font: Arial, 10pt
- **Column G:** Flags or notes
  - May have yellow fill (`FFFFFF00`) if value = "Y"

---

## Professional Fees Tab

### Sheet Dimensions
- **Rows:** 32
- **Columns:** 6 (A through F)

### Column Widths
```
A: 42.44140625  (Vendor/Firm name)
B: 16.6640625   (Month name label)
C: 10.44140625  (Dec 2025)
D: 13.0         (Jan 2026)
E: 11.44140625  (Feb 2026)
F: 13.0         (Note/Variance)
```

### Merged Cells
- None

### Freeze Panes
- None

### Header Structure

#### Row 3: Pivot Table Labels
```
A3: "Sum of Net"
   Font: Arial, 10pt, normal
B3: "Column Labels"
   Font: Arial, 10pt, normal
```

#### Row 4: Column Headers
All cells use Arial 7pt bold
```
A4: "Row Labels"
B4: "Dec 2025"
C4: "Jan 2026"
D4: "Feb 2026"
E4: "Variance"
F4: "Note"
```

### Data Rows (5+)
Vendor/firm detail rows (flat list, no GL grouping):
- **Column A:** Vendor/firm name
  - Font: Arial, 10pt, normal
- **Columns C, D, E:** Monthly amounts
  - Number Format: `"$"#,##0.00_);\("$"#,##0.00\)` or `_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)`
  - Font: Arial, 10pt
- **Column F:** Notes or variance
  - Font: Arial, 10pt

**Key Difference from COGS/Contractors:**
- No GL account grouping (flat vendor list)
- Single indentation level

---

## Software Tab

### Sheet Dimensions
- **Rows:** 136
- **Columns:** 14 (A through N)

### Column Widths
```
A: 36.44140625  (Full list - Software name)
B: 16.6640625   (Month label)
C: 11.44140625  (Dec 2025)
D: 13.0         (Jan 2026)
E: 12.77734375  (Feb 2026)
F: 37.77734375  (Top movers - Software name)
G: 10.44140625  (Dec 2025)
H: 13.0         (Jan 2026)
I: 13.109375    (Feb 2026)
J: 11.109375    (Var)
K: 12.6640625   (Abs Var)
L: 18.109375    (Note)
M: 8.77734375   (Unused/Flag)
N: 13.0         (Unused)
```

### Merged Cells
- None

### Freeze Panes
- None

### Layout Structure
**Two-section layout on same sheet:**
- **Left section (A-E):** Full software list (all items)
- **Right section (F-L):** Top movers by absolute variance (manually curated)

### Header Rows

#### Row 2: Prepaids Label
```
L2: "Check for Prepaids"
   Font: Arial, 8pt, normal
```

#### Row 3: Pivot Table Labels
```
A3: "Sum of Net"
   Font: Arial, 10pt, normal
B3: "Column Labels"
   Font: Arial, 10pt, normal
F3: (blank or repeat label)
```

#### Row 4: Column Headers

**Left Section (Full List):**
All cells use Arial 10pt (varies - some bold for right section)
```
A4: "Row Labels"
   Font: Arial, 10pt
B4: "Dec 2025"
   Font: Arial, 10pt
C4: "Jan 2026"
   Font: Arial, 10pt
D4: "Feb 2026"
   Font: Arial, 10pt
```

**Right Section (Top Movers):**
All use Arial 8pt bold with yellow fill (`FFFFFF00`)
```
F4: "Row Labels"
   Font: Arial, 8pt, bold
G4: "Dec 2025"
   Font: Arial, 8pt, bold
H4: "Jan 2026"
   Font: Arial, 8pt, bold
I4: "Feb 2026"
   Font: Arial, 8pt, bold
J4: "Var"
   Font: Arial, 8pt, bold
K4: "Abs Var"
   Font: Arial, 8pt, bold
L4: "Note"
   Font: Arial, 8pt, bold
```

### Data Rows

**Left Section (5+):**
- **Column A:** Software name
  - Font: Arial, 10pt
- **Columns C, D, E:** Monthly amounts
  - Number Format: `"$"#,##0.00_);\("$"#,##0.00\)`
  - Font: Arial, 10pt

**Right Section (5+):**
- **Column F:** Software name (top movers)
  - Font: Arial, 10pt
- **Columns G, H, I:** Monthly amounts
  - Number Format: `_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)`
  - Fill: Yellow (`FFFFFF00`)
  - Font: Arial, 10pt
- **Column J:** Variance (Jan to Feb)
  - Number Format: `_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)`
  - Fill: Yellow (`FFFFFF00`)
- **Column K:** Absolute Variance
  - Number Format: `_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)`
  - Fill: Yellow (`FFFFFF00`)
- **Column L:** Note/description
  - Font: Arial, 10pt
  - Fill: Yellow (`FFFFFF00`)
  - Example: "$30k Brex purchase - Infrastructure"

### Right Section Curation
- Manually selected items (not auto-generated)
- Sorted by Absolute Variance (descending)
- Updated monthly to show material changes
- Provides quick view of top cost drivers and changes

---

## Implementation Guide for openpyxl

### Cell Formatting Template

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# Header style (gray background)
header_fill = PatternFill(start_color="FFD0D0D0", end_color="FFD0D0D0", fill_type="solid")
header_font = Font(name='Arial', size=7, bold=True)

# Currency format
currency_format = '"$"#,##0.00_);\\("$"#,##0.00\\)'

# Percentage format
percent_format = '0.0%'

# Yellow highlight (for flags)
yellow_fill = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")

# Example: Setting up Income Statement header
ws = wb['Income Statement']
ws.column_dimensions['A'].width = 47.0
ws.column_dimensions['B'].width = 14.109375
# ... set all column widths

# Merge cells
ws.merge_cells('A1:D1')
ws['A1'] = 'Acme, Inc.'
ws['A1'].font = Font(name='Arial', size=12, bold=True)

# Set freeze panes
ws.freeze_panes = 'B9'

# Header row formatting
for col in ['A', 'B', 'C', 'D', 'E', 'F', 'H', 'K', 'L']:
    ws[f'{col}8'].fill = header_fill
    ws[f'{col}8'].font = header_font
    if col in ['E', 'F']:
        ws[f'{col}8'].number_format = percent_format if col == 'F' else currency_format

# Data rows (starting row 9)
for row in range(9, 135):
    for col in ['B', 'C', 'D', 'E', 'K', 'L', 'M']:
        ws[f'{col}{row}'].number_format = currency_format
    ws[f'F{row}'].number_format = percent_format
```

### Quick Reference: Column Width Assignments

**Use this approach to set all widths per sheet:**
```python
column_widths = {
    'A': 47.0,
    'B': 14.109375,
    'C': 13.44140625,
    # ... etc
}
for col, width in column_widths.items():
    ws.column_dimensions[col].width = width
```

### Quick Reference: Number Formats

| Format Purpose | openpyxl Format String | Example Value | Display |
|---|---|---|---|
| Currency (accounting) | `"$"#,##0.00_);\("$"#,##0.00\)` | 1234.56 | $X,XXX.XX |
| Percentage | `0.0%` | 0.085 | 8.5% |
| Thousands with decimals | `_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)` | 1234.56 | 1,234.56 |

---

## Notes and Special Cases

### Flag Column (Yellow Highlighting)
When a cell in column N (Income Statement) or G (Contractors) contains "Y":
- Apply yellow fill: `PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")`
- Typically used to mark accounts with notes or unusual activity

### Flux Meeting Columns (Income Statement K-N)
These columns are used for post-meeting adjustments:
- **K:** Feb 2026 amount per flux meeting decision
- **L:** Post-meeting adjustments (if needed)
- **M:** Change indicator (difference between K and actual)
- **N:** Meeting notes/comments

### GL Account Grouping
- **Income Statement:** Groups by GL account code (e.g., "411000 - Revenue")
  - Group rows are bold
  - Sub-accounts are indented or listed below
- **Balance Sheet:** Groups by asset/liability/equity categories
  - Section headers (ASSETS, LIABILITIES, etc.) are bold
  - Sub-accounts listed below with indentation
- **COGS/Contractors:** May have vendor grouping by department
- **Professional Fees:** Flat list (no grouping)

### Software Tab Right Section
The right section is **manually curated each month**:
1. Not auto-calculated from the left section
2. Represents the top variance items by absolute change
3. Updated by the analyst to highlight items of concern
4. Sorted by impact, not alphabetically

### Spacer Rows
Rows labeled as "spacer" (empty cells with formatting) are used for visual separation:
- May contain non-breaking space (`\xa0`) in column A
- Have the same font formatting as headers but empty content
- Help readability by breaking up dense data

---

## Related Files and Context

- **Source Workbook:** 2026-02 Flux Analysis v2.xlsx
- **Analysis Date:** April 6, 2026
- **Data Period:** February 2026 (with comparisons to Dec 2025 and Jan 2026)
- **Consolidation:** Acme Holdings (Consolidated)
