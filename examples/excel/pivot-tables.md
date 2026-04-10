# Pivot Table Showcase

Comprehensive demo of all OfficeCLI pivot table features across 11 sheets.

**Script:** [pivot-tables.sh](pivot-tables.sh)
**Output:** [pivot-tables.xlsx](pivot-tables.xlsx)

```bash
bash pivot-tables.sh
```

## Source Data

- **Sheet1**: 50 rows, 10 columns (Region, Category, Product, Quarter, Sales, Quantity, Cost, Channel, Priority, Date) spanning 2024-2025
- **CNData**: 12 rows of Chinese sales data for locale sort demo

## Sheets

| # | Sheet | Layout | Rows | Cols | Values | Key Features |
|---|-------|--------|------|------|--------|-------------|
| 1 | Sales Overview | tabular | Region > Category | Quarter | Sales:sum, Qty:sum, Cost:**%row** | **repeatLabels**, dual filters, desc sort |
| 2 | Market Share | outline | Region | Category | Sales:**%col** | percent-of-column display |
| 3 | Product Deep Dive | tabular | Category > Product | — | Sales:sum/avg/max, Qty:sum, Cost:sum | **5 value fields**, no col axis |
| 4 | Channel Analysis | outline | Channel | Quarter | Sales:**%total**, Qty:sum | percent-of-total, no filters |
| 5 | Priority Matrix | tabular | Priority > Region | Category | Sales:sum, Cost:%row | **blankRows** between groups |
| 6 | Compact 3-Level | **compact** | Region > Category > Product | — | Sales:sum, Qty:sum | **3-level hierarchy** with indentation |
| 7 | No Subtotals | tabular | Region > Category | Quarter | Sales:sum | **subtotals=off**, grandTotals=cols |
| 8 | Date Grouping | outline | **Date:year > Date:quarter** | — | Sales:sum, Cost:sum | **automatic date grouping** |
| 9 | Top 5 Products | tabular | Product | — | Sales:sum, Qty:sum, Cost:sum | **topN=5**, grandTotals=none |
| 10 | Ultimate | tabular | Region > Category | Quarter | Sales:sum, Qty:avg, Cost:%row | repeatLabels + blankRows + dual filters |
| 11 | Chinese Locale | tabular | Region > Category | — | Sales:sum | **locale sort** (pinyin), grandTotalCaption |

## Features Covered

- **Layouts**: compact, outline, tabular
- **Report Layout**: repeatLabels (Repeat All Item Labels), blankRows (Insert Blank Line After Each Item)
- **Filters**: 0, 1, or 2 page filters
- **Row hierarchy**: 1, 2, or 3 levels
- **Column axis**: with or without
- **Value fields**: 1 to 5 simultaneous data fields
- **Aggregations**: sum, average, max
- **Show Data As**: percent_of_row, percent_of_col, percent_of_total
- **Grand totals**: both, rows, cols, none
- **Subtotals**: on, off
- **Sort**: asc, desc, locale (Chinese pinyin)
- **Date grouping**: year + quarter automatic hierarchy
- **Top-N**: filter to top 5 by value
- **Custom caption**: grandTotalCaption for localized labels
- **Styles**: 7 different PivotStyle themes (Light, Medium, Dark)
