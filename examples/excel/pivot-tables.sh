#!/bin/bash
# Pivot Table Showcase — demonstrates all pivot table features in OfficeCLI
# Generates: pivot-tables.xlsx (11 pivot table sheets + 2 data sheets)
#
# Uses the batch JSON file (pivot-tables-batch.json) which contains all
# commands in a single array for one open/save cycle.
#
# ============================================================================
# Pivot tables created (see pivot-tables.md for full details):
#
#  1-Sales Overview     — tabular + repeatLabels + dual filters + 3 values
#                         with percent_of_row + desc sort
#  2-Market Share       — outline + percent_of_col display
#  3-Product Deep Dive  — tabular + 5 value fields (sum/avg/max) + no col axis
#  4-Channel Analysis   — outline + percent_of_total + no filters
#  5-Priority Matrix    — tabular + blankRows between groups
#  6-Compact 3-Level    — compact + 3-level row hierarchy with indentation
#  7-No Subtotals       — tabular + subtotals=off + grandTotals=cols only
#  8-Date Grouping      — outline + automatic year>quarter date hierarchy
#  9-Top 5 Products     — tabular + topN=5 + grandTotals=none
# 10-Ultimate           — tabular + repeatLabels + blankRows + dual filters
#                         + 3 mixed values + row-only grand totals
# 11-Chinese Locale     — tabular + sort=locale (pinyin) + grandTotalCaption
#
# ============================================================================
# Key pivot table properties demonstrated:
#
#   source           — data range including headers (e.g. Sheet1!A1:J51)
#   rows             — comma-separated row fields; multi-level creates hierarchy
#   cols             — column axis fields
#   values           — Field:func[:showDataAs] (e.g. Sales:sum:percent_of_row)
#   filters          — page filter fields (dropdown above pivot)
#   layout           — compact (indented), outline (grouped), tabular (flat)
#   repeatLabels     — repeat outer row labels on every data row
#   blankRows        — insert blank line after each outer group
#   grandTotals      — both | rows | cols | none
#   subtotals        — on | off
#   sort             — asc | desc | locale (pinyin) | locale-desc
#   topN             — keep only top N items by first value field
#   grandTotalCaption— custom label for grand total row
#   style            — PivotStyleLight/Medium/Dark + number
#   name             — pivot table name (unique within workbook)
#
# Aggregation functions: sum, count, average, max, min, product, stddev, var
# Show Data As: percent_of_row, percent_of_col, percent_of_total, running_total
# Date grouping: Field:year, Field:quarter, Field:month, Field:day
# ============================================================================

set -e

FILE="pivot-tables.xlsx"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

rm -f "$FILE"

# Create a blank workbook
officecli create "$FILE"

# Run all commands in a single batch (one open/save cycle)
officecli batch "$FILE" --force --input "$SCRIPT_DIR/pivot-tables-batch.json"

echo ""
echo "Done! Generated: $FILE"
echo "  13 sheets (Sheet1 + CNData + 11 pivot tables)"
