# CopyFilteredWithFormat
Excel XLAM macro to copy filtered rows with formatting (merged-cell safe)

# CopyFilteredWithFormat (Excel XLAM)

A production-ready Excel VBA macro that copies **only filtered (visible) rows**
from a selected range and pastes them **with formatting**, including **merged
cells**, into another column.

## Features
- Works only on visible (filtered) rows
- Preserves formatting and formulas
- Merged-cell safe (no duplication or corruption)
- Native Excel UI using `Application.InputBox(Type:=8)`
- Performance optimized for large datasets
- XLAM-ready for global use

## How to Use
1. Apply a filter to your data
2. Select the source range
3. Run `CopyFiltered_WithFormat_Final`
4. Choose any cell in the destination column
5. Done

## Installation (XLAM)
1. Open Excel → ALT + F11
2. Import `CopyFiltered_WithFormat_Final.bas`
3. Save workbook as **Excel Add-In (*.xlam)**
4. Enable via Excel → Options → Add-Ins

## Author
Akshay Solanki

## License
MIT
