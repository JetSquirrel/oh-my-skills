---
name: excel-sheet-reference
description: Create Excel files with multiple sheets and cross-sheet formulas including COUNTIFS, VLOOKUP, and MATCH. Keep source data in one sheet and create summary sheets with formulas referencing the source data.
---

# Excel Sheet Reference Skill

This skill helps you create Excel (.xlsx) files with multiple sheets and use cross-sheet references with formulas. This is useful when you need to maintain source data in one sheet and create summary or analysis sheets that reference and calculate results from the source data.

## Use Cases

- Summarizing datasets while keeping raw data intact
- Creating report sheets that dynamically calculate from source data
- Building dashboards with formulas referencing multiple data sheets
- Maintaining data integrity by using formulas instead of copying values

## Prerequisites

You'll need Python with the `openpyxl` library to create Excel files with formulas:

```bash
pip install openpyxl
```

## Creating Multi-Sheet Excel Files with Formulas

### Basic Structure

When creating Excel files with cross-sheet references:

1. **Source Data Sheet**: Contains raw data
2. **Summary/Report Sheet**: Contains formulas that reference the source sheet

### Cross-Sheet Reference Syntax

To reference cells in another sheet:
- Basic syntax: `SheetName!CellRange`
- If sheet name has spaces: `'Sheet Name'!CellRange`
- Range examples: `Sheet1!A1`, `Sheet1!A1:A10`, `'Data Sheet'!B2:B100`

## Common Formulas with Cross-Sheet References

### 1. COUNTIFS - Count cells that meet multiple criteria

**Syntax**: `COUNTIFS(range1, criteria1, [range2, criteria2], ...)`

**Example**: Count how many times "apple" appears in Sheet1
```
=COUNTIFS(Sheet1!A:A,"apple")
```

**Multiple criteria example**:
```
=COUNTIFS(Sheet1!A:A,"apple",Sheet1!B:B,">100")
```

### 2. VLOOKUP - Lookup and retrieve values from a table

**Syntax**: `VLOOKUP(lookup_value, table_array, column_index_num, [range_lookup])`

**Parameters**:
- `lookup_value`: The value to search for (in the first column)
- `table_array`: The range of cells containing the data table
- `column_index_num`: Which column to return (1-based index)
- `range_lookup`: FALSE for exact match, TRUE for approximate (default)

**Example**: Look up price for a product
```
=VLOOKUP(A2,Sheet1!A:C,3,FALSE)
```

### 3. MATCH - Find the position of a value in a range

**Syntax**: `MATCH(lookup_value, lookup_array, [match_type])`

**Parameters**:
- `lookup_value`: The value to find
- `lookup_array`: The range to search in
- `match_type`: 0 = exact match, 1 = largest value ≤ lookup_value, -1 = smallest value ≥ lookup_value

**Example**: Find position of "banana" in a list
```
=MATCH("banana",Sheet1!A:A,0)
```

### 4. SUMIF / SUMIFS - Sum values based on criteria

**Example**: Sum values from Sheet1 where category is "apple"
```
=SUMIF(Sheet1!A:A,"apple",Sheet1!B:B)
```

## Python Implementation Example

### Creating Excel with Cross-Sheet References

```python
from openpyxl import Workbook

# Create a new workbook
wb = Workbook()

# Create source data sheet
source_sheet = wb.active
source_sheet.title = "SourceData"

# Add headers and data
source_sheet['A1'] = 'Fruit'
source_sheet['A2'] = 'apple'
source_sheet['A3'] = 'apple'
source_sheet['A4'] = 'banana'
source_sheet['A5'] = 'orange'
source_sheet['A6'] = 'apple'
source_sheet['A7'] = 'banana'

# Create summary sheet
summary_sheet = wb.create_sheet("Summary")

# Add headers
summary_sheet['A1'] = 'Fruit'
summary_sheet['B1'] = 'Count'

# Add fruit names
summary_sheet['A2'] = 'apple'
summary_sheet['A3'] = 'banana'
summary_sheet['A4'] = 'orange'

# Add formulas that reference SourceData sheet
# COUNTIFS formula to count each fruit
summary_sheet['B2'] = '=COUNTIFS(SourceData!A:A,A2)'
summary_sheet['B3'] = '=COUNTIFS(SourceData!A:A,A3)'
summary_sheet['B4'] = '=COUNTIFS(SourceData!A:A,A4)'

# Save the workbook
wb.save('fruits_summary.xlsx')
```

### Advanced Example with VLOOKUP

```python
from openpyxl import Workbook

wb = Workbook()

# Create price list sheet
price_sheet = wb.active
price_sheet.title = "PriceList"
price_sheet['A1'] = 'Product'
price_sheet['B1'] = 'Price'
price_sheet['C1'] = 'Category'
price_sheet['A2'] = 'apple'
price_sheet['B2'] = 1.50
price_sheet['C2'] = 'fruit'
price_sheet['A3'] = 'banana'
price_sheet['B3'] = 0.80
price_sheet['C3'] = 'fruit'
price_sheet['A4'] = 'carrot'
price_sheet['B4'] = 1.20
price_sheet['C4'] = 'vegetable'

# Create order sheet with lookups
order_sheet = wb.create_sheet("Orders")
order_sheet['A1'] = 'Product'
order_sheet['B1'] = 'Quantity'
order_sheet['C1'] = 'Price'
order_sheet['D1'] = 'Total'
order_sheet['E1'] = 'Category'

order_sheet['A2'] = 'apple'
order_sheet['B2'] = 5
# VLOOKUP to get price from PriceList
order_sheet['C2'] = '=VLOOKUP(A2,PriceList!A:C,2,FALSE)'
# Calculate total
order_sheet['D2'] = '=B2*C2'
# VLOOKUP to get category
order_sheet['E2'] = '=VLOOKUP(A2,PriceList!A:C,3,FALSE)'

order_sheet['A3'] = 'banana'
order_sheet['B3'] = 10
order_sheet['C3'] = '=VLOOKUP(A3,PriceList!A:C,2,FALSE)'
order_sheet['D3'] = '=B3*C3'
order_sheet['E3'] = '=VLOOKUP(A3,PriceList!A:C,3,FALSE)'

wb.save('orders_with_lookup.xlsx')
```

### Example with MATCH and INDEX

```python
from openpyxl import Workbook

wb = Workbook()

# Data sheet
data_sheet = wb.active
data_sheet.title = "Data"
data_sheet['A1'] = 'Name'
data_sheet['B1'] = 'Age'
data_sheet['C1'] = 'City'
data_sheet['A2'] = 'Alice'
data_sheet['B2'] = 25
data_sheet['C2'] = 'New York'
data_sheet['A3'] = 'Bob'
data_sheet['B3'] = 30
data_sheet['C3'] = 'London'
data_sheet['A4'] = 'Charlie'
data_sheet['B4'] = 35
data_sheet['C4'] = 'Paris'

# Lookup sheet
lookup_sheet = wb.create_sheet("Lookup")
lookup_sheet['A1'] = 'Search Name'
lookup_sheet['B1'] = 'Position'
lookup_sheet['C1'] = 'Age'

lookup_sheet['A2'] = 'Bob'
# MATCH to find position
lookup_sheet['B2'] = '=MATCH(A2,Data!A:A,0)'
# INDEX with MATCH to get age
lookup_sheet['C2'] = '=INDEX(Data!B:B,MATCH(A2,Data!A:A,0))'

wb.save('data_with_match.xlsx')
```

## Important Notes

### Formula Storage vs. Calculation

When using `openpyxl`:
- Formulas are stored as strings in the Excel file
- Excel will calculate the results when the file is opened
- To see calculated values in Python, you need to load the file after Excel has calculated it

### Sheet Name Rules

- Avoid special characters in sheet names
- Use single quotes around sheet names with spaces: `'My Sheet'!A1`
- Keep sheet names under 31 characters

### Cell Reference Formats

- **Relative**: `A1` - Changes when formula is copied
- **Absolute**: `$A$1` - Stays fixed when formula is copied
- **Mixed**: `$A1` or `A$1` - Partially fixed

### Common Formula Patterns

**Count unique values from another sheet**:
```python
# Using COUNTIFS with specific criteria
sheet['A1'] = '=COUNTIFS(SourceData!A:A,A2)'
```

**Sum with multiple conditions**:
```python
sheet['A1'] = '=SUMIFS(Data!C:C,Data!A:A,"apple",Data!B:B,">10")'
```

**Conditional lookup**:
```python
# VLOOKUP with IFERROR for handling missing values
sheet['A1'] = '=IFERROR(VLOOKUP(A2,Data!A:C,3,FALSE),"Not Found")'
```

## Complete Working Example

Here's a complete example that demonstrates the use case from the issue:

```python
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

def create_fruit_summary_excel():
    """
    Create an Excel file with source data and summary sheet.
    Summary sheet uses COUNTIFS to count occurrences of each fruit.
    """
    wb = Workbook()
    
    # Sheet 1: Source Data
    sheet1 = wb.active
    sheet1.title = "sheet1"
    
    # Add header
    sheet1['A1'] = 'fruit'
    
    # Add data
    fruits_data = ['apple', 'apple', 'banana', 'orange', 'apple', 
                   'banana', 'grape', 'apple', 'orange', 'banana']
    
    for idx, fruit in enumerate(fruits_data, start=2):
        sheet1[f'A{idx}'] = fruit
    
    # Sheet 2: Summary with formulas
    sheet2 = wb.create_sheet("sheet2")
    
    # Add headers
    sheet2['A1'] = 'fruit'
    sheet2['B1'] = 'count'
    
    # Get unique fruits for summary
    unique_fruits = ['apple', 'banana', 'orange', 'grape']
    
    # Add formulas to count each fruit
    for idx, fruit in enumerate(unique_fruits, start=2):
        sheet2[f'A{idx}'] = fruit
        # COUNTIFS formula referencing sheet1
        sheet2[f'B{idx}'] = f'=COUNTIFS(sheet1!A:A,A{idx})'
    
    # Save workbook
    wb.save('fruit_summary.xlsx')
    print("Created fruit_summary.xlsx with cross-sheet references")
    
    return wb

if __name__ == "__main__":
    create_fruit_summary_excel()
```

## Tips and Best Practices

1. **Test formulas in Excel first**: Create a sample Excel file manually to verify formula syntax
2. **Use named ranges**: For complex references, consider using named ranges for clarity
3. **Error handling**: Wrap formulas in IFERROR to handle missing data gracefully
4. **Data validation**: Ensure source data is clean before creating formulas
5. **Performance**: For large datasets, consider using pivot tables or Power Query instead of complex formulas
6. **Documentation**: Add comments or a README sheet explaining the workbook structure

## Troubleshooting

**Formula shows as text instead of calculating**:
- Check that you're assigning the formula as a string starting with `=`
- Verify sheet names are correct and properly quoted if they contain spaces

**#REF! error in Excel**:
- Sheet name might be incorrect
- Referenced range might not exist
- Check for typos in sheet names

**#NAME? error**:
- Formula function name might be misspelled
- Sheet reference syntax might be incorrect

**Formula not updating**:
- Excel calculates formulas on load
- Force recalculation: Ctrl+Alt+F9 (Windows) or Cmd+Shift+K (Mac)

## Additional Resources

- [openpyxl documentation](https://openpyxl.readthedocs.io/)
- [Excel formula reference](https://support.microsoft.com/en-us/office/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188)
- [Excel cross-sheet references](https://support.microsoft.com/en-us/office/create-an-external-reference-link-to-a-cell-range-in-another-workbook-c98d1803-dd75-4668-ac6a-d7cca2a9b95f)
