# Excel Sheet Reference Examples

This document provides additional examples and use cases for working with cross-sheet references in Excel files.

## Example 1: Simple Data Summary

This matches the exact example from the issue description.

### Scenario
You have raw fruit data in one sheet and want to count occurrences in a summary sheet.

### Implementation

```python
from openpyxl import Workbook

wb = Workbook()

# sheet1 - Source data
sheet1 = wb.active
sheet1.title = "sheet1"
sheet1['A1'] = 'fruit'
sheet1['A2'] = 'apple'
sheet1['A3'] = 'apple'
sheet1['A4'] = 'banana'

# sheet2 - Summary with formulas
sheet2 = wb.create_sheet("sheet2")
sheet2['A1'] = 'fruit'
sheet2['B1'] = 'count'
sheet2['A2'] = 'apple'
sheet2['B2'] = '=COUNTIFS(sheet1!A:A,A2)'
sheet2['A3'] = 'banana'
sheet2['B3'] = '=COUNTIFS(sheet1!A:A,A3)'

wb.save('fruit_count.xlsx')
```

### Result in Excel

**sheet1**:
```
| A      |
|--------|
| fruit  |
| apple  |
| apple  |
| banana |
```

**sheet2**:
```
| A      | B     |
|--------|-------|
| fruit  | count |
| apple  | 2     |
| banana | 1     |
```

## Example 2: Sales Analysis with VLOOKUP

### Scenario
You have a product catalog and want to automatically populate order details using VLOOKUP.

### Implementation

```python
from openpyxl import Workbook

wb = Workbook()

# Products sheet
products = wb.active
products.title = "Products"
products.append(['Product', 'Price', 'Supplier'])
products.append(['Widget A', 19.99, 'Acme Corp'])
products.append(['Widget B', 29.99, 'Beta Inc'])
products.append(['Widget C', 39.99, 'Gamma LLC'])

# Orders sheet
orders = wb.create_sheet("Orders")
orders['A1'] = 'Product'
orders['B1'] = 'Quantity'
orders['C1'] = 'Unit Price'
orders['D1'] = 'Subtotal'
orders['E1'] = 'Supplier'

# Add order with VLOOKUP formulas
orders['A2'] = 'Widget A'
orders['B2'] = 5
orders['C2'] = '=VLOOKUP(A2,Products!A:C,2,FALSE)'
orders['D2'] = '=B2*C2'
orders['E2'] = '=VLOOKUP(A2,Products!A:C,3,FALSE)'

orders['A3'] = 'Widget C'
orders['B3'] = 3
orders['C3'] = '=VLOOKUP(A3,Products!A:C,2,FALSE)'
orders['D3'] = '=B3*C3'
orders['E3'] = '=VLOOKUP(A3,Products!A:C,3,FALSE)'

wb.save('sales_orders.xlsx')
```

## Example 3: Dynamic Dashboard with Multiple Criteria

### Scenario
Create a dashboard that filters and summarizes sales data based on multiple criteria.

### Implementation

```python
from openpyxl import Workbook

wb = Workbook()

# Sales data
sales = wb.active
sales.title = "SalesData"
sales.append(['Date', 'Product', 'Region', 'Amount'])
sales.append(['2024-01-01', 'Widget A', 'North', 1500])
sales.append(['2024-01-02', 'Widget B', 'South', 2000])
sales.append(['2024-01-03', 'Widget A', 'North', 1800])
sales.append(['2024-01-04', 'Widget C', 'East', 2500])
sales.append(['2024-01-05', 'Widget A', 'South', 1600])

# Dashboard
dashboard = wb.create_sheet("Dashboard")
dashboard['A1'] = 'Metric'
dashboard['B1'] = 'Value'

# Count total sales
dashboard['A2'] = 'Total Sales Count'
dashboard['B2'] = '=COUNTA(SalesData!A:A)-1'

# Sum total amount
dashboard['A3'] = 'Total Revenue'
dashboard['B3'] = '=SUM(SalesData!D:D)'

# Count Widget A sales
dashboard['A4'] = 'Widget A Sales Count'
dashboard['B4'] = '=COUNTIFS(SalesData!B:B,"Widget A")'

# Sum Widget A revenue
dashboard['A5'] = 'Widget A Revenue'
dashboard['B5'] = '=SUMIF(SalesData!B:B,"Widget A",SalesData!D:D)'

# Count North region sales
dashboard['A6'] = 'North Region Sales'
dashboard['B6'] = '=COUNTIFS(SalesData!C:C,"North")'

# Widget A in North region
dashboard['A7'] = 'Widget A in North'
dashboard['B7'] = '=COUNTIFS(SalesData!B:B,"Widget A",SalesData!C:C,"North")'

# Average sale amount
dashboard['A8'] = 'Average Sale'
dashboard['B8'] = '=AVERAGE(SalesData!D:D)'

wb.save('sales_dashboard.xlsx')
```

## Example 4: Inventory Management with INDEX-MATCH

### Scenario
Track inventory and use INDEX-MATCH for more flexible lookups than VLOOKUP.

### Implementation

```python
from openpyxl import Workbook

wb = Workbook()

# Inventory
inventory = wb.active
inventory.title = "Inventory"
inventory.append(['SKU', 'Product', 'Quantity', 'Location'])
inventory.append(['SKU001', 'Widget A', 150, 'Warehouse 1'])
inventory.append(['SKU002', 'Widget B', 200, 'Warehouse 2'])
inventory.append(['SKU003', 'Widget C', 100, 'Warehouse 1'])
inventory.append(['SKU004', 'Widget D', 175, 'Warehouse 3'])

# Stock Check
check = wb.create_sheet("StockCheck")
check['A1'] = 'Product Name'
check['B1'] = 'SKU'
check['C1'] = 'Current Stock'
check['D1'] = 'Location'

# Search by product name
check['A2'] = 'Widget B'
# INDEX-MATCH to find SKU (column 1) by matching product name (column 2)
check['B2'] = '=INDEX(Inventory!A:A,MATCH(A2,Inventory!B:B,0))'
# INDEX-MATCH to find quantity
check['C2'] = '=INDEX(Inventory!C:C,MATCH(A2,Inventory!B:B,0))'
# INDEX-MATCH to find location
check['D2'] = '=INDEX(Inventory!D:D,MATCH(A2,Inventory!B:B,0))'

check['A3'] = 'Widget D'
check['B3'] = '=INDEX(Inventory!A:A,MATCH(A3,Inventory!B:B,0))'
check['C3'] = '=INDEX(Inventory!C:C,MATCH(A3,Inventory!B:B,0))'
check['D3'] = '=INDEX(Inventory!D:D,MATCH(A3,Inventory!B:B,0))'

wb.save('inventory_lookup.xlsx')
```

## Example 5: Grade Book with Conditional Formatting Data

### Scenario
Calculate student statistics from a grade sheet.

### Implementation

```python
from openpyxl import Workbook

wb = Workbook()

# Grades sheet
grades = wb.active
grades.title = "Grades"
grades.append(['Student', 'Math', 'Science', 'English'])
grades.append(['Alice', 85, 90, 88])
grades.append(['Bob', 78, 82, 80])
grades.append(['Charlie', 92, 88, 95])
grades.append(['Diana', 88, 85, 87])

# Statistics sheet
stats = wb.create_sheet("Statistics")
stats['A1'] = 'Subject'
stats['B1'] = 'Average'
stats['C1'] = 'Highest'
stats['D1'] = 'Lowest'
stats['E1'] = 'Count >85'

stats['A2'] = 'Math'
stats['B2'] = '=AVERAGE(Grades!B:B)'
stats['C2'] = '=MAX(Grades!B:B)'
stats['D2'] = '=MIN(Grades!B:B)'
stats['E2'] = '=COUNTIF(Grades!B:B,">85")'

stats['A3'] = 'Science'
stats['B3'] = '=AVERAGE(Grades!C:C)'
stats['C3'] = '=MAX(Grades!C:C)'
stats['D3'] = '=MIN(Grades!C:C)'
stats['E3'] = '=COUNTIF(Grades!C:C,">85")'

stats['A4'] = 'English'
stats['B4'] = '=AVERAGE(Grades!D:D)'
stats['C4'] = '=MAX(Grades!D:D)'
stats['D4'] = '=MIN(Grades!D:D)'
stats['E4'] = '=COUNTIF(Grades!D:D,">85")'

wb.save('gradebook.xlsx')
```

## Example 6: Budget Tracking with Multiple Sheets

### Scenario
Track expenses across categories and create a summary budget view.

### Implementation

```python
from openpyxl import Workbook

wb = Workbook()

# Expenses
expenses = wb.active
expenses.title = "Expenses"
expenses.append(['Date', 'Category', 'Description', 'Amount'])
expenses.append(['2024-01-05', 'Food', 'Groceries', 150.00])
expenses.append(['2024-01-06', 'Transport', 'Gas', 45.00])
expenses.append(['2024-01-07', 'Food', 'Restaurant', 65.00])
expenses.append(['2024-01-08', 'Utilities', 'Electric Bill', 120.00])
expenses.append(['2024-01-09', 'Food', 'Groceries', 180.00])
expenses.append(['2024-01-10', 'Transport', 'Parking', 20.00])

# Budget Summary
budget = wb.create_sheet("BudgetSummary")
budget['A1'] = 'Category'
budget['B1'] = 'Total Spent'
budget['C1'] = 'Number of Transactions'
budget['D1'] = 'Average per Transaction'

categories = ['Food', 'Transport', 'Utilities']

for idx, category in enumerate(categories, start=2):
    budget[f'A{idx}'] = category
    # SUMIF to sum expenses by category
    budget[f'B{idx}'] = f'=SUMIF(Expenses!B:B,A{idx},Expenses!D:D)'
    # COUNTIF to count transactions
    budget[f'C{idx}'] = f'=COUNTIF(Expenses!B:B,A{idx})'
    # Average calculation
    budget[f'D{idx}'] = f'=B{idx}/C{idx}'

# Add totals
total_row = len(categories) + 2
budget[f'A{total_row}'] = 'TOTAL'
budget[f'B{total_row}'] = '=SUM(Expenses!D:D)'
budget[f'C{total_row}'] = f'=SUM(C2:C{total_row-1})'

wb.save('budget_tracker.xlsx')
```

## Example 7: Cross-Reference with Error Handling

### Scenario
Create formulas that gracefully handle missing data using IFERROR.

### Implementation

```python
from openpyxl import Workbook

wb = Workbook()

# Master list
master = wb.active
master.title = "MasterList"
master.append(['ID', 'Name', 'Department'])
master.append(['E001', 'Alice', 'Engineering'])
master.append(['E002', 'Bob', 'Marketing'])
master.append(['E003', 'Charlie', 'Sales'])

# Lookup with error handling
lookup = wb.create_sheet("Lookup")
lookup['A1'] = 'Employee ID'
lookup['B1'] = 'Name'
lookup['C1'] = 'Department'
lookup['D1'] = 'Status'

# Valid ID
lookup['A2'] = 'E001'
lookup['B2'] = '=IFERROR(VLOOKUP(A2,MasterList!A:C,2,FALSE),"Not Found")'
lookup['C2'] = '=IFERROR(VLOOKUP(A2,MasterList!A:C,3,FALSE),"Not Found")'
lookup['D2'] = '=IF(ISNA(VLOOKUP(A2,MasterList!A:C,1,FALSE)),"Invalid","Valid")'

# Invalid ID
lookup['A3'] = 'E999'
lookup['B3'] = '=IFERROR(VLOOKUP(A3,MasterList!A:C,2,FALSE),"Not Found")'
lookup['C3'] = '=IFERROR(VLOOKUP(A3,MasterList!A:C,3,FALSE),"Not Found")'
lookup['D3'] = '=IF(ISNA(VLOOKUP(A3,MasterList!A:C,1,FALSE)),"Invalid","Valid")'

wb.save('error_handled_lookup.xlsx')
```

## Tips for Complex Formulas

### 1. Building Complex COUNTIFS
Count items matching multiple criteria:
```python
# Count apples in North region with quantity > 10
sheet['A1'] = '=COUNTIFS(Data!A:A,"apple",Data!B:B,"North",Data!C:C,">10")'
```

### 2. Nested VLOOKUP
Look up values from different sheets:
```python
# Get price from Products, then find discount in DiscountTable
sheet['A1'] = '=VLOOKUP(VLOOKUP(A2,Products!A:B,2,FALSE),DiscountTable!A:C,2,FALSE)'
```

### 3. Using SUMIFS with Multiple Conditions
```python
# Sum sales for specific product and region
sheet['A1'] = '=SUMIFS(Sales!D:D,Sales!B:B,"Widget A",Sales!C:C,"North")'
```

### 4. Combining INDEX-MATCH for Two-Way Lookup
```python
# Find value at intersection of row and column
sheet['A1'] = '=INDEX(Data!A:Z,MATCH(lookup_row,Data!A:A,0),MATCH(lookup_col,Data!1:1,0))'
```

## Common Patterns

### Pattern 1: Data Validation List
Use data from another sheet for dropdown lists:
```python
# In openpyxl, you can set data validation to reference another sheet
from openpyxl.worksheet.datavalidation import DataValidation

dv = DataValidation(type="list", formula1="Products!$A$2:$A$100")
sheet.add_data_validation(dv)
dv.add(sheet['A2'])
```

### Pattern 2: Dynamic Named Ranges
Reference entire columns that may grow:
```python
# Use full column reference
sheet['A1'] = '=SUM(Data!A:A)'
```

### Pattern 3: Conditional Aggregation
Sum or count based on conditions:
```python
# Sum all sales > 1000 for specific product
sheet['A1'] = '=SUMIFS(Sales!C:C,Sales!A:A,"Widget A",Sales!C:C,">1000")'
```

## Best Practices Recap

1. **Always use exact sheet names** - Case-sensitive and must match exactly
2. **Quote sheet names with spaces** - Use 'My Sheet'!A1 format
3. **Use FALSE for exact matches** - In VLOOKUP and similar functions
4. **Wrap in IFERROR** - Handle missing data gracefully
5. **Use full column references** - More flexible as data grows (A:A instead of A1:A100)
6. **Test in Excel first** - Verify formula syntax before coding
7. **Document your formulas** - Add comments explaining complex formulas

## Running the Examples

To run any of these examples:

1. Save the code to a Python file
2. Install openpyxl: `pip install openpyxl`
3. Run: `python your_file.py`
4. Open the generated .xlsx file in Excel to see the results

For a complete working script with all examples, use:
```bash
python scripts/create_excel_with_references.py
```
