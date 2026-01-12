# CSV Data Summary - Examples

This document provides practical examples of using the CSV Data Summary skill for different types of data analysis.

## Example 1: Sales Data Analysis

### Input Data (sales_data.csv)

```csv
order_date,product,quantity,price,customer_segment,region
2024-01-05,Laptop,1,1200.00,Business,North
2024-01-05,Mouse,2,25.00,Consumer,South
2024-01-06,Keyboard,1,75.00,Consumer,East
...
```

### Usage

```python
from scripts.analyze import summarize_csv

# Analyze sales data
summary = summarize_csv('sales_data.csv')
print(summary)
```

### What You Get

- **Data Overview**: Number of orders, products, date range
- **Revenue Analysis**: Total revenue, average order value, revenue by product
- **Time Series**: Daily/weekly revenue trends
- **Customer Segments**: Distribution across Business vs Consumer
- **Regional Patterns**: Sales distribution by region
- **Visualizations**:
  - Time series plot showing revenue over time
  - Product distribution bar chart
  - Regional sales breakdown
  - Price vs quantity correlation heatmap

## Example 2: Customer Demographics

### Input Data (customers.csv)

```csv
customer_id,age,gender,city,signup_date,total_purchases
C001,25,Female,New York,2024-01-01,5
C002,34,Male,Los Angeles,2024-01-02,3
C003,45,Female,Chicago,2024-01-03,8
...
```

### Usage

```bash
python scripts/analyze.py customers.csv
```

### Analysis Includes

- Age distribution histogram
- Gender breakdown
- City frequency analysis
- Signup trends over time
- Purchase behavior patterns
- Correlation between age and purchase frequency

## Example 3: Financial Transactions

### Input Data (transactions.csv)

```csv
transaction_date,amount,category,payment_method,status
2024-01-01,150.50,Groceries,Credit Card,Completed
2024-01-01,45.00,Transport,Debit Card,Completed
2024-01-02,1200.00,Rent,Bank Transfer,Completed
...
```

### Key Insights Generated

- Transaction volume trends
- Average transaction amount by category
- Payment method preferences
- Status distribution (completed vs pending)
- Category spending patterns
- Monthly/weekly spending trends

## Example 4: Survey Responses

### Input Data (survey.csv)

```csv
respondent_id,age_group,satisfaction,product_rating,would_recommend
R001,25-34,Very Satisfied,5,Yes
R002,35-44,Satisfied,4,Yes
R003,18-24,Neutral,3,Maybe
...
```

### Analysis Output

- Satisfaction score distribution
- Product rating breakdown
- Recommendation likelihood by age group
- Correlation between satisfaction and rating
- Response completeness check

## Example 5: Operational Metrics

### Input Data (operations.csv)

```csv
timestamp,machine_id,temperature,pressure,output_units,downtime_minutes
2024-01-01 08:00:00,M001,85.5,120,1000,0
2024-01-01 09:00:00,M001,87.2,122,1050,0
2024-01-01 10:00:00,M002,90.1,118,980,15
...
```

### Metrics Analyzed

- Average temperature and pressure by machine
- Output efficiency patterns
- Downtime analysis
- Hourly/daily trends
- Machine performance comparison
- Correlation between temperature and output

## Command Line Examples

### Basic Analysis

```bash
# Analyze a CSV file
python scripts/analyze.py data/my_data.csv

# Use the included sample
python scripts/analyze.py resources/sample.csv
```

### Python Script Integration

```python
from scripts.analyze import summarize_csv
import sys

# Analyze multiple files
files = ['sales_q1.csv', 'sales_q2.csv', 'sales_q3.csv']
for file in files:
    print(f"\n\n=== Analysis of {file} ===")
    print(summarize_csv(file))
```

### Batch Processing

```bash
# Process all CSV files in a directory
for file in data/*.csv; do
    echo "Processing $file"
    python scripts/analyze.py "$file"
done
```

## Understanding the Visualizations

### Correlation Heatmap

Shows how numeric variables relate to each other:
- **Dark red**: Strong positive correlation
- **Dark blue**: Strong negative correlation
- **White**: No correlation

**Use Case**: Identify which metrics move together (e.g., temperature and output)

### Time Series Plot

Displays trends over time:
- **Upward trend**: Growth or increase
- **Downward trend**: Decline or decrease
- **Flat line**: Stable metric
- **Spikes**: Anomalies or seasonal patterns

**Use Case**: Spot trends, seasonality, and anomalies

### Distribution Histograms

Shows the spread and shape of numeric data:
- **Bell curve**: Normal distribution
- **Skewed left/right**: Data concentrated on one side
- **Multiple peaks**: Multiple groups in the data

**Use Case**: Understand data range and identify outliers

### Categorical Bar Charts

Displays frequency of categories:
- **Tall bars**: Common values
- **Short bars**: Rare values
- **Even heights**: Uniform distribution

**Use Case**: Identify most/least common categories

## Tips for Best Results

### 1. Prepare Your Data

```python
# Good column names
order_date, customer_name, total_amount

# Avoid
OrderDate, Customer Name, Total_Amount_$
```

### 2. Date Column Naming

Ensure date columns contain 'date' or 'time':
```csv
order_date,created_time,shipment_date  ✓
date1,date2,date3                      ✓
col1,col2,col3                         ✗ (won't be detected)
```

### 3. Numeric Data

Ensure numbers are not formatted as text:
```csv
1200.00    ✓
$1,200.00  ✗ (remove currency symbols)
1200       ✓
```

### 4. Missing Values

The tool automatically handles missing values:
```csv
order_date,product,quantity
2024-01-01,Laptop,5
2024-01-02,,3        ← Will be reported as missing
2024-01-03,Mouse,   ← Will be reported as missing
```

## Advanced Examples

### Custom Analysis Workflow

```python
from scripts.analyze import summarize_csv
import pandas as pd

# Load and pre-process
df = pd.read_csv('raw_data.csv')
df = df.dropna()  # Remove rows with missing data
df.to_csv('cleaned_data.csv', index=False)

# Analyze cleaned data
summary = summarize_csv('cleaned_data.csv')
print(summary)
```

### Combining Multiple Datasets

```python
import pandas as pd
from scripts.analyze import summarize_csv

# Merge multiple files
df1 = pd.read_csv('january.csv')
df2 = pd.read_csv('february.csv')
df3 = pd.read_csv('march.csv')

combined = pd.concat([df1, df2, df3], ignore_index=True)
combined.to_csv('q1_combined.csv', index=False)

# Analyze combined dataset
summary = summarize_csv('q1_combined.csv')
print(summary)
```

### Filtering Before Analysis

```python
import pandas as pd
from scripts.analyze import summarize_csv

# Load and filter
df = pd.read_csv('all_sales.csv')
high_value = df[df['amount'] > 1000]
high_value.to_csv('high_value_sales.csv', index=False)

# Analyze subset
summary = summarize_csv('high_value_sales.csv')
print(summary)
```

## Comparison with Excel Sheet Reference Skill

| Feature | CSV Data Summary | Excel Sheet Reference |
|---------|------------------|----------------------|
| **Primary Use** | Data analysis & insights | Excel file creation with formulas |
| **Input Format** | CSV files | N/A (creates new files) |
| **Output** | Statistics, charts (PNG) | Excel files with formulas |
| **Analysis Type** | Exploratory data analysis | Formula-based calculations |
| **Visualizations** | Automatic charts | None (Excel creates them) |
| **Best For** | Understanding data patterns | Maintaining live formulas in Excel |

### When to Use Which Skill

**Use CSV Data Summary when you want to:**
- Quickly understand a dataset
- Generate visualizations automatically
- Get statistical insights
- Identify data quality issues

**Use Excel Sheet Reference when you want to:**
- Create Excel files with multiple sheets
- Use VLOOKUP, COUNTIFS, and other formulas
- Maintain dynamic calculations in Excel
- Create reference-based reports

## Troubleshooting Common Issues

### Issue: "No visualizations created"

**Cause**: Dataset might have only one column or no numeric data

**Solution**: Ensure your CSV has multiple columns with numeric data

### Issue: "Date columns not detected"

**Cause**: Column names don't contain 'date' or 'time'

**Solution**: Rename columns to include these keywords:
```python
df.rename(columns={'created_at': 'created_date'}, inplace=True)
```

### Issue: "Correlation heatmap shows only 1x1"

**Cause**: Only one numeric column in the dataset

**Solution**: This is normal - correlations require at least 2 numeric columns

### Issue: "Plot is too crowded"

**Cause**: Too many categories or data points

**Solution**: The script automatically limits to top 10 categories. For custom limits, modify `analyze.py`

## Additional Resources

- [pandas User Guide](https://pandas.pydata.org/docs/user_guide/index.html)
- [Data Visualization Best Practices](https://www.storytellingwithdata.com/)
- [CSV Format Specification](https://tools.ietf.org/html/rfc4180)
- [Statistical Analysis Basics](https://www.khanacademy.org/math/statistics-probability)
