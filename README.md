
## Excel Spreadsheet Processing with Python
---

The following Python script demonstrates how to manipulate an Excel spreadsheet using the `openpyxl` library. 
This script loads an Excel file named `transaction.xlsx`, reads and modifies data, and creates a bar chart to visualize the processed data. 

### Code Explanation

```python
import openpyxl as xl
from openpyxl.chart import BarChart, Reference

# Load the workbook and select the first worksheet
wb = xl.load_workbook('transaction.xlsx')
sheet = wb['Sheet1']

# Process each row in the worksheet, starting from the second row
for row in range(2, sheet.max_row + 1):
    # Retrieve the value from the third column (C) of the current row
    cell = sheet.cell(row, 3)
    # Apply a 10% discount to the value
    corrected_price = cell.value * 0.9
    # Write the corrected price to the fourth column (D) of the current row
    corrected_price_cell = sheet.cell(row, 4)
    corrected_price_cell.value = corrected_price

# Define the range for the bar chart data
values = Reference(sheet,
                   min_row=2,
                   max_row=sheet.max_row,
                   min_col=4,
                   max_col=4)

# Create a bar chart
chart = BarChart()
chart.add_data(values)

# Add the chart to the worksheet at position E2
sheet.add_chart(chart, 'E2')

# Save the modified workbook
wb.save('transaction.xlsx')
```

### Step-by-Step Breakdown

1. **Import Libraries**:
    ```python
    import openpyxl as xl
    from openpyxl.chart import BarChart, Reference
    ```
    - `openpyxl`: A library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files.
    - `BarChart`, `Reference`: Classes from `openpyxl.chart` to create and manage charts in Excel.

2. **Load the Workbook**:
    ```python
    wb = xl.load_workbook('transaction.xlsx')
    sheet = wb['Sheet1']
    ```
    - Load the workbook named `transaction.xlsx`.
    - Select the worksheet named `Sheet1`.

3. **Iterate Through Rows**:
    ```python
    for row in range(2, sheet.max_row + 1):
        cell = sheet.cell(row, 3)
        corrected_price = cell.value * 0.9
        corrected_price_cell = sheet.cell(row, 4)
        corrected_price_cell.value = corrected_price
    ```
    - Loop through each row starting from the second row to the last row.
    - Read the value in the third column (C).
    - Apply a 10% discount to the value.
    - Write the corrected price in the fourth column (D).

4. **Define the Data Range for the Chart**:
    ```python
    values = Reference(sheet,
                       min_row=2,
                       max_row=sheet.max_row,
                       min_col=4,
                       max_col=4)
    ```
    - Create a `Reference` object defining the range of the corrected prices (column D) to be used in the bar chart.

5. **Create and Add the Bar Chart**:
    ```python
    chart = BarChart()
    chart.add_data(values)
    sheet.add_chart(chart, 'E2')
    ```
    - Instantiate a `BarChart` object.
    - Add the data range to the chart.
    - Place the chart on the worksheet at cell E2.

6. **Save the Workbook**:
    ```python
    wb.save('transaction.xlsx')
    ```
    - Save the changes made to the workbook.
---

### Finaly 

This script is a simple yet powerful example of how Python can be used to automate data processing and visualization tasks in Excel spreadsheets. 
By leveraging the `openpyxl` library, you can efficiently handle large datasets and enhance your productivity with automated workflows.


#### Input 

<img width="215" alt="Screenshot 2024-05-25 at 01 10 53" src="https://github.com/DumindUdara/Excel_Spreadsheet/assets/98957798/d6ae5cdd-6bd1-4416-b02e-1540505291f9">

#### Output 

<img width="512" alt="Screenshot 2024-05-25 at 01 11 48" src="https://github.com/DumindUdara/Excel_Spreadsheet/assets/98957798/181381d5-cd90-4ed1-80c1-d356b6876b9a">

