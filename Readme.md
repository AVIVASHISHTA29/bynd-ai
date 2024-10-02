# README

## Overview

This Python script transforms a raw Excel file (`source.xlsx`) into a target Excel file (`transformed_final.xlsx`) by identifying cells that should contain formulas based on arithmetic relationships between numbers. It then inserts the appropriate formulas into those cells using the `openpyxl` library.

The script handles addition and subtraction operations involving multiple cells, processes all columns consistently, and correctly interprets numbers with commas and parentheses (e.g., `(28)` as `-28` and `13,507` as `13507`). Empty cells are treated as zeros in calculations.

## Features

- **Identifies formula cells**: Detects cells that should contain formulas based on bold formatting.
- **Handles arithmetic relationships**: Considers addition and subtraction operations, including combinations of multiple cells.
- **Parses formatted numbers**: Correctly parses numbers with commas and parentheses.
- **Processes empty cells**: Treats empty cells as zeros in calculations.
- **Consistent across columns**: Applies the same logic to all columns in the Excel sheet.
- **Preserves formatting**: Copies cell formats and styles from the source to the target workbook.

## Requirements

- Python 3.x

## Dependencies

All dependencies are listed in the `requirements.txt` file. Install them using:

```bash
pip install -r requirements.txt
```

**Contents of `requirements.txt`:**

```
openpyxl
```

## Instructions

### 1. Prepare the Environment

Ensure you have Python 3.x installed on your system. Check your Python version using:

```bash
python --version
```

### 2. Install Dependencies

Install the required Python packages using the `requirements.txt` file:

```bash
pip install -r requirements.txt
```

### 3. Prepare the Source File

- **Place `source.xlsx` in the Same Directory**: Ensure the `source.xlsx` file is in the same directory as the script.
- **Bold the Formula Cells**: In `source.xlsx`, apply bold formatting to the cells that should contain formulas.
- **Data Consistency**:
  - Ensure there is no unintended data or formatting beyond your data columns (e.g., columns beyond the last data column).
  - Remove any extra columns or rows that do not contain data.
- **Data Formatting**: You don't need to modify the formatting of numbers with commas or parentheses; the script handles them.

### 4. Run the Script

Execute the script using the command line:

```bash
python main.py
```

Replace `main.py` with the actual name of your script file.

### 5. Check the Output

After running the script, a new Excel file named `transformed_final.xlsx` will be created in the same directory.

- **Verify the Formulas**: Open `transformed_final.xlsx` and ensure that the formulas have been correctly applied to the bolded cells.
- **Check for Extra Columns**: Confirm that there are no extra columns (e.g., columns beyond your data columns).
- **Consistency Across Columns**: Ensure that formulas are applied consistently across all columns.

## Script Explanation

### Overview

The script performs the following steps:

1. **Load the Source Workbook**: Opens `source.xlsx` and selects the active sheet.
2. **Create a New Workbook**: Initializes a new workbook (`transformed_wb`) for the transformed data.
3. **Copy Data and Formats**: Copies data from the source sheet to the transformed sheet, preserving formats and styles.
4. **Parse and Store Cell Values**:
   - Parses numeric values, handling commas and parentheses.
   - Stores the numeric values in a dictionary `cell_values`.
   - Treats empty cells as zeros.
5. **Identify and Apply Formulas**:
   - Iterates over each bolded cell that should contain a formula.
   - Searches for arithmetic relationships with cells above the target cell in the same column.
   - Considers combinations of previous cells, trying all possible sequences of '+' and '-' operators.
   - Constructs and applies the appropriate formula if a match is found.
6. **Save the Transformed Workbook**: Writes the changes to `transformed_final.xlsx`.

### Key Functions

#### `parse_number(value)`

Parses a number from a string that may contain commas and parentheses.

```python
def parse_number(value):
    if isinstance(value, (int, float)):
        return float(value)
    elif isinstance(value, str):
        value = value.strip()
        # Remove commas
        value = value.replace(',', '')
        # Handle negative numbers in parentheses
        if value.startswith('(') and value.endswith(')'):
            value = '-' + value[1:-1]
        # Remove any additional formatting
        value = value.replace('$', '')  # Remove dollar signs if present
        try:
            return float(value)
        except ValueError:
            return None
    else:
        return None
```

#### `find_formula(ws, target_cell, data_start_row, cell_values, max_depth=10, tolerance=0.01)`

Finds an appropriate formula for a target cell by checking arithmetic relationships with other cells in the same column.

```python
def find_formula(ws, target_cell, data_start_row, cell_values, max_depth=10, tolerance=0.01):
    # Function implementation as in the script
```

- **Parameters**:
  - `ws`: Worksheet object.
  - `target_cell`: The cell for which to find the formula.
  - `data_start_row`: The row where your data starts.
  - `cell_values`: Dictionary of cell values.
  - `max_depth`: Maximum number of previous cells to consider (default is `10`).
  - `tolerance`: Acceptable difference between calculated and target values (default is `0.01`).

#### `evaluate_expression(expr, cell_values)`

Evaluates arithmetic expressions by replacing cell references with their values from `cell_values`.

```python
def evaluate_expression(expr, cell_values):
    # Function implementation as in the script
```

### Script Logic

- **Data Copying**: Copies each cell's value and formatting from the source to the transformed workbook.
- **Cell Value Parsing**: Converts all numeric values to floats, handling special formats.
- **Formula Detection**:
  - Only processes bolded cells, as these are the ones expected to contain formulas.
  - Uses combinations of previous cells to attempt to recreate the target value using addition and subtraction.
- **Formula Application**:
  - If a matching combination is found, the corresponding formula is applied to the cell.
  - Updates `cell_values` with the result for potential use in subsequent calculations.

### Parameters and Adjustments

- **`data_start_row`**: Set to `3` by default. Adjust if your data starts on a different row.
- **`data_start_col`**: Set to `2` by default (column B). Adjust if your data starts on a different column.
- **`max_depth`**: Increased to `10` to consider more combinations of cells.
- **`tolerance`**: Set to `0.01` to account for floating-point precision issues.

## Troubleshooting

### Extra Columns Appearing

- **Issue**: An extra column (e.g., column F) appears in the output file.
- **Solution**:
  - Ensure there is no unintended data or formatting in the source file beyond your data columns.
  - Adjust the `max_col` calculation in the script to use the source worksheet's maximum column.

### Incorrect Formulas

- **Issue**: Formulas are not being applied correctly.
- **Solution**:
  - Verify that the cells intended to contain formulas are bolded.
  - Ensure `data_start_row` and `data_start_col` correctly reflect where your data starts.
  - Increase `max_depth` if necessary.

### Performance Issues

- **Issue**: The script runs slowly with large datasets.
- **Solution**:
  - Consider reducing `max_depth` to limit the number of combinations evaluated.
  - Ensure your data does not contain unnecessary rows or columns.

## Summary

This script automates the process of identifying and applying formulas in an Excel sheet based on arithmetic relationships. By following the instructions, you can transform a raw Excel file into one with the correct formulas applied, saving time and reducing the potential for manual errors.
