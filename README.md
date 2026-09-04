# Model Info Converter

A Streamlit app that converts legacy Levi's Model Info data into the new standardized format.

## Input File Requirements

Upload an Excel (.xlsx) file containing the following columns:

| Column |
|----------|
| PC-9 |
| PC9 |
| Model_Info |

## Example Conversion

Input:

```text
Model is 5'9", Waist 25", Wearing size 26 x 27
```

Output:

```text
Model is 175cm (5'9"), Waist 25", Wearing a Size 26 x 27
```

## Features

- Converts feet and inches to centimeters
- Preserves original height in parentheses
- Standardizes apparel sizes:
  - Small → S
  - Medium → M
  - Large → L
- Removes HTML line break tags (`<br>`)
- Removes inch symbols from size values
- Generates a downloadable output Excel file

## Output

The app creates a new column:

```text
Converted_Model_Info
```

and generates a downloadable:

```text
output.xlsx
```

## Built With

- Streamlit
- OpenPyXL
- Python
``
