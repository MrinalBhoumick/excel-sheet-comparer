Here's a structured README file for your script that guides users on how to use it. You can include this in your project's repository or documentation:

---

# Compare Sheets and Highlight Differences

## Overview

This Python script compares multiple sheets within an Excel workbook and highlights differences between them. The differences are highlighted in orange and are saved in a new sheet called "Differences."

## Prerequisites

- **Python 3.x**: Ensure you have Python installed. This script was developed and tested with Python 3.x.
- **openpyxl Library**: The script uses the `openpyxl` library to handle Excel files. You can install it using pip.

## Installation

1. **Clone the Repository** (if applicable):

    ```bash
    git clone https://github.com/MrinalBhoumick/excel-sheet-comparer.git
    cd excel-sheet-comparer
    ```

2. **Install Dependencies**:

    Ensure you have `openpyxl` installed. You can install it using pip:

    ```bash
    pip install openpyxl
    ```

## Usage

1. **Prepare Your Excel File**:

    Ensure that your Excel workbook file is properly formatted and contains the sheets you want to compare. The script assumes that the first row of each sheet contains column headers.

2. **Run the Script**:

    Save the script as `compare_sheets.py` and run it with the file path of your Excel workbook as an argument:

    ```bash
    python compare_sheets.py <path_to_your_excel_file>
    ```

    Replace `<path_to_your_excel_file>` with the path to your Excel workbook.

3. **Check Results**:

    The script will add a new sheet named "Differences" to the workbook, highlighting the cells where differences were found. The highlighted differences will be shown in orange. The file will be saved with these changes.

## Script Details

### Functions

- **`compare_sheets(file_path)`**: Compares sheets in the specified Excel workbook. Highlights differences in a new sheet named "Differences."

### Key Features

- Compares all sheets in the workbook.
- Highlights differences in orange.
- Copies formatting (excluding fill color) from the original cells.
- Adds a new sheet for differences.

### Example

For an example Excel file `StepFunctions-Sheet.xlsx`, you can run:

```bash
python compare_sheets.py StepFunctions-Sheet.xlsx
```

The differences will be highlighted in orange and saved in a new sheet named "Differences."

## Troubleshooting

- **No differences highlighted**: Ensure that there are actual differences between the sheets. The script compares cells in all sheets.
- **Script errors**: Check for any syntax errors or issues with the Excel file formatting.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contact

For any questions or issues, please contact [Mrinal Bhoumick](mrinalbhoumick0610@example.com).

---

Feel free to adjust the contact information and other details as needed.