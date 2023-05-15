# Excel File Converter

The Excel File Converter repository contains a Python script that allows users to convert old Excel file formats (`.xls`) into newer ones (`.xlsx`), while preserving the data and structure of the original file. This tool is useful for users who have legacy Excel files and wish to upgrade them to a newer format that is more universally compatible with modern systems.

## Features

- **Excel File Conversion**: This script reads each Excel file in the input directory and converts it to the newer `.xlsx` format.
- **Error Handling**: Invalid Excel files are skipped and logged for review.
- **Logging**: Each operation (successful or failed) is logged with a timestamp for future reference and debugging purposes.

## Warning

Please be aware that this tool only converts the **first sheet** of the original Excel file. If your `.xls` file contains multiple sheets, only the first one will be included in the converted `.xlsx` file. Make sure to back up your data or use a separate tool if you need to convert all sheets.

## How it Works

The script begins by asking for an input and output directory from the user. It then creates a log file in the output directory.

The script then iterates over every file in the input directory. If the file is an Excel file, the script tries to open it. If it's an older `.xls` file, it uses `xlrd` to open the file. If it's a newer `.xlsx` file, it uses `openpyxl` to open it. If an error occurs while opening the file (due to the file not being a valid Excel file), the script logs the error and skips to the next file.

For each valid Excel file, the script creates a new `.xlsx` file, and copies the data from the first sheet of the original file to the new file. The new file is then saved in the output directory.

Finally, the script logs the completion of the operation and continues to the next file. If there are no more files, the script informs the user that the conversion is complete.

## Requirements

- Python 3.7+
- `xlrd` package
- `openpyxl` package
- `logging` package
- `datetime` package

## Usage

Run the Python script and input the read path (input directory) and write path (output directory) when prompted. The script will then begin the conversion process, logging its progress along the way.

After completion, you'll find the converted `.xlsx` files in the output directory, along with a log file that details the operations performed by the script.
## Installation

Before you can run the script, you need to install the required Python packages. If you don't have them installed, you can do so using pip, Python's package installer. Open your terminal and type the following commands:

```
pip install xlrd
pip install openpyxl
```

## Running the Script

1. Clone this repository to your local machine using `git clone <repository_url>`.
2. Navigate to the cloned repository's directory.
3. Run the Python script using the command `python <script_name>.py`.
4. When prompted, input the path to the directory containing the Excel files you wish to convert (the "read path").
5. Next, input the path to the directory where you want the converted files to be saved (the "write path").
6. The script will now begin converting the files. Progress will be logged in a file named `Process.log`, located in the write path directory.

## Preserving All Sheets

If you need to preserve all the sheets in the original Excel file during conversion, follow the steps below to modify the script.

1. Locate the following code blocks in the script:

   For `.xls` files:

   ```python
   # create new Excel file with .xlsx extension
   new_filename = filename.split('.')[0] + '.xlsx'
   new_wb = Workbook()
   # copy first sheet from original file to new file
   new_sheet = new_wb.active
   new_sheet.name = sheet.name
   for row_idx in range(1, sheet.nrows):
       row = []
       for col_idx in range(sheet.ncols):
           row.append(sheet.cell_value(row_idx, col_idx))
       new_sheet.append(row)
   ```

   For `.xlsx` files:

   ```python
   # create new Excel file with .xlsx extension
   new_filename = filename.split('.')[0] + '.xlsx'
   new_wb = Workbook()
   # copy first sheet from original file to new file
   new_sheet = new_wb.active
   new_sheet.title = sheet.title
   for row in sheet.iter_rows(values_only=True):
       new_sheet.append(row)
   ```

2. Replace the above code blocks with the following:

   For `.xls` files:

   ```python
   # create new Excel file with .xlsx extension
   new_filename = filename.split('.')[0] + '.xlsx'
   new_wb = Workbook()
   new_wb.remove(new_wb.active)  # remove default sheet

   # copy all sheets from original file to new file
   for sheet_idx in range(wb.nsheets):
       sheet = wb.sheet_by_index(sheet_idx)
       new_sheet = new_wb.create_sheet(sheet.name)
       for row_idx in range(sheet.nrows):
           row = []
           for col_idx in range(sheet.ncols):
               row.append(sheet.cell_value(row_idx, col_idx))
           new_sheet.append(row)
   ```

   For `.xlsx` files:

   ```python
   # create new Excel file with .xlsx extension
   new_filename = filename.split('.')[0] + '.xlsx'
   new_wb = Workbook()
   new_wb.remove(new_wb.active)  # remove default sheet

   # copy all sheets from original file to new file
   for sheet_name in wb.sheetnames:
       sheet = wb[sheet_name]
       new_sheet = new_wb.create_sheet(sheet.title)
       for row in sheet.iter_rows(values_only=True):
           new_sheet.append(row)
   ```

3. Save the changes and run the script as described in the "Running the Script" section.

By making these changes, the script will now copy all the sheets from the original Excel file to the converted `.xlsx` file, preserving their data and structure.

## Troubleshooting

If you encounter any issues while running the script, refer to the `Process.log` file for more information. Each operation is logged with a timestamp, which should help you identify when and where the problem occurred.

If an Excel file can't be converted, ensure that it's not open in another program, as this may prevent the script from accessing it. If the problem persists, the file may be corrupt or not a valid Excel file.

## Contributions

Contributions are welcome! Please feel free to submit a Pull Request or open an Issue if you find a bug or think of a new feature that could be added.

## License

This project is licensed under the MIT License. See the LICENSE file for more details.
