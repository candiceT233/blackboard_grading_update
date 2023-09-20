# Blackboard Grading Sheet Updater

## Overview

This Python script allows you to update a master grading sheet with information from a partial grading sheet. 
It particularly updates the assignment grades and the comments for displaying in HTML format.

It's especially useful for managing student grades in a Blackboard course.

## Prerequisites

Before using this script, ensure you have the following:

- Python 3.x installed on your system.
- Required Python packages installed. You can install them using `pip`:

    ```
    pip install pandas
    ```

## Prepare the `master_grade_data.xlsx`

1. Follow this tutorial to download the grading sheet for the whole class: https://help.blackboard.com/Learn/Instructor/Original/Grade/Grading_Tasks/Work_Offline_With_Grade_Data

2. Make sure to download your sheet with:

- `Selected Column` 
- checking the box ` Include Comments for this Column`
- `Delimiter Type : ` choose `Comma`

3. Open your downloaded `master_grade_data.csv` file, open it and save as `master_grade_data.xlsx` to avoid python string formatting problem

In general, make sure the master grading data file has these columns exactly:

- `First Name`,`Last Name`
- `Total Pts`: for recording student grades 
- `Grading Notes`: for grading comments


## Prepare the `partial_grade_data.csv`

1. Format your workbook with columns `First Name`, `Last Name`, `Final Score`, and `All Comments`

2. `All Comments` example format as below:
```txt
Q1: good
Q2: (-0.25) incorrect.
Q3: Great
Q4: -
Q5: b:
(a) comment for a.
(d) comment for b.
Total:(-3 pts)
Q6: (-1 pt) c: Some comment
General Comment: Good.
```
- Each problem and sub-problem has comment and point subtraction within its own line.


## Usage

1. Clone this repository to your local machine or download the script file.

2. Open your terminal or command prompt.

3. Run the script with the following command:

    ```
    python update_grading_sheets.py master_grade_data.xlsx partial_grade_data.csv updated_master_grade_data.csv
    ```

    - Replace `master_grade_data.xlsx` with the path to your master grading sheet (in EXCEL format).
    - Replace `partial_grade_data.csv` with the path to your partial grading sheet (in CSV format).
    - Replace `updated_master_grade_data.csv` with the desired output file name for the updated master grading sheet.

4. The script will merge the two sheets based on "Last Name" and "First Name" columns, update the master grading sheet, and save the updated sheet as `updated_master_grade_data.csv`.

## CSV File Format

- **Master Grading Sheet**: The master grading sheet should be in CSV format and include at least the "Last Name" and "First Name" columns.

- **Partial Grading Sheet**: The partial grading sheet should also be in CSV format, with columns containing additional grading information.

Please make sure the column names for "Last Name" and "First Name" are identical in both sheets.

## Additional Notes

- The script assumes that the "partial grading sheet" contains additional columns that need to be merged into the "master grading sheet." It will update the master grading sheet with any missing data from the partial grading sheet.

- If you encounter any issues or have specific requirements for the update process, you can customize the script to suit your needs.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

