# Excel Column Filter  
  
This application allows you to filter and save specific columns from an Excel file. The filtered data is saved in a new Excel file with a name that includes the original file name and the current date and time. The new file is saved in the `out` directory.  
  
## Requirements  
  
- Python 3.x  
- pandas  
- openpyxl
  
## Installation  
  
1. Clone the repository or download the source code.  
2. Navigate to the project directory.  
3. Install the required Python packages using `requirements.txt`:  
  
    ```sh  
    pip install -r requirements.txt  
    ```  
  
## Usage  
  
1. Place the Excel files you want to process in the `source` directory.  
2. Run the script:  
  
    ```sh  
    python script_name.py  
    ```  
  
    Replace `script_name.py` with the actual name of your script file.  
  
3. Follow the prompts:  
    - Select the Excel file by entering the corresponding number from the displayed list.  
    - Enter the column numbers you want to keep, separated by commas (e.g., `1, 3, 5`). **Use a comma (`,`) to input multiple column numbers.**  
    - The script will create a new Excel file in the `out` directory with the filtered columns.  
  
## Example  
  
Assume you have an Excel file named `example.xlsx` in the `source` directory with the following columns:  
  
- Name  
- Age  
- Email  
- Country  
  
When you run the script, you will be prompted to select the file and the columns to keep. If you choose columns 1 and 3 (`Name` and `Email`), the script will create a new file in the `out` directory with a name like `example_20220329_123456.xlsx`, containing only the `Name` and `Email` columns.  
  
## Notes  
  
- The script will prompt you to enter the exact file name if you provide an invalid file number.  
- The output file will be saved in the `out` directory.  
- The output file name format is `{original_name}_{date_and_time}.xlsx`.  
- **Use a comma (`,`) to input multiple column numbers when prompted to enter column numbers to keep.**  
  
## License  
  
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.  