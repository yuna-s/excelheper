import os  
import pandas as pd  
from datetime import datetime  
  
def list_excel_files(directory):  
    """List all Excel files in the given directory."""  
    return [f for f in os.listdir(directory) if f.endswith('.xlsx') or f.endswith('.xls')]  
  
def filter_excel_columns(source_file, destination_file, columns_to_keep):  
    try:  
        # Step 1: Read the source Excel file  
        df = pd.read_excel(source_file)  
          
        # Step 2: Check if specified columns exist  
        missing_columns = [col for col in columns_to_keep if col not in df.columns]  
        if missing_columns:  
            raise ValueError(f"Columns not found in the source file: {missing_columns}")  
          
        # Step 3: Filter the columns  
        filtered_df = df[columns_to_keep]  
          
        # Step 4: Write the filtered data to a new Excel file  
        filtered_df.to_excel(destination_file, index=False)  
        print(f"Filtered Excel file has been created: {destination_file}")  
      
    except FileNotFoundError:  
        print(f"Source file not found: {source_file}")  
    except ValueError as ve:  
        print(ve)  
    except Exception as e:  
        print(f"An error occurred: {e}")  
  
if __name__ == "__main__":  
    source_dir = 'source'  
    out_dir = 'out'  
      
    # Ensure the output directory exists  
    if not os.path.exists(out_dir):  
        os.makedirs(out_dir)  
      
    # List Excel files in the source directory  
    excel_files = list_excel_files(source_dir)  
      
    # Check if there are any Excel files in the directory  
    if not excel_files:  
        print(f"No Excel files found in the directory: {source_dir}")  
    else:  
        # Display the list of Excel files  
        print("Excel files available in the 'source' directory:")  
        for index, file in enumerate(excel_files):  
            print(f"{index + 1}. {file}")  
          
        # Ask the user to select an Excel file  
        try:  
            file_choice = int(input(f"Select the Excel file by number (1-{len(excel_files)}) or enter the file name: "))  
            if 1 <= file_choice <= len(excel_files):  
                source_file = os.path.join(source_dir, excel_files[file_choice - 1])  
            else:  
                raise ValueError("Invalid selection")  
        except ValueError:  
            # If user directly enters the filename  
            source_file = os.path.join(source_dir, input("Enter the exact file name: "))  
          
        # Read the selected Excel file to get column names  
        df = pd.read_excel(source_file)  
        print("Available columns in the selected file:")  
        for index, column in enumerate(df.columns):  
            print(f"{index + 1}. {column}")  
          
        # User input for columns to keep  
        columns_to_keep_indices = input("Enter the column numbers to keep (separated by commas): ").split(',')  
        columns_to_keep = [df.columns[int(index) - 1] for index in columns_to_keep_indices if index.strip().isdigit()]  
          
        # Generate the default destination file name  
        original_name = os.path.splitext(os.path.basename(source_file))[0]  
        date_time_str = datetime.now().strftime("%Y%m%d_%H%M%S")  
        destination_file = os.path.join(out_dir, f"{original_name}_filtered_{date_time_str}.xlsx")  
          
        # Call the function to filter columns and create the new Excel file  
        filter_excel_columns(source_file, destination_file, columns_to_keep) 