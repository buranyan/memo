# 2025/2/22 ver1.0.0
import pandas as pd
import time

def excel_sheet_to_csv(excel_file, sheet_name, output_csv):
    """
    Convert a specific sheet from an Excel file to a CSV file.
    :param excel_file: Path to the input Excel file
    :param sheet_name: Name of the sheet to be converted
    :param output_csv: Path to the output CSV file
    """
    try:
        # Load the specified sheet into a DataFrame
        df = pd.read_excel(excel_file, sheet_name=sheet_name, engine='openpyxl')
        # Save as CSV
        df.to_csv(output_csv, index=False, encoding='utf-8-sig')
        print(f"Sheet '{sheet_name}' has been successfully saved as '{output_csv}'.")
    except Exception as e:
        print(f"Error: {e}")

# Example usage
if __name__ == "__main__":
    excel_file = "input_file_large.xlsx"  # Input Excel file
    sheet_name = "Sheet1"        # Sheet to convert
    output_csv = "panda.csv"     # Output CSV file

    time_start = time.time()
    print("<処理開始>")
    
    excel_sheet_to_csv(excel_file, sheet_name, output_csv)

    time_end = time.time() - time_start
    print("<処理時間>", time_end)
