# Working Code
import pandas as pd
import sys
import os
def accounting_to_float(value):
    """Converts accounting-style numbers to float, handling commas and non-numeric values."""
    if isinstance(value, str):
        value = value.replace(',', '')
    try:
        return float(value)
    except ValueError:
        return 0.0
def process_csv(file_path):
    try:
        column_types = {
            "Grant Code": str,
        }
        df = pd.read_csv(file_path, dtype=column_types, encoding="utf-8", low_memory=False)
        required_columns = {"Debit", "Credit", "Check Address Code"}
        missing_columns = required_columns - set(df.columns)
        if missing_columns:
            print(f"Error: Missing columns in the file: {missing_columns}")
            return None
        df["Credit"] = df["Credit"].apply(accounting_to_float)
        df["Debit"] = df["Debit"].apply(accounting_to_float)
        df[["Credit", "Debit"]] = df[["Credit", "Debit"]].fillna(0.0)
        credit_index = df.columns.get_loc("Credit")
        df.insert(credit_index + 1, "NET", df["Debit"] - df["Credit"])
        output_file = os.path.splitext(file_path)[0] + "_rectified.csv"
        df.to_csv(output_file, index=False, encoding="utf-8")
        print(f"Rectified file saved: {output_file}")
        return output_file
    except Exception as e:
        print(f"Error processing CSV file: {e}")
        return None
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python process_csv.py <file_path>")
    else:
        file_path = sys.argv[1]
        if os.path.exists(file_path):
            rectified_file = process_csv(file_path)
            if rectified_file:
                print(f"Pass this file to the web app: {rectified_file}")
        else:
            print(f"File not found: {file_path}")



# # Actual tested code which is working perfectly for rectifying the csv file:
# def accounting_to_float(value):
#     if isinstance(value, str):
#         value = value.replace(',', '')
#     try:
#         return float(value)
#     except ValueError:
#         return 0.0
# data['Credit'] = data['Credit'].apply(accounting_to_float)
# data['Debit'] = data['Debit'].apply(accounting_to_float)
# data[['Credit', 'Debit']] = data[['Credit', 'Debit']].fillna(0.0)
# data.insert(data.columns.get_loc('Credit') + 1, 'NET', data['Debit'] - data['Credit'])