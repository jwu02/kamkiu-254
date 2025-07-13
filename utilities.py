import os
import pandas as pd

from PyQt6.QtWidgets import (
    QMessageBox, QFileDialog
)

def load_cpk_tolerance_map():
    df_cpk_7457 = pd.read_csv('./data/尺寸公差/尺寸公差_7457.csv')
    df_cpk_7461 = pd.read_csv('./data/尺寸公差/尺寸公差_7461.csv')
    df_cpk_7487 = pd.read_csv('./data/尺寸公差/尺寸公差_7487.csv')

    cpk_tolerance_map = {
        'KAP-7457上U-A76-50': df_cpk_7457,
        'KAP-7461中板-A76-50': df_cpk_7461,
        'KAP-7487下U-A76-50': df_cpk_7487,
    }

    return cpk_tolerance_map

def check_cpk_conformance(cpk_path, df_cpk_tolerance) -> str:
    """
    Check CPk conformance given cpk_path and a df containing CPK tolerances

    TODO: understand how CPk works, might not even be using the correct CPK tolerance
    """
    try:
        # cpk_path = os.path.join(path, matching_files[0])

        # wb = load_workbook(filename=cpk_path)
        # sheet = wb.active
        
        # print(f"\nContents of {target_file}:")
        # print("----------------------------------------")
        # # for row in sheet.iter_rows(values_only=True):
        # #     print(row)
        # print("----------------------------------------")

        # # ================ get data from existing CPK files
        # df = pd.read_excel(
        #     cpk_path,
        #     engine='openpyxl',
        #     header=None,  # No headers (since we're reading raw cells)
        #     usecols="AH:AT",  # Columns from AH to AT
        #     skiprows=10,  # Skip first 10 rows (to start at row 11)
        #     nrows=56,  # Read 66 rows (11 to 76 → 76-10=66)
        # )
        
        return "🟢 存在"

    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return "🔴 错误"

def find_files_with_substrings(directory: str, substrings: list[str]) -> list[str]:
    """
    Return list of filenames in directory that contain substring
    """
    matches = []

    for f in os.listdir(directory):
        if all(sub in f for sub in substrings):
            matches.append(os.path.join(directory, f))

    return matches

def condense_row(row):
    values = [v for v in row if not pd.isna(v)]
    return values + [None] * (len(row) - len(values))

def get_mechanical_electrical_df_mask(df):
    # Create new DataFrame with False values
    df_mask = df.copy().astype(bool)
    df_mask.loc[:, :] = False
    df_mask.iloc[:, :2] = True  # Columns 0 and 1 (A and B)
    df_mask.iloc[-4, :] = True  # Entire row

    return df_mask

def get_metallographic_df_mask(df):
    df_mask = df.copy().astype(bool)
    df_mask.loc[:, :] = False
    df_mask.iloc[:, :2] = True  # Columns 0 and 1 (A and B)

    return df_mask

def show_error(msg: str):
    # Create and show a warning message box
    msg_box = QMessageBox()
    msg_box.setIcon(QMessageBox.Icon.Warning)
    msg_box.setText(msg)
    msg_box.setWindowTitle("注意")
    msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
    msg_box.exec()

def show_info(msg: str):
    # Create and show a informative message box
    msg_box = QMessageBox()
    msg_box.setIcon(QMessageBox.Icon.Information)
    msg_box.setText(msg)
    msg_box.setWindowTitle("信息")
    msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
    msg_box.exec()
