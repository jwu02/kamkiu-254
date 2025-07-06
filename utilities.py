from datetime import datetime
from enum import Enum
import pandas as pd
import os

# Get last two digits of current year as a string
year_suffix = str(datetime.now().year)[-2:]  # '25' for 2025

def transform_extrusion_batch_code(code):
    if '-' not in code and len(code)==15:
        return code

    parts = code.split('-')
    
    xx = parts[0]
    die_code = parts[1] if len(parts[1])==4 else '0'+parts[1]
    extrusion_date = year_suffix + parts[2]

    return xx + die_code + extrusion_date


customer_match = {
    '精密': '无锡精密',
    'EPZ': '无锡精密',
    '金属': '无锡比亚迪',
}

location_match = {
    '华阳': 'HY',
    '郎克斯': 'LKS',
    '朗克斯': 'LKS',
    'LKS': 'LKS'
}

def extract_location(name):
    matched_location = [v for k, v in location_match.items() if k in name]
    location = sorted(set(matched_location))

    return location[0] if location else '地区未录入'

def extract_customer(name):
    matched_customer = [v for k, v in customer_match.items() if k in name]
    customer = sorted(set(matched_customer))

    return customer[0] if customer else '客户未录入'


# # Normalize 时效批 for 流程卡二维码记录
# def normalize_ageing_batch_code(code: str):
#     return code[:8]

class CheckStatus(Enum):
    NOT_CHECKED = '⚪️ 未检查'
    OK = '🟢 OK'
    NG = '🔴 NG'
    NO_DATA = '🟠 找不到数据'

MODEL_CODE_MAPPINGS = {
    'KAP-7457上U-A76-50': {
        'cpk': {
            'path': r'\\192.168.3.18\品质qe小组\A-PM\254\CPK\EVT\发货\新版\7457',
            'tolerance': './data/尺寸公差/尺寸公差_7457.csv',
            'num_rows': 51, # Num rows to copy from CPK datasheet
        },
        'schema_code': '806-55322-04',
        'customer_part_code': '18242741-00',
        'composition': {
            'start_row': 81, # 从第81行开始粘贴数据
            'start_column': 9, # 从第9列（I）开始黏贴数据
        }
    },
    'KAP-7461中板-A76-50': {
        'cpk': {
            'path': r'\\192.168.3.18\品质qe小组\A-PM\254\CPK\EVT\发货\新版\7461',
            'tolerance': './data/尺寸公差/尺寸公差_7461.csv',
            'num_rows': 56, # Num rows to copy from CPK datasheet
        },
        'schema_code': '806-55327-09',
        'customer_part_code': '18242780-00',
        'composition': {
            'start_row': 96, # 从第96行开始粘贴数据
            'start_column': 9, # 从第9列（I）开始黏贴数据
        }
    },
    'KAP-7487下U-A76-50': {
        'cpk': {
            'path': r'\\192.168.3.18\品质qe小组\A-PM\254\CPK\EVT\发货\新版\7487',
            'tolerance': './data/尺寸公差/尺寸公差_7487.csv',
            'num_rows': 64, # Num rows to copy from CPK datasheet
        },
        'schema_code': '806-55323-05',
        'customer_part_code': '18242767-00',
        'composition': {
            'start_row': 94, # 从第94行开始粘贴数据
            'start_column': 9, # 从第9列（I）开始黏贴数据
        }
    },
}

# 型号 自定义排序
MODEL_CODE_ORDER = [
    'KAP-7461中板-A76-50', 'KAP-7461中板-A76-85', 
    'KAP-7457上U-A76-50', 'KAP-7457上U-A76-85',
    'KAP-7487下U-A76-50', 'KAP-7487下U-A76-85'
]

# 成分样品类型
SAMPLE_TYPES = [
    # '1350-铸造1/2L送样（3米）',
    '08-型材成分检验(尾）',
    '08-型材成分检验（尾）',
]

def get_sample_size(batch_size):
    sample_size = 0
    if batch_size < 2: return "不够数量"
    elif batch_size <= 32: return "全检"
    elif batch_size <= 500: sample_size = 32
    elif batch_size <= 3200: sample_size = 125
    elif batch_size <= 10000: sample_size = 200
    elif batch_size <= 35000: sample_size = 315
    elif batch_size <= 150000: sample_size = 500
    elif batch_size <= 500000: sample_size = 800
    else: sample_size = 1250

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


def get_report_name(
    model_code: str, 
    customer_part_code: str, 
    shipment_batch_quantity: int, 
    location: str, 
    furnace_code: str, 
    extrusion_batch_code: str
):
    return f"MANCHESTER {model_code} {customer_part_code} {shipment_batch_quantity} ({location}) {furnace_code} {extrusion_batch_code}"

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

def find_files_with_substring(directory, substring):
    """Return list of filenames in directory that contain substring"""
    matching_files = []
    for filename in os.listdir(directory):
        if substring.lower() in filename.lower():
            matching_files.append(filename)
    return matching_files
