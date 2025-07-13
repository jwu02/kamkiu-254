from enum import Enum

class CheckStatus(Enum):
    NOT_CHECKED = '⚪️ 未检查'
    OK = '🟢 OK'
    NG = '🔴 NG'
    NO_DATA = '🟠 找不到数据'

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

# oqc样品
OQC_RETENTION_SAMPLE_CODES = [
    ['Hot 1#','Cool 1#','Hot 2#','Cool 2#',],
    ['First Tail-AVG','Last Tail-AVG',],
    ['First Tail','Last Tail',],
    ['Cavity 1 First Tail-AVG','Cavity 1 Last Tail-AVG','Cavity 1 First Head-AVG','Cavity 1 Last Head-AVG',],
    ['Cavity 1 First Tail','Cavity 1 Last Tail','Cavity 1 First Head','Cavity 1 Last Head',],
]

REPORT_OUTPUT_PATH = './报告输出'

MODEL_CODE_MAPPINGS = {
    'KAP-7457上U-A76-50': {
        'cpk': {
            # 'path': r'\\192.168.3.18\品质qe小组\A-PM\254\CPK\EVT\发货\新版\7457',      # CPK路径
            'path': './test_files/cpk_datasheets/7457',
            'tolerance': './data/尺寸公差/尺寸公差_7457.csv',                            # CPK尺寸公差 来检查CPK合不合格（未有此检查功能，用不到）
            'num_rows': 51,                                                             # 从CPK数据表复制多少行数据
        },
        '性能': {
            'start_row': 65,
            'start_column': 9, # I
        },
        'composition': {
            # 报告模板 成分部分开始的单元格点位 用来决定从报告模板里哪里粘贴数据
            'start_row': 81, # 第81行
            'start_column': 9, # 第9列（I）
        },
        '重量': {
            'lower_limit': 291,                                                         # 重量下限
            'upper_limit': 299,                                                         # 重量上限
            # 报告模板 重量部分开始的单元格点位 用来决定从报告模板里哪里粘贴数据
            'starting_row': 108,
            'starting_column': 9, # I
        },
    },
    'KAP-7461中板-A76-50': {
        'cpk': {
            # 'path': r'\\192.168.3.18\品质qe小组\A-PM\254\CPK\EVT\发货\新版\7461',
            'path': './test_files/cpk_datasheets/7461',
            'tolerance': './data/尺寸公差/尺寸公差_7461.csv',
            'num_rows': 56,
        },
        '性能': {
            'start_row': 70,
            'start_column': 9,
        },
        'composition': {
            'start_row': 96,
            'start_column': 9,
        },
        '重量': {
            'lower_limit': 2572,
            'upper_limit': 2585,
            'starting_row': 123,
            'starting_column': 9,
        },
    },
    'KAP-7487下U-A76-50': {
        'cpk': {
            # 'path': r'\\192.168.3.18\品质qe小组\A-PM\254\CPK\EVT\发货\新版\7487',
            'path': './test_files/cpk_datasheets/7487',
            'tolerance': './data/尺寸公差/尺寸公差_7487.csv',
            'num_rows': 64,
        },
        '性能': {
            'start_row': 78,
            'start_column': 9,
        },
        'composition': {
            'start_row': 94,
            'start_column': 9,
        },
        '重量': {
            'lower_limit': 221,
            'upper_limit': 229,
            'starting_row': 121,
            'starting_column': 9,
        },
    },
}

# 客户料号
CUSTOMER_PART_CODE = {
    'KAP-7461中板-A76-50': '18242780-00',
    'KAP-7457上U-A76-50': '18242741-00',
    'KAP-7487下U-A76-50': '18242767-00',
}

# 图号
SCHEMA_CODE = {
    'KAP-7461中板-A76-50': '806-55327-09',
    'KAP-7457上U-A76-50': '806-55322-04',
    'KAP-7487下U-A76-50': '806-55323-05',
}

# 品名
PART_NAME = {
    'KAP-7461中板-A76-50': '铝挤板_156.8x81.4x11.1MM_7R03_C',
    'KAP-7457上U-A76-50': '铝挤板_81.6x17.2x10.7MM_7R03_C',
    'KAP-7487下U-A76-50': '铝挤板_81.6x17.8x9.7MM_7R03_C',
}

# 型号 自定义排序
MODEL_CODE_ORDER = [
    'KAP-7461中板-A76-50', 'KAP-7461中板-A76-85', 
    'KAP-7457上U-A76-50', 'KAP-7457上U-A76-85',
    'KAP-7487下U-A76-50', 'KAP-7487下U-A76-85',
]

# 成分样品类型
SAMPLE_TYPES = [
    # '1350-铸造1/2L送样（3米）',
    '08-型材成分检验(尾）',
    '08-型材成分检验（尾）',
]