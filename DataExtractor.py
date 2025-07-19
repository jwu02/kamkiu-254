from datetime import datetime
import itertools

import pandas as pd

from constants import (
    customer_match, 
    location_match,
    CheckStatus,
    MODEL_CODE_ORDER,
    SAMPLE_TYPES,
    OQC_RETENTION_SAMPLE_CODES,
    SCHEMA_CODE,
    CUSTOMER_PART_CODE,
    PART_NAME,
    CUSTOMER_CODE_EN,
)

class DataExtractor:
    def extract_shipment_batch_data(self, response_data: list) -> pd.DataFrame:
        title_list = [list(title_obj.keys())[0] for title_obj in response_data['titleList']]
        data = response_data['list']
        df = pd.DataFrame(data, columns=title_list)

        df_shipment_batch = df.reindex(columns=[
            '地区',
            '项目',
            'zfhs', # 发货数
            'zfhrq', # 发货日期
            'zbm', # 型号
            '图号',
            '合金',
            '回收比',
            'jy_no2', # 第二列 挤压批号（有两列）
            '挤压批（二维码）',
            '挤压批次二维码',
            '模号',
            'smelt_lot', #炉号
            '熔铸批号',
            'sx_no', # 时效批号
            '时效炉',
            '时效批次（二维码）',
            '客户料号',
            '客户批号',
            '客户',
        ])

        # Create a mapping dictionary
        column_mapping = self.get_column_name_mapping(response_data['titleList'])

        # Rename the columns
        df_shipment_batch = df_shipment_batch.rename(columns=column_mapping)
        
        # Apply to DataFrame
        df_shipment_batch['地区'] = df['zkhdq'].apply(self.extract_location) # 客户/地区
        df_shipment_batch['客户'] = df['zkhdq'].apply(self.extract_customer)
        df_shipment_batch['项目'] = 'Manchester'
        df_shipment_batch['图号'] = df_shipment_batch['型号'].apply(lambda model_code: SCHEMA_CODE[model_code])
        df_shipment_batch['合金'] = '7R03'
        df_shipment_batch['回收比'] = '50%'
        df_shipment_batch['挤压批号'] = df_shipment_batch['挤压批号'].apply(self.transform_extrusion_batch_code)
        df_shipment_batch['模号'] = df_shipment_batch['挤压批号'].apply(self.extract_die_code)
        df_shipment_batch['时效炉'] = df_shipment_batch['时效批号'].apply(lambda model_code: model_code[3:5])
        df_shipment_batch['客户料号'] = df_shipment_batch['型号'].apply(lambda model_code: CUSTOMER_PART_CODE[model_code])
        df_shipment_batch['客户批号'] = ''

        df_shipment_batch['型号'] = pd.Categorical(
            df_shipment_batch['型号'], 
            categories=MODEL_CODE_ORDER, 
            ordered=True
        )
        df_shipment_batch.sort_values(by=['地区', '客户', '型号', '炉号', '发货数', '挤压批号', '时效批号'], inplace=True)
        df_shipment_batch.reset_index(drop=True, inplace=True)

        df_shipment_batch['CPK'] = CheckStatus.NOT_CHECKED.value
        df_shipment_batch['性能'] = CheckStatus.NOT_CHECKED.value
        df_shipment_batch['成分'] = CheckStatus.NOT_CHECKED.value

        return df_shipment_batch
    
    def get_column_name_mapping(self, unflattened: list) -> dict:
        # Create a mapping dictionary
        column_mapping = {}
        for item in unflattened:
            column_mapping.update(item)

        print(column_mapping)
        return column_mapping

    def extract_location(self, name: str) -> str:
        matched_location = [v for k, v in location_match.items() if k in name]
        location = sorted(set(matched_location))

        return location[0] if location else '地区未录入'

    def extract_customer(self, name: str) -> str:
        matched_customer = [v for k, v in customer_match.items() if k in name]
        customer = sorted(set(matched_customer))

        return customer[0] if customer else '客户未录入'

    def transform_extrusion_batch_code(self, code: str) -> str:
        if '-' not in code and len(code)==15:
            return code

        parts = code.split('-')
        
        xx = parts[0]
        die_code = parts[1] if len(parts[1])==4 else '0'+parts[1]

        # Get last two digits of current year as a string
        year_suffix = str(datetime.now().year)[-2:]  # '25' for 2025
        extrusion_date = year_suffix + parts[2]

        return xx + die_code + extrusion_date

    def extract_die_code(self, extrusion_batch_code: str) -> str:
        die_code = extrusion_batch_code[2:6]
        if die_code[0] == '0':
            die_code = die_code[1:]
        
        return die_code

    def extract_ageing_qrcode_data(self, response_data: dict) -> pd.DataFrame:
        df = pd.DataFrame(response_data['list'])

        # Rename the columns we need
        # Manually create column name mapping, column heading codes doesnt match with obj keys
        df = df.rename(columns={
            'zbm': '型号',
            'jyPrd': '生产挤压批',
            'smeltLot': '铝棒炉号',
            'jyCode': '挤压批',
            'rzCode': '熔铸批号',
        })

        df_ageing_qrcode = df[[
            '型号',
            '生产挤压批',
            '铝棒炉号',
            '挤压批',
            '熔铸批号',
        ]]

        # df_ageing_qrcode.sort_values(by=['型号', '铝棒炉号', '生产挤压批'], inplace=True)
        # df_ageing_qrcode.reset_index(drop=True, inplace=True)

        return df_ageing_qrcode

    def extract_process_card_qrcode_data(self, response_data: dict) -> pd.DataFrame:
        df = pd.DataFrame(response_data['list'])

        df = df.rename(columns={
            'zbm': '型号',
            'jyNo': '挤压批号',
            'zlh': '炉号',
            'sfc': '时效批',
            'qrcode': '二维码',
        })

        df_process_card_qrcode = df[[
            '型号',
            '挤压批号',
            '炉号',
            '时效批',
            '二维码',
        ]]
        df_process_card_qrcode['时效批'] = df_process_card_qrcode['时效批'].apply(lambda x: x[:8])

        return df_process_card_qrcode

    def fill_data_from_ageing_qrcode(self, df_shipment_batch: pd.DataFrame, df_ageing_qrcode: pd.DataFrame) -> pd.DataFrame:
        """
        填入 挤压批 & 熔铸批号 二维码
        """
        for index, row in df_shipment_batch.iterrows():
            model_code = row['型号']
            extrusion_batch_code = row['挤压批号']
            furnace_code = row['炉号']

            df = df_ageing_qrcode[
                (df_ageing_qrcode['型号'] == model_code) &
                (df_ageing_qrcode['生产挤压批'] == extrusion_batch_code) &
                (df_ageing_qrcode['铝棒炉号'] == furnace_code)
            ]
            
            df_shipment_batch.at[index, '挤压批（二维码）'] = df.iloc[0]['挤压批'] if len(df)>0 else "🟠 没记录"
            df_shipment_batch.at[index, '熔铸批号'] = df.iloc[0]['熔铸批号'][2:] if len(df)>0 else "🟠 没记录"

        df_shipment_batch['挤压批次二维码'] = df_shipment_batch['挤压批（二维码）'].apply(lambda x: str(x).split('+')[-1])

        return df_shipment_batch

    def fill_data_from_process_card_qrcode(self, df_shipment_batch: pd.DataFrame, df_process_card_qrcode: pd.DataFrame) -> pd.DataFrame:
        """
        从 流程卡二维码记录 采取 时效批次二维码
        """
        for index, row in df_shipment_batch.iterrows():
            model_code = row['型号']
            extrusion_batch_code = row['挤压批号']
            furnace_code = row['炉号']
            ageing_code = row['时效批号']

            df = df_process_card_qrcode[
                (df_process_card_qrcode['型号'] == model_code) &
                (df_process_card_qrcode['挤压批号'] == extrusion_batch_code) &
                (df_process_card_qrcode['炉号'] == furnace_code) &
                (df_process_card_qrcode['时效批'] == ageing_code)
            ]

            df_shipment_batch.at[index, '时效批次（二维码）'] = df.iloc[0]['二维码'][-4:] if len(df)>0 else "🟠 没记录"
        
        return df_shipment_batch

    def extract_mechanical_properties_data(self, response_data: dict) -> pd.DataFrame:
        df = pd.DataFrame(response_data['list'])

        # TODO: Create dataframe column title remapping/rename

        df_functional_properties = df[[
            '检测项目',
            '型号',
            # '挤压批次',
            '铝棒炉号',
            '时效炉号',
            'oqc样号',
            '点位',
            '硬度值',
            '电导率',
            '非比例延伸强度',
            '抗拉强度',
            '断后伸长率',
            '平均截距',
            '最大晶粒尺寸',
            '横纵比',
            '第二相尺寸',
        ]]

        df_functional_properties = df_functional_properties[
            df_functional_properties['oqc样号'].isin(list(itertools.chain.from_iterable(OQC_RETENTION_SAMPLE_CODES)))
        ]
        
        # df_functional_properties.sort_values(by=[
        #     '型号',
        #     '时效炉号',
        #     # '铝棒炉号',
        #     '检测项目',
        # ], inplace=True)
        # df_functional_properties.reset_index(drop=True, inplace=True)

        return df_functional_properties
    
    def extract_chemical_composition_data(self, response_data: dict, compositions: list[str]) -> pd.DataFrame:
        columns = self.get_column_name_mapping(response_data['titleList'])
        df = pd.DataFrame(response_data['list'])
        df = df.rename(columns={
            'process_lot': '炉号',
            'type': '类型'
        })
        df = df.rename(columns=columns)
        print(f"Removing duplicated columns from chemical composition DataFrame: {df.columns[df.columns.duplicated()].tolist()}")
        df = df.loc[:, ~df.columns.duplicated()]

        df_composition = df.reindex(columns=['炉号', '类型', *compositions])
        df_composition['Mn+Cr'] = 0
        df_composition = df_composition[df_composition['类型'].isin(SAMPLE_TYPES)]
        df_composition = df_composition.replace('-', pd.NA).dropna(how='any')

        # df_composition.sort_values(by=['炉号', '类型'], ascending=[True, False], inplace=True)
        # df_composition.reset_index(drop=True, inplace=True)
        df_composition['Mn+Cr'] = round(df_composition['Mn'].astype(float) + df_composition['Cr'].astype(float), 5)

        return df_composition
    
    def extract_customer_shipment_details(self, df_shipment_batch: pd.DataFrame) -> pd.DataFrame:
        df_customer_shipment_details = df_shipment_batch.reindex(columns=[
            '客户料号',
            '品名',
            '发货日期',
            '挤压批号',
            '挤压批（二维码）',
            '发货数',
            '炉号',
            '熔铸批号',
            '型材厂商',
            'Makeup',
            '标签',
            '阶段',
            '地区',
            '客户',
        ])

        df_customer_shipment_details['品名'] = df_shipment_batch['型号'].map(PART_NAME)
        df_customer_shipment_details['型材厂商'] = 'KAP'
        df_customer_shipment_details['Makeup'] = '50% prime + 50% IP'
        df_customer_shipment_details['标签'] = ''
        df_customer_shipment_details['阶段'] = 'MP'
        df_customer_shipment_details['客户'] = df_customer_shipment_details['客户'].map(CUSTOMER_CODE_EN)

        return df_customer_shipment_details
    
    def extract_test_commission_form_data(self, response_data: dict) -> pd.DataFrame:
        df = pd.DataFrame(response_data['list'])

        df_test_commission_form = df[[
            '委托单号',
            '检测项目',
            '检验结果',
            '型号',
            '挤压批次',
            '铝棒炉号',
            '时效炉号',
        ]]

        # df_test_commission_form.sort_values(by=[
        #     '型号',
        #     '时效炉号',
        #     '铝棒炉号',
        #     '检测项目',
        # ], inplace=True)

        df_test_commission_form = df_test_commission_form.replace('-', None)

        return df_test_commission_form