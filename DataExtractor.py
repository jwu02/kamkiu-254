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
)

class DataExtractor:
    def extract_shipment_batch_data(self, df: pd.DataFrame) -> pd.DataFrame:
        df_shipment_batch = df.reindex(columns=[
            '地区',
            '项目',
            '发货数',
            '发货日期',
            '型号',
            '图号',
            '合金',
            '回收比',
            '挤压批号', # 第二列 挤压批号（有两列）
            '挤压批（二维码）',
            '挤压批次二维码',
            '模号',
            '炉号',
            '熔铸批号',
            '时效批号',
            '时效炉',
            '时效批次（二维码）',
            '客户料号',
            '客户批号',
            '客户',
        ])
        
        # Apply to DataFrame
        df_shipment_batch['地区'] = df['客户/地区'].apply(self.extract_location)
        df_shipment_batch['客户'] = df['客户/地区'].apply(self.extract_customer)
        df_shipment_batch['项目'] = 'Manchester'
        df_shipment_batch['图号'] = df_shipment_batch['型号'].apply(lambda model_code: SCHEMA_CODE[model_code])
        df_shipment_batch['合金'] = '7R03'
        df_shipment_batch['回收比'] = '50%'
        df_shipment_batch['挤压批号'] = df_shipment_batch['挤压批号'].apply(self.transform_extrusion_batch_code)
        df_shipment_batch['模号'] = ''
        df_shipment_batch['时效炉'] = ''
        df_shipment_batch['客户料号'] = df_shipment_batch['型号'].apply(lambda model_code: CUSTOMER_PART_CODE[model_code])
        df_shipment_batch['客户批号'] = ''
        df_shipment_batch['时效批号（sfc）'] = df_shipment_batch['时效批号'].apply(lambda x: x+'*')

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

    def extract_functional_properties_data(self, df: pd.DataFrame) -> pd.DataFrame:
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
    
    def extract_chemical_composition_data(self, df: pd.DataFrame, compositions: list[str]) -> pd.DataFrame:
        df = df.reindex(columns=['炉号', '类型', *compositions])
        df['Mn+Cr'] = 0
        df = df[df['类型'].isin(SAMPLE_TYPES)]
        df = df.replace('-', pd.NA).dropna(how='any')

        # df.sort_values(by=['炉号', '类型'], ascending=[True, False], inplace=True)
        # df.reset_index(drop=True, inplace=True)
        df['Mn+Cr'] = round(df['Mn'].astype(float) + df['Cr'].astype(float), 5)

        return df

    def extract_ageing_qrcode_data(self, df: pd.DataFrame) -> pd.DataFrame:
        df_filtered = df[[
            '型号',
            '生产挤压批',
            '铝棒炉号',
            '内部时效批',
            '挤压批',
            '熔铸批号',
        ]]

        # df_filtered.sort_values(by=['型号', '铝棒炉号', '生产挤压批'], inplace=True)
        # df_filtered.reset_index(drop=True, inplace=True)

        return df_filtered

    def extract_process_card_qrcode_data(self, df: pd.DataFrame) -> pd.DataFrame:
        df_filtered = df[[
            '型号',
            '挤压批号',
            '炉号',
            'sfc',
            '二维码',
        ]]

        df_filtered.rename(columns={'sfc': '时效批'}, inplace=True)
        df_filtered['时效批'] = df_filtered['时效批'].apply(lambda x: x[:8])

        return df_filtered