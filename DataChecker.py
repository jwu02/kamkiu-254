import os
import pandas as pd

from utilities import (
    load_cpk_tolerance_map,
    find_files_with_substrings,
    show_info,
    show_error,
)

from constants import (
    MODEL_CODE_MAPPINGS,
    SampleDeliveryTestResult,
    TestGroup,
)

from ShipmentBatch import ShipmentBatch

class DataChecker:
    def __init__(self):
        self.cpk_tolerance_map = load_cpk_tolerance_map()

    def check_chemical_composition_conformance(
        self,
        df_shipment_batch: pd.DataFrame,
        df_chemical_composition: pd.DataFrame,
        df_chemical_composition_limits: pd.DataFrame
    ) -> pd.DataFrame:
        """
        检查 化学成分
        """
        
        for index, row in df_shipment_batch.iterrows():
            furnace_code = row['炉号']

            # check if corresponding furnace
            df = df_chemical_composition[
                df_chemical_composition['炉号'] == furnace_code
            ]

            if len(df)>0:
                first_row = df.iloc[0]
                
                for index2, row2 in df_chemical_composition_limits.iterrows():
                    element = row2['成分']
                    upper_limit = row2['上限']
                    lower_limit = row2['下限']

                    value = float(first_row[element])
                    
                    if not (lower_limit <= value <= upper_limit):
                        df_shipment_batch.at[index, '成分'] = "🔴 不合格"
                        break
                    df_shipment_batch.at[index, '成分'] = "🟢 合格"
            else:
                df_shipment_batch.at[index, '成分'] = "🟠 找不到炉号"
    
        return df_shipment_batch
    
    def check_cpk_path(self, df_shipment_batch: pd.DataFrame) -> pd.DataFrame:
        error_path = []

        for index, row in df_shipment_batch.iterrows():
            model_code = row['型号']
            extrusion_batch = str(row['挤压批号']).strip()
            path = MODEL_CODE_MAPPINGS[model_code]['cpk']['path']

            if not path or not os.path.isdir(path):
                df_shipment_batch.at[index, 'CPK'] = "🔴 错误"
                if path not in error_path:
                    show_error(f"{model_code} 型号的路径找不到：${path}")
                    error_path.append(path)
                return

            # Check if any file contains the extrusion batch string
            matching_files = find_files_with_substrings(path, [extrusion_batch])
            file_count = len(matching_files)

            if file_count == 0:
                df_shipment_batch.at[index, 'CPK'] = "🟠 不存在"
            else:
                # check CPK conformance
                if file_count > 1:
                    df_shipment_batch.at[index, 'CPK'] = "🟠 多数CPK存在"
                else:
                    # df_shipment_batch.at[index, 'CPK'] = "🟢 存在"

                    file_path = os.path.join(path, matching_files[0])
                    df_shipment_batch.at[index, 'CPK'] = self.check_cpk_conformance(file_path, self.cpk_tolerance_map[model_code])
        
        return df_shipment_batch


    def check_cpk_conformance(self, cpk_path, df_cpk_tolerance) -> str:
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
    
    def check_functional_conformance(self, shipment_batch: ShipmentBatch, df_test_commission_form: pd.DataFrame) -> str:
        """
        Check mechanical function conformance of a shipment batch entry using sample test results data exported from wtd1 
        """
        for tg in TestGroup:
            condition = (
                (df_test_commission_form['型号'] == shipment_batch.model_code) &
                (df_test_commission_form['检测项目'] == tg.value)
            )

            if tg == TestGroup.METALLOGRAPHIC_STRUCTURE:
                condition &= (df_test_commission_form['铝棒炉号'] == shipment_batch.casting_furnace_code)
            else:
                condition &= (df_test_commission_form['时效炉号'] == shipment_batch.ageing_batch_code)
            
            df_filtered = df_test_commission_form[condition]

            if len(df_filtered)==0:
                print(df_filtered)
                return f"🟠 {tg.value} 无送样记录"
            
            if not df_filtered["检验结果"].eq('Y-合格').any():
                if not df_filtered["检验结果"].eq('N-不合格').any():
                    return f"🟠 {tg.value} 送样结果未出"
                else:
                    test_commission_form_code = df_filtered.iloc[0]['委托单号']
                    return f"🔴 {tg.value}NG {test_commission_form_code}"
        
        return "🟢 合格"