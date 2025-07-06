import os
import sys
import pandas as pd
from openpyxl import load_workbook
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QPushButton, QHBoxLayout, QVBoxLayout,
    QFileDialog, QTableWidgetItem, QLabel, QMessageBox
)

from MultiSelectionTable import MultiSelectionTable

from utilities import (
    load_cpk_tolerance_map,
    check_cpk_conformance,
    MODEL_CODE_MAPPINGS,
    find_files_with_substring,
    get_report_name,
    format_date,
    generate_random_weights,
    get_batch_quantity_by_furnace_code,
    extract_shipment_batch_data,
    extract_chemical_composition_data,
    extract_process_card_qrcode_data,
    extract_ageing_qrcode_data,
    show_info,
    show_error,
)

# https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
pd.options.mode.copy_on_write = True

class KamKiu254(QMainWindow):
    DOCUMENT_NOT_UPLOADED = ""
    UPLOAD_CSV = "上传 CSV"

    def __init__(self):
        super().__init__()

        self.df_shipment_batch = None
        self.df_chemical_composition = None
        self.df_ageing_qrcode = None
        self.df_process_card_qrcode = None

        self.cpk_tolerance_map = load_cpk_tolerance_map()
        self.setDesiredElements()
        self.setDesiredSampleTypes()

        self.init_ui()
        

    def init_ui(self):
        """
        创造界面
        """
        self.setWindowTitle("254 发货报告主动生成")
        # self.resize(800, 600)

        # 上传 发货批次表
        self.shipment_batch_layout = QHBoxLayout()
        self.shipment_batch_layout.addWidget(QLabel("发货批次表："))
        self.shipment_path_label = QLabel(self.DOCUMENT_NOT_UPLOADED)
        self.shipment_batch_layout.addWidget(self.shipment_path_label)
        self.shipment_upload_button = QPushButton(self.UPLOAD_CSV)
        self.shipment_upload_button.clicked.connect(self.upload_shipment_batch_csv)
        self.shipment_batch_layout.addWidget(self.shipment_upload_button)
        self.shipment_batch_layout.addStretch()

        # 上传 化学成分
        self.composition_layout = QHBoxLayout()
        self.composition_layout.addWidget(QLabel("化学成分："))
        self.composition_path_label = QLabel(self.DOCUMENT_NOT_UPLOADED)
        self.composition_layout.addWidget(self.composition_path_label)
        self.composition_upload_button = QPushButton(self.UPLOAD_CSV)
        self.composition_upload_button.clicked.connect(self.upload_chemical_composition_csv)
        self.composition_layout.addWidget(self.composition_upload_button)
        self.composition_layout.addStretch()

        # 上传 型材时效二维码
        self.ageing_qrcode_layout = QHBoxLayout()
        self.ageing_qrcode_layout.addWidget(QLabel("型材时效二维码："))
        self.ageing_qrcode_path_label = QLabel(self.DOCUMENT_NOT_UPLOADED)
        self.ageing_qrcode_layout.addWidget(self.ageing_qrcode_path_label)
        self.ageing_qrcode_upload_button = QPushButton(self.UPLOAD_CSV)
        self.ageing_qrcode_upload_button.clicked.connect(self.upload_ageing_qrcode_csv)
        self.ageing_qrcode_layout.addWidget(self.ageing_qrcode_upload_button)
        self.ageing_qrcode_layout.addStretch()

        # 上传 流程卡二维码记录
        self.process_card_qrcode_layout = QHBoxLayout()
        self.process_card_qrcode_layout.addWidget(QLabel("流程卡二维码记录："))
        self.process_card_qrcode_path_label = QLabel(self.DOCUMENT_NOT_UPLOADED)
        self.process_card_qrcode_layout.addWidget(self.process_card_qrcode_path_label)
        self.process_card_qrcode_upload_button = QPushButton(self.UPLOAD_CSV)
        self.process_card_qrcode_upload_button.clicked.connect(self.upload_process_card_qrcode_csv)
        self.process_card_qrcode_layout.addWidget(self.process_card_qrcode_upload_button)
        self.process_card_qrcode_layout.addStretch()

        # 其他功能
        self.other_functionalities_layout = QHBoxLayout()
        # 检查CPk
        self.check_cpk_button = QPushButton("检查CPK")
        self.check_cpk_button.clicked.connect(self.check_cpk_path)
        self.other_functionalities_layout.addWidget(self.check_cpk_button)
        # 检查化学成分
        self.check_chemical_compositions_button = QPushButton("检查化学成分")
        self.check_chemical_compositions_button.clicked.connect(self.check_chemical_composition_conformance)
        self.other_functionalities_layout.addWidget(self.check_chemical_compositions_button)
        # 采取 挤压批此（二维码） & 熔铸批号
        self.fill_extrusion_batch_button = QPushButton("采取 挤压批此（二维码） & 熔铸批号")
        self.fill_extrusion_batch_button.clicked.connect(self.extract_data_from_ageing_qrcode)
        self.other_functionalities_layout.addWidget(self.fill_extrusion_batch_button)
        # 采取 时效批次（二维码）
        self.fill_ageing_batch_button = QPushButton("采取 时效批次（二维码）")
        self.fill_ageing_batch_button.clicked.connect(self.extract_ageing_batch_qrcode)
        self.other_functionalities_layout.addWidget(self.fill_ageing_batch_button)

        # 对下报告数量
        self.check_batch_quantity_button = QPushButton("对数量")
        self.check_batch_quantity_button.clicked.connect(self.check_batch_quantity)
        self.other_functionalities_layout.addWidget(self.check_batch_quantity_button)

        self.main_table = MultiSelectionTable()

        layout = QVBoxLayout()
        layout.addLayout(self.shipment_batch_layout)
        layout.addLayout(self.composition_layout)
        layout.addLayout(self.ageing_qrcode_layout)
        layout.addLayout(self.process_card_qrcode_layout)
        layout.addLayout(self.other_functionalities_layout)
        layout.addWidget(self.main_table)

        # Main widget and layout
        main_widget = QWidget()
        main_widget.setLayout(layout)
        self.setCentralWidget(main_widget)


    def upload_shipment_batch_csv(self):
        """
        上传 发货批次表
        """
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open CSV File", "", "CSV Files (*.csv)"
        )

        if file_path:
            self.shipment_path_label.setText(f"{file_path}")
            try:
                df = pd.read_csv(file_path)
                self.df_shipment_batch = extract_shipment_batch_data(df)

                self.display_dataframe(self.df_shipment_batch)
                self.display_report_generation_buttons()
                
            except Exception as e:
                self.shipment_path_label.setText(f"读取文件出错: {str(e)}")

    def display_dataframe(self, df):
        """
        显示 数据框架
        """
        self.main_table.setRowCount(len(df.index))
        self.main_table.setColumnCount(len(df.columns))
        self.main_table.setHorizontalHeaderLabels(df.columns.astype(str).tolist())

        for row in range(len(df.index)):
            for col in range(len(df.columns)):
                item = QTableWidgetItem(str(df.iat[row, col]))
                self.main_table.setItem(row, col, item)
        
        self.main_table.resizeColumnsToContents()
    
    def display_report_generation_buttons(self):
        num_table_cols = self.main_table.columnCount()
        self.main_table.insertColumn(num_table_cols)
        # Set the column header
        self.main_table.setHorizontalHeaderItem(num_table_cols, QTableWidgetItem("操作"))
        
        for index, row in self.df_shipment_batch.iterrows():
            # Create a button for the third column
            button = QPushButton('生成报告')
            button.clicked.connect(lambda _, i=index, r=row: self.generate_report(i, r))
            self.main_table.setCellWidget(index, num_table_cols, button)
    
    def generate_report(self, index, row):
        """
        报告模板生成函数
        填：型号、出货日期、图号、炉号、批量（出货数）、客户料号
        - CPK：查看对应的型号的CPK路径，再找对应挤压批号的CPK，复制数据过去模板
        - 
        """
        model_code = row['型号']
        extrusion_batch_code = row['挤压批号']
        furnace_code = row['炉号']
        batch_quantity = get_batch_quantity_by_furnace_code(self.df_shipment_batch, row)

        cpk_path_str = MODEL_CODE_MAPPINGS[model_code]['cpk']['path']
        num_rows_to_extract = MODEL_CODE_MAPPINGS[model_code]['cpk']['num_rows']

        matching_files = find_files_with_substring(cpk_path_str, extrusion_batch_code)

        file_count = len(matching_files)
        if file_count > 0:
            cpk_path = os.path.join(cpk_path_str, matching_files[0])

            # Read data from existing CPK datasheet
            df_cpk_datasheet = pd.read_excel(
                cpk_path,
                engine='openpyxl',
                header=None,  # No headers (since we're reading raw cells)
                usecols="AH:AT",  # Columns from AH to AT
                skiprows=10,  # Skip first 10 rows (to start at row 11)
                nrows=num_rows_to_extract,  # Read x rows
            )

            # Define template and output paths
            template_file = f"./报告模板/{model_code}.xlsx"  # Original template
            output_name = get_report_name(
                model_code,
                MODEL_CODE_MAPPINGS[model_code]['customer_part_code'],
                batch_quantity,
                row['地区'],
                furnace_code,
                extrusion_batch_code
            )
            output_dir = './output' # Where to save reports
            output_file = os.path.join(output_dir, f"{output_name}.xlsx")
            
            # Check if output directory exists, create if not
            os.makedirs(output_dir, exist_ok=True)
            
            # # Check if the DataFrame fits in the target range
            # if df.shape[0] > 64 or df.shape[1] > 11:
            #     show_error("Extracted data is too large for the target range")
            #     return
            
            # Load the template workbook
            wb = load_workbook(template_file)
            ws = wb.active

            # Fill in basic report information
            ws.cell(row=3, column= 14, value=model_code) # 填 型号
            ws.cell(row=4, column=3, value=format_date(row['发货日期'])) # 填 发货日期
            ws.cell(row=4, column=7, value=MODEL_CODE_MAPPINGS[model_code]['schema_code']) # 填 图号
            ws.cell(row=4, column=11, value=row['炉号']) # 填 炉号
            ws.cell(row=4, column=14, value=batch_quantity)# 填 发货数
            ws.cell(row=4, column=19, value=MODEL_CODE_MAPPINGS[model_code]['customer_part_code']) # 填 客户料号
            
            # Write CPK data to report template
            for r_idx, row_data in enumerate(df_cpk_datasheet.values, start=12):  # Start at row 12
                for c_idx, value in enumerate(row_data, start=9):  # Start at column I (9)
                    ws.cell(row=r_idx, column=c_idx, value=value)
            
            # Write chemical composition data to report template
            if self.df_chemical_composition is None:
                show_error("请上传 化学成分 数据")
                return
            else:
                compositions = self.df_chemical_composition_limits['成分'].tolist()
                df_chemical_composition_filtered = self.df_chemical_composition[self.df_chemical_composition['炉号'] == furnace_code]
                if df_chemical_composition_filtered.empty:
                    show_error(f"找不到 对应 炉号 {furnace_code} 的数据")
                else:

                    composition_data = df_chemical_composition_filtered.iloc[0].reindex(compositions)
                    for r_idx, value in enumerate(composition_data, start=MODEL_CODE_MAPPINGS[model_code]['composition']['start_row']):
                        ws.cell(row=r_idx, column=MODEL_CODE_MAPPINGS[model_code]['composition']['start_column'], value=round(float(value), 4))

            weights = generate_random_weights(model_code)
            # Weight cell starting values / coordinates
            weight_starting_row = MODEL_CODE_MAPPINGS[model_code]['weight']['starting_row']
            weight_starting_column = MODEL_CODE_MAPPINGS[model_code]['weight']['starting_column']
            for c_index, weight_val in enumerate(weights, start=weight_starting_column):
                ws.cell(row=weight_starting_row, column=c_index, value=weight_val)

            wb.save(output_file) # Save as new file

            show_info(f"报告成功生成：{output_file}")
        else:
            show_error("对应的 CPK 不存在")
    
    
    def upload_chemical_composition_csv(self):
        """
        上传 化学成分
        """
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open CSV File", "", "CSV Files (*.csv)"
        )

        if file_path:
            self.composition_path_label.setText(f"{file_path}")
            try:
                df = pd.read_csv(file_path)
                compositions = self.df_chemical_composition_limits['成分'].tolist()
                self.df_chemical_composition = extract_chemical_composition_data(df, compositions)

                # self.display_dataframe(self.df_chemical_composition)
                
            except Exception as e:
                self.composition_path_label.setText(f"读取文件出错: {str(e)}")
    
    def setDesiredElements(self):
        try:
            self.df_chemical_composition_limits = pd.read_csv("./data/成分_元素条件.csv")
            self.df_chemical_composition_limits = self.df_chemical_composition_limits.astype({ 
                '上限': float, 
                '下限': float 
            })
        except Exception as e:
            show_error(f"读取文件出错: {str(e)}")
    
    def setDesiredSampleTypes(self):
        try:
            self.desired_sample_types = pd.read_csv("./data/成分_类型条件.csv")
        except Exception as e:
            show_error(f"读取文件出错: {str(e)}")

    def upload_ageing_qrcode_csv(self):
        """
        上传 型材时效二维码
        """
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open CSV File", "", "CSV Files (*.csv)"
        )

        if file_path:
            self.ageing_qrcode_path_label.setText(f"{file_path}")
            try:
                df = pd.read_csv(file_path)
                self.df_ageing_qrcode = extract_ageing_qrcode_data(df)

                # self.display_dataframe(self.df_ageing_qrcode)
                
            except Exception as e:
                self.ageing_qrcode_path_label.setText(f"读取文件出错: {str(e)}")
    
    def upload_process_card_qrcode_csv(self):
        """
        上传 流程卡二维码记录
        """
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open CSV File", "", "CSV Files (*.csv)"
        )

        if file_path:
            self.process_card_qrcode_path_label.setText(f"{file_path}")
            try:
                df = pd.read_csv(file_path)
                self.df_process_card_qrcode = extract_process_card_qrcode_data(df)

                # self.display_dataframe(self.df_process_card_qrcode)
                
            except Exception as e:
                self.process_card_qrcode_path_label.setText(f"读取文件出错: {str(e)}")

    def check_single_record_cpk(self, index, row):
        error_path = []

        model_code = row['型号']
        extrusion_batch = str(row['挤压批号']).strip()
        path = MODEL_CODE_MAPPINGS[model_code]['cpk']['path']

        if not path or not os.path.isdir(path):
            self.df_shipment_batch.at[index, 'CPK'] = "🔴 错误"
            if path not in error_path:
                show_error(f"{model_code} 型号的路径找不到：${path}")
                error_path.append(path)
            return

        # Check if any file contains the extrusion batch string
        matching_files = find_files_with_substring(path, extrusion_batch)
        file_count = len(matching_files)

        if file_count == 0:
            self.df_shipment_batch.at[index, 'CPK'] = "🟠 不存在"
        else:
            # check CPK conformance
            if file_count > 1:
                self.df_shipment_batch.at[index, 'CPK'] = "🟠 多数CPK存在"
            else:
                # self.df_shipment_batch.at[index, 'CPK'] = "🟢 存在"

                file_path = os.path.join(path, matching_files[0])
                self.df_shipment_batch.at[index, 'CPK'] = check_cpk_conformance(file_path, self.cpk_tolerance_map[model_code])
            
    def check_cpk_path(self):
        if self.df_shipment_batch is None:
            show_error("请上传 发货批次表 数据")
            return

        for index, row in self.df_shipment_batch.iterrows():
            self.check_single_record_cpk(index, row)

        self.display_dataframe(self.df_shipment_batch)
        self.display_report_generation_buttons()
    
    def check_chemical_composition_conformance(self):
        """
        检查 化学成分
        """
        if self.df_shipment_batch is None:
            show_error("请上传 发货批次表 数据")
            return 
        if self.df_chemical_composition is None:
            show_error("请上传 化学成分 数据")
            return
        
        for index, row in self.df_shipment_batch.iterrows():
            furnace_code = row['炉号']

            # check if corresponding furnace
            df = self.df_chemical_composition[
                self.df_chemical_composition['炉号'] == furnace_code
            ]

            if len(df)>0:
                first_row = df.iloc[0]
                
                for index2, row2 in self.df_chemical_composition_limits.iterrows():
                    element = row2['成分']
                    upper_limit = row2['上限']
                    lower_limit = row2['下限']

                    value = float(first_row[element])
                    
                    if not (lower_limit <= value <= upper_limit):
                        self.df_shipment_batch.at[index, '成分'] = "🔴 不合格"
                        break
                    self.df_shipment_batch.at[index, '成分'] = "🟢 合格"
            else:
                self.df_shipment_batch.at[index, '成分'] = "🟠 找不到炉号"
        
        self.display_dataframe(self.df_shipment_batch)
        self.display_report_generation_buttons()


    def extract_data_from_ageing_qrcode(self):
        """
        填入 挤压批 & 熔铸批号 二维码
        """
        if self.df_shipment_batch is None:
            show_error("请上传 发货批次表 数据")
            return
        if self.df_ageing_qrcode is None:
            show_error("请上传 型材时效二维码 数据")
            return

        for index, row in self.df_shipment_batch.iterrows():
            model_code = row['型号']
            extrusion_batch_code = row['挤压批号']
            furnace_code = row['炉号']

            df = self.df_ageing_qrcode[
                (self.df_ageing_qrcode['型号'] == model_code)
                & (self.df_ageing_qrcode['生产挤压批'] == extrusion_batch_code)
                & (self.df_ageing_qrcode['铝棒炉号'] == furnace_code)
            ]

            self.df_shipment_batch.at[index, '挤压批（二维码）'] = df.iloc[0]['挤压批'] if len(df)>0 else "🟠 没记录"
            self.df_shipment_batch.at[index, '熔铸批号'] = df.iloc[0]['熔铸批号'][2:] if len(df)>0 else "🟠 没记录"
        
        self.display_dataframe(self.df_shipment_batch)
        self.display_report_generation_buttons()
    
    def extract_ageing_batch_qrcode(self):
        """
        从 流程卡二维码记录 采取 时效批次二维码
        """
        if self.df_shipment_batch is None:
            show_error("请上传 发货批次表 数据")
            return
        if self.df_process_card_qrcode is None:
            show_error("请上传 流程卡二维码记录 数据")
            return
            
        for index, row in self.df_shipment_batch.iterrows():
            model_code = row['型号']
            extrusion_batch_code = row['挤压批号']
            furnace_code = row['炉号']
            ageing_code = row['时效批号']

            df = self.df_process_card_qrcode[
                (self.df_process_card_qrcode['型号'] == model_code)
                & (self.df_process_card_qrcode['挤压批号'] == extrusion_batch_code)
                & (self.df_process_card_qrcode['炉号'] == furnace_code)
                & (self.df_process_card_qrcode['时效批'] == ageing_code)
            ]

            self.df_shipment_batch.at[index, '时效批次（二维码）'] = df.iloc[0]['二维码'][-4:] if len(df)>0 else "🟠 没记录"
        
        self.display_dataframe(self.df_shipment_batch)
        self.display_report_generation_buttons()

    def check_batch_quantity(self):
        """
        报选定文件夹各个型号的发货数量
        """
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select Directory",
            os.path.expanduser("~"),
            # QFileDialog.ShowDirsOnly
        )
        if not folder:
            return

        # List files (non-recursive)
        files = [
            fname for fname in os.listdir(folder)
            if os.path.isfile(os.path.join(folder, fname)) and 'MANCHESTER' in fname
        ]

        batch_quantities = {}
        for s in files:
            _, model_code, _, quantity, _, _, _ = s.split()
            if model_code not in batch_quantities:
                batch_quantities[model_code] = int(quantity)
            else:
                batch_quantities[model_code] += int(quantity)
        
        show_info('\n'.join(f"{k}: {v}" for k, v in batch_quantities.items()))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = KamKiu254()
    window.showMaximized()
    sys.exit(app.exec())
