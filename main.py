import os
import sys
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QPushButton, QHBoxLayout, QVBoxLayout,
    QFileDialog, QTableWidgetItem, QLabel, QMessageBox
)

from MultiSelectionTable import MultiSelectionTable

from ShipmentBatch import ShipmentBatch
from DataExtractor import DataExtractor
from DataChecker import DataChecker

from utilities import (
    show_info,
    show_error,
)

from errors import (
    NonConformantError,
)

# https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
pd.options.mode.copy_on_write = True

class KamKiu254(QMainWindow):
    DOCUMENT_NOT_UPLOADED = ""
    UPLOAD_CSV = "上传 CSV"

    def __init__(self):
        super().__init__()

        self.data_extractor = DataExtractor()
        self.data_checker = DataChecker()

        self.df_shipment_batch = None
        self.df_chemical_composition = None
        self.df_ageing_qrcode = None
        self.df_process_card_qrcode = None

        self.df_functional_properties = None
        self.df_function_ageing = None
        self.df_function_casting = None

        self.setDesiredElements()
        self.setDesiredSampleTypes()

        self.mid_plate_report_functional_requirements = None
        self.u_part_report_functional_requirements = None
        self.load_report_functional_requirements()

        self.df_customer_shipment_details = None
        self.df_test_commission_form = None

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

        # 上传 wtdmx 数据 
        self.functional_properties_ageing_filter_layout = QHBoxLayout()
        self.functional_properties_ageing_filter_layout.addWidget(QLabel("性能 - 硬度、电导率、拉伸（按时效批号搜索）："))
        self.function_ageing_path_label = QLabel(self.DOCUMENT_NOT_UPLOADED)
        self.functional_properties_ageing_filter_layout.addWidget(self.function_ageing_path_label)
        self.function_ageing_filter_button = QPushButton(self.UPLOAD_CSV)
        self.function_ageing_filter_button.clicked.connect(self.upload_functional_properties_ageing_filter_xlsx)
        self.functional_properties_ageing_filter_layout.addWidget(self.function_ageing_filter_button)
        self.functional_properties_ageing_filter_layout.addStretch()

        # 上传 wtdmx 数据
        self.functional_properties_furnace_filter_layout = QHBoxLayout()
        self.functional_properties_furnace_filter_layout.addWidget(QLabel("性能 - 金相（按熔铸炉号搜索）："))
        self.function_furnace_path_label = QLabel(self.DOCUMENT_NOT_UPLOADED)
        self.functional_properties_furnace_filter_layout.addWidget(self.function_furnace_path_label)
        self.function_furance_upload_button = QPushButton(self.UPLOAD_CSV)
        self.function_furance_upload_button.clicked.connect(self.upload_functional_properties_furnace_filter_xlsx)
        self.functional_properties_furnace_filter_layout.addWidget(self.function_furance_upload_button)
        self.functional_properties_furnace_filter_layout.addStretch()

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

        # 上传 检测委托单 wtd1
        self.test_commission_form_layout = QHBoxLayout()
        self.test_commission_form_layout.addWidget(QLabel("检测委托单（wtd1）："))
        self.test_commission_form_path_label = QLabel(self.DOCUMENT_NOT_UPLOADED)
        self.test_commission_form_layout.addWidget(self.test_commission_form_path_label)
        self.test_commission_form_upload_button = QPushButton(self.UPLOAD_CSV)
        self.test_commission_form_upload_button.clicked.connect(self.upload_test_commission_form_csv)
        self.test_commission_form_layout.addWidget(self.test_commission_form_upload_button)
        self.test_commission_form_layout.addStretch()
        
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

        self.generate_all_reports_button = QPushButton("生成全部报告")
        self.generate_all_reports_button.clicked.connect(self.generate_all_reports)
        self.other_functionalities_layout.addWidget(self.generate_all_reports_button)

        self.generate_customer_shipment_details_button = QPushButton("生成客户出货明细")
        self.generate_customer_shipment_details_button.clicked.connect(self.generate_customer_shipment_details)
        self.other_functionalities_layout.addWidget(self.generate_customer_shipment_details_button)

        self.main_table = MultiSelectionTable()

        layout = QVBoxLayout()
        layout.addLayout(self.shipment_batch_layout)
        layout.addLayout(self.functional_properties_ageing_filter_layout)
        layout.addLayout(self.functional_properties_furnace_filter_layout)
        layout.addLayout(self.composition_layout)
        layout.addLayout(self.ageing_qrcode_layout)
        layout.addLayout(self.process_card_qrcode_layout)
        layout.addLayout(self.test_commission_form_layout)
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
            self, "上传 发货批次表 CSV", "", "CSV Files (*.csv)"
        )

        if file_path:
            self.shipment_path_label.setText(f"{file_path}")
            try:
                self.df_shipment_batch = self.data_extractor.extract_shipment_batch_data(file_path)
                self.display_dataframe(self.df_shipment_batch)
                self.display_report_generation_buttons()
            except Exception as e:
                self.shipment_path_label.setText(f"读取文件出错: {str(e)}")

    def display_dataframe(self, df: pd.DataFrame):
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
    
    def upload_functional_properties_ageing_filter_xlsx(self):
        """
        上传 性能数据
        """
        file_path, _ = QFileDialog.getOpenFileName(
            self, "上传 性能 XLSX", "", "Excel Files (*.xlsx)"
        )

        if file_path:
            self.function_ageing_path_label.setText(f"{file_path}")
            try:
                self.df_function_ageing = self.data_extractor.extract_functional_properties_data(file_path)
                if self.df_function_casting is not None:
                    self.df_functional_properties = pd.concat([self.df_function_ageing, self.df_function_casting], axis=0)
            except Exception as e:
                self.function_ageing_path_label.setText(f"读取文件出错: {str(e)}")
    
    def upload_functional_properties_furnace_filter_xlsx(self):
        """
        上传 性能数据
        """
        file_path, _ = QFileDialog.getOpenFileName(
            self, "上传 性能 XLSX", "", "Excel Files (*.xlsx)"
        )

        if file_path:
            self.function_furnace_path_label.setText(f"{file_path}")
            try:
                self.df_function_casting = self.data_extractor.extract_functional_properties_data(file_path)
                if self.df_function_ageing is not None:
                    self.df_functional_properties = pd.concat([self.df_function_ageing, self.df_function_casting], axis=0)
            except Exception as e:
                self.function_furnace_path_label.setText(f"读取文件出错: {str(e)}")
    
    def display_report_generation_buttons(self):
        num_table_cols = self.main_table.columnCount()
        self.main_table.insertColumn(num_table_cols)
        # Set the column header
        self.main_table.setHorizontalHeaderItem(num_table_cols, QTableWidgetItem("操作"))
        
        for index, row in self.df_shipment_batch.iterrows():
            # Create a button for the third column
            button = QPushButton('生成报告')
            button.clicked.connect(lambda _, i=index, r=row: self.safe_generate_report(i, r))
            self.main_table.setCellWidget(index, num_table_cols, button)
    
    def generate_all_reports(self):
        try:
            self.check_data_uploaded(self.df_shipment_batch, "请上传发货批次表数据")
            self.check_data_uploaded(self.df_function_ageing, "请上传经过时效批号搜索的性能数据")
            self.check_data_uploaded(self.df_function_casting, "请上传经过熔铸炉号搜索的性能数据")
            self.check_data_uploaded(self.df_chemical_composition, "请上传经化学成分数据")
        except Exception as e:
            return
        
        for index, row in self.df_shipment_batch.iterrows():
            sb = ShipmentBatch(row)

            try:
                self.df_shipment_batch.at[index, '性能'] = self.data_checker.check_functional_conformance(sb, self.df_test_commission_form)

                sb.generate_report(
                    self.df_shipment_batch,
                    self.df_chemical_composition, 
                    self.df_chemical_composition_limits,
                    self.df_functional_properties,
                    self.mid_plate_report_functional_requirements,
                    self.u_part_report_functional_requirements
                )

                show_info("Report generated at :")
            # except NonConformantError as e:
            #     self.df_shipment_batch.at[index, '性能'] = e.message
            #     print(e)
            except Exception as e:
                # print errors to terminal so user don't get bombarded with popups
                print(e)
        
        self.display_dataframe(self.df_shipment_batch)
        show_info("生成完毕")
        
    def safe_generate_report(self, index, row):
        try:
            self.check_data_uploaded(self.df_function_ageing, "请上传经过时效批号搜索的性能数据")
            self.check_data_uploaded(self.df_function_casting, "请上传经过熔铸炉号搜索的性能数据")
            self.check_data_uploaded(self.df_chemical_composition, "请上传经化学成分数据")
        except Exception as e:
            print(e)
        
        sb = ShipmentBatch(row)
        output_report_path = None

        self.df_shipment_batch.at[index, '性能'] = self.data_checker.check_functional_conformance(sb, self.df_test_commission_form)
        self.display_dataframe(self.df_shipment_batch)
        self.display_report_generation_buttons()

        try:
            output_report_path = sb.generate_report(
                self.df_shipment_batch, 
                self.df_chemical_composition, 
                self.df_chemical_composition_limits,
                self.df_functional_properties,
                self.mid_plate_report_functional_requirements,
                self.u_part_report_functional_requirements
            )
        except Exception as e:
            print(e)
            show_error(str(e))
        finally:
            if output_report_path is not None:
                show_info(f"报告成功生成：{output_report_path}")
    
    def load_report_functional_requirements(self):
        try:
            self.mid_plate_report_functional_requirements = pd.read_csv('./data/点位/202507_中板.csv')
            self.u_part_report_functional_requirements = pd.read_csv('./data/点位/202507_U件.csv')
        except Exception as e:
            show_error(f"读取性能要求与点位数据出错: {str(e)}")

    def upload_chemical_composition_csv(self):
        """
        上传 化学成分
        """
        file_path, _ = QFileDialog.getOpenFileName(
            self, "上传 化学成分 CSV", "", "CSV Files (*.csv)"
        )

        if file_path:
            self.composition_path_label.setText(f"{file_path}")
            try:
                compositions = self.df_chemical_composition_limits['成分'].tolist()
                self.df_chemical_composition = self.data_extractor.extract_chemical_composition_data(file_path, compositions)
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
            self, "上传 型材时效二维码 CSV", "", "CSV Files (*.csv)"
        )

        if file_path:
            self.ageing_qrcode_path_label.setText(f"{file_path}")
            try:
                self.df_ageing_qrcode = self.data_extractor.extract_ageing_qrcode_data(file_path)
                # self.display_dataframe(self.df_ageing_qrcode)
            except Exception as e:
                self.ageing_qrcode_path_label.setText(f"读取文件出错: {str(e)}")
    
    def upload_process_card_qrcode_csv(self):
        """
        上传 流程卡二维码记录
        """
        file_path, _ = QFileDialog.getOpenFileName(
            self, "上传 流程卡二维码记录 CSV", "", "CSV Files (*.csv)"
        )

        if file_path:
            self.process_card_qrcode_path_label.setText(f"{file_path}")
            try:
                self.df_process_card_qrcode = self.data_extractor.extract_process_card_qrcode_data(file_path)
                # self.display_dataframe(self.df_process_card_qrcode)
            except Exception as e:
                self.process_card_qrcode_path_label.setText(f"读取文件出错: {str(e)}")
    
    def upload_test_commission_form_csv(self):
        """
        上传 检测委托单wtd1
        """
        file_path, _ = QFileDialog.getOpenFileName(
            self, "上传 检测委托单 CSV", "", "CSV Files (*.csv)"
        )

        if file_path:
            self.test_commission_form_path_label.setText(f"{file_path}")
            try:
                self.df_test_commission_form = self.data_extractor.extract_test_commission_form_data(file_path)
                # self.display_dataframe(self.df_test_commission_form)
            except Exception as e:
                self.test_commission_form_path_label.setText(f"读取文件出错: {str(e)}")
    
    def check_data_uploaded(self, var_to_check, message: str):
        if var_to_check is None:
            show_error(message)
            raise ValueError(message)

    def check_cpk_path(self):
        try:
            self.check_data_uploaded(self.df_shipment_batch, "请上传 发货批次表 数据")
        except Exception as e:
            print(e)
        
        self.df_shipment_batch = self.data_checker.check_cpk_path(self.df_shipment_batch)
        self.display_dataframe(self.df_shipment_batch)
        self.display_report_generation_buttons()
    
    def check_chemical_composition_conformance(self):
        """
        检查 化学成分
        """
        try:
            self.check_data_uploaded(self.df_shipment_batch, "请上传 发货批次表 数据")
            self.check_data_uploaded(self.df_chemical_composition, "请上传 化学成分 数据")
        except Exception as e:
            print(e)
        
        self.df_shipment_batch = self.data_checker.check_chemical_composition_conformance(
            self.df_shipment_batch,
            self.df_chemical_composition,
            self.df_chemical_composition_limits
        )
        
        self.display_dataframe(self.df_shipment_batch)
        self.display_report_generation_buttons()


    def extract_data_from_ageing_qrcode(self):
        """
        填入 挤压批 & 熔铸批号 二维码
        """
        try:
            self.check_data_uploaded(self.df_shipment_batch, "请上传 发货批次表 数据")
            self.check_data_uploaded(self.df_ageing_qrcode, "请上传 型材时效二维码 数据")
        except Exception as e:
            print(e)

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

        self.df_shipment_batch['挤压批次二维码'] = self.df_shipment_batch['挤压批（二维码）'].apply(lambda x: str(x).split('+')[-1])
        
        self.display_dataframe(self.df_shipment_batch)
        self.display_report_generation_buttons()
    
    def extract_ageing_batch_qrcode(self):
        """
        从 流程卡二维码记录 采取 时效批次二维码
        """
        try:
            self.check_data_uploaded(self.df_shipment_batch, "请上传 发货批次表 数据")
            self.check_data_uploaded(self.df_process_card_qrcode, "请上传 流程卡二维码记录 数据")
        except Exception as e:
            print(e)
            
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
            "选择文件夹",
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


    def generate_customer_shipment_details(self):
        try:
            self.check_data_uploaded(self.df_shipment_batch, "请上传 发货批次表 数据")
            self.check_data_uploaded(self.df_process_card_qrcode, "请上传 流程卡二维码记录 数据")

            self.df_customer_shipment_details = self.data_extractor.extract_customer_shipment_details(self.df_shipment_batch)
            self.display_dataframe(self.df_customer_shipment_details)
        except Exception as e:
            print(e)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = KamKiu254()
    window.showMaximized()
    sys.exit(app.exec())
