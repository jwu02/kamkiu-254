import os
import sys
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QPushButton, QHBoxLayout, QVBoxLayout,
    QFileDialog, QTableWidgetItem, QLabel, QMessageBox
)

from MultiSelectionTable import MultiSelectionTable

from ShipmentBatch import ShipmentBatch
from DataRequester import DataRequester
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

    def __init__(self):
        super().__init__()

        self.data_requester = DataRequester()
        self.data_extractor = DataExtractor()
        self.data_checker = DataChecker()

        self.df_shipment_batch = None
        self.df_chemical_composition = None
        self.df_mechanical_properties = None

        self.setDesiredElements()

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
        self.shipment_upload_button = QPushButton("读取 发货批次表")
        self.shipment_upload_button.clicked.connect(self.display_shipment_batch_data_full)
        self.shipment_batch_layout.addWidget(self.shipment_upload_button)
        self.shipment_batch_layout.addStretch()

        # 上传 wtdmx 数据 
        self.mechanical_properties_layout = QHBoxLayout()
        self.mechanical_properties_layout.addWidget(QLabel("性能："))
        self.mechanical_properties_path_label = QLabel(self.DOCUMENT_NOT_UPLOADED)
        self.mechanical_properties_layout.addWidget(self.mechanical_properties_path_label)
        self.mechanical_properties_upload_button = QPushButton("读取 机械性能")
        self.mechanical_properties_upload_button.clicked.connect(self.request_mechanical_properties_data)
        self.mechanical_properties_layout.addWidget(self.mechanical_properties_upload_button)
        self.mechanical_properties_layout.addStretch()

        # 上传 化学成分
        self.composition_layout = QHBoxLayout()
        self.composition_layout.addWidget(QLabel("化学成分："))
        self.composition_path_label = QLabel(self.DOCUMENT_NOT_UPLOADED)
        self.composition_layout.addWidget(self.composition_path_label)
        self.composition_upload_button = QPushButton("读取 化学成分")
        self.composition_upload_button.clicked.connect(self.request_chemical_composition_data)
        self.composition_layout.addWidget(self.composition_upload_button)
        self.composition_layout.addStretch()

        # # 上传 检测委托单 wtd1
        # self.test_commission_form_layout = QHBoxLayout()
        # self.test_commission_form_layout.addWidget(QLabel("检测委托单（wtd1）："))
        # self.test_commission_form_path_label = QLabel(self.DOCUMENT_NOT_UPLOADED)
        # self.test_commission_form_layout.addWidget(self.test_commission_form_path_label)
        # self.test_commission_form_upload_button = QPushButton(self.UPLOAD_CSV)
        # self.test_commission_form_upload_button.clicked.connect(self.request_test_commission_form_data)
        # self.test_commission_form_layout.addWidget(self.test_commission_form_upload_button)
        # self.test_commission_form_layout.addStretch()
        
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
        layout.addLayout(self.mechanical_properties_layout)
        layout.addLayout(self.composition_layout)
        # layout.addLayout(self.test_commission_form_layout)
        layout.addLayout(self.other_functionalities_layout)
        layout.addWidget(self.main_table)

        # Main widget and layout
        main_widget = QWidget()
        main_widget.setLayout(layout)
        self.setCentralWidget(main_widget)

    def display_shipment_batch_data_full(self):
        self.request_shipment_batch_data()

        model_code_list = self.df_shipment_batch['型号'].tolist()
        extrusion_batch_code_list = self.df_shipment_batch['挤压批号'].tolist()

        response_data = self.data_requester.request_ageing_qrcode(model_code_list, extrusion_batch_code_list)
        df_ageing_qrcode = self.data_extractor.extract_ageing_qrcode_data(response_data)
        self.df_shipment_batch = self.data_extractor.fill_data_from_ageing_qrcode(self.df_shipment_batch, df_ageing_qrcode)

        response_data = self.data_requester.request_process_card_qrcode(model_code_list, extrusion_batch_code_list)
        df_process_card_qrcode = self.data_extractor.extract_process_card_qrcode_data(response_data)
        self.df_shipment_batch = self.data_extractor.fill_data_from_process_card_qrcode(self.df_shipment_batch, df_process_card_qrcode)

        self.display_dataframe(self.df_shipment_batch)
        self.display_report_generation_buttons()

    def request_shipment_batch_data(self):
        """
        读取 发货批次表 数据
        """
        try:
            response_data = self.data_requester.request_shipment_details()
            self.df_shipment_batch = self.data_extractor.extract_shipment_batch_data(response_data)

            # self.display_dataframe(self.df_shipment_batch)
        except Exception as e:
            self.shipment_path_label.setText(f"读取数据错误: {str(e)}")

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
    
    def request_mechanical_properties_data(self):
        """
        读取 机械性能 数据
        """
        try:
            model_code_list = self.df_shipment_batch['型号'].tolist()
            ageing_furnace_code_list = self.df_shipment_batch['时效批号'].tolist()
            smelting_furnace_code_list = self.df_shipment_batch['炉号'].tolist()

            response_data = self.data_requester.request_mechanical_properties(
                model_code_list=model_code_list, 
                ageing_furnace_code_list=ageing_furnace_code_list
            )
            df_mechanical_properties_ageing = self.data_extractor.extract_mechanical_properties_data(response_data)

            response_data = self.data_requester.request_mechanical_properties(
                model_code_list=model_code_list, 
                smelting_furnace_code_list=smelting_furnace_code_list
            )
            df_mechanical_properties_smelting = self.data_extractor.extract_mechanical_properties_data(response_data)

            self.df_mechanical_properties = pd.concat([df_mechanical_properties_ageing, df_mechanical_properties_smelting], axis=0)

        except Exception as e:
            self.mechanical_properties_path_label.setText(f"读取文件出错: {str(e)}")
    
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
        for index, row in self.df_shipment_batch.iterrows():
            sb = ShipmentBatch(row)

            try:
                self.df_shipment_batch.at[index, '性能'] = self.data_checker.check_functional_conformance(sb, self.df_test_commission_form)

                sb.generate_report(
                    self.df_shipment_batch,
                    self.df_chemical_composition, 
                    self.df_chemical_composition_limits,
                    self.df_mechanical_properties,
                    self.mid_plate_report_functional_requirements,
                    self.u_part_report_functional_requirements
                )

                # show_info("Report generated at :")
            # except NonConformantError as e:
            #     self.df_shipment_batch.at[index, '性能'] = e.message
            #     print(e)
            except Exception as e:
                # print errors to terminal so user don't get bombarded with popups
                print(e)
        
        self.display_dataframe(self.df_shipment_batch)
        show_info("生成完毕")
        
    def safe_generate_report(self, index, row):
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
                self.df_mechanical_properties,
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

    def request_chemical_composition_data(self):
        """
        读取 化学成分 数据
        """
        try:
            compositions = self.df_chemical_composition_limits['成分'].tolist()
            smelt_lot_list = self.df_shipment_batch['炉号'].tolist()
            response_data = self.data_requester.request_chemical_composition(smelt_lot_list)
            self.df_chemical_composition = self.data_extractor.extract_chemical_composition_data(response_data, compositions)

            self.display_dataframe(self.df_chemical_composition)
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

    def check_cpk_path(self):
        self.df_shipment_batch = self.data_checker.check_cpk_path(self.df_shipment_batch)

        self.display_dataframe(self.df_shipment_batch)
        self.display_report_generation_buttons()
    
    def check_chemical_composition_conformance(self):
        """
        检查 化学成分
        """
        if self.df_shipment_batch is None:
            self.display_shipment_batch_data_full()
        if self.df_chemical_composition is None:
            self.request_chemical_composition_data()
        
        self.df_shipment_batch = self.data_checker.check_chemical_composition_conformance(
            self.df_shipment_batch,
            self.df_chemical_composition,
            self.df_chemical_composition_limits
        )
        
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
            self.df_customer_shipment_details = self.data_extractor.extract_customer_shipment_details(self.df_shipment_batch)
            self.display_dataframe(self.df_customer_shipment_details)
        except Exception as e:
            print(e)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = KamKiu254()
    window.showMaximized()
    sys.exit(app.exec())
