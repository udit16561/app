import sys
import os
import pandas as pd
import joblib
from prophet import Prophet
from PySide6.QtWidgets import (QApplication, QMainWindow, QDialog, QVBoxLayout, QHBoxLayout,
                              QLabel, QLineEdit, QPushButton, QMessageBox, QFileDialog,
                              QTabWidget, QWidget, QScrollArea, QCheckBox, 
                              QTableWidget, QTableWidgetItem, QProgressDialog, QTextEdit,
                              QGroupBox, QSizePolicy)
from PySide6.QtGui import QFont, QColor, QPalette
from PySide6.QtCore import Qt, QDate, QTime, Signal
from PyInstaller.utils.hooks import collect_all, collect_data_files ,collect_dynamic_libs



# Constants for CUF calculation
PLANT_CAPACITY = 50  # MW
HOURS_PER_DAY = 24
DAYS_PER_MONTH = 30
CUF_DENOMINATOR = PLANT_CAPACITY * HOURS_PER_DAY * DAYS_PER_MONTH

MODEL_PATH = r"D:\active power\prophet_model.joblib"

datas = collect_data_files('prophet')
binaries = collect_dynamic_libs('prophet')
hiddenimports = ['pystan', 'prophet.plot', 'prophet.diagnostics']

def resource_path(relative_path):
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS2
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class LoginDialog(QDialog):
    def __init__(self, current_time, current_user):
        super().__init__()
        self.current_time = current_time
        self.current_user = current_user
        self.setWindowTitle("Login")
        self.setGeometry(100, 100, 800, 600)
        self.setStyleSheet("background-color: #2C3E50; color: white; border-radius: 10px;")

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignCenter)

        # Title
        title = QLabel("WTG Maintenance Scheduler")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: #ECF0F1; margin-bottom: 15px;")

        # Current time display
        time_label = QLabel(f"Current Time (UTC): {self.current_time}")
        time_label.setFont(QFont("Arial", 10))
        time_label.setAlignment(Qt.AlignCenter)
        time_label.setStyleSheet("color: #ECF0F1; margin-bottom: 10px;")

        # Current user display
        user_label = QLabel(f"Logged in as: {self.current_user}")
        user_label.setFont(QFont("Arial", 10))
        user_label.setAlignment(Qt.AlignCenter)
        user_label.setStyleSheet("color: #ECF0F1; margin-bottom: 20px;")

        # Login inputs
        self.user_input = QLineEdit()
        self.user_input.setPlaceholderText("Username")
        self.user_input.setFont(QFont("Arial", 11))
        self.user_input.setFixedWidth(250)
        self.user_input.setStyleSheet("""
            QLineEdit {
                background-color: white;
                color: black;
                padding: 6px;
                border-radius: 5px;
            }
        """)

        self.pass_input = QLineEdit()
        self.pass_input.setPlaceholderText("Password")
        self.pass_input.setFont(QFont("Arial", 11))
        self.pass_input.setFixedWidth(250)
        self.pass_input.setEchoMode(QLineEdit.Password)
        self.pass_input.setStyleSheet("""
            QLineEdit {
                background-color: white;
                color: black;
                padding: 6px;
                border-radius: 5px;
            }
        """)

        # Login button
        self.login_button = QPushButton("Login")
        self.login_button.setFont(QFont("Arial", 11, QFont.Bold))
        self.login_button.setFixedWidth(120)
        self.login_button.setStyleSheet("""
            QPushButton {
                background-color: #3498DB;
                color: white;
                padding: 8px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #2980B9;
            }
        """)
        self.login_button.clicked.connect(self.check_login)

        # Copyright
        copyright_label = QLabel("© 2025 All rights reserved.")
        copyright_label.setFont(QFont("Arial", 9))
        copyright_label.setAlignment(Qt.AlignCenter)
        copyright_label.setStyleSheet("color: #ECF0F1; margin-top: 200px;")

        # Add widgets to layout
        layout.addWidget(title)
        layout.addWidget(time_label)
        layout.addWidget(user_label)
        layout.addWidget(self.user_input, alignment=Qt.AlignCenter)
        layout.addWidget(self.pass_input, alignment=Qt.AlignCenter)
        layout.addWidget(self.login_button, alignment=Qt.AlignCenter)
        layout.addWidget(copyright_label)

        self.setLayout(layout)

    def check_login(self):
        if self.user_input.text() == "admin" and self.pass_input.text() == "admin":
            self.accept()
        else:
            QMessageBox.warning(self, "Error", "Invalid Username or Password")

class DataUploadWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.selected_file = None
        self.selected_sheets = []
        self.sheets = []
        self.sheet_checkboxes = []
        self.is_csv = False
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        # Title
        title = QLabel("Upload Monthly Data")
        title.setFont(QFont("Arial", 12, QFont.Bold))
        title.setStyleSheet("color: #2C3E50; margin: 10px;")
        
        # File selection
        file_layout = QHBoxLayout()
        self.file_label = QLabel("No file selected")
        self.file_label.setStyleSheet("color: #7F8C8D;")
        
        upload_btn = QPushButton("Select File")
        upload_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498DB;
                color: white;
                padding: 8px;
                border-radius: 5px;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #2980B9;
            }
        """)
        upload_btn.clicked.connect(self.select_file)
        
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(upload_btn)

        # Sheet selection (for Excel files)
        self.sheet_label = QLabel("Select Turbine Sheets:")
        self.sheet_label.setFont(QFont("Arial", 10, QFont.Bold))
        
        # Create scroll area for checkboxes
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setStyleSheet("""
            QScrollArea {
                border: 1px solid #BDC3C7;
                border-radius: 5px;
                background-color: white;
            }
        """)
        
        self.sheet_container = QWidget()
        self.sheet_layout = QVBoxLayout(self.sheet_container)
        self.scroll.setWidget(self.sheet_container)
        self.scroll.setMaximumHeight(200)

        # Select All Checkbox
        self.select_all_cb = QCheckBox("Select All Sheets")
        self.select_all_cb.setStyleSheet("""
            QCheckBox {
                color: #2C3E50;
                font-weight: bold;
                padding: 5px;
            }
            QCheckBox::indicator {
                width: 15px;
                height: 15px;
            }
        """)
        self.select_all_cb.clicked.connect(self.toggle_all_sheets)

        # Status
        self.status_label = QLabel("Status: Waiting for file upload")
        self.status_label.setStyleSheet("color: #7F8C8D;")

        # Add all widgets to main layout
        layout.addWidget(title)
        layout.addLayout(file_layout)
        layout.addWidget(self.sheet_label)
        layout.addWidget(self.select_all_cb)
        layout.addWidget(self.scroll)
        layout.addWidget(self.status_label)
        layout.addStretch()

        self.setLayout(layout)

    def select_file(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Select Monthly Data File",
            "",
            "Data Files (*.xlsx *.xls *.csv);;Excel Files (*.xlsx *.xls);;CSV Files (*.csv);;All Files (*)"
        )
        
        if file_name:
            try:
                # Verify file exists
                if not os.path.exists(file_name):
                    raise FileNotFoundError(f"File not found: {file_name}")
                
                # Clear previous selections
                self.selected_file = None
                self.selected_sheets = []
                
                # Handle CSV files
                if file_name.lower().endswith('.csv'):
                    self.is_csv = True
                    # Test CSV reading
                    test_df = pd.read_csv(file_name)
                    self.validate_columns(test_df)
                    self.handle_csv_file(file_name)
                # Handle Excel files
                else:
                    self.is_csv = False
                    # Test Excel reading
                    test_df = pd.read_excel(file_name)
                    self.validate_columns(test_df)
                    self.handle_excel_file(file_name)
                    
                self.selected_file = file_name
                self.file_label.setText(os.path.basename(file_name))
                self.status_label.setText("Status: File loaded successfully")
                self.status_label.setStyleSheet("color: #27AE60;")

            except Exception as e:
                self.handle_file_error(e)

    def validate_columns(self, df):
        """Check for required columns in the dataframe"""
        required_cols = {'date', 'Energy', 'CUF'}
        if not required_cols.issubset(df.columns):
            missing = required_cols - set(df.columns)
            raise ValueError(f"Missing required columns: {', '.join(missing)}")
        
        # Check for wind speed column with any of these common names
        wind_speed_cols = {'wind_speed_avg', 'WindSpeed', 'wind_speed'}
        if not any(col in df.columns for col in wind_speed_cols):
            print("Warning: No wind speed column found - wind speed predictions will be skipped")

    def handle_csv_file(self, file_name):
        """Special handling for CSV files"""
        # Hide sheet selection UI elements for CSV
        self.sheet_label.hide()
        self.select_all_cb.hide()
        self.scroll.hide()
        
        # Automatically select the CSV data
        self.selected_sheets = ['csv_data']
        self.status_label.setText("Status: CSV file loaded")
        self.status_label.setStyleSheet("color: #27AE60;")

    def handle_excel_file(self, file_name):
        """Special handling for Excel files"""
        # Show sheet selection UI elements for Excel
        self.sheet_label.show()
        self.select_all_cb.show()
        self.scroll.show()
        
        excel_file = pd.ExcelFile(file_name)
        self.sheets = excel_file.sheet_names
        self.create_sheet_checkboxes()
        
        # Verify first sheet's date column
        df = excel_file.parse(self.sheets[0])
        if not self.validate_date_column(df['date']):
            raise ValueError(f"Could not parse date column in sheet: {self.sheets[0]}")

    def validate_date_column(self, date_series):
        """Try parsing dates with multiple formats"""
        try:
            pd.to_datetime(date_series)
            return True
        except:
            return False

    def handle_file_error(self, error):
        """Show detailed error message to user"""
        self.selected_file = None
        self.file_label.setText("No file selected")
        self.status_label.setText("Status: Error loading file")
        self.status_label.setStyleSheet("color: #C0392B;")
        
        error_msg = str(error)
        if "Missing required columns" in error_msg:
            error_msg += "\n\nRequired columns:\n- date\n- Energy\n- CUF"
        elif "parse date column" in error_msg:
            error_msg += "\n\nDate format should be YYYY-MM-DD or similar recognizable format"
        
        QMessageBox.critical(
            self,
            "File Error",
            f"Could not load file:\n\n{error_msg}\n\n"
            "Please check:\n"
            "1. File is not open in another program\n"
            "2. Contains all required columns\n"
            "3. Uses correct date format"
        )

    def create_sheet_checkboxes(self):
        for cb in self.sheet_checkboxes:
            self.sheet_layout.removeWidget(cb)
            cb.deleteLater()
        self.sheet_checkboxes.clear()
        
        for sheet_name in self.sheets:
            cb = QCheckBox(sheet_name)
            cb.setStyleSheet("""
                QCheckBox {
                    color: #2C3E50;
                    padding: 5px;
                }
                QCheckBox::indicator {
                    width: 15px;
                    height: 15px;
                }
            """)
            cb.stateChanged.connect(self.update_selected_sheets)
            self.sheet_checkboxes.append(cb)
            self.sheet_layout.addWidget(cb)

    def toggle_all_sheets(self, state):
        for cb in self.sheet_checkboxes:
            cb.setChecked(state)

    def update_selected_sheets(self):
        self.selected_sheets = [cb.text() for cb in self.sheet_checkboxes if cb.isChecked()]
        num_selected = len(self.selected_sheets)
        if num_selected > 0:
            self.status_label.setText(f"Status: {num_selected} sheet{'s' if num_selected > 1 else ''} selected")
            self.status_label.setStyleSheet("color: #27AE60;")
        else:
            self.status_label.setText("Status: No sheets selected")
            self.status_label.setStyleSheet("color: #E74C3C;")

    def get_selected_sheets(self):
        """Return the selected sheets or 'csv_data' for CSV files"""
        if self.is_csv:
            return ['csv_data']
        return self.selected_sheets if self.selected_sheets else []

class EnergyPredictor:
    def __init__(self):
        self.models = None
    
    def load_models(self):
        """Load pre-trained models from specified path"""
        try:
            if not os.path.exists(MODEL_PATH):
                raise FileNotFoundError(f"Model file not found at {MODEL_PATH}")
                
            self.models = joblib.load(MODEL_PATH)
            if not all(k in self.models for k in ['energy_model', 'cuf_model']):
                raise ValueError("Model file missing required components")
            return True
        except Exception as e:
            raise RuntimeError(f"Failed to load models: {str(e)}")
    
    def predict(self, periods=12, freq='M'):
        """Generate predictions with proper future dataframe"""
        if not self.models:
            raise ValueError("Models not loaded - call load_models() first")
            
        future = self.models['energy_model'].make_future_dataframe(
            periods=periods, 
            freq=freq,
            include_history=False
        )
        
        # Get predictions
        energy = self.models['energy_model'].predict(future)
        cuf = self.models['cuf_model'].predict(future)
        
        # Prepare output DataFrame
        results = pd.DataFrame({
            'Date': future['ds'].dt.strftime('%Y-%m-%d'),
            'Predicted Energy (MWh)': energy['yhat'].round(2),
            'Energy Lower Bound': energy['yhat_lower'].round(2),
            'Energy Upper Bound': energy['yhat_upper'].round(2),
            'Predicted CUF (%)': cuf['yhat'].round(4),
            'CUF Lower Bound': cuf['yhat_lower'].round(4),
            'CUF Upper Bound': cuf['yhat_upper'].round(4),
            'Derived CUF (%)': ((energy['yhat'] * 100) / CUF_DENOMINATOR).round(4)
        })
        
        # Add wind speed if model exists
        if 'wind_speed_model' in self.models:
            wind = self.models['wind_speed_model'].predict(future)
            results['Predicted Wind Speed (m/s)'] = wind['yhat'].round(2)
            results['Wind Speed Lower Bound'] = wind['yhat_lower'].round(2)
            results['Wind Speed Upper Bound'] = wind['yhat_upper'].round(2)
        
        return results

class ModelPreprocessingTab(QWidget):
    def __init__(self, result_tab):
        super().__init__()
        self.result_tab = result_tab
        self.energy_predictor = EnergyPredictor()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()
        
        # Title
        title = QLabel("Wind Turbine Data Processing")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setStyleSheet("color: #2C3E50; margin: 10px;")
        title.setAlignment(Qt.AlignCenter)
        
        # Model control group
        model_group = QGroupBox("Model Configuration")
        model_group.setStyleSheet("""
            QGroupBox {
                background-color: white;
                border: 1px solid #BDC3C7;
                border-radius: 5px;
                margin: 10px;
                padding: 15px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px;
                color: black;
            }
        """)
        
        model_layout = QVBoxLayout(model_group)
        
        # Model path display
        self.model_path_label = QLabel(f"Model Path: {MODEL_PATH}")
        self.model_path_label.setStyleSheet("color: #2C3E50;")
        self.model_status_label = QLabel("Status: Model not loaded")
        self.model_status_label.setStyleSheet("color: #E74C3C;")
        
        # Model buttons
        btn_layout = QHBoxLayout()
        
        self.load_model_btn = QPushButton("Load Model")
        self.load_model_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498DB;
                color: white;
                padding: 8px;
                border-radius: 5px;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #2980B9;
            }
        """)
        self.load_model_btn.clicked.connect(self.load_model)
        
        self.change_model_btn = QPushButton("Change Model")
        self.change_model_btn.setStyleSheet("""
            QPushButton {
                background-color: #9B59B6;
                color: white;
                padding: 8px;
                border-radius: 5px;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #8E44AD;
            }
        """)
        self.change_model_btn.clicked.connect(self.change_model)
        
        btn_layout.addWidget(self.load_model_btn)
        btn_layout.addWidget(self.change_model_btn)
        
        model_layout.addWidget(self.model_path_label)
        model_layout.addWidget(self.model_status_label)
        model_layout.addLayout(btn_layout)
        
        # Data upload widget
        self.monthly_upload = DataUploadWidget()
        
        # Process button
        self.process_button = QPushButton("Generate Forecast")
        self.process_button.setStyleSheet("""
            QPushButton {
                background-color: #2ECC71;
                color: white;
                padding: 10px;
                border-radius: 5px;
                font-weight: bold;
                margin: 20px;
            }
            QPushButton:hover {
                background-color: #27AE60;
            }
            QPushButton:disabled {
                background-color: #BDC3C7;
            }
        """)
        self.process_button.clicked.connect(self.process_data)
        self.process_button.setEnabled(False)
        
        # Add widgets to layout
        layout.addWidget(title)
        layout.addWidget(model_group)
        layout.addWidget(self.monthly_upload)
        layout.addWidget(self.process_button, alignment=Qt.AlignCenter)
        
        self.setLayout(layout)

    def load_model(self):
        try:
            if self.energy_predictor.load_models():
                self.model_status_label.setText("Status: Model loaded successfully")
                self.model_status_label.setStyleSheet("color: #27AE60;")
                self.process_button.setEnabled(True)
                QMessageBox.information(self, "Success", "Model loaded successfully!")
            else:
                self.model_status_label.setText("Status: Model loading failed")
                self.model_status_label.setStyleSheet("color: #E74C3C;")
                self.process_button.setEnabled(False)
        except Exception as e:
            self.model_status_label.setText(f"Status: Error - {str(e)}")
            self.model_status_label.setStyleSheet("color: #E74C3C;")
            self.process_button.setEnabled(False)
            QMessageBox.critical(self, "Error", f"Failed to load model:\n{str(e)}")

    def change_model(self):
        global MODEL_PATH
        new_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Model File",
            os.path.dirname("D:\active power\prophet_model.joblib"),
            "Joblib Files (*.joblib);;All Files (*)"
        )
        
        if new_path:
            MODEL_PATH = new_path
            self.model_path_label.setText(f"Model Path: {MODEL_PATH}")
            self.model_status_label.setText("Status: Model not loaded")
            self.model_status_label.setStyleSheet("color: #E74C3C;")
            self.process_button.setEnabled(False)

    def process_data(self):
        try:
            if not hasattr(self.energy_predictor, 'models') or not self.energy_predictor.models:
                raise ValueError("Please load the prediction model first")
            
            progress = QProgressDialog("Generating forecast...", "Cancel", 0, 100, self)
            progress.setWindowModality(Qt.WindowModal)
            
            progress.setLabelText("Creating forecast...")
            QApplication.processEvents()
            
            predictions = self.energy_predictor.predict()
            progress.setValue(50)
            
            progress.setLabelText("Preparing results...")
            self.result_tab.add_result("Energy Forecast", predictions)
            
            # Save to CSV
            default_path = os.path.join(os.path.dirname(MODEL_PATH), "energy_forecast.csv")
            predictions.to_csv(default_path, index=False)
            
            progress.setValue(100)
            QMessageBox.information(
                self, 
                "Success", 
                f"Forecast generated successfully!\n\nSaved to: {default_path}"
            )
            
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
        finally:
            if 'progress' in locals():
                progress.close()
    
    def validate_columns(self, df):
        """Validate that required columns exist in the dataframe"""
        required_cols = {'date', 'Energy', 'CUF'}
        if not required_cols.issubset(df.columns):
            missing = required_cols - set(df.columns)
            raise ValueError(f"Missing required columns: {', '.join(missing)}")

    def run_energy_model(self, data_dict, sheets):
        """Process energy data with Prophet with proper error handling"""
        predictions = {}
        for sheet in sheets:
            try:
                df = data_dict[sheet]
                
                if len(df) < 12:
                    QMessageBox.warning(
                        self,
                        "Insufficient Data",
                        f"Sheet '{sheet}' has only {len(df)} months of data. "
                        "At least 12 months are required for accurate forecasting."
                    )
                    continue
                
                self.energy_predictor.train_models(df)
                pred_df = self.energy_predictor.predict_energy_and_cuf()
                
                predictions[sheet] = pred_df.reset_index()
                predictions[sheet].to_csv(f'energy_cuf_predictions_{sheet}.csv', index=False)
                self.result_tab.add_result(f"Energy & CUF - {sheet}", predictions[sheet])
                
            except Exception as e:
                QMessageBox.critical(
                    self,
                    "Prediction Error",
                    f"Failed to generate predictions for sheet '{sheet}':\n\n{str(e)}"
                )
                continue
        
        return predictions

class ResultTab(QWidget):
    results_updated = Signal(dict)
    
    def __init__(self):
        super().__init__()
        self.current_results = {}
        self.initUI()

    def initUI(self):
        self.main_layout = QVBoxLayout(self)
        
        # Title
        title = QLabel("Prediction Results")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setStyleSheet("color: #2C3E50; margin: 10px;")
        title.setAlignment(Qt.AlignCenter)
        self.main_layout.addWidget(title)

        # Clear button
        clear_button = QPushButton("Clear Results")
        clear_button.setStyleSheet("""
            QPushButton {
                background-color: #E74C3C;
                color: white;
                padding: 8px;
                border-radius: 5px;
                max-width: 150px;
            }
            QPushButton:hover {
                background-color: #C0392B;
            }
        """)
        clear_button.clicked.connect(self.clear_results)
        self.main_layout.addWidget(clear_button, alignment=Qt.AlignRight)

        # Create scroll area for results
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("QScrollArea { border: none; }")
        
        # Container widget for scroll area
        self.container = QWidget()
        self.container_layout = QVBoxLayout(self.container)
        self.container_layout.setAlignment(Qt.AlignTop)
        self.container_layout.addStretch()
        
        self.scroll_area.setWidget(self.container)
        self.main_layout.addWidget(self.scroll_area)

    def add_result(self, sheet, df):
        # Energy and CUF predictions
        energy_group = QGroupBox(f"{sheet} - Energy & CUF Forecast")
        energy_group.setStyleSheet("""
            QGroupBox {
                background-color: white;
                border: 1px solid #BDC3C7;
                border-radius: 5px;
                margin: 10px;
                padding: 15px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px;
                color: black;
            }
        """)
        
        energy_layout = QVBoxLayout(energy_group)
        
        energy_columns = [col for col in df.columns if 'Energy' in col or 'CUF' in col or 'Date' in col]
        energy_df = df[energy_columns]
        
        energy_table = QTableWidget()
        energy_table.setRowCount(energy_df.shape[0])
        energy_table.setColumnCount(energy_df.shape[1])
        energy_table.setHorizontalHeaderLabels(energy_df.columns)
        
        energy_table.setStyleSheet("""
            QTableWidget {
                color: black;
            }
            QHeaderView::section {
                color: black;
                background-color: #f0f0f0;
            }
        """)
        
        for i in range(energy_df.shape[0]):
            for j in range(energy_df.shape[1]):
                item = QTableWidgetItem(str(energy_df.iat[i, j]))
                item.setFlags(item.flags() ^ Qt.ItemIsEditable)
                item.setForeground(QColor("black"))
                energy_table.setItem(i, j, item)
        
        energy_table.resizeColumnsToContents()
        energy_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        energy_layout.addWidget(energy_table)
        self.container_layout.insertWidget(self.container_layout.count() - 1, energy_group)
        
        # Wind speed predictions if available
        if any('Wind Speed' in col for col in df.columns):
            wind_group = QGroupBox(f"{sheet} - Wind Speed Forecast")
            wind_group.setStyleSheet("""
                QGroupBox {
                    background-color: white;
                    border: 1px solid #BDC3C7;
                    border-radius: 5px;
                    margin: 10px;
                    padding: 15px;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    left: 10px;
                    padding: 0 3px;
                    color: black;
                }
            """)
            
            wind_layout = QVBoxLayout(wind_group)
            
            wind_columns = [col for col in df.columns if 'Wind Speed' in col or 'Date' in col]
            wind_df = df[wind_columns]
            
            wind_table = QTableWidget()
            wind_table.setRowCount(wind_df.shape[0])
            wind_table.setColumnCount(wind_df.shape[1])
            wind_table.setHorizontalHeaderLabels(wind_df.columns)
            
            wind_table.setStyleSheet("""
                QTableWidget {
                    color: black;
                }
                QHeaderView::section {
                    color: black;
                    background-color: #f0f0f0;
                }
            """)
            
            for i in range(wind_df.shape[0]):
                for j in range(wind_df.shape[1]):
                    item = QTableWidgetItem(str(wind_df.iat[i, j]))
                    item.setFlags(item.flags() ^ Qt.ItemIsEditable)
                    item.setForeground(QColor("black"))
                    wind_table.setItem(i, j, item)
            
            wind_table.resizeColumnsToContents()
            wind_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            
            wind_layout.addWidget(wind_table)
            self.container_layout.insertWidget(self.container_layout.count() - 1, wind_group)

        self.current_results[sheet] = (energy_group, wind_group if any('Wind Speed' in col for col in df.columns) else None)
        self.results_updated.emit({sheet: df})

    def clear_results(self):
        for i in reversed(range(self.container_layout.count() - 1)):
            widget = self.container_layout.itemAt(i).widget()
            if widget is not None:
                widget.deleteLater()
        self.current_results = {}

class DocumentationTab(QWidget):
    def __init__(self, result_data=None):
        super().__init__()
        self.result_data = {}
        self.initUI()
        
    def initUI(self):
        layout = QVBoxLayout()
        
        # Title
        title = QLabel("Documentation and Results")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setStyleSheet("color: #2C3E50; margin: 10px;")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        # Documentation text
        doc_text = QLabel(
            "<h2>Wind Power Energy Forecast Documentation</h2>"
            "<p>This application predicts energy production, Capacity Utilization Factor (CUF), "
            "and wind speed for wind turbines using Prophet time series forecasting.</p>"
            "<h3>How to Use:</h3>"
            "<ol>"
            "<li>Upload monthly data file (CSV or Excel) containing date, Energy, and CUF columns</li>"
            "<li>Include WindSpeed column for wind speed predictions (optional)</li>"
            "<li>Click 'Process Data' to generate forecasts</li>"
            "<li>View results in the 'View Forecast' tab</li>"
            "<li>See summary results below</li>"
            "</ol>"
            "<h3>Results Summary:</h3>"
        )
        doc_text.setWordWrap(True)
        doc_text.setStyleSheet("color: #2C3E50; margin: 10px;")
        
        # Save button
        self.save_button = QPushButton("Save Predictions to CSV")
        self.save_button.setStyleSheet("""
            QPushButton {
                background-color: #3498DB;
                color: white;
                padding: 8px;
                border-radius: 5px;
                min-width: 150px;
            }
            QPushButton:hover {
                background-color: #2980B9;
            }
        """)
        self.save_button.clicked.connect(self.save_predictions)
        
        # Results display area with black text
        self.results_display = QTextEdit()
        self.results_display.setReadOnly(True)
        self.results_display.setStyleSheet("""
            QTextEdit {
                background-color: white;
                border: 1px solid #BDC3C7;
                border-radius: 5px;
                padding: 10px;
                color: black;
            }
        """)
        
        # Add widgets to layout
        layout.addWidget(doc_text)
        layout.addWidget(self.save_button, alignment=Qt.AlignRight)
        layout.addWidget(self.results_display)
        
        self.setLayout(layout)
    
    def update_results(self, results):
        """Update the documentation tab with results summary"""
        self.result_data = results
        if not results:
            self.results_display.setPlainText("No results available yet. Please process data first.")
            return
            
        summary_text = "<h3 style='color:black;'>Latest Forecast Results:</h3><ul style='color:black;'>"
        
        for sheet_name, df in results.items():
            summary_text += f"<li><b>{sheet_name}</b>:"
            summary_text += f"<br>First date: {df['Date'].iloc[0]}"
            summary_text += f"<br>Last date: {df['Date'].iloc[-1]}"
            summary_text += f"<br>Avg Energy: {df['Predicted Energy (MWh)'].mean():.2f} MWh"
            summary_text += f"<br>Avg CUF: {df['Predicted CUF (%)'].mean():.2f}%"
            
            if 'Predicted Wind Speed (m/s)' in df.columns:
                summary_text += f"<br>Avg Wind Speed: {df['Predicted Wind Speed (m/s)'].mean():.2f} m/s"
            
            summary_text += "</li>"
        
        summary_text += "</ul>"
        self.results_display.setHtml(summary_text)

    def save_predictions(self):
        """Save the prediction results to CSV files"""
        if not self.result_data:
            QMessageBox.warning(self, "No Data", "No prediction results available to save")
            return
            
        try:
            # Get save directory
            save_dir = QFileDialog.getExistingDirectory(
                self,
                "Select Directory to Save Predictions",
                "",
                QFileDialog.ShowDirsOnly
            )
            
            if not save_dir:
                return  # User cancelled
                
            # Save each sheet's predictions
            saved_files = []
            for sheet_name, df in self.result_data.items():
                # Clean sheet name for filename
                clean_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '_')).rstrip()
                filename = os.path.join(save_dir, f"predictions_{clean_name}.csv")
                df.to_csv(filename, index=False)
                saved_files.append(filename)
            
            # Show success message
            QMessageBox.information(
                self,
                "Save Successful",
                f"Saved {len(saved_files)} prediction files to:\n{save_dir}"
            )
            
        except Exception as e:
            QMessageBox.critical(
                self,
                "Save Error",
                f"Failed to save predictions:\n\n{str(e)}"
            )

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Wind Power Energy Forecast")
        self.setGeometry(100, 100, 1024, 768)
        self.setStyleSheet("background-color: #ECF0F1;")

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        self.result_tab = ResultTab()
        self.model_preprocessing_tab = ModelPreprocessingTab(self.result_tab)
        self.documentation_tab = DocumentationTab()

        self.tabs.addTab(self.model_preprocessing_tab, "Data Processing")
        self.tabs.addTab(self.result_tab, "View Forecast")
        self.tabs.addTab(self.documentation_tab, "Documentation")
        
        self.result_tab.results_updated.connect(self.documentation_tab.update_results)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    app.setStyle("Fusion")
    palette = app.palette()
    palette.setColor(QPalette.Window, QColor(236, 240, 241))
    palette.setColor(QPalette.WindowText, QColor(44, 62, 80))
    app.setPalette(palette)
    
    current_time = QDate.currentDate().toString(Qt.ISODate) + " " + \
                  QTime.currentTime().toString(Qt.ISODate)
    current_user = "Admin"
    
    login = LoginDialog(current_time, current_user)
    if login.exec() == QDialog.Accepted:
        window = MainWindow()
        window.show()
        sys.exit(app.exec())