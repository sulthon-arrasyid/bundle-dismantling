from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QVBoxLayout, QFileDialog,
    QLineEdit, QComboBox, QTextEdit, QMessageBox
)
import pandas as pd
import threading
from datetime import datetime
import sys

class dismantlingProcessor:
    """Handles the core process of dismantling bundle products."""
    def __init__(self, ui):
        self.ui = ui  # Reference to UI for logging

    def process_bundle(self, order_file, master_file, order_sheet, master_sheet):
        """Handles the processing of bundle dismantling."""
        
        # Validation
        if not order_file or not master_file or not order_sheet or not master_sheet:
            self.ui.log_message("Validation failed: Missing fields")
            QMessageBox.critical(self.ui, "Error", "Please fill in all fields")
            return
        
        try:
            # Start process
            self.ui.log_message("Started processing...")

            # Load Excel sheets
            order_df = pd.read_excel(order_file, sheet_name=order_sheet)
            master_df = pd.read_excel(master_file, sheet_name=master_sheet)

            # Check for empty data
            if order_df.empty or master_df.empty:
                self.ui.log_message("Error: One or both sheets are empty.")
                QMessageBox.critical(self.ui, "Error", "One or both selected sheets are empty.")
                return
            
            # Merge data
            self.ui.log_message("Merge data...")

            merged_df = order_df.merge(master_df, left_on='SKU', right_on='Parent Code', how='left')
            merged_df['Child Code'] = merged_df['Child Code'].fillna(merged_df['SKU'])
            merged_df['Quantity'] = merged_df.apply(
                lambda row: row['Quantity_x'] * row['Quantity_y'] if pd.notna(row['Quantity_y']) else row['Quantity_x'], axis=1
            )

            # Select relevant columns
            result_df = merged_df[['Payment Time', 'Order Number', 'Order Status', 'Channel', 'Store Name', 'Ref No', 'Child Code', 'Quantity']]
            self.ui.log_message("Processing completed successfully...")

            # Save the results
            self.ui.save_result(result_df)
            self.ui.log_message("Saving Result...")

        except Exception as e:
            self.ui.log_message(f"Error while processing: {str(e)}")
            QMessageBox.critical(self.ui, "Error", f"Error while processing: {str(e)}")

class dismantlingUI(QWidget):
    """Handles the UI using PyQt5."""

    def __init__(self):
        super().__init__()
        self.processor = dismantlingProcessor(self)  # Create an instance of the processor
        self.init_ui()

    def init_ui(self):
        """Initialize the UI components."""
        self.setWindowTitle("Bundle Dismantling App")
        self.setGeometry(100, 100, 600, 400)

        layout = QVBoxLayout()

        # File inputs
        self.order_file_label = QLabel("Order File:")
        self.order_file_input = QLineEdit()
        self.order_file_button = QPushButton("Browse Order File")
        self.order_file_button.clicked.connect(self.load_order_file)

        self.order_sheet_label = QLabel("Order Sheet Name:")
        self.order_sheet_combo = QComboBox()

        self.master_file_label = QLabel("Master Bundle File:")
        self.master_file_input = QLineEdit()
        self.master_file_button = QPushButton("Browse Master File")
        self.master_file_button.clicked.connect(self.load_master_file)

        self.master_sheet_label = QLabel("Master Sheet Name:")
        self.master_sheet_combo = QComboBox()

        # Process button
        self.process_button = QPushButton("Process")
        self.process_button.clicked.connect(self.start_processing)

        # Log output
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)

        # Layout arrangement
        layout.addWidget(self.order_file_label)
        layout.addWidget(self.order_file_input)
        layout.addWidget(self.order_file_button)
        layout.addWidget(self.order_sheet_label)
        layout.addWidget(self.order_sheet_combo)

        layout.addWidget(self.master_file_label)
        layout.addWidget(self.master_file_input)
        layout.addWidget(self.master_file_button)
        layout.addWidget(self.master_sheet_label)
        layout.addWidget(self.master_sheet_combo)

        layout.addWidget(self.process_button)
        layout.addWidget(QLabel("Log:"))
        layout.addWidget(self.log_text)

        self.setLayout(layout)

    def load_file(self, input_widget, combo_widget):
        """Load an Excel file and populate sheet combobox."""
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx)")
        if file_path:
            input_widget.setText(file_path)
            try:
                xl = pd.ExcelFile(file_path)
                combo_widget.addItems(xl.sheet_names)
                self.log_message(f"Loaded file: {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load the file: {str(e)}")
                self.log_message(f"Error loading file: {str(e)}")
    
    def load_order_file(self):
        """Load order file."""
        self.load_file(self.order_file_input, self.order_sheet_combo)

    def load_master_file(self):
        """Load master file."""
        self.load_file(self.master_file_input, self.master_sheet_combo)

    def start_processing(self):
        """Start processing in a separate thread to keep UI responsive."""
        threading.Thread(target=self.processor.process_bundle,
                         args=(self.order_file_input.text(),
                               self.master_file_input.text(),
                               self.order_sheet_combo.currentText(),
                               self.master_sheet_combo.currentText())).start()
    
    def save_result(self, result_df):
        """Save the result DataFrame."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"Result_{timestamp}.xlsx"

        save_path, _ = QFileDialog.getSaveFileName(self, "Save Result", default_filename, "Excel Files (*.xlsx);;CSV Files (*.csv)")
        if save_path:
            try:
                if save_path.endswith(".csv"):
                    result_df.to_csv(save_path, index=False)
                else:
                    result_df.to_excel(save_path, index=False)

                QMessageBox.information(self, "Success", f"Process Completed!\nResult saved as {save_path}")
                self.log_message(f"File saved: {save_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error while saving the file: {str(e)}")
                self.log_message(f"Error while saving: {str(e)}")
    
    def log_message(self, message):
        """Log messages in the UI."""
        self.log_text.append(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = dismantlingUI()
    window.show()
    sys.exit(app.exec_())