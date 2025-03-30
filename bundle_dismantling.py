import tkinter as tk
import pandas as pd
import threading
from tkinter import filedialog, messagebox, ttk
from datetime import datetime

class BundleBreakdownApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Bundle Breakdown App")

        # Labels
        ttk.Label(root, text="Order File:").grid(row=0, column=0, padx=10, pady=10)
        ttk.Label(root, text="Order Sheet Name:").grid(row=1, column=0, padx=10, pady=10)
        ttk.Label(root, text="Master Bundle File:").grid(row=2, column=0, padx=10, pady=10)
        ttk.Label(root, text="Master Sheet Name:").grid(row=3, column=0, padx=10, pady=10)
        self.log_text = tk.Text(root, height=10, width=75, state='disabled')

        # Entries
        self.order_file_entry = ttk.Entry(root, state="readonly", width=65)
        self.master_file_entry = ttk.Entry(root, state="readonly", width=65)

        self.order_file_entry.grid(row=0, column=1, padx=10, pady=10)
        self.master_file_entry.grid(row=2, column=1, padx=10, pady=10)

        # ComboBoxes for sheet names
        self.order_sheet_combo = ttk.Combobox(root, state="readonly", width=62)
        self.master_sheet_combo = ttk.Combobox(root, state="readonly", width=62)

        self.order_sheet_combo.grid(row=1, column=1, padx=10, pady=10)
        self.master_sheet_combo.grid(row=3, column=1, padx=10, pady=10)

        # Buttons
        ttk.Button(root, text="Browse Order File", command=self.load_order_file).grid(row=0, column=2, padx=10, pady=10)
        ttk.Button(root, text="Browse Master File", command=self.load_master_file).grid(row=2, column=2, padx=10, pady=10)
        ttk.Button(root, text="Process", command=self.start_threaded_processing).grid(row=4, column=1, padx=10, pady=20)
        ttk.Button(root, text="Master Template", command=self.download_template_master).grid(row=7, column=0, padx=10, pady=10)
        ttk.Button(root, text="Order Template", command=self.download_template_order).grid(row=7, column=1, padx=10, pady=10)
        
        # Logging area
        ttk.Label(root, text="Log:").grid(row=5, column=0, padx=10, pady=10)
        self.log_text.grid(row=6, column=0, columnspan=3, padx=10, pady=10)

    def load_file(self, entry_widget, combo_widget):
        """Helper function to load an Excel file and populate the sheet combobox."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            entry_widget.config(state='normal')
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, file_path)
            try:
                xl = pd.ExcelFile(file_path)
                combo_widget.config(values=xl.sheet_names)
                combo_widget.current(0)  # Select first sheet by default
                self.log_message(f"Loaded file: {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load the file: {str(e)}")
                self.log_message(f"Error loading file: {str(e)}")

    def load_order_file(self):
        self.load_file(self.order_file_entry, self.order_sheet_combo)

    def load_master_file(self):
        self.load_file(self.master_file_entry, self.master_sheet_combo)

    def start_threaded_processing(self):
        """Start processing in a separate thread to keep the GUI responsive."""
        processing_thread = threading.Thread(target=self.start_processing)
        processing_thread.start()

    def start_processing(self):
        """Process the data after files and sheets are selected."""
        order_file = self.order_file_entry.get()
        master_file = self.master_file_entry.get()
        order_sheet = self.order_sheet_combo.get()
        master_sheet = self.master_sheet_combo.get()

        # Ensure all fields are filled
        if not order_file or not master_file or not order_sheet or not master_sheet:
            messagebox.showerror("Error", "Please fill in all fields")
            self.log_message("Validation failed: Missing fields")
            return

        try:
            self.log_message("Started processing...")
            
            # Load the selected sheets
            order_df = pd.read_excel(order_file, sheet_name=order_sheet)
            master_df = pd.read_excel(master_file, sheet_name=master_sheet)

            # Check if DataFrames are empty
            if order_df.empty or master_df.empty:
                messagebox.showerror("Error", "One or both selected sheets are empty.")
                self.log_message("Error: One or both sheets are empty.")
                return

            # Merge the DataFrames on 'SKU' and 'Parent Code'
            merged_df = order_df.merge(master_df, left_on='SKU', right_on='Parent Code', how='left')

            # Process the merged DataFrame
            merged_df['Child Code'] = merged_df['Child Code'].fillna(merged_df['SKU'])
            merged_df['Quantity'] = merged_df.apply(
                lambda row: row['Quantity_x'] * row['Quantity_y'] if pd.notna(row['Quantity_y']) else row['Quantity_x'], axis=1
            )

            # Select relevant columns for the result
            result_df = merged_df[['Payment Time', 'Order Number', 'Order Status', 'Channel', 'Store Name', 'Ref No', 'Child Code', 'Quantity']]

            # Save the result to a file
            self.save_config(result_df)
            self.log_message("Processing completed successfully.")

        except Exception as e:
            messagebox.showerror("Error", f"Error while processing: {str(e)}")
            self.log_message(f"Error while processing: {str(e)}")

    def save_config(self, result_df):
        """Save the result DataFrame to Excel or CSV."""
        # Get current timestamp for default file name
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"Result_{timestamp}"

        # Prompt user to select file format
        filetypes = [("Excel Workbook", "*.xlsx"), ("CSV (Comma delimited)", "*.csv")]
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=filetypes,
            initialfile=default_filename,
            title="Save As"
        )

        if save_path:
            try:
                if save_path.endswith(".csv"):
                    result_df.to_csv(save_path, index=False)
                else:
                    result_df.to_excel(save_path, index=False)

                messagebox.showinfo("Success", f"Process Completed!\nResult saved as {save_path}")
                self.log_message(f"File saved: {save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Error while saving the file: {str(e)}")
                self.log_message(f"Error while saving: {str(e)}")

    def download_template_master(self):
        """Allow users to download the template file."""
        # Specify the template structure as a DataFrame
        template_data = {
            "Parent Code": ["PARENT_01", "PARENT_01", "PARENT_02"],
            "Product Name": ["BUNDLE PRODUCT 1", "BUNDLE PRODUCT 1", "BUNDLE PRODUCT 2"],
            "Child Code": ["CHILD_001", "CHILD_002", "CHILD_001"],
            "Quantity": ["1", "1", "2"],
        }
        template_df = pd.DataFrame(template_data)

        # Prompt user for download location
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            initialfile="master_template.xlsx",
            title="Save As"
        )

        if save_path:
            try:
                template_df.to_excel(save_path, index=False)
                messagebox.showinfo("Success", f"Template downloaded as {save_path}")
                self.log_message(f"Template downloaded: {save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Error while downloading template: {str(e)}")
                self.log_message(f"Error while downloading template: {str(e)}")

    def download_template_order(self):
        """Allow users to download the template file."""
        # Specify the template structure as a DataFrame
        template_data = {
            "Payment Time": ["2024-07-01 00:00:37", "2024-07-01 00:00:37", "2024-07-01 00:00:37"],
            "Order Number": ["240701EM5T0QM3", "240701EM5T0QM3", "240701EM5T0QM3"],
            "Order Status": ["Completed", "Completed", "Completed"],
            "Channel": ["Shopee", "Shopee", "Shopee"],
            "Store Name": ["Store Name A", "Store Name A", "Store Name A"],
            "Ref No": ["EM5T0QM3", "EM5T0QM3", "EM5T0QM3"],
            "SKU": ["SINGLE_01", "BUNDLE_01", "SINGLE_02"],
            "Quantity": ["1", "1", "1"],
        }
        template_df = pd.DataFrame(template_data)

        # Prompt user for download location
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            initialfile="order_template.xlsx",
            title="Save As"
        )

        if save_path:
            try:
                template_df.to_excel(save_path, index=False)
                messagebox.showinfo("Success", f"Template downloaded as {save_path}")
                self.log_message(f"Template downloaded: {save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Error while downloading template: {str(e)}")
                self.log_message(f"Error while downloading template: {str(e)}")

    def log_message(self, message):
        """Log messages to the log text widget."""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')

if __name__ == "__main__":
    root = tk.Tk()
    app = BundleBreakdownApp(root)
    root.mainloop()
