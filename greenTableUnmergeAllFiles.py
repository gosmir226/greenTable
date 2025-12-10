import sys
import os
import glob
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QFileDialog, 
                             QLineEdit, QTextEdit, QProgressBar, QCheckBox)
from PyQt5.QtCore import Qt

class ExcelParserApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('Excel Parser for Non-Standard Tables')
        self.setGeometry(100, 100, 800, 600)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QVBoxLayout()
        
        # Directory selection
        dir_layout = QHBoxLayout()
        self.dir_label = QLabel('Selected directory:')
        self.dir_path = QLineEdit()
        self.browse_btn = QPushButton('Browse')
        self.browse_btn.clicked.connect(self.browse_directory)
        dir_layout.addWidget(self.dir_label)
        dir_layout.addWidget(self.dir_path)
        dir_layout.addWidget(self.browse_btn)
        
        # Output file selection
        output_layout = QHBoxLayout()
        self.output_label = QLabel('Output file:')
        self.output_path = QLineEdit('output.xlsx')
        self.output_browse_btn = QPushButton('Browse Output')
        self.output_browse_btn.clicked.connect(self.browse_output_file)
        output_layout.addWidget(self.output_label)
        output_layout.addWidget(self.output_path)
        output_layout.addWidget(self.output_browse_btn)
        
        # Data only option
        self.data_only_checkbox = QCheckBox("Read calculated values only (ignore formulas)")
        self.data_only_checkbox.setChecked(True)
        
        # Progress bar
        self.progress = QProgressBar()
        
        # Process button
        self.process_btn = QPushButton('Process Excel Files')
        self.process_btn.clicked.connect(self.process_directory)
        
        # Log output
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        
        # Add all to main layout
        layout.addLayout(dir_layout)
        layout.addLayout(output_layout)
        layout.addWidget(self.data_only_checkbox)
        layout.addWidget(self.progress)
        layout.addWidget(self.process_btn)
        layout.addWidget(self.log)
        
        central_widget.setLayout(layout)
        
    def browse_directory(self):
        directory = QFileDialog.getExistingDirectory(self, 'Select Directory')
        if directory:
            self.dir_path.setText(directory)
            
    def browse_output_file(self):
        file_name, _ = QFileDialog.getSaveFileName(self, 'Save Output File', 'output.xlsx', 'Excel Files (*.xlsx)')
        if file_name:
            self.output_path.setText(file_name)
            
    def log_message(self, message):
        self.log.append(message)
        QApplication.processEvents()
    
    def get_cell_value(self, sheet, merged_cell_ranges, row, col, processed_merged_cells):
        """Получить значение ячейки с учетом объединенных ячеек, избегая дублирования."""
        for merged_range in merged_cell_ranges:
            if merged_range.min_row <= row <= merged_range.max_row and \
               merged_range.min_col <= col <= merged_range.max_col:
                
                merged_key = (merged_range.min_row, merged_range.min_col,
                              merged_range.max_row, merged_range.max_col)
                
                if merged_key not in processed_merged_cells:
                    processed_merged_cells[merged_key] = True
                    return sheet.cell(merged_range.min_row, merged_range.min_col).value
                else:
                    return None
        
        return sheet.cell(row, col).value

    def process_sheet(self, workbook, sheet_name):
        """Обработка отдельного листа с новой логикой для объединенных ячеек"""
        sheet = workbook[sheet_name]
        self.log_message(f"Processing sheet: {sheet_name}")
        
        merged_ranges = list(sheet.merged_cells.ranges)
        self.log_message(f"Found {len(merged_ranges)} merged cell ranges.")
        
        # Поиск пустых строк для разделения блоков
        empty_rows = []
        for row_idx in range(1, sheet.max_row + 1):
            is_empty = True
            has_fill = False
            
            for col_idx in range(1, sheet.max_column + 1):
                cell_value = sheet.cell(row=row_idx, column=col_idx).value
                if cell_value is not None and str(cell_value).strip() != '':
                    is_empty = False
                    break
                
                cell = sheet.cell(row=row_idx, column=col_idx)
                if cell.fill.start_color.index != '00000000':
                    has_fill = True
                    break
            
            if is_empty and not has_fill:
                empty_rows.append(row_idx)
        
        # Разделение на цепочки
        chain_ranges = []
        start_row = 1
        
        for empty_row in empty_rows:
            if empty_row > start_row:
                chain_ranges.append((start_row, empty_row - 1))
            start_row = empty_row + 1
        
        if start_row <= sheet.max_row:
            chain_ranges.append((start_row, sheet.max_row))
        
        all_data = []
        
        for start_row, end_row in chain_ranges:
            # Определяем количество столбцов в блоке динамически
            col_start = 1
            
            while col_start <= sheet.max_column:
                processed_merged_cells = {}
                
                # Ищем дату в текущем блоке
                date_value = None
                date_row = start_row
                
                while date_row <= end_row and date_value is None:
                    cell_value = self.get_cell_value(sheet, merged_ranges, date_row, col_start, processed_merged_cells)
                    if cell_value and str(cell_value).strip():
                        date_value = cell_value
                    date_row += 1
                
                if not date_value:
                    col_start += 1
                    continue
                
                # Определяем конец текущего блока
                col_end = col_start
                while col_end <= sheet.max_column:
                    has_data = False
                    for check_row in range(start_row, min(start_row + 10, end_row)):
                        cell_value = self.get_cell_value(sheet, merged_ranges, check_row, col_end, {})
                        if cell_value and str(cell_value).strip():
                            has_data = True
                            break
                    
                    if not has_data and col_end > col_start:
                        col_end -= 1
                        break
                    
                    col_end += 1
                
                if col_end >= sheet.max_column:
                    col_end = sheet.max_column
                
                block_cols = col_end - col_start + 1
                
                # Ищем значение печи
                furnace_value = None
                if "УВНК" in sheet_name:
                    furnace_value = sheet_name
                else:
                    for r in range(start_row, end_row):
                        for c in range(col_start, col_end + 1):
                            cell_val = self.get_cell_value(sheet, merged_ranges, r, c, {})
                            if cell_val and "УППФ" in str(cell_val):
                                furnace_value = cell_val
                                break
                        if furnace_value:
                            break
                
                if not furnace_value:
                    furnace_value = ""
                
                # Определяем границы данных (исключая заголовки)
                data_start_row = start_row
                while data_start_row <= end_row:
                    has_numeric = False
                    for c in range(col_start, col_end + 1):
                        val = self.get_cell_value(sheet, merged_ranges, data_start_row, c, {})
                        if isinstance(val, (int, float)):
                            has_numeric = True
                            break
                    if has_numeric:
                        break
                    data_start_row += 1
                
                data_end_row = end_row
                while data_end_row > data_start_row:
                    has_data = False
                    for c in range(col_start, col_end + 1):
                        val = self.get_cell_value(sheet, merged_ranges, data_end_row, c, {})
                        if val and str(val).strip():
                            has_data = True
                            break
                    if has_data:
                        break
                    data_end_row -= 1
                
                if data_start_row >= data_end_row:
                    col_start = col_end + 1
                    continue
                
                # Собираем данные из блока
                for row in range(data_start_row, data_end_row + 1):
                    row_data = {
                        'Date': date_value,
                        'Furnace': furnace_value,
                        'Row': row - data_start_row + 1
                    }
                    
                    row_processed_cells = {}
                    
                    col_values = []
                    current_col = col_start
                    
                    while current_col <= col_end:
                        value = self.get_cell_value(sheet, merged_ranges, row, current_col, row_processed_cells)
                        
                        if value is not None:
                            col_values.append(value)
                            
                            for merged_range in merged_ranges:
                                if merged_range.min_row <= row <= merged_range.max_row and \
                                   merged_range.min_col <= current_col <= merged_range.max_col:
                                    current_col = merged_range.max_col
                                    break
                        
                        current_col += 1
                    
                    for i, value in enumerate(col_values):
                        row_data[f'Value{i+1}'] = value
                    
                    all_data.append(row_data)
                
                col_start = col_end + 1
        
        return all_data
        
    def process_directory(self):
        try:
            directory = self.dir_path.text()
            output_file = self.output_path.text()
            
            if not directory:
                self.log_message("Please select a directory first.")
                return
                
            self.log_message("Scanning directory for Excel files...")
            excel_files = glob.glob(os.path.join(directory, "**", "*.xlsx"), recursive=True)
            excel_files.extend(glob.glob(os.path.join(directory, "**", "*.xls"), recursive=True))
            
            if not excel_files:
                self.log_message("No Excel files found in the directory.")
                return
                
            self.log_message(f"Found {len(excel_files)} Excel files.")
            
            all_data = []
            total_files = len(excel_files)
            
            for file_idx, input_file in enumerate(excel_files):
                self.progress.setValue(int(file_idx / total_files * 100))
                self.log_message(f"Processing file {file_idx+1}/{total_files}: {os.path.basename(input_file)}")
                
                try:
                    if self.data_only_checkbox.isChecked():
                        workbook = load_workbook(filename=input_file, data_only=True)
                    else:
                        workbook = load_workbook(filename=input_file)
                    
                    for sheet_name in workbook.sheetnames:
                        sheet_data = self.process_sheet(workbook, sheet_name)
                        all_data.extend(sheet_data)
                        self.log_message(f"Extracted {len(sheet_data)} rows from {sheet_name}")
                    
                    workbook.close()
                    
                except Exception as e:
                    self.log_message(f"Error processing file {input_file}: {str(e)}")
            
            if all_data:
                df = pd.DataFrame(all_data)
                cols = ['Date', 'Furnace', 'Row'] + [col for col in df.columns if col not in ['Date', 'Furnace', 'Row']]
                df = df[cols]
                df.to_excel(output_file, index=False)
                self.log_message(f"Data successfully saved to {output_file}. Total rows: {len(all_data)}")
            else:
                self.log_message("No data was extracted.")
                
            self.progress.setValue(100)
            
        except Exception as e:
            self.log_message(f"Error: {str(e)}")
            self.progress.setValue(0)

def main():
    app = QApplication(sys.argv)
    window = ExcelParserApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()