"""
DOCX to XML Converter Application
Chuyển đổi file DOCX sang XML theo cấu trúc câu hỏi
"""

import sys
import os
from pathlib import Path
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QListWidget, 
                             QFileDialog, QProgressBar, QTextEdit, QGroupBox,
                             QMessageBox, QSplitter)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon
import traceback

from docx_processor import DocxProcessor


class ProcessingThread(QThread):
    """Thread xử lý file để không block UI"""
    progress = pyqtSignal(str)  # Thông báo tiến trình
    finished = pyqtSignal(bool, str)  # Kết quả: (success, message)
    file_progress = pyqtSignal(int, int)  # (current_file, total_files)
    
    def __init__(self, input_files, output_dir):
        super().__init__()
        self.input_files = input_files
        self.output_dir = output_dir
        self.processor = DocxProcessor()
        
    def run(self):
        try:
            total_files = len(self.input_files)
            success_count = 0
            
            for idx, input_file in enumerate(self.input_files, 1):
                self.file_progress.emit(idx, total_files)
                
                file_name = Path(input_file).stem
                self.progress.emit(f"🔄 Đang xử lý: {file_name}.docx...")
                
                try:
                    # Xử lý file
                    xml_content = self.processor.process_docx(input_file)
                    
                    # Lưu file XML
                    output_file = os.path.join(self.output_dir, f"{file_name}.xml")
                    with open(output_file, 'w', encoding='utf-8') as f:
                        f.write(xml_content)
                    
                    self.progress.emit(f"✅ Hoàn thành: {file_name}.xml")
                    success_count += 1
                    
                except Exception as e:
                    error_msg = f"❌ Lỗi khi xử lý {file_name}.docx: {str(e)}"
                    self.progress.emit(error_msg)
                    self.progress.emit(f"   Chi tiết: {traceback.format_exc()}")
            
            # Thông báo kết quả
            if success_count == total_files:
                self.finished.emit(True, 
                    f"✅ Xử lý thành công {success_count}/{total_files} file!")
            elif success_count > 0:
                self.finished.emit(True, 
                    f"⚠️ Xử lý thành công {success_count}/{total_files} file. "
                    f"Có {total_files - success_count} file lỗi.")
            else:
                self.finished.emit(False, 
                    "❌ Không có file nào được xử lý thành công!")
                
        except Exception as e:
            self.finished.emit(False, f"❌ Lỗi nghiêm trọng: {str(e)}")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.input_files = []
        self.output_dir = ""
        self.processing_thread = None
        self.init_ui()
        
    def init_ui(self):
        """Khởi tạo giao diện"""
        self.setWindowTitle("DOCX to XML Converter - Công cụ chuyển đổi câu hỏi")
        self.setGeometry(100, 100, 1000, 700)
        
        # Widget chính
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # Header
        header_label = QLabel("📄 DOCX to XML Converter")
        header_label.setFont(QFont("Arial", 18, QFont.Bold))
        header_label.setAlignment(Qt.AlignCenter)
        header_label.setStyleSheet("""
            QLabel {
                color: #2c3e50;
                padding: 15px;
                background-color: #ecf0f1;
                border-radius: 8px;
            }
        """)
        main_layout.addWidget(header_label)
        
        # Splitter cho 2 phần chính
        splitter = QSplitter(Qt.Horizontal)
        
        # ===== PHẦN TRÁI: Chọn file =====
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        
        # Group box danh sách file
        file_group = QGroupBox("📁 Danh sách file DOCX")
        file_group.setFont(QFont("Arial", 10, QFont.Bold))
        file_layout = QVBoxLayout()
        
        self.file_list = QListWidget()
        self.file_list.setStyleSheet("""
            QListWidget {
                border: 2px solid #3498db;
                border-radius: 5px;
                padding: 5px;
                background-color: white;
            }
        """)
        file_layout.addWidget(self.file_list)
        
        # Buttons cho file
        file_btn_layout = QHBoxLayout()
        
        self.add_files_btn = QPushButton("➕ Thêm file")
        self.add_files_btn.setStyleSheet(self.get_button_style("#3498db"))
        self.add_files_btn.clicked.connect(self.add_files)
        
        self.remove_file_btn = QPushButton("➖ Xóa file")
        self.remove_file_btn.setStyleSheet(self.get_button_style("#e74c3c"))
        self.remove_file_btn.clicked.connect(self.remove_selected_file)
        
        self.clear_files_btn = QPushButton("🗑️ Xóa tất cả")
        self.clear_files_btn.setStyleSheet(self.get_button_style("#95a5a6"))
        self.clear_files_btn.clicked.connect(self.clear_files)
        
        file_btn_layout.addWidget(self.add_files_btn)
        file_btn_layout.addWidget(self.remove_file_btn)
        file_btn_layout.addWidget(self.clear_files_btn)
        file_layout.addLayout(file_btn_layout)
        
        file_group.setLayout(file_layout)
        left_layout.addWidget(file_group)
        
        # Chọn thư mục đầu ra
        output_group = QGroupBox("💾 Thư mục lưu kết quả")
        output_group.setFont(QFont("Arial", 10, QFont.Bold))
        output_layout = QVBoxLayout()
        
        self.output_label = QLabel("Chưa chọn thư mục")
        self.output_label.setStyleSheet("""
            QLabel {
                padding: 10px;
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 5px;
            }
        """)
        self.output_label.setWordWrap(True)
        output_layout.addWidget(self.output_label)
        
        self.select_output_btn = QPushButton("📂 Chọn thư mục")
        self.select_output_btn.setStyleSheet(self.get_button_style("#27ae60"))
        self.select_output_btn.clicked.connect(self.select_output_dir)
        output_layout.addWidget(self.select_output_btn)
        
        output_group.setLayout(output_layout)
        left_layout.addWidget(output_group)
        
        # Nút xử lý
        self.process_btn = QPushButton("🚀 BẮT ĐẦU XỬ LÝ")
        self.process_btn.setFont(QFont("Arial", 12, QFont.Bold))
        self.process_btn.setMinimumHeight(50)
        self.process_btn.setStyleSheet(self.get_button_style("#16a085", 50))
        self.process_btn.clicked.connect(self.start_processing)
        left_layout.addWidget(self.process_btn)
        
        splitter.addWidget(left_widget)
        
        # ===== PHẦN PHẢI: Log và tiến trình =====
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        
        # Progress bar
        progress_group = QGroupBox("📊 Tiến trình xử lý")
        progress_group.setFont(QFont("Arial", 10, QFont.Bold))
        progress_layout = QVBoxLayout()
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #3498db;
                border-radius: 5px;
                text-align: center;
                height: 25px;
            }
            QProgressBar::chunk {
                background-color: #3498db;
            }
        """)
        progress_layout.addWidget(self.progress_bar)
        
        self.progress_label = QLabel("Sẵn sàng")
        self.progress_label.setAlignment(Qt.AlignCenter)
        progress_layout.addWidget(self.progress_label)
        
        progress_group.setLayout(progress_layout)
        right_layout.addWidget(progress_group)
        
        # Log area
        log_group = QGroupBox("📋 Nhật ký xử lý")
        log_group.setFont(QFont("Arial", 10, QFont.Bold))
        log_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet("""
            QTextEdit {
                border: 2px solid #95a5a6;
                border-radius: 5px;
                background-color: #2c3e50;
                color: #ecf0f1;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 10pt;
            }
        """)
        log_layout.addWidget(self.log_text)
        
        self.clear_log_btn = QPushButton("🧹 Xóa log")
        self.clear_log_btn.setStyleSheet(self.get_button_style("#7f8c8d"))
        self.clear_log_btn.clicked.connect(lambda: self.log_text.clear())
        log_layout.addWidget(self.clear_log_btn)
        
        log_group.setLayout(log_layout)
        right_layout.addWidget(log_group)
        
        splitter.addWidget(right_widget)
        
        # Set tỷ lệ cho splitter
        splitter.setSizes([400, 600])
        main_layout.addWidget(splitter)
        
        # Status bar
        self.statusBar().showMessage("Sẵn sàng xử lý file")
        
    def get_button_style(self, color, height=40):
        """Tạo style cho button"""
        return f"""
            QPushButton {{
                background-color: {color};
                color: white;
                border: none;
                padding: 10px;
                border-radius: 5px;
                font-weight: bold;
                min-height: {height}px;
            }}
            QPushButton:hover {{
                background-color: {self.darken_color(color)};
            }}
            QPushButton:pressed {{
                background-color: {self.darken_color(color, 0.8)};
            }}
            QPushButton:disabled {{
                background-color: #bdc3c7;
            }}
        """
    
    def darken_color(self, hex_color, factor=0.9):
        """Làm tối màu"""
        hex_color = hex_color.lstrip('#')
        r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        r, g, b = int(r * factor), int(g * factor), int(b * factor)
        return f"#{r:02x}{g:02x}{b:02x}"
    
    def add_files(self):
        """Thêm file DOCX"""
        files, _ = QFileDialog.getOpenFileNames(
            self, "Chọn file DOCX", "", "Word Documents (*.docx)"
        )
        if files:
            for file in files:
                if file not in self.input_files:
                    self.input_files.append(file)
                    self.file_list.addItem(Path(file).name)
            self.log(f"✅ Đã thêm {len(files)} file")
            self.statusBar().showMessage(f"Đã có {len(self.input_files)} file")
    
    def remove_selected_file(self):
        """Xóa file đã chọn"""
        current_row = self.file_list.currentRow()
        if current_row >= 0:
            removed = self.input_files.pop(current_row)
            self.file_list.takeItem(current_row)
            self.log(f"🗑️ Đã xóa: {Path(removed).name}")
            self.statusBar().showMessage(f"Còn {len(self.input_files)} file")
    
    def clear_files(self):
        """Xóa tất cả file"""
        if self.input_files:
            reply = QMessageBox.question(
                self, "Xác nhận", 
                "Bạn có chắc muốn xóa tất cả file?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.input_files.clear()
                self.file_list.clear()
                self.log("🗑️ Đã xóa tất cả file")
                self.statusBar().showMessage("Danh sách file trống")
    
    def select_output_dir(self):
        """Chọn thư mục đầu ra"""
        dir_path = QFileDialog.getExistingDirectory(self, "Chọn thư mục lưu file XML")
        if dir_path:
            self.output_dir = dir_path
            self.output_label.setText(dir_path)
            self.log(f"📂 Thư mục đầu ra: {dir_path}")
    
    def log(self, message):
        """Thêm log"""
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
    
    def start_processing(self):
        """Bắt đầu xử lý"""
        # Validate
        if not self.input_files:
            QMessageBox.warning(self, "Cảnh báo", "Vui lòng chọn ít nhất 1 file DOCX!")
            return
        
        if not self.output_dir:
            QMessageBox.warning(self, "Cảnh báo", "Vui lòng chọn thư mục lưu kết quả!")
            return
        
        # Disable buttons
        self.set_buttons_enabled(False)
        self.progress_bar.setValue(0)
        self.log("\n" + "="*60)
        self.log("🚀 BẮT ĐẦU XỬ LÝ...")
        self.log("="*60)
        
        # Start processing thread
        self.processing_thread = ProcessingThread(self.input_files, self.output_dir)
        self.processing_thread.progress.connect(self.log)
        self.processing_thread.file_progress.connect(self.update_progress)
        self.processing_thread.finished.connect(self.processing_finished)
        self.processing_thread.start()
    
    def update_progress(self, current, total):
        """Cập nhật progress bar"""
        progress = int((current / total) * 100)
        self.progress_bar.setValue(progress)
        self.progress_label.setText(f"Đang xử lý file {current}/{total}")
        self.statusBar().showMessage(f"Tiến trình: {current}/{total} file")
    
    def processing_finished(self, success, message):
        """Xử lý xong"""
        self.log("\n" + "="*60)
        self.log(message)
        self.log("="*60 + "\n")
        
        self.progress_bar.setValue(100)
        self.progress_label.setText("Hoàn thành!")
        self.set_buttons_enabled(True)
        
        # Show message box
        if success:
            QMessageBox.information(self, "Thành công", message)
            # Mở thư mục kết quả
            reply = QMessageBox.question(
                self, "Mở thư mục", 
                "Bạn có muốn mở thư mục kết quả?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                os.startfile(self.output_dir)
        else:
            QMessageBox.critical(self, "Lỗi", message)
    
    def set_buttons_enabled(self, enabled):
        """Enable/disable buttons"""
        self.add_files_btn.setEnabled(enabled)
        self.remove_file_btn.setEnabled(enabled)
        self.clear_files_btn.setEnabled(enabled)
        self.select_output_btn.setEnabled(enabled)
        self.process_btn.setEnabled(enabled)


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()