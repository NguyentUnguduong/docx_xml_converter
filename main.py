"""
DOCX to XML Converter Application
Chuyển đổi file DOCX sang XML theo cấu trúc câu hỏi
"""

import subprocess
import sys
import os
from pathlib import Path
import tempfile
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QListWidget, 
                             QFileDialog, QProgressBar, QTextEdit, QGroupBox,QDialog,
                             QMessageBox, QSplitter)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon
import traceback
import requests
from packaging  import version
import json

from docx_processor import DocxProcessor # Import lớp đã cập nhật

class ProcessingThread(QThread):
    """Thread xử lý file để không block UI"""
    progress = pyqtSignal(str)  # Thông báo tiến trình

    finished = pyqtSignal(bool, str, dict)  # Kết quả: (overall_success, overall_message, file_results)

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
            failed_count = 0
            file_results = {} # Dictionary để lưu kết quả cho từng file

            for idx, input_file in enumerate(self.input_files, 1):
                self.file_progress.emit(idx, total_files)
                
                file_name = Path(input_file).stem
                self.progress.emit(f"🔄 Đang xử lý: {file_name}.docx...")
                
                try:
                    # GỌI process_docx MỚI - Trả về xml_content và danh sách lỗi
                    xml_content, errors = self.processor.process_docx(input_file)
                    
                    if errors:
                        file_results[file_name] = {
                            'status': 'error',
                            'errors': errors
                        }
                        self.progress.emit(f"⚠️ Hoàn thành có lỗi: {file_name}.docx")
                        for err in errors:
                             self.progress.emit(f"   - {err}")
                        failed_count += 1
                    else:
                        file_results[file_name] = {
                            'status': 'success',
                            'errors': []
                        }
                        self.progress.emit(f"✅ Hoàn thành: {file_name}.xml")
                        success_count += 1
                    
                    # Luôn lưu file, ngay cả khi có lỗi (nếu có thể)
                    output_file = os.path.join(self.output_dir, f"{file_name}.xml")
                    with open(output_file, 'w', encoding='utf-8') as f:
                        f.write(xml_content)
                    
                except Exception as e:
                    error_msg = f"❌ Lỗi nghiêm trọng khi xử lý {file_name}.docx: {str(e)}"
                    self.progress.emit(error_msg)
                    self.progress.emit(f"   Chi tiết: {traceback.format_exc()}")
                    file_results[file_name] = {
                        'status': 'critical_error',
                        'errors': [str(e)]
                    }
                    failed_count += 1
            
            # Tạo thông báo tổng thể
            overall_success = failed_count == 0
            if success_count == total_files:
                overall_message = f"✅ Xử lý thành công {success_count}/{total_files} file!"
            elif success_count > 0:
                overall_message = f"⚠️ Xử lý xong {total_files}/{total_files} file. " \
                                  f"{success_count} thành công, {failed_count} có lỗi."
            else:
                overall_message = f"❌ Không có file nào được xử lý thành công hoàn toàn! {failed_count} file có lỗi."

            # Gửi tín hiệu hoàn thành với kết quả chi tiết
            self.finished.emit(overall_success, overall_message, file_results)
                
        except Exception as e:
            self.finished.emit(False, f"❌ Lỗi nghiêm trọng trong thread: {str(e)}", {})


# CURRENT_VERSION = "1.0.0"  # <-- Bạn tự cập nhật mỗi lần release
# CURRENT_VERSION = get_current_version()
GITHUB_REPO = "NguyentUnguduong/docx_xml_converter"  # Ví dụ: "nguyenvanA/my-docx-xml-converter"

def get_version_file_path():
    """Trả về đường dẫn tới version.json đúng vị trí"""
    if getattr(sys, "frozen", False):
        # Nếu là exe
        base_path = os.path.dirname(sys.executable)
    else:
        # Nếu chạy từ Python
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, "version.json")

def get_current_version():
    """Đọc version hiện tại từ version.json"""
    version_file = get_version_file_path()
    print(f"[DEBUG] Đang đọc version từ: {version_file}")
    try:
        if os.path.exists(version_file):
            with open(version_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                print(f"[DEBUG] Nội dung version.json: {data}")
            return data.get("version", "0.0.0")
        else:
            return "0.0.0"
    except Exception as e:
        print(f"Lỗi đọc version.json: {e}")
        return "0.0.0"

def update_local_version(new_version):
    """Ghi version mới vào version.json cùng thư mục với exe hoặc main.py"""
    version_file = get_version_file_path()
    try:
        with open(version_file, "w", encoding="utf-8") as f:
            json.dump({"version": new_version}, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"Lỗi ghi version.json: {e}")
        
def check_for_update():
    """Kiểm tra update từ GitHub, trả về (has_update, exe_url, latest_ver)"""
    try:
        CURRENT_VERSION = get_current_version()
        url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        response = requests.get(url, timeout=10)
        data = response.json()
        latest_tag = data.get("tag_name", "0.0.0").lstrip("vV")
        assets = data.get("assets", [])

        # Tìm file .exe trong assets
        exe_url = None
        for asset in assets:
            if asset["name"].endswith(".exe"):
                exe_url = asset["browser_download_url"]
                break

        if not exe_url:
            return False, None, latest_tag

        if version.parse(latest_tag) > version.parse(CURRENT_VERSION):
            return True, exe_url, latest_tag
        return False, None, latest_tag
    except Exception:
        return False, None, None


def download_and_update(download_url, latest_version):
    """Tải file exe mới, thay thế, ghi version.json và restart app"""
    try:
        # Tải file vào temp
        temp_dir = tempfile.gettempdir()
        new_exe = os.path.join(temp_dir, "updated_app.exe")

        with requests.get(download_url, stream=True) as r:
            r.raise_for_status()
            with open(new_exe, 'wb') as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)

        if not getattr(sys, "frozen", False):
            QMessageBox.warning(None, "Không thể cập nhật",
                                "Cập nhật chỉ hoạt động khi chạy file .exe đã đóng gói.")
            return False

        current_exe = sys.executable
        exe_name = os.path.basename(current_exe).lower()

        # Kiểm tra exe nhạy cảm
        forbidden_exes = ["python.exe", "python313.exe"]
        if exe_name in forbidden_exes:
            QMessageBox.critical(None, "Cảnh báo",
                                 f"Không thể cập nhật từ {exe_name}")
            return False

        # Tạo batch script xóa exe cũ và replace
        bat_script = os.path.join(temp_dir, "update.bat")
        with open(bat_script, "w", encoding="utf-8") as bat:
            bat.write(f'''
@echo off
timeout /t 2 /nobreak >nul
del "{current_exe}"
move "{new_exe}" "{current_exe}"
start "" "{current_exe}"
''')

        # Ghi version.json mới
        update_local_version(latest_version)

        # Chạy batch và thoát app
        subprocess.Popen([bat_script], shell=True)
        sys.exit(0)

    except Exception as e:
        QMessageBox.critical(None, "Lỗi cập nhật", f"Không thể cập nhật:\n{str(e)}")
        return False


class DownloadWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)  # truyền đường dẫn file tải xong
    error = pyqtSignal(str)

    def __init__(self, url, save_path):
        super().__init__()
        self.url = url
        self.save_path = save_path

    def run(self):
        try:
            with requests.get(self.url, stream=True, timeout=30) as r:
                r.raise_for_status()
                total_size = int(r.headers.get('content-length', 0))
                downloaded = 0
                with open(self.save_path, 'wb') as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)
                            downloaded += len(chunk)
                            if total_size > 0:
                                perc = int(100 * downloaded / total_size)
                                self.progress.emit(perc)
                self.finished.emit(self.save_path)
        except Exception as e:
            self.error.emit(str(e))


class UpdateDialog(QDialog):
    def __init__(self, current_version, latest_version, download_url, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Cập nhật phần mềm")

        self.setFixedSize(450, 250)  # giảm chiều cao ban đầu

        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)


        self.download_url = download_url

        self.current_version = current_version

        self.latest_version = latest_version

        self.temp_exe_path = None

        self.worker = None

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        layout.setContentsMargins(20, 20, 20, 20)

        layout.setSpacing(15)

        # Title
        title = QLabel("🔔 Có bản cập nhật mới!")
        title.setFont(QFont("Segoe UI", 14, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # Version info
        info = QLabel(
            f"<b>Phiên bản hiện tại:</b> v{self.current_version}<br>"
            f"<b>Phiên bản mới:</b> v{self.latest_version}"
        )
        info.setFont(QFont("Segoe UI", 11))

        info.setAlignment(Qt.AlignCenter)

        layout.addWidget(info)

        # Status label (ẩn ban đầu)
        self.status_label = QLabel("")

        self.status_label.setFont(QFont("Segoe UI", 10))

        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.hide()
        layout.addWidget(self.status_label)

        # Progress bar (ẩn ban đầu)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.hide()
        layout.addWidget(self.progress_bar)

        # Buttons
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(20)

        self.btn_update = QPushButton("Cập nhật")
        self.btn_later = QPushButton("Để sau")

        self.btn_update.setStyleSheet("""
            QPushButton {
                background-color: #28a745;
                color: white;
                padding: 8px 18px;
                border-radius: 6px;
                font-size: 12pt;
            }
            QPushButton:hover {
                background-color: #218838;
            }
        """)

        self.btn_later.setStyleSheet("""
            QPushButton {
                background-color: #cccccc;
                color: black;
                padding: 8px 18px;
                border-radius: 6px;
                font-size: 12pt;
            }
            QPushButton:hover {
                background-color: #b6b6b6;
            }
        """)

        btn_layout.addWidget(self.btn_update)

        btn_layout.addWidget(self.btn_later)

        layout.addLayout(btn_layout)

        self.setLayout(layout)

        self.btn_update.clicked.connect(self.start_update)

        self.btn_later.clicked.connect(self.reject)

    def start_update(self):
        # Ẩn nút, hiện progress
        self.btn_update.hide()
        self.btn_later.hide()
        self.status_label.setText("Đang tải bản cập nhật...")
        self.status_label.show()
        self.progress_bar.show()
        self.setFixedSize(450, 280)

        # Tạo thư mục TEMP/app_update
        temp_root = tempfile.gettempdir()
        update_folder = os.path.join(temp_root, "app_update")

        os.makedirs(update_folder, exist_ok=True)

        # File exe mới nằm trong thư mục tạm cố định
        self.temp_exe_path = os.path.join(update_folder, "new_app.exe")
        self.update_folder = update_folder

        # Bắt đầu tải
        self.worker = DownloadWorker(self.download_url, self.temp_exe_path)
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.on_download_finished)
        self.worker.error.connect(self.on_download_error)
        self.worker.start()

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def on_download_finished(self, file_path):
        self.status_label.setText("Đang áp dụng cập nhật...")

        # Ghi version mới
        self.update_local_version(self.latest_version)

        current_exe = sys.executable
        update_folder = self.update_folder
        bat_script = os.path.join(update_folder, "update.bat")
        log_file = os.path.join(update_folder, "update.log")

        # Tạo nội dung batch an toàn
        bat_content = fr'''@echo off
    chcp 65001 >nul
    set "LOGFILE={log_file}"
    set "UPDATE_FOLDER={update_folder}"
    set "CURRENT_EXE={current_exe}"
    set "NEW_EXE={file_path}"

    echo =============================== >> "%LOGFILE%"
    echo Update process started at %date% %time% >> "%LOGFILE%"
    echo Current EXE: %CURRENT_EXE% >> "%LOGFILE%"
    echo New EXE: %NEW_EXE% >> "%LOGFILE%"

    :: Đợi 5 giây để đảm bảo app cũ hoàn toàn thoát
    timeout /t 5 /nobreak >nul

    :: Thử xóa file cũ — nếu fail thì ghi log và tiếp tục
    echo [INFO] Deleting old EXE... >> "%LOGFILE%"
    del /f /q "%CURRENT_EXE%" >> "%LOGFILE%" 2>&1

    :: Di chuyển file mới vào vị trí
    echo [INFO] Moving new EXE into place... >> "%LOGFILE%"
    move /y "%NEW_EXE%" "%CURRENT_EXE%" >> "%LOGFILE%" 2>&1

    :: Kiểm tra file mới tồn tại
    if not exist "%CURRENT_EXE%" (
        echo [ERROR] Failed to replace EXE! >> "%LOGFILE%"
        pause
        exit /b 1
    )

    echo [SUCCESS] EXE replaced successfully. >> "%LOGFILE%"

    :: Khởi động lại app — dùng start để tách tiến trình
    echo [INFO] Restarting application... >> "%LOGFILE%"
    start "" "%CURRENT_EXE%" >> "%LOGFILE%" 2>&1

    :: Dọn dẹp sau 10 giây — tránh lock
    echo [INFO] Scheduling cleanup... >> "%LOGFILE%"
    (
        timeout /t 10 /nobreak >nul
        rmdir /s /q "%UPDATE_FOLDER%" >nul 2>&1
    ) >nul 2>&1 &

    exit
    '''

        try:
            with open(bat_script, "w", encoding="utf-8-sig") as f:
                f.write(bat_content)

            # Đảm bảo app hiện tại thoát hoàn toàn
            self.accept()  # Đóng dialog
            QApplication.quit()  # Đóng Qt
            # DÙNG subprocess.Popen để chạy batch, rồi exit
            subprocess.Popen([bat_script], shell=True, creationflags=subprocess.CREATE_NEW_CONSOLE)
            sys.exit(0)  # Thoát hoàn toàn

        except Exception as e:
            self.show_error(f"Không thể áp dụng cập nhật:\n{str(e)}")

    def on_download_error(self, error_msg):
        self.show_error(f"Lỗi khi tải cập nhật:\n{error_msg}")

    def show_error(self, msg):
        self.status_label.setText("❌ Cập nhật thất bại")
        QMessageBox.critical(self, "Lỗi cập nhật", msg)
        self.reject()

    def update_local_version(self, new_version):
        """Ghi version.json (giống hàm toàn cục, nhưng có thể reuse)"""
        if getattr(sys, "frozen", False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        version_file = os.path.join(base_path, "version.json")
        try:
            with open(version_file, "w", encoding="utf-8") as f:
                json.dump({"version": new_version}, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Lỗi ghi version.json: {e}")

    def closeEvent(self, event):
        # Đảm bảo luồng được dừng (nếu cần)
        if self.worker and self.worker.isRunning():
            self.worker.quit()
            self.worker.wait()
        super().closeEvent(event)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.input_files = []
        self.output_dir = ""
        self.processing_thread = None
        self.detail_results_text = ""
        self.init_ui()
        self.check_update_on_start()

    # def check_update_on_start(self):
    #     """Kiểm tra cập nhật ngay khi app mở"""
    #     try:
    #         current_version = get_current_version()
    #         has_update, url, latest_ver = check_for_update()
    #         if has_update and url:
    #             dialog = UpdateDialog(current_version, latest_ver, url, self)
    #             choice = dialog.exec_()
    #             if choice == "update":
    #                 download_and_update(url, latest_ver)
    #     except Exception as e:
    #         print(f"Lỗi khi kiểm tra cập nhật: {e}")
    def check_update_on_start(self):
        """Kiểm tra cập nhật ngay khi app mở"""
        try:
            current_version = get_current_version()
            has_update, url, latest_ver = check_for_update()
            if has_update and url:
                # Hiển thị dialog có tiến trình tải
                dialog = UpdateDialog(current_version, latest_ver, url, self)
                dialog.exec_()  # dialog sẽ tự xử lý tải + cập nhật + thoát nếu cần
                # ⚠️ Nếu cập nhật thành công, app đã exit rồi → dòng dưới KHÔNG CHẠY
                # Nếu người dùng bấm "Để sau", exec_() trả về và app tiếp tục bình thường
        except Exception as e:
            print(f"[Lỗi khi kiểm tra cập nhật]: {e}")
            # Có thể hiện QMessageBox nếu muốn, nhưng không bắt buộc
        
    def init_ui(self):
        """Khởi tạo giao diện"""
        # ... (phần code UI cũ giữ nguyên) ...
        self.setWindowTitle("Công cụ chuyển đổi file docx sang XML")
        self.setGeometry(100, 100, 1000, 700)
        
        # Widget chính
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # Header
        header_label = QLabel("📄 Công cụ chuyển đổi từ file DOCX sang file XML")
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
        self.process_btn = QPushButton("🚀 Bắt đầu chuyển đổi")
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
        # CẬP NHẬT: Nhận thêm file_results
        self.processing_thread.finished.connect(self.processing_finished)
        self.processing_thread.start()
    
    def update_progress(self, current, total):
        """Cập nhật progress bar"""
        progress = int((current / total) * 100)
        self.progress_bar.setValue(progress)
        self.progress_label.setText(f"Đang xử lý file {current}/{total}")
        self.statusBar().showMessage(f"Tiến trình: {current}/{total} file")
    
    def processing_finished(self, overall_success, overall_message, file_results):
        """Xử lý xong - CẬP NHẬT để nhận file_results và tạo nội dung chi tiết"""
        self.log("\n" + "="*60)
        self.log("KẾT QUẢ TỔNG THỂ:")
        self.log(overall_message)

        # **Tạo chuỗi văn bản chi tiết để hiển thị khi nhấn nút**
        detailed_text = "📄 KẾT QUẢ CHI TIẾT CHO TỪNG FILE\n"
        detailed_text += "="*50 + "\n"

        has_errors = any(result['status'] != 'success' for result in file_results.values())
        if has_errors:
            detailed_text += "\n--- 📌 CHI TIẾT LỖI ---\n"
            for file_name, result in file_results.items():
                if result['status'] == 'success':
                    detailed_text += f"✅ {file_name}.docx: Thành công - Không có lỗi\n"
                else: # error hoặc critical_error
                    status_icon = "❌" if result['status'] == 'critical_error' else "⚠️"
                    detailed_text += f"{status_icon} {file_name}.docx:\n"
                    for err in result['errors']:
                        detailed_text += f"      • {err}\n"
        else:
            detailed_text += "\n🎉 Tất cả các file đều được xử lý thành công!\n"
        
        detailed_text += "\n" + "="*50 + "\n"
        self.detailed_results_text = detailed_text

        # In tóm tắt vào log chính
        self.log(detailed_text)

        self.progress_bar.setValue(100)
        self.progress_label.setText("Hoàn thành!")
        self.set_buttons_enabled(True)
        
        # Show message box với nút tùy chỉnh
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Xử lý hoàn tất")
        msg_box.setText(overall_message)

        # Thiết lập icon dựa trên overall_success
        msg_box.setIcon(QMessageBox.Information if overall_success else QMessageBox.Warning)

        # Thêm các nút
        view_details_btn = msg_box.addButton("🔍 Xem Chi Tiết", QMessageBox.ActionRole)
        open_folder_btn = msg_box.addButton("📂 Mở Thư Mục Kết Quả", QMessageBox.AcceptRole)
        close_btn = msg_box.addButton("Đóng", QMessageBox.RejectRole)

        # Hiển thị hộp thoại
        msg_box.exec_()

        # Kiểm tra nút nào được nhấn
        clicked_button = msg_box.clickedButton()
        if clicked_button == view_details_btn:
            # Hiển thị một hộp thoại thông tin khác với nội dung chi tiết
            self.show_detail_results()
        elif clicked_button == open_folder_btn:
            # Mở thư mục kết quả
            try:
                os.startfile(self.output_dir)
            except Exception as e:
                QMessageBox.critical(self, "Lỗi", f"Không thể mở thư mục: {str(e)}")
    
    def show_detail_results(self):
        """Hiển thị popup chứa chi tiết kết quả"""
        dlg = QDialog(self)
        dlg.setWindowTitle("Chi tiết kết quả xử lý")
        dlg.setMinimumSize(600, 500)

        layout = QVBoxLayout(dlg)

        text = QTextEdit()
        text.setReadOnly(True)
        text.setText(self.detailed_results_text)
        layout.addWidget(text)

        close_btn = QPushButton("Đóng")
        close_btn.clicked.connect(dlg.close)
        layout.addWidget(close_btn)

        dlg.exec_()

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
