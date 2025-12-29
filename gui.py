import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QFileDialog, QTextEdit, 
                             QGroupBox, QMessageBox, QTabWidget, QProgressBar, QFrame)
from PyQt5.QtCore import Qt, QSettings
from PyQt5.QtGui import QFont, QPalette, QColor, QIcon

class ModernFileInput(QWidget):
    def __init__(self, title, file_types="تمام فایل‌ها (*)"):
        super().__init__()
        self.file_types = file_types
        self.setup_ui(title)
        
    def setup_ui(self, title):
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        self.setLayout(layout)
        
        self.line_edit = QLineEdit()
        self.line_edit.setPlaceholderText(f"فایل {title} را اینجا رها کنید یا با کلیک بر روی دکمه انتخاب کنید")
        self.line_edit.setAcceptDrops(True)
        self.line_edit.dragEnterEvent = self.dragEnterEvent
        self.line_edit.dropEvent = self.dropEvent
        self.line_edit.setStyleSheet("""
            QLineEdit {
                padding: 8px;
                border: 2px solid #ddd;
                border-radius: 4px;
                background: white;
                font-family: Arial;
                font-size: 12pt;
            }
            QLineEdit:hover {
                border-color: #aaa;
            }
        """)
        
        self.browse_btn = QPushButton("انتخاب فایل...")
        self.browse_btn.setStyleSheet("""
            QPushButton {
                padding: 8px 12px;
                background: #4a6fa5;
                color: white;
                border: none;
                border-radius: 4px;
                font-family: Arial;
                font-size: 12pt;
            }
            QPushButton:hover {
                background: #3a5a8f;
            }
        """)
        self.browse_btn.clicked.connect(self.browse_file)
        
        layout.addWidget(self.line_edit, 1)
        layout.addWidget(self.browse_btn)
        
    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, f"انتخاب فایل {self.file_types}", "", self.file_types)
        if file_path:
            self.line_edit.setText(file_path)
            
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            
    def dropEvent(self, event):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if os.path.isfile(file_path):
                self.line_edit.setText(file_path)
                break
                
    def get_path(self):
        return self.line_edit.text()
    
    def set_path(self, path):
        self.line_edit.setText(path)

class XmlProcessorGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("XML_Processor", "پردازشگر_XML")
        self.setup_ui()
        self.load_settings()
        
    def setup_ui(self):
        self.setWindowTitle("پردازشگر پیشرفته XML")
        self.setMinimumSize(900, 650)
        
        # تنظیم تم روشن
        self.set_light_theme()
        
        # تنظیم فونت فارسی
        font = QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        self.setFont(font)
        
        # ویجت مرکزی
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        central_widget.setLayout(main_layout)
        
        # تب‌های اصلی
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 10px;
            }
            QTabBar::tab {
                padding: 8px 12px;
                background: #f0f0f0;
                border: 1px solid #ddd;
                border-bottom: none;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                font-family: Arial;
                font-size: 12pt;
            }
            QTabBar::tab:selected {
                background: white;
                border-bottom: 1px solid white;
                margin-bottom: -1px;
            }
        """)
        main_layout.addWidget(self.tabs)
        
        # تب اصلی
        self.main_tab = QWidget()
        self.tabs.addTab(self.main_tab, "  پردازش  ")
        self.setup_main_tab()
        
        # تب راهنما
        self.help_tab = QWidget()
        self.tabs.addTab(self.help_tab, "  راهنما   ")
        self.setup_help_tab()
        
        # نوار وضعیت
        self.status_bar = self.statusBar()
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximumSize(300, 20)
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #ddd;
                border-radius: 4px;
                text-align: center;
                font-family: Arial;
                font-size: 10pt;
            }
            QProgressBar::chunk {
                background: #4a6fa5;
            }
        """)
        self.status_bar.addPermanentWidget(self.progress_bar)
    
    def set_light_theme(self):
        self.setPalette(QApplication.style().standardPalette())
    
    def setup_main_tab(self):
        layout = QVBoxLayout()
        layout.setSpacing(15)
        self.main_tab.setLayout(layout)
        
        # گروه فایل‌های ورودی
        input_group = QGroupBox("فایل‌های ورودی")
        input_group.setStyleSheet("""
            QGroupBox {
                border: 1px solid #ddd;
                border-radius: 4px;
                margin-top: 10px;
                padding-top: 15px;
                font-family: Arial;
                font-size: 12pt;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
            }
        """)
        input_layout = QVBoxLayout()
        input_layout.setSpacing(10)
        input_group.setLayout(input_layout)
        
        self.excel_input = ModernFileInput("اکسل", "فایل‌های اکسل (*.xlsx *.xls)")
        self.xml_input = ModernFileInput("XML ورودی", "فایل‌های XML (*.xml)")
        self.output_input = ModernFileInput("XML خروجی", "فایل‌های XML (*.xml)")
        
        input_layout.addWidget(self.excel_input)
        input_layout.addWidget(self.xml_input)
        input_layout.addWidget(self.output_input)
        
        # دکمه پردازش
        self.process_btn = QPushButton("شروع پردازش")
        self.process_btn.setStyleSheet("""
            QPushButton {
                padding: 12px;
                background: #4a6fa5;
                color: white;
                border: none;
                border-radius: 4px;
                font-weight: bold;
                font-family: Arial;
                font-size: 14pt;
            }
            QPushButton:hover {
                background: #3a5a8f;
            }
            QPushButton:pressed {
                background: #2a4a7f;
            }
        """)
        self.process_btn.clicked.connect(self.process_files)
        
        # بخش پیش‌نمایش
        self.preview_label = QLabel("نتایج:")
        self.preview_label.setStyleSheet("font-family: Arial; font-size: 12pt;")
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setStyleSheet("""
            QTextEdit {
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 8px;
                background: white;
                font-family: Arial;
                font-size: 12pt;
            }
        """)
        
        # افزودن به لیآوت اصلی
        layout.addWidget(input_group)
        layout.addWidget(self.process_btn)
        layout.addWidget(self.preview_label)
        layout.addWidget(self.preview_text, 1)
    
    def setup_help_tab(self):
        layout = QVBoxLayout()
        self.help_tab.setLayout(layout)
        
        help_text = QTextEdit()
        help_text.setReadOnly(True)
        help_text.setHtml("""
        <div style='font-family:Arial; font-size:12pt; text-align:left; direction:rtl;'>
            <h1 style='color:#4a6fa5;'>راهنمای پردازشگر XML</h1>
            
            <h2>راهنمای استفاده:</h2>
            <ol>
                <li>فایل اکسل حاوی تنظیمات را انتخاب کنید</li>
                <li>فایل XML ورودی را انتخاب کنید</li>
                <li>مسیر ذخیره فایل XML خروجی را مشخص کنید</li>
                <li>دکمه "شروع پردازش" را کلیک کنید</li>
            </ol>
            
            <h2>ویژگی‌های برنامه:</h2>
            <ul>
                <li>نمایش صحیح متن‌های فارسی و عربی در خروجی</li>
                <li>تبدیل کاراکترهای خاص جایگزینی خودکار  & ، &lt; ، &gt; ، "و'  با معادل .</li>
                <li>تنظیم طول خط کنترل خودکار طول خطوط متن بر اساس حداقل و حداکثر تعیین‌شده.</li>
                <li>ذخیره تنظیمات به خاطر سپردن آخرین فایل‌های استفاده‌شده برای استفاده بعدی.</li>
            </ul>
            
            <h2>پشتیبانی:</h2>
            <p>در صورت بروز هرگونه مشکل فنی با ما تماس بگیرید.</p>
            <p>یوتیوب @LegendOfGreat_Battles</p>
            <p>تلگرام @LegendOfGreat_Battles</p>
            <p>وبلاگ legend-of-great-battles.blogfa.com</p>
        </div>
        """)
        help_text.setStyleSheet("""
            QTextEdit {
                border: none;
                padding: 10px;
                background: transparent;
            }
        """)
        
        layout.addWidget(help_text)
    
    def load_settings(self):
        self.excel_input.set_path(self.settings.value("excel_path", ""))
        self.xml_input.set_path(self.settings.value("xml_path", ""))
        self.output_input.set_path(self.settings.value("output_path", ""))
    
    def save_settings(self):
        self.settings.setValue("excel_path", self.excel_input.get_path())
        self.settings.setValue("xml_path", self.xml_input.get_path())
        self.settings.setValue("output_path", self.output_input.get_path())
    
    def validate_inputs(self):
        if not os.path.isfile(self.excel_input.get_path()):
            self.preview_text.setPlainText("خطا: لطفاً یک فایل اکسل معتبر انتخاب کنید.")
            return False
            
        if not os.path.isfile(self.xml_input.get_path()):
            self.preview_text.setPlainText("خطا: لطفاً یک فایل XML ورودی معتبر انتخاب کنید.")
            return False
            
        output_path = self.output_input.get_path()
        if not output_path:
            self.preview_text.setPlainText("خطا: لطفاً مسیر ذخیره فایل خروجی را مشخص کنید.")
            return False
            
        return True
    
    def process_files(self):
        self.preview_text.clear()
        
        if not self.validate_inputs():
            return
            
        reply = QMessageBox.question(
            self, "تأیید پردازش",
            "آیا مطمئن هستید که می‌خواهید پردازش را شروع کنید؟",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.No:
            return
            
        try:
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            
            # شبیه‌سازی پیشرفت پردازش
            for i in range(1, 101):
                self.progress_bar.setValue(i)
                QApplication.processEvents()
            
            # پردازش اصلی
            from xml_processor import process_excel_to_xml
            
            excel_path = self.excel_input.get_path()
            xml_path = self.xml_input.get_path()
            output_path = self.output_input.get_path()
            
            success, message = process_excel_to_xml(excel_path, xml_path, output_path)
            
            # نمایش نتیجه در بخش پیش‌نمایش
            self.preview_text.setPlainText(message)
            self.save_settings()
            
        except Exception as e:
            self.preview_text.setPlainText(f"خطا در پردازش:\n{str(e)}")
        finally:
            self.progress_bar.setVisible(False)
    
    def closeEvent(self, event):
        self.save_settings()
        event.accept()