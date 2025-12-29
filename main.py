from gui import XmlProcessorGUI
import sys
from PyQt5.QtWidgets import QApplication

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    # تنظیم فونت پیش‌فرض
    font = QApplication.font()
    font.setFamily("Arial")
    font.setPointSize(10)
    app.setFont(font)
    
    window = XmlProcessorGUI()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()