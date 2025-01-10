import sys

from PySide6.QtCore import Qt, QSettings, QCoreApplication
from PySide6.QtWidgets import QApplication

from windows.main_window import MainWindow

if __name__ == "__main__":
    app = QApplication(sys.argv)
    QCoreApplication.setOrganizationName("Obsidian Exporter")
    QCoreApplication.setApplicationName("Obsidian Exporter")

    # Show the main window
    main_window = MainWindow()
    main_window.show()

    app.exec()
