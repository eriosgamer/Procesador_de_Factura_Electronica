from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QPushButton,
    QFileDialog,
    QLabel,
    QMessageBox,
    QProgressBar,
)
from PySide6.QtCore import Qt
import sys
import os

import main  # Importa la lógica como librería
import extract_attachments  # Importa el script de extracción


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Factura Electrónica - Procesador")
        self.setMinimumSize(400, 250)
        layout = QVBoxLayout()

        self.label = QLabel("Procesar carpeta de facturas:")
        layout.addWidget(self.label)

        self.btn_facturas = QPushButton("Seleccionar y procesar facturas")
        self.btn_facturas.clicked.connect(self.procesar_facturas)
        layout.addWidget(self.btn_facturas)

        self.label_eml = QLabel("Procesar carpeta de correos:")
        layout.addWidget(self.label_eml)

        self.btn_eml = QPushButton("Seleccionar y procesar correos")
        self.btn_eml.clicked.connect(self.procesar_eml)
        layout.addWidget(self.btn_eml)

        self.file_label = QLabel("")
        layout.addWidget(self.file_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        self.setLayout(layout)

    def procesar_facturas(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Seleccione carpeta de facturas"
        )
        if not folder:
            return
        try:

            def progress_callback(idx, total, fname):
                self.file_label.setText(f"Procesando: {fname}")
                self.progress_bar.setMaximum(total)
                self.progress_bar.setValue(idx)
                self.progress_bar.setFormat(f"{idx}/{total}")
                QApplication.processEvents()

            main.main_procesar_facturas(folder, progress_callback)
            self.file_label.setText("")
            QMessageBox.information(self, "Procesamiento", "Procesamiento terminado.")
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat("")
        except Exception as e:
            self.file_label.setText("")
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat("")
            QMessageBox.critical(self, "Error", f"Ocurrió un error:\n{str(e)}")

    def procesar_eml(self):
        folder = QFileDialog.getExistingDirectory(self, "Seleccione carpeta de correos")
        if not folder:
            return
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            output_folder = os.path.join(base_dir, "extraidos")
            eml_files = [f for f in os.listdir(folder) if f.lower().endswith(".eml")]
            total = len(eml_files)
            self.progress_bar.setMaximum(total)
            self.progress_bar.setValue(0)

            def progress_callback(idx, total, fname):
                self.file_label.setText(f"Procesando: {fname}")
                self.progress_bar.setValue(idx)
                self.progress_bar.setFormat(f"{idx}/{total}")
                QApplication.processEvents()

            extract_attachments.extract_attachments(
                folder, output_folder, progress_callback=progress_callback
            )
            self.file_label.setText("")
            QMessageBox.information(
                self,
                "Procesamiento de correos",
                f"Extracción de adjuntos terminada.\nArchivos guardados en: {output_folder}",
            )
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat("")
        except Exception as e:
            self.file_label.setText("")
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat("")
            QMessageBox.critical(self, "Error", f"Ocurrió un error:\n{str(e)}")


if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        window = MainWindow()
        window.show()
        sys.exit(app.exec())
    except KeyboardInterrupt:
        QMessageBox.warning(None, "Advertencia", "Proceso cancelado por el usuario.")
        sys.exit(0)
    except Exception as e:
        QMessageBox.critical(None, "Error", f"Ocurrió un error:\n{str(e)}")
