# Импортируем нужный объект из библиотеки
from docxtpl import DocxTemplate
import pandas as pd

import sys
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QFileDialog,
    QGridLayout,
    QPushButton,
    QLabel,
    QMessageBox,


)




class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('WordfromCSV')
        self.setGeometry(500, 500, 500, 450)
        self.layout = QGridLayout()
        self.setLayout(self.layout)

        self.btnWordTmpl = QPushButton("1. Выберите шаблон Word")
        self.lblWordTmpl = QLabel("шаблон Word:")

        self.layout.addWidget(self.btnWordTmpl)
        self.layout.addWidget(self.lblWordTmpl)
        self.btnWordTmpl.clicked.connect(self.openWordTmpl)

        self.btnCsvFile = QPushButton("2. Выберите CSV файл для слияния:")
        self.lblCsvFile = QLabel("CSV для слияния:")
        self.btnCsvFile.setEnabled(False)
        self.layout.addWidget(self.btnCsvFile)
        self.layout.addWidget(self.lblCsvFile)
        self.btnCsvFile.clicked.connect(self.openCsvFile)

        self.btnStart = QPushButton("3. СТАРТ")
        self.labelStart = QLabel("")
        self.labelLoad = QLabel("Сформирвано:")
        self.labelNotLoad = QLabel("not loaded, record exist:")
        self.labelProgress = QLabel("Всего в списке:")

        self.btnStart.setEnabled(False)
        self.layout.addWidget(self.btnStart)
        self.btnStart.clicked.connect(self.start1)


        self.WordTmplfilename=None
        self.CsvFilefilename=None

    def openWordTmpl(self):
        self.WordTmplfilename = QFileDialog.getOpenFileName(self, "Open File", "", "шаблон  (*.xlsx)")[0]
        self.lblWordTmpl.setText(u"file: " + self.WordTmplfilename)
        self.btnCsvFile.setEnabled(True)

    def openCsvFile(self):
        self.CsvFilefilename = QFileDialog.getOpenFileName(self, "Open File", "", "CSV шаблон  (*.csv)")[0]
        self.lblCsvFile.setText(u"file: " + self.CsvFilefilename)
        self.btnStart.setEnabled(True)

    def start1(self):
        QApplication.setOverrideCursor(Qt.WaitCursor)


        j = 0
        self.btnStart.setEnabled(False)
        self.labelStart.setText(u"Начали загрузку...")
        # Загрузка шаблона
        doc = DocxTemplate(self.WordTmplfilename)
        file_name=self.WordTmplfilename+'_1_'

        df = pd.read_csv(self.CsvFilefilename,delimiter=';',encoding='windows-1251')
        file_name_number = 1

        for index, row in df.iterrows():
            #for col in df.columns:
            # value = df.loc[index,col]
            doc.render(row)

            # Сохранение документа
            doc.save(file_name + str(file_name_number)  +  '.xlsx')
            file_name_number = file_name_number + 1
            j=j+1
        QApplication.restoreOverrideCursor()
        msgBox = QMessageBox()
        msgBox.setText("Закончили, всего в списке:"+" " +str(j) )
        msgBox.exec()


if __name__ == "__main__":

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())






