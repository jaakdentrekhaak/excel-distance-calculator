from PyQt5.QtWidgets import QWidget, QApplication, QPushButton, QLabel, QLineEdit, QFileDialog
import sys
import os
from PyQt5 import QtGui, QtCore
from PyQt5.QtGui import QCursor, QDesktopServices, QColor

import webscraper


class Window(QWidget):
    def __init__(self):
        super().__init__()

        self.title = "Excel distance and time calculator by Jens Putzeys"

        self.path = None

        self.screen_dim = (1600, 900)

        self.width = 550
        self.height = 340

        self.left = int(self.screen_dim[0] / 2 - self.width / 2)
        self.top = int(self.screen_dim[1] / 2 - self.height / 2)

        self.init_window()

    def init_window(self):
        # self.setWindowIcon(QtGui.QIcon('templates/youtube_logo.png'))
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.setStyleSheet('background-color: rgb(52, 50, 51);')

        self.create_layout()

        self.show()

    def start(self):

        allright = True

        if self.path is None:
            self.label8.setText('Select an excel file.')
            allright = False

        path_workbook = self.path

        if len(self.line_edit_sheet_name.text()) == 0 or len(self.line_edit_column1.text()) == 0 or len(
                self.line_edit_column2.text()) == 0 or len(self.line_edit_column3.text()) == 0 or len(
            self.line_edit_column4.text()) == 0:
            self.label8.setText('Fill in all the fields.')
            allright = False
        name_worksheet = self.line_edit_sheet_name.text()
        column1 = self.line_edit_column1.text()
        column2 = self.line_edit_column2.text()
        column1_insert = self.line_edit_column3.text()
        column2_insert = self.line_edit_column4.text()
        if len(self.line_edit_row1.text()) == 0:
            self.label8.setText('Row begin must be an integer.')
            allright = False
        else:
            beginning_row = int(self.line_edit_row1.text())
        if len(self.line_edit_row2.text()) == 0:
            self.label8.setText('Row end must be an integer.')
            allright = False
        else:
            ending_row = int(self.line_edit_row2.text())

        if allright:
            webscraper.webscrape(path_workbook, name_worksheet, column1, column2, column1_insert, column2_insert,
                                 beginning_row, ending_row)



    def get_path(self):
        # self.path = str(QFileDialog.getExistingDirectory(self, "Select Directory"))  # e.g. C:/Users/jensb/Desktop
        # self.path = str(QFileDialog.getOpenFileName(self, "Select file"))
        self.path = str(QFileDialog.getOpenFileName(self)[0])
        index = self.path.rfind(r'/')  # rfind: find last occurrence
        filename = self.path[index + 1:]
        # Make it display the name of the file
        self.label6.setText(filename)

    def create_layout(self):
        # buttons
        button_dim = (80, 40)
        window_dim = (self.width, self.height)

        self.button_start = QPushButton('Start', self)
        self.button_start.setGeometry(window_dim[0] / 2 - 35, 260,
                                      button_dim[0], button_dim[1])
        self.button_start.setStyleSheet("color: rgb(52, 50, 51); background-color: rgb(255, 209, 82);"
                                        "border-width: 1.5px; border-radius: 5px; border-color: rgb(255, 209, 82);"
                                        "font-size: 20px; font-family: Verdana;")
        self.button_start.clicked.connect(self.start)
        self.button_start.setCursor(QCursor(QtCore.Qt.PointingHandCursor))

        self.button_path = QPushButton('Select Excel File', self)
        self.button_path.setGeometry(window_dim[0] / 2 - 190 / 2, 170,
                                     200, button_dim[1])
        self.button_path.setStyleSheet("color: rgb(52, 50, 51); background-color: rgb(255, 209, 82);"
                                       "border-width: 1.5px; border-radius: 5px; border-color: rgb(255, 209, 82);"
                                       "font-size: 20px; font-family: Verdana;")
        self.button_path.clicked.connect(self.get_path)
        self.button_path.setCursor(QCursor(QtCore.Qt.PointingHandCursor))

        # line edits
        line_edit_dim = (self.width - 50, 30)

        self.line_edit_column1 = QLineEdit(self)
        self.line_edit_column1.setGeometry(50, 30, 30, 30)
        self.line_edit_column1.setStyleSheet("color: beige; background-color: rgb(52, 50, 51); border-width: 1.5px; "
                                             "border-radius: 0px; border-color: rgb(255, 209, 82); font-size: 20px; "
                                             "font-family: Verdana; border-style: solid")

        self.line_edit_column2 = QLineEdit(self)
        self.line_edit_column2.setGeometry(190, 30, 30, 30)
        self.line_edit_column2.setStyleSheet("color: beige; background-color: rgb(52, 50, 51); border-width: 1.5px; "
                                             "border-radius: 0px; border-color: rgb(255, 209, 82); font-size: 20px; "
                                             "font-family: Verdana; border-style: solid")

        self.line_edit_column3 = QLineEdit(self)
        self.line_edit_column3.setGeometry(330, 30, 30, 30)
        self.line_edit_column3.setStyleSheet("color: beige; background-color: rgb(52, 50, 51); border-width: 1.5px; "
                                             "border-radius: 0px; border-color: rgb(255, 209, 82); font-size: 20px; "
                                             "font-family: Verdana; border-style: solid")

        self.line_edit_column4 = QLineEdit(self)
        self.line_edit_column4.setGeometry(470, 30, 30, 30)
        self.line_edit_column4.setStyleSheet("color: beige; background-color: rgb(52, 50, 51); border-width: 1.5px; "
                                             "border-radius: 0px; border-color: rgb(255, 209, 82); font-size: 20px; "
                                             "font-family: Verdana; border-style: solid")

        self.line_edit_row1 = QLineEdit(self)
        self.line_edit_row1.setGeometry(50 + 30, 120, 30, 30)
        self.line_edit_row1.setStyleSheet("color: beige; background-color: rgb(52, 50, 51); border-width: 1.5px; "
                                          "border-radius: 0px; border-color: rgb(255, 209, 82); font-size: 20px; "
                                          "font-family: Verdana; border-style: solid")

        self.line_edit_row2 = QLineEdit(self)
        self.line_edit_row2.setGeometry(190 + 30, 120, 30, 30)
        self.line_edit_row2.setStyleSheet("color: beige; background-color: rgb(52, 50, 51); border-width: 1.5px; "
                                          "border-radius: 0px; border-color: rgb(255, 209, 82); font-size: 20px; "
                                          "font-family: Verdana; border-style: solid")

        self.line_edit_sheet_name = QLineEdit(self)
        self.line_edit_sheet_name.setGeometry(190 + 140, 120, 150, 30)
        self.line_edit_sheet_name.setStyleSheet("color: beige; background-color: rgb(52, 50, 51); border-width: 1.5px; "
                                                "border-radius: 0px; border-color: rgb(255, 209, 82); font-size: 20px; "
                                                "font-family: Verdana; border-style: solid")

        # labels
        self.label1 = QLabel(self)
        self.label1.setText('Column cities from')
        self.label1.setGeometry(0, 0, 150, 30)
        self.label1.setStyleSheet('color: beige; font-size: 13px; font-family: Verdana')

        self.label2 = QLabel(self)
        self.label2.setText('Column cities to')
        self.label2.setGeometry(150, 0, 150, 30)
        self.label2.setStyleSheet('color: beige; font-size: 13px; font-family: Verdana')

        self.label3 = QLabel(self)
        self.label3.setText('Column distance')
        self.label3.setGeometry(300, 0, 150, 30)
        self.label3.setStyleSheet('color: beige; font-size: 13px; font-family: Verdana')

        self.label4 = QLabel(self)
        self.label4.setText('Column time')
        self.label4.setGeometry(450, 0, 150, 30)
        self.label4.setStyleSheet('color: beige; font-size: 13px; font-family: Verdana')

        self.label5 = QLabel(self)
        self.label5.setText('Row begin')
        self.label5.setGeometry(30 + 30, 80, 150, 30)
        self.label5.setStyleSheet('color: beige; font-size: 13px; font-family: Verdana')

        self.label5 = QLabel(self)
        self.label5.setText('Row end')
        self.label5.setGeometry(180 + 30, 80, 150, 30)
        self.label5.setStyleSheet('color: beige; font-size: 13px; font-family: Verdana')

        self.label7 = QLabel(self)
        self.label7.setText('Working sheet name')
        self.label7.setGeometry(340, 80, 150, 30)
        self.label7.setStyleSheet('color: beige; font-size: 13px; font-family: Verdana')

        self.label6 = QLabel(self)
        self.label6.setText('')
        self.label6.setGeometry(180, 220, line_edit_dim[0], 20)
        self.label6.setStyleSheet('color: beige; font-size: 15px; font-family: Verdana')

        self.label8 = QLabel(self)
        self.label8.setText('')
        self.label8.setGeometry(0, 300, line_edit_dim[0], 20)
        self.label8.setStyleSheet('color: beige; font-size: 15px; font-family: Verdana')


if __name__ == '__main__':
    App = QApplication(sys.argv)
    window = Window()
    sys.exit(App.exec())
