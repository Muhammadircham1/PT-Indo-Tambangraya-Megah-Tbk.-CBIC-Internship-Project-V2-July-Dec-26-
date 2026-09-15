from PyQt6 import QtCore, QtGui, QtWidgets


class Ui_dialog(object):
    def setupUi(self, dialog, status="success"):
        dialog.setObjectName("dialog")
        dialog.resize(298, 100)
        self.label = QtWidgets.QLabel(parent=dialog)
        self.label.setGeometry(QtCore.QRect(30, 20, 251, 31))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        self.label.setFont(font)
        self.label.setObjectName("label")

        self.pushButton_ok = QtWidgets.QPushButton(parent=dialog)
        self.pushButton_ok.setGeometry(QtCore.QRect(110, 60, 75, 31))
        font_btn = QtGui.QFont()
        font_btn.setFamily("Arial")
        font_btn.setPointSize(9)
        font_btn.setBold(True)
        self.pushButton_ok.setFont(font_btn)
        self.pushButton_ok.setObjectName("pushButton_ok")

        self.retranslateUi(dialog, status)
        QtCore.QMetaObject.connectSlotsByName(dialog)

    def retranslateUi(self, dialog, status):
        _translate = QtCore.QCoreApplication.translate

        if status == "success":
            dialog.setWindowTitle(_translate("dialog", "Finish"))
            self.label.setText(_translate("dialog", "Automation Process Completed"))
            self.label.setStyleSheet("color: green;")
        elif status == "error":
            dialog.setWindowTitle(_translate("dialog", "Error"))
            self.label.setText(_translate("dialog", "Automation Process Error"))
            self.label.setStyleSheet("color: red;")

        self.pushButton_ok.setText(_translate("dialog", "OK"))