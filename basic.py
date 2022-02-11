# _*_ coding: utf-8 _*_

# Comp Publish Tool
#
# Description : Shot 폴더 내의 플레이트 및 소스를 지정한 경로로 복사하는 툴

import os
import sys

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import fileCopyTool
import uiSetting_plate
import uiSetting_edit
import fileCopy

# global variables
OPT_PATH = os.getenv("OPT_PATH")

PWD = os.path.dirname(os.path.abspath(__file__))
uiFileName = "{}/fileCopyTool.ui".format(PWD)

from PySide2 import QtWidgets, QtUiTools, QtCore

class FileCopyTool():
    def __init__(self):
        self.ui = QtUiTools.QUiLoader().load(uiFileName, None)
        self.ui.show()

        self.version = fileCopyTool.__version__ # 버전
        self.proJson = "{}/TD/sceneopener/code/openersetCopy/teamInfo.json".format(OPT_PATH)
        self.plate_column_headers = ["ShotCode", "Plate", "r_Plate", "p_src", "src", "Error"]
        self.edit_column_headers = ["ShotCode", "Edit", "Error"]
        self.plate_column_idx_lookup = {"ShotCode": 0, "Plate": 1, "r_Plate": 2, "p_src": 3, "src": 4, "Error": 5}
        self.edit_column_idx_lookup = {"ShotCode": 0, "Edit": 1, "Error": 2}
        self.shotCodeDict = {}
        
        # ui 초기화
        self.ui.version_label.setText(self.version)
        uiSetting_plate.initUIFunc(self)
        uiSetting_edit.initUIFunc(self)

        # 플레이트 복사 탭 UI
        self.ui.plate_shotgun_radioButton.toggled.connect(lambda: uiSetting_plate.initUIForShotgunFunc(self))
        self.ui.plate_excel_radioButton.toggled.connect(lambda: uiSetting_plate.initUIForExcelFunc(self))
        self.ui.plate_project_comboBox.currentIndexChanged.connect(lambda: uiSetting_plate.setVendorComboBoxFunc(self))
        self.ui.plate_vendor_comboBox.currentIndexChanged.connect(lambda: uiSetting_plate.initTableFunc(self))
        self.ui.plate_pathToCopy_pushButton.pressed.connect(lambda: uiSetting_plate.openBrowserFunc(self))
        self.ui.plate_excelFile_pushButton.pressed.connect(lambda: uiSetting_plate.openBrowserForExcelFunc(self))
        self.ui.plate_excelLoad_pushButton.pressed.connect(lambda: uiSetting_plate.shotcodeSettingByExcelFunc(self))
        self.ui.plate_vendorLoad_pushButton.pressed.connect(lambda: uiSetting_plate.shotcodeSettingByVendorFunc(self))
        self.ui.plate_copy_pushButton.pressed.connect(lambda: fileCopy.palteCopyFunc(self))
        self.ui.plate_export_pushButton.pressed.connect(lambda: fileCopy.exportPlateExcelFunc(self))
        self.ui.plate_cancel_pushButton.pressed.connect(lambda: self.ui.close())

        # 편집본 복사 탭 UI
        self.ui.edit_shotgun_radioButton.toggled.connect(lambda: uiSetting_edit.initUIForShotgunFunc(self))
        self.ui.edit_excel_radioButton.toggled.connect(lambda: uiSetting_edit.initUIForExcelFunc(self))
        self.ui.edit_project_comboBox.currentIndexChanged.connect(lambda: uiSetting_edit.setVendorComboBoxFunc(self))
        self.ui.edit_vendor_comboBox.currentIndexChanged.connect(lambda: uiSetting_edit.initTableFunc(self))
        self.ui.edit_pathToCopy_pushButton.pressed.connect(lambda: uiSetting_edit.openBrowserFunc(self))
        self.ui.edit_excelFile_pushButton.pressed.connect(lambda: uiSetting_edit.openBrowserForExcelFunc(self))
        self.ui.edit_excelLoad_pushButton.pressed.connect(lambda: uiSetting_edit.shotcodeSettingByExcelFunc(self))
        self.ui.edit_vendorLoad_pushButton.pressed.connect(lambda: uiSetting_edit.shotcodeSettingByVendorFunc(self))
        self.ui.edit_copy_pushButton.pressed.connect(lambda: fileCopy.editCopyFunc(self))
        self.ui.edit_export_pushButton.pressed.connect(lambda: fileCopy.exportEditExcelFunc(self))
        self.ui.edit_cancel_pushButton.pressed.connect(lambda: self.ui.close())

if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_ShareOpenGLContexts)
    app = QtWidgets.QApplication(sys.argv)
    win = FileCopyTool()
    sys.exit(app.exec_())