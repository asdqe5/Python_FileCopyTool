# _*_ coding: utf-8 _*_

# Comp Publish Tool
#
# Description : UI를 세팅하는 스크립트
#

import os
import sys
import json
import re
import xlrd

# global variables
TD_PATH = os.getenv("TD_PATH")

from PySide2 import QtWidgets, QtCore, QtGui

if "{}/api/pathapi".format(TD_PATH) not in sys.path:
    sys.path.append("{}/api/pathapi".format(TD_PATH))
import rdpath

if "{}/api" not in sys.path:
    sys.path.append('{}/api'.format(TD_PATH))

import shotgun_api3

sg = shotgun_api3.Shotgun("https://road101.shotgunstudio.com", script_name="authentication_script", api_key="vZnk#jbmopl6oqispvodazfyn")

def initUIFunc(self):
    '''
    UI 초기 세팅
    '''
    projectList, err = rdpath.projects()
    if err != None:
        dial = QtWidgets.QMessageBox()
        dial.setWindowTitle("Error")
        dial.setText(err)
        ok_btn = dial.addButton("OK", QtWidgets.QMessageBox.YesRole)
        dial.exec_()
        return

    if projectList == None:
        dial = QtWidgets.QMessageBox()
        dial.setWindowTitle("Error")
        dial.setText("프로젝트가 존재하지 않습니다.")
        ok_btn = dial.addButton("OK", QtWidgets.QMessageBox.YesRole)
        dial.exec_()
        return
        
    self.ui.edit_project_comboBox.clear()
    # json 파일에서 old 프로젝트 리스트를 가져옴
    f = open(self.proJson, "r")
    jsonData = json.load(f)
    f.close()
    jsonProjectList = []
    newProjectList = []
    for p in jsonData["nuke"]["projects"]:
        jsonProjectList.append(p["Name"])
    for p in projectList:
        if p not in jsonProjectList:
            newProjectList.append(p)
    self.ui.edit_project_comboBox.addItems([""] + newProjectList)

    self.ui.edit_pathtocopy_lineEdit.clear()
    self.ui.edit_pathtocopy_lineEdit.setReadOnly(True)
    self.ui.edit_excelFile_lineEdit.clear()
    self.ui.edit_excelFile_lineEdit.setReadOnly(True)

    # 시퀀스 샷 테이블 설정
    initTableFunc(self)
    self.ui.edit_fileCopy_tableWidget.setColumnCount(len(self.edit_column_headers))
    self.ui.edit_fileCopy_tableWidget.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
    self.ui.edit_fileCopy_tableWidget.setHorizontalHeaderLabels(self.edit_column_headers)
    self.ui.edit_fileCopy_tableWidget.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)
    self.ui.edit_fileCopy_tableWidget.horizontalHeader().setMinimumSectionSize(190)

    # 초기에는 엑셀로 설정이 되어있다고 생각
    self.ui.edit_excel_radioButton.setChecked(True)
    self.ui.edit_vendor_comboBox.setEnabled(False)
    self.ui.edit_vendorLoad_pushButton.setEnabled(False)

def initUIForShotgunFunc(self):
    '''
    shotgun 라디오 버튼을 눌렀을 경우
    shotgun 라디오 버튼이 체크되어있으면
    그것에 맞게 UI를 세팅하는 함수이디.
    '''
    self.ui.edit_project_comboBox.setCurrentIndex(0)
    self.ui.edit_pathtocopy_lineEdit.clear()
    self.ui.edit_vendor_comboBox.setEnabled(True)
    self.ui.edit_vendor_comboBox.clear()
    self.ui.edit_vendorLoad_pushButton.setEnabled(True)
    self.ui.edit_excelFile_lineEdit.clear()
    self.ui.edit_excelFile_pushButton.setEnabled(False)
    self.ui.edit_excelLoad_pushButton.setEnabled(False)
    initTableFunc(self)
    self.ui.edit_copy_pushButton.setEnabled(True)

def initUIForExcelFunc(self):
    '''
    excel 라디오 버튼을 눌렀을 경우
    excel 라디오 버튼이 체크되어있으면 
    그것에 맞게 UI를 세팅하는 함수이다
    '''
    self.ui.edit_project_comboBox.setCurrentIndex(0)
    self.ui.edit_pathtocopy_lineEdit.clear()
    self.ui.edit_vendor_comboBox.setEnabled(False)
    self.ui.edit_vendor_comboBox.clear()
    self.ui.edit_vendorLoad_pushButton.setEnabled(False)
    self.ui.edit_excelFile_lineEdit.clear()
    self.ui.edit_excelFile_pushButton.setEnabled(True)
    self.ui.edit_excelLoad_pushButton.setEnabled(True)
    initTableFunc(self)
    self.ui.edit_copy_pushButton.setEnabled(True)

def initTableFunc(self):
    '''
    샷코드 테이블을 초기화하는 
    함수이다.
    '''
    self.ui.edit_fileCopy_tableWidget.clearContents()
    while (self.ui.edit_fileCopy_tableWidget.rowCount() > 0):
        self.ui.edit_fileCopy_tableWidget.removeRow(0)
    self.ui.edit_fileCopy_tableWidget.setRowCount(0)

def setVendorComboBoxFunc(self):
    '''
    Shotgun 기준으로 체크가 되어있을 경우
    프로젝트를 선택했을 때 해당 프로젝트에
    샷리스트에 적힌 벤더 목록을 가져온다.
    '''
    initTableFunc(self)
    self.ui.edit_pathtocopy_lineEdit.clear()
    self.ui.edit_copy_pushButton.setEnabled(True)

    if self.ui.edit_excel_radioButton.isChecked():
        self.ui.edit_excelFile_lineEdit.clear()
        return

    self.ui.edit_vendor_comboBox.clear()
    if self.ui.edit_project_comboBox.currentText() == "":
        return

    projectName = self.ui.edit_project_comboBox.currentText()
    project_dict = sg.find_one('Project', [['name', 'is', projectName]])
    shotCodeList = sg.find('Shot', [['project', 'is', project_dict], ['sg_ven', "is_not", None]], ['code', 'sg_ven'])
    
    self.shotCodeDict = {}
    vendorList = []
    for shotCode in shotCodeList:
        # 샷코드가 규칙에 맞는지 확인
        if not re.match("s[0-9]+_c[0-9]+", shotCode['code']):
            if not re.match("s[0-9]+_c[a-z]+", shotCode['code']):
                continue
        
        vendors = shotCode['sg_ven']
        for v in vendors:
            if not v['name'] in vendorList:
                vendorList.append(v['name'])

            if not v['name'] in self.shotCodeDict:
                self.shotCodeDict[v['name']] = [shotCode['code']]
            else:
                self.shotCodeDict[v['name']].append(shotCode['code'])
    
    if vendorList == []:
        dial = QtWidgets.QMessageBox()
        dial.setWindowTitle("Error")
        dial.setText(u"Shotgun에서 {} 프로젝트의 벤더 정보가 존재하지 않습니다.".format(projectName))
        ok_btn = dial.addButton("OK", QtWidgets.QMessageBox.YesRole)
        dial.exec_()
        return

    self.ui.edit_vendor_comboBox.addItems([""] + vendorList)

def openBrowserFunc(self):
    '''
    파일 브라우저를 열어서 폴더 경로를 받아오는 함수
    '''
    folderPath = QtWidgets.QFileDialog.getExistingDirectory(None, "Select Folder", None, QtWidgets.QFileDialog.ShowDirsOnly)
    if folderPath == "":
        return
    
    # 받아온 경로를 lineEdit에 텍스트로 표시
    self.ui.edit_pathtocopy_lineEdit.setText(folderPath)

def openBrowserForExcelFunc(self):
    '''
    파일 브라우저를 열어서 엑셀 파일 경로를 받아오는 함수
    '''
    filePath = QtWidgets.QFileDialog.getOpenFileName(None, "Select Excel File", None, "xlsx files(*.xlsx)")
    if filePath[0] == "":
        return

    # 받아온 경로를 excel lineEdit에 텍스트로 표시
    self.ui.edit_excelFile_lineEdit.setText(filePath[0])

def shotcodeSettingByExcelFunc(self):
    '''
    입력 받은 엑셀 파일을 읽어서 
    테이블에 샷코드를 정리하는 함수
    '''
    excelPath = self.ui.edit_excelFile_lineEdit.text()
    if excelPath == "":
        dial = QtWidgets.QMessageBox()
        dial.setWindowTitle("Error")
        dial.setText(u"엑셀 파일을 먼저 선택해주세요.")
        ok_btn = dial.addButton("OK", QtWidgets.QMessageBox.YesRole)
        dial.exec_()
        return

    # tableWidget 초기화
    initTableFunc(self)

    workbook = xlrd.open_workbook(excelPath)
    worksheet = workbook.sheet_by_index(0)

    # 엑셀 파일에서 데이터 가져오기
    for row in range(worksheet.nrows):
        shotCode = worksheet.cell_value(row, 0)
        if not shotCode:
            continue

        rowPosition = self.ui.edit_fileCopy_tableWidget.rowCount()
        self.ui.edit_fileCopy_tableWidget.insertRow(rowPosition)

        # 샷코드 추가
        item = QtWidgets.QTableWidgetItem(shotCode)
        item.setTextAlignment(QtCore.Qt.AlignVCenter | QtCore.Qt.AlignHCenter)
        self.ui.edit_fileCopy_tableWidget.setItem(rowPosition, self.edit_column_idx_lookup["ShotCode"], item)
    
    self.ui.edit_fileCopy_tableWidget.resizeColumnsToContents()
    self.ui.edit_fileCopy_tableWidget.resizeRowsToContents()

    self.ui.edit_copy_pushButton.setEnabled(True)

def shotcodeSettingByVendorFunc(self):
    '''
    선택한 벤더를 기준으로
    샷코드 테이블에 샷코드를 
    세팅하는 함수이다.
    '''
    if self.ui.edit_project_comboBox.currentText() == "":
        dial = QtWidgets.QMessageBox()
        dial.setWindowTitle("Error")
        dial.setText(u"벤더 정보를 가져올 프로젝트를 먼저 선택해주세요.")
        ok_btn = dial.addButton("OK", QtWidgets.QMessageBox.YesRole)
        dial.exec_()
        return
    
    if self.ui.edit_vendor_comboBox.currentText() == "":
        dial = QtWidgets.QMessageBox()
        dial.setWindowTitle("Error")
        dial.setText(u"샷코드 정보를 가져올 벤더를 먼저 선택해주세요.")
        ok_btn = dial.addButton("OK", QtWidgets.QMessageBox.YesRole)
        dial.exec_()
        return
    
    vendor = self.ui.edit_vendor_comboBox.currentText()

    # tableWidget 초기화
    initTableFunc(self)

    shotCodeList = self.shotCodeDict[vendor.encode('utf8')]
    for shotCode in shotCodeList:
        rowPosition = self.ui.edit_fileCopy_tableWidget.rowCount()
        self.ui.edit_fileCopy_tableWidget.insertRow(rowPosition)

        # 샷코드 추가
        item = QtWidgets.QTableWidgetItem(shotCode)
        item.setTextAlignment(QtCore.Qt.AlignVCenter | QtCore.Qt.AlignHCenter)
        self.ui.edit_fileCopy_tableWidget.setItem(rowPosition, self.edit_column_idx_lookup["ShotCode"], item)
    
    self.ui.edit_fileCopy_tableWidget.resizeColumnsToContents()
    self.ui.edit_fileCopy_tableWidget.resizeRowsToContents()

    self.ui.edit_copy_pushButton.setEnabled(True)
