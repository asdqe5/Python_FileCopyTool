# File Copy Tool
<br>

#### [소스 코드]

##### 메인 코드
- [basic.py](../basic.py) : FileCopyTool 클래스가 구현되어 있는 메인 코드
   - 기본적으로 실행을 할 때, ui 파일을 로드하여 보여준다.
   - UI 테이블을 구성하기 위한 컬럼에 대한 정의가 되어있다.
   - 버튼 클릭, radio 버튼 토글 변경, 콤보 박스 변경에 따른 시그널 함수들이 정의되어 있다.

##### 서브 코드
- [uiSetting_plate.py](../uiSetting_plate.py) : Plate 탭 UI를 세팅하는 코드
  - pathapi, shotgunapi 로드를 해온다.
  - 시그널이 발생하면 실행되는 함수들이 정의되어 있다.
  - 초기에 UI를 세팅하는 함수(initUIFunc)에서 rdpath.py 의 projects() 함수를 사용한다.
  - shotgun api를 이용하여 선택한 프로젝트의 샷건에 올라온 샷 정보를 읽고 벤더 정보를 가져온다.
  ```
  projectName = self.ui.plate_project_comboBox.currentText()
  project_dict = sg.find_one('Project', [['name', 'is', projectName]])
  shotCodeList = sg.find('Shot', [['project', 'is', project_dict], ['sg_ven', "is_not", None]], ['code', 'sg_ven'])
  ```
  - 엑셀 파일을 읽어오기 위해서 xlrd 라이브러리를 사용한다.
  ```
  import xlrd

  workbook = xlrd.open_workbook(excelPath)
  worksheet = workbook.sheet_by_index(0)
  ```

- [uiSetting_edit.py](../uiSetting_edit.py) : Edit 탭 UI를 세팅하는 코드

- [fileCopy.py](../fileCopy.py) : 세팅된 테이블 데이터를 기준으로 파일을 복사하거나 엑셀 추출을 하는 코드
  - 복사를 시작할 때 UI내의 버튼과 같은 툴 메뉴들을 비활성화한다.
  - 에러가 나거나 정상적으로 복사가 완료되었을 때, 아래의 코드를 통해 실시간으로 이벤트를 보여준다.
  ```
  QtCore.QCoreApplication.processEvents()
  ```
  - 플레이트, 리타임 플레이트 / 소스 / 플레이트 소스를 복사할 떄, 가장 높은 버전의 파일만 복사를 한다.
  - 해당 경로에 복사하는 버전보다 낮은 버전의 파일이 남아있다면 그 파일은 삭제한다.
  ```
  for cp in copiedPlate:
    version = int(cp.split("v")[1])
    if version < maxValue:
        shutil.rmtree(os.path.join(os.path.dirname(pathToCopyPlate), cp))
  ```
  - 테이블 데이터 엑셀 추출을 위해 openpyxl 라이브러리를 사용한다.
  ```
  import openpyxl

  # 엑셀 만들기
  excelFileName = "{}_plate_copy.xlsx".format(projectName)
  excelFilePath = os.path.join(pathToCopy, excelFileName)
  
  write_wb = openpyxl.Workbook()
  write_ws = write_wb.create_sheet('plate_copy')
  write_ws = write_wb.active
  ```




     