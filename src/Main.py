from View import View
from Handlers import *

GUI = View("600x110")

GUI.addField("사업자 정보 엑셀", "파일 선택", onExcelButtonClick, 0)
GUI.addField("통판 사업자 엑셀", "파일 선택", onExcelButtonClick, 1)
GUI.addField("한글 상장 서식", "파일 선택", onHWPButtonClick, 2)

GUI.addRunButton("매크로 실행", onMacroButtonClick, 3)

GUI.run()