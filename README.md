# merge_excel
헤더가 동일할 때, 엑셀파일(.xlsx) 또는 시트 결합하는 파일


# 실행법  윈도우 사용 cmd
.py 파일과 동일 폴더에서
pyinstaller -F merge_excel2.py --hidden-import pandas --hidden-import tkinter --hidden-import openpyxl


# 실행법
1. dist 폴더
2. 실행

# 주의점
1. 엑셀(*.xlsx)에 가능합니다.
2. 시트와 엑셀이 모두 동일한 헤더를 가져야 합니다.
 - 동일하지 않을 시, 원하는 형태로 결합이 되지 않습니다.
