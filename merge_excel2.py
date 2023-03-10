import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk

root = tk.Tk()
root.geometry('400x200')
root.title("엑셀 시트, 파일 결합")

# 파일 선택 대화상자 열기
file_paths = filedialog.askopenfilenames()

# 프로그레스바 생성
progress = ttk.Progressbar(root, orient='horizontal', mode='determinate')
progress.grid(row=1, column=0, sticky='ew', padx=100 , pady=100)

# 파일 결합 중에 작업중입니다 메세지 띄우기, 이것도 안뜸...
message_label = ttk.Label(root, text='파일 결합 중입니다...')
message_label.grid(row=0, column=0, sticky='w', padx=100, pady=100)

# 프로그레스바 maximum 값 설정
progress['maximum'] = 100

# 창의 크기가 변경될 때, 프로그레스바가 계속 중앙에 위치하도록 설정
root.rowconfigure(0, weight=1)
root.rowconfigure(2, weight=1)
root.columnconfigure(0, weight=1)

# root.update() 함수 호출
root.update_idletasks()

# 파일 결합
all_dfs = []
for i, file_path in enumerate(file_paths):
    xls = pd.ExcelFile(file_path, engine='openpyxl')
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        all_dfs.append(df)
    # 작업이 1/n 완료되었음을 프로그레스바에 업데이트
    progress['value'] = (i+1) * 100 / len(file_paths)
    progress.update()
    root.update_idletasks() 

# 결과 출력, 이것도 안뜸
message_label.config(text='결과를 저장하는 중입니다...')
merged_df = pd.concat(all_dfs, axis=0, ignore_index=True)
file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
merged_df.to_excel(file_path, index=False)

# 완료 메시지 출력, 이거 안뜸
message_label.config(text='작업이 완료되었습니다.')
progress.destroy()

# GUI 창 닫기
root.quit()

