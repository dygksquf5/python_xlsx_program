from tkinter import *
import tkinter.messagebox as message
import tkinter.ttk as ttk
from tkinter import filedialog
import csv
from xlsxwriter.workbook import Workbook


root = Tk()
root.title("GUI_yosuniiiii")


def add_file():
    files = filedialog.askopenfilenames(title="csv 파일을 선택하세요",
                                        filetypes=(("CVS 파일", "*.csv"),
                                                   ("모든파일", "*.*")),
                                        initialdir="/Users/yosuniiiii/Downloads")

    for file in files:
        list_file.insert(END, file)


# 선택삭제
def del_file():
    for index in reversed(list_file.curselection()):
        list_file.delete(index)


# 저장경로 (폴더)
def browse_dest_path():
    folder_selected = filedialog.askdirectory()
    txt_dest_path.delete(0, END) # 엔트리로 만들어서 0 END 라고 하면됨
    txt_dest_path.insert(0, folder_selected)


def changing(file_path):
    wb = Workbook(file_path[:-4] + '.xlsx')
    ws = wb.add_worksheet()
    with open(file_path, 'rt', encoding='utf8') as f:
        reader = csv.reader(f)
        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                ws.write(r, c, col)
    wb.close()

# 시작
def start():
    # 파일목록 확인하는 구간
    if list_file.size() == 0:
        message.showwarning("!", "파일을 추가하세요! ")
        return
    try:
        num = 0
        for x in list_file.get(0, END):
            num += 1
            progress = (num + 1) / int(list_file.size()) * 100  # 퍼센트 정보 계산
            p_bar.set(progress)
            progress_bar.update()
            changing(x)
        message.showinfo("알림", "작업이 완료되었습니다! ")
    except Exception as err:
        message.showerror("에러", err)



# 파일 프레임 (파일추가, 선택삭제)
file_frame = Frame(root)
file_frame.pack(fill="x", padx=5, pady=5) # pad 로 간격주기

btn_add_file = Button(file_frame, padx=5, pady=5, width=12, text="파일추가", command=add_file)
btn_add_file.pack(side="left")

btn_del_file = Button(file_frame, padx=5, pady=5, width=12, text="선택 삭제", command=del_file)
btn_del_file.pack(side="right")


# 리스트 프레임
list_frame = Frame(root)
list_frame.pack(fill="both", padx=5, pady=5)

scrollbar = Scrollbar(list_frame)
scrollbar.pack(side="right", fill="y")


list_file = Listbox(list_frame, selectmode="extended", height=15, yscrollcommand=scrollbar.set)
list_file.pack(side="left", fill="both", expand=True)
scrollbar.config(command=list_file.yview)



# 저장경로 프레임!
path_frame = LabelFrame(root, text="")
path_frame.pack(fill="x" , padx=5, pady=5, ipady=3)

txt_dest_path = Entry(path_frame)
txt_dest_path.pack(side="left", fill="x", expand=True, padx=5, pady=5, ipady=2) # i pad y , 즉 i = inner pad의 길이를 늘린다 y 축 길이를 늘릴거야!
btn_dest_path = Button(path_frame, text="저장경로지정불가", width=10, command=browse_dest_path)
btn_dest_path.pack(side="right",padx=5, pady=5)



# 진행상황 progress bar
frame_progress = LabelFrame(root, text="진행상황")
frame_progress.pack(fill="x", padx=5, pady=5, ipady=3)

p_bar = DoubleVar()
progress_bar = ttk.Progressbar(frame_progress, maximum=100, variable=p_bar)
progress_bar.pack(fill="x",padx=5, pady=5)

# 실행프레임
frame_run = Frame(root)
frame_run.pack(fill="x", padx=5, pady=5)


btn_close = Button(frame_run, padx=5, pady=5, text="닫기", width=12, command=root.quit)
btn_close.pack(side="right",padx=5, pady=5)

btn_start = Button(frame_run, padx=5, pady=5, text="시작", width=12, command=start)
btn_start.pack(side="right",padx=5, pady=5)



root.resizable(False, False)
root.mainloop()




