import ttkbootstrap as tb
from ttkbootstrap.constants import *
import tkinter as tk

root = tb.Window(themename="minty")  # 테마만 바꿨는데 갑자기 예뻐짐
root.title("한전 원격검침 비교")
root.geometry("600x100")

progress = tb.Progressbar(root, orient='horizontal', mode='determinate', bootstyle=SUCCESS)
progress.pack(fill="x", padx=10, pady=5)

log = tb.Label(root, text="파일 - 열기 - ...", anchor="w")
log.pack(fill="x", padx=10)

root.mainloop()
