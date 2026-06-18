import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import shutil

# 변경할 서체 목록: (찾을 서체, 바꿀 서체)
FONT_MAP = [
    ("맑은 고딕", "함초롬돋움"),
    ("바탕",      "함초롬바탕"),
]

def patch_font_data(data, src_font, dst_font):
    """
    한글 파일 내부의 바이트 패턴을 찾아 안전하게 교체합니다.
    """
    src_bytes = src_font.encode("utf-16-le")
    dst_bytes = dst_font.encode("utf-16-le")
    
    # 두 글꼴 이름의 바이트 길이가 다를 경우 빈 바이트를 채워 정렬을 맞춤
    if len(src_bytes) > len(dst_bytes):
        dst_bytes = dst_bytes + b"\x00" * (len(src_bytes) - len(dst_bytes))
    elif len(src_bytes) < len(dst_bytes):
        src_bytes = src_bytes + b"\x00" * (len(dst_bytes) - len(src_bytes))
        
    if src_bytes in data:
        return data.replace(src_bytes, dst_bytes), True
    return data, False

def change_font_in_file_direct(filepath):
    with open(filepath, "rb") as f:
        data = f.read()
        
    replaced_any = False
    for src_font, dst_font in FONT_MAP:
        data, r1 = patch_font_data(data, src_font, dst_font)
        # 특수 속성 구문도 함께 패치
        data, r2 = patch_font_data(data, f"{src_font},TrueType", f"{dst_font},TrueType")
        
        if r1 or r2:
            replaced_any = True
            
    if replaced_any:
        with open(filepath, "wb") as f:
            f.write(data)
    return replaced_any

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("HWP 서체 일괄 변경 (최종 수정본)")
        self.resizable(False, False)
        self.configure(bg="#F7F8FA")
        self._build_ui()
        self._center_window(480, 500)

    def _center_window(self, w, h):
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def _build_ui(self):
        PAD = 24
        BG, CARD = "#F7F8FA", "#FFFFFF"
        ACCENT, GRAY, BORDER = "#5B6EF5", "#6B7280", "#E5E7EB"
        
        tf = tk.Frame(self, bg=BG)
        tf.pack(fill="x", padx=PAD, pady=(PAD, 8))
        tk.Label(tf, text="HWP 서체 일괄 변경", font=("맑은 고딕", 14, "bold"), bg=BG).pack(anchor="w")
        
        card = tk.Frame(self, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        card.pack(fill="x", padx=PAD, pady=4)

        tk.Label(card, text="📁 변경할 폴더", font=("맑은 고딕", 10), bg=CARD).pack(anchor="w", padx=16, pady=(14,6))
        fr = tk.Frame(card, bg=CARD)
        fr.pack(fill="x", padx=16, pady=(0,14))
        self.folder_var = tk.StringVar()
        tk.Entry(fr, textvariable=self.folder_var, font=("맑은 고딕", 10), bg="#F3F4F6", relief="flat").pack(side="left", fill="x", expand=True, ipady=6, padx=(0,8))
        tk.Button(fr, text="선택", command=lambda: self.folder_var.set(filedialog.askdirectory()), bg=ACCENT, fg="white", bd=0, padx=10).pack(side="right")

        self.run_btn = tk.Button(self, text="변경 시작", command=self._run, bg=ACCENT, fg="white", font=("맑은 고딕", 11, "bold"), pady=10, bd=0)
        self.run_btn.pack(fill="x", padx=PAD, pady=20)
        
        self.status_var = tk.StringVar(value="대기 중...")
        tk.Label(self, textvariable=self.status_var, bg=BG).pack()

    def _run(self):
        folder = self.folder_var.get()
        if not folder: return
        hwp_files = [f for f in os.listdir(folder) if f.lower().endswith(".hwp")]
        
        self.run_btn.config(state="disabled")
        threading.Thread(target=self._worker, args=(folder, hwp_files), daemon=True).start()

    def _worker(self, folder, files):
        success = 0
        for fname in files:
            fpath = os.path.join(folder, fname)
            shutil.copy2(fpath, fpath + ".bak") # 안전한 백업
            try:
                if change_font_in_file_direct(fpath):
                    success += 1
            except Exception as e:
                print(f"Error: {e}")
        self.after(0, lambda: messagebox.showinfo("완료", f"성공적으로 처리되었습니다. ({success}개 파일)"))
        self.after(0, lambda: self.run_btn.config(state="normal"))

if __name__ == "__main__":
    app = App()
    app.mainloop()
