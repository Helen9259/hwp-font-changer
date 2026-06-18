import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import shutil
import time  # 👈 안정적인 대기를 위해 추가

FONT_MAP = [
    ("맑은 고딕", "함초롬돋움"),
    ("바탕",      "함초롬바탕"),
]

try:
    import win32com.client
    import win32com.client.gencache as gencache
    HWP_AVAILABLE = True
except ImportError:
    HWP_AVAILABLE = False


def check_hwp():
    if not HWP_AVAILABLE:
        messagebox.showerror("오류", "pywin32가 설치되어 있지 않습니다.\npip install pywin32 명령어로 설치해주세요.")
        return False
    try:
        try:
            hwp = gencache.EnsureDispatch("HWPFrame.HwpObject")
        except Exception:
            hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
        hwp.Quit()
        return True
    except Exception as e:
        messagebox.showerror("한글 프로그램 없음", f"한컴 한글(HWP) 연동 실패.\n\n[상세 에러]: {e}")
        return False


def change_font_in_file(hwp, filepath, src_font, dst_font):
    # 파일 열기 시도
    open_res = hwp.Open(filepath, "HWP", "forceopen:true")
    if not open_res:
        raise Exception(f"한글 파일을 열 수 없습니다.\n경로: {filepath}")

    # 한글 내부 액션 엔진이 준비될 시간을 0.2초 줌 (NoneType 에러 방지)
    time.sleep(0.2)

    act = hwp.CreateAction("FindReplace")
    if not act:
        raise Exception("한글 'FindReplace' 액션을 생성하지 못했습니다. 한글 프로그램 상태를 확인하세요.")

    param = act.CreateSet()
    if not param:
        raise Exception("한글 'FindReplace' 파라미터 세트를 생성하지 못했습니다. (NoneType 에러 발생 지점)")

    param.SetItem("FindString", "")
    param.SetItem("ReplaceString", "")
    param.SetItem("FindType", 1)
    param.SetItem("ReplaceType", 1)
    param.SetItem("WholeWordOnly", 0)
    param.SetItem("MatchCase", 0)
    param.SetItem("AllWordForms", 0)
    param.SetItem("SeveralWords", 0)
    param.SetItem("UseWildCards", 0)
    param.SetItem("CircularSearch", 0)
    param.SetItem("StartPosCur", 0)
    param.SetItem("ReverseFind", 0)
    param.SetItem("FindJaso", 0)
    param.SetItem("FindRegEx", 0)

    find_face = param.CreateItemArray("FindCharShape", 1)
    for key in ["FaceNameHangul","FaceNameLatin","FaceNameHanja","FaceNameJapanese","FaceNameOther","FaceNameSymbol","FaceNameUser"]:
        find_face.SetItem(key, src_font)

    replace_face = param.CreateItemArray("ReplaceCharShape", 1)
    for key in ["FaceNameHangul","FaceNameLatin","FaceNameHanja","FaceNameJapanese","FaceNameOther","FaceNameSymbol","FaceNameUser"]:
        replace_face.SetItem(key, dst_font)

    act.Execute(param)  # 파라미터를 명시적으로 전달하여 실행 안정성 확보
    hwp.Save()
    hwp.Close()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("HWP 서체 일괄 변경")
        self.resizable(False, False)
        self.configure(bg="#F7F8FA")
        self._build_ui()
        self._center_window(480, 480)

    def _center_window(self, w, h):
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def _build_ui(self):
        PAD = 24
        BG, CARD = "#F7F8FA", "#FFFFFF"
        ACCENT, GRAY, BORDER = "#5B6EF5", "#6B7280", "#E5E7EB"
        F_TITLE = ("맑은 고딕", 14, "bold")
        F_LABEL = ("맑은 고딕", 10)
        F_SMALL = ("맑은 고딕", 9)

        tf = tk.Frame(self, bg=BG)
        tf.pack(fill="x", padx=PAD, pady=(PAD, 8))
        tk.Label(tf, text="HWP 서체 일괄 변경", font=F_TITLE, bg=BG, fg="#111827").pack(anchor="w")
        tk.Label(tf, text="폴더 안의 모든 HWP 파일 서체를 한 번에 바꿉니다", font=F_SMALL, bg=BG, fg=GRAY).pack(anchor="w", pady=(2,0))

        card = tk.Frame(self, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        card.pack(fill="x", padx=PAD, pady=4)

        tk.Label(card, text="📁  변경할 HWP 파일 폴더", font=F_LABEL, bg=CARD, fg="#374151").pack(anchor="w", padx=16, pady=(14,6))
        fr = tk.Frame(card, bg=CARD)
        fr.pack(fill="x", padx=16, pady=(0,14))
        self.folder_var = tk.StringVar()
        tk.Entry(fr, textvariable=self.folder_var, font=F_LABEL, relief="flat",
                 bg="#F3F4F6", fg="#374151", highlightthickness=1, highlightbackground=BORDER
                 ).pack(side="left", fill="x", expand=True, ipady=6, padx=(0,8))
        tk.Button(fr, text="폴더 선택", command=self._pick_folder,
                  bg=ACCENT, fg="white", font=F_LABEL, relief="flat",
                  cursor="hand2", padx=12, pady=4, bd=0,
                  activebackground="#4B5EE4", activeforeground="white").pack(side="right")

        tk.Frame(card, bg=BORDER, height=1).pack(fill="x", padx=16, pady=4)

        tk.Label(card, text="🔄  변환 서체 목록", font=F_LABEL, bg=CARD, fg="#374151").pack(anchor="w", padx=16, pady=(4,6))
        for src, dst in FONT_MAP:
            row = tk.Frame(card, bg=CARD)
            row.pack(fill="x", padx=24, pady=2)
            tk.Label(row, text=src, font=F_LABEL, bg="#F3F4F6", fg="#374151",
                     padx=10, pady=4, relief="flat").pack(side="left")
            tk.Label(row, text=" → ", font=("맑은 고딕", 11), bg=CARD, fg=ACCENT).pack(side="left")
            tk.Label(row, text=dst, font=F_LABEL, bg="#F3F4F6", fg="#374151",
                     padx=10, pady=4, relief="flat").pack(side="left")

        self.backup_var = tk.BooleanVar(value=True)
        tk.Checkbutton(card, text=" 원본 파일 백업 (.bak 파일 생성)", variable=self.backup_var,
                       bg=CARD, fg=GRAY, font=F_SMALL, activebackground=CARD,
                       selectcolor=CARD, cursor="hand2").pack(anchor="w", padx=16, pady=(8,14))

        self.run_btn = tk.Button(self, text="서체 일괄 변경 시작", command=self._run,
                                 bg=ACCENT, fg="white", font=("맑은 고딕", 11, "bold"),
                                 relief="flat", cursor="hand2", pady=10, bd=0,
                                 activebackground="#4B5EE4", activeforeground="white")
        self.run_btn.pack(fill="x", padx=PAD, pady=(12,8))

        pf = tk.Frame(self, bg=BG)
        pf.pack(fill="x", padx=PAD)
        self.status_var = tk.StringVar(value="")
        tk.Label(pf, textvariable=self.status_var, font=F_SMALL, bg=BG, fg=GRAY).pack(anchor="w")
        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill="x", padx=PAD, pady=(4, PAD))

    def _pick_folder(self):
        path = filedialog.askdirectory(title="HWP 파일 폴더 선택")
        if path:
            self.folder_var.set(path)

    def _run(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("입력 오류", "폴더를 선택해주세요.")
            return
        if not check_hwp():
            return
        hwp_files = [f for f in os.listdir(folder) if f.lower().endswith(".hwp")]
        if not hwp_files:
            messagebox.showinfo("파일 없음", "선택한 폴더에 HWP 파일이 없습니다.")
            return
        self.run_btn.config(state="disabled", text="변경 중...")
        self.progress["maximum"] = len(hwp_files)
        self.progress["value"] = 0
        threading.Thread(target=self._worker, args=(folder, hwp_files), daemon=True).start()

    def _worker(self, folder, files):
        success, fail = 0, 0
        try:
            try:
                hwp = gencache.EnsureDispatch("HWPFrame.HwpObject")
            except Exception:
                hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
            
            # 💡 중요: 백그라운드에서 에러가 숨는 것을 막기 위해 창을 눈에 보이게 켭니다.
            hwp.XHwpWindows.Item(0).Visible = True
            time.sleep(0.5)  # 한글 창이 완전히 로드될 때까지 대기
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("오류", f"한글 프로그램 초기화 실패.\n\n[상세 에러]: {e}"))
            self.after(0, self._reset_btn)
            return

        for i, fname in enumerate(files, 1):
            fpath = os.path.join(folder, fname)
            self.after(0, lambda n=fname, i=i, t=len(files):
                       self.status_var.set(f"처리 중 ({i}/{t}): {n}"))
            try:
                if self.backup_var.get():
                    shutil.copy2(fpath, fpath + ".bak")
                
                abs_path = os.path.abspath(fpath)
                for src_font, dst_font in FONT_MAP:
                    change_font_in_file(hwp, abs_path, src_font, dst_font)
                success += 1
            except Exception as ex:
                fail += 1
                self.after(0, lambda f=fname, e=ex: messagebox.showerror(
                    "파일 변환 실패", 
                    f"파일 처리 중 에러가 발생했습니다:\n[{f}]\n\n[상세 에러 내용]:\n{e}"
                ))
                
            self.after(0, lambda v=i: self.progress.config(value=v))

        try:
            hwp.Quit()
        except Exception:
            pass

        msg = f"✅ 완료!  성공 {success}개"
        if fail:
            msg += f"  /  실패 {fail}개"
        self.after(0, lambda: self.status_var.set(msg))
        self.after(0, lambda: messagebox.showinfo("완료", msg))
        self.after(0, self._reset_btn)

    def _reset_btn(self):
        self.run_btn.config(state="normal", text="서체 일괄 변경 시작")


if __name__ == "__main__":
    app = App()
    app.mainloop()
