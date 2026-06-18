import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import shutil

FONT_MAP = [
    ("맑은 고딕", "함초롬돋움"),
    ("바탕",      "함초롬바탕"),
]

try:
    import win32com.client
    import win32com.client.gencache as gencache  # PyInstaller 빌드 시 안정성을 위해 gencache 명시적 임포트 
    HWP_AVAILABLE = True
except ImportError:
    HWP_AVAILABLE = False


def check_hwp():
    if not HWP_AVAILABLE:
        messagebox.showerror("오류", "pywin32가 설치되어 있지 않습니다.\npip install pywin32 명령어로 설치해주세요. ")
        return False
    try:
        # 가상 가상 환경 및 다양한 윈도우 환경 대응을 위해 2중 바인딩 시도 
        try:
            hwp = gencache.EnsureDispatch("HWPFrame.HwpObject") [cite: 1]
        except Exception:
            hwp = win32com.client.Dispatch("HWPFrame.HwpObject") [cite: 1]
        hwp.Quit() [cite: 1]
        return True
    except Exception:
        messagebox.showerror("한글 프로그램 없음", "한컴 한글(HWP)이 설치되어 있지 않습니다.\n한글 프로그램 설치 후 다시 실행해주세요. ")
        return False


def change_font_in_file(hwp, filepath, src_font, dst_font):
    hwp.Open(filepath, "HWP", "forceopen:true") [cite: 1]
    act = hwp.CreateAction("FindReplace") [cite: 1]
    param = act.CreateSet() [cite: 1]
    param.SetItem("FindString", "") [cite: 1]
    param.SetItem("ReplaceString", "") [cite: 1]
    param.SetItem("FindType", 1) [cite: 1]
    param.SetItem("ReplaceType", 1) [cite: 1]
    param.SetItem("WholeWordOnly", 0) [cite: 1]
    param.SetItem("MatchCase", 0) [cite: 1]
    param.SetItem("AllWordForms", 0) [cite: 1]
    param.SetItem("SeveralWords", 0) [cite: 1]
    param.SetItem("UseWildCards", 0) [cite: 1]
    param.SetItem("CircularSearch", 0) [cite: 1]
    param.SetItem("StartPosCur", 0) [cite: 1]
    param.SetItem("ReverseFind", 0) [cite: 1]
    param.SetItem("FindJaso", 0) [cite: 1]
    param.SetItem("FindRegEx", 0) [cite: 1]

    find_face = param.CreateItemArray("FindCharShape", 1) [cite: 1]
    for key in ["FaceNameHangul","FaceNameLatin","FaceNameHanja","FaceNameJapanese","FaceNameOther","FaceNameSymbol","FaceNameUser"]: [cite: 1]
        find_face.SetItem(key, src_font) [cite: 1]

    replace_face = param.CreateItemArray("ReplaceCharShape", 1) [cite: 1]
    for key in ["FaceNameHangul","FaceNameLatin","FaceNameHanja","FaceNameJapanese","FaceNameOther","FaceNameSymbol","FaceNameUser"]: [cite: 1]
        replace_face.SetItem(key, dst_font) [cite: 1]

    act.Execute() [cite: 1]
    hwp.Save() [cite: 1]
    hwp.Close() [cite: 1]


class App(tk.Tk):
    def __init__(self):
        super().__init__() [cite: 1]
        self.title("HWP 서체 일괄 변경") [cite: 1]
        self.resizable(False, False) [cite: 1]
        self.configure(bg="#F7F8FA") [cite: 1]
        self._build_ui() [cite: 1]
        self._center_window(480, 480) [cite: 1]

    def _center_window(self, w, h):
        self.update_idletasks() [cite: 1]
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight() [cite: 1]
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}") [cite: 1]

    def _build_ui(self):
        PAD = 24 [cite: 1]
        BG, CARD = "#F7F8FA", "#FFFFFF" [cite: 1]
        ACCENT, GRAY, BORDER = "#5B6EF5", "#6B7280", "#E5E7EB" [cite: 1]
        F_TITLE = ("맑은 고딕", 14, "bold") [cite: 1]
        F_LABEL = ("맑은 고딕", 10) [cite: 1]
        F_SMALL = ("맑은 고딕", 9) [cite: 1]

        # 타이틀
        tf = tk.Frame(self, bg=BG) [cite: 1]
        tf.pack(fill="x", padx=PAD, pady=(PAD, 8)) [cite: 1]
        tk.Label(tf, text="HWP 서체 일괄 변경", font=F_TITLE, bg=BG, fg="#111827").pack(anchor="w") [cite: 1]
        tk.Label(tf, text="폴더 안의 모든 HWP 파일 서체를 한 번에 바꿉니다", font=F_SMALL, bg=BG, fg=GRAY).pack(anchor="w", pady=(2,0)) [cite: 1]

        # 카드
        card = tk.Frame(self, bg=CARD, highlightthickness=1, highlightbackground=BORDER) [cite: 1]
        card.pack(fill="x", padx=PAD, pady=4) [cite: 1]

        # 폴더 선택
        tk.Label(card, text="📁  변경할 HWP 파일 폴더", font=F_LABEL, bg=CARD, fg="#374151").pack(anchor="w", padx=16, pady=(14,6)) [cite: 1]
        fr = tk.Frame(card, bg=CARD) [cite: 1]
        fr.pack(fill="x", padx=16, pady=(0,14)) [cite: 1]
        self.folder_var = tk.StringVar() [cite: 1]
        tk.Entry(fr, textvariable=self.folder_var, font=F_LABEL, relief="flat", [cite: 1]
                 bg="#F3F4F6", fg="#374151", highlightthickness=1, highlightbackground=BORDER [cite: 1]
                 ).pack(side="left", fill="x", expand=True, ipady=6, padx=(0,8)) [cite: 1]
        tk.Button(fr, text="폴더 선택", command=self._pick_folder, [cite: 1]
                  bg=ACCENT, fg="white", font=F_LABEL, relief="flat", [cite: 1]
                  cursor="hand2", padx=12, pady=4, bd=0, [cite: 1]
                  activebackground="#4B5EE4", activeforeground="white").pack(side="right") [cite: 1]

        # 구분선
        tk.Frame(card, bg=BORDER, height=1).pack(fill="x", padx=16, pady=4) [cite: 1]

        # 변환 목록 표시
        tk.Label(card, text="🔄  변환 서체 목록", font=F_LABEL, bg=CARD, fg="#374151").pack(anchor="w", padx=16, pady=(4,6)) [cite: 1]
        for src, dst in FONT_MAP: [cite: 1]
            row = tk.Frame(card, bg=CARD) [cite: 1]
            row.pack(fill="x", padx=24, pady=2) [cite: 1]
            tk.Label(row, text=src, font=F_LABEL, bg="#F3F4F6", fg="#374151", [cite: 1]
                     padx=10, pady=4, relief="flat").pack(side="left") [cite: 1]
            tk.Label(row, text=" → ", font=("맑은 고딕", 11), bg=CARD, fg=ACCENT).pack(side="left") [cite: 1]
            tk.Label(row, text=dst, font=F_LABEL, bg="#F3F4F6", fg="#374151", [cite: 1]
                     padx=10, pady=4, relief="flat").pack(side="left") [cite: 1]

        # 백업 체크박스
        self.backup_var = tk.BooleanVar(value=True) [cite: 1]
        tk.Checkbutton(card, text=" 원본 파일 백업 (.bak 파일 생성)", variable=self.backup_var, [cite: 1]
                       bg=CARD, fg=GRAY, font=F_SMALL, activebackground=CARD, [cite: 1]
                       selectcolor=CARD, cursor="hand2").pack(anchor="w", padx=16, pady=(8,14)) [cite: 1]

        # 실행 버튼
        self.run_btn = tk.Button(self, text="서체 일괄 변경 시작", command=self._run, [cite: 1]
                                 bg=ACCENT, fg="white", font=("맑은 고딕", 11, "bold"), [cite: 1]
                                 relief="flat", cursor="hand2", pady=10, bd=0, [cite: 1]
                                 activebackground="#4B5EE4", activeforeground="white") [cite: 1]
        self.run_btn.pack(fill="x", padx=PAD, pady=(12,8)) [cite: 1]

        # 진행 상황
        pf = tk.Frame(self, bg=BG) [cite: 1]
        pf.pack(fill="x", padx=PAD) [cite: 1]
        self.status_var = tk.StringVar(value="") [cite: 1]
        tk.Label(pf, textvariable=self.status_var, font=F_SMALL, bg=BG, fg=GRAY).pack(anchor="w") [cite: 1]
        self.progress = ttk.Progressbar(self, mode="determinate") [cite: 1]
        self.progress.pack(fill="x", padx=PAD, pady=(4, PAD)) [cite: 1]

    def _pick_folder(self):
        path = filedialog.askdirectory(title="HWP 파일 폴더 선택") [cite: 1]
        if path: [cite: 1]
            self.folder_var.set(path) [cite: 1]

    def _run(self):
        folder = self.folder_var.get().strip() [cite: 1]
        if not folder or not os.path.isdir(folder): [cite: 1]
            messagebox.showwarning("입력 오류", "폴더를 선택해주세요.") [cite: 1]
            return
        if not check_hwp(): [cite: 1]
            return
        hwp_files = [f for f in os.listdir(folder) if f.lower().endswith(".hwp")] [cite: 1]
        if not hwp_files: [cite: 1]
            messagebox.showinfo("파일 없음", "선택한 폴더에 HWP 파일이 없습니다.") [cite: 1]
            return
        self.run_btn.config(state="disabled", text="변경 중...") [cite: 1]
        self.progress["maximum"] = len(hwp_files) [cite: 1]
        self.progress["value"] = 0 [cite: 1]
        threading.Thread(target=self._worker, args=(folder, hwp_files), daemon=True).start() [cite: 1]

    def _worker(self, folder, files):
        success, fail = 0, 0 [cite: 1]
        try:
            # 2중 바인딩구조 적용 
            try:
                hwp = gencache.EnsureDispatch("HWPFrame.HwpObject") [cite: 1]
            except Exception:
                hwp = win32com.client.Dispatch("HWPFrame.HwpObject") [cite: 1]
            hwp.XHwpWindows.Item(0).Visible = False [cite: 1]
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("오류", f"한글 실행 실패: {e}")) [cite: 1]
            self.after(0, self._reset_btn) [cite: 1]
            return

        for i, fname in enumerate(files, 1): [cite: 1]
            fpath = os.path.join(folder, fname) [cite: 1]
            self.after(0, lambda n=fname, i=i, t=len(files): [cite: 1]
                       self.status_var.set(f"처리 중 ({i}/{t}): {n}")) [cite: 1]
            try:
                if self.backup_var.get(): [cite: 1]
                    shutil.copy2(fpath, fpath + ".bak") [cite: 1]
                for src_font, dst_font in FONT_MAP: [cite: 1]
                    change_font_in_file(hwp, os.path.abspath(fpath), src_font, dst_font) [cite: 1]
                success += 1 [cite: 1]
            except Exception as ex:
                print(f"[실패] {fname}: {ex}") [cite: 1]
                fail += 1 [cite: 1]
            self.after(0, lambda v=i: self.progress.config(value=v)) [cite: 1]

        try:
            hwp.Quit() [cite: 1]
        except Exception:
            pass [cite: 1]

        msg = f"✅ 완료!  성공 {success}개" [cite: 1]
        if fail: [cite: 1]
            msg += f"  /  실패 {fail}개" [cite: 1]
        self.after(0, lambda: self.status_var.set(msg)) [cite: 1]
        self.after(0, lambda: messagebox.showinfo("완료", msg)) [cite: 1]
        self.after(0, self._reset_btn) [cite: 1]

    def _reset_btn(self):
        self.run_btn.config(state="normal", text="서체 일괄 변경 시작") [cite: 1]


if __name__ == "__main__":
    app = App() [cite: 1]
    app.mainloop() [cite: 1]
