import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import shutil
import sys

try:
    import win32com.client
    HWP_AVAILABLE = True
except ImportError:
    HWP_AVAILABLE = False

import logging
LOG_PATH = os.path.join(os.path.expanduser("~"), "Desktop", "hwp_log.txt")
logging.basicConfig(
    filename=LOG_PATH,
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8"
)

# 고정 폰트 매핑
FONT_MAP = [
    ("맑은 고딕", "함초롬돋움"),
    ("바탕",      "함초롬바탕"),
]


def get_hwp():
    """한글 COM 인스턴스 반환. 실패 시 None."""
    for prog_id in [
        "HwpAutomationApp2.HwpAutomation",
        "HwpAutomationApp2.HwpAutomation.1",
        "HwpAutomationApp2.HwpAutomation.2",
        "HWPFrame.HwpObject",
        "HWPFrame.HwpObject.1",
    ]:
        try:
            hwp = win32com.client.Dispatch(prog_id)
            return hwp
        except Exception:
            continue
    return None


def register_security(hwp):
    """한글 2022 보안 모듈 등록 시도 (실패해도 계속 진행)"""
    try:
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
    except Exception:
        pass
    try:
        hwp.SetMessageBoxMode(0x00)  # 메시지박스 자동 닫기
    except Exception:
        pass


def change_font(hwp, src_font, dst_font):
    """현재 열린 문서에서 폰트 교체"""
    act = hwp.CreateAction("FindReplace")
    param = act.CreateSet()

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
    for key in ["FaceNameHangul", "FaceNameLatin", "FaceNameHanja",
                "FaceNameJapanese", "FaceNameOther", "FaceNameSymbol", "FaceNameUser"]:
        find_face.SetItem(key, src_font)

    replace_face = param.CreateItemArray("ReplaceCharShape", 1)
    for key in ["FaceNameHangul", "FaceNameLatin", "FaceNameHanja",
                "FaceNameJapanese", "FaceNameOther", "FaceNameSymbol", "FaceNameUser"]:
        replace_face.SetItem(key, dst_font)

    act.Execute()


def process_file(hwp, filepath):
    """파일 열기 → 폰트 교체 → 저장 → 닫기"""
    hwp.Open(filepath, "HWP", "forceopen:true")
    register_security(hwp)
    for src, dst in FONT_MAP:
        change_font(hwp, src, dst)
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

        # 타이틀
        tf = tk.Frame(self, bg=BG)
        tf.pack(fill="x", padx=PAD, pady=(PAD, 8))
        tk.Label(tf, text="HWP 서체 일괄 변경", font=F_TITLE,
                 bg=BG, fg="#111827").pack(anchor="w")
        tk.Label(tf, text="폴더 안의 모든 HWP 파일 서체를 한 번에 바꿉니다",
                 font=F_SMALL, bg=BG, fg=GRAY).pack(anchor="w", pady=(2, 0))

        # 카드
        card = tk.Frame(self, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        card.pack(fill="x", padx=PAD, pady=4)

        # 폴더 선택
        tk.Label(card, text="📁  변경할 HWP 파일 폴더", font=F_LABEL,
                 bg=CARD, fg="#374151").pack(anchor="w", padx=16, pady=(14, 6))
        fr = tk.Frame(card, bg=CARD)
        fr.pack(fill="x", padx=16, pady=(0, 14))
        self.folder_var = tk.StringVar()
        tk.Entry(fr, textvariable=self.folder_var, font=F_LABEL, relief="flat",
                 bg="#F3F4F6", fg="#374151",
                 highlightthickness=1, highlightbackground=BORDER
                 ).pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 8))
        tk.Button(fr, text="폴더 선택", command=self._pick_folder,
                  bg=ACCENT, fg="white", font=F_LABEL, relief="flat",
                  cursor="hand2", padx=12, pady=4, bd=0,
                  activebackground="#4B5EE4", activeforeground="white").pack(side="right")

        tk.Frame(card, bg=BORDER, height=1).pack(fill="x", padx=16, pady=4)

        # 변환 목록
        tk.Label(card, text="🔄  변환 서체 목록", font=F_LABEL,
                 bg=CARD, fg="#374151").pack(anchor="w", padx=16, pady=(4, 6))
        for src, dst in FONT_MAP:
            row = tk.Frame(card, bg=CARD)
            row.pack(fill="x", padx=24, pady=3)
            tk.Label(row, text=src, font=F_LABEL, bg="#F3F4F6", fg="#374151",
                     padx=10, pady=4).pack(side="left")
            tk.Label(row, text=" → ", font=("맑은 고딕", 11),
                     bg=CARD, fg=ACCENT).pack(side="left")
            tk.Label(row, text=dst, font=F_LABEL, bg="#F3F4F6", fg="#374151",
                     padx=10, pady=4).pack(side="left")

        # 백업 체크박스
        self.backup_var = tk.BooleanVar(value=True)
        tk.Checkbutton(card, text=" 원본 파일 백업 (.bak 파일 생성)",
                       variable=self.backup_var,
                       bg=CARD, fg=GRAY, font=F_SMALL,
                       activebackground=CARD, selectcolor=CARD,
                       cursor="hand2").pack(anchor="w", padx=16, pady=(8, 14))

        # 실행 버튼
        self.run_btn = tk.Button(self, text="서체 일괄 변경 시작",
                                 command=self._run,
                                 bg=ACCENT, fg="white",
                                 font=("맑은 고딕", 11, "bold"),
                                 relief="flat", cursor="hand2",
                                 pady=10, bd=0,
                                 activebackground="#4B5EE4", activeforeground="white")
        self.run_btn.pack(fill="x", padx=PAD, pady=(16, 8))

        # 상태 / 진행바
        pf = tk.Frame(self, bg=BG)
        pf.pack(fill="x", padx=PAD)
        self.status_var = tk.StringVar(value="")
        tk.Label(pf, textvariable=self.status_var,
                 font=F_SMALL, bg=BG, fg=GRAY).pack(anchor="w")
        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill="x", padx=PAD, pady=(4, PAD))

    def _pick_folder(self):
        path = filedialog.askdirectory(title="HWP 파일 폴더 선택")
        if path:
            self.folder_var.set(path)

    def _run(self):
        if not HWP_AVAILABLE:
            messagebox.showerror("오류",
                "pywin32가 설치되지 않았습니다.\n"
                "이 exe는 Windows 전용입니다.")
            return
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("입력 오류", "폴더를 선택해주세요.")
            return
        hwp_files = [f for f in os.listdir(folder) if f.lower().endswith(".hwp")]
        if not hwp_files:
            messagebox.showinfo("파일 없음", "선택한 폴더에 HWP 파일이 없습니다.")
            return

        self.run_btn.config(state="disabled", text="변경 중...")
        self.progress["maximum"] = len(hwp_files)
        self.progress["value"] = 0
        threading.Thread(target=self._worker,
                         args=(folder, hwp_files), daemon=True).start()

    def _worker(self, folder, files):
        # 한글 실행
        hwp = get_hwp()
        if hwp is None:
            logging.error("한글 COM 인스턴스 생성 실패")
            self.after(0, lambda: messagebox.showerror(
                "한글 없음",
                "한컴 한글(HWP)을 찾을 수 없습니다.\n"
                "한글 프로그램이 설치되어 있는지 확인해주세요."))
            self.after(0, self._reset_btn)
            return

        logging.info(f"한글 실행 성공")

        # 한글 창 숨기기
        try:
            hwp.XHwpWindows.Item(0).Visible = False
        except Exception as e:
            logging.warning(f"창 숨기기 실패(무시): {e}")

        success, fail, fail_list = 0, 0, []

        for i, fname in enumerate(files, 1):
            fpath = os.path.abspath(os.path.join(folder, fname))
            self.after(0, lambda n=fname, i=i, t=len(files):
                       self.status_var.set(f"처리 중 ({i}/{t}): {n}"))
            try:
                if self.backup_var.get():
                    shutil.copy2(fpath, fpath + ".bak")
                logging.info(f"처리 시작: {fpath}")
                process_file(hwp, fpath)
                logging.info(f"처리 성공: {fname}")
                success += 1
            except Exception as ex:
                fail += 1
                fail_list.append(fname)
                logging.error(f"처리 실패: {fname} → {ex}", exc_info=True)
            self.after(0, lambda v=i: self.progress.config(value=v))

        try:
            hwp.Quit()
        except Exception:
            pass

        # 결과 메시지
        msg = f"✅ 완료!  성공 {success}개"
        if fail:
            msg += f"  /  실패 {fail}개"
            detail = "\n".join(fail_list[:10])
            if len(fail_list) > 10:
                detail += f"\n... 외 {len(fail_list)-10}개"
            self.after(0, lambda: messagebox.showwarning(
                "완료 (일부 실패)", f"{msg}\n\n실패 파일:\n{detail}"))
        else:
            self.after(0, lambda: messagebox.showinfo("완료", msg))

        self.after(0, lambda: self.status_var.set(msg))
        self.after(0, self._reset_btn)

    def _reset_btn(self):
        self.run_btn.config(state="normal", text="서체 일괄 변경 시작")


if __name__ == "__main__":
    app = App()
    app.mainloop()
