import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import shutil
import time

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
    try:
        hwp.Clear(1) 
    except Exception:
        pass

    open_res = hwp.Open(filepath, "HWP", "forceopen:true")
    if not open_res:
        raise Exception("한글 파일 열기 실패")

    time.sleep(0.5)

    act = hwp.CreateAction("FindReplace")
    if not act:
        raise Exception("한글 엔진이 'FindReplace' 액션 생성을 거부했습니다. (한글 프로그램 창을 확인해주세요)")

    param = act.CreateSet()
    if not param:
        raise Exception("FindReplace 파라미터 세트 생성 실패")

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
    font_keys = ["FaceNameHangul", "FaceNameLatin", "FaceNameHanja", "FaceNameJapanese", "FaceNameOther", "FaceNameSymbol", "FaceNameUser"]
    for key in font_keys:
        find_face.SetItem(key, src_font)

    replace_face = param.CreateItemArray("ReplaceCharShape", 1)
    for key in font_keys:
        replace_face.SetItem(key, dst_font)

    act.Execute(param)
    
    hwp.Save()
    hwp.Clear(1)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("HWP 서체 일괄 변경")
        self.resizable(False, False)
        self.configure(bg="#F7F8FA")
        self._build_ui()
        self._center_window(480, 540)

    def _center_window(self, w, h):
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.
