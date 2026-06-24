# -*- coding: utf-8 -*-
"""
hwp2docx_gui.py — HWP -> DOCX 변환기 (간단 GUI)

추가 설치 없이 동작(파이썬 기본 내장 tkinter 사용).
같은 폴더의 hwp2docx.py 를 불러서 변환한다.

실행:
  - "HWP를 DOCX로 변환.bat" 더블클릭, 또는
  - python hwp2docx_gui.py
"""

import os
import sys
import threading
import traceback
import queue

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# 같은 폴더의 변환 모듈 사용
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hwp2docx
import docx2hwpx

# 변환 방향 정의: (라벨, 모듈, 입력확장자, 파일대화상자 필터)
DIRECTIONS = {
    "hwp2docx": ("HWP → DOCX", hwp2docx, ".hwp",
                 [("한글 문서", "*.hwp"), ("모든 파일", "*.*")]),
    "docx2hwpx": ("DOCX → HWPX", docx2hwpx, ".docx",
                  [("워드 문서", "*.docx"), ("모든 파일", "*.*")]),
}


class App:
    def __init__(self, root):
        self.root = root
        self.q = queue.Queue()
        self.last_outdir = None
        self.busy = False

        self.direction = tk.StringVar(value="hwp2docx")

        root.title("한글 ↔ 워드 변환기")
        root.geometry("640x520")
        root.minsize(520, 440)

        # 상단 안내
        head = tk.Label(root, text="한글 ↔ 워드 문서 변환",
                        font=("맑은 고딕", 13, "bold"))
        head.pack(pady=(14, 4))
        sub = tk.Label(root, text="변환 방향을 고르고, 파일이나 폴더를 선택하세요.",
                       fg="#555", font=("맑은 고딕", 9))
        sub.pack(pady=(0, 8))

        # 변환 방향 선택
        dirframe = tk.LabelFrame(root, text="변환 방향", font=("맑은 고딕", 9))
        dirframe.pack(pady=4)
        tk.Radiobutton(dirframe, text="HWP → DOCX  (한글 → 워드)",
                       variable=self.direction, value="hwp2docx",
                       font=("맑은 고딕", 10)).grid(row=0, column=0, sticky="w",
                                                   padx=10, pady=2)
        tk.Radiobutton(dirframe, text="DOCX → HWPX  (워드 → 한글)",
                       variable=self.direction, value="docx2hwpx",
                       font=("맑은 고딕", 10)).grid(row=1, column=0, sticky="w",
                                                   padx=10, pady=2)

        # 버튼 영역
        btns = tk.Frame(root)
        btns.pack(pady=4)
        self.btn_files = ttk.Button(btns, text="📄  HWP 파일 선택",
                                    width=22, command=self.pick_files)
        self.btn_files.grid(row=0, column=0, padx=6)
        self.btn_folder = ttk.Button(btns, text="📁  폴더 전체 변환",
                                     width=22, command=self.pick_folder)
        self.btn_folder.grid(row=0, column=1, padx=6)

        # 로그 영역
        logframe = tk.Frame(root)
        logframe.pack(fill="both", expand=True, padx=14, pady=10)
        self.log = tk.Text(logframe, wrap="word", font=("맑은 고딕", 10),
                           state="disabled", bg="#fafafa")
        sb = ttk.Scrollbar(logframe, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log.pack(side="left", fill="both", expand=True)
        self.log.tag_config("ok", foreground="#1a7f37")
        self.log.tag_config("err", foreground="#cf222e")
        self.log.tag_config("info", foreground="#0550ae")

        # 하단 상태/열기 버튼
        bottom = tk.Frame(root)
        bottom.pack(fill="x", padx=14, pady=(0, 12))
        self.status = tk.Label(bottom, text="대기 중", anchor="w", fg="#555")
        self.status.pack(side="left")
        self.btn_open = ttk.Button(bottom, text="결과 폴더 열기",
                                   command=self.open_outdir, state="disabled")
        self.btn_open.pack(side="right")

        self._write("준비되었습니다. 변환 방향을 고른 뒤 파일 또는 폴더를 선택하세요.\n"
                    "  • HWP → DOCX : 한글 문서를 워드로\n"
                    "  • DOCX → HWPX : 워드 문서를 한글로 (한글에서 열어 확인 권장)\n",
                    "info")
        self.root.after(120, self._drain)

    # ---- 로그 ----
    def _write(self, text, tag=None):
        self.log.configure(state="normal")
        self.log.insert("end", text, tag or ())
        self.log.see("end")
        self.log.configure(state="disabled")

    def _drain(self):
        """워커 스레드 메시지를 UI에 반영"""
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "log":
                    self._write(payload[0], payload[1])
                elif kind == "status":
                    self.status.config(text=payload)
                elif kind == "done":
                    self._finish(payload)
        except queue.Empty:
            pass
        self.root.after(120, self._drain)

    # ---- 파일/폴더 선택 ----
    def _cur(self):
        return DIRECTIONS[self.direction.get()]

    def pick_files(self):
        if self.busy:
            return
        label, _, ext, ftypes = self._cur()
        files = filedialog.askopenfilenames(
            title="변환할 파일 선택 (%s)" % label, filetypes=ftypes)
        if files:
            self.start(list(files))

    def pick_folder(self):
        if self.busy:
            return
        label, _, ext, _ = self._cur()
        d = filedialog.askdirectory(title="파일이 있는 폴더 선택 (%s)" % label)
        if not d:
            return
        files = [os.path.join(d, f) for f in os.listdir(d)
                 if f.lower().endswith(ext)]
        if not files:
            messagebox.showinfo("안내", "선택한 폴더에 %s 파일이 없습니다." % ext)
            return
        self.start(files)

    # ---- 변환 ----
    def start(self, files):
        self.busy = True
        self.btn_files.config(state="disabled")
        self.btn_folder.config(state="disabled")
        self.btn_open.config(state="disabled")
        label, module, _, _ = self._cur()
        self._write("\n──────── %s 변환 시작 (%d개) ────────\n"
                    % (label, len(files)), "info")
        t = threading.Thread(target=self._worker, args=(files, module),
                             daemon=True)
        t.start()

    def _worker(self, files, module):
        ok = 0
        outdir = None
        for i, src in enumerate(files, 1):
            name = os.path.basename(src)
            self.q.put(("status", "변환 중 %d/%d: %s" % (i, len(files), name)))
            self.q.put(("log", ("• %s ... " % name, None)))
            try:
                dst = module.convert(src)
                outdir = os.path.dirname(os.path.abspath(dst))
                self.q.put(("log", ("완료 → %s\n" % os.path.basename(dst), "ok")))
                ok += 1
            except Exception as e:
                self.q.put(("log", ("실패: %s\n" % e, "err")))
                traceback.print_exc()
        self.q.put(("done", (ok, len(files), outdir)))

    def _finish(self, payload):
        ok, total, outdir = payload
        self.busy = False
        self.btn_files.config(state="normal")
        self.btn_folder.config(state="normal")
        self.last_outdir = outdir
        if outdir:
            self.btn_open.config(state="normal")
        msg = "완료: 성공 %d / 전체 %d" % (ok, total)
        self.status.config(text=msg)
        self._write("──────── %s ────────\n" % msg,
                    "ok" if ok == total else "err")
        if ok:
            messagebox.showinfo("변환 완료",
                                "%s\n\n결과 파일은 원본과 같은 폴더에 저장되었습니다." % msg)

    def open_outdir(self):
        if self.last_outdir and os.path.isdir(self.last_outdir):
            try:
                os.startfile(self.last_outdir)
            except Exception:
                messagebox.showinfo("경로", self.last_outdir)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
