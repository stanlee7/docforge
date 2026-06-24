"""DocForge - Tkinter GUI. 엑셀 행 → 개인별 AI 문서 일괄 생성."""
import os
import threading
import traceback
import webbrowser
from datetime import datetime
from tkinter import (Tk, Toplevel, StringVar, BooleanVar, Text, Listbox,
                     filedialog, messagebox, END, MULTIPLE,
                     N, S, E, W)
from tkinter import ttk

from . import config, engine, excel_io, pipeline, setup_env


class App:
    def __init__(self, root: Tk):
        self.root = root
        root.title("DocForge — 엑셀 행 → 개인별 AI 문서 일괄 생성")
        root.geometry("780x720")
        root.minsize(720, 640)

        self.excel = StringVar()
        self.out_dir = StringVar()
        self.preset = StringVar(value=config.DEFAULT_PRESET)
        self.fmt = StringVar(value=config.DEFAULT_FORMAT)
        self.fname = StringVar()
        self.fast = BooleanVar(value=False)
        self.status = StringVar(value="엑셀 명단 파일을 선택하세요.")
        self.headers = []
        self.rows = []
        self.cancel_flag = False
        self.last_out_dir = None

        frm = ttk.Frame(root, padding=14)
        frm.grid(sticky=(N, S, E, W))
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        frm.columnconfigure(1, weight=1)
        r = 0

        # 1) 엑셀 선택
        ttk.Label(frm, text="엑셀 명단").grid(row=r, column=0, sticky=W, pady=4)
        ttk.Entry(frm, textvariable=self.excel).grid(row=r, column=1, sticky=(E, W), padx=8)
        ttk.Button(frm, text="찾아보기…", command=self.pick_excel).grid(row=r, column=2)
        r += 1

        # 2) 문서 유형 + 출력 형식
        opt = ttk.Frame(frm)
        opt.grid(row=r, column=0, columnspan=3, sticky=W, pady=4)
        ttk.Label(opt, text="문서 유형").grid(row=0, column=0, sticky=W)
        self.preset_cb = ttk.Combobox(opt, textvariable=self.preset, state="readonly",
                                      values=list(config.PRESETS.keys()), width=16)
        self.preset_cb.grid(row=0, column=1, padx=(6, 18))
        self.preset_cb.bind("<<ComboboxSelected>>", lambda e: self.apply_preset())
        ttk.Label(opt, text="출력 형식").grid(row=0, column=2, sticky=W)
        ttk.Combobox(opt, textvariable=self.fmt, state="readonly",
                     values=config.OUTPUT_FORMATS, width=8).grid(row=0, column=3, padx=(6, 18))
        ttk.Checkbutton(opt, text="빠른 모드 (저사양·4B)", variable=self.fast).grid(row=0, column=4, sticky=W)
        r += 1

        # 3) AI에 넣을 컬럼
        ttk.Label(frm, text="AI에 넣을 컬럼").grid(row=r, column=0, sticky=(N, W), pady=4)
        self.col_list = Listbox(frm, selectmode=MULTIPLE, height=5, exportselection=False)
        self.col_list.grid(row=r, column=1, columnspan=2, sticky=(E, W), padx=8, pady=4)
        r += 1

        # 4) 파일명 패턴
        ttk.Label(frm, text="파일명 패턴").grid(row=r, column=0, sticky=W, pady=4)
        ttk.Entry(frm, textvariable=self.fname).grid(row=r, column=1, sticky=(E, W), padx=8)
        ttk.Label(frm, text="예: {이름}_성과평가서", foreground="#888").grid(row=r, column=2, sticky=W)
        r += 1

        # 5) 작성 지침
        ttk.Label(frm, text="작성 지침").grid(row=r, column=0, sticky=(N, W), pady=4)
        self.instr = Text(frm, height=4, wrap="word")
        self.instr.grid(row=r, column=1, columnspan=2, sticky=(E, W), padx=8, pady=4)
        r += 1

        # 6) 출력 폴더
        ttk.Label(frm, text="출력 폴더").grid(row=r, column=0, sticky=W, pady=4)
        ttk.Entry(frm, textvariable=self.out_dir).grid(row=r, column=1, sticky=(E, W), padx=8)
        ttk.Button(frm, text="폴더 선택…", command=self.pick_outdir).grid(row=r, column=2)
        r += 1

        # 7) 미리보기 + 1행 미리 생성
        prev = ttk.Frame(frm)
        prev.grid(row=r, column=0, columnspan=3, sticky=W, pady=(8, 2))
        ttk.Label(prev, text="미리보기").grid(row=0, column=0, sticky=W)
        ttk.Button(prev, text="1행 미리 생성", command=self.preview_one).grid(row=0, column=1, padx=12)
        r += 1
        self.tree = ttk.Treeview(frm, show="headings", height=6)
        self.tree.grid(row=r, column=0, columnspan=3, sticky=(N, S, E, W), pady=4)
        self.tree.tag_configure("bad", background="#ffe0e0")
        frm.rowconfigure(r, weight=1)
        r += 1

        # 8) 실행 줄
        runrow = ttk.Frame(frm)
        runrow.grid(row=r, column=0, columnspan=3, sticky=W, pady=6)
        self.run_btn = ttk.Button(runrow, text="일괄 생성 실행", command=self.run)
        self.run_btn.grid(row=0, column=0)
        self.cancel_btn = ttk.Button(runrow, text="취소", command=self.do_cancel, state="disabled")
        self.cancel_btn.grid(row=0, column=1, padx=8)
        self.ready_lbl = ttk.Label(runrow, text="환경 확인 중…", foreground="#888")
        self.ready_lbl.grid(row=0, column=2, padx=12)
        self.setup_btn = ttk.Button(runrow, text="환경 준비", command=self.open_setup)
        self.setup_btn.grid(row=0, column=3)
        r += 1

        self.bar = ttk.Progressbar(frm, mode="determinate")
        self.bar.grid(row=r, column=0, columnspan=3, sticky=(E, W), pady=4)
        r += 1
        ttk.Label(frm, textvariable=self.status, foreground="#444").grid(
            row=r, column=0, columnspan=3, sticky=W, pady=2)
        r += 1
        self.open_btn = ttk.Button(frm, text="출력 폴더 열기", command=self.open_folder,
                                   state="disabled")
        self.open_btn.grid(row=r, column=0, sticky=W, pady=4)

        self.apply_preset()
        self.refresh_ready(auto_open=True)

    # ---------- 프리셋 ----------
    def apply_preset(self):
        p = config.PRESETS[self.preset.get()]
        self.fname.set(p["default_filename"])
        self.instr.delete("1.0", END)
        self.instr.insert("1.0", p.get("default_instructions", ""))

    # ---------- 파일 선택 ----------
    def pick_excel(self):
        path = filedialog.askopenfilename(
            title="엑셀 명단 선택", filetypes=[("엑셀", "*.xlsx *.xlsm")])
        if not path:
            return
        self.excel.set(path)
        if not self.out_dir.get():
            self.out_dir.set(os.path.join(os.path.dirname(path), "생성문서"))
        threading.Thread(target=self._load_excel, args=(path,), daemon=True).start()

    def _load_excel(self, path):
        try:
            headers, rows = excel_io.read_excel(path)
            self.root.after(0, self._on_excel, headers, rows, None)
        except Exception as e:  # noqa: BLE001
            self.root.after(0, self._on_excel, [], [], str(e))

    def _on_excel(self, headers, rows, err):
        if err:
            messagebox.showerror("엑셀 읽기 오류", err)
            return
        self.headers, self.rows = headers, rows
        # 컬럼 리스트 갱신(기본: 전체 선택)
        self.col_list.delete(0, END)
        for h in headers:
            self.col_list.insert(END, h)
        self.col_list.select_set(0, END)
        self._refresh_preview()
        self.status.set(f"{len(rows)}개 행, 컬럼: {', '.join(headers)}")

    def _refresh_preview(self):
        self.tree.delete(*self.tree.get_children())
        cols = self.headers + ["→ 파일명"]
        self.tree["columns"] = cols
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=120, anchor=W)
        pattern = self.fname.get()
        for row in self.rows[:50]:
            fname = excel_io.fill_filename(pattern, row)
            unknown = excel_io.find_unknown_placeholders(pattern, self.headers)
            vals = [row.get(h, "") for h in self.headers] + [fname]
            tag = "bad" if (unknown or fname == "문서") else ""
            self.tree.insert("", "end", values=vals, tags=(tag,))

    def pick_outdir(self):
        d = filedialog.askdirectory(title="출력 폴더 선택")
        if d:
            self.out_dir.set(d)

    def _selected_columns(self):
        sel = [self.col_list.get(i) for i in self.col_list.curselection()]
        return sel or self.headers

    # ---------- 1행 미리 생성 ----------
    def preview_one(self):
        if not self.rows:
            messagebox.showwarning("미리보기", "먼저 엑셀을 선택하세요.")
            return
        self.status.set("1행 생성 중…")
        threading.Thread(target=self._preview_work, daemon=True).start()

    def _preview_work(self):
        model = config.MODEL_FAST if self.fast.get() else config.MODEL_FULL
        try:
            engine.ensure_server()
            engine.keep_warm(model)
            messages = engine.build_messages(
                self.preset.get(), self.instr.get("1.0", END).strip(),
                self.rows[0], columns=self._selected_columns())
            text = engine.generate(messages, model)
            self.root.after(0, lambda: (self.status.set("미리보기 완료"),
                                        self._show_preview(text)))
        except Exception:
            err = traceback.format_exc()
            self.root.after(0, lambda: (self.status.set("미리보기 오류"),
                                        messagebox.showerror("오류", err[-1200:])))

    def _show_preview(self, text):
        win = Toplevel(self.root)
        win.title("1행 미리 생성 결과")
        win.geometry("560x520")
        t = Text(win, wrap="word", padx=10, pady=10)
        t.pack(fill="both", expand=True)
        t.insert("1.0", text)

    # ---------- 환경 준비 ----------
    def refresh_ready(self, auto_open=False):
        def work():
            st = setup_env.check_readiness()
            self.root.after(0, self._apply_ready, st, auto_open)
        threading.Thread(target=work, daemon=True).start()

    def _apply_ready(self, st, auto_open):
        if st["ready"]:
            self.ready_lbl.config(text="● 준비 완료", foreground="#1a7f37")
            self.run_btn.config(state="normal")
            self.setup_btn.config(state="disabled")
        else:
            msg = ("Ollama 미설치" if not st["installed"]
                   else "Ollama 서버 꺼짐" if not st["server"] else "모델 미설치")
            self.ready_lbl.config(text=f"● {msg} — '환경 준비' 필요", foreground="#cf222e")
            self.run_btn.config(state="disabled")
            self.setup_btn.config(state="normal")
            if auto_open:
                self.open_setup()

    def open_setup(self):
        SetupDialog(self.root, on_done=lambda: self.refresh_ready())

    # ---------- 실행 ----------
    def run(self):
        if not self.rows:
            messagebox.showwarning("확인", "엑셀 명단을 선택하세요.")
            return
        out_dir = self.out_dir.get().strip()
        if not out_dir:
            messagebox.showwarning("확인", "출력 폴더를 선택하세요.")
            return
        self.cancel_flag = False
        self.run_btn.config(state="disabled")
        self.open_btn.config(state="disabled")
        self.cancel_btn.config(state="normal")
        threading.Thread(target=self._work, args=(out_dir,), daemon=True).start()

    def do_cancel(self):
        self.cancel_flag = True
        self.status.set("취소 중…")

    def _work(self, out_dir):
        model = config.MODEL_FAST if self.fast.get() else config.MODEL_FULL
        try:
            self._set_status(f"모델 준비 중… ({model})")

            def prog(done, total, name):
                self.root.after(0, self._on_prog, done, total, name)

            rows = pipeline.generate_documents(
                self.excel.get(), out_dir,
                preset_key=self.preset.get(),
                user_instructions=self.instr.get("1.0", END).strip(),
                columns=self._selected_columns(),
                filename_pattern=self.fname.get(),
                fmt=self.fmt.get(), model=model,
                progress=prog, cancel=lambda: self.cancel_flag)
            res_path = os.path.join(out_dir, f"_생성결과_{datetime.now():%Y%m%d_%H%M}.xlsx")
            pipeline.write_result_excel(rows, res_path)
            self.last_out_dir = out_dir
            self.root.after(0, self._on_done, rows, out_dir)
        except Exception:
            err = traceback.format_exc()
            self.root.after(0, self._on_error, err)

    def _on_prog(self, done, total, name):
        self.bar.config(maximum=max(total, 1), value=done)
        self.status.set(f"[{done}/{total}] 생성 중… {name}")

    def _on_done(self, rows, out_dir):
        ok = sum(1 for r in rows if r["상태"] == "완료")
        self.status.set(f"완료: {ok}/{len(rows)}건 성공  →  {out_dir}")
        self.run_btn.config(state="normal")
        self.cancel_btn.config(state="disabled")
        self.open_btn.config(state="normal")

    def _on_error(self, err):
        self.run_btn.config(state="normal")
        self.cancel_btn.config(state="disabled")
        self.status.set("오류 발생")
        messagebox.showerror("오류", err[-1500:])

    def _set_status(self, msg):
        self.root.after(0, self.status.set, msg)

    def open_folder(self):
        if self.last_out_dir and os.path.isdir(self.last_out_dir):
            os.startfile(self.last_out_dir)


class SetupDialog(Toplevel):
    """첫 실행 온보딩: 설치 → ASCII 경로 → 서버 재기동 → 모델 pull.

    (DocBatch SetupDialog 재사용 — 문구만 DocForge에 맞춤. 모델 경로를 공유하므로
    DocBatch에서 이미 받은 모델이면 다운로드를 건너뛴다.)
    """

    def __init__(self, parent, on_done=None):
        super().__init__(parent)
        self.on_done = on_done
        self.title("환경 준비")
        self.geometry("560x360")
        self.transient(parent)
        self.grab_set()

        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="온디바이스 실행에 필요한 환경을 자동으로 준비합니다.",
                  font=("", 10, "bold")).pack(anchor=W)
        ttk.Label(frm, text=f"· Ollama 설치 확인\n· 모델 경로 ASCII 고정 ({config.ASCII_MODELS_DIR})\n"
                            f"· 서버 재기동\n· 모델 다운로드 ({config.MODEL_FULL}, 약 15GB · RAM 16GB+ 권장 · 이미 있으면 건너뜀)",
                  foreground="#444", justify="left").pack(anchor=W, pady=8)

        self.bar = ttk.Progressbar(frm, mode="determinate")
        self.bar.pack(fill="x", pady=6)
        self.msg = StringVar(value="준비되면 아래 버튼을 누르세요.")
        ttk.Label(frm, textvariable=self.msg, foreground="#444").pack(anchor=W)

        btns = ttk.Frame(frm)
        btns.pack(side="bottom", anchor=E, pady=8)
        self.start_btn = ttk.Button(btns, text="자동 준비 시작", command=self.start)
        self.start_btn.grid(row=0, column=0, padx=4)
        self.close_btn = ttk.Button(btns, text="닫기", command=self.destroy)
        self.close_btn.grid(row=0, column=1, padx=4)

    def _set(self, pct, text):
        if pct is not None and pct >= 0:
            self.bar.config(value=pct)
        self.msg.set(text)

    def start(self):
        self.start_btn.config(state="disabled")
        self.close_btn.config(state="disabled")
        threading.Thread(target=self._flow, daemon=True).start()

    def _ui(self, pct, text):
        self.after(0, self._set, pct, text)

    def _flow(self):
        try:
            if not setup_env.is_installed():
                if setup_env.winget_available():
                    self._ui(0, "Ollama 설치 중… (winget, 수 분 소요)")
                    ok, m = setup_env.install_ollama_winget()
                    if not ok:
                        self._ui(-1, f"{m}. 다운로드 페이지를 엽니다.")
                        webbrowser.open(setup_env.OLLAMA_DOWNLOAD_URL)
                        return self._fail("Ollama 설치 후 다시 시도하세요.")
                else:
                    webbrowser.open(setup_env.OLLAMA_DOWNLOAD_URL)
                    return self._fail("Ollama를 수동 설치한 뒤 다시 시도하세요.")

            self._ui(0, "모델 경로 고정 및 서버 재기동 중…")
            if not setup_env.restart_server_ascii():
                return self._fail("Ollama 서버 기동 실패.")

            # 품질(26B)·빠른(4B) 두 모델 모두 확보 → 어느 모드를 켜도 '모델 없음' 방지.
            for mdl in (config.MODEL_FULL, config.MODEL_FAST):
                if not setup_env.has_model(mdl):
                    self._ui(0, f"모델 다운로드 중… {mdl}")
                    ok = setup_env.pull_model(
                        mdl,
                        progress=lambda p, s: self._ui(p, f"다운로드 {p}% · {s}" if p >= 0 else f"오류: {s}"))
                    if not ok:
                        return self._fail(f"모델 다운로드 실패: {mdl}")

            self._ui(100, "준비 완료!")
            self.after(0, self._success)
        except Exception:
            self._fail(traceback.format_exc()[-600:])

    def _success(self):
        self.bar.config(value=100)
        if self.on_done:
            self.on_done()
        messagebox.showinfo("완료", "환경 준비가 끝났습니다.", parent=self)
        self.destroy()

    def _fail(self, text):
        self.after(0, lambda: (self._set(-1, text),
                               self.start_btn.config(state="normal"),
                               self.close_btn.config(state="normal")))


def main():
    root = Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
