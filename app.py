from __future__ import annotations

import argparse
import os
import queue
import subprocess
import sys
import threading
from pathlib import Path

from converter import (
    ConversionCancelled,
    ConversionError,
    ConversionOptions,
    convert_pdf,
    inspect_pdf,
    pdf_requires_password,
)


APP_NAME = "PDF 转 Word"
VERSION = "1.0.0"


def _open_path(path: Path) -> None:
    if sys.platform == "win32":
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def run_cli(args: argparse.Namespace) -> int:
    source = Path(args.input)
    output = Path(args.output) if args.output else source.with_suffix(".docx")

    def show_progress(current: int, total: int, message: str) -> None:
        print(f"[{current}/{total}] {message}")

    try:
        result = convert_pdf(
            source,
            output,
            ConversionOptions(
                mode=args.mode,
                pages=args.pages,
                password=args.password,
                ocr_language=args.ocr_language,
                ocr_dpi=args.ocr_dpi,
            ),
            callback=show_progress,
        )
    except ConversionError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1
    print(f"已生成：{result}")
    return 0


def run_gui() -> int:
    import tkinter as tk
    from tkinter import filedialog, messagebox, simpledialog, ttk
    from tkinter.scrolledtext import ScrolledText

    if sys.platform == "win32":
        try:
            import ctypes

            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

    root = tk.Tk()
    root.title(f"{APP_NAME}  {VERSION}")
    root.geometry("790x660")
    root.minsize(720, 610)

    style = ttk.Style(root)
    if sys.platform == "win32" and "vista" in style.theme_names():
        style.theme_use("vista")
    elif "clam" in style.theme_names():
        style.theme_use("clam")
    style.configure("Title.TLabel", font=("TkDefaultFont", 20, "bold"))
    style.configure("Subtitle.TLabel", foreground="#5f6368")
    style.configure("Accent.TButton", font=("TkDefaultFont", 10, "bold"), padding=(14, 9))
    style.configure("Card.TLabelframe", padding=14)
    style.configure("Card.TLabelframe.Label", font=("TkDefaultFont", 10, "bold"))

    input_var = tk.StringVar()
    output_var = tk.StringVar()
    pages_var = tk.StringVar()
    mode_var = tk.StringVar(value="保留版式（推荐，可编辑）")
    ocr_lang_var = tk.StringVar(value="eng")
    status_var = tk.StringVar(value="请选择一个 PDF 文件")
    info_var = tk.StringVar(value="尚未选择文件")
    mode_help_var = tk.StringVar()
    events: queue.Queue[tuple] = queue.Queue()
    cancel_event = threading.Event()
    current_output: list[Path | None] = [None]
    worker: list[threading.Thread | None] = [None]

    mode_ids = {
        "保留版式（推荐，可编辑）": "layout",
        "完全一致（每页作为图片）": "exact",
        "OCR 识别（扫描件可编辑）": "ocr",
    }
    mode_help = {
        "layout": "尽量保留原来的文字、图片、表格和排版；适合绝大多数电子 PDF。",
        "exact": "外观最稳定，但页面内容会作为图片插入，文字不可编辑。",
        "ocr": "识别扫描件中的文字并写入 Word；需要安装 Tesseract，复杂排版可能简化。",
    }

    outer = ttk.Frame(root, padding=22)
    outer.pack(fill="both", expand=True)
    ttk.Label(outer, text=APP_NAME, style="Title.TLabel").pack(anchor="w")
    ttk.Label(
        outer,
        text="离线转换，不上传文件；无需安装 Microsoft Word 或 Adobe Acrobat。",
        style="Subtitle.TLabel",
    ).pack(anchor="w", pady=(3, 18))

    files_box = ttk.LabelFrame(outer, text="文件", style="Card.TLabelframe")
    files_box.pack(fill="x")
    files_box.columnconfigure(1, weight=1)
    ttk.Label(files_box, text="PDF 文件").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
    input_entry = ttk.Entry(files_box, textvariable=input_var)
    input_entry.grid(row=0, column=1, sticky="ew", pady=5)
    ttk.Label(files_box, text="DOCX 输出").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=5)
    output_entry = ttk.Entry(files_box, textvariable=output_var)
    output_entry.grid(row=1, column=1, sticky="ew", pady=5)

    def refresh_info(*_args) -> None:
        raw = input_var.get().strip()
        if not raw:
            info_var.set("尚未选择文件")
            return
        path = Path(raw)
        if not path.is_file():
            info_var.set("文件不存在")
            return
        try:
            if pdf_requires_password(path):
                info_var.set(f"{path.name} · 已加密（转换时需要密码）")
                return
            info = inspect_pdf(path)
            size_mb = info.size_bytes / (1024 * 1024)
            info_var.set(f"{path.name} · {info.pages} 页 · {size_mb:.2f} MB")
        except ConversionError as exc:
            info_var.set(str(exc))

    def choose_input() -> None:
        filename = filedialog.askopenfilename(title="选择 PDF", filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")])
        if filename:
            input_var.set(filename)
            if not output_var.get().strip():
                output_var.set(str(Path(filename).with_suffix(".docx")))
            refresh_info()

    def choose_output() -> None:
        source = Path(input_var.get()) if input_var.get().strip() else Path("converted.pdf")
        filename = filedialog.asksaveasfilename(
            title="保存 DOCX",
            initialfile=f"{source.stem}.docx",
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")],
        )
        if filename:
            output_var.set(filename)

    ttk.Button(files_box, text="浏览…", command=choose_input).grid(row=0, column=2, padx=(10, 0), pady=5)
    ttk.Button(files_box, text="选择…", command=choose_output).grid(row=1, column=2, padx=(10, 0), pady=5)
    ttk.Label(files_box, textvariable=info_var, style="Subtitle.TLabel").grid(
        row=2, column=1, columnspan=2, sticky="w", pady=(2, 0)
    )

    options_box = ttk.LabelFrame(outer, text="转换选项", style="Card.TLabelframe")
    options_box.pack(fill="x", pady=(14, 0))
    options_box.columnconfigure(1, weight=1)
    ttk.Label(options_box, text="转换模式").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
    mode_combo = ttk.Combobox(options_box, textvariable=mode_var, values=list(mode_ids), state="readonly")
    mode_combo.grid(row=0, column=1, columnspan=3, sticky="ew", pady=5)
    ttk.Label(options_box, textvariable=mode_help_var, style="Subtitle.TLabel", wraplength=600).grid(
        row=1, column=1, columnspan=3, sticky="w", pady=(0, 7)
    )
    ttk.Label(options_box, text="页码范围").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=5)
    ttk.Entry(options_box, textvariable=pages_var, width=22).grid(row=2, column=1, sticky="w", pady=5)
    ttk.Label(options_box, text="留空=全部，例如 1-3,5", style="Subtitle.TLabel").grid(
        row=2, column=2, columnspan=2, sticky="w", padx=(10, 0), pady=5
    )
    ocr_label = ttk.Label(options_box, text="OCR 语言")
    ocr_combo = ttk.Combobox(
        options_box,
        textvariable=ocr_lang_var,
        values=("eng", "chi_sim+eng", "chi_tra+eng", "jpn+eng", "kor+eng", "fra+eng", "deu+eng"),
        width=18,
    )
    ocr_hint = ttk.Label(options_box, text="需安装对应语言包", style="Subtitle.TLabel")
    ocr_label.grid(row=3, column=0, sticky="w", padx=(0, 10), pady=5)
    ocr_combo.grid(row=3, column=1, sticky="w", pady=5)
    ocr_hint.grid(row=3, column=2, sticky="w", padx=(10, 0), pady=5)

    def update_mode(*_args) -> None:
        mode = mode_ids[mode_var.get()]
        mode_help_var.set(mode_help[mode])
        state = "normal" if mode == "ocr" else "disabled"
        ocr_combo.configure(state=state)
        if state == "normal":
            ocr_label.configure(state="normal")
            ocr_hint.configure(state="normal")
        else:
            ocr_label.configure(state="disabled")
            ocr_hint.configure(state="disabled")

    mode_combo.bind("<<ComboboxSelected>>", update_mode)
    update_mode()

    progress = ttk.Progressbar(outer, mode="determinate", maximum=100)
    progress.pack(fill="x", pady=(18, 5))
    ttk.Label(outer, textvariable=status_var).pack(anchor="w")

    log = ScrolledText(outer, height=6, wrap="word", state="disabled", font=("TkFixedFont", 9))
    log.pack(fill="both", expand=True, pady=(8, 12))

    button_row = ttk.Frame(outer)
    button_row.pack(fill="x")
    convert_button = ttk.Button(button_row, text="开始转换", style="Accent.TButton")
    convert_button.pack(side="left")
    cancel_button = ttk.Button(button_row, text="取消", state="disabled")
    cancel_button.pack(side="left", padx=(9, 0))
    open_button = ttk.Button(button_row, text="打开生成的 DOCX", state="disabled")
    open_button.pack(side="right")

    def append_log(text: str) -> None:
        log.configure(state="normal")
        log.insert("end", text.rstrip() + "\n")
        log.see("end")
        log.configure(state="disabled")

    def set_running(running: bool) -> None:
        state = "disabled" if running else "normal"
        convert_button.configure(state=state)
        input_entry.configure(state=state)
        output_entry.configure(state=state)
        mode_combo.configure(state="disabled" if running else "readonly")
        cancel_button.configure(state="normal" if running else "disabled")
        if not running:
            update_mode()

    def do_convert() -> None:
        source_text = input_var.get().strip()
        output_text = output_var.get().strip()
        if not source_text:
            messagebox.showwarning(APP_NAME, "请先选择 PDF 文件。")
            return
        source = Path(source_text)
        if not source.is_file():
            messagebox.showerror(APP_NAME, "所选 PDF 文件不存在。")
            return
        if not output_text:
            output_text = str(source.with_suffix(".docx"))
            output_var.set(output_text)
        destination = Path(output_text)
        if destination.suffix.lower() != ".docx":
            destination = destination.with_suffix(".docx")
            output_var.set(str(destination))
        if destination.exists() and not messagebox.askyesno(APP_NAME, "输出文件已经存在，是否覆盖？"):
            return

        password = ""
        try:
            if pdf_requires_password(source):
                password = simpledialog.askstring(APP_NAME, "这个 PDF 已加密，请输入打开密码：", show="*") or ""
                if not password:
                    return
                inspect_pdf(source, password)
        except ConversionError as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return

        cancel_event.clear()
        current_output[0] = None
        open_button.configure(state="disabled")
        progress["value"] = 0
        status_var.set("准备转换…")
        append_log(f"输入：{source}")
        append_log(f"输出：{destination}")
        set_running(True)

        options = ConversionOptions(
            mode=mode_ids[mode_var.get()],
            pages=pages_var.get(),
            password=password,
            ocr_language=ocr_lang_var.get().strip() or "eng",
        )

        def callback(current: int, total: int, message: str) -> None:
            events.put(("progress", current, total, message))

        def task() -> None:
            try:
                result = convert_pdf(source, destination, options, callback, cancel_event)
                events.put(("done", result))
            except ConversionCancelled as exc:
                events.put(("cancelled", str(exc)))
            except Exception as exc:
                events.put(("error", str(exc)))

        worker[0] = threading.Thread(target=task, daemon=True)
        worker[0].start()

    def cancel() -> None:
        cancel_event.set()
        status_var.set("正在取消，请稍候…")
        append_log("已请求取消。版式分析阶段可能需要等待当前步骤结束。")
        cancel_button.configure(state="disabled")

    def poll_events() -> None:
        try:
            while True:
                event = events.get_nowait()
                kind = event[0]
                if kind == "progress":
                    _, current, total, message = event
                    progress["value"] = current / max(1, total) * 100
                    status_var.set(message)
                    append_log(message)
                elif kind == "done":
                    result = Path(event[1])
                    current_output[0] = result
                    progress["value"] = 100
                    status_var.set("转换完成")
                    append_log(f"完成：{result}")
                    set_running(False)
                    open_button.configure(state="normal")
                    messagebox.showinfo(APP_NAME, f"转换完成！\n\n{result}")
                elif kind == "cancelled":
                    status_var.set("转换已取消")
                    append_log(event[1])
                    set_running(False)
                elif kind == "error":
                    status_var.set("转换失败")
                    append_log(f"错误：{event[1]}")
                    set_running(False)
                    messagebox.showerror(APP_NAME, event[1])
        except queue.Empty:
            pass
        root.after(120, poll_events)

    def open_result() -> None:
        if current_output[0] and current_output[0].exists():
            _open_path(current_output[0])

    def on_close() -> None:
        if worker[0] and worker[0].is_alive():
            if not messagebox.askyesno(APP_NAME, "转换仍在进行，确定退出吗？"):
                return
            cancel_event.set()
        root.destroy()

    convert_button.configure(command=do_convert)
    cancel_button.configure(command=cancel)
    open_button.configure(command=open_result)
    input_var.trace_add("write", lambda *_: root.after(250, refresh_info))
    root.protocol("WM_DELETE_WINDOW", on_close)
    root.after(120, poll_events)
    root.update_idletasks()
    x = max(0, (root.winfo_screenwidth() - root.winfo_width()) // 2)
    y = max(0, (root.winfo_screenheight() - root.winfo_height()) // 3)
    root.geometry(f"+{x}+{y}")
    root.mainloop()
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="在 Windows 和 macOS 上将 PDF 转换为 DOCX。")
    parser.add_argument("input", nargs="?", help="输入 PDF；省略时打开图形界面")
    parser.add_argument("-o", "--output", help="输出 DOCX 路径")
    parser.add_argument("--mode", choices=("layout", "exact", "ocr"), default="layout")
    parser.add_argument("--pages", default="", help="页码范围，例如 1-3,5；留空表示全部")
    parser.add_argument("--password", default="", help="PDF 打开密码")
    parser.add_argument("--ocr-language", default="eng", help="Tesseract 语言代码，例如 chi_sim+eng")
    parser.add_argument("--ocr-dpi", type=int, default=220, help="OCR 渲染 DPI")
    parser.add_argument("--version", action="version", version=f"%(prog)s {VERSION}")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    return run_cli(args) if args.input else run_gui()


if __name__ == "__main__":
    raise SystemExit(main())

