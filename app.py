"""A small Russian-language desktop UI. Run with python app.py on Windows."""
from __future__ import annotations

import json
import os
import queue
import shutil
import tempfile
import threading
import uuid
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from converter import convert

CONFIG = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "ExcelSignaturePDF" / "settings.json"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel → PDF с подписью")
        self.geometry("720x460")
        self.minsize(650, 440)
        self.temp = tempfile.TemporaryDirectory(prefix="signature-preview-", ignore_cleanup_errors=True)
        self.preview = Path(self.temp.name) / "preview.pdf"
        self.events = queue.Queue()
        self.busy = False
        self.excel = tk.StringVar()
        self.signature = tk.StringVar()
        self.sheet = tk.StringVar()
        self.width = tk.StringVar(value="45")
        self.landscape = tk.BooleanVar(value=False)
        self.print_area = tk.BooleanVar(value=False)
        self.status = tk.StringVar(value="Выберите таблицу и картинку подписи.")
        try:
            config = json.loads(CONFIG.read_text(encoding="utf-8"))
            self.signature.set(config.get("signature", ""))
            self.width.set(str(config.get("width", 45)))
        except (OSError, ValueError, TypeError, AttributeError):
            pass
        panel = ttk.Frame(self, padding=22)
        panel.pack(fill="both", expand=True)
        panel.columnconfigure(1, weight=1)
        ttk.Label(panel, text="Таблица и подпись → PDF", font=("Segoe UI", 18, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 18))
        for row, title, var, command in (
            (1, "Excel (.xlsx)", self.excel, self.choose_excel),
            (2, "Подпись", self.signature, self.choose_signature),
        ):
            ttk.Label(panel, text=title).grid(row=row, column=0, sticky="w", padx=(0, 12))
            ttk.Entry(panel, textvariable=var).grid(row=row, column=1, sticky="ew", pady=6)
            ttk.Button(panel, text="Выбрать…", command=command).grid(row=row, column=2, padx=(8, 0))
        ttk.Label(panel, text="Название листа").grid(row=3, column=0, sticky="w")
        ttk.Entry(panel, textvariable=self.sheet).grid(row=3, column=1, sticky="ew", pady=6)
        ttk.Label(panel, text="Пусто — первый видимый лист").grid(row=4, column=1, sticky="w")
        ttk.Label(panel, text="Ширина подписи, мм").grid(row=5, column=0, sticky="w", pady=10)
        ttk.Spinbox(panel, from_=20, to=90, width=8, textvariable=self.width).grid(row=5, column=1, sticky="w")
        ttk.Checkbutton(panel, text="Альбомный A4 (для широких таблиц)", variable=self.landscape).grid(row=6, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(panel, text="Использовать заданную в Excel область печати", variable=self.print_area).grid(row=7, column=0, columnspan=3, sticky="w", pady=6)
        ttk.Label(panel, text="Подпись — справа внизу последней страницы. Текст можно копировать.", wraplength=640).grid(row=8, column=0, columnspan=3, sticky="w", pady=10)
        actions = ttk.Frame(panel)
        actions.grid(row=9, column=0, columnspan=3, sticky="w", pady=8)
        self.create_button = ttk.Button(actions, text="Создать и посмотреть", command=self.start)
        self.create_button.pack(side="left")
        self.save_button = ttk.Button(actions, text="Сохранить PDF…", command=self.save, state="disabled")
        self.save_button.pack(side="left", padx=10)
        ttk.Label(panel, textvariable=self.status, wraplength=640).grid(row=10, column=0, columnspan=3, sticky="w", pady=8)
        for var in (self.excel, self.signature, self.sheet, self.width, self.landscape, self.print_area):
            var.trace_add("write", self.invalidate)
        self.protocol("WM_DELETE_WINDOW", self.close)
        self.after(100, self.poll)

    def invalidate(self, *_):
        self.save_button.configure(state="disabled")
        if not self.busy:
            self.status.set("Параметры изменены. Нажмите «Создать и посмотреть».")

    def choose_excel(self):
        name = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if name:
            self.excel.set(name)

    def choose_signature(self):
        name = filedialog.askopenfilename(filetypes=[("Подпись PNG / JPEG", "*.png *.jpg *.jpeg")])
        if name:
            self.signature.set(name)

    def values(self):
        return (self.excel.get(), self.signature.get(), self.sheet.get(),
                self.width.get(), self.landscape.get(), self.print_area.get())

    def start(self):
        if self.busy:
            return
        self.started_values = self.values()
        source, signature, sheet, width, landscape, print_area = self.started_values
        if not source or not signature:
            messagebox.showerror("Не выбран файл", "Выберите Excel-файл и картинку подписи.")
            return
        try:
            width = float(width.replace(",", "."))
        except ValueError:
            messagebox.showerror("Размер подписи", "Введите ширину числом, например 45.")
            return
        self.busy = True
        self.preview = Path(self.temp.name) / (uuid.uuid4().hex + ".pdf")
        self.create_button.configure(state="disabled")
        self.save_button.configure(state="disabled")
        self.status.set("Excel готовит PDF. Дождитесь завершения…")

        def worker():
            try:
                result = convert(Path(source), Path(signature), self.preview,
                                 sheet, width, landscape, print_area)
                self.events.put((True, result))
            except Exception as exc:
                self.events.put((False, str(exc)))
        threading.Thread(target=worker, daemon=True).start()

    def poll(self):
        try:
            ok, result = self.events.get_nowait()
        except queue.Empty:
            self.after(100, self.poll)
            return
        self.busy = False
        self.create_button.configure(state="normal")
        if ok and self.values() != self.started_values:
            self.status.set("Параметры изменились во время создания. Создайте PDF заново.")
        elif ok:
            pages, sheet = result
            self.save_button.configure(state="normal")
            self.status.set(f"Готово: лист «{sheet}», страниц: {pages}. Проверьте PDF и сохраните.")
            try:
                CONFIG.parent.mkdir(parents=True, exist_ok=True)
                CONFIG.write_text(json.dumps({"signature": self.signature.get(), "width": self.width.get()}, ensure_ascii=False), encoding="utf-8")
            except OSError:
                pass  # Settings persistence must not prevent saving the PDF.
            try:
                os.startfile(self.preview)
            except (OSError, AttributeError):
                messagebox.showinfo("Предпросмотр", "Не удалось открыть просмотрщик PDF. Сохраните файл и откройте его вручную.")
        else:
            self.status.set("Не удалось создать PDF.")
            messagebox.showerror("Ошибка", str(result))
        self.after(100, self.poll)

    def save(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
            initialfile=Path(self.excel.get()).stem + "_signed.pdf",
        )
        if not filename:
            return
        try:
            # Write beside the destination; preserve an existing file on failure.
            target = Path(filename)
            if target.suffix.lower() != ".pdf":
                raise OSError("Сохраните результат с расширением .pdf.")
            with tempfile.TemporaryDirectory(dir=target.parent) as temp:
                staged = Path(temp) / "result.pdf"
                shutil.copyfile(self.preview, staged)
                os.replace(staged, target)
            self.status.set(f"Сохранено: {filename}")
        except OSError as exc:
            messagebox.showerror("Ошибка сохранения", str(exc))

    def close(self):
        if self.busy:
            messagebox.showinfo("Идёт экспорт", "Дождитесь завершения экспорта Excel.")
            return
        try:
            self.temp.cleanup()
        except OSError:
            pass  # A PDF viewer may hold the preview open on Windows.
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
