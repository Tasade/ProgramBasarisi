"""Word -> PDF Donusturucu (Windows masaustu uygulamasi).

.doc, .docx, .docm, .dot, .dotx, .rtf, .odt gibi Microsoft Word'un actigi
TUM belge bicimlerini/surumlerini PDF'e cevirir.

Nasil calisir:
    Bilgisayarda kurulu olan Microsoft Word, COM otomasyonu (pywin32) ile
    arka planda (gorunmez sekilde) baslatilir; her belge Word'un kendisi
    tarafindan acilip PDF olarak kaydedilir. Boylece donusum, Word'un o
    surumde belgeyi nasil goruntulediginin birebir ayni sonucunu verir ve
    eski .doc (Word 97-2003) dosyalari da dahil olmak uzere Word'un
    destekledigi her surum guvenilir sekilde calisir.

Gereksinimler:
    - Windows isletim sistemi
    - Bilgisayarda Microsoft Word (herhangi bir surum) kurulu olmali
    - pip install pywin32

Tek dosya .exe olarak derlemek icin (Windows uzerinde calistirin):
    pip install -r requirements.txt pyinstaller
    pyinstaller --onefile --windowed --name WordToPDFDonusturucu word_to_pdf_gui.py
    -> cikti dist/WordToPDFDonusturucu.exe olarak tek bir dosyadir.
"""

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_TITLE = "Word'den PDF'e Donusturucu"
WORD_EXTENSIONS = (".doc", ".docx", ".docm", ".dot", ".dotx", ".rtf", ".odt")
FILE_DIALOG_TYPES = [
    ("Word belgeleri", " ".join(f"*{ext}" for ext in WORD_EXTENSIONS)),
    ("Tum dosyalar", "*.*"),
]
WD_FORMAT_PDF = 17  # Word'un wdFormatPDF sabiti


class ConversionError(Exception):
    """Word baslatilamadigi veya bulunamadigi durumlarda firlatilir."""


def convert_files(file_paths, output_dir, progress_queue):
    """Ayri bir is parcaciginda (thread) calisir; ilerlemeyi progress_queue'ya yazar."""
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    word = None
    try:
        try:
            word = win32.DispatchEx("Word.Application")
        except Exception as exc:
            progress_queue.put(("fatal", None, None, str(exc)))
            return

        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone

        for index, path in enumerate(file_paths, start=1):
            file_name = os.path.basename(path)
            progress_queue.put(("status", index, file_name, "Donusturuluyor..."))
            doc = None
            try:
                target_dir = output_dir or os.path.dirname(os.path.abspath(path))
                base_name = os.path.splitext(file_name)[0]
                pdf_path = os.path.join(target_dir, base_name + ".pdf")

                doc = word.Documents.Open(
                    os.path.abspath(path),
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    ConfirmConversions=False,
                )
                doc.SaveAs(os.path.abspath(pdf_path), FileFormat=WD_FORMAT_PDF)
                progress_queue.put(("ok", index, file_name, pdf_path))
            except Exception as exc:
                progress_queue.put(("error", index, file_name, str(exc)))
            finally:
                if doc is not None:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
        progress_queue.put(("done", None, None, None))


class WordToPdfApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("680x520")
        self.minsize(600, 440)

        self.file_paths = []
        self.output_dir = None
        self.progress_queue = queue.Queue()
        self.worker_thread = None

        self._build_widgets()
        self._warn_if_not_windows()

    # ------------------------------------------------------------------ UI
    def _build_widgets(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        ttk.Button(top, text="Dosya Ekle...", command=self.add_files).pack(side="left")
        ttk.Button(top, text="Secileni Kaldir", command=self.remove_selected).pack(
            side="left", padx=(8, 0)
        )
        ttk.Button(top, text="Listeyi Temizle", command=self.clear_files).pack(
            side="left", padx=(8, 0)
        )

        list_frame = ttk.Frame(self, padding=(10, 0))
        list_frame.pack(fill="both", expand=True)

        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")

        self.listbox = tk.Listbox(
            list_frame, selectmode="extended", yscrollcommand=scrollbar.set
        )
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.listbox.yview)

        output_frame = ttk.LabelFrame(self, text="Cikti klasoru", padding=10)
        output_frame.pack(fill="x", padx=10, pady=(10, 0))

        self.output_mode = tk.StringVar(value="same")
        ttk.Radiobutton(
            output_frame,
            text="Kaynak dosya ile ayni klasore kaydet",
            variable=self.output_mode,
            value="same",
            command=self._on_output_mode_change,
        ).pack(anchor="w")

        custom_row = ttk.Frame(output_frame)
        custom_row.pack(fill="x", anchor="w")
        ttk.Radiobutton(
            custom_row,
            text="Farkli klasore kaydet:",
            variable=self.output_mode,
            value="custom",
            command=self._on_output_mode_change,
        ).pack(side="left")
        self.output_dir_label = ttk.Label(
            custom_row, text="(secilmedi)", foreground="#555555"
        )
        self.output_dir_label.pack(side="left", padx=(6, 6))
        self.choose_dir_button = ttk.Button(
            custom_row, text="Klasor Sec...", command=self.choose_output_dir, state="disabled"
        )
        self.choose_dir_button.pack(side="left")

        action_frame = ttk.Frame(self, padding=10)
        action_frame.pack(fill="x")

        self.convert_button = ttk.Button(
            action_frame, text="PDF'e Donustur", command=self.start_conversion
        )
        self.convert_button.pack(side="left")

        self.progress = ttk.Progressbar(action_frame, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=(10, 0))

        log_frame = ttk.LabelFrame(self, text="Durum", padding=10)
        log_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        log_scroll = ttk.Scrollbar(log_frame)
        log_scroll.pack(side="right", fill="y")

        self.log_text = tk.Text(
            log_frame, height=8, state="disabled", yscrollcommand=log_scroll.set
        )
        self.log_text.pack(side="left", fill="both", expand=True)
        log_scroll.config(command=self.log_text.yview)

    def _warn_if_not_windows(self):
        if os.name != "nt":
            messagebox.showwarning(
                APP_TITLE,
                "Bu uygulama Microsoft Word'u Windows COM otomasyonu ile "
                "kullanir; bu nedenle sadece Windows uzerinde ve Microsoft "
                "Word kuruluyken calisir.",
            )

    # -------------------------------------------------------------- events
    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Word dosyalarini secin", filetypes=FILE_DIALOG_TYPES
        )
        for path in paths:
            if path not in self.file_paths:
                self.file_paths.append(path)
                self.listbox.insert("end", path)

    def remove_selected(self):
        for index in reversed(self.listbox.curselection()):
            self.listbox.delete(index)
            del self.file_paths[index]

    def clear_files(self):
        self.listbox.delete(0, "end")
        self.file_paths.clear()

    def _on_output_mode_change(self):
        if self.output_mode.get() == "custom":
            self.choose_dir_button.config(state="normal")
            if not self.output_dir:
                self.choose_output_dir()
        else:
            self.choose_dir_button.config(state="disabled")

    def choose_output_dir(self):
        directory = filedialog.askdirectory(title="Cikti klasorunu secin")
        if directory:
            self.output_dir = directory
            self.output_dir_label.config(text=directory)
        elif not self.output_dir:
            self.output_mode.set("same")
            self.choose_dir_button.config(state="disabled")

    def _log(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def start_conversion(self):
        if not self.file_paths:
            messagebox.showinfo(APP_TITLE, "Once donusturulecek en az bir Word dosyasi ekleyin.")
            return
        if self.output_mode.get() == "custom" and not self.output_dir:
            messagebox.showinfo(APP_TITLE, "Lutfen bir cikti klasoru secin.")
            return

        self.convert_button.config(state="disabled")
        self.progress.config(maximum=len(self.file_paths), value=0)
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

        output_dir = self.output_dir if self.output_mode.get() == "custom" else None
        self.worker_thread = threading.Thread(
            target=convert_files,
            args=(list(self.file_paths), output_dir, self.progress_queue),
            daemon=True,
        )
        self.worker_thread.start()
        self.after(100, self._poll_queue)

    def _poll_queue(self):
        try:
            while True:
                kind, index, name, detail = self.progress_queue.get_nowait()
                if kind == "status":
                    self._log(f"[{index}/{len(self.file_paths)}] {name}: {detail}")
                elif kind == "ok":
                    self.progress.config(value=index)
                    self._log(f"[{index}/{len(self.file_paths)}] {name}: Tamamlandi -> {detail}")
                elif kind == "error":
                    self.progress.config(value=index)
                    self._log(f"[{index}/{len(self.file_paths)}] {name}: HATA - {detail}")
                elif kind == "fatal":
                    messagebox.showerror(
                        APP_TITLE,
                        "Microsoft Word baslatilamadi. Bilgisayarinizda Microsoft "
                        f"Word kurulu oldugundan emin olun.\n\nDetay: {detail}",
                    )
                    self.convert_button.config(state="normal")
                    return
                elif kind == "done":
                    self.convert_button.config(state="normal")
                    self._log("Islem tamamlandi.")
                    messagebox.showinfo(APP_TITLE, "Donusturme islemi tamamlandi.")
                    return
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)


def main():
    app = WordToPdfApp()
    app.mainloop()


if __name__ == "__main__":
    main()
