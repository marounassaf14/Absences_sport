import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from tkinter import ttk

import webbrowser
import urllib.parse

from attendance import find_students_with_absences
from mailer import SUBJECT, BODY_TEMPLATE, format_dates


class AttendanceApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ENSTA • Absence Management")
        self.geometry("1100x700")
        self.configure(bg="#f4f6f8")

        self.excel_path = None
        self.students = []

        self._setup_styles()
        self._build_header()
        self._build_controls()
        self._build_main_area()

    # ---------- STYLES ----------
    def _setup_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("Header.TLabel",
                        font=("Segoe UI", 20, "bold"),
                        background="#1f3a5f",
                        foreground="white")

        style.configure("SubHeader.TLabel",
                        font=("Segoe UI", 10),
                        background="#1f3a5f",
                        foreground="white")

        style.configure("Card.TFrame",
                        background="white")

        style.configure("Primary.TButton",
                        font=("Segoe UI", 10, "bold"),
                        padding=8)

        style.configure("Treeview",
                        font=("Segoe UI", 9),
                        rowheight=26)

        style.configure("Treeview.Heading",
                        font=("Segoe UI", 9, "bold"))

    # ---------- HEADER ----------
    def _build_header(self):
        header = tk.Frame(self, bg="#1f3a5f", height=80)
        header.pack(fill="x")

        tk.Label(
            header,
            text="ENSTA – Absence Management",
            font=("Segoe UI", 20, "bold"),
            bg="#1f3a5f",
            fg="white"
        ).pack(anchor="w", padx=20, pady=(15, 0))

        tk.Label(
            header,
            text="Scan attendance • Review absences • Contact students",
            font=("Segoe UI", 10),
            bg="#1f3a5f",
            fg="#dce3ea"
        ).pack(anchor="w", padx=20)

    # ---------- TOP CONTROLS ----------
    def _build_controls(self):
        card = ttk.Frame(self, style="Card.TFrame", padding=15)
        card.pack(fill="x", padx=20, pady=15)

        # File row
        file_row = tk.Frame(card, bg="white")
        file_row.pack(fill="x")

        self.file_label = tk.Label(
            file_row,
            text="No Excel file selected",
            font=("Segoe UI", 9),
            bg="white",
            fg="#555"
        )
        self.file_label.pack(side="left", expand=True, fill="x")

        ttk.Button(
            file_row,
            text="Choose Excel File",
            command=self.choose_file
        ).pack(side="right")

        # Options row
        options = tk.Frame(card, bg="white")
        options.pack(fill="x", pady=(10, 0))

        '''tk.Label(options, text="Minimum absences:", bg="white").pack(side="left")
        self.min_abs = tk.IntVar(value=2)
        ttk.Entry(options, textvariable=self.min_abs, width=6).pack(side="left", padx=8)'''

        ttk.Button(
            options,
            text="Scan Attendance",
            style="Primary.TButton",
            command=self.scan
        ).pack(side="right")

    # ---------- MAIN AREA ----------
    def _build_main_area(self):
        main = tk.PanedWindow(self, orient=tk.HORIZONTAL, bg="#f4f6f8", sashwidth=4)
        main.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        # Left card (table)
        left = ttk.Frame(main, style="Card.TFrame", padding=10)
        main.add(left, width=720)

        tk.Label(
            left,
            text="Absent students",
            font=("Segoe UI", 11, "bold"),
            bg="white"
        ).pack(anchor="w", pady=(0, 8))

        columns = ("Sport", "Nom", "Prénom", "Email", "Absences", "Dates")
        self.tree = ttk.Treeview(left, columns=columns, show="headings", selectmode="extended")

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="w")

        self.tree.column("Email", width=220)
        self.tree.column("Absences", width=80)
        self.tree.pack(fill="both", expand=True)

        # Buttons
        btns = tk.Frame(left, bg="white")
        btns.pack(fill="x", pady=10)

        ttk.Button(btns, text="Select all", command=self.select_all).pack(side="left")
        ttk.Button(btns, text="Clear selection", command=self.clear_selection).pack(side="left", padx=8)

        ttk.Button(
            btns,
            text="Open Outlook compose (selected)",
            style="Primary.TButton",
            command=self.open_compose_selected
        ).pack(side="right")

        # Right card (log)
        right = ttk.Frame(main, style="Card.TFrame", padding=10)
        main.add(right, width=360)

        tk.Label(
            right,
            text="Activity log",
            font=("Segoe UI", 11, "bold"),
            bg="white"
        ).pack(anchor="w", pady=(0, 8))

        self.preview = ScrolledText(
            right,
            wrap="word",
            font=("Consolas", 9),
            height=10
        )
        self.preview.pack(fill="both", expand=True)

        self.preview.insert("end", "1. Choose Excel file\n2. Scan\n3. Select students\n4. Open Outlook\n")

    # ---------- LOGIC ----------
    def choose_file(self):
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls")],
        )
        if path:
            self.excel_path = path
            self.file_label.config(text=path)
            self.preview.delete("1.0", "end")
            self.preview.insert("end", "File selected.\n")

    def scan(self):
        if not self.excel_path:
            messagebox.showwarning("Missing file", "Select an Excel file first.")
            return

        self.students = find_students_with_absences(
            self.excel_path,
            min_absences=2
        )

        for item in self.tree.get_children():
            self.tree.delete(item)

        self.preview.delete("1.0", "end")

        if not self.students:
            self.preview.insert("end", "No absent students found.\n")
            return

        self.preview.insert("end", f"Total absent students: {len(self.students)}\n\n")

        for idx, s in enumerate(self.students):
            self.tree.insert(
                "",
                "end",
                iid=str(idx),
                values=(
                    s["Sport"],
                    s["Nom"],
                    s["Prénom"],
                    s["Email"],
                    s["Absent_Count"],
                    ", ".join(s["Absent_Dates"]),
                )
            )

    def build_subject_body(self, student):
        subject = SUBJECT
        body = BODY_TEMPLATE.format(
            prenom=student["Prénom"],
            nb=student["Absent_Count"],
            dates=format_dates(student["Absent_Dates"]),
        )
        return subject, body

    def open_compose_selected(self):
        sel = list(self.tree.selection())
        if not sel:
            messagebox.showinfo("No selection", "Select at least one student.")
            return

        if not messagebox.askyesno("Confirm", f"Open {len(sel)} email(s) in Outlook?"):
            return

        for iid in sel:
            s = self.students[int(iid)]
            subject, body = self.build_subject_body(s)
            query = urllib.parse.urlencode(
                {"subject": subject, "body": body},
                quote_via=urllib.parse.quote
            )
            webbrowser.open(f"mailto:{s['Email']}?{query}")

    def select_all(self):
        for item in self.tree.get_children():
            self.tree.selection_add(item)

    def clear_selection(self):
        self.tree.selection_remove(self.tree.selection())


if __name__ == "__main__":
    AttendanceApp().mainloop()
