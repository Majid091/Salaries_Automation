import tkinter as tk
from tkinter import ttk, messagebox
import os
from datetime import datetime
from google_sheets_reader import GoogleSheetsReader
from pdf_generator import PDFGenerator
from email_sender import EmailSender
import threading

# App Colors
TEAL_LIGHT = '#0e8282'
TEAL_DARK = '#073630'
WHITE = '#FFFFFF'
LIGHT_BG = '#f8f9fa'
DARK_TEXT = '#333333'
GRAY_TEXT = '#666666'


class RoundedButton(tk.Canvas):
    """Custom rounded button widget"""
    def __init__(self, parent, text, command, bg_color, fg_color, hover_color=None,
                 width=200, height=45, corner_radius=22, font=("Segoe UI", 11, "bold")):
        super().__init__(parent, width=width, height=height, bg=parent["bg"],
                        highlightthickness=0)

        self.command = command
        self.bg_color = bg_color
        self.fg_color = fg_color
        self.hover_color = hover_color or bg_color
        self.corner_radius = corner_radius
        self.text = text
        self.font = font
        self._width = width
        self._height = height

        self.draw_button(self.bg_color)

        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)
        self.bind("<Button-1>", self.on_click)

    def draw_button(self, color):
        self.delete("all")
        r = self.corner_radius
        w, h = self._width, self._height

        # Draw rounded rectangle
        self.create_arc(0, 0, 2*r, 2*r, start=90, extent=90, fill=color, outline=color)
        self.create_arc(w-2*r, 0, w, 2*r, start=0, extent=90, fill=color, outline=color)
        self.create_arc(0, h-2*r, 2*r, h, start=180, extent=90, fill=color, outline=color)
        self.create_arc(w-2*r, h-2*r, w, h, start=270, extent=90, fill=color, outline=color)

        self.create_rectangle(r, 0, w-r, h, fill=color, outline=color)
        self.create_rectangle(0, r, w, h-r, fill=color, outline=color)

        # Draw text
        self.create_text(w//2, h//2, text=self.text, fill=self.fg_color, font=self.font)

    def on_enter(self, e):
        self.draw_button(self.hover_color)
        self.config(cursor="hand2")

    def on_leave(self, e):
        self.draw_button(self.bg_color)

    def on_click(self, e):
        if self.command:
            self.command()


class GradientFrame(tk.Canvas):
    """Canvas with gradient background"""
    def __init__(self, parent, color1, color2, **kwargs):
        super().__init__(parent, **kwargs)
        self.color1 = color1
        self.color2 = color2
        self.bind("<Configure>", self.draw_gradient)

    def draw_gradient(self, event=None):
        self.delete("gradient")
        width = self.winfo_width()
        height = self.winfo_height()

        r1, g1, b1 = self.hex_to_rgb(self.color1)
        r2, g2, b2 = self.hex_to_rgb(self.color2)

        for i in range(height):
            ratio = i / height
            r = int(r1 + (r2 - r1) * ratio)
            g = int(g1 + (g2 - g1) * ratio)
            b = int(b1 + (b2 - b1) * ratio)
            color = f"#{r:02x}{g:02x}{b:02x}"
            self.create_line(0, i, width, i, fill=color, tags="gradient")

        self.tag_lower("gradient")

    def hex_to_rgb(self, hex_color):
        hex_color = hex_color.lstrip('#')
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))


class SalaryAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Salary Automation")
        self.root.geometry("550x620")
        self.root.minsize(500, 550)
        self.root.resizable(True, True)

        # Initialize components
        self.sheets_reader = None
        self.pdf_generator = PDFGenerator()
        self.email_sender = EmailSender()

        # Variables
        self.selected_month = tk.StringVar()
        self.selected_year = tk.StringVar()
        self.status_text = tk.StringVar(value="Ready")
        self.progress_var = tk.DoubleVar()
        self.employee_records = []
        self.selected_employees = []

        # Card references
        self.card = None
        self.shadow = None

        self.setup_ui()
        self.root.bind("<Configure>", self.on_resize)

    def setup_ui(self):
        # Gradient background
        self.gradient_bg = GradientFrame(
            self.root, TEAL_LIGHT, TEAL_DARK,
            highlightthickness=0
        )
        self.gradient_bg.pack(fill=tk.BOTH, expand=True)

        # Create card (will be positioned in on_resize)
        self.shadow = tk.Frame(self.gradient_bg, bg='#05504a')
        self.card = tk.Frame(self.gradient_bg, bg=WHITE)

        # Card content with padding
        self.content = tk.Frame(self.card, bg=WHITE)
        self.content.pack(fill=tk.BOTH, expand=True, padx=30, pady=25)

        # App Title
        tk.Label(
            self.content,
            text="Salary Automation",
            font=("Segoe UI", 20, "bold"),
            bg=WHITE,
            fg=TEAL_DARK
        ).pack(anchor=tk.W)

        tk.Label(
            self.content,
            text="Generate and send salary slips easily",
            font=("Segoe UI", 9),
            bg=WHITE,
            fg=GRAY_TEXT
        ).pack(anchor=tk.W, pady=(2, 15))

        # Divider
        tk.Frame(self.content, bg=TEAL_LIGHT, height=2).pack(fill=tk.X, pady=(0, 15))

        # Period Selection
        tk.Label(
            self.content,
            text="SELECT PERIOD",
            font=("Segoe UI", 8, "bold"),
            bg=WHITE,
            fg=GRAY_TEXT
        ).pack(anchor=tk.W)

        period_frame = tk.Frame(self.content, bg=WHITE)
        period_frame.pack(fill=tk.X, pady=(8, 15))

        # Month
        month_frame = tk.Frame(period_frame, bg=WHITE)
        month_frame.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 8))

        tk.Label(month_frame, text="Month", font=("Segoe UI", 9), bg=WHITE, fg=DARK_TEXT).pack(anchor=tk.W)

        months = ["January", "February", "March", "April", "May", "June",
                  "July", "August", "September", "October", "November", "December"]

        month_combo = ttk.Combobox(month_frame, textvariable=self.selected_month, values=months,
                                   state="readonly", font=("Segoe UI", 10))
        month_combo.pack(fill=tk.X, pady=(3, 0), ipady=4)
        month_combo.current(datetime.now().month - 1)

        # Year
        year_frame = tk.Frame(period_frame, bg=WHITE)
        year_frame.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(8, 0))

        tk.Label(year_frame, text="Year", font=("Segoe UI", 9), bg=WHITE, fg=DARK_TEXT).pack(anchor=tk.W)

        current_year = datetime.now().year
        years = [str(y) for y in range(current_year - 5, current_year + 2)]
        year_combo = ttk.Combobox(year_frame, textvariable=self.selected_year, values=years,
                                  state="readonly", font=("Segoe UI", 10))
        year_combo.pack(fill=tk.X, pady=(3, 0), ipady=4)
        year_combo.set(str(current_year))

        # Generate Button Frame
        self.generate_btn_frame = tk.Frame(self.content, bg=WHITE)
        self.generate_btn_frame.pack(fill=tk.X, pady=(8, 15))

        self.generate_btn = RoundedButton(
            self.generate_btn_frame,
            text="Generate PDFs",
            command=self.generate_pdfs,
            bg_color=TEAL_LIGHT,
            fg_color=WHITE,
            hover_color=TEAL_DARK,
            width=380,
            height=42,
            corner_radius=21
        )
        self.generate_btn.pack()

        # Email Section
        tk.Label(
            self.content,
            text="SEND EMAILS",
            font=("Segoe UI", 8, "bold"),
            bg=WHITE,
            fg=GRAY_TEXT
        ).pack(anchor=tk.W, pady=(5, 8))

        # Email buttons frame
        self.email_btn_frame = tk.Frame(self.content, bg=WHITE)
        self.email_btn_frame.pack(fill=tk.X, pady=(0, 15))

        self.send_all_btn = RoundedButton(
            self.email_btn_frame,
            text="Send to All",
            command=lambda: self.send_emails("all"),
            bg_color=TEAL_LIGHT,
            fg_color=WHITE,
            hover_color=TEAL_DARK,
            width=182,
            height=40,
            corner_radius=20
        )
        self.send_all_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.send_selected_btn = RoundedButton(
            self.email_btn_frame,
            text="Send to Selected",
            command=lambda: self.send_emails("selective"),
            bg_color=TEAL_LIGHT,
            fg_color=WHITE,
            hover_color=TEAL_DARK,
            width=182,
            height=40,
            corner_radius=20
        )
        self.send_selected_btn.pack(side=tk.LEFT)

        # Progress Section
        self.progress_frame = tk.Frame(self.content, bg=LIGHT_BG, padx=12, pady=12)
        self.progress_frame.pack(fill=tk.X, pady=(5, 12))

        self.status_label = tk.Label(
            self.progress_frame,
            textvariable=self.status_text,
            font=("Segoe UI", 9),
            bg=LIGHT_BG,
            fg=DARK_TEXT
        )
        self.status_label.pack(anchor=tk.W)

        self.progress_bar_bg = tk.Frame(self.progress_frame, bg='#e0e0e0', height=6)
        self.progress_bar_bg.pack(fill=tk.X, pady=(6, 0))
        self.progress_bar_bg.pack_propagate(False)

        self.progress_fill = tk.Frame(self.progress_bar_bg, bg=TEAL_LIGHT, height=6, width=0)
        self.progress_fill.place(x=0, y=0, relheight=1)

        # Quick Tips
        tips_frame = tk.Frame(self.content, bg=WHITE)
        tips_frame.pack(fill=tk.X, pady=(8, 0))

        tk.Label(
            tips_frame,
            text="Quick Tips",
            font=("Segoe UI", 9, "bold"),
            bg=WHITE,
            fg=DARK_TEXT
        ).pack(anchor=tk.W)

        tips = [
            "• Select month and year, then generate PDFs",
            "• PDFs are saved in the 'pdfs' folder",
            "• Use 'Send to Selected' to choose specific employees"
        ]
        for tip in tips:
            tk.Label(
                tips_frame,
                text=tip,
                font=("Segoe UI", 8),
                bg=WHITE,
                fg=GRAY_TEXT
            ).pack(anchor=tk.W, pady=1)

    def on_resize(self, event=None):
        """Handle window resize - make card responsive"""
        if event and event.widget != self.root:
            return

        self.root.update_idletasks()
        win_width = self.root.winfo_width()
        win_height = self.root.winfo_height()

        # Calculate card size (responsive)
        card_width = min(440, win_width - 60)
        card_height = min(530, win_height - 60)

        card_x = (win_width - card_width) // 2
        card_y = (win_height - card_height) // 2

        # Position shadow and card
        self.shadow.place(x=card_x+3, y=card_y+3, width=card_width, height=card_height)
        self.card.place(x=card_x, y=card_y, width=card_width, height=card_height)

    def update_progress_bar(self, value):
        self.root.update_idletasks()
        progress_width = self.progress_bar_bg.winfo_width()
        fill_width = int((value / 100) * progress_width)
        self.progress_fill.configure(width=fill_width)

    def _get_sheet_name(self):
        return f"{self.selected_month.get()} {self.selected_year.get()}"

    def _get_email_from_record(self, record):
        email_keys = ['email address', 'Email Address', 'email', 'Email', 'EMAIL', 'E-mail', 'e-mail']
        for key in email_keys:
            if key in record and record[key]:
                return str(record[key]).strip()
        record_lower = {k.lower(): v for k, v in record.items()}
        if 'email address' in record_lower and record_lower['email address']:
            return str(record_lower['email address']).strip()
        if 'email' in record_lower and record_lower['email']:
            return str(record_lower['email']).strip()
        return ''

    def _get_name_from_record(self, record):
        name_keys = ['Name', 'name', 'NAME', 'Employee Name', 'employee name']
        for key in name_keys:
            if key in record and record[key]:
                return str(record[key]).strip()
        return 'Unknown'

    def generate_pdfs(self):
        if not self.selected_month.get() or not self.selected_year.get():
            messagebox.showwarning("Warning", "Please select both month and year")
            return
        thread = threading.Thread(target=self._generate_pdfs_thread)
        thread.daemon = True
        thread.start()

    def _update_ui(self, status=None, progress=None, message_type=None, message=None):
        if status is not None:
            self.status_text.set(status)
        if progress is not None:
            self.progress_var.set(progress)
            self.update_progress_bar(progress)
        if message_type and message:
            if message_type == "info":
                messagebox.showinfo("Info", message)
            elif message_type == "success":
                messagebox.showinfo("Success", message)
            elif message_type == "error":
                messagebox.showerror("Error", message)
            elif message_type == "warning":
                messagebox.showwarning("Warning", message)

    def _generate_pdfs_thread(self):
        try:
            self.root.after(0, lambda: self._update_ui(status="Connecting to Google Sheets...", progress=0))

            self.sheets_reader = GoogleSheetsReader()
            sheet_name = self._get_sheet_name()

            self.root.after(0, lambda: self._update_ui(status=f"Fetching data for {sheet_name}..."))

            records = self.sheets_reader.get_month_data(sheet_name)
            self.employee_records = records

            company_info = self.sheets_reader.get_company_info()
            self.pdf_generator.set_company_info(
                company_info.get('company_name', ''),
                company_info.get('app_name', '')
            )

            if not records:
                self.root.after(0, lambda: self._update_ui(
                    status="Ready", progress=0,
                    message_type="info", message=f"No records found for {sheet_name}"
                ))
                return

            os.makedirs("pdfs", exist_ok=True)
            total = len(records)

            for idx, record in enumerate(records, 1):
                self.root.after(0, lambda i=idx, t=total: self._update_ui(status=f"Generating PDF {i}/{t}..."))
                self.pdf_generator.create_pdf(record, sheet_name)
                progress = (idx / total) * 100
                self.root.after(0, lambda p=progress: self._update_ui(progress=p))

            self.root.after(0, lambda: self._update_ui(
                status="Ready", progress=0,
                message_type="success", message=f"Generated {total} PDFs successfully!"
            ))

        except Exception as e:
            error_msg = str(e)
            self.root.after(0, lambda msg=error_msg: self._update_ui(
                status="Ready", progress=0,
                message_type="error", message=f"Error: {msg}"
            ))

    def send_emails(self, mode):
        if not self.selected_month.get() or not self.selected_year.get():
            messagebox.showwarning("Warning", "Please select both month and year")
            return

        if mode == "selective":
            self.show_employee_selector()
        else:
            if not self.employee_records:
                try:
                    self.status_text.set("Loading employees...")
                    self.root.update()
                    if not self.sheets_reader:
                        self.sheets_reader = GoogleSheetsReader()
                    self.employee_records = self.sheets_reader.get_month_data(self._get_sheet_name())
                    self.status_text.set("Ready")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to load employees: {str(e)}")
                    return

            if not self.employee_records:
                messagebox.showinfo("Info", "No employees found")
                return

            count = len(self.employee_records)
            if messagebox.askyesno("Confirm", f"Send emails to all {count} employees?"):
                self.selected_employees = self.employee_records
                thread = threading.Thread(target=self._send_emails_thread)
                thread.daemon = True
                thread.start()

    def show_employee_selector(self):
        """Show improved employee selection dialog"""
        if not self.employee_records:
            try:
                self.status_text.set("Loading employees...")
                self.root.update()
                if not self.sheets_reader:
                    self.sheets_reader = GoogleSheetsReader()
                self.employee_records = self.sheets_reader.get_month_data(self._get_sheet_name())
                self.status_text.set("Ready")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load employees: {str(e)}")
                return

        if not self.employee_records:
            messagebox.showinfo("Info", "No employees found")
            return

        # Create dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Select Employees")
        dialog.geometry("500x580")
        dialog.minsize(450, 500)
        dialog.resizable(True, True)
        dialog.transient(self.root)
        dialog.grab_set()

        # Center dialog
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 250
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 290
        dialog.geometry(f"500x580+{x}+{y}")

        # Gradient background
        gradient_bg = GradientFrame(dialog, TEAL_LIGHT, TEAL_DARK, highlightthickness=0)
        gradient_bg.pack(fill=tk.BOTH, expand=True)

        # Card container frame
        card_container = tk.Frame(gradient_bg, bg=gradient_bg["bg"])
        card_container.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

        # Shadow and card
        shadow = tk.Frame(card_container, bg='#05504a')
        shadow.place(x=3, y=3, relwidth=1, relheight=1)

        card = tk.Frame(card_container, bg=WHITE)
        card.pack(padx=0, pady=0)

        # Content
        content = tk.Frame(card, bg=WHITE, padx=25, pady=20)
        content.pack(fill=tk.BOTH, expand=True)

        # Variables
        checkbox_vars = []
        employee_frames = []

        # Header
        header_frame = tk.Frame(content, bg=WHITE)
        header_frame.pack(fill=tk.X, pady=(0, 12))

        tk.Label(
            header_frame,
            text="Select Employees",
            font=("Segoe UI", 16, "bold"),
            bg=WHITE,
            fg=TEAL_DARK
        ).pack(side=tk.LEFT)

        selected_count_var = tk.StringVar(value="0 selected")
        tk.Label(
            header_frame,
            textvariable=selected_count_var,
            font=("Segoe UI", 9),
            bg=WHITE,
            fg=TEAL_LIGHT
        ).pack(side=tk.RIGHT, pady=(5, 0))

        # Divider
        tk.Frame(content, bg=TEAL_LIGHT, height=2).pack(fill=tk.X, pady=(0, 12))

        # Search box
        search_frame = tk.Frame(content, bg=LIGHT_BG, padx=10, pady=8)
        search_frame.pack(fill=tk.X, pady=(0, 8))

        search_var = tk.StringVar()
        search_entry = tk.Entry(
            search_frame,
            textvariable=search_var,
            font=("Segoe UI", 10),
            bg=LIGHT_BG,
            fg=DARK_TEXT,
            relief=tk.FLAT,
            highlightthickness=0
        )
        search_entry.pack(fill=tk.X, ipady=3)
        search_entry.insert(0, "Search employees...")
        search_entry.config(fg=GRAY_TEXT)

        def on_search_focus_in(e):
            if search_entry.get() == "Search employees...":
                search_entry.delete(0, tk.END)
                search_entry.config(fg=DARK_TEXT)

        def on_search_focus_out(e):
            if not search_entry.get():
                search_entry.insert(0, "Search employees...")
                search_entry.config(fg=GRAY_TEXT)

        search_entry.bind("<FocusIn>", on_search_focus_in)
        search_entry.bind("<FocusOut>", on_search_focus_out)

        # Action buttons
        action_frame = tk.Frame(content, bg=WHITE)
        action_frame.pack(fill=tk.X, pady=(0, 8))

        def update_count():
            count = sum(1 for var in checkbox_vars if var.get())
            selected_count_var.set(f"{count} selected")

        def select_all():
            for var in checkbox_vars:
                var.set(True)
            update_count()

        def deselect_all():
            for var in checkbox_vars:
                var.set(False)
            update_count()

        select_all_btn = RoundedButton(
            action_frame,
            text="Select All",
            command=select_all,
            bg_color=TEAL_LIGHT,
            fg_color=WHITE,
            hover_color=TEAL_DARK,
            width=90,
            height=30,
            corner_radius=15,
            font=("Segoe UI", 9, "bold")
        )
        select_all_btn.pack(side=tk.LEFT, padx=(0, 6))

        deselect_btn = RoundedButton(
            action_frame,
            text="Deselect All",
            command=deselect_all,
            bg_color='#e0e0e0',
            fg_color=DARK_TEXT,
            hover_color='#d0d0d0',
            width=90,
            height=30,
            corner_radius=15,
            font=("Segoe UI", 9, "bold")
        )
        deselect_btn.pack(side=tk.LEFT)

        # Employee list container
        list_container = tk.Frame(content, bg=WHITE, height=280)
        list_container.pack(fill=tk.BOTH, expand=True)
        list_container.pack_propagate(False)

        canvas = tk.Canvas(list_container, bg=WHITE, highlightthickness=0)
        scrollbar = tk.Scrollbar(list_container, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=WHITE)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas_frame = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        def on_canvas_configure(e):
            canvas.itemconfig(canvas_frame, width=e.width)

        canvas.bind("<Configure>", on_canvas_configure)
        canvas.configure(yscrollcommand=scrollbar.set)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", on_mousewheel)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Create employee items
        for idx, record in enumerate(self.employee_records):
            var = tk.BooleanVar(value=False)
            checkbox_vars.append(var)

            name = self._get_name_from_record(record)
            email = self._get_email_from_record(record)

            emp_frame = tk.Frame(scrollable_frame, bg=WHITE, pady=2)
            emp_frame.pack(fill=tk.X)
            employee_frames.append((emp_frame, name, email))

            inner_frame = tk.Frame(emp_frame, bg=LIGHT_BG, padx=10, pady=8)
            inner_frame.pack(fill=tk.X)

            cb = tk.Checkbutton(
                inner_frame,
                variable=var,
                bg=LIGHT_BG,
                activebackground=LIGHT_BG,
                command=update_count
            )
            cb.pack(side=tk.LEFT)

            info_frame = tk.Frame(inner_frame, bg=LIGHT_BG)
            info_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))

            tk.Label(
                info_frame,
                text=name,
                font=("Segoe UI", 10, "bold"),
                bg=LIGHT_BG,
                fg=DARK_TEXT
            ).pack(anchor=tk.W)

            if email:
                tk.Label(
                    info_frame,
                    text=email,
                    font=("Segoe UI", 8),
                    bg=LIGHT_BG,
                    fg=GRAY_TEXT
                ).pack(anchor=tk.W)

            def make_toggle(v):
                def toggle(e):
                    v.set(not v.get())
                    update_count()
                return toggle

            inner_frame.bind("<Button-1>", make_toggle(var))
            for child in inner_frame.winfo_children():
                if not isinstance(child, tk.Checkbutton):
                    child.bind("<Button-1>", make_toggle(var))

        # Search filter
        def filter_employees(*args):
            search_text = search_var.get().lower()
            if search_text == "search employees...":
                search_text = ""
            for emp_frame, name, email in employee_frames:
                if search_text in name.lower() or search_text in email.lower():
                    emp_frame.pack(fill=tk.X)
                else:
                    emp_frame.pack_forget()

        search_var.trace("w", filter_employees)

        # Bottom buttons
        btn_frame = tk.Frame(content, bg=WHITE)
        btn_frame.pack(fill=tk.X, pady=(12, 0))

        def on_send():
            selected_indices = [i for i, var in enumerate(checkbox_vars) if var.get()]
            if not selected_indices:
                messagebox.showwarning("Warning", "Please select at least one employee", parent=dialog)
                return
            self.selected_employees = [self.employee_records[i] for i in selected_indices]
            count = len(self.selected_employees)
            canvas.unbind_all("<MouseWheel>")
            if messagebox.askyesno("Confirm", f"Send emails to {count} selected employee(s)?", parent=dialog):
                dialog.destroy()
                thread = threading.Thread(target=self._send_emails_thread)
                thread.daemon = True
                thread.start()

        def on_cancel():
            canvas.unbind_all("<MouseWheel>")
            dialog.destroy()

        cancel_btn = RoundedButton(
            btn_frame,
            text="Cancel",
            command=on_cancel,
            bg_color='#e0e0e0',
            fg_color=DARK_TEXT,
            hover_color='#d0d0d0',
            width=100,
            height=38,
            corner_radius=19,
            font=("Segoe UI", 10, "bold")
        )
        cancel_btn.pack(side=tk.RIGHT, padx=(8, 0))

        send_btn = RoundedButton(
            btn_frame,
            text="Send Emails",
            command=on_send,
            bg_color=TEAL_LIGHT,
            fg_color=WHITE,
            hover_color=TEAL_DARK,
            width=120,
            height=38,
            corner_radius=19,
            font=("Segoe UI", 10, "bold")
        )
        send_btn.pack(side=tk.RIGHT)

        dialog.protocol("WM_DELETE_WINDOW", on_cancel)

        # Responsive card sizing for dialog
        def on_dialog_resize(event=None):
            if event and event.widget != dialog:
                return
            dialog.update_idletasks()
            w = dialog.winfo_width()
            h = dialog.winfo_height()
            card_w = min(420, w - 50)
            card_h = min(520, h - 40)
            card_container.configure(width=card_w, height=card_h)
            shadow.place(x=3, y=3, width=card_w, height=card_h)
            card.configure(width=card_w, height=card_h)

        dialog.bind("<Configure>", on_dialog_resize)
        on_dialog_resize()

    def _send_emails_thread(self):
        try:
            self.root.after(0, lambda: self._update_ui(status="Sending emails...", progress=0))

            sheet_name = self._get_sheet_name()
            records = self.selected_employees

            pdfs_dir = "pdfs"
            if not os.path.exists(pdfs_dir):
                self.root.after(0, lambda: self._update_ui(
                    status="Ready", progress=0,
                    message_type="warning", message="PDFs not found. Please generate PDFs first."
                ))
                return

            total = len(records)
            success = 0
            fail = 0

            for idx, record in enumerate(records, 1):
                email = self._get_email_from_record(record)
                name = self._get_name_from_record(record)

                if not email or '@' not in email:
                    fail += 1
                    continue

                self.root.after(0, lambda i=idx, t=total, n=name: self._update_ui(
                    status=f"Sending {i}/{t}: {n}..."
                ))

                pdf_filename = self.pdf_generator.get_pdf_filename(record, sheet_name)
                pdf_path = os.path.join(pdfs_dir, pdf_filename)

                if not os.path.exists(pdf_path):
                    fail += 1
                    continue

                try:
                    self.email_sender.send_email(
                        to_email=email,
                        subject=f"Salary Statement - {sheet_name}",
                        body=self._get_email_body(record, sheet_name),
                        pdf_path=pdf_path
                    )
                    success += 1
                except Exception as e:
                    print(f"Error sending to {email}: {str(e)}")
                    fail += 1

                progress = (idx / total) * 100
                self.root.after(0, lambda p=progress: self._update_ui(progress=p))

            self.root.after(0, lambda s=success, f=fail: self._update_ui(
                status="Ready", progress=0,
                message_type="info", message=f"Emails sent!\nSuccess: {s}\nFailed: {f}"
            ))

        except Exception as e:
            error_msg = str(e)
            self.root.after(0, lambda msg=error_msg: self._update_ui(
                status="Ready", progress=0,
                message_type="error", message=f"Error: {msg}"
            ))

    def _get_email_body(self, record, month_name):
        name = self._get_name_from_record(record)
        return f"""Dear {name},

Please find attached your salary statement for {month_name}.

If you have any questions, please contact HR.

Best regards,
HR Department"""


if __name__ == "__main__":
    root = tk.Tk()
    app = SalaryAutomationApp(root)
    root.mainloop()
