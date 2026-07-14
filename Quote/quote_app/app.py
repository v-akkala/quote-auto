from __future__ import annotations

import calendar
import json
import os
import subprocess
import sys
from datetime import date
from pathlib import Path
from tkinter import END, LEFT, BOTH, RIGHT, VERTICAL, BooleanVar, Listbox, Menu, StringVar, Text, Tk, Toplevel, filedialog, messagebox
from tkinter import ttk

from generator import (
    ProductItem,
    TestItem,
    generate_quote,
    load_quote,
    product_key,
    product_quote_label,
    today_string,
)

QUOTE_MADE_BY_OPTIONS = ["Vijayalakshmi", "Jyothi", "Nethra", "Nithin", "Roshan"]
TECHNICAL_ENGINEER_OPTIONS = ["Dinesh", "Akshay", "Shailesh", "Ajesh", "Dada", "Somashekar", "Madan", "Jagan"]
QUOTE_PREFIX = "SKC2627"
ROOT_DIR = Path(__file__).resolve().parent.parent
TEST_STANDARDS_PATH = ROOT_DIR / "data" / "test_standards.json"


def load_test_catalog() -> tuple[list[str], dict[str, list[str]], dict[str, str]]:
    fallback = {
        "High Temperature": [],
        "Thermal Shock": [],
        "Power Burn-In": [],
        "Thermal Cycling": [],
        "Low Temperature": [],
        "Acceleration": [],
        "Visual Examination": [],
        "Vibration": [],
        "Altitude": [],
        "Solar Radiation": [],
        "Tropical Exposure": [],
        "Mould Growth": [],
        "Corrosion": [],
        "Toppling": [],
        "Bump": [],
        "Shock": [],
        "Damp Heat": [],
        "Dust": [],
    }
    if not TEST_STANDARDS_PATH.exists():
        return list(fallback), fallback, {}

    try:
        raw_items = json.loads(TEST_STANDARDS_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return list(fallback), fallback, {}

    type_to_standards: dict[str, list[str]] = {}
    standard_to_type: dict[str, str] = {}
    for item in raw_items:
        test_type = str(item.get("test_type", "")).strip()
        if not test_type:
            continue
        standards = []
        for standard in item.get("standards", []):
            standard_text = str(standard).strip()
            if standard_text and standard_text not in standards:
                standards.append(standard_text)
                standard_to_type.setdefault(standard_text, test_type)
        type_to_standards[test_type] = standards

    if not type_to_standards:
        return list(fallback), fallback, {}
    return sorted(type_to_standards), type_to_standards, standard_to_type


TEST_NAMES, TEST_STANDARDS_BY_TYPE, TEST_TYPE_BY_STANDARD = load_test_catalog()
ALL_TEST_STANDARDS = sorted(TEST_TYPE_BY_STANDARD)


def prefix_matches(values: list[str], prefix: str) -> list[str]:
    clean_prefix = prefix.strip().lower()
    if not clean_prefix:
        return list(values)
    return [value for value in values if value.lower().startswith(clean_prefix)]


class AutocompleteCombobox(ttk.Frame):
    def __init__(self, parent, textvariable: StringVar, values: list[str], **kwargs):
        super().__init__(parent)
        width = kwargs.pop("width", None)
        self.variable = textvariable
        self.all_values = list(values)
        self.matches = list(values)
        self.match_index = -1
        self.cycle_prefix = ""
        self.popup: Toplevel | None = None
        self.listbox: Listbox | None = None

        self.entry = ttk.Entry(self, textvariable=textvariable, width=width)
        self.entry.pack(side=LEFT, fill="x", expand=True)
        ttk.Button(self, text="v", width=2, command=self.show_all_matches).pack(side=LEFT, padx=(4, 0))

        self.entry.bind("<KeyRelease>", self.on_key_release)
        self.entry.bind("<Tab>", self.cycle_match)
        self.entry.bind("<Return>", self.accept_match)
        self.entry.bind("<Escape>", self.hide_popup)
        self.entry.bind("<Down>", self.cycle_match)
        self.entry.bind("<Up>", self.cycle_match_reverse)
        self.entry.bind("<FocusIn>", self.on_focus_in)
        self.entry.bind("<FocusOut>", self.on_focus_out)

    def set_values(self, values: list[str]) -> None:
        self.all_values = list(values)
        self.cycle_prefix = self.get().strip()
        self.update_matches(reset_index=True)
        self.refresh_popup()

    def update_matches(self, prefix: str | None = None, reset_index: bool = False) -> None:
        typed = self.get().strip() if prefix is None else prefix
        matches = prefix_matches(self.all_values, typed)
        self.matches = matches
        if reset_index:
            self.match_index = 0 if matches else -1
        elif self.match_index >= len(matches):
            self.match_index = len(matches) - 1

    def on_key_release(self, event=None):
        if event and event.keysym in {"Tab", "Return", "Escape", "Up", "Down", "Left", "Right"}:
            return
        self.cycle_prefix = self.get().strip()
        self.update_matches(self.cycle_prefix, reset_index=True)
        self.refresh_popup()

    def on_focus_in(self, _event=None):
        self.cycle_prefix = self.get().strip()
        self.update_matches(self.cycle_prefix, reset_index=True)
        self.refresh_popup(force=True)

    def on_focus_out(self, _event=None):
        self.after(150, self.hide_popup)

    def show_all_matches(self) -> None:
        self.entry.focus_set()
        self.cycle_prefix = self.get().strip()
        self.update_matches(self.cycle_prefix, reset_index=True)
        self.refresh_popup(force=True)

    def cycle_match(self, _event=None):
        if not self.cycle_prefix:
            self.cycle_prefix = self.get().strip()
        self.update_matches(self.cycle_prefix)
        if not self.matches:
            return "break"
        self.match_index = (self.match_index + 1) % len(self.matches)
        self.refresh_popup(force=True)
        self.entry.focus_set()
        self.entry.icursor(END)
        return "break"

    def cycle_match_reverse(self, _event=None):
        if not self.cycle_prefix:
            self.cycle_prefix = self.get().strip()
        self.update_matches(self.cycle_prefix)
        if not self.matches:
            return "break"
        self.match_index = (self.match_index - 1) % len(self.matches)
        self.refresh_popup(force=True)
        self.entry.focus_set()
        self.entry.icursor(END)
        return "break"

    def accept_match(self, _event=None):
        if self.matches and self.match_index >= 0:
            self.set(self.matches[self.match_index])
        self.entry.icursor(END)
        self.cycle_prefix = self.get().strip()
        self.hide_popup()
        self.event_generate("<<AutocompleteAccepted>>")
        self.event_generate("<<ComboboxSelected>>")
        return "break"

    def refresh_popup(self, force: bool = False) -> None:
        if not self.matches:
            self.hide_popup()
            return
        if not force and not self.get().strip():
            self.hide_popup()
            return
        if self.popup is None or not self.popup.winfo_exists():
            self.popup = Toplevel(self)
            self.popup.wm_overrideredirect(True)
            self.listbox = Listbox(self.popup, height=min(8, len(self.matches)), exportselection=False)
            self.listbox.pack(fill=BOTH, expand=True)
            self.listbox.bind("<ButtonRelease-1>", self.accept_clicked_match)
        assert self.listbox is not None
        self.listbox.delete(0, END)
        for value in self.matches[:100]:
            self.listbox.insert(END, value)
        self.listbox.configure(height=min(8, len(self.matches)))
        self.highlight_match()
        self.position_popup()
        self.popup.deiconify()

    def highlight_match(self) -> None:
        if self.listbox is None or self.match_index < 0:
            return
        visible_index = min(self.match_index, self.listbox.size() - 1)
        self.listbox.selection_clear(0, END)
        self.listbox.selection_set(visible_index)
        self.listbox.activate(visible_index)
        self.listbox.see(visible_index)

    def position_popup(self) -> None:
        if self.popup is None:
            return
        self.update_idletasks()
        x = self.entry.winfo_rootx()
        y = self.entry.winfo_rooty() + self.entry.winfo_height()
        width = max(self.entry.winfo_width(), 360)
        self.popup.geometry(f"{width}x{min(8, len(self.matches)) * 22}+{x}+{y}")

    def accept_clicked_match(self, _event=None):
        if self.listbox is not None:
            selection = self.listbox.curselection()
            if selection:
                self.match_index = int(selection[0])
        return self.accept_match()

    def hide_popup(self, _event=None):
        if self.popup is not None and self.popup.winfo_exists():
            self.popup.withdraw()
        return None

    def get(self) -> str:
        return self.variable.get()

    def set(self, value: str) -> None:
        self.variable.set(value)

    def icursor(self, index) -> None:
        self.entry.icursor(index)

    def focus_set(self) -> None:
        self.entry.focus_set()

    def bind(self, sequence=None, func=None, add=None):
        if sequence in {"<FocusOut>", "<FocusIn>"}:
            return self.entry.bind(sequence, func, add)
        return super().bind(sequence, func, add)


class MultiSelectMenu(ttk.Frame):
    def __init__(self, parent, options: list[str], variable: StringVar, width: int = 34):
        super().__init__(parent)
        self.options = options
        self.variable = variable
        self.display_var = StringVar(value="Select names")
        self.option_vars = {option: BooleanVar(value=False) for option in options}
        self.button = ttk.Menubutton(self, textvariable=self.display_var, width=width)
        self.menu = Menu(self.button, tearoff=False)
        for option in options:
            self.menu.add_checkbutton(
                label=option,
                variable=self.option_vars[option],
                command=self.sync_from_checks,
            )
        self.menu.add_separator()
        self.menu.add_command(label="Clear selection", command=self.clear)
        self.button["menu"] = self.menu
        self.button.pack(fill="x", expand=True)
        self.set_selection(variable.get())

    def sync_from_checks(self) -> None:
        selected = [option for option in self.options if self.option_vars[option].get()]
        value = ", ".join(selected)
        self.variable.set(value)
        self.display_var.set(value or "Select names")

    def set_selection(self, value: str) -> None:
        selected = {part.strip() for part in value.split(",") if part.strip()}
        for option, option_var in self.option_vars.items():
            option_var.set(option in selected)
        self.sync_from_checks()

    def clear(self) -> None:
        for option_var in self.option_vars.values():
            option_var.set(False)
        self.sync_from_checks()


class DateEntry(ttk.Frame):
    def __init__(self, parent, textvariable: StringVar, width: int = 18):
        super().__init__(parent)
        self.variable = textvariable
        self.entry = ttk.Entry(self, textvariable=textvariable, width=width)
        self.entry.pack(side=LEFT, fill="x", expand=True)
        ttk.Button(self, text="Select", command=self.open_picker).pack(side=LEFT, padx=(6, 0))

    def open_picker(self) -> None:
        DatePicker(self, self.variable)


class DatePicker(Toplevel):
    def __init__(self, parent, variable: StringVar):
        super().__init__(parent)
        self.title("Select date")
        self.resizable(False, False)
        self.variable = variable
        self.selected = date.today()
        self.year = self.selected.year
        self.month = self.selected.month
        self.header = ttk.Label(self, anchor="center", font=("Segoe UI", 10, "bold"))
        self.header.grid(row=0, column=1, columnspan=5, sticky="ew", pady=8)
        ttk.Button(self, text="<", width=3, command=self.previous_month).grid(row=0, column=0, padx=4)
        ttk.Button(self, text=">", width=3, command=self.next_month).grid(row=0, column=6, padx=4)
        self.days_frame = ttk.Frame(self)
        self.days_frame.grid(row=1, column=0, columnspan=7, padx=8, pady=(0, 8))
        self.render_days()
        self.transient(parent.winfo_toplevel())
        self.grab_set()

    def previous_month(self) -> None:
        self.month -= 1
        if self.month == 0:
            self.month = 12
            self.year -= 1
        self.render_days()

    def next_month(self) -> None:
        self.month += 1
        if self.month == 13:
            self.month = 1
            self.year += 1
        self.render_days()

    def render_days(self) -> None:
        for child in self.days_frame.winfo_children():
            child.destroy()
        self.header.configure(text=f"{calendar.month_name[self.month]} {self.year}")
        for col, day_name in enumerate(("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")):
            ttk.Label(self.days_frame, text=day_name, anchor="center", width=5).grid(row=0, column=col)
        for row_index, week in enumerate(calendar.monthcalendar(self.year, self.month), start=1):
            for col_index, day_num in enumerate(week):
                if day_num == 0:
                    ttk.Label(self.days_frame, text="", width=5).grid(row=row_index, column=col_index)
                    continue
                ttk.Button(
                    self.days_frame,
                    text=str(day_num),
                    width=5,
                    command=lambda day=day_num: self.choose(day),
                ).grid(row=row_index, column=col_index, padx=1, pady=1)

    def choose(self, day: int) -> None:
        self.variable.set(f"{self.month}-{day}-{self.year}")
        self.destroy()


class QuoteApp(Tk):
    def __init__(self):
        super().__init__()
        self.title("SKC Service Quote Generator")
        self.geometry("1120x760")
        self.minsize(980, 680)

        self.tests: list[TestItem] = []
        self.products: list[ProductItem] = []
        self.editing_index: int | None = None
        self.product_editing_index: int | None = None
        self.last_output: Path | None = None
        self.fields: dict[str, StringVar] = {}
        self.use_discount_var = BooleanVar(value=False)
        self.discount_percent_var = StringVar()
        self.instate_var = BooleanVar(value=True)
        self.assign_products_var = BooleanVar(value=False)

        self._configure_style()
        self._build()

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        if sys.platform == "win32":
            style.theme_use("vista")
        style.configure("Title.TLabel", font=("Segoe UI", 12, "bold"))
        style.configure("Section.TLabelframe.Label", font=("Segoe UI", 10, "bold"))

    def _build(self) -> None:
        container = ttk.Frame(self, padding=14)
        container.pack(fill=BOTH, expand=True)

        top = ttk.Frame(container)
        top.pack(fill="x", pady=(0, 10))
        ttk.Label(top, text="SKC Service Quote Generator", style="Title.TLabel").pack(side=LEFT)

        notebook = ttk.Notebook(container)
        notebook.pack(fill=BOTH, expand=True)
        self.info_tab = ttk.Frame(notebook, padding=12)
        self.tests_tab = ttk.Frame(notebook, padding=12)
        notebook.add(self.info_tab, text="Quote Info")
        notebook.add(self.tests_tab, text="Tests")

        self._build_info_tab()
        self._build_tests_tab()
        self._build_footer(container)

    def _build_info_tab(self) -> None:
        self._set_field("quote_date", today_string())
        self._set_field("email_date", today_string())
        self._set_field("quote_suffix", "")
        self._set_field("quote_made_by", "")
        self._set_field("technical_engineers", "")
        self._set_field("payment_terms", "As per agreement with the Customer")

        quote_frame = ttk.LabelFrame(self.info_tab, text="Quote", style="Section.TLabelframe", padding=12)
        quote_frame.pack(fill="x", pady=(0, 12))
        quote_ref_widget = ttk.Frame(quote_frame)
        ttk.Label(quote_ref_widget, text=QUOTE_PREFIX).pack(side=LEFT)
        ttk.Entry(quote_ref_widget, textvariable=self.fields["quote_suffix"], width=12).pack(side=LEFT, padx=(4, 0))
        self._add_labeled_widget(quote_frame, "Service Quote Ref No.", quote_ref_widget, 0, 0)
        self._add_labeled_widget(quote_frame, "Dated", DateEntry(quote_frame, self.fields["quote_date"]), 0, 2)
        self._add_labeled_widget(quote_frame, "Your email dated", DateEntry(quote_frame, self.fields["email_date"]), 1, 0)
        self.quote_made_by_menu = MultiSelectMenu(quote_frame, QUOTE_MADE_BY_OPTIONS, self.fields["quote_made_by"])
        self._add_labeled_widget(quote_frame, "Quote Made By", self.quote_made_by_menu, 1, 2)
        self.technical_engineers_menu = MultiSelectMenu(
            quote_frame,
            TECHNICAL_ENGINEER_OPTIONS,
            self.fields["technical_engineers"],
        )
        self._add_labeled_widget(quote_frame, "Technical Engineer(s)", self.technical_engineers_menu, 2, 0)

        payment_controls = ttk.Frame(quote_frame)
        self.discount_check = ttk.Checkbutton(
            payment_controls,
            text="Apply discount",
            variable=self.use_discount_var,
            command=self.toggle_discount,
        )
        self.discount_check.pack(side=LEFT)
        self.discount_frame = ttk.Frame(payment_controls)
        ttk.Label(self.discount_frame, text="Discount %").pack(side=LEFT, padx=(12, 4))
        ttk.Entry(self.discount_frame, textvariable=self.discount_percent_var, width=8).pack(side=LEFT)
        ttk.Checkbutton(payment_controls, text="In-state", variable=self.instate_var).pack(side=LEFT, padx=(18, 0))
        self._add_labeled_widget(quote_frame, "Payment", payment_controls, 2, 2)
        ttk.Entry(quote_frame, textvariable=self.fields["payment_terms"], width=48).grid(
            row=3,
            column=1,
            columnspan=3,
            sticky="ew",
            padx=(0, 18),
            pady=6,
        )
        ttk.Label(quote_frame, text="Payment Terms").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=6)
        self.toggle_discount()

        sample_frame = ttk.LabelFrame(
            self.info_tab,
            text="Test Sample and Test Details",
            style="Section.TLabelframe",
            padding=12,
        )
        sample_frame.pack(fill=BOTH, expand=True, pady=(0, 12))
        sample_frame.columnconfigure(0, weight=2)
        sample_frame.columnconfigure(1, weight=1)
        sample_frame.rowconfigure(1, weight=1)

        product_editor = ttk.Frame(sample_frame)
        product_editor.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 12))
        product_editor.columnconfigure(1, weight=1)

        self.product_name_var = StringVar()
        self.product_dim_vars = [StringVar(), StringVar(), StringVar()]
        self.product_dim_unit_var = StringVar(value="mm")
        self.product_weight_var = StringVar()
        self.product_weight_unit_var = StringVar(value="kg")

        ttk.Label(product_editor, text="Product name").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=6)
        ttk.Entry(product_editor, textvariable=self.product_name_var, width=38).grid(
            row=0,
            column=1,
            sticky="ew",
            pady=6,
        )

        dims = ttk.Frame(product_editor)
        for index, variable in enumerate(self.product_dim_vars):
            ttk.Entry(dims, textvariable=variable, width=10).pack(side=LEFT)
            if index < 2:
                ttk.Label(dims, text=" x ", padding=(4, 0)).pack(side=LEFT)
        ttk.Combobox(dims, textvariable=self.product_dim_unit_var, values=("mm", "cm", "m"), width=5, state="readonly").pack(
            side=LEFT,
            padx=(6, 0),
        )
        ttk.Label(product_editor, text="Dimensions").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=6)
        dims.grid(row=1, column=1, sticky="w", pady=6)

        weight = ttk.Frame(product_editor)
        ttk.Entry(weight, textvariable=self.product_weight_var, width=12).pack(side=LEFT)
        ttk.Combobox(
            weight,
            textvariable=self.product_weight_unit_var,
            values=("kg", "g"),
            width=5,
            state="readonly",
        ).pack(side=LEFT, padx=(6, 0))
        ttk.Label(product_editor, text="Weight").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=6)
        weight.grid(row=2, column=1, sticky="w", pady=6)

        product_buttons = ttk.Frame(product_editor)
        product_buttons.grid(row=3, column=1, sticky="e", pady=(8, 0))
        ttk.Button(product_buttons, text="Clear", command=self.clear_product_form).pack(side=LEFT, padx=(0, 8))
        self.save_product_button = ttk.Button(product_buttons, text="Save Product", command=self.save_product)
        self.save_product_button.pack(side=LEFT)

        product_side = ttk.Frame(sample_frame)
        product_side.grid(row=0, column=1, rowspan=2, sticky="nsew")
        product_side.rowconfigure(1, weight=1)
        ttk.Label(product_side, text="Current Products", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        self.products_tree = ttk.Treeview(product_side, columns=("name", "dimensions", "weight"), show="headings", height=6)
        self.products_tree.heading("name", text="Product")
        self.products_tree.heading("dimensions", text="Dimensions")
        self.products_tree.heading("weight", text="Weight")
        self.products_tree.column("name", width=120)
        self.products_tree.column("dimensions", width=120)
        self.products_tree.column("weight", width=70)
        self.products_tree.grid(row=1, column=0, sticky="nsew", pady=(8, 8))
        self.products_tree.bind("<<TreeviewSelect>>", self.load_selected_product)

        product_scrollbar = ttk.Scrollbar(product_side, orient=VERTICAL, command=self.products_tree.yview)
        product_scrollbar.grid(row=1, column=1, sticky="ns", pady=(8, 8))
        self.products_tree.configure(yscrollcommand=product_scrollbar.set)

        product_actions = ttk.Frame(product_side)
        product_actions.grid(row=2, column=0, sticky="ew")
        ttk.Button(product_actions, text="Add New", command=self.clear_product_form).pack(side=LEFT)
        ttk.Button(product_actions, text="Remove", command=self.remove_selected_product).pack(side=LEFT, padx=(8, 0))
        ttk.Button(product_actions, text="Duplicate", command=self.duplicate_selected_product).pack(side=LEFT, padx=(8, 0))

        customer_frame = ttk.LabelFrame(self.info_tab, text="Customer", style="Section.TLabelframe", padding=12)
        customer_frame.pack(fill="x", pady=(0, 12))
        self._add_entry(customer_frame, "Customer", "customer", 0, 0)
        self._add_entry(customer_frame, "Designation", "designation", 0, 2)
        self._add_entry(customer_frame, "Company", "company", 1, 0)
        self._add_entry(customer_frame, "Tel No.", "tel_no", 1, 2)
        self._add_entry(customer_frame, "E-mail", "email", 2, 0)
        self._add_entry(customer_frame, "Address", "address", 2, 2, width=48)

        for frame in (quote_frame, sample_frame, customer_frame):
            frame.columnconfigure(1, weight=1)
            frame.columnconfigure(3, weight=1)

    def _build_tests_tab(self) -> None:
        toolbar = ttk.Frame(self.tests_tab)
        toolbar.pack(fill="x", pady=(0, 10))
        ttk.Button(toolbar, text="Add New Test", command=self.clear_test_form).pack(side=LEFT)
        ttk.Button(toolbar, text="Remove Selected Test", command=self.remove_selected_test).pack(side=LEFT, padx=(8, 0))
        ttk.Button(toolbar, text="Duplicate Selected", command=self.duplicate_selected_test).pack(side=LEFT, padx=(8, 0))

        body = ttk.Frame(self.tests_tab)
        body.pack(fill=BOTH, expand=True)
        body.columnconfigure(0, weight=2)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        editor = ttk.LabelFrame(body, text="Test Details", style="Section.TLabelframe", padding=12)
        editor.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        editor.columnconfigure(1, weight=1)
        editor.rowconfigure(4, weight=1)

        self.test_name_var = StringVar()
        self.test_standard_var = StringVar()
        ttk.Label(editor, text="Test type").grid(row=0, column=0, sticky="w", pady=4)
        self.test_name_combo = AutocompleteCombobox(editor, textvariable=self.test_name_var, values=TEST_NAMES)
        self.test_name_combo.grid(row=0, column=1, sticky="ew", pady=4)
        self.test_name_combo.bind("<<ComboboxSelected>>", self.test_type_changed, add="+")
        self.test_name_combo.bind("<<AutocompleteAccepted>>", self.test_type_changed, add="+")
        self.test_name_combo.bind("<FocusOut>", self.test_type_changed, add="+")

        ttk.Label(editor, text="Test standard").grid(row=1, column=0, sticky="w", pady=4)
        self.test_standard_combo = AutocompleteCombobox(
            editor,
            textvariable=self.test_standard_var,
            values=ALL_TEST_STANDARDS,
        )
        self.test_standard_combo.grid(row=1, column=1, sticky="ew", pady=4)
        self.test_standard_combo.bind("<<ComboboxSelected>>", self.test_standard_changed, add="+")
        self.test_standard_combo.bind("<<AutocompleteAccepted>>", self.test_standard_changed, add="+")
        self.test_standard_combo.bind("<FocusOut>", self.test_standard_changed, add="+")

        product_assignment = ttk.Frame(editor)
        product_assignment.grid(row=2, column=1, sticky="ew", pady=4)
        self.assign_products_check = ttk.Checkbutton(
            product_assignment,
            text="Assign this test to product(s)",
            variable=self.assign_products_var,
            command=self.toggle_product_assignment,
        )
        self.assign_products_check.pack(side=LEFT)
        self.test_products_listbox = Listbox(product_assignment, selectmode="multiple", height=4, exportselection=False)
        self.test_products_listbox.pack(side=LEFT, fill="x", expand=True, padx=(12, 0))
        ttk.Label(editor, text="Products").grid(row=2, column=0, sticky="nw", pady=4)

        small_fields = ttk.Frame(editor)
        small_fields.grid(row=3, column=1, sticky="w", pady=4)
        self.qty_var = StringVar(value="1")
        self.batch_var = StringVar(value="1")
        self.cost_var = StringVar()
        self._small_entry(small_fields, "Sample Qty", self.qty_var, 10)
        self._small_entry(small_fields, "Test Qty", self.batch_var, 10)
        self._small_entry(small_fields, "Total Cost", self.cost_var, 16, suffix="Rs.")
        ttk.Label(editor, text="").grid(row=3, column=0, sticky="w", pady=4)

        ttk.Label(editor, text="Test method / requirements").grid(row=4, column=0, sticky="nw", pady=4)
        self.requirements_text = Text(editor, height=18, wrap="word", undo=True)
        self.requirements_text.grid(row=4, column=1, sticky="nsew", pady=4)

        editor_buttons = ttk.Frame(editor)
        editor_buttons.grid(row=5, column=1, sticky="e", pady=(10, 0))
        ttk.Button(editor_buttons, text="Clear", command=self.clear_test_form).pack(side=LEFT, padx=(0, 8))
        self.save_test_button = ttk.Button(editor_buttons, text="Save Test", command=self.save_test)
        self.save_test_button.pack(side=LEFT)

        side = ttk.Frame(body)
        side.grid(row=0, column=1, sticky="nsew")
        side.rowconfigure(1, weight=1)
        ttk.Label(side, text="Current Tests", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        self.tests_tree = ttk.Treeview(
            side,
            columns=("name", "standard", "products", "qty", "batch", "cost"),
            show="headings",
            height=14,
        )
        self.tests_tree.heading("name", text="Test")
        self.tests_tree.heading("standard", text="Standard")
        self.tests_tree.heading("products", text="Products")
        self.tests_tree.heading("qty", text="Sample Qty")
        self.tests_tree.heading("batch", text="Test Qty")
        self.tests_tree.heading("cost", text="Cost")
        self.tests_tree.column("name", width=150)
        self.tests_tree.column("standard", width=145)
        self.tests_tree.column("products", width=130)
        self.tests_tree.column("qty", width=82, anchor="center")
        self.tests_tree.column("batch", width=72, anchor="center")
        self.tests_tree.column("cost", width=80, anchor="e")
        self.tests_tree.grid(row=1, column=0, sticky="nsew", pady=(8, 12))
        self.tests_tree.bind("<<TreeviewSelect>>", self.load_selected_test)

        scrollbar = ttk.Scrollbar(side, orient=VERTICAL, command=self.tests_tree.yview)
        scrollbar.grid(row=1, column=1, sticky="ns", pady=(8, 12))
        self.tests_tree.configure(yscrollcommand=scrollbar.set)

        ttk.Label(side, text="Available test types").grid(row=2, column=0, sticky="w")
        names = "\n".join(TEST_NAMES[:120])
        available = Text(side, height=10, width=32, wrap="word")
        available.insert("1.0", names)
        available.configure(state="disabled")
        available.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=(4, 0))
        self.refresh_test_product_options()

    def _build_footer(self, parent) -> None:
        footer = ttk.Frame(parent)
        footer.pack(fill="x", pady=(12, 0))
        ttk.Button(footer, text="Generate Excel Quote", command=self.generate).pack(side=RIGHT)
        ttk.Button(footer, text="Open Output Folder", command=self.open_output_folder).pack(side=RIGHT, padx=(0, 8))
        ttk.Button(footer, text="Load Existing Quote", command=self.load_existing_quote).pack(side=RIGHT, padx=(0, 8))
        self.status_var = StringVar(value="Ready")
        ttk.Label(footer, textvariable=self.status_var).pack(side=LEFT)

    def _set_field(self, key: str, value: str = "") -> None:
        self.fields[key] = StringVar(value=value)

    def _add_entry(self, parent, label: str, key: str, row: int, col: int, width: int = 28) -> None:
        self._set_field(key, "")
        entry = ttk.Entry(parent, textvariable=self.fields[key], width=width)
        self._add_labeled_widget(parent, label, entry, row, col)

    def _add_labeled_widget(self, parent, label: str, widget, row: int, col: int) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=col, sticky="w", padx=(0, 8), pady=6)
        widget.grid(row=row, column=col + 1, sticky="ew", padx=(0, 18), pady=6)

    def _small_entry(self, parent, label: str, variable: StringVar, width: int, suffix: str = "") -> None:
        ttk.Label(parent, text=label).pack(side=LEFT, padx=(0, 4))
        ttk.Entry(parent, textvariable=variable, width=width).pack(side=LEFT, padx=(0, 12))
        if suffix:
            ttk.Label(parent, text=suffix).pack(side=LEFT, padx=(0, 12))

    def toggle_discount(self) -> None:
        if self.use_discount_var.get():
            self.discount_frame.pack(side=LEFT)
        else:
            self.discount_frame.pack_forget()

    def toggle_product_assignment(self) -> None:
        state = "normal" if self.assign_products_var.get() else "disabled"
        self.test_products_listbox.configure(state=state)
        if not self.assign_products_var.get():
            self.test_products_listbox.selection_clear(0, END)

    def refresh_test_product_options(self) -> None:
        if not hasattr(self, "test_products_listbox"):
            return
        selected_keys = set(self.selected_test_product_keys())
        self.test_products_listbox.configure(state="normal")
        self.test_products_listbox.delete(0, END)
        for product in self.products:
            self.test_products_listbox.insert(END, product_quote_label(product))
            if product_key(product) in selected_keys:
                self.test_products_listbox.selection_set(END)
        self.toggle_product_assignment()

    def selected_test_product_keys(self) -> list[str]:
        if not hasattr(self, "test_products_listbox"):
            return []
        keys = []
        for index in self.test_products_listbox.curselection():
            if index < len(self.products):
                keys.append(product_key(self.products[index]))
        return keys

    def set_selected_test_product_keys(self, keys: list[str]) -> None:
        self.refresh_test_product_options()
        key_set = set(keys)
        self.test_products_listbox.configure(state="normal")
        self.test_products_listbox.selection_clear(0, END)
        for index, product in enumerate(self.products):
            if product_key(product) in key_set:
                self.test_products_listbox.selection_set(index)
        self.assign_products_var.set(bool(keys))
        self.toggle_product_assignment()

    def test_type_changed(self, _event=None) -> None:
        test_type = self.test_name_var.get().strip()
        standards = TEST_STANDARDS_BY_TYPE.get(test_type, ALL_TEST_STANDARDS)
        self.test_standard_combo.set_values(standards)

    def test_standard_changed(self, _event=None) -> None:
        standard = self.test_standard_var.get().strip()
        if not standard:
            return
        if not self.test_name_var.get().strip():
            matched_type = TEST_TYPE_BY_STANDARD.get(standard)
            if matched_type:
                self.test_name_var.set(matched_type)
                self.test_type_changed()

    def clear_test_form(self) -> None:
        self.editing_index = None
        self.test_name_var.set("")
        self.test_standard_var.set("")
        self.test_standard_combo.set_values(ALL_TEST_STANDARDS)
        self.assign_products_var.set(False)
        self.refresh_test_product_options()
        self.qty_var.set("1")
        self.batch_var.set("1")
        self.cost_var.set("")
        self.requirements_text.delete("1.0", END)
        self.save_test_button.configure(text="Save Test")
        self.tests_tree.selection_remove(self.tests_tree.selection())
        self.test_name_combo.focus_set()

    def save_test(self) -> None:
        name = self.test_name_var.get().strip()
        standard = self.test_standard_var.get().strip()
        requirements = self.requirements_text.get("1.0", END).strip()
        if not name and not standard:
            messagebox.showwarning("Missing test", "Enter a test type or test standard before saving.")
            return
        product_keys = self.selected_test_product_keys() if self.assign_products_var.get() else []
        if self.assign_products_var.get() and not product_keys:
            messagebox.showwarning("Missing products", "Select at least one product for this test, or turn off product assignment.")
            return
        item = TestItem(
            name=name,
            requirements=requirements,
            qty=self.qty_var.get().strip() or "1",
            batch=self.batch_var.get().strip() or "1",
            total_cost=self.cost_var.get().strip(),
            standard=standard,
            product_keys=product_keys,
        )
        if self.editing_index is None:
            self.tests.append(item)
        else:
            self.tests[self.editing_index] = item
        self.refresh_tests()
        self.clear_test_form()

    def selected_index(self) -> int | None:
        selection = self.tests_tree.selection()
        if not selection:
            return None
        return int(selection[0])

    def load_selected_test(self, _event=None) -> None:
        index = self.selected_index()
        if index is None or index >= len(self.tests):
            return
        item = self.tests[index]
        self.editing_index = index
        self.test_name_var.set(item.name)
        self.test_standard_var.set(item.standard)
        self.test_type_changed()
        self.set_selected_test_product_keys(item.product_keys)
        self.qty_var.set(item.qty)
        self.batch_var.set(item.batch)
        self.cost_var.set(item.total_cost)
        self.requirements_text.delete("1.0", END)
        self.requirements_text.insert("1.0", item.requirements)
        self.save_test_button.configure(text="Update Test")

    def remove_selected_test(self) -> None:
        index = self.selected_index()
        if index is None:
            messagebox.showinfo("Remove test", "Select a test to remove.")
            return
        if not messagebox.askyesno("Remove test", "Are you sure you want to remove the selected test?"):
            return
        del self.tests[index]
        self.refresh_tests()
        self.clear_test_form()

    def duplicate_selected_test(self) -> None:
        index = self.selected_index()
        if index is None:
            messagebox.showinfo("Duplicate test", "Select a test to duplicate.")
            return
        item = self.tests[index]
        self.tests.insert(
            index + 1,
            TestItem(
                item.name,
                item.requirements,
                item.qty,
                item.batch,
                item.total_cost,
                item.standard,
                list(item.product_keys),
            ),
        )
        self.refresh_tests()

    def refresh_tests(self) -> None:
        for item_id in self.tests_tree.get_children():
            self.tests_tree.delete(item_id)
        for index, item in enumerate(self.tests):
            products_text = ", ".join(item.product_keys)
            self.tests_tree.insert(
                "",
                END,
                iid=str(index),
                values=(item.name, item.standard, products_text, item.qty, item.batch, item.total_cost),
            )
        count = len(self.tests)
        self.status_var.set(f"{count} test{'s' if count != 1 else ''} ready")

    def current_product_from_form(self) -> ProductItem | None:
        name = self.product_name_var.get().strip()
        if not name:
            return None
        return ProductItem(
            name=name,
            dim1=self.product_dim_vars[0].get().strip(),
            dim2=self.product_dim_vars[1].get().strip(),
            dim3=self.product_dim_vars[2].get().strip(),
            dim_unit=self.product_dim_unit_var.get().strip() or "mm",
            weight=self.product_weight_var.get().strip(),
            weight_unit=self.product_weight_unit_var.get().strip() or "kg",
        )

    def clear_product_form(self) -> None:
        self.product_editing_index = None
        self.product_name_var.set("")
        for variable in self.product_dim_vars:
            variable.set("")
        self.product_dim_unit_var.set("mm")
        self.product_weight_var.set("")
        self.product_weight_unit_var.set("kg")
        self.save_product_button.configure(text="Save Product")
        self.products_tree.selection_remove(self.products_tree.selection())

    def save_product(self) -> None:
        item = self.current_product_from_form()
        if item is None:
            messagebox.showwarning("Missing product", "Enter a product name before saving.")
            return
        if self.product_editing_index is None:
            self.products.append(item)
        else:
            old_key = product_key(self.products[self.product_editing_index])
            self.products[self.product_editing_index] = item
            new_key = product_key(item)
            if old_key != new_key:
                for test in self.tests:
                    test.product_keys = [new_key if key == old_key else key for key in test.product_keys]
                self.refresh_tests()
        self.refresh_products()
        self.clear_product_form()

    def product_selected_index(self) -> int | None:
        selection = self.products_tree.selection()
        if not selection:
            return None
        return int(selection[0])

    def load_selected_product(self, _event=None) -> None:
        index = self.product_selected_index()
        if index is None or index >= len(self.products):
            return
        item = self.products[index]
        self.product_editing_index = index
        self.product_name_var.set(item.name)
        for variable, value in zip(self.product_dim_vars, (item.dim1, item.dim2, item.dim3)):
            variable.set(value)
        self.product_dim_unit_var.set(item.dim_unit or "mm")
        self.product_weight_var.set(item.weight)
        self.product_weight_unit_var.set(item.weight_unit or "kg")
        self.save_product_button.configure(text="Update Product")

    def remove_selected_product(self) -> None:
        index = self.product_selected_index()
        if index is None:
            messagebox.showinfo("Remove product", "Select a product to remove.")
            return
        if not messagebox.askyesno("Remove product", "Are you sure you want to remove the selected product?"):
            return
        removed_key = product_key(self.products[index])
        del self.products[index]
        for test in self.tests:
            test.product_keys = [key for key in test.product_keys if key != removed_key]
        self.refresh_products()
        self.refresh_tests()
        self.clear_product_form()

    def duplicate_selected_product(self) -> None:
        index = self.product_selected_index()
        if index is None:
            messagebox.showinfo("Duplicate product", "Select a product to duplicate.")
            return
        item = self.products[index]
        self.products.insert(
            index + 1,
            ProductItem(item.name, item.dim1, item.dim2, item.dim3, item.dim_unit, item.weight, item.weight_unit),
        )
        self.refresh_products()

    def refresh_products(self) -> None:
        for item_id in self.products_tree.get_children():
            self.products_tree.delete(item_id)
        for index, item in enumerate(self.products):
            dimensions = "x".join(part for part in (item.dim1, item.dim2, item.dim3) if part)
            if dimensions:
                dimensions = f"{dimensions} {item.dim_unit}"
            weight = f"{item.weight} {item.weight_unit}" if item.weight else ""
            self.products_tree.insert("", END, iid=str(index), values=(item.name, dimensions, weight))
        count = len(self.products)
        self.status_var.set(f"{count} product{'s' if count != 1 else ''} ready")
        self.refresh_test_product_options()

    def products_for_generation(self) -> list[ProductItem]:
        products = list(self.products)
        pending = self.current_product_from_form()
        if pending is not None:
            if self.product_editing_index is None:
                products.append(pending)
            elif self.product_editing_index < len(products):
                products[self.product_editing_index] = pending
        return products

    def collect_data(self) -> dict[str, str]:
        data = {key: variable.get() for key, variable in self.fields.items()}
        data["discount_enabled"] = "1" if self.use_discount_var.get() else ""
        data["discount_percent"] = self.discount_percent_var.get()
        data["instate"] = "1" if self.instate_var.get() else ""
        return data

    def generate(self) -> None:
        try:
            output_path = generate_quote(self.collect_data(), self.tests, self.products_for_generation())
        except Exception as exc:
            messagebox.showerror("Could not generate quote", str(exc))
            return

        self.last_output = output_path
        self.status_var.set(f"Generated {output_path.name}")
        if messagebox.askyesno("Quote generated", f"Saved:\n{output_path}\n\nOpen the Excel file now?"):
            self.open_path(output_path)

    def load_existing_quote(self) -> None:
        path = filedialog.askopenfilename(
            title="Load existing quote",
            filetypes=(("Excel workbooks", "*.xlsx"), ("All files", "*.*")),
        )
        if not path:
            return
        try:
            data, tests, products = load_quote(path)
        except Exception as exc:
            messagebox.showerror("Could not load quote", str(exc))
            return

        for key, variable in self.fields.items():
            if key in data:
                variable.set(data[key])
        self.quote_made_by_menu.set_selection(self.fields["quote_made_by"].get())
        self.technical_engineers_menu.set_selection(self.fields["technical_engineers"].get())
        self.use_discount_var.set(data.get("discount_enabled") == "1")
        self.discount_percent_var.set(data.get("discount_percent", ""))
        self.instate_var.set(data.get("instate") == "1")
        self.toggle_discount()

        self.tests = tests
        self.products = products
        self.refresh_tests()
        self.refresh_products()
        self.clear_test_form()
        self.clear_product_form()
        self.status_var.set(f"Loaded {Path(path).name}")

    def open_output_folder(self) -> None:
        output_dir = Path(__file__).resolve().parent.parent / "outputs"
        output_dir.mkdir(exist_ok=True)
        self.open_path(output_dir)

    def open_path(self, path: Path) -> None:
        if sys.platform == "win32":
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)


if __name__ == "__main__":
    app = QuoteApp()
    app.mainloop()
