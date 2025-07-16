import tkinter as tk
from tkinter import filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import pandas as pd
import os

# === Global State ===
df = None
opened_file_path = ""
file_extension = ""
current_theme = "dark"

# === Themes ===
light_theme = {
    "bg": "#ffffff",
    "fg": "#000000",
    "entry_bg": "#ffffff",
    "entry_fg": "#000000",
    "button_bg": "#f0f0f0",
    "label_bg": "#e0e0e0"
}

dark_theme = {
    "bg": "#2d2d2d",
    "fg": "#ffffff",
    "entry_bg": "#3c3c3c",
    "entry_fg": "#ffffff",
    "button_bg": "#444444",
    "label_bg": "#3c3c3c",
    "info_bg": "#000000"
}

def configure_theme():
    theme = dark_theme if current_theme == "dark" else light_theme

    app.configure(bg=theme["bg"])
    top_controls.configure(bg=theme["bg"])
    drop_label.configure(bg=theme["label_bg"], fg=theme["fg"])
    message_label.configure(bg=theme["bg"], fg=theme["fg"])
    canvas.configure(bg=theme["bg"])
    table_frame.configure(bg=theme["bg"])
    table_inner_frame.configure(bg=theme["bg"])
    message_label.configure(bg=theme["info_bg"])

    for widget in top_controls.winfo_children():
        widget.configure(bg=theme["button_bg"], fg=theme["fg"])

    for widget in table_inner_frame.winfo_children():
        if isinstance(widget, tk.Entry):
            widget.configure(bg=theme["entry_bg"], fg=theme["entry_fg"])
        elif isinstance(widget, tk.Label):
            widget.configure(bg=theme["label_bg"], fg=theme["fg"])
        elif isinstance(widget, tk.Button):
            widget.configure(bg=theme["button_bg"], fg=theme["fg"])

def toggle_theme():
    global current_theme
    current_theme = "dark" if current_theme == "light" else "light"
    configure_theme()

# === Open File Explorer ===

def open_excel_file_from_dialog():
    file_path = filedialog.askopenfilename(filetypes=[("Excel/CSV files", "*.xlsx *.csv")])
    if file_path:
        open_file(file_path)

# === Drag & Drop File Open ===

def on_drop(event):
    file_path = event.data.strip('{}')
    if file_path.lower().endswith(('.xlsx', '.csv')):
        open_file(file_path)
    else:
        message_label.config(text="Only .xlsx or .csv files are supported.")

def toggle_edit_buttons(state):
    for b in [save_button, save_close_button, close_button, undo_button, redo_button, stack_button]:
        b.config(state=tk.NORMAL if state else tk.DISABLED)

def open_file(file_path):
    global df, opened_file_path, file_extension
    opened_file_path = file_path
    file_extension = os.path.splitext(file_path)[-1].lower()

    try:
        if file_extension == ".csv":
            df = pd.read_csv(file_path, header=None).fillna(0)
        elif file_extension == ".xlsx":
            df = pd.read_excel(file_path, header=None).fillna(0)
        else:
            message_label.config(text="Unsupported file type.")
            return
    except Exception as e:
        message_label.config(text=f"Error opening file: {e}")
        return

    
    display_data(df)
    toggle_edit_buttons(True)
    drop_label.pack_forget()
    message_label.config(text=f"Opened: {file_path}")

def save_file():
    if not opened_file_path:
        message_label.config(text="No file opened.")
        return

    updated_df = get_updated_data()
    try:
        if file_extension == ".csv":
            updated_df.to_csv(opened_file_path, header=False, index=False)
        elif file_extension == ".xlsx":
            updated_df.to_excel(opened_file_path, header=False, index=False)
        message_label.config(text=f"File saved: {opened_file_path}")
    except Exception as e:
        message_label.config(text=f"Error saving file: {e}")

def save_and_close():
    save_file()
    unload_sheet()

def unload_sheet():
    for widget in table_inner_frame.winfo_children():
        widget.destroy()
    toggle_edit_buttons(False)
    drop_label.pack(after=top_controls, fill=tk.X, padx=10, pady=5)
    global df, opened_file_path, file_extension
    df = None
    opened_file_path = ""
    file_extension = ""
    message_label.config(text="Sheet closed.")

def get_column_letter(n):
    # Convert 1-based number to Excel-style letter (A, B, ..., Z, AA, AB, ...)
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result

def display_data(dataframe):
    def clear_undo_stack_on_edit(event):
        undo_stack.clear()
        redo_stack.clear()
        history.clear()
        update_undo_window()
    for widget in table_inner_frame.winfo_children():
        widget.destroy()

    global df
    df = dataframe.copy()
    entry_widgets = []
    num_rows, num_cols = df.shape

    col_widths = []
    for c in range(num_cols):
        max_len = max([len(str(df.iat[r, c])) for r in range(num_rows)] + [1])
        col_widths.append(max(min(max_len, 30), 10))

    def remove_column(c):
        updated_df = get_updated_data()
        updated_df.drop(columns=[c], inplace=True)
        updated_df = updated_df.reset_index(drop=True)
        push_undo(f"Deleted Column {get_column_letter(c + 1)}")
        display_data(updated_df)

    def remove_row(r):
        updated_df = get_updated_data()
        updated_df.drop(index=[r], inplace=True)
        updated_df = updated_df.reset_index(drop=True)
        push_undo(f"Deleted Row {str(r + 1)}")
        display_data(updated_df)

    def auto_resize(event, col):
        updated_df = get_updated_data()
        max_len = max([len(str(updated_df.iloc[r, col])) for r in range(len(updated_df))] + [1])
        width = max(min(max_len, 30), 10)
        for w in table_inner_frame.grid_slaves():
            info = w.grid_info()
            if info["column"] == col + 1 and isinstance(w, tk.Entry):
                w.config(width=width)
    
    def auto_resize_unlimited(event, col):
        updated_df = get_updated_data()
        max_len = max([len(str(updated_df.iloc[r, col])) for r in range(len(updated_df))] + [1])
        width = max(max_len, 10)
        for w in table_inner_frame.grid_slaves():
            info = w.grid_info()
            if info["column"] == col + 1 and isinstance(w, tk.Entry):
                w.config(width=width)

    def bind_entry_navigation(entry, r, c):
        def move_focus(dr, dc):
            nr, nc = r + dr, c + dc
            if dr == 1 and nr >= display_data.num_rows:
                add_row()
            if dc == 1 and nc >= display_data.num_cols:
                add_column()
            if 0 <= nr < display_data.num_rows and 0 <= nc < display_data.num_cols:
                idx = nr * display_data.num_cols + nc
                if idx < len(display_data.entry_widgets):
                    target = display_data.entry_widgets[idx]
                    target.focus_set()
                    target.icursor(tk.END)

        entry.bind("<Up>", lambda e: (move_focus(-1, 0), "break"))
        entry.bind("<Down>", lambda e: (move_focus(1, 0), "break"))
        entry.bind("<Return>", lambda e: (move_focus(1, 0), "break"))
        entry.bind("<Tab>", lambda e: (move_focus(0, 1), "break"))
        entry.bind("<Shift-Tab>", lambda e: (move_focus(0, -1), "break"))

    for r in range(num_rows + 2):
        for c in range(num_cols + 2):
            if r == 0 and c == 0:
                tk.Label(table_inner_frame, text="", borderwidth=1, relief="solid", width=5).grid(row=r, column=c)
            elif r == 0 and 1 <= c <= num_cols:
                lbl = tk.Label(table_inner_frame, text=get_column_letter(c), borderwidth=1, relief="solid", width=col_widths[c - 1])
                lbl.grid(row=r, column=c)
                lbl.bind("<Double-Button-1>", lambda e, col=c - 1: auto_resize(e, col))
                lbl.bind("<Button-1>", lambda e, col=c - 1: auto_resize_unlimited(e, col))
                lbl.bind("<Button-3>", lambda e, col=c - 1: remove_column(col))
            elif r == 0 and c == num_cols + 1:
                tk.Button(table_inner_frame, text="+", command=add_column, width=4).grid(row=r, column=c)
            elif c == 0 and 1 <= r <= num_rows:
                lbl = tk.Label(table_inner_frame, text=str(r), borderwidth=1, relief="solid", width=5)
                lbl.grid(row=r, column=c)
                lbl.bind("<Button-3>", lambda e, row=r - 1: remove_row(row))
            elif c == 0 and r == num_rows + 1:
                tk.Button(table_inner_frame, text="+", command=add_row, width=4).grid(row=r, column=c)
            elif 1 <= r <= num_rows and 1 <= c <= num_cols:
                value = df.iat[r - 1, c - 1]
                entry = tk.Entry(table_inner_frame, width=col_widths[c - 1])
                entry.insert(0, str(value))
                entry.bind("<KeyRelease>", clear_undo_stack_on_edit)
                entry.grid(row=r, column=c, sticky="nsew")
                bind_entry_navigation(entry, r - 1, c - 1)
                entry_widgets.append(entry)

    display_data.entry_widgets = entry_widgets
    display_data.num_rows = num_rows
    display_data.num_cols = num_cols

    table_inner_frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))
    configure_theme()

def get_updated_data():
    entries = display_data.entry_widgets
    num_rows = display_data.num_rows
    num_cols = display_data.num_cols

    if len(entries) != num_rows * num_cols:
        message_label.config(text="Entry size mismatch; rebuilding.")
        return df.copy()

    data = []
    for r in range(num_rows):
        row_data = []
        for c in range(num_cols):
            index = r * num_cols + c
            row_data.append(entries[index].get())
        data.append(row_data)

    return pd.DataFrame(data)

def add_row():
    updated_df = get_updated_data()
    if updated_df is not None:
        new_row = [0] * display_data.num_cols
        updated_df.loc[len(updated_df)] = new_row
        push_undo("Add Row")
        display_data(updated_df)
        

def add_column():
    updated_df = get_updated_data()
    if updated_df is not None:
        updated_df[len(updated_df.columns)] = 0
        push_undo("Add Column")
        display_data(updated_df)

# === Undo/Redo History ===
undo_stack = []
redo_stack = []
history = []
undo_window = None
undo_listbox = None

def push_undo(action):
    undo_stack.append(get_updated_data())
    history.append(action)
    if len(undo_stack) > 50:
        undo_stack.pop(0)
        history.pop(0)
    redo_stack.clear()
    update_undo_window()

def undo_last_action(event=None):
    if len(undo_stack) > 0:
        redo_stack.append(get_updated_data())
        last_df = undo_stack.pop()
        history.pop()
        display_data(last_df)
        update_undo_window()

def redo_last_action(event=None):
    if redo_stack:
        undo_stack.append(get_updated_data())
        history.append("Redo")
        next_df = redo_stack.pop()
        display_data(next_df)
        update_undo_window()

def update_undo_window():
    global undo_listbox
    if undo_listbox and undo_listbox.winfo_exists():
        try:
            undo_listbox.delete(0, tk.END)
        except:
            undo_listbox = None
        
        for i, df in enumerate(undo_stack):
            if history[i] == "Add Column":
                undo_listbox.insert(tk.END, f"{i+1}) {history[i]} {get_column_letter(df.shape[1] + 1)}")
            elif history[i] == "Add Row":
                undo_listbox.insert(tk.END, f"{i+1}) {history[i]} {df.shape[0] + 1}")
            elif history[i] == "Redo":
                redo_df = get_updated_data()
                undo_listbox.insert(tk.END, f"{i+1}) {history[i]} ({redo_df.shape[0]}x{get_column_letter(redo_df.shape[1])})")
            else:
                undo_listbox.insert(tk.END, f"{i+1}) {history[i]}")


def open_undo_window():
    global undo_window, undo_listbox
    if undo_window and undo_window.winfo_exists():
        undo_window.lift()
        return
    undo_window = tk.Toplevel(app)
    undo_window.title("Undo Stack")
    undo_window.geometry("250x300")
    undo_listbox = tk.Listbox(undo_window)
    undo_listbox.pack(fill=tk.BOTH, expand=True)
    update_undo_window()

# === GUI Setup ===
app = TkinterDnD.Tk()
app.title("Better Excel")
app.geometry("1000x700")

# === Top Controls ===
top_controls = tk.Frame(app)
top_controls.pack(fill=tk.X, padx=10, pady=5)

open_button = tk.Button(top_controls, text="Open File", command=open_excel_file_from_dialog)
open_button.pack(side=tk.LEFT, padx=5)

save_button = tk.Button(top_controls, text="Save Changes", command=save_file, state=tk.DISABLED)
save_button.pack(side=tk.LEFT, padx=5)

save_close_button = tk.Button(top_controls, text="Save & Close", command=save_and_close, state=tk.DISABLED)
save_close_button.pack(side=tk.LEFT, padx=5)

close_button = tk.Button(top_controls, text="Close Without Saving", command=unload_sheet, state=tk.DISABLED)
close_button.pack(side=tk.LEFT, padx=5)

theme_button = tk.Button(top_controls, text="Toggle Theme", command=toggle_theme)
theme_button.pack(side=tk.RIGHT, padx=5)

undo_button = tk.Button(top_controls, text="Undo", command=undo_last_action, state=tk.DISABLED)
undo_button.pack(side=tk.RIGHT, padx=2)

redo_button = tk.Button(top_controls, text="Redo", command=redo_last_action, state=tk.DISABLED)
redo_button.pack(side=tk.RIGHT, padx=2)

stack_button = tk.Button(top_controls, text="Show Undo Stack", command=open_undo_window, state=tk.DISABLED)
stack_button.pack(side=tk.RIGHT, padx=2)

# === Drag & Drop Label ===
drop_label = tk.Label(app, text="Or drag and drop an Excel (.xlsx) or CSV (.csv) file here",
                      relief="groove", height=5)
drop_label.pack(after=top_controls, fill=tk.X, padx=10, pady=10)
drop_label.drop_target_register(DND_FILES)
drop_label.dnd_bind('<<Drop>>', on_drop)

# === Scrollable Table Area ===
table_frame = tk.Frame(app)
table_frame.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)

canvas = tk.Canvas(table_frame)
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scrollbar_y = tk.Scrollbar(table_frame, orient="vertical", command=canvas.yview)
scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)

scrollbar_x = tk.Scrollbar(app, orient="horizontal", command=canvas.xview)
scrollbar_x.pack(fill=tk.X)

canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

table_inner_frame = tk.Frame(canvas)
canvas.create_window((0, 0), window=table_inner_frame, anchor="nw")

def on_mousewheel(event):
    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

def on_shift_mousewheel(event):
    canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")

canvas.bind_all("<MouseWheel>", on_mousewheel)
canvas.bind_all("<Shift-MouseWheel>", on_shift_mousewheel)
canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

# === Keyboard Bindings ===
app.bind("<Control-z>", undo_last_action)
app.bind("<Control-y>", redo_last_action)

# === Message Label ===

message_label = tk.Label(app, text="",height=1)
message_label.pack(side=tk.RIGHT, fill=tk.X, expand=True)

configure_theme()
app.mainloop()
