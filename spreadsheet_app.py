import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

class SpreadsheetApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Spreadsheet App")
        self.root.geometry("1200x700")
        self.root.configure(bg="#f0f0f0")

        self.data = pd.DataFrame()
        self.tree = ttk.Treeview(self.root, show="headings")

        self.create_menu()
        self.create_table()

    def create_menu(self):
        menu_bar = tk.Menu(self.root)
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Open CSV", command=self.open_csv)
        file_menu.add_command(label="Save CSV", command=self.save_csv)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menu_bar.add_cascade(label="File", menu=file_menu)
        self.root.config(menu=menu_bar)

    def create_table(self):
        self.tree.pack(expand=True, fill="both")

        self.tree.tag_configure('header', background='black', foreground='white', font=('Arial', 12, 'bold'))
        self.tree.bind("<Double-1>", self.edit_cell)

    def open_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path:
            try:
                self.data = pd.read_csv(file_path)
                self.display_data()
            except Exception as e:
                messagebox.showerror("Error", f"Unable to open file: {e}")

    def save_csv(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if file_path:
            try:
                self.data.to_csv(file_path, index=False)
                messagebox.showinfo("Success", "File saved successfully.")
            except Exception as e:
                messagebox.showerror("Error", f"Unable to save file: {e}")

    def display_data(self):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(self.data.columns)

        for col in self.data.columns:
            self.tree.heading(col, text=col, anchor="center")
            self.tree.column(col, width=100, anchor="center")

        for _, row in self.data.iterrows():
            self.tree.insert("", "end", values=list(row), tags=('data',))

        self.tree.tag_configure('data', font=('Arial', 10))

    def edit_cell(self, event):
        selected_item = self.tree.selection()[0]
        column = self.tree.identify_column(event.x)
        column_index = int(column.replace("#", "")) - 1
        value = self.tree.item(selected_item, "values")[column_index]

        self.popup = tk.Toplevel(self.root)
        self.popup.title("Edit Cell")
        self.popup.geometry("200x100")
        entry = tk.Entry(self.popup)
        entry.insert(0, value)
        entry.pack(pady=10)
        tk.Button(self.popup, text="Save", command=lambda: self.save_edit(selected_item, column_index, entry.get())).pack()

    def save_edit(self, item, col, new_value):
        current_values = list(self.tree.item(item, "values"))
        current_values[col] = new_value
        self.tree.item(item, values=current_values)
        self.popup.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = SpreadsheetApp(root)
    root.mainloop()
