import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
import pandas as pd

class SpreadsheetApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Python Spreadsheet")
        self.root.geometry("1200x700")
        self.root.configure(bg="#e0e0e0")
        
        self.create_ui()
        
    def create_ui(self):
        style = ttk.Style()
        style.configure("Treeview", font=("Arial", 12), rowheight=30, background="#ffffff", foreground="#333333")
        style.configure("Treeview.Heading", font=("Arial", 14, "bold"), background="#0078D7", foreground="black", padding=10)
        style.map("Treeview.Heading", background=[("active", "#005bb5")])
        
        self.frame = ttk.Frame(self.root, padding=15)
        self.frame.pack(fill=tk.BOTH, expand=True)
        
        self.tree = ttk.Treeview(self.frame, show="headings", height=20, selectmode="browse")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.tree.bind("<Double-1>", self.edit_cell)
        
        self.menu = tk.Menu(self.root)
        self.root.config(menu=self.menu)
        
        file_menu = tk.Menu(self.menu, tearoff=0)
        file_menu.add_command(label="Open", command=self.load_file)
        file_menu.add_command(label="Save", command=self.save_file)
        self.menu.add_cascade(label="File", menu=file_menu)
        
        function_menu = tk.Menu(self.menu, tearoff=0)
        function_menu.add_command(label="SUM", command=lambda: self.apply_function("SUM"))
        function_menu.add_command(label="AVERAGE", command=lambda: self.apply_function("AVERAGE"))
        function_menu.add_command(label="MAX", command=lambda: self.apply_function("MAX"))
        function_menu.add_command(label="MIN", command=lambda: self.apply_function("MIN"))
        function_menu.add_command(label="COUNT", command=lambda: self.apply_function("COUNT"))
        self.menu.add_cascade(label="Functions", menu=function_menu)
        
        format_menu = tk.Menu(self.menu, tearoff=0)
        format_menu.add_command(label="Bold", command=self.apply_bold)
        format_menu.add_command(label="Change Font Color", command=self.change_font_color)
        self.menu.add_cascade(label="Format", menu=format_menu)
    
    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path:
            df = pd.read_csv(file_path)
            self.populate_table(df)
    
    def save_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if file_path:
            data = [[self.tree.item(item, "values")[i] for i in range(len(self.tree["columns"]))] for item in self.tree.get_children()]
            df = pd.DataFrame(data, columns=self.tree["columns"])
            df.to_csv(file_path, index=False)
    
    def populate_table(self, df):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        
        for col in df.columns:
            self.tree.heading(col, text=col.upper(), anchor="center")
            self.tree.column(col, width=150, anchor="center")
        
        for row in df.itertuples(index=False):
            self.tree.insert("", tk.END, values=row)
    
    def edit_cell(self, event):
        selected_item = self.tree.selection()[0]
        col = self.tree.identify_column(event.x)
        col_index = int(col[1:]) - 1
        
        entry_popup = tk.Toplevel(self.root)
        entry_popup.title("Edit Cell")
        entry_popup.geometry("250x120")
        entry_popup.configure(bg="#f8f9fa")
        
        entry = tk.Entry(entry_popup, font=("Arial", 14))
        entry.pack(pady=10, padx=10)
        entry.insert(0, self.tree.item(selected_item, "values")[col_index])
        
        def save_edit():
            new_value = entry.get()
            values = list(self.tree.item(selected_item, "values"))
            values[col_index] = new_value
            self.tree.item(selected_item, values=values)
            entry_popup.destroy()
        
        save_button = tk.Button(entry_popup, text="Save", command=save_edit, font=("Arial", 12), bg="#0078D7", fg="white")
        save_button.pack(pady=5)
    
    def apply_function(self, func):
        values = []
        for item in self.tree.get_children():
            row_values = [self.tree.item(item, "values")[i] for i in range(len(self.tree["columns"]))]
            values.extend([float(v) for v in row_values if v.replace('.', '', 1).isdigit()])
        
        if not values:
            messagebox.showerror("Error", "No numeric data found.")
            return
        
        result = None
        if func == "SUM":
            result = sum(values)
        elif func == "AVERAGE":
            result = sum(values) / len(values)
        elif func == "MAX":
            result = max(values)
        elif func == "MIN":
            result = min(values)
        elif func == "COUNT":
            result = len(values)
        
        messagebox.showinfo(func, f"Result: {result}")
    
    def apply_bold(self):
        messagebox.showinfo("Info", "Bold formatting applied (UI update needed)")
    
    def change_font_color(self):
        color = colorchooser.askcolor()[1]
        if color:
            messagebox.showinfo("Info", f"Font color changed to {color} (UI update needed)")

if __name__ == "__main__":
    root = tk.Tk()
    app = SpreadsheetApp(root)
    root.mainloop()
