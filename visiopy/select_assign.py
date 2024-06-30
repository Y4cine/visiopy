import tkinter as tk
from tkinter import ttk
import win32com.client
import pythoncom
import time

class SelectedShapeUpdater:
    def __init__(self, vWin):
        print("Initializing SelectedShapeUpdater...")
        self.vWin = vWin
        self.selected_field = ""
        self.selected_value = ""
        self.check_active = False

        self.init_visio()
        self.previous_selection = [shape.ID for shape in self.vWin.Selection]  # Set initial selection
        self.create_gui()

    def init_visio(self):
        print("Initializing Visio variables...")
        self.vApp = win32com.client.Dispatch("Visio.Application")
        self.vDoc = self.vApp.ActiveDocument
        self.vPg = self.vApp.ActivePage
        self.vWin = self.vApp.ActiveWindow

    def set_value(self, shape, field, value):
        try:
            if shape.CellExists(f"prop.{field}", False):
                shape.Cells(f"prop.{field}.Value").FormulaU = f'"{value}"'
        except Exception as e:
            print(f"Error setting value: {e}")

    def batch_modify_shapes(self):
        try:
            selection = self.vWin.Selection
            if selection.Count > 0 and self.check_active:
                for i in range(1, selection.Count + 1):
                    shape = selection.Item(i)
                    self.set_value(shape, self.selected_field, self.selected_value)
        except Exception as e:
            print(f"Error in batch_modify_shapes: {e}")

    def poll_selection_changes(self):
        if not self.check_active:
            return
        
        try:
            current_selection = [shape.ID for shape in self.vWin.Selection]
            if current_selection != self.previous_selection:
                self.batch_modify_shapes()
                self.previous_selection = current_selection
        except Exception as e:
            print(f"Error in poll_selection_changes: {e}")
        
        self.root.after(1000, self.poll_selection_changes)  # Poll every second

    def toggle_active(self):
        self.check_active = not self.check_active
        print(f"Active state toggled: {self.check_active}")
        if self.check_active:
            self.poll_selection_changes()

    def create_gui(self):
        self.root = tk.Tk()
        self.root.title("SelectedShapeUpdater")

        explanation = ttk.Label(self.root, text="This dialog updates the selected shape properties in Visio.")
        explanation.grid(row=0, column=0, columnspan=2, padx=10, pady=5)

        self.active_var = tk.BooleanVar()
        self.active_check = ttk.Checkbutton(self.root, text="Active", variable=self.active_var, command=self.toggle_active)
        self.active_check.grid(row=1, column=0, columnspan=2, padx=10, pady=5)

        ttk.Label(self.root, text="Property Row Name:").grid(row=2, column=0, padx=10, pady=5, sticky=tk.E)
        self.field_entry = ttk.Entry(self.root)
        self.field_entry.grid(row=2, column=1, padx=10, pady=5, sticky=tk.W)
        self.field_entry.bind("<KeyRelease>", self.on_field_change)

        ttk.Label(self.root, text="New Value:").grid(row=3, column=0, padx=10, pady=5, sticky=tk.E)
        self.value_entry = ttk.Entry(self.root)
        self.value_entry.grid(row=3, column=1, padx=10, pady=5, sticky=tk.W)
        self.value_entry.bind("<KeyRelease>", self.on_value_change)

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.mainloop()

    def on_field_change(self, event):
        self.selected_field = self.field_entry.get()
        print(f"Field changed: {self.selected_field}")

    def on_value_change(self, event):
        self.selected_value = self.value_entry.get()
        print(f"Value changed: {self.selected_value}")

    def on_closing(self):
        print("Closing application...")
        self.check_active = False
        self.root.after_cancel(self.poll_selection_changes)
        self.root.destroy()
        print("Application closed.")