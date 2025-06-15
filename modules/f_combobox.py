import tkinter as tk
import customtkinter as ctk

class FilterableCombobox(ctk.CTkFrame):
    def __init__(self, parent, values, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.values = values
        self.filtered_values = values

        self.entry = ctk.CTkEntry(self)
        self.entry.pack(fill="x", expand=True)

        self.listbox = tk.Listbox(self, height=5)
        self.listbox.pack(fill="x", expand=True)
        self.listbox.bind("<ButtonRelease-1>", self._on_select)

        self.entry.bind("<KeyRelease>", self._filter_values)

        self._update_listbox()

    def _update_listbox(self):
        """Update the Listbox with current filtered values."""
        self.listbox.delete(0, tk.END)
        for value in self.filtered_values:
            self.listbox.insert(tk.END, value)

    def _filter_values(self, event=None):
        """Filter values based on entry content."""
        search_term = self.entry.get().lower()
        self.filtered_values = [
            value for value in self.values if search_term in value.lower()
        ]
        self._update_listbox()

    def _on_select(self, event=None):
        """Set the selected value in the Entry field."""
        selection = self.listbox.get(self.listbox.curselection())
        self.entry.delete(0, tk.END)
        self.entry.insert(0, selection)
        self.listbox.pack_forget()  # Hide the Listbox after selection

    def get(self):
        """Return the current value of the Entry."""
        return self.entry.get()

    def set(self, value):
        """Set the value of the Entry."""
        self.entry.delete(0, tk.END)
        self.entry.insert(0, value)