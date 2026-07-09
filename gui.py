"""
gui.py
------
GUI layer only.
Handles user interaction and delegates all data operations
to data_manager.DataManager.

Features:
    - Search box (Everything-style live search)
    - Import CSV / Excel files
    - Drag & Drop files directly into main window
    - Delete selected rows
    - Double click cell editing
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from data_manager import DataManager


# Optional drag & drop support
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False


class ImportDialog(tk.Toplevel):
    """
    File import dialog.
    Only supports selecting CSV / Excel files.
    """

    def __init__(self, master, on_import):
        super().__init__(master)

        self.title("Import File")
        self.geometry("450x220")

        self.transient(master)
        self.grab_set()

        self.on_import = on_import
        self.file_path = None

        tk.Label(
            self,
            text="Select a CSV or Excel file to import:",
            anchor="w"
        ).pack(
            fill="x",
            padx=15,
            pady=(20, 10)
        )

        self.file_label = tk.Label(
            self,
            text="No file selected",
            fg="gray"
        )

        self.file_label.pack(pady=10)


        button_frame = tk.Frame(self)
        button_frame.pack(fill="x", padx=15, pady=20)


        tk.Button(
            button_frame,
            text="Browse",
            command=self._choose_file
        ).pack(side="left")


        tk.Button(
            button_frame,
            text="Cancel",
            command=self.destroy
        ).pack(side="right")


        tk.Button(
            button_frame,
            text="Import",
            command=self._confirm
        ).pack(
            side="right",
            padx=8
        )


    def _choose_file(self):

        path = filedialog.askopenfilename(
            title="Select CSV or Excel File",
            filetypes=[
                (
                    "Spreadsheet files",
                    "*.csv *.xlsx *.xls *.tsv *.txt"
                ),
                (
                    "All files",
                    "*.*"
                )
            ]
        )

        if path:
            self.file_path = path
            self.file_label.config(
                text=path,
                fg="black"
            )


    def _confirm(self):

        if not self.file_path:
            messagebox.showwarning(
                "No File",
                "Please select a file first."
            )
            return

        try:

            columns, rows = DataManager.import_file(
                self.file_path
            )

            if not rows:
                messagebox.showwarning(
                    "No Data",
                    "No valid rows were found."
                )
                return


            self.on_import(
                columns,
                rows
            )

            self.destroy()


        except Exception as e:

            messagebox.showerror(
                "Import Failed",
                str(e)
            )



class App:

    def __init__(self, root, data_manager: DataManager):

        self.root = root
        self.dm = data_manager

        self.root.title(
            "Excel Entry Manager"
        )

        self.root.geometry(
            "1000x600"
        )


        self._editing_entry = None


        self._build_widgets()


        # Enable dropping files directly to main window
        if DND_AVAILABLE:

            self.root.drop_target_register(
                DND_FILES
            )

            self.root.dnd_bind(
                "<<Drop>>",
                self._on_drop_file
            )


        self._refresh_table()



    # ---------------- UI ----------------


    def _build_widgets(self):

        top = tk.Frame(self.root)

        top.pack(
            fill="x",
            padx=10,
            pady=10
        )


        tk.Label(
            top,
            text="Search:"
        ).pack(
            side="left"
        )


        self.search_var = tk.StringVar()


        self.search_var.trace_add(
            "write",
            lambda *_:
                self._refresh_table()
        )


        tk.Entry(
            top,
            textvariable=self.search_var
        ).pack(
            side="left",
            fill="x",
            expand=True,
            padx=10
        )


        tk.Button(
            top,
            text="Add",
            command=self._open_import_dialog
        ).pack(
            side="left",
            padx=5
        )


        tk.Button(
            top,
            text="Delete Selected",
            command=self._delete_selected
        ).pack(
            side="left"
        )



        table_frame = tk.Frame(self.root)

        table_frame.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=(0,10)
        )


        self.tree = ttk.Treeview(
            table_frame,
            show="headings",
            selectmode="extended"
        )


        self.tree.pack(
            side="left",
            fill="both",
            expand=True
        )


        scrollbar = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=self.tree.yview
        )


        scrollbar.pack(
            side="right",
            fill="y"
        )


        self.tree.configure(
            yscrollcommand=scrollbar.set
        )


        self.tree.bind(
            "<Double-1>",
            self._on_double_click
        )


        self.tree.bind(
            "<Delete>",
            lambda e:
                self._delete_selected()
        )

        # Keyboard shortcuts
        self.tree.bind(
            "<Alt-a>",
            self._select_all
        )

        self.tree.bind(
            "<Control-a>",
            self._select_all
        )

        self.tree.bind(
            "<Control-c>",
            self._copy_selected
        )

        self.tree.bind(
            "<Control-x>",
            self._cut_selected
        )

        self._set_columns(
            self.dm.columns
        )

    def _select_all(self, event=None):
        """
        Select all rows in the table.
        """

        items = self.tree.get_children()

        self.tree.selection_set(items)

        return "break"

    def _copy_selected(self, event=None):
        """
        Copy selected rows as tab-separated text.
        Compatible with Excel paste.
        """

        selected = self.tree.selection()

        if not selected:
            return "break"


        lines = []


        for item in selected:

            values = self.tree.item(
                item,
                "values"
            )

            lines.append(
                "\t".join(
                    str(v)
                    for v in values
                )
            )


        text = "\n".join(lines)


        self.root.clipboard_clear()

        self.root.clipboard_append(
            text
        )

        self.root.update()
        return "break"

    def _cut_selected(self, event=None):
        """
        Copy selected rows and remove them.
        """

        self._copy_selected()

        self._delete_selected()

        return "break"

    def _set_columns(self, columns):

        self.tree["columns"] = columns


        for col in columns:

            self.tree.heading(
                col,
                text=col
            )

            self.tree.column(
                col,
                width=160
            )



    # ---------------- Import ----------------


    def _open_import_dialog(self):

        ImportDialog(
            self.root,
            self._on_import_finished
        )


    def _on_import_finished(
        self,
        columns,
        rows
    ):

        self.dm.add_rows(
            columns,
            rows
        )

        self._refresh_table()



    def _on_drop_file(self, event):

        files = self.root.tk.splitlist(
            event.data
        )


        total = 0


        try:

            for path in files:

                columns, rows = DataManager.import_file(
                    path
                )

                result = self.dm.add_rows(
                    columns,
                    rows
                )

                total += len(rows)


            self._refresh_table()


            #messagebox.showinfo(
            #   "Import Complete",
            #    f"Imported: {result['added']} rows\n"
            #    f"Skipped duplicates: {result['skipped']}"
            #)

        except Exception as e:

            messagebox.showerror(
                "Import Failed",
                str(e)
            )



    # ---------------- Table ----------------


    def _refresh_table(self):

        if tuple(self.dm.columns) != self.tree["columns"]:

            self._set_columns(
                self.dm.columns
            )


        for item in self.tree.get_children():

            self.tree.delete(item)



        keyword = self.search_var.get().strip()


        for idx in self.dm.search(keyword):

            row = self.dm.rows[idx]

            values = [
                row.get(col, "")
                for col in self.dm.columns
            ]


            self.tree.insert(
                "",
                "end",
                iid=str(idx),
                values=values
            )



    # ---------------- Delete ----------------


    def _delete_selected(self):

        selected = self.tree.selection()


        if not selected:
            return


        if not messagebox.askyesno(
            "Confirm Delete",
            f"Delete {len(selected)} selected row(s)?"
        ):
            return


        self.dm.delete_rows(
            [
                int(i)
                for i in selected
            ]
        )


        self._refresh_table()



    # ---------------- Edit ----------------


    def _on_double_click(self,event):

        region = self.tree.identify(
            "region",
            event.x,
            event.y
        )


        if region != "cell":
            return


        row_id = self.tree.identify_row(
            event.y
        )

        col_id = self.tree.identify_column(
            event.x
        )


        if not row_id:
            return


        col_index = int(
            col_id.replace("#","")
        ) - 1


        column = self.dm.columns[col_index]


        bbox = self.tree.bbox(
            row_id,
            col_id
        )


        if not bbox:
            return


        x,y,w,h = bbox


        value = self.tree.set(
            row_id,
            column
        )


        self._destroy_editor()


        entry = tk.Entry(
            self.tree
        )


        entry.place(
            x=x,
            y=y,
            width=w,
            height=h
        )


        entry.insert(
            0,
            value
        )


        entry.focus()


        def save(_=None):

            new_value = entry.get()

            self.tree.set(
                row_id,
                column,
                new_value
            )


            self.dm.update_cell(
                int(row_id),
                column,
                new_value
            )


            self._destroy_editor()



        entry.bind(
            "<Return>",
            save
        )

        entry.bind(
            "<FocusOut>",
            save
        )


        self._editing_entry = entry



    def _destroy_editor(self):

        if self._editing_entry:

            self._editing_entry.destroy()

            self._editing_entry = None



def run_app():

    dm = DataManager()


    root = (
        TkinterDnD.Tk()
        if DND_AVAILABLE
        else tk.Tk()
    )


    App(
        root,
        dm
    )


    root.mainloop()