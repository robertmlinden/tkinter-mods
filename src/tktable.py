import tkinter as tk
import re

class Spreadsheet(tk.Frame):
    def _spreadsheet_callback(self, var_name):
        var_number = int(re.sub('[^0-9]', '', var_name))
        row = var_number // self.spreadsheet_columns
        column = var_number % self.spreadsheet_columns
        index = (row, column)
        print(index, '!!')
        print(self._spreadsheet_entry[index])
        print(self._spreadsheet_svs[index].get())

    def _on_spreadsheet_cell_exit(self, e, v):
        entry_widget = self.nametowidget(e)
        try:
            float(v)
            entry_widget.config(justify='right')
        except ValueError:
            entry_widget.config(justify='left')

        entry_widget.config(state='disabled', cursor='plus')

        return True

    def fullname(self, o):
        return o.__module__ + "." + o.__class__.__qualname__

    def _restore_borders(self, new):
        for old in self._selected_cells:
            if old not in new:
                old.config(highlightbackground = 'ghost white')

    def _select_cell(self, entry_widget):
        entry_widget.config(highlightbackground = 'darkgreen')
        self._selected_cells.append(entry_widget)

    def _on_spreadsheet_click(self, event):
        entry_widget = self.nametowidget(event.widget)

        if not self.focus_get() == event.widget:
            if [entry_widget] == self._selected_cells:
                entry_widget.config(state='normal', cursor='xterm')
                entry_widget.focus_set()
                entry_widget.selection_range(0, 'end')
                entry_widget.icursor(tk.END)

                if entry_widget.get():
                    return 'break'
            else:
                self.containing_frame.focus_set()
                self._restore_borders([entry_widget])
                self._anchor_cell = self._spreadsheet_entry_inverse[entry_widget]
                self._select_cell(entry_widget)

    def _on_spreadsheet_control_click(self, event):
        focus_widget = self.nametowidget(self.focus_get())
        entry_widget = self.nametowidget(event.widget)
        self._select_cell(entry_widget)
        if isinstance(focus_widget, tk.Entry):
            self.containing_frame.focus_set()
            self._select_cell(focus_widget)
        # Let's do this!!

    def _on_spreadsheet_left(self, event):
        entry_widget = self.nametowidget(event.widget)
        if entry_widget.selection_present():
            entry_widget.icursor(entry_widget.index(tk.ANCHOR) + 1)
            

    def _on_spreadsheet_right(self, event):
        entry_widget = self.nametowidget(event.widget)
        if entry_widget.selection_present():
            entry_widget.icursor(entry_widget.index(tk.ANCHOR) - 1)

    def _on_spreadsheet_mouse_motion(self, event):
        # Feature in progress
        entry = self.winfo_containing(event.x_root, event.y_root)
        if entry.config()['state'] == 'disabled':
            # Do disabled things here
            pass
        print(entry)

    def __init__(self, parent, rows, columns):
        tk.Frame.__init__(self, parent)

        self.containing_frame = parent

        self._spreadsheet_entry = {}
        self._spreadsheet_entry_inverse = {}
        self._spreadsheet_svs = {}
        self.spreadsheet_rows = rows
        self.spreadsheet_columns = columns

        self._selected_cells = []
        self._anchor_cell = None

        # register a command to use for validation
        vcmd = (self.register(self._on_spreadsheet_cell_exit), '%W', '%P')

        # create the table of widgets
        for row in range(self.spreadsheet_rows):
            for column in range(self.spreadsheet_columns):
                index = (row, column)
                self._spreadsheet_svs[index] = tk.StringVar()
                self._spreadsheet_svs[index].trace_add('write', lambda var_name, mode, _: self._spreadsheet_callback(var_name))#self._spreadsheet_callback(self._spreadsheet_svs[index]))
                print(self._spreadsheet_svs[index])
                print(self)
                print(vcmd)
                e = tk.Entry(self, textvariable=self._spreadsheet_svs[index], validate="focusout", validatecommand=vcmd)
                e.grid(row=row, column=column, stick="nsew")
                e.config(justify="left", state='disabled', cursor='plus', highlightthickness = 1, highlightbackground = 'ghost white')
                e.bind("<Button-1>", self._on_spreadsheet_click)
                e.bind("<Control-Button-1>", self._on_spreadsheet_control_click)
                e.bind("<Left>", self._on_spreadsheet_left)
                e.bind('<Right>', self._on_spreadsheet_right)
                e.bind("<B1-Motion>", self._on_spreadsheet_mouse_motion)
                self._spreadsheet_entry[index] = e
                self._spreadsheet_entry_inverse[e] = index

        # adjust column weights so they all expand equally
        for column in range(self.spreadsheet_columns):
            self.grid_columnconfigure(column, weight=1)
        # designate a final, empty row to fill up any extra space
        self.grid_rowconfigure(rows, weight=1)

    def get(self):
        '''Return a list of lists, containing the data in the table'''
        result = []
        for row in range(self.spreadsheet_rows):
            current_row = []
            for column in range(self.spreadsheet_columns):
                index = (row, column)
                current_row.append(self._spreadsheet_entry[index].get())
            result.append(current_row)
        return result

class Example(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent)
        self.table = Spreadsheet(self, 3, 4)
        self.submit = tk.Button(self, text="Submit", command=self.on_submit)
        self.table.pack(side="top", fill="both", expand=True)
        self.submit.pack(side="bottom")

    def on_submit(self):
        print(self.table.get())