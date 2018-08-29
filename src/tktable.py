import tkinter as tk
import re

class Spreadsheet(tk.Frame):
    def _on_entry_keystroke(self, sv):
        print(sv.get())

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

    def _deselect_all_but(self, entries, anchor=False):
        if type(entries) != list:
            entries = [entries]
        print(self._selected_cells)
        self._restore_borders(entries)
        self._selected_cells = []
        for cell in entries:
            self._select_cell(cell)

        if anchor:
            self._anchor_cell = entries[0]

    def _deselect_cell(self, cell):
        self._selected_cells.remove(cell)
        cell.config(highlightbackground = 'ghost white')

    def _select_cells(self, entry_widgets, flip=False):
        for cell in entry_widgets:
            self._select_cell(cell, flip)

    def _select_cell(self, entry_widget, anchor = False, flip=False):
        entry_widget.config(highlightbackground = 'darkgreen')
        if entry_widget in self._selected_cells:
            if flip:
                self._deselect_cell(entry_widget)
        else:
            self._selected_cells.append(entry_widget)
        
        if anchor:
            self._anchor_cell = entry_widget

    def _on_spreadsheet_click(self, event):
        entry_widget = self.nametowidget(event.widget)

        self._motion_anchor_cell = entry_widget
        print(self._motion_anchor_cell)

        if not self.focus_get() == event.widget:
            if [entry_widget] == self._selected_cells:
                entry_widget.config(state='normal', cursor='xterm')
                entry_widget.focus_set()
                entry_widget.selection_range(0, 'end')
                entry_widget.icursor(tk.END)

                if entry_widget.get():
                    return 'break'
            else:
                self.god_entry.focus_set()
                self._deselect_all_but([entry_widget], anchor=True)

    def _on_spreadsheet_control_click(self, event):
        focus_widget = self.nametowidget(self.focus_get())
        entry_widget = self.nametowidget(event.widget)
        self._motion_anchor_cell = entry_widget
        self._select_cell(entry_widget, anchor=True)
        if isinstance(focus_widget, tk.Entry):
            self.containing_frame.focus_set()
            self._select_cell(focus_widget, flip=True)

    def _closed_range(self, a, b):
        start = None
        stop = None
        
        if a < b:
            start = a
            stop = b
        else:
            start = b
            stop = a

        return range(start, stop + 1)

    def _on_spreadsheet_shift_click(self, event):
        self._select_range(self._anchor_cell, self.nametowidget(event.widget), exclusive = True)

    def _on_spreadsheet_control_shift_click(self, event):
        self._select_range(self._anchor_cell, self.nametowidget(event.widget))

    def _select_range(self, anchor_widget, reel_widget, exclusive = False):
        anchor_coordinates = (ax, ay) = self._spreadsheet_entry_inverse[anchor_widget]
        reel_coordinates = (rx, ry) = self._spreadsheet_entry_inverse[reel_widget]
        
        cells = []

        for x in self._closed_range(ax, rx):
            for y in self._closed_range(ay, ry):
                cells.append(self._spreadsheet_entry[(x, y)])

        if exclusive:
            self._deselect_all_but(cells)
        else:
            self._select_cells(cells)

        print(self._selected_cells)

    def _on_spreadsheet_mouse_motion(self, event):
        motion_reel = self.winfo_containing(event.x_root, event.y_root)
        if motion_reel.config().get('state') != 'normal':
            print(motion_reel)
            self._select_range(self._motion_anchor_cell, motion_reel)

    def _on_spreadsheet_backspace(self, event):
        self.erase_selected_cell_contents()

    def _on_spreadsheet_delete(self, event):
        self.erase_selected_cell_contents()

    def erase_selected_cell_contents(self):
        print(self._selected_cells)
        for e in self._selected_cells:
            print(e)
            print(e.get())
            e.delete(0, tk.END)

    def _on_spreadsheet_left(self, event):
        entry_widget = self.nametowidget(event.widget)
        if entry_widget.selection_present():
            entry_widget.icursor(entry_widget.index(tk.ANCHOR) + 1)
            

    def _on_spreadsheet_right(self, event):
        entry_widget = self.nametowidget(event.widget)
        if entry_widget.selection_present():
            entry_widget.icursor(entry_widget.index(tk.ANCHOR) - 1)


    def __init__(self, parent, rows, columns):
        tk.Frame.__init__(self, parent)

        self.containing_frame = parent

        self._spreadsheet_entry = {}
        self._spreadsheet_entry_inverse = {}
        self.spreadsheet_rows = rows
        self.spreadsheet_columns = columns

        self._selected_cells = []
        self._anchor_cell = None

        self._motion_anchor_cell = None

        # register a command to use for validation
        vcmd = (self.register(self._on_spreadsheet_cell_exit), '%W', '%P')

        sv = tk.StringVar()
        sv.trace_add('write', lambda idc, idc2, idc3, sv=sv: self._on_entry_keystroke(sv))
        self.god_entry = tk.Entry(self, textvariable = sv)
        self.god_entry.grid(row=0, column=0)
        print(self.god_entry)

        # create the table of widgets
        for row in range(self.spreadsheet_rows):
            for column in range(self.spreadsheet_columns):
                index = (row, column)
                e = tk.Entry(self, validate="focusout", validatecommand=vcmd)
                e.grid(row=row, column=column, stick="nsew")
                e.config(justify="left", state='disabled', cursor='plus', highlightthickness = 1, highlightbackground = 'ghost white')
                e.bind("<Button-1>", self._on_spreadsheet_click)
                e.bind("<Control-Button-1>", self._on_spreadsheet_control_click)
                e.bind("<Shift-Button-1>", self._on_spreadsheet_shift_click)
                e.bind("<Control-Shift-Button-1>", self._on_spreadsheet_control_shift_click)
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

        self.god_entry.bind('<BackSpace>', self._on_spreadsheet_backspace)
        self.god_entry.bind('<Delete>', self._on_spreadsheet_delete)
        self.god_entry.bind('<Up>', self._on_spreadsheet_up)

        self.god_entry.focus_set()

    def _on_spreadsheet_up(self, event):
        anchor_widget = self.nametowidget(self._anchor_cell)
        anchor_x, anchor_y = self._spreadsheet_entry_inverse[anchor_widget][1]
        if anchor_y > 0:
            self._deselect_all_but(self._spreadsheet_entry[(anchor_x, anchor_y - 1)], anchor=True)


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
        self.table = Spreadsheet(self, 4, 5)
        self.submit = tk.Button(self, text="Submit", command=self.on_submit)
        self.table.pack(side="top", fill="both", expand=True)
        self.submit.pack(side="bottom")

    def on_submit(self):
        print(self.table.get())