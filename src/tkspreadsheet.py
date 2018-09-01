import tkinter as tk
import re

from tkinter import messagebox

import src.utils as utils

from PIL import Image, ImageTk

from operator import add

import os

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

        sv = self._spreadsheet_svs[self._spreadsheet_entry_inverse[entry_widget]]

        sv.set(sv.get().strip())

        return True

    def _restore_borders(self, new, method='normal'):
        hlbg = 'darkgreen' if method == 'selected' else 'ghost white'

        if type(new) != list:
            new = [new]
        for old in self._selected_cells:
            if old not in new:
                old.config(highlightbackground = hlbg)

    def _deselect_cells(self, cells, drag=False):
        [self._deselect_cell(cell, drag=drag) for cell in cells]

    def _deselect_cell(self, cell, drag=False):
        #hlbg = 'darkgreen' if cell in self._selected_cells else 'ghost white'
        hlbg = 'ghost white'
        cell.config(highlightbackground = hlbg)
        if drag:
            self._selected_motion_cells.remove(cell)
        else:
            self._selected_cells.remove(cell)

    def _select_cells(self, entry_widgets, exclusive=True, drag=False):
        if exclusive:
            self._restore_borders(entry_widgets)
            self._selected_cells = []

        self._selected_motion_cells = []
        
        [self._select_cell(cell, exclusive=False, drag=drag) for cell in entry_widgets]

    def _set_anchor(self, cell = None):
        if not cell:
            self._anchor_cell = None
            self._restore_borders(None, method='selected')

        if type(cell) == tuple:
            cell = self.nametowidget(self._spreadsheet_entry[cell])

        self._anchor_cell = cell
        cell.config(highlightbackground = 'goldenrod')
        self._restore_borders(self._anchor_cell, method='selected')


    def _select_cell(self, entry_widget, anchor = False, exclusive=False, drag=False, flip=False):
        if exclusive:
            self._restore_borders(entry_widget)
            self._selected_cells = []

        if anchor:
            self._set_anchor(entry_widget)
        elif entry_widget != self._anchor_cell:
            entry_widget.config(highlightbackground = 'darkgreen')

        if entry_widget in self._selected_cells:
            if flip:
                print('flip')
                self._deselect_cell(entry_widget)
            else:
                return

        if drag:
            self._selected_motion_cells.append(entry_widget)
        else:
            self._selected_cells.append(entry_widget)
        
        self.god_entry.focus_set()

    def _entry_focus(self, entry_widget, highlight=True):
        entry_widget.config(state='normal', cursor='xterm')
        entry_widget.focus_set()
        if highlight:
            entry_widget.selection_range(0, 'end')
        entry_widget.icursor(tk.END)

        if entry_widget.get():
            return 'break'

    def _on_spreadsheet_click(self, event):
        entry_widget = self.nametowidget(event.widget)

        print(entry_widget)

        self._motion_anchor_cell = entry_widget

        if not self.focus_get() == event.widget:
            print(self._selected_cells)
            if [entry_widget] == self._selected_cells:
                print('here')
                return self._entry_focus(entry_widget)
            else:
                self.god_entry.focus_set()
                self._select_cell(entry_widget, anchor=True, exclusive=True, drag=False)

    def _on_spreadsheet_control_click(self, event):
        focus_widget = self.nametowidget(self.focus_get())
        entry_widget = self.nametowidget(event.widget)
        self._motion_anchor_cell = entry_widget
        self._select_cell(entry_widget, anchor=True, flip=True)

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

    def _select_range(self, anchor_widget, reel_widget, exclusive = False, drag=False, flip=False):
        anchor_coordinates = (a_row, a_column) = self._spreadsheet_entry_inverse[anchor_widget]
        reel_coordinates = (r_row, r_column) = self._spreadsheet_entry_inverse[reel_widget]
        
        cells = []

        row_range = self._closed_range(a_row, r_row)
        column_range = self._closed_range(a_column, r_column)

        if drag:
            _min_motion_row = row_range[0]
            _min_motion_column = column_range[0]
            _max_motion_row = row_range[-1]
            _max_motion_column = column_range[-1]

            if _min_motion_row > self._min_motion_row:
                self._deselect_cells([self._spreadsheet_entry[(self._min_motion_row, column)] for column in column_range], drag=True)

            elif _max_motion_row < self._max_motion_row:
                self._deselect_cells([self._spreadsheet_entry[(self._max_motion_row, column)] for column in column_range], drag=True)

            
            if _min_motion_column > self._min_motion_column:
                self._deselect_cells([self._spreadsheet_entry[(row, self._min_motion_column)] for row in row_range], drag=True)

            elif _max_motion_column < self._max_motion_column:
                self._deselect_cells([self._spreadsheet_entry[(row, self._max_motion_column)] for row in row_range], drag=True)



            self._min_motion_row = _min_motion_row
            self._min_motion_column = _min_motion_column
            self._max_motion_row = _max_motion_row
            self._max_motion_column = _max_motion_column

            print(_min_motion_row, _min_motion_column, _max_motion_row, _max_motion_column)


        for x in row_range:
            for y in column_range:
                cells.append(self._spreadsheet_entry[(x, y)])

        self._select_cells(cells, exclusive=exclusive, drag=drag)


    def _on_spreadsheet_mouse_motion(self, event):
        motion_reel = self.winfo_containing(event.x_root, event.y_root)
        if motion_reel.config().get('state') != 'normal':
            self._select_range(self._motion_anchor_cell, motion_reel, drag=True)

    def _on_spreadsheet_backspace(self, event):
        self._erase_selected_cell_contents()

    def _on_spreadsheet_delete(self, event):
        self._erase_selected_cell_contents()

    def _erase_selected_cell_contents(self):
        for e in self._selected_cells:
            self._erase_cell_contents(e)

    def _erase_cell_contents(self, entry_widget):
        sv = self._spreadsheet_svs[self._spreadsheet_entry_inverse[entry_widget]]
        sv.set('')

    def _on_spreadsheet_typing_left(self, event):
        entry_widget = self.nametowidget(event.widget)
        if entry_widget.selection_present():
            entry_widget.icursor(entry_widget.index(tk.ANCHOR) + 1)
            

    def _on_spreadsheet_typing_right(self, event = None):
        entry_widget = self.nametowidget(event.widget)
        if entry_widget.selection_present():
            entry_widget.icursor(entry_widget.index(tk.ANCHOR) - 1)

    def _on_spreadsheet_mouse_release(self, event):
        self._selected_cells.extend(self._selected_motion_cells)
        self._selected_motion_cells = []

        self._min_motion_row = self.spreadsheet_columns
        self._max_motion_row = -1
        self._min_motion_column = self.spreadsheet_rows
        self._max_motion_column = -1

        self._reel_cell = self.nametowidget(event.widget)

    def _on_spreadsheet_control_backspace(self, event):
        if self.focus_get() == event.widget:
            entry_widget = self.nametowidget(event.widget)
            self._erase_cell_contents(entry_widget)
            entry_widget.icursor(0)
            print(entry_widget.focus_get())
            self._guarantee_widget_focus = entry_widget
            print(self._guarantee_widget_focus)
            return 'break'

    def _on_entry_focus_out(self, event=None):
        print('here!')
        if self._guarantee_widget_focus:
            self.nametowidget(event.widget).focus_set()

        self._guarantee_widget_focus = None

    def _on_column_label_click(self, event, exclusive = True):
        column = self._column_labels.index(self.nametowidget(event.widget))
        self._select_column(column, exclusive = exclusive)

    def _on_column_label_control_click(self, event):
        self._on_column_label_click(event, exclusive = False)

    def _on_column_label_shift_click(self, event):
        (_, anchor_col) = self._get_anchor_coords()
        event_col = self._column_labels.index(self.nametowidget(event.widget))
        for col in self._closed_range(anchor_col, event_col):
            self._select_column(col, exclusive = False, anchor=False)
        self._set_anchor(self._spreadsheet_entry[(0, anchor_col)])

    def _select_column(self, column, exclusive = True, anchor = True):
        self._select_cells([self._spreadsheet_entry[(row, column)] for row in range(self.spreadsheet_rows)], exclusive = exclusive)
        if anchor:
            self._set_anchor((0, column))

    def _on_row_label_click(self, event, exclusive = True):
        row = self._row_labels.index(self.nametowidget(event.widget))
        self._select_row(row, exclusive = exclusive)

    def _on_row_label_control_click(self, event):
        self._on_row_label_click(event, exclusive = False)

    def _select_row(self, row, exclusive = True):
        self._select_cells([self._spreadsheet_entry[(row, column)] for column in range(self.spreadsheet_columns)], exclusive = exclusive)
        self._set_anchor((row, 0))

    def _select_cell_indices(self, indices):
        self._select_cells([self._spreadsheet_entry[index] for index in indices])

    def _select_all(self, event):
        for column in range(self.spreadsheet_columns):
            self._select_column(column, exclusive = False)

        self._set_anchor((0, 0))


    def __init__(self, program_paths, parent, rows, columns):
        tk.Frame.__init__(self, parent)

        self.program_paths = program_paths

        self.containing_frame = parent

        self._spreadsheet_entry = {}
        self._spreadsheet_entry_inverse = {}
        self._spreadsheet_svs = {}
        self.spreadsheet_rows = rows
        self.spreadsheet_columns = columns

        self._selected_cells = []
        self._anchor_cell = None
        self._reel_cell = None

        self._column_labels = []
        self._row_labels = []


        self._motion_anchor_cell = None

        self._selected_motion_cells = []

        self._min_motion_row = self.spreadsheet_columns
        self._max_motion_row = -1
        self._min_motion_column = self.spreadsheet_rows
        self._max_motion_column = -1

        self._guarantee_widget_focus = None

        # register a command to use for validation
        vcmd = (self.register(self._on_spreadsheet_cell_exit), '%W', '%P')

        self.gsv = tk.StringVar()
        self.gsv.trace_add('write', lambda idc, idc2, idc3: self._go_to_entry_widget())
        self.god_entry = tk.Entry(self, textvariable = self.gsv)
        self.god_entry.grid(row=1, column=1)

        im = Image.open(os.path.join(self.program_paths['icons'], 'select_all_cropped.gif')).resize((10, 10), Image.ANTIALIAS)
        ph = ImageTk.PhotoImage(im)

        image_label = tk.Label(self, image=ph)
        image_label.grid(row=0, column=0, sticky='se')
        image_label.image=ph  #need to keep the reference of your image to avoid garbage collection

        image_label.bind('<Button-1>', self._select_all)
        image_label.config(cursor='lr_angle')

        for column in range(self.spreadsheet_columns):
            l = tk.Label(self, text=utils.get_column_letters(column))
            l.grid(row=0, column=column+1, sticky='ew')
            l.config(cursor='sb_down_arrow')
            l.bind('<Button-1>', self._on_column_label_click)
            l.bind('<Control-Button-1>', self._on_column_label_control_click)
            l.bind('<Shift-Button-1>', self._on_column_label_shift_click)
            self._column_labels.append(l)

        for row in range(self.spreadsheet_rows):
            l = tk.Label(self, text='  ' + str(row + 1) + '  ')
            l.grid(row=row + 1, column=0)
            l.config(cursor='sb_right_arrow')
            l.bind('<Button-1>', self._on_row_label_click)
            l.bind('<Control-Button-1>', self._on_row_label_control_click)
            self._row_labels.append(l)

        # create the table of widgets
        for row in range(self.spreadsheet_rows):
            for column in range(self.spreadsheet_columns):
                index = (row, column)
                sv = tk.StringVar()
                e = tk.Entry(self, textvariable = sv, validate="focusout", validatecommand=vcmd)
                sv.trace('w', lambda idc1, idc2, idc3, e=e: self._update_all_selected_entries_typing(e))
                self._spreadsheet_svs[index] = sv
                e.grid(row=row+1, column=column+1, stick="nsew")
                e.config(justify="left", state='disabled', cursor='plus', highlightthickness = 1, highlightbackground = 'ghost white',
                            disabledbackground='white', highlightcolor = 'goldenrod')
                e.bind("<Button-1>", self._on_spreadsheet_click)
                e.bind("<Control-Button-1>", self._on_spreadsheet_control_click)
                e.bind("<Shift-Button-1>", self._on_spreadsheet_shift_click)
                e.bind("<Control-Shift-Button-1>", self._on_spreadsheet_control_shift_click)
                e.bind("<Left>", self._on_spreadsheet_typing_left)
                e.bind('<Right>', self._on_spreadsheet_typing_right)
                e.bind("<B1-Motion>", self._on_spreadsheet_mouse_motion)
                e.bind('<ButtonRelease-1>', self._on_spreadsheet_mouse_release)
                e.bind('<Control-BackSpace>', self._on_spreadsheet_control_backspace)
                e.bind('<Tab>', self._next)
                e.bind('<Shift-Tab>', self._prev)
                e.bind('<Return>', self._on_spreadsheet_down)
                e.bind('<Down>', self._on_spreadsheet_down)
                e.bind('<Up>', self._on_spreadsheet_up)
                e.bind('<Escape>', self._on_spreadsheet_typing_escape)
                self._spreadsheet_entry[index] = e
                self._spreadsheet_entry_inverse[e] = index

        # adjust column weights so they all expand equally
        for column in range(self.spreadsheet_columns):
            self.grid_columnconfigure(column, weight=1)
        # designate a final, empty row to fill up any extra space
        self.grid_rowconfigure(rows, weight=1)

        self.god_entry.bind('<BackSpace>', self._on_spreadsheet_backspace)
        self.god_entry.bind('<Delete>', self._on_spreadsheet_delete)
        self.god_entry.bind('<Up>', lambda event: self._on_spreadsheet_up(event, exclusive=True))
        self.god_entry.bind('<Shift-Up>', self._on_spreadsheet_shift_up)
        self.god_entry.bind('<Control-Up>', lambda event: self._on_spreadsheet_up(event, exclusive=False))
        self.god_entry.bind('<Down>', lambda event: self._on_spreadsheet_down(event, exclusive=True))
        self.god_entry.bind('<Shift-Down>', self._on_spreadsheet_shift_down)
        self.god_entry.bind('<Control-Down>', lambda event: self._on_spreadsheet_down(event, exclusive=False))
        self.god_entry.bind('<Left>', lambda event: self._on_spreadsheet_left(event, exclusive=True))
        self.god_entry.bind('<Shift-Left>', self._on_spreadsheet_shift_left)
        self.god_entry.bind('<Control-Left>', lambda event: self._on_spreadsheet_left(event, exclusive=False))
        self.god_entry.bind('<Right>', lambda event: self._on_spreadsheet_right(event, exclusive=True))
        self.god_entry.bind('<Shift-Right>', self._on_spreadsheet_shift_right)
        self.god_entry.bind('<Control-Right>', lambda event: self._on_spreadsheet_right(event, exclusive=False))
        self.god_entry.bind('<Tab>', self._on_spreadsheet_tab)
        self.god_entry.bind('<Shift-Tab>', self._on_spreadsheet_shift_tab)
        self.god_entry.bind('<Control-Tab>', self._on_spreadsheet_shift_tab)
        self.god_entry.bind('<FocusOut>', self._on_god_entry_focus_out)
        self.god_entry.bind('<FocusIn>', self._on_entry_focus_out)
        self.god_entry.bind('<Return>', self._on_spreadsheet_enter_key)
        self.god_entry.bind('<Control-Key-d>', self._on_spreadsheet_control_d)
        self.god_entry.bind('<Control-Key-a>', self._select_all)
        self.god_entry.bind('<Control-Key-D>', self._calculate_determinant)
        self.god_entry.bind('<Control-Key-I>', self._convert_to_inverse)
        self.god_entry.bind('<Control-Key-i>', self._print_inverse)

        self._guarantee_focus = False

        self.god_entry.focus_set()

    def _is_square_selection(self, event):
        pass

    # Assume 3x3 from (0, 0)
    def _calculate_determinant(self, event):
        a00 = self._spreadsheet_entry[(0, 0)].get().strip()
        a01 = self._spreadsheet_entry[(0, 1)].get().strip()
        a02 = self._spreadsheet_entry[(0, 2)].get().strip()
        a10 = self._spreadsheet_entry[(1, 0)].get().strip()
        a11 = self._spreadsheet_entry[(1, 1)].get().strip()
        a12 = self._spreadsheet_entry[(1, 2)].get().strip()
        a20 = self._spreadsheet_entry[(2, 0)].get().strip()
        a21 = self._spreadsheet_entry[(2, 1)].get().strip()
        a22 = self._spreadsheet_entry[(2, 2)].get().strip()
        
        try:
            d1 = a11 * a22 - a21 * a12
            print('1')
            d2 = a10 * a22 - a20 * a22
            print('2')
            d3 = a10 * a21 - a20 * a11
            print('3')
            d = a00 * d1 - a01 * d2 + a02 * d3
            print('d')
            print(d)
        except TypeError:
            messagebox.showinfo('Try again', 'Not all entries are integers')

    def _convert_to_inverse(self, event):
        pass

    def _print_inverse(self, event):
        pass

    def _on_spreadsheet_control_d(self, event):
        self._update_all_selected_entries_typing()

    def _on_spreadsheet_typing_escape(self, event):
        # TBD, keep old entry somewhere
        pass

    def _on_spreadsheet_enter_key(self, event):
        return self._entry_focus(self._anchor_cell, highlight=False)

    def _on_god_entry_focus_out(self, event):
        if self._guarantee_focus:
            self.god_entry.focus_set()

        self._guarantee_focus = False

    def _get_anchor_sv(self):
        return self._spreadsheet_svs[self._spreadsheet_entry_inverse[self._anchor_cell]]

    def _go_to_entry_widget(self):
        asv = self._get_anchor_sv()
        asv.set(self.gsv.get())
        self.gsv.set('')
        return self._entry_focus(self._anchor_cell, highlight=False)

    def _update_all_selected_entries_typing(self, entry = None):
        if entry == self._anchor_cell:
            for entry_widget in self._selected_cells:
                print(self._spreadsheet_entry_inverse[entry_widget])
                if entry_widget == self._anchor_cell:
                    continue
                sv = self._spreadsheet_svs[self._spreadsheet_entry_inverse[entry_widget]]
                sv.set(self._anchor_cell.get())
                entry_widget.update()

    def _on_spreadsheet_tab(self, event = None):
        self._next()

    def _on_spreadsheet_shift_tab(self, event=None):
        self._prev()

    def _prev(self, event = None):
        if len(self._selected_cells) == 1:
            self._on_spreadsheet_left()
        elif len(self._selected_cells) > 1:
            pass
        else:
            pass

        self._guarantee_focus = True
        
        # Not effective for some reason
        # return 'break'

    def _next(self, event = None):
        if len(self._selected_cells) == 1:
            self._on_spreadsheet_right()
        elif len(self._selected_cells) > 1:
            pass
        else:
            pass

        self._guarantee_focus = True

        # Not effective for some reason
        # return 'break'

    def _get_anchor_coords(self):
        anchor_widget = self.nametowidget(self._anchor_cell)
        return self._spreadsheet_entry_inverse[anchor_widget]

    def _get_reel_coords(self):
        reel_widget = self.nametowidget(self._reel_cell)
        return self._spreadsheet_entry_inverse[reel_widget]        

    def _on_spreadsheet_shift_dir(self, event=None):
        self._select_range(self._anchor_cell, self._reel_cell)

    def _on_spreadsheet_shift_up(self, event=None):
        reel_coords = reel_row, reel_col = self._get_reel_coords()
        offset = (0, 0)
        if reel_row > 0:
            offset = (-1, 0)
        self._reel_cell = self._spreadsheet_entry[tuple(map(add, reel_coords, offset))]
        self._select_range(self._anchor_cell, self._reel_cell, exclusive=True)


    def _on_spreadsheet_shift_down(self, event=None):
        pass

    def _on_spreadsheet_shift_left(self, event=None):
        pass

    def _on_spreadsheet_shift_right(self, event=None):
        pass

    def _on_spreadsheet_up(self, event = None, exclusive = True):
        anchor_row, anchor_col = self._get_anchor_coords()
        if anchor_row > 0:
            self._select_cell(self._spreadsheet_entry[(anchor_row - 1, anchor_col)], anchor=True, exclusive=exclusive, drag=False)
        return 'break'

    def _on_spreadsheet_down(self, event = None, exclusive = True):
        anchor_row, anchor_col = self._get_anchor_coords()
        if anchor_row < self.spreadsheet_rows - 1:
            self._select_cell(self._spreadsheet_entry[(anchor_row + 1, anchor_col)], anchor=True, exclusive=exclusive, drag=False)
        return 'break'

    def _on_spreadsheet_left(self, event = None, exclusive = True):
        anchor_row, anchor_col = self._get_anchor_coords()
        if anchor_col > 0:
            self._select_cell(self._spreadsheet_entry[(anchor_row, anchor_col - 1)], anchor=True, exclusive=exclusive, drag=False)
        elif anchor_row > 0:
            self._select_cell(self._spreadsheet_entry[(anchor_row - 1, self.spreadsheet_columns - 1)], anchor=True, exclusive=exclusive, drag=False)

    def _on_spreadsheet_right(self, event = None, exclusive = True):
        anchor_row, anchor_col = self._get_anchor_coords()
        if anchor_col < self.spreadsheet_columns - 1:
            self._select_cell(self._spreadsheet_entry[(anchor_row, anchor_col + 1)], anchor=True, exclusive=exclusive, drag=False)
        elif anchor_row < self.spreadsheet_rows - 1:
            self._select_cell(self._spreadsheet_entry[(anchor_row + 1, 0)], anchor=True, exclusive=exclusive, drag=False)


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
    def __init__(self, parent, program_paths):
        tk.Frame.__init__(self, parent)
        self.table = Spreadsheet(program_paths, self, 11, 11)
        self.submit = tk.Button(self, text="Submit", command=self.on_submit)
        self.table.pack(side="top", fill="both", expand=True)
        self.submit.pack(side="bottom")

    def on_submit(self):
        print(self.table.get())