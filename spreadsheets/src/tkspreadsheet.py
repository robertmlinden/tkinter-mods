import tkinter as tk
import re

from tkinter import messagebox, filedialog

import src.utils as utils
import src.expr_evaluator as arithmetic_evaluator

from PIL import Image, ImageTk

from operator import add

import os

import csv

class Cell(tk.Entry):
    def __init__(self, *args, **kw):
        tk.Entry.__init__(self, *args, **{key:kw[key] for key in kw if key!='cell_index'})
        self.sv = kw['textvariable']
        self.cell_index = kw['cell_index']

    @property
    def display_value(self):
        "I am the 'x' property."
        return self.sv.get()

    @display_value.setter
    def display_value(self, value):
        self.sv.set(value)

class Spreadsheet(tk.Frame):
    def _on_entry_keystroke(self, sv):
        print(sv.get())

    def _get_cell_value(self, cell_index, stringify):
        column_letter = re.search(r'[a-zA-Z]+', cell_index)[0]

        column = utils.get_column_index(column_letter)
        row = int(re.search(r'[0-9]+', cell_index)[0]) - 1

        cell = self._cells[(row, column)]

        value =  cell.get()

        if stringify:
            value = "'" + value + "'"

        return value

    def _cell_convert(self, value, stringify=False):
        return re.sub(r'\[.*?\]', lambda match: self._get_cell_value(match[0], stringify), value)

    def _process_formula(self, formula, number_based = True):
        if formula and formula[0] == '=' and len(formula) > 1:
            converted_value = self._cell_convert(formula[1:], not number_based)
            if number_based:
                return str(arithmetic_evaluator.evaluate_expression(converted_value))#value[1:]))
            else:
                return eval(converted_value)
        else:
            return formula

    def _on_spreadsheet_cell_exit(self, c, v):
        print('leaving cell!')

        cell = self.nametowidget(c)

        value = cell.display_value.strip()

        try:
            value = self._process_formula(value)
        except TypeError:
            value = self._process_formula(value, number_based = False)
        try:
            float(value)
            cell.config(justify='right')
        except ValueError:
            cell.config(justify='left')

        cell.config(state='disabled', cursor='plus')

        cell.display_value = value

        return True

    def _restore_borders(self, new, method='normal'):
        print('restore borders')
        hlbg = 'darkgreen' if method == 'selected' else 'ghost white'

        if type(new) != list:
            new = [new]
        for old in self._solidified_cells:
            if old not in new:
                old.config(highlightbackground = hlbg)

    def _deselect_all(self, solidify=False):
        print('deselect all')

        cells = self._solidified_cells if solidify else self._selected_cells

        self._deselect_cells(cells, solidify=solidify)

    def _deselect_cells(self, cells, solidify=False):
        print('Deselecting cells ' + str([cell.cell_index for cell in cells]))
        [self._deselect_cell(cell, solidify=solidify) for cell in cells]

    def _deselect_cell(self, cell, solidify=False):
        #print('deselect cell : ' + str(cell))
        try:
            cells = self._solidified_cells if solidify else self._selected_cells
            cells.remove(cell)
            cell.config(highlightbackground = 'ghost white')
            if cell == self._anchor_cell:
                anchor = self._selected_cells[-1] if self._selected_cells else self._solidified_cells[-1] if self._solidified_cells else None
                self._set_anchor(anchor)
        except ValueError:
            pass

    def _select_all(self, event):
        print('Selecting all cells in the grid!')
        for column in range(self.spreadsheet_columns):
            self._select_column(column, exclusive = False, solidify=False)

        self._set_anchor((0, 0))

    def _select_cells(self, entry_widgets, exclusive=True, solidify=False):
        print('Selecting cells ' + str([cell.cell_index for cell in entry_widgets]))
        if exclusive:
            self._restore_borders(entry_widgets)
            self._solidified_cells = []
        
        [self._select_cell(cell, exclusive=False, solidify=solidify) for cell in entry_widgets]

    def _select_cell(self, entry_widget, anchor=False, exclusive=False, solidify=False, flip=False):
        #print(('Solidify-' if solidify else 'Soft-') + 'Selecting cell ' + str(entry_widget))

        cells = self._solidified_cells if solidify else self._selected_cells

        if exclusive:
                     

        if entry_widget in cells:
            if flip:
                print('flip')
                self._deselect_cell(entry_widget)
            else:
                return
        elif anchor:
            self._set_anchor(entry_widget)
        else:
            entry_widget.config(highlightbackground = 'darkgreen')

        cells.append(entry_widget)

        self.god_entry.focus_set()

        #print(('self._solidified_cells' if solidify else 'self._selected_cells') + ' now is ' + str(cells))
        #print(('self._solidified_cells' if not solidify else 'self._selected_cells') + ' now is ' + 
        #            str(self._solidified_cells) if not solidify else str(self._selected_cells))

    def _set_reel(self, cell):
        self._reel_cell, self._prev_reel_cell = cell, self._reel_cell

    def _set_anchor(self, cell = None):
        print('Setting cell ' + str(cell) + ' to anchor')
        
        if not cell:
            self._anchor_cell = None
            #self._restore_borders(None, method='selected')
            return

        if type(cell) == tuple:
            cell = self.nametowidget(self._cells[cell])

        self._anchor_cell = cell
        cell.config(highlightbackground = 'goldenrod')
        self._restore_borders(self._anchor_cell, method='selected')

    def _solidify(self):
        self._solidified_cells.extend(self._selected_cells)

    def _export_to_csv(self, event=None):
        values = self.get()

        filename = filedialog.asksaveasfilename(title='Export to CSV', initialdir=self.program_paths['index'], filetypes=[('Comma Separated Values', '*.csv')])
        
        if not filename:
            return

        if filename[-4:].lower() != '.csv':
            filename += '.csv'

        with open(filename, "w+", newline='') as f:
            writer = csv.writer(f, skipinitialspace=True)
            writer.writerows(values) 

        os.startfile(filename)


    def _entry_focus(self, entry_widget, highlight=True):
        print('Focus on entry: ' + str(entry_widget))
        if not entry_widget:
            return

        entry_widget.config(state='normal', cursor='xterm')
        entry_widget.focus_set()
        if highlight:
            entry_widget.selection_range(0, 'end')
        entry_widget.icursor(tk.END)

        if entry_widget.get():
            return 'break'

    def _on_spreadsheet_click(self, event):
        print(event.type + ': <Button-1>')
        entry_widget = self.nametowidget(event.widget)

        (self._min_motion_row, self._min_motion_column) = \
        (self._max_motion_row, self._max_motion_column) = self._cells_inverse[entry_widget]

        if not self.focus_get() == event.widget:
            if [entry_widget] == self._solidified_cells:
                return self._entry_focus(entry_widget)
            else:
                self._select_cell(entry_widget, anchor=True, exclusive=True, solidify=False)
                self.god_entry.focus_set()

    def _on_spreadsheet_control_click(self, event):
        print(event.type + ': <Control-Button-1>')
        focus_widget = self.nametowidget(self.focus_get())
        entry_widget = self.nametowidget(event.widget)
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
        print(event.type + ': <Shift-Button-1>')
        self._set_reel(self.nametowidget(event.widget))
        self._deselect_all(solidify=True)
        self._select_range(exclusive = True)

    def _on_spreadsheet_control_shift_click(self, event):
        print(event.type + ': <Control-Shift-Button-1>')
        self._set_reel(self.nametowidget(event.widget))
        self._select_range()

    def _select_range(self, exclusive = False, flip=False):
        print('Selecting range: ', end='')
        print(self._anchor_cell.cell_index + ' to ' + self._reel_cell.cell_index)

        anchor_coordinates = (a_row, a_column) = self._cells_inverse[self._anchor_cell]
        prev_reel_coordinates = (p_row, p_column) = self._cells_inverse[self._prev_reel_cell]
        reel_coordinates = (r_row, r_column) = self._cells_inverse[self._reel_cell]
        
        cells = []

        row_range = self._closed_range(a_row, r_row)[::-1]
        prev_row_range = self._closed_range(a_row, p_row)[::-1]

        column_range = self._closed_range(a_column, r_column)[::-1]
        prev_column_range = self._closed_range(a_column, p_column)[::-1]
    
        _min_motion_row = row_range[-1]
        _min_motion_column = column_range[-1]
        _max_motion_row = row_range[0]
        _max_motion_column = column_range[0]


        if _min_motion_row > self._min_motion_row:
            print('The row minimum increased')
            for row in range(self._min_motion_row, _min_motion_row):
                self._deselect_cells([self._cells[(row, column)] for column in prev_column_range], solidify=False)

        elif _max_motion_row < self._max_motion_row:
            print('The row maximum decreased')
            for row in range(self._max_motion_row, _max_motion_row, -1):
                self._deselect_cells([self._cells[(row, column)] for column in prev_column_range], solidify=False)

        if _min_motion_row < self._min_motion_row:
            print('The row minimum decreased')
            for row in range(_min_motion_row, self._min_motion_row):
                self._select_cells([self._cells[(row, column)] for column in prev_column_range], exclusive=exclusive, solidify=False)

        elif _max_motion_row > self._max_motion_row:
            print('The row maximum increased')
            for row in range(_max_motion_row, self._max_motion_row, -1):
                self._select_cells([self._cells[(row, column)] for column in prev_column_range], exclusive=exclusive, solidify=False)

        

        
        if _min_motion_column > self._min_motion_column:
            print('The column minimum increased')
            for column in range(self._min_motion_column, _min_motion_column):
                self._deselect_cells([self._cells[(row, column)] for row in row_range], solidify=False)

        elif _max_motion_column < self._max_motion_column:
            print('The column maximum decreased')
            for column in range(self._max_motion_column, _max_motion_column, -1):
                self._deselect_cells([self._cells[(row, column)] for row in row_range], solidify=False)


        if _min_motion_column < self._min_motion_column:
            print('The column minimum decreased')
            for column in range(_min_motion_column, self._min_motion_column):
                self._select_cells([self._cells[(row, column)] for row in row_range], exclusive=exclusive, solidify=False)

        elif _max_motion_column > self._max_motion_column:
            print('The column maximum increased')
            for column in range(_max_motion_column, self._max_motion_column, -1):
                self._select_cells([self._cells[(row, column)] for row in row_range], exclusive=exclusive, solidify=False)



        self._min_motion_row = _min_motion_row
        self._min_motion_column = _min_motion_column
        self._max_motion_row = _max_motion_row
        self._max_motion_column = _max_motion_column

    def _on_spreadsheet_mouse_motion(self, event):
        self._set_reel(self.winfo_containing(event.x_root, event.y_root))
        if self._reel_cell not in self._cells.values():
            return
        if self._reel_cell.config().get('state') != 'normal':
            self._select_range(exclusive=False)

    def _on_spreadsheet_backspace(self, event):
        self._erase_selected_cell_contents()

    def _on_spreadsheet_delete(self, event):
        self._erase_selected_cell_contents()

    def _erase_selected_cell_contents(self):
        for e in self._solidified_cells:
            self._erase_cell_contents(e)

    def _erase_cell_contents(self, entry_widget):
        entry_widget.display_value = ''

    def _on_spreadsheet_typing_left(self, event):
        entry_widget = self.nametowidget(event.widget)
        if entry_widget.selection_present():
            entry_widget.icursor(entry_widget.index(tk.ANCHOR) + 1)
            

    def _on_spreadsheet_typing_right(self, event = None):
        entry_widget = self.nametowidget(event.widget)
        if entry_widget.selection_present():
            entry_widget.icursor(entry_widget.index(tk.ANCHOR) - 1)

    def _on_spreadsheet_mouse_release(self, event):
        self._solidified_cells.extend(self._selected_cells)
        self._selected_cells = []

        self._set_reel(self.winfo_containing(event.x_root, event.y_root))

    def _on_spreadsheet_control_backspace(self, event):
        if self.focus_get() == event.widget:
            entry_widget = self.nametowidget(event.widget)
            self._erase_cell_contents(entry_widget)
            entry_widget.icursor(0)
            self._guarantee_widget_focus = entry_widget

    def _on_entry_focus_out(self, event=None):
        print('Focus out!')
        if self._guarantee_widget_focus:
            self.nametowidget(event.widget).focus_set()

        self._guarantee_widget_focus = None

    def _on_column_label_click(self, event, exclusive = True):
        column = self._column_labels.index(self.nametowidget(event.widget))
        self._select_column(column, exclusive = exclusive)

    def _on_column_label_control_click(self, event):
        self._on_column_label_click(event, exclusive = False)

    def _on_column_label_shift_click(self, event, exclusive=True):
        (_, anchor_col) = self._get_anchor_coords()
        event_col = self._column_labels.index(self.nametowidget(event.widget))
        if exclusive:
            self._deselect_all()
        for col in self._closed_range(anchor_col, event_col):
            self._select_column(col, exclusive = False, anchor=False)
        self._set_anchor(self._cells[(0, anchor_col)])

    def _on_row_label_shift_click(self, event, exclusive=True):
        (anchor_row, _) = self._get_anchor_coords()
        event_row = self._row_labels.index(self.nametowidget(event.widget))
        if exclusive:
            self._deselect_all()
        for row in self._closed_range(anchor_row, event_row):
            self._select_row(row, exclusive = False, anchor=False)
        self._set_anchor(self._cells[(anchor_row, 0)])

    def _select_column(self, column, exclusive = True, anchor = True):
        self._select_cells([self._cells[(row, column)] for row in range(self.spreadsheet_rows)], exclusive = exclusive)
        if anchor:
            self._set_anchor((0, column))

    def _on_row_label_click(self, event, exclusive = True):
        row = self._row_labels.index(self.nametowidget(event.widget))
        self._select_row(row, exclusive = exclusive)

    def _on_row_label_control_click(self, event):
        self._on_row_label_click(event, exclusive = False)

    def _select_row(self, row, exclusive = True, anchor=True):
        self._select_cells([self._cells[(row, column)] for column in range(self.spreadsheet_columns)], exclusive = exclusive)
        if anchor:
            self._set_anchor((row, 0))

    def _select_cell_indices(self, indices):
        self._select_cells([self._cells[index] for index in indices])


    def __init__(self, program_paths, parent, rows, columns):
        tk.Frame.__init__(self, parent)

        self.program_paths = program_paths

        self.containing_frame = parent

        self._cells = {}
        self._cells_inverse = {}
        self._cells = {}
        self.spreadsheet_rows = rows
        self.spreadsheet_columns = columns

        self._solidified_cells = []
        self._anchor_cell = None
        self._reel_cell = None
        self._prev_reel_cell = None


        self._column_labels = []
        self._row_labels = []

        self._selected_cells = []

        self._guarantee_widget_focus = None

        # register a command to use for validation
        vcmd = (self.register(self._on_spreadsheet_cell_exit), '%W', '%P')

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
            l.bind('<Control-Shift-Button-1>', lambda event: self._on_column_label_shift_click(event, exclusive=False))
            self._column_labels.append(l)

        for row in range(self.spreadsheet_rows):
            l = tk.Label(self, text='  ' + str(row + 1) + '  ')
            l.grid(row=row + 1, column=0)
            l.config(cursor='sb_right_arrow')
            l.bind('<Button-1>', self._on_row_label_click)
            l.bind('<Control-Button-1>', self._on_row_label_control_click)
            l.bind('<Shift-Button-1>', self._on_row_label_shift_click)
            l.bind('<Control-Shift-Button-1>', lambda event: self._on_row_label_shift_click(event, exclusive=False))
            self._row_labels.append(l)

        # create the table of widgets
        for row in range(self.spreadsheet_rows):
            for column in range(self.spreadsheet_columns):
                index = (row, column)
                sv = tk.StringVar()
                c = Cell(self, textvariable = sv, validate="focusout", validatecommand=vcmd, cell_index = utils.get_cell_index(row, column))
                c.grid(row=row+1, column=column+1, stick="nsew")
                c.config(justify="left", state='disabled', cursor='plus', highlightthickness = 1, highlightbackground = 'ghost white',
                            disabledbackground='white', highlightcolor = 'goldenrod')
                c.bind("<Button-1>", self._on_spreadsheet_click)
                c.bind("<Control-Button-1>", self._on_spreadsheet_control_click)
                c.bind("<Shift-Button-1>", self._on_spreadsheet_shift_click)
                c.bind("<Control-Shift-Button-1>", self._on_spreadsheet_control_shift_click)
                c.bind("<Left>", self._on_spreadsheet_typing_left)
                c.bind('<Right>', self._on_spreadsheet_typing_right)
                c.bind("<B1-Motion>", self._on_spreadsheet_mouse_motion)
                c.bind('<ButtonRelease-1>', self._on_spreadsheet_mouse_release)
                c.bind('<Control-BackSpace>', self._on_spreadsheet_control_backspace)
                c.bind('<Tab>', self._next)
                c.bind('<Shift-Tab>', self._prev)
                c.bind('<Return>', self._on_spreadsheet_down)
                c.bind('<Down>', self._on_spreadsheet_down)
                c.bind('<Up>', self._on_spreadsheet_up)
                c.bind('<Escape>', self._on_spreadsheet_escape)
                c.bind('<FocusOut>', self._on_entry_focus_out)
                c.bind('<Control-Key-d>', self._on_spreadsheet_control_d)
                c.bind('<Control-Key-E>', self._export_to_csv)

                self._cells[index] = c
                self._cells_inverse[c] = index

        # adjust column weights so they all expand equally
        for column in range(self.spreadsheet_columns):
            self.grid_columnconfigure(column, weight=1)
        # designate a final, empty row to fill up any extra space
        self.grid_rowconfigure(rows, weight=1)

        self.gsv = tk.StringVar()
        self.gsv.trace_add('write', lambda idc, idc2, idc3: self._go_to_entry_widget())
        self.god_entry = tk.Entry(self, textvariable = self.gsv)
        self.god_entry.grid(row=1, column=1)
        self.god_entry.lower()

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
        self.god_entry.bind('<Return>', self._on_spreadsheet_enter_key)
        self.god_entry.bind('<Control-Return>', self._on_spreadsheet_control_enter_key)
        self.god_entry.bind('<Control-Key-d>', self._on_spreadsheet_control_d)
        self.god_entry.bind('<Control-Key-a>', self._select_all)
        self.god_entry.bind('<Control-Key-D>', self._print_determinant)
        self.god_entry.bind('<Control-Key-I>', self._convert_to_inverse)
        self.god_entry.bind('<Control-Key-i>', self._print_inverse)
        self.god_entry.bind('<Control-Key-E>', self._export_to_csv)
        self.god_entry.bind('<Escape>', self._on_spreadsheet_escape)

        self._guarantee_focus = False

        self.god_entry.focus_set()

    def _on_spreadsheet_escape(self, event):
        self._select_cell(self._anchor_cell, exclusive=True, anchor=True)

    def _is_square_selection(self, event):
        pass

    def _print_determinant(self, event):
        message = None
        try:
            determinant = self._calculate_determinant()
            message = ('Determinant Calculation', str(determinant))
        except (TypeError, ValueError):
            message = ('Try again', 'Not all entries are integers')
        
        messagebox.showinfo(*message)

    # Assume 3x3 from (0, 0)
    def _calculate_determinant(self, event):
        a00 = float(self._cells[(0, 0)].get().strip())
        a01 = float(self._cells[(0, 1)].get().strip())
        a02 = float(self._cells[(0, 2)].get().strip())
        a10 = float(self._cells[(1, 0)].get().strip())
        a11 = float(self._cells[(1, 1)].get().strip())
        a12 = float(self._cells[(1, 2)].get().strip())
        a20 = float(self._cells[(2, 0)].get().strip())
        a21 = float(self._cells[(2, 1)].get().strip())
        a22 = float(self._cells[(2, 2)].get().strip())
        
        d1 = a11 * a22 - a21 * a12
        d2 = a10 * a22 - a20 * a22
        d3 = a10 * a21 - a20 * a11
        d = a00 * d1 - a01 * d2 + a02 * d3

        return d
        

    def _convert_to_inverse(self):
        a00 = self._cells[(0, 0)]
        a01 = self._cells[(0, 1)]
        a02 = self._cells[(0, 2)]
        a10 = self._cells[(1, 0)]
        a11 = self._cells[(1, 1)]
        a12 = self._cells[(1, 2)]
        a20 = self._cells[(2, 0)]
        a21 = self._cells[(2, 1)]
        a22 = self._cells[(2, 2)]

        sv00 = self._cells[self._cells_inverse[a00]].sv
        sv01 = self._cells[self._cells_inverse[a01]].sv
        sv02 = self._cells[self._cells_inverse[a02]].sv
        sv10 = self._cells[self._cells_inverse[a10]].sv
        sv11 = self._cells[self._cells_inverse[a11]].sv
        sv12 = self._cells[self._cells_inverse[a12]].sv
        sv20 = self._cells[self._cells_inverse[a20]].sv
        sv21 = self._cells[self._cells_inverse[a21]].sv
        sv22 = self._cells[self._cells_inverse[a22]].sv

        determinant = self._calculate_determinant()

        result00 = float(a00.get()) / determinant
        result01 = float(a01.get()) / determinant
        result02 = float(a02.get()) / determinant
        result10 = float(a10.get()) / determinant
        result11 = float(a11.get()) / determinant
        result12 = float(a12.get()) / determinant
        result20 = float(a20.get()) / determinant
        result21 = float(a21.get()) / determinant
        result22 = float(a22.get()) / determinant

        if int(result00) == result00: result00 = int(result00)
        if int(result01) == result01: result01 = int(result01)
        if int(result02) == result02: result02 = int(result02)
        if int(result10) == result10: result10 = int(result10)
        if int(result11) == result11: result11 = int(result11)
        if int(result12) == result12: result12 = int(result12)
        if int(result20) == result20: result20 = int(result20)
        if int(result21) == result21: result21 = int(result21)
        if int(result22) == result22: result22 = int(result22)

        sv00.set(result00)
        sv01.set(result01)
        sv02.set(result02)
        sv10.set(result10)
        sv11.set(result11)
        sv12.set(result12)
        sv20.set(result20)
        sv21.set(result21)
        sv22.set(result22)

        a00.update()
        a01.update()
        a02.update()
        a10.update()
        a11.update()
        a12.update()
        a20.update()
        a21.update()
        a22.update()



    def _print_inverse(self, event):
        pass

    def _on_spreadsheet_control_d(self, event):
        self._copy_from_anchor_to_selected(self._anchor_cell)

    def _on_spreadsheet_typing_escape(self, event):
        # TBD, keep old entry somewhere
        pass

    def _on_spreadsheet_enter_key(self, event):
        return self._entry_focus(self._anchor_cell, highlight=False)

    def _on_spreadsheet_control_enter_key(self, event):
        self._deselect_cell(self._anchor_cell)

    def _on_god_entry_focus_out(self, event):
        if self._guarantee_focus:
            self.god_entry.focus_set()

        self._guarantee_focus = False

    def _get_anchor_sv(self):
        return self._cells[self._cells_inverse[self._anchor_cell]].sv

    def _go_to_entry_widget(self):
        self._anchor_cell.display_value = self.gsv.get()
        self.gsv.set('')
        return self._entry_focus(self._anchor_cell, highlight=False)

    def _copy_from_anchor_to_selected(self, entry = None):
        print('Updating all entries!')
        if entry == self._anchor_cell:
            for entry_widget in self._solidified_cells:
                if entry_widget == self._anchor_cell:
                    continue
                entry_widget.display_value = self._anchor_cell.display_value
                entry_widget.update()

    def _on_spreadsheet_tab(self, event = None):
        self._next()

    def _on_spreadsheet_shift_tab(self, event=None):
        self._prev()

    def _prev(self, event = None):
        if len(self._solidified_cells) == 1:
            self._on_spreadsheet_left(wrap=True)
        elif len(self._solidified_cells) > 1:
            pass
        else:
            pass

        self._guarantee_focus = True
        
        # Not effective for some reason
        # return 'break'

    def _next(self, event = None):
        if len(self._solidified_cells) == 1:
            self._on_spreadsheet_right(wrap=True)
        elif len(self._solidified_cells) > 1:
            pass
        else:
            pass

        self._guarantee_focus = True

        # Not effective for some reason
        # return 'break'

    def _get_anchor_coords(self):
        anchor_widget = self.nametowidget(self._anchor_cell)
        return self._cells_inverse[anchor_widget]

    def _get_reel_coords(self):
        reel_widget = self.nametowidget(self._reel_cell)
        return self._cells_inverse[reel_widget]        

    def _on_spreadsheet_shift_dir(self, event=None):
        self._select_range()

    def _on_spreadsheet_shift_up(self, event=None): # exclusive for control-shift
        reel_coords = reel_row, reel_col = self._get_reel_coords()
        offset = (0, 0)
        if reel_row > 0:
            offset = (-1, 0)
        self._set_reel(self._cells[tuple(map(add, reel_coords, offset))])
        self._select_range(exclusive=True)

    def _on_spreadsheet_alt_up(self, event=None):
        # This should move the anchor one up, nothing else should happen
        pass

    def _on_spreadsheet_shift_down(self, event=None):
        reel_coords = reel_row, reel_col = self._get_reel_coords()
        offset = (0, 0)
        if reel_row < self.spreadsheet_rows - 1:
            offset = (1, 0)
        self._set_reel(self._cells[tuple(map(add, reel_coords, offset))])
        self._select_range(exclusive=True)

    def _on_spreadsheet_shift_left(self, event=None):
        reel_coords = reel_row, reel_col = self._get_reel_coords()
        offset = (0, 0)
        if reel_col > 0:
            offset = (0, -1)
        self._set_reel(self._cells[tuple(map(add, reel_coords, offset))])
        self._select_range(exclusive=True)


    def _on_spreadsheet_shift_right(self, event=None):
        reel_coords = reel_row, reel_col = self._get_reel_coords()
        offset = (0, 0)
        if reel_col < self.spreadsheet_columns - 1:
            offset = (0, 1)
        self._set_reel(self._cells[tuple(map(add, reel_coords, offset))])
        self._select_range(exclusive=True)

    def _on_spreadsheet_up(self, event = None, exclusive = True):
        if self._anchor_cell:
            anchor_row, anchor_col = self._get_anchor_coords()
            if anchor_row > 0:
                self._select_cell(self._cells[(anchor_row - 1, anchor_col)], anchor=True, exclusive=exclusive, solidify=True)
            return 'break'
        else:
            self._set_anchor(self._cells[(0, 0)])

    def _on_spreadsheet_down(self, event = None, exclusive = True):
        anchor_row, anchor_col = self._get_anchor_coords()
        if anchor_row < self.spreadsheet_rows - 1:
            self._select_cell(self._cells[(anchor_row + 1, anchor_col)], anchor=True, exclusive=exclusive, solidify=True)
        return 'break'

    def _on_spreadsheet_left(self, event = None, exclusive = True, wrap=False):
        anchor_row, anchor_col = self._get_anchor_coords()
        if anchor_col > 0:
            self._select_cell(self._cells[(anchor_row, anchor_col - 1)], anchor=True, exclusive=exclusive, solidify=True)
        elif wrap and anchor_row > 0:
            self._select_cell(self._cells[(anchor_row - 1, self.spreadsheet_columns - 1)], anchor=True, exclusive=exclusive, solidify=True)

    def _on_spreadsheet_right(self, event = None, exclusive = True, wrap=False):
        anchor_row, anchor_col = self._get_anchor_coords()
        if anchor_col < self.spreadsheet_columns - 1:
            self._select_cell(self._cells[(anchor_row, anchor_col + 1)], anchor=True, exclusive=exclusive, solidify=True)
        elif wrap and anchor_row < self.spreadsheet_rows - 1:
            self._select_cell(self._cells[(anchor_row + 1, 0)], anchor=True, exclusive=exclusive, solidify=True)


    def get(self):
        '''Return a list of lists, containing the data in the table'''
        result = []
        for row in range(self.spreadsheet_rows):
            current_row = []
            for column in range(self.spreadsheet_columns):
                index = (row, column)
                current_row.append(self._cells[index].get())
            result.append(current_row)
        return result

class Example(tk.Frame):
    def __init__(self, parent, program_paths):
        tk.Frame.__init__(self, parent)
        self._spreadsheet = Spreadsheet(program_paths, self, 10, 10)
        self.submit = tk.Button(self, text="Submit", command=self.on_submit)
        self._spreadsheet.pack(side="top", fill="both", expand=True)
        self.submit.pack(side="bottom")

    def on_submit(self):
        print_(self.spreadsheet.get())