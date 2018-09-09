import tkinter as tk
import re

import importlib

from tkinter import messagebox, filedialog

import src.utils as utils
import src.expr_evaluator as arithmetic_evaluator

from PIL import Image, ImageTk

from operator import add

import os

import csv

class Cell(tk.Entry):
    def __init__(self, *args, **kw):
        sv = tk.StringVar()
      
        kw['textvariable'] = sv
        self.cell_index = kw.pop('cell_index')

        super().__init__(*args, **kw)
        
        self.sv = kw['textvariable']

        self.bind('<Control-BackSpace>', self._on_spreadsheet_control_backspace)
        self.bind("<Left>", self._on_spreadsheet_typing_left)
        self.bind('<Right>', self._on_spreadsheet_typing_right)

        self._formula_value = ''
        self._computed_value = ''

        self._mode = 'formula'

    def _on_spreadsheet_control_backspace(self, event):
        print('Control-backspace')
        cell = self.nametowidget(event.widget)
        self._erase_cell_contents()
        cell.icursor(0)
        self.focus_set()

    def _erase_cell_contents(self):
        self._formula_value = ''
        

    def _on_spreadsheet_typing_left(self, event):
        if self.selection_present():
            self.icursor(self.index(tk.ANCHOR) + 1)
            

    def _on_spreadsheet_typing_right(self, event = None):
        if self.selection_present():
            self.icursor(self.index(tk.ANCHOR) - 1)

    @property
    def formula_value(self):
        return self._formula_value
    

    @formula_value.setter
    def formula_value(self, value):
        self._formula_value = value
        self._update_display()


    @property
    def computed_value(self):
        return self._computed_value
    

    @computed_value.setter
    def computed_value(self, value):
        self._computed_value = value
        self._update_display()

    
    @property
    def coordinates(self):
        return utils.normalize_cell_notation(self.cell_index)

    def _set_background(self, color_string):
        self.config(background=color_string, disabledbackground=color_string)

    def _update_display(self):
        if self._mode == 'formula':
            self.sv.set(self._formula_value)
        else:
            self.sv.set(self._computed_value)

    def mode(self, _mode):
        self._mode = _mode
        self._update_display()

    def _update_formula_value(self):
        self._formula_value = self.sv.get()

    def _align_based_on_entry_type(self):
        print('aligning!')
        try:
            float(self.computed_value)
            self.config(justify='right')
        except ValueError:
            self.config(justify='left')

    def __repr__(self):
        return self.cell_index

class Spreadsheet(tk.Frame):
    def __init__(self, program_paths, parent, rows, columns):
        tk.Frame.__init__(self, parent)

        self.program_paths = program_paths

        self.containing_frame = parent

        self.cells = {}
        self._cells_inverse = {}
        self.spreadsheet_rows = rows
        self.spreadsheet_columns = columns

        self._anchor_cell = None
        self._reel_cell = None
        self._prev_reel_cell = None


        self._column_labels = []
        self._row_labels = []

        self._selected_cells = []

        self._guarantee_widget_focus = None

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
            l.bind('<B1-Motion>', self._on_column_label_mouse_motion)
            self._column_labels.append(l)

        for row in range(self.spreadsheet_rows):
            l = tk.Label(self, text='  ' + str(row + 1) + '  ')
            l.grid(row=row + 1, column=0)
            l.config(cursor='sb_right_arrow')
            l.bind('<Button-1>', self._on_row_label_click)
            l.bind('<Control-Button-1>', self._on_row_label_control_click)
            l.bind('<Shift-Button-1>', self._on_row_label_shift_click)
            l.bind('<Control-Shift-Button-1>', lambda event: self._on_row_label_shift_click(event, exclusive=False))
            l.bind('<B1-Motion>', self._on_row_label_mouse_motion)
            self._row_labels.append(l)

        # create the table of widgets
        for row in range(self.spreadsheet_rows):
            for column in range(self.spreadsheet_columns):
                index = (row, column)
                c = Cell(self, cell_index = utils.get_cell_index(row, column))
                c.grid(row=row+1, column=column+1, stick="nsew")
                c.config(justify="left", state='disabled', cursor='plus', highlightthickness = 1, highlightbackground = 'ghost white',
                            disabledbackground='white', highlightcolor = 'goldenrod')
                c.bind("<Button-1>", self._on_spreadsheet_click)
                c.bind("<Control-Button-1>", self._on_spreadsheet_control_click)
                c.bind("<Shift-Button-1>", self._on_spreadsheet_shift_click)
                c.bind("<Control-Shift-Button-1>", self._on_spreadsheet_control_shift_click)
                c.bind("<B1-Motion>", self._on_spreadsheet_mouse_motion)
                c.bind('<Tab>', self._next)
                c.bind('<Shift-Tab>', self._prev)
                c.bind('<Return>', self._on_spreadsheet_down)
                c.bind('<Down>', self._on_spreadsheet_down)
                c.bind('<Up>', self._on_spreadsheet_up)
                c.bind('<Escape>', self._on_spreadsheet_escape)
                c.bind('<FocusIn>', self._on_cell_begin_typing)
                c.bind('<FocusOut>', self._on_exit_cell_typing)
                c.bind('<Control-Key-d>', self._on_spreadsheet_control_d)
                c.bind('<Control-Key-E>', self._export_to_csv)

                self.cells[index] = c
                self._cells_inverse[c] = index

        # adjust column weights so they all expand equally
        for column in range(self.spreadsheet_columns):
            self.grid_columnconfigure(column, weight=1)
        # designate a final, empty row to fill up any extra space
        self.grid_rowconfigure(rows, weight=1)

        self.gsv = tk.StringVar()
        self.gsv.trace_add('write', lambda idc, idc2, idc3: self._start_anchor_entry_cursor())
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
        self.god_entry.bind('<Control-Key-E>', self._export_to_csv)
        self.god_entry.bind('<Control-Key-p>', self._dot_plot)
        self.god_entry.bind('<Escape>', self._on_spreadsheet_escape)
        self.god_entry.bind('<Control-Key-m>', self._create_macro)
        self.god_entry.bind('<Control-Key-l>', self._import_macro)

        self._guarantee_focus = False

        self.containing_frame.focus_set()

    def set_formula(self, formula, *cell_refs):
        cells = []
        row = None
        for cell in cell_refs:
            if row:
                column = cell
                cells.append(self.cells[utils.normalize_cell_notation(None, row, column)])
                row = None
            elif type(cell) == Cell:
                cells.append(cell)
            elif type(cell) == str or type(cell) == tuple:
                cells.append(self.cells[utils.normalize_cell_notation(None, cell)])
            elif type(cell) == int:
                row = cell
            else:
                raise ValueError('Cell reference ' + str(cell) + ' is illegal')

        for cell in cells:
            cell.formula = formula
            self._update_display_based_on_formula(cell)



    def _dot_plot(self, event):
        []

    def _get_formatted_value(self, match, stringify=False):
        cell = self.cells[utils.normalize_cell_notation(self, match)]
        return "'" + cell.computed_value + "'" if stringify else cell.computed_value

    def _cell_convert(self, value, stringify=False):
        return re.sub(r'\[.*?\]', lambda match: self._get_formatted_value(match[0], stringify), value)

    def _process_formula(self, formula, number_based = True):
        if formula and formula[0] == '=' and len(formula) > 1:
            converted_value = self._cell_convert(formula[1:], not number_based)
            if number_based:
                return str(arithmetic_evaluator.evaluate_expression(converted_value))
            else:
                return eval(converted_value)
        else:
            return formula

    def _select_all(self, event):
        print('Selecting all cells in the grid!')
        self._select_range(exclusive=True, anchor='A1', reel=(-1, -1))

    def _select_cells(self, cells, exclusive=False, flip=False):
        print('Selecting cells ' + str([cell.cell_index for cell in cells]))
        if exclusive:
            self._deselect_all()
        
        [self._select_cell(cell, exclusive=False, flip=flip) for cell in cells]

    def _select_cell(self, cell, anchor=False, exclusive=False, flip=False):
        print('Selecting cell ' + repr(cell))

        if exclusive:
            self._deselect_all()          

        if cell in self._selected_cells:
            if flip:
                print('flip')
                self._deselect_cell(cell)
            else:
                if anchor:
                    self._set_anchor(cell)
                return
        elif anchor:
            self._set_anchor(cell)
            self._selected_cells.append(cell)
        else:
            cell.config(highlightbackground = 'darkgreen')
            self._selected_cells.append(cell)

        self.god_entry.focus_set()

        print('self._selected_cells now is ' + str(self._selected_cells))



    def _deselect_all(self, but=[]):
        print('deselect all')

        self._deselect_cells([cell for cell in self._selected_cells if cell not in but])

    def _deselect_cells(self, cells):
        print('Deselecting cells ' + repr(cells))
        [self._deselect_cell(cell) for cell in cells]

    def _deselect_cell(self, cell):
        print('deselect cell : ' + repr(cell))
        try:
            self._selected_cells.remove(cell)
            cell.config(highlightbackground = 'ghost white')
            if cell == self._anchor_cell:
                print('Anchor')
                anchor = self._selected_cells[-1] if self._selected_cells else None
                self._set_anchor(anchor)
        except ValueError:
            pass

        if not self._selected_cells:
            print('no selected cells')
            self.containing_frame.focus_set()


    def _restore_borders(self, new, method='normal'):
        print('restore borders')
        hlbg = 'darkgreen' if method == 'selected' else 'ghost white'

        if type(new) != list:
            new = [new]
        for old in self._selected_cells:
            if old not in new:
                old.config(highlightbackground = hlbg)

    def _select_range(self, keepanchor = True, exclusive = False, flip=False, anchor = None, reel = None):
        print('Selecting ' + ('exclusive' if exclusive else '') +  ' range: ', end='')

        if exclusive:
            but = [self._anchor_cell] if keepanchor else []
            self._deselect_all(but=but)

        if anchor:
            self._set_anchor(anchor, add=True)

        if reel:
            self._set_reel(reel)
        
        anchor_coordinates = (a_row, a_column) = self._cells_inverse[self._anchor_cell]

        prev_reel_cell = self._anchor_cell if exclusive else self._prev_reel_cell

        prev_reel_coordinates = (p_row, p_column) = self._cells_inverse[prev_reel_cell]
        reel_coordinates = (r_row, r_column) = self._cells_inverse[self._reel_cell]
    
        row_range = utils.closed_range(a_row, r_row)[::-1]
        if self._prev_reel_cell:
            prev_row_range = utils.closed_range(a_row, p_row)[::-1]

        column_range = utils.closed_range(a_column, r_column)[::-1]
        if self._prev_reel_cell:
            prev_column_range = utils.closed_range(a_column, p_column)[::-1]

        if row_range[-1] > prev_row_range[-1]:
            print('The row minimum increased')
            for row in range(prev_row_range[-1], row_range[-1]):
                self._deselect_cells([self.cells[(row, column)] for column in prev_column_range])

        elif row_range[0] < prev_row_range[0]:
            print('The row maximum decreased')
            for row in range(prev_row_range[0], row_range[0], -1):
                self._deselect_cells([self.cells[(row, column)] for column in prev_column_range])

        if row_range[-1] < prev_row_range[-1]:
            print('The row minimum decreased')
            for row in range(row_range[-1], prev_row_range[-1]):
                self._select_cells([self.cells[(row, column)] for column in prev_column_range], flip=flip)

        elif row_range[0] > prev_row_range[0]:
            print('The row maximum increased')
            for row in range(row_range[0], prev_row_range[0], -1):
                self._select_cells([self.cells[(row, column)] for column in prev_column_range], flip=flip)

        

        
        if column_range[-1] > prev_column_range[-1]:
            print('The column minimum increased')
            for column in range(prev_column_range[-1], column_range[-1]):
                self._deselect_cells([self.cells[(row, column)] for row in row_range])

        elif column_range[0] < prev_column_range[0]:
            print('The column maximum decreased')
            for column in range(prev_column_range[0], column_range[0], -1):
                self._deselect_cells([self.cells[(row, column)] for row in row_range])


        if column_range[-1] < prev_column_range[-1]:
            print('The column minimum decreased')
            for column in range(column_range[-1], prev_column_range[-1]):
                self._select_cells([self.cells[(row, column)] for row in row_range], flip=flip)

        elif column_range[0] > prev_column_range[0]:
            print('The column maximum increased')
            for column in range(column_range[0], prev_column_range[0], -1):
                self._select_cells([self.cells[(row, column)] for row in row_range], flip=flip)



    def _set_anchor(self, cell, col = None, add = False, flip=False):        
        if not cell and cell != 0:
            self._anchor_cell = None
            return

        if type(cell) != Cell:
            cell = self.cells[utils.normalize_cell_notation(self, cell, col)]

        if flip and cell in self._selected_cells:
            self._deselect_cell(cell)
            return

        print('Setting cell ' + repr(cell) + ' to anchor')

        self._anchor_cell = cell
        cell.config(highlightbackground = 'goldenrod')
        self._restore_borders(self._anchor_cell, method='selected')

        self._set_reel(cell)

        if add and cell not in self._selected_cells:
            self._selected_cells.append(cell)

    def _set_reel(self, cell, col=None):
        if type(cell) != Cell:
            cell = self.cells[utils.normalize_cell_notation(self, cell, col)]

        self._reel_cell, self._prev_reel_cell = cell, self._reel_cell

        print('Reel cell is now ' + repr(self._reel_cell))

    def _on_spreadsheet_click(self, event):
        print(event.type + ': <Button-1>')
        cell = self.nametowidget(event.widget)

        if not self.focus_get() == event.widget:
            if [cell] == self._selected_cells:
                cell.focus_set()
            else:
                self._select_cell(cell, anchor=True, exclusive=True)
                self.god_entry.focus_set()

    def _on_spreadsheet_control_click(self, event):
        print(event.type + ': <Control-Button-1>')
        cell = self.nametowidget(event.widget)
        self._select_cell(cell, anchor=True, flip=True)

    def _on_spreadsheet_shift_click(self, event):
        print(event.type + ': <Shift-Button-1>')
        self._deselect_all(but=[self._anchor_cell])
        self._select_range(reel = self.nametowidget(event.widget), exclusive = True)

    def _on_spreadsheet_control_shift_click(self, event):
        print(event.type + ': <Control-Shift-Button-1>')
        self._select_range(reel = self.nametowidget(event.widget))

    def _on_spreadsheet_mouse_motion(self, event):
        self._set_reel(self.winfo_containing(event.x_root, event.y_root))
        if self._reel_cell not in self.cells.values():
            print(str(self._reel_cell) + ' is out of bounds')
            return
        if self._prev_reel_cell in self.cells.values():
            self._select_range(exclusive=False)
        else:
            print("Coming from out of bounds")
            self._select_range(exclusive = True)

    def _on_spreadsheet_shift_up(self, event=None): # exclusive for control-shift
        reel_coords = reel_row, reel_col = self._get_reel_coords()
        offset = (0, 0)
        if reel_row > 0:
            offset = (-1, 0)
        self._set_reel(self.cells[tuple(map(add, reel_coords, offset))])
        self._select_range()

    def _on_spreadsheet_shift_down(self, event=None):
        reel_coords = reel_row, reel_col = self._get_reel_coords()
        offset = (0, 0)
        if reel_row < self.spreadsheet_rows - 1:
            offset = (1, 0)
        self._set_reel(self.cells[tuple(map(add, reel_coords, offset))])
        self._select_range()

    def _on_spreadsheet_shift_left(self, event=None):
        reel_coords = reel_row, reel_col = self._get_reel_coords()
        offset = (0, 0)
        if reel_col > 0:
            offset = (0, -1)
        self._set_reel(self.cells[tuple(map(add, reel_coords, offset))])
        self._select_range()


    def _on_spreadsheet_shift_right(self, event=None):
        reel_coords = reel_row, reel_col = self._get_reel_coords()
        offset = (0, 0)
        if reel_col < self.spreadsheet_columns - 1:
            offset = (0, 1)
        self._set_reel(self.cells[tuple(map(add, reel_coords, offset))])
        self._select_range()

    def _on_spreadsheet_up(self, event = None, exclusive = True):
        if self._anchor_cell:
            anchor_row, anchor_col = self._get_anchor_coords()
            if anchor_row > 0:
                self._select_cell(self.cells[(anchor_row - 1, anchor_col)], anchor=True, exclusive=exclusive)
            return 'break'
        else:
            self._set_anchor(self.cells[(0, 0)])

    def _on_spreadsheet_down(self, event = None, exclusive = True):
        anchor_row, anchor_col = self._get_anchor_coords()
        if anchor_row < self.spreadsheet_rows - 1:
            self._select_cell(self.cells[(anchor_row + 1, anchor_col)], anchor=True, exclusive=exclusive)
        return 'break'

    def _on_spreadsheet_left(self, event = None, exclusive = True, wrap=False):
        anchor_row, anchor_col = self._get_anchor_coords()
        if anchor_col > 0:
            self._select_cell(self.cells[(anchor_row, anchor_col - 1)], anchor=True, exclusive=exclusive)
        elif wrap and anchor_row > 0:
            self._select_cell(self.cells[(anchor_row - 1, self.spreadsheet_columns - 1)], anchor=True, exclusive=exclusive)

    def _on_spreadsheet_right(self, event = None, exclusive = True, wrap=False):
        anchor_row, anchor_col = self._get_anchor_coords()
        if anchor_col < self.spreadsheet_columns - 1:
            self._select_cell(self.cells[(anchor_row, anchor_col + 1)], anchor=True, exclusive=exclusive)
        elif wrap and anchor_row < self.spreadsheet_rows - 1:
            self._select_cell(self.cells[(anchor_row + 1, 0)], anchor=True, exclusive=exclusive)


    def _on_column_label_click(self, event, exclusive = True, flip=False):
        column = self._column_labels.index(self.nametowidget(event.widget))
        self._select_column(column, exclusive = exclusive, flip=flip)
        self._column_y = event.y_root

    def _on_column_label_control_click(self, event):
        column = self._column_labels.index(self.nametowidget(event.widget))
        column_cells = [self.cells[(row, column)] for row in range(self.spreadsheet_rows)]
        num_cells_selected = sum([column_cell in self._selected_cells for column_cell in column_cells])

        if num_cells_selected > self.spreadsheet_rows / 2:
            self._deselect_cells(column_cells)
        else:
            self._select_cells(column_cells)
            self._set_anchor(column_cells[0])

        self._column_y = event.y_root

    def _on_column_label_shift_click(self, event, exclusive=True):
        (_, anchor_col) = self._cells_inverse[self._anchor_cell]
        event_col = self._column_labels.index(self.nametowidget(event.widget))
        self._select_range(anchor=(0, anchor_col), reel=(self.spreadsheet_rows - 1, event_col), keepanchor = True, exclusive = exclusive)
        self._column_y = event.y_root

    def _on_column_label_mouse_motion(self, event):
        reel_col = self._column_labels.index(self.winfo_containing(event.x_root, self._column_y))
        self._select_range(exclusive=False, keepanchor=True, reel=(-1, reel_col))

    def _select_column(self, column, exclusive = True, flip = False):
        print('Selecting column ' + str(column))
        self._select_range(anchor = (0, column), keepanchor = False, reel = (self.spreadsheet_rows - 1, column), exclusive = exclusive, flip=True)

    def _on_row_label_click(self, event, exclusive = True):
        row = self._row_labels.index(self.nametowidget(event.widget))
        self._select_row(row, exclusive = exclusive)
        self._row_x = event.x_root

    def _on_row_label_control_click(self, event):
        self._on_row_label_click(event, exclusive = False)

    def _on_row_label_shift_click(self, event, exclusive=True):
        (anchor_row, _) = self._cells_inverse[self._anchor_cell]
        event_row = self._row_labels.index(self.nametowidget(event.widget))
        self._select_range(anchor=(anchor_row, 0), reel=(event_row, -1), keepanchor = True, exclusive = exclusive)

    def _on_row_label_mouse_motion(self, event):
        reel_row = self._row_labels.index(self.winfo_containing(self._row_x, event.y_root))
        self._select_range(exclusive=False, keepanchor=True, reel=(reel_row, -1))

    def _select_row(self, row, exclusive = True):
        print('Selecting row ' + str(row))
        self._select_range(anchor = (row, 0), keepanchor = False, reel = (row, -1), exclusive = exclusive)

    def _on_spreadsheet_backspace(self, event):
        self._erase_selected_cell_contents()

    def _on_spreadsheet_delete(self, event):
        self._erase_selected_cell_contents()

    def _erase_selected_cell_contents(self):
        for cell in self._selected_cells:
            cell._erase_cell_contents()

    def _on_spreadsheet_control_d(self, event):
        self._copy_from_anchor_to_selected(self._anchor_cell)

    def _on_spreadsheet_tab(self, event = None):
        self._next()

    def _on_spreadsheet_shift_tab(self, event=None):
        self._prev()

    def _on_spreadsheet_enter_key(self, event):
        self._select_cell(self._anchor_cell, exclusive=True, anchor=True)
        if self._anchor_cell:
            self.focus_set()

    def _on_spreadsheet_control_enter_key(self, event):
        self._deselect_cell(self._anchor_cell)

    def _on_spreadsheet_escape(self, event):
        self._select_cell(self._anchor_cell, exclusive=True, anchor=True)

    def _convert_to_inverse(self):
        a00 = self.cells[(0, 0)]
        a01 = self.cells[(0, 1)]
        a02 = self.cells[(0, 2)]
        a10 = self.cells[(1, 0)]
        a11 = self.cells[(1, 1)]
        a12 = self.cells[(1, 2)]
        a20 = self.cells[(2, 0)]
        a21 = self.cells[(2, 1)]
        a22 = self.cells[(2, 2)]

        sv00 = self.cells[self._cells_inverse[a00]].sv
        sv01 = self.cells[self._cells_inverse[a01]].sv
        sv02 = self.cells[self._cells_inverse[a02]].sv
        sv10 = self.cells[self._cells_inverse[a10]].sv
        sv11 = self.cells[self._cells_inverse[a11]].sv
        sv12 = self.cells[self._cells_inverse[a12]].sv
        sv20 = self.cells[self._cells_inverse[a20]].sv
        sv21 = self.cells[self._cells_inverse[a21]].sv
        sv22 = self.cells[self._cells_inverse[a22]].sv

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

    def _export_to_csv(self, event=None):
        values = self.get()

        filename = filedialog.asksaveasfilename(title='Export to CSV', initialdir=self.program_paths['index'], filetypes=[('Comma Separated Values', '*.csv')])
        
        if not filename:
            return

        if filename[-4:] != '.csv':
            filename += '.csv'

        with open(filename, "w+", newline='') as f:
            writer = csv.writer(f, skipinitialspace=True)
            writer.writerows(values) 

        os.startfile(filename)

    def _compute_formula(self, value):
        try:
            value = self._process_formula(value)
        except TypeError:
            value = self._process_formula(value, number_based = False)
        except SyntaxError:
            # Tell the user their syntax was off
            pass

        return value


    def _update_display_based_on_formula(self, cell):
        cell.computed_value = self._compute_formula(cell.formula_value.strip())
        cell.config(state='disabled', cursor='plus')
        cell._align_based_on_entry_type()
        cell.mode('computed')

    def _on_cell_begin_typing(self, event, highlight=False):
        cell = self.nametowidget(event.widget)
        cell.mode('formula')
        cell.config(state='normal', cursor='xterm')
        if highlight:
            cell.selection_range(0, 'end')
        cell.icursor(tk.END)

    def _on_exit_cell_typing(self, event=None):
        if self._guarantee_widget_focus:
            print("Not leaving the cell quite yet!")
            self.nametowidget(event.widget).focus_set()
        else:
            print('leaving cell!')
            cell = self.nametowidget(event.widget)
            cell._update_formula_value()
            self._update_display_based_on_formula(cell)

            return True
            

        self._guarantee_widget_focus = None


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
        a00 = float(self.cells[(0, 0)].get().strip())
        a01 = float(self.cells[(0, 1)].get().strip())
        a02 = float(self.cells[(0, 2)].get().strip())
        a10 = float(self.cells[(1, 0)].get().strip())
        a11 = float(self.cells[(1, 1)].get().strip())
        a12 = float(self.cells[(1, 2)].get().strip())
        a20 = float(self.cells[(2, 0)].get().strip())
        a21 = float(self.cells[(2, 1)].get().strip())
        a22 = float(self.cells[(2, 2)].get().strip())
        
        d1 = a11 * a22 - a21 * a12
        d2 = a10 * a22 - a20 * a22
        d3 = a10 * a21 - a20 * a11
        d = a00 * d1 - a01 * d2 + a02 * d3

        return d
        



    def _on_spreadsheet_typing_escape(self, event):
        # TBD, keep old entry somewhere
        pass


    def _on_god_entry_focus_out(self, event):
        if self._guarantee_focus:
            self.god_entry.focus_set()

        self._guarantee_focus = False

    def _start_anchor_entry_cursor(self):
        if self._anchor_cell:
            print('Go to the anchor')
            self._anchor_cell.formula_value = self.gsv.get()
            self.gsv.set('')
            self._anchor_cell.focus_set()

    def _copy_from_anchor_to_selected(self, cell = None):
        print('Updating all entries!')
        if cell == self._anchor_cell:
            for cell in self._selected_cells:
                if cell == self._anchor_cell:
                    continue
                cell.formula_value = self._anchor_cell.formula_value
                self._update_display_based_on_formula(cell)
                cell._align_based_on_entry_type()

    def _prev(self, event = None):
        if len(self._selected_cells) == 1:
            self._on_spreadsheet_left(wrap=True)
        elif len(self._selected_cells) > 1:
            pass
        else:
            pass

        self._guarantee_focus = True
        
        # Not effective for some reason
        # return 'break'

    def _next(self, event = None):
        if len(self._selected_cells) == 1:
            self._on_spreadsheet_right(wrap=True)
        elif len(self._selected_cells) > 1:
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

    def _create_macro(self, event):
        filename = filedialog.asksaveasfilename(title='Create Macro File', initialdir=self.program_paths['index'], filetypes=[('Python File', '*.py')])

        if not filename:
            return

        if filename[-3:] != '.py':
            filename += '.py'

        with open(filename, "w+") as python:
            python.write('from pysheets import *\n\n')
            # Comment out examples of ways to use the API

        os.startfile(filename)

    def _run_macro(self, event):
        self.active_macro_import.run(self)

    def _import_macro(self, event):
        active_macro_fullpath = filedialog.askopenfilename(title='Create Macro File', initialdir=self.program_paths['index'], filetypes=[('Python File', '*.py')])

        self.active_macro_name = os.path.split(active_macro_fullpath)[-1]

        print(active_macro_fullpath)

        macro_import_init_file = os.path.join(self.program_paths['index'], 'macros', '__init__.py')

        if not os.path.exists('macros'):
            os.makedirs('macros')

        thingy = 'macros.' + self.active_macro_name[:-3]

        with open(macro_import_init_file, "a") as init_file:
            print('got here')
            init_file.write('\nimport ' + thingy + '\n')

        self.active_macro_import = importlib.import_module(thingy)

        # Detect changes to file
        importlib.reload(self.active_macro_import)


        self.god_entry.bind('<Control-Key-R>', self._run_macro)
        self.god_entry.bind('<Control-Key-r>', self._run_macro)


    def get(self):
        '''Return a list of lists, containing the data in the table'''
        result = []
        for row in range(self.spreadsheet_rows):
            current_row = []
            for column in range(self.spreadsheet_columns):
                index = (row, column)
                current_row.append(self.cells[index].get())
            result.append(current_row)
        return result