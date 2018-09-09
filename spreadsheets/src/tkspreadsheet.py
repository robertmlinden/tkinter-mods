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
        self.sv = tk.StringVar()

        kw['textvariable'] = self.sv
        self.cell_index = kw.pop('cell_index')
        self.__ss = kw.pop('spreadsheet')

        super().__init__(*args, **kw)

        self.bind('<Control-BackSpace>', self.__control_backspace)
        self.bind("<Left>", self.__typing_left)
        self.bind('<Right>', self.__typing_right)
        self.bind('<KeyRelease>', lambda _ : self._absorb_display())

        self.__formula_value = ''
        self.__computed_value = ''

        self.__mode = 'formula'

    def _absorb_display(self):
        print('absorb!!')
        if self.__mode == 'formula':
            print(self.sv.get())
            self.__formula_value = self.sv.get()
        else:
            self.__computed_value = self.sv.get()

    def __control_backspace(self, event):
        print('Control-backspace')
        self.erase_cell_contents()
        self.icursor(0)
        self.focus_set()

    def erase_cell_contents(self):
        self.formula_value = ''
        self.computed_value = ''
        

    def __typing_left(self, event):
        if self.selection_present():
            self.icursor(self.index(tk.ANCHOR) + 1)
            

    def __typing_right(self, event = None):
        if self.selection_present():
            self.icursor(self.index(tk.ANCHOR) - 1)

    @property
    def formula_value(self):
        return self.__formula_value
    

    @formula_value.setter
    def formula_value(self, value):
        self.__formula_value = value
        self.__update_display()


    @property
    def computed_value(self):
        return self.__computed_value
    

    @computed_value.setter
    def computed_value(self, value):
        self.__computed_value = value
        self.__update_display()

    
    @property
    def coordinates(self):
        return utils.normalize_cell_notation(None, self.cell_index)

    def __set_background(self, color_string):
        self.config(background=color_string, disabledbackground=color_string)

    def __update_display(self):
        if self.__mode == 'formula':
            print('update formula')
            self.sv.set(self.__formula_value)
        else:
            print('update compute')
            self.sv.set(self.__computed_value)

    def mode(self, _mode):
        self.__mode = _mode
        self.__update_display()

    def align_based_on_entry_type(self):
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

        self.__cells = {}

        self.rows = rows
        self.columns = columns
        self.cols = columns

        self.__anchor_cell = None
        self.__reel_cell = None
        self.__prev_reel_cell = None


        self.__column_labels = []
        self.__row_labels = []

        self.__selected_cells = []

        self.__guarantee_widget_focus = None

        im = Image.open(os.path.join(self.program_paths['icons'], 'select_all_cropped.gif')).resize((10, 10), Image.ANTIALIAS)
        ph = ImageTk.PhotoImage(im)

        image_label = tk.Label(self, image=ph)
        image_label.grid(row=0, column=0, sticky='se')
        image_label.image=ph  #need to keep the reference of your image to avoid garbage collection

        image_label.bind('<Button-1>', self.__select_all)
        image_label.config(cursor='lr_angle')

        for column in range(self.columns):
            l = tk.Label(self, text=utils.get_column_letters(column))
            l.grid(row=0, column=column+1, sticky='ew')
            l.config(cursor='sb_down_arrow')
            l.bind('<Button-1>', self.__on_column_label_click)
            l.bind('<Control-Button-1>', self.__on_column_label_control_click)
            l.bind('<Shift-Button-1>', self.__on_column_label_shift_click)
            l.bind('<Control-Shift-Button-1>', lambda event: self.__on_column_label_shift_click(event, exclusive=False))
            l.bind('<B1-Motion>', self.__on_column_label_mouse_motion)
            self.__column_labels.append(l)

        for row in range(self.rows):
            l = tk.Label(self, text='  ' + str(row + 1) + '  ')
            l.grid(row=row + 1, column=0)
            l.config(cursor='sb_right_arrow')
            l.bind('<Button-1>', self.__on_row_label_click)
            l.bind('<Control-Button-1>', self.__on_row_label_control_click)
            l.bind('<Shift-Button-1>', self.__on_row_label_shift_click)
            l.bind('<Control-Shift-Button-1>', lambda event: self.__on_row_label_shift_click(event, exclusive=False))
            l.bind('<B1-Motion>', self.__on_row_label_mouse_motion)
            self.__row_labels.append(l)

        # create the table of widgets
        for row in range(self.rows):
            for column in range(self.columns):
                index = (row, column)
                c = Cell(self, cell_index = utils.get_cell_index(row, column), spreadsheet = self)
                c.grid(row=row+1, column=column+1, stick="nsew")
                c.config(justify="left", state='disabled', cursor='plus', highlightthickness = 1, highlightbackground = 'ghost white',
                            disabledbackground='white', highlightcolor = 'goldenrod')
                c.bind("<Button-1>", self.__click)
                c.bind("<Control-Button-1>", self.__control_click)
                c.bind("<Shift-Button-1>", self.__shift_click)
                c.bind("<Control-Shift-Button-1>", self.__control_shift_click)
                c.bind("<B1-Motion>", self.__mouse_motion)
                c.bind('<Tab>', self.__next)
                c.bind('<Shift-Tab>', self.__prev)
                c.bind('<Return>', self.__down)
                c.bind('<Down>', self.__down)
                c.bind('<Up>', self.__up)
                c.bind('<Escape>', self.__escape)
                c.bind('<FocusIn>', self.__on_cell_focus_in)
                c.bind('<FocusOut>', self.__on_exit_cell_typing)
                c.bind('<Control-Key-d>', self.__control_d)
                c.bind('<Control-Key-E>', self.__export_to_csv)
                c.bind('<Control-Key-a>', self.__select_all)

                self.__cells[index] = c

        # adjust column weights so they all expand equally
        for column in range(self.columns):
            self.grid_columnconfigure(column, weight=1)
        # designate a final, empty row to fill up any extra space
        self.grid_rowconfigure(rows, weight=1)

        self.gsv = tk.StringVar()
        self.gsv.trace_add('write', lambda idc, idc2, idc3: self.__start_anchor_entry_cursor())
        self.god_entry = tk.Entry(self, textvariable = self.gsv)
        self.god_entry.grid(row=1, column=1)
        self.god_entry.lower()

        self.god_entry.bind('<BackSpace>', self.__backspace)
        self.god_entry.bind('<Delete>', self.__delete)
        self.god_entry.bind('<Up>', lambda event: self.__up(event, exclusive=True))
        self.god_entry.bind('<Shift-Up>', self.__shift_up)
        self.god_entry.bind('<Control-Up>', lambda event: self.__up(event, exclusive=False))
        self.god_entry.bind('<Down>', lambda event: self.__down(event, exclusive=True))
        self.god_entry.bind('<Shift-Down>', self.__shift_down)
        self.god_entry.bind('<Control-Down>', lambda event: self.__down(event, exclusive=False))
        self.god_entry.bind('<Left>', lambda event: self.__left(event, exclusive=True))
        self.god_entry.bind('<Shift-Left>', self.__shift_left)
        self.god_entry.bind('<Control-Left>', lambda event: self.__left(event, exclusive=False))
        self.god_entry.bind('<Right>', lambda event: self.__right(event, exclusive=True))
        self.god_entry.bind('<Shift-Right>', self.__shift_right)
        self.god_entry.bind('<Control-Right>', lambda event: self.__right(event, exclusive=False))
        self.god_entry.bind('<Tab>', self.__tab)
        self.god_entry.bind('<Shift-Tab>', self.__shift_tab)
        self.god_entry.bind('<Control-Tab>', self.__shift_tab)
        self.god_entry.bind('<FocusOut>', self.__on_god_entry_focus_out)
        self.god_entry.bind('<Return>', self.__enter_key)
        self.god_entry.bind('<Control-Return>', self.__control_enter_key)
        self.god_entry.bind('<Shift-Return>', self.__shift_enter_key)
        self.god_entry.bind('<Control-Key-d>', self.__control_d)
        self.god_entry.bind('<Control-Key-a>', self.__select_all)
        self.god_entry.bind('<Control-Key-D>', self.__print_determinant)
        self.god_entry.bind('<Control-Key-I>', self.__convert_to_inverse)
        self.god_entry.bind('<Control-Key-E>', self.__export_to_csv)
        self.god_entry.bind('<Control-Key-p>', self.__dot_plot)
        self.god_entry.bind('<Escape>', self.__escape)
        self.god_entry.bind('<Control-Key-m>', self.__create_macro)
        self.god_entry.bind('<Control-Key-l>', self.__import_macro)

        self.__guarantee_focus = False

        self.containing_frame.focus_set()

    def __set_formula(self, formula, *cell_refs):
        cells = []
        row = None
        for cell in cell_refs:
            if row:
                column = cell
                if type(column) != int:
                    raise ValueError('Cell reference ' + str(cell) + ' needed to be an integer to couple with the previous argument')
                cells.append(self.__cells[utils.normalize_cell_notation(None, row, column)])
                row = None
            elif type(cell) == Cell:
                cells.append(cell)
            elif type(cell) == str or type(cell) == tuple:
                cells.append(self.__cells[utils.normalize_cell_notation(None, cell)])
            elif type(cell) == int:
                row = cell
            else:
                raise ValueError('Cell reference ' + str(cell) + ' is illegal')

        for cell in cells:
            cell.formula = formula
            self.__update_display_based_on_formula(cell)



    def __dot_plot(self, event):
        []

    def __get_formatted_value(self, match, stringify=False):
        cell = self.__cells[utils.normalize_cell_notation(self, match)]
        return "'" + cell.computed_value + "'" if stringify else cell.computed_value

    def __cell_convert(self, value, stringify=False):
        return re.sub(r'\[.*?\]', lambda match: self.__get_formatted_value(match[0], stringify), value)

    def __process_formula(self, formula, number_based = True):
        if formula and formula[0] == '=' and len(formula) > 1:
            converted_value = self.__cell_convert(formula[1:], not number_based)
            if number_based:
                return str(arithmetic_evaluator.evaluate_expression(converted_value))
            else:
                return eval(converted_value)
        else:
            return formula

    def __select_all(self, event):
        print('Selecting all cells in the grid!')
        self.__select_range(exclusive=True, anchor='A1', reel=(-1, -1))
        return 'break'

    def __select_cells(self, cells, exclusive=False, flip=False):
        print('Selecting cells ' + str([cell.cell_index for cell in cells]))
        if exclusive:
            self.__deselect_all()
        
        [self.__select_cell(cell, exclusive=False, flip=flip) for cell in cells]

    def __select_cell(self, cell, anchor=False, exclusive=False, flip=False):
        print('Selecting cell ' + repr(cell))

        if exclusive:
            self.__deselect_all()          

        if cell in self.__selected_cells:
            if flip:
                print('flip')
                self.__deselect_cell(cell)
            else:
                if anchor:
                    self.__set_anchor(cell)
                return
        elif anchor:
            self.__set_anchor(cell)
            self.__selected_cells.append(cell)
        else:
            cell.config(highlightbackground = 'darkgreen')
            self.__selected_cells.append(cell)

        self.god_entry.focus_set()

        print('self.__selected_cells now is ' + str(self.__selected_cells))



    def __deselect_all(self, but=[]):
        print('deselect all')

        self.__deselect_cells([cell for cell in self.__selected_cells if cell not in but])

    def __deselect_cells(self, cells):
        print('Deselecting cells ' + repr(cells))
        [self.__deselect_cell(cell) for cell in cells]

    def __deselect_cell(self, cell):
        print('deselect cell : ' + repr(cell))
        try:
            self.__selected_cells.remove(cell)
            cell.config(highlightbackground = 'ghost white')
            if cell == self.__anchor_cell:
                print('Anchor')
                anchor = self.__selected_cells[-1] if self.__selected_cells else None
                self.__set_anchor(anchor)
        except ValueError:
            pass

        if not self.__selected_cells:
            print('no selected cells')
            self.containing_frame.focus_set()


    def __restore_borders(self, new, method='normal'):
        print('restore borders')
        hlbg = 'darkgreen' if method == 'selected' else 'ghost white'

        if type(new) != list:
            new = [new]
        for old in self.__selected_cells:
            if old not in new:
                old.config(highlightbackground = hlbg)

    def __select_range(self, keepanchor = True, exclusive = False, flip=False, anchor = None, reel = None):
        print('Selecting ' + ('exclusive' if exclusive else '') +  ' range: ', end='')

        if exclusive:
            but = [self.__anchor_cell] if keepanchor else []
            self.__deselect_all(but=but)

        if anchor:
            self.__set_anchor(anchor, add=True)

        if reel:
            self.__set_reel(reel)
        
        prev_reel_cell = self.__anchor_cell if exclusive else self.__prev_reel_cell


        anchor_coordinates = (a_row, a_column) = self.__get_anchor_coords()
        prev_reel_coordinates = (p_row, p_column) = self.__get_prev_reel_coords()
        reel_coordinates = (r_row, r_column) = self.__get_reel_coords()

    
        row_range = utils.closed_range(a_row, r_row)[::-1]
        if self.__prev_reel_cell:
            prev_row_range = utils.closed_range(a_row, p_row)[::-1]

        column_range = utils.closed_range(a_column, r_column)[::-1]
        if self.__prev_reel_cell:
            prev_column_range = utils.closed_range(a_column, p_column)[::-1]

        if row_range[-1] > prev_row_range[-1]:
            print('The row minimum increased')
            for row in range(prev_row_range[-1], row_range[-1]):
                self.__deselect_cells([self.__cells[(row, column)] for column in prev_column_range])

        elif row_range[0] < prev_row_range[0]:
            print('The row maximum decreased')
            for row in range(prev_row_range[0], row_range[0], -1):
                self.__deselect_cells([self.__cells[(row, column)] for column in prev_column_range])

        if row_range[-1] < prev_row_range[-1]:
            print('The row minimum decreased')
            for row in range(row_range[-1], prev_row_range[-1]):
                self.__select_cells([self.__cells[(row, column)] for column in prev_column_range], flip=flip)

        elif row_range[0] > prev_row_range[0]:
            print('The row maximum increased')
            for row in range(row_range[0], prev_row_range[0], -1):
                self.__select_cells([self.__cells[(row, column)] for column in prev_column_range], flip=flip)

        

        
        if column_range[-1] > prev_column_range[-1]:
            print('The column minimum increased')
            for column in range(prev_column_range[-1], column_range[-1]):
                self.__deselect_cells([self.__cells[(row, column)] for row in row_range])

        elif column_range[0] < prev_column_range[0]:
            print('The column maximum decreased')
            for column in range(prev_column_range[0], column_range[0], -1):
                self.__deselect_cells([self.__cells[(row, column)] for row in row_range])


        if column_range[-1] < prev_column_range[-1]:
            print('The column minimum decreased')
            for column in range(column_range[-1], prev_column_range[-1]):
                self.__select_cells([self.__cells[(row, column)] for row in row_range], flip=flip)

        elif column_range[0] > prev_column_range[0]:
            print('The column maximum increased')
            for column in range(column_range[0], prev_column_range[0], -1):
                self.__select_cells([self.__cells[(row, column)] for row in row_range], flip=flip)



    def __set_anchor(self, cell, col = None, add = False, flip=False):        
        if not cell and cell != 0:
            self.__anchor_cell = None
            return

        if type(cell) != Cell:
            cell = self.__cells[utils.normalize_cell_notation(self, cell, col)]

        if flip and cell in self.__selected_cells:
            self.__deselect_cell(cell)
            return

        print('Setting cell ' + repr(cell) + ' to anchor')

        self.__anchor_cell = cell
        cell.config(highlightbackground = 'goldenrod')
        self.__restore_borders(self.__anchor_cell, method='selected')

        self.__set_reel(cell)

        if add and cell not in self.__selected_cells:
            self.__selected_cells.append(cell)

    def __set_reel(self, cell, col=None):
        if type(cell) != Cell:
            cell = self.__cells[utils.normalize_cell_notation(self, cell, col)]

        self.__reel_cell, self.__prev_reel_cell = cell, self.__reel_cell

        print('Reel cell is now ' + repr(self.__reel_cell))

    def __click(self, event):
        print(event.type + ': <Button-1>')
        cell = self.nametowidget(event.widget)

        if not self.focus_get() == event.widget:
            if [cell] == self.__selected_cells:
                cell.focus_set()
            else:
                self.__select_cell(cell, anchor=True, exclusive=True)
                self.god_entry.focus_set()

    def __control_click(self, event):
        print(event.type + ': <Control-Button-1>')
        cell = self.nametowidget(event.widget)
        self.__select_cell(cell, anchor=True, flip=True)

    def __shift_click(self, event):
        print(event.type + ': <Shift-Button-1>')
        self.__deselect_all(but=[self.__anchor_cell])
        self.__select_range(reel = self.nametowidget(event.widget), exclusive = True)

    def __control_shift_click(self, event):
        print(event.type + ': <Control-Shift-Button-1>')
        self.__select_range(reel = self.nametowidget(event.widget))

    def __mouse_motion(self, event):
        self.__set_reel(self.winfo_containing(event.x_root, event.y_root))
        if self.__reel_cell not in self.__cells.values():
            print(str(self.__reel_cell) + ' is out of bounds')
            return
        if self.__prev_reel_cell in self.__cells.values():
            self.__select_range(exclusive=False)
        else:
            print("Coming from out of bounds")
            self.__select_range(exclusive = True)

    def __shift_up(self, event=None): # exclusive for control-shift
        reel_coords = reel_row, reel_col = self.__get_reel_coords()
        offset = (0, 0)
        if reel_row > 0:
            offset = (-1, 0)
        self.__set_reel(self.__cells[tuple(map(add, reel_coords, offset))])
        self.__select_range()

    def __shift_down(self, event=None):
        reel_coords = reel_row, reel_col = self.__get_reel_coords()
        offset = (0, 0)
        if reel_row < self.rows - 1:
            offset = (1, 0)
        self.__set_reel(self.__cells[tuple(map(add, reel_coords, offset))])
        self.__select_range()

    def __shift_left(self, event=None):
        reel_coords = reel_row, reel_col = self.__get_reel_coords()
        offset = (0, 0)
        if reel_col > 0:
            offset = (0, -1)
        self.__set_reel(self.__cells[tuple(map(add, reel_coords, offset))])
        self.__select_range()


    def __shift_right(self, event=None):
        reel_coords = reel_row, reel_col = self.__get_reel_coords()
        offset = (0, 0)
        if reel_col < self.columns - 1:
            offset = (0, 1)
        self.__set_reel(self.__cells[tuple(map(add, reel_coords, offset))])
        self.__select_range()

    def __up(self, event = None, exclusive = True):
        if self.__anchor_cell:
            anchor_row, anchor_col = self.__get_anchor_coords()
            if anchor_row > 0:
                self.__select_cell(self.__cells[(anchor_row - 1, anchor_col)], anchor=True, exclusive=exclusive)
            return 'break'
        else:
            self.__set_anchor(self.__cells[(0, 0)])

    def __down(self, event = None, exclusive = True):
        anchor_row, anchor_col = self.__get_anchor_coords()
        if anchor_row < self.rows - 1:
            self.__select_cell(self.__cells[(anchor_row + 1, anchor_col)], anchor=True, exclusive=exclusive)
        return 'break'

    def __left(self, event = None, exclusive = True, wrap=False):
        anchor_row, anchor_col = self.__get_anchor_coords()
        if anchor_col > 0:
            self.__select_cell(self.__cells[(anchor_row, anchor_col - 1)], anchor=True, exclusive=exclusive)
        elif wrap and anchor_row > 0:
            self.__select_cell(self.__cells[(anchor_row - 1, self.columns - 1)], anchor=True, exclusive=exclusive)

    def __right(self, event = None, exclusive = True, wrap=False):
        anchor_row, anchor_col = self.__get_anchor_coords()
        if anchor_col < self.columns - 1:
            self.__select_cell(self.__cells[(anchor_row, anchor_col + 1)], anchor=True, exclusive=exclusive)
        elif wrap and anchor_row < self.rows - 1:
            self.__select_cell(self.__cells[(anchor_row + 1, 0)], anchor=True, exclusive=exclusive)


    def __on_column_label_click(self, event, exclusive = True, flip=False):
        column = self.__column_labels.index(self.nametowidget(event.widget))
        self.__select_column(column, exclusive = exclusive, flip=flip)
        self.__column_y = event.y_root

    def __on_column_label_control_click(self, event):
        column = self.__column_labels.index(self.nametowidget(event.widget))
        column_cells = [self.__cells[(row, column)] for row in range(self.rows)]
        num_cells_selected = sum([column_cell in self.__selected_cells for column_cell in column_cells])

        if num_cells_selected > self.rows / 2:
            self.__deselect_cells(column_cells)
        else:
            self.__select_cells(column_cells)
            self.__set_anchor(column_cells[0])

        self.__column_y = event.y_root

    def __on_column_label_shift_click(self, event, exclusive=True):
        (_, anchor_col) = self.__get_anchor_coords()
        event_col = self.__column_labels.index(self.nametowidget(event.widget))
        self.__select_range(anchor=(0, anchor_col), reel=(self.rows - 1, event_col), keepanchor = True, exclusive = exclusive)
        self.__column_y = event.y_root

    def __on_column_label_mouse_motion(self, event):
        reel_col = self.__column_labels.index(self.winfo_containing(event.x_root, self.__column_y))
        self.__select_range(exclusive=False, keepanchor=True, reel=(-1, reel_col))

    def __select_column(self, column, exclusive = True, flip = False):
        print('Selecting column ' + str(column))
        self.__select_range(anchor = (0, column), keepanchor = False, reel = (self.rows - 1, column), exclusive = exclusive, flip=True)

    def __on_row_label_click(self, event, exclusive = True):
        row = self.__row_labels.index(self.nametowidget(event.widget))
        self.__select_row(row, exclusive = exclusive)
        self.__row_x = event.x_root

    def __on_row_label_control_click(self, event):
        self.__on_row_label_click(event, exclusive = False)

    def __on_row_label_shift_click(self, event, exclusive=True):
        (anchor_row, _) = self.__get_anchor_coords()
        event_row = self.__row_labels.index(self.nametowidget(event.widget))
        self.__select_range(anchor=(anchor_row, 0), reel=(event_row, -1), keepanchor = True, exclusive = exclusive)

    def __on_row_label_mouse_motion(self, event):
        reel_row = self.__row_labels.index(self.winfo_containing(self.__row_x, event.y_root))
        self.__select_range(exclusive=False, keepanchor=True, reel=(reel_row, -1))

    def __select_row(self, row, exclusive = True):
        print('Selecting row ' + str(row))
        self.__select_range(anchor = (row, 0), keepanchor = False, reel = (row, -1), exclusive = exclusive)

    def __backspace(self, event):
        self.__erase_selected_cell_contents()

    def __delete(self, event):
        self.__erase_selected_cell_contents()

    def __erase_selected_cell_contents(self):
        for cell in self.__selected_cells:
            cell.erase_cell_contents()

    def __control_d(self, event):
        print(self.__anchor_cell.formula_value)
        self.god_entry.focus_set()
        self.__copy_from_anchor_to_selected(self.__anchor_cell)

    def __tab(self, event = None):
        self.__next()

    def __shift_tab(self, event=None):
        self.__prev()

    def __enter_key(self, event):
        print('<Enter>')
        if self.__anchor_cell:
            self.__anchor_cell.focus_set()

    def __shift_enter_key(self, event):
        print('<Shift-Enter>')
        if self.__anchor_cell:
            self.__anchor_cell.focus_set()
            self.__anchor_cell.selection_range(0, 'end')
   

    def __control_enter_key(self, event):
        self.__deselect_cell(self.__anchor_cell)

    def __escape(self, event):
        self.__select_cell(self.__anchor_cell, exclusive=True, anchor=True)

    def __convert_to_inverse(self, event=None):
        sv00 = self.__cells[(0, 0)].sv
        sv01 = self.__cells[(0, 1)].sv
        sv02 = self.__cells[(0, 2)].sv
        sv10 = self.__cells[(1, 0)].sv
        sv11 = self.__cells[(1, 1)].sv
        sv12 = self.__cells[(1, 2)].sv
        sv20 = self.__cells[(2, 0)].sv
        sv21 = self.__cells[(2, 1)].sv
        sv22 = self.__cells[(2, 2)].sv

        determinant = self.__calculate_determinant()

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

    def __export_to_csv(self, event=None):
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

    def __compute_formula(self, value):
        try:
            value = self.__process_formula(value)
        except TypeError:
            value = self.__process_formula(value, number_based = False)
        except SyntaxError:
            # Tell the user their syntax was off
            pass

        return value


    def __update_display_based_on_formula(self, cell):
        cell.computed_value = self.__compute_formula(cell.formula_value.strip())
        cell.config(state='disabled', cursor='plus')
        cell.align_based_on_entry_type()
        cell.mode('computed')

    def __on_cell_focus_in(self, event):
        cell = self.nametowidget(event.widget)
        cell.mode('formula')
        cell.config(state='normal', cursor='xterm')
        cell.icursor(tk.END)

    def __on_exit_cell_typing(self, event=None):
        if self.__guarantee_widget_focus:
            print("Not leaving the cell quite yet!")
            self.nametowidget(event.widget).focus_set()
        else:
            print('leaving cell!')
            cell = self.nametowidget(event.widget)
            cell._absorb_display()
            self.__update_display_based_on_formula(cell)

            return True
            

        self.__guarantee_widget_focus = None


    def __print_determinant(self, event):
        message = None
        try:
            determinant = self.__calculate_determinant()
            message = ('Determinant Calculation', str(determinant))
        except (TypeError, ValueError):
            message = ('Try again', 'Not all entries are integers')
        
        messagebox.showinfo(*message)

    # Assume 3x3 from (0, 0)
    def __calculate_determinant(self, event):
        a00 = float(self.__cells[(0, 0)].get().strip())
        a01 = float(self.__cells[(0, 1)].get().strip())
        a02 = float(self.__cells[(0, 2)].get().strip())
        a10 = float(self.__cells[(1, 0)].get().strip())
        a11 = float(self.__cells[(1, 1)].get().strip())
        a12 = float(self.__cells[(1, 2)].get().strip())
        a20 = float(self.__cells[(2, 0)].get().strip())
        a21 = float(self.__cells[(2, 1)].get().strip())
        a22 = float(self.__cells[(2, 2)].get().strip())
        
        d1 = a11 * a22 - a21 * a12
        d2 = a10 * a22 - a20 * a22
        d3 = a10 * a21 - a20 * a11
        d = a00 * d1 - a01 * d2 + a02 * d3

        return d

    def __on_god_entry_focus_out(self, event):
        if self.__guarantee_focus:
            self.god_entry.focus_set()

        self.__guarantee_focus = False

    def __start_anchor_entry_cursor(self):
        if self.__anchor_cell:
            print('Go to the anchor')
            self.__anchor_cell.formula_value = self.gsv.get()
            self.gsv.set('')
            self.__anchor_cell.focus_set()

    def __copy_from_anchor_to_selected(self, cell = None):
        print('Updating all entries!')
        if cell == self.__anchor_cell:
            for cell in self.__selected_cells:
                if cell == self.__anchor_cell:
                    continue
                cell.formula_value = self.__anchor_cell.formula_value
                self.__update_display_based_on_formula(cell)
                cell.align_based_on_entry_type()

    def __prev(self, event = None):
        if len(self.__selected_cells) == 1:
            self.__left(wrap=True)
        elif len(self.__selected_cells) > 1:
            pass
        else:
            pass

        self.__guarantee_focus = True


    def __next(self, event = None):
        if len(self.__selected_cells) == 1:
            self.__right(wrap=True)
        elif len(self.__selected_cells) > 1:
            pass
        else:
            pass

        self.__guarantee_focus = True

    def __get_anchor_coords(self):
        return self.nametowidget(self.__anchor_cell).coordinates 

    def __get_reel_coords(self):
        return self.nametowidget(self.__reel_cell).coordinates

    def __get_prev_reel_coords(self):
        return self.nametowidget(self.__prev_reel_cell).coordinates

    def __create_macro(self, event):
        filename = filedialog.asksaveasfilename(title='Create Macro File', initialdir=self.program_paths['index'], filetypes=[('Python File', '*.py')])

        if not filename:
            return

        if filename[-3:] != '.py':
            filename += '.py'

        with open(filename, "w+") as python:
            python.write('from pysheets import *\n\n')
            # Comment out examples of ways to use the API

        os.startfile(filename)

    def __run_macro(self, event):
        self.active_macro_import.run(self)

    def __import_macro(self, event):
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


        self.god_entry.bind('<Control-Key-R>', self.__run_macro)
        self.god_entry.bind('<Control-Key-r>', self.__run_macro)









    ###### Public API

    def __getattr__(self, attr):
        pass
        # All
        # Row
        # Column/ Col

    def __getitem__(self, input):
        if type(input) == int:
            pass
        else:
            pass


class CellsView(object):
    def __init__(self, ss, *cell_refs):
        self.__ss = ss

        self.__cells = []
        row = None
        for cell in cell_refs:
            if row:
                column = cell
                if type(column) != int:
                    raise ValueError('Cell reference ' + str(cell) + ' needed to be an integer to couple with the previous argument')
                self.__cells.append(self.__cells[utils.normalize_cell_notation(None, row, column)])
                row = None
            elif type(cell) == Cell:
                self.__cells.append(cell)
            elif type(cell) == str or type(cell) == tuple:
                self.__cells.append(self.__cells[utils.normalize_cell_notation(None, cell)])
            elif type(cell) == int:
                row = cell
            else:
                raise ValueError('Cell reference ' + str(cell) + ' is illegal')

    def keys(self):
        _keys = []
        for cell in self.__cells:
            _keys.append(cell.cell_index)
        return _keys

    def __sub__(self, other):
        cell_refs = list(set(self.keys()) - set(other.keys()))

        return CellsView(self.__ss, *cell_refs)

    def __add__(self, other):
        cell_refs = list(set(self.keys()) + set(other.keys()))

        return CellsView(self.__ss, *cell_refs)

    def apply_formula(self, formula):
        for cell in self.__cells:
            cell.formula = formula

    def get_formula(self):
        pass
        # If the formula of all of the cells is the same, return the formula
        # Else throw an exception or something
        # Or even better return a dictionary of cell to formula

    
    def get_computed_value(self):
        pass

    get_display_value = get_computed_value


