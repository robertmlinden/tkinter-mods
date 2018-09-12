import tkinter as tk
import re

import importlib

from tkinter import messagebox, filedialog

import src.utils as utils
import src.expr_evaluator as arithmetic_evaluator

from PIL import Image, ImageTk

from operator import add

from simpleeval import simple_eval

from copy import copy

import os

import csv

class Spreadsheet(tk.Frame):
    class _Cell(tk.Entry):
        def __init__(self, *args, **kw):
            self.sv = tk.StringVar()

            kw['textvariable'] = self.sv
            self.cell_index = kw.pop('cell_index')
            self.__ss = kw.pop('spreadsheet')

            super().__init__(*args, **kw)

            self.bind('<Control-BackSpace>', self.__control_backspace)
            self.bind("<Left>", self.__typing_left)
            self.bind('<Right>', self.__typing_right)

            self.__formula_value = ''
            self.__computed_value = ''

            self.__mode = 'formula'

            self.__options = {}
            self.__options['selected'] = 'dark green'

        def config(self, **options):
            to_be_popped_keys = []
            for option_key, option_value in options.items():
                self.__options[option_key] = option_value
                if option_key not in self.keys():
                    to_be_popped_keys.append(option_key)
                if option_key == 'highlightbackground' and option_value == 'selected':
                    options[option_key] = self.__options['selected']

            for key in to_be_popped_keys:
                options.pop(key)

            super().config(**options)

        def __update_computed_based_on_formula(self):
            self.computed_value = self.__ss.compute_formula(self.formula_value.strip())
            print('computed value = ' + str(self.computed_value))
            self.__align_based_on_entry_type()
            self.mode('computed')

        def _absorb_display(self):
            if self.__mode == 'formula':
                self.formula_value = self.sv.get()
            else:
                self.computed_value = self.sv.get()

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
            print('Setting formula value')
            self.__formula_value = value
            self.__update_computed_based_on_formula()
            self.__update_display()
            self.__align_based_on_entry_type()


        @property
        def computed_value(self):
            return self.__computed_value

        @computed_value.setter
        def computed_value(self, value):
            self.__computed_value = value

        
        @property
        def coordinates(self):
            return utils.normalize_cell_notation(self.__ss.rows, self.__ss.columns, self.cell_index)

        def __set_background(self, color_string):
            self.config(background=color_string, disabledbackground=color_string)

        def __update_display(self):
            if self.__mode == 'formula':
                print('update formula')
                self.sv.set(self.__formula_value)
            else:
                print('update compute')
                self.sv.set(str(self.__computed_value))

        def mode(self, _mode):
            self.__mode = _mode
            self.__update_display()

        def __align_based_on_entry_type(self):
            try:
                float(self.computed_value)
                self.config(justify='right')
            except ValueError:
                self.config(justify='left')

        def __repr__(self):
            return self.cell_index


    def __init__(self, program_paths, parent, rows, columns):
        tk.Frame.__init__(self, parent)

        self.program_paths = program_paths

        self.containing_frame = parent

        self.rows = rows
        self.columns = columns
        self.cols = columns

        self.__cells = [[None for j in range(self.columns)] for i in range(self.rows)]

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
        image_label.image=ph  #need to keep the reference of image to avoid garbage collection

        image_label.bind('<Button-1>', self.__select_all_cells)
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

        # create the table of cell widgets
        for row in range(self.rows):
            for column in range(self.columns):
                index = (row, column)
                c = self._Cell(self, cell_index = utils.get_cell_index(row, column), spreadsheet = self)
                print(utils.get_cell_index(row, column))
                c.grid(row=row+1, column=column+1, stick="nsew")
                c.config(justify="left", state='disabled', cursor='plus', highlightthickness = 1, highlightbackground = 'ghost white',
                            disabledbackground='white', highlightcolor = 'goldenrod', fg='black', disabledforeground='#101010')
                c.bind("<Button-1>", self.__click_cell)
                c.bind("<Control-Button-1>", self.__control_click_cell)
                c.bind("<Shift-Button-1>", self.__shift_click_cell)
                c.bind("<Control-Shift-Button-1>", self.__control_shift_click_cell)
                c.bind("<B1-Motion>", self.__on_cell_mouse_motion)
                c.bind('<Tab>', self.__next_cell)
                c.bind('<Shift-Tab>', self.__prev_cell)
                c.bind('<Return>', self.__down)
                c.bind('<Down>', self.__down)
                c.bind('<Up>', self.__up)
                c.bind('<Escape>', self.__escape)
                c.bind('<FocusIn>', self.__on_cell_focus_in)
                c.bind('<FocusOut>', self.__on_exit_cell_typing)
                c.bind('<Control-Key-d>', self.__control_d)
                c.bind('<Control-Key-E>', self.__export_to_csv)
                c.bind('<Control-Key-a>', self.__select_all_cells)

                self.__cells[row][column] = c

        self.__cells_flattened = self.__flatten_cells()

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
        self.god_entry.bind('<Control-Key-a>', self.__select_all_cells)
        self.god_entry.bind('<Control-Key-E>', self.__export_to_csv)
        self.god_entry.bind('<Escape>', self.__escape)
        self.god_entry.bind('<Control-Key-m>', self.__create_macro)
        self.god_entry.bind('<Control-Key-l>', self.__import_macro)

        self.__guarantee_focus = False

        self.containing_frame.focus_set()

    def __flatten_cells(self):
        return [cell for row in self.__cells for cell in row]

    def __select_all_cells(self, event):
        print('Selecting all cells in the grid!')
        self.__select_range(exclusive=True, anchor='A1', reel=(-1, -1))
        return 'break'

    def __select_cells(self, cells, exclusive=False, flip=False):
        if exclusive:
            self.__deselect_all_cells()
        
        [self.__select_cell(cell, exclusive=False, flip=flip) for cell in cells]

    def __select_cell(self, cell, anchor=False, exclusive=False, flip=False):
        if exclusive:
            self.__deselect_all_cells()          

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
            cell.config(highlightbackground = 'selected')
            self.__selected_cells.append(cell)

        self.god_entry.focus_set()


    def __deselect_all_cells(self, but=[]):
        print('deselect all')

        self.__deselect_cells([cell for cell in self.__selected_cells if cell not in but])

    def __deselect_cells(self, cells):
        [self.__deselect_cell(cell) for cell in cells]

    def __deselect_cell(self, cell):
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

    def __select_range(self, keepanchor = True, exclusive = False, flip=False, anchor = None, reel = None):
        #print('Selecting ' + ('exclusive' if exclusive else '') +  ' range: ', end='')

        if exclusive:
            but = [self.__anchor_cell] if keepanchor else []
            self.__deselect_all_cells(but=but)

        if anchor:
            self.__set_anchor(anchor, add=True)

        if reel:
            self.__set_reel_cell(reel)
        
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
                self.__deselect_cells([self.__cells[row][column] for column in prev_column_range])

        elif row_range[0] < prev_row_range[0]:
            print('The row maximum decreased')
            for row in range(prev_row_range[0], row_range[0], -1):
                self.__deselect_cells([self.__cells[row][column] for column in prev_column_range])

        if row_range[-1] < prev_row_range[-1]:
            print('The row minimum decreased')
            for row in range(row_range[-1], prev_row_range[-1]):
                self.__select_cells([self.__cells[row][column] for column in prev_column_range], flip=flip)

        elif row_range[0] > prev_row_range[0]:
            print('The row maximum increased')
            for row in range(row_range[0], prev_row_range[0], -1):
                self.__select_cells([self.__cells[row][column] for column in prev_column_range], flip=flip)

        

        
        if column_range[-1] > prev_column_range[-1]:
            print('The column minimum increased')
            for column in range(prev_column_range[-1], column_range[-1]):
                self.__deselect_cells([self.__cells[row][column] for row in row_range])

        elif column_range[0] < prev_column_range[0]:
            print('The column maximum decreased')
            for column in range(prev_column_range[0], column_range[0], -1):
                self.__deselect_cells([self.__cells[row][column] for row in row_range])


        if column_range[-1] < prev_column_range[-1]:
            print('The column minimum decreased')
            for column in range(column_range[-1], prev_column_range[-1]):
                self.__select_cells([self.__cells[row][column] for row in row_range], flip=flip)

        elif column_range[0] > prev_column_range[0]:
            print('The column maximum increased')
            for column in range(column_range[0], prev_column_range[0], -1):
                self.__select_cells([self.__cells[row][column] for row in row_range], flip=flip)



    def __set_anchor(self, cell, col = None, add = False, flip=False):        
        if not cell and cell != 0:
            self.__anchor_cell = None
            return

        if type(cell) != self._Cell:
            row, column = utils.normalize_cell_notation(self.rows, self.columns, cell, col)
            cell = self.__cells[row][column]

        if flip and cell in self.__selected_cells:
            self.__deselect_cell(cell)
            return

        print('Setting cell ' + repr(cell) + ' to anchor')

        if self.__anchor_cell in self.__selected_cells:
            self.__anchor_cell.config(highlightbackground = 'selected')
        self.__anchor_cell = cell
        self.__anchor_cell.config(highlightbackground = 'goldenrod')

        self.__set_reel_cell(cell)

        if add and cell not in self.__selected_cells:
            self.__selected_cells.append(cell)

    def __set_reel_cell(self, cell, col=None):
        if type(cell) != self._Cell:
            row, column = utils.normalize_cell_notation(self.rows, self.columns, cell, col)
            if type(row) == int and type(column) == int:
                cell = self.__cells[row][column]

        self.__reel_cell, self.__prev_reel_cell = cell, self.__reel_cell

        #print('Reel cell is now ' + repr(self.__reel_cell))

    def __init_click_events(self, event):
        self.__min_x = self.__cells[0][0].winfo_rootx()
        self.__min_y = self.__cells[0][0].winfo_rooty()

        row, column = utils.convert_coordinates_from_negative(self.rows, self.cols, -1, -1)
        last_cell = self.__cells[row][column]

        self.__max_x = last_cell.winfo_rootx() + last_cell.winfo_width()
        self.__max_y = last_cell.winfo_rooty() + last_cell.winfo_height()

    def __click_cell(self, event):
        print(event.type + ': <Button-1>')
        cell = self.nametowidget(event.widget)

        self.__init_click_events(event)

        if not self.focus_get() == event.widget:
            if [cell] == self.__selected_cells:
                cell.focus_set()
            else:
                self.__select_cell(cell, anchor=True, exclusive=True)
                self.god_entry.focus_set()

    def __control_click_cell(self, event):
        print(event.type + ': <Control-Button-1>')
        self.__init_click_events(event)
        cell = self.nametowidget(event.widget)
        self.__select_cell(cell, anchor=True, flip=True)

    def __shift_click_cell(self, event):
        print(event.type + ': <Shift-Button-1>')
        self.__init_click_events(event)
        self.__deselect_all_cells(but=[self.__anchor_cell])
        self.__select_range(reel = self.nametowidget(event.widget), exclusive = True)

    def __control_shift_click_cell(self, event):
        print(event.type + ': <Control-Shift-Button-1>')
        self.__init_click_events(event)
        self.__select_range(reel = self.nametowidget(event.widget))


    def __on_cell_mouse_motion(self, event):
        if not (self.focus_get() == self.containing_frame or self.focus_get() == self.god_entry):
            return

        x = event.x_root
        if x <= self.__min_x:
            x = self.__min_x + 1
        elif x >= self.__max_x:
            x = self.__max_x - 1

        y = event.y_root
        if y <= self.__min_y:
            y = self.__min_y + 1
        elif y >= self.__max_y:
            y = self.__max_y - 1

        
        self.__set_reel_cell(self.winfo_containing(x, y))
        if self.__reel_cell not in self.__cells_flattened:
            print(str(self.__reel_cell) + ' is out of bounds')
            return
        if self.__prev_reel_cell in self.__cells_flattened:
            self.__select_range(exclusive=False)
        else:
            print("Coming from out of bounds")
            self.__select_range(exclusive = True)

    def __shift_up(self, event=None): # exclusive for control-shift
        reel_coords = reel_row, reel_col = self.__get_reel_coords()
        offset = (0, 0)
        if reel_row > 0:
            offset = (-1, 0)
            row, column = tuple(map(add, reel_coords, offset))
        self.__set_reel_cell(self.__cells[row][column])
        self.__select_range()

    def __shift_down(self, event=None):
        reel_coords = reel_row, reel_col = self.__get_reel_coords()
        offset = (0, 0)
        if reel_row < self.rows - 1:
            offset = (1, 0)
            row, column = tuple(map(add, reel_coords, offset))
        self.__set_reel_cell(self.__cells[row][column])
        self.__select_range()

    def __shift_left(self, event=None):
        reel_coords = reel_row, reel_col = self.__get_reel_coords()
        offset = (0, 0)
        if reel_col > 0:
            offset = (0, -1)
            row, column = tuple(map(add, reel_coords, offset))
        self.__set_reel_cell(self.__cells[row][column])
        self.__select_range()


    def __shift_right(self, event=None):
        reel_coords = reel_row, reel_col = self.__get_reel_coords()
        offset = (0, 0)
        if reel_col < self.columns - 1:
            offset = (0, 1)
            row, column = tuple(map(add, reel_coords, offset))
        self.__set_reel_cell(self.__cells[row][column])
        self.__select_range()

    def __up(self, event = None, exclusive = True):
        if self.__anchor_cell:
            anchor_row, anchor_col = self.__get_anchor_coords()
            if anchor_row > 0:
                self.__select_cell(self.__cells[anchor_row - 1][anchor_col], anchor=True, exclusive=exclusive)
            return 'break'
        else:
            self.__set_anchor(self.__cells[(0, 0)])

    def __down(self, event = None, exclusive = True):
        anchor_row, anchor_col = self.__get_anchor_coords()
        if anchor_row < self.rows - 1:
            self.__select_cell(self.__cells[anchor_row + 1][anchor_col], anchor=True, exclusive=exclusive)
        return 'break'

    def __left(self, event = None, exclusive = True, wrap=False):
        anchor_row, anchor_col = self.__get_anchor_coords()
        if anchor_col > 0:
            self.__select_cell(self.__cells[anchor_row][anchor_col - 1], anchor=True, exclusive=exclusive)
        elif wrap and anchor_row > 0:
            self.__select_cell(self.__cells[anchor_row - 1][self.columns - 1], anchor=True, exclusive=exclusive)

    def __right(self, event = None, exclusive = True, wrap=False):
        anchor_row, anchor_col = self.__get_anchor_coords()
        if anchor_col < self.columns - 1:
            self.__select_cell(self.__cells[anchor_row][anchor_col + 1], anchor=True, exclusive=exclusive)
        elif wrap and anchor_row < self.rows - 1:
            self.__select_cell(self.__cells[anchor_row + 1][0], anchor=True, exclusive=exclusive)


    def __on_column_label_click(self, event, exclusive = True, flip=False):
        column = self.__column_labels.index(self.nametowidget(event.widget))
        self.__select_column(column, exclusive = exclusive, flip=flip)
        self.__column_y = event.y_root

    def __on_column_label_control_click(self, event):
        column = self.__column_labels.index(self.nametowidget(event.widget))
        column_cells = [self.__cells[row][column] for row in range(self.rows)]
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
        column_label = self.winfo_containing(event.x_root, self.__column_y)
        if column_label not in self.__column_labels:
            return
        reel_col = self.__column_labels.index(column_label)
        self.__select_range(exclusive=False, keepanchor=True, reel=(-1, reel_col))

    def __select_column(self, column, exclusive = True, flip = False):
        print('Selecting column ' + str(column))
        self.__select_range(anchor = (0, column), keepanchor = False, reel = (self.rows - 1, column), exclusive = exclusive, flip=True)

    def __on_row_label_click(self, event, exclusive = True, flip=False):
        row = self.__row_labels.index(self.nametowidget(event.widget))
        self.__select_row(row, exclusive = exclusive, flip=flip)
        self.__row_x = event.x_root

    def __on_row_label_control_click(self, event):
        row = self.__row_labels.index(self.nametowidget(event.widget))
        row_cells = [self.__cells[row][column] for column in range(self.columns)]
        num_cells_selected = sum([row_cell in self.__selected_cells for row_cell in row_cells])

        if num_cells_selected > self.columns / 2:
            self.__deselect_cells(row_cells)
        else:
            self.__select_cells(row_cells)
            self.__set_anchor(row_cells[0])

        self.__row_x = event.x_root

    def __on_row_label_shift_click(self, event, exclusive=True):
        (anchor_row, _) = self.__get_anchor_coords()
        event_row = self.__row_labels.index(self.nametowidget(event.widget))
        self.__select_range(anchor=(anchor_row, 0), reel=(event_row, self.columns - 1), keepanchor = True, exclusive = exclusive)
        self.__row_x = event.x_root

    def __on_row_label_mouse_motion(self, event):
        row_label = self.winfo_containing(self.__row_x, event.y_root)
        if row_label not in self.__row_labels:
            return
        reel_row = self.__row_labels.index(row_label)
        self.__select_range(exclusive=False, keepanchor=True, reel=(reel_row, -1))

    def __select_row(self, row, exclusive = True, flip = False):
        print('Selecting row ' + str(row))
        self.__select_range(anchor = (row, 0), keepanchor = False, reel = (row, self.columns - 1), exclusive = exclusive, flip=True)

    def __backspace(self, event):
        self.__erase_selected_cell_contents()

    def __delete(self, event):
        self.__erase_selected_cell_contents()

    def __erase_selected_cell_contents(self):
        for cell in self.__selected_cells:
            cell.erase_cell_contents()

    def __control_d(self, event):
        print('Control-d')
        self.__anchor_cell._absorb_display()
        self.god_entry.focus_set()
        self.__copy_from_anchor_to_selected(self.__anchor_cell)

    def __tab(self, event = None):
        self.__next_cell()

    def __shift_tab(self, event=None):
        self.__prev_cell()

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

    def __get_formatted_value(self, match):
        row, column = utils.normalize_cell_notation(self.rows, self.columns, match)
        cell = self.__cells[row][column]
        return cell.computed_value

    def __cell_convert(self, value):
        return re.sub(r'\[.*?\]', lambda match: self.__get_formatted_value(match[0]), value)

    def compute_formula(self, formula):
        try:
            if formula and formula[0] == '=' and len(formula) > 1:
                converted_value = self.__cell_convert(formula[1])
                return simple_eval(converted_value)
            else:
                return formula
        except SyntaxError:
            print("Syntax Error")

        return value

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
            print('leaving cell ' + repr(self.nametowidget(event.widget)))
            cell = self.nametowidget(event.widget)
            cell._absorb_display()
            cell.config(state='disabled', cursor='plus')

            return True
            

        self.__guarantee_widget_focus = None

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

    def __prev_cell(self, event = None):
        if len(self.__selected_cells) == 1:
            self.__left(wrap=True)
        elif len(self.__selected_cells) > 1:
            pass
        else:
            pass

        self.__guarantee_focus = True


    def __next_cell(self, event = None):
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
        # Detect changes to file
        importlib.reload(self.active_macro_import)

        self.active_macro_import.run(self)

    def __import_macro(self, event):
        active_macro_fullpath = filedialog.askopenfilename(title='Create Macro File', initialdir=os.path.join(self.program_paths['index'], 'macros'), filetypes=[('Python File', '*.py')])

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


        self.god_entry.bind('<Control-Key-R>', self.__run_macro)
        self.god_entry.bind('<Control-Key-r>', self.__run_macro)


    ###### Public API

    def __getattr__(self, attr):
        if attr.lower() == 'all' or attr.lower() == 'everything':
            print('all')
            return celllist([self.__cells[row][column] for column in range(self.columns) for row in range(self.rows)])

        elif attr.lower() == 'row':
            print('row')
            return cellList([celllist([self.__cells[row][column] for column in range(self.columns)]) for row in range(self.rows)])
        elif attr.lower() == 'column' or attr.lower() == 'col':
            print('col')
            return celllist([celllist([self.__cells[row][column] for row in range(self.rows)]) for column in range(self.columns)])
        elif attr.lower() == 'selected':
            print('selected')
            return celllist(self.__selected_cells)
        else:
            raise AttributeError('"' + attr + '" is not an attribute of ' + repr(self) + '\n' + 
                                    'Attributes include all/everything, row, column/col')   

    def __getitem__(self, index):
        if isinstance(index, slice):
            ifnone = lambda a, b: b if a is None else a

            row_start = 0
            row_stop = self.rows

            col_start = 0
            col_stop = self.columns

            if index.start:
                print('start')
                if type(index.start) == int:
                    row_start = index.start
                elif type(index.start) == str:
                    row_start, col_start = utils.get_cell_coordinates(index.start)
                elif type(index.start) == tuple:
                    row_start, col_start = index

            if index.stop:
                print('stop')
                if type(index.stop) == int:
                    row_stop = index.stop
                elif type(index.stop) == str:
                    row_stop, col_stop = utils.get_cell_coordinates(index.stop)
                elif type(index.stop) == tuple:
                    row_stop, col_stop = index

            row_step = 1 if row_start < row_stop else -1
            col_step = 1 if col_start < col_stop else -1

            cell_step = ifnone(index.step, 1)

            row_range = range(row_start, row_stop, row_step)
            col_range = range(col_start, col_stop, col_step)
            return celllist([self.__cells[row][column] for column in col_range for row in row_range][::cell_step])

        elif isinstance(index, str):
            if index.isalpha():
               column = utils.get_column_letters(index)
               return celllist([self.__cells[row][column] for row in range(self.rows)])
            elif index.isnumeric():
                row = int(index) - 1
                return celllist([self.__cells[row][column] for column in range(self.columns)])
            else:
                row, column = utils.normalize_cell_notation(self.rows, self.columns, index)
                return self.__cells[row][column]

        elif isinstance(index, tuple):
            row, column = index
            return celllist(self.__cells[row][column])

        elif isinstance(index, int):
            row = index
            return celllist([self.__cells[row][column] for column in range(self.columns)])


class celllist(list):

    @property
    def formula_value(self):
        formulas = []
        for item in self:
            if isinstance(item, celllist):
                formulas.append([subitem.formula_value for subitem in item])
            else: # isinstance(item, _Cell)
                formulas.append(item.formula_value)
            return formulas


    @formula_value.setter
    def formula_value(self, formula_value):
        for item in self:
            item.formula_value = formula_value

    def config(self, **options):
        for item in self:
            item.config(**options)
    
    def __sub__(self, other):
        return [item for item in self if item not in other]
    
    def __getitem__(self, index):
        if type(index) == str:
            index = utils.get_column_index(index)

        elif type(index) == slice:
            start = utils.get_column_index(index.start) if index.start else index.start
            stop = utils.get_column_index(index.stop) if index.stop else index.stop
            
            slce = slice(start, stop, index.step)
            return super().__getitem__(slce)

        return super().__getitem__(index)
