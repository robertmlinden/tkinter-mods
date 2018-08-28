import tkinter as tk
import re

class SimpleTableInput(tk.Frame):
    def __callback(self, var_name):
        var_number = int(re.sub('[^0-9]', '', var_name))
        row = var_number // self.columns
        column = var_number % self.columns
        index = (row, column)
        print(index, '!!')
        print(self._entry[index])
        print(self._svs[index].get())

    def __on_exit(self, e, v):
        entry_widget = self.nametowidget(e)
        try:
            float(v)
            entry_widget.config(justify='right')
        except ValueError:
            entry_widget.config(justify='left')

        entry_widget.config(state='disabled', cursor='plus')

        return True

    def __on_click(self, event):
        entry_widget = self.nametowidget(event.widget)

        self.lift()

        if not self.focus_get() == event.widget:
            entry_widget.config(state='normal', cursor='xterm')
            entry_widget.focus_set()
            entry_widget.selection_range(0, 'end')
            entry_widget.icursor(tk.END)

            if entry_widget.get():
                return 'break'

    def __on_left(self, event):
        entry_widget = self.nametowidget(event.widget)
        if entry_widget.selection_present():
            entry_widget.icursor(entry_widget.index(tk.ANCHOR) + 1)
            

    def __on_right(self, event):
        entry_widget = self.nametowidget(event.widget)
        if entry_widget.selection_present():
            entry_widget.icursor(entry_widget.index(tk.ANCHOR) - 1)

    def __on_motion(self, event):
        entry = self.winfo_containing(event.x_root, event.y_root)
        if entry.config[state] == 'disabled':
            # Do disabled things here
            pass
        print(entry)

    def __init__(self, parent, rows, columns):
        tk.Frame.__init__(self, parent)

        self._entry = {}
        self._svs = {}
        self.rows = rows
        self.columns = columns

        self.bind_all("<B1-Motion>", self.__on_motion)

        # register a command to use for validation
        vcmd = (self.register(self.__on_exit), '%W', '%P')

        # create the table of widgets
        for row in range(self.rows):
            for column in range(self.columns):
                index = (row, column)
                self._svs[index] = tk.StringVar()
                self._svs[index].trace_add('write', lambda var_name, mode, _: self.__callback(var_name))#self.__callback(self._svs[index]))
                e = tk.Entry(self, textvariable=self._svs[index], validate="focusout", validatecommand=vcmd)
                e.grid(row=row, column=column, stick="nsew")
                e.config(justify="left", state='disabled', cursor='plus')
                e.bind("<Button-1>", self.__on_click)
                e.bind("<Left>", self.__on_left)
                e.bind('<Right>', self.__on_right)
                self._entry[index] = e

        # adjust column weights so they all expand equally
        for column in range(self.columns):
            self.grid_columnconfigure(column, weight=1)
        # designate a final, empty row to fill up any extra space
        self.grid_rowconfigure(rows, weight=1)

    def get(self):
        '''Return a list of lists, containing the data in the table'''
        result = []
        for row in range(self.rows):
            current_row = []
            for column in range(self.columns):
                index = (row, column)
                current_row.append(self._entry[index].get())
            result.append(current_row)
        return result

    def _validate(self, P):
        '''Perform input validation. 

        Allow only an empty value, or a value that can be converted to a float
        '''
        if P.strip() == "":
            return True

        try:
            f = float(P)
        except ValueError:
            self.bell()
            return False
        return True

class Example(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent)
        self.table = SimpleTableInput(self, 3, 4)
        self.submit = tk.Button(self, text="Submit", command=self.on_submit)
        self.table.pack(side="top", fill="both", expand=True)
        self.submit.pack(side="bottom")

    def on_submit(self):
        print(self.table.get())