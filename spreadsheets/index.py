from src.tkspreadsheet import Example
import tkinter as tk
import os

def main():
	root = tk.Tk()

	program_paths = {}
	program_paths['index'] = os.path.dirname(os.path.realpath(__file__))
	program_paths['icons'] = os.path.join(program_paths['index'], 'resources', 'icons')

	Example(root, program_paths = program_paths).pack(side="top", fill="both", expand=True)
	root.mainloop()

if __name__ == '__main__':
	main()

class Example(tk.Frame):
    def __init__(self, parent, program_paths):
        tk.Frame.__init__(self, parent)
        self._spreadsheet = Spreadsheet(program_paths, self, 10, 10)
        self.submit = tk.Button(self, text="Submit", command=self.on_submit)
        self._spreadsheet.pack(side="top", fill="both", expand=True)
        self.submit.pack(side="bottom")

    def on_submit(self):
        print_(self.spreadsheet.get())