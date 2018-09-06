import tkinter as tk
from tkinter import *

import os

from src.text_editor import TextEditor

class Example(tk.Frame):
    def __init__(self, parent, program_paths):
        tk.Frame.__init__(self, parent)
        self._text_editor = TextEditor(self, program_paths)
        self._text_editor.pack(side="top", fill="both", expand=True)
        self._text_editor.grid(row=0, column=0)

def main():
	root = tk.Tk()

	program_paths = {}
	program_paths['index'] = os.path.dirname(os.path.realpath(__file__))

	Example(root, program_paths = program_paths)
	
	root.mainloop()


if __name__ == '__main__':
	main()