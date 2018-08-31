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