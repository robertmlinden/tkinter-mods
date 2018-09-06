import tkinter as tk
from tkinter import *

class TextEditor(tk.Text):
	def __init__(self, parent, program_paths = None, *args, **kwargs):
		tk.Text.__init__(self, *args, **kwargs)
		self.containing_frame = parent
		self.program_paths = program_paths
		print(dir(self))
		print([thing for thing in dir(self) if 'weight' in thing])