from tkinter import *

from tkinter.filedialog  import *

root = Tk()

filename=askopenfilename(parent = root)  #打开
#filename=asksaveasfilename(parent = root,defaultextension='.xls')  #另存


print(filename) 

root.mainloop()