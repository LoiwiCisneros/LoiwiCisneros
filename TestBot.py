from tkinter import *

root = Tk()

myLabel1 = Label(root, text="Hello world!").grid(row=0, column=0)
myLabel2 = Label(root, text="My name is Loiwi Cisneros").grid(row=1, column=0)

myButton = Button(root, text="Click me!", state=DISABLED).grid(row=2, column=0)

root.mainloop()