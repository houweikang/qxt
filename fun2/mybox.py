import tkinter as tk
from tkinter import filedialog


class MyBox:

    def __init__(self):
        width, height = 400, 150
        self.root = tk.Tk()
        self.root.width = width
        self.root.height = height
        self.canvas = tk.Canvas(self.root, height=height, width=width)
        self.canvas.pack()

    def my_label(self, text, font=('helvetica', 18),
                 relx=0.5, rely=0.15, relwidth=0.8, relheight=0.2, anchor="center"):
        self.label = tk.Label(self.canvas, text=text)
        self.label.config(font=font)
        self.label.place(relx=relx, rely=rely,
                         relwidth=relwidth, relheight=relheight, anchor=anchor)

    def my_entry(self, default='', font=('helvetica', 15), justify='center',
                 relx=0.5, rely=0.4, relwidth=0.8, relheight=0.3, anchor="center"):
        self.content = tk.StringVar()
        self.content.set(default)
        self.entry = tk.Entry(self.canvas, font=font,
                              justify=justify, textvariable=self.content)
        self.entry.place(relx=relx, rely=rely,
                         relwidth=relwidth, relheight=relheight, anchor=anchor)
        self.entry.focus_set()
        if default:
            self.entry.selection_range(0, 100)
        return self.entry

    def my_button(self, text="Yes", bg='#cccccc', font=('helvetica', 18),
                  relx=0.5, rely=0.8, relwidth=0.3, relheight=0.3, anchor="center"):
        self.buttonInputBox = tk.Button(
            self.canvas, text=text, bg=bg, font=font)
        self.buttonInputBox.place(relx=relx, rely=rely,
                                  relwidth=relwidth, relheight=relheight, anchor=anchor)
        return self.buttonInputBox

    def my_inputbox(self, label_text, entry_default):
        value = ['']
        self.my_label(text=label_text)
        entry = self.my_entry(default=entry_default)
        button = self.my_button()
        button['command'] = lambda: self.inputbox_def(value, entry.get())
        self.root.mainloop()
        return value[0]

    def inputbox_def(self, value, other_value):
        value[0] = other_value
        self.root.destroy()

    def my_onlyok(self, label_text):
        self.my_label(text=label_text, rely=0.4)
        button = self.my_button(text='OK')
        button['command'] = lambda: self.root.destroy()
        self.root.mainloop()

    def file_path(self):
        self.root.withdraw()
        return filedialog.askopenfilename()


def inputbox(mylabel, default):
    return MyBox().my_inputbox(mylabel, default)


def only_ok(default='Ok!'):
    return MyBox().my_onlyok(default)


def file_path():
    return MyBox().file_path()


if __name__ == '__main__':
    print(inputbox('haha', 123))
    only_ok('haha')