import tkinter as tk
from tkinter.constants import INSERT
from tkinter.filedialog import askopenfilename, askdirectory
from tkinter.font import ROMAN
from xml.dom.minidom import Text

'''
Create a row with a file selector and label to display the selected button
'''


class FileSelectorWidget(tk.Frame):
    def select_file(self, parent, label):
        filename = askopenfilename()
        label.delete(1.0, tk.END)
        label.insert(INSERT, filename)
        parent.filename = filename

    def __init__(self, parent, text, row):
        tk.Frame.__init__(self, parent)
        self.parent = parent

        label_text = tk.Text(self.parent, width=25, height=1)
        label_text.insert(INSERT, text)
        file_name = tk.Text(self.parent, width=100, height=1)
        file_name.insert(INSERT, 'test')
        button = tk.Button(self.parent, text='...', width=5,
                           command=lambda: self.select_file(self, file_name))
        label_text.grid(row=row, column=0)
        file_name.grid(row=row, column=1)
        button.grid(row=row, column=2)


class ProcessButton(tk.Frame):
    def __init__(self, parent, process_function, row):
        tk.Frame.__init__(self, parent)
        self.parent = parent

        button = tk.Button(self.parent, text='Process', width=10,
                           command=process_function)
        button.grid(row=row)


class MainApplication(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        self.parent.title('Boilerplate Filler')
        self.excel_template = FileSelectorWidget(self, 'Excel File', 0)
        self.word_template = FileSelectorWidget(self, 'Word Template', 1)
        self.process_button = ProcessButton(self, lambda: print('test'), 2)

    def process_files(main_application):
        excelpath = main_application.excel_template.filename
        wordpath = main_application.word_template.filename
        # TODO export main function and plug in these values and run


if __name__ == "__main__":
    root = tk.Tk()
    MainApplication(root).pack()
    root.mainloop()
