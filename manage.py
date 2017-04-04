import shutil
import glob
import xlrd
import xlwt
from xlutils.copy import copy
import tkinter as tk
from tkinter import *
from tkinter import filedialog

class ParentWindow(Frame):
    def __init__(self, master, *args, **kwargs):
        Frame.__init__(self, master, *args, **kwargs)
        self.master = master
        self.master.minsize(700,220) #(width, height)
        self.master.maxsize(700,220)
        self.master.title("Auto Excel")
        self.master.configure(bg="#F0F0F0")
        arg = self.master
        load_gui(self)




def load_gui(self):
    self.lbl_excelTag = tk.Label(self.master,text='Excel file:')
    self.lbl_excelTag.grid(row=0,column=0,columnspan=1,padx=(27,0),pady=(10,0),sticky=N+W)
    self.lbl_link = tk.Label(self.master,text='Link Column')
    self.lbl_link.grid(row=1,column=0,columnspan=1,padx=(27,0),pady=(10,0),sticky=N+W)
    self.lbl_output= tk.Label(self.master,text='Destination Column')
    self.lbl_output.grid(row=2,column=0,columnspan=1,padx=(27,0),pady=(10,0),sticky=N+W)
    self.txt_excel_location = tk.Label(self.master,text='', bg="#FFFFFF", borderwidth = '2px',width='50')
    self.txt_excel_location.grid(row=0,column=2,rowspan=1,columnspan=3,padx=(30,40),pady=(10,0),sticky=N+E)
    self.txt_link = tk.Entry(self.master, bg="#FFFFFF",bd = '2px',width='58')
    self.txt_link.grid(row=1,column=2,rowspan=1,columnspan=3,padx=(30,40),pady=(10,0),sticky=N+E)
    self.txt_output = tk.Entry(self.master, bg="#FFFFFF",bd = '2px',width='58')
    self.txt_output.grid(row=2,column=2,rowspan=1,columnspan=3,padx=(30,40),pady=(10,0),sticky=N+E)
    self.btn_srctag = tk.Button(self.master,text="Choose Source",command = lambda:select_source(self))
    self.btn_srctag.grid(row=0,column=5,rowspan=1,columnspan=2,padx=(30,40),pady=(10,0),sticky=N+E)
    self.btn_submit = tk.Button(self.master, text = "Do Chantel's Work",height = '4',width = '18',bd='3px',command = lambda:excel_main(self))
    self.btn_submit.grid(row = 4, column = 2,rowspan=2,columnspan=2,padx=(30,40),pady=(25,0),sticky=N+W)
    self.btn_submit = tk.Button(self.master, text = "CANCEL",height = '4',width = '15',bd='3px')
    self.btn_submit.grid(row = 4, column = 4,rowspan=2,columnspan=2,padx=(30,40),pady=(25,0),sticky=N+W)

def select_source(self):
    src_folder = filedialog.askopenfilename()
    self.txt_excel_location.config(text = src_folder)

def excel_main(self):
    file_location = self.txt_excel_location.cget("text")
    workbook = xlrd.open_workbook(file_location)
    wbcopy = copy(workbook)
    wbcopy_sheet = wbcopy.get_sheet(0)
    sheet = workbook.sheet_by_index(0)
    input_col = self.txt_link.get()
    output_col = self.txt_output.get()
    input_col = letter_to_num(input_col.lower())
    output_col = letter_to_num(output_col.lower()) 
    #print (inputt)
    #print (output)
    for row in range(sheet.nrows):
        input_value = sheet.cell_value(row,input_col)
        if input_value[0:4] == 'http':
            value_str = input_value.split('.com')
            input_value = value_str[0]
            value_str = input_value.split('.net')
            input_value = value_str[0]
            value_str = input_value.split('.org')
            input_value = value_str[0]
            value_str = input_value.split('//')
            input_value = value_str[1]
            if input_value[0:4] == 'www.':
                value_str = input_value.split('www.')
                output_value = value_str[1]
            else:
                output_value = input_value
        else:
            output_value = input_value
        output_value = output_value.title()
        if row == 0:
            output_value = sheet.cell_value(row,input_col)
        wbcopy_sheet.write(row,output_col,output_value)
        #print (output_value.title())
    wbcopy.save('Outputexcel.xlsx')
    
def letter_to_num(letter):
    if letter == 'a':
        return 0
    elif letter == 'b':
        return 1
    elif letter == 'c':
        return 2
    elif letter == 'd':
        return 3
    elif letter == 'e':
        return 4
    elif letter == 'f':
        return 5
    elif letter == 'g':
        return 6
    elif letter == 'h':
        return 7
    elif letter == 'i':
        return 8
    elif letter == 'j':
        return 9
    elif letter == 'k':
        return 10
    
    
if __name__ == "__main__":
    root = tk.Tk()
    App = ParentWindow(root)
    root.mainloop()
