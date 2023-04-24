import os
import pathlib
from datetime import date
from tkinter import *
from tkinter import filedialog, messagebox
from tkinter.ttk import Combobox

import openpyxl  # Excel
import xlrd  # Excel .xls only
from PIL import Image, ImageTk

os.chdir(os.path.dirname(__file__))


background = "#06283D"
framebg = "#EDEDED"
framefg = "#06283D"

class StudentRegistrationSystem(Tk):
    
    def __init__(self):
        super().__init__()
        
        self.title("Student Registration System")
        self.geometry("1200x675+210+50")
        self.resizable(False,False)
        self.config(bg=background)
        #self.email = Label(text="Email: ",
        #   width=10,
        #    height=3,      
        #   bg="#f0687c",
        #    anchor='e')
        #self.email.pack(side=TOP, fill=X)
        
        # Student details
        self.lframe = LabelFrame(
            text="Student's Details",
            font=20,
            bd=2,
            width=900,
            bg=framebg,
            fg=framefg,
            height=250,
            relief=GROOVE)
        self.lframe.place(relx=0.05, rely=0.2)
        
        # Parent details
        self.pframe = LabelFrame(
            text="Parent's Details",
            font=20,
            bd=2,
            width=900,
            bg=framebg,
            fg=framefg,
            height=220,
            relief=GROOVE)
        self.pframe.place(relx=0.05, rely=0.6)
        
        self.__create_student_base()
        self.__create_top_widgets()
        self.__create_date_registration()
        self.__create_details()
    
    def __create_details(self):
        full_name = Label(self.lframe,text="Full name:", font="Arial 13", bg=framebg, fg=framefg)
        full_name.place(relx=0.07, rely=0.25)
        
        self.Name = StringVar()
        self.name_entry = Entry()
        
        full_name = Label(self.lframe,text="Date of Birth:", font="Arial 13", bg=framebg, fg=framefg)
        full_name.place(relx=0.06, rely=0.14)
        
        full_name = Label(self.lframe,text="Gender:", font="Arial 13", bg=framebg, fg=framefg)
        full_name.place(relx=0.075, rely=0.4)
        
        full_name = Label(self.lframe,text="Class:", font="Arial 13", bg=framebg, fg=framefg)
        full_name.place(relx=0.075, rely=0.45)
        
        full_name = Label(self.lframe,text="Religion:", font="Arial 13", bg=framebg, fg=framefg)
        full_name.place(relx=0.2, rely=0.32)
        
        full_name = Label(self.lframe,text="Date Skills:", font="Arial 13", bg=framebg, fg=framefg)
        full_name.place(relx=0.2, rely=0.4)

    # Top widgets
    def __create_top_widgets(self):
        self.label = Label(
            text="STUDENT REGISTRATION",
            width=12,
            height=3,
            bg="#c36464",
            fg="#fff",
            font="Arial 18 bold",
            )
        self.label.pack(side=TOP, fill=X)
        
        self.Search = StringVar()
        self.enry = Entry(textvariable=self.Search, 
                width=12,
                bd=2,
                font="arial 20")
        self.enry.place(relx=0.64,rely=0.05)
    
        # search box to update
        self.searchimage = PhotoImage(file="Images/search.png")
        self.search_button = Button(
            text="Search",
            compound=LEFT,
            width=123,
            height=2,
            bd=2,
            bg="#68ddfa",
            fg="black",
            image=self.searchimage,
            font="Arial 13 bold")
        self.search_button.place(relx=0.82, rely=0.04, relheight=0.07)       
    
        self.layerimg = PhotoImage(file="Images/Layer 4.png")
        self.update_button = Button( 
            bg="#c36464",
            image=self.layerimg) 
        self.update_button.place(relx=0.1, rely=0.04)
        
    # Date and Registration
    def __create_date_registration(self):
        self.lreg = Label(
            text="Registration No:", 
            font="Arial 13",
            bg = "#06283D", # framefg = "#06283D" 
            fg="#EDEDED") # framebg
        self.lreg.place(relx=0.04, rely=0.15)
        self.ldate = Label(
            text="Date:", 
            font="Arial 13",
            bg = "#06283D", # framefg = "#06283D" 
            fg="#EDEDED") # framebg
        self.ldate.place(relx=0.4, rely=0.15)
        
        self.Registration = StringVar()
        self.Date = StringVar()
        
        self.reg_entry = Entry(
            textvariable=self.Registration,
            width=14,
            font="Arial 10"
        )
        self.reg_entry.place(relx=0.15, rely=0.15)
        
        # Get todays date
        today = date.today()
        d1 = today.strftime("%d/%m/%Y")
        self.data_entry = Entry(
            textvariable=self.Date,
            width=15,
            font="Arial 10"
        )
        self.data_entry.place(relx=0.44, rely=0.15)
        
        self.Date.set(d1)
        
    
    def __create_student_base(self):
        self.student_data = pathlib.Path("Student_data.xlsx")
        if self.student_data.exists():
            pass
        else: 
            file = openpyxl.Workbook()
            sheet = file.active
            sheet['A1'] = "Registration No."
            sheet['B1'] = "Name"
            sheet['C1'] = "Class"
            sheet['D1'] = "Gender"
            sheet['E1'] = "DOB"
            sheet['F1'] = "Date of Registration"
            sheet['G1'] = "Regilion"
            sheet['H1'] = "SKill"
            sheet['I1'] = "Father Name"
            sheet['J1'] = "Mother Name"
            sheet['K1'] = "Father's Occupation"
            sheet['L1'] = "Mother's Occupation"             

            self.student_data.save('Student_data.xlsx')
        




if __name__ == "__main__":
    MyApp = StudentRegistrationSystem()
    MyApp.mainloop()