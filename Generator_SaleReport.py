import tkinter as tk
from tkinter import *
from tkinter.messagebox import Message
from datetime import datetime,date
import Generator
class Application(tk.Frame):
    def __init__(self, master=None):
        tk.Frame.__init__(self, master)
        self.pack()
        self.createWidgets()

    def createWidgets(self):
        self.lbl_title=tk.Label(self,text='Customer Report:')
        self.lbl_title.pack(side='top')

        self.lbl_start=tk.Label(self,text='Start Date:')
        self.lbl_start.pack(side='top')

        self.date_start=tk.Text(self,height=1)
        self.date_start.insert(INSERT,date(date.today().year,date.today().month,1))
        self.date_start.pack(side='top')

        self.lbl_end=tk.Label(self,text='End Date:')
        self.lbl_end.pack(side='top')

        self.date_end=tk.Text(self,height=1)
        self.date_end.insert(INSERT,date.today())
        self.date_end.pack(side='top')

        self.generate=tk.Button(self,text='Generate Report')
        self.generate.pack(side='top')
        self.generate['command']=self.generate_report

    def generate_report(self):
        date_start=self.date_start.get('0.0',END).strip()
        date_end=self.date_end.get('0.0',END).strip()
        try:
            Generator.ReportGenerator(date_start,date_end,'SaleReport').generate_reports()
            Generator.ReportGenerator(date_start,date_end,'CustomerReport').generate_reports()
        except:
            Message().show(title='Error',message="Maybe can't connection to server, please try again later.")
        print('Finish')


root = tk.Tk(screenName='NTS Myanmar')
app = Application(master=root)
app.mainloop()