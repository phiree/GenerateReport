import tkinter as tk
from tkinter import *
from datetime import datetime,date
from tkinter.messagebox import Message
import Generator
class Application(tk.Frame):
    def __init__(self, master=None):
        tk.Frame.__init__(self, master)
        self.pack()
        self.createWidgets()

    def createWidgets(self):
        self.lbl_title=tk.Label(self,text='BillRelation Report:')
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
           Generator.ReportGenerator(date_start,date_end,'Bill_Report').generate_reports()
        except Exception  as ex:
            Message().show(title='Error',message="Maybe can't connection to server, please try again later.")
            raise ex
        print('Finish')
        Message().show(title='Success',message="Reports generated successfully into 'Report' directory.")


root = tk.Tk(screenName='NTS Myanmar')
app = Application(master=root)
app.mainloop()