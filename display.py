import customtkinter as ck
import tkinter as tk
from tkinter import ttk
from functions import *
import sys


class month_record:
    def __init__(self,root,year,month,user): 
        self.root=root
        self.month=month
        self.user=user
        self.year=year
        self.root.title(f'{user}の{year}年{month}月記録')
        db=Db_setting('ndb.db')
        db.create('time_card')
        self.data=db.get_data(str(year)+str(month),user)
        columns=tuple(self.data[1])

        # create table
        self.table=ttk.Treeview(root,columns=columns,show='headings')
        # add heading
        
        for column in columns:
            self.table.heading(column,text=column)
            self.table.column(column,width=100,anchor='center')
        self.table.pack(fill='both',expand=True)
        # add record
        for record in range(2,len(self.data)):
            self.table.insert(parent='',index='end',value=tuple(self.data[record]))
        #create excel generation button
        self.button=ttk.Button(root,text='イクセルに出力',command=self.gen_excel,width=15)
        self.button.pack(anchor='nw')
        #create pdf generation button
        self.button=ttk.Button(root,text='pdfに出力',command=self.gen_pdf,width=15)
        self.button.pack(anchor='nw')

    def gen_excel(self):
        month=self.year+self.month
        if month!='':
            data=self.data
            excel=Write_excel()
            excel.create_workbook(data,f'{self.user}{str(month)}.xlsx')
            time.sleep(1)
            os.system('start excel.exe ' +f'{self.user}{str(month)}.xlsx')
        else:
            pass
    
    def gen_pdf(self):
        month=self.year+self.month
        data=self.data
        pdf=Write_excel()
        pdf.write_pdf(data,f'{str(month)}.xlsx',f'{self.user}{str(month)}.pdf')
        os.system(f'explorer {self.user}{str(month)}.pdf')

    def show(month,user):
        root=ck.CTk()
        root.geometry('1400x600')
        app=month_record(root,month[0:4],month[4:6],user)
        root.mainloop()

if __name__=='__main__':
    root=ck.CTk()
    root.geometry('1300x600')
    app=month_record(root,sys.argv[1][0:4],sys.argv[1][4:6],sys.argv[2])
    root.mainloop()