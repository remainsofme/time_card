import tkinter as tk
import customtkinter as ck
from tkinter import messagebox
from functions import *


#cal salary to correct! 2024/4/16
class TimeCardApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Time Card")
        
        # Create labels
        self.label = ck.CTkLabel(root, text="Time Card", font=("Helvetica", 20))
        self.label.pack(pady=10)
        self.label = ck.CTkLabel(root, text="Name", font=("Helvetica", 10))
        self.label.pack(pady=10)

        self.user_combobox=ck.CTkComboBox(root,state='readonly',values=['呉','李','周','王'])
        self.user_combobox.pack(pady=5)


        self.start_button =ck.CTkButton(root, text="Start Time Check",command=self.start_time_check)
        self.start_button.pack(pady=5)

        self.end_button = ck.CTkButton(root, text="End Time Check",command=self.end_time_check)
        self.end_button.pack(pady=5)

        self.label = ck.CTkLabel(root, text="year and month", font=("Helvetica", 10))
        self.label.pack(pady=10)

        self.month_entry=ck.CTkEntry(root)
        self.month_entry.pack(pady=5)
        
        self.payslip_button = ck.CTkButton(root, text="Create Payslip", command=self.cal_salary)
        self.payslip_button.pack(pady=5)

    def choose_user(self):
        self.user=self.user_combobox.get()
      
    def start_time_check(self):
        self.choose_user()
        print(self.user+'hello')
        start_time=Salary_cal.get_time()
        db=Db_setting('ndb.db')
        db.create('time_card')
        if self.user=='':
            messagebox.showinfo('info','名前を入れてください')
        else:
            if  db.get_start_info(self.user)==False:
                db.add_record('start_time',start_time[0],self.user)
                print(f'start time {start_time[1]}')
                messagebox.showinfo('info',f'{self.user}は{start_time[1]}に出勤しました')
            else:
                messagebox.showinfo('info',f'{self.user}の今日の出勤記録はすでにされています' )
    
    def end_time_check(self):
        self.choose_user()
        end_time=Salary_cal.get_time()
        db=Db_setting('ndb.db')
        db.create('time_card')
        if self.user=='':
            messagebox.showinfo('info','名前を入れてください')
        else:
            if db.get_start_info(self.user)==False:
                messagebox.showinfo('info',f'{self.user}の出勤記録がありません。先に出勤記録してください' )
            else:
                db.add_record('end_time',end_time[0],self.user)
                print(f'end time {end_time[1]}')
                messagebox.showinfo('info',f'{self.user}は{end_time[1]}に退勤しました')

    def cal_salary(self):
        #get id from the date entry, get name from name entry, check read record, fill in records user=self.user id=self.month_entry.get()
        #consider break xls process and refill process
        #refill process, get all record, use for to loop through them, do the refill
        #xls process, read date and name from entry, read record, make list, output
        rest_time=3600
        self.choose_user()
        db=Db_setting('ndb.db')
        db.create('time_card')
        records=db.read_all_records(self.user)
        if records!=None:
            for record in records:
                if record!=None:
                    if record[2] and record[3] and record[11]!=None:
                        db.fill_record(record)
                    else:
                        messagebox.showinfo('info','出退勤情報がありませんので、作成できません。')
                        
                else:
                    messagebox.showinfo('info','記録がありませんので、作成できません')
        else:
            messagebox.showinfo('info','記録がありませんので、作成できません')


        month=self.month_entry.get()
        if month!='':
            data=db.get_data(month,self.user)
            excel=Write_excel()
            excel.create_workbook(data,str(month)+'.xlsx')
            time.sleep(1)
            os.system('start excel.exe ' +str(month)+'.xlsx')
        else:
            messagebox.showinfo('info','日付を入力してください。例202404')



if __name__=='__main__':
    root=ck.CTk()
    root.attributes('-topmost',True)
    app=TimeCardApp(root)
    root.mainloop()
        