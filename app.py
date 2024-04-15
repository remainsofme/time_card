import tkinter as tk
from tkinter import messagebox
from functions import *
import math

class TimeCardApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Time Card")
        
        # Create labels
        self.label = tk.Label(root, text="Time Card", font=("Helvetica", 20))
        self.label.pack(pady=10)
        
        # Create buttons
        self.start_button = tk.Button(root, text="Start Time Check",command=self.start_time_check)
        self.start_button.pack(pady=5)

        self.end_button = tk.Button(root, text="End Time Check",command=self.end_time_check)
        self.end_button.pack(pady=5)

        self.month_entry=tk.Entry(root)
        self.month_entry.pack(pady=5)
        
        self.payslip_button = tk.Button(root, text="Create Payslip", command=self.cal_salary)
        self.payslip_button.pack(pady=5)

    def start_time_check(self):
        start_time=Salary_cal.get_time()
        db=Db_setting('ndb.db')
        db.create('time_card')
        if  db.get_start_info()==False:
            db.add_record('start_time',start_time[0])
            print(f'start time {start_time[1]}')
            messagebox.showinfo('info',f'{start_time[1]}に出勤しました')
        else:
            messagebox.showinfo('info','今日の出勤記録はすでにされています' )
    
    def end_time_check(self):
        end_time=Salary_cal.get_time()
        db=Db_setting('ndb.db')
        db.create('time_card')
        if db.get_start_info()==False:
            messagebox.showinfo('info','出勤記録がありません。先に出勤記録してください' )
        else:
            db.add_record('end_time',end_time[0])
            print(f'end time {end_time[1]}')
            messagebox.showinfo('info',f'{end_time[1]}に退勤しました')

    def cal_salary(self):
        db=Db_setting('ndb.db')
        db.create('time_card')
        ids=db.get_ids()
        if ids==[]:
            messagebox.showinfo('info','出退勤情報がありませんので、作成できません。')
        else:
            for id in ids:
                record=db.read_record(id)
                print('!!!!')
                rest_time=1*3600
                if record[2] and record[3] and record[11]!=None: 
                    print('!!')
                    salary_call=Salary_cal(record[2],record[3],rest_time,1188,190000,1.25,1.35,record[11])
                    work_time=salary_call.get_working_time()[0]
                    salary=salary_call.calculate_salary()
                    db.add_record('work_time',work_time)
                    db.add_record('salary',salary)
                    db.add_record('rest_time',rest_time)
                    db.add_record('start_time_formated',salary_call.format_time(record[2]))
                    db.add_record('end_time_formated',salary_call.format_time(record[3]))
                    db.add_record('work_time_formated',f'{math.floor(work_time/3600)}時間{round(((work_time % 3600)/60),1)}分')
                    db.add_record('rest_time_formated',f'{math.floor(rest_time/3600)}時間{round(((rest_time % 3600)/60),1)}分')
                    if not self.month_entry.get():
                        messagebox.showerror('error','出力したい勤務表の年月を入力してください。例：202404')
                    else:
                        print('!!!!!!!!!!!!!!!!!!!')
                        a=int(self.month_entry.get())
                        data=db.get_data(a)
                        excel=Write_excel()
                        excel.create_workbook(data,str(a)+'.xlsx')
                        time.sleep(1)
                        os.system('start excel.exe ' +str(a)+'.xlsx')
                elif record[2]==None:
                    print(2)
                    messagebox.showinfo('info','出勤記録がありません')
                elif record[3]==None:
                    print(3)
                    messagebox.showinfo('info','退勤記録がありません')
                elif record[11]==None:
                    print(11)
                    messagebox.showinfo('info','深刻なエラーが発生しました。管理人にご連絡ください。')
                    print(record[11])
                else:
                    print(500)
                    messagebox.showinfo('info','出退勤記録がありません')

if __name__=='__main__':
    root=tk.Tk()
    root.attributes('-topmost',True)
    app=TimeCardApp(root)
    root.mainloop()
        