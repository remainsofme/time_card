import time
import datetime
import sqlite3
import os
from openpyxl import Workbook

class Db_setting:
    def __init__(self,database_name):
        self.database_name=database_name
        pass
    
    def create(self,table_name):
        if self.database_name not in os.listdir():
            with open(self.database_name,'w') as f:
                pass
        else:
            pass
        conn=sqlite3.connect(self.database_name)
        cursor=conn.cursor()
        cursor.execute(f'''
            CREATE TABLE IF NOT EXISTS {table_name}(
                                id TEXT PRIME KEY,
                                name TEXT,
                                start_time REAL,
                                end_time REAL,
                                rest_time REAL,
                                work_time REAL,
                                salary REAL,
                                start_time_formated TEXT,
                                end_time_formated TEXT,
                                work_time_formated REAL,
                                rest_time_formated REAL,
                                week_day INTEGER
            );
            ''' )
        cursor.close()
        conn.commit()
        conn.close()
        self.table_name=table_name

    def add_record(self,attribute,input):
        conn=sqlite3.connect(self.database_name)
        cursor=conn.cursor()
        id_format='%Y%m%d'
        date=datetime.datetime.now()
        id=date.strftime(id_format)
        cursor.execute(f'SELECT id FROM {self.table_name};')
        rows=cursor.fetchall()
        ids=[row[0] for row in rows]
        print(ids)
        if id in ids:
            cursor.execute(f'''
                UPDATE {self.table_name} SET {attribute}='{input}' WHERE id={id};
                ''')
        elif id not in ids:
            cursor.execute(f'''
                INSERT INTO {self.table_name} (id) VALUES ({id});
                ''')
            cursor.execute(f'''
                UPDATE {self.table_name} SET {attribute}='{input}' WHERE id={id};
                ''')
            cursor.execute(f'''
                UPDATE {self.table_name} SET week_day='{date.weekday()}' WHERE id={id};
                ''')
        conn.commit()
        cursor.close()
        conn.close()
    
    def get_ids(self):
        conn=sqlite3.connect(self.database_name)
        cursor=conn.cursor()
        id_format='%Y%m%d'
        date=datetime.datetime.now()
        id=date.strftime(id_format)
        cursor.execute(f'SELECT id FROM {self.table_name};')
        rows=cursor.fetchall()
        ids=[row[0] for row in rows]
        return ids
    
    def read_record(self,id):
        conn=sqlite3.connect(self.database_name)
        cursor=conn.cursor()
        cursor.execute(f'''
            SELECT * FROM {self.table_name} WHERE id={id}
            ''')
        conn.commit()
        record=cursor.fetchone()
        cursor.close()
        conn.close()
        return list(record) if record else None
    
    def get_data(self,month):
        ids=self.get_ids()
        data=[]
        data.append([f'{month}月'])
        data.append([
            '日付',
            '曜日',
            '出勤時間',
            '退勤時間',
            '休憩時間',
            '労働時間',
            '給料'
        ])

        for id in ids:
            if int(str(id)[0:6])==month:
                record=self.read_record(id)
                row=[
                    str(record[0])[0:4]+'年'+str(record[0])[4:6]+'月'+str(record[0])[6:8]+'日',
                    record[11]+1,
                    record[7],
                    record[8],
                    record[10],
                    record[9],
                    round(record[6],1)
                ]
                data.append(row)
            else:
                print('not found')
        print(data)
        return data
    
    def get_start_info(self):
        date=datetime.datetime.now()
        id=date.strftime('%Y%m%d')
        print(id)
        if self.read_record(id)==None:
            self.have_start_time=False
            print('no no start time')
        else:
            start_time=self.read_record(id)[2]
            if start_time!=None:
                self.have_start_time=True
                print('already checked in')
            else:
                print('no start_time')
                self.have_start_time=False
        return self.have_start_time
    
            

class Salary_cal:
    def __init__(self,start,end,rest,hour_salary,month_salary,overtime_rate,holiday_rate,weekday):
        self.start=start
        self.end=end
        self.rest=rest
        self.hour_salary=hour_salary
        self.month_salary=month_salary
        self.legal_work_time=3600*8
        self.overtime_rate=overtime_rate
        self.holiday_rate=holiday_rate
        self.weekday=weekday

    def get_time():
        time_format='%Y年%m月%d日%H時%M分%S秒'
        timestamp=time.time()
        datetime_object=datetime.datetime.fromtimestamp(timestamp)
        formated_time=datetime_object.strftime(time_format)
        return [timestamp,formated_time]
    
    def format_time(self,time_stamp):
        time_format='%Y年%m月%d日%H時%M分%S秒'
        datetime_object=datetime.datetime.fromtimestamp(time_stamp)
        formated_time=datetime_object.strftime(time_format)
        return formated_time

    
    def get_working_time(self):
        if self.end-self.start<6*3600:
            working_time=self.end-self.start
        else:
            working_time=self.end-self.start-self.rest
        formated_time=working_time/3600
        return [working_time,formated_time]
    
    def calculate_salary(self):
        if self.end-self.start<6*3600:
            working_time=self.end-self.start
        else:
            working_time=self.end-self.start-self.rest
        # 平日の割増賃金含む一日の賃金
        if working_time<=0:
            salary=0
        else:
            if self.weekday<5:
                if working_time>self.legal_work_time:
                    salary=self.hour_salary*self.legal_work_time/3600+(working_time/3600-self.legal_work_time/3600)*self.hour_salary*self.overtime_rate
                #　割増賃金ないの場合の一日の賃金
                elif working_time<=self.legal_work_time:
                    salary=self.hour_salary*working_time/3600
            else:
                salary=self.hour_salary*working_time*self.holiday_rate/3600
        return salary
        

class Write_excel:
    def __init__(self):
        pass
    def create_workbook(self,data,file_name):
        wb=Workbook()
        ws=wb.active
        for row in data:
            ws.append(row)
        wb.save(file_name)






if __name__=="__main__":
    db=Db_setting('ndb.db')
    db.create('time_card')
    db.add_record('start_time',1)
    db.add_record('id', 33)
    print(db.read_record(1))
 
        
        