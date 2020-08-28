import openpyxl
import pandas as pd
import random 
from  openpyxl.styles import Border,Side
def date():
    dd = pd.date_range('2020-4-07','2020-8-28')
    date_list = [pd.Timestamp(x).strftime("%m-%d") for x in dd.values]
    return date_list
    
def temperature(day):
    i=0
    temperature_list=[]
    while(i<=day):
        temperature_list.append("36.{}".format(random.randint(0,9)))
        i=i+1
    return temperature_list
    
def Track(day):
    i=0
    track_list=[]
    track_pool=["无","下楼买菜，步行","下楼倒垃圾，步行","采购日用品，步行"]
    while(i<=day):
        track_list.append("{}".format(track_pool[random.randint(0,3)]))
        i=i+1
    return track_list
    
date_list = date()
day = len(date_list)
print(day)
temperature_list1=temperature(day)
temperature_list2=temperature(day)
track_list = Track(day)
location="" #居住地地址  

workbook = openpyxl.load_workbook(".\demo.xlsx") #模板表格
worksheet = workbook.worksheets[0]

i = 0
while(i<day):
    text=[]
    text.append(date_list[i])
    text.append(location)
    text.append(temperature_list1[i])
    text.append(temperature_list2[i])
    text.append("无")
    text.append(track_list[i])
    text.append("无")
    print(text)
    worksheet.append(text)
    i=i+1


filename='.xlsx' #生成的电子表格名称  
workbook.save(filename)