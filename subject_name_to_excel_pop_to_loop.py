from selenium import webdriver
import openpyxl



xl=openpyxl.Workbook()
sh=xl.active
sh.title="Sheet"


arr=['OS','CN','OOAD','PPL','ST','BD','OOAD LAB','OS LAB', 'SOCIAL VALUES & ETHICS']
count=0
for i in range(1,27,3):
    c=sh.cell(1,i+2)
    count+=1
    c.value=arr[count-1]
    
  

xl.save("xcel.xlsx")
