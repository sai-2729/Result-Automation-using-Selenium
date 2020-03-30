from selenium import webdriver
import openpyxl



xl=openpyxl.Workbook()
sh=xl.active
sh.title="Sheet"


arr=['OS','CN','OOAD','PPL','BD','OOAD LAB','OS LAB', 'SOCIAL VALUES & ETHICS']
arr.reverse()
print(arr)
for i in range(1,25,3):
   
    c=sh.cell(1,i+2)
    c.value=arr.pop()
    
  

xl.save("xcel.xlsx")
