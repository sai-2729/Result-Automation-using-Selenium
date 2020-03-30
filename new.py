import openpyxl

xl=openpyxl.Workbook()

data=xl.active

data.title="result"

data["A4"].value="12"


xl.save("new.xlsx")