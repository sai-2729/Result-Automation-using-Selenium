from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import openpyxl

jntua=webdriver.Chrome(executable_path="../chromedriver/chromedriver.exe")



def user_in(hall_ticket):

    xl=openpyxl.Workbook()

    data=xl.active



    code=int(hall_ticket[7:])

    for i in range(5):
        code+=1
        hall_ticket=hall_ticket[:7]+str(code)
        print(hall_ticket)
        jntua.get('https://jntuaresults.ac.in/view-results-56736237.html')
        jntua.find_element_by_class_name('txt').send_keys(hall_ticket)
        jntua.find_element_by_class_name('ci').click()



        try:
            xl.create_sheet(hall_ticket)
            data.title=hall_ticket
            data=xl[hall_ticket]
            jntua.implicitly_wait(10)
            jntua.find_element_by_xpath('/html/body/div/div[1]/div/div/center/div[1]/table')
        except NoSuchElementException:
            xl.remove_sheet(xl[hall_ticket])
            print('Student Details Not Found')
            continue

        rows=len(jntua.find_elements_by_xpath('//*[@id="rs"]/table/tbody/tr'))
        cols=len(jntua.find_elements_by_xpath('//*[@id="rs"]/table/tbody/tr[1]/th'))


        for r in range(2,rows):
            for c in range(1,cols+1):
                res=jntua.find_element_by_xpath("//*[@id='rs']/table/tbody/tr["+str(r)+"]/td["+str(c)+"]").text
                cell_data=data.cell(r,c)
                cell_data.value=res

        xl.save("new.xlsx")




user_in("179f1a0520")