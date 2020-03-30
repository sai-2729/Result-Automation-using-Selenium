from selenium import webdriver

#jntua=webdriver.Chrome(executable_path="chromedriver/chromedriver.exe")
jntua.get("https://jntuaresults.ac.in/view-results-56736237.html")


jntua.find_element_by_class_name("txt").send_keys("179f1a0501")
jntua.find_element_by_class_name("ci").click()
jntua.implicitly_wait(10)
value=jntua.find_element_by_xpath("/html/body/div/div[1]/div/div/center/div[1]/table/tbody/tr[2]/td[3]").text
print(value)
