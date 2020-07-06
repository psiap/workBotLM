from selenium import webdriver
from selenium.webdriver.support.ui import Select
import time
import os
from threading import Thread

class TreadParserRun(Thread):
    def __init__(self,filter):
        
        Thread.__init__(self)
        self.__filter = filter

    def run(self):
        self.ParsRunWork()

    def ParsRunWork(self):
        listLogin = loginText()
        temp_count = 0
        while(True):
            for elLogin in listLogin:
                driver = webdriver.Chrome()
                parser = ParseWork(driver,"https://linkum.ru",elLogin.split(":")[0],elLogin.split(":")[1],self.__filter,configText()[1])
                parser.starting_a_page()
                temp_count += 1
                print("Циклов - ",temp_count)

class ParseWork(object):

    def __init__(self, driver, link,login,password,temp_option,time_count):
        self.__driver = driver
        self.__link = link
        self.__login = login
        self.__password = password
        self.__temp_option = temp_option
        self.__time_count = float(time_count)

    

    def starting_a_page(self):
        count = False
        self.singUp()
        try:
            self.error_check = self.__driver.find_elements_by_xpath('//*[text()="«Требует доработки»."]')
        except Exception:
            print("Доработок нету.")
        self.startAkCheck()
        self.__driver.get("https://linkum.ru/user/crowd/")
        self.parsOptions()
        while(count != True):
            if not self.error_check:
                try:
                    self.__driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                except Exception:
                    print("Скрола нету.")
                if self.parseWork() == True:
                    print("Bot: refresh")
                elif self.parseWork() == "ERROR":
                    count = True
                else:
                    print("Заданий сейчас нету")
            else:
                count = True
        try:
            self.startAkCheck()
            self.__driver.close()
        except Exception:
            self.errorText("ERROR: self.__driver.close()")
    

    def singUp(self):
        os.system("cls")
        self.loginPrint()
        self.__driver.get(self.__link)
        self.__driver.maximize_window()
        self.timeSleep()
        try:
            self.__driver.find_element_by_class_name("authin").click()
            self.timeSleep()
            self.__driver.find_element_by_id("login").send_keys(self.__login)
            self.__driver.find_element_by_id("pass").send_keys(self.__password)     
            self.__driver.find_element_by_id("loginsubmit").click()
        except Exception:
            self.errorText("ERROR: class_name(authin).click() BLOCK")

    def startAkCheck(self):
        for i in range(0,1):
         try:
            self.__driver.get("https://linkum.ru/user/crowd/links/")
            self.__driver.execute_script("window.scrollTo(document.body.scrollHeight, 0);")
            self.__driver.find_element_by_xpath('//*[text()=". Необходимо выполнить сначала срочные работы."]')
            self.__driver.find_element_by_class_name("crowd_cancel").click()
            self.__driver.switch_to_alert().accept()
            self.timeSleep()
            self.__driver.get("https://linkum.ru/user/crowd/")
            self.timeSleep()
            self.parsOptions()
            self.timeSleep()
            return True
         except Exception:
            print("Bot: kxmmm.....")         

    def checkFast(self):
        try:
            self.__driver.find_element_by_xpath('//*[text()=". У вас в работе уже несколько срочных работ. Выполните сначала их."]')
            self.timeSleep()
            self.__driver.get("https://linkum.ru/user/crowd/links/")
            self.timeSleep()
            self.__driver.execute_script("window.scrollTo(document.body.scrollHeight, 0);")
            self.timeSleep()
            self.__driver.find_element_by_class_name("crowd_cancel").click()
            self.timeSleep()
            self.__driver.switch_to_alert().accept()
            self.timeSleep()
            self.__driver.get("https://linkum.ru/user/crowd/")
            self.timeSleep()
            self.parsOptions()
            self.timeSleep()
            return True
        except Exception:
            print("Bot: kxmmm.....")
        try:
            self.__driver.find_element_by_xpath('//*[text()="ЗАДАНИЕ ДЛЯ ИСПОЛНИТЕЛЯ"]')
            self.startAkCheck()
            self.__driver.get("https://linkum.ru/user/crowd/")
            self.parsOptions()
        except Exception:
            print("Bot: kxmmm.....")
            

    def parsOptions(self):
        os.system("cls")
        self.loginPrint()
        try:
            element = self.__driver.find_element_by_xpath("//select[@name='placement_type']")
            all_options = element.find_elements_by_tag_name("option")
            for option in all_options:
                print("Порядковая цифра: %s) %s" % (option.get_attribute("value"),option.get_attribute("text")))
            
            for option in all_options:
                if option.get_attribute("value") == self.__temp_option:
                   option.click()
                   self.__driver.find_element_by_xpath('//*[text()="Выбрать"]').click()
                   break
        except Exception:
            self.errorText("ERROR: option BLOCK")

       
    
    def alertCheck(self):
         self.timeSleep()
         self.timeSleep()
         try:
             ale = self.__driver.switch_to_alert()
         except Exception:
             self.errorText("ale = self.__driver.switch_to_alert()")
             return True
         if ale.text == "У вас в работе уже больше 50 заказов. Пожалуйста, выполните сначала их.":
             ale.accept()
             return "ERROR"
         elif ale.text == "К сожалению, данный заказ недоступен.":
             ale.accept()
             return "ERROR"
         else:
             ale.accept()
             return True


    def parseWork(self):
        os.system("cls")
        self.loginPrint()
        self.__driver.refresh()
        try:
            elementValue = self.__driver.find_element_by_xpath("//select[@name='pp']")
            all_optionsValue = elementValue.find_elements_by_tag_name("option")
            for option in all_optionsValue:
                if option.get_attribute("value") == "100":
                   option.click()
        except Exception:
            print("Увеличений нету.")
        try:
            elem_work = self.__driver.find_elements_by_xpath('//*[text()="Посмотреть задание"]')
        except Exception:
            print("Данный блок не нашло: elem_work")
        try:
           elem_rab = self.__driver.find_elements_by_xpath('//*[text()="Взять в работу"]')
        except Exception:
            print("Данный блок не нашло: elem_rab")


        if elem_rab:
            for el_work in reversed(elem_rab):
                
                try:
                    el_work.click()
                    self.timeSleep()
                    self.checkFast()
                except Exception:
                    self.errorText("ERROR: el_work.click()")
                    return True
                try:
                    self.__driver.find_element_by_xpath('//*[text()="СРОЧНАЯ"]')
                    return True
                    #zero_check = self.__driver.find_element_by_class_name("error")
                except Exception:
                    print("Zero check.")
                    zero_check = []
                if not zero_check:
                    print("Нашло")
                    try:
                        self.timeSleep()
                        self.timeSleep()
                        btn = self.__driver.find_element_by_xpath("//input[@type='submit']").click()
                        self.timeSleep()
                    except Exception:
                        self.errorText("ERROR: self.__driver.find_element_by_xpath(//*[text()=Например,]).click()")
                        return True
                    

                    return self.alertCheck()
                else:
                    try:
                        self.__driver.find_element_by_class_name("simplemodal-close").click()
                    except Exception:
                        self.errorText("ERROR: simplemodal-close")
                    self.timeSleep()
                    return True
            return False
        else:
            for el_work in reversed(elem_work):
                
                try:
                    el_work.click()
                    self.timeSleep()
                    self.checkFast()
                except Exception:
                    self.errorText("ERROR: el_work()_MOD2")

                try:
                    self.__driver.find_element_by_xpath('//*[text()="СРОЧНАЯ"]')
                    return True
                except Exception:
                    print("Zero check")
                    zero_check = []
                if not zero_check:
                    print("Нашло")
                    try:
                        self.timeSleep()
                        self.timeSleep()
                        self.__driver.find_element_by_xpath('//*[text()="Задание:"]').click()

                    except Exception:
                        print("ERROR: Задание не взялось.")
                        return True
                    
                    return self.alertCheck()
                else:
                    try:
                        self.__driver.find_element_by_class_name("simplemodal-close").click()
                    except Exception:
                        self.errorText("ERROR: simplemodal-close")
                    self.timeSleep()
                    return True

    def timeSleep(self):
        print("Time SLEEP")
        time.sleep(self.__time_count)
       
    def loginPrint(self):
        print("workBotLM  -- Login: %s | Password: %s\n\n" % (self.__login,self.__password))

    def errorText(self,error):
        f2 = open('errorLog.txt', 'a')
        f2.write(error + " " + time.ctime() + "\n")
        f2.close()
            
            

        
        
def configText():
    f1 = open('config.txt', 'r')
    s = f1.readline()
    f1.close()
    return s.split(" ")

def loginText():
    f2 = open('login.txt', 'r')
    l = list([line.rstrip() for line in f2.readlines()])
    f2.close()
    return l



def main():
    cT = configText()[3::]
    for item, numcT in enumerate(cT):
        thread = TreadParserRun(numcT)
        thread.start()


if __name__ == "__main__":
    main()