# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import pyautogui
import calendar
import autoit
import unittest, time, re, os, webbrowser, shutil
import win32api, win32con
import io
from PIL import Image
from python_anticaptcha import AnticaptchaClient, ImageToTextTask
from selenium.common.exceptions import NoSuchElementException


def click(x,y):
        win32api.SetCursorPos((x,y))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)

class UntitledTestCase(unittest.TestCase):
    def setUp(self):

        razao_social = 'SER'

        today = datetime.today()
        Meses=('Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro')
        mes = (today.month-1)
        mes_nome = Meses[mes-1]+ ' - '


        string_1 = "01-NF-e emitidas"
        dir = 'C:\\' + razao_social + '\\' + mes_nome + str(mes) + '\\' + string_1

        if os.path.exists(dir):
            print("O diretório z existe.")
        else:
            os.makedirs(dir)
      
        #options = Options()
        #options.headless = True
        profile = webdriver.FirefoxProfile()
        profile.set_preference("browser.download.folderList", 2)
        profile.set_preference("browser.download.manager.showWhenStarting", False)
        profile.set_preference("browser.download.dir", dir)
        profile.set_preference("browser.download.panel.shown", False)
        profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf, text/plain, text/csv, text/xml")
        profile.set_preference("plugin.disable_full_page_plugin_for_types", "application/pdf, text/plain, text/csv, text/xml")
        profile.set_preference("pdfjs.disabled", True)

        self.driver = webdriver.Firefox(firefox_profile=profile, executable_path=r'C:\\geckodriver.exe')
    
        self.driver.implicitly_wait(10)

        self.base_url = "https://www.katalon.com/"
        self.verificationErrors = []
        self.accept_next_alert = True

    def test_untitled_test_case(self):
        driver = self.driver
        driver.get("https://nfse.salvador.ba.gov.br/")
        
        login = driver.find_element_by_id("txtLogin")
        senha = driver.find_element_by_id("txtSenha")
        login.send_keys("18906494000133")
        senha.send_keys("ehst2013")
        
        #teste captcha começa aqui
        found = False
        while not found:
            try:
                element = driver.find_element_by_id("cvCodigo")
                found = True
                
                image = driver.find_element_by_xpath(u'//*[@id="img1"]') .screenshot_as_png
                imageStream = io.BytesIO(image)
                im = Image.open(imageStream)
                im.save('screen.png')

                api_key = 'b9195adfcf2ebee8a84904766fcba4dc'
                captcha_fp = open('screen.png', 'rb')
                client = AnticaptchaClient(api_key)
                task = ImageToTextTask(captcha_fp)
                job = client.createTask(task)
                job.join()
                print(job.get_captcha_text())
                captcha_text = job.get_captcha_text()
                captcha = driver.find_element_by_id("tbCaptcha")
                captcha.send_keys(captcha_text)
                driver.find_element_by_id("cmdLogin").click()
            except NoSuchElementException:
                break
        
        #teste captcha termina aqui
        
        #print(pyautogui.position())
        pyautogui.moveTo(1054,328)
        time.sleep(2)
        pyautogui.click()
        time.sleep(5)
        driver.find_element_by_xpath(u"(.//*[normalize-space(text()) and normalize-space(.)='Emissão Guias'])[1]/following::span[1]").click()
        driver.find_element_by_id("ddlMes").click()
        Select(driver.find_element_by_id("ddlMes")).select_by_visible_text("5")
        driver.find_element_by_xpath(u"(.//*[normalize-space(text()) and normalize-space(.)='Mês'])[1]/following::option[10]").click()
        driver.find_element_by_id("btEmitidas").click()
        time.sleep(4)
        
        driver.switch_to.window(driver.window_handles[1])
        driver.find_element_by_id("ddlTipoArquivo").click()
        Select(driver.find_element_by_id("ddlTipoArquivo")).select_by_visible_text("PDF")
        driver.find_element_by_xpath("(.//*[normalize-space(text()) and normalize-space(.)='Para exportar as Notas, selecione o formato do arquivo:'])[1]/following::option[4]").click()
        driver.find_element_by_id("btGerar").click()
        time.sleep(5)
        autoit.send("^p")
        time.sleep(5)
        autoit.send("{ENTER}")
        time.sleep(10)

        autoit.send("{N}")
        time.sleep(5)
        autoit.send("{F}")
        time.sleep(5)
        autoit.send("{S}")
        time.sleep(5)
        autoit.send("{e}")
        time.sleep(5)
        autoit.send("{ENTER}")
        autoit.send("{ENTER}")
        time.sleep(10) # Pause to allow you to inspect the browser.
        # Move a file from the directory d1 to d2
        shutil.move('C:\\Users\\leka\\Documents\\NFSe.pdf', 'C:\\SER\\Outubro - 10\\01-NF-e emitidas\\NFSe.pdf')
        time.sleep(5)
        driver.refresh()

        driver.find_element_by_id("ddlTipoArquivo").click()
        Select(driver.find_element_by_id("ddlTipoArquivo")).select_by_visible_text("Planilha (CSV)")
        driver.find_element_by_id("ddlTipoArquivo").click()
        driver.find_element_by_id("btGerar").click()
        time.sleep(1) # Pause to allow you to inspect the browser.
        driver.refresh()
        driver.find_element_by_id("ddlTipoArquivo").click()
        Select(driver.find_element_by_id("ddlTipoArquivo")).select_by_visible_text("XML")
        driver.find_element_by_id("ddlTipoArquivo").click()
        driver.find_element_by_id("btGerar").click()
        time.sleep(5)
        pyautogui.moveTo(499,463)
        pyautogui.click()
        time.sleep(1)
        pyautogui.moveTo(498,488)
        pyautogui.click()
        time.sleep(1)
        pyautogui.moveTo(770,568)
        pyautogui.click()
        time.sleep(5)
        driver.refresh()
        time.sleep(3)
        driver.find_element_by_id("ddlTipoArquivo").click()
        Select(driver.find_element_by_id("ddlTipoArquivo")).select_by_visible_text("TXT (Tabulado)")
        driver.find_element_by_id("ddlTipoArquivo").click()
        driver.find_element_by_id("btGerar").click()



        driver.switch_to.window(driver.window_handles[0])
        driver.find_element_by_id("hypSair").click()

        time.sleep(10)

        driver.quit() 
        
    def is_element_present(self, how, what):
        try: self.driver.find_element(by=how, value=what)
        except NoSuchElementException as e: return False
        return True
    
    def is_alert_present(self):
        try: self.driver.switch_to_alert()
        except NoAlertPresentException as e: return False
        return True
    
    def close_alert_and_get_its_text(self):
        try:
            alert = self.driver.switch_to_alert()
            alert_text = alert.text
            if self.accept_next_alert:
                alert.accept()
            else:
                alert.dismiss()
            return alert_text
        finally: self.accept_next_alert = True
    
    def tearDown(self):
        self.driver.quit()
        self.assertEqual([], self.verificationErrors)

#if __name__ == "__main__":
    #unittest.main()