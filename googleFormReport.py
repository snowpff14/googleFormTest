import os
import time
from datetime import datetime
from tkinter import Tk, messagebox

import traceback

import numpy as np
import openpyxl as excel
import pandas as pd
import xlrd
import configparser

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import Select, WebDriverWait
from seleniumOperationBase import SeleniumOperationBase
from utils.logger import LoggerObj
from webBase import WebExecuteBase
from openpyxl.styles.borders import Border, Side
import openpyxl as excel
from openpyxl.utils import get_column_letter

# driver=webdriver.Chrome('C:/webdrivers/chromedriver.exe')
binary = FirefoxBinary('C:\\Program Files\\Mozilla Firefox\\firefox.exe')
caps = DesiredCapabilities().FIREFOX
caps["marionette"] = True
driver = webdriver.Firefox(
        capabilities=caps,
        firefox_binary=binary,
        executable_path='C:\webdrivers\geckodriver.exe')

# driver=webdriver.Firefox('C:\webdrivers\geckodriver.exe')
# driver=webdriver.Firefox()

# ログイン情報などを取得
iniFile=configparser.ConfigParser()

PULLDOWN_BUTTON_1='//div[1]/div/div[2]/div[1]/div[2]'
PULLDOWN_TYPE_1='//div[1]/div/div[2]/div[1]/div[1]'
TERM_1='//div[2]/div/div[2]/div/div[1]/div/div[1]/input'
PULLDOWN_BUTTON_2='//div[3]/div/div[2]/div[1]/div[2]'
PULLDOWN_TYPE_2='//div[3]/div/div[2]/div[1]/div[1]'
TERM_2='//div[4]/div/div[2]/div/div[1]/div/div[1]/input'
PULLDOWN_BUTTON_3='//div[5]/div/div[2]/div[1]/div[2]'
PULLDOWN_TYPE_3='//div[5]/div/div[2]/div[1]/div[1]'
TERM_3='//div[6]/div/div[2]/div/div[1]/div/div[1]/input'

SEND_BUTTON='/html/body/div/div[2]/form/div/div[2]/div[3]/div[1]/div/div/content/span'

class GoogleFormReport(SeleniumOperationBase):

    def __init__(self,driver,log,screenShotBaseName='screenShotName'):
        super().__init__(driver,log,screenShotBaseName)
    

    def inputReport(self,reportSheet,targetUrl):
        reportSheetDict=reportSheet.to_dict('index')

        for index,data in reportSheetDict.items():

            number=data['項番']
            if number !=number:
                # 報告日がない状態であれば登録処理を終了
                break
            pulldown1=data['プルダウン1']
            pulldown2=data['プルダウン2']
            pulldown3=data['プルダウン3']
            term1=data['任意項目1']
            term2=data['任意項目2']
            term3=data['任意項目3']


            try:
                self.selectPullDownGoogleForm(PULLDOWN_BUTTON_1,pulldown1,0)
                super().sendTextWaitDisplay(TERM_1,term1)

                super().moveScroll(TERM_1)
                self.selectPullDownGoogleForm(PULLDOWN_BUTTON_2,pulldown2,1)
                super().sendTextWaitDisplay(TERM_2,term2)

                super().moveScroll(TERM_3)
                self.selectPullDownGoogleForm(PULLDOWN_BUTTON_3,pulldown3,2)
                super().sendTextWaitDisplay(TERM_3,term3)
                # super().webElementClick(SEND_BUTTON)
                num=str(index+1)

                super().createOkDialog(num+':人目入力完了','送信完了後にOKを押してください')

                driver.get(targetUrl)

                

            except TimeoutError :
                super().log.error('画面構成が想定外のため失敗:')
                super().log.error(data)
                super().getScreenShot(screenShotName='ErrorInfo',sleepTime=2)
            except  :
                super().log.error(traceback.format_exc())
                super().log.error('登録失敗:')
                super().log.error(data)
                super().getScreenShot(screenShotName='ErrorInfo',sleepTime=2)





class SeleniumTestSite(WebExecuteBase):
    inputFile=''
    def init(self,inputFile,mode=0,filePaths=['resources/appConfig.ini']):
        super().init(iniFile,mode=mode,filePaths=filePaths)
        self.inputFile=inputFile
    
    def mainExecute(self):
        TARGET_URL=iniFile.get('info','url')
        logClass=LoggerObj()
        log=logClass.createLog()
        driver.get(TARGET_URL)
        excelFile=pd.ExcelFile(self.inputFile)
        reportSheet=excelFile.parse(sheet_name='登録内容',dtype='str')
        print(reportSheet.head())


        googleFormReport=GoogleFormReport(driver,log,'test')
        # グーグルフォームの内容を登録
        googleFormReport.inputReport(reportSheet,TARGET_URL)


        time.sleep(2)

        googleFormReport.createOkDialog('処理完了','登録処理完了')

        #driver.close()


        

# メイン処理
if __name__=="__main__":
    seleniumTestSite =SeleniumTestSite()
    seleniumTestSite.init('data/フォーム.xlsx')
    seleniumTestSite.mainExecute()



