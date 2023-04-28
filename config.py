import xlwings as xw
import os
import time
import datetime
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from dotenv import load_dotenv
load_dotenv()

today = datetime.date.today().strftime("%Y/%m/%d")

spfa_id = os.environ.get("USER_ID")
spfa_pass = os.environ.get("PASS_WORD")
repair_excel = os.environ.get("REPAIR_EXCEL")
spfa_top = os.environ.get("SPFA_TOP")
res_calc = os.environ.get("RES_CALC")
