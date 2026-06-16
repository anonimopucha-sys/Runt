import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver import Firefox
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
from selenium.webdriver.common.keys import Keys
import os
import pyperclip
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
import logging
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options

