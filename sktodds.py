# -*- coding: utf-8 -*-

import sys
import os
from os.path import join, dirname
from dotenv import load_dotenv
import time
import gspread
import pyocr.builders
from PIL import Image
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.keys import Keys as keys

if __name__ == "__main__":

    dotenv_path = join(dirname(__file__), '.env')
    load_dotenv(dotenv_path)

    SKT_JSON_KEYFILE = os.environ.get("SKT_JSON_KEYFILE")
    SKT_SPREAD_NAME = os.environ.get("SKT_SPREAD_NAME")
    SKT_WORKSHEET_NAME = os.environ.get("SKT_WORKSHEET_NAME")
    SKT_CHROMEDRIVER_PATH = os.environ.get("SKT_CHROMEDRIVER_PATH")
    SKT_TWITTER_LOGIN = os.environ.get("SKT_TWITTER_LOGIN")
    SKT_TWITTER_PWD = os.environ.get("SKT_TWITTER_PWD")

    SCOPE = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(SKT_JSON_KEYFILE, SCOPE)
    gc = gspread.authorize(credentials)

    wb = gc.open(SKT_SPREAD_NAME)
    wks1 = wb.sheet1
    wks = wb.worksheet(SKT_WORKSHEET_NAME)

    infolist = [
        ["デポ1", "H2", "G2", "dep01"],
        ["デポ2", "H4", "G4", "dep02"],
        ["デポ3", "H6", "G6", "dep03"],
        ["ケン1", "H8", "G8", "ken01"],
        ["ケン2", "H10", "G10", "ken02"],
        ["ケン3", "H12", "G12", "ken03"],
        ["イシ1", "H14", "G14", "isi01"],
        ["イシ2", "H16", "G16", "isi02"],
        ["イシ3", "H18", "G18", "isi03"],
        ["ケレ1", "H20", "G20", "ker01"],
    ]

    driver = webdriver.Chrome(executable_path=SKT_CHROMEDRIVER_PATH)
    driver.set_window_size(1000, 1200)

    ####
    # Twitterログイン start
    ####
    try:
        url = "https://twitter.com/login"
        driver.get(url)
        driver.implicitly_wait(3)
        u_xpath = '//input[@name="session[username_or_email]"]'
        p_xpath = '//input[@name="session[password]"]'
        username = driver.find_element_by_xpath(u_xpath)
        password = driver.find_element_by_xpath(p_xpath)
        username.send_keys(SKT_TWITTER_LOGIN)
        password.send_keys(SKT_TWITTER_PWD)
        password.send_keys(keys.ENTER)
        driver.implicitly_wait(3)
    except:
        print("exception occured")
    ####
    # Twitterログイン end
    ####

    for info in infolist:

        servername = info[0]
        poll_url_cell = info[1]
        odds_cell = info[2]
        servercode = info[3]

        poll_url = wks.acell(poll_url_cell).value
        # print(poll_url)

        ####
        # オッズ取得 start
        ####
        driver.set_window_size(1000, 1200)

        driver.implicitly_wait(2)
        time.sleep(2)

        url = "https://twitter.com/login"
        driver.get(url)
        driver.implicitly_wait(2)
        time.sleep(2)

        driver.get(poll_url)

        driver.implicitly_wait(2)
        time.sleep(2)

        # # get width and height of the page
        # w = driver.execute_script("return document.body.scrollWidth;")
        # h = driver.execute_script("return document.body.scrollHeight;")
        #
        # # set window size
        # driver.set_window_size(w, h)

        driver.set_window_size(1000, 1200)

        SKT_SCREENSHOT_SAVE_FILENAME = join(dirname(__file__), servercode + '_screen.png')

        # Take Screen Shot
        driver.save_screenshot(SKT_SCREENSHOT_SAVE_FILENAME)

    for info in infolist:

        servername = info[0]
        poll_url_cell = info[1]
        odds_cell = info[2]
        servercode = info[3]

        SKT_SCREENSHOT_SAVE_FILENAME = join(dirname(__file__), servercode + '_screen.png')

        tools = pyocr.get_available_tools()
        if len(tools) == 0:
            print("No OCR tool found")
            sys.exit(1)

        tool = tools[0]
        # print("Will use tool '%s'" % (tool.get_name()))

        txt = tool.image_to_string(
            Image.open(SKT_SCREENSHOT_SAVE_FILENAME),
            lang="jpn",
            builder=pyocr.builders.TextBuilder(tesseract_layout=6)
        )
        # print('---------------------------------------------------')
        # print(txt)
        # print('---------------------------------------------------')
        words = txt.split()
        # print(words)

        odds = "-"
        for word in words:
            print(word)
            if "%" in word:
                odds = word
                break
        ####
        # オッズ取得 end
        ####

        driver.implicitly_wait(3)

        if odds == '-':
            print("failed to get odds continuing")
            continue

        odds_tmp = float(odds.strip('%'))
        if odds_tmp > 100:
            odds_num = odds_tmp / 10
        elif odds_tmp == 100:
            odds_num = 99
        elif odds_tmp == 0:
            odds_num = 1
        else:
            odds_num = odds_tmp

        wks.update_acell(odds_cell, str(odds_num) + "%")
        print(servername + ", " + wks.acell(odds_cell).value + ", " + poll_url)

        ####
        # アンケートURL取得 end
        ####
