# coding=gbk

import logging.config
import logging
import os
import time

import Image
from pytesseract import pytesseract
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
import win32api
import ConfigParser

BASE_DIR = os.path.dirname(__file__)  # 运行目录
logging.config.fileConfig("logging.ini")
logger = logging.getLogger("xzs")
VK_CODE = {
    'backspace': 0x08,
    'tab': 0x09,
    'clear': 0x0C,
    'enter': 0x0D,
    'shift': 0x10,
    'ctrl': 0x11,
    'alt': 0x12,
    'pause': 0x13,
    'caps_lock': 0x14,
    'esc': 0x1B,
    'spacebar': 0x20,
    'page_up': 0x21,
    'page_down': 0x22,
    'end': 0x23,
    'home': 0x24,
    'left_arrow': 0x25,
    'up_arrow': 0x26,
    'right_arrow': 0x27,
    'down_arrow': 0x28,
    'select': 0x29,
    'print': 0x2A,
    'execute': 0x2B,
    'print_screen': 0x2C,
    'ins': 0x2D,
    'del': 0x2E,
    'help': 0x2F,
    '0': 0x30,
    '1': 0x31,
    '2': 0x32,
    '3': 0x33,
    '4': 0x34,
    '5': 0x35,
    '6': 0x36,
    '7': 0x37,
    '8': 0x38,
    '9': 0x39,
    'a': 0x41,
    'b': 0x42,
    'c': 0x43,
    'd': 0x44,
    'e': 0x45,
    'f': 0x46,
    'g': 0x47,
    'h': 0x48,
    'i': 0x49,
    'j': 0x4A,
    'k': 0x4B,
    'l': 0x4C,
    'm': 0x4D,
    'n': 0x4E,
    'o': 0x4F,
    'p': 0x50,
    'q': 0x51,
    'r': 0x52,
    's': 0x53,
    't': 0x54,
    'u': 0x55,
    'v': 0x56,
    'w': 0x57,
    'x': 0x58,
    'y': 0x59,
    'z': 0x5A,
    'numpad_0': 0x60,
    'numpad_1': 0x61,
    'numpad_2': 0x62,
    'numpad_3': 0x63,
    'numpad_4': 0x64,
    'numpad_5': 0x65,
    'numpad_6': 0x66,
    'numpad_7': 0x67,
    'numpad_8': 0x68,
    'numpad_9': 0x69,
    'multiply_key': 0x6A,
    'add_key': 0x6B,
    'separator_key': 0x6C,
    'subtract_key': 0x6D,
    'decimal_key': 0x6E,
    'divide_key': 0x6F,
    'F1': 0x70,
    'F2': 0x71,
    'F3': 0x72,
    'F4': 0x73,
    'F5': 0x74,
    'F6': 0x75,
    'F7': 0x76,
    'F8': 0x77,
    'F9': 0x78,
    'F10': 0x79,
    'F11': 0x7A,
    'F12': 0x7B,
    'F13': 0x7C,
    'F14': 0x7D,
    'F15': 0x7E,
    'F16': 0x7F,
    'F17': 0x80,
    'F18': 0x81,
    'F19': 0x82,
    'F20': 0x83,
    'F21': 0x84,
    'F22': 0x85,
    'F23': 0x86,
    'F24': 0x87,
    'num_lock': 0x90,
    'scroll_lock': 0x91,
    'left_shift': 0xA0,
    'right_shift ': 0xA1,
    'left_control': 0xA2,
    'right_control': 0xA3,
    'left_menu': 0xA4,
    'right_menu': 0xA5,
    'browser_back': 0xA6,
    'browser_forward': 0xA7,
    'browser_refresh': 0xA8,
    'browser_stop': 0xA9,
    'browser_search': 0xAA,
    'browser_favorites': 0xAB,
    'browser_start_and_home': 0xAC,
    'volume_mute': 0xAD,
    'volume_Down': 0xAE,
    'volume_up': 0xAF,
    'next_track': 0xB0,
    'previous_track': 0xB1,
    'stop_media': 0xB2,
    'play/pause_media': 0xB3,
    'start_mail': 0xB4,
    'select_media': 0xB5,
    'start_application_1': 0xB6,
    'start_application_2': 0xB7,
    'attn_key': 0xF6,
    'crsel_key': 0xF7,
    'exsel_key': 0xF8,
    'play_key': 0xFA,
    'zoom_key': 0xFB,
    'clear_key': 0xFE,
    '+': 0xBB,
    ',': 0xBC,
    '-': 0xBD,
    '.': 0xBE,
    '/': 0xBF,
    '`': 0xC0,
    ';': 0xBA,
    '[': 0xDB,
    '\\': 0xDC,
    ']': 0xDD,
    "'": 0xDE,
    '`': 0xC0}

cf = ConfigParser.ConfigParser()
cf.read("main.ini")

# download conf
address = cf.get("download", "addr")
file_img = os.path.join(address, cf.get("download", "file"))

# user conf
username = cf.get("user", "username")
pwd = cf.get("user", "pwd")


def delete_file(filename):
    """delete file"""
    logger.debug(filename)
    if os.path.exists(filename):
        os.remove(filename)


def try_more(web_driver, try_times, f):
    """if failure try more
       f: 高阶函数
    """
    ini_times = try_times
    while try_times > 0:
        logger.debug("try %d", ini_times - try_times + 1)
        result = f(web_driver)
        if result:
            break
        try_times -= 1
    return try_times


def get_driver():
    """set up browser"""
    logger.debug("get_driver begin")
    browser_url = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    chrome_option = webdriver.ChromeOptions()
    chrome_option.binary_location = browser_url
    driver_loc = os.path.join(BASE_DIR, "chromedriver.exe")
    web_driver = webdriver.Chrome(executable_path=driver_loc, chrome_options=chrome_option)
    web_driver.set_page_load_timeout(60)
    web_driver.get("http://www.xbpex.com/")
    web_driver.maximize_window()
    logger.debug("get_driver end")
    return web_driver


def entry_home_try(web_driver):
    """entry home and find login link"""
    try:
        entry_link = web_driver.find_element(By.LINK_TEXT, u"联通入口")
        entry_link.click()

        time.sleep(3)
        all_handlers = web_driver.window_handles
        while len(all_handlers) != 2:
            time.sleep(3)
            all_handlers = web_driver.window_handles

        web_driver.close()
        web_driver.switch_to.window(all_handlers[1])
    except:
        return False

    return True


def entry_home(web_driver):
    """find login link, then close home page"""
    logger.debug("entry_home begin")

    try_times = try_more(web_driver, 3, entry_home_try)
    if try_times == 0:
        raise Exception("find login page more than 3")

    logger.debug("entry_home end")


def entry_login_try(web_driver):
    WebDriverWait(web_driver, 10).until(lambda x: x.find_element_by_id("traderId"))
    element_username = web_driver.find_element_by_id("traderId")
    element_username.clear()
    element_username.send_keys(username)
    element_pwd = web_driver.find_element_by_id("tpassword")
    element_pwd.clear()
    element_pwd.send_keys(pwd)

    delete_file(file_img)

    # 右键菜单
    captcha = web_driver.find_element_by_id("tlaim")
    action = ActionChains(web_driver).move_to_element(captcha)
    action.context_click(captcha)
    action.perform()
    win32api.keybd_event(VK_CODE["down_arrow"], 0, 0, 0)
    win32api.keybd_event(VK_CODE["down_arrow"], 0, 0, 0)
    win32api.keybd_event(VK_CODE["enter"], 0, 0, 0)
    time.sleep(3)
    win32api.keybd_event(VK_CODE["enter"], 0, 0, 0)

    time.sleep(3)

    try:
        image = Image.open(file_img)
        vcode = pytesseract.image_to_string(image)
        element_catpcha = web_driver.find_element_by_id("timgcode")
        element_catpcha.send_keys(vcode)
        logger.debug(vcode)
    except:
        # the character maybe cannot decode, and send to input box
        return False

    entry_link = web_driver.find_element_by_id("logon")
    entry_link.click()

    # if login failure for error ocs resolve
    time.sleep(3)
    try:
        submit = web_driver.find_element_by_id("asyncbox_error_ok")
        if submit.is_displayed():
            submit.click()
            return False
        else:
            # maybe won't be here
            return True
    except:
        # success login into
        return True


def entry_login(web_driver):
    logger.debug("entry_login begin")

    try_times = try_more(web_driver, 9, entry_login_try)
    if try_times == 0:
        raise Exception("try ocs more than 9")

    time.sleep(3)
    all_handlers = web_driver.window_handles
    web_driver.switch_to.window(all_handlers[0])

    logger.debug("entry_login end")


def entry_main(web_driver):
    logger.debug("entry_main begin")

    element_jjyw = web_driver.find_element(By.CSS_SELECTOR, "#modulebtndiv>span:nth-child(3)>div>a")
    element_jjyw.click()

    logger.debug("entry_main end")


def workflow():
    driver = get_driver()
    entry_home(driver)
    entry_login(driver)
    entry_main(driver)


try:
    workflow()
    raw_input()  # wait until enter
except:
    logger.exception('Got exception on main handler')
