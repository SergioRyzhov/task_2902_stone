import asyncio
import os
import sys
from contextlib import asynccontextmanager
from datetime import datetime, timedelta
from time import sleep

import dotenv
import pandas as pd
import telebot
from selenium import webdriver
from selenium.common import (TimeoutException,
                             NoSuchElementException,
                             ElementNotInteractableException)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.expected_conditions import presence_of_element_located
from selenium.webdriver.support.wait import WebDriverWait
from typing import List
import logging

logging.basicConfig(level=logging.INFO)

dotenv.load_dotenv()

# load data from sku.xlsx
df = pd.read_excel(os.getenv('FILE_NAME'), header=None)
SKU_LIST = df.iloc[:, 0].tolist()

URL = os.getenv('URL')

# max waiting time for load page
timeout = 10

DATA = []
# time delta days for old feedbacks
TIME_DELTA = 10

API_KEY = os.getenv('TELEGRAM_TOKEN')
CHANNEL_ID = os.getenv('CHANNEL')


def collect_all_feedbacks(driver_var: WebDriver, time: int) -> List[WebElement]:
    """
    Collects all feedbacks from a webpage until a specified time threshold is reached.

    Args:
        driver_var (WebDriver): The Selenium WebDriver instance.
        time (int): The time threshold in days. Feedbacks older than this threshold will be collected.

    Returns:
        List[WebElement]: A list of WebElement objects representing the collected feedback items.
    """
    feedbacks_list = []
    new_feedbacks_list = driver_var.find_elements(By.CLASS_NAME, 'comments__item')

    while new_feedbacks_list:
        feedback_border = None
        for index, feedback in enumerate(new_feedbacks_list):
            feedback_date_str = feedback.find_element(By.CLASS_NAME, 'feedback__date').get_attribute('content')
            feedback_date = datetime.strptime(feedback_date_str, '%Y-%m-%dT%H:%M:%SZ')
            # check old feedbacks and stop if found
            if feedback_date < (datetime.now()-timedelta(days=time)):
                feedback_border = index
                break
        if feedback_border:
            feedbacks_list.extend(new_feedbacks_list[:feedback_border])
            break
        elif feedback_border == 0:
            break
        else:
            feedbacks_list.extend(new_feedbacks_list)

        try:
            new_feedbacks_list[-1].click()
            sleep(1)
        except ElementNotInteractableException:
            feedbacks_list.extend(new_feedbacks_list)
            break
        new_feedbacks_list = driver_var.find_elements(By.CLASS_NAME, 'comments__item')[len(feedbacks_list):]
    logging.info(f'{len(feedbacks_list)} elements found')
    return feedbacks_list


async def handle_page(driver_var: WebDriver, url: str, sku: str) -> None:
    """
    Handles the page with the specified URL and SKU, collecting relevant data.

    Args:
        driver_var (WebDriver): The Selenium WebDriver instance.
        url (str): The URL of the page to handle.
        sku (str): The SKU of the product.

    Returns:
        None
    """
    driver_var.get(url)
    try:
        # on the feedback page
        WebDriverWait(driver_var, timeout).until(presence_of_element_located(
            (By.XPATH, '//*[@id="app"]/div[2]/div/section/div[3]/div/div[1]/div/b')))
        # rating
        rating = driver_var.find_element(By.XPATH,
                                 '//*[@id="app"]/div[2]/div/section/div[3]/div/div[1]/div/b').text
        rating = float(rating)
        if rating < 5:
            # product name
            product_name = driver_var.find_element(By.XPATH,
                                 '/html/body/div[1]/main/div[2]/div/div[2]/div/div[2]/div/div[2]/div[1]/a/b').text
            # messages
            all_feedbacks = collect_all_feedbacks(driver_var, TIME_DELTA)

            for feedback in all_feedbacks:
                feedback_name = None
                feedback_text = None
                feedback_stars = None
                try:
                    feedback_stars = feedback.find_element(By.CLASS_NAME, 'feedback__rating').get_attribute('class')[-1]
                    if feedback_stars == '5':
                        continue
                except NoSuchElementException:
                    pass
                try:
                    feedback_name = feedback.find_element(By.CLASS_NAME, 'feedback__header').text
                except NoSuchElementException:
                    pass
                try:
                    feedback_text = feedback.find_element(By.CLASS_NAME, 'feedback__text').text
                except NoSuchElementException:
                    pass

                feedback_data_item = f"{feedback_name}/{product_name}/{sku}/{feedback_stars}/{feedback_text}/{rating}"
                logging.info(f'Collected: {feedback_data_item}')
                DATA.append(feedback_data_item)

    except TimeoutException:
        pass


@asynccontextmanager
async def async_chrome_driver(options):
    driver = webdriver.Chrome(options=options)
    try:
        yield driver
    finally:
        driver.quit()


async def main():
    """
    Main entry point of the program. Initializes the WebDriver with specified options and handles the pages.

    Returns:
        None
    """
    logging.info(f'Started...')

    # maintain options
    options = Options()
    options.add_argument('--window-size=1524,1580')
    options.add_argument("--headless")
    options.add_argument('--incognito')
    options.add_argument('--disable-infobars')
    options.add_argument('--disable-extensions')
    options.add_argument('--disable-notifications')
    options.add_argument('--disable-default-apps')
    options.add_argument('--disable-bundled-ppapi-flash')
    options.add_argument('--disable-modal-animations')
    options.add_argument('--disable-login-animations')
    options.add_argument('--disable-pull-to-refresh-effect')
    options.add_argument('--blink-settings=imagesEnabled=false')
    options.add_argument('--autoplay-policy=document-user-activation-required')
    options.add_experimental_option(
        "prefs", {"profile.default_content_setting_values.notifications": 1})
    options.add_experimental_option(
        "prefs", {"profile.managed_default_content_settings.images": 2})
    # options.add_experimental_option("excludeSwitches", ["enable-logging"])

    async with async_chrome_driver(options) as driver:
        tasks = []
        for sku in SKU_LIST:
            url = URL + str(sku) + '/feedbacks'
            task = asyncio.create_task(handle_page(driver, url, sku))
            tasks.append(task)

        await asyncio.gather(*tasks)

# run main
asyncio.run(main())


# telegram bot

bot = telebot.TeleBot(API_KEY)


def sent_data():
    """
    Sent data to telegram channel

    :return: None
    """
    for message in DATA:
        try:
            bot.send_message(CHANNEL_ID, message)
        except telebot.apihelper.ApiTelegramException as err:
            if err.error_code == 429:
                retry_after = int(err.error_code)
                print(f"Rate limit exceeded. Retrying after {retry_after} seconds...")
            else:
                print(f"Error: {err}")
    stop_bot()


def stop_bot():
    """
    Stops running bot
    :return:
    """
    bot.stop_polling()

    logging.info(f'Finished')
    sys.exit()


sent_data()
bot.polling()
