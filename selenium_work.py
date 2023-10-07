import time
import calendar
import logging

from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from datetime import datetime
from datetime import date
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from dateutil.relativedelta import *
from selenium.webdriver import ActionChains

import pandas as pd
from tkinter import messagebox as mbox

module_logger = logging.getLogger("main.selenium")


def format_date_with_dot(_date):
    _date = _date[:2] + '.' + _date[2:4] + '.' + _date[4:]
    return _date


def get_float(text):
    return float(text.replace(',', '.').replace('\xa0', ''))


def open_browser(dict_with_data):
    logger = logging.getLogger("main.selenium.open_browser")

    try:
        debt_sum = float(dict_with_data['debt_sum'])
        first_day = int(dict_with_data['first_day'])
        first_month = int(dict_with_data['first_month'])
        first_year = int(dict_with_data['first_year'])
        directory_for_result = dict_with_data['directory']
        var_visual = dict_with_data['var_visualization']
        var_delay = dict_with_data['var_delay']
        bet_option_var = dict_with_data['bet_option_var']
        last_day = int(dict_with_data['last_day'])
        last_month = int(dict_with_data['last_month'])
        last_year = int(dict_with_data['last_year'])
        choose_date = dict_with_data['choose_date']

        try:
            options = Options()
            if var_visual:
                options.add_argument("--start-maximized")
            else:
                options.add_argument('--headless')
                options.add_argument('--disable-gpu')
            driver = webdriver.Chrome(options=options, service=ChromeService(ChromeDriverManager().install()))
            driver.get("https://calc.consultant.ru/npd")
        except Exception as ex:
            text = f"Текст ошибки:\n\n{ex.msg}\n\nПроверьте версии Google Chrome и ChromeDriver."
            mbox.showerror("Ошибка запуска браузера", text)
            return

        driver.implicitly_wait(5)

        banner = driver.find_element(By.ID, 'bannerPopup')

        driver.execute_script("arguments[0].style.display = 'none';", banner)

        time.sleep(3)

        total_sum = debt_sum

        # Массивы данных для будущей таблицы
        dates = []
        date_period = []
        day = []
        bet_cb = []
        penalty = []
        credit_and_penalty = []

        date_from = driver.find_element(By.ID, 'dateFrom')
        shadow_root_date_from = driver.execute_script('return arguments[0].shadowRoot', date_from)
        date_from_input = shadow_root_date_from.find_element(By.CLASS_NAME, 'dateinput')

        date_to = driver.find_element(By.ID, 'dateTo')
        shadow_root_date_to = driver.execute_script('return arguments[0].shadowRoot', date_to)
        date_to_input = shadow_root_date_to.find_element(By.CLASS_NAME, 'dateinput')

        use_date = date(first_year, first_month, first_day)

        start_date = date(first_year, first_month, first_day)

        end_date = date(first_year, first_month, first_day)

        ruble_input_field = driver.find_element(By.CLASS_NAME, 'main-sum-field-input')
        ruble_input_field.send_keys(debt_sum, Keys.ENTER)

        rate_type_input = driver.find_element(By.ID, 'rateTypeInput')
        field_div = rate_type_input.find_element(By.CSS_SELECTOR, '.field-div dropdown-select')
        shadow_root_rate_type_input = driver.execute_script('return arguments[0].shadowRoot', field_div)
        ddc = shadow_root_rate_type_input.find_element(By.CLASS_NAME, 'ddc')
        ActionChains(driver).click(ddc).perform()

        dropelem = ddc.find_element(By.CSS_SELECTOR, '.dropelems div:last-child')
        ActionChains(driver).click(dropelem).perform()

        bank_rate_type_input = driver.find_element(By.ID, 'bankRateTypeInput')
        share = bank_rate_type_input.find_element(By.ID, 'share')
        ActionChains(driver).click(share).perform()

        share_input = driver.find_element(By.ID, 'shareInput')
        base_input = share_input.find_element(By.CLASS_NAME, 'base-input')
        base_input.send_keys("1300", Keys.ENTER)

        rate_period_input = driver.find_element(By.ID, 'ratePeriodInput')
        field_div = rate_period_input.find_element(By.CSS_SELECTOR, '.field-div dropdown-select')
        shadow_root_rate_type_input = driver.execute_script('return arguments[0].shadowRoot', field_div)
        ddc = shadow_root_rate_type_input.find_element(By.CLASS_NAME, 'ddc')
        ActionChains(driver).click(ddc).perform()

        dropelems = ddc.find_element(By.CSS_SELECTOR, '.dropelems')

        if bet_option_var == 0:
            dropelem = dropelems.find_element(By.CSS_SELECTOR, '.dropelems div:first-child')
        elif bet_option_var == 1:
            dropelem = dropelems.find_element(By.CSS_SELECTOR, '.dropelems div:nth-child(2)')
        elif bet_option_var == 2:
            dropelem = dropelems.find_element(By.CSS_SELECTOR, '.dropelems div:nth-child(3)')
        elif bet_option_var == 3:
            dropelem = dropelems.find_element(By.CSS_SELECTOR, '.dropelems div:last-child')

        ActionChains(driver).click(dropelem).perform()

        if bet_option_var == 3:
            rate_date = driver.find_element(By.ID, 'rateDate')
            shadow_root_rate_date = driver.execute_script('return arguments[0].shadowRoot', rate_date)
            dateinput = shadow_root_rate_date.find_element(By.CLASS_NAME, 'dateinput')
            dateinput.send_keys(choose_date, Keys.ENTER)

        time.sleep(3)

        start_period = list()
        end_period = list()

        for year in range(first_year, last_year + 1, 1):
            for month in range(1, 13, 1):
                if year == first_year and month < first_month:
                    continue

                # блок счета с начала месяца

                if year == last_year and month > last_month:
                    break

                last_day_of_month = calendar.monthrange(use_date.year, use_date.month)[1]

                if year == first_year and month == first_month:
                    if last_day_of_month >= first_day:
                        end_date = end_date.replace(day=last_day_of_month)

                else:
                    use_date = use_date + relativedelta(months=+1)
                    use_date = use_date.replace(day=1)
                    start_date = start_date.replace(year=use_date.year, month=use_date.month, day=use_date.day)
                    last_day_of_month = calendar.monthrange(use_date.year, use_date.month)[1]
                    end_date = end_date.replace(year=use_date.year, month=use_date.month, day=last_day_of_month)

                if year == last_year and use_date.month == last_month:
                    end_date = end_date.replace(day=last_day)

                start_period = list(reversed(start_date.isoformat().split('-')))
                end_period = list(reversed(end_date.isoformat().split('-')))

                current_date_from = f"{start_period[0]}{start_period[1]}{start_period[2]}"

                date_from_input.clear()
                date_from_input.send_keys(current_date_from, Keys.ENTER)

                time.sleep(var_delay)

                current_date_to = f"{end_period[0]}{end_period[1]}{end_period[2]}"

                date_to_input.clear()
                date_to_input.send_keys(current_date_to, Keys.ENTER)

                time.sleep(var_delay)

                ruble_input_field.clear()
                ruble_input_field.send_keys(debt_sum, Keys.ARROW_RIGHT)
                time.sleep(var_delay)
                driver.find_element(By.ID, 'calcSubmit').click()
                time.sleep(var_delay)

                rate_info_render = driver.find_element(By.CLASS_NAME, 'rate-info-render')
                rate = rate_info_render.text

                table = driver.find_element(By.CLASS_NAME, 'detalization-div')
                inner_html_table = table.get_attribute('innerHTML')
                pandas_table = pd.read_html(inner_html_table, thousands=None)
                pandas_table = pandas_table[0]
                for row, cell in enumerate(pandas_table.values):
                    if row > 0:
                        dates.append(format_date_with_dot(current_date_from))
                        date_period.append(cell[0])
                        day.append(get_float(cell[1]))

                        if bet_option_var == 0:
                            bet_cb.append(cell[2])
                        else:
                            bet_cb.append(rate)

                        penalty.append(get_float(cell[3]))
                        total_sum += penalty[-1]
                        credit_and_penalty.append(total_sum)

                time.sleep(0.1)

        dict_result = {
            'Дата месяц/год': dates,
            'период': date_period,
            'дней': day,
            'ставка ЦБ': bet_cb,
            'неустойка': penalty,
            'долг и неуст.': credit_and_penalty
        }

        df = pd.DataFrame(dict_result)
        df.to_excel(f'{directory_for_result}/result_{debt_sum}_{datetime.now().second}.xlsx')

        time.sleep(3)
        logger.info("Successful starting of the bot")

    except:
        logger.exception("Bot error: ")
