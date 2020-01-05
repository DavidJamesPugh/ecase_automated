"""
Some functions needed for ecase_downloader
"""
from calendar import monthrange

from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait


def click_previous_month_button(driver, count: int):
    """
    Clicks the eCase previous_month button the same amount of times as count
    :param driver: The selenium object
    :param count: Number of times to click the previous month button
    :return:
    """
    wait = WebDriverWait(driver, 10)
    while count > 0:
        # the next three statements are necessary, as the previous month
        # button in ecase is finnicky
        # assigns the previous month button to a variable, and waits for it to be available
        prev_month = wait.until(ec.element_to_be_clickable((By.CSS_SELECTOR,
                                                            'body > div.Zebra_DatePicker.dp_visible > '
                                                            'table.dp_header > tbody > tr > td.dp_previous')))
        # hovering over previous month button
        ActionChains(driver).move_to_element(prev_month).perform()
        # clicking the date that we hovered over
        driver.find_element_by_css_selector(
            'body > div.Zebra_DatePicker.dp_visible > table.dp_header > tbody > tr > td.dp_previous').click()
        count -= 1


def date_selector(month: int, year: int):
    """
    Returns a tuple with the starting day, ending day of the month for use in
    ecases date selector. These three values give the positional data needed
    to select the first and last day in a given month
    :param month: int from 1-12 denoting month
    :param year: 4 digit int denoting the year

    :return
    Calendar_rows denotes which row the end_date is placed on
    start_date is an int from 1-7, denotes what day of the week starts the month
    end_date is an int from 1-7, denotes what day of the week ends the month
    Sunday is 7

    TESTS WRITTEN AND WORKING 12/2019 in tests.py DateSelectorTests
    """

    end_cal_dates = monthrange(year, month)
    start_date = end_cal_dates[0] + 1
    end_date = end_cal_dates[0] + end_cal_dates[1] % 4
    calendar_rows = 6

    if end_date > 7:
        end_date -= 7
        calendar_rows += 1

    return calendar_rows, start_date, end_date
