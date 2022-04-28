from bs4 import BeautifulSoup
from openpyxl import Workbook
from selenium import webdriver
from selenium.common.exceptions import *
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import time
from webdriver_manager.chrome import ChromeDriverManager
from urllib import parse
import urllib.request as req
import urllib3.exceptions


class Crawler:
    @classmethod
    def __get_element(cls, driver, xpath):
        element = WebDriverWait(driver, 10).until(
            ec.element_to_be_clickable((By.XPATH, xpath))
        )
        return element

    @classmethod
    def __get_raw_members(cls, server, guild):
        try:
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
            driver.get(
                "https://" + parse.quote("maple.gg/guild/%s/%s" % (server, guild))
            )

            try:
                xpath_sync = "//button[@id='btn-sync']"
                xpath_load = "//*[@id='guild-content']/section/div[1]/div[1]/section/div[2]/div/div[1]/b/a"

                element = cls.__get_element(driver, xpath_sync)
                element.click()

                alert = driver.switch_to.alert
                alert.accept()

                time.sleep(5)
                _ = cls.__get_element(driver, xpath_load)
            finally:
                html = driver.page_source

            driver.quit()

        except (
            SessionNotCreatedException,
            TimeoutException,
            urllib3.exceptions.ProtocolError,
        ):
            return None
        return html

    @classmethod
    def __get_mulung(cls, name):
        res = req.urlopen("https://" + parse.quote("maple.gg/u/%s" % name)).read()
        data = BeautifulSoup(res, "html.parser")

        notice = data.select(
            "#app > div.card.border-bottom-0 > div > section > div.row.text-center "
            "> div:nth-child(1) > section > div > div > div > h1"
        )

        return notice[0].text.split("\n")[0] if notice else "?"

    @classmethod
    def get_members(cls, server, guild):
        html = cls.__get_raw_members(server, guild)
        data = BeautifulSoup(html, "html.parser").findChildren("div")
        lines = list(filter(lambda x: x != "", data[0].text.split("\n")))
        result = []
        for i in range(len(lines)):
            if "마지막 활동일" in lines[i]:
                result.append([lines[i - 2], lines[i - 1], lines[i]])

        return list(
            map(
                lambda x: [
                    x[0],
                    x[1].split("/")[0],
                    x[1].split("/")[1],
                    cls.__get_mulung(x[0]) + "층",
                    x[2].split(": ")[1],
                ],
                result[3:],
            )
        )


if __name__ == "__main__":
    members = Crawler.get_members("scania", "아이엠캔들")
    write_wb = Workbook()
    write_ws = write_wb.active
    for member in members:
        write_ws.append(member)
    write_wb.save("D:/Downloads/아이엠캔들/2022/아이엠캔들길드원0428list.xlsx")
