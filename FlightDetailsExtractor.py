from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from FlightExtractorHelper import flight_schools_extractor, check_page_status, get_url_list
import asyncio
import sys
import os


async def main():
    chrome_options = Options()
    if hasattr(sys, '_MEIPASS'):
        extension_path = os.path.join(sys._MEIPASS, 'adblock.crx')
    else:
        extension_path = 'adblock.crx'
    # chrome_options.add_extension(extension_path)
    chrome_options.add_argument("--disable-features=NetworkService")
    # chrome_options.add_argument("--headless")
    driver = webdriver.Chrome(options=chrome_options)
    urls = await get_url_list("https://www.bestaviation.net/flight_school/", driver)
    try:
        for url in urls:
            status = await check_page_status(url)
            if status == 200:
                driver.get(url)
                await flight_schools_extractor(driver)
            else:
                print(f"Got {status} from {url}")
    except Exception as e:
        print(f"Exception in for loop main: {e}")
    finally:
        driver.close()


if __name__ == "__main__":
    asyncio.run(main())
