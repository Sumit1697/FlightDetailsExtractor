from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from FlightExtractorHelper import go_to_url_and_proceed_login, get_nbaa_urls, go_to_url_and_start_scraping
from concurrent.futures import ThreadPoolExecutor
from selenium import webdriver
import sys, os, asyncio


async def main():
    try:
        chrome_options = Options()
        if hasattr(sys, '_MEIPASS'):
            extension_path = os.path.join(sys._MEIPASS, 'adblock.crx')
        else:
            extension_path = 'adblock.crx'
        # chrome_options.add_extension(extension_path)
        chrome_options.add_argument("--disable-features=NetworkService")
        # chrome_options.add_argument("--headless")
        driver = webdriver.Chrome(options=chrome_options)
        wait = WebDriverWait(driver, 2)
        is_logged_in = go_to_url_and_proceed_login("nanicsit@gmail.com", "Test#123", driver,
                                                   "https://connect.nbaa.org/", wait)
        print(f"Is logged in: {is_logged_in}")
        if is_logged_in:
            # actions = ActionChains(driver)
            urls = get_nbaa_urls(driver, wait)
            if len(urls) > 0:
                [await go_to_url_and_start_scraping(url, wait, driver) for url in urls[::-1]]
    except Exception as e:
        print(f"Exception While Initializing Main Function: {e}")
    finally:
        input("Press Enter to close the browser")
        # driver.close()


if __name__ == "__main__":
    asyncio.run(main())
