import scrapy
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
import time

class LazadaFeedbackSpider(scrapy.Spider):
    name = 'lazada_feedback'

    def __init__(self, *args,**kwargs):
        super(LazadaFeedbackSpider, self).__init__(*args, **kwargs)
        self.workbook = Workbook()
        self.sheet = self.workbook.active
        self.sheet.title = "Lazada Feedback"
        self.sheet.append(["Content"])  # Write header
        self.comments_count = 0

        # Set up Selenium
        chrome_options = Options()
        # chrome_options.add_argument("--headless")  # Commented out for debugging
        self.driver = webdriver.Chrome(options=chrome_options)

    def start_requests(self):
        url = 'https://www.lazada.com.ph/products/easypc-rakk-sari-v2-rgb-usb-gaming-keyboard-for-better-gaming-experience-for-laptop-and-desktop-pc-i3439607685-s17641150548.html?c=&channelLpJumpArgs=&freeshipping=1&fs_ab=2&fuse_fs=&lang=en&location=Metro%20Manila~Pasig&price=771&priceCompare=skuId%3A17641150548%3Bsource%3Alazada-search-voucher%3Bsn%3A9593c41546eae682aaaea0898572f852%3BunionTrace%3Aa3b54d9917222641588933455e%3BoriginPrice%3A77100%3BvoucherPrice%3A77100%3BdisplayPrice%3A77100%3BsinglePromotionId%3A-1%3BsingleToolCode%3AmockedSalePrice%3BvoucherPricePlugin%3A1%3BbuyerId%3A0%3Btimestamp%3A1722264159397&ratingscore=4.916666666666667&request_id=9593c41546eae682aaaea0898572f852&review=120&sale=375&search=1&source=search&stock=1'
        self.logger.info(f"Starting requests with URL: {url}")
        yield scrapy.Request(url=url, callback=self.parse, dont_filter=True)

    def parse(self, response):
        self.logger.info(f"Parsing URL: {response.url}")
        self.driver.get(response.url)

        # Wait for the reviews to load
        try:
            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.mod-reviews"))
            )
            self.logger.info("Reviews section loaded successfully")
        except Exception as e:
            self.logger.error(f"Timed out waiting for page to load: {str(e)}")
            self.logger.info(f"Current page source: {self.driver.page_source[:500]}...")
            return

        # Scroll down to load more reviews
        scroll_attempts = 0
        while scroll_attempts < 5:
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            scroll_attempts += 1
            self.logger.info(f"Scrolled {scroll_attempts} times")

        # Extract feedback
        try:
            content_elements = self.driver.find_elements(By.CSS_SELECTOR, "div.mod-reviews div.item div.item-content:not(.item-content--seller-reply) > div.content")
            self.logger.info(f"Found {len(content_elements)} content elements")
        except Exception as e:
            self.logger.error(f"Error finding content elements: {str(e)}")
            content_elements = []

        if not content_elements:
            self.logger.warning("No content elements found. Current page structure:")
            self.logger.warning(self.driver.execute_script("return document.body.innerHTML;")[:2000])

        for content in content_elements:
            try:
                content_text = content.text.strip()
                if content_text:
                    self.sheet.append([content_text])
                    self.comments_count += 1
                    self.logger.info(f"Extracted comment {self.comments_count}: {content_text[:50]}...")
                    yield {'content': content_text}
                else:
                    self.logger.warning("Empty content for a feedback element")
            except Exception as e:
                self.logger.error(f"Error extracting feedback: {str(e)}")

        # Save the workbook after processing each page
        try:
            self.workbook.save('lazada_feedback_content.xlsx')
            self.logger.info(f"Saved {self.comments_count} comments to Excel file")
        except Exception as e:
            self.logger.error(f"Error saving Excel file: {str(e)}")

        # Check for next page
        try:
            next_page = self.driver.find_element(By.CSS_SELECTOR, "button.next-pagination-item.next")
            if next_page.is_enabled():
                self.logger.info("Clicking next page")
                next_page.click()
                time.sleep(5)
                yield scrapy.Request(url=self.driver.current_url, callback=self.parse, dont_filter=True)
            else:
                self.logger.info("No more pages to follow")
        except Exception as e:
            self.logger.error(f"Error checking for next page: {str(e)}")

    def closed(self, reason):
        try:
            self.workbook.save('lazada_feedback_content.xlsx')
            self.logger.info(f"Final save: Total comments extracted: {self.comments_count}")
        except Exception as e:
            self.logger.error(f"Error during final save: {str(e)}")
        self.driver.quit()

# To run the spider, use:
# scrapy runspider lazada_feedback_spider.py -s LOG_LEVEL=DEBUG
