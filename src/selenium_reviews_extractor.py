import time

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

DRIVER_PATH = "C://Users//Mukes//validus//dataservice-dataservice//pythonProject//chromedriver.exe"
WEB_URL = "https://www.meta.com/experiences/yeeps-hide-and-seek/7276525889052788/?utm_source=vrdb.app"
# Maximum number of "Show more reviews" clicks
MAX_SMR_CLICKS = 20

class MetaReviewsExtractor:
    def __init__(self):
        self.chrome_options = Options()
        # self.chrome_options.add_argument('--headless')
        self.chrome_options.add_argument('--no-sandbox')
        self.chrome_options.add_argument('--disable-ev-shm-usage')
        self.driver = None

    def start_driver(self):
        """Initializes the Driver"""
        self.driver = webdriver.Chrome(options=self.chrome_options)

    def close_driver(self):
        """Close the Webdriver"""
        if self.driver:
            self.driver.quit()

    def extract_reviews(self):
        # Locate all review containers with the specified classes
        review_divs = self.driver.find_elements(
            By.XPATH,
            "//div[contains(@class, 'xeuugli') and contains(@class, 'x2lwn1j') and contains(@class, 'x78zum5') and contains(@class, 'xdt5ytf') and contains(@class, 'xgpatz3') and contains(@class, 'x40hh3e')]"
        )

        reviews = []
        for review_div in review_divs:
            # Extract title
            try:
                title_element = review_div.find_element(
                    By.XPATH,
                    ".//div[contains(@class, 'x16g9bbj') and contains(@class, 'x17gzxuv') and contains(@class, "
                    "'xv6bue1') and contains(@class, 'xm5vtmc') and contains(@class, 'xsp84uj') and contains(@class, "
                    "'x1j402mz') and contains(@class, 'x1wsgf3v') and contains(@class, 'x14imz66') and contains("
                    "@class, 'x1k03ns3') and contains(@class, 'xcxolhg') and contains(@class, 'xjbavb') and contains("
                    "@class, 'x1npfmwo') and contains(@class, 'x16b4c32') and contains(@class, 'xrm2kyc') and "
                    "contains(@class, 'x1i6xp69') and contains(@class, 'xvyeec0') and contains(@class, 'x12429cg') "
                    "and contains(@class, 'x6tc29j') and contains(@class, 'xbq7h4v') and contains(@class, 'x6jdkww') "
                    "and contains(@class, 'xq9mrsl')]"
                )
                title = title_element.text
            except:
                title = 'N/A'

            # Extract rating (stars)
            try:
                stars_div = review_div.find_element(By.CLASS_NAME, 'x3nfvp2')
                rating = len(stars_div.find_elements(By.CLASS_NAME, 'xjbqb8w'))
            except:
                rating = 0

            # Extract time
            try:
                time_element = review_div.find_element(
                    By.XPATH,
                    ".//span[contains(@class, 'x16g9bbj') and contains(@class, 'x17gzxuv') and contains(@class, "
                    "'x3a6nna') and contains(@class, 'xm5vtmc') and contains(@class, 'x1t2x7uc') and contains(@class, "
                    "'x1o1n6r0') and contains(@class, 'x1wsgf3v') and contains(@class, 'x1c773n9') and contains("
                    "@class, 'x1k03ns3') and contains(@class, 'xpbi8i2') and contains(@class, 'x9820fh') and "
                    "contains(@class, 'x1npfmwo') and contains(@class, 'xhj0du5') and contains(@class, 'xrm2kyc') and "
                    "contains(@class, 'xjprkx4') and contains(@class, 'xlu1awn') and contains(@class, 'x12429cg') and "
                    "contains(@class, 'x6tc29j') and contains(@class, 'xbq7h4v') and contains(@class, 'x6jdkww') and "
                    "contains(@class, 'xq9mrsl')]"
                )
                review_time = time_element.text
            except:
                review_time = 'N/A'

            # Extract review content
            try:
                review_element = review_div.find_element(
                    By.XPATH,
                    ".//div[contains(@class, 'x17gzxuv') and contains(@class, 'x3a6nna') and contains(@class, "
                    "'xm5vtmc') and contains(@class, 'x1t2x7uc') and contains(@class, 'x1o1n6r0') and contains("
                    "@class, 'x1wsgf3v') and contains(@class, 'x1c773n9') and contains(@class, 'x1k03ns3') and "
                    "contains(@class, 'xpbi8i2') and contains(@class, 'x9820fh') and contains(@class, 'x1npfmwo') and "
                    "contains(@class, 'xhj0du5') and contains(@class, 'xrm2kyc') and contains(@class, 'xjprkx4') and "
                    "contains(@class, 'xlu1awn')]"
                )
                review_content = review_element.text
            except:
                review_content = 'N/A'

            # Extract author
            try:
                author_element = review_div.find_element(
                    By.XPATH,
                    ".//span[contains(@class, 'x16g9bbj') and contains(@class, 'x17gzxuv') and contains(@class, "
                    "'x1rujz1s') and contains(@class, 'xm5vtmc') and contains(@class, 'x3voqp2') and contains(@class, "
                    "'x658qfi') and contains(@class, 'x1wsgf3v') and contains(@class, 'xn1wy4v') and contains(@class, "
                    "'x1k03ns3') and contains(@class, 'xpbi8i2') and contains(@class, 'xh2n1af') and contains(@class, "
                    "'x1npfmwo') and contains(@class, 'xg94uf4') and contains(@class, 'xrm2kyc') and contains(@class, "
                    "'xjprkx4') and contains(@class, 'xawl3gl') and contains(@class, 'x12429cg') and contains(@class, "
                    "'x6tc29j') and contains(@class, 'xbq7h4v') and contains(@class, 'x6jdkww') and contains(@class, "
                    "'xq9mrsl')]"
                )
                author = author_element.text
            except:
                author = 'N/A'

            reviews.append({
                'title': title,
                'rating': rating,
                'time': review_time,
                'content': review_content,
                'author': author
            })

        return reviews

    def scrape_reviews(self, url, MAX_SMR_CLICKS=5):
        self.start_driver()
        self.driver.get(url)
        time.sleep(3)  # Wait for the page to load
        click_counts = 0
        reviews = []

        try:
            # Loop to click "Show more reviews" button
            while click_counts <= MAX_SMR_CLICKS:
                try:
                    show_more_button = self.driver.find_element(
                        By.XPATH,
                        "//div[contains(@class, 'x78zum5') and contains(@class, 'xl56j7k')]/span[text()='Show more reviews']"
                    )

                    show_more_button.click()
                    time.sleep(10)  # Wait for the next set of reviews to load
                    click_counts += 1

                except NoSuchElementException:
                    print("No more 'Show more reviews' buttons found.")
                    break

            reviews = self.extract_reviews()

        except Exception as e:
            print(f"Exception while scraping: {e}")
        finally:
            self.close_driver()

        return reviews


if __name__ == '__main__':
    try:
        meta_extractor = MetaReviewsExtractor()
        meta_extractor.scrape_reviews(WEB_URL)
    except Exception as e:
        print(f"Exception - {e}")
