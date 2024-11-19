import os
import time

import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

WEB_URL = "https://www.meta.com/experiences/yeeps-hide-and-seek/7276525889052788/?utm_source=vrdb.app"
# Maximum number of "Show more reviews" clicks
MAX_SMR_CLICKS = 20

# Sleep Time post "Show more reviews" clicks to let content load
SMR_SLEEP_TIME = 10


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
        print("Trying to extract reviews.")
        # Locate all review containers with the specified classes
        review_divs = self.driver.find_elements(
            By.XPATH,
            "//div[contains(@class, 'xeuugli') and contains(@class, 'x2lwn1j') and contains(@class, 'x78zum5') and "
            "contains(@class, 'xdt5ytf') and contains(@class, 'xgpatz3') and contains(@class, 'x40hh3e')]"
        )

        reviews = []
        for review_div in review_divs:
            # Extract title
            try:
                title_element = review_div.find_element(
                    By.XPATH,
                    ".//div[@class='x16g9bbj x17gzxuv xv6bue1 xm5vtmc xsp84uj x1j402mz x1wsgf3v x14imz66 x1k03ns3 "
                    "xcxolhg xjbavb x1npfmwo x16b4c32 xrm2kyc x1i6xp69 xvyeec0 x12429cg x6tc29j xbq7h4v x6jdkww "
                    "xq9mrsl']"
                )
                title = title_element.text
            except:
                try:
                    title = review_div.text.split('\n')[0]
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
                try:
                    review_time = review_div.text.split('\n')[1]
                except:
                    review_time = 'N/A'

            # Extract review comments
            try:
                review_element = review_div.find_element(
                    By.XPATH,
                    ".//div[@class='x17gzxuv x3a6nna xm5vtmc x1t2x7uc x1o1n6r0 x1wsgf3v x1c773n9 x1k03ns3 xpbi8i2 "
                    "x9820fh x1npfmwo xhj0du5 xrm2kyc xjprkx4 xlu1awn']"
                )
                review_content = review_element.text
            except:
                try:
                    review_content = review_div.text.split('\n')[2]
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
                try:
                    author = review_div.text.split('\n')[3]
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
        print("Sleeping ...")
        time.sleep(20)  # Wait for the page to load
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
                    print(f"Clicking \"Show more reviews\" Button. Count - {click_counts + 1}")

                    time.sleep(SMR_SLEEP_TIME)  # Wait for the next set of reviews to load
                    click_counts += 1

                except NoSuchElementException:
                    print("No more 'Show more reviews' buttons found.")
                    break

            reviews = self.extract_reviews()
            print(f"Reviews Extracted - {len(reviews)}")
        except Exception as e:
            print(f"Exception while scraping: {e}")
        finally:
            self.close_driver()

        return reviews

    def save_game_reviews(self, reviews, game_name):
        df = pd.DataFrame(reviews)
        directory_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "Games Reviews"))
        os.makedirs(directory_path, exist_ok=True)
        file_path = os.path.join(directory_path, game_name + '.xlsx')
        df.to_excel(file_path, index=False)

        print(f"Reviews saved to: {file_path}")


if __name__ == '__main__':
    try:
        print("Trying to fetch data from VR_Games_Data.xlsx")
        file_path = os.path.join(os.path.dirname(__file__), "..", "VR_Games_Data.xlsx")

        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            exit(1)

        games_list_df = pd.read_excel(file_path)

        if games_list_df.empty:
            print("The file is empty or invalid.")
            exit(1)

        if 'store_link' not in games_list_df.columns:
            print("The required column 'store_link' is missing.")
            exit(1)

        meta_extractor = MetaReviewsExtractor()

        for index, row in games_list_df.iterrows():
            store_link = row['store_link']
            print(f"Processing Game - {row['name']}")
            if not isinstance(store_link, str) or not store_link.strip():
                print(f"Invalid store link at row {index}. Skipping...")
                continue

            try:
                reviews = meta_extractor.scrape_reviews(store_link)
                if reviews:
                    game_name = store_link.split('/')[-1].split('?')[0]
                    meta_extractor.save_game_reviews(reviews, game_name)
                else:
                    print(f"No reviews found for the link: {store_link}")
            except Exception as inner_e:
                print(f"Error processing link {store_link} - {inner_e}")

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
