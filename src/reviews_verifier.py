import os
import re
import time
import pandas as pd
from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import traceback

from src.selenium_reviews_extractor import MetaReviewsExtractor


class MetaReviewScraper:
    def __init__(self, games_folder, output_log_path):
        """
        Initialize the scraper with folder paths and logging

        Args:
            games_folder (str): Path to folder containing Excel files
            output_log_path (str): Path to log Excel file
        """
        self.games_folder = games_folder
        self.output_log_path = output_log_path

        # Setup Chrome options
        chrome_options = Options()
        chrome_options.add_argument("--headless")  # Run in background

        # Setup WebDriver
        self.driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=chrome_options
        )

        # Create output log if not exists
        if not os.path.exists(output_log_path):
            pd.DataFrame(columns=['filename', 'url', 'ratings', 'reviews', 'file_reviews', 'processed_at']).to_excel(
                output_log_path, index=False)

    def extract_reviews(self, meta_url):
        """
        Extract reviews from Meta experience page using full class name XPATH

        Args:
            meta_url (str): URL of Meta experience

        Returns:
            tuple: (ratings, reviews) or (None, None) if extraction fails
        """
        try:
            meta_extractor = MetaReviewsExtractor()
            meta_extractor.start_driver()
            meta_extractor.driver.get(meta_url)
            print("Sleeping ...")
            time.sleep(20)  # Wait for the page to load

            try:
                # Try to find the ratings and reviews span using multiple strategies
                try:
                    # Strategy 1: Direct class match
                    ratings_reviews_span = meta_extractor.driver.find_element(
                        By.XPATH,
                        "//span[contains(@class, 'x16g9bbj') and contains(text(), 'ratings') and contains(text(), "
                        "'reviews')]"
                    )
                except:
                    print("Can't Find Number of Reviews.")
                    return None, None

                # Extract text and parse ratings and reviews
                ratings_reviews_text = ratings_reviews_span.text
                match = re.search(r'(\d+)\s*ratings,\s*(\d+)\s*reviews', ratings_reviews_text)
                if match:
                    ratings, reviews = match.groups()
                    print(f"Ratings: {ratings}, Reviews: {reviews}")
                    return ratings, reviews
            except NoSuchElementException:
                print("No more 'Show more reviews' buttons found.")
                return None, None

        except Exception as e:
            print(f"Error extracting reviews: {e}")
            traceback.print_exc()

        return None, None

    def process_files(self):
        """
        Process new Excel files in the games folder
        """
        # Read existing log
        log_df = pd.read_excel(self.output_log_path)
        processed_files = set(log_df['filename'])

        # Find new files
        for filename in os.listdir(self.games_folder):
            if filename.endswith('.xlsx') and filename not in processed_files:
                file_path = os.path.join(self.games_folder, filename)

                try:
                    # Extract ID from filename
                    meta_id = os.path.splitext(filename)[0]
                    meta_url = f"https://www.meta.com/experiences/{meta_id}"

                    # Read Excel file to count reviews
                    df = pd.read_excel(file_path)
                    file_review_count = len(df)

                    # Extract online reviews
                    ratings, reviews = self.extract_reviews(meta_url)

                    # Log results
                    new_entry = pd.DataFrame({
                        'filename': [filename],
                        'url': [meta_url],
                        'ratings': [ratings],
                        'reviews': [reviews],
                        'file_reviews': [file_review_count],
                        'processed_at': [pd.Timestamp.now()]
                    })

                    log_df = pd.concat([log_df, new_entry], ignore_index=True)
                    log_df.to_excel(self.output_log_path, index=False)

                    print(f"Processed {filename}: {reviews} online reviews, {file_review_count} file reviews")

                except Exception as e:
                    print(f"Error processing {filename}: {e}")

    def continuous_monitoring(self, interval=60):
        """
        Continuously monitor folder for new files

        Args:
            interval (int): Seconds between checks, default 60
        """
        while True:
            self.process_files()
            time.sleep(interval)

    def __del__(self):
        """Close browser on object deletion"""
        if hasattr(self, 'driver'):
            self.driver.quit()


# Example Usage
if __name__ == "__main__":
    scraper = MetaReviewScraper(
        games_folder=os.path.join(os.path.dirname(__file__), "..", "Games Reviews"),
        output_log_path=os.path.join(os.path.dirname(__file__), "..", "Reviews_Verification.xlsx")
    )
    scraper.continuous_monitoring()
