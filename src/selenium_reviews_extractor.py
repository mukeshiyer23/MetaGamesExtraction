import datetime
import os
import time
from functools import partial
from multiprocessing import Pool
from typing import List, Dict, Any

import numpy as np
import pandas as pd
from selenium import webdriver
from selenium.common import NoSuchElementException, TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

# Maximum number of "Show more reviews" clicks
MAX_SMR_CLICKS = 2000

# Sleep Time post "Show more reviews" clicks to let content load
SMR_SLEEP_TIME = 15

# Determine number of processes
NUM_PROCESSES = max(os.cpu_count() - 15, 1)


class MetaReviewsExtractor:
    def __init__(self):
        self.chrome_options = Options()
        self.chrome_options.add_argument('--headless')
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

            # Extract helpfulness
            try:
                helpful_votes_element = review_div.find_element(By.CSS_SELECTOR,
                                                                ".x1heor9g.x17gzxuv.x1rujz1s.xex5isp.xsp84uj.x658qfi"
                                                                ".x1wsgf3v.xn1wy4v.xby3lk6.xcxolhg.xh2n1af.x1npfmwo"
                                                                ".xg94uf4.x1yyhlu9.x1i6xp69.xawl3gl.x12429cg.x6tc29j"
                                                                ".xbq7h4v.x6jdkww.xq9mrsl.x1iorvi4.xjkvuk6")
                helpfulness = int(helpful_votes_element.text.strip().split()[2])
            except:
                helpfulness = 'N/A'

            try:
                time_ago_element = review_div.find_element(By.CSS_SELECTOR,
                                                           ".x16g9bbj.x17gzxuv.x3a6nna.xm5vtmc.x1t2x7uc.x1o1n6r0.x1wsgf3v"
                                                           ".x1c773n9.x1k03ns3.xpbi8i2.x9820fh.x1npfmwo.xhj0du5.xrm2kyc"
                                                           ".xjprkx4.xlu1awn.x12429cg.x6tc29j.xbq7h4v.x6jdkww.xq9mrsl")
                time_str = time_ago_element.text.strip()

                now = datetime.datetime.now()

                if "ago" in time_str:
                    if "A" or "An" in time_str:
                        time_diff = 1
                    else:
                        time_diff = int(time_str.split()[0])

                    if "day" in time_str:
                        past_time = now - datetime.timedelta(days=time_diff)
                    elif "hour" in time_str:
                        past_time = now - datetime.timedelta(hours=time_diff)

                    review_time = past_time.strftime('%m-%d-%Y %H:%M')

                else:
                    if "," not in time_str:
                        month_day = time_str.split(" at ")[0]
                        time_part = time_str.split(" at ")[1]
                        time_str = f"{month_day}, {now.year} at {time_part}"

                    review_time = datetime.datetime.strptime(time_str, '%b %d, %Y at %I:%M %p').strftime(
                        '%m-%d-%Y %H:%M')
            except:
                review_time = 'N/A'

            reviews.append({
                'title': title,
                'rating': rating,
                'time': review_time,
                'content': review_content,
                'author': author,
                'helpful_votes': helpfulness
            })

        return reviews

    def extract_ad(self, data):
        # Define the keys we want to extract
        keys_to_extract = [
            'Game modes', 'Multiplayer', 'Supported player modes',
            'Supported controllers', 'Supported platforms',
            'Category', 'Genres', 'Languages', 'Version',
            'Developer', 'Publisher', 'Website', 'Release date',
            'Space required', 'Comfort level', 'Internet connection'
        ]

        # Dictionary to store the extracted details
        game_details = {}

        # Iterate through the list
        for i in range(len(data)):
            # Check if the current item is in our keys to extract
            if data[i] in keys_to_extract:
                # If the key exists, add it to the dictionary with its next value
                if i + 1 < len(data):
                    game_details[data[i]] = data[i + 1]

        return game_details

    def extract_additional_games_details(self):
        try:
            print("Trying to game details.")
            target_div = self.driver.find_element(
                By.XPATH,
                ".//div[contains(@class, 'x78zum5') and contains(@class, 'x1l7klhg') and contains(@class, 'x1iyjqo2') "
                "and contains(@class, 'x2lah0s') and contains(@class, 'x1a02dak') and contains(@class, 'xd2bs7b') and "
                "contains(@class, 'x5bj0eh') and contains(@class, 'x1sje56t') and contains(@class, 'x2b88hg') and "
                "contains(@class, 'x17tu2g0') and contains(@class, 'xnjo89n') and contains(@class, 'xo2o5nc') and "
                "contains(@class, 'xv9pgs7') and contains(@class, 'xjfzuef')]"
            )
            data = target_div.text.split('\n')
            result = self.extract_ad(data)
        except:
            return {}
        return result

    def extract_pegi_rating(self):
        try:
            return None

        except Exception as e:
            print(f"Rating extraction failed: {e}")
            return None

    def extract_descriptions(self):
        try:
            show_more_button = self.driver.find_element(
                By.XPATH,
                "//div[@role='button' and contains(@class, 'x1i10hfl') and contains(text(), 'more...')]"
            )

            show_more_button.click()
            print("Clicked 'more' button.")
        except Exception as e:
            print(f"'More' button not found or could not be clicked: {e}")

        try:
            target_div = self.driver.find_element(
                By.XPATH,
                ".//div[contains(@class, 'xeuugli') and contains(@class, 'x2lwn1j') and contains(@class, 'x78zum5') "
                "and contains(@class, 'xdt5ytf') and contains(@class, 'xozqiw3') and contains(@class, 'x3pnbk8')]"
            )
            data = target_div.text.split('\n')
            result = ' '.join(data[:-1])
            return {'description': result}
        except Exception as e:
            print(f"Target div not found or data extraction failed: {e}")
            return {}

    def scrape_reviews(self, url, row, MAX_SMR_CLICKS=5):
        self.start_driver()
        self.driver.get(url)
        print("Sleeping ...")
        time.sleep(20)  # Wait for the page to load
        click_counts = 0
        reviews = []

        try:
            additional_game_details = self.extract_additional_games_details()
            description_details = self.extract_descriptions()

            # Add additional game details to the row
            if additional_game_details:
                for key, value in additional_game_details.items():
                    row.loc[key] = value

            # Add description if available
            if description_details and 'description' in description_details:
                row.loc['description'] = description_details['description']

            # Loop to click "Show more reviews" button
            while click_counts <= MAX_SMR_CLICKS:
                try:
                    show_more_button = self.driver.find_element(
                        By.XPATH,
                        "//div[contains(@class, 'x78zum5') and contains(@class, 'xl56j7k')]/span[text()='Show more "
                        "reviews']"
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
        xlsx_file_path = os.path.join(directory_path, 'xlsx_games_reviews', game_name + '.xlsx')
        csv_file_path = os.path.join(directory_path, 'csv_games_reviews', f"{game_name}_{len(reviews)}.csv")
        df.to_excel(xlsx_file_path, index=False)
        df.to_csv(csv_file_path, index=False)
        print(f"Reviews saved to: {file_path}")


class ParallelMetaReviewsExtractor:
    @staticmethod
    def chunk_dataframe(df: pd.DataFrame, chunks: int) -> List[pd.DataFrame]:
        """Split dataframe into roughly equal chunks."""
        return np.array_split(df, chunks)

    @staticmethod
    def process_chunk(chunk: pd.DataFrame, chunk_id: int) -> List[Dict[str, Any]]:
        """Process a chunk of the dataframe in a separate process."""
        meta_extractor = MetaReviewsExtractor()  # Creates new Selenium instance
        results = []
        games_processed = 0  # Counter for processed games
        COOLDOWN_INTERVAL = 10  # Process 10 games before cooling
        COOLDOWN_DURATION = 0  # Cool down for 60 seconds

        try:
            print(f"Process {chunk_id} started with {len(chunk)} games")

            for index, row in chunk.iterrows():
                store_link = row['store_link']
                print(f"Process {chunk_id} - Processing Game - {row['name']}")

                if not isinstance(store_link, str) or not store_link.strip():
                    print(f"Process {chunk_id} - Invalid store link at row {index}. Skipping...")
                    continue

                try:
                    game_name = store_link.split('/')[-1].split('?')[0]
                    print(f"Processing Game - {game_name}")

                    reviews = meta_extractor.scrape_reviews(store_link, row, MAX_SMR_CLICKS)
                    if reviews:
                        meta_extractor.save_game_reviews(reviews, game_name)
                        games_processed += 1  # Increment counter only for successful extractions

                        # Check if we need a cooling period
                        if games_processed % COOLDOWN_INTERVAL == 0:
                            print(
                                f"Process {chunk_id} - Cooling down for {COOLDOWN_DURATION} seconds after processing {COOLDOWN_INTERVAL} games...")
                            time.sleep(COOLDOWN_DURATION)
                    else:
                        print(f"Process {chunk_id} - No reviews found for: {store_link}")

                except Exception as e:
                    print(f"Process {chunk_id} - Error processing {store_link}: {str(e)}")

        except Exception as e:
            print(f"Process {chunk_id} encountered an error: {str(e)}")
        finally:
            meta_extractor.driver.quit()  # Clean up Selenium driver
            # Append the updated chunk to the Excel file
            try:
                # Check if file exists
                if os.path.exists('output.xlsx'):
                    with pd.ExcelWriter('output.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                        chunk.to_excel(writer, index=False, header=False, startrow=writer.sheets['Sheet1'].max_row)
                else:
                    # Create a new file if it doesn't exist
                    chunk.to_excel('output.xlsx', index=False)

                print(f"Process {chunk_id} - Updated data appended to {'output.xlsx'}")
            except Exception as e:
                print(f"Process {chunk_id} - Error appending updated data: {str(e)}")
        return results


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

        # Split dataframe into chunks
        chunks = ParallelMetaReviewsExtractor.chunk_dataframe(games_list_df, NUM_PROCESSES)

        # Create process pool and process chunks in parallel
        with Pool(processes=NUM_PROCESSES) as pool:
            # Create partial function with chunk_id
            process_chunk_with_id = partial(
                ParallelMetaReviewsExtractor.process_chunk
            )
            # Map chunks to processes with their IDs
            all_results = pool.starmap(
                process_chunk_with_id,
                [(chunk, i) for i, chunk in enumerate(chunks)]
            )

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
