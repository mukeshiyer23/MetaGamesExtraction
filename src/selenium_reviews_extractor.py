import json
import msvcrt
import os
import re
import time
from functools import partial
from multiprocessing import Pool
from typing import List, Dict, Any
from selenium.webdriver.support import expected_conditions as EC

import numpy as np
import pandas as pd
from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

# Maximum number of "Show more reviews" clicks
MAX_SMR_CLICKS = 2000

# Sleep Time post "Show more reviews" clicks to let content load
SMR_SLEEP_TIME = 15

# Determine number of processes
NUM_PROCESSES = max(os.cpu_count() - 13, 1)


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

            # Extract helpfulness
            try:
                helpfulness_element = review_div.find_element(
                    By.XPATH,
                    ".//span[contains(@class, 'x1heor9g') and contains(@class, 'x17gzxuv') and contains(@class, "
                    "'x1rujz1s') and contains(@class, 'xex5isp') and contains(@class, 'xsp84uj') and contains(@class, "
                    "'x658qfi') and contains(@class, 'x1wsgf3v') and contains(@class, 'xn1wy4v') and contains(@class, "
                    "'xby3lk6') and contains(@class, 'xcxolhg') and contains(@class, 'xh2n1af') and contains(@class, "
                    "'x1npfmwo') and contains(@class, 'xg94uf4') and contains(@class, 'x1yyhlu9') and contains(@class, "
                    "'x1i6xp69') and contains(@class, 'xawl3gl') and contains(@class, 'x12429cg') and contains(@class, "
                    "'x6tc29j') and contains(@class, 'xbq7h4v') and contains(@class, 'x6jdkww') and contains(@class, "
                    "'xq9mrsl')]"
                )
                helpfulness = helpfulness_element.text
            except:
                try:
                    helpfulness = review_div.text.split('\n')[3]
                except:
                    helpfulness = 'N/A'

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

        return result

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

        # Dynamic wait for page load
        WebDriverWait(self.driver, 30).until(
            lambda driver: driver.execute_script('return document.readyState') == 'complete'
        )

        click_counts = 0
        reviews = []

        game_name = url.split('/')[-1].split('?')[0]
        directory_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "Games Reviews"))

        try:
            time.sleep(5)
            additional_game_details = self.extract_additional_games_details()
            description_details = self.extract_descriptions()

            row_df = pd.DataFrame([row])

            # Add additional game details to the row
            if additional_game_details:
                for key, value in additional_game_details.items():
                    row_df[key] = value

            # Add description if available
            if description_details and 'description' in description_details:
                row_df['description'] = description_details['description']

            self.save_to_files(row_df, {
                'excel': os.path.join(directory_path,'games.xlsx'),
                'json': os.path.join(directory_path,'games.json'),
            })

            # Loop to click "Show more reviews" button
            try:
                while click_counts <= MAX_SMR_CLICKS:
                    try:
                        show_more_button = WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH,
                                                        "//div[contains(@class, 'x78zum5') and contains(@class, "
                                                        "'xl56j7k')]/span[text()='Show more reviews']"))
                        )
                        show_more_button.click()
                        print(f"Clicking \"Show more reviews\" Button. Count - {click_counts + 1}")

                        click_counts += 1

                    except NoSuchElementException:
                        print("No more 'Show more reviews' buttons found.")
                        break
                    except Exception as e:
                        print(f"No more 'Show more reviews' buttons found. - {e}")
                        break
            except Exception as e:
                print(f"No more 'Show more reviews' buttons found. - {e}")

            reviews = self.extract_reviews()
            print(f"Reviews Extracted - {len(reviews)}")

        finally:
            # Check for low review count and log
            if len(reviews) <= 25:
                lock_file_path = os.path.join(directory_path, 'skipped_games.txt.lock')
                skipped_games_path = os.path.join(directory_path, 'skipped_games.txt')

                try:
                    # Open lock file
                    with open(lock_file_path, 'w') as lock_file:
                        try:
                            # Attempt to lock the file
                            msvcrt.locking(lock_file.fileno(), msvcrt.LK_LOCK, 1)

                            # Write to skipped games file
                            with open(skipped_games_path, 'a') as f:
                                f.write(f"Skipping game: {game_name}\n")

                            # Clear reviews
                            reviews = []

                        finally:
                            # Unlock the file
                            msvcrt.locking(lock_file.fileno(), msvcrt.LK_UNLCK, 1)

                except Exception as e:
                    print(f"Error writing to skipped games file: {str(e)}")

            self.close_driver()

        return reviews

    def save_to_files(self, data, output_files: dict) -> None:
        """
        Save data to multiple file types with thread-safe file writing

        :param data: DataFrame or dictionary to be saved
        :param output_files: Dict with file paths and their types
                             Example: {
                                 'excel': '/path/to/output.xlsx',
                                 'json': '/path/to/output.json',
                             }
        """
        try:
            for file_path in output_files.values():
                os.makedirs(os.path.dirname(file_path), exist_ok=True)

            first_file_path = list(output_files.values())[0]
            lock_file_path = first_file_path + '.lock'

            with open(lock_file_path, 'w') as lock_file:
                try:
                    msvcrt.locking(lock_file.fileno(), msvcrt.LK_LOCK, 1)

                    for file_type, file_path in output_files.items():
                        if file_type == 'excel':
                            self._save_to_excel(data, file_path)
                        elif file_type == 'json':
                            self._save_to_json(data, file_path)
                        else:
                            raise ValueError(f"Unsupported file type: {file_type}")

                finally:
                    # Release the lock
                    msvcrt.locking(lock_file.fileno(), msvcrt.LK_UNLCK, 1)
        except Exception as e:
            print(f"Error saving data to files: {str(e)}")
            raise

    def _save_to_excel(self, data, file_path):
        """Save DataFrame to Excel"""
        try:
            existing_df = pd.read_excel(file_path) if os.path.exists(file_path) else pd.DataFrame()

            if not existing_df.empty:
                data = data.reindex(columns=existing_df.columns)

            combined_df = pd.concat([existing_df, data], ignore_index=True)

            columns_to_drop = ['genres', 'developer', 'publisher']
            combined_df = combined_df.drop(columns=[col for col in columns_to_drop if col in combined_df.columns])

            combined_df.columns = [self._to_camel_case(col) for col in combined_df.columns]

            with pd.ExcelWriter(file_path, mode='w', engine='openpyxl') as writer:
                combined_df.to_excel(writer, index=False, header=True, sheet_name='Data')

        except Exception as e:
            print(f"Error saving to Excel: {str(e)}")
            raise

    def _save_to_json(self, data, file_path):
        """Save DataFrame to JSON"""
        try:
            existing_data = []
            if os.path.exists(file_path):
                with open(file_path, 'r') as f:
                    existing_data = json.load(f)

            new_data = data.to_dict(orient='records') if isinstance(data, pd.DataFrame) else data

            combined_data = existing_data + new_data

            with open(file_path, 'w') as f:
                json.dump(combined_data, f, indent=2)

        except Exception as e:
            print(f"Error saving to JSON: {str(e)}")
            raise

    def _save_to_txt(self, data, file_path):
        """Save data to TXT"""
        try:
            if isinstance(data, pd.DataFrame):
                data_str = data.to_csv(index=False)
            elif isinstance(data, list):
                data_str = '\n'.join(str(item) for item in data)
            else:
                data_str = str(data)

            mode = 'a' if os.path.exists(file_path) else 'w'
            with open(file_path, mode) as f:
                f.write(data_str + '\n')

        except Exception as e:
            print(f"Error saving to TXT: {str(e)}")
            raise

    def _to_camel_case(self, s):
        """Convert string to camelCase"""
        parts = re.split(r'[_\s]+', s)
        return parts[0].lower() + ''.join(word.capitalize() for word in parts[1:])

    def save_game_reviews(self, reviews, game_name):
        df = pd.DataFrame(reviews)
        directory_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "Games Reviews"))
        os.makedirs(directory_path, exist_ok=True)
        xlsx_file_path = os.path.join(directory_path, 'xlsx_games_reviews', game_name + '.xlsx')
        csv_file_path = os.path.join(directory_path, 'csv_games_reviews', f"{game_name}_{len(reviews)}.csv")
        df.to_excel(xlsx_file_path, index=False)
        df.to_csv(csv_file_path, index=False)


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
                    directory_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "Games Reviews"))
                    vr_games_data_path = os.path.join(directory_path, 'VR_Games_data.xlsx')
                    skipped_games_path = os.path.join(directory_path, 'skipped_games.txt')
                    game_name = store_link.split('/')[-1].split('?')[0]

                    def is_game_processed(store_link):
                        if os.path.exists(vr_games_data_path):
                            df = pd.read_excel(vr_games_data_path)

                            if 'store_link' in df.columns and store_link in df['store_link'].values:
                                print(f"The store link '{store_link}' is already processed.")
                                return True

                        if os.path.exists(skipped_games_path):
                            with open(skipped_games_path, 'r') as f:
                                skipped_games = f.readlines()

                            for line in skipped_games:
                                if f"Skipping game: {game_name}" in line:
                                    print(f"The game '{game_name}' is already skipped.")
                                    return True

                        return False

                    if not is_game_processed(store_link):
                        print(f"Processing Game - {game_name}")

                        reviews = meta_extractor.scrape_reviews(store_link, row, MAX_SMR_CLICKS)
                        if len(reviews) > 0:
                            meta_extractor.save_game_reviews(reviews, game_name)
                            games_processed += 1

                            if games_processed % COOLDOWN_INTERVAL == 0:
                                print(
                                    f"Process {chunk_id} - Cooling down for {COOLDOWN_DURATION} seconds after processing {COOLDOWN_INTERVAL} games...")
                                time.sleep(COOLDOWN_DURATION)
                        else:
                            print(f"Process {chunk_id} - No reviews found for: {store_link}")
                    else:
                        print(f"Game - {game_name} already processed")

                except Exception as e:
                    print(f"Process {chunk_id} - Error processing {store_link}: {str(e)}")

        except Exception as e:
            print(f"Process {chunk_id} encountered an error: {str(e)}")
        finally:
            meta_extractor.driver.quit()  # Clean up Selenium driver

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
