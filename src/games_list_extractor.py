import json

import requests
import pandas as pd
from bs4 import BeautifulSoup
import re
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class VRDBExtractor:
    def __init__(self):
        self.base_url = "https://vrdb.app/games"

    def fetch_data(self, page: int) -> str:
        """Fetch data from a specific page of the VRDB website."""
        try:
            response = requests.get(f"{self.base_url}?page={page}", timeout=30)
            response.raise_for_status()
            return response.text
        except requests.RequestException as e:
            logger.error(f"Error fetching data from VRDB on page {page}: {str(e)}")
            raise

    def extract_script_content(self, html_content: str) -> str:
        """Extract the relevant script content from HTML"""
        try:
            soup = BeautifulSoup(html_content, 'html.parser')
            scripts = soup.find_all('script')

            for script in scripts:
                if script.string and '__sveltekit_11vwm6r.resolve' in script.string:
                    return script.string

            raise ValueError("Could not find relevant script content")
        except Exception as e:
            logger.error(f"Error extracting script content: {str(e)}")
            raise

    def parse_game_data(self, text: str) -> pd.DataFrame:
        """Regular expressions for capturing required fields"""

        game_pattern = re.compile(
            r'{\s*id:\s*"(?P<id>\d+)",\s*name:\s*"(?P<name>[^"]+)",.*?'
            r'genres:\s*(?P<genres>\[.*?\]),\s*rating_score:\s*(?P<rating_score>\d+(\.\d+)?),.*?'
            r'release_date:\s*"(?P<release_date>[^"]*)",.*?developer:\s*"(?P<developer>[^"]*)",.*?'
            r'publisher:\s*"(?P<publisher>[^"]*)",.*?platforms:\s*(?P<platforms>\[.*?\]),.*?'
            r'store_link:\s*"(?P<store_link>[^"]*)".*?}'
            , re.DOTALL
        )

        # Extract games using regex
        games = []
        for match in game_pattern.finditer(text):
            game = {
                "id": match.group("id"),
                "name": match.group("name"),
                "genres": json.loads(match.group("genres")),
                "rating_score": float(match.group("rating_score")),
                "release_date": match.group("release_date"),
                "developer": match.group("developer"),
                "publisher": match.group("publisher"),
                "platforms": json.loads(match.group("platforms")),
                "store_link": match.group("store_link")
            }
            games.append(game)

        return pd.DataFrame(games)

    def save_to_excel(self, df: pd.DataFrame, output_file: str) -> None:
        """Save DataFrame to Excel, either creating a new file or appending to an existing one."""
        try:

            with pd.ExcelWriter(output_file, mode='w', engine='openpyxl') as writer:
                df.to_excel(writer, index=False, header=True, sheet_name='VR_Games')

        except Exception as e:
            logger.error(f"Error saving data to Excel: {str(e)}")
            raise

    def run(self) -> str:
        """Main method to run the entire extraction process with pagination and Excel file handling."""
        page = 1
        original_df = pd.DataFrame()
        output_file = "VR_Games_Data.xlsx"
        first_write = True
        while True:
            try:

                logger.info(f"Page : {page} Data {'written' if first_write else 'appended'} to {output_file} ")
                html_content = self.fetch_data(page)
                script_content = self.extract_script_content(html_content)
                games_df = self.parse_game_data(script_content)
                logger.info(f"Page : {page} Number of data extracted - {len(games_df)}")
                if games_df.empty:
                    logger.info(f"No more data found on page {page}. Ending extraction.")
                    break
                original_df = pd.concat([original_df, games_df], ignore_index=True)
                first_write = False
                page += 1
            except ValueError:
                # Stop if there are no more relevant pages
                logger.info("No more relevant data found. Stopping extraction.")
                break
            except Exception as e:
                logger.error(f"Error in extraction process on page {page}: {str(e)}")
                raise

        self.save_to_excel(original_df, output_file)
        return output_file


if __name__ == "__main__":
    try:
        extractor = VRDBExtractor()
        output_file = extractor.run()
        print(f"Successfully extracted VR games data to: {output_file}")
    except Exception as e:
        print(f"Error: {str(e)}")
