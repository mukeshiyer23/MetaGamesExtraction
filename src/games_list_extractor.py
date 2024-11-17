import json
from typing import Dict, Any, List

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
                if script.string and 'return {id:1' in script.string:
                    return script.string

            raise ValueError("Could not find relevant script content")
        except Exception as e:
            logger.error(f"Error extracting script content: {str(e)}")
            raise

    def parse_game_data(self, text: str) -> pd.DataFrame:
        """
        Parse game data text into a pandas DataFrame with mandatory and optional fields.
        """
        # Mandatory fields pattern
        mandatory_pattern = re.compile(
            r'{\s*'
            r'id:\s*"(?P<id>\d+)".*?'
            r'name:\s*"(?P<name>[^"]+)".*?'
            r'genres:\s*(?P<genres>\[.*?\])'
            r'.*?store_link:\s*"(?P<store_link>[^"]*)"'
            r'.*?}',
            re.DOTALL
        )

        # Optional fields patterns
        optional_patterns = {
            'developer': r'developer:\s*"(?P<developer>[^"]*)"',
            'publisher': r'publisher:\s*"(?P<publisher>[^"]*)"',
            'platforms': r'platforms:\s*(?P<platforms>\[.*?\])',
            'release_date': r'release_date:\s*"(?P<release_date>[^"]*)"',
            'rating_score': r'rating_score:\s*(?P<rating_score>\d+(?:\.\d+)?)',
            'rating_count': r'rating_count:\s*(?P<rating_count>\d+)',
            'game_mode': r'game_mode:\s*"(?P<game_mode>[^"]*)"',
            'languages': r'languages:\s*(?P<languages>\[.*?\])',
            'age_rating': r'age_rating:\s*"(?P<age_rating>[^"]*)"',
            'space_required': r'space_required:\s*"(?P<space_required>[^"]*)"',
            'price_USD_amount': r'price_USD_amount:\s*(?P<price_USD_amount>\d+)',
            'price_USD_formatted': r'price_USD_formatted:\s*"(?P<price_USD_formatted>[^"]*)"'
        }

        def parse_array_field(field_value: str) -> List[str]:
            """Parse array fields from string to list"""
            try:
                return json.loads(field_value)
            except json.JSONDecodeError:
                return []

        def extract_optional_fields(game_text: str) -> Dict[str, Any]:
            """Extract optional fields from game text"""
            optional_fields = {}

            for field, pattern in optional_patterns.items():
                match = re.search(pattern, game_text)
                if match:
                    value = match.group(field)
                    # Handle array fields
                    if field in ['platforms', 'languages']:
                        optional_fields[field] = parse_array_field(value)
                    # Handle numeric fields
                    elif field in ['rating_score', 'rating_count', 'price_USD_amount']:
                        try:
                            optional_fields[field] = float(value) if '.' in value else int(value)
                        except ValueError:
                            continue
                    else:
                        optional_fields[field] = value

            return optional_fields

        games = []
        for match in mandatory_pattern.finditer(text):
            try:
                game_data = {
                    "id": match.group("id"),
                    "name": match.group("name"),
                    "genres": parse_array_field(match.group("genres")),
                    "store_link": match.group("store_link")
                }

                game_text = match.group(0)
                optional_fields = extract_optional_fields(game_text)

                game_data.update(optional_fields)

                games.append(game_data)

            except (AttributeError, json.JSONDecodeError) as e:
                print(f"Error parsing game data: {e}")
                continue
        df = pd.DataFrame(games)

        mandatory_fields = {'id', 'name', 'genres', 'store_link'}
        missing_fields = mandatory_fields - set(df.columns)
        if missing_fields:
            raise ValueError(f"Missing mandatory fields: {missing_fields}")

        return df

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
