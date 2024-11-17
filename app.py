from src.games_list_extractor import VRDBExtractor

if __name__ == "__main__":
    try:
        extractor = VRDBExtractor()
        output_file = extractor.run()
        print(f"Successfully extracted VR games data to: {output_file}")
    except Exception as e:
        print(f"Error: {str(e)}")