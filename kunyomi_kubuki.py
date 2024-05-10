import xlwings as xw
import requests
import logging
import sys
#Set up logging once, using a file handler to log to a file and a stream handler to log to stdout.
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
file_handler = logging.FileHandler('kunyomi_kabuki_log.txt')
stream_handler = logging.StreamHandler(sys.stdout)
formatter = logging.Formatter('%(asctime)s - %(message)s')
file_handler.setFormatter(formatter)
stream_handler.setFormatter(formatter)
logger.addHandler(file_handler)
logger.addHandler(stream_handler)

# Constants should be defined at the top level of your script
EXCEL_PATH = "C:\\Users\\dkbar\\OneDrive\\Documents\\Spreadsheets\\TangoVocabSheet.xlsm"
SHEET_NAME = 'Raw'

def get_kanji_details(kanji):
    """Fetches details for a given kanji from the KanjiAPI."""
    try:
        response = requests.get(f"https://kanjiapi.dev/v1/kanji/{kanji}")
        response.raise_for_status()
        return response.json()
    except requests.RequestException as e:
        logging.error(f"Error fetching details for {kanji}: {e}")
        return None
        
def update_kanji_details(wb, sheet_raw):
    """
    Updates kanji details in an Excel sheet using xlwings.
    This function is intended to be triggered by a ribbon button and doesn't close the workbook after execution.
    """ 
    for column in ['L', 'P', 'T']:
        changes_made = False  # Flag to track if any changes are made to the sheet
        last_row = sheet_raw.range(f"{column}1048576").end('up').row
        
        for i in range(2, last_row + 1):
            cell = sheet_raw.range(f"{column}{i}")
            if cell.value is not None and cell.offset(0, 1).value is None:
                details = get_kanji_details(cell.value)
                if details:
                    cell.offset(0, 1).value = ', '.join(details.get('meanings', [])[:5])
                    # Process kun readings to remove characters after periods and remove duplicates
                    kun_readings = set(reading.split('.')[0] for reading in details.get('kun_readings', []))
                    cell.offset(0, 2).value = ', '.join(sorted(kun_readings))  # Sorted to maintain consistent order
                    cell.offset(0, 3).value = ', '.join(details.get('on_readings', []))
                    changes_made = True

        if changes_made:
            logger.info(f"Kanji details updated in column {column}.")



def main():
    try:
        # Open the workbook; assumes that Excel is already running
        app = xw.apps.active
        wb = app.books.active
        sheet_raw= wb.sheets[SHEET_NAME]  # Directly access the sheet since it must exist as per assumption above
        # Check if "Raw" sheet exists
        if SHEET_NAME in [sheet.name for sheet in wb.sheets]:
            sheet_raw = wb.sheets[SHEET_NAME]
            update_kanji_details(wb, sheet_raw)
            wb.save(EXCEL_PATH)  # Save changes to the workbook
        else:
            print("Sheet 'Raw' does not exist.")
            logging.info("Script completed successfully")
    except Exception as e:
        print(f"An error occurred while updating the kanji details: {e}")
        logging.error(f"An error occurred: {e}")

if __name__ == "__main__":
   main()
