import xlwings as xw
import os
import spacy
import json
import logging

logging.basicConfig(level=logging.DEBUG, format="%(asctime)s [%(levelname)s] %(message)s")
EXCEL_PATH = "C:\\Users\\dkbar\\OneDrive\\Documents\\Spreadsheets\\TangoVocabSheet.xlsm"
SHEET_NAME = 'Raw'  # The name of the worksheet we will be working on.

try:
    nlp = spacy.load("ja_ginza")
except Exception as e:
    logging.error(f"Failed to load spaCy language model: {e}", exc_info=True)
    logging.exception("Unhandled exception occurred")
    raise

def get_conjugations(words):
    # Load the JSON files generated by the Node.js scripts
    with open('C:\\Projects\\Kamiya\\output\\output_verb.json', 'r', encoding='utf-8') as f:
        verb_conjugations = json.load(f)
    with open('C:\\Projects\\Kamiya\\output\\output_adj.json', 'r', encoding='utf-8') as f:
        adj_conjugations = json.load(f)


def update_excel():
    try:
        app = xw.App(visible=False)
        wb = app.books.open(EXCEL_PATH)
        sheet = wb.sheets[SHEET_NAME]
        last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row

        logging.info(f"Starting to process rows 2 to {last_row}")

        for i in range(2, last_row + 1):
            words = sheet.range(f'A{i}').value
            part_of_speech = sheet.range(f'H{i}').value or ""
            skip_row = False  # Define 'skip_row' before using it in the loop

            for col in ['D', 'F']:
                sentence = sheet.range(f'{col}{i}').value

                if sentence and '<strong>' in sentence:
                    logging.info(f"Row {i} already processed, skipping...")
                    skip_row = True
                    break

            if skip_row:
                continue

            if part_of_speech not in ['Godan う', 'い-adjective', 'Godan る', 'Ichidan verb']:
                for col in ['D', 'F']:
                    sentence = sheet.range(f'{col}{i}').value
                    if sentence:
                        doc = nlp(sentence)
                        for ent in doc:
                            if ent.text.strip() in words.split(','):
                                sentence = sentence.replace(ent.text, f'<strong>{ent.text}</strong>')
                        logging.debug(f"Setting value in Excel at column {col}, row {i}: {sentence}")
                        sheet.range(f'{col}{i}').value = sentence
                        logging.debug(f"Row {i}, Column A value: {words}")
                        logging.info(f"Updated non-conjugated word in row {i}, column {col}")
            else:
                
        # Save changes to the workbook
        wb.save()
        logging.info("Excel workbook saved successfully.")

    except Exception as e:
        logging.exception(f"An error occurred while updating Excel: {e}")

if __name__ == "__main__":
    try:
        logging.info("Script started.")
        update_excel()
        logging.info("Script finished successfully.")
    except Exception as e:
        logging.error(f"Script did not complete due to an error: {e}", exc_info=True)
        logging.exception("Unhandled exception occurred")