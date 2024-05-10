import json
import xlwings as xw
import subprocess
import logging
import sys

# Configure logging
logging.basicConfig(stream=sys.stdout, level=logging.DEBUG,
                    format='%(asctime)s:%(levelname)s:%(message)s')

# Global variables
EXCEL_PATH = r"C:\Users\dkbar\OneDrive\Documents\Spreadsheets\TangoVocabSheet.xlsm"
SHEET_NAME = 'Raw'
VERB_CRITERIA = ['Godan う', 'Godan る', 'Ichidan verb']
ADJ_CRITERIA = ['い-adjective']
VERB_JSON_PATH = r"C:\Projects\Kamiya\input\input_verb.json"
ADJ_JSON_PATH = r"C:\Projects\Kamiya\input\input_adj.json"
NODE_VERB_SCRIPT_PATH = r"C:\Projects\Kamiya\scripts\kamiyacodec_verbs.js"
NODE_ADJ_SCRIPT_PATH = r"C:\Projects\Kamiya\scripts\kamiyacodec_adj.js"

def read_and_filter_excel():
    try:
        app = xw.App(visible=False)
        wb = app.books.open(EXCEL_PATH)
        sheet = wb.sheets[SHEET_NAME]
        last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row

        verbs = []
        adjectives = []

        for i in range(2, last_row + 1):
            word = sheet.range(f'A{i}').value
            part_of_speech = sheet.range(f'H{i}').value or ""
            simplified_pos = part_of_speech.split(",")[0]
            if any(criterion in simplified_pos for criterion in VERB_CRITERIA):
                verbs.append(word)
            elif any(criterion in simplified_pos for criterion in ADJ_CRITERIA):
                adjectives.append(word)

        
        logging.info(f"Extracted {len(verbs)} verbs and {len(adjectives)} adjectives from Excel.")
        wb.close()
        app.quit()
        return verbs, adjectives

    except Exception as e:
        logging.error("Error reading and filtering Excel file: " + str(e))
        raise

def write_json(data, file_path):
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
            f.flush()  # Force flushing the I/O buffer
        logging.info(f"Data written to {file_path} successfully")
    except Exception as e:
        logging.error(f"Error writing data to JSON: {file_path}, error: {str(e)}")
        raise

def trigger_node_script(input_file, script_path):
    try:
        logging.info(f"Attempting to run Node script: {script_path} with input file: {input_file}")
        result = subprocess.run(['node', script_path, input_file], check=True, capture_output=True, text=True, encoding='utf-8')
        logging.info(f"Node script {script_path} executed with result: {result.stdout}")
        if result.stdout:
            logging.info(f"Node script {script_path} output:\n{result.stdout}")
        if result.stderr:
            logging.warning(f"Node script {script_path} error output:\n{result.stderr}")
    except subprocess.CalledProcessError as e:
        logging.error(f"Node script execution failed: {e.stderr}")
        raise
    except Exception as e:
        logging.error(f"Unexpected error during Node script execution: {str(e)}")
        raise

def main():
    try:
        verbs, adjectives = read_and_filter_excel()
        write_json(verbs, VERB_JSON_PATH)
        write_json(adjectives, ADJ_JSON_PATH)

        trigger_node_script(VERB_JSON_PATH, NODE_VERB_SCRIPT_PATH)
        trigger_node_script(ADJ_JSON_PATH, NODE_ADJ_SCRIPT_PATH)

    except Exception as e:
        logging.error("Failed to complete main function: " + str(e))

if __name__ == "__main__":
    main()
