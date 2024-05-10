import xlwings as xw
import requests
from concurrent.futures import ThreadPoolExecutor
import logging

# Setup logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# Constants
REPLACEMENTS = {
    "Godan verb with 'u' ending": 'Godan う',
    "I-adjective (keiyoushi)": 'い-adjective',
    "Transitive verb": 'Transitive',
    "Intransitive verb": 'Intransitive',
    "Na-adjective (keiyodoshi)": 'な-adjective',
    "Godan verb with 'ru' ending": 'Godan る',
    "Noun which may take the genitive case particle 'no'": '+の',
    "Adverb (fukushi)": 'Adverb',
    "Suru verb": '+する',
    "jlpt-n5": 'N5',
    "jlpt-n4": 'N4',
    "jlpt-n3": 'N3',
    "jlpt-n2": 'N2',
    "jlpt-n1": 'N1'
}
EXCLUDED_PARTS_OF_SPEECH = {'Wikipedia definition', 'place', 'person', 'organization'}

EXCEL_PATH = "C:\\Users\\dkbar\\OneDrive\\Documents\\Spreadsheets\\TangoVocabSheet.xlsm"
SHEET_NAME = 'Raw'

def fetch_data(vocab):
    """Fetch data from Jisho API."""
    url = f"https://jisho.org/api/v1/search/words?keyword={vocab}"
    try:
        response = requests.get(url)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"Failed to fetch data for {vocab}: {e}")
        return None

def format_parts_of_speech(pos_list):
    """Replace parts of speech based on predefined patterns, excluding certain parts."""
    # Filter out excluded parts of speech and replace as necessary
    return [REPLACEMENTS.get(pos, pos) for pos in pos_list if pos not in EXCLUDED_PARTS_OF_SPEECH]


def format_jlpt(jlpt_levels, is_common):
    """Format the JLPT level or commonality."""
    if jlpt_levels:
        # Replace JLPT codes with their corresponding values
        return ', '.join(REPLACEMENTS.get(level, level) for level in sorted(jlpt_levels))
    return 'common' if is_common else 'uncommon'

def process_data(vocab, session):
    """Process vocabulary data fetched from Jisho API, excluding certain parts of speech."""
    data = fetch_data(vocab)
    if not data or not data.get('data'):
        return None

    item = data['data'][0]
    
    # Filter senses to exclude certain parts of speech
    filtered_senses = [s for s in item['senses'] if not any(pos in EXCLUDED_PARTS_OF_SPEECH for pos in s['parts_of_speech'])]

    processed_data = {
        'Vocab': vocab,
        'Hiragana': item['japanese'][0].get('reading', ''),
        'Meaning': "; ".join(sum([s['english_definitions'] for s in filtered_senses], [])),
        'Part of Speech': ", ".join(sorted(set(sum([format_parts_of_speech(s['parts_of_speech']) for s in filtered_senses], [])))),
        'JLPT': format_jlpt(item.get('jlpt'), item.get('is_common', False)),
        'Other Forms': '  '.join(j['word'] for j in item['japanese'][1:] if 'word' in j)
    }
    return processed_data


def update_spreadsheet():
    """Update the active spreadsheet with processed data using concurrent API calls."""
    app = xw.apps.active  # Get the currently active Excel application
    wb = app.books.active  # Get the active workbook
    sht = wb.sheets[SHEET_NAME]
    vocabs = [sht.range(f'A{i}').value for i in range(2, sht.range('A' + str(sht.cells.last_cell.row)).end('up').row + 1) if sht.range(f'A{i}').value and not sht.range(f'B{i}').value]

    with requests.Session() as session:
        with ThreadPoolExecutor(max_workers=10) as executor:
            results = list(executor.map(lambda vocab: process_data(vocab, session), vocabs))

        for idx, result in enumerate(results, start=2):
            if result:
                sht.range(f'B{idx}').value = result['Hiragana']
                sht.range(f'C{idx}').value = result['Meaning']
                sht.range(f'H{idx}').value = result['Part of Speech']
                sht.range(f'I{idx}').value = result['JLPT']
                sht.range(f'K{idx}').value = result['Other Forms']

    # Do not save the workbook here if you want it to stay open; user can save manually
    logging.info("Workbook updated successfully.")

# Call the function directly without running as main if triggered from a Ribbon button
update_spreadsheet()







