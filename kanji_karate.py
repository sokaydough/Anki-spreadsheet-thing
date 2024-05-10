import xlwings as xw
import re

EXCEL_PATH = "C:\\Users\\dkbar\\OneDrive\\Documents\\Spreadsheets\\TangoVocabSheet.xlsm"
SHEET_NAME = 'Raw'
def extract_kanji(text):
    """Extracts individual Kanji characters from the text."""
    return re.findall(r'[\u4e00-\u9faf]', text)

def update_kanji_details():
    app = xw.apps.active  # Get the currently active Excel application
    wb = app.books.active 
    sht = wb.sheets[SHEET_NAME]  # Target sheet

    # Define the range for the 'Vocab' column (e.g., column A)
    vocab_range = sht.range('A2:A' + str(sht.range('A' + str(sht.cells.last_cell.row)).end('up').row))
    
    # Iterate over each cell in the range
    for cell in vocab_range:
        kanji_list = extract_kanji(cell.value)
        # Write extracted kanji to designated columns
        if len(kanji_list) > 0:
            cell.offset(0, 11).value = kanji_list[0]  # Assuming 'Kanji A' is at column L
        if len(kanji_list) > 1:
            cell.offset(0, 15).value = kanji_list[1]  # Assuming 'Kanji B' is at column P
        if len(kanji_list) > 2:
            cell.offset(0, 19).value = kanji_list[2]  # Assuming 'Kanji C' is at column T

    wb.save()  # Save changes

if __name__ == "__main__":
    update_kanji_details()
