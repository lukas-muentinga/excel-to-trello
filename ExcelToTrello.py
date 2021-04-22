from trello.Trello import Trello
from trello.Boards import create_board
from trello.Cards import create_card
from excelutil.excelutil import table_to_dictlist
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string


######## EXCEL PARAMS
EXCEL_FILE_PATH = 'TestData.xlsx'
EXCEL_SHEET = 'Tabelle1'
TOP_LEFT_CELL = 'A1'
BOTTOM_RIGHT_CELL = 'D5'

######## TRELLO PARAMS (set API-key and board token in trello/config.py)
# Board Parameters
BOARD_NAME = 'Test-Board'
BOARD_DESCRIPTION = 'Example Description'
# Task-Status / Board-Lists

# Map from table headline to trello card fields
CARD_MAP = {
    'name': 'Location', # TODO: implement possibility of combining multiple columns
    'description': 'Description', # TODO: implement possibility to combine multiple columns
    'due_date': 'Due',
    'state': 'Status'
    # TOOD
}
# List of Subtasks if needed
SUBTASKS = []


if __name__ == '__main__':

    # TODO: implement

    # Read in Excel sheet
    start = list(coordinate_from_string(TOP_LEFT_CELL))
    start[0] = column_index_from_string(start[0])
    end = list(coordinate_from_string(BOTTOM_RIGHT_CELL))
    end[0] = column_index_from_string(end[0])
    input = table_to_dictlist(EXCEL_FILE_PATH, EXCEL_SHEET, tuple(start), tuple(end))

    # Open Trello Session
    trello = Trello()

    # Create Board
    board = create_board(trello,
                         name = BOARD_NAME,
                         description = BOARD_DESCRIPTION,
                         use_default_lists=False)

    # Get possible states
    states = sorted(list(set([d[CARD_MAP['state']] for d in input])))
    # Create list for each state
    for state in states:
        
        result_list = board.create_list(state)
        
        # Create Cards on each list
        matching_rows = [card for card in input if card[CARD_MAP['state']] == state]
        for row in matching_rows:
            card = create_card(trello=trello,
                               list_id=result_list.id(),
                               name=row[CARD_MAP['name']],
                               description=row[CARD_MAP['description']],
                               due=row[CARD_MAP['due_date']])
