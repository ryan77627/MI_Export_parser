try:
    import openpyxl
except ImportError:
    print("Please install the openpyxl module")
    quit()

try:
    import fuzzywuzzy
except ImportError:
    print("Please install the fuzzywuzzy module")
    quit()


class item:
    # Internal vars to represent an item.
    # Name: Name of an item
    # ID: InfoGenesis Item ID
    # Price: InfoGenesis Price
    # Inserted: Internal var that gets set when placed in the spreadsheet

    def __init__(self, name, ID, price):
        self.name = name
        self.inserted = 0
        self.ID = ID
        self.price = price


def main():
    import_items()
    sort_items()
    create_spreadsheet()
