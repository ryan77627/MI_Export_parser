try:
    import openpyxl
except ImportError:
    print("Please install the openpyxl module")
    quit()

try:
    from fuzzywuzzy import fuzz
except ImportError:
    print("Please install the fuzzywuzzy module")
    quit()

import queue
import multiprocessing

items = []  # MOVE THIS

class item:
    # Internal vars to represent an item.
    # Name: Name of an item
    # ID: InfoGenesis Item ID
    # Price: InfoGenesis Price
    # Inserted: Internal var that gets set when placed in the spreadsheet

    def __init__(self, name, line):
        self.name = name
        self.inserted = 0
        self.line = line


def import_items():
    file = input("Please enter file path to MI_EXP file: ")
    with open(file) as f:
        lines = f.readlines()
        for i in range(0, len(lines)):
            # Split, basic parsing, store into object
            entry = lines[i]
            entry = entry.split(",")
            object = item(entry[2].strip("\""), i)
            items.append(object)
            print(f"added {object.name}!")

    # We should have a list with all items created here


def sort_items():
    # Agh. Suuuuper inefficient
    sorted_list = []
    counter = 0
    for item in items:
        progress = "{:.2f}".format((counter / len(items)) * 100)
        print(f"Current Progress: {progress}%", end='\r')
        if item.inserted != 1:
            # Need to add this and similar items to the sorted list
            sorted_list.append(item)
            item.inserted = 1
            counter += 1
            for item1 in items:
                # This is about o(N^2). Yikes.
                if item1.inserted != 1:
                    # Not in the list, is it similar?
                    if fuzz.token_sort_ratio(item.name, item1.name) >= 70:
                        sorted_list.append(item1)
                        item1.inserted = 1
                        counter += 1

    # For now, write names to list to check how we're doing
    out = open("output.txt", "w")
    for item in sorted_list:
        out.write(item.name + '\n')

    out.close()


def main():
    import_items()
    sort_items()
    # create_spreadsheet()


if __name__ == "__main__":
    main()
