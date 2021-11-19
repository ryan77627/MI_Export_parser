try:
    import openpyxl
except ImportError:
    print("Please install the openpyxl module")
    quit()

try:
    from rapidfuzz import fuzz
except ImportError:
    print("Please install the fuzzywuzzy module")
    quit()

import multiprocessing as mp
# mp.set_start_method('spawn')
import csv
from time import sleep


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


def import_items():  # TODO: Parse items to class ONCE, then sort into new lists based on parsed info
    file = input("Please enter file path to MI_EXP file: ")
    igCategories = []
    counter = 0
    with open(file) as f:
        lines = f.readlines()
        for category in range(0, 50):
            catItems = []
            for i in range(0, len(lines)):
                # Split, basic parsing, store into object
                # Standardize line because of brackets
                entry = ""
                for char in lines[i]:
                    if char == '{':
                        entry += '"'
                        entry += char
                    elif char == '}':
                        entry += char
                        entry += '"'
                    else:
                        entry += char
                entry = [ "{}".format(x) for x in next(csv.reader([entry], delimiter=',', quotechar='"')) ]
                if int(entry[8]) == category:
                    object = item(entry[2].strip("\""), i)
                    catItems.append(object)
                    counter += 1
                    print(f"Parsed {counter} items from export file...", end='\r')
                    # print(f"added {object.name} for revenue class {category}!")
            if len(catItems) > 0:
                # we had some items for this rev group, add to todo
                igCategories.append(catItems)

    # We should have a list with all items in sublists
    print()
    return igCategories


def sort_items(todo, counter, results):
    # Agh. Suuuuper inefficient
    sorted_list = []
    unsorted = todo.get()
    allItems = unsorted.copy()
    for item in unsorted:
        if item.inserted != 1:
            # Need to add this and similar items to the sorted list
            sorted_list.append(item)
            item.inserted = 1
            counter.put(1)
            for item1 in allItems:
                # This is about o(N^2). Yikes.
                if item1.inserted != 1:
                    # Not in the list, is it similar?
                    if fuzz.token_sort_ratio(item.name, item1.name) >= 80:
                        sorted_list.append(item1)
                        item1.inserted = 1
                        counter.put(1)

    # we should be done with this block now
    results.put(sorted_list)
    counter.put(None)
    return


def init_sort(igCategories):
    # This actually spawns the sort threads (processes in this case)
    sorted_list = []
    processPool = mp.Pool()
    poolManager = mp.Manager()
    todo = poolManager.Queue()
    sorted_results = poolManager.Queue()
    counter_queue = poolManager.Queue()
    counter = 0
    workers = len(igCategories)

    # Add tasks to todo (all lists to be sorted)
    for items in igCategories:
        todo.put(items)

    items = poolManager.list()
    for i in igCategories:  # Make proper var names
        for j in i:
            items.append(j)

    for i in igCategories:
        res = processPool.apply_async(sort_items, (todo, counter_queue, sorted_results))

    while workers > 0:  # Change this for number of results expected
        status = counter_queue.get()
        if status == None:
            workers -= 1
        else:
            counter += status  # this increments the counter for every item sorted
            progress = "{:.2f}".format((counter / len(items)) * 100)
            print(f"Current Progress: {progress}%", end='\r')
    print()

    processPool.close()
    processPool.join()

    with open("output.txt", "w") as out:
        # We should have results
        while not sorted_results.empty():
            result = sorted_results.get()
            for item in result:
                out.write(item.name + '\n')

        out.close()



def main():
    igCategories = import_items()
    init_sort(igCategories)
    # create_spreadsheet()


if __name__ == "__main__":
    main()
