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

from multiprocessing import Queue
import multiprocessing as mp
#import logging

#logger = mp.log_to_stderr()
#logger.setLevel(mp.SUBDEBUG)
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
            #print(f"added {object.name}!")

    # We should have a list with all items created here


def sort_items(todo, allItems, counter, results):
    # Agh. Suuuuper inefficient
    sorted_list = []
    unsorted = todo.get()
    for item in unsorted:
        #print(f"Eval {item.name} currently...")
        if item.inserted != 1:
            # Need to add this and similar items to the sorted list
            sorted_list.append(item)
            item.inserted = 1
            counter.put(1)
            for item1 in allItems:
                #(f"SUB Eval {item1.name}")
                # This is about o(N^2). Yikes.
                if item1.inserted != 1:
                    # Not in the list, is it similar?
                    if fuzz.token_sort_ratio(item.name, item1.name) >= 70:
                        sorted_list.append(item1)
                        item1.inserted = 1
                        counter.put(1)

    # we should be done with this block now
    results.put(sorted_list)
    counter.put(None)
    return

def init_sort():
    # This actually spawns the sort threads (processes in this case)
    sorted_list = []
    todo = Queue()
    sorted_results = Queue()
    counter_queue = Queue()
    counter = 0
    workers = 1 # make this dynamic

    # Add tasks to todo (all lists to be sorted)
    todo.put(items)

    # Let's start with only one thread for now, expand when we know this works
    p1 = mp.Process(target = sort_items, args=(todo, items, counter_queue, sorted_results))
    p1.start()

    while workers > 0: #Change this for number of results expected
        status = counter_queue.get()
        if status == None:
            workers -= 1
        else:
            counter += status # this increments the counter for every item sorted
            progress = "{:.2f}".format((counter / len(items)) * 100)
            print(f"Current Progress: {progress}%", end='\r')

def main():
    import_items()
    init_sort()
    # create_spreadsheet()


if __name__ == "__main__":
    main()
