"""
Author: Ethan Zornosa
Email: ethansito@tutanota.com
Date Created: 6/10/2025
Last Updated: 6/22/2025
Note: This program is dependent on the CIS csv format in place as of the date of this program. If the format of that
file is changed, this program could become broken.
"""

import csv
import sys
import statistics
import tkinter as tk
from tkinter import *
from tkinter import filedialog
import pickle
import webbrowser
import os

def smooth(column_num, z_thresh, num_samples):
    global rows
    global title_row
    global last_touched
    global final_rows
    print("Title Row: " + str(title_row))
    print("Length: " + str(len(rows)))

    # For each data point, gather 5 leading and 5 trailing data points
    for j in range(title_row, len(rows) - 1):
        line = rows[j]
        if line[column_num] == "":
            continue
        point_of_interest = float(line[column_num])

        target = j
        last_touched = target
        # Gathering points before point of interest
        first_search = point_search(num_samples, 0, target, column_num)
        second_search = point_search(num_samples + abs(len(first_search) - num_samples), 1, target, column_num)
        third_search = point_search(num_samples - len(second_search), 0, last_touched, column_num)

        collection = first_search + second_search + third_search
        mean = statistics.mean(collection)
        standard_deviation = statistics.stdev(collection, mean)
        median = statistics.median(collection)
        point = float(line[column_num])
        # Avoid zero division error
        if standard_deviation == 0:
            continue
        z_score = (point - mean) / standard_deviation

        # Remedy Spikes
        if abs(z_score) >= abs(z_thresh):
            final_rows[j][column_num] = str(median)
            if median == 0:
                print("MEDIAN IS 0!")

        # Set up Future Runs by overwriting original rows
        rows.clear()
        for thingy in final_rows:
            rows.append(thingy.copy())

# Search for specified number of data points in specified direction. Return back data points found.
def point_search(points_required, direction, target, column_num):
    global rows
    global title_row
    global last_touched
    points = []
    points_found = 0
    while points_found < points_required:
        if direction == 0:
            target -= 1
        elif direction == 1:
            target += 1
        if target >= len(rows) or target <= title_row:
            print("Unable to find required points.")
            break
        potential_line = rows[target]
        potential_point = potential_line[column_num]
        if potential_point == "":
            continue
        points_found += 1
        points.append(float(potential_point))
        last_touched = target
    return points

def file_explore(entry: tk.Entry):
    global files_to_explore
    files_to_explore = filedialog.askopenfilenames(title="Open CIS File", filetypes=(("CIS (.csv)", "*.csv"), ("All Files", "*.*")))
    if len(files_to_explore) > 0:
        entry.delete(0, 'end')
        number = len(files_to_explore)
        for i in range(0, number):
            if i < number - 1:
                entry.insert('end', files_to_explore[i] + ", ")
            else:
                entry.insert('end', files_to_explore[i])


def folder_explore(entry: tk.Entry):
    folder = filedialog.askdirectory(title="Open Directory")
    if len(folder) > 0:
        entry.delete(0, 'end')
        entry.insert(0, folder)

def operate(z_thresh, num_runs, num_samples, input_entry: tk.Entry, output_entry: tk.Entry, root, runs_entry: tk.Entry, samples_entry: tk.Entry):
    global last_touched
    global rows
    global title_row
    global final_rows
    global files_to_explore

    try:
        z_thresh = float(z_thresh)
    except ValueError or TypeError:
        error_box("Z Score Must Be Integer Or Decimal.", root)
        # Reset the globals
        rows = []
        last_touched = 0
        title_row = 0
        final_rows = []
        return
    try:
        num_runs = int(num_runs)
    except ValueError or TypeError:
        error_box("Number Of Runs Must Be Integer.", root)
        # Reset the globals
        rows = []
        last_touched = 0
        title_row = 0
        final_rows = []
        return
    try:
        num_samples = int(num_samples)
    except ValueError or TypeError:
        error_box("Number Of Samples Must Be Integer.", root)
        # Reset the globals
        rows = []
        last_touched = 0
        title_row = 0
        final_rows = []
        return

    if num_runs > 2 or num_runs <= 0:
        runs_entry.delete(0, "end")
        runs_entry.insert(0, "2")
        num_runs = 2

    if num_samples < 1:
        samples_entry.delete(0, "end")
        samples_entry.insert(0, "1")
        num_samples = 1

    if len(files_to_explore) == 0:
        files_to_explore = input_entry.get().split(", ")
    for file in files_to_explore:
        # Setup
        path = file
        fields = []

        # read the file
        try:
            with open(path, 'r') as csvfile:
                csvreader = csv.reader(csvfile)
                for row in csvreader:
                    rows.append(row)
        except OSError:
            error_box("Invalid File Name:\n" + path, root)
            return

        # Find row with titles
        found = False
        for row in rows:
            if not found:
                title_row += 1
            for i in range(len(row)):
                thing = row[i]
                if thing.lower() == "dc on":
                    print("Located Title Row")
                    fields = row
                    found = True
        if not found:
            error_box("DC ON Column Does Not Exist For File: \n " + str(file), root)

        # Make a deep copy of the data. Necessary as we want to reference original points while doing calculations.
        for thing in rows:
            final_rows.append(thing.copy())

        # For each point, find mean and standard deviation of surrounding points to generate Z score.
        for i in range(0, num_runs):
            dc_on_column = fields.index("DC ON")
            smooth(dc_on_column, z_thresh, num_samples)
            dc_off_column = fields.index("DC OFF")
            smooth(dc_off_column, z_thresh, num_samples)
            try:
                ac_on_column = fields.index("AC ON")
                smooth(ac_on_column, z_thresh, num_samples)
                ac_off_column = fields.index("AC OFF")
                smooth(ac_off_column, z_thresh, num_samples)
            except ValueError:
                pass

        # Slice the input to get the file name
        slash_found = False
        i = 1
        while not slash_found:
            # print(input_entry.get()[-i])
            if file[-i] == "/" or file[-i] == "\\":
                slash_found = True
            else:
                i += 1
        file_name = file[-i:]
        # Output to a new csv file
        with open(output_entry.get() + "/" + file_name[0:-4] + "-Fixed.csv", 'w', newline='') as csvfile:
            csvwriter = csv.writer(csvfile)
            csvwriter.writerows(final_rows)
        error_box("Task successfully failed!", root)

        # Reset the globals
        rows = []
        last_touched = 0
        title_row = 0
        final_rows = []

    # Pickle the paths
    database_inputs = []
    for file in files_to_explore:
        database_inputs.append(file)
    database_output = output_entry.get()
    with open("database.pickle", "wb") as db_file:
        pickle.dump(database_inputs, db_file)
        pickle.dump(database_output, db_file)

def error_box(message, root):
    global rows
    global last_touched
    global title_row
    global final_rows
    global files_to_explore
    rows = []
    last_touched = 0
    title_row = 0
    final_rows = []
    files_to_explore = ()

    window = Toplevel(root)
    window.geometry("400x100+440+260")
    error_label = tk.Label(window, text=message, font=("Arial", 16))
    error_label.pack()
    error_button = tk.Button(window, text="I'm a failure.", command= lambda: window.destroy(), font=("Arial", 16))
    error_button.pack()

def mouse_click(event):
    webbrowser.open("https://www.linkedin.com/in/ethan-zornosa/")

def mouse_enter(event, sig_label: tk.Label):
    sig_label.configure(fg="Blue")

def mouse_leave(event, sig_label: tk.Label):
    sig_label.configure(fg="Black")

def edit_input_entry(event):
    global files_to_explore
    files_to_explore = ()

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # set by PyInstaller at runtime
    except AttributeError:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def main():
    root = Tk()
    try:
        photo = PhotoImage(file="Spike-B-Gone.png")
    except TclError:
        photo = PhotoImage(file=resource_path("Spike-B-Gone.png"))
    root.iconphoto(True, photo)
    root.geometry("480x360+400+200")
    root.title("Spike-B-Gone")
    label = tk.Label(root, text="Spike-B-Gone", font=("Comic Sans MS", 24))
    label.pack()
    image = photo.subsample(16, 16)
    photo_label = tk.Label(root, image=image)
    photo_label.place(x=60, y=10)
    description_label = tk.Label(root, text="CIS Data Processor", font=("Arial", 14))
    description_label.pack()
    input_label = tk.Label(root, text="Input")
    input_label.pack()
    input_frame = tk.Frame(root)
    input_frame.pack()
    input_entry = tk.Entry(input_frame, width=50)
    input_entry.bind("<Key>", edit_input_entry)
    input_entry.grid(row=0, column=0)
    input_button = tk.Button(input_frame, text="Browse", command= lambda: file_explore(input_entry))
    input_button.grid(row=0, column=1)
    output_label = tk.Label(root, text="Output")
    output_label.pack()
    output_frame = tk.Frame(root)
    output_frame.pack()
    output_entry = tk.Entry(output_frame, width=50)
    output_entry.grid(row=0, column=0)
    output_button = tk.Button(output_frame, text="Browse", command= lambda: folder_explore(output_entry))
    output_button.grid(row=0, column=1)
    settings_frame = tk.Frame(root)
    settings_frame.pack(pady=25)
    z_score_entry = tk.Entry(settings_frame, justify="center", width=3)
    z_score_entry.insert(0, "1")
    z_score_entry.grid(row=0, column=0)
    z_score_label = tk.Label(settings_frame, text="Z Score Threshold")
    z_score_label.grid(row=0, column=1, padx=(0, 10))
    runs_entry = tk.Entry(settings_frame, justify="center", width=3)
    runs_entry.grid(row=0, column=2)
    runs_entry.insert(0, "1")
    runs_label = tk.Label(settings_frame, text="Runs")
    runs_label.grid(row=0, column=3, padx=(0, 10))
    points_entry = tk.Entry(settings_frame, justify="center", width=3)
    points_entry.grid(row=0, column=4)
    points_entry.insert(0, "15")
    points_label = tk.Label(settings_frame, text="Num. Samples (Both Sides)")
    points_label.grid(row=0, column=5)
    go_button = tk.Button(root, text="GO", bg="red", fg="white", font=("Comic Sans MS", 18), width=200,
                          command= lambda: operate(z_score_entry.get(), runs_entry.get(),
                                                   points_entry.get(), input_entry, output_entry, root, runs_entry, points_entry))
    go_button.pack()
    signature_label = tk.Label(root, text="""Ethan Zornosa
    ethansito@tutanota.com""", font=("Arial", 8))
    signature_label.pack(side="left")
    signature_label.bind("<Button-1>", mouse_click)
    signature_label.bind("<Enter>", lambda event: mouse_enter(event, sig_label=signature_label))
    signature_label.bind("<Leave>", lambda event: mouse_leave(event, sig_label=signature_label))
    # If they exist, unpickle input and output paths
    try:
        with open("database.pickle", "rb") as db_file:
            try:
                database_input = pickle.load(db_file)
                database_output = pickle.load(db_file)
                input_entry.delete(0, 'end')
                number = len(database_input)
                for i in range(0, number):
                    if i < number - 1:
                        input_entry.insert('end', database_input[i] + ", ")
                    else:
                        input_entry.insert('end', database_input[i])
                output_entry.delete(0, "end")
                output_entry.insert(0, database_output)
            except EOFError:
                pass
    except FileNotFoundError:
        pass
    root.mainloop()

if __name__ == "__main__":
    rows = []
    last_touched = 0
    title_row = 0
    final_rows = []
    files_to_explore = ()
    main()
