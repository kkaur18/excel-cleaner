import tkinter as tk
from tkinter import filedialog
from os import write
import pandas as pd
import numpy as np
import csv
from scipy import stats

def browse_file(entry_widget):
    file_path = filedialog.askopenfilename()
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, file_path)

def clean_files(file1_path, file2_path):
    print(f"Input File: {file1_path}")
    print(f"Output File: {file2_path}")
    df = pd.read_csv(file1_path)
    deleted = ["ALPHABRODER", "BULLET LINE LLC", "CIC - ALPHARETTA", "D. PEYSER/M V SPORT", "DAVID PEYSER - ALABAMA", "ECOMPANYSTORE.COM", "ECOMPANYSTORE.COM-OFFICE", "HUETONE IMPRINTS,INC.", "LC MARKETING", "PRIME RESOURCES OUTBOUND"]

    df_filtered = df.drop(df[df["Shipper Name"].isin(deleted)].index)
    for index,row in df_filtered.iterrows():
        row["Reference Number(s)"] = str(row["Reference Number(s)"])
        row["Reference Number(s)"] = row["Reference Number(s)"].replace("PONUM","")
        row["Reference Number(s)"] = row["Reference Number(s)"].replace("PO#","")
        numbers = row["Reference Number(s)"].split("|")
        row["Reference Number(s)"] = None
        for j in numbers:
            if j[0] == "4" and len(j) == 6:
                row["Reference Number(s)"] = j
                break
    df_filtered.dropna(subset=["Reference Number(s)"], inplace=True)
    with open(file2_path, "w",newline='') as file:
        writer = csv.writer(file)
        writer.writerow(["PONumber","DateShipped","TrackingNumber"])
        for index,row in df_filtered.iterrows():
            writer.writerow([row["Reference Number(s)"], row["Tracking Number"], row["Manifest Date"]])
    file.close()


def on_submit(file1_entry, file2_entry):
    file1_path = file1_entry.get()
    file2_path = file2_entry.get()

    if file1_path and file2_path:
        clean_files(file1_path, file2_path)
        result_label.config(text="Files cleaned successfully!")
    else:
        result_label.config(text="Please select input AND output files.")

# GUI setup
root = tk.Tk()
root.title("Excel Cleaning Tool")

# File 1 input
file1_label = tk.Label(root, text="Input File:")
file1_label.grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)

file1_entry = tk.Entry(root, width=40)
file1_entry.grid(row=0, column=1, padx=10, pady=5)

file1_button = tk.Button(root, text="Browse", command=lambda: browse_file(file1_entry))
file1_button.grid(row=0, column=2, padx=5, pady=5)

# File 2 input
file2_label = tk.Label(root, text="Output File:")
file2_label.grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)

file2_entry = tk.Entry(root, width=40)
file2_entry.grid(row=1, column=1, padx=10, pady=5)

file2_button = tk.Button(root, text="Browse", command=lambda: browse_file(file2_entry))
file2_button.grid(row=1, column=2, padx=5, pady=5)

# Submit button
submit_button = tk.Button(root, text="Submit", command=lambda: on_submit(file1_entry, file2_entry))
submit_button.grid(row=2, column=0, columnspan=3, pady=10)

# Result label
result_label = tk.Label(root, text="")
result_label.grid(row=3, column=0, columnspan=3)

root.mainloop()