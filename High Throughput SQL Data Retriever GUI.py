import pandas as pd
import pymssql
import tkinter as tk
from tkinter import filedialog, messagebox

def main(ID_path,short_path,save_to_path):
    IDs = pd.read_excel(ID_path,header=None).iloc[:,0]
    measurements = pd.read_excel(short_path,header=None).iloc[:,0]

    conn = pymssql.connect(server='XXXXXX', port=00000)

    def concatenate_string(SeriesData,low_index,high_index):
        string = "("
        for item in SeriesData[low_index:high_index]:
            string = string + "'" + item + "',"
        string = string + "'" + SeriesData[high_index] + "')"
        return string

    measurements_string = concatenate_string(measurements,0,len(measurements)-1)

    def retrieve_data(IDs_string,measurements_string):
        sqlCommand = "SELECT sample_ID, measurement, measured_value from table_for_sample_measurements WHERE sample_ID IN "+IDs_string+" AND measurement IN "+measurements_string
        cursor = conn.cursor()
        cursor.execute(sqlCommand)
        data = cursor.fetchall()
        return data

    if len(IDs) <= 5000:
        IDs_string = concatenate_string(IDs,0,len(IDs)-1)
        data = retrieve_data(IDs_string,measurements_string)
    else:
        data = []
        low = 0
        high = 4999
        while high <= len(IDs)-1:
            IDs_string = concatenate_string(IDs,low,high)
            data += retrieve_data(IDs_string,measurements_string)
            low += 5000
            high += 5000
        IDs_string = concatenate_string(IDs,low,len(IDs)-1)
        data += retrieve_data(IDs_string,measurements_string)
        
    conn.close()

    data = pd.DataFrame(data,columns=["Sample ID","Measurement","Measured Value"])
    data = data.pivot(index=['Sample ID'], columns='Measurement', values='Measured Value')
    data = data.reset_index()
    data.to_excel(save_to_path+"/Retrieved data.xlsx", index=False)

def run_code():
    ID_path = ID_path_entry.get()
    short_path = short_path_entry.get()
    save_to_path = save_to_folder.get()
    main(ID_path,short_path,save_to_path)
    messagebox.showinfo("High Throughput SQL Data Retriever", "Data have been retrieved and saved to the destination folder")

root = tk.Tk()
root.title("High Throughput Data Retriever")

empty = tk.Label(root, text="")
empty.grid(row=0, column=1)

empty = tk.Label(root, text="This tool is designed to retrieve large volume of data from the SQL database",font=('Times New Roman', 20, 'bold'))
empty.grid(row=1, column=1)

empty = tk.Label(root, text="")
empty.grid(row=2, column=1)
empty = tk.Label(root, text="")
empty.grid(row=3, column=1)

def open_ID_file():
    file_path = filedialog.askopenfilename()
    ID_path_entry.delete(0, tk.END)
    ID_path_entry.insert(0, file_path)

ID_path_label = tk.Label(root, text="Excel file (xlsx) for sample IDs:", font=('Times New Roman', 15, 'bold'))
ID_path_label.grid(row=4, column=0)
ID_path_entry = tk.Entry(root, width=120)
ID_path_entry.grid(row=4, column=1)
ID_path_button = tk.Button(root, text="Select File", command=open_ID_file)
ID_path_button.grid(row=4, column=2)
ID_path_note = tk.Label(root, text='What is this for: Read sample IDs from a Excel spreadsheet. The sample IDs should be put in the first column in the spreadsheet.')
ID_path_note.grid(row=5, column=1)

empty = tk.Label(root, text="")
empty.grid(row=6, column=1)
empty = tk.Label(root, text="")
empty.grid(row=7, column=1)

def open_short_file():
    file_path = filedialog.askopenfilename()
    short_path_entry.delete(0, tk.END)
    short_path_entry.insert(0, file_path)

short_path_label = tk.Label(root, text="Excel file (xlsx) for measurement names:", font=('Times New Roman', 15, 'bold'))
short_path_label.grid(row=8, column=0)
short_path_entry = tk.Entry(root, width=120)
short_path_entry.grid(row=8, column=1)
short_path_button = tk.Button(root, text="Select File", command=open_short_file)
short_path_button.grid(row=8, column=2)
short_path_note = tk.Label(root, text='What is this for: Read measurment names from a Excel spreadsheet. The measurement names should be put in the first column in the spreadsheet.')
short_path_note.grid(row=9, column=1)

empty = tk.Label(root, text="")
empty.grid(row=10, column=1)
empty = tk.Label(root, text="")
empty.grid(row=11, column=1)

save_to_folder = tk.StringVar()
def select_save_to_folder():
    save_to_folder.set(filedialog.askdirectory())

copy_to_label = tk.Label(root, text="Destination folder for the retrieved data:", font=('Times New Roman', 15, 'bold'))
copy_to_label.grid(row=12, column=0)
copy_to_entry = tk.Entry(root, width=120, textvariable=save_to_folder)
copy_to_entry.grid(row=12, column=1)
copy_to_button = tk.Button(root, text="Browse", command=select_save_to_folder)
copy_to_button.grid(row=12, column=2)

empty = tk.Label(root, text="")
empty.grid(row=13, column=1)
empty = tk.Label(root, text="")
empty.grid(row=14, column=1)

run_button = tk.Button(root, text="Retrieve data", command=run_code, font=('Times New Roman', 15, 'bold'))
run_button.grid(row=15, column=1)

root.mainloop()