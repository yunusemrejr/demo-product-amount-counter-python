import tkinter as tk
import os
import openpyxl
import re
from tkinter import filedialog


def search_in_excel(path, query):
    file_path = path
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
    except Exception as e:
        return f"An ERROR occured while reading the excel file! : {str(e)}"
  
    query = query.lower()  
    matching_rows = []
    lisans_sayisi = 0
  
    for row in sheet.iter_rows(min_col=1, max_col=1, values_only=True): 
        cell_value = row[0]
        if cell_value is not None:
            cell_value = str(cell_value).lower()
            mini_row_arr = cell_value.split(',')
            for mini_row in mini_row_arr:

                regex_pattern = r'(\d+x\s*){}\b'.format(re.escape(query))
                license_counts = re.findall(regex_pattern, mini_row, re.IGNORECASE)  
                
                if license_counts:
                    matching_rows.append(mini_row)
                    lisans_sayisi += sum(int(count.split('x')[0]) for count in license_counts)
    
    if matching_rows:
        return matching_rows, lisans_sayisi
    else:
        return [], 0

def is_excel():
    current_directory = os.getcwd()
    files_in_directory = os.listdir(current_directory)
    
    excel_files = [file_name for file_name in files_in_directory if file_name.endswith((".xlsx", ".xls"))]
    
    if len(excel_files) == 1:
        return excel_files[0]
    elif len(excel_files) > 1:
        return excel_files
    else:
        return False

def start_procedure():
    search_term = input_box.get()
    
    if len(search_term) != 0:

        excel_file = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        
        if excel_file:
            matching_rows, lisans_sayisi = search_in_excel(excel_file, search_term)
            result_label.config(text=f"\n\nYour query: {search_term}\n\nYour spreadsheet: {excel_file}\n\nTotal LICENSE count: {lisans_sayisi}\n\n")
            
            licenses_text.config(state=tk.NORMAL)
            licenses_text.delete("1.0", tk.END)
            licenses_text.insert(tk.END, f"{len(matching_rows)} rows matched with your query.(case-insensitive).\n\nTotal LICENSE COUNT -->: ")
            licenses_text.insert(tk.END, str(lisans_sayisi), "italic")
            licenses_text.tag_configure("italic", font=("Helvetica", 12, "italic"))
            licenses_text.config(state=tk.DISABLED)
        else:
            result_label.config(text="No Excel file selected. Please choose an Excel file.")
    else:
        result_label.config(text="You need to give me a query!")

root = tk.Tk()
root.title("Demo License Counter")

 

frame = tk.Frame(root)
frame.grid(row=0, column=0, padx=10, pady=10)

window_width = 800
window_height = 600
root.geometry(f"{window_width}x{window_height}")

label = tk.Label(frame, text="License Type:")
label.grid(row=0, column=0)

input_box = tk.Entry(frame)
input_box.grid(row=0, column=1)

search_button = tk.Button(frame, text="Search", command=start_procedure)
search_button.grid(row=0, column=2)

result_label = tk.Label(frame, text="", fg="yellow")
result_label.grid(row=1, column=0, columnspan=3)

licenses_text = tk.Text(frame, wrap=tk.WORD, height=5, width=50)
licenses_text.grid(row=2, column=0, columnspan=3)
licenses_text.config(state=tk.DISABLED)  

root.mainloop()
