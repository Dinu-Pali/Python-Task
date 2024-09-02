import pandas as pd
import tkinter as tk
from tkinter import ttk
import pyperclip
from docx import Document

# Load the Excel file
file_path = 'Snr Appt.xlsx'
df = pd.read_excel(file_path)

# Load the Word file
word_file_path = 'REG NEW.docx'
document = Document(word_file_path)

# Define the correct column names for searching
search_column = 'PRESENT APPT'
regt_no_column = 'REGT NO'

# Create the main application window
root = tk.Tk()
root.title("Excel and Word Search Tool")

# Function to update suggestions based on the input
def update_suggestions(*args):
    search_term = entry_var.get().lower()
    suggestions = df[df[search_column].str.contains(search_term, case=False, na=False)][search_column].tolist()
    listbox.delete(0, tk.END)
    for suggestion in suggestions:
        listbox.insert(tk.END, suggestion)

# Function to display the row data when a suggestion is selected
def show_details(event):
    selected_item = listbox.get(listbox.curselection())
    row_data = df[df[search_column] == selected_item].iloc[0]
    regt_no = row_data[regt_no_column]
    details_var.set("\n".join([f"{col}: {row_data[col]}" for col in df.columns]))
    
    # Search the Word document for the REGT NO value
    word_details = search_word_file(regt_no)
    word_details_var.set(word_details)

# Function to search the Word file based on REGT NO
def search_word_file(regt_no):
    for table in document.tables:
        for row in table.rows:
            if row.cells[0].text.strip() == str(regt_no):
                return "\n".join([cell.text for cell in row.cells])
    return "No matching details found in the Word file."

# Function to copy the details to clipboard
def copy_details():
    pyperclip.copy(details_var.get())

# UI Elements
entry_var = tk.StringVar()
entry_var.trace("w", update_suggestions)

entry_label = tk.Label(root, text="Enter Search Term:")
entry_label.pack(pady=5)

entry_box = tk.Entry(root, textvariable=entry_var, width=50)
entry_box.pack(pady=5)

listbox = tk.Listbox(root, height=10, width=50)
listbox.pack(pady=5)
listbox.bind("<<ListboxSelect>>", show_details)

details_label = tk.Label(root, text="Excel Details:")
details_label.pack(pady=5)

details_var = tk.StringVar()
details_display = tk.Label(root, textvariable=details_var, justify="left", wraplength=400, relief="sunken", anchor="w")
details_display.pack(pady=5)

copy_button = tk.Button(root, text="Copy Details", command=copy_details)
copy_button.pack(pady=5)

word_details_label = tk.Label(root, text="Word File Details:")
word_details_label.pack(pady=5)

word_details_var = tk.StringVar()
word_details_display = tk.Label(root, textvariable=word_details_var, justify="left", wraplength=400, relief="sunken", anchor="w")
word_details_display.pack(pady=5)

# Start the Tkinter event loop
root.mainloop()
