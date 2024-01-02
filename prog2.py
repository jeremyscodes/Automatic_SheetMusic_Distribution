# pip install -r requirements.txt
#To run from cmd terminal: python prog2.py


# So Far: 
# 1. Reads google survey csv, to get email address_ instrument mapping
# 2. Reads zip file, to get extract instrument names from pdf titles
# 3. Matches instrument names to email addresses
# 4. Sends emails to email addresses with pdfs attached

# 5. For PDFs that were not matched to an email address, allow user to manually label the pdfs with the right instruments
# Consider getting newest version of google form from online #DONE
#TODO 
# 6. package program and send to jyo for testing
import tkinter as tk
from tkinter import filedialog
import csv
import zipfile
import io
import os
import smtplib
import threading
import time
from email.message import EmailMessage

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

def get_updated_csv():

    from assets.googe import create_service
    CLIENT_SECRET_FILE = 'assets/credentials.json'
    API_NAME = 'sheets'
    API_VERSION = 'v4'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

    service = create_service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

    import pandas as pd
    response = service.spreadsheets().values().get(
        spreadsheetId='1l2yTzu1oEmUt27LbqiaFEAa660hVhs7mCJgxgtJDrEM',
        majorDimension='ROWS',
        range='Form responses 1'
    ).execute()

    columns = response['values'][0]
    data = response['values'][1:]
    df = pd.DataFrame(data, columns=columns)
    # convert to csv
    df.to_csv('assets/devnodes.csv', index=False)
    return 'assets/devnodes.csv'

# This method reads in the survey csv and extracts email - instrument pairs
def parse_csv(filepath):
    instruments_mapping = {}  # Dictionary to store instruments mapped to email addresses

    try:
        print("Opening csv")
        with open(filepath, 'r') as csv_file:
            csv_reader = csv.reader(csv_file)
            headers = next(csv_reader)  # Skip the header row
            # print(headers)
            # print("Processing csv")
            for row in csv_reader:
                print(row)
                if row != ['', '', '']:
                    timestamp, email, instrument = row
                    instruments_mapping[email] = instrument.lower()

        return instruments_mapping
    except FileNotFoundError:
        print("File not found.")
        return None
    except Exception as e:
        print("Error:", str(e))
        return None

# This method opens a zip and extracts the survey information
csv_instrument_mapping = None  # Variable to hold the instrument mapping from CSV

def import_csv(button):
    global csv_instrument_mapping

    filepath = get_updated_csv()
    if filepath:
        csv_instrument_mapping = parse_csv(filepath)
        if csv_instrument_mapping:
            for email, instrument in csv_instrument_mapping.items():
                print(f"Email: {email}, Instrument: {instrument}")
            # You can do further processing with the parsed data here
        # button.config(background='#90EE90')
        button.config(highlightbackground='#90EE90') 
        return csv_instrument_mapping
    else:
        print(filepath)

def loading_task(fetch_button):
    # Display a loading indication (you can use a label or any other method)
    fetch_button.config(highlightbackground="#e5ed4e") #yellow
    print("Loading...")

# Function to start the import process in a separate thread
def start_import(button):

    loading_thread = threading.Thread(target=lambda: loading_task(button))

    import_thread = threading.Thread(target=import_csv, args=(button,))
    
    loading_thread.start()  # Start the loading indication thread
    import_thread.start()  # Start the CSV import thread


import re
instrument_list = ['violin', 'viola', 'cello','cellos', 'contrabass', 'basses','bass guitar', 'bass trombone',
                   'bassoon', 'clarinet','cls',  'flute', 'glockenspiel', 'horn', 'oboe',
                   'percussion', 'perc' , 'drum', 'piccolo', 'timpani', 'timp','bells', 'trombone', 'trumpet', 'tuba', 'sax', 'harp', 'piano','score']

def extract_proper_instrument(file_name):
    file_lower = file_name.lower()  # Convert filename to lowercase for case insensitivity
    
    # identify instrument and group into instrument type
    for instrument in instrument_list:
        if instrument in file_lower:
            if instrument == 'cls':
                return 'woodwinds'
            
            
            # PERCUSSION
            elif instrument == 'perc':
                return 'percussion'
            elif instrument == 'timp' or instrument == 'timpani':
                return 'percussion'
            elif instrument == 'bells':
                return 'percussion'
            elif instrument == 'drum' or instrument == 'drum set':
                return 'percussion'
            elif instrument == 'glockenspiel':
                return 'percussion'
            
            # BRASS
            elif instrument == 'bass trombone':
                return 'brass'
            elif instrument == 'trombone':
                return 'brass'
            elif instrument == 'trumpet':
                return 'brass'
            elif instrument == 'tuba':
                return 'brass'
            elif instrument == 'horn':
                return 'brass'
            elif instrument == 'sax':
                return 'brass'
            elif instrument == 'trumpet':
                return 'brass'
            
            # WOODWINDS
            elif instrument == 'flute':
                return 'woodwinds'
            elif instrument == 'piccolo':
                return 'woodwinds'
            elif instrument == 'oboe':
                return 'woodwinds'
            elif instrument == 'clarinet':
                return 'woodwinds'
            elif instrument == 'bassoon':
                return 'woodwinds'
            

            elif instrument == 'basses':
                return 'contrabass'
            elif instrument == 'cellos':
                return 'cello/contrabass'
            elif instrument == 'cello':
                return 'cello/contrabass'
            elif instrument == 'basses':
                return 'cello/contrabass'
            elif instrument == 'bass':
                return 'cello/contrabass'
            elif instrument == 'contrabass':
                return 'cello/contrabass'
            
            elif instrument == 'bass guitar':
                return 'bass/electric guitar'
            elif instrument == 'electric guitar':
                return 'bass/electric guitar'
            
            elif instrument == 'piano':
                return 'piano/harp'
            elif instrument == 'harp':
                return 'piano/harp'
            
            elif instrument == 'score':
                return 'conductor'
            
            # else:
            #     print("instrument did not match any of the cases", instrument)
            return instrument
    else: return None  # Return None if no valid instrument is found
    

import tempfile

def extract_pdf_info(zip_path):
    pdf_instrument = {}  # Dictionary to store instrument of the pdf
    temp_folder = tempfile.mkdtemp()  # Create a temporary folder to extract PDFs
    
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        for file_name in zip_ref.namelist():
            if not file_name.startswith('__MACOSX/') and file_name.lower().endswith('.pdf'):
                extract_to = os.path.join(temp_folder, file_name)
                zip_ref.extract(file_name, temp_folder)
                instrument_name = extract_proper_instrument(file_name)
                
                pdf_instrument[extract_to] = instrument_name  # Storing path instead of file_name    
                    
    return pdf_instrument

def match_pdfs_to_emails(pdf_instrument, instrument_mapping):
    email_pdf_dict = {}  # Dictionary to store matched PDFs with instruments
    
    # Iterate through the instrument mapping to match PDFs to emails
    for email, instrument in instrument_mapping.items():
        for pdf_name, pdf_instrument_name in pdf_instrument.items():
            if pdf_instrument_name == instrument:
                if email not in email_pdf_dict:
                    email_pdf_dict[email] = []
                email_pdf_dict[email].append(pdf_name)
    
    # Display email_pdf_dict nicely in prints
    for email, pdf_paths in email_pdf_dict.items():
        print(email, ":", pdf_paths)
    
    return email_pdf_dict


sheet_music_name =""
email_pdf_dict = None  # Dictionary to store email - pdf pairs
pdf_instrument = None  # Dictionary to store pdf - instrument pairs
def upload_sheet_music(music_button,mail_button):
    global csv_instrument_mapping  # Use the global variable
    global sheet_music_name
    global email_pdf_dict
    global pdf_instrument
    mail_button.config(background='SystemButtonFace')
    music_button.config(background='SystemButtonFace')

    if csv_instrument_mapping:  # Check if the CSV data has been imported
        zip_path = filedialog.askopenfilename(filetypes=[("Zip Files", "*.zip")])
        sheet_music_name = os.path.splitext(os.path.basename(zip_path))[0]
        if zip_path:
            pdf_instrument = extract_pdf_info(zip_path)
            # print(pdf_instrument)
            email_pdf_dict = match_pdfs_to_emails(pdf_instrument, csv_instrument_mapping)
            # print("email_pdf_dict:   " ,email_pdf_dict)
    else:
        print("Please import CSV first.")
    if email_pdf_dict:
        # music_button.config(background="#90EE90")
        music_button.config(highlightbackground="#90EE90")

import zipfile
import os
import tempfile

def extract_pdfs_from_zip(zip_path):
    extract_to = tempfile.mkdtemp()
    pdf_files = []
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        for filename in zip_ref.namelist():
            if filename.endswith('.pdf'):
                zip_ref.extract(filename, extract_to)
                pdf_files.append(os.path.join(extract_to, filename))
    # print(pdf_files)
    return pdf_files

import yagmail

yag = yagmail.SMTP(email_sender, email_password)



import tkinter as tk
from tkinter import ttk

def handle_instrument_selection(pdf_title, instrument_var):
    global email_pdf_dict
    print(f"Selected instrument for '{pdf_title}': {instrument_var}")
    pdf_instrument[pdf_title] = instrument_var
    email_pdf_dict = match_pdfs_to_emails(pdf_instrument, csv_instrument_mapping)


def send_emails(button):
    global email_pdf_dict
    
    if email_pdf_dict:
        window = tk.Tk()
        window.title("Select Instrument Groups")
        
        instrument_sections = ['violin', 'viola', 'cello/contrabass', 'woodwinds', 'brass', 'percussion', 'piano/harp', 'conductor','bass/electric guitar']
        
        max_per_column = 8  # Maximum number of columns to display
        current_column = 0
        frame = None
        
        all_pdf_titles = list(pdf_instrument.keys())
        
        for pdf_title in all_pdf_titles:
            # Create a new frame for each column
            if current_column % max_per_column == 0:
                frame = ttk.Frame(window)
                frame.pack(side='left', padx=10)
            n=os.path.splitext(os.path.basename(pdf_title))[0]
            label_pdf = tk.Label(frame, text=f"{n}")
            label_pdf.pack(anchor='w')
            
            instrument_var = tk.StringVar()
            instrument_dropdown = ttk.Combobox(frame, textvariable=instrument_var)
            instrument_dropdown['values'] = instrument_sections
            
            # Set the initial instrument selection if available
            instrument_group = pdf_instrument[pdf_title]
            if instrument_group in instrument_sections:
                instrument_dropdown.set(instrument_group)
            
            instrument_dropdown.pack()
            
            # Bind the instrument selection to the handle_instrument_selection function
            # instrument_dropdown.bind('<<ComboboxSelected>>', lambda event, pdf_title=pdf_title, instrument_var=instrument_var: handle_instrument_selection(pdf_title, instrument_var.get()))
            # Python
            instrument_dropdown.bind('<<ComboboxSelected>>', lambda event, pdf_title=pdf_title, instrument_dropdown=instrument_dropdown: handle_instrument_selection(pdf_title, instrument_dropdown.get()))
            current_column += 1
                
        

        def show_email():
            try:

                # Create a new window to display the output
                output_window = tk.Tk()
                output_window.title("Email Sending Confirmation")
                
                
                # Create a Text widget to display the output
                output_text = tk.Text(output_window, wrap="word")
                output_text.pack(fill="both", expand=True)
                # Function to write the print statements to the Text widget
                

                num_scores=0
                for email, pdf_paths in email_pdf_dict.items():
                    contents = ['']
                    contents.extend(pdf_paths)  # Add PDF files to email contents
                    num_scores+=len(pdf_paths)

                    output_text.insert(tk.END, f"\nEmail sent to: {email}\n")
                    instrument = csv_instrument_mapping[email]
                    output_text.insert(tk.END, f"Instrument played: {instrument}\n")
                    # extract basename from all pdf paths and join them with commas
                    sent_pdfs = [os.path.splitext(os.path.basename(pdf_path))[0] for pdf_path in pdf_paths]
                    output_text.insert(tk.END, f"PDFs to send: {', '.join(sent_pdfs)}\n")
                    
                output_text.insert(tk.END, f"Total number of scores to be sent: {num_scores}\n")
                
                # Confirm button
                confirm_button = tk.Button(output_window, text="Confirm", command=lambda: confirm_send(output_window))

                confirm_button.pack(side=tk.LEFT, padx=10, pady=10)
                # Cancel button
                def destroy_window(wind):
                    wind.destroy()
                cancel_button = tk.Button(output_window, text="Cancel", command=lambda: destroy_window(output_window))
                cancel_button.pack(side=tk.RIGHT, padx=10, pady=10)


                output_window.mainloop()
            except Exception as e:
                print("Error:", str(e))

        def actual_send():
            try:
                # sending emails
                for email, pdf_paths in email_pdf_dict.items():
                    subject = 'New Sheet Music From JYO: ' + sheet_music_name
                    contents = ['']
                    contents.extend(pdf_paths)  # Add PDF files to email contents
                    num_scores = len(pdf_paths)
                    
                    yag.send(email, subject, contents)
                                                
                print("Emails Sent Successfully")

            except Exception as e:
                print("Error:", str(e))

        def confirm_send(wind):
            output_window = tk.Toplevel()
            output_window.title("Email Sending Status")
            
            output_text = tk.Text(output_window, wrap="word")
            output_text.pack(fill="both", expand=True)
            output_text.insert(tk.END, "\nPlease wait...\n")

            # Create a separate thread for email sending
            email_thread = threading.Thread(target=actual_send)
            email_thread.start()

            # Update the output window when the email thread completes
            def check_email_thread():
                if email_thread.is_alive():
                    output_text.delete(1.0, tk.END)
                    output_text.insert(tk.END, "\nPlease wait...\n")
                    output_text.insert(tk.END, "\nSending emails...")
                    output_text.after(1000, check_email_thread)
                else:
                    output_text.delete(1.0, tk.END)
                    output_text.insert(tk.END, "\nEmails Sent Successfully\n")

            check_email_thread()
            button.after(100, lambda: button.config(highlightbackground='#90EE90'))
            # button.after(100, lambda: button.config(background='#90EE90'))

        mini_send_button = tk.Button(window, text="Next", command=show_email)
        mini_send_button.pack()
        
        window.mainloop()
        
    else:
        print("No PDFs to display.")

class PrintRedirector:
    def __init__(self, widget, print_func):
        self.widget = widget
        self.print_func = print_func

    def write(self, text):
        self.widget.insert(tk.END, text)
        self.widget.see(tk.END)
# Python
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None

    def show_tooltip(self):
        x = self.widget.winfo_rootx() + self.widget.winfo_width() // 2
        y = self.widget.winfo_rooty() + self.widget.winfo_height()

        # Creates a toplevel window
        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)  # Removes the window decorations
        self.tooltip.wm_geometry(f"+{x}+{y}")

        message = tk.Message(self.tooltip, text=self.text, width=200)
        message.pack()

    def hide_tooltip(self):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None



# Creating the GUI
root = tk.Tk()
root.title("Sheet Music Distributor")


# Adjust the window size to full screen upon launch
root.geometry("{0}x{1}+0+0".format(root.winfo_screenwidth(), root.winfo_screenheight()))


# Python
import tkinter as tk
from tkinter import ttk  # ttk is tkinter's themed widget library

# Function to update button size based on window size
def update_button_size(event):
    button_width = root.winfo_width() // 40  # Adjust based on the desired ratio
    button_height = root.winfo_height() // 200  # Adjust based on the desired ratio
    import_button.config(width=button_width, height=button_height)
    upload_button.config(width=button_width, height=button_height)
    send_button.config(width=button_width, height=button_height)

# Create a style object
style = ttk.Style()

# Configure the style
style.configure("TButton",
                foreground="midnight blue",
                background="light sky blue",
                font=("Helvetica", 16, "bold"),
                padding=10)

# style.configure("TLabel",
#                 foreground="midnight blue",
#                 background="white",
#                 font=("Helvetica", 24, "bold"),
#                 padding=10)

# Create a frame to hold the buttons
frame = ttk.Frame(root, padding="10 10 10 10")
frame.pack(fill=tk.BOTH, expand=True)


# Button to import CSV
import_button = tk.Button(frame, text="Fetch Updated Form Responses", command=lambda: start_import(import_button),font=('Helvetica', '20'))
import_button.grid(column=0, row=0, padx=10, pady=10)

# Button to upload sheet music
upload_button = tk.Button(frame, text="Upload Sheet Music", command=lambda: upload_sheet_music(upload_button,send_button), font=('Helvetica', '20'))
upload_button.grid(column=0, row=1, padx=10, pady=10)

# Button to send emails
send_button = tk.Button(frame, text="Send Emails", command=lambda: send_emails(send_button),  font=('Helvetica', '20'))
send_button.grid(column=0, row=2, padx=10, pady=10)

# Query default background and foreground colors
default_bg_color = send_button.cget("background")
default_fg_color = send_button.cget("foreground")

print(f"Default Background Color: {default_bg_color}")
print(f"Default Foreground Color: {default_fg_color}")

# Create a button with a question mark
help_button = ttk.Button(frame, text="?", style="TButton")
help_button.grid(column=0, row=0, padx=10, pady=10, sticky='w')


# Create a ToolTip for the button
tooltip = ToolTip(help_button, "1) Click 'Import CSV to get the up-to-date email addresses for orchestra members from the google form. \n 2) Upload a ZIP folder containing the sheet music for distribution. \n 3) Click 'Send Emails' to confirm the music allocations and distribute the sheet music to the orchestra members.")
help_button.bind("<Enter>", lambda e: tooltip.show_tooltip())
help_button.bind("<Leave>", lambda e: tooltip.hide_tooltip())


root.bind("<Configure>", update_button_size)

# Make the grid columns and rows expand when the window is resized
frame.columnconfigure(0, weight=1)
for i in range(4):
    frame.rowconfigure(i, weight=1)
root.mainloop()

