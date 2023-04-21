import os
import smtplib
import email
import collections
from email.mime.base import MIMEBase
from email import encoders
import imapclient
import datetime
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from tkcalendar import DateEntry
import tkinter.messagebox
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pdf2image import convert_from_path
from PIL import Image, ImageTk
import PyPDF2
import docx


frame_index = 0
global results_listbox

downloaded_files_list = []


def close_application():
    root.destroy()

def load_output_folder():
    try:
        with open('output_folder.txt', 'r') as f:
            return f.read().strip()
    except FileNotFoundError:
        return ""

def save_output_folder(folder_path):
    with open('output_folder.txt', 'w') as f:
        f.write(folder_path)

unique_keywords = set()

def load_keywords():
    global unique_keywords
    try:
        with open('keywords.txt', 'r') as f:
            unique_keywords = {line.strip() for line in f.readlines()}
            return list(unique_keywords)
    except FileNotFoundError:
        return []

def save_keyword(keyword):
    global unique_keywords
    unique_keywords.add(keyword)
    with open('keywords.txt', 'w') as f:
        for kw in unique_keywords:
            f.write(f"{kw}\n")

def remove_keyword():
    global unique_keywords
    selected_keyword = search_keyword_entry.get()
    if selected_keyword in unique_keywords:
        unique_keywords.remove(selected_keyword)
        with open('keywords.txt', 'w') as f:
            for kw in unique_keywords:
                f.write(f"{kw}\n")
        search_keyword_entry["values"] = list(unique_keywords)
        if unique_keywords:
            search_keyword_entry.set(list(unique_keywords)[-1])
        else:
            search_keyword_entry.set("")

def clean_file_name(file_name):
    invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|', '\r', '\n']
    for char in invalid_chars:
        file_name = file_name.replace(char, '_')
    return file_name

def select_output_folder():
    folder_selected = filedialog.askdirectory()
    output_folder_entry.delete(0, tk.END)
    output_folder_entry.insert(0, folder_selected)
    save_output_folder(folder_selected)


def send_downloaded_files(sender_email, recipient_email, password, subject, body, files):
    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = recipient_email
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain"))

    for file in files:
        with open(file, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())

        encoders.encode_base64(part)

        part.add_header(
            "Content-Disposition",
            f"attachment; filename={os.path.basename(file)}",
        )

        msg.attach(part)

    context = ssl.SSLContext(ssl.PROTOCOL_TLS)
    context.verify_mode = ssl.CERT_REQUIRED
    context.check_hostname = True
    context.load_default_certs()

    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, recipient_email, msg.as_string())


def download_attachments(email_address, password, search_keyword, start_date, end_date,
                         output_folder, download_all_attachments, file_extension_filter, print_output=True):

    imap = imapclient.IMAPClient("imap.gmail.com", ssl=True)
    imap.login(email_address, password)
    imap.select_folder("INBOX", readonly=True)
    if not search_keyword.strip():
        return []

    # Remove newline characters from search_keyword
    search_keyword = search_keyword.replace("\n", " ").replace("\r", "")

    date_range = f'SINCE "{start_date.strftime("%d-%b-%Y")}" BEFORE "{(end_date + datetime.timedelta(days=1)).strftime("%d-%b-%Y")}"'
    search_subject = f'SUBJECT "{search_keyword}" {date_range}'
    search_body = f'BODY "{search_keyword}" {date_range}'
    if print_output:
        print(search_subject)
        print(search_body)

    messages_subject = imap.search(search_subject, charset="UTF-8")
    messages_body = imap.search(search_body, charset="UTF-8")
    messages = list(set(messages_subject) | set(messages_body))

    downloaded_files = []

    for msg_id in messages:
        msg_data = imap.fetch([msg_id], ["BODY.PEEK[]", "FLAGS"])[msg_id][b"BODY[]"]
        msg = email.message_from_bytes(msg_data)
        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue

            filename = part.get_filename()
            if not filename or (not download_all_attachments and not filename.lower().endswith(file_extension_filter.lower())):
                continue

            local_filename = clean_file_name(filename)
            local_filepath = os.path.join(output_folder, local_filename)

            with open(local_filepath, "wb") as f:
                f.write(part.get_payload(decode=True))

            downloaded_files.append(local_filepath)

    imap.logout()

    return downloaded_files

def import_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        file_name = os.path.basename(file_path)
        files_listbox.insert(tk.END, file_name)
        downloaded_files_list.append(file_path)


def delete_selected_file():
    selected_file_indices = files_listbox.curselection()
    if not selected_file_indices:
        tk.messagebox.showwarning("No File Selected", "Please select a file to delete.")
        return

    for index in reversed(selected_file_indices):
        files_listbox.delete(index)
        del downloaded_files_list[index]

    tk.messagebox.showinfo("File Deleted", "Selected file(s) have been deleted.")

def download_attachments_gui():
    global downloaded_files_list
    email_address = email_address_entry.get()
    password = password_entry.get()
    search_keyword = search_keyword_entry.get().strip()

    if not search_keyword:
        tk.messagebox.showwarning("Invalid Input", "Please enter a search keyword.")
        return
    save_keyword(search_keyword)

    start_date = start_date_entry.get_date()
    end_date = end_date_entry.get_date()
    output_folder = output_folder_entry.get()
    download_all_attachments = all_attachments_var.get()
    file_extension_filter = file_extension_entry.get()

    downloaded_files_list = download_attachments(email_address, password, search_keyword, start_date, end_date,
                                                 output_folder, download_all_attachments, file_extension_filter)

    files_listbox.delete(0, tk.END)
    for file in downloaded_files_list:
        files_listbox.insert(tk.END, os.path.basename(file))

    tk.messagebox.showinfo("Download Complete", "Alle bijlagen zijn gedownload.")


def send_selected_files():
    email_address = email_address_entry.get()
    password = password_entry.get()
    recipient_email = recipient_email_entry.get()

    selected_files_indices = files_listbox.curselection()
    selected_files = [downloaded_files_list[i] for i in selected_files_indices]

    send_downloaded_files(email_address, recipient_email, password, "Downloaded Attachments",
                          "Here are the downloaded attachments.", selected_files)
    tk.messagebox.showinfo("Sending Complete", "De geselecteerde bijlagen zijn verzonden.")

def clear_downloaded_files_list():
    global downloaded_files_list
    downloaded_files_list.clear()
    files_listbox.delete(0, tk.END)

def generate_summary():
    file_counter = collections.Counter()
    total_size = 0

    for file in downloaded_files_list:
        file_extension = os.path.splitext(file)[1]
        file_size = os.path.getsize(file)
        file_counter[file_extension] += 1
        total_size += file_size

    summary_text = f"Total attachments: {len(downloaded_files_list)}\n"
    summary_text += f"Total size: {total_size / 1024:.2f} KB\n"
    summary_text += "\nFile type breakdown:\n"

    for file_ext, count in file_counter.items():
        summary_text += f"{file_ext}: {count}\n"

    tk.messagebox.showinfo("Summary", summary_text)

def download_attachments_all_keywords_gui():
    global downloaded_files_list
    email_address = email_address_entry.get()
    password = password_entry.get()
    start_date = start_date_entry.get_date()
    end_date = end_date_entry.get_date()
    output_folder = output_folder_entry.get()
    download_all_attachments = all_attachments_var.get()
    file_extension_filter = file_extension_entry.get()

    downloaded_files_list = []
    for keyword in unique_keywords:
        current_downloaded_files = download_attachments(email_address, password, keyword, start_date, end_date,
                                                        output_folder, download_all_attachments, file_extension_filter,
                                                        print_output=False)

        downloaded_files_list.extend(current_downloaded_files)

    files_listbox.delete(0, tk.END)
    for file in downloaded_files_list:
        files_listbox.insert(tk.END, os.path.basename(file))

    tk.messagebox.showinfo("Download Complete", "Alle bijlagen zijn gedownload.")

def open_pdf_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        images = convert_from_path(file_path, dpi=300)
        if images:
            image = images[0]
            image.thumbnail((500, 500))
            photo = ImageTk.PhotoImage(image)
            pdf_preview_label.config(image=photo)
            pdf_preview_label.image = photo
        else:
            messagebox.showerror("Error", "Failed to load the PDF file.")


def update_ascii_animation():
    global root, ascii_animation_label, ascii_animation_frames, frame_index

    if root.winfo_exists():  # Check if the main GUI window still exists
        frame_index = (frame_index + 1) % len(ascii_animation_frames)
        ascii_animation_label.config(text=ascii_animation_frames[frame_index])
        root.after(1000, update_ascii_animation)

def save_email_addresses(email_address, recipient_email):
    with open("email_addresses.txt", "w") as file:
        file.write(email_address + "\n")
        file.write(recipient_email + "\n")

def load_email_addresses():
    try:
        with open("email_addresses.txt", "r") as file:
            email_address = file.readline().strip()
            recipient_email = file.readline().strip()
        return email_address, recipient_email
    except FileNotFoundError:
        return "", ""


root = tk.Tk()
root.title("Python Email Downloader")

# E-mail Address
email_address_label = tk.Label(root, text="E-mailadres:")
email_address_label.grid(row=0, column=0, padx=5, pady=5, sticky="W")
email_address_entry = tk.Entry(root)
email_address_entry.grid(row=0, column=1, padx=5, pady=5, sticky="WE")

# Password
password_label = tk.Label(root, text="Wachtwoord:")
password_label.grid(row=1, column=0, padx=5, pady=5, sticky="W")
password_entry = tk.Entry(root, show="*")
password_entry.grid(row=1, column=1, padx=5, pady=5, sticky="WE")

# Search Keyword
search_keyword_label = tk.Label(root, text="Zoekwoord:")
search_keyword_label.grid(row=2, column=0, padx=5, pady=5, sticky="W")
search_keyword_entry = ttk.Combobox(root)
search_keyword_entry.grid(row=2, column=1, padx=5, pady=5, sticky="WE")
remove_keyword_button = tk.Button(root, text="Verwijderen", command=remove_keyword)
remove_keyword_button.grid(row=2, column=2, padx=5, pady=5)

# Start Date
start_date_label = tk.Label(root, text="Startdatum:")
start_date_label.grid(row=3, column=0, padx=5, pady=5, sticky="W")
start_date_entry = DateEntry(root, date_pattern="dd-mm-yyyy")
start_date_entry.grid(row=3, column=1, padx=5, pady=5, sticky="WE")

# End Date
end_date_label = tk.Label(root, text="Einddatum:")
end_date_label.grid(row=4, column=0, padx=5, pady=5, sticky="W")
end_date_entry = DateEntry(root, date_pattern="dd-mm-yyyy")
end_date_entry.grid(row=4, column=1, padx=5, pady=5, sticky="WE")

# Output Folder
output_folder_label = tk.Label(root, text="Uitvoermap:")
output_folder_label.grid(row=5, column=0, padx=5, pady=5, sticky="W")
output_folder_entry = tk.Entry(root)
output_folder_entry.grid(row=5, column=1, padx=5, pady=5, sticky="WE")
browse_button = tk.Button(root, text="Bladeren", command=select_output_folder)
browse_button.grid(row=5, column=2, padx=5, pady=5)

# All Attachments Checkbox
all_attachments_var = tk.IntVar()
all_attachments_checkbox = tk.Checkbutton(root, text="Download alle bijlagen", variable=all_attachments_var)
all_attachments_checkbox.grid(row=6, column=0, columnspan=2, padx=5, pady=5)

# Download All Attachments
download_all_button = tk.Button(root, text="Download All Attachments", command=download_attachments_all_keywords_gui)
download_all_button.grid(row=6, column=3, columnspan=2, pady=(10, 0), padx=5, sticky="w")

# File Extension Filter
file_extension_label = tk.Label(root, text="Bestandsextensie filter:")
file_extension_label.grid(row=7, column=0, padx=5, pady=5, sticky="W")
file_extension_entry = tk.Entry(root)
file_extension_entry.grid(row=7, column=1, padx=5, pady=5, sticky="WE")
file_extension_entry.insert(0, ".pdf")

# Recipient Email
recipient_email_label = tk.Label(root, text="Ontvanger E-mail:")
recipient_email_label.grid(row=8, column=0, padx=5, pady=5, sticky="W")
recipient_email_entry = tk.Entry(root)
recipient_email_entry.grid(row=8, column=1, padx=5, pady=5, sticky="WE")

# Download Button
download_button = tk.Button(root, text="Downloaden", command=download_attachments_gui)
download_button.grid(row=9, column=0, padx=5, pady=10)

# Files Listbox
files_listbox = tk.Listbox(root, selectmode=tk.MULTIPLE)
files_listbox.grid(row=9, column=1, padx=5, pady=10, sticky="WE")

# Send Button
send_button = tk.Button(root, text="Verzenden", command=send_selected_files)
send_button.grid(row=9, column=2, padx=5, pady=10)

# Clear Button
clear_button = tk.Button(root, text="Lijst wissen", command=clear_downloaded_files_list)
clear_button.grid(row=9, column=3, padx=5, pady=10)

# Summary Button
summary_button = tk.Button(root, text="Summary", command=generate_summary)
summary_button.grid(row=9, column=4, padx=5, pady=10)

# Progress Bar
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_var)
progress_bar.grid(row=10, column=0, columnspan=3, padx=5, pady=5, sticky="WE")

# Delete Button
delete_button = tk.Button(root, text="Delete Selected", command=delete_selected_file)
delete_button.grid(row=8, column=3, columnspan=2, pady=(10, 0), padx=10, sticky="WE")

pdf_button = tk.Button(root, text="Open PDF", command=open_pdf_file)
pdf_button.grid(row=11, column=0, padx=5, pady=5)

quit_button = tk.Button(root, text="Afsluiten", command=close_application)
quit_button.grid(row=14, column=4, padx=5, pady=10)

pdf_preview_label = tk.Label(root)
pdf_preview_label.grid(row=11, column=1, padx=5, pady=5, columnspan=4)

import_pdf_button = tk.Button(root, text="Import PDF", command=import_pdf)
import_pdf_button.grid(row=12, column=4, sticky="W", padx=(10, 0), pady=(10, 10))

ascii_animation_frames = [
    """

     __  __            _      ______               _ _                                
    |  \/  |          | |    |___  /              | | |                               
    | \  / | __ _ _ __| | __    / / __ _ _ __   __| | |__   ___ _ __ __ _  ___ _ __   
    | |\/| |/ _` | '__| |/ /   / / / _` | '_ \ / _` | '_ \ / _ \ '__/ _` |/ _ \ '_ \  
    | |  | | (_| | |  |   <   / /_| (_| | | | | (_| | |_) |  __/ | | (_| |  __/ | | | 
    |_|  |_|\__,_|_|  |_|\_\ /_____\__,_|_| |_|\__,_|_.__/ \___|_|  \__, |\___|_| |_| 
                                                                     __/ |            
                                                                    |___/             
    """,
    """


     __  __            _      ______               _ _                                
    |  \/  |          | |    |___  /              | | |                               
    | \  / | __ _ _ __| | __    / / __ _ _ __   __| | |__   ___ _ __ __ _  ___ _ __   
    | |\/| |/ _` | '__| |/ /   / / / _` | '_ \ / _` | '_ \ / _ \ '__/ _` |/ _ \ '_ \  
    | |  | | (_| | |  |   <   / /_| (_| | | | | (_| | |_) |  __/ | | (_| |  __/ | | | 
    |_|  |_|\__,_|_|  |_|\_\ /_____\__,_|_| |_|\__,_|_.__/ \___|_|  \__, |\___|_| |_| 
                                                                     __/ |            
                                                                    |___/             
    """,
    """
     __  __            _      ______               _ _                                
    |  \/  |          | |    |___  /              | | |                               
    | \  / | __ _ _ __| | __    / / __ _ _ __   __| | |__   ___ _ __ __ _  ___ _ __   
    | |\/| |/ _` | '__| |/ /   / / / _` | '_ \ / _` | '_ \ / _ \ '__/ _` |/ _ \ '_ \  
    | |  | | (_| | |  |   <   / /_| (_| | | | | (_| | |_) |  __/ | | (_| |  __/ | | | 
    |_|  |_|\__,_|_|  |_|\_\ /_____\__,_|_| |_|\__,_|_.__/ \___|_|  \__, |\___|_| |_| 
                                                                     __/ |            
                                                                    |___/             
    """
]

# ASCII Animation
ascii_animation_label = tk.Label(root, text="", wraplength=400, height=9, width=50)
ascii_animation_label.grid(row=12, column=0, columnspan=3, pady=(10, 0), padx=10)

email_address, recipient_email = load_email_addresses()
email_address_entry.insert(0, email_address)
recipient_email_entry.insert(0, recipient_email)

# Window Configuration
window_width = 680
window_height = 1100
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_coord = int((screen_width / 2) - (window_width / 2))
y_coord = int((screen_height / 2) - (window_height / 2))
root.geometry(f"{window_width}x{window_height}+{x_coord}+{y_coord}")

update_ascii_animation()


stored_keywords = load_keywords()
if stored_keywords:
    search_keyword_entry["values"] = stored_keywords
    search_keyword_entry.set(stored_keywords[-1])

stored_output_folder = load_output_folder()
if stored_output_folder:
    output_folder_entry.insert(0, stored_output_folder)

def on_closing():
    email_address = email_address_entry.get()
    recipient_email = recipient_email_entry.get()
    save_email_addresses(email_address, recipient_email)
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)

root.mainloop()

