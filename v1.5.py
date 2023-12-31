import smtplib
import time
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, Menu, ttk, simpledialog
import threading
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import json
import subprocess
import warnings
import urllib.error, urllib.request
from tidylib import tidy_document
import os
import socket
import csv

os.environ['PYTHONWARNINGS'] = 'ignore'
warnings.simplefilter("ignore", DeprecationWarning)
warnings.simplefilter("ignore", category=PendingDeprecationWarning)
warnings.simplefilter("ignore", category=DeprecationWarning)

emails_loaded = False
should_stop = False


error_dialog = None



import re

emailRegex = re.compile(r'''
#example :
#something-.+_@somedomain.com
(
([a-zA-Z0-9_.+]+
@
[a-zA-Z0-9_.+]+)
)
''', re.VERBOSE)


class HighlightText(tk.Text):
    def __init__(self, *args, **kwargs):
        tk.Text.__init__(self, *args, **kwargs)
        self.configure_tags()
        self.bind("<KeyRelease>", self.on_key_release)
        self.re_highlight_pattern = re.compile(r"<[^<>]+>")

    def configure_tags(self):
        self.tag_configure("HTML_TAG", foreground="blue")

    def on_key_release(self, event=None):
        self.remove_tags("1.0", tk.END)
        self.apply_tags("1.0", tk.END)

    def remove_tags(self, start, end):
        self.tag_remove("HTML_TAG", start, end)

    def apply_tags(self, start, end):
        text = self.get(start, end)
        for match in self.re_highlight_pattern.finditer(text):
            self.tag_add("HTML_TAG", f"{start} + {match.start()} chars", f"{start} + {match.end()} chars")


def open_email_finder_gui():
    email_finder_win = tk.Toplevel()
    email_finder_win.title("Email Finder")

    # Left side for URL input
    url_input_label = tk.Label(email_finder_win, text="Enter URLs:")
    url_input_label.pack(pady=5, anchor=tk.W)
    url_input_text = scrolledtext.ScrolledText(email_finder_win, width=40, height=20)
    url_input_text.pack(side=tk.LEFT, padx=10, pady=10)

    # Right side for displaying emails
    email_display_label = tk.Label(email_finder_win, text="Found Emails:")
    email_display_label.pack(pady=5, anchor=tk.W)
    email_display_text = scrolledtext.ScrolledText(email_finder_win, width=40, height=20)
    email_display_text.pack(side=tk.RIGHT, padx=10, pady=10)

    progress_label = tk.Label(email_finder_win, text="Progress:")
    progress_label.pack(pady=5, anchor=tk.W)
    progress = ttk.Progressbar(email_finder_win, orient="horizontal", length=200, mode="determinate")
    progress.pack(pady=5)

    def find_emails():
        urls = url_input_text.get("1.0", tk.END).strip().split('\n')
        total_urls = len(urls)
        progress["maximum"] = total_urls
        found_emails = []

        for url in urls:
            try:
                urllib.parse.urlparse(url)
            except ValueError:
                show_warning("Error", f"Invalid URL: {url}")
                progress["maximum"] -= 1
                continue

            found_emails.extend(find_emails_from_single_url(url))
            progress["value"] += 1
            email_finder_win.update_idletasks()

        # Display the found emails
        for email in found_emails:
            email_display_text.insert(tk.END, email + "\n")

        progress["value"] = total_urls

    def clear_inputs():

        url_input_text.delete('1.0', tk.END)
        email_display_text.delete('1.0', tk.END)
        progress["value"] = 0

    def save_list():
        emails_to_save = email_display_text.get("1.0", tk.END).strip()
        if emails_to_save:
            try:
                file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
                if file_path:
                    with open(file_path, 'w') as file:
                        file.write(emails_to_save)
                    messagebox.showinfo("Success", "List saved successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Error saving file: {str(e)}")
        else:
            messagebox.showwarning("Warning", "No emails to save.")

    find_button = tk.Button(email_finder_win, text="Find Emails", command=find_emails)
    find_button.pack(pady=10)
    save_list_button = tk.Button(email_finder_win, text="Save Emails", command=save_list)
    save_list_button.pack(pady=10)
    clear_button = tk.Button(email_finder_win, text="Clear", command=clear_inputs)
    clear_button.pack(pady=10)


def find_emails_from_single_url(url):

    try:

        urllib.parse.urlparse(url)

        urlText = htmlPageRead(url)
        return extractEmailsFromUrlText(urlText)
    except ValueError:
        show_warning("Error", f"Invalid URL: {url}")

        return []
    except urllib.error.HTTPError as err:
        if err.code == 404:
            cache_url = 'http://webcache.googleusercontent.com/search?q=cache:' + url
            try:
                urlText = htmlPageRead(cache_url)
                return extractEmailsFromUrlText(urlText)
            except:
                show_warning("Error", f"Could not find emails for URL: {url}")
                return []
        else:
            show_warning("Error", f"HTTP Error {err.code} occurred for URL: {url}")
            return []

def show_warning(title, message):
    tk.messagebox.showwarning(title, message)

def extractEmailsFromUrlText(urlText):
    extractedEmail = emailRegex.findall(urlText)
    emails = [email[0] for email in extractedEmail]
    return emails

def htmlPageRead(url):
    headers = { 'User-Agent' : 'Mozilla/5.0' }
    request = urllib.request.Request(url, None, headers)
    response = urllib.request.urlopen(request)
    urlHtmlPageRead = response.read()
    urlText = urlHtmlPageRead.decode()
    return urlText

def emailsFromUrl(url):
    try:
        urlText = htmlPageRead(url)
        return extractEmailsFromUrlText(urlText)
    except urllib.error.HTTPError as err:
        if err.code == 404:
            cache_url = 'http://webcache.googleusercontent.com/search?q=cache:' + url
            urlText = htmlPageRead(cache_url)
            return extractEmailsFromUrlText(urlText)
        else:
            return []

def find_emails_from_urls(url_list):
    all_emails = []
    for url in url_list:
        emails = emailsFromUrl(url)
        all_emails.extend(emails)
    return list(set(all_emails))


def open_html_editor():
    html_win = tk.Toplevel()
    html_win.title("HTML Editor")
    menu_bar = tk.Menu(html_win)
    html_win.config(menu=menu_bar)
    advanced_html_editor_frame = tk.Frame(html_win)
    advanced_html_editor_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    advanced_html_editor = HighlightText(advanced_html_editor_frame, wrap=tk.WORD)
    advanced_html_editor.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = tk.Scrollbar(advanced_html_editor_frame, command=advanced_html_editor.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    advanced_html_editor.config(yscrollcommand=scrollbar.set)

    current_content = html_editor.get("1.0", tk.END)
    advanced_html_editor.insert("1.0", current_content)



    def cut_text():
        advanced_html_editor.event_generate("<<Cut>>")

    def copy_text():
        advanced_html_editor.event_generate("<<Copy>>")

    def paste_text():
        advanced_html_editor.event_generate("<<Paste>>")

    def undo():
        advanced_html_editor.event_generate("<<Undo>>")

    def redo():
        advanced_html_editor.event_generate("<<Redo>>")

    def find_text():
        find_str = simpledialog.askstring("Find", "Enter text to find:")
        if find_str:
            start = advanced_html_editor.search(find_str, "1.0", stopindex=tk.END)
            if start:
                end = f"{start}+{len(find_str)}c"
                advanced_html_editor.tag_add(tk.SEL, start, end)
                advanced_html_editor.mark_set(tk.INSERT, end)
                advanced_html_editor.see(tk.INSERT)

    def validate_html():
        html_text = advanced_html_editor.get("1.0", tk.END)
        document, errors = tidy_document(html_text, options={'output-xhtml': 1, 'show-warnings': 0})

        if not errors:
            messagebox.showinfo('Validation Result', 'HTML is valid.')
        else:
            messagebox.showerror('Validation Result', f'HTML is not valid. {errors}')

    def on_right_click(event):
        context_menu.post(event.x_root, event.y_root)

    def cancel_context_menu():
        context_menu.unpost()

    context_menu = tk.Menu(html_win, tearoff=0)
    context_menu.add_command(label="Cut", command=cut_text)
    context_menu.add_command(label="Copy", command=copy_text)
    context_menu.add_command(label="Paste", command=paste_text)
    context_menu.add_separator()
    context_menu.add_command(label="Cancel", command=cancel_context_menu)

    advanced_html_editor.bind("<Button-3>", on_right_click)

    def save_to_file():
        file_path = filedialog.asksaveasfilename(defaultextension=".html", filetypes=[("HTML Files", "*.html")])
        if not file_path:
            return

        with open(file_path, "w") as file:
            html_content = advanced_html_editor.get("1.0", tk.END)
            file.write(html_content)

        messagebox.showinfo("Info", "Email saved successfully!")

    def export_to_main():
        content = advanced_html_editor.get("1.0", tk.END)
        html_editor.delete("1.0", tk.END)
        html_editor.insert("1.0", content)
        html_win.destroy()

    def preview_html():
        import webbrowser
        with open("temp_preview.html", "w") as file:
            html_content = advanced_html_editor.get("1.0", tk.END)
            file.write(html_content)
        webbrowser.open("temp_preview.html")

    edit_menu = Menu(menu_bar, tearoff=0)
    menu_bar.add_cascade(label="Edit", menu=edit_menu)
    edit_menu.add_command(label="Load", command=load_email_from_html)
    edit_menu.add_command(label="Save", command=save_email_to_html)
    edit_menu.add_separator()
    edit_menu.add_command(label="Cut", command=cut_text)
    edit_menu.add_command(label="Copy", command=copy_text)
    edit_menu.add_command(label="Paste", command=paste_text)
    edit_menu.add_separator()
    edit_menu.add_command(label="Undo", command=undo)
    edit_menu.add_command(label="Redo", command=redo)
    edit_menu.add_separator()
    edit_menu.add_command(label="Find", command=find_text)
    edit_menu.add_separator()
    edit_menu.add_command(label="Validate HTML", command=validate_html)

    save_button = tk.Button(html_win, text="Save", command=save_to_file)
    save_button.pack(pady=5)

    export_button = tk.Button(html_win, text="Export", command=export_to_main)
    export_button.pack(pady=5)

    preview_button = tk.Button(html_win, text="Preview", command=preview_html)
    preview_button.pack(pady=5)


def close_error_dialog():
    global error_dialog
    if error_dialog:
        error_dialog.destroy()

def display_error(message):
    global error_dialog
    error_dialog = tk.Toplevel()
    error_dialog.title("Error")

    error_label = tk.Label(error_dialog, text=message, padx=20, pady=20)
    error_label.pack()

    close_button = tk.Button(error_dialog, text="Close", command=close_error_dialog)
    close_button.pack()

    error_dialog.protocol("WM_DELETE_WINDOW", close_error_dialog)


    error_dialog.transient(root)
    error_dialog.grab_set()
    root.wait_window(error_dialog)

def close_error_dialog():
    global error_dialog
    if error_dialog:
        error_dialog.destroy()



def load_configuration():
    file_path = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
    if not file_path or isinstance(file_path, tuple):
        return

    with open(file_path, 'r') as config_file:
        configuration = json.load(config_file)

        smtp_server.set(configuration.get("smtp_server", ""))
        port_entry.set(configuration.get("port", "465"))
        email_user.set(configuration.get("email_user", ""))
        email_pass.set(configuration.get("email_pass", ""))
        subject_entry.delete(0, tk.END)
        subject_entry.insert(0, configuration.get("subject", ""))
        use_local_smtp.set(configuration.get("use_local_smtp", False))

        messagebox.showinfo("Info", "Configuration loaded successfully!")


def clear_fields():
    smtp_server.set("")
    port_entry.set("")
    email_user.set("")
    email_pass.set("")
    subject_entry.delete(0, tk.END)
    html_editor.delete("1.0", tk.END)
    email_dict.clear()
    email_listbox.delete(0, tk.END)

def save_email_to_html():
    file_path = filedialog.asksaveasfilename(defaultextension=".html", filetypes=[("HTML Files", "*.html")])
    if not file_path:
        return  # User canceled the save dialog

    with open(file_path, "w") as file:
        html_content = html_editor.get("1.0", tk.END)
        file.write(html_content)

    messagebox.showinfo("Info", "Email saved successfully!")


def load_email_from_html():
    file_path = filedialog.askopenfilename(filetypes=[("HTML Files", "*.html")])
    if not file_path or isinstance(file_path, tuple):  # Checking for Cancel or invalid input
        return
    with open(file_path, 'r') as file:
        html_content = file.read()
        html_editor.delete("1.0", tk.END)
        html_editor.insert("1.0", html_content)

def save_configuration():
    configuration = {
        "smtp_server": smtp_server.get(),
        "port": port_entry.get(),
        "email_user": email_user.get(),
        "email_pass": email_pass.get(),
        "subject": subject_entry.get(),
        "use_local_smtp": use_local_smtp.get(),
    }

    file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON Files", "*.json")])
    if not file_path:
        return

    with open(file_path, "w") as config_file:
        json.dump(configuration, config_file)

    messagebox.showinfo("Info", "Configuration saved successfully!")

def start_local_smtp_server():
    global local_smtp_server
    local_smtp_server = subprocess.Popen(["python", "-m", "smtpd", "-n", "-c", "DebuggingServer", "localhost:1025"])


def stop_local_smtp_server():
    global local_smtp_server
    if local_smtp_server:
        local_smtp_server.terminate()
        local_smtp_server = None



        smtp_server.set("")
        port_entry.set("")
        server_entry.delete(0, tk.END)
        port_entry_entry.delete(0, tk.END)



def is_valid_email(email):
    # Use a regular expression to check if the email has a valid format
    pattern = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return re.match(pattern, email) is not None


def load_email_list():
    global emails_loaded
    global email_dict

    file_path = filedialog.askopenfilename()
    if not file_path or isinstance(file_path, tuple):
        return

    valid_emails = []
    invalid_emails = []

    if file_path.endswith('.csv'):
        valid_emails, invalid_emails = find_emails_in_csv(file_path)
    else:
        with open(file_path, 'r') as file:
            for line in file:
                email = line.strip()
                if is_valid_email(email):
                    valid_emails.append(email)
                    email_dict[email] = ""
                else:
                    invalid_emails.append(email)

    if valid_emails:
        update_email_listbox()
        emails_loaded = True

    if invalid_emails:
        messagebox.showwarning("Invalid Emails", "The list provided contains invalid email format")

def find_emails_in_csv(file_path):
    valid_emails = []
    invalid_emails = []

    with open(file_path, 'r') as file:
        csv_reader = csv.reader(file)
        for row in csv_reader:
            for item in row:
                emails = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b', item)
                for email in emails:
                    if is_valid_email(email):
                        valid_emails.append(email)
                        email_dict[email] = ""
                    else:
                        invalid_emails.append(email)

    return valid_emails, invalid_emails



def stop_sending():
    global should_stop
    should_stop = True
    root.after(0, lambda: log_display.insert(tk.END, "Email sending process stopped by the user.\n"))

def test_smtp_connection():
    if use_local_smtp.get():
    
        try:
            server = smtplib.SMTP(smtp_server.get())
            server.quit()
            messagebox.showinfo("Success", "SMTP connection successful!")
        except Exception as e:
            display_error(f"SMTP Connection Error: {str(e)}")
    else:
       
        if not all((smtp_server.get(), port_entry.get(), email_user.get(), email_pass.get())):
            display_error("Please fill in all fields.")
            return

        try:
            server = smtplib.SMTP_SSL(smtp_server.get(), int(port_entry.get()))
            server.quit()
            messagebox.showinfo("Success", "SMTP connection successful!")
        except Exception as e:
            display_error(f"SMTP Connection Error: {str(e)}")





def send_emails():
    def threaded_send():

        global should_stop
        should_stop = False



        if not emails_loaded:
            messagebox.showwarning("Warning", "No emails loaded.")
            return

        if not use_local_smtp.get():
            if not smtp_server.get() or not email_user.get() or not email_pass.get():
                messagebox.showwarning("Error", "SMTP Server, Email User, and Email Password are required fields.")
                return
        else:
            if not smtp_server.get():
                messagebox.showwarning("Error", "SMTP Server is a required field.")
                return

      
        local_email_dict = {}
        email_list = list(email_dict.keys())

        if not email_list:
            messagebox.showwarning("Warning", "No emails loaded.")
            return

        total_emails = len(email_list)
        if no_wait_var.get():
            delay_between_emails = 0
        else:
            delay_between_emails = (24 * 60 * 60) / total_emails

        try:
            print("Sending emails...")
            port = int(port_entry.get())

            if not is_server_reachable(smtp_server.get(), port):
                display_error("SMTP Server is not reachable.")
                return

            if use_local_smtp.get():
                server = smtplib.SMTP(smtp_server.get(), port)
            else:
                server = smtplib.SMTP_SSL(smtp_server.get(), port)
                server.login(email_user.get(), email_pass.get())

            for recipient in email_list:
                if should_stop:  
                    print("Email sending process halted by user.")
                    should_stop = False  
                    break
                msg = MIMEMultipart("alternative")
                msg['From'] = email_user.get()
                msg['To'] = recipient
                msg['Subject'] = subject_entry.get()

                html_content = html_editor.get("1.0", tk.END)
                html_part = MIMEText(html_content, "html")
                msg.attach(html_part)

                try:
                    server.sendmail(email_user.get(), recipient, msg.as_string())
                    local_email_dict[recipient] = "Sent"
                    print(f"Sent email to {recipient}")
                except:
                    local_email_dict[recipient] = "Failed"
                    print(f"Failed to send email to {recipient}")

               
                with email_dict_lock:
                    email_dict[recipient] = local_email_dict[recipient]
                update_email_listbox()

                time.sleep(delay_between_emails)

            server.quit()
            print("Emails sent successfully.")
        except smtplib.SMTPAuthenticationError:
            display_error("SMTP Authentication Error: Please check your username and password.")
        except smtplib.SMTPException as e:
            print(f"Failure .. {str(e)}")
            display_error(f"SMTP Error: {str(e)}")

        except Exception as e:
            print(f"Failure .. {str(e)}")
            display_error(f"Error: {str(e)}")
        finally:
        
            with email_dict_lock:
                email_dict.update(local_email_dict)
            update_email_listbox()

    t = threading.Thread(target=threaded_send)
    t.start()


email_dict_lock = threading.Lock()


def is_server_reachable(server, port):
    try:
        socket.create_connection((server, port), timeout=10)
        return True
    except (ConnectionRefusedError, socket.timeout):
        return False


def clear_email_list():
    email_dict.clear()
    email_listbox.delete(0, tk.END)


def about_dialog():
    def copy_crypto_address():
        crypto_text.clipboard_clear()
        crypto_text.clipboard_append(crypto_text.selection_get())

    def on_right_click(event):
        context_menu.post(event.x_root, event.y_root)

    about_win = tk.Toplevel()
    about_win.title("About EmailKiller")


    context_menu = tk.Menu(about_win, tearoff=0)
    context_menu.add_command(label="Copy", command=copy_crypto_address)
    # Header content
    header_content = """
EmailKiller\n\nVersion 1.4.2\n\n
EmailKiller: Send up to 10k emails per day.\n
Takes 24 hours to send all emails to prevent spam detection.\n
Send from SMTP or local server.\n
Fully supports HTML.\n
Developed by Derek Johnston 2023\n\n
If you found this software helpful, consider donating:\n
    """

    header_label = tk.Label(about_win, text=header_content, padx=20, pady=20)
    header_label.pack()

    # Crypto addresses in a Text widget
    crypto_content = """
ETH: 0xB139a7f6A2398fd4F50BbaC9970da8BE57E6F539
BTC: bc1qeyuvfap99mx3r269htxm60qs04xuq4a9ahpjvt
    """
    crypto_text = tk.Text(about_win, height=3, width=70, wrap=tk.WORD, padx=20, pady=20)
    crypto_text.insert(tk.END, crypto_content)
    crypto_text.config(state=tk.DISABLED)  # Make it read-only
    crypto_text.pack()

    # Enable copy/paste on the Text widget
    crypto_text.bind("<Control-c>", lambda e: about_win.clipboard_append(crypto_text.selection_get()))
    crypto_text.bind("<Control-v>", lambda e: None)  # Optional: Disable pasting

    context_menu = tk.Menu(about_win, tearoff=0)
    context_menu.add_command(label="Copy", command=copy_crypto_address)

    crypto_text.bind("<Button-3>", on_right_click)

    close_button = tk.Button(about_win, text="Close", command=about_win.destroy)
    close_button.pack(pady=10)

    about_win.transient(root)
    about_win.grab_set()
    root.wait_window(about_win)


def help_dialog():
    help_win = tk.Toplevel()
    help_win.title("Help for EmailKiller")

    help_text = """    To use EmailKiller:
    1. Load or type in your SMTP configuration.
    2. Use 'Load Email List' to load the list of recipients.
    3. Load or type in the content of the HTML email.
    4. Press 'Send Emails' to start sending.
    
    To use EmailFinder:
    1. Enter each URL on new line. URLs must be formatted properly (https://website.com)
    2. Press 'Find Emails' to start search
    3. Emails found will be displayed in opposite column once complete
    4. Use Ctrl+C/Ctrl+V to copy and paste 
    

    Note: Make sure your SMTP settings are correct."""

    help_label = tk.Label(help_win, text=help_text, padx=20, pady=20, justify=tk.LEFT)
    help_label.pack()

    close_button = tk.Button(help_win, text="Close", command=help_win.destroy)
    close_button.pack(pady=10)

    help_win.transient(root)
    help_win.grab_set()
    root.wait_window(help_win)


def update_email_listbox():
    email_listbox.delete(0, tk.END)
    for email, status in email_dict.items():
        email_listbox.insert(tk.END, f"{email} - {status}")

def toggle_use_local_smtp():
    if use_local_smtp.get():
        smtp_server.set("localhost")
        port_entry.set("1025")
        start_local_smtp_server()
        messagebox.showinfo("Info", "Local SMTP server started on port 1025")
    else:
        stop_local_smtp_server()
        messagebox.showinfo("Info", "Local SMTP server stopped")


root = tk.Tk()
root.title("EmailKiller")
root.configure(bg='black')
menu_bar = Menu(root)
root.config(menu=menu_bar)
no_wait_var = tk.BooleanVar()
no_wait_var.set(False)  # Default is to wait

use_local_smtp = tk.BooleanVar()
use_local_smtp.set(False)  
local_smtp_server = None  

# Description
desc_label = tk.Label(root, text="EmailKiller\n\nSends up to 10k emails a day",
                      bg='black', fg='white', wraplength=500)
desc_label.pack(pady=15)

smtp_server = tk.StringVar(value="smtp.gmail.com")
port_entry = tk.StringVar(value="465")  # Default port number
email_user = tk.StringVar()
email_pass = tk.StringVar()
email_dict = {}

file_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="File", menu=file_menu)


file_menu.add_command(label="Save Configuration", command=save_configuration)
file_menu.add_command(label="Load Configuration", command=load_configuration)
file_menu.add_command(label="Save Email", command=save_email_to_html)

file_menu.add_command(label="Load Email", command=load_email_from_html)
file_menu.add_command(label="Load Email List", command=load_email_list)
file_menu.add_separator()
file_menu.add_command(label="Exit", command=root.quit)

options_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Options", menu=options_menu)
options_menu.add_command(label="Test Connection", command=test_smtp_connection)
options_menu.add_checkbutton(label="Use Local SMTP", variable=use_local_smtp, command=toggle_use_local_smtp)
options_menu.add_separator()
options_menu.add_checkbutton(label="Disable 24hr Rule", variable=no_wait_var)
options_menu.add_separator()

options_menu.add_command(label="Email Finder", command=open_email_finder_gui)
Edit_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Edit", menu=Edit_menu)
# Add this line
Edit_menu.add_command(label="HTML Editor", command=open_html_editor)

Edit_menu.add_command(label="Clear Fields", command=clear_fields)

# Help Menu


help_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Help", menu=help_menu)
help_menu.add_command(label="About", command=about_dialog)
help_menu.add_command(label="Help", command=help_dialog)


frame_email_list = tk.Frame(root, bg='white')
frame_email_list.pack(side=tk.LEFT, padx=10, pady=10, fill=tk.BOTH, expand=True)

scrollbar = tk.Scrollbar(frame_email_list, orient="vertical")
email_listbox = tk.Listbox(frame_email_list, yscrollcommand=scrollbar.set, width=25, bg='white', fg='black')
scrollbar.config(command=email_listbox.yview)
email_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

frame_controls = tk.Frame(root, bg='black')
frame_controls.pack(side=tk.RIGHT, padx=10, pady=10, fill=tk.BOTH)


# Label and Entry Widgets in the First Column
tk.Label(frame_controls, text="Server:", bg='black', fg='white').grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
server_entry = tk.Entry(frame_controls, textvariable=smtp_server, width=40)
server_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)

tk.Label(frame_controls, text="Port:", bg='black', fg='white').grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
port_entry_entry = tk.Entry(frame_controls, textvariable=port_entry, width=40)
port_entry_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)

tk.Label(frame_controls, text="User:", bg='black', fg='white').grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
tk.Entry(frame_controls, textvariable=email_user, width=40).grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)

tk.Label(frame_controls, text="Password:", bg='black', fg='white').grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
password_entry = tk.Entry(frame_controls, textvariable=email_pass, show='*', width=40)
password_entry.grid(row=3, column=1, padx=5, pady=5, sticky=tk.W)

email_label = tk.Label(frame_controls, text="Compose Email", bg='black', fg='white', font=("Arial", 12, "bold"))
email_label.grid(row=4, column=1, padx=5, pady=15, sticky=tk.W)


tk.Label(frame_controls, text="Subject:", bg='black', fg='white').grid(row=5, column=0, padx=5, pady=5, sticky=tk.W)
subject_entry = tk.Entry(frame_controls, width=40) # Declaration of subject_entry
subject_entry.grid(row=5, column=1, padx=5, pady=5, sticky=tk.W)

# HTML Editor (using scrolledtext)
html_editor = scrolledtext.ScrolledText(frame_controls, width=60, height=13, bg='white', fg='black')
html_editor.grid(row=6, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W + tk.E + tk.N + tk.S)

# Email Listbox
email_listbox.pack(fill=tk.BOTH, expand=True)

# Send Button in the Second Column
tk.Button(frame_controls, text="Load Email List", command=load_email_list).grid(row=8, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W + tk.E)


tk.Button(frame_controls, text="Send Emails", command=send_emails).grid(row=7, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W + tk.E)

error_label = tk.Label(root, text="", bg='black')  
error_label.pack(pady=10)

log_display = tk.Text(root, height=25, width=40, bg='black', fg='white', borderwidth=0, highlightthickness=0)
log_display.pack(pady=10)

import sys

class IORedirector(object):
    def __init__(self, widget):
        self.widget = widget

    def write(self, text):
        self.widget.insert(tk.END, text)
        self.widget.see(tk.END)  

    def flush(self):
        pass

sys.stdout = IORedirector(log_display)


copyright_label = tk.Label(root, text="© Derek Johnston 2023", bg='black', fg='white')
copyright_label.pack(side=tk.BOTTOM, pady=10)
stop_button = tk.Button(root, text="Stop Sending", command=stop_sending, bg='black', fg='white')
stop_button.pack(pady=10, before=log_display)

test_connection_button = tk.Button(frame_controls, text="Test SMTP Connection", command=test_smtp_connection)
test_connection_button.grid(row=9, column=1, columnspan=1, padx=5, pady=5, sticky=tk.W + tk.E)


clear_button = tk.Button(frame_controls, text="Clear Email List", command=clear_email_list)
clear_button.grid(row=9, column=0, columnspan=1, padx=5, pady=5, sticky=tk.W + tk.E)

root.mainloop()
