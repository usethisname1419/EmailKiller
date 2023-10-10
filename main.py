import smtplib
import time
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, Menu
import threading
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import json
import subprocess
import warnings

warnings.simplefilter("ignore", DeprecationWarning)

emails_loaded = False


error_dialog = None



import re

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


def open_html_editor():
    html_win = tk.Toplevel()
    html_win.title("HTML Editor")

    advanced_html_editor_frame = tk.Frame(html_win)
    advanced_html_editor_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    advanced_html_editor = HighlightText(advanced_html_editor_frame, wrap=tk.WORD)
    advanced_html_editor.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = tk.Scrollbar(advanced_html_editor_frame, command=advanced_html_editor.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    advanced_html_editor.config(yscrollcommand=scrollbar.set)

    current_content = html_editor.get("1.0", tk.END)
    advanced_html_editor.insert("1.0", current_content)

    def save_and_close():
        content = advanced_html_editor.get("1.0", tk.END)
        html_editor.delete("1.0", tk.END)
        html_editor.insert("1.0", content)
        html_win.destroy()

    save_button = tk.Button(html_win, text="Save", command=save_and_close)
    save_button.pack(pady=10)


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

    # Make the error dialog a modal window
    error_dialog.transient(root)
    error_dialog.grab_set()
    root.wait_window(error_dialog)



def load_configuration():
    file_path = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
    if not file_path or isinstance(file_path, tuple):  # Checking for Cancel or invalid input
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
    port_entry.set("465")  # Set the default port number
    email_user.set("")
    email_pass.set("")
    subject_entry.delete(0, tk.END)
    html_editor.delete("1.0", tk.END)


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
        return  # User canceled the save dialog

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


        # Clear the server and port fields when stopping the local SMTP server
        smtp_server.set("")
        port_entry.set("")
        server_entry.delete(0, tk.END)
        port_entry_entry.delete(0, tk.END)


def load_email_list():
    global emails_loaded
    file_path = filedialog.askopenfilename()
    if not file_path or isinstance(file_path, tuple):
        return
    with open(file_path, 'r') as file:
        for line in file:
            email_dict[line.strip()] = ""
    update_email_listbox()
    emails_loaded = True


def send_emails():
    def threaded_send():
        # Validate required fields

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

        # Create a separate email_dict for this thread
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
            if use_local_smtp.get():
                server = smtplib.SMTP(smtp_server.get(), port)
            else:
                server = smtplib.SMTP_SSL(smtp_server.get(), port)
                server.login(email_user.get(), email_pass.get())

            for recipient in email_list:
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

                # Update the email_dict with the status of the current email and refresh the listbox
                with email_dict_lock:
                    email_dict[recipient] = local_email_dict[recipient]
                update_email_listbox()

                time.sleep(delay_between_emails)

            server.quit()
            print("Emails sent successfully.")
        except smtplib.SMTPAuthenticationError:
            display_error("SMTP Authentication Error: Please check your username and password.")
        except smtplib.SMTPException as e:
            display_error(f"SMTP Error: {str(e)}")
        except Exception as e:
            display_error(f"Error: {str(e)}")
        finally:
            # Use a lock to update the global email_dict with the results from this thread
            with email_dict_lock:
                email_dict.update(local_email_dict)
            update_email_listbox()

    t = threading.Thread(target=threaded_send)
    t.start()

# Add a lock to protect access to email_dict
email_dict_lock = threading.Lock()





def clear_email_list():
    email_dict.clear()
    email_listbox.delete(0, tk.END)


def about_dialog():
    about_win = tk.Toplevel()
    about_win.title("About EmailKiller")

    about_content = """
EmailKiller\n\nVersion 1.0\n\n
EmailKiller: Send up to 10k emails per day.\n
Takes 24 hours to send all emails to prevent spam detection.\n
Send from SMTP or local server.\n
Fully supports HTML.\n
Developed by Derek Johnston 2023\n\n
If you found this software helpful, consider donating:\n
ETH: 0xB139a7f6A2398fd4F50BbaC9970da8BE57E6F539\n
BTC: bc1qeyuvfap99mx3r269htxm60qs04xuq4a9ahpjvt
    """

    about_label = tk.Label(about_win, text=about_content, padx=20, pady=20)
    about_label.pack()

    close_button = tk.Button(about_win, text="Close", command=about_win.destroy)
    close_button.pack(pady=10)

    about_win.transient(root)
    about_win.grab_set()
    root.wait_window(about_win)


def help_dialog():
    help_win = tk.Toplevel()
    help_win.title("Help for EmailKiller")

    help_text = """To use EmailKiller:
    1. Load or type in your SMTP configuration.
    2. Use 'Load Email List' to load the list of recipients.
    3. Load or type in the content of the HTML email.
    4. Press 'Send Emails' to start sending.

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
use_local_smtp.set(False)  # Set the initial value to False
local_smtp_server = None  # Store the local SMTP server process

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
file_menu.add_separator()
file_menu.add_command(label="Exit", command=root.quit)

options_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Options", menu=options_menu)
options_menu.add_checkbutton(label="Disable 24hr Rule", variable=no_wait_var)
options_menu.add_checkbutton(label="Use Local SMTP", variable=use_local_smtp, command=toggle_use_local_smtp)

options_menu.add_command(label="Clear Fields", command=clear_fields)
# Add this line
options_menu.add_command(label="HTML Editor", command=open_html_editor)


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

# Label and Entry Widgets in the Second Column
tk.Label(frame_controls, text="Subject:", bg='black', fg='white').grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)
subject_entry = tk.Entry(frame_controls, width=40)
subject_entry.grid(row=4, column=1, padx=5, pady=5, sticky=tk.W)

# HTML Editor (using scrolledtext)
html_editor = scrolledtext.ScrolledText(frame_controls, width=50, height=10, bg='white', fg='black')
html_editor.grid(row=5, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W + tk.E + tk.N + tk.S)

# Email Listbox
email_listbox.pack(fill=tk.BOTH, expand=True)

# Send Button in the Second Column
tk.Button(frame_controls, text="Load Email List", command=load_email_list).grid(row=7, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W + tk.E)

# Send Button in the Second Column
tk.Button(frame_controls, text="Send Emails", command=send_emails).grid(row=8, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W + tk.E)

error_label = tk.Label(root, text="", bg='black')  # Error label for displaying error messages in red.
error_label.pack(pady=10)

log_display = tk.Text(root, height=10, width=40, bg='black', fg='white', borderwidth=0, highlightthickness=0)
log_display.pack(pady=10)

import sys

class IORedirector(object):
    def __init__(self, widget):
        self.widget = widget

    def write(self, text):
        self.widget.insert(tk.END, text)
        self.widget.see(tk.END)  # Auto-scroll to the end

    def flush(self):
        pass

sys.stdout = IORedirector(log_display)

copyright_label = tk.Label(root, text="© Derek Johnston 2023", bg='black', fg='white')
copyright_label.pack(side=tk.BOTTOM, pady=10)

clear_button = tk.Button(frame_controls, text="Clear Email List", command=clear_email_list)
clear_button.grid(row=9, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W + tk.E)

root.mainloop()