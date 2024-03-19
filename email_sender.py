import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import time
import os
from dotenv import load_dotenv

load_dotenv()


class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automated Email Sender")
        self.root.geometry("600x400")
        self.root.resizable(0, 0)

        self.sender_email_label = ttk.Label(root, text="Sender Email:")
        self.sender_email_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.EW)
        self.sender_email_entry = ttk.Entry(root)
        self.sender_email_entry.insert(0, os.getenv('SENDER_EMAIL'))
        self.sender_email_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)


        self.sender_password_label = ttk.Label(root, text="Sender Password:")
        self.sender_password_label.grid(row=1, column=0, padx=5, pady=5, sticky=tk.EW)
        self.sender_password_entry = ttk.Entry(root, show='*')
        self.sender_password_entry.insert(0, os.getenv('SENDER_PASSWORD'))
        self.sender_password_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.EW)

        self.subject_label = ttk.Label(root, text="Subject:")
        self.subject_label.grid(row=2, column=0, padx=5, pady=5, sticky=tk.EW)
        self.subject_entry = ttk.Entry(root)
        self.subject_entry.insert(0, os.getenv('SUBJECT'))
        self.subject_entry.grid(row=2, column=1, padx=5, pady=5, sticky=tk.EW)


        self.message_label = ttk.Label(root, text="Message:")
        self.message_label.grid(row=3, column=0, padx=5, pady=5, sticky=tk.EW)
        self.message_text = tk.Text(root,width=40, height=10)
        self.message_text.insert(tk.END, os.getenv("BODY"))
        self.message_text.grid(row=3, column=1, padx=5, pady=5, sticky=tk.EW)

        self.attachment_file_paths = []

        self.recipients_label = ttk.Label(root, text="Attachment List:")
        self.recipients_label.grid(row=4, column=0, padx=5, pady=5, sticky=tk.EW)
        self.recipients_button = ttk.Button(root, text="Choose File", command=self.choose_file)
        self.recipients_button.grid(row=4, column=1, padx=5, pady=5, sticky=tk.EW)

        self.schedule_label = ttk.Label(root, text="Schedule Time(YYYY-MM-DD HH:MM):")
        self.schedule_label.grid(row=5, column=0, padx=5, pady=5, sticky=tk.EW)
        self.schedule_entry = ttk.Entry(root)
        self.schedule_entry.grid(row=5, column=1, padx=5, pady=5, sticky=tk.EW)


        self.send_email_button = ttk.Button(root, text="Send Emails", command=self.send_emails)
        self.send_email_button.grid(row=6, columnspan=2, padx=5, pady=5)



    def choose_file(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("All files", "*.*")])

        if file_paths:
            self.attachment_file_paths.extend(file_paths)

            file_names = [os.path.basename(path) for path in file_paths]
            self.recipients_button.config(text=f"Selected Files: {', '.join(file_names)}")

        self.attachment_file_paths = file_paths
        

    def send_emails(self):
        sender_email = self.sender_email_entry.get()
        sender_password = self.sender_password_entry.get()
        subject = self.subject_entry.get()
        message = self.message_text.get("1.0", "end-1c")
        schedule_time_str = self.schedule_entry.get()

        if not (sender_email and sender_password and subject and  message and schedule_time_str):
            messagebox.showerror("Error", "Please fill in all fields.")
            return
        
        try:
            schedule_time = datetime.strptime(schedule_time_str, "%Y-%m-%d %H:%M")
        except ValueError:
            messagebox.showerror("Error", "Invalid schedule time formate. Please use YYYY-MM-DD HH:MM.")
            return
        
        recipients_str = os.getenv('TO_EMAILS')
        recipients = recipients_str.split(',')

        while datetime.now() < schedule_time:
            time.sleep(60)

        for recipient in recipients:
            self.send_email(sender_email, sender_password, recipient.strip(), subject, message)
            messagebox.showinfo("Success", "Email sent successfully.")
            self.root.destroy()
            return

    def send_email(self, sender_email, sender_password, recipient, subject, message):
        smtp_server = 'smtp.gmail.com'
        smtp_port = 587

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient
        msg['Subject'] = subject
        msg.attach(MIMEText(message, 'plain'))

        if hasattr(self, 'attachment_file_paths') and self.attachment_file_paths:
            for attachment_path in self.attachment_file_paths:
                try:
                    with open(attachment_path, 'rb') as attachment:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f"attachment; filename={os.path.basename(attachment_path)}")
                    msg.attach(part)
                except FileNotFoundError:
                    print(f"File not found: {attachment_path}")
        else:
           print("No attachments selected")
        
        with smtplib.SMTP(smtp_server, smtp_port) as smtp:
            smtp.starttls()
            smtp.login(sender_email, sender_password)
            smtp.send_message(msg)


if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()