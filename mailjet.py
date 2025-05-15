import os
import base64
import logging
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import pandas as pd
import re
import json
from mailjet_rest import Client

CONFIG_FILE = "config.json"


class MailSenderApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Certificate Email Sender - Mailjet GUI")
        self.master.geometry("760x650")
        self.master.resizable(False, False)

        self.config = {
            "file_type": "excel",
            "disable_attachments": False,
            "from_email": "",
            "mailjet_api_key": "",
            "mailjet_api_secret": "",
            "file_path": "",
            "attachments_folder": ""
        }

        self.setup_widgets()
        self.load_config()
        self.master.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_widgets(self):
        def label(text, row):
            tk.Label(self.master, text=text).grid(
                row=row, column=0, sticky="w", padx=10, pady=5)

        def entry(name, row, show=None):
            label(name, row)
            ent = tk.Entry(self.master, width=60, show=show)
            ent.grid(row=row, column=1, padx=10, pady=5)
            return ent

        self.from_email_entry = entry("From Email", 0)
        self.api_key_entry = entry("Mailjet API Key", 1, show="*")
        self.api_secret_entry = entry("Mailjet API Secret", 2, show="*")

        label("File Type (excel/csv)", 3)
        self.file_type_combo = ttk.Combobox(
            self.master, values=["excel", "csv"])
        self.file_type_combo.set("excel")
        self.file_type_combo.grid(row=3, column=1, padx=10, pady=5)

        label("Data File", 4)
        self.file_path_entry = tk.Entry(self.master, width=50)
        self.file_path_entry.grid(row=4, column=1, padx=10, pady=5)
        tk.Button(self.master, text="Browse", command=self.browse_file).grid(
            row=4, column=2, padx=5)

        label("Attachments Folder", 5)
        self.attachments_entry = tk.Entry(self.master, width=50)
        self.attachments_entry.grid(row=5, column=1, padx=10, pady=5)
        tk.Button(self.master, text="Browse", command=self.browse_folder).grid(
            row=5, column=2, padx=5)

        self.disable_attachments_var = tk.BooleanVar()
        tk.Checkbutton(self.master, text="Disable Attachments", variable=self.disable_attachments_var).grid(
            row=6, column=1, sticky="w", padx=10
        )

        tk.Button(self.master, text="Start Sending Emails", command=self.send_emails, bg="green", fg="white").grid(
            row=7, column=1, pady=10
        )

        self.progress = ttk.Progressbar(
            self.master, orient="horizontal", length=600, mode="determinate")
        self.progress.grid(row=8, column=0, columnspan=3, pady=10)

        label("Logs", 9)
        self.log_area = scrolledtext.ScrolledText(
            self.master, height=15, width=90, state="disabled")
        self.log_area.grid(row=10, column=0, columnspan=3, padx=10, pady=5)

    def log(self, message):
        self.log_area.configure(state="normal")
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.configure(state="disabled")
        self.log_area.see(tk.END)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
        if file_path:
            self.file_path_entry.delete(0, tk.END)
            self.file_path_entry.insert(0, file_path)

    def browse_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.attachments_entry.delete(0, tk.END)
            self.attachments_entry.insert(0, folder_path)

    def send_emails(self):
        self.config.update({
            "from_email": self.from_email_entry.get(),
            "mailjet_api_key": self.api_key_entry.get(),
            "mailjet_api_secret": self.api_secret_entry.get(),
            "file_type": self.file_type_combo.get().lower(),
            "file_path": self.file_path_entry.get(),
            "attachments_folder": self.attachments_entry.get(),
            "disable_attachments": self.disable_attachments_var.get()
        })
        self.save_config()

        if not self.validate_config():
            return

        try:
            df = pd.read_excel(self.config["file_path"]) if self.config["file_type"] == "excel" else pd.read_csv(
                self.config["file_path"])
        except Exception as e:
            self.log(f"❌ Error reading file: {e}")
            return

        required_columns = {"full_name", "email", "cert_no"}
        if not required_columns.issubset(df.columns):
            self.log(
                f"❌ Missing columns: {required_columns - set(df.columns)}")
            return

        self.progress["maximum"] = len(df)
        self.progress["value"] = 0

        mailjet = Client(auth=(
            self.config["mailjet_api_key"], self.config["mailjet_api_secret"]), version='v3.1')
        summary = {"sent": 0, "failed": 0}

        def process_row(index):
            if index >= len(df):
                self.finalize_summary(mailjet, summary, len(df))
                return

            row = df.iloc[index]
            full_name = str(row.get("full_name", "Participant"))
            to_email = str(row.get("email", "")).strip()
            cert_no = str(row.get("cert_no", "")).strip()

            email_regex = r'^[\w\.-]+@[\w\.-]+\.\w+$'
            if not re.match(email_regex, to_email):
                self.log(f"❌ Invalid email for {full_name}: {to_email}")
                summary["failed"] += 1
            else:
                msg_data = {
                    "From": {"Email": self.config["from_email"], "Name": "NIELIT Chandigarh"},
                    "To": [{"Email": to_email, "Name": full_name}],
                    "Subject": "Congratulations! Your Certificate is Ready",
                    "HTMLPart": f"""<html><body><p>Dear {full_name},</p>
                                    <p>Your certificate is attached.</p>
                                    <p>Best regards,<br/>NIELIT Chandigarh</p></body></html>"""
                }

                attachment_path = os.path.join(
                    self.config["attachments_folder"], f"{cert_no}.pdf")
                if not self.config["disable_attachments"] and os.path.exists(attachment_path):
                    try:
                        with open(attachment_path, 'rb') as f:
                            content = base64.b64encode(f.read()).decode()
                        msg_data["Attachments"] = [{
                            "ContentType": "application/pdf",
                            "Filename": f"{cert_no}.pdf",
                            "Base64Content": content
                        }]
                    except Exception as e:
                        self.log(f"⚠️ Couldn't attach {cert_no}.pdf: {e}")

                try:
                    response = mailjet.send.create(
                        data={"Messages": [msg_data]})
                    if response.status_code == 200:
                        self.log(f"✅ Sent to {to_email}")
                        summary["sent"] += 1
                    else:
                        self.log(
                            f"❌ Failed to send to {to_email}: {response.status_code} {response.text}")
                        summary["failed"] += 1
                except Exception as e:
                    self.log(f"❌ Exception sending to {to_email}: {e}")
                    summary["failed"] += 1

            self.progress["value"] = index + 1
            self.master.update_idletasks()
            self.master.after(100, process_row, index + 1)

        process_row(0)

    def finalize_summary(self, mailjet, summary, total):
        msg = f"Total: {total}\nSent: {summary['sent']}\nFailed: {summary['failed']}"
        self.log("📬 Summary:\n" + msg)
        try:
            summary_msg = {
                "Messages": [{
                    "From": {"Email": self.config["from_email"], "Name": "NIELIT Chandigarh"},
                    "To": [{"Email": self.config["from_email"], "Name": "Administrator"}],
                    "Subject": "Email Sending Summary",
                    "TextPart": msg
                }]
            }
            mailjet.send.create(data=summary_msg)
        except Exception as e:
            self.log(f"⚠️ Summary email failed: {e}")

        messagebox.showinfo("Summary", msg)

    def validate_config(self):
        required = ["from_email", "mailjet_api_key",
                    "mailjet_api_secret", "file_path", "attachments_folder"]
        for field in required:
            if not self.config[field]:
                self.log(f"❌ {field.replace('_', ' ').title()} is required.")
                return False
        return True

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    data = json.load(f)
                    self.config.update(data)
                    self.from_email_entry.insert(
                        0, self.config.get("from_email", ""))
                    self.api_key_entry.insert(
                        0, self.config.get("mailjet_api_key", ""))
                    self.api_secret_entry.insert(
                        0, self.config.get("mailjet_api_secret", ""))
                    self.file_type_combo.set(
                        self.config.get("file_type", "excel"))
                    self.file_path_entry.insert(
                        0, self.config.get("file_path", ""))
                    self.attachments_entry.insert(
                        0, self.config.get("attachments_folder", ""))
                    self.disable_attachments_var.set(
                        self.config.get("disable_attachments", False))
            except Exception as e:
                self.log(f"⚠️ Failed to load config: {e}")

    def save_config(self):
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(self.config, f, indent=4)
        except Exception as e:
            self.log(f"⚠️ Failed to save config: {e}")

    def on_close(self):
        self.save_config()
        self.master.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = MailSenderApp(root)
    root.mainloop()
