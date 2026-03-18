---

# 🤖 Intelligent Outlook Attachment Downloader

This automation robot streamlines the process of downloading specific attachments from your Outlook Inbox by automatically filtering, extracting, and organizing them into date-stamped folders.

---

## ⚙️ How to Customize
Before running or compiling the script, locate the **"SEARCH SETTINGS"** block in the Python code to tailor the behavior:

* **SEARCH_TERM**: The specific company name or keyword in the subject line (e.g., "Invoice", "Report", "Vendor X").
* **TARGET_EXTENSION**: The specific file type to extract (e.g., ".pdf", ".xlsx", ".zip").
* **MESSAGE_LIMIT**: Defines how many recent emails the robot should scan before stopping.

---

## 🚀 Generating the Executable (.exe)
To ensure the robot functions correctly on corporate machines with antivirus or network restrictions, use the following compilation method:

1.  Open your terminal or command prompt in the project folder.
2.  Execute the command below:
    
    ```bash
    python -m PyInstaller --onedir --nowindowed --hidden-import="win32timezone" --collect-submodules="win32com" --name "Outlook_Downloader" Outlook_Downloader.py
    ```

3.  Once finished, provide the user with the entire folder found in `dist/Outlook_Downloader`.

---

## ⚠️ Requirements & Security
* **Outlook Desktop**: Classic Outlook must be installed, configured, and open.
* **Security Prompt**: Outlook may display a security alert asking for permission to access your mailbox. Select **"Allow access for 10 minutes"** and confirm.
* **Optimal Location**: Always run the application from a local drive (C: Drive or Desktop). Avoid running the `.exe` directly from USB flash drives or network mapped drives to prevent permission errors.

---
