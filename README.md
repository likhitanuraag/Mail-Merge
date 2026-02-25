# PyQt5 Rich-Text Mail Merge Utility ✉️

A robust, GUI-driven desktop application for sending bulk customized emails via Microsoft Outlook. This tool uses **Microsoft Word (.docx)** files as rich-text templates and **Microsoft Excel (.xlsx)** files as the data source, preserving your formatting, tables, and styles perfectly in the final email.

## ✨ Features

* **Rich-Text HTML Support:** Uses `mammoth` to convert Word documents to clean HTML, preserving bold, italics, lists, and tables.
* **Smart Mapping Interface:** Easily map `{{placeholders}}` in your Word document to columns in your Excel sheet.
* **Dynamic Routing:** Select Excel columns for **To**, **CC**, and **BCC** fields.
* **Dynamic Subject Lines:** Use `{{placeholders}}` directly in the subject line.
* **Live HTML Preview:** Review the exact HTML rendering and placeholder replacements before sending.
* **Save/Load Configurations:** Save your column mappings and settings to a `.json` file to run recurring jobs instantly.
* **Resilient Processing:** If a single email fails (e.g., bad email address), the app logs the error and continues processing the rest of the batch, providing a detailed summary at the end.
* **Built-in Sample Generator:** First-time users can generate sample Word and Excel files directly from the "Help" menu to test the application.
* **Asynchronous Execution:** Uses PyQt5 `QThread` to send emails in the background, keeping the UI responsive and preventing freezing.

---

## 🛠️ Prerequisites

To run this application, you must have the following installed on your machine:
* **Python 3.7+**
* **Microsoft Outlook** (A local, desktop installation is required as this script uses the Outlook COM API).

## 📦 Installation

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/yourusername/pyqt5-mail-merge.git](https://github.com/yourusername/pyqt5-mail-merge.git)
   cd pyqt5-mail-merge