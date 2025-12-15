# MergeRelay
Automated Bulk Email Distribution for Helpdesk &amp; IT Teams


### What is MergeRelay?

MergeRelay automates bulk personalized emails using a Word mail-merge template and Outlook. It reduces manual sending, supports multiple recipients, CC/BCC, multiple attachments, and can save messages as drafts.

**Key behaviors (from the embedded macro):**
- Prompts for the email subject.
- Option to preview the first email before processing.
- Option to send immediately or save all messages as drafts.
- Forces `From` using the `From` column (uses `SentOnBehalfOfName`).
- Supports multiple recipients, CC, BCC, and attachments separated by semicolons.
- Writes an error log in the project folder for failed records.

---

## Quick Start

1. Open `SETUP.docm` in Word and enable macros.
2. Connect an Excel data source to `SETUP.docm`:
   - Go to **Mailings** → **Select Recipients** → **Use an Existing List**.
3. Browse and select your Excel file  
   (a sample is provided as `sample_data.xlsx`).
4. Compose and format the email body and signature in the active Word document.
   - This content is copied directly into Outlook.
   - Use Word formatting exactly as you want it to appear.
   - The email subject is **not** part of the document; you will be prompted to enter it when the macro runs.
5. Run the macro `ExecuteMailingProcess`:
   - **Windows (Word 2010–365):**
     - Ribbon → View → Macros  
     - or Developer → Macros  
     - or shortcut `Alt + F8`
   - **Mac (modern Word):**
     - Ribbon → View → Macros  
     - or Tools → Macros  
     - or shortcut `Option + F8`

---

## Excel Sheet Format

This Excel file contains recipient data, attachments, and merge fields.

Prepare the sheet with at least these columns:

- **Required**
  - `Email`
  - `From`
- **Optional**
  - `CC`
  - `BCC`
  - `Attachment`
  - `Contact_Person`
  - `FirstName`
  - `Company`
  - Any other fields you want to merge

**Rules:**
- Use semicolons to separate multiple values  
  - Emails: `a@x.com;b@y.com`
  - Attachments: `C:\file1.pdf;C:\file2.docx`

> **Note:** You can use the structure of `sample_data.xlsx` directly.

---

## Merge Fields and Formatting

- Use Word merge fields such as:
  - `«Contact_Person»`
  - `«FirstName»`
  - `«Company»`
  - `«Email»`
  - `«CC»`
  - `«BCC»`
  - `«Attachment»`
- Field names **must match Excel column headers exactly**.
- MergeRelay processes each record, renders the merged document, and copies the content into Outlook.
- All formatting is controlled in Word, not Outlook.
- You can save your composed template at any time using `Ctrl + S` in `SETUP.docm`.

---

## Excel Columns (Summary)

- `Email` (required): recipient(s), semicolon-separated
- `From` (required): sender alias or address
- `CC` (optional): semicolon-separated addresses
- `BCC` (optional): semicolon-separated addresses
- `Attachment` (optional): semicolon-separated full file paths
- Personalization fields: `Contact_Person`, `FirstName`, `Company`, etc.

---

## Specific Notes

- `SETUP.docm` must be connected to a data source formatted like `sample_data.xlsx`.
- If the `From` column is provided, MergeRelay sets `SentOnBehalfOfName`.
  - You must have permission to send from that address.
- Attachments must be full file paths.
  - Missing files are logged and skipped for that record.

---

## System Requirements

- Windows
- Microsoft Word (2016+) and Outlook (2016+)
- Macros enabled in Word
- Word and Outlook must be signed into the same email account

---

## Troubleshooting (Common)

- **No active Word document found**  
  Open your Word mail-merge template.
- **No mail merge data source found**  
  Attach your Excel file via Mailings → Select Recipients.
- Invalid email addresses are logged and skipped.
- **Could not initialize Outlook**  
  Ensure Outlook is installed and running; restart if needed.
- Drafts created but not sent  
  Open Outlook Drafts and send manually.

---

## Security and Usage

- MergeRelay runs entirely locally.
- No external data transmission.
- Reads Word and Excel files only.
- Writes to Outlook and local error logs.
- Use only for legitimate, authorized communication.
- Review the macro code (`SETUP.docm` → `Alt + F11`) before use.

---

## Support

- Check `MailMerge_ErrorLog_*.txt` in the project folder for failure details.

---

## Version

- **MergeRelay v1.0.0**

Open `SETUP.docm` to begin.
