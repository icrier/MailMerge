# MailMerge for Microsoft Outlook

Mail merge with **To, CC, BCC and attachments**: something Outlook and Word can't do on their own.

Load your recipient list from Excel or CSV (or add people by hand), write one email template with placeholders like `{Name}`, attach files per person or for everyone, and let Outlook send a personalised email to each recipient.

**No installation required.** The tool is a PowerShell script with a simple Windows interface, and PowerShell is built into every Windows PC.

---

## Features

- **To, CC and BCC per recipient.** Each row can have its own CC/BCC. Separate multiple addresses with `;`.
- **Common CC / BCC.** Add addresses that are copied on every email (for example your manager or accounts team).
- **Per-recipient attachments.** Each person gets their own files (for example their own invoice).
- **Common attachments.** Files sent to everyone (for example terms and conditions).
- **Attachments folder.** Write only the file name in Excel (`INV-1001.pdf`) and the tool looks for it in the chosen folder.
- **Placeholders.** Any column in your sheet can be used in the subject, body or attachment name as `{ColumnName}`.
- **Three sending modes:**
  - *Preview*: each email opens in Outlook so you can check it and click Send yourself.
  - *Save to Drafts*: emails are saved to your Outlook Drafts folder.
  - *Send directly*: emails are sent automatically.
- **Safety checks.** Emails with a missing attachment are skipped (optional), and you are asked to confirm before sending directly.
- **Delay between emails** to avoid hitting mail server limits.
- **Your Outlook signature** is added automatically.
- **Log** of every email (sent / skipped / error) that can be saved as CSV.
- **Editable grid.** Fix any recipient detail directly in the table before sending.

---

## Requirements

- Windows 10 or 11
- Microsoft Outlook (desktop app), installed and signed in to your mail account
- Microsoft Excel, only if you want to load `.xlsx` files. Without Excel, save your sheet as CSV and load that instead.

---

## Getting started

1. Download **`MailMerge.zip`** from this repository and extract it to a folder.
2. If Windows blocks the files, right-click **`Start_MailMerge.bat`** and **`MailMerge.ps1`**, open **Properties**, tick **Unblock** and click **OK**.
3. Double-click **`Start_MailMerge.bat`** to open the tool.

### Files

| File | Purpose |
|---|---|
| `Start_MailMerge.bat` | Double-click to start the tool |
| `MailMerge.ps1` | The application (PowerShell) |
| `Recipients_Template.xlsx` | Sample recipient sheet to fill in |

---

## How to use

### 1. Recipients tab

Click **Excel / CSV Load karein** and choose your file, or fill in the **Manually recipient add karein** form and click **+ Add Recipient**.

Your sheet's first row must contain headers. Only **To** is required:

| To | CC | BCC | Name | InvoiceNo | Amount | DueDate | Attachments |
|---|---|---|---|---|---|---|---|
| rahul@example.com | accounts@example.com | | Rahul Sharma | INV-1001 | Rs 12,500 | 15-10-2026 | C:\Invoices\INV-1001.pdf |
| priya@example.com | manager@example.com; team@example.com | | Priya Verma | INV-1002 | Rs 8,200 | 15-10-2026 | INV-1002.pdf; Terms.pdf |

- Separate multiple email addresses or files with `;`.
- Use a full path for attachments, or just the file name together with the **Attachments folder** setting.
- You can add any extra columns you like (City, Phone, etc.) and use them as placeholders.
- A column named `Email` is accepted in place of `To`.

### 2. Email Content tab

- **Subject**, for example `Invoice {InvoiceNo} - {Name}`
- **Common CC / BCC**, added to every email
- **Attachments folder**, where file names from the sheet are looked up
- **Common attachments**, sent with every email
- **Body**, for example:

  ```
  Dear {Name},

  Please find attached invoice {InvoiceNo} for amount {Amount}, due on {DueDate}.

  Regards,
  ```

  Plain text works fine. HTML (`<b>`, `<br>`, etc.) is also supported.

### 3. Preview & Send tab

1. Click **Pehle record ka Preview** to check the first email's To, CC, subject and attachments.
2. Choose a mode: **Preview**, **Drafts** or **Send**.
3. Click **Mail Merge Start**. Use **Stop** to cancel at any time.
4. Click **Log save karein** to save the results as CSV.

> **Tip:** For the first run, put your own email address in the To column and use **Preview** mode.

---

## Troubleshooting

| Problem | Solution |
|---|---|
| Nothing happens when double-clicking the `.bat` file | Unblock both files (see Getting started step 2). Make sure both files are in the same folder. |
| "Outlook start nahi hua" | Open Outlook first and make sure it is signed in. Don't run the tool as Administrator if Outlook runs normally (or vice versa). |
| "Is PC me MS Excel nahi mila" | Open the sheet, use **Save As → CSV**, and load the CSV file. |
| Outlook shows "A program is trying to send email" | Click **Allow**. This is an Outlook security prompt and depends on your antivirus / organisation settings. |
| Email skipped with "missing" in the log | The attachment path is wrong or the file doesn't exist. Check the path or the Attachments folder setting. |
| Dates appear as numbers | Format the date column in Excel as a date before loading, or type the date as text. |
| Some emails not sent when sending many | Increase the delay between emails. Many mail servers limit how many emails you can send per minute or per day. |

---

## Notes

- Emails are sent from your default Outlook account, using your default signature.
- The tool does not store your data or send it anywhere other than through your own Outlook.
- Please follow your organisation's email policy and anti-spam rules when sending bulk email.

---

## License

Free to use and modify.
