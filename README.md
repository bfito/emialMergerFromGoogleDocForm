```markdown
# 📄 Personalized Document Generator with Google Apps Script

This Google Apps Script project generates personalized Google Docs for student applicants using data from a Google Sheet and a linked template Google Doc. It's ideal for charter schools or other organizations managing large-scale enrollment or onboarding via templated communication.

---

## ✅ Features

- Reads student and parent data from a Google Sheet
- Uses a reusable Google Doc template with placeholders (e.g., `{{StudentFirstName}}`)
- Automatically replaces placeholders with real data
- Bolds all inserted values in the generated documents
- Creates uniquely named documents like: `Jane Doe - Enrollment Email`
- Returns a clickable link to each generated Doc back into the sheet

---

## 🛠 Setup Instructions

### 1. Spreadsheet Preparation

- Open your Google Sheet
- Row 1 should be your header with these (or similar) column names:

```

StudentFirstName | StudentLastName | ParentFirstName | ParentLastName | StudentGradeFall2025

```

- Fill in the data starting in row 2.

---

### 2. Create Your Google Doc Template

- Open Google Drive → New → Google Doc
- Paste your letter or email body
- Use placeholders where you'd like dynamic content inserted, such as:

```

Dear {{ParentFirstName}} {{ParentLastName}},

We are pleased to offer {{StudentFirstName}} {{StudentLastName}} a seat in {{StudentGradeFall2025}} grade for the 2025–26 school year.

Please complete the Enrollment Packet 2025.

```

- Optionally:
  - Add a hyperlink to “Enrollment Packet 2025” manually via **Insert > Link**
  - Format your template with bold, italic, or colored text as needed
- Save the document and **copy the ID** from the URL:
```

[https://docs.google.com/document/d/THIS\_PART\_IS\_THE\_ID/edit](https://docs.google.com/document/d/THIS_PART_IS_THE_ID/edit)

````

---

### 3. Paste Script into Apps Script

- In your spreadsheet, go to **Extensions > Apps Script**
- Delete any placeholder code
- Paste the contents of `Code.gs` from this repo
- Replace the placeholder in the script:
```js
const templateId = "PASTE_YOUR_TEMPLATE_DOC_ID_HERE";
````

---

### 4. Run the Script

* From the script editor, select `generateCustomEmails` and click ▶️ Run
* The first time, Google will ask for authorization — accept it
* Each row will generate a new document
* The Google Doc link will appear in the next empty column

---

## 🧠 How It Works

* The script makes a copy of your template for each row
* It replaces placeholders in `{{HeaderName}}` format with actual values
* It bolds each inserted value for emphasis
* A unique filename is generated to avoid overwriting
* Each document is saved in your Google Drive

---

## 📌 Future Ideas

* Automatically email the documents as PDFs
* Save all files into a specific Google Drive folder
* Add a Sheet button to run the script without opening the editor
* Add error handling/logging

---

## 📎 Example Output

For input data like:

| StudentFirstName | StudentLastName | ParentFirstName | ParentLastName | StudentGradeFall2025 |
| ---------------- | --------------- | --------------- | -------------- | -------------------- |
| Ava              | Johnson         | Sam             | Johnson        | 3rd                  |

The script will create a document titled:

```
Ava Johnson - Enrollment Email
```

With a letter like:

```
Dear Sam Johnson,

We are pleased to offer Ava Johnson a seat in 3rd grade for the 2025–26 school year.
```

---
NOTE: The original letter and verbage was more extensive and warmer. Eg is for exmaple sake.
## License

MIT License

```

---

```
