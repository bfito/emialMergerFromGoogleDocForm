/**
 * Main function to generate personalized documents for each student row
 * using a shared template Google Doc and inserting spreadsheet data.
 */
function generateCustomEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0]; // First row = column headers

  // 🔗 Replace this with your actual Google Doc template ID
  const templateId = "PASTE_YOUR_TEMPLATE_DOC_ID_HERE";

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // 🔠 Build student full name for document title
    const studentFirst = row[headers.indexOf("StudentFirstName")];
    const studentLast = row[headers.indexOf("StudentLastName")];
    const studentName = `${studentFirst} ${studentLast}`;
    const baseDocName = `${studentName} - Enrollment Email`;
    const docName = generateUniqueDocName(baseDocName);

    // 📄 Make a copy of the template document
    const templateFile = DriveApp.getFileById(templateId);
    const newFile = templateFile.makeCopy(docName);
    const doc = DocumentApp.openById(newFile.getId());
    const body = doc.getBody();

    // 🔁 Replace all placeholders like {{ParentFirstName}}, etc.
    headers.forEach((header, index) => {
      const placeholder = `{{${header}}}`;
      const value = row[index] || '';
      body.replaceText(placeholder, value);

      // ✨ Optional: Bold the inserted values
      const found = body.findText(value);
      if (found) {
        const element = found.getElement();
        if (element.editAsText) {
          const text = element.asText();
          const start = found.getStartOffset();
          const end = found.getEndOffsetInclusive();
          text.setBold(start, end, true);
        }
      }
    });

    // 💾 Save and close the generated document
    doc.saveAndClose();

    // 🔗 Write the generated Doc URL to the next empty column in the sheet
    const lastCol = sheet.getLastColumn();
    sheet.getRange(i + 1, lastCol + 1).setValue(doc.getUrl());
  }
}

/**
 * Ensure each generated document has a unique name.
 * Appends (1), (2), etc., if a doc with the same name already exists.
 */
function generateUniqueDocName(baseName) {
  let name = baseName;
  let counter = 1;
  let files = DriveApp.getFilesByName(name);

  while (files.hasNext()) {
    name = `${baseName} (${counter})`;
    files = DriveApp.getFilesByName(name);
    counter++;
  }

  return name;
}
