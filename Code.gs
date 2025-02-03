function myFunction(e) {
  if (!e || !e.range) {
    console.error("This function must be triggered by an edit in the spreadsheet.");
    return;
  }

  const range = e.range;  // The range of cells that was edited
  const sheet = range.getSheet();
  const editedRow = range.getRow();

  // Skip the header row and only process rows below it
  if (editedRow === 1) {
    console.log("Header row edited. Ignoring...");
    return;
  }

  // Ensure the edit was made in the correct sheet
  if (sheet.getName() !== "Sheet1") {
    console.log("Edit made in an untracked sheet. Ignoring...");
    return;
  }

  const url = "https://14f6-122-187-117-179.ngrok-free.app/trigger-email";  // Replace with your Flask API URL
  const options = {
    method: "POST",
    contentType: "application/json",
    payload: JSON.stringify({
      Name: sheet.getRange(editedRow, 1).getValue(),  // Assuming Name is in column 1
      "Email Adress": sheet.getRange(editedRow, 2).getValue()  // Assuming Email is in column 2
    })
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    console.log(`Triggered API for row ${editedRow}, response: ${response.getContentText()}`);
  } catch (error) {
    console.error(`Error triggering API for row ${editedRow}: ${error.message}`);
  }
}
