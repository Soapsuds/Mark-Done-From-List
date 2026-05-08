Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("run-update").onclick = processMatches;
  }
});

async function processMatches() {
  await Excel.run(async (context) => {
    // 1. Get the data from your HTML UI
    const rawList = document.getElementById("patient-list").value;
    const externalNames = rawList.split("\n")
      .map(n => n.trim().toLowerCase())
      .filter(n => n !== "");

    if (externalNames.length === 0) return;

    // 2. Interact with Excel
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const table = sheet.tables.getItemAt(0);

    const bodyRange = table.getDataBodyRange();

    // 3. Establish the Firewall BEFORE doing anything else
    // Grab only the currently visible cells based on how you have the sheet filtered right now
    const visibleCells = bodyRange.getSpecialCellsOrNullObject(Excel.SpecialCellType.visible);
    
    // Load the data we need into memory
    const headerRange = table.getHeaderRowRange().load("values");
    bodyRange.load(["values", "rowIndex"]);
    visibleCells.load("rowIndex"); // This gets the absolute row index of the first visible row

    await context.sync();

    // Safety check: If the entire table is hidden, stop running
    if (visibleCells.isNullObject) {
      console.log("No visible rows found to process.");
      return;
    }

    // Capture our safe starting point and the table's starting point
    const firstVisibleSheetRow = visibleCells.rowIndex;
    const bodyStartRow = bodyRange.rowIndex;

    // 4. Find column indexes dynamically
    const headers = headerRange.values[0];
    const patientNameIndex = headers.indexOf("Patient Name");
    
    let statusColIndex = headers.indexOf(" Status");
    if (statusColIndex === -1) statusColIndex = headers.indexOf("Status");

    if (patientNameIndex === -1 || statusColIndex === -1) {
      console.error("Could not find 'Patient Name' or 'Status' column.");
      return;
    }

    const bodyValues = bodyRange.values;

    // 5. Loop through the data array in memory
    for (let i = 0; i < bodyValues.length; i++) {
      // Calculate exactly where this row lives on the spreadsheet
      const currentSheetRow = bodyStartRow + i;

      // THE FIREWALL: If this row is physically located above the first visible row (e.g., above 8881), skip it completely!
      if (currentSheetRow < firstVisibleSheetRow) {
        continue; 
      }

      const currentStatus = bodyValues[i][statusColIndex];
      const patientName = bodyValues[i][patientNameIndex];

      // Only act on rows that are currently Pending
      if (currentStatus === "Pending" && patientName) {
        const cleanName = patientName.toString().trim().toLowerCase();

        if (externalNames.includes(cleanName)) {
          // Surgically write "Done" only to the matching, approved cells
          bodyRange.getCell(i, statusColIndex).values = [["Done"]];
        }
      }
    }

    // 6. Push all the cell updates back to the spreadsheet at once
    await context.sync(); 
  });
}