function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("index");
}

// API endpoint to search ticket
function checkTicket(ticketID) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Checkin"); // change if your sheet name is different

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const idxID = headers.indexOf("TicketID");
  const idxName = headers.indexOf("Name");
  const idxPax = headers.indexOf("Pax");
  const idxSeat = headers.indexOf("Seat");
  const idxStatus = headers.indexOf("Status");

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxID]) === String(ticketID)) {
      return {
        found: true,
        TicketID: data[i][idxID],
        Name: data[i][idxName],
        Pax: data[i][idxPax],
        Seat: data[i][idxSeat],
        Status: data[i][idxStatus],
        row: i + 1
      };
    }
  }
  return { found: false };
}

// Mark status as "Checked in"
function markCheckedIn(rowNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Checkin");
  sheet.getRange(rowNumber, 5).setValue("Checked in");
  return true;
}
