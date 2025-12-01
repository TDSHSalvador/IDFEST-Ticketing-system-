// ===============================================================
// ExecuteEmail — main entry point
// ===============================================================
function ExecuteEmail() {
  const sheetNames = [
    'Form Responses(Early)',
    'Form Responses(Presale)',
    'Form Responses(Normal)'
  ];
  sendEmailsWithFormattedTicketID(sheetNames);
}

// ===============================================================
// Email Sender
// ===============================================================
function sendEmailsWithFormattedTicketID(sheetNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get draft
  const draftSubject = "name of email draft";
  const drafts = GmailApp.getDrafts();
  let draftEmailMessage = null;

  for (const d of drafts) {
    if (d.getMessage().getSubject() === draftSubject) {
      draftEmailMessage = d.getMessage();
      break;
    }
  }

  if (!draftEmailMessage) {
    Logger.log("Draft with subject not found: " + draftSubject);
    return;
  }

  const draftBodyRaw = draftEmailMessage.getBody();

  // Get last ticket number across all sheets
  let nextNumeric = getLastTicketID(sheetNames, "ticket id");

  // Process sheets
  for (const sheetName of sheetNames) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) continue;

    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const header = data[0].map(h => (h ? String(h).toLowerCase().trim() : ''));

    const findCol = name => header.indexOf(name.toLowerCase()) + 1;

    const emailCol = findCol("email address");
    const nameCol = findCol("name");
    const statusCol = findCol("status");
    const emailStatusCol = findCol("email status");
    const ticketIdCol = findCol("ticket id");

    // Hard-coded ticket quantity (Column E = 5)
    const ticketQtyCol = 5;

    if (!emailCol || !statusCol || !emailStatusCol || !ticketIdCol) {
      Logger.log(`Missing required columns in sheet ${sheetName}`);
      continue;
    }

    // Loop rows
    for (let r = 2; r <= lastRow; r++) {
      const row = data[r - 1];

      const email       = row[emailCol - 1];
      const name        = row[nameCol - 1];
      const status      = row[statusCol - 1];
      const emailStatus = row[emailStatusCol - 1];
      const oldTicketID = row[ticketIdCol - 1];
      const qty         = row[ticketQtyCol - 1];

      if (status !== "Verified") continue;
      if (String(emailStatus).toLowerCase() === "sent") continue;

      // ---------------------------
      // ASSIGN TICKET ID
      // ---------------------------
      let ticketID;

      if (!oldTicketID || oldTicketID === "") {
        nextNumeric++;
        ticketID = `IDF25${String(nextNumeric).padStart(3, '0')}`;
      } else {
        ticketID = oldTicketID;
        Logger.log(`Reusing existing TicketID ${ticketID} on row ${r}`);
      }

      // ---------------------------
      // GENERATE QR CODE BLOB
      // ---------------------------
      const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${encodeURIComponent(ticketID)}`;
      let qrBlob;

      try {
        const response = UrlFetchApp.fetch(qrUrl);
        qrBlob = response.getBlob().setName(ticketID + ".png");
      } catch (err) {
        Logger.log("QR Generation Failed for " + ticketID + ": " + err);
        continue;
      }

      // ---------------------------
      // BUILD EMAIL BODY
      // ---------------------------
      let emailBody = draftBodyRaw;

      const replaceMap = {
        "{NAME}": name,
        "{Ticket ID}": ticketID,
        "{Ticket}": qty,
      };

      for (const key in replaceMap) {
        emailBody = emailBody.replace(new RegExp(key, "g"), replaceMap[key]);
      }

      // Inject QR Code <img> tag
      const qrImgTag = `
        <img src="cid:qrCode"
             style="margin: 0 auto; padding-top: 18%; display: block; max-width: 500px; width: 50%; height: auto;"
             alt="qrCode">
      `;

      emailBody = emailBody.replace(/\{QRCODE\}/g, qrImgTag);

      // ---------------------------
      // SEND EMAIL
      // ---------------------------
      try {
        GmailApp.sendEmail(String(email), draftSubject, "", {
          htmlBody: emailBody,
          inlineImages: { qrCode: qrBlob }
        });
      } catch (err) {
        Logger.log("FAILED to send email to " + email + ": " + err);
        if (!oldTicketID) nextNumeric--; // rollback only if NEW ID failed
        continue;
      }

      // ---------------------------
      // UPDATE SHEET
      // ---------------------------
      sheet.getRange(r, emailStatusCol).setValue("Sent");
      sheet.getRange(r, ticketIdCol).setValue(ticketID);

      Logger.log(`Email successfully sent to ${email} — Ticket: ${ticketID}`);
    }
  }
}

// ===============================================================
// Get Highest Existing Ticket ID across all sheets
// ===============================================================
function getLastTicketID(sheetNames, targetColName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let maxNum = 0;

  for (const sheetName of sheetNames) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) continue;

    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const header = data[0].map(h => (h ? String(h).toLowerCase().trim() : ''));

    const colIndex = header.indexOf(targetColName.toLowerCase().trim());
    if (colIndex < 0) continue;

    for (let r = 2; r < data.length; r++) {
      const val = data[r][colIndex];
      if (typeof val === "string" && val.startsWith("ticketid")) {
        const num = parseInt(val.substring(5), 10);
        if (!isNaN(num) && num > maxNum) maxNum = num;
      }
    }
  }

  return maxNum;
}
