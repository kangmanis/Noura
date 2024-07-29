function onEdit(e) {
  // Check if the edit is in the Quotation sheet
  var sheetName = 'Quotation';
  var sheet = e.source.getSheetByName(sheetName);
  if (e.range.getSheet().getName() === sheetName) {
    // Check if the edit is in the relevant range (e.g., new row added)
    if (e.range.getRow() > 1) { // Skip header row
      createQuotationFromTemplate();
    }
  }
}

function createQuotationFromTemplate() {
  try {
    var clientSheetName = "Client Info";
    var eventSheetName = "Event Info";
    var packageSheetName = "Base Packages";
    var addOnSheetName = "Adds on";
    var quotationSheetName = "Quotation";
    var templateId = "14HfqsHQlCTkp0gttkAbJbSbhj4c7R137zsAq-sLmsXQ"; // Replace with your Google Docs template ID
    var folderId = "1pAz4NLMZjS28Zs8m4a6QsHof1oEhCyAM"; // Replace with your Google Drive folder ID

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var clientSheet = spreadsheet.getSheetByName(clientSheetName);
    var eventSheet = spreadsheet.getSheetByName(eventSheetName);
    var packageSheet = spreadsheet.getSheetByName(packageSheetName);
    var addOnSheet = spreadsheet.getSheetByName(addOnSheetName);
    var quotationSheet = spreadsheet.getSheetByName(quotationSheetName);

    if (!clientSheet || !eventSheet || !packageSheet || !addOnSheet || !quotationSheet) {
      throw new Error("One or more sheets not found. Please check sheet names.");
    }

    var quotationData = quotationSheet.getRange("A2:L" + quotationSheet.getLastRow()).getValues(); // Adjust range as needed
    var clientData = clientSheet.getRange("A2:J" + clientSheet.getLastRow()).getValues();
    var eventData = eventSheet.getRange("A2:I" + eventSheet.getLastRow()).getValues();
    var packageData = packageSheet.getRange("A2:E" + packageSheet.getLastRow()).getValues();
    var addOnData = addOnSheet.getRange("A2:C" + addOnSheet.getLastRow()).getValues();

    if (!quotationData || quotationData.length === 0) {
      throw new Error("No data found in the quotation sheet.");
    }

    for (var i = 0; i < quotationData.length; i++) {
      var row = quotationData[i];
      var quotationId = row[0];
      var clientId = row[1];
      var packageId = row[2];
      var basePackagePrice = row[3];
      var addOnIds = row[4] ? row[4].split(", ") : [];
      var addOnCosts = row[5];
      var totalCost = row[6];
      var status = row[7];
      var dateIssued = row[8];
      var expiryDate = row[9];

      var client = clientData.find(row => row[0] === clientId);
      if (!client) {
        throw new Error("Client data not found for ClientID " + clientId);
      }
      var [ , firstName, lastName, emailAddress, phoneNumber, address, weddingDate, numberOfGuests, preferredContactMethod, notes ] = client;

      var event = eventData.find(row => row[1] === clientId);
      if (!event) {
        throw new Error("Event data not found for ClientID " + clientId);
      }
      var [ , , eventType, date, startTime, endTime, venue, details, theme ] = event;

      var package = packageData.find(row => row[0] === packageId);
      if (!package) {
        throw new Error("Package data not found for PackageID " + packageId);
      }
      var [ , packageName, description, basePrice ] = package;

      var addOns = addOnIds.map(id => {
        var addOn = addOnData.find(row => row[0] === id);
        return addOn ? { name: addOn[1], description: addOn[2], price: addOn[3] } : null;
      }).filter(Boolean);

      var templateFile = DriveApp.getFileById(templateId);
      var newFileName = "Quotation for " + firstName + " " + lastName;
      var newFile = templateFile.makeCopy(newFileName, DriveApp.getFolderById(folderId));

      var doc = DocumentApp.openById(newFile.getId());
      var body = doc.getBody();

      body.replaceText("{{DateIssued}}", dateIssued || "");
      body.replaceText("{{FirstName}}", firstName || "");
      body.replaceText("{{LastName}}", lastName || "");
      body.replaceText("{{EmailAddress}}", emailAddress || "");
      body.replaceText("{{PhoneNumber}}", phoneNumber || "");
      body.replaceText("{{Address}}", address || "");
      body.replaceText("{{WeddingDate}}", weddingDate || "");
      body.replaceText("{{NumberOfGuests}}", numberOfGuests || "");
      body.replaceText("{{PreferredContactMethod}}", preferredContactMethod || "");
      body.replaceText("{{Notes}}", notes || "");
      body.replaceText("{{EventType}}", eventType || "");
      body.replaceText("{{Date}}", date || "");
      body.replaceText("{{StartTime}}", startTime || "");
      body.replaceText("{{EndTime}}", endTime || "");
      body.replaceText("{{Venue}}", venue || "");
      body.replaceText("{{Theme}}", theme || "");
      body.replaceText("{{Details}}", details || "");
      body.replaceText("{{PackageName}}", packageName || "");
      body.replaceText("{{Description}}", description || "");
      body.replaceText("{{BasePrice}}", basePrice || "");
      body.replaceText("{{BasePackagePrice}}", basePackagePrice || "");
      body.replaceText("{{AddOnNames}}", addOns.map(addOn => addOn.name).join(", ") || "");
      body.replaceText("{{AddOnDescriptions}}", addOns.map(addOn => addOn.description).join(", ") || "");
      body.replaceText("{{AddOnCosts}}", addOnCosts || "");
      body.replaceText("{{TotalCost}}", totalCost || "");
      body.replaceText("{{Status}}", status || "");
      body.replaceText("{{ExpiryDate}}", expiryDate || "");

      // Send the document via email
      MailApp.sendEmail({
        to: emailAddress,
        subject: "Your Quotation - " + newFileName,
        body: "Dear " + firstName + " " + lastName + ",\n\nPlease find attached your quotation.\n\nBest regards,\nYour Company",
        attachments: [DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF)]
      });

      // Optionally move the document to the folder
      DriveApp.getFolderById(folderId).addFile(DriveApp.getFileById(doc.getId()));
      DriveApp.getRootFolder().removeFile(DriveApp.getFileById(doc.getId())); // Remove from root folder
    }
  } catch (error) {
    Logger.log("Error: " + error.message);
  }
}
