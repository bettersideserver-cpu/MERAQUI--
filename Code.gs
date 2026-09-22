function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No POST data received.');
    }

    var data = JSON.parse(e.postData.contents);
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // SITE VISIT BOOKING
    // Primary check: type === "siteVisit"
    // Fallback check: date + time are present, so older cached website builds
    // can still be routed to the Visit_Form sheet correctly.
    var isSiteVisit = data.type === "siteVisit" ||
      (!!data.date && !!data.time && !!(data.name || data.fullName) && !!(data.phone || data.telephone));

    if (isSiteVisit) {
      var siteVisitSheet = spreadsheet.getSheetByName("Visit_Form");

      if (!siteVisitSheet) {
        throw new Error('Sheet "Visit_Form" not found. Rename your second tab to Visit_Form.');
      }

      siteVisitSheet.appendRow([
        new Date(),
        data.name || data.fullName || "",
        data.phone || data.telephone || "",
        data.date || "",
        data.time || ""
      ]);

      return jsonResponse({
        success: true,
        message: "Site visit saved successfully"
      });
    }

    // NORMAL ENQUIRY
    var enquirySheet = spreadsheet.getSheetByName("Enquiry_Sheet");

    if (!enquirySheet) {
      throw new Error('Sheet "Enquiry_Sheet" not found. Rename your first tab to Enquiry_Sheet.');
    }

    enquirySheet.appendRow([
      new Date(),
      data.fullName || data.name || "",
      data.telephone || data.phone || "",
      data.email || "",
      data.residence || data.interest || "",
      data.message || ""
    ]);

    return jsonResponse({
      success: true,
      message: "Enquiry saved successfully"
    });

  } catch (error) {
    return jsonResponse({
      success: false,
      error: error.toString()
    });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
