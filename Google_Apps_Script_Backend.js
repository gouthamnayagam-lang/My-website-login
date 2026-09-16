// Code.gs
function doPost(e) {
  try {
    // 1. Get the active sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // 2. Parse incoming JSON data
    var data = JSON.parse(e.postData.contents);
    var timestamp = new Date();
    
    // 3. Append data to the next empty row
    sheet.appendRow([
      timestamp, 
      data.name, 
      data.email, 
      data.projectType, 
      data.budget, 
      data.details
    ]);
    
    // 4. Send Email Notification
    var adminEmail = Session.getEffectiveUser().getEmail(); // Sends to your Google Account email
    var subject = "🚀 New Project Request from " + data.name;
    var body = "You have received a new project order!\n\n" +
               "Name: " + data.name + "\n" +
               "Email: " + data.email + "\n" +
               "Project Type: " + data.projectType + "\n" +
               "Budget: " + data.budget + "\n\n" +
               "Project Details:\n" + data.details + "\n\n" +
               "Log in to your Google Sheet to view full details.";
               
    MailApp.sendEmail(adminEmail, subject, body);
    
    // 5. Return success response
    return ContentService.createTextOutput(JSON.stringify({
      "status": "success", 
      "message": "Order processed successfully."
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch(error) {
    // Handle errors
    return ContentService.createTextOutput(JSON.stringify({
      "status": "error", 
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Handle CORS preflight request for fetch() API
function doOptions(e) {
  var headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  };
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.JSON)
    .setHeaders(headers);
}
