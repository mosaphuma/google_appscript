//This funntion sent invoice to email that has register 
//it generate invoice number  year-0000

function submitFormdocTelegram() {
  try {
  var form = FormApp.openById('1_mdfUPvQ6NRNQDgFcYXxFVxtYSgfDrpudh3PfiYeOm8');
  var formResponses = form.getResponses();

  // Get the latest form response
  var latestResponse = formResponses[formResponses.length - 1];
  var itemResponses = latestResponse.getItemResponses();

  var submitterName = itemResponses[0].getResponse();
  var particPant = itemResponses[5].getResponse();
  var phone = itemResponses[3].getResponse();
  var submitterEmail = itemResponses[4].getResponse();
  var submissionDate = Utilities.formatDate(latestResponse.getTimestamp(), Session.getScriptTimeZone(), 'dd-MM-YYYY ');
  Logger.log('Processing form response for ' + submitterName + ' (' + submitterEmail + ')');

// Generate a common invoice number
  var invoiceNumber = generateInvoiceNumber();

  // Generate the invoice and get the PDF blob
  var pdfBlob = generateInvoice(submitterName, particPant, phone, submitterEmail, submissionDate,invoiceNumber);

  // Send the PDF to Telegram
  sendPdfToTelegram(pdfBlob, submitterName,'Invoice');



  //reciept
// Generate the receipt and get the PDF blob
//var receiptBlob = generateReceipt(submitterName, particPant, phone, submitterEmail, submissionDate,invoiceNumber);

  // Save the PDF blob to Google Drive
  //saveToDrive(receiptBlob, 'Receipt_for_' + submitterName);

  // Send the PDF to Telegram
  //sendPdfToTelegram(receiptBlob, submitterName, 'Receipt');

  // Update the Google Sheet with invoice and receipt numbers
  //updateSheetWithNumbers(latestResponse.getId(), invoiceNumber);

  // Insert a new row into the Google Sheet
  //insertRowIntoSheet(submitterName, particPant, phone, submitterEmail, submissionDate, invoiceNumber);
  //insertNumbersIntoSheet(submitterName, invoiceNumber);
insertRowIntoSheet(submitterName, particPant, phone, submitterEmail, submissionDate, invoiceNumber);
  //insertNumbersIntoSheet(submitterName, invoiceNumber);
  } catch (error) {
    Logger.log('Error in submitFormdocTelegram: ' + error.message);

Logger.log('Updated sheet with invoice number: ' + invoiceNumber);
  }
}

function insertRowIntoSheet(submitterName, particPant, phone, submitterEmail, submissionDate, invoiceNumber) {
  try {
    var sheet = SpreadsheetApp.openById('1k-NnU7YV7Em1rt56cK8_ySQIohMuOzGTaZVQVWnkI3o').getActiveSheet();
    
    // Get all data from the sheet
    var data = sheet.getDataRange().getValues();
    
    // Find the row with the submitter's name
    var rowIndex = -1;
    for (var i = 0; i < data.length; i++) {
      if (data[i][1] == submitterName && data[i][5] == submitterEmail) {
        rowIndex = i + 1; // Adding 1 because sheet rows are 1-indexed
        break;
      }
    }

    if (rowIndex !== -1) {
      // Update values in the found row
     // sheet.getRange(rowIndex, 1).setValue(submitterName);
      //sheet.getRange(rowIndex, 2).setValue(particPant);
     // sheet.getRange(rowIndex, 3).setValue(phone);
     // sheet.getRange(rowIndex, 4).setValue(submitterEmail);
      //sheet.getRange(rowIndex, 5).setValue(submissionDate);
      sheet.getRange(rowIndex, 13).setValue(invoiceNumber);  // Column H
      //sheet.getRange(rowIndex, 14).setValue(invoiceNumber);  // Column I

      Logger.log('Updated row in the sheet for ' + submitterName);
    } else {
      Logger.log('Submitter not found in the sheet: ' + submitterName);
    }
  } catch (error) {
    Logger.log('Error updating row in the sheet: ' + error.message);
  }
}

/*
function generateReceipt(submitterName, particPant, phone, submitterEmail, submissionDate,invoiceNumber) {
  try {
    // Create a new Google Doc as a template
    var templateId = '1E04yQjSiA9efiS1QkNUwX0xK1g121wpFBBT802zVpDk';
    var templateDoc = DriveApp.getFileById(templateId);
    var copiedFile = templateDoc.makeCopy('Receipt_for_' + submitterName +'_'+ invoiceNumber);
    var copiedDoc = DocumentApp.openById(copiedFile.getId());
    // Generate an invoice number
   // var invoiceNumber = generateInvoiceNumber();
    // Generate placeholders for the receipt template
    var placeholders = {
      "{{name}}": submitterName,
      "{{Particpant}}": particPant,
      "{{InvoiceNumber}}": invoiceNumber,
      "{{email}}": submitterEmail,
      "{{phone}}": phone,
      "{{Date}}": submissionDate
    };

    // Replace placeholders in the copied document
    var body = copiedDoc.getBody();
    replacePlaceholdersInDoc(body, placeholders);

    // Save changes in the copied document
    copiedDoc.saveAndClose();

    // Convert the document to PDF
    var pdfBlob = DriveApp.getFileById(copiedFile.getId()).getAs(MimeType.PDF);

    // Delete the temporary template document
    DriveApp.getFileById(copiedFile.getId()).setTrashed(true);

    // Log success
    Logger.log('Receipt generated successfully for ' + submitterName + ' (' + submitterEmail + ')');

    return pdfBlob;
  } catch (error) {
    Logger.log('Error generating receipt for ' + submitterName + ' (' + submitterEmail + '): ' + error.message);
    return null;
  }
}
*/

function generateInvoice(submitterName, particPant, phone, submitterEmail, submissionDate,invoiceNumber) {
  try {
    // Create a new Google Doc as a template
    var templateId = '1OOUP1fIs41NJbTECT6emW7Qp2S9FC6B1aql1fWMVW0M';
    var templateDoc = DriveApp.getFileById(templateId);
    var copiedFile = templateDoc.makeCopy('Invoice_for_ ' + submitterName+'_'+ invoiceNumber);
    var copiedDoc = DocumentApp.openById(copiedFile.getId());

    // Generate an invoice number
    //var invoiceNumber = generateInvoiceNumber();

    // Replace placeholders in the copied document
    var body = copiedDoc.getBody();
    var placeholders = {
      "{{name}}": submitterName,
      "{{Particpant}}": particPant,
      "{{InvoiceNumber}}": invoiceNumber,
      "{{email}}": submitterEmail,
      "{{phone}}": phone,
      "{{Date}}": submissionDate
    };

    replacePlaceholdersInDoc(body, placeholders);

    // Save changes in the copied document
    copiedDoc.saveAndClose();

    // Convert the document to PDF
    var pdfBlob = DriveApp.getFileById(copiedFile.getId()).getAs(MimeType.PDF);



   // Send the email with HTML body and PDF attachment
    MailApp.sendEmail({
      to: submitterEmail,
      subject: "Invoice for " + submitterName,
      htmlBody: "Dear " + submitterName + ",<br><br>" +
            "គោរពជូន<br>" +
            "លោកឧកញ៉ា លោកជំទាវ លោក លោកស្រី ប្រធានសហគ្រាស<br><br>" +
            "និយ័តករគណនេយ្យនិងសវនកម្ម នៃអាជ្ញាធរសេវាហិរញ្ញវត្ថុមិនមែនធនាគារនឹងរៀបចំសិក្ខាសាលាផ្សព្វផ្សាយស្តីពី"+
            '"<b>វគ្គបណ្តុះបណ្តាលស្តង់ដារបាយការណ៍ហិរញ្ញវត្ថុអន្តរជាតិនៃកម្ពុជាសម្រាប់សហគ្រាសធុនតូចនិងមធ្យម (CIFRS for SMEs)</b>"<br><br>' +
            "សិក្ខាសាលានេះនឹងប្រព្រឹត្តទៅនា ៖<br>"+
            "ថ្ងៃទី១៧ ខែមករា ឆ្នាំ២០២៤<br><br>" +
            "វេលាម៉ោង៨:០០នាទីព្រឹក នៅសណ្ឋាគារសុខា ខេត្តព្រះសីហនុ។<br>"+
            "សម្រាប់ព័ត៌មានបន្ថែមសូមទំនាក់ទំនង<br>"+
            "កញ្ញា សុខ ណាវី <br>"+
            "០៦៩៥៥១៥០៧<br>"+
            "សូមអរគុណចំពោះការបញ្ជូនរបស់អ្នក។ សូមស្វែងរកវិក្កយបត្រដែលភ្ជាប់មកជាមួយសម្រាប់អ្នកជាឯកសារយោង។",
      attachments: [pdfBlob]
    });
    
    
    // Delete the temporary template document
    DriveApp.getFileById(copiedFile.getId()).setTrashed(true);

    // Log success
    Logger.log('Invoice generated successfully for ' + submitterName + ' (' + submitterEmail + ')');

    return pdfBlob;
  } catch (error) {
    Logger.log('Error generating invoice for ' + submitterName + ' (' + submitterEmail + '): ' + error.message);
    return null;
  }

}

function sendPdfToTelegram(pdfBlob, submitterName) {
  try {
    // Upload the PDF to Google Drive
    var folder = DriveApp.getFolderById('1t5vvXS3t2-YvEJSHMltt9wVw7m54lhXO'); // Replace with your actual folder ID
    var file = folder.createFile(pdfBlob);

    // Get the URL of the uploaded PDF
    var pdfUrl = file.getUrl();

    // Replace 'YOUR_BOT_TOKEN' and 'CHAT_ID' with your Telegram bot token and chat ID
    var botToken = '1909271216:AAEsk3wInqTDEG_omIPHcEctkfXO3sAvmm4';
    var chatId = '418789011';

    // Compose the Telegram message with the PDF link
    var message = "Invoice for " + submitterName + "\n" + pdfUrl;

    // Send the message to Telegram
    var apiUrl = 'https://api.telegram.org/bot' + botToken + '/sendMessage';
    var payload = {
      method: 'post',
      payload: {
        'chat_id': chatId,
        'text': message,
      }
    };

    UrlFetchApp.fetch(apiUrl, payload);
  } catch (error) {
    Logger.log('Error sending PDF to Telegram: ' + error.message);
  }
}

// Helper function to generate a unique invoice number with a sequential part
function generateInvoiceNumber() {
    // Get the current year
    var currentYear = new Date().getFullYear();
  
    // Get the stored sequential number from Script Properties
    var scriptProperties = PropertiesService.getScriptProperties();
    var sequentialNumber = scriptProperties.getProperty('sequentialNumber');
  
    // If the sequential number is not stored or is not a number, initialize it to 1
    if (!sequentialNumber || isNaN(sequentialNumber)) {
      sequentialNumber = 1;
    } else {
      // Increment the sequential number for each new submission
      sequentialNumber++;
    }
  
    // Store the updated sequential number back to Script Properties
    scriptProperties.setProperty('sequentialNumber', sequentialNumber);
  
    // Format the sequential part as four digits (e.g., "0001")
    var formattedSequentialNumber = ('0000' + sequentialNumber).slice(-4);
  
    // Combine the current year and sequential number to form the invoice number
    return currentYear + '-' + formattedSequentialNumber;
  }

// Helper function to replace placeholders in Google Slides
function replacePlaceholdersInSlides(presentation, placeholders) {
    var slides = presentation.getSlides();
    slides.forEach(function (slide) {
      var shapes = slide.getShapes();
      shapes.forEach(function (shape) {
        if (shape.getText) {
          var text = shape.getText();
          Object.keys(placeholders).forEach(function (placeholder) {
            text.replaceAllText(placeholder, placeholders[placeholder]);
          });
        }
      });
    });
  }
  
  // Helper function to replace placeholders in Google Docs
function replacePlaceholdersInDoc(body, placeholders) {
    Object.keys(placeholders).forEach(function (placeholder) {
      // Use replaceText to correctly replace placeholders in Google Docs
      body.replaceText(placeholder, placeholders[placeholder]);
    });
  }


/*

function updateSheetWithNumbers1(responseId, invoiceNumber) {
  // Open the Google Sheet associated with the form responses
  var spreadsheetId = '1o9ugGmQetjPEDntx1dVHHfUfOvMPe4u6D8tHEXy2G04'; // Replace with your actual spreadsheet ID
  var sheet = SpreadsheetApp.openById(spreadsheetId).getActiveSheet();

  // Find the row corresponding to the form response
  var responses = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues(); // Assuming form responses start from row 2
  var rowIndex = responses.findIndex(function (response) {
    return response[0] === responseId;
  });

  // Update the sheet with invoice and receipt numbers
  if (rowIndex !== -1) {
    sheet.getRange(rowIndex + 2, 8).setValue(invoiceNumber); // Assuming column H is for invoice number
    sheet.getRange(rowIndex + 2, 9).setValue(''); // Assuming column I is for receipt number (initially empty)
  }
}


function updateSheetWithNumbers(submitterName, type, number) {
  try {
    var sheet = SpreadsheetApp.openById('1o9ugGmQetjPEDntx1dVHHfUfOvMPe4u6D8tHEXy2G04').getActiveSheet();
    var lastRow = sheet.getLastRow() + 1;

    // Assuming invoice number in column H and receipt number in column I
    var column = (type === 'Invoice') ? 'H' : 'I';

    sheet.getRange(lastRow, getColumnNumber(column)).setValue(number);
    Logger.log('Updated sheet with ' + type + ' number ' + number + ' for ' + submitterName);
  } catch (error) {
    Logger.log('Error updating sheet: ' + error.message);
  }
}

function getColumnNumber(column) {
  return column.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
}
*/
