function onEdit(e) {
  try {
    var sheetId = '1o9ugGmQetjPEDntx1dVHHfUfOvMPe4u6D8tHEXy2G04';
    var sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();

    var range = e.range;
    var selectedValue = range.getValue();
    var column = range.getColumn();
    var statusColumn = column + 1;

    Logger.log('Edit event - Column: %s, Selected Value: %s', column, selectedValue);

    if (column === 15 && selectedValue === 'Paid') {
      var row = range.getRow();
      var submitterName = sheet.getRange(row, 2).getValue();
      var particPant = sheet.getRange(row, 5).getValue();
      var phone = sheet.getRange(row, 3).getValue();
      var submitterEmail = sheet.getRange(row, 4).getValue();
      var submissionDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-YYYY');

      // Generate Receipt Number and set it in Google Sheet column N (16)
      var receiptNumber = generateReceiptNumber();
      sheet.getRange(row, 14).setValue(receiptNumber);

      var statusCell = sheet.getRange(row, statusColumn);
      var emailCell = sheet.getRange(row, 6); // Assuming email is in column F

      var emailAddress = emailCell.getValue();
      var pdfBlob = generateReceipt(submitterName, particPant, phone, submitterEmail, submissionDate, receiptNumber);

      // Send the PDF to Telegram
      sendPdfToTelegram(pdfBlob, submitterName,'Receipt');
      

      // Send the PDF to Telegram
     // sendPdfToTelegram(pdfBlob, submitterName,'Receipt');
      
      // Attempt to send the email
      var emailSent = true;
      try {
        var emailResult = MailApp.sendEmail({
          to: emailAddress,
          subject: 'Payment Confirmation',
          body: 'Thank you for your payment. Please find the attached receipt.',
          attachments: [pdfBlob],
        });

        if (emailResult) {
          emailSent = false;
          // Update the status value
          statusCell.setValue('Receipt Was not Sent');
          Logger.log('Email sent to: %s', emailAddress);
        } else {
          // Handle the case where the email result is false
          statusCell.setValue('Receipt Was  Sent');
          Logger.log('Receipt sending email to: %s', emailAddress);
        }
      } catch (error) {
        // Handle any exceptions during email sending
        statusCell.setValue('Email Sending Exception');
        Logger.log('Exception sending email to: %s. Error: %s', emailAddress, error);
      }
    }
  } catch (error) {
    Logger.log('Error in onEdit: %s', error);
  }
}


// Helper function to generate a unique receipt number with a sequential part
function generateReceiptNumber() {
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

  // Combine the current year and sequential number to form the receipt number
  return currentYear + '-' + formattedSequentialNumber;
}

function generateReceipt(submitterName, particPant, phone, submitterEmail, submissionDate, receiptNumber) {
  try {
    // Create a new Google Doc as a template
    var templateId = '1E04yQjSiA9efiS1QkNUwX0xK1g121wpFBBT802zVpDk';
    var templateDoc = DriveApp.getFileById(templateId);
    var copiedFile = templateDoc.makeCopy('Receipt_for_' + submitterName + '_' + receiptNumber);
    var copiedDoc = DocumentApp.openById(copiedFile.getId());

    // Generate placeholders for the receipt template
    var placeholders = {
      "{{name}}": submitterName,
      "{{Particpant}}": particPant,
      "{{InvoiceNumber}}": receiptNumber,
      "{{email}}": submitterEmail,
      "{{phone}}": phone,
      "{{Date}}": submissionDate,
    };

    // Replace placeholders in the copied document
    replacePlaceholdersInDoc(copiedDoc.getBody(), placeholders);

    // Save changes in the copied document
    copiedDoc.saveAndClose();

    // Convert the document to PDF
    var pdfBlob = copiedFile.getAs(MimeType.PDF);

    // Delete the temporary template document
    copiedFile.setTrashed(true);

    // Log success
    Logger.log('Receipt generated successfully for ' + submitterName + ' (' + submitterEmail + ')');

    return pdfBlob;
  } catch (error) {
    Logger.log('Error generating receipt for ' + submitterName + ' (' + submitterEmail + '): ' + error.message);
    return null;
  }
}

// Helper function to replace placeholders in Google Docs
function replacePlaceholdersInDoc(body, placeholders) {
  Object.keys(placeholders).forEach(function (placeholder) {
    // Use replaceText to correctly replace placeholders in Google Docs
    body.replaceText(placeholder, placeholders[placeholder]);
  });
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
    var message = "Receipt for " + submitterName + "\n" + pdfUrl;

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

