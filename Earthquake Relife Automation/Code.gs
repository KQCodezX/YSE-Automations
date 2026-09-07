function onFormSubmit(e) {
  // Extract form data (UPDATED INDICES)
  const responses = e.values;
  const name = responses[1]; // Name
  const invoiceDate = responses[2]; // Invoice Date
  const email = responses[3]; // Email
  const phone = responses[4]; // Phone Number
  const donationAmount = responses[5]; // Donation Amount
  const paymentMethod = responses[6]; // Payment Method
  const receiptUrls = responses[7]; // Receipt URLs

  // Configuration
  const TEMPLATE_ID = 'XXXXXX'; // Your template ID
  const DRIVE_FOLDER_ID = 'XXXXX'; // Folder for temp files

  try {
    // Create invoice document
    const templateFile = DriveApp.getFileById(TEMPLATE_ID);
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const newFile = templateFile.makeCopy(`Invoice_${name}_${invoiceDate}`, folder);
    const doc = DocumentApp.openById(newFile.getId());
    const body = doc.getBody();

    // Replace placeholders in the document
    body.replaceText('<<Name>>', name);
    body.replaceText('<<Phone Number>>', phone);
    body.replaceText('<<Email>>', email);
    body.replaceText('<<Payment Method>>', paymentMethod);
    body.replaceText('<<Donation Amount>>', donationAmount);
    body.replaceText('<<Invoice Date>>', invoiceDate);
    doc.saveAndClose();

    // Prepare PDF invoice
    const pdfBlob = newFile.getAs('application/pdf');
    pdfBlob.setName(`Invoice_${name}.pdf`);

    // Process receipt files
    const attachments = [pdfBlob]; // Start with invoice PDF
    if (receiptUrls) {
      receiptUrls.split(',').forEach(url => {
        try {
          // Extract file ID from URL (fixed regex)
          const fileId = url.trim().match(/\/file\/d\/([^\/]+)/)[1];
          const file = DriveApp.getFileById(fileId);
          attachments.push(file.getAs('application/pdf').setName(file.getName()));
        } catch (err) {
          Logger.log(`Failed to attach file from URL: ${url}`);
        }
      });
    }

    // Customize email
    const subject = `${name} - YSE Earthquake 2025 Donation Receipt`;
    const emailBody = `
      Dear ${name},

      We are deeply grateful for your generous donation to the YSE တူတူရှိနေမယ် Campaign in support of Myanmar earthquake relief efforts. Your contribution will help provide critical aid and support to those affected by this disaster.

      Attached is your official donation invoice for your records. Please don’t hesitate to reach out if you have any questions or need further assistance.

      Once again, thank you for standing with us in this time of need. Together, we can make a difference.

      With gratitude,
      YSE Team
    `;

    // Send email with attachments
    GmailApp.sendEmail(email, subject, emailBody, {
      htmlBody: emailBody.replace(/\n/g, '<br>'), // Format as HTML
      attachments: attachments,
      bcc: 'X@gmail.com, X@gmail.com'
    });

    // Cleanup
    newFile.setTrashed(true);

  } catch (error) {
    Logger.log('Error: ' + error.toString());
  }
}