const SLIDES_TEMPLATE_ID = '1j90GUtznN7EcoXOLcA-CzKzwYUmsaQLLEH7VcQ-mTI8';
const OUTPUT_FOLDER_ID = '1wDQqV0wQvhPkqlNtM-Tq6QP2C2EgV4s7';

function handleFormSubmit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = sheet.getRange(row, 1, 1, headers.length).getValues()[0];

  const data = {};
  headers.forEach((h, i) => data[h] = rowData[i]);

  if (data.status && data.status !== 'Pending') return;

  // ===== 1. שכפול תבנית Slides =====
  const copy = DriveApp.getFileById(SLIDES_TEMPLATE_ID)
    .makeCopy(`temp_${Date.now()}`, DriveApp.getFolderById(OUTPUT_FOLDER_ID));

  const presentation = SlidesApp.openById(copy.getId());
  const slides = presentation.getSlides();

  // ===== 2. החלפת משתנים =====
  slides.forEach(slide => {
    slide.replaceAllText('{{recipient}}', data.recipient || '');
    slide.replaceAllText('{{service}}', data.service || '');
    slide.replaceAllText('{{purchaser}}', data.purchaser || '');
    const expiryFormatted = data.expiry
      ? Utilities.formatDate(
          data.expiry instanceof Date ? data.expiry : new Date(data.expiry),
          Session.getScriptTimeZone(),
          'dd/MM/yyyy'
        )
      : '';
    slide.replaceAllText('{{expiry}}', expiryFormatted);
    slide.replaceAllText('{{voucher_id}}', data.voucher_id || '');
  });

  // מזהה השקופית הראשונה
  const slideId = slides[0].getObjectId();

  presentation.saveAndClose();

  // ===== 3. שם הקובץ ותיקייה ייעודית =====
  const createdDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const finalName = `${createdDate}__שובר מתנה__${data.voucher_id}`;
  const purchaserSafe = data.purchaser || 'unknown';
  const recipientSafe = data.recipient || 'unknown';
  const folderName = `${createdDate}__${data.voucher_id}__${purchaserSafe}__${recipientSafe}`;

  // יצירת תת-תיקייה בתוך התיקייה הראשית
  const outputFolder = DriveApp.getFolderById(OUTPUT_FOLDER_ID).createFolder(folderName);

  // ===== 4. יצירת PDF =====
  const pdfBlob = copy.getAs(MimeType.PDF).setName(finalName + '.pdf');
  const pdfFile = outputFolder.createFile(pdfBlob);

  // ===== 5. יצירת PNG =====
  const exportUrl = `https://docs.google.com/presentation/d/${copy.getId()}/export/png?pageid=${slideId}`;
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + token }
  });
  const imageBlob = response.getBlob().setName(finalName + '.png');
  const imageFile = outputFolder.createFile(imageBlob);

  // ===== 6. עדכון הגיליון =====
  const statusCol = headers.indexOf('status') + 1;
  const pdfCol = headers.indexOf('pdf_url') + 1;
  const imageCol = headers.indexOf('image_url') + 1;
  const createdCol = headers.indexOf('created_at') + 1;

  if (statusCol) sheet.getRange(row, statusCol).setValue('Created');
  if (pdfCol) sheet.getRange(row, pdfCol).setValue(pdfFile.getUrl());
  if (imageCol) sheet.getRange(row, imageCol).setValue(imageFile.getUrl());
  if (createdCol) sheet.getRange(row, createdCol).setValue(new Date());

  // ===== 7. שליחת מייל =====
    const expiryFormatted = data.expiry
      ? Utilities.formatDate(
          data.expiry instanceof Date ? data.expiry : new Date(data.expiry),
          Session.getScriptTimeZone(),
          'dd/MM/yyyy'
        )
      : '';
    const fixedEmail = 'nomik.lif@gmail.com'; // כתובת קבועה
    const fixedCC = 'rinatkom+voucher@gmail.com'; // כתובת קבועה
    const subject = `שובר מתנה ${data.voucher_id} מ${data.From} ל${data.To}`;
    const body = `היי,\n\nמצורף השובר של ${data.From} עבור ${data.To} 🎁\n\nהקדשה:\n${data.recipient}\n\nשובר:\n${data.service}\n\nברכה:\n${data.purchaser}\n\nתוקף: ${expiryFormatted}\n\nתודה רבה,\nרינת ליפשיץ`;

    MailApp.sendEmail({
      to: fixedEmail,
      cc: fixedCC,
      subject: subject,
      body: body,
      attachments: [pdfFile.getBlob(), imageFile.getBlob()]
    });

  // ===== 8. מחיקת המצגת הזמנית =====
  DriveApp.getFileById(copy.getId()).setTrashed(true);
}