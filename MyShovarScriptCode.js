const SLIDES_TEMPLATE_ID = '1j90GUtznN7EcoXOLcA-CzKzwYUmsaQLLEH7VcQ-mTI8';
const OUTPUT_FOLDER_ID = '1wDQqV0wQvhPkqlNtM-Tq6QP2C2EgV4s7';

function handleFormSubmit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  //const editUrl = e.response.getEditResponseUrl(); // כתובת לעדכון הטופס

  const data = {};
  headers.forEach((h, i) => data[h] = rowData[i]); // טעינת המידע

  const mode = data.Mode;
  const isUpdate = mode && mode.includes('לא');


  const expiryFormatted = data.expiry  // עיצוב התוקף בפורמט הנכון
    ? Utilities.formatDate(
      data.expiry instanceof Date ? data.expiry : new Date(data.expiry),
      Session.getScriptTimeZone(),
      'dd/MM/yyyy'
    )
    : '';


  const statusCol = headers.indexOf('status') + 1;
  const pdfCol = headers.indexOf('pdf_url') + 1;
  const imageCol = headers.indexOf('image_url') + 1;
  const createdCol = headers.indexOf('created_at') + 1;

  const folderIdCol = headers.indexOf('folder_id') + 1;
  const pdfIdCol = headers.indexOf('pdf_file_id') + 1;
  const imageIdCol = headers.indexOf('image_file_id') + 1;
  const lastUpdatedCol = headers.indexOf('LastUpdated') + 1;
  const errorCol = headers.indexOf('ErrorMsg') + 1;

  if (statusCol < 1 || pdfCol < 1 || imageCol < 1 || createdCol < 1) {
    throw new Error('Ooops - Missing required columns in spreadsheet');
  }

  let targetRow = row;



  // =============================
  // מצב: תיקון שובר אחרון
  // =============================
  if (isUpdate) {

    const lastRow = sheet.getLastRow();

    for (let i = lastRow; i > 1; i--) {
      const statusValue = sheet.getRange(i, statusCol).getValue();
      if (statusValue === 'Created') {
        targetRow = i;
        break;
      }
    }

    if (!targetRow) return;

    // מעדכנים רק שדות תוכן ולא מזהים פנימיים
    const fieldsToUpdate = ['recipient', 'service', 'purchaser', 'expiry'];

    fieldsToUpdate.forEach(field => {
      const col = headers.indexOf(field) + 1;
      if (col > 0) {
        sheet.getRange(targetRow, col).setValue(data[field]);
      }
    });
    sheet.getRange(targetRow, lastUpdatedCol).setValue(new Date());


    // מסמנים את השורה החדשה כבקשת תיקון בלבד
    sheet.getRange(row, statusCol).setValue('Updated !');


  } else {

    // =============================
    // מצב: שובר חדש
    // =============================
    if (data.status === 'Created') return;

  }



  // =============================
  // יצירת השובר (חדש או דריסה)
  // =============================
  // ===== 1. שכפול תבנית Slides בשביל שובר חדש =====
  let copy;
  let slideId;
  
  try {
    copy = DriveApp.getFileById(SLIDES_TEMPLATE_ID)
      .makeCopy(`temp_${Date.now()}`, DriveApp.getFolderById(OUTPUT_FOLDER_ID));

    const presentation = SlidesApp.openById(copy.getId());
    const slides = presentation.getSlides();

    // ===== 2. החלפת משתנים בשובר בנתוני אמת =====
    slides.forEach(slide => {
      slide.replaceAllText('{{recipient}}', data.recipient || '');
      slide.replaceAllText('{{service}}', data.service || '');
      slide.replaceAllText('{{purchaser}}', data.purchaser || '');
      slide.replaceAllText('{{expiry}}', expiryFormatted);
      slide.replaceAllText('{{voucher_id}}', data.voucher_id || '');
    });


    // מזהה השקופית הראשונה
    slideId = slides[0].getObjectId();
    presentation.saveAndClose();

  } catch (err) {

    Logger.log('Voucher generation failed: ' + err);

    if (errorCol > 0) {
      sheet.getRange(row, errorCol).setValue(err.toString());
    }

    if (statusCol > 0) {
      sheet.getRange(row, statusCol).setValue('ERROR');
    }

    return;
  }


  // ===== 3. שם הקובץ ותיקייה ייעודית =====
  const createdDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const finalName = `${createdDate}__שובר מתנה__${data.voucher_id}`;

  //const purchaserSafe = data.purchaser || 'unknown';
  //const recipientSafe = data.recipient || 'unknown';
  //const folderNamePart = `__${data.voucher_id}__${purchaserSafe}__${recipientSafe}`;



  const parentFolder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
  let outputFolder;

  let existingFolderId = sheet.getRange(targetRow, folderIdCol).getValue();

  if (isUpdate && existingFolderId) {

    // שימוש בתיקייה קיימת לפי ID
    outputFolder = DriveApp.getFolderById(existingFolderId);

    // מחיקת קבצים קיימים לפי ID
    const existingPdfId = sheet.getRange(targetRow, pdfIdCol).getValue();
    const existingImageId = sheet.getRange(targetRow, imageIdCol).getValue();

    if (existingPdfId) {
      try { DriveApp.getFileById(existingPdfId).setTrashed(true); } catch (e) { }
    }

    if (existingImageId) {
      try { DriveApp.getFileById(existingImageId).setTrashed(true); } catch (e) { }
    }

  } else {

    // יצירת תיקייה חדשה
    const purchaserSafe = sanitize(data.purchaser);
    const recipientSafe = sanitize(data.recipient);
    const folderName = `${createdDate}__${data.voucher_id}__${purchaserSafe}__${recipientSafe}`;
    outputFolder = parentFolder.createFolder(folderName);

  }




  // ===== 4. יצירת PDF =====
  const pdfBlob = copy.getAs(MimeType.PDF).setName(finalName + '.pdf');
  const pdfFile = outputFolder.createFile(pdfBlob);



  // ===== 5. יצירת PNG =====
  const exportUrl = `https://docs.google.com/presentation/d/${copy.getId()}/export/png?pageid=${slideId}`;
  const token = ScriptApp.getOAuthToken();
  let imageFile;

  try {
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: 'Bearer ' + token }
    });
    const imageBlob = response.getBlob().setName(finalName + '.png');
    imageFile = outputFolder.createFile(imageBlob);
  } catch (imgErr) {
    Logger.log('PNG export failed: ' + imgErr);
  }

  // שמירת מזהים פנימיים בגיליון
  sheet.getRange(targetRow, folderIdCol).setValue(outputFolder.getId());
  sheet.getRange(targetRow, pdfIdCol).setValue(pdfFile.getId());
  sheet.getRange(targetRow, imageIdCol).setValue(imageFile.getId());




  // עדכון שורה רלוונטית
  sheet.getRange(targetRow, statusCol).setValue('Created');
  sheet.getRange(targetRow, pdfCol).setValue(pdfFile.getUrl());
  sheet.getRange(targetRow, imageCol).setValue(imageFile.getUrl());
  sheet.getRange(targetRow, createdCol).setValue(new Date());






  // ===== 7. שליחת מייל =====
  let ModeText = '';
  if (isUpdate) { ModeText = 'מתוקן' };

  const fixedEmail = 'nomik.lif@gmail.com'; // כתובת קבועה
  const fixedCC = 'rinatkom+voucher@gmail.com'; // כתובת קבועה
  const subject = `שובר מתנה ${ModeText} ${data.voucher_id}  מ${data.From} ל${data.To}`;
  const body = `
    מצורף השובר שנרכש ע"י  <b>${data.From}</b>  עבור  <b>${data.To}</b> 🎁
    <br><br><b>הקדשה:</b><br>${data.recipient}
    <br><br><b>קיבלת שובר מתנה:</b><br>${data.service}
    <br><br><b>ברכה:</b><br>${data.purchaser}
    <br><br><b>תוקף:</b> ${expiryFormatted}
    <br><b>הערה:</b> ${data.Note}
    <br><br>תודה רבה,<br>השובר-בוט שלך`;

  MailApp.sendEmail({
    to: fixedEmail,
    //cc: fixedCC,
    subject: subject,
    htmlBody: body,
    attachments: [pdfFile.getBlob(), imageFile.getBlob()]
  });



  // ===== 8. מחיקת המצגת הזמנית =====
  DriveApp.getFileById(copy.getId()).setTrashed(true);

}

// פוקנציית עזר לניקוי תווים מיוחדים שעלולים להכשיל את קוד יצירת תיקייה
function sanitize(str) {
  return String(str).replace(/[\\\/:*?"<>|]/g, '').trim();
}
