/****************************************************
 * LINE LIFF SALES REPORT
 * Google Apps Script + Google Sheets + Google Drive
 ****************************************************/

const CONFIG = {
  SPREADSHEET_ID: '1XZ8tlL8aLcgCCdgWb1KPEO4ZI1Mf9t-BrXXeqmeXp90',

  SALES_SHEET: 'รายงานยอดขาย',
  STAFF_SHEET: 'พนักงาน',

  // Folder สำหรับเก็บรูปใบส่งยอด
  DRIVE_FOLDER_ID: '1fNe5bMW9N1SfSTVdLgFJLPAlRCuvh2ft',
};

/**
 * GET API
 *
 * ใช้สำหรับ
 * - getShops
 * - getStaffByShop
 */
function doGet(e) {

  const action = e.parameter.action || '';
  const callback = e.parameter.callback || '';

  let result;

  try {

    if (action === 'getShops') {

      result = {
        success: true,
        data: getShops()
      };

    }

    else if (action === 'getStaffByShop') {

      result = {
        success: true,
        data: getStaffByShop(
          e.parameter.shop || ''
        )
      };

    }

    else {

      result = {
        success: false,
        message: 'Unknown action'
      };

    }

  }

  catch (error) {

    result = {
      success: false,
      message: error.message
    };

  }


  /**
   * JSONP
   *
   * GitHub Pages สามารถเรียก Apps Script
   * ผ่าน <script> ได้
   */

  if (callback) {

    return ContentService
      .createTextOutput(
        callback + '(' +
        JSON.stringify(result) +
        ')'
      )
      .setMimeType(
        ContentService.MimeType.JAVASCRIPT
      );

  }


  return ContentService
    .createTextOutput(
      JSON.stringify(result)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}


/**
 * POST API
 *
 * ใช้สำหรับบันทึกรายงาน
 */
function doPost(e) {

  try {

    const data =
      JSON.parse(
        e.postData.contents
      );


    const result =
      saveSalesReport(data);


    return ContentService
      .createTextOutput(
        JSON.stringify(result)
      )
      .setMimeType(
        ContentService.MimeType.JSON
      );

  }

  catch (error) {

    return ContentService
      .createTextOutput(
        JSON.stringify({
          success: false,
          message: error.message
        })
      )
      .setMimeType(
        ContentService.MimeType.JSON
      );

  }

}


/**
 * GET SHOPS
 */
function getShops() {

  const ss =
    SpreadsheetApp.openById(
      CONFIG.SPREADSHEET_ID
    );

  const sheet =
    ss.getSheetByName(
      CONFIG.STAFF_SHEET
    );

  if (!sheet) {
    throw new Error(
      'ไม่พบ Sheet: ' +
      CONFIG.STAFF_SHEET
    );
  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        2
      )
      .getValues();


  const shops =
    values
      .map(row =>
        String(row[0]).trim()
      )
      .filter(Boolean);


  return [
    ...new Set(shops)
  ];
}


/**
 * GET STAFF BY SHOP
 */
function getStaffByShop(shop) {

  if (!shop) {
    return [];
  }

  const ss =
    SpreadsheetApp.openById(
      CONFIG.SPREADSHEET_ID
    );

  const sheet =
    ss.getSheetByName(
      CONFIG.STAFF_SHEET
    );

  if (!sheet) {
    throw new Error(
      'ไม่พบ Sheet: ' +
      CONFIG.STAFF_SHEET
    );
  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        2
      )
      .getValues();


  const staff =
    values
      .filter(row =>
        String(row[0]).trim() ===
        String(shop).trim()
      )
      .map(row =>
        String(row[1]).trim()
      )
      .filter(Boolean);


  return [
    ...new Set(staff)
  ];
}


/**
 * SAVE REPORT
 */
function saveSalesReport(data) {

  if (!data) {
    throw new Error(
      'ไม่พบข้อมูล'
    );
  }


  validateData(data);


  const ss =
    SpreadsheetApp.openById(
      CONFIG.SPREADSHEET_ID
    );


  const sheet =
    ss.getSheetByName(
      CONFIG.SALES_SHEET
    );


  if (!sheet) {
    throw new Error(
      'ไม่พบ Sheet: ' +
      CONFIG.SALES_SHEET
    );
  }


  let imageUrl = '';


  /**
   * SAVE IMAGE
   */
  if (
    data.image &&
    data.image.data
  ) {

    imageUrl =
      saveImageToDrive(
        data.image.data,
        data.image.name ||
          'ใบส่งยอด.jpg',
        data.image.mimeType ||
          'image/jpeg',
        data.shop,
        data.date
      );

  }


  /**
   * EMPLOYEES
   */
  const employeeNames =
    Array.isArray(data.employees)
      ? data.employees.join(', ')
      : '';


  /**
   * SAVE SHEET
   */
  sheet.appendRow([

    new Date(),

    data.lineUserId || '',

    data.lineName || '',

    data.shop || '',

    data.date || '',

    data.shift || '',

    Number(data.totalSales) || 0,

    Number(data.cash) || 0,

    Number(data.credit) || 0,

    Number(data.qr) || 0,

    Number(data.employeeTotal) || 0,

    Number(data.employeeCash) || 0,

    Number(data.employeeQr) || 0,

    Number(data.simSales) || 0,

    employeeNames,

    imageUrl

  ]);


  return {
    success: true,
    message:
      'บันทึกรายงานเรียบร้อยแล้ว',
    imageUrl: imageUrl
  };
}


/**
 * VALIDATE
 */
function validateData(data) {

  if (!data.shop) {
    throw new Error(
      'กรุณาเลือกร้าน'
    );
  }

  if (!data.date) {
    throw new Error(
      'กรุณาเลือกวันที่'
    );
  }

  if (!data.shift) {
    throw new Error(
      'กรุณาเลือกกะ'
    );
  }

  if (
    !Array.isArray(data.employees) ||
    data.employees.length === 0
  ) {
    throw new Error(
      'กรุณาเลือกพนักงานอย่างน้อย 1 คน'
    );
  }


  /**
   * SALES
   */
  const totalSales =
    Number(data.totalSales) || 0;

  const cash =
    Number(data.cash) || 0;

  const credit =
    Number(data.credit) || 0;

  const qr =
    Number(data.qr) || 0;


  const salesSum =
    cash +
    credit +
    qr;


  if (
    Math.abs(
      totalSales - salesSum
    ) > 0.01
  ) {

    throw new Error(
      'ยอดขายไม่ตรงกัน'
    );

  }


  /**
   * EMPLOYEE SALES
   */
  const employeeTotal =
    Number(data.employeeTotal) || 0;

  const employeeCash =
    Number(data.employeeCash) || 0;

  const employeeQr =
    Number(data.employeeQr) || 0;


  const employeeSum =
    employeeCash +
    employeeQr;


  if (
    Math.abs(
      employeeTotal -
      employeeSum
    ) > 0.01
  ) {

    throw new Error(
      'ยอดพนักงานซื้อไม่ตรงกัน'
    );

  }

}


/**
 * SAVE IMAGE
 */
function saveImageToDrive(
  base64,
  fileName,
  mimeType,
  shop,
  date
) {

  const folder =
    DriveApp.getFolderById(
      CONFIG.DRIVE_FOLDER_ID
    );


  const cleanBase64 =
    base64.replace(
      /^data:[^;]+;base64,/,
      ''
    );


  const bytes =
    Utilities.base64Decode(
      cleanBase64
    );


  const blob =
    Utilities.newBlob(
      bytes,
      mimeType,
      fileName
    );


  const timestamp =
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyyMMdd_HHmmss'
    );


  const finalName =
    timestamp +
    '_' +
    shop +
    '_' +
    date +
    '_' +
    fileName;


  const file =
    folder.createFile(blob);


  file.setName(finalName);


  return file.getUrl();
}
