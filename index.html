/****************************************************
 * LINE LIFF SALES REPORT
 * Google Apps Script + Google Sheets + Google Drive
 ****************************************************/

const CONFIG = {

  SPREADSHEET_ID:
    '1XZ8tlL8aLcgCCdgWb1KPEO4ZI1Mf9t-BrXXeqmeXp90',

  SALES_SHEET:
    'รายงานยอดขาย',

  // ROOT FOLDER
  DRIVE_FOLDER_ID:
    '1fNe5bMW9N1SfSTVdLgFJLPAlRCuvh2ft',

  SHOPS: [
    'M1',
    'M85',
    'M95',
    'M89'
  ],

  SHIFTS: [
    '09.00-14.00',
    '14.00-18.00',
    '18.00-23.00'
  ]

};


/****************************************************
 * GET
 *
 * ใช้สำหรับทดสอบ API เท่านั้น
 *
 * ระบบไม่โหลดร้าน/พนักงานจาก Google Sheet แล้ว
 ****************************************************/

function doGet(e) {

  try {

    const action =
      e &&
      e.parameter
        ? e.parameter.action || ''
        : '';

    const callback =
      e &&
      e.parameter
        ? e.parameter.callback || ''
        : '';

    let result;


    if (action === 'test') {

      result = {
        success: true,
        message: 'Apps Script ทำงานปกติ',
        time:
          new Date().toISOString()
      };

    }

    else {

      result = {
        success: false,
        message: 'Unknown action',
        receivedAction: action
      };

    }


    if (callback) {

      return ContentService
        .createTextOutput(
          callback +
          '(' +
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


/****************************************************
 * POST
 ****************************************************/

function doPost(e) {

  try {

    if (
      !e ||
      !e.postData ||
      !e.postData.contents
    ) {

      throw new Error(
        'ไม่พบข้อมูล POST'
      );

    }


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


/****************************************************
 * SAVE SALES REPORT
 ****************************************************/

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


  /**************************************************
   * SAVE IMAGE
   **************************************************/

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

        data.date,

        data.shift

      );

  }


  /**************************************************
   * EMPLOYEES
   **************************************************/

  const employeeNames =
    Array.isArray(data.employees)
      ? data.employees.join(', ')
      : '';


  /**************************************************
   * SHIFT
   **************************************************/

  const shiftText =
    normalizeShifts(
      data.shift
    );


  /**************************************************
   * SAVE SHEET
   **************************************************/

  sheet.appendRow([

    new Date(),

    data.lineUserId || '',

    data.lineName || '',

    data.shop || '',

    data.date || '',

    shiftText,

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

    imageUrl:
      imageUrl

  };

}


/****************************************************
 * VALIDATE
 ****************************************************/

function validateData(data) {

  /**************************************************
   * SHOP
   **************************************************/

  if (!data.shop) {

    throw new Error(
      'ไม่พบร้าน'
    );

  }


  if (
    CONFIG.SHOPS.indexOf(
      String(data.shop).trim()
    ) === -1
  ) {

    throw new Error(
      'ร้านไม่ถูกต้อง: ' +
      data.shop
    );

  }


  /**************************************************
   * DATE
   **************************************************/

  if (!data.date) {

    throw new Error(
      'กรุณาเลือกวันที่'
    );

  }


  /**************************************************
   * SHIFT
   **************************************************/

  if (!data.shift) {

    throw new Error(
      'กรุณาเลือกกะอย่างน้อย 1 กะ'
    );

  }


  const shifts =
    parseShifts(
      data.shift
    );


  if (!shifts.length) {

    throw new Error(
      'ไม่พบกะที่เลือก'
    );

  }


  const invalidShifts =
    shifts.filter(
      function(shift) {

        return CONFIG.SHIFTS.indexOf(
          shift
        ) === -1;

      }
    );


  if (invalidShifts.length) {

    throw new Error(
      'พบกะไม่ถูกต้อง: ' +
      invalidShifts.join(', ')
    );

  }


  /**************************************************
   * EMPLOYEES
   **************************************************/

  if (
    !Array.isArray(data.employees) ||
    data.employees.length === 0
  ) {

    throw new Error(
      'กรุณาเลือกพนักงานอย่างน้อย 1 คน'
    );

  }


  /**************************************************
   * SALES
   **************************************************/

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
      totalSales -
      salesSum
    ) > 0.01
  ) {

    throw new Error(
      'ยอดขายไม่ตรงกัน: ' +
      'รวมยอดขาย ' +
      totalSales +
      ' ≠ เงินสด + เครดิต + QR ' +
      salesSum
    );

  }


  /**************************************************
   * EMPLOYEE SALES
   **************************************************/

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
      'ยอดพนักงานซื้อไม่ตรงกัน: ' +
      'รวม ' +
      employeeTotal +
      ' ≠ เงินสด + QR ' +
      employeeSum
    );

  }

}


/****************************************************
 * SHIFT
 ****************************************************/

function parseShifts(shift) {

  if (Array.isArray(shift)) {

    return shift
      .map(function(value) {

        return String(value).trim();

      })
      .filter(Boolean);

  }


  return String(
    shift || ''
  )
    .split(',')
    .map(function(value) {

      return String(value).trim();

    })
    .filter(Boolean);

}


function normalizeShifts(shift) {

  return parseShifts(
    shift
  ).join(', ');

}


/****************************************************
 * SAVE IMAGE
 *
 * ROOT
 * └── 2569
 *     └── 08สิงหาคม
 *
 * ตัวอย่างชื่อ:
 *
 * 20260817_M85_09.00-14.00_14.00-18.00_ใบส่งยอด.jpg
 ****************************************************/

function saveImageToDrive(
  base64,
  fileName,
  mimeType,
  shop,
  date,
  shift
) {

  const rootFolder =
    DriveApp.getFolderById(
      CONFIG.DRIVE_FOLDER_ID
    );


  const dateObj =
    parseReportDate(
      date
    );


  if (!dateObj) {

    throw new Error(
      'รูปแบบวันที่ไม่ถูกต้อง: ' +
      date
    );

  }


  /**************************************************
   * พ.ศ.
   **************************************************/

  const buddhistYear =
    dateObj.getFullYear() +
    543;


  /**************************************************
   * เดือน
   **************************************************/

  const month =
    dateObj.getMonth();


  const monthNames = [

    'มกราคม',
    'กุมภาพันธ์',
    'มีนาคม',
    'เมษายน',
    'พฤษภาคม',
    'มิถุนายน',
    'กรกฎาคม',
    'สิงหาคม',
    'กันยายน',
    'ตุลาคม',
    'พฤศจิกายน',
    'ธันวาคม'

  ];


  const monthNumber =
    String(
      month + 1
    ).padStart(
      2,
      '0'
    );


  const monthFolderName =
    monthNumber +
    monthNames[month];


  /**************************************************
   * YEAR FOLDER
   **************************************************/

  const yearFolder =
    getOrCreateFolder(
      rootFolder,
      String(buddhistYear)
    );


  /**************************************************
   * MONTH FOLDER
   **************************************************/

  const monthFolder =
    getOrCreateFolder(
      yearFolder,
      monthFolderName
    );


  /**************************************************
   * BASE64
   **************************************************/

  const cleanBase64 =
    String(base64)
      .replace(
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
      mimeType || 'image/jpeg',
      fileName || 'ใบส่งยอด.jpg'
    );


  /**************************************************
   * FILE NAME
   **************************************************/

  const yyyy =
    dateObj.getFullYear();


  const mm =
    String(
      dateObj.getMonth() + 1
    ).padStart(
      2,
      '0'
    );


  const dd =
    String(
      dateObj.getDate()
    ).padStart(
      2,
      '0'
    );


  const shiftName =
    parseShifts(
      shift
    ).join('_');


  const cleanShop =
    sanitizeFileName(
      shop ||
      'ไม่ระบุร้าน'
    );


  const cleanOriginalName =
    sanitizeFileName(
      fileName ||
      'ใบส่งยอด.jpg'
    );


  let finalName =
    yyyy +
    mm +
    dd +
    '_' +
    cleanShop;


  if (shiftName) {

    finalName +=
      '_' +
      shiftName;

  }


  finalName +=
    '_' +
    cleanOriginalName;


  /**************************************************
   * CREATE FILE
   **************************************************/

  const file =
    monthFolder.createFile(
      blob
    );


  file.setName(
    finalName
  );


  return file.getUrl();

}


/****************************************************
 * PARSE DATE
 ****************************************************/

function parseReportDate(date) {

  if (
    Object.prototype
      .toString
      .call(date) ===
      '[object Date]'
  ) {

    return date;

  }


  const text =
    String(date).trim();


  /**************************************************
   * YYYY-MM-DD
   **************************************************/

  const match =
    text.match(
      /^(\d{4})-(\d{1,2})-(\d{1,2})$/
    );


  if (match) {

    return new Date(

      Number(match[1]),

      Number(match[2]) - 1,

      Number(match[3])

    );

  }


  /**************************************************
   * DD/MM/YYYY
   **************************************************/

  const match2 =
    text.match(
      /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
    );


  if (match2) {

    let year =
      Number(match2[3]);


    if (year > 2400) {

      year -= 543;

    }


    return new Date(

      year,

      Number(match2[2]) - 1,

      Number(match2[1])

    );

  }


  return null;

}


/****************************************************
 * GET / CREATE FOLDER
 ****************************************************/

function getOrCreateFolder(
  parentFolder,
  folderName
) {

  const folders =
    parentFolder.getFoldersByName(
      folderName
    );


  if (folders.hasNext()) {

    return folders.next();

  }


  return parentFolder.createFolder(
    folderName
  );

}


/****************************************************
 * CLEAN FILE NAME
 ****************************************************/

function sanitizeFileName(name) {

  return String(
    name || ''
  )
    .replace(
      /[\\\/:*?"<>|]/g,
      '_'
    )
    .replace(
      /\s+/g,
      ' '
    )
    .trim();

}
