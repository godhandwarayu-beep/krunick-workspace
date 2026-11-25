function authorizeDocs() {
  DocumentApp.create("Test Authorize");
}
const FOLDER_ID = "1G2LXDCBoWUH3VnAFGm2fxaPktGrSseSU";
const SHEET_NAME = "ตารางงาน";
const PLC_SHEET_NAME = "PLC";
const AUTHORIZED_EMAIL = Session.getEffectiveUser().getEmail();

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🌟 ระบบจัดการงาน ว.PA')
    .addItem('1. เริ่มต้นสร้าง/อัปเดตตาราง', 'setupSheets')
    .addItem('2. เปิดระบบจัดการงาน', 'showWebApp')
    .addToUi();
}

function showWebApp() {
  const html = HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('ระบบจัดการงานและเอกสารวิทยฐานะ')
      .setWidth(1280)
      .setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function doGet(e) {
  const userEmail = Session.getActiveUser().getEmail();
  if (userEmail !== AUTHORIZED_EMAIL) {
    return HtmlService.createHtmlOutput("<div style='font-family:Sarabun;text-align:center;margin-top:50px;'><h3>❌ คุณไม่มีสิทธิ์เข้าถึงระบบนี้</h3></div>");
  }
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Krunick Workspace')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  const headers = ["ID", "วันที่", "ปีการศึกษา", "ชื่องาน/โครงการ", "หมวดหมู่ ว.PA", "รายละเอียด", "ลิงก์ไฟล์แนบ", "สถานะ", "สรุปผลการดำเนินงาน", "ภาพที่ 1", "ภาพที่ 2", "ภาพที่ 3", "ภาพที่ 4", "ภาพที่ 5"];
  if(sheet.getMaxColumns() < headers.length) { sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns()); }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#1e293b").setFontColor("white");
  sheet.setFrozenRows(1);

  // 🤝 สร้าง/อัปเดตชีตบันทึกกิจกรรม PLC (ชุมชนแห่งการเรียนรู้ทางวิชาชีพ)
  let plcSheet = ss.getSheetByName(PLC_SHEET_NAME);
  if (!plcSheet) plcSheet = ss.insertSheet(PLC_SHEET_NAME);
  const plcHeaders = ["ID", "วันที่", "ครั้งที่", "ปีการศึกษา", "ภาคเรียน", "ชื่อกลุ่ม/ทีม PLC", "บทบาทของตนเอง", "สมาชิกที่เข้าร่วม", "ประเด็นปัญหา/หัวข้อที่ประชุม", "เป้าหมายการพัฒนา", "กระบวนการ/สรุปการประชุม", "แนวทาง/ข้อตกลงที่ได้", "การนำไปใช้และผลลัพธ์", "จำนวนชั่วโมง", "ภาพที่ 1", "ภาพที่ 2", "ภาพที่ 3", "ภาพที่ 4", "ภาพที่ 5", "ไฟล์แนบ 1", "ไฟล์แนบ 2", "ไฟล์แนบ 3", "ไฟล์แนบ 4", "ไฟล์แนบ 5"];
  if (plcSheet.getMaxColumns() < plcHeaders.length) { plcSheet.insertColumnsAfter(plcSheet.getMaxColumns(), plcHeaders.length - plcSheet.getMaxColumns()); }
  plcSheet.getRange(1, 1, 1, plcHeaders.length).setValues([plcHeaders]).setFontWeight("bold").setBackground("#0f766e").setFontColor("white");
  plcSheet.setFrozenRows(1);

  // 🔥 แอบฝังคำสั่งนี้ไว้ เพื่อบังคับให้ระบบเด้งขอสิทธิ์ Google Docs ทันทีที่รัน setupSheets
  try { DocumentApp.getActiveDocument(); } catch(e) {}
}

function getCalendarEvents() {
  const tasks = getTasks("ทั้งหมด", "ทั้งหมด");
  return tasks.map(t => {
    let dateStr = t['วันที่'];
    if(dateStr.includes('/')) {
       const parts = dateStr.split('/');
       if(parts.length === 3) dateStr = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
    }
    return {
      id: t['ID'], title: t['ชื่องาน/โครงการ'], start: dateStr, color: t['สถานะ'] === 'เสร็จสิ้น' ? '#10b981' : '#f59e0b', allDay: true,
      extendedProps: { category: t['หมวดหมู่ ว.PA'], description: t['รายละเอียด'], academicYear: t['ปีการศึกษา'], fileUrl: t['ลิงก์ไฟล์แนบ'], status: t['สถานะ'] }
    };
  });
}

function getTasks(yearFilter, catFilter) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0];
  let results = [];
  for(let i=data.length-1; i>=1; i--) {
    if(yearFilter && yearFilter !== "ทั้งหมด" && data[i][2] !== yearFilter) continue;
    if(catFilter && catFilter !== "ทั้งหมด" && data[i][4] !== catFilter) continue;
    let obj = {};
    headers.forEach((h, idx) => obj[h] = data[i][idx]);
    results.push(obj);
  }
  return results;
}

function processForm(formObject) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const folder = DriveApp.getFolderById(FOLDER_ID);

    let existingUrls = [formObject.exFile1 || "", formObject.exFile2 || "", formObject.exFile3 || "", formObject.exFile4 || "", formObject.exFile5 || ""];
    let files = [formObject.file1, formObject.file2, formObject.file3, formObject.file4, formObject.file5];

    for (let i = 0; i < 5; i++) { if (files[i] && files[i].name) { existingUrls[i] = folder.createFile(files[i]).getUrl(); } }
    let combinedUrl = existingUrls.join(" | ");

    if (formObject.taskId) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === formObject.taskId) {
          sheet.getRange(i + 1, 2, 1, 6).setValues([[formObject.date, formObject.academicYear, formObject.taskName, formObject.category, formObject.description, combinedUrl]]);
          return "อัปเดตข้อมูลสำเร็จ";
        }
      }
    }
    const newId = "PA-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmmss");
    sheet.appendRow([newId, formObject.date, formObject.academicYear, formObject.taskName, formObject.category, formObject.description, combinedUrl, "รอดำเนินการ", "", "", "", "", "", ""]);
    return "บันทึกข้อมูลงานสำเร็จ";
  } catch(e) { throw new Error(e.message); }
}

function saveTaskReport(formObject) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const folder = DriveApp.getFolderById(FOLDER_ID);
    let imgs = [formObject.existingImg1, formObject.existingImg2, formObject.existingImg3, formObject.existingImg4, formObject.existingImg5];
    const files = [formObject.img1, formObject.img2, formObject.img3, formObject.img4, formObject.img5];
    for (let j = 0; j < 5; j++) { if (files[j] && files[j].name) imgs[j] = folder.createFile(files[j]).getUrl(); }
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === formObject.reportTaskId) {
        sheet.getRange(i + 1, 8, 1, 7).setValues([["เสร็จสิ้น", formObject.reportSummary, imgs[0], imgs[1], imgs[2], imgs[3], imgs[4]]]);
        return "บันทึกรายงานสำเร็จ";
      }
    }
  } catch(e) { throw new Error(e.message); }
}

function deleteTask(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (data[i][0] === id) { sheet.deleteRow(i + 1); return "ลบข้อมูลเรียบร้อย"; } }
}

function getYears() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  let years = new Set();
  for(let i=1; i<data.length; i++) if(data[i][2]) years.add(data[i][2].toString());
  return Array.from(years).sort().reverse();
}

function getMultipleBase64FromDriveUrls(urls) {
  let results = [];
  for (let i = 0; i < urls.length; i++) {
    try {
      if (!urls[i]) continue;
      const match = urls[i].match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (!match || !match[1]) continue;
      const fileId = match[1];
      const file = DriveApp.getFileById(fileId);
      const blob = file.getBlob();
      const base64 = Utilities.base64Encode(blob.getBytes());
      const mimeType = blob.getContentType() || 'image/jpeg';
      results.push(`data:${mimeType};base64,${base64}`);
    } catch (e) { console.log("Error fetching image " + i + ": " + e.message); }
  }
  return results;
}

// 🔥 สร้างไฟล์ Google Docs อัตโนมัติ (Phase 2)
function exportToDoc(taskId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    let task = null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === taskId) {
        task = {
          date: data[i][1], year: data[i][2], title: data[i][3], category: data[i][4],
          description: data[i][5], status: data[i][7], summary: data[i][8],
          img1: data[i][9], img2: data[i][10], img3: data[i][11], img4: data[i][12], img5: data[i][13]
        };
        break;
      }
    }

    if (!task) throw new Error("ไม่พบข้อมูลงานในระบบ");

    const docName = "รายงาน ว.PA_" + task.title;
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();

    const titleStyle = {};
    titleStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Sarabun';
    titleStyle[DocumentApp.Attribute.FONT_SIZE] = 18;
    titleStyle[DocumentApp.Attribute.BOLD] = true;

    const normalStyle = {};
    normalStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Sarabun';
    normalStyle[DocumentApp.Attribute.FONT_SIZE] = 16;
    normalStyle[DocumentApp.Attribute.BOLD] = false;

    let header = body.appendParagraph("รายงานสรุปผลการปฏิบัติงาน");
    header.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setAttributes(titleStyle);
    let subHeader = body.appendParagraph("โรงเรียนวัดเชิงท่า | สพป.พระนครศรีอยุธยา เขต 2");
    subHeader.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setAttributes(normalStyle);
    body.appendParagraph("");

    body.appendParagraph("📌 ชื่องาน/โครงการ: " + task.title).setAttributes(titleStyle);
    body.appendParagraph("📅 วันที่: " + task.date + "   |   🎓 ปีการศึกษา: " + task.year).setAttributes(normalStyle);
    body.appendParagraph("📂 หมวดหมู่: " + task.category).setAttributes(normalStyle);
    body.appendParagraph("");

    body.appendParagraph("📝 สรุปผลการดำเนินงาน:").setAttributes(titleStyle);
    body.appendParagraph(task.summary ? task.summary : "-").setAttributes(normalStyle);
    body.appendParagraph("");

    body.appendParagraph("📸 ภาพกิจกรรมประกอบ:").setAttributes(titleStyle);
    body.appendParagraph("");

    const imageUrls = [task.img1, task.img2, task.img3, task.img4, task.img5];
    for (let i = 0; i < imageUrls.length; i++) {
      if (imageUrls[i]) {
        let match = imageUrls[i].match(/\/d\/([a-zA-Z0-9-_]+)/);
        if (match && match[1]) {
          try {
            let file = DriveApp.getFileById(match[1]);
            let blob = file.getBlob();
            let img = body.appendImage(blob);

            let width = img.getWidth();
            let height = img.getHeight();
            if (width > 450) {
              let ratio = height / width;
              img.setWidth(450);
              img.setHeight(450 * ratio);
            }
            body.appendParagraph("");
          } catch(imgErr) {
            body.appendParagraph("[ไม่สามารถแทรกภาพที่ " + (i+1) + " ได้]").setAttributes(normalStyle);
          }
        }
      }
    }

    const targetFolder = DriveApp.getFolderById(FOLDER_ID);
    const docFile = DriveApp.getFileById(doc.getId());
    docFile.moveTo(targetFolder);
    doc.saveAndClose();

    return doc.getUrl();

  } catch (e) {
    throw new Error(e.message);
  }
}

// 🎯 ตั้งเป้าหมายรายด้าน ว.PA (เก็บไว้ต่อผู้ใช้ ด้วย PropertiesService)
function getTargets() {
  const props = PropertiesService.getUserProperties();
  const raw = props.getProperty('CATEGORY_TARGETS');
  if (raw) {
    try { return JSON.parse(raw); } catch(e) {}
  }
  return {
    'ด้านที่ 1 การจัดการเรียนรู้': 0,
    'ด้านที่ 2 การส่งเสริมสนับสนุน': 0,
    'ด้านที่ 3 การพัฒนาตนเอง/วิชาชีพ': 0,
    'งานพิเศษ/หน้าที่อื่น': 0
  };
}

function saveTargets(targets) {
  PropertiesService.getUserProperties().setProperty('CATEGORY_TARGETS', JSON.stringify(targets));
  return 'บันทึกเป้าหมายเรียบร้อย';
}

// 📚 สร้างพอร์ตโฟลิโอรวมทั้งปี (รวมทุกงานที่เสร็จสิ้น จัดกลุ่มตามหมวดหมู่ ว.PA เป็นไฟล์ Word เล่มเดียว)
function exportPortfolioToDoc(yearFilter) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const categories = ['ด้านที่ 1 การจัดการเรียนรู้', 'ด้านที่ 2 การส่งเสริมสนับสนุน', 'ด้านที่ 3 การพัฒนาตนเอง/วิชาชีพ', 'งานพิเศษ/หน้าที่อื่น'];

    let tasksByCategory = {};
    categories.forEach(c => tasksByCategory[c] = []);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = row[7];
      if (status !== 'เสร็จสิ้น') continue;
      const year = row[2] ? row[2].toString() : '';
      if (yearFilter && yearFilter !== 'ทั้งหมด' && year !== yearFilter) continue;
      const cat = row[4] || 'งานพิเศษ/หน้าที่อื่น';
      if (!tasksByCategory[cat]) tasksByCategory[cat] = [];
      tasksByCategory[cat].push({
        date: row[1], year: row[2], title: row[3], category: cat, description: row[5],
        summary: row[8], img1: row[9], img2: row[10], img3: row[11], img4: row[12], img5: row[13]
      });
    }

    const docName = "พอร์ตโฟลิโอ ว.PA" + (yearFilter && yearFilter !== 'ทั้งหมด' ? " ปี " + yearFilter : " ทุกปี");
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();

    const titleStyle = {};
    titleStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Sarabun';
    titleStyle[DocumentApp.Attribute.FONT_SIZE] = 20;
    titleStyle[DocumentApp.Attribute.BOLD] = true;

    const headStyle = {};
    headStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Sarabun';
    headStyle[DocumentApp.Attribute.FONT_SIZE] = 16;
    headStyle[DocumentApp.Attribute.BOLD] = true;

    const subStyle = {};
    subStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Sarabun';
    subStyle[DocumentApp.Attribute.FONT_SIZE] = 14;
    subStyle[DocumentApp.Attribute.BOLD] = true;

    const normalStyle = {};
    normalStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Sarabun';
    normalStyle[DocumentApp.Attribute.FONT_SIZE] = 14;
    normalStyle[DocumentApp.Attribute.BOLD] = false;

    let header = body.appendParagraph("พอร์ตโฟลิโอสรุปผลการปฏิบัติงาน ว.PA");
    header.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setAttributes(titleStyle);
    let sub = body.appendParagraph("โรงเรียนวัดเชิงท่า | สพป.พระนครศรีอยุธยา เขต 2" + (yearFilter && yearFilter !== 'ทั้งหมด' ? " | ปีการศึกษา " + yearFilter : ""));
    sub.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setAttributes(normalStyle);
    body.appendParagraph("");
    body.appendHorizontalRule();

    categories.forEach(cat => {
      const list = tasksByCategory[cat];
      body.appendParagraph("");
      let catHeader = body.appendParagraph("📂 " + cat + " (" + list.length + " รายการ)");
      catHeader.setAttributes(headStyle);

      if (list.length === 0) {
        body.appendParagraph("ยังไม่มีงานที่เสร็จสิ้นในหมวดนี้").setAttributes(normalStyle);
        body.appendHorizontalRule();
        return;
      }

      list.forEach((task, idx) => {
        let dateStr = task.date;
        if (task.date instanceof Date) {
          try { dateStr = Utilities.formatDate(task.date, Session.getScriptTimeZone(), "dd/MM/yyyy"); } catch(e) {}
        }
        body.appendParagraph("");
        body.appendParagraph((idx + 1) + ". " + task.title).setAttributes(subStyle);
        body.appendParagraph("📅 วันที่: " + dateStr + "   |   🎓 ปีการศึกษา: " + task.year).setAttributes(normalStyle);
        body.appendParagraph("📝 สรุปผล: " + (task.summary ? task.summary : "-")).setAttributes(normalStyle);

        const imageUrls = [task.img1, task.img2, task.img3, task.img4, task.img5];
        for (let i = 0; i < imageUrls.length; i++) {
          if (imageUrls[i]) {
            let match = imageUrls[i].match(/\/d\/([a-zA-Z0-9-_]+)/);
            if (match && match[1]) {
              try {
                let file = DriveApp.getFileById(match[1]);
                let blob = file.getBlob();
                let img = body.appendImage(blob);
                let width = img.getWidth();
                let height = img.getHeight();
                if (width > 300) {
                  let ratio = height / width;
                  img.setWidth(300);
                  img.setHeight(300 * ratio);
                }
              } catch (imgErr) {}
            }
          }
        }
        body.appendParagraph("");
      });
      body.appendHorizontalRule();
    });

    const targetFolder = DriveApp.getFolderById(FOLDER_ID);
    const docFile = DriveApp.getFileById(doc.getId());
    docFile.moveTo(targetFolder);
    doc.saveAndClose();

    return doc.getUrl();
  } catch (e) {
    throw new Error(e.message);
  }
}

// ============================================================
// 🤝 ระบบบันทึกกิจกรรม PLC (ชุมชนแห่งการเรียนรู้ทางวิชาชีพ)
// ============================================================

function getPLCRecords(yearFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PLC_SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return [];
  const headers = data[0];
  let results = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (yearFilter && yearFilter !== "ทั้งหมด" && data[i][3] !== yearFilter) continue;
    let obj = {};
    headers.forEach((h, idx) => obj[h] = data[i][idx]);
    results.push(obj);
  }
  return results;
}

function getPLCYears() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PLC_SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  let years = new Set();
  for (let i = 1; i < data.length; i++) if (data[i][3]) years.add(data[i][3].toString());
  return Array.from(years).sort().reverse();
}

function savePLCRecord(formObject) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(PLC_SHEET_NAME);
    if (!sheet) { setupSheets(); sheet = ss.getSheetByName(PLC_SHEET_NAME); }
    const data = sheet.getDataRange().getValues();
    const folder = DriveApp.getFolderById(FOLDER_ID);

    let existingImgUrls = [formObject.exPlcImg1 || "", formObject.exPlcImg2 || "", formObject.exPlcImg3 || "", formObject.exPlcImg4 || "", formObject.exPlcImg5 || ""];
    let imgFiles = [formObject.plcImg1, formObject.plcImg2, formObject.plcImg3, formObject.plcImg4, formObject.plcImg5];
    for (let i = 0; i < 5; i++) { if (imgFiles[i] && imgFiles[i].name) { existingImgUrls[i] = folder.createFile(imgFiles[i]).getUrl(); } }

    let existingUrls = [formObject.exPlcFile1 || "", formObject.exPlcFile2 || "", formObject.exPlcFile3 || "", formObject.exPlcFile4 || "", formObject.exPlcFile5 || ""];
    let files = [formObject.plcFile1, formObject.plcFile2, formObject.plcFile3, formObject.plcFile4, formObject.plcFile5];
    for (let i = 0; i < 5; i++) { if (files[i] && files[i].name) { existingUrls[i] = folder.createFile(files[i]).getUrl(); } }

    const rowValues = [
      formObject.plcDate, formObject.plcSession, formObject.plcYear, formObject.plcSemester,
      formObject.plcGroup, formObject.plcRole, formObject.plcMembers, formObject.plcTopic,
      formObject.plcGoal, formObject.plcProcess, formObject.plcSolution, formObject.plcResult,
      formObject.plcHours,
      existingImgUrls[0], existingImgUrls[1], existingImgUrls[2], existingImgUrls[3], existingImgUrls[4],
      existingUrls[0], existingUrls[1], existingUrls[2], existingUrls[3], existingUrls[4]
    ];

    if (formObject.plcId) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === formObject.plcId) {
          sheet.getRange(i + 1, 2, 1, rowValues.length).setValues([rowValues]);
          return "อัปเดตบันทึก PLC สำเร็จ";
        }
      }
    }
    const newId = "PLC-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmmss");
    sheet.appendRow([newId].concat(rowValues));
    return "บันทึกกิจกรรม PLC สำเร็จ";
  } catch (e) { throw new Error(e.message); }
}

function deletePLCRecord(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PLC_SHEET_NAME);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (data[i][0] === id) { sheet.deleteRow(i + 1); return "ลบบันทึก PLC เรียบร้อย"; } }
}

function getPLCTarget() {
  const raw = PropertiesService.getUserProperties().getProperty('PLC_HOUR_TARGET');
  return raw ? parseFloat(raw) : 50;
}

function savePLCTarget(target) {
  PropertiesService.getUserProperties().setProperty('PLC_HOUR_TARGET', String(target));
  return 'บันทึกเป้าหมายชั่วโมง PLC เรียบร้อย';
}

// 📚 สร้างรายงานสรุปกิจกรรม PLC รวม (Word) เพื่อใช้เป็นหลักฐานประกอบการส่งผลงาน
function exportPLCToDoc(yearFilter) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PLC_SHEET_NAME);
    if (!sheet) throw new Error("ยังไม่มีข้อมูล PLC ในระบบ");
    const data = sheet.getDataRange().getValues();
    let records = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const year = row[3] ? row[3].toString() : '';
      if (yearFilter && yearFilter !== 'ทั้งหมด' && year !== yearFilter) continue;
      records.push({
        id: row[0], date: row[1], session: row[2], year: row[3], semester: row[4],
        group: row[5], role: row[6], members: row[7], topic: row[8], goal: row[9],
        process: row[10], solution: row[11], result: row[12], hours: row[13],
        img1: row[14], img2: row[15], img3: row[16], img4: row[17], img5: row[18],
        file1: row[19], file2: row[20], file3: row[21], file4: row[22], file5: row[23]
      });
    }
    if (records.length === 0) throw new Error("ไม่พบข้อมูล PLC ในปีที่เลือก");

    records.sort(function(a, b) {
      const da = a.date instanceof Date ? a.date.getTime() : 0;
      const db = b.date instanceof Date ? b.date.getTime() : 0;
      return da - db;
    });

    let totalHours = 0;
    records.forEach(r => { totalHours += parseFloat(r.hours) || 0; });

    const docName = "สรุปบันทึก PLC" + (yearFilter && yearFilter !== 'ทั้งหมด' ? " ปี " + yearFilter : " ทุกปี");
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();

    const titleStyle = {};
    titleStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Sarabun';
    titleStyle[DocumentApp.Attribute.FONT_SIZE] = 20;
    titleStyle[DocumentApp.Attribute.BOLD] = true;

    const headStyle = {};
    headStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Sarabun';
    headStyle[DocumentApp.Attribute.FONT_SIZE] = 15;
    headStyle[DocumentApp.Attribute.BOLD] = true;

    const normalStyle = {};
    normalStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Sarabun';
    normalStyle[DocumentApp.Attribute.FONT_SIZE] = 14;
    normalStyle[DocumentApp.Attribute.BOLD] = false;

    const boldInline = {};
    boldInline[DocumentApp.Attribute.FONT_FAMILY] = 'Sarabun';
    boldInline[DocumentApp.Attribute.FONT_SIZE] = 14;
    boldInline[DocumentApp.Attribute.BOLD] = true;

    let header = body.appendParagraph("สรุปบันทึกกิจกรรมชุมชนแห่งการเรียนรู้ทางวิชาชีพ (PLC)");
    header.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setAttributes(titleStyle);
    let sub = body.appendParagraph("โรงเรียนวัดเชิงท่า | สพป.พระนครศรีอยุธยา เขต 2" + (yearFilter && yearFilter !== 'ทั้งหมด' ? " | ปีการศึกษา " + yearFilter : ""));
    sub.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setAttributes(normalStyle);
    let summaryLine = body.appendParagraph("จำนวนครั้งที่ประชุมทั้งหมด " + records.length + " ครั้ง   |   จำนวนชั่วโมงสะสมรวม " + totalHours + " ชั่วโมง");
    summaryLine.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setAttributes(headStyle);
    body.appendParagraph("");

    const tableData = [["ครั้งที่", "วันที่", "หัวข้อ/ประเด็น", "ชั่วโมง"]];
    records.forEach(function(r) {
      let dateStr = r.date instanceof Date ? Utilities.formatDate(r.date, Session.getScriptTimeZone(), "dd/MM/yyyy") : r.date;
      tableData.push([String(r.session || "-"), dateStr, r.topic || "-", String(r.hours || 0)]);
    });
    const table = body.appendTable(tableData);
    table.getRow(0).editAsText().setBold(true);
    for (let r = 0; r < table.getNumRows(); r++) {
      for (let c = 0; c < table.getRow(r).getNumCells(); c++) {
        table.getRow(r).getCell(c).editAsText().setFontFamily('Sarabun').setFontSize(11);
      }
    }
    body.appendParagraph("");
    body.appendHorizontalRule();

    records.forEach(function(r, idx) {
      let dateStr = r.date instanceof Date ? Utilities.formatDate(r.date, Session.getScriptTimeZone(), "dd/MM/yyyy") : r.date;
      body.appendParagraph("");
      body.appendParagraph("ครั้งที่ " + (r.session || (idx + 1)) + "   บันทึกกิจกรรม PLC").setAttributes(headStyle);
      body.appendParagraph("📅 วันที่: " + dateStr + "   |   🎓 ปีการศึกษา: " + r.year + " ภาคเรียนที่ " + r.semester + "   |   ⏱ จำนวนชั่วโมง: " + r.hours).setAttributes(normalStyle);
      body.appendParagraph("👥 กลุ่ม/ทีม PLC: " + (r.group || "-") + "   |   บทบาท: " + (r.role || "-")).setAttributes(normalStyle);
      body.appendParagraph("สมาชิกที่เข้าร่วม: " + (r.members || "-")).setAttributes(normalStyle);
      body.appendParagraph("");
      body.appendParagraph("ประเด็นปัญหา/หัวข้อที่ประชุม").setAttributes(boldInline);
      body.appendParagraph(r.topic || "-").setAttributes(normalStyle);
      body.appendParagraph("เป้าหมายการพัฒนา").setAttributes(boldInline);
      body.appendParagraph(r.goal || "-").setAttributes(normalStyle);
      body.appendParagraph("กระบวนการ/สรุปการประชุม").setAttributes(boldInline);
      body.appendParagraph(r.process || "-").setAttributes(normalStyle);
      body.appendParagraph("แนวทาง/ข้อตกลงที่ได้จากที่ประชุม").setAttributes(boldInline);
      body.appendParagraph(r.solution || "-").setAttributes(normalStyle);
      body.appendParagraph("การนำไปใช้จริงและผลลัพธ์ที่เกิดขึ้น").setAttributes(boldInline);
      body.appendParagraph(r.result || "-").setAttributes(normalStyle);

      const imgs = [r.img1, r.img2, r.img3, r.img4, r.img5];
      for (let i = 0; i < imgs.length; i++) {
        if (imgs[i]) {
          let match = imgs[i].match(/\/d\/([a-zA-Z0-9-_]+)/);
          if (match && match[1]) {
            try {
              let file = DriveApp.getFileById(match[1]);
              let blob = file.getBlob();
              let img = body.appendImage(blob);
              let width = img.getWidth();
              let height = img.getHeight();
              if (width > 300) {
                let ratio = height / width;
                img.setWidth(300);
                img.setHeight(300 * ratio);
              }
            } catch (imgErr) {}
          }
        }
      }

      const attachFiles = [r.file1, r.file2, r.file3, r.file4, r.file5];
      const hasFiles = attachFiles.some(function(u) { return !!u; });
      if (hasFiles) {
        body.appendParagraph("");
        body.appendParagraph("ไฟล์แนบเอกสารประกอบ").setAttributes(boldInline);
        attachFiles.forEach(function(url, idx) {
          if (!url) return;
          let fname = "ไฟล์แนบ " + (idx + 1);
          let match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
          try { if (match && match[1]) fname = DriveApp.getFileById(match[1]).getName(); } catch(e2) {}
          let p = body.appendParagraph("📎 " + fname);
          p.setAttributes(normalStyle);
          try { p.editAsText().setLinkUrl(0, p.getText().length - 1, url); } catch(e3) {}
        });
      }

      body.appendParagraph("");
      body.appendHorizontalRule();
    });

    const targetFolder = DriveApp.getFolderById(FOLDER_ID);
    const docFile = DriveApp.getFileById(doc.getId());
    docFile.moveTo(targetFolder);
    doc.saveAndClose();

    return doc.getUrl();
  } catch (e) {
    throw new Error(e.message);
  }
}
