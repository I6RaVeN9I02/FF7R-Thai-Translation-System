/**
 * Gemini Chat: 🖥️ FF7 Rebirth Glossary UI Workbench
 * File: Code.gs (Server-Side Logic)
 * (สถานะ: V1.0 "Sidebar Search")
 * (ณ: 2025-11-25 03:25) โดย Gemini
 *
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Translation Engine')
      .addItem('เปิดโต๊ะแปล (Final)', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar_Search')
      .setTitle('Translator Workbench (Final)')
      .setWidth(500); 
  SpreadsheetApp.getUi().showSidebar(html);
}

function getAllData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('DB_Entities');
    var metaSheet = ss.getSheetByName('SYS_Metadata');
    
    // 🛡️ Safety Check 1
    if (!sheet) throw new Error("CRITICAL: ไม่พบชีต 'DB_Entities'");

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    
    // 🛡️ Safety Check 2
    if (values.length < 2) { 
      return { database: [], options: getMetaLists(metaSheet) };
    }

    var headers = values.shift(); 
    var colMap = getColMap(headers); 

    // 🛡️ Safety Check 3
    if (colMap.id === -1) throw new Error("CRITICAL: ไม่พบคอลัมน์ 'Master_ID'");

    var db = values.map(function(row, i) {
      var getVal = function(idx) { 
        return (idx > -1 && row[idx] != null) ? row[idx] : ""; 
      };
      
      // แปลง Date เป็น String ฝั่ง Server เพื่อความชัวร์ในการแสดงผล
      var dateVal = getVal(colMap.lastUp);
      var dateStr = "";
      if (dateVal instanceof Date) {
        dateStr = dateVal.toLocaleString('th-TH'); 
      } else if (dateVal) {
        dateStr = String(dateVal);
      }

      return {
        rowIndex: i + 2,
        id: getVal(colMap.id),
        alias: getVal(colMap.alias),
        lid3: getVal(colMap.lid3),
        
        cat1: getVal(colMap.cat1),
        cat2: getVal(colMap.cat2),
        cat3: getVal(colMap.cat3),
        
        jp: getVal(colMap.jp),
        us: getVal(colMap.us),
        enjp: getVal(colMap.enjp),
        
        th_w: getVal(colMap.th_w),
        th_f: getVal(colMap.th_f),
        th_ai: getVal(colMap.th_ai), // ✅ เพิ่มบรรทัดนี้ครับ (ใส่ต่อจาก th_f ก็ได้)
        status: getVal(colMap.status),
        
        context: getVal(colMap.context),
        genNote: getVal(colMap.genNote),
        
        gender: getVal(colMap.gender),
        age: getVal(colMap.age),
        
        m_f: getVal(colMap.m_f),
        m_ee: getVal(colMap.m_ee),
        m_ti: getVal(colMap.m_ti),
        m_pe: getVal(colMap.m_pe),
        
        lastUp: dateStr
      };
    });

    return { database: db, options: getMetaLists(metaSheet) };

  } catch (e) {
    throw new Error(e.message); 
  }
}

function getMetaLists(metaSheet) {
    var metaLists = { statuses:[], cat1:[], cat2:[], cat3:[], gender:[], age:[] };
    try {
      if (metaSheet) {
          var last = metaSheet.getLastRow();
          if (last > 1) {
              var range = metaSheet.getRange("A2:F" + last).getValues();
              metaLists.statuses = range.map(r => r[0]).filter(String);
              metaLists.cat1 = range.map(r => r[1]).filter(String);
              metaLists.cat2 = range.map(r => r[2]).filter(String);
              metaLists.cat3 = range.map(r => r[3]).filter(String);
              metaLists.gender = range.map(r => r[4]).filter(String);
              metaLists.age = range.map(r => r[5]).filter(String);
          }
      }
    } catch(e) {
      console.warn("Metadata Error: " + e.message);
    }
    return metaLists;
}

function saveTranslation(data) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_Entities');
    if (!sheet) throw new Error("Sheet not found");

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colMap = getColMap(headers);
    
    var r = parseInt(data.rowIndex); 

    // Cell-by-Cell Writing (Safe & Robust)
    sheet.getRange(r, colMap.alias + 1).setValue(data.alias);
    sheet.getRange(r, colMap.cat1 + 1).setValue(data.cat1);
    sheet.getRange(r, colMap.cat2 + 1).setValue(data.cat2);
    sheet.getRange(r, colMap.cat3 + 1).setValue(data.cat3);
    
    sheet.getRange(r, colMap.th_w + 1).setValue(data.th_w);
    sheet.getRange(r, colMap.th_f + 1).setValue(data.th_f);
    sheet.getRange(r, colMap.status + 1).setValue(data.status);
    
    sheet.getRange(r, colMap.context + 1).setValue(data.context);
    sheet.getRange(r, colMap.genNote + 1).setValue(data.genNote);
    
    sheet.getRange(r, colMap.gender + 1).setValue(data.gender);
    sheet.getRange(r, colMap.age + 1).setValue(data.age);
    
    sheet.getRange(r, colMap.m_f + 1).setValue(data.m_f);
    sheet.getRange(r, colMap.m_ee + 1).setValue(data.m_ee);
    sheet.getRange(r, colMap.m_ti + 1).setValue(data.m_ti);
    sheet.getRange(r, colMap.m_pe + 1).setValue(data.m_pe);

    // ✅ บันทึกเวลาที่ส่งมาจาก Client
    sheet.getRange(r, colMap.lastUp + 1).setValue(data.clientTime);

    SpreadsheetApp.flush(); 

    return { success: true };
  } catch (e) {
    throw new Error(e.message);
  }
}

function getColMap(headers) {
  return {
    id: headers.indexOf('Master_ID'),
    alias: headers.indexOf('Alias_ID'),
    lid3: headers.indexOf('3LID'),
    
    cat1: headers.indexOf('Category_1'),
    cat2: headers.indexOf('Category_2'),
    cat3: headers.indexOf('Category_3'),
    
    jp: headers.indexOf('Name_JP'),
    us: headers.indexOf('Name_US'),
    enjp: headers.indexOf('Name_ENJP_Ref'),
    
    th_w: headers.indexOf('Name_TH_Working'),
    th_f: headers.indexOf('Name_TH_Final'),
    th_ai: headers.indexOf('Name_TH_AI'), // ✅ เพิ่มบรรทัดนี้ครับ
    
    context: headers.indexOf('Translator_Context'),
    status: headers.indexOf('Term_Status'),
    genNote: headers.indexOf('General_Notes'),
    
    gender: headers.indexOf('Gender'),
    age: headers.indexOf('Age_Range'),
    
    m_f: headers.indexOf('Metric_F'),
    m_ee: headers.indexOf('Metric_EE'),
    m_ti: headers.indexOf('Metric_TI'),
    m_pe: headers.indexOf('Metric_PE'),
    
    lastUp: headers.indexOf('Last_Updated')
  };
}

// ฟังก์ชันสำหรับไฮไลท์แถวใน Sheet ตามที่ส่งมา
function highlightSheetRow(rowIndex) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_Entities');
    if (!sheet) return;
    
    // ตรวจสอบความถูกต้องของแถว
    var r = parseInt(rowIndex);
    if (isNaN(r) || r < 1) return;

    // สั่งให้ Select ทั้งแถว (ตั้งแต่คอลัมน์ 1 ถึงสุดขอบ)
    sheet.getRange(r, 1, 1, sheet.getLastColumn()).activate();
    
  } catch (e) {
    // เงียบไว้ ไม่ต้องโวยวายถ้า Error (เช่น สับเปลี่ยนเร็วจัดจนหาแถวไม่ทัน)
  }
}


// ฟังก์ชันสำหรับ Sidebar คอยเช็คว่านายท่านเลือกแถวไหนอยู่
function getActiveSelectionId() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_Entities');
    var activeSheet = SpreadsheetApp.getActiveSheet();
    
    // ถ้าไม่ได้อยู่ที่ชีต DB_Entities ไม่ต้องทำอะไร
    if (activeSheet.getName() !== 'DB_Entities') return null;

    var r = activeSheet.getActiveCell().getRow();
    var lastRow = activeSheet.getLastRow();

    // ถ้าคลิกนอกเหนือเขตข้อมูล หรือคลิกหัวตาราง (แถว 1) ให้ข้าม
    if (r < 2 || r > lastRow) return null;

    // ดึงค่า ID จากคอลัมน์ A (Master_ID) ของแถวนั้น
    // สมมติว่า Master_ID อยู่คอลัมน์ 1 เสมอ (เพื่อความเร็วในการประมวลผล)
    return activeSheet.getRange(r, 1).getValue();

  } catch (e) {
    return null;
  }
}
