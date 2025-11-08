/**
 * =============================================================
 * File: Code.gs (Template V3+Libraries Version)
 * (สถานะ: V1.4 "Fix-Validation")
 *
 * นี่คือสคริปต์ "เดียว" ที่ผูกติดกับชีต Template
 *
 * V1.4 (Changelog):
 * - [FIX] แก้ไข `wrapper_runMainBuild`
 * - [FIX] ปรับปรุง `ui.prompt` ให้แสดงกฎที่ถูกต้อง (1-10) และ (600+)
 * - [FIX] แก้ไข "กฎการตรวจสอบ" (Validation) ให้รองรับ 600+
 * (ย้อนกลับจากการใช้ 2-Call ... กลับมาใช้ 1-Call ที่ถูกต้อง)
 *
 * V1.3 (Changelog):
 * - [แก้ไข] เปลี่ยนชื่อเมนูย่อย "สร้างรายชื่อตัวละคร (Run 7.4)"
 * =============================================================
 */

// Global Variable สำหรับชีตนี้
const ss = SpreadsheetApp.getActiveSpreadsheet();
const ui = SpreadsheetApp.getUi();

/**
 * ===========================================
 * สร้างเมนูเมื่อเปิดไฟล์
 * ===========================================
 */
function onOpen() {
  // --- เมนูที่ 1: WTP V3 (WSP Builders) ---
  const menuWTP = ui.createMenu('WTP V3 (Library)');
  
  menuWTP.addItem('▶ [RUN] Build All (Default Mode)', 'wrapper_runFullWorkspace');
  menuWTP.addSeparator();
  menuWTP.addItem('Build ONLY: MAIN (Advanced Split)', 'wrapper_runMainBuild');
  menuWTP.addItem('Build ONLY: QST (Advanced Mode)', 'wrapper_runQstBuild');
  menuWTP.addSeparator();
  menuWTP.addItem('Build ONLY: NPC (Standard)', 'wrapper_runNpcBuild');
  menuWTP.addItem('Build ONLY: ETC (Standard)', 'wrapper_runEtcBuild');
  menuWTP.addToUi();

  // --- เมนูที่ 2: Actor Analyzer (V1.3 แก้ไขชื่อ) ---
  const menuSherlock = ui.createMenu('สร้างรายชื่อตัวละคร');
  menuSherlock.addItem('7.4 FullAuto + CaseFiles', 'wrapper_runSherlock');
  menuSherlock.addSeparator();
  menuSherlock.addItem('ล้างชีต Temp/Audit ของ Sherlock', 'wrapper_runSherlockCleanup');
  menuSherlock.addToUi();
}

/**
 * ===========================================
 * "ตัวกลาง" (Wrapper) สำหรับ WTP V3
 * ===========================================
 */

// 1. ตัวกลางสำหรับ Build All
function wrapper_runFullWorkspace() {
  ui.alert('WTP V3 (Library)', 'กระบวนการ "Build All" จะทำการรันทั้ง 4 โมดูลด้วย "โหมดพื้นฐาน" (Default Mode) โดยอัตโนมัติ (MAIN และ QST จะใช้โหมด 1)\n\nกด "OK" เพื่อเริ่มทำงาน', ui.ButtonSet.OK);
  
  // เรียกไลบรารี (API ที่ 1)
  const result = TranslationEngine.runFullWorkspace(ss);
  
  if (result.success) {
    ui.alert('WTP V3 (Library)', 'กระบวนการ "Build All" เสร็จสมบูรณ์!', ui.ButtonSet.OK);
  } else {
    ui.alert('WTP V3 (Library)', `เกิดข้อผิดพลาด: ${result.error}`, ui.ButtonSet.OK);
  }
}

// 2. ตัวกลางสำหรับ MAIN (V1.4: "Smart 1-Call" Wrapper)
function wrapper_runMainBuild() {
  
  // --- (V1.4) "1-Call" แบบที่ฉลาดขึ้น (Smart 1-Call) ---
  // (เราไม่รู้จำนวนแถวล่วงหน้า แต่เรา "รู้" กฎที่ถูกต้อง)
  
  let promptTitle = `Build ONLY: MAIN (โหมด 1-Call)`;
  let promptMessage = `กรุณาป้อน "กฎ" การแบ่ง:\n` +
                   `(ระบบไม่ทราบจำนวนแถวล่วงหน้าในโหมดนี้)\n\n` +
                   `• ป้อน '1' = ไม่แบ่ง (สร้าง 1 ชีต)\n` +
                   `• ป้อน '2' - '10' = "หาร N ชิ้น"\n` +
                   `• ป้อน '600' ขึ้นไป = "แบ่งชิ้นละ N แถว"`;

  const prompt = ui.prompt(promptTitle, promptMessage, ui.ButtonSet.OK_CANCEL);
  
  if (prompt.getSelectedButton() == ui.Button.OK) {
    const splitInputText = prompt.getResponseText();
    const splitInput = parseInt(splitInputText);
    
    // --- (V1.4) "กฎการตรวจสอบ" ใหม่ ที่รองรับ 600+ ---
    // (นี่คือจุดที่แก้ไขปัญหาของนายท่านค่ะ)
    if (isNaN(splitInput) || splitInput < 1) {
      ui.alert('ข้อมูลไม่ถูกต้อง', 'กรุณาใส่ตัวเลขที่มากกว่า 0 เท่านั้น', ui.ButtonSet.OK);
      return;
    }
    
    // (กฎสำคัญที่แก้ไขให้ถูกต้องแล้ว)
    if (splitInput > 10 && splitInput < 600) {
      ui.alert('ข้อมูลไม่ถูกต้อง', `คุณป้อน "${splitInputText}"\n\nกฎการแบ่งแบบ "ชิ้นละ N แถว" จะต้องมีค่า "600" ขึ้นไปค่ะ`, ui.ButtonSet.OK);
      return;
    }
    
    // --- "ทำงานจริง" (Execute) ---
    ui.alert('WTP V3 (Library)', `กำลังเรียกไลบรารี (Call 1) ด้วยกฎ "${splitInputText}"...\n\n(กระบวนการนี้อาจใช้เวลา 1-2 นาที)`, ui.ButtonSet.OK);
    
    // (เรียก API ที่ 2 ตัวเดิม ... แต่ส่ง "กฎ" ที่ฉลาดแล้วเข้าไป)
    const result = TranslationEngine.runMainBuild(ss, splitInput);
    
    if (result.success) {
      ui.alert('WTP V3 (Library)', 'Build MAIN เสร็จสมบูรณ์!', ui.ButtonSet.OK);
    } else {
      ui.alert('WTP V3 (Library)', `เกิดข้อผิดพลาด: ${result.error}`, ui.ButtonSet.OK);
    }
  }
}

// 3. ตัวกลางสำหรับ QST (มี Prompt)
function wrapper_runQstBuild() {
  const prompt = ui.prompt(
    'Build ONLY: QST (Advanced Mode)',
    'กรุณาเลือกโหมดการแบ่งชีต QST:\n\n[ 1 ] = โหมดมาตรฐาน (Standard)\n[ 2 ] = โหมดแบ่งย่อย (Fragmented)\n[ 3 ] = โหมด (Smart Splitting)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (prompt.getSelectedButton() == ui.Button.OK) {
    const splitMode = parseInt(prompt.getResponseText());
    if (![1, 2, 3].includes(splitMode)) {
      ui.alert('ข้อมูลไม่ถูกต้อง', 'กรุณาใส่ตัวเลข 1, 2, หรือ 3 เท่านั้น', ui.ButtonSet.OK);
      return;
    }
    
    ui.alert('WTP V3 (Library)', `กำลังเริ่ม Build QST (โหมด ${splitMode})...`, ui.ButtonSet.OK);
    
    // เรียกไลบรารี (API ที่ 2)
    const result = TranslationEngine.runQstBuild(ss, splitMode);
    
    if (result.success) {
      ui.alert('WTP V3 (Library)', 'Build QST เสร็จสมบูรณ์!', ui.ButtonSet.OK);
    } else {
      ui.alert('WTP V3 (Library)', `เกิดข้อผิดพลาด: ${result.error}`, ui.ButtonSet.OK);
    }
  }
}

// 4. ตัวกลางสำหรับ NPC (ไม่มี Prompt)
function wrapper_runNpcBuild() {
  ui.alert('WTP V3 (Library)', 'กำลังเริ่ม Build NPC (โหมด Standard)...', ui.ButtonSet.OK);
  
  // เรียกไลบรารี (API ที่ 2)
  const result = TranslationEngine.runNpcBuild(ss);
  
  if (result.success) {
    ui.alert('WTP V3 (Library)', 'Build NPC เสร็จสมบูรณ์!', ui.ButtonSet.OK);
  } else {
    ui.alert('WTP V3 (Library)', `เกิดข้อผิดพลาด: ${result.error}`, ui.ButtonSet.OK);
  }
}

// 5. ตัวกลางสำหรับ ETC (ไม่มี Prompt)
function wrapper_runEtcBuild() {
  ui.alert('WTP V3 (Library)', 'กำลังเริ่ม Build ETC (โหมด Standard)...', ui.ButtonSet.OK);
  
  // เรียกไลบรารี (API ที่ 2)
  const result = TranslationEngine.runEtcBuild(ss);
  
  if (result.success) {
    ui.alert('WTP V3 (Library)', 'Build ETC เสร็จสมบูรณ์!', ui.ButtonSet.OK);
  } else {
    ui.alert('WTP V3 (Library)', `เกิดข้อผิดพลาด: ${result.error}`, ui.ButtonSet.OK);
  }
}


/**
 * ===========================================
 * "ตัวกลาง" (Wrapper) สำหรับ Project Sherlock
 * ===========================================
 */

// 6. ตัวกลางสำหรับ Sherlock Build
function wrapper_runSherlock() {
  ui.alert('Project Sherlock (Library)', 'กระบวนการนี้จะทำการวิเคราะห์ชื่อตัวละคร (Actor Names) จากชีตข้อมูลดิบ (JP, US, ENJP) และสร้างผลลัพธ์ลงในชีต "Names_JP_TH_EN" และ "Sherlock_Case_Files"\n\nกด "OK" เพื่อเริ่มทำงาน', ui.ButtonSet.OK);
  
  // เรียกไลบรารี (API ที่ 3)
  const result = TranslationEngine.runSherlockBuild(ss);
  
  if (result.success) {
    // (แสดงผลลัพธ์ที่ Sherlock ส่งกลับมา)
    ui.alert('Project Sherlock (Library)', result.message, ui.ButtonSet.OK);
  } else {
    ui.alert('Project Sherlock (Library)', `เกิดข้อผิดพลาด: ${result.error}`, ui.ButtonSet.OK);
  }
}

// 7. ตัวกลางสำหรับ Sherlock Cleanup
function wrapper_runSherlockCleanup() {
  const confirm = ui.alert('Project Sherlock (Library)', 'คุณแน่ใจหรือไม่ว่าต้องการ "ลบ" ชีต Temp และ Audit ทั้งหมดของ Sherlock? (ยกเว้น Case Files)\n\nชีตที่ถูกลบไปแล้วจะไม่สามารถกู้คืนได้', ui.ButtonSet.OK_CANCEL);
  
  if (confirm == ui.Button.OK) {
    // เรียกไลบรารี (API ที่ 3)
    const result = TranslationEngine.runSherlockCleanup(ss);
    
    // (แสดงผลลัพธ์ที่ Sherlock ส่งกลับมา)
    ui.alert('Project Sherlock (Library)', result.message, ui.ButtonSet.OK);
  }
}
