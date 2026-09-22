/**
 * ==============================================================================
 * TIF Safety Incident Report - Google Apps Script
 * ระบบรับข้อมูลรายงานอุบัติเหตุ/เหตุการณ์ผิดปกติ TIF บันทึกลง Google Sheet และส่งอีเมล
 * ==============================================================================
 * 
 * วิธีติดตั้ง (ทำเพียงครั้งเดียว):
 * 1. เปิด Google Sheet ใหม่ (เช่น https://sheet.new) ตั้งชื่อไฟล์ เช่น "TIF Safety Reports"
 * 2. ไปที่เมนู "ส่วนขยาย" (Extensions) -> "Apps Script"
 * 3. ลบโค้ดเดิมทั้งหมดออก แล้ววางโค้ดทั้งหมดนี้ลงไป
 * 4. กดปุ่ม "บันทึก" (Save - รูปแผ่นดิสก์)
 * 5. กดปุ่มสีน้ำเงิน "ทำให้ใช้งานได้" (Deploy) -> "การทำให้ใช้งานได้รายการใหม่" (New deployment)
 * 6. เลือกประเภทเป็น "เว็บแอป" (Web app)
 *    - คำอธิบาย: TIF Safety Form
 *    - เรียกใช้ในฐานะ (Execute as): ตัวฉัน (Me)
 *    - ผู้ที่มีสิทธิ์เข้าถึง (Who has access): "ทุกคน" (Anyone) **สำคัญมาก ต้องเลือก Anyone**
 * 7. กด "ทำให้ใช้งานได้" (Deploy) -> กดยอมรับสิทธิ์ (Authorize access)
 * 8. คัดลอก "URL เว็บแอป" (Web app URL) ที่ได้ มาวางในไฟล์ index.html ที่ตัวแปร GOOGLE_SCRIPT_URL
 * ==============================================================================
 */

const TARGET_EMAIL = "pratana.pat@tif.ac.th";  // อีเมลปลายทางหลัก
const CC_EMAIL = "";                                // เช่น pratana.pat@tif.ac.th (หากต้องการ CC)

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    let data;
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter;
      }
    } else {
      data = e.parameter;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getActiveSheet();

    // สร้างหัวตารางอัตโนมัติหากเป็นชีทใหม่
    if (sheet.getLastRow() === 0) {
      const headers = [
        "วันเวลาที่ส่งข้อมูล",
        "เลขที่รายงาน (Report No.)",
        "วันเวลาเกิดเหตุ (Date-Time)",
        "สถานที่ (Location)",
        "ผู้พบเหตุ (Reported by)",
        "ตำแหน่ง (Position)",
        "ช่องทางติดต่อ (Contact)",
        "รายละเอียดเหตุการณ์ (Details)",
        "การดำเนินการเบื้องต้น (Initial Action)",
        "ผู้ลงนาม (Signature)",
        "วันที่ลงนาม (Date)"
      ];
      sheet.appendRow(headers);
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#12235a");
      headerRange.setFontColor("#ffffff");
      sheet.setFrozenRows(1);
    }

    const timestamp = Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm:ss");

    // บันทึกข้อมูลลงแถวใหม่
    sheet.appendRow([
      timestamp,
      data.ref || "-",
      data.datetime || "-",
      data.location || "-",
      data.reporter || "-",
      data.position || "-",
      data.contact || "-",
      data.details || "-",
      data.initial_action || "-",
      data.signature || "-",
      data.signature_date || "-"
    ]);

    // สร้างเนื้อหาอีเมล HTML สวยงาม
    const emailSubject = `[TIF Incident Report] ${data.ref} — ${data.location}`;
    const emailHtml = `
      <div style="font-family: 'Segoe UI', Tahoma, sans-serif; max-width: 640px; margin: 0 auto; border: 1px solid #dce4f0; border-radius: 14px; overflow: hidden; box-shadow: 0 4px 20px rgba(0,0,0,0.06); background: #ffffff;">
        <div style="background: linear-gradient(100deg, #12235a 0%, #2f4fd0 50%, #7c4dff 100%); color: #ffffff; padding: 22px 26px;">
          <h2 style="margin: 0; font-size: 18px; letter-spacing: 0.3px;">แบบฟอร์มรายงานอุบัติเหตุ / เหตุการณ์ผิดปกติ</h2>
          <div style="font-size: 12.5px; opacity: 0.9; margin-top: 4px;">Thai Inter Flying Co., Ltd. • Incident Report (${data.ref})</div>
        </div>
        <div style="padding: 22px 26px;">
          <table style="width: 100%; border-collapse: collapse; font-size: 14px; line-height: 1.6;">
            <tr>
              <td style="width: 32%; padding: 9px 0; color: #5f6b7e; font-weight: bold;">เลขที่รายงาน:</td>
              <td style="padding: 9px 0; color: #12235a; font-weight: bold;">${data.ref}</td>
            </tr>
            <tr style="border-top: 1px solid #edf2f7;">
              <td style="padding: 9px 0; color: #5f6b7e; font-weight: bold;">วัน/เวลาเกิดเหตุ:</td>
              <td style="padding: 9px 0; color: #1a2233;">${data.datetime}</td>
            </tr>
            <tr style="border-top: 1px solid #edf2f7;">
              <td style="padding: 9px 0; color: #5f6b7e; font-weight: bold;">สถานที่:</td>
              <td style="padding: 9px 0; color: #d6336c; font-weight: bold;">${data.location}</td>
            </tr>
            <tr style="border-top: 1px solid #edf2f7;">
              <td style="padding: 9px 0; color: #5f6b7e; font-weight: bold;">ผู้พบเหตุ / ผู้รายงาน:</td>
              <td style="padding: 9px 0; color: #1a2233;">${data.reporter} (${data.position || '-'})</td>
            </tr>
            <tr style="border-top: 1px solid #edf2f7;">
              <td style="padding: 9px 0; color: #5f6b7e; font-weight: bold;">ช่องทางติดต่อกลับ:</td>
              <td style="padding: 9px 0; color: #1a2233;">${data.contact || '-'}</td>
            </tr>
            <tr style="border-top: 1px solid #edf2f7;">
              <td style="padding: 9px 0; color: #5f6b7e; font-weight: bold; vertical-align: top;">รายละเอียดเหตุการณ์:</td>
              <td style="padding: 10px; color: #1a2233; background: #f8fafc; border-radius: 8px; white-space: pre-wrap;">${data.details}</td>
            </tr>
            <tr style="border-top: 1px solid #edf2f7;">
              <td style="padding: 9px 0; color: #5f6b7e; font-weight: bold; vertical-align: top;">การดำเนินการเบื้องต้น:</td>
              <td style="padding: 9px 0; color: #1a2233; white-space: pre-wrap;">${data.initial_action || '-'}</td>
            </tr>
            <tr style="border-top: 1px solid #edf2f7;">
              <td style="padding: 9px 0; color: #5f6b7e; font-weight: bold;">ผู้ลงนาม:</td>
              <td style="padding: 9px 0; color: #1a2233;">${data.signature || '-'} (วันที่ลงนาม: ${data.signature_date || '-'})</td>
            </tr>
          </table>
        </div>
        <div style="background: #f8fafc; border-top: 1px solid #edf2f7; padding: 14px 26px; text-align: center; font-size: 12px; color: #64748b;">
          อีเมลนี้ถูกส่งอัตโนมัติจากระบบรายงานความปลอดภัยออนไลน์ Thai Inter Flying Co., Ltd.
        </div>
      </div>
    `;

    // ส่งอีเมล
    MailApp.sendEmail({
      to: TARGET_EMAIL,
      cc: CC_EMAIL || undefined,
      subject: emailSubject,
      htmlBody: emailHtml
    });

    return ContentService.createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: "error", error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
