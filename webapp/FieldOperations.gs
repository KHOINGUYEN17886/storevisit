// -*- coding: utf-8 -*-
/**
 * StoreVisit/webapp/FieldOperations.gs
 * ═══════════════════════════════════════════════════════════════
 * MODULE MỞ RỘNG KIỂM TRA THỰC ĐỊA & BÁO CÁO SỰ VỤ THEO TAG
 * (NON-BREAKING ADDITIVE MODULE — HOÀN TOÀN TÁCH BIỆT KHẢO SÁT 7 PHẦN CŨ)
 * ═══════════════════════════════════════════════════════════════
 */

/**
 * Ghi nhận lượt viếng thăm thực địa Top 16 Cửa Hàng Thâm Hụt
 * @param {Object} visitData { asm_name, store_code, visit_date, actions_taken, measured_outcome, photo_url }
 */
function recordFieldVisitLog(visitData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetName = "TB_FIELD_VISIT_LOGS";
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow([
        "ID", "ASM Quản Lý", "Mã Cửa Hàng", "Ngày Thực Địa",
        "Hành Động Tại Chỗ", "Kết Quả Đo Lường", "Link Ảnh Minh Chứng", "Thời Gian Ghi Nhận"
      ]);
      sheet.getRange("A1:H1").setFontWeight("bold").setBackground("#1F3864").setFontColor("#FFFFFF");
    }
    
    var newId = sheet.getLastRow();
    sheet.appendRow([
      newId,
      visitData.asm_name || "",
      visitData.store_code || "",
      visitData.visit_date || Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd"),
      visitData.actions_taken || "",
      visitData.measured_outcome || "",
      visitData.photo_url || "",
      new Date()
    ]);
    
    return { success: true, message: "Đã lưu nhật ký thực địa thành công!", id: newId };
  } catch (err) {
    Logger.log("Lỗi recordFieldVisitLog: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * Ghi nhận Sự Vụ Chất Lượng KCS Theo Tag Động
 * @param {Object} incidentData { store_code, incident_tag, product_code, defect_type, defect_qty, photo_url }
 */
function recordQualityIncident(incidentData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetName = "TB_QUALITY_INCIDENTS";
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow([
        "ID", "Mã Cửa Hàng", "Tag Sự Vụ", "Mã Sản Phẩm",
        "Loại Lỗi KCS", "Số Lượng Lỗi", "Link Ảnh Minh Chứng", "Trạng Thái", "Thời Gian Khai Báo"
      ]);
      sheet.getRange("A1:I1").setFontWeight("bold").setBackground("#800020").setFontColor("#FFFFFF");
    }
    
    var newId = sheet.getLastRow();
    sheet.appendRow([
      newId,
      incidentData.store_code || "",
      incidentData.incident_tag || "GENERAL_DEFECT",
      incidentData.product_code || "",
      incidentData.defect_type || "",
      incidentData.defect_qty || 1,
      incidentData.photo_url || "",
      "OPEN",
      new Date()
    ]);
    
    return { success: true, message: "Đã ghi nhận sự vụ chất lượng thành công!", id: newId };
  } catch (err) {
    Logger.log("Lỗi recordQualityIncident: " + err.toString());
    return { success: false, error: err.toString() };
  }
}
