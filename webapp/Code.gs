function doGet() {
  var template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('Báo cáo Kiểm tra Cửa hàng (StoreVisit Pro)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// -------------------------------------------------------------
// SETUP HELPER - Chạy hàm này 1 lần từ Apps Script Editor
// để cấu hình quyền truy cập cho Python Worker Service Account
// -------------------------------------------------------------
function setupWorkerServiceAccount() {
  var WORKER_EMAIL = "storevisit-bot@avid-vine-499907-g0.iam.gserviceaccount.com";
  PropertiesService.getScriptProperties().setProperty("WORKER_SERVICE_ACCOUNT_EMAIL", WORKER_EMAIL);
  
  // Grant access to existing StoreVisit_Photos folder
  try {
    var spreadsheetFile = DriveApp.getFileById(SPREADSHEET_ID);
    var parentFolders = spreadsheetFile.getParents();
    var parentFolder = parentFolders.hasNext() ? parentFolders.next() : DriveApp.getRootFolder();
    var folderIterator = parentFolder.getFoldersByName("StoreVisit_Photos");
    if (folderIterator.hasNext()) {
      var folder = folderIterator.next();
      folder.addViewer(WORKER_EMAIL);
      Logger.log("✅ Đã cấp quyền viewer cho " + WORKER_EMAIL + " vào folder StoreVisit_Photos (ID: " + folder.getId() + ")");
      
      // Also grant to all existing files in folder
      var files = folder.getFiles();
      var fileCount = 0;
      while (files.hasNext()) {
        var file = files.next();
        try { file.addViewer(WORKER_EMAIL); fileCount++; } catch(e) {}
      }
      Logger.log("✅ Đã cấp quyền cho " + fileCount + " files đã tồn tại.");
    } else {
      Logger.log("⚠️ Chưa tìm thấy folder StoreVisit_Photos. Folder sẽ được tạo tự động khi có ảnh upload đầu tiên.");
    }
  } catch(e) {
    Logger.log("❌ Lỗi cấp quyền folder: " + e.toString());
  }
  
  Logger.log("✅ Script Property WORKER_SERVICE_ACCOUNT_EMAIL đã được cài: " + WORKER_EMAIL);
  Logger.log("🎉 Setup hoàn tất! Python worker có thể tải ảnh từ Google Drive.");
}

// -------------------------------------------------------------
// CONFIGURATION & ENUMS
// -------------------------------------------------------------
var SPREADSHEET_ID = "1Qetn0_6EYEaKQID3Ig4uwX55HhI6Au90Sz8LnVChmEM";
var SHEET_NAME = "Form Responses 1";

var REQUIRED_CHECKLIST_SCHEMA = {
  "frontage": ["F1", "F2", "F3", "F4", "F5"],
  "inner": ["I1", "I2", "I3", "I4", "I5"],
  "merch_ap": ["AP1", "AP2", "AP3", "AP4", "AP5"],
  "merch_pie": ["PIE1", "PIE2", "PIE3", "PIE4", "PIE5"],
  "merch_anamai": ["AN1", "AN2", "AN3", "AN4", "AN5"],
  "merch_bonjour": ["BJ1", "BJ2", "BJ3", "BJ4", "BJ5"],
  "merch_pk": ["PK1", "PK2", "PK3", "PK4", "PK5", "PK6", "PK7", "PK8", "PK9", "PK10", "PK11", "PK12", "PK13", "PK14"],
  "stockroom": ["K1", "K2"],
  "fitting_room": ["K3"],
  "toilet": ["K4"],
  "fire_safety": ["K5"],
  "cashier": ["TN1", "TN2", "TN4"],
  "packaging_security": ["TN3", "TN5"],
  "staff": ["S1", "S2", "S3", "S4"]
};

var SERIOUS_ITEMS = ["K5", "TN1", "TN5"];

var VALID_EVALS = [
  "Chưa kiểm tra",
  "Đạt",
  "Không đạt",
  "Không áp dụng"
];

var VALID_SEVERITIES = [
  "Nhẹ",
  "Trung bình",
  "Nghiêm trọng"
];

var VALID_RESOLVED = [
  "Có",
  "Không"
];

var VALID_RATINGS = [
  "Chưa đánh giá",
  "Tốt",
  "Đạt",
  "Chưa đạt",
  "Không áp dụng"
];

var STORE_MAPPING_SHEET = "StoreMapping";
// Sheet 'StoreMapping' must have header: StoreCode | StoreName | Region | ASM
// Upload store_mapping_clean.csv to this sheet. Do NOT include Department/Level.
// The CSV is already pre-filtered (CHG + Level3 + non-operational excluded).

// NON_STORE_ASM: blocklist for Code.gs dynamic reads (if StoreMapping sheet
// contains raw DimStore data WITH Department & Level columns).
// If sheet has clean CSV (4 cols only), this list still runs as a safety net.
var NON_STORE_ASM = [
  'Closed', 'Event', 'BGD', 'CSKH', 'GiftCard',
  'KDTT', 'KĦ Phân phối', 'KH Phân phối', 'Marketing', 'TLNB',
  'HN',                // Virtual Hà Nội stores not yet operational
  'Thắng', 'Thắng (TT)' // Online/KDTT managers, not field ASMs
];

// NON_STORE_REGION: blocklist for non-geographical region values.
var NON_STORE_REGION = [
  'Closed', 'Event', 'BGD', 'CSKH', 'GiftCard',
  'KDTT', 'KH Phân phối', 'Marketing', 'TLNB',
  'HN'                 // HN as Region = HN stores not yet opened
];

/**
 * getStoreData()
 * Reads the 'StoreMapping' sheet (columns: StoreCode, StoreName, Region, ASM)
 * and returns a structured object for the frontend dropdown.
 *
 * Sheet format expected:
 *   Row 1: StoreCode | StoreName | Region | ASM   (header)
 *   Row 2+: data rows
 *
 * Returns: {
 *   asms: ["Khôi", "Quân", ...],
 *   regions: ["HCM", "Miền Tây", ...],
 *   mapping_by_asm: { "Khôi": ["AEONBD - AEON Bình Dương", ...], ... },
 *   mapping_by_region: { "HCM": [...], ... },
 *   source: "sheet" | "fallback"
 * }
 */
function getStoreData() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get("store_data_json");
    if (cached) {
      return JSON.parse(cached);
    }
  } catch(e) {
    console.warn("Cache read error: " + e.toString());
  }

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(STORE_MAPPING_SHEET);
    
    if (!sheet || sheet.getLastRow() < 2) {
      console.warn('getStoreData: Sheet "' + STORE_MAPPING_SHEET + '" not found or empty. Using fallback.');
      return _buildFallbackStoreData();
    }
    
    var data = sheet.getDataRange().getValues();
    var header = data[0].map(function(h) { return String(h).trim(); });
    var iCode   = header.indexOf('StoreCode');
    var iName   = header.indexOf('StoreName');
    var iRegion = header.indexOf('Region');
    var iAsm    = header.indexOf('ASM');
    // Optional columns (only present if sheet has raw DimStore data)
    var iDept   = header.indexOf('Department');
    var iLevel  = header.indexOf('Level');
    
    if (iCode < 0 || iName < 0 || iRegion < 0 || iAsm < 0) {
      console.warn('getStoreData: Missing required columns in StoreMapping sheet. Using fallback.');
      return _buildFallbackStoreData();
    }
    
    var mapping_by_asm    = {};
    var mapping_by_region = {};
    var asmSet    = {};
    var regionSet = {};
    
    for (var i = 1; i < data.length; i++) {
      var row    = data[i];
      var scode  = String(row[iCode]  || '').trim();
      var sname  = String(row[iName]  || '').trim();
      var region = String(row[iRegion]|| '').trim();
      var asm    = String(row[iAsm]   || '').trim();
      
      if (!scode || !sname || !region || !asm) continue;
      
      // If raw DimStore data (has Department + Level columns): apply CHG/Level3 filter
      if (iDept >= 0 && iLevel >= 0) {
        var dept = String(row[iDept] || '').trim().toUpperCase();
        var lvl  = String(row[iLevel] || '').trim();
        if (dept !== 'CHG' || (lvl !== '3' && lvl !== '3.0')) continue;
      }
      
      // Always exclude non-operational / virtual ASMs and Regions
      if (NON_STORE_ASM.indexOf(asm) >= 0) continue;
      if (NON_STORE_REGION.indexOf(region) >= 0) continue;
      
      var label = scode + ' - ' + sname;
      
      if (!mapping_by_asm[asm])    mapping_by_asm[asm]    = [];
      if (!mapping_by_region[region]) mapping_by_region[region] = [];
      
      // Avoid duplicates
      if (mapping_by_asm[asm].indexOf(label) < 0)
        mapping_by_asm[asm].push(label);
      if (mapping_by_region[region].indexOf(label) < 0)
        mapping_by_region[region].push(label);
      
      asmSet[asm]       = true;
      regionSet[region] = true;
    }
    
    // Sort stores within each ASM/Region
    Object.keys(mapping_by_asm).forEach(function(k)    { mapping_by_asm[k].sort(); });
    Object.keys(mapping_by_region).forEach(function(k) { mapping_by_region[k].sort(); });
    
    var finalAsms = Object.keys(asmSet).sort();
    var finalRegions = Object.keys(regionSet).sort();
    
    // BẮT BỆNH: Nếu sheet có nhưng data bị filter sạch, ép dùng fallback
    if (finalAsms.length === 0) {
      console.warn("Sheet StoreMapping tồn tại nhưng filter không trả về cửa hàng nào. Kích hoạt Fallback map.");
      return _buildFallbackStoreData();
    }
    
    var result = {
      asms:             finalAsms,
      regions:          finalRegions,
      mapping_by_asm:   mapping_by_asm,
      mapping_by_region: mapping_by_region,
      source:           'sheet'
    };

    try {
      var cache = CacheService.getScriptCache();
      cache.put("store_data_json", JSON.stringify(result), 3600);
    } catch(e) {
      console.warn("Cache write error: " + e.toString());
    }

    return result;
    
  } catch (e) {
    console.error('getStoreData error: ' + e.message);
    return _buildFallbackStoreData();
  }
}

/** Internal: build the old hardcoded data as an emergency fallback */
function _buildFallbackStoreData() {
  var asms    = Object.keys(STORE_DATA_MAP.mapping_by_asm);
  var regions = Object.keys(STORE_DATA_MAP.mapping_by_region);
  return {
    asms:             asms,
    regions:          regions,
    mapping_by_asm:   STORE_DATA_MAP.mapping_by_asm,
    mapping_by_region: STORE_DATA_MAP.mapping_by_region,
    source:           'fallback'
  };
}

// Legacy hardcoded map — kept as emergency fallback only.
// Source: DimStore_Final.xlsx filtered CHG + Level3 + exclusion lists applied.
// DO NOT edit manually; run build_store_data_map.py to regenerate.
var STORE_DATA_MAP = {
  "mapping_by_asm": {
      "Tiên": ["AEONBT - AEON Bình Tân", "AUCO - Âu Cơ", "CHOA - Cộng Hòa (Số 8)", "CUCHI - Củ Chi", "DIAMOND - Diamond Plaza", "LVKHUONG - Lê Văn Khương - Q12", "LYTT - Lý Tự Trọng", "NGA6 - 03 Nguyễn Trãi", "NGANHTHU - Nguyễn Ảnh Thủ", "NGHUE - Nguyễn Huệ", "NTQ1 - 74 Nguyễn Trãi", "OIKHIEM - Ông Ích Khiêm", "PDL - Phan Đăng Lưu", "PDL2 - Phan Đăng Lưu 2", "PTER - Pasteur", "TAKA - Takashimaya", "TCHINH2 - Trường Chinh 2 - Q12", "VINCOM - Vincom Q.1"],
      "Tín": ["BCDN - BigC Đồng Nai", "BENLUC - Bến Lức", "BHNAQUOC - Biên Hòa - Nguyễn Ái Quốc", "BIENHOA - Biên Hoà", "BINHLONG - Bình Long", "BPHUOC - Bình Phước", "LONGAN - Long An", "LONGKHANH - Long Khánh", "LONGTHANH - Long Thành", "TAMHIEP - Tam Hiệp", "TAYNINH - Tây Ninh", "THUDUC - Thủ Đức", "THUDUC2 - Thủ Đức 2", "VINCOMBH - VinCom Biên Hòa"],
      "Dũng": ["BAOLOC - Bảo Lộc", "CHOA3 - 454 Cộng Hòa", "CMT8 - Cách Mạng Tháng 8", "DAKLAK - Daklak", "DAKLAK2 - Daklak - Nguyễn Tất Thành", "DAKLAK5 - Buôn Hồ", "DAKNONG - ĐẮK NÔNG (Tôn Đức Thắng)", "DALAT - Đà Lạt", "DALAT2 - Đà Lạt 2", "GIALAI - Gia Lai", "KONTUM - Kon Tum (Trần Hưng Đạo)", "LBBICH - Lũy Bán Bích", "PTHIET - Phan Thiết", "QBINH - Quảng Bình", "TCHINH - Trường Chinh"],
      "Linh": ["126_3T2 - 126 Đường 3T2", "185_3T2 - 185 Đường 3T2", "901QT - 901 Quang Trung", "AEONTP - AEON Tân Phú", "CAOTH - Cao Thắng", "GOVAP - Gò Vấp", "HBT - Hai Bà Trưng", "LETRONGTAN - Lê Trọng Tấn", "LOTTEGV - Lotte Mart Gò Vấp", "LQDINH - Nguyên Hồng", "LVS - Lê Văn Sỹ", "NDC - Nguyễn Đình Chiểu", "NGUYENOANH - Nguyễn Oanh", "NTMK - Nguyễn Thị Minh Khai", "NVTROI - Nguyễn Văn Trỗi", "PDP - Phan Đình Phùng", "SO1 - Số 1"],
      "Quân": ["BIGCHUE - Aeon Huế", "CAMRANH - Cam Ranh", "DN - Đà Nẵng (Hoàng Diệu)", "DN2 - Đà Nẵng 2 (Lê Duẩn)", "DN3 - Đà Nẵng 3 - Nguyễn Văn Linh", "DN4 - Vincom Đà Nẵng", "DN5DBP - Đà Nẵng 5 (Điện Biên Phủ )", "HUE2 - Huế 2 (Hùng Vương)", "NHT2 - Nha Trang 2 - Lý Tự Trọng", "NHT3 - Nha Trang 1 - Thái Nguyên", "PHANRANG2 - Phan Rang 2", "PHUYEN - Phú Yên", "QNGAI - Quảng Ngãi", "QUANGTRI - QUẢNG TRỊ", "QUINHON2 - Quy Nhơn - Phan Bội Châu", "QUYNHON - Quy Nhơn - Trần Hưng Đạo", "QUYNHON3 - Quy Nhơn 2 - Lý Thường Kiệt", "TAMKY - Tam Kỳ"],
      "Hương": ["BACLIEU - Bạc Liêu", "BACLIEU2 - Bạc Liêu 2 - Trần Huỳnh", "BENTRE - Bến Tre", "CANTHO3T2 - Cần Thơ 3T2", "CMAU - Cà Mau", "CMAU2 - Cà Mau 2 - Nguyễn Tất Thành", "CTHO - Cần Thơ (NT)", "CTHO2 - Cần Thơ (LTT) - Mậu Thân", "CTHO3 - Cần Thơ (HB)", "CTHO6 - Cần Thơ (Nguyễn Văn Cừ)", "HAUGIANG - Hậu Giang - Q6", "KINHDV - Kinh Dương Vương", "STR - Sóc Trăng", "STR2 - Sóc Trăng 2 - Trần Hưng Đạo", "TRAVINH - Trà Vinh", "VINHLONG - Vĩnh Long", "VINHLONG2 - Vĩnh Long 2 - Phạm Thái Bường"],
      "Nhi": ["APHAUGIANG - An Phước Vị Thanh", "CAOLANH - Cao Lãnh", "CAYLAY - Cai Lậy", "CHAUDOC - Châu Đốc - An Giang", "FLDLTTON - FLD Lý Thánh Tôn - Nha Trang", "HATIEN - Hà Tiên", "HONGNGU - Hồng Ngự", "LONGXUYEN2 - Long xuyên 2", "LONGXUYEN3 - Long Xuyên 3 - (Trần Hưng Đạo)", "LXUYEN - Long Xuyên 1 - Hai Bà Trưng", "MYTHO - Mỹ Tho", "MYTHO2 - Mỹ Tho 2", "RACHGIA3 - Rạch Giá 3", "RGIA - Rạch Giá", "RGIA2 - Rạch Giá 2", "SADEC - AN PHƯỚC SA ĐÉC"],
      "Khôi": ["AEONBD - AEON Bình Dương", "BARIA - Bà Rịa (Nguyễn Hữu Thọ)", "BDUONG - Bình Dương", "CREMALL - Crescent Mall Q7", "DBTRAC - Dương Bá Trạc", "DIANBD - Dĩ An - Bình Dương", "DLBD - Đại Lộ Bình Dương", "FLDBDUONG - FLD Nguyễn Đình Chiểu BD", "KHANHHOI - Khánh Hội", "LOTTEQ7 - Lotte Mart Q7", "NTT - Nguyễn Thị Thập", "PMHNDC - Phú Mỹ Hưng 1 - Nguyễn Đức Cảnh", "PMHNLB - Phú Mỹ Hưng 2 - Nguyễn Lương Bằng", "PMHNVL - Phú Mỹ Hưng 3 - Nguyễn Văn Linh", "VTAU2 - Vũng Tàu 2 - Ba Cu", "VTAU3004 - Vũng Tàu 3 - Đường 30-04", "VTAU4 - Vũng Tàu 1 - Lê Hồng Phong"],
      "Lâm": ["HVPLAZA - Hùng Vương Plaza", "SENSECITY - Sense City", "VANHANH - Vạn Hạnh Mall", "VINCOMLVV - Vincom Lê Văn Việt", "VINCOMQ2 - Vincom Thảo Điền", "VINCOMTD - Vincom Grand Park"]
  },
  "mapping_by_region": {
      "HCM": ["126_3T2 - 126 Đường 3T2", "185_3T2 - 185 Đường 3T2", "901QT - 901 Quang Trung", "AEONBT - AEON Bình Tân", "AEONTP - AEON Tân Phú", "AUCO - Âu Cơ", "CAOTH - Cao Thắng", "CHOA - Cộng Hòa (Số 8)", "CHOA3 - 454 Cộng Hòa", "CMT8 - Cách Mạng Tháng 8", "CREMALL - Crescent Mall Q7", "CUCHI - Củ Chi", "DBTRAC - Dương Bá Trạc", "DIAMOND - Diamond Plaza", "GOVAP - Gò Vấp", "HAUGIANG - Hậu Giang - Q6", "HBT - Hai Bà Trưng", "HVPLAZA - Hùng Vương Plaza", "KHANHHOI - Khánh Hội", "KINHDV - Kinh Dương Vương", "LBBICH - Lũy Bán Bích", "LETRONGTAN - Lê Trọng Tấn", "LOTTEGV - Lotte Mart Gò Vấp", "LOTTEQ7 - Lotte Mart Q7", "LQDINH - Nguyên Hồng", "LVKHUONG - Lê Văn Khương - Q12", "LVS - Lê Văn Sỹ", "LYTT - Lý Tự Trọng", "NDC - Nguyễn Đình Chiểu", "NGA6 - 03 Nguyễn Trãi", "NGANHTHU - Nguyễn Ảnh Thủ", "NGHUE - Nguyễn Huệ", "NGUYENOANH - Nguyễn Oanh", "NTMK - Nguyễn Thị Minh Khai", "NTQ1 - 74 Nguyễn Trãi", "NTT - Nguyễn Thị Thập", "NVTROI - Nguyễn Văn Trỗi", "OIKHIEM - Ông Ích Khiêm", "PDL - Phan Đăng Lưu", "PDL2 - Phan Đăng Lưu 2", "PDP - Phan Đình Phùng", "PMHNDC - Phú Mỹ Hưng 1 - Nguyễn Đức Cảnh", "PMHNLB - Phú Mỹ Hưng 2 - Nguyễn Lương Bằng", "PMHNVL - Phú Mỹ Hưng 3 - Nguyễn Văn Linh", "PTER - Pasteur", "SENSECITY - Sense City", "SO1 - Số 1", "TAKA - Takashimaya", "TCHINH - Trường Chinh", "TCHINH2 - Trường Chinh 2 - Q12", "THUDUC - Thủ Đức", "THUDUC2 - Thủ Đức 2", "VANHANH - Vạn Hạnh Mall", "VINCOM - Vincom Q.1", "VINCOMLVV - Vincom Lê Văn Việt", "VINCOMQ2 - Vincom Thảo Điền", "VINCOMTD - Vincom Grand Park"],
      "Miền Trung- Tây Nguyên": ["BAOLOC - Bảo Lộc", "BIGCHUE - Aeon Huế", "CAMRANH - Cam Ranh", "DAKLAK - Daklak", "DAKLAK2 - Daklak - Nguyễn Tất Thành", "DAKLAK5 - Buôn Hồ", "DAKNONG - ĐẮK NÔNG (Tôn Đức Thắng)", "DALAT - Đà Lạt", "DALAT2 - Đà Lạt 2", "DN - Đà Nẵng (Hoàng Diệu)", "DN2 - Đà Nẵng 2 (Lê Duẩn)", "DN3 - Đà Nẵng 3 - Nguyễn Văn Linh", "DN4 - Vincom Đà Nẵng", "DN5DBP - Đà Nẵng 5 (Điện Biên Phủ )", "FLDLTTON - FLD Lý Thánh Tôn - Nha Trang", "GIALAI - Gia Lai", "HUE2 - Huế 2 (Hùng Vương)", "KONTUM - Kon Tum (Trần Hưng Đạo)", "NHT2 - Nha Trang 2 - Lý Tự Trọng", "NHT3 - Nha Trang 1 - Thái Nguyên", "PHANRANG2 - Phan Rang 2", "PHUYEN - Phú Yên", "PTHIET - Phan Thiết", "QBINH - Quảng Bình", "QNGAI - Quảng Ngãi", "QUANGTRI - QUẢNG TRỊ", "QUINHON2 - Quy Nhơn - Phan Bội Châu", "QUYNHON - Quy Nhơn - Trần Hưng Đạo", "QUYNHON3 - Quy Nhơn 2 - Lý Thường Kiệt", "TAMKY - Tam Kỳ"],
      "Miền Tây": ["APHAUGIANG - An Phước Vị Thanh", "BACLIEU - Bạc Liêu", "BACLIEU2 - Bạc Liêu 2 - Trần Huỳnh", "BENLUC - Bến Lức", "BENTRE - Bến Tre", "CANTHO3T2 - Cần Thơ 3T2", "CAOLANH - Cao Lãnh", "CAYLAY - Cai Lậy", "CHAUDOC - Châu Đốc - An Giang", "CMAU - Cà Mau", "CMAU2 - Cà Mau 2 - Nguyễn Tất Thành", "CTHO - Cần Thơ (NT)", "CTHO2 - Cần Thơ (LTT) - Mậu Thân", "CTHO3 - Cần Thơ (HB)", "CTHO6 - Cần Thơ (Nguyễn Văn Cừ)", "HATIEN - Hà Tiên", "HONGNGU - Hồng Ngự", "LONGAN - Long An", "LONGXUYEN2 - Long xuyên 2", "LONGXUYEN3 - Long Xuyên 3 - (Trần Hưng Đạo)", "LXUYEN - Long Xuyên 1 - Hai Bà Trưng", "MYTHO - Mỹ Tho", "MYTHO2 - Mỹ Tho 2", "RACHGIA3 - Rạch Giá 3", "RGIA - Rạch Giá", "RGIA2 - Rạch Giá 2", "SADEC - AN PHƯỚC SA ĐÉC", "STR - Sóc Trăng", "STR2 - Sóc Trăng 2 - Trần Hưng Đạo", "TRAVINH - Trà Vinh", "VINHLONG - Vĩnh Long", "VINHLONG2 - Vĩnh Long 2 - Phạm Thái Bường"],
      "Miền Đông": ["AEONBD - AEON Bình Dương", "BARIA - Bà Rịa (Nguyễn Hữu Thọ)", "BCDN - BigC Đồng Nai", "BDUONG - Bình Dương", "BHNAQUOC - Biên Hòa - Nguyễn Ái Quốc", "BIENHOA - Biên Hoà", "BINHLONG - Bình Long", "BPHUOC - Bình Phước", "DIANBD - Dĩ An - Bình Dương", "DLBD - Đại Lộ Bình Dương", "FLDBDUONG - FLD Nguyễn Đình Chiểu BD", "LONGKHANH - Long Khánh", "LONGTHANH - Long Thành", "TAMHIEP - Tam Hiệp", "TAYNINH - Tây Ninh", "VINCOMBH - VinCom Biên Hòa", "VTAU2 - Vũng Tàu 2 - Ba Cu", "VTAU3004 - Vũng Tàu 3 - Đường 30-04", "VTAU4 - Vũng Tàu 1 - Lê Hồng Phong"]
  }
};


function verifyMappingConsistency() {
  var dynamicStoreData = getStoreData();
  var asmStores = [];
  for (var asm in dynamicStoreData.mapping_by_asm) {
    asmStores = asmStores.concat(dynamicStoreData.mapping_by_asm[asm]);
  }
  var regionStores = [];
  for (var r in dynamicStoreData.mapping_by_region) {
    regionStores = regionStores.concat(dynamicStoreData.mapping_by_region[r]);
  }
  
  // Clean duplicates
  asmStores = asmStores.filter(function(item, pos) { return asmStores.indexOf(item) === pos; });
  regionStores = regionStores.filter(function(item, pos) { return regionStores.indexOf(item) === pos; });
  
  asmStores.sort();
  regionStores.sort();
  
  var diff = [];
  asmStores.forEach(function(s) {
    if (!regionStores.includes(s) && !diff.includes(s)) {
      diff.push("Thiếu ở region: " + s);
    }
  });
  regionStores.forEach(function(s) {
    if (!asmStores.includes(s) && !diff.includes(s)) {
      diff.push("Thiếu ở ASM: " + s);
    }
  });
  
  if (diff.length > 0) {
    console.error("Lệch mapping cửa hàng: " + JSON.stringify(diff));
    return { success: false, errors: diff };
  }
  return { success: true };
}

// -------------------------------------------------------------
// HELPERS
// -------------------------------------------------------------
function getOrCreateStorePhotosFolder() {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty("PHOTOS_FOLDER_ID");
  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch(e) {
      console.warn("Cached PHOTOS_FOLDER_ID invalid or inaccessible, recreating: " + e.toString());
    }
  }

  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  var spreadsheetFile = DriveApp.getFileById(sheet.getParent().getId());
  var parentFolders = spreadsheetFile.getParents();
  var parentFolder = parentFolders.hasNext() ? parentFolders.next() : DriveApp.getRootFolder();
  var folderIterator = parentFolder.getFoldersByName("StoreVisit_Photos");
  var folder;
  if (folderIterator.hasNext()) {
    folder = folderIterator.next();
  } else {
    folder = parentFolder.createFolder("StoreVisit_Photos");
  }
  // Share viewer access to worker service account if configured (done once on creation/update)
  grantViewerAccessToWorker(folder);
  
  try {
    props.setProperty("PHOTOS_FOLDER_ID", folder.getId());
  } catch(e) {
    console.warn("Failed to cache PHOTOS_FOLDER_ID: " + e.toString());
  }
  return folder;
}

function validateSubmissionId(id) {
  if (!id || typeof id !== 'string') return false;
  var regex = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
  return regex.test(id);
}

function checkFileExists(fileId) {
  if (!fileId) return false;
  try {
    var file = DriveApp.getFileById(fileId);
    return file !== null;
  } catch(e) {
    return false;
  }
}

function sanitizeFormulaInjection(value) {
  if (typeof value === 'string') {
    var trimmed = value.trim();
    if (trimmed.startsWith('=') || trimmed.startsWith('+') || trimmed.startsWith('-') || trimmed.startsWith('@')) {
      return "'" + value;
    }
  }
  return value;
}

function normalizeHeader(str) {
  if (!str) return "";
  return str.toString()
    .toLowerCase()
    .replace(/đ/g, "d")
    .replace(/[àáảãạăằắẳẵặâầấẩẫậ]/g, "a")
    .replace(/[èéẻẽẹêềếểễệ]/g, "e")
    .replace(/[ìíỉĩị]/g, "i")
    .replace(/[òóỏõọôồốổỗộơờớởỡợ]/g, "o")
    .replace(/[ùúủũụưừứửữự]/g, "u")
    .replace(/[ỳýỷỹỵ]/g, "y")
    .replace(/[^a-z0-9]/g, "")
    .replace(/\s+/g, "")
    .trim();
}

function getFileUrl(fileId) {
  if (!fileId) return "";
  return "https://drive.google.com/open?id=" + fileId;
}

// -------------------------------------------------------------
// SECURE FILE VERIFICATION AND ACCESS SHARING
// -------------------------------------------------------------
function grantViewerAccessToWorker(driveItem) {
  try {
    var email = PropertiesService.getScriptProperties().getProperty("WORKER_SERVICE_ACCOUNT_EMAIL");
    if (email && email.trim()) {
      driveItem.addViewer(email.trim());
    } else {
      console.log("Quyền worker chưa được cấu hình (WORKER_SERVICE_ACCOUNT_EMAIL trống).");
    }
  } catch(e) {
    console.warn("Không thể cấp quyền viewer cho worker: " + e.toString());
  }
}

function isValidUploadSlot(slot) {
  if (!slot || typeof slot !== 'string' || slot.length > 50) return false;
  
  // Strip index suffix if present (e.g. frontage_main_0 -> frontage_main, A1_before_1 -> A1_before)
  var cleanSlot = slot.replace(/_\d+$/, "");
  
  var generalWhitelist = ["frontage_main", "frontage_left", "frontage_right", "inner_entrance", "inner_left", "inner_right", "stockroom", "fitting_room", "cashier"];
  if (generalWhitelist.includes(cleanSlot)) return true;
  
  var competitorWhitelist = ["photoComp1", "photoComp2", "photoComp3", "photoCSVC1"];
  if (competitorWhitelist.includes(cleanSlot)) return true;
  
  var match = cleanSlot.match(/^([A-Z0-9]+)_(before|after)$/);
  if (match) {
    var itemId = match[1];
    for (var secKey in REQUIRED_CHECKLIST_SCHEMA) {
      if (REQUIRED_CHECKLIST_SCHEMA[secKey].includes(itemId)) {
        return true;
      }
    }
  }
  
  return false;
}

function verifySubmissionFile(fileId, submissionId, expectedSlot) {
  if (!fileId) return false;
  try {
    var file = DriveApp.getFileById(fileId);
    if (!file) return false;
    
    var mime = file.getMimeType();
    if (!mime || !mime.startsWith("image/")) return false;
    
    var name = file.getName();
    
    // Check if filename starts with submissionId + "__" + expectedSlot
    var expectedPrefix = submissionId + "__" + expectedSlot;
    if (name.indexOf(expectedPrefix) !== 0) return false;
    
    // Verify the delimiter is either "__" or "_\d+__"
    var remaining = name.substring(expectedPrefix.length);
    if (!remaining.startsWith("__") && !/^_\d+__/.test(remaining)) return false;
    
    var parents = file.getParents();
    if (!parents.hasNext()) return false;
    var parentFolder = parents.next();
    
    var expectedFolder = getOrCreateStorePhotosFolder();
    if (parentFolder.getId() !== expectedFolder.getId()) return false;
    
    return true;
  } catch(e) {
    console.warn("Lỗi xác minh file " + fileId + ": " + e.toString());
    return false;
  }
}

// -------------------------------------------------------------
// SEQUENTIAL IMAGE UPLOAD ENDPOINT
// -------------------------------------------------------------
function uploadSubmissionImage(payload) {
  try {
    if (!payload || typeof payload !== 'object') {
      throw new Error("Payload không hợp lệ.");
    }
    var submissionId = payload.submission_id;
    var slot = payload.slot;
    var base64Data = payload.base64;
    
    if (!validateSubmissionId(submissionId)) {
      throw new Error("Mã submission_id không đúng định dạng UUID.");
    }
    if (!slot || typeof slot !== 'string') {
      throw new Error("Thiếu thông tin slot ảnh.");
    }
    if (!/^[a-zA-Z0-9_]+$/.test(slot)) {
      throw new Error("Tên slot chứa ký tự không hợp lệ.");
    }
    if (slot.includes(submissionId) || slot.length > 30) {
      throw new Error("Slot không được chứa submission_id hoặc quá dài.");
    }
    if (!isValidUploadSlot(slot)) {
      throw new Error("Slot upload ảnh không hợp lệ hoặc bị từ chối bởi chính sách bảo mật.");
    }
    if (!base64Data || typeof base64Data !== 'string') {
      throw new Error("Thiếu dữ liệu Base64 của ảnh.");
    }
    
    var parts = base64Data.split(',');
    if (parts.length !== 2) {
      throw new Error("Dữ liệu ảnh không đúng định dạng Base64.");
    }
    var header = parts[0];
    var base64Content = parts[1];
    
    var mimeType = header.split(';')[0].split(':')[1] || '';
    if (!mimeType.startsWith('image/')) {
      throw new Error("Chỉ cho phép tải lên file hình ảnh (MIME nhận được: " + mimeType + ").");
    }
    
    var approxSize = base64Content.length * 0.75;
    if (approxSize > 10 * 1024 * 1024) {
      throw new Error("Dung lượng ảnh vượt quá giới hạn 10MB.");
    }
    
    var folder = getOrCreateStorePhotosFolder();
    var ext = "jpg";
    if (mimeType === "image/png") ext = "png";
    else if (mimeType === "image/gif") ext = "gif";
    else if (mimeType === "image/webp") ext = "webp";
    
    var timestamp = new Date().getTime();
    var cleanFileName = submissionId + "__" + slot + "__" + timestamp + "." + ext;
    
    var decodedBlob = Utilities.newBlob(Utilities.base64Decode(base64Content), mimeType, cleanFileName);
    var file = folder.createFile(decodedBlob);
    
    // Cấp quyền viewer cho email service account (Redundant: permissions are inherited from folder)
    // grantViewerAccessToWorker(file);
    
    return {
      success: true,
      fileId: file.getId(),
      slot: slot,
      name: cleanFileName
    };
  } catch(e) {
    return {
      success: false,
      error: e.toString()
    };
  }
}

// -------------------------------------------------------------
// SERVER-SIDE SCORING LOGIC
// -------------------------------------------------------------
function calculateServerRating(secKey, items) {
  var totalApplicable = 0;
  var totalPassed = 0;
  var hasSeriousFailure = false;
  var pendingCount = 0;
  
  items.forEach(function(item) {
    if (item.eval === "Chưa kiểm tra") {
      pendingCount++;
    } else if (item.eval !== "Không áp dụng") {
      totalApplicable++;
      if (item.eval === "Đạt") {
        totalPassed++;
      } else if (item.eval === "Không đạt") {
        var isSerious = SERIOUS_ITEMS.includes(item.id) || item.severity === "Nghiêm trọng";
        if (isSerious) {
          hasSeriousFailure = true;
        }
      }
    }
  });
  
  var rating = "Chưa đánh giá";
  var passRate = 0;
  
  if (pendingCount > 0) {
    rating = "Chưa đánh giá";
  } else if (totalApplicable === 0) {
    rating = "Không áp dụng";
  } else {
    passRate = (totalPassed / totalApplicable) * 100;
    if (hasSeriousFailure) {
      rating = "Chưa đạt";
    } else if (passRate === 100) {
      rating = "Tốt";
    } else if (passRate >= 80) {
      rating = "Đạt";
    } else {
      rating = "Chưa đạt";
    }
  }
  
  return {
    rating: rating,
    hasSeriousFailure: hasSeriousFailure
  };
}

function getOverallCategoryServerRating(overallKey, subSecKeys, sections) {
  var subRatings = [];
  var allNA = true;
  var anyUnchecked = false;
  var anySerious = false;
  
  subSecKeys.forEach(function(k) {
    var sec = sections[k];
    if (sec) {
      var secScore = calculateServerRating(k, sec.items);
      var secRating = sec.adjusted ? sec.rating : secScore.rating;
      
      if (secScore.hasSeriousFailure) {
        secRating = "Chưa đạt";
        anySerious = true;
      }
      
      subRatings.push(secRating);
      if (secRating !== "Không áp dụng") allNA = false;
      if (secRating === "Chưa đánh giá") anyUnchecked = true;
    }
  });
  
  var autoRating = "Chưa đánh giá";
  if (anyUnchecked) {
    autoRating = "Chưa đánh giá";
  } else if (allNA) {
    autoRating = "Không áp dụng";
  } else {
    var activeRatings = subRatings.filter(function(r) {
      return r !== "Không áp dụng" && r !== "Chưa đánh giá";
    });
    if (activeRatings.includes("Chưa đạt") || anySerious) {
      autoRating = "Chưa đạt";
    } else if (activeRatings.every(function(r) { return r === "Tốt"; })) {
      autoRating = "Tốt";
    } else {
      autoRating = "Đạt";
    }
  }
  
  return {
    autoRating: autoRating,
    anySerious: anySerious
  };
}

// -------------------------------------------------------------
// MAIN SUBMISSION PROCESS
// -------------------------------------------------------------
function processForm(formObject) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // Wait up to 30 seconds
  } catch(e) {
    return { success: false, error: "Hệ thống đang bận xử lý yêu cầu khác, vui lòng thử lại sau ít phút." };
  }

  try {
    // 0. Auto-check mapping consistency
    var mappingCheck = verifyMappingConsistency();
    if (!mappingCheck.success) {
      throw new Error("Lỗi hệ thống: Mapping cửa hàng giữa ASM và Region bị lệch. Chi tiết: " + JSON.stringify(mappingCheck.errors));
    }

    // 1. Server-side validations
    if (!formObject.submission_id || !validateSubmissionId(formObject.submission_id)) {
      throw new Error("Mã submission_id không tồn tại hoặc sai định dạng.");
    }
    
    var isDraftFlag = (formObject.isDraft === "true" || formObject.isDraft === true);
    
    if (isDraftFlag) {
      if (!formObject.modeSelect || !["own", "cross"].includes(formObject.modeSelect)) {
        formObject.modeSelect = "own";
      }
      if (!formObject.asmName) {
        formObject.asmName = "Draft ASM";
      }
      if (!formObject.storeCode) {
        formObject.storeCode = "DRAFT_STORE";
      }
      if (!formObject.reportDate || !/^\d{4}-\d{2}-\d{2}$/.test(formObject.reportDate)) {
        formObject.reportDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      if (!formObject.chtName || !formObject.chtName.trim()) {
        formObject.chtName = "Draft CHT";
      }
      if (!formObject.timeStart || !/^\d{2}:\d{2}$/.test(formObject.timeStart)) {
        formObject.timeStart = "08:00";
      }
      if (!formObject.timeEnd || !/^\d{2}:\d{2}$/.test(formObject.timeEnd)) {
        formObject.timeEnd = "17:00";
      }
      if (!formObject.nvCount || !/^\d+$/.test(String(formObject.nvCount))) {
        formObject.nvCount = "0";
      }
    }
    
    // Validate Mode Select
    if (!isDraftFlag) {
      if (!formObject.modeSelect || !["own", "cross"].includes(formObject.modeSelect)) {
        throw new Error("Hình thức kiểm tra không hợp lệ (chỉ chấp nhận 'own' hoặc 'cross').");
      }
    }
    
    // Fetch dynamic store mapping for validation
    var dynamicStoreData = getStoreData();
    
    // Validate ASM Name
    if (!isDraftFlag) {
      if (!formObject.asmName || !dynamicStoreData.asms.includes(formObject.asmName)) {
        throw new Error("Tên ASM không hợp lệ hoặc không có trong danh sách.");
      }
    }
    
    // Validate Region Select
    if (!isDraftFlag && formObject.modeSelect === "cross") {
      if (!formObject.regionSelect || !dynamicStoreData.regions.includes(formObject.regionSelect)) {
        throw new Error("Khu vực (Region) bắt buộc và phải hợp lệ khi chọn kiểm tra chéo.");
      }
    }
    
    // Validate Store Code
    if (!isDraftFlag) {
      if (!formObject.storeCode) {
        throw new Error("Thiếu mã cửa hàng.");
      }
      if (formObject.modeSelect === "own") {
        var storesForAsm = dynamicStoreData.mapping_by_asm[formObject.asmName] || [];
        if (!storesForAsm.includes(formObject.storeCode)) {
          throw new Error("Cửa hàng " + formObject.storeCode + " không thuộc quyền quản lý của ASM " + formObject.asmName);
        }
      } else if (formObject.modeSelect === "cross") {
        var storesForRegion = dynamicStoreData.mapping_by_region[formObject.regionSelect] || [];
        if (!storesForRegion.includes(formObject.storeCode)) {
          throw new Error("Cửa hàng " + formObject.storeCode + " không thuộc khu vực " + formObject.regionSelect);
        }
      }
    }
    
    // Validate Date (Valid calendar date check)
    if (!isDraftFlag) {
      if (!formObject.reportDate || !/^\d{4}-\d{2}-\d{2}$/.test(formObject.reportDate)) {
        throw new Error("Ngày kiểm tra không hợp lệ hoặc sai định dạng YYYY-MM-DD.");
      }
      var dateParts = formObject.reportDate.split("-");
      var year = parseInt(dateParts[0], 10);
      var month = parseInt(dateParts[1], 10) - 1; // 0-indexed
      var day = parseInt(dateParts[2], 10);
      var testDate = new Date(year, month, day);
      if (testDate.getFullYear() !== year || testDate.getMonth() !== month || testDate.getDate() !== day) {
        throw new Error("Ngày kiểm tra không phải là ngày hợp lệ thực tế.");
      }
    }
    
    // Validate CHT Name
    if (!isDraftFlag) {
      if (!formObject.chtName || !formObject.chtName.trim()) {
        throw new Error("Thiếu tên Cửa Hàng Trưởng / Quản lý ca.");
      }
    }
    
    // Validate Time (00:00 - 23:59 range check)
    if (!isDraftFlag) {
      if (!formObject.timeStart || !/^\d{2}:\d{2}$/.test(formObject.timeStart)) {
        throw new Error("Giờ bắt đầu không đúng định dạng HH:MM.");
      }
      var startParts = formObject.timeStart.split(":");
      var startH = parseInt(startParts[0], 10);
      var startM = parseInt(startParts[1], 10);
      if (startH < 0 || startH > 23 || startM < 0 || startM > 59) {
        throw new Error("Giờ bắt đầu phải nằm trong khoảng 00:00 - 23:59.");
      }

      if (!formObject.timeEnd || !/^\d{2}:\d{2}$/.test(formObject.timeEnd)) {
        throw new Error("Giờ kết thúc không đúng định dạng HH:MM.");
      }
      var endParts = formObject.timeEnd.split(":");
      var endH = parseInt(endParts[0], 10);
      var endM = parseInt(endParts[1], 10);
      if (endH < 0 || endH > 23 || endM < 0 || endM > 59) {
        throw new Error("Giờ kết thúc phải nằm trong khoảng 00:00 - 23:59.");
      }
    }
    
    // Validate Employee Count (Strictly digits only)
    if (!isDraftFlag) {
      if (!formObject.nvCount || !/^\d+$/.test(String(formObject.nvCount))) {
        throw new Error("Số nhân viên có mặt phải là số nguyên không âm.");
      }
    }
    var parsedNvCount = parseInt(formObject.nvCount, 10);
    
    // Check user email
    var userEmail = "";
    try {
      userEmail = Session.getActiveUser().getEmail() || "unknown";
    } catch(e) {
      userEmail = "unknown";
    }
    
    // Connect to sheet
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error("Không tìm thấy Sheet có tên: " + SHEET_NAME);
    }
    
    // 2. Prevent duplicate submission under ScriptLock
    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var submissionIdColIdx = -1;
    
    if (lastCol > 0 && lastRow > 0) {
      var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      for (var j = 0; j < headers.length; j++) {
        var headerStrNorm = normalizeHeader(headers[j]);
        if (headerStrNorm === "submissionid" || headerStrNorm === "submission_id") {
          submissionIdColIdx = j + 1;
          break;
        }
      }
      if (!(formObject.isEdit === "true" || formObject.isEdit === true) && submissionIdColIdx !== -1 && lastRow > 1) {
        var submissionIds = sheet.getRange(2, submissionIdColIdx, lastRow - 1, 1).getValues();
        for (var r = 0; r < submissionIds.length; r++) {
          if (String(submissionIds[r][0]) === formObject.submission_id) {
            return { success: true, duplicate: true, msg: "Báo cáo đã được gửi thành công trước đó." };
          }
        }
      }
    }
    
    // 3. Parse JSONs and validate
    // Google Sheets giới hạn 50.000 ký tự/cell — chặn sớm với thông báo rõ ràng
    var CELL_CHAR_LIMIT = 49500;
    if (formObject.checklist_json && formObject.checklist_json.length > CELL_CHAR_LIMIT) {
      throw new Error("Dữ liệu checklist quá lớn (" + formObject.checklist_json.length + " ký tự, giới hạn " + CELL_CHAR_LIMIT + "). Vui lòng rút gọn các ghi chú/nhận xét dài rồi gửi lại.");
    }
    if (formObject.survey_json && formObject.survey_json.length > CELL_CHAR_LIMIT) {
      throw new Error("Dữ liệu khảo sát quá lớn (" + formObject.survey_json.length + " ký tự, giới hạn " + CELL_CHAR_LIMIT + "). Vui lòng rút gọn các câu trả lời rồi gửi lại.");
    }
    var checklistObj = {};
    if (!formObject.checklist_json) {
      throw new Error("Thiếu dữ liệu checklist_json.");
    }
    try {
      checklistObj = JSON.parse(formObject.checklist_json);
    } catch(e) {
      throw new Error("Dữ liệu checklist_json không đúng định dạng JSON.");
    }
    
    var surveyObj = {};
    if (formObject.survey_json) {
      try {
        surveyObj = JSON.parse(formObject.survey_json);
      } catch(e) {
        throw new Error("Dữ liệu survey_json không đúng định dạng JSON.");
      }
    } else {
      throw new Error("Thiếu dữ liệu survey_json.");
    }
    
    // Validate survey keys A1-E7
    var surveyKeys = [];
    for (var i = 1; i <= 5; i++) surveyKeys.push("A" + i);
    for (var i = 1; i <= 4; i++) surveyKeys.push("B" + i);
    for (var i = 1; i <= 4; i++) surveyKeys.push("C" + i);
    for (var i = 1; i <= 6; i++) surveyKeys.push("D" + i);
    for (var i = 1; i <= 7; i++) surveyKeys.push("E" + i);
    
    if (!isDraftFlag) {
      surveyKeys.forEach(function(key) {
        if (!surveyObj[key] || typeof surveyObj[key] !== 'object' || !surveyObj[key].hasOwnProperty('answer')) {
          throw new Error("Thiếu câu trả lời cho khảo sát câu: " + key);
        }
        var ans = surveyObj[key].answer;
        if (ans === null || ans === undefined || typeof ans !== 'string' || ans.trim() === "") {
          throw new Error("Câu trả lời cho khảo sát câu " + key + " phải là một chuỗi và không được để trống.");
        }
        if (ans.length > 5000) {
          throw new Error("Câu trả lời khảo sát câu " + key + " quá dài.");
        }
      });
    }
    
    var serverPendingIssues = [];
    
    // Validate Checklist items & Drive file existences
    var sections = checklistObj.sections || {};
    var requiredSections = Object.keys(REQUIRED_CHECKLIST_SCHEMA);
    
    requiredSections.forEach(function(secKey) {
      var secVal = sections[secKey];
      if (!secVal) {
        throw new Error("Thiếu dữ liệu đánh giá phần: " + secKey);
      }
      var items = secVal.items || [];
      if (!Array.isArray(items) || items.length === 0) {
        throw new Error("Dữ liệu phần " + secKey + " rỗng hoặc không hợp lệ.");
      }
      
      var requiredIds = REQUIRED_CHECKLIST_SCHEMA[secKey];
      var receivedIds = items.map(function(it) { return it.id; });
      
      // Duplicates check
      var hasDup = receivedIds.some(function(item, idx) { return receivedIds.indexOf(item) !== idx; });
      if (hasDup) throw new Error("Trùng lặp tiêu chí trong phần: " + secKey);
      
      // Check completeness
      requiredIds.forEach(function(reqId) {
        if (!receivedIds.includes(reqId)) {
          throw new Error("Thiếu tiêu chí bắt buộc " + reqId + " trong phần " + secKey);
        }
      });
      receivedIds.forEach(function(recId) {
        if (!requiredIds.includes(recId)) {
          throw new Error("Tiêu chí lạ không hợp lệ " + recId + " trong phần " + secKey);
        }
      });
      
      items.forEach(function(item) {
        if (isDraftFlag) {
          if (item.eval === "Không đạt" && item.resolved === "Không") {
            serverPendingIssues.push({
              issue_id: item.id,
              source_section: secKey,
              item_id: item.id,
              item_label: item.label,
              note: item.note || "",
              severity: item.severity || "Thường",
              assignee: item.assignee || "",
              deadline: item.deadline || "",
              photo_before: item.photo_before || "",
              status: "Chưa xử lý"
            });
          }
          return;
        }
        
        if (!VALID_EVALS.includes(item.eval)) {
          throw new Error("Trạng thái đánh giá tiêu chí " + item.id + " không hợp lệ.");
        }
        if (item.eval === "Chưa kiểm tra") {
          throw new Error("Tiêu chí " + item.id + " của phần " + secKey + " chưa được đánh giá.");
        }
        if (item.eval === "Không áp dụng" && (!item.na_reason || !item.na_reason.trim())) {
          throw new Error("Tiêu chí " + item.id + " chọn không áp dụng nhưng thiếu lý do.");
        }
        
        if (item.eval === "Không đạt") {
          if (!VALID_SEVERITIES.includes(item.severity)) {
            throw new Error("Mức độ nghiêm trọng của tiêu chí " + item.id + " không hợp lệ.");
          }
          if (!VALID_RESOLVED.includes(item.resolved)) {
            throw new Error("Trạng thái khắc phục của tiêu chí " + item.id + " không hợp lệ.");
          }
          if (!item.note || !item.note.trim()) {
            throw new Error("Tiêu chí " + item.id + " đánh giá không đạt nhưng thiếu ghi chú lỗi.");
          }
          var isPhotoRequired = !["S1", "S2", "S3", "S4", "TN1"].includes(item.id);
          if (isPhotoRequired && !item.photo_before) {
            throw new Error("Tiêu chí " + item.id + " đánh giá không đạt nhưng thiếu ảnh hiện trạng.");
          }
          if (item.photo_before) {
            var beforeIds = item.photo_before.split(",");
            for (var j = 0; j < beforeIds.length; j++) {
              var bid = beforeIds[j].trim();
              if (bid) {
                if (!verifySubmissionFile(bid, formObject.submission_id, item.id + "_before")) {
                  throw new Error("Ảnh hiện trạng của tiêu chí " + item.id + " không hợp lệ hoặc không thuộc submission này.");
                }
              }
            }
          }
          
          if (item.resolved === "Có") {
            if (isPhotoRequired && !item.photo_after) {
              throw new Error("Tiêu chí " + item.id + " đã khắc phục nhưng thiếu ảnh sau khắc phục.");
            }
            if (item.photo_after) {
              var afterIds = item.photo_after.split(",");
              for (var j = 0; j < afterIds.length; j++) {
                var aid = afterIds[j].trim();
                if (aid) {
                  if (!verifySubmissionFile(aid, formObject.submission_id, item.id + "_after")) {
                    throw new Error("Ảnh sau khắc phục của tiêu chí " + item.id + " không hợp lệ hoặc không thuộc submission này.");
                  }
                }
              }
            }
          } else {
            // Lỗi chưa xử lý
            if (!item.assignee || !item.assignee.trim()) {
              throw new Error("Tiêu chí " + item.id + " chưa đạt nhưng thiếu người chịu trách nhiệm.");
            }
            if (!item.deadline || !item.deadline.trim()) {
              throw new Error("Tiêu chí " + item.id + " chưa đạt nhưng thiếu hạn xử lý.");
            }
            if (!/^\d{4}-\d{2}-\d{2}$/.test(item.deadline)) {
              throw new Error("Hạn xử lý tiêu chí " + item.id + " sai định dạng YYYY-MM-DD.");
            }
            var reportDateVal = new Date(formObject.reportDate);
            var deadlineVal = new Date(item.deadline);
            reportDateVal.setHours(0,0,0,0);
            deadlineVal.setHours(0,0,0,0);
            if (deadlineVal < reportDateVal) {
              throw new Error("Hạn xử lý tiêu chí " + item.id + " không được trước ngày kiểm tra.");
            }
            
            // Collect for server-side pending issues
            serverPendingIssues.push({
              issue_id: item.id,
              source_section: secKey,
              item_id: item.id,
              item_label: item.label,
              note: item.note,
              severity: item.severity,
              assignee: item.assignee,
              deadline: item.deadline,
              photo_before: item.photo_before,
              status: "Chưa xử lý"
            });
          }
        } else {
          // Eval !== "Không đạt" thì resolved phải là ""
          if (item.resolved !== "") {
            throw new Error("Tiêu chí " + item.id + " có trạng thái " + item.eval + " nhưng mang resolved là '" + item.resolved + "' (bắt buộc phải là chuỗi rỗng).");
          }
        }
      });
      
      // Tính rating tự động và áp dụng manual override an toàn
      var secScore = calculateServerRating(secKey, items);
      
      if (isDraftFlag) {
        secVal.auto_rating = secScore.rating;
        secVal.has_serious_failure = secScore.hasSeriousFailure;
        
        var totalApplicable = 0;
        var totalPassed = 0;
        var pendingCount = 0;
        items.forEach(function(item) {
          if (item.eval === "Chưa kiểm tra") pendingCount++;
          else if (item.eval !== "Không áp dụng") {
            totalApplicable++;
            if (item.eval === "Đạt") totalPassed++;
          }
        });
        secVal.pass_rate = totalApplicable > 0 ? Math.round((totalPassed / totalApplicable) * 100) : 0;
        secVal.pending_count = pendingCount;
      } else {
        if (secVal.adjusted === true) {
          if (!secVal.adjust_reason || !secVal.adjust_reason.trim()) {
            throw new Error("Phần " + secKey + " có điều chỉnh xếp loại nhưng thiếu lý do điều chỉnh.");
          }
          if (secVal.rating === secScore.rating) {
            throw new Error("Phần " + secKey + " có trạng thái adjusted=true nhưng xếp loại thủ công trùng xếp loại tự động.");
          }
          if (secVal.rating === "Chưa đánh giá" || secVal.rating === "Không áp dụng") {
            throw new Error("Không được điều chỉnh xếp loại thủ công phần " + secKey + " thành Chưa đánh giá hoặc Không áp dụng.");
          }
          if (secScore.rating === "Chưa đánh giá" || secScore.rating === "Không áp dụng") {
            throw new Error("Không được điều chỉnh xếp loại thủ công phần " + secKey + " khi trạng thái tự động là Chưa đánh giá hoặc Không áp dụng.");
          }
          if (secScore.rating === "Chưa đạt" && (secVal.rating === "Đạt" || secVal.rating === "Tốt")) {
            throw new Error("Không được thay đổi xếp loại thành Đạt/Tốt cho phần " + secKey + " khi kết quả tự động là Chưa đạt.");
          }
        } else if (secVal.adjusted === false) {
          if (secVal.rating !== secScore.rating) {
            throw new Error("Phần " + secKey + " có xếp loại thủ công khác xếp loại tự động mà adjusted=false.");
          }
        } else {
          throw new Error("Trạng thái adjusted của phần " + secKey + " không hợp lệ.");
        }
        
        if (secScore.hasSeriousFailure) {
          if (secVal.rating !== "Chưa đạt") {
            throw new Error("Phần " + secKey + " có lỗi nghiêm trọng nhưng xếp loại là " + secVal.rating + " (bắt buộc phải là Chưa đạt).");
          }
        }
        
        if (!VALID_RATINGS.includes(secVal.rating)) {
          throw new Error("Xếp loại của phần " + secKey + " không hợp lệ: " + secVal.rating);
        }
        
        secVal.auto_rating = secScore.rating;
        secVal.has_serious_failure = secScore.hasSeriousFailure;
        
        var totalApplicable = 0;
        var totalPassed = 0;
        var pendingCount = 0;
        items.forEach(function(item) {
          if (item.eval === "Chưa kiểm tra") pendingCount++;
          else if (item.eval !== "Không áp dụng") {
            totalApplicable++;
            if (item.eval === "Đạt") totalPassed++;
          }
        });
        secVal.pass_rate = totalApplicable > 0 ? Math.round((totalPassed / totalApplicable) * 100) : 0;
        secVal.pending_count = pendingCount;
      }
    });
    
    if (!isDraftFlag) {
      // Validate General Photos
      var generalPhotos = checklistObj.general_photos || {};
      var requiredGeneral = ["frontage_main", "inner_entrance", "inner_left", "inner_right", "stockroom", "fitting_room", "cashier"];
      requiredGeneral.forEach(function(slot) {
        var fileId = generalPhotos[slot];
        if (!fileId) {
          throw new Error("Thiếu ảnh tổng quan bắt buộc: " + slot);
        }
        var ids = fileId.split(",");
        for (var j = 0; j < ids.length; j++) {
          var id = ids[j].trim();
          if (id) {
            if (!verifySubmissionFile(id, formObject.submission_id, slot)) {
              throw new Error("Ảnh tổng quan slot " + slot + " không hợp lệ hoặc không thuộc submission này.");
            }
          }
        }
      });
      
      // Check optional/N/A general photos
      ["frontage_left", "frontage_right"].forEach(function(slot) {
        if (generalPhotos[slot + "_na"]) {
          var reason = generalPhotos[slot + "_na_reason"];
          if (!reason || !reason.trim()) {
            throw new Error("Ảnh tổng quan slot " + slot + " chọn Không áp dụng (N/A) nhưng thiếu lý do.");
          }
        } else {
          var fileId = generalPhotos[slot];
          if (!fileId) {
            throw new Error("Thiếu ảnh tổng quan bắt buộc: " + slot);
          }
          var ids = fileId.split(",");
          for (var j = 0; j < ids.length; j++) {
            var id = ids[j].trim();
            if (id) {
              if (!verifySubmissionFile(id, formObject.submission_id, slot)) {
                throw new Error("Ảnh tổng quan slot " + slot + " không hợp lệ hoặc không thuộc submission này.");
              }
            }
          }
        }
      });
    }
    
    if (!isDraftFlag) {
      // Validate Competitor
      var competitorObj = checklistObj.competitor || {};
      if (competitorObj.has_competitor === true) {
        if (!competitorObj.name || !competitorObj.name.trim()) {
          throw new Error("Thiếu tên đối thủ trực tiếp khi có đối thủ.");
        }
        var threatLevel = competitorObj.threat_level || "";
        if (threatLevel && !["Cao", "Vừa", "Thấp"].includes(threatLevel)) {
          throw new Error("Mức độ đe dọa đối thủ không hợp lệ.");
        }
        if (!competitorObj.photo1) {
          throw new Error("Thiếu ảnh minh chứng đối thủ 1 bắt buộc khi có đối thủ.");
        }
        if (!verifySubmissionFile(competitorObj.photo1, formObject.submission_id, "photoComp1")) {
          throw new Error("Ảnh đối thủ 1 không hợp lệ hoặc không thuộc submission này.");
        }
        for (var i = 2; i <= 3; i++) {
          var fileId = competitorObj["photo" + i];
          if (fileId && !verifySubmissionFile(fileId, formObject.submission_id, "photoComp" + i)) {
            throw new Error("Ảnh đối thủ " + i + " không hợp lệ hoặc không thuộc submission này.");
          }
        }
      } else if (competitorObj.has_competitor === false) {
        if (!competitorObj.no_competitor_reason || !competitorObj.no_competitor_reason.trim()) {
          throw new Error("Thiếu nhận xét tình hình lân cận khi không có đối thủ.");
        }
      } else {
        throw new Error("Trạng thái has_competitor của đối thủ không hợp lệ.");
      }
    }
    
    // Validate CSVC additional photo
    if (!isDraftFlag && formObject.photoCSVC1) {
      if (!verifySubmissionFile(formObject.photoCSVC1, formObject.submission_id, "photoCSVC1")) {
        throw new Error("Ảnh CSVC bổ sung không hợp lệ hoặc không thuộc submission này.");
      }
    }
    
    // Server-side Rating Consolidated Verification (FIX-04 & FIX-05)
    var overallRatings = {};
    
    var overallCategories = [
      { key: "frontage", subSecs: ["frontage"] },
      { key: "inner", subSecs: ["inner"] },
      { key: "merch", subSecs: ["merch_ap", "merch_pie", "merch_anamai", "merch_bonjour", "merch_pk"] },
      { key: "csvc", subSecs: ["stockroom", "fitting_room", "toilet", "fire_safety", "cashier", "packaging_security"] },
      { key: "staff", subSecs: ["staff"] }
    ];
    
    overallCategories.forEach(function(cat) {
      var catKey = cat.key;
      var res = getOverallCategoryServerRating(catKey, cat.subSecs, sections);
      var overallObj = checklistObj.overall[catKey];
      if (!overallObj) {
        overallObj = { rating: "Chưa đánh giá", adjusted: false, adjust_reason: "", auto_rating: "Chưa đánh giá" };
        checklistObj.overall[catKey] = overallObj;
      }
      
      if (isDraftFlag) {
        overallObj.auto_rating = res.autoRating;
        overallRatings[catKey] = overallObj.rating || "Chưa đánh giá";
        return;
      }
      
      // Validate adjusted
      if (overallObj.adjusted === true) {
        if (!overallObj.adjust_reason || !overallObj.adjust_reason.trim()) {
          throw new Error("Đánh giá chung phần " + catKey + " có điều chỉnh xếp loại nhưng thiếu lý do.");
        }
        if (overallObj.rating === res.autoRating) {
          throw new Error("Đánh giá chung phần " + catKey + " có trạng thái adjusted=true nhưng xếp loại trùng xếp loại tự động.");
        }
        if (overallObj.rating === "Chưa đánh giá" || overallObj.rating === "Không áp dụng") {
          throw new Error("Không được xếp loại thủ công phần chung " + catKey + " thành Chưa đánh giá hoặc Không áp dụng.");
        }
        if (res.autoRating === "Chưa đánh giá" || res.autoRating === "Không áp dụng") {
          throw new Error("Không được xếp loại thủ công phần chung " + catKey + " khi chưa đánh giá xong hoặc không áp dụng.");
        }
        if (res.autoRating === "Chưa đạt" && (overallObj.rating === "Đạt" || overallObj.rating === "Tốt")) {
          throw new Error("Không được thay đổi xếp loại thành Đạt/Tốt cho phần chung " + catKey + " khi kết quả tự động là Chưa đạt.");
        }
      } else if (overallObj.adjusted === false) {
        if (overallObj.rating !== res.autoRating) {
          throw new Error("Đánh giá chung phần " + catKey + " có xếp loại thủ công khác xếp loại tự động mà adjusted=false.");
        }
      } else {
        throw new Error("Trạng thái adjusted của Đánh giá chung phần " + catKey + " không hợp lệ.");
      }
      
      // Ép rating thành Chưa đạt nếu có serious failure
      if (res.anySerious) {
        if (overallObj.rating !== "Chưa đạt") {
          throw new Error("Đánh giá chung phần " + catKey + " có lỗi nghiêm trọng nhưng xếp loại là " + overallObj.rating + " (bắt buộc phải là Chưa đạt).");
        }
      }
      
      if (!VALID_RATINGS.includes(overallObj.rating)) {
        throw new Error("Xếp loại chung của phần " + catKey + " không hợp lệ.");
      }
      
      overallObj.auto_rating = res.autoRating;
      overallRatings[catKey] = overallObj.rating;
    });

    // Write-back the server-generated pending_issues into checklistObj
    checklistObj.pending_issues = serverPendingIssues;

    // 4. Map columns with exact aliases
    var COLUMN_DEFINITIONS = [
      { key: "timestamp", defaultName: "Timestamp", aliases: ["timestamp", "thoi gian"] },
      { key: "submission_id", defaultName: "submission_id", aliases: ["submissionid", "submission_id", "ma gui"] },
      { key: "user_email", defaultName: "user_email", aliases: ["useremail", "user_email", "email"] },
      { key: "store_code", defaultName: "Mã cửa hàng", aliases: ["macuahang", "storecode"] },
      { key: "report_date", defaultName: "Ngày kiểm tra", aliases: ["ngaykiemtra", "date"] },
      { key: "asm_name", defaultName: "QLKD/ASM", aliases: ["qlkdasm", "qlkd", "asm", "nguoikiemsau"] },
      { key: "cht_name", defaultName: "Tên CHT", aliases: ["tencht", "cuahangtruong"] },
      { key: "time_start", defaultName: "Giờ bắt đầu", aliases: ["giobatdau", "timestart"] },
      { key: "time_end", defaultName: "Giờ kết thúc", aliases: ["gioketthuc", "timeend"] },
      { key: "nv_count", defaultName: "Số NV", aliases: ["sonv", "nhanviencomat"] },
      
      { key: "rating_frontage", defaultName: "rating_frontage", aliases: ["ratingfrontage", "danhgiamattien"] },
      { key: "comment_frontage", defaultName: "Nhận xét mặt tiền", aliases: ["nhanxetmattien"] },
      { key: "photo_frontage", defaultName: "Ảnh mặt tiền", aliases: ["anhmattien", "exteriorphoto"] },
      
      { key: "rating_inner", defaultName: "rating_inner", aliases: ["ratinginner", "danhgiabentrong"] },
      { key: "comment_inner", defaultName: "Nhận xét bên trong", aliases: ["nhanxetbentrong"] },
      { key: "photo_inner", defaultName: "Ảnh bên trong", aliases: ["anhbentrong", "innerphoto"] },
      
      { key: "rating_merch", defaultName: "rating_merch", aliases: ["ratingmerch", "danhgiahanghoa"] },
      { key: "comment_merch", defaultName: "Nhận xét hàng hóa", aliases: ["nhanxethanghoa"] },
      { key: "photo_merch", defaultName: "Ảnh hàng hóa", aliases: ["anhhanghoa", "merchandisephoto"] },
      
      { key: "rating_staff", defaultName: "rating_staff", aliases: ["ratingstaff", "danhgianhansu"] },
      { key: "comment_staff", defaultName: "Nhận xét nhân sự", aliases: ["nhanxetnhansu"] },
      { key: "photo_staff", defaultName: "Ảnh nhân sự", aliases: ["anhnhansu", "staffphoto"] },
      
      { key: "rating_csvc", defaultName: "rating_csvc", aliases: ["ratingcsvc", "danhgiacsvc"] },
      { key: "comment_csvc", defaultName: "Nhận xét CSVC", aliases: ["nhanxetcsvc"] },
      { key: "photo_csvc", defaultName: "Ảnh CSVC", aliases: ["anhcsvc", "csvcphoto"] },
      
      { key: "pending_issues", defaultName: "pending_issues", aliases: ["pendingissues", "vandetondong"] },
      { key: "action_plan", defaultName: "action_plan", aliases: ["actionplan", "kehoachkhacphuc"] },
      { key: "action_deadline", defaultName: "action_deadline", aliases: ["actiondeadline", "thoihanxuly"] },
      
      { key: "store_recommendation", defaultName: "Đề xuất phát triển", aliases: ["dexuatphattrien", "storerecommendation"] },
      { key: "checklist_json", defaultName: "checklist_json", aliases: ["checklistjson", "checklist"] },
      { key: "survey_json", defaultName: "survey_json", aliases: ["surveyjson", "survey"] },
      { key: "general_photos_json", defaultName: "general_photos_json", aliases: ["generalphotosjson", "general_photos_json"] },
      { key: "competitor_json", defaultName: "competitor_json", aliases: ["competitorjson", "competitor_json"] },
      { key: "inspection_mode", defaultName: "inspection_mode", aliases: ["inspection_mode", "hình thức kiểm tra", "loại kiểm tra", "hinh thuc kiem tra", "loai kiem tra"] },
      { key: "inspection_region", defaultName: "inspection_region", aliases: ["inspection_region", "khu vực kiểm tra", "region kiểm tra", "khu vuc kiem tra", "region kiem tra"] },
      
      { key: "status", defaultName: "Status", aliases: ["status"] },
      { key: "created_at", defaultName: "created_at", aliases: ["createdat", "created_at"] },
      { key: "updated_at", defaultName: "updated_at", aliases: ["updatedat", "updated_at"] }
    ];
    
    // Read headers
    var headers = [];
    if (sheet.getLastColumn() > 0) {
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    }
    
    var colIndexes = {};
    
    COLUMN_DEFINITIONS.forEach(function(def) {
      var matchedIdx = -1;
      for (var j = 0; j < headers.length; j++) {
        var headerStrNorm = normalizeHeader(headers[j]);
        for (var k = 0; k < def.aliases.length; k++) {
          if (headerStrNorm === normalizeHeader(def.aliases[k])) {
            matchedIdx = j;
            break;
          }
        }
        if (matchedIdx !== -1) break;
      }
      
      if (matchedIdx !== -1) {
        colIndexes[def.key] = matchedIdx + 1;
      } else {
        var newColIdx = headers.length + 1;
        sheet.getRange(1, newColIdx).setValue(def.defaultName);
        headers.push(def.defaultName);
        colIndexes[def.key] = newColIdx;
      }
    });
    
    // 5. Gather photo URLs for traditional photo columns
    // 5. Gather photo URLs for traditional photo columns
    function collectSectionPhotos(secKey, checkObj) {
      var urls = [];
      if (checkObj && checkObj.sections && checkObj.sections[secKey]) {
        var items = checkObj.sections[secKey].items || [];
        for (var i = 0; i < items.length; i++) {
          var item = items[i];
          if (item.photo_before) {
            item.photo_before.split(",").forEach(function(fid) {
              if (fid.trim()) urls.push(getFileUrl(fid.trim()));
            });
          }
          if (item.photo_after) {
            item.photo_after.split(",").forEach(function(fid) {
              if (fid.trim()) urls.push(getFileUrl(fid.trim()));
            });
          }
        }
      }
      return urls;
    }
    
    var frontageUrls = [];
    if (generalPhotos.frontage_main) {
      generalPhotos.frontage_main.split(",").forEach(function(fid) {
        if (fid.trim()) frontageUrls.push(getFileUrl(fid.trim()));
      });
    }
    if (generalPhotos.frontage_left && !generalPhotos.frontage_left_na) {
      generalPhotos.frontage_left.split(",").forEach(function(fid) {
        if (fid.trim()) frontageUrls.push(getFileUrl(fid.trim()));
      });
    }
    if (generalPhotos.frontage_right && !generalPhotos.frontage_right_na) {
      generalPhotos.frontage_right.split(",").forEach(function(fid) {
        if (fid.trim()) frontageUrls.push(getFileUrl(fid.trim()));
      });
    }
    frontageUrls = frontageUrls.concat(collectSectionPhotos("frontage", checklistObj));
    var frontageUrlsStr = frontageUrls.join("\n");
    
    var innerUrls = [];
    if (generalPhotos.inner_entrance) {
      generalPhotos.inner_entrance.split(",").forEach(function(fid) {
        if (fid.trim()) innerUrls.push(getFileUrl(fid.trim()));
      });
    }
    if (generalPhotos.inner_left) {
      generalPhotos.inner_left.split(",").forEach(function(fid) {
        if (fid.trim()) innerUrls.push(getFileUrl(fid.trim()));
      });
    }
    if (generalPhotos.inner_right) {
      generalPhotos.inner_right.split(",").forEach(function(fid) {
        if (fid.trim()) innerUrls.push(getFileUrl(fid.trim()));
      });
    }
    innerUrls = innerUrls.concat(collectSectionPhotos("inner", checklistObj));
    var innerUrlsStr = innerUrls.join("\n");
    
    var merchUrls = [];
    ["merch_ap", "merch_pie", "merch_anamai", "merch_bonjour", "merch_pk"].forEach(function(k) {
      merchUrls = merchUrls.concat(collectSectionPhotos(k, checklistObj));
    });
    var merchUrlsStr = merchUrls.join("\n");
    
    var staffUrlsStr = collectSectionPhotos("staff", checklistObj).join("\n");
    
    var csvcUrls = [];
    if (generalPhotos.stockroom) {
      generalPhotos.stockroom.split(",").forEach(function(fid) {
        if (fid.trim()) csvcUrls.push(getFileUrl(fid.trim()));
      });
    }
    if (generalPhotos.fitting_room) {
      generalPhotos.fitting_room.split(",").forEach(function(fid) {
        if (fid.trim()) csvcUrls.push(getFileUrl(fid.trim()));
      });
    }
    if (generalPhotos.cashier) {
      generalPhotos.cashier.split(",").forEach(function(fid) {
        if (fid.trim()) csvcUrls.push(getFileUrl(fid.trim()));
      });
    }
    if (formObject.photoCSVC1) {
      formObject.photoCSVC1.split(",").forEach(function(fid) {
        if (fid.trim()) csvcUrls.push(getFileUrl(fid.trim()));
      });
    }
    ["stockroom", "fitting_room", "toilet", "fire_safety", "cashier", "packaging_security"].forEach(function(k) {
      csvcUrls = csvcUrls.concat(collectSectionPhotos(k, checklistObj));
    });
    var csvcUrlsStr = csvcUrls.join("\n");
    
    // Parse numeric fields
    var parsedNvCount = parseInt(formObject.nvCount, 10);
    if (isNaN(parsedNvCount)) parsedNvCount = 0;
    
    var valuesMap = {
      timestamp: new Date(),
      submission_id: formObject.submission_id,
      user_email: userEmail,
      store_code: formObject.storeCode,
      report_date: formObject.reportDate,
      asm_name: formObject.asmName,
      cht_name: formObject.chtName,
      time_start: formObject.timeStart,
      time_end: formObject.timeEnd,
      nv_count: parsedNvCount,
      
      rating_frontage: overallRatings.frontage,
      comment_frontage: formObject.commentFrontage,
      photo_frontage: frontageUrlsStr,
      
      rating_inner: overallRatings.inner,
      comment_inner: formObject.commentInner,
      photo_inner: innerUrlsStr,
      
      rating_merch: overallRatings.merch,
      comment_merch: formObject.commentMerch,
      photo_merch: merchUrlsStr,
      
      rating_staff: overallRatings.staff,
      comment_staff: formObject.commentStaff,
      photo_staff: staffUrlsStr,
      
      rating_csvc: overallRatings.csvc,
      comment_csvc: formObject.commentCSVC,
      photo_csvc: csvcUrlsStr,
      
      pending_issues: formObject.pendingIssues,
      action_plan: formObject.actionPlan,
      action_deadline: formObject.actionDeadline,
      store_recommendation: formObject.storeRecommendation,
      checklist_json: JSON.stringify(checklistObj),
      survey_json: JSON.stringify(surveyObj),
      general_photos_json: JSON.stringify(generalPhotos),
      competitor_json: JSON.stringify(competitorObj),
      inspection_mode: formObject.modeSelect,
      inspection_region: formObject.regionSelect || "",
      status: isDraftFlag ? "draft" : "pending",
      created_at: new Date(),
      updated_at: new Date()
    };
    
    // Build row and apply formula injection protection to user strings
    var newRow = new Array(headers.length);
    for (var i = 0; i < newRow.length; i++) {
      newRow[i] = "";
    }
    
    for (var key in valuesMap) {
      if (valuesMap.hasOwnProperty(key) && colIndexes[key]) {
        var val = valuesMap[key];
        newRow[colIndexes[key] - 1] = sanitizeFormulaInjection(val);
      }
    }
    
    var rowIdx = -1;
    if (formObject.isEdit === "true" || formObject.isEdit === true) {
      // Find the row by submission_id
      var data = sheet.getDataRange().getValues();
      var headersData = data[0].map(function(h) { return String(h).trim(); });
      var responseIdCol = headersData.indexOf("submission_id");
      if (responseIdCol === -1) responseIdCol = headersData.indexOf("submissionid");
      if (responseIdCol === -1) responseIdCol = headersData.indexOf("ResponseId");
      if (responseIdCol === -1) responseIdCol = headersData.indexOf("Timestamp");
      
      if (responseIdCol !== -1) {
        for (var idxRow = 1; idxRow < data.length; idxRow++) {
          if (String(data[idxRow][responseIdCol]).trim() === formObject.submission_id) {
            rowIdx = idxRow + 1; // 1-based index
            break;
          }
        }
      }
    }
    
    if (rowIdx !== -1) {
      var range = sheet.getRange(rowIdx, 1, 1, newRow.length);
      range.setValues([newRow]);
    } else {
      sheet.appendRow(newRow);
    }

    // Clear caches
    try {
      var cache = CacheService.getScriptCache();
      cache.remove("store_data_json");
      cache.remove("historical_submissions_json");
    } catch(e) {
      console.warn("Failed to clear caches: " + e.toString());
    }

    return { success: true };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function cleanupSubmissionUploads(submissionId, fileIds) {
  if (!validateSubmissionId(submissionId)) return { success: false, error: "Invalid submission ID" };
  if (!Array.isArray(fileIds) || fileIds.length === 0) return { success: true };
  
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch(e) {
    return { success: false, error: "System busy" };
  }
  
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error("Sheet not found");
    
    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    
    var activeFileIds = [];
    var subRowIdx = -1;
    
    if (lastCol > 0 && lastRow > 1) {
      var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      var subIdColIdx = -1;
      for (var j = 0; j < headers.length; j++) {
        var hNorm = normalizeHeader(headers[j]);
        if (hNorm === "submissionid" || hNorm === "submission_id") {
          subIdColIdx = j + 1;
          break;
        }
      }
      if (subIdColIdx !== -1) {
        var submissionIds = sheet.getRange(2, subIdColIdx, lastRow - 1, 1).getValues();
        for (var r = 0; r < submissionIds.length; r++) {
          if (String(submissionIds[r][0]) === submissionId) {
            subRowIdx = r + 2; // Real sheet row index
            break;
          }
        }
      }
    }
    
    // If submission record exists, collect all actively used fileIds in this row
    if (subRowIdx !== -1) {
      var rowValues = sheet.getRange(subRowIdx, 1, 1, lastCol).getValues()[0];
      rowValues.forEach(function(val) {
        if (!val) return;
        var valStr = String(val);
        
        // Match standard drive file ids inside open?id=...
        var urlMatches = valStr.match(/id=([a-zA-Z0-9_-]{25,50})/g);
        if (urlMatches) {
          urlMatches.forEach(function(m) {
            activeFileIds.push(m.replace("id=", ""));
          });
        }
        
        // Match raw json quoted IDs
        var rawMatches = valStr.match(/"[a-zA-Z0-9_-]{25,50}"/g);
        if (rawMatches) {
          rawMatches.forEach(function(m) {
            activeFileIds.push(m.replace(/"/g, ""));
          });
        }
        
        // Fallback for line-based items
        var lines = valStr.split("\n");
        lines.forEach(function(line) {
          var lineMatches = line.match(/id=([a-zA-Z0-9_-]{25,50})/);
          if (lineMatches) {
            activeFileIds.push(lineMatches[1]);
          }
        });
      });
    }
    
    var folder = getOrCreateStorePhotosFolder();
    var deletedCount = 0;
    
    fileIds.forEach(function(fileId) {
      if (!fileId) return;
      
      // Safety check: Never delete a fileId if it's currently stored on Google Sheet record
      if (activeFileIds.includes(fileId)) {
        console.warn("File " + fileId + " is active in Sheets response, skipping cleanup.");
        return;
      }
      
      try {
        var file = DriveApp.getFileById(fileId);
        if (file) {
          var parents = file.getParents();
          if (parents.hasNext() && parents.next().getId() === folder.getId()) {
            var fileName = file.getName();
            if (fileName.startsWith(submissionId + "__")) {
              file.setTrashed(true);
              deletedCount++;
            }
          }
        }
      } catch(e) {
        console.warn("Cannot delete file " + fileId + ": " + e.toString());
      }
    });
    
    return { success: true, deletedCount: deletedCount };
  } catch(e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// -------------------------------------------------------------
// MAIN WEBHOOK FOR PYTHON WORKER
// -------------------------------------------------------------
// -------------------------------------------------------------
// CLIENT ERROR TELEMETRY — nhận lỗi runtime từ webapp, ghi vào sheet ClientErrors
// -------------------------------------------------------------
function logClientError(payload) {
  try {
    if (!payload || typeof payload !== 'object') {
      return { success: false, error: "Payload không hợp lệ." };
    }
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("ClientErrors");
    if (!sheet) {
      sheet = ss.insertSheet("ClientErrors");
      sheet.appendRow(["Timestamp", "Message", "Stack", "Source", "UserAgent", "URL", "SubmissionId"]);
    }
    // Giữ sheet gọn: vượt 5000 dòng thì cắt 1000 dòng cũ nhất
    if (sheet.getLastRow() > 5000) {
      sheet.deleteRows(2, 1000);
    }
    sheet.appendRow([
      new Date(),
      sanitizeFormulaInjection(String(payload.message || "").slice(0, 2000)),
      sanitizeFormulaInjection(String(payload.stack || "").slice(0, 3000)),
      sanitizeFormulaInjection(String(payload.source || "").slice(0, 500)),
      sanitizeFormulaInjection(String(payload.userAgent || "").slice(0, 500)),
      sanitizeFormulaInjection(String(payload.url || "").slice(0, 500)),
      sanitizeFormulaInjection(String(payload.appContext || "").slice(0, 100))
    ]);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function doPost(e) {
  try {
    var postData = JSON.parse(e.postData.contents);
    var action = postData.action;
    var payload = postData.payload;
    var result = {};
    
    if (action === "send_email") {
      result = sendReportEmail(postData);
    } else if (action === "getStoreData") {
      result = getStoreData();
    } else if (action === "getPendingIssues") {
      result = getPendingIssues(payload);
    } else if (action === "loginUser") {
      if (Array.isArray(payload)) {
        result = loginUser(payload[0], payload[1]);
      } else {
        result = loginUser(postData.username || payload, postData.password);
      }
    } else if (action === "changeUserPassword") {
      if (Array.isArray(payload)) {
        result = changeUserPassword(payload[0], payload[1], payload[2]);
      } else {
        result = changeUserPassword(postData.username, postData.oldPassword, postData.newPassword);
      }
    } else if (action === "getHistoricalSubmissions") {
      if (Array.isArray(payload)) {
        result = getHistoricalSubmissions(payload[0], payload[1], payload[2]);
      } else {
        result = getHistoricalSubmissions();
      }
    } else if (action === "getAllUsers") {
      var reqUser = Array.isArray(payload) ? payload[0] : payload;
      result = getAllUsers(reqUser);
    } else if (action === "saveUserAccount") {
      if (Array.isArray(payload)) {
        result = saveUserAccount(payload[0], payload[1]);
      } else {
        result = saveUserAccount(postData.requesterUsername, payload);
      }
    } else if (action === "resetUserPasswordByAdmin") {
      if (Array.isArray(payload)) {
        result = resetUserPasswordByAdmin(payload[0], payload[1], payload[2]);
      } else {
        result = resetUserPasswordByAdmin(postData.requesterUsername, postData.targetUsername, postData.newPassword);
      }
    } else if (action === "toggleUserStatus") {
      if (Array.isArray(payload)) {
        result = toggleUserStatus(payload[0], payload[1], payload[2]);
      } else {
        result = toggleUserStatus(postData.requesterUsername, postData.targetUsername, postData.newStatus);
      }
    } else if (action === "syncASMUsersFromStoresInfo") {
      if (Array.isArray(payload)) {
        result = syncASMUsersFromStoresInfo(payload[0], payload[1]);
      } else {
        result = syncASMUsersFromStoresInfo(postData.requesterUsername, postData.storeDataList);
      }
    } else if (action === "processForm") {
      result = processForm(payload);
    } else if (action === "updateIssueResolution") {
      result = updateIssueResolution(payload);
    } else if (action === "uploadSubmissionImage") {
      result = uploadSubmissionImage(payload);
    } else if (action === "getSubmissionDetail") {
      result = getSubmissionDetail(payload);
    } else if (action === "logClientError") {
      result = logClientError(payload);
    } else if (action === "cleanupSubmissionUploads") {
      if (Array.isArray(payload)) {
        result = cleanupSubmissionUploads(payload[0], payload[1]);
      } else {
        result = { success: false, error: "cleanupSubmissionUploads: payload phải là mảng [submissionId, fileIds]" };
      }
    } else {
      result = { success: false, error: "Invalid action: " + action };
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// -------------------------------------------------------------
// ISSUE FOLLOW-UP & RESOLUTION (PHASE 2)
// -------------------------------------------------------------
function getPendingIssues(storeCode) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });
    
    var storeCodeCol = headers.indexOf("StoreCode");
    if (storeCodeCol === -1) storeCodeCol = headers.indexOf("Mã cửa hàng");
    var checklistJsonCol = headers.indexOf("checklist_json");
    var responseIdCol = headers.indexOf("submission_id");
    if (responseIdCol === -1) responseIdCol = headers.indexOf("submissionid");
    if (responseIdCol === -1) responseIdCol = headers.indexOf("ResponseId");
    if (responseIdCol === -1) responseIdCol = headers.indexOf("Timestamp");
    
    if (storeCodeCol === -1 || checklistJsonCol === -1) {
      return { success: false, error: "Không tìm thấy cấu trúc cột dữ liệu phù hợp." };
    }
    
    var latestRowIdx = -1;
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][storeCodeCol]).trim() === storeCode) {
        latestRowIdx = i;
        break;
      }
    }
    
    if (latestRowIdx === -1) {
      return { success: true, issues: [], message: "Cửa hàng này chưa có báo cáo nào." };
    }
    
    var rowData = data[latestRowIdx];
    var jsonStr = rowData[checklistJsonCol];
    var responseId = String(rowData[responseIdCol]);
    
    if (!jsonStr) {
      return { success: true, issues: [], message: "Báo cáo gần nhất không chứa dữ liệu checklist." };
    }
    
    var checklist = JSON.parse(jsonStr);
    var sections = checklist.sections || {};
    var pendingIssues = [];
    
    var secLabels = {
      "frontage": "Mặt tiền",
      "inner": "Không gian trong",
      "merch_ap": "Trưng bày AP",
      "merch_pie": "Trưng bày PIE",
      "merch_ab": "Trưng bày AB",
      "merch_anamai": "Trưng bày Anamai",
      "merch_bonjour": "Trưng bày Bonjour",
      "merch_pk": "Phụ kiện",
      "warehouse": "Kho/Phòng thử",
      "stockroom": "Kho hàng",
      "fitting_room": "Phòng thử đồ",
      "toilet": "Nhà vệ sinh",
      "fire_safety": "PCCC & Thoát hiểm",
      "cashier": "Thu ngân",
      "packaging_security": "Bao bì & An ninh",
      "staff": "Nhân sự"
    };
    
    for (var secKey in sections) {
      var sec = sections[secKey];
      if (sec && sec.items) {
        sec.items.forEach(function(item) {
          if (item.eval === "Không đạt" && item.resolved !== "Có") {
            pendingIssues.push({
              submissionId: responseId,
              rowIdx: latestRowIdx + 1,
              sectionKey: secKey,
              sectionLabel: secLabels[secKey] || secKey,
              id: item.id,
              label: item.label || "",
              note: item.note || "",
              severity: item.severity || "Trung bình",
              assignee: item.assignee || "CHT",
              deadline: item.deadline || "-",
              photoBefore: item.photo_before || ""
            });
          }
        });
      }
    }
    
    return { success: true, issues: pendingIssues, submissionId: responseId, rowIdx: latestRowIdx + 1 };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function updateIssueResolution(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch(e) {
    return { success: false, error: "Hệ thống đang bận, vui lòng thử lại sau." };
  }
  
  try {
    var rowIdx = Number(payload.rowIdx);
    var sectionKey = payload.sectionKey;
    var itemId = payload.itemId;
    var photoAfter = payload.photoAfter;
    var notes = payload.notes || "";
    
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });
    
    var checklistJsonCol = headers.indexOf("checklist_json");
    var statusCol = headers.indexOf("Status");
    if (statusCol === -1) {
      statusCol = headers.length;
      sheet.getRange(1, statusCol + 1).setValue("Status");
    }
    
    if (checklistJsonCol === -1) {
      return { success: false, error: "Không tìm thấy cột checklist_json." };
    }
    
    var jsonStr = sheet.getRange(rowIdx, checklistJsonCol + 1).getValue();
    if (!jsonStr) {
      return { success: false, error: "Không có dữ liệu checklist tại dòng " + rowIdx };
    }
    
    var checklist = JSON.parse(jsonStr);
    var sections = checklist.sections || {};
    var sec = sections[sectionKey];
    if (!sec || !sec.items) {
      return { success: false, error: "Không tìm thấy phần hoặc tiêu chí phù hợp." };
    }
    
    var itemFound = false;
    sec.items.forEach(function(item) {
      if (item.id === itemId) {
        item.resolved = "Có";
        item.photo_after = photoAfter;
        item.resolved_notes = notes;
        item.resolved_at = new Date().toISOString();
        itemFound = true;
      }
    });
    
    if (!itemFound) {
      return { success: false, error: "Không tìm thấy tiêu chí " + itemId };
    }
    
    sheet.getRange(rowIdx, checklistJsonCol + 1).setValue(JSON.stringify(checklist));
    sheet.getRange(rowIdx, statusCol + 1).setValue("pending");
    
    // Clear history cache
    try {
      var cache = CacheService.getScriptCache();
      cache.remove("historical_submissions_json");
    } catch(e) {
      console.warn("Failed to clear history cache: " + e.toString());
    }

    return { success: true, message: "Đã cập nhật trạng thái khắc phục thành công." };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// -------------------------------------------------------------
// HISTORY & RE-EDITING (PHASE 3)
// -------------------------------------------------------------
function getHistoricalSubmissions(username, role, storesAllowedStr) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });
    
    var storeCodeCol = headers.indexOf("StoreCode");
    if (storeCodeCol === -1) storeCodeCol = headers.indexOf("Mã cửa hàng");
    var storeNameCol = headers.indexOf("StoreName");
    if (storeNameCol === -1) storeNameCol = headers.indexOf("Tên cửa hàng");
    var dateCol = headers.indexOf("ReportDate");
    if (dateCol === -1) dateCol = headers.indexOf("Ngày kiểm tra");
    if (dateCol === -1) dateCol = headers.indexOf("Ngày báo cáo");
    var asmCol = headers.indexOf("ASM");
    if (asmCol === -1) asmCol = headers.indexOf("QLKD/ASM");
    var responseIdCol = headers.indexOf("submission_id");
    if (responseIdCol === -1) responseIdCol = headers.indexOf("submissionid");
    if (responseIdCol === -1) responseIdCol = headers.indexOf("ResponseId");
    if (responseIdCol === -1) responseIdCol = headers.indexOf("Timestamp");
    var statusCol = headers.indexOf("Status");
    var checklistJsonCol = headers.indexOf("checklist_json");
    
    var isMaster = (role === "master");
    var allowedStores = [];
    if (!isMaster && storesAllowedStr && storesAllowedStr !== "ALL") {
      allowedStores = storesAllowedStr.split(",").map(function(s) { return s.trim().toUpperCase(); });
    }

    var submissions = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var storeCodeStr = storeCodeCol !== -1 ? String(row[storeCodeCol]).trim().toUpperCase() : "";
      var cleanCode = storeCodeStr.indexOf(" - ") !== -1 ? storeCodeStr.split(" - ")[0].trim() : storeCodeStr;
      var asmNameStr = asmCol !== -1 ? String(row[asmCol]).trim() : "";

      // RBAC Filter: Master sees all. ASM sees assigned stores or submissions by their name/username.
      if (!isMaster && allowedStores.length > 0) {
        var matchStore = allowedStores.some(function(as) { return cleanCode.indexOf(as) !== -1 || as.indexOf(cleanCode) !== -1; });
        var matchAsm = username && asmNameStr.toLowerCase().indexOf(String(username).toLowerCase()) !== -1;
        if (!matchStore && !matchAsm) continue;
      }

      var rawDate = row[dateCol];
      var dateStr = "";
      if (rawDate instanceof Date) {
        dateStr = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
      } else {
        dateStr = String(rawDate);
      }
      
      submissions.push({
        rowIdx: i + 1,
        responseId: String(row[responseIdCol]),
        storeCode: storeCodeCol !== -1 ? String(row[storeCodeCol]) : "",
        storeName: storeNameCol !== -1 ? String(row[storeNameCol]) : "",
        reportDate: dateStr,
        asmName: asmNameStr,
        status: statusCol !== -1 ? String(row[statusCol]) : "pending",
        hasChecklist: checklistJsonCol !== -1 && String(row[checklistJsonCol]).trim().length > 0
      });
    }
    
    submissions.reverse();
    return { success: true, submissions: submissions };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getSubmissionDetail(rowIdx) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim(); });
    var rowValues = sheet.getRange(rowIdx, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    var rowData = {};
    headers.forEach(function(h, idx) {
      rowData[h] = rowValues[idx];
    });
    
    return { success: true, data: rowData };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// -------------------------------------------------------------
// AUTOMATED EMAIL NOTIFICATION (PHASE 4)
// -------------------------------------------------------------
function getAsmEmail(asmName) {
  // 1. Try to find in StoreMapping sheet
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(STORE_MAPPING_SHEET);
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      var headers = data[0].map(function(h) { return String(h).trim(); });
      var asmCol = headers.indexOf("ASM");
      var emailCol = headers.indexOf("ASMEmail");
      if (emailCol === -1) emailCol = headers.indexOf("Email");
      
      if (asmCol !== -1 && emailCol !== -1) {
        for (var i = 1; i < data.length; i++) {
          if (String(data[i][asmCol]).trim() === asmName) {
            var email = String(data[i][emailCol]).trim();
            if (email) return email;
          }
        }
      }
    }
  } catch(e) {
    console.warn("Error getting ASM email from sheet: " + e.toString());
  }
  
  // 2. Fallback to hardcoded mapping
  var fallbackMap = {
    "Khôi": "khoind@anphuoc.com.vn",
    "Tiên": "tienntk@anphuoc.com.vn",
    "Tín": "tintv@anphuoc.com.vn",
    "Dũng": "dungnv@anphuoc.com.vn",
    "Linh": "linhnt@anphuoc.com.vn",
    "Quân": "quannh@anphuoc.com.vn",
    "Hương": "huonglt@anphuoc.com.vn",
    "Nhi": "nhinh@anphuoc.com.vn",
    "Lâm": "lamnt@anphuoc.com.vn"
  };
  return fallbackMap[asmName] || "khoind@anphuoc.com.vn";
}

function sendReportEmail(postData) {
  try {
    var folder = getOrCreateReportsFolder();
    
    // Save base64 files if present and convert to fileIds
    if (postData.pdfBase64 && postData.pdfName) {
      try {
        var decoded = Utilities.base64Decode(postData.pdfBase64.split(",")[1]);
        var blob = Utilities.newBlob(decoded, "application/pdf", postData.pdfName);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        postData.pdfFileId = file.getId();
      } catch(e) {
        console.warn("Lỗi lưu file PDF từ base64: " + e.toString());
      }
    }
    
    if (postData.docxBase64 && postData.docxName) {
      try {
        var decoded = Utilities.base64Decode(postData.docxBase64.split(",")[1]);
        var blob = Utilities.newBlob(decoded, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", postData.docxName);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        postData.docxFileId = file.getId();
      } catch(e) {
        console.warn("Lỗi lưu file DOCX từ base64: " + e.toString());
      }
    }
    
    if (postData.pptxBase64 && postData.pptxName) {
      try {
        var decoded = Utilities.base64Decode(postData.pptxBase64.split(",")[1]);
        var blob = Utilities.newBlob(decoded, "application/vnd.openxmlformats-officedocument.presentationml.presentation", postData.pptxName);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        postData.pptxFileId = file.getId();
      } catch(e) {
        console.warn("Lỗi lưu file PPTX từ base64: " + e.toString());
      }
    }
    
    if (postData.xlsxBase64 && postData.xlsxName) {
      try {
        var decoded = Utilities.base64Decode(postData.xlsxBase64.split(",")[1]);
        var blob = Utilities.newBlob(decoded, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", postData.xlsxName);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        postData.xlsxFileId = file.getId();
      } catch(e) {
        console.warn("Lỗi lưu file XLSX từ base64: " + e.toString());
      }
    }

    var attachments = [];
    var totalAttachmentSize = 0;
    var maxAttachmentSize = 20 * 1024 * 1024; // 20 MB
    
    if (postData.pdfFileId) {
      try {
        var pdfFile = DriveApp.getFileById(postData.pdfFileId);
        var pdfBlob = pdfFile.getBlob();
        if (pdfBlob.getBytes().length <= maxAttachmentSize) {
          attachments.push(pdfBlob);
          totalAttachmentSize += pdfBlob.getBytes().length;
        }
      } catch(e) {
        console.warn("Lỗi đính kèm PDF: " + e.toString());
      }
    }
    
    if (postData.docxFileId) {
      try {
        var docxFile = DriveApp.getFileById(postData.docxFileId);
        var docxBlob = docxFile.getBlob();
        if (totalAttachmentSize + docxBlob.getBytes().length <= maxAttachmentSize) {
          attachments.push(docxBlob);
        }
      } catch(e) {
        console.warn("Lỗi đính kèm Word: " + e.toString());
      }
    }
    
    var pptxUrl = "";
    var xlsxUrl = "";
    if (postData.pptxFileId) {
      try {
        var pptxFile = DriveApp.getFileById(postData.pptxFileId);
        pptxFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        pptxUrl = pptxFile.getUrl();
      } catch(e) {}
    }
    if (postData.xlsxFileId) {
      try {
        var xlsxFile = DriveApp.getFileById(postData.xlsxFileId);
        xlsxFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        xlsxUrl = xlsxFile.getUrl();
      } catch(e) {}
    }
    
    var htmlBody = '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px;">' +
      '<h2 style="color: #0a2342; border-bottom: 2px solid #0a2342; padding-bottom: 10px;">Báo cáo Kiểm tra Cửa hàng - StoreVisit Pro</h2>' +
      '<p>Kính gửi Ban Giám đốc và Quản lý,</p>' +
      '<p>Hệ thống xin thông báo kết quả kiểm tra cửa hàng định kỳ:</p>' +
      '<table style="width: 100%; border-collapse: collapse; margin: 20px 0;">' +
      '<tr style="background-color: #f4f7f9;">' +
      '<td style="padding: 10px; font-weight: bold; border: 1px solid #e0e0e0; width: 40%;">Cửa hàng:</td>' +
      '<td style="padding: 10px; border: 1px solid #e0e0e0;">' + (postData.storeName || "") + '</td>' +
      '</tr>' +
      '<tr>' +
      '<td style="padding: 10px; font-weight: bold; border: 1px solid #e0e0e0;">Ngày kiểm tra:</td>' +
      '<td style="padding: 10px; border: 1px solid #e0e0e0;">' + (postData.reportDate || "") + '</td>' +
      '</tr>' +
      '<tr style="background-color: #f4f7f9;">' +
      '<td style="padding: 10px; font-weight: bold; border: 1px solid #e0e0e0;">ASM thực hiện:</td>' +
      '<td style="padding: 10px; border: 1px solid #e0e0e0;">' + (postData.asmName || "") + '</td>' +
      '</tr>' +
      '</table>' +
      '<p>Báo cáo chi tiết được đính kèm trực tiếp trong email này (file PDF và Word).</p>' +
      '<div style="margin: 25px 0; padding: 15px; background-color: #f9f9f9; border-left: 4px solid #2e5b88;">' +
      '<h4 style="margin-top: 0; color: #2e5b88;">Đường link tải file báo cáo dung lượng lớn:</h4>' +
      '<ul style="padding-left: 20px; margin-bottom: 0;">' +
      (pptxUrl ? '<li><strong>Báo cáo PowerPoint (PPTX):</strong> <a href="' + pptxUrl + '" target="_blank">Tải xuống báo cáo PPTX</a></li>' : "") +
      (xlsxUrl ? '<li><strong>Bảng tính chi tiết (Excel):</strong> <a href="' + xlsxUrl + '" target="_blank">Tải xuống bảng tính Excel</a></li>' : "") +
      '</ul>' +
      '</div>' +
      '<p style="font-size: 12px; color: #7f8c8d; margin-top: 30px; border-top: 1px solid #e0e0e0; padding-top: 10px;">' +
      'Email này được gửi tự động từ hệ thống StoreVisit Pro. Vui lòng không trả lời trực tiếp email này.' +
      '</p>' +
      '</div>';
    
    var asmEmail = getAsmEmail(postData.asmName);
    var recipients = asmEmail;
    var ccList = ["khoind@anphuoc.com.vn", "dkhoi86@gmail.com"];
    ccList = ccList.filter(function(email) { return email !== asmEmail; });
    var ccRecipients = ccList.join(",");

    MailApp.sendEmail({
      to: recipients,
      cc: ccRecipients,
      subject: "[StoreVisit Pro] Báo cáo Kiểm tra Cửa hàng - " + (postData.storeName || "") + " (" + (postData.reportDate || "") + ")",
      htmlBody: htmlBody,
      attachments: attachments
    });
    
    return { success: true, message: "Đã gửi email báo cáo thành công.", pdfFileId: postData.pdfFileId, docxFileId: postData.docxFileId, pptxFileId: postData.pptxFileId, xlsxFileId: postData.xlsxFileId };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getOrCreateReportsFolder() {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty("REPORTS_FOLDER_ID");
  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch(e) {
      console.warn("Cached REPORTS_FOLDER_ID invalid or inaccessible, recreating: " + e.toString());
    }
  }

  var spreadsheetFile = DriveApp.getFileById(SPREADSHEET_ID);
  var parentFolders = spreadsheetFile.getParents();
  var parentFolder = parentFolders.hasNext() ? parentFolders.next() : DriveApp.getRootFolder();
  var folderIterator = parentFolder.getFoldersByName("StoreVisit_Reports");
  var folder;
  if (folderIterator.hasNext()) {
    folder = folderIterator.next();
  } else {
    folder = parentFolder.createFolder("StoreVisit_Reports");
  }

  try {
    props.setProperty("REPORTS_FOLDER_ID", folder.getId());
  } catch(e) {
    console.warn("Failed to cache REPORTS_FOLDER_ID: " + e.toString());
  }
  return folder;
}

// -------------------------------------------------------------
// USER AUTHENTICATION & ROLE-BASED ACCESS CONTROL (RBAC)
// -------------------------------------------------------------
var USER_SHEET_NAME = "ASM_Users";

function initASMUsersSheet() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(USER_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(USER_SHEET_NAME);
      var headers = ["username", "password", "full_name", "role", "region", "stores", "status", "created_at"];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#0A2342").setFontColor("#FFFFFF");
      
      var nowIso = new Date().toISOString();
      // 10 Official ASM Accounts Seeded from StoresInfo.xlsx (184 stores total)
      sheet.appendRow(["khoi", "khoi123", "ASM Khôi", "master", "ALL", "ALL", "Active", nowIso]);
      sheet.appendRow(["dung", "123456", "ASM Dũng", "asm", "HCM, Miền Trung- Tây Nguyên", "CMT8, CHOA3, TCHINH, LBBICH, QBINH, KONTUM, GIALAI, PTHIET, DAKLAK, DAKLAK2, DAKLAK5, DAKNONG, DALAT, DALAT2, BAOLOC", "Active", nowIso]);
      sheet.appendRow(["hn", "123456", "ASM Hà Nội", "asm", "HN", "HN1, HN2, HN3, HN4, HN5, HN6, HN8, HN10, HN11, HN12, HTMAU, HN15, HN16, HN17, HN18, HN19, HN20, HN21, HN22, HN23, HN24, HN25, HN26, CAUGIAYHN, HNTN, TUYENQUANG, HP, HP3, HP4, LACHTRAYHP, TBINH, QNINH, THANHHOA, VINH1, VINH2, HATINH, NINHBINH, HAGIANG, VIETTRI, BACNINH, NAMDINH, BACGIANG, VINHYEN, LAOCAI, YENBAI", "Active", nowIso]);
      sheet.appendRow(["huong", "123456", "ASM Hương", "asm", "HCM, Miền Trung- Tây Nguyên, Miền Tây", "KINHDV, HAUGIANG, MYTHO, MYTHO2, CAYLAY, CAOLANH, HONGNGU, SADEC, LXUYEN, LONGXUYEN2, LONGXUYEN3, CHAUDOC, RGIA, RGIA2, RACHGIA3, HATIEN, BENTRE, VINHLONG, VINHLONG2, TRAVINH, CTHO, CTHO2, CTHO3, CTHO6, CANTHO3T2, APHAUGIANG, STR, STR2, BACLIEU, BACLIEU2, CMAU, CMAU2, FLDLTTON", "Active", nowIso]);
      sheet.appendRow(["linh", "123456", "ASM Linh", "asm", "HCM", "SO1, HBT, CAOTH, NTMK, NDC, LVS, 185_3T2, 126_3T2, NVTROI, PDP, GOVAP, LOTTEGV, 901QT, LQDINH, NGUYENOANH, AEONTP, LETRONGTAN", "Active", nowIso]);
      sheet.appendRow(["lam", "123456", "ASM Lâm", "asm", "HCM", "HVPLAZA, VANHANH, SENSECITY, VINCOMTD, VINCOMLVV, VINCOMQ2", "Active", nowIso]);
      sheet.appendRow(["tien", "123456", "ASM Tiên", "asm", "HCM", "DIAMOND, LYTT, NGA6, NTQ1, PTER, NGHUE, VINCOM, TAKA, OIKHIEM, AUCO, CHOA, TCHINH2, NGANHTHU, LVKHUONG, PDL, PDL2, AEONBT, CUCHI", "Active", nowIso]);
      sheet.appendRow(["tin", "123456", "ASM Tín", "asm", "HCM, Miền Đông", "THUDUC, THUDUC2, BPHUOC, BINHLONG, BCDN, VINCOMBH, BIENHOA, BHNAQUOC, TAMHIEP, LONGKHANH, LONGTHANH, TAYNINH, BENLUC, LONGAN", "Active", nowIso]);
      sheet.appendRow(["quan", "123456", "ASM Quân", "asm", "Miền Trung- Tây Nguyên", "QUANGTRI, BIGCHUE, HUE2, DN, DN2, DN3, DN4, DN5DBP, TAMKY, QNGAI, QUINHON2, QUYNHON3, NHT3, NHT2, CAMRANH, PHANRANG2, PHUYEN, QUYNHON", "Active", nowIso]);
      sheet.appendRow(["ni", "123456", "ASM Ni", "asm", "HCM", "ONLINEWEB", "Active", nowIso]);
      sheet.appendRow(["khoind", "khoi123", "ASM Khôi", "master", "ALL", "ALL", "Active", nowIso]);
      Logger.log("✅ Đã khởi tạo sheet " + USER_SHEET_NAME + " với 10 tài khoản ASM chính thức (184 cửa hàng).");
    }
    return sheet;
  } catch(e) {
    Logger.log("Lỗi khởi tạo ASM_Users sheet: " + e.toString());
    return null;
  }
}

function loginUser(username, password) {
  try {
    if (!username || !password) {
      return { success: false, error: "Vui lòng nhập tên đăng nhập và mật khẩu." };
    }
    var sheet = initASMUsersSheet();
    if (!sheet) {
      return { success: false, error: "Không thể mở cơ sở dữ liệu người dùng." };
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: false, error: "Chưa có tài khoản nào được tạo." };
    }
    
    var headers = data[0];
    var uIdx = headers.indexOf("username");
    var pIdx = headers.indexOf("password");
    var nIdx = headers.indexOf("full_name");
    var rIdx = headers.indexOf("role");
    var regIdx = headers.indexOf("region");
    var sIdx = headers.indexOf("stores");
    
    var searchUser = String(username).trim().toLowerCase();
    var searchPass = String(password).trim();
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var uVal = String(row[uIdx]).trim().toLowerCase();
      var pVal = String(row[pIdx]).trim();
      
      if (uVal === searchUser) {
        if (pVal === searchPass) {
          var userObj = {
            username: String(row[uIdx]).trim(),
            fullName: String(row[nIdx] || row[uIdx]).trim(),
            role: String(row[rIdx] || "asm").trim().toLowerCase(),
            region: String(row[regIdx] || "").trim(),
            stores: String(row[sIdx] || "ALL").trim()
          };
          return { success: true, user: userObj };
        } else {
          return { success: false, error: "Mật khẩu không chính xác." };
        }
      }
    }
    
    return { success: false, error: "Tên đăng nhập không tồn tại." };
  } catch(e) {
    return { success: false, error: "Lỗi hệ thống đăng nhập: " + e.toString() };
  }
}

function changeUserPassword(username, oldPassword, newPassword) {
  try {
    if (!username || !oldPassword || !newPassword) {
      return { success: false, error: "Thiếu thông tin mật khẩu cũ hoặc mật khẩu mới." };
    }
    if (String(newPassword).trim().length < 4) {
      return { success: false, error: "Mật khẩu mới phải có ít nhất 4 ký tự." };
    }
    
    var sheet = initASMUsersSheet();
    if (!sheet) return { success: false, error: "Không mở được bảng người dùng." };
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var uIdx = headers.indexOf("username");
    var pIdx = headers.indexOf("password");
    
    var searchUser = String(username).trim().toLowerCase();
    var searchOldPass = String(oldPassword).trim();
    
    for (var i = 1; i < data.length; i++) {
      var uVal = String(data[i][uIdx]).trim().toLowerCase();
      var pVal = String(data[i][pIdx]).trim();
      
      if (uVal === searchUser) {
        if (pVal !== searchOldPass) {
          return { success: false, error: "Mật khẩu hiện tại không đúng." };
        }
        sheet.getRange(i + 1, pIdx + 1).setValue(String(newPassword).trim());
        return { success: true, message: "Đổi mật khẩu thành công!" };
      }
    }
    return { success: false, error: "Không tìm thấy tài khoản để đổi mật khẩu." };
  } catch(e) {
    return { success: false, error: "Lỗi đổi mật khẩu: " + e.toString() };
  }
}

function getInspectionHistory(username, role, storesAllowedStr) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { success: false, error: "Không tìm thấy sheet báo cáo." };
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, history: [] };
    
    var headers = data[0];
    var subIdIdx = headers.indexOf("submission_id");
    var dateIdx = headers.indexOf("Ngày kiểm tra");
    var storeIdx = headers.indexOf("Mã cửa hàng");
    var asmIdx = headers.indexOf("QLKD/ASM");
    var chtIdx = headers.indexOf("Tên CHT");
    var frontageIdx = headers.indexOf("rating_frontage");
    var innerIdx = headers.indexOf("rating_inner");
    var merchIdx = headers.indexOf("rating_merch");
    var staffIdx = headers.indexOf("rating_staff");
    var csvcIdx = headers.indexOf("rating_csvc");
    var issuesIdx = headers.indexOf("Vấn đề tồn đọng");
    var planIdx = headers.indexOf("Kế hoạch khắc phục");
    var deadlineIdx = headers.indexOf("Thời hạn xử lý");
    var jsonIdx = headers.indexOf("checklist_json");
    var statusIdx = headers.indexOf("Status");
    
    var history = [];
    var isMaster = (role === "master");
    
    var allowedStores = [];
    if (!isMaster && storesAllowedStr && storesAllowedStr !== "ALL") {
      allowedStores = storesAllowedStr.split(",").map(function(s) { return s.trim().toUpperCase(); });
    }
    
    // Reverse loop to get latest submissions first
    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      var storeCode = String(row[storeIdx] || "").trim().toUpperCase();
      if (storeCode.indexOf(" - ") !== -1) {
        storeCode = storeCode.split(" - ")[0].trim();
      }
      
      // Filter logic: Master sees all; ASM sees assigned stores or created by them
      if (!isMaster && allowedStores.length > 0) {
        var asmRowName = String(row[asmIdx] || "").trim().toLowerCase();
        var matchStore = allowedStores.some(function(as) { return storeCode.indexOf(as) !== -1 || as.indexOf(storeCode) !== -1; });
        var matchAsm = (asmRowName.indexOf(String(username).toLowerCase()) !== -1);
        if (!matchStore && !matchAsm) continue;
      }
      
      history.push({
        rowIndex: i + 1,
        submissionId: String(row[subIdIdx] || ""),
        reportDate: String(row[dateIdx] || ""),
        storeCode: String(row[storeIdx] || ""),
        asmName: String(row[asmIdx] || ""),
        chtName: String(row[chtIdx] || ""),
        ratingFrontage: String(row[frontageIdx] || "Chưa đánh giá"),
        ratingInner: String(row[innerIdx] || "Chưa đánh giá"),
        ratingMerch: String(row[merchIdx] || "Chưa đánh giá"),
        ratingStaff: String(row[staffIdx] || "Chưa đánh giá"),
        ratingCSVC: String(row[csvcIdx] || "Chưa đánh giá"),
        pendingIssues: String(row[issuesIdx] || ""),
        actionPlan: String(row[planIdx] || ""),
        actionDeadline: String(row[deadlineIdx] || ""),
        status: String(row[statusIdx] || ""),
        checklistJson: String(row[jsonIdx] || "")
      });
      
      if (history.length >= 150) break; // Limit 150 recent items for fast load
    }
    
    return { success: true, history: history };
  } catch(e) {
    return { success: false, error: "Lỗi tải lịch sử báo cáo: " + e.toString() };
  }
}

// -------------------------------------------------------------
// USER MANAGEMENT FUNCTIONS FOR MASTER & ADMIN ACCOUNTS
// -------------------------------------------------------------
function isUserAdminOrMaster(username) {
  if (!username) return false;
  var searchUser = String(username).trim().toLowerCase();
  if (searchUser === "khoi" || searchUser === "khoind" || searchUser === "admin") return true;
  var sheet = initASMUsersSheet();
  if (!sheet) return false;
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return false;
  var headers = data[0];
  var uIdx = headers.indexOf("username");
  var rIdx = headers.indexOf("role");
  
  for (var i = 1; i < data.length; i++) {
    var uVal = String(data[i][uIdx]).trim().toLowerCase();
    if (uVal === searchUser) {
      var role = String(data[i][rIdx] || "").trim().toLowerCase();
      return (role === "master" || role === "admin");
    }
  }
  return false;
}

function getAllUsers(requesterUsername) {
  try {
    if (!isUserAdminOrMaster(requesterUsername)) {
      return { success: false, error: "Bạn không có quyền quản trị tài khoản người dùng." };
    }
    var sheet = initASMUsersSheet();
    if (!sheet) return { success: false, error: "Không mở được bảng người dùng." };
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, users: [] };
    
    var headers = data[0];
    var uIdx = headers.indexOf("username");
    var nIdx = headers.indexOf("full_name");
    var rIdx = headers.indexOf("role");
    var regIdx = headers.indexOf("region");
    var sIdx = headers.indexOf("stores");
    var stIdx = headers.indexOf("status");
    if (stIdx === -1) stIdx = headers.length;
    
    var users = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[uIdx]) continue;
      users.push({
        username: String(row[uIdx]).trim(),
        fullName: String(row[nIdx] || row[uIdx]).trim(),
        role: String(row[rIdx] || "asm").trim().toLowerCase(),
        region: String(row[regIdx] || "").trim(),
        stores: String(row[sIdx] || "ALL").trim(),
        status: String(row[stIdx] || "Active").trim()
      });
    }
    return { success: true, users: users };
  } catch(e) {
    return { success: false, error: "Lỗi tải danh sách người dùng: " + e.toString() };
  }
}

function saveUserAccount(requesterUsername, userData) {
  try {
    if (!isUserAdminOrMaster(requesterUsername)) {
      return { success: false, error: "Bạn không có quyền thực hiện thao tác này." };
    }
    if (!userData || !userData.username) {
      return { success: false, error: "Tên đăng nhập không được để trống." };
    }
    
    var sheet = initASMUsersSheet();
    if (!sheet) return { success: false, error: "Không mở được bảng người dùng." };
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var uIdx = headers.indexOf("username");
    var pIdx = headers.indexOf("password");
    var nIdx = headers.indexOf("full_name");
    var rIdx = headers.indexOf("role");
    var regIdx = headers.indexOf("region");
    var sIdx = headers.indexOf("stores");
    var stIdx = headers.indexOf("status");
    if (stIdx === -1) {
      stIdx = headers.length;
      sheet.getRange(1, stIdx + 1).setValue("status").setFontWeight("bold");
    }
    
    var targetUser = String(userData.username).trim().toLowerCase();
    var foundIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][uIdx]).trim().toLowerCase() === targetUser) {
        foundIndex = i + 1; // 1-indexed row in sheet
        break;
      }
    }
    
    var roleVal = String(userData.role || "asm").trim().toLowerCase();
    var regionVal = String(userData.region || "").trim();
    var storesVal = String(userData.stores || "ALL").trim();
    var fullNameVal = String(userData.fullName || userData.username).trim();
    var statusVal = String(userData.status || "Active").trim();
    
    if (foundIndex > 0) {
      // Update existing user
      sheet.getRange(foundIndex, nIdx + 1).setValue(fullNameVal);
      sheet.getRange(foundIndex, rIdx + 1).setValue(roleVal);
      sheet.getRange(foundIndex, regIdx + 1).setValue(regionVal);
      sheet.getRange(foundIndex, sIdx + 1).setValue(storesVal);
      sheet.getRange(foundIndex, stIdx + 1).setValue(statusVal);
      if (userData.password && String(userData.password).trim()) {
        sheet.getRange(foundIndex, pIdx + 1).setValue(String(userData.password).trim());
      }
      return { success: true, message: "Đã cập nhật thông tin tài khoản " + userData.username + " thành công." };
    } else {
      // Add new user
      var newPass = userData.password ? String(userData.password).trim() : "123456";
      var newRow = [];
      newRow[uIdx] = userData.username.trim();
      newRow[pIdx] = newPass;
      newRow[nIdx] = fullNameVal;
      newRow[rIdx] = roleVal;
      newRow[regIdx] = regionVal;
      newRow[sIdx] = storesVal;
      newRow[stIdx] = statusVal;
      sheet.appendRow(newRow);
      return { success: true, message: "Đã tạo mới tài khoản " + userData.username + " thành công." };
    }
  } catch(e) {
    return { success: false, error: "Lỗi lưu tài khoản: " + e.toString() };
  }
}

function resetUserPasswordByAdmin(requesterUsername, targetUsername, newPassword) {
  try {
    if (!isUserAdminOrMaster(requesterUsername)) {
      return { success: false, error: "Bạn không có quyền reset mật khẩu tài khoản khác." };
    }
    if (!targetUsername || !newPassword) {
      return { success: false, error: "Thiếu thông tin tài khoản hoặc mật khẩu mới." };
    }
    
    var sheet = initASMUsersSheet();
    if (!sheet) return { success: false, error: "Không mở được bảng người dùng." };
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var uIdx = headers.indexOf("username");
    var pIdx = headers.indexOf("password");
    
    var target = String(targetUsername).trim().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][uIdx]).trim().toLowerCase() === target) {
        sheet.getRange(i + 1, pIdx + 1).setValue(String(newPassword).trim());
        return { success: true, message: "Đã reset mật khẩu cho tài khoản " + targetUsername + " thành công!" };
      }
    }
    return { success: false, error: "Không tìm thấy tài khoản " + targetUsername };
  } catch(e) {
    return { success: false, error: "Lỗi reset mật khẩu: " + e.toString() };
  }
}

function toggleUserStatus(requesterUsername, targetUsername, newStatus) {
  try {
    if (!isUserAdminOrMaster(requesterUsername)) {
      return { success: false, error: "Bạn không có quyền đổi trạng thái tài khoản." };
    }
    var sheet = initASMUsersSheet();
    if (!sheet) return { success: false, error: "Không mở được bảng người dùng." };
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var uIdx = headers.indexOf("username");
    var stIdx = headers.indexOf("status");
    if (stIdx === -1) {
      stIdx = headers.length;
      sheet.getRange(1, stIdx + 1).setValue("status").setFontWeight("bold");
    }
    
    var target = String(targetUsername).trim().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][uIdx]).trim().toLowerCase() === target) {
        sheet.getRange(i + 1, stIdx + 1).setValue(newStatus);
        return { success: true, message: "Đã đổi trạng thái tài khoản " + targetUsername + " thành " + newStatus };
      }
    }
    return { success: false, error: "Không tìm thấy tài khoản " + targetUsername };
  } catch(e) {
    return { success: false, error: "Lỗi cập nhật trạng thái: " + e.toString() };
  }
}

function syncASMUsersFromStoresInfo(requesterUsername, storeDataList) {
  try {
    if (!isUserAdminOrMaster(requesterUsername)) {
      return { success: false, error: "Bạn không có quyền thực hiện đồng bộ tài khoản." };
    }
    if (!storeDataList || !Array.isArray(storeDataList) || storeDataList.length === 0) {
      return { success: false, error: "Dữ liệu danh sách cửa hàng không hợp lệ." };
    }
    
    // Group stores by ASM name
    var asmMap = {};
    var validStoreCount = 0;
    
    for (var i = 0; i < storeDataList.length; i++) {
      var item = storeDataList[i];
      if (!item || !item.code) continue;
      var code = String(item.code).trim().toUpperCase();
      var asmName = String(item.asm || "Khác").trim();
      var region = String(item.region || "").trim();
      if (!code) continue;
      
      validStoreCount++;
      if (!asmMap[asmName]) {
        asmMap[asmName] = {
          asmName: asmName,
          regions: [],
          stores: []
        };
      }
      if (region && asmMap[asmName].regions.indexOf(region) === -1) {
        asmMap[asmName].regions.push(region);
      }
      if (asmMap[asmName].stores.indexOf(code) === -1) {
        asmMap[asmName].stores.push(code);
      }
    }
    
    var sheet = initASMUsersSheet();
    if (!sheet) return { success: false, error: "Không mở được bảng ASM_Users." };
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var uIdx = headers.indexOf("username");
    var pIdx = headers.indexOf("password");
    var nIdx = headers.indexOf("full_name");
    var rIdx = headers.indexOf("role");
    var regIdx = headers.indexOf("region");
    var sIdx = headers.indexOf("stores");
    var stIdx = headers.indexOf("status");
    if (stIdx === -1) {
      stIdx = headers.length;
      sheet.getRange(1, stIdx + 1).setValue("status").setFontWeight("bold");
    }
    
    // Map existing rows by username
    var existingRows = {};
    for (var r = 1; r < data.length; r++) {
      var uName = String(data[r][uIdx]).trim().toLowerCase();
      if (uName) existingRows[uName] = r + 1; // 1-indexed row
    }
    
    function removeAccents(str) {
      return str.normalize('NFD').replace(/[\u0300-\u036f]/g, '')
                .replace(/đ/g, 'd').replace(/Đ/g, 'D')
                .replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
    }
    
    var syncedASMCount = 0;
    var asmKeys = Object.keys(asmMap);
    
    for (var k = 0; k < asmKeys.length; k++) {
      var asmName = asmKeys[k];
      var asmObj = asmMap[asmName];
      var uname = removeAccents(asmName);
      if (!uname) uname = "asm_" + k;
      
      var fullName = "ASM " + asmName;
      var role = (uname === "khoi" || uname === "khoind") ? "master" : "asm";
      var regionStr = asmObj.regions.join(", ");
      var storesStr = (uname === "khoi" || uname === "khoind") ? "ALL" : asmObj.stores.join(", ");
      
      if (existingRows[uname]) {
        // Update existing row
        var rowNum = existingRows[uname];
        sheet.getRange(rowNum, nIdx + 1).setValue(fullName);
        sheet.getRange(rowNum, rIdx + 1).setValue(role);
        sheet.getRange(rowNum, regIdx + 1).setValue(regionStr);
        sheet.getRange(rowNum, sIdx + 1).setValue(storesStr);
      } else {
        // Create new row
        var defaultPass = (uname === "khoi") ? "khoi123" : "123456";
        var newRow = [];
        newRow[uIdx] = uname;
        newRow[pIdx] = defaultPass;
        newRow[nIdx] = fullName;
        newRow[rIdx] = role;
        newRow[regIdx] = regionStr;
        newRow[sIdx] = storesStr;
        newRow[stIdx] = "Active";
        sheet.appendRow(newRow);
      }
      syncedASMCount++;
    }
    
    return {
      success: true,
      validStoreCount: validStoreCount,
      syncedASMCount: syncedASMCount,
      message: "Đã đồng bộ thành công " + validStoreCount + " cửa hàng và " + syncedASMCount + " tài khoản ASM từ StoresInfo.xlsx!"
    };
  } catch(e) {
    return { success: false, error: "Lỗi đồng bộ dữ liệu Excel: " + e.toString() };
  }
}
