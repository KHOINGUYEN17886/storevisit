import re
from datetime import datetime

def remove_accents(str_val: str) -> str:
    if not str_val:
        return ""
    s = str(str_val).strip()
    s = re.sub(r'[àáạảãâầấậẩẫăằắặẳẵ]', 'a', s)
    s = re.sub(r'[ÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶẲẴ]', 'A', s)
    s = re.sub(r'[èéẹẻẽêềếệểễ]', 'e', s)
    s = re.sub(r'[ÈÉẸẺẼÊỀẾỆỂỄ]', 'E', s)
    s = re.sub(r'[ìíịỉĩ]', 'i', s)
    s = re.sub(r'[ÌÍỊỈĨ]', 'I', s)
    s = re.sub(r'[òóọỏõôồốộổỗơờớợởỡ]', 'o', s)
    s = re.sub(r'[ÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠ]', 'O', s)
    s = re.sub(r'[ùúụủũưừứựửữ]', 'u', s)
    s = re.sub(r'[ÙÚỤỦŨƯỪỨỰỬỮ]', 'U', s)
    s = re.sub(r'[ỳýỵỷỹ]', 'y', s)
    s = re.sub(r'[ỲÝỴỶỸ]', 'Y', s)
    s = re.sub(r'[đ]', 'd', s)
    s = re.sub(r'[Đ]', 'D', s)
    return s

def clean_name_for_filename(name: str) -> str:
    """Removes accents and non-alphanumeric characters for clean, safe filenames."""
    if not name:
        return ""
    no_acc = remove_accents(name)
    clean = re.sub(r'[^a-zA-Z0-9\s\-]', '', no_acc)
    clean = re.sub(r'\s+', ' ', clean).strip()
    return clean

def format_date_ddmmyyyy(report_date: str) -> str:
    """Formats any date string into DDMMYYYY format."""
    if not report_date:
        return datetime.now().strftime("%d%m%Y")
    
    date_str = str(report_date).strip()
    m_iso = re.match(r'^(\d{4})[\-/](\d{1,2})[\-/](\d{1,2})', date_str)
    m_vn = re.match(r'^(\d{1,2})[\-/](\d{1,2})[\-/](\d{4})', date_str)
    
    if m_iso:
        y, m, d = m_iso.group(1), m_iso.group(2).zfill(2), m_iso.group(3).zfill(2)
        return f"{d}{m}{y}"
    elif m_vn:
        d, m, y = m_vn.group(1).zfill(2), m_vn.group(2).zfill(2), m_vn.group(3)
        return f"{d}{m}{y}"
    
    digits = re.sub(r'\D', '', date_str)
    if len(digits) == 8:
        return digits
    return datetime.now().strftime("%d%m%Y")

def format_store_output_filename(store_name: str, asm_name: str, report_date: str, ext: str) -> str:
    """
    Formats single store report filenames as: [Tên Cửa Hàng]-[ASM]-[Ngày kiểm tra].[ext]
    Example: 'An Phuoc Binh Duong-Nguyen Van Dung-25072026.pptx'
    """
    clean_store = clean_name_for_filename(store_name) or "Store"
    clean_asm = clean_name_for_filename(asm_name) or "ASM"
    date_part = format_date_ddmmyyyy(report_date)
    clean_ext = ext.lstrip(".")
    return f"{clean_store}-{clean_asm}-{date_part}.{clean_ext}"

def format_cluster_output_filename(cluster_name: str, asm_name: str, report_date: str, ext: str) -> str:
    """
    Formats cluster report filenames as: BaoCao_Cum_[TenCum]-[ASM]-[DDMMYYYY].[ext]
    Example: 'BaoCao_Cum_TPHCM-NguyenVanDung-18082026.pptx'
    """
    clean_cum = clean_name_for_filename(cluster_name) or "Cum"
    clean_asm = clean_name_for_filename(asm_name) or "ASM"
    date_part = format_date_ddmmyyyy(report_date)
    clean_ext = ext.lstrip(".")
    return f"BaoCao_Cum_{clean_cum}-{clean_asm}-{date_part}.{clean_ext}"

def format_executive_output_filename(period_type: str, asm_name: str, report_date: str, ext: str) -> str:
    """
    Formats executive combo report filenames as: BaoCao_Executive_[PeriodTag]-[ASM]-[DDMMYYYY].[ext]
    Example: 'BaoCao_Executive_Tuan-Khoi-31072026.pptx'
    """
    period_tag = "Tuan" if period_type == "weekly" else ("Thang" if period_type == "monthly" else "Quy")
    clean_asm = clean_name_for_filename(asm_name) or "TongHop"
    date_part = format_date_ddmmyyyy(report_date)
    clean_ext = ext.lstrip(".")
    return f"BaoCao_Executive_{period_tag}-{clean_asm}-{date_part}.{clean_ext}"
