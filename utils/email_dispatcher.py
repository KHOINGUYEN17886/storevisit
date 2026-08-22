import os
import re
import smtplib
import base64
import requests
import yaml
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from typing import Dict, List, Optional, Tuple, Any

class EmailDispatcher:
    """
    Enterprise Multi-Provider Email Dispatcher.
    Hỗ trợ 3 kênh gửi email doanh nghiệp:
      1. Direct SMTP (Gmail, Office365, An Phước Mail Server)
      2. Google Apps Script WebApp Dispatcher (Cloud Serverless)
      3. Microsoft Outlook Desktop COM Automation (Windows Client)
    """

    def __init__(self, config_dict: Optional[Dict[str, Any]] = None, config_path: Optional[str] = None):
        self.config_path = config_path or os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "config", "app_config.yaml"))
        self.config = config_dict or self._load_config_from_disk()

    def _load_config_from_disk(self) -> Dict[str, Any]:
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    return yaml.safe_load(f) or {}
            except Exception as e:
                print(f"[EmailDispatcher] Error loading config from {self.config_path}: {e}")
        return {}

    def get_email_config(self) -> Dict[str, Any]:
        """Lấy nhánh cấu hình email từ config hiện hành."""
        email_cfg = self.config.get("email", {})
        if not email_cfg:
            webapp_url = self.config.get("google", {}).get("webapp_url", "")
            email_cfg = {
                "active_provider": "google_webapp",
                "smtp": {
                    "host": "smtp.gmail.com",
                    "port": 587,
                    "use_tls": True,
                    "sender_name": "Hệ thống Báo cáo StoreVisit - An Phước",
                    "sender_email": "khoind@anphuoc.com.vn",
                    "sender_password": ""
                },
                "google_webapp": {
                    "url": webapp_url
                },
                "default_cc": "",
                "default_subject_template": "BÁO CÁO KIỂM TRA CỬA HÀNG {store_name} ({store_code}) - NGÀY {report_date}",
                "default_body_template": "Kính gửi Ban Giám Đốc và Quản Lý Khu Vực,\n\nĐính kèm là bộ tài liệu báo cáo kiểm tra thực tế tại cửa hàng {store_name} ({store_code}) thực hiện vào ngày {report_date}.\n\nTrân trọng,\n{asm_name} - QLKD phụ trách"
            }
        return email_cfg

    def save_email_config(self, new_email_cfg: Dict[str, Any]) -> bool:
        """Cập nhật và ghi đè cấu hình email vào file app_config.yaml an toàn."""
        try:
            current_cfg = self._load_config_from_disk()
            current_cfg["email"] = new_email_cfg
            
            if "google_webapp" in new_email_cfg and "url" in new_email_cfg["google_webapp"]:
                if "google" not in current_cfg:
                    current_cfg["google"] = {}
                current_cfg["google"]["webapp_url"] = new_email_cfg["google_webapp"]["url"]

            with open(self.config_path, "w", encoding="utf-8") as f:
                yaml.dump(current_cfg, f, allow_unicode=True, default_flow_style=False, sort_keys=False)
            
            self.config = current_cfg
            return True
        except Exception as e:
            print(f"[EmailDispatcher] Error saving email config: {e}")
            return False

    @staticmethod
    def render_template(template_str: str, context: Dict[str, Any]) -> str:
        """Điền các biến động vào mẫu nội dung/tiêu đề email."""
        if not template_str:
            return ""
        rendered = template_str
        for k, v in context.items():
            placeholder = "{" + str(k) + "}"
            rendered = rendered.replace(placeholder, str(v or ""))
        return rendered

    def send_report_email(
        self,
        to_email: str,
        subject: str,
        body: str,
        cc_emails: Optional[str] = "",
        attachments: Optional[List[str]] = None,
        context: Optional[Dict[str, Any]] = None,
        provider: Optional[str] = None
    ) -> Tuple[bool, str]:
        """
        Hàm gửi email chính: Tự động điều phối theo Provider được chọn hoặc chỉ định.
        Returns: (success: bool, message: str)
        """
        to_email = (to_email or "").strip()
        if not to_email or "@" not in to_email:
            return False, "Địa chỉ email người nhận (To) không hợp lệ."

        email_cfg = self.get_email_config()
        selected_provider = provider or email_cfg.get("active_provider", "google_webapp")
        
        cc_list = []
        if cc_emails:
            for item in re.split(r"[,;]+", cc_emails):
                item = item.strip()
                if item and "@" in item:
                    cc_list.append(item)

        valid_attachments = []
        if attachments:
            for fpath in attachments:
                if fpath and os.path.exists(fpath) and os.path.getsize(fpath) > 0:
                    valid_attachments.append(fpath)

        context = context or {}

        print(f"[EmailDispatcher] Dispatching email via provider '{selected_provider}' to '{to_email}' (CC: {cc_list})...")

        if selected_provider == "smtp":
            return self._send_via_smtp(to_email, cc_list, subject, body, valid_attachments, email_cfg.get("smtp", {}))
        elif selected_provider == "outlook":
            return self._send_via_outlook(to_email, cc_list, subject, body, valid_attachments)
        else:
            return self._send_via_gas(to_email, cc_list, subject, body, valid_attachments, context, email_cfg.get("google_webapp", {}))

    def _send_via_smtp(
        self,
        to_email: str,
        cc_list: List[str],
        subject: str,
        body: str,
        attachments: List[str],
        smtp_cfg: Dict[str, Any]
    ) -> Tuple[bool, str]:
        """Gửi email trực tiếp qua giao thức SMTP (SSL / TLS)."""
        host = smtp_cfg.get("host", "smtp.gmail.com")
        port = int(smtp_cfg.get("port", 587))
        use_tls = smtp_cfg.get("use_tls", True)
        sender_email = smtp_cfg.get("sender_email", "")
        sender_name = smtp_cfg.get("sender_name", "StoreVisit Report")
        sender_password = smtp_cfg.get("sender_password", "")

        if not sender_email or not sender_password:
            return False, "Thiếu cấu hình Email hoặc Mật khẩu SMTP trong phần Cài đặt."

        try:
            msg = MIMEMultipart()
            msg["From"] = f"{sender_name} <{sender_email}>"
            msg["To"] = to_email
            if cc_list:
                msg["Cc"] = ", ".join(cc_list)
            msg["Subject"] = subject
            msg["Date"] = datetime.now().strftime("%a, %d %b %Y %H:%M:%S +0700")

            msg.attach(MIMEText(body, "plain", "utf-8"))

            for file_path in attachments:
                fname = os.path.basename(file_path)
                with open(file_path, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f'attachment; filename="{fname}"')
                msg.attach(part)

            all_recipients = [to_email] + cc_list

            if port == 465:
                with smtplib.SMTP_SSL(host, port, timeout=30) as server:
                    server.login(sender_email, sender_password)
                    server.sendmail(sender_email, all_recipients, msg.as_string())
            else:
                with smtplib.SMTP(host, port, timeout=30) as server:
                    if use_tls:
                        server.starttls()
                    server.login(sender_email, sender_password)
                    server.sendmail(sender_email, all_recipients, msg.as_string())

            return True, f"Đã gửi thành công qua SMTP ({host}) tới {to_email}."
        except Exception as e:
            err_msg = str(e)
            print(f"[EmailDispatcher] SMTP sending failed: {err_msg}")
            return False, f"Lỗi gửi SMTP: {err_msg}"

    def _send_via_outlook(
        self,
        to_email: str,
        cc_list: List[str],
        subject: str,
        body: str,
        attachments: List[str]
    ) -> Tuple[bool, str]:
        """Gửi email qua Microsoft Outlook Desktop COM Automation."""
        try:
            import win32com.client
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            
            mail.To = to_email
            if cc_list:
                mail.CC = "; ".join(cc_list)
            mail.Subject = subject
            mail.Body = body
            
            for fpath in attachments:
                mail.Attachments.Add(os.path.abspath(fpath))
                
            mail.Send()
            return True, f"Đã gửi thành công qua Microsoft Outlook Desktop tới {to_email}."
        except ImportError:
            return False, "Thư viện win32com chưa được cài đặt trên hệ thống."
        except Exception as e:
            err_msg = str(e)
            print(f"[EmailDispatcher] Outlook COM sending failed: {err_msg}")
            return False, f"Lỗi gửi qua Outlook Desktop: {err_msg}"

    def _send_via_gas(
        self,
        to_email: str,
        cc_list: List[str],
        subject: str,
        body: str,
        attachments: List[str],
        context: Dict[str, Any],
        gas_cfg: Dict[str, Any]
    ) -> Tuple[bool, str]:
        """Gửi email qua Google Apps Script WebApp."""
        webapp_url = gas_cfg.get("url", "")
        if not webapp_url:
            return False, "Chưa cấu hình Google Apps Script WebApp URL trong phần Cài đặt."

        def get_b64(fpath):
            if fpath and os.path.exists(fpath):
                try:
                    with open(fpath, "rb") as f:
                        return base64.b64encode(f.read()).decode("utf-8")
                except Exception as e:
                    print(f"[EmailDispatcher] Error reading attachment {fpath}: {e}")
            return ""

        pdf_b64 = ""
        docx_b64 = ""
        pptx_b64 = ""
        xlsx_b64 = ""
        
        pdf_name = ""
        docx_name = ""
        pptx_name = ""
        xlsx_name = ""

        for fpath in attachments:
            fname = os.path.basename(fpath).lower()
            if fname.endswith(".pdf") and not pdf_b64:
                pdf_b64 = f"data:application/pdf;base64,{get_b64(fpath)}"
                pdf_name = os.path.basename(fpath)
            elif fname.endswith(".docx") and not docx_b64:
                docx_b64 = f"data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{get_b64(fpath)}"
                docx_name = os.path.basename(fpath)
            elif fname.endswith(".pptx") and not pptx_b64:
                pptx_b64 = f"data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{get_b64(fpath)}"
                pptx_name = os.path.basename(fpath)
            elif fname.endswith(".xlsx") and not xlsx_b64:
                xlsx_b64 = f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{get_b64(fpath)}"
                xlsx_name = os.path.basename(fpath)

        payload = {
            "action": "send_email",
            "confirmSend": True,
            "recipientEmail": to_email,
            "targetEmail": to_email,
            "ccEmails": ", ".join(cc_list),
            "customSubject": subject,
            "customBody": body,
            "storeName": context.get("store_name", context.get("store_code", "Cửa hàng")),
            "reportDate": context.get("report_date", datetime.now().strftime("%d/%m/%Y")),
            "asmName": context.get("asm_name", "QLKD"),
            "pdfBase64": pdf_b64,
            "pdfName": pdf_name,
            "docxBase64": docx_b64,
            "docxName": docx_name,
            "pptxBase64": pptx_b64,
            "pptxName": pptx_name,
            "xlsxBase64": xlsx_b64,
            "xlsxName": xlsx_name
        }

        try:
            resp = requests.post(webapp_url, json=payload, timeout=60, allow_redirects=True)
            res_text = resp.text
            print(f"[EmailDispatcher] GAS Response ({resp.status_code}): {res_text[:300]}")
            
            try:
                res_json = resp.json()
                if res_json.get("success"):
                    return True, f"Đã gửi email thành công qua Google Apps Script tới {to_email}."
                else:
                    return False, f"Google Apps Script báo lỗi: {res_json.get('error', 'Không xác định')}"
            except Exception:
                if resp.status_code == 200 and "success" in res_text.lower():
                    return True, f"Đã gửi email thành công qua Google Apps Script tới {to_email}."
                return False, f"Phản hồi không hợp lệ từ GAS (HTTP {resp.status_code}): {res_text[:200]}"
        except Exception as e:
            err_msg = str(e)
            print(f"[EmailDispatcher] GAS sending failed: {err_msg}")
            return False, f"Lỗi kết nối tới Google Apps Script: {err_msg}"

    def test_connection(self, provider: str, target_email: str = "") -> Tuple[bool, str]:
        """Kiểm tra đường truyền / kết nối cho một Provider cụ thể."""
        email_cfg = self.get_email_config()
        
        if provider == "smtp":
            smtp_cfg = email_cfg.get("smtp", {})
            host = smtp_cfg.get("host", "smtp.gmail.com")
            port = int(smtp_cfg.get("port", 587))
            use_tls = smtp_cfg.get("use_tls", True)
            sender_email = smtp_cfg.get("sender_email", "")
            sender_password = smtp_cfg.get("sender_password", "")
            
            if not sender_email or not sender_password:
                return False, "Vui lòng điền Email gửi và Mật khẩu SMTP trước khi kiểm tra."
            
            try:
                if port == 465:
                    with smtplib.SMTP_SSL(host, port, timeout=15) as server:
                        server.login(sender_email, sender_password)
                else:
                    with smtplib.SMTP(host, port, timeout=15) as server:
                        if use_tls:
                            server.starttls()
                        server.login(sender_email, sender_password)
                return True, f"✓ Kết nối SMTP ({host}:{port}) và xác thực tài khoản {sender_email} THÀNH CÔNG!"
            except Exception as e:
                return False, f"❌ Kết nối SMTP thất bại: {str(e)}"
                
        elif provider == "outlook":
            try:
                import win32com.client
                outlook = win32com.client.Dispatch("Outlook.Application")
                mail = outlook.CreateItem(0)
                del mail
                del outlook
                return True, "✓ Kết nối tới ứng dụng Microsoft Outlook Desktop THÀNH CÔNG!"
            except Exception as e:
                return False, f"❌ Không thể kết nối tới Microsoft Outlook Desktop: {str(e)}"
                
        elif provider == "google_webapp":
            webapp_url = email_cfg.get("google_webapp", {}).get("url", "")
            if not webapp_url:
                return False, "Vui lòng nhập Google Apps Script WebApp URL trước khi kiểm tra."
            try:
                resp = requests.post(webapp_url, json={"action": "getStoreData"}, timeout=15, allow_redirects=True)
                if resp.status_code == 200:
                    return True, "✓ Kết nối tới Google Apps Script WebApp THÀNH CÔNG!"
                else:
                    return False, f"❌ Google Apps Script phản hồi mã lỗi HTTP {resp.status_code}."
            except Exception as e:
                return False, f"❌ Lỗi kết nối tới WebApp URL: {str(e)}"
                
        return False, f"Không hỗ trợ kiểm tra provider: {provider}"
