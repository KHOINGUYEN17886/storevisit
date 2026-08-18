import os
import sys
import json
import uuid
import time
from datetime import datetime
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import yaml

_APP_DIR = os.path.dirname(os.path.abspath(__file__))
if _APP_DIR not in sys.path:
    sys.path.insert(0, _APP_DIR)
import ap_theme  # design system An Phước (navy/crimson + Be Vietnam Pro)

class StoreVisitApp:
    def __init__(self, root):
        self.root = root
        self.root.title("StoreVisit - Tự Động Hóa Báo Cáo Kiểm Tra Cụm Cửa Hàng")
        self.root.geometry("1020x840")
        self.root.minsize(900, 650)
        
        # Paths
        self.root_dir = os.path.dirname(os.path.abspath(__file__))
        self.config_path = os.path.join(self.root_dir, "config/app_config.yaml")
        
        # Load configuration
        with open(self.config_path, "r", encoding="utf-8") as f:
            self.config = yaml.safe_load(f)
            
        self.dim_store_path = self.config["paths"]["dim_store"]
        if not os.path.isabs(self.dim_store_path):
            self.dim_store_path = os.path.join(self.root_dir, self.dim_store_path)
            
        # Watchdog timeouts
        self.timeouts = self.config.get("timeouts", {})
        
        # State variables
        self.worker_process = None
        self.current_job_id = None
        self.cancel_file_path = None
        self.last_event_time = 0
        self.watchdog_thread = None
        self.is_running = False
        self.latest_error_text = ""
        
        # Initialize Google Form Response Cache & Market Survey Cache
        self.form_cache = None
        self.survey_cache = None
        google_config = self.config.get("google", {})
        try:
            from data.form_response_cache import FormResponseCache
            cache_path = os.path.join(self.root_dir, "data/form_cache.json")
            self.form_cache = FormResponseCache(
                cache_path=cache_path,
                credentials_path=google_config.get("credentials_path", ""),
                spreadsheet_id=google_config.get("spreadsheet_id", ""),
                photo_cache_dir=google_config.get("drive_photos_cache_dir", "")
            )
        except Exception as e:
            print(f"Warning: Failed to init FormResponseCache: {e}")
            
        try:
            from data.form_response_cache import MarketSurveyCache
            survey_cache_path = google_config.get("survey_cache_path", os.path.join(self.root_dir, "data/survey_cache.json"))
            self.survey_cache = MarketSurveyCache(
                cache_path=survey_cache_path,
                credentials_path=google_config.get("credentials_path", ""),
                spreadsheet_id=google_config.get("spreadsheet_id", ""),
                photo_cache_dir=google_config.get("survey_photos_dir", "")
            )
        except Exception as e:
            print(f"Warning: Failed to init MarketSurveyCache: {e}")
            
        # Styles & Themes — An Phước design system (navy/crimson + Be Vietnam Pro)
        ap_theme.load_fonts()
        self.T = ap_theme.apply(self.root)
        self.root.configure(bg=ap_theme.SURFACE)
        self.style = ttk.Style(self.root)
        fam = self.T.family
        # Ánh xạ các style-name sẵn có của app sang bảng màu thương hiệu
        self.style.configure("Header.TLabel", font=(fam, 15, "bold"), foreground=ap_theme.NAVY)
        self.style.configure("Accent.TButton", font=(fam, 10, "bold"), foreground="white",
                             background=ap_theme.NAVY, borderwidth=0, focusthickness=0, padding=(12, 7))
        self.style.map("Accent.TButton",
                       background=[("active", ap_theme.NAVY_DARK), ("disabled", "#9AA3B5")])
        self.style.configure("Cancel.TButton", font=(fam, 10, "bold"), foreground="white",
                             background=ap_theme.CRIMSON, borderwidth=0, focusthickness=0, padding=(12, 7))
        self.style.map("Cancel.TButton",
                       background=[("active", ap_theme.CRIMSON_DK), ("disabled", "#9AA3B5")])
        self.style.configure("Card.TFrame", background=ap_theme.CARD)
        self.style.configure("Progress.Horizontal.TProgressbar", thickness=10,
                             background=ap_theme.CRIMSON, troughcolor=ap_theme.BORDER, borderwidth=0)
        
        self._build_ui()
        self._load_data()

    def _build_ui(self):
        # Header thương hiệu + Footer mark (An Phước) — bọc quanh layout gốc
        ap_theme.Header(
            self.root, title="StoreVisit Pro",
            subtitle="Tự động hóa Báo cáo Công tác Kiểm tra Cụm Cửa hàng",
            app_tag="RETAIL COMMANDER",
        ).pack(fill="x")
        ap_theme.Footer(self.root, extra="StoreVisit Pro").pack(side="bottom", fill="x")

        # Main Layout container
        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Notebook for Tabs
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        tab_core = ttk.Frame(notebook, padding=10)
        tab_google = ttk.Frame(notebook, padding=10)
        tab_survey = ttk.Frame(notebook, padding=10)
        
        notebook.add(tab_core, text="📊 Báo cáo theo tuần (Core)")
        notebook.add(tab_google, text="📥 Google Sync (StoreVisit)")
        notebook.add(tab_survey, text="📋 Khảo sát thị trường (Survey)")
        
        # --- TAB 1: Core ---
        # Top Panel: Selection Box
        selection_frame = ttk.LabelFrame(tab_core, text=" LỰA CHỌN PHẠM VI BÁO CÁO ", padding=10)
        selection_frame.pack(fill=tk.X, pady=(0, 10))
        
        # ASM Dropdown
        asm_lbl_frame = ttk.Frame(selection_frame)
        asm_lbl_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(asm_lbl_frame, text="Chọn ASM / Quản Lý Kinh Doanh:", font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(0, 10))
        self.asm_var = tk.StringVar()
        self.asm_combo = ttk.Combobox(asm_lbl_frame, textvariable=self.asm_var, state="readonly", width=30)
        self.asm_combo.pack(side=tk.LEFT)
        self.asm_combo.bind("<<ComboboxSelected>>", self._on_asm_selected)
        
        # Stores checklist frame
        ttk.Label(selection_frame, text="Chọn danh sách cửa hàng công tác:", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(5, 5))
        
        # Scrollable Frame for Stores Checklist
        self.canvas_frame = ttk.Frame(selection_frame)
        self.canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        self.store_canvas = tk.Canvas(self.canvas_frame, height=90, bd=0, highlightthickness=0, background="#F8F9FA")
        scrollbar = ttk.Scrollbar(self.canvas_frame, orient="vertical", command=self.store_canvas.yview)
        
        self.store_checkbox_frame = ttk.Frame(self.store_canvas, style="Card.TFrame")
        self.store_checkbox_frame.bind(
            "<Configure>",
            lambda e: self.store_canvas.configure(scrollregion=self.store_canvas.boundingbox("all"))
        )
        
        self.store_canvas.create_window((0, 0), window=self.store_checkbox_frame, anchor="nw")
        self.store_canvas.configure(yscrollcommand=scrollbar.set)
        
        self.store_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.store_vars = {} # Maps store abbreviation to IntVar
        
        # --- PERIOD / DATE ANCHOR SELECTOR ---
        period_anchor_frame = ttk.LabelFrame(tab_core, text=" MỐC THỜI GIAN BÁO CÁO EXECUTIVE ", padding=10)
        period_anchor_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.exec_date_mode_var = tk.StringVar(value="auto")
        
        rb_auto = ttk.Radiobutton(period_anchor_frame, text="🟢 Tự động (Lấy dữ liệu mới nhất)", variable=self.exec_date_mode_var, value="auto")
        rb_auto.pack(side=tk.LEFT, padx=(0, 15))
        
        rb_prev = ttk.Radiobutton(period_anchor_frame, text="⏪ Kỳ trước (Tuần/Tháng trước)", variable=self.exec_date_mode_var, value="prev")
        rb_prev.pack(side=tk.LEFT, padx=(0, 15))

        rb_today = ttk.Radiobutton(period_anchor_frame, text="📅 Hôm nay (Thời gian thực)", variable=self.exec_date_mode_var, value="today")
        rb_today.pack(side=tk.LEFT)

        btn_frame_core = ttk.Frame(tab_core)
        btn_frame_core.pack(fill=tk.X, pady=(10, 0))
        self.btn_run = ttk.Button(btn_frame_core, text="TẠO BÁO CÁO CỤM (CLUSTER)", style="Accent.TButton", command=self._start_report_job)
        self.btn_run.pack(side=tk.LEFT, padx=(0, 10))
        
        self.btn_exec_weekly = ttk.Button(btn_frame_core, text="🚀 BÁO CÁO TUẦN EXECUTIVE", style="Accent.TButton", command=lambda: self._start_executive_job("weekly"))
        self.btn_exec_weekly.pack(side=tk.LEFT, padx=(0, 10))
        
        self.btn_exec_monthly = ttk.Button(btn_frame_core, text="📊 BÁO CÁO THÁNG EXECUTIVE", style="Accent.TButton", command=lambda: self._start_executive_job("monthly"))
        self.btn_exec_monthly.pack(side=tk.LEFT, padx=(0, 10))
        
        self.btn_email_core = ttk.Button(btn_frame_core, text="📧 GỬI MAIL BÁO CÁO", command=self._open_manual_email_modal)
        self.btn_email_core.pack(side=tk.LEFT)
        
        # --- TAB 2: Google Sync ---
        # Sync control
        sync_frame = ttk.LabelFrame(tab_google, text=" ĐỒNG BỘ DỮ LIỆU TỪ GOOGLE FORMS ", padding=10)
        sync_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.btn_sync = ttk.Button(sync_frame, text="Đồng bộ ngay", style="Accent.TButton", command=self._sync_google_forms)
        self.btn_sync.pack(side=tk.LEFT, padx=(0, 15))
        
        self.lbl_sync_status = ttk.Label(sync_frame, text="Lần sync gần nhất: Chưa đồng bộ", font=("Segoe UI", 9, "italic"))
        self.lbl_sync_status.pack(side=tk.LEFT, anchor="center")
        
        # Multi-Criteria Filter Bar
        filter_frame = ttk.LabelFrame(tab_google, text=" 🔍 BỘ LỌC ĐA TIÊU CHÍ (MULTI-CRITERIA FILTER) ", padding=8)
        filter_frame.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(filter_frame, text="ASM:").grid(row=0, column=0, padx=(0, 4), sticky="w")
        self.filter_asm_var = tk.StringVar(value="Tất cả ASM")
        self.combo_filter_asm = ttk.Combobox(filter_frame, textvariable=self.filter_asm_var, state="readonly", width=18)
        self.combo_filter_asm.grid(row=0, column=1, padx=(0, 15), sticky="w")
        self.combo_filter_asm.bind("<<ComboboxSelected>>", lambda e: self._refresh_treeview())

        ttk.Label(filter_frame, text="Cửa hàng:").grid(row=0, column=2, padx=(0, 4), sticky="w")
        self.filter_store_var = tk.StringVar(value="")
        self.entry_filter_store = ttk.Entry(filter_frame, textvariable=self.filter_store_var, width=14)
        self.entry_filter_store.grid(row=0, column=3, padx=(0, 15), sticky="w")
        self.entry_filter_store.bind("<KeyRelease>", lambda e: self._refresh_treeview())

        ttk.Label(filter_frame, text="Trạng thái:").grid(row=0, column=4, padx=(0, 4), sticky="w")
        self.filter_status_var = tk.StringVar(value="Tất cả trạng thái")
        self.combo_filter_status = ttk.Combobox(filter_frame, textvariable=self.filter_status_var, values=["Tất cả trạng thái", "Chưa xử lý (pending)", "Đã xử lý (done)", "Xử lý bị lỗi (error)"], state="readonly", width=16)
        self.combo_filter_status.grid(row=0, column=5, padx=(0, 15), sticky="w")
        self.combo_filter_status.bind("<<ComboboxSelected>>", lambda e: self._refresh_treeview())

        btn_reset_filter = ttk.Button(filter_frame, text="🔄 Đặt lại", command=self._reset_tab2_filters)
        btn_reset_filter.grid(row=0, column=6, padx=(5, 0), sticky="w")

        # Responses List
        list_frame = ttk.LabelFrame(tab_google, text=" DANH SÁCH BẢN GHI ĐANG CHỜ TẠO BÁO CÁO ", padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.tree = ttk.Treeview(list_frame, columns=("id", "store", "date", "asm", "cht", "status"), show="headings", selectmode="extended", height=4)
        self.tree.heading("id", text="Row ID")
        self.tree.heading("store", text="Mã CH")
        self.tree.heading("date", text="Ngày")
        self.tree.heading("asm", text="ASM/QLKD")
        self.tree.heading("cht", text="CHT")
        self.tree.heading("status", text="Trạng thái")
        
        self.tree.column("id", width=60, anchor="center")
        self.tree.column("store", width=80, anchor="center")
        self.tree.column("date", width=100, anchor="center")
        self.tree.column("asm", width=150, anchor="w")
        self.tree.column("cht", width=120, anchor="w")
        self.tree.column("status", width=90, anchor="center")
        
        tree_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Mode options
        opt_frame = ttk.LabelFrame(tab_google, text=" CHẾ ĐỘ TẠO BÁO CÁO ", padding=10)
        opt_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.google_mode_var = tk.StringVar(value="merge")
        ttk.Radiobutton(opt_frame, text="Gộp các cửa hàng đã chọn thành 1 file báo cáo cụm duy nhất", variable=self.google_mode_var, value="merge").pack(anchor="w", pady=2)
        ttk.Radiobutton(opt_frame, text="Báo cáo riêng lẻ từng cửa hàng (Tạo các tệp PPTX/PDF tách biệt)", variable=self.google_mode_var, value="separate").pack(anchor="w", pady=2)
        
        btn_frame_google = ttk.Frame(tab_google)
        btn_frame_google.pack(fill=tk.X)
        self.btn_run_google = ttk.Button(btn_frame_google, text="TẠO BÁO CÁO TỪ GOOGLE FORMS", style="Accent.TButton", command=self._start_google_report_job)
        self.btn_run_google.pack(side=tk.LEFT, padx=(0, 10))
        
        self.btn_email_google = ttk.Button(btn_frame_google, text="📧 GỬI MAIL BÁO CÁO", command=self._open_manual_email_modal)
        self.btn_email_google.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_diag_google = ttk.Button(btn_frame_google, text="🛠️ CHUẨN ĐOÁN LỖI", command=self._open_diagnostic_inspector)
        self.btn_diag_google.pack(side=tk.LEFT)

        # --- TAB 3: Market Survey ---
        # Sync control
        sync_frame_survey = ttk.LabelFrame(tab_survey, text=" ĐỒNG BỘ DỮ LIỆU KHẢO SÁT THỊ TRƯỜNG ", padding=10)
        sync_frame_survey.pack(fill=tk.X, pady=(0, 10))
        
        self.btn_sync_survey = ttk.Button(sync_frame_survey, text="Đồng bộ khảo sát", style="Accent.TButton", command=self._sync_market_surveys)
        self.btn_sync_survey.pack(side=tk.LEFT, padx=(0, 15))
        
        self.lbl_survey_sync_status = ttk.Label(sync_frame_survey, text="Lần sync gần nhất: Chưa đồng bộ", font=("Segoe UI", 9, "italic"))
        self.lbl_survey_sync_status.pack(side=tk.LEFT, anchor="center")
        
        # Responses List for Survey
        list_frame_survey = ttk.LabelFrame(tab_survey, text=" DANH SÁCH KHẢO SÁT ĐANG CHỜ TỔNG HỢP (QC: APPROVED) ", padding=10)
        list_frame_survey.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.tree_survey = ttk.Treeview(list_frame_survey, columns=("id", "store", "region", "date", "asm", "qc_status", "status"), show="headings", selectmode="extended", height=4)
        self.tree_survey.heading("id", text="Row ID")
        self.tree_survey.heading("store", text="Mã CH")
        self.tree_survey.heading("region", text="Khu vực/Cụm")
        self.tree_survey.heading("date", text="Ngày")
        self.tree_survey.heading("asm", text="ASM/QLKD")
        self.tree_survey.heading("qc_status", text="Duyệt QC")
        self.tree_survey.heading("status", text="Trạng thái")
        
        self.tree_survey.column("id", width=60, anchor="center")
        self.tree_survey.column("store", width=80, anchor="center")
        self.tree_survey.column("region", width=120, anchor="center")
        self.tree_survey.column("date", width=100, anchor="center")
        self.tree_survey.column("asm", width=150, anchor="w")
        self.tree_survey.column("qc_status", width=90, anchor="center")
        self.tree_survey.column("status", width=90, anchor="center")
        
        survey_scroll = ttk.Scrollbar(list_frame_survey, orient="vertical", command=self.tree_survey.yview)
        self.tree_survey.configure(yscrollcommand=survey_scroll.set)
        
        self.tree_survey.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        survey_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        btn_frame_survey = ttk.Frame(tab_survey)
        btn_frame_survey.pack(fill=tk.X)
        self.btn_run_survey = ttk.Button(btn_frame_survey, text="TỔNG HỢP BÁO CÁO KHẢO SÁT (EXCEL)", style="Accent.TButton", command=self._start_survey_consolidate_job)
        self.btn_run_survey.pack(side=tk.LEFT)

        # --- SHARED PANELS (Progress & Log & Diagnostic Station) ---
        # Middle Panel: Progress & Operations
        ops_frame = ttk.LabelFrame(main_frame, text=" TIẾN TRÌNH XỬ LÝ (JOB STATUS) ", padding=10)
        ops_frame.pack(fill=tk.X, pady=(0, 8))
        
        btn_frame = ttk.Frame(ops_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.btn_cancel = ttk.Button(btn_frame, text="🛑 HỦY BỎ JOB ĐANG CHẠY", style="Cancel.TButton", state=tk.DISABLED, command=self._trigger_cancel)
        self.btn_cancel.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_unlock = ttk.Button(btn_frame, text="🔓 MỞ KHÓA KHẨN CẤP (RESET)", style="Accent.TButton", command=self._emergency_unlock)
        self.btn_unlock.pack(side=tk.LEFT)
        
        self.progress_lbl = ttk.Label(ops_frame, text="Trạng thái: Sẵn sàng thực hiện", font=("Segoe UI", 10, "bold"))
        self.progress_lbl.pack(anchor="w", pady=(4, 2))
        
        self.progress_bar = ttk.Progressbar(ops_frame, style="Progress.Horizontal.TProgressbar", mode="determinate")
        self.progress_bar.pack(fill=tk.X)
        
        # Bottom Panel: Integrated Notebook (Live Console & Live Diagnostic Inspector)
        self.bottom_notebook = ttk.Notebook(main_frame)
        self.bottom_notebook.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

        # Tab A: Live Log Console
        tab_console = ttk.Frame(self.bottom_notebook, padding=6)
        self.bottom_notebook.add(tab_console, text=" 📜 NHẬT KÝ THỜI GIAN THỰC (LIVE LOG) ")

        console_tools = ttk.Frame(tab_console)
        console_tools.pack(fill=tk.X, pady=(0, 4))
        
        btn_copy_log = ttk.Button(console_tools, text="📋 Sao chép Log", command=self._copy_live_log)
        btn_copy_log.pack(side=tk.LEFT, padx=(0, 6))

        btn_clear_log = ttk.Button(console_tools, text="🧹 Xóa trắng Log", command=self._clear_live_log)
        btn_clear_log.pack(side=tk.LEFT)

        log_container = ttk.Frame(tab_console)
        log_container.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(log_container, height=6, font=("Consolas", 10), background="#0F1729", foreground="#E6EAF2", wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        log_scroll = ttk.Scrollbar(log_container, orient="vertical", command=self.log_text.yview)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=log_scroll.set)

        # Tab B: Live Diagnostic Inspector & Real-Time Error Pane
        tab_diag = ttk.Frame(self.bottom_notebook, padding=6)
        self.bottom_notebook.add(tab_diag, text=" 🛠️ BẢNG CHẨN ĐOÁN HỆ THỐNG & BẮT LỖI TỰ ĐỘNG ")

        diag_tools = ttk.Frame(tab_diag)
        diag_tools.pack(fill=tk.X, pady=(0, 4))

        self.lbl_diag_status_bar = ttk.Label(diag_tools, text="Trạng thái hệ thống: Đang khởi tạo...", font=("Segoe UI", 9, "bold"), foreground="#0A2342")
        self.lbl_diag_status_bar.pack(side=tk.LEFT, padx=(0, 10))

        btn_refresh_diag = ttk.Button(diag_tools, text="🔄 Quét lại Chẩn đoán", command=self._update_diagnostic_dashboard)
        btn_refresh_diag.pack(side=tk.LEFT, padx=(0, 6))

        btn_copy_diag = ttk.Button(diag_tools, text="📋 Sao chép Báo cáo", command=self._copy_diagnostic_report)
        btn_copy_diag.pack(side=tk.LEFT, padx=(0, 6))

        diag_container = ttk.Frame(tab_diag)
        diag_container.pack(fill=tk.BOTH, expand=True)

        self.diag_text = tk.Text(diag_container, height=6, font=("Consolas", 10), background="#0B132B", foreground="#00FF66", wrap=tk.WORD)
        self.diag_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        diag_scroll = ttk.Scrollbar(diag_container, orient="vertical", command=self.diag_text.yview)
        diag_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.diag_text.configure(yscrollcommand=diag_scroll.set)
        
        # Bind close window behavior
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _load_data(self):
        """Load unique ASMs from the DimStore Excel file."""
        if not os.path.exists(self.dim_store_path):
            messagebox.showerror("Lỗi Cấu Hình", f"Không tìm thấy tệp danh mục cửa hàng tại: {self.dim_store_path}")
            return
            
        try:
            df = pd.read_excel(self.dim_store_path)
            # Find unique ASMs
            asms = sorted(df["ASM"].dropna().unique())
            self.asm_combo["values"] = asms
            self.dim_store_df = df
            self._write_log("Đã tải thành công danh sách ASM từ DimStore_Final.xlsx")
            self._refresh_treeview()
            self._refresh_survey_treeview()
            self._update_diagnostic_dashboard()
        except Exception as e:
            messagebox.showerror("Lỗi Đọc Dữ Liệu", f"Không thể tải tệp DimStore_Final.xlsx: {str(e)}")

    def _on_asm_selected(self, event=None):
        """Update stores checkbox list based on selected ASM."""
        # Clear existing check buttons
        for widget in self.store_checkbox_frame.winfo_children():
            widget.destroy()
        self.store_vars.clear()
        
        selected_asm = self.asm_var.get()
        if not selected_asm:
            return
            
        # Get stores matching selected ASM
        df_filtered = self.dim_store_df[self.dim_store_df["ASM"] == selected_asm]
        
        # In field_mapping we have store abbreviation to standard StoreCode mapping.
        # Let's find abbreviation for displaying to user.
        # For simplicity, we search for matching StoreCode and display the store description.
        store_code_mapping = self.config.get("field_mapping", {}).get("store_code_mapping", {})
        # Build reverse mapping
        rev_mapping = {v: k for k, v in store_code_mapping.items()}
        
        # Grid placement helper variables
        col = 0
        row = 0
        max_cols = 3
        
        for _, r_store in df_filtered.iterrows():
            code = r_store["StoreCode"]
            name = r_store["StoreName"]
            
            # Find display abbreviation (abbreviation used in weekly_json and excel name prefix)
            abbrev = rev_mapping.get(code, code)
            
            var = tk.IntVar()
            self.store_vars[abbrev] = var
            
            chk = ttk.Checkbutton(
                self.store_checkbox_frame, 
                text=f"{abbrev} - {name}", 
                variable=var,
                padding=(10, 5)
            )
            chk.grid(row=row, column=col, sticky="w")
            
            col += 1
            if col >= max_cols:
                col = 0
                row += 1

    def _write_log(self, msg: str):
        self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_text.see(tk.END)

    def _start_report_job(self):
        """Collect params and launch app_worker subprocess."""
        asm = self.asm_var.get()
        if not asm:
            messagebox.showwarning("Thiếu thông tin", "Vui lòng chọn ASM trước.")
            return
            
        stores = [k for k, v in self.store_vars.items() if v.get() == 1]
        if not stores:
            messagebox.showwarning("Thiếu thông tin", "Vui lòng chọn ít nhất 1 cửa hàng để tạo báo cáo.")
            return
            
        self._start_job_thread(asm, stores, form_response_ids=None, no_merge=False)

    def _start_google_report_job(self):
        if not self.form_cache:
            return
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Thiếu thông tin", "Vui lòng chọn ít nhất 1 dòng phản hồi từ bảng trên.")
            return
            
        selected_responses = []
        for item in selected_items:
            vals = self.tree.item(item, "values")
            rid = vals[0]
            resp = self.form_cache.get_by_id(rid)
            if resp:
                selected_responses.append(resp)
                
        stores = [r.store_code for r in selected_responses]
        rids = [r.response_id for r in selected_responses]
        asm = selected_responses[0].asm_name or "ASM"
        
        no_merge = (self.google_mode_var.get() == "separate")
        self._start_job_thread(asm, stores, form_response_ids=rids, no_merge=no_merge)

    def _start_executive_job(self, period_type: str):
        """Launch Executive Combo Weekly or Monthly report job."""
        asm = self.asm_var.get() or "ALL"
        stores = [k for k, v in self.store_vars.items() if v.get() == 1]
        if not stores:
            stores = ["ALL"]
        ref_date = self.exec_date_mode_var.get() if hasattr(self, "exec_date_mode_var") else "auto"
        self._start_job_thread(asm, stores, form_response_ids=None, no_merge=False, schema="executive_combo", period_type=period_type, reference_date=ref_date)

    def _start_job_thread(self, asm: str, stores: list, form_response_ids: list = None, no_merge: bool = False, schema: str = "store_visit", period_type: str = "weekly", reference_date: str = "auto"):
        # Initialize job state
        self.current_job_id = str(uuid.uuid4())[:8]
        self.cancel_file_path = os.path.join(self.root_dir, f"temp/cancel_{self.current_job_id}.sig")
        os.makedirs(os.path.dirname(self.cancel_file_path), exist_ok=True)
        
        # Block inputs
        self._set_ui_blocked(True)
        self.progress_bar["value"] = 0
        self.log_text.delete("1.0", tk.END)
        
        mode_str = "Tách biệt" if no_merge else "Gộp cụm"
        if schema == "executive_combo":
            self._write_log(f"Bắt đầu tạo Báo cáo Executive Combo ({period_type.upper()}) | Mốc: {reference_date.upper()} | ASM: {asm}")
        elif form_response_ids:
            self._write_log(f"Bắt đầu tạo báo cáo từ Google Forms ({mode_str}) | Cửa hàng: {', '.join(stores)}")
        else:
            self._write_log(f"Bắt đầu tạo báo cáo cụm cho ASM: {asm} | Cửa hàng: {', '.join(stores)}")
            
        # Start worker thread
        self.is_running = True
        self.last_event_time = time.time()
        
        worker_thread = threading.Thread(target=self._run_subprocess, args=(asm, stores, form_response_ids, no_merge, schema, period_type, reference_date))
        worker_thread.daemon = True
        worker_thread.start()
        
        # Start Watchdog
        self.watchdog_thread = threading.Thread(target=self._run_watchdog)
        self.watchdog_thread.daemon = True
        self.watchdog_thread.start()

    def _set_ui_blocked(self, blocked: bool):
        state = tk.DISABLED if blocked else tk.NORMAL
        combo_state = "disabled" if blocked else "readonly"
        
        self.asm_combo["state"] = combo_state
        self.btn_run["state"] = state
        if hasattr(self, "btn_exec_weekly"): self.btn_exec_weekly["state"] = state
        if hasattr(self, "btn_exec_monthly"): self.btn_exec_monthly["state"] = state
        self.btn_cancel["state"] = tk.NORMAL if blocked else tk.DISABLED
        
        # Block checkboxes
        for child in self.store_checkbox_frame.winfo_children():
            try:
                child["state"] = state
            except Exception:
                pass
                
        # Block Google & Survey buttons
        try:
            self.btn_sync["state"] = state
            self.btn_run_google["state"] = state
            self.btn_sync_survey["state"] = state
            self.btn_run_survey["state"] = state
        except AttributeError:
            pass

    def _run_subprocess(self, asm: str, stores: list, form_response_ids: list = None, no_merge: bool = False, schema: str = "store_visit", period_type: str = "weekly", reference_date: str = "auto"):
        """Execute the worker subprocess with parameters."""
        worker_script = os.path.join(self.root_dir, "app_worker.py")
        venv_python = os.path.join(self.root_dir, ".venv/Scripts/python.exe")
        
        cmd = [
            venv_python, "-u", worker_script, # -u is for unbuffered output
            "--job-id", self.current_job_id,
            "--asm", asm,
            "--stores", ",".join(stores),
            "--cancel-file", self.cancel_file_path,
            "--schema", schema,
            "--period-type", period_type,
            "--reference-date", reference_date
        ]
        if form_response_ids:
            cmd.extend(["--form-response-ids", ",".join(form_response_ids)])
        if no_merge:
            cmd.append("--no-merge")
        
        try:
            # CREATE_NO_WINDOW hides CMD popup window on Windows
            self.worker_process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            # Read stdout line-by-line (async JSON stream)
            for line in iter(self.worker_process.stdout.readline, ""):
                if not self.is_running:
                    break
                line = line.strip()
                if line:
                    self._parse_ipc_line(line)
                    
            self.worker_process.wait()
            
            # Read stderr if process failed
            stderr_output = self.worker_process.stderr.read()
            if stderr_output.strip() and self.worker_process.returncode != 0:
                self.root.after(0, lambda: self._on_job_failed(f"Subprocess exit code: {self.worker_process.returncode}", stderr_output))
                
        except Exception as e:
            self.root.after(0, lambda: self._on_job_failed(str(e), traceback.format_exc()))
        finally:
            self.root.after(0, self._finalize_job)

    def _parse_ipc_line(self, line: str):
        """Parse structured JSON lines from worker stdout."""
        try:
            data = json.loads(line)
            # Update last event timestamp for watchdog
            self.last_event_time = time.time()
            
            msg_type = data.get("type")
            payload = data.get("payload", {})
            
            # Process events on main GUI thread
            self.root.after(0, lambda: self._handle_ipc_event(msg_type, payload))
        except Exception:
            # If line is not JSON, log as warning
            self.root.after(0, lambda: self._write_log(f"[Worker log]: {line}"))

    def _handle_ipc_event(self, msg_type: str, payload: dict):
        if msg_type == "job_started":
            self.progress_lbl["text"] = "Trạng thái: Đang khởi chạy job..."
            self._write_log("Tiến trình worker đã kết nối thành công.")
            self._update_diagnostic_dashboard()
            
        elif msg_type == "stage_started":
            stage = payload.get("stage", "")
            self.progress_lbl["text"] = f"Trạng thái: Bắt đầu giai đoạn '{stage}'"
            self._write_log(f"-> Giai đoạn mới: {stage}")
            
        elif msg_type == "progress":
            step = payload.get("current_step", 0)
            msg = payload.get("message", "")
            self.progress_bar["value"] = step
            self.progress_lbl["text"] = f"Tiến độ: {step}% — {msg}"
            self._write_log(f"   [{step}%] {msg}")
            
        elif msg_type == "warning":
            warn = payload.get("warning", "")
            self._write_log(f"⚠ Cảnh báo: {warn}")
            
        elif msg_type == "stage_completed":
            stage = payload.get("stage", "")
            self._write_log(f"✓ Hoàn thành giai đoạn: {stage}")
            
        elif msg_type == "job_completed":
            self.progress_bar["value"] = 100
            self.progress_lbl["text"] = "Trạng thái: Tạo báo cáo thành công!"
            self._write_log(f"🎉 JOB HOÀN THÀNH!")
            self._write_log(f"   PPTX: {payload.get('pptx')}")
            self._write_log(f"   PDF: {payload.get('pdf')}")
            
            pptx_path = payload.get("pptx", "")
            pdf_path = payload.get("pdf", "")
            docx_path = payload.get("docx", "")
            xlsx_path = payload.get("xlsx", "")
            store_code = payload.get("store_code", payload.get("store", "CH"))
            store_name = payload.get("store_name", store_code)
            asm_name = payload.get("asm", payload.get("asm_name", "ASM"))
            report_date = payload.get("report_date", datetime.now().strftime("%d/%m/%Y"))

            # Prompt for email sending with interactive confirmation (Topmost, No Grab Lock)
            self.root.after(100, lambda: self._prompt_and_send_email(
                store_code, store_name, asm_name, report_date, pptx_path, pdf_path, docx_path, xlsx_path
            ))

            # Open output folder
            if pptx_path:
                output_dir = os.path.dirname(pptx_path)
                try:
                    os.startfile(output_dir)
                except Exception:
                    pass
            self._finalize_job()
            
        elif msg_type == "job_cancelled":
            self.progress_lbl["text"] = "Trạng thái: Đã hủy bỏ job."
            self._write_log("Job đã được hủy bỏ theo yêu cầu người dùng.")
            self._finalize_job()
            
        elif msg_type == "job_failed":
            err = payload.get("error", "Lỗi chưa xác định")
            tb = payload.get("traceback", "")
            self._on_job_failed(err, tb)

    def _prompt_and_send_email(self, store_code: str, store_name: str, asm_name: str, report_date: str, pptx_path: str, pdf_path: str = "", docx_path: str = "", xlsx_path: str = ""):
        """Pop up an interactive email confirmation dialog allowing the user to view/edit recipient email before sending."""
        default_email = ""
        if hasattr(self, "loader") and self.loader:
            try:
                default_email = self.loader.get_asm_email(asm_name)
            except Exception:
                default_email = ""
                
        # Create Toplevel Window (NO grab_set to prevent Tkinter input freezing)
        win = tk.Toplevel(self.root)
        win.title("📧 XÁC NHẬN GỬI EMAIL BÁO CÁO CÔNG TÁC")
        win.geometry("580x340")
        win.resizable(False, False)
        win.transient(self.root)
        win.attributes("-topmost", True)
        win.lift()
        win.focus_force()

        # Center on parent window
        win.geometry("+%d+%d" % (self.root.winfo_x() + 50, self.root.winfo_y() + 50))

        # Title Label
        lbl_title = ttk.Label(win, text="📧 XÁC NHẬN GỬI EMAIL BÁO CÁO CÔNG TÁC", font=("Segoe UI", 12, "bold"), foreground="#0A2342")
        lbl_title.pack(anchor="w", padx=20, pady=(20, 5))

        lbl_info = ttk.Label(win, text=f"• Cửa hàng: {store_code} - {store_name}\n• ASM Phụ trách: {asm_name} | Ngày: {report_date}", font=("Segoe UI", 10))
        lbl_info.pack(anchor="w", padx=20, pady=(0, 15))

        # Input Frame
        frame_input = ttk.LabelFrame(win, text=" ĐỊA CHỈ EMAIL NGƯỜI NHẬN ", padding=10)
        frame_input.pack(fill=tk.X, padx=20, pady=(0, 15))

        lbl_input_hint = ttk.Label(frame_input, text="Vui lòng kiểm tra địa chỉ Email bên dưới (Có thể gõ sửa trực tiếp):", font=("Segoe UI", 9, "italic"))
        lbl_input_hint.pack(anchor="w", pady=(0, 5))

        email_var = tk.StringVar(value=default_email)
        entry_email = ttk.Entry(frame_input, textvariable=email_var, font=("Segoe UI", 10))
        entry_email.pack(fill=tk.X, pady=(0, 5))
        entry_email.focus_set()

        lbl_status = ttk.Label(win, text="", font=("Segoe UI", 9, "italic"), foreground="#007ACC")
        lbl_status.pack(anchor="w", padx=20, pady=(0, 10))

        def close_dialog():
            win.destroy()
            self.root.focus_set()

        def do_send():
            target_email = email_var.get().strip()
            if not target_email or "@" not in target_email:
                messagebox.showwarning("Email Chưa Hợp Lệ", "Vui lòng nhập địa chỉ Email hợp lệ (chứa ký tự '@').", parent=win)
                return

            btn_send["state"] = tk.DISABLED
            btn_skip["state"] = tk.DISABLED
            lbl_status["text"] = f"Đang gửi email báo cáo tới '{target_email}'..."

            def send_bg():
                try:
                    from app_worker import send_email_for_store
                    success = send_email_for_store(
                        self.loader, store_code, store_name, asm_name, report_date,
                        pdf_path, docx_path, pptx_path, xlsx_path, target_email=target_email
                    )
                    if success:
                        self.root.after(0, lambda: self._write_log(f"✓ ĐÃ GỬI MAIL BÁO CÁO THÀNH CÔNG TỚI: {target_email}"))
                        self.root.after(0, lambda: messagebox.showinfo("Thành Công", f"Đã gửi email báo cáo thành công tới:\n{target_email}", parent=self.root))
                    else:
                        self.root.after(0, lambda: self._write_log(f"❌ GỬI MAIL THẤT BẠI TỚI: {target_email}"))
                        self.root.after(0, lambda: messagebox.showerror("Gửi Mail Thất Bại", f"Không thể gửi email tới {target_email}. Vui lòng kiểm tra lại cấu hình WebApp URL.", parent=self.root))
                except Exception as ex:
                    self.root.after(0, lambda: self._write_log(f"❌ LỖI GỬI MAIL: {ex}"))
                    self.root.after(0, lambda: messagebox.showerror("Lỗi Gửi Mail", f"Lỗi phát sinh khi gửi mail: {ex}", parent=self.root))
                finally:
                    self.root.after(0, close_dialog)

            t = threading.Thread(target=send_bg)
            t.daemon = True
            t.start()

        # Button Frame
        frame_btns = ttk.Frame(win)
        frame_btns.pack(fill=tk.X, padx=20, pady=(0, 15))

        btn_send = ttk.Button(frame_btns, text="📧 GỬI EMAIL NGAY", style="Accent.TButton", command=do_send)
        btn_send.pack(side=tk.LEFT, padx=(0, 10))

        btn_skip = ttk.Button(frame_btns, text="❌ BỎ QUA / KHÔNG GỬI", command=close_dialog)
        btn_skip.pack(side=tk.LEFT)

    def _open_manual_email_modal(self):
        """Open email prompt modal manually for selected store/response."""
        selected_items = self.tree.selection() if hasattr(self, "tree") else []
        store_code = ""
        store_name = ""
        asm_name = ""
        report_date = ""

        if selected_items:
            item = selected_items[0]
            vals = self.tree.item(item, "values")
            if vals:
                store_code = str(vals[1]).strip()
                report_date = str(vals[2]).strip()
                asm_name = str(vals[3]).strip()
                store_name = store_code
        else:
            checked_stores = [code for code, var in self.store_vars.items() if var.get() == 1]
            if checked_stores:
                store_code = checked_stores[0]
                store_name = store_code
            else:
                messagebox.showwarning("Chọn Cửa Hàng", "Vui lòng chọn 1 bản ghi trong danh sách (Tab Google Forms) hoặc tích chọn 1 cửa hàng (Tab Báo cáo Cụm) để gửi mail.")
                return

        if hasattr(self, "loader") and self.loader:
            try:
                df_store = self.loader.load_dim_store()
                match = df_store[df_store["StoreCode"] == store_code]
                if not match.empty:
                    store_name = match.iloc[0]["StoreName"]
                    if not asm_name:
                        asm_name = str(match.iloc[0]["ASM"])
            except Exception:
                pass

        if not store_name:
            store_name = store_code

        output_dir = self.loader.get_path("output_dir") if hasattr(self, "loader") and self.loader else "output"
        pptx_path = ""
        pdf_path = ""
        docx_path = ""
        xlsx_path = ""

        from utils.filename_formatter import remove_accents
        if os.path.exists(output_dir):
            for root_dir, dirs, files in os.walk(output_dir):
                for f in files:
                    clean_f = f.lower()
                    if store_code.lower() in clean_f or (store_name and remove_accents(store_name).lower() in clean_f):
                        full_p = os.path.join(root_dir, f)
                        if clean_f.endswith(".pptx"): pptx_path = full_p
                        elif clean_f.endswith(".pdf"): pdf_path = full_p
                        elif clean_f.endswith(".docx"): docx_path = full_p
                        elif clean_f.endswith(".xlsx"): xlsx_path = full_p

        self._prompt_and_send_email(
            store_code=store_code,
            store_name=store_name,
            asm_name=asm_name,
            report_date=report_date or datetime.now().strftime("%d/%m/%Y"),
            pptx_path=pptx_path,
            pdf_path=pdf_path,
            docx_path=docx_path,
            xlsx_path=xlsx_path
        )

    def _on_job_failed(self, error: str, tb: str):
        self.progress_lbl["text"] = "Trạng thái: Gặp lỗi trong quá trình tạo báo cáo."
        self._write_log(f"❌ THẤT BẠI: {error}")
        if tb:
            self._write_log(f"Traceback chi tiết:\n{tb}")
        
        self.latest_error_text = f"THỜI GIAN GẶP LỖI: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\nLỖI CHI TIẾT: {error}\n\nTRACEBACK ĐẦY ĐỦ:\n{tb}"
        self._update_diagnostic_dashboard()
        
        # Automatically switch to the Diagnostic Tab so user immediately sees the root cause!
        try:
            if hasattr(self, "bottom_notebook"):
                self.bottom_notebook.select(1)
        except Exception:
            pass
            
        self._finalize_job()
        messagebox.showerror("Thất Bại", f"Quá trình tạo báo cáo thất bại:\n{error}\n\nĐã ghi nhận toàn bộ chẩn đoán tại tab '🛠️ BẢNG CHẨN ĐOÁN HỆ THỐNG' bên dưới.")

    def _trigger_cancel(self):
        """Write cancel signal file."""
        if self.cancel_file_path and self.is_running:
            self.btn_cancel["state"] = tk.DISABLED
            self.progress_lbl["text"] = "Đang gửi tín hiệu hủy bỏ job..."
            self._write_log("Đang yêu cầu hủy bỏ job...")
            try:
                with open(self.cancel_file_path, "w") as f:
                    f.write("CANCEL")
            except Exception as e:
                self._write_log(f"Lỗi gửi tín hiệu hủy: {e}")
                self._force_kill_subprocess()

    def _emergency_unlock(self):
        """Emergency reset to unblock UI, force release locks, and kill orphan workers."""
        self._write_log("⚠️ [MỞ KHÓA KHẨN CẤP] Đang tiến hành giải phóng toàn bộ khóa và tiến trình con...")
        self.is_running = False
        self._force_kill_subprocess()
        try:
            from job_lock import JobLock
            JobLock().release(force=True)
        except Exception:
            pass
        
        # Clean any lock & sig files
        try:
            import glob
            for f in glob.glob(os.path.join(self.root_dir, "temp", "*.sig")):
                try: os.remove(f)
                except Exception: pass
            for f in glob.glob(os.path.join(self.root_dir, "temp", "*.lock")):
                try: os.remove(f)
                except Exception: pass
        except Exception:
            pass

        self._set_ui_blocked(False)
        self.progress_lbl["text"] = "Trạng thái: Đã mở khóa thành công! Hệ thống sẵn sàng."
        self.progress_bar["value"] = 0
        self._write_log("✓ ĐÃ MỞ KHÓA HỆ THỐNG THÀNH CÔNG! Giao diện đã sẵn sàng nhận lệnh mới.")
        self._update_diagnostic_dashboard()
        messagebox.showinfo("Đã Mở Khóa Khẩn Cấp", "✓ Đã mở khóa và khôi phục giao diện thành công!\nToàn bộ tiến trình con và file khóa đã được dọn sạch.")

    def _copy_live_log(self):
        """Copy entire live console log text to clipboard."""
        try:
            content = self.log_text.get("1.0", tk.END).strip()
            if content:
                self.root.clipboard_clear()
                self.root.clipboard_append(content)
                messagebox.showinfo("Sao Chép", "Đã sao chép toàn bộ Nhật ký Live Log vào Clipboard!")
            else:
                messagebox.showinfo("Sao Chép", "Nhật ký hiện đang trống.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể sao chép: {e}")

    def _clear_live_log(self):
        """Clear live console log."""
        self.log_text.delete("1.0", tk.END)
        self._write_log("Đã xóa trắng nhật ký.")

    def _copy_diagnostic_report(self):
        """Copy diagnostic dashboard content to clipboard."""
        try:
            content = self.diag_text.get("1.0", tk.END).strip()
            if content:
                self.root.clipboard_clear()
                self.root.clipboard_append(content)
                messagebox.showinfo("Sao Chép", "Đã sao chép Báo cáo Chẩn đoán Hệ thống vào Clipboard!")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể sao chép: {e}")

    def _update_diagnostic_dashboard(self):
        """Perform real-time system health check and refresh the diagnostic inspector tab."""
        if not hasattr(self, "diag_text"):
            return

        diag_lines = []
        diag_lines.append("=========================================================================")
        diag_lines.append(f"  STOREVISIT PRO REALTIME SYSTEM DIAGNOSTICS - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        diag_lines.append("=========================================================================\n")

        # 0. LATEST ERROR FINGERPRINT (IF ANY)
        if self.latest_error_text:
            diag_lines.append("🚨 [LỖI PHÁT SINH GẦN NHẤT / LATEST ERROR FINGERPRINT]:")
            diag_lines.append("-------------------------------------------------------------------------")
            diag_lines.append(self.latest_error_text)
            diag_lines.append("-------------------------------------------------------------------------\n")
            status_text = "Hệ thống ghi nhận LỖI! Vui lòng kiểm tra báo cáo bên dưới."
            status_color = "#C00000"
        else:
            diag_lines.append("✓ [TRẠNG THÁI HOẠT ĐỘNG]: Không ghi nhận lỗi runtime tồn đọng.\n")
            status_text = "Hệ thống: SẴN SÀNG | Mọi thành phần hoạt động ổn định"
            status_color = "#008000"

        # 1. Environment & Dependencies
        diag_lines.append("[1] MÔI TRƯỜNG PYTHON & THƯ VIỆN BẮT BUỘC:")
        diag_lines.append(f"  • Operating System: Windows ({sys.platform})")
        diag_lines.append(f"  • Python Interpreter: {sys.executable}")
        diag_lines.append(f"  • Project Directory: {self.root_dir}")

        for pkg in ["pandas", "pptx", "openpyxl", "requests", "yaml"]:
            try:
                mod = __import__(pkg)
                ver = getattr(mod, "__version__", "OK")
                diag_lines.append(f"  ✓ Thư viện '{pkg}': Sẵn sàng (v{ver})")
            except Exception as ex:
                diag_lines.append(f"  ❌ Thư viện '{pkg}': THIẾU ({ex})")

        # 2. Configuration & Credentials
        diag_lines.append("\n[2] TỆP CẤU HÌNH & GOOGLE CREDENTIALS:")
        cfg_path = os.path.join(self.root_dir, "config", "app_config.yaml")
        diag_lines.append(f"  • app_config.yaml: {'✓ Đầy đủ' if os.path.exists(cfg_path) else '❌ THIẾU'}")

        cred_path = self.config.get("google", {}).get("credentials_path", "")
        if not os.path.isabs(cred_path): cred_path = os.path.join(self.root_dir, cred_path)
        diag_lines.append(f"  • google_credentials.json: {'✓ Hợp lệ' if os.path.exists(cred_path) else '❌ THIẾU'}")

        w_url = self.config.get("google", {}).get("webapp_url", "")
        diag_lines.append(f"  • Google Apps Script WebApp URL: {'✓ Đã cấu hình' if w_url else '⚠️ CHƯA CẤU HÌNH'}")

        # 3. Data Cache
        diag_lines.append("\n[3] TRẠNG THÁI DỮ LIỆU & CACHE:")
        fc_path = os.path.join(self.root_dir, "data", "form_cache.json")
        if os.path.exists(fc_path):
            try:
                with open(fc_path, "r", encoding="utf-8") as f:
                    fc_data = json.load(f)
                diag_lines.append(f"  ✓ form_cache.json: {len(fc_data)} phản hồi Google Forms")
            except Exception as e:
                diag_lines.append(f"  ❌ form_cache.json: Lỗi đọc file ({e})")
        else:
            diag_lines.append("  ⚠️ form_cache.json: Chưa có tệp cache")

        # 4. Lock & Subprocess State
        diag_lines.append("\n[4] TIẾN TRÌNH WORKER & KHÓA NỀN (LOCK STATE):")
        lock_p = os.path.join(self.root_dir, "temp", "app.lock")
        if os.path.exists(lock_p):
            try:
                with open(lock_p, "r") as f:
                    pid = f.read().strip()
                diag_lines.append(f"  ⚠ File khóa temp/app.lock đang tồn tại (PID: {pid})")
            except Exception:
                diag_lines.append("  ⚠ File khóa temp/app.lock đang tồn tại")
        else:
            diag_lines.append("  ✓ File khóa temp/app.lock: SẠCH (Không bị chiếm dụng)")

        worker_st = "ĐANG CHẠY (RUNNING)" if self.is_running else "RẢNH (IDLE)"
        diag_lines.append(f"  • Trạng thái Worker: {worker_st}")

        diag_content = "\n".join(diag_lines)
        
        self.diag_text.delete("1.0", tk.END)
        self.diag_text.insert(tk.END, diag_content)

        if hasattr(self, "lbl_diag_status_bar"):
            self.lbl_diag_status_bar.config(text=f"Trạng thái: {status_text} | Worker: {worker_st}", foreground=status_color)

    def _run_watchdog(self):
        """Monitor event timeouts in background thread."""
        watchdog_timeout = 300 
        
        while self.is_running:
            time.sleep(2)
            if not self.is_running:
                break
                
            elapsed = time.time() - self.last_event_time
            if elapsed > watchdog_timeout:
                self.root.after(0, self._on_watchdog_timeout)
                break

    def _on_watchdog_timeout(self):
        self._write_log(f"❌ THỜI GIAN PHẢN HỒI QUÁ HẠN (Watchdog Timeout): Không nhận được sự kiện mới trong {300} giây.")
        messagebox.showerror("Watchdog Timeout", "Tiến trình con bị treo hoặc hết thời gian phản hồi. Hệ thống sẽ buộc dừng tiến trình con để bảo vệ tài nguyên.")
        self._force_kill_subprocess()

    def _force_kill_subprocess(self):
        """Forcibly kill the worker process and clean up."""
        self._write_log("Đang tiến hành buộc dừng tiến trình worker con...")
        if self.worker_process:
            try:
                # Force kill PID on Windows
                subprocess.run(f"taskkill /PID {self.worker_process.pid} /F /T", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except Exception:
                pass
            self.worker_process = None
            
        # Clean lock
        try:
            from job_lock import JobLock
            JobLock().release(force=True)
        except Exception:
            pass
            
        self._finalize_job()

    def _finalize_job(self):
        self.is_running = False
        self.worker_process = None
        self.current_job_id = None
        self.cancel_file_path = None
        try:
            from job_lock import JobLock
            JobLock().release(force=True)
        except Exception:
            pass
        self._set_ui_blocked(False)
        self._refresh_treeview()
        self._update_diagnostic_dashboard()

    def _sync_google_forms(self):
        if not self.form_cache:
            messagebox.showerror("Lỗi cấu hình", "Chưa thể khởi tạo Google Cache. Vui lòng kiểm tra config/app_config.yaml.")
            return
            
        google_config = self.config.get("google", {})
        if not google_config.get("spreadsheet_id") or not os.path.exists(google_config.get("credentials_path", "")):
            messagebox.showwarning("Thiếu cấu hình", "Vui lòng cấu hình 'spreadsheet_id' và kiểm tra tệp 'google_credentials.json' trước khi đồng bộ.")
            return
            
        self.btn_sync["state"] = tk.DISABLED
        self.lbl_sync_status["text"] = "Đang kết nối Google API..."
        self._write_log("Bắt đầu đồng bộ từ Google Forms...")
        
        def run_sync():
            try:
                def cb(msg, progress):
                    self.root.after(0, lambda: self.lbl_sync_status.config(text=msg))
                    self.root.after(0, lambda: self._write_log(msg))
                    
                new_count = self.form_cache.sync_from_google(progress_callback=cb)
                self.root.after(0, lambda: self._on_sync_success(new_count))
            except Exception as e:
                self.root.after(0, lambda: self._on_sync_failed(str(e)))
                
        t = threading.Thread(target=run_sync)
        t.daemon = True
        t.start()

    def _on_sync_success(self, new_count):
        self.btn_sync["state"] = tk.NORMAL
        now_str = datetime.now().strftime("%H:%M:%S")
        self.lbl_sync_status["text"] = f"Lần sync gần nhất: {now_str} (Thêm mới: {new_count})"
        self._write_log(f"✓ Đồng bộ thành công! Lấy được {new_count} phản hồi mới.")
        self._refresh_treeview()
        self._update_diagnostic_dashboard()
        
    def _on_sync_failed(self, err_msg):
        self.btn_sync["state"] = tk.NORMAL
        self.lbl_sync_status["text"] = "Đồng bộ thất bại"
        self._write_log(f"❌ Đồng bộ thất bại: {err_msg}")
        self.latest_error_text = f"LỖI ĐỒNG BỘ GOOGLE SHEETS:\n{err_msg}"
        self._update_diagnostic_dashboard()
        messagebox.showerror("Lỗi Đồng Bộ", f"Không thể đồng bộ dữ liệu từ Google Sheets:\n{err_msg}")

    def _reset_tab2_filters(self):
        if hasattr(self, "filter_asm_var"): self.filter_asm_var.set("Tất cả ASM")
        if hasattr(self, "filter_store_var"): self.filter_store_var.set("")
        if hasattr(self, "filter_status_var"): self.filter_status_var.set("Tất cả trạng thái")
        self._refresh_treeview()

    def _refresh_treeview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        if not self.form_cache:
            return
            
        self.form_cache.load()
        responses = self.form_cache.get_all()

        # Update ASM filter dropdown values dynamically
        if hasattr(self, "combo_filter_asm"):
            asm_list = sorted(list({r.asm_name for r in responses if r.asm_name}))
            curr_vals = list(self.combo_filter_asm["values"])
            new_vals = ["Tất cả ASM"] + asm_list
            if curr_vals != new_vals:
                self.combo_filter_asm["values"] = new_vals

        sel_asm = self.filter_asm_var.get() if hasattr(self, "filter_asm_var") else "Tất cả ASM"
        search_store = self.filter_store_var.get().strip().upper() if hasattr(self, "filter_store_var") else ""
        sel_status = self.filter_status_var.get() if hasattr(self, "filter_status_var") else "Tất cả trạng thái"

        for r in responses:
            # ASM Filter
            if sel_asm != "Tất cả ASM" and r.asm_name != sel_asm:
                continue

            # Store Search Filter
            if search_store and (search_store not in r.store_code.upper() and search_store not in r.store_name.upper()):
                continue

            # Status Filter
            if sel_status == "Chưa xử lý (pending)" and r.status != "pending":
                continue
            elif sel_status == "Đã xử lý (done)" and r.status != "done":
                continue
            elif sel_status == "Xử lý bị lỗi (error)" and r.status != "error":
                continue

            self.tree.insert("", tk.END, values=(
                r.response_id,
                r.store_code,
                r.report_date,
                r.asm_name,
                r.cht_name,
                r.status
            ))

    def _open_diagnostic_inspector(self):
        """Switch to the integrated Diagnostic Inspector Tab."""
        if hasattr(self, "bottom_notebook"):
            self.bottom_notebook.select(1)
        self._update_diagnostic_dashboard()

    def _refresh_survey_treeview(self):
        for item in self.tree_survey.get_children():
            self.tree_survey.delete(item)
            
        if not self.survey_cache:
            return
            
        self.survey_cache.load()
        responses = self.survey_cache.get_all()
        for r in responses:
            self.tree_survey.insert("", tk.END, values=(
                r.response_id,
                r.store_code,
                r.region,
                r.survey_date,
                r.qlkd_asm,
                r.qc_status,
                r.status
            ))

    def _sync_market_surveys(self):
        if not self.survey_cache:
            messagebox.showerror("Lỗi cấu hình", "Chưa thể khởi tạo Survey Cache. Vui lòng kiểm tra config/app_config.yaml.")
            return
            
        google_config = self.config.get("google", {})
        if not google_config.get("spreadsheet_id") or not os.path.exists(google_config.get("credentials_path", "")):
            messagebox.showwarning("Thiếu cấu hình", "Vui lòng cấu hình 'spreadsheet_id' và kiểm tra tệp 'google_credentials.json' trước khi đồng bộ.")
            return
            
        self.btn_sync_survey["state"] = tk.DISABLED
        self.lbl_survey_sync_status["text"] = "Đang kết nối Google API..."
        self._write_log("Bắt đầu đồng bộ khảo sát từ Google Sheets...")
        
        def run_sync():
            try:
                def cb(msg, progress):
                    self.root.after(0, lambda: self.lbl_survey_sync_status.config(text=msg))
                    self.root.after(0, lambda: self._write_log(msg))
                    
                survey_sheet_name = google_config.get("survey_sheet_name", "MarketSurvey_Responses")
                new_count = self.survey_cache.sync_from_google(sheet_name=survey_sheet_name, progress_callback=cb)
                self.root.after(0, lambda: self._on_survey_sync_success(new_count))
            except Exception as e:
                self.root.after(0, lambda: self._on_survey_sync_failed(str(e)))
                
        t = threading.Thread(target=run_sync)
        t.daemon = True
        t.start()

    def _on_survey_sync_success(self, new_count):
        self.btn_sync_survey["state"] = tk.NORMAL
        now_str = datetime.now().strftime("%H:%M:%S")
        self.lbl_survey_sync_status["text"] = f"Lần sync gần nhất: {now_str} (Thêm mới: {new_count})"
        self._write_log(f"✓ Đồng bộ khảo sát thành công! Lấy được {new_count} phản hồi mới.")
        self._refresh_survey_treeview()
        
    def _on_survey_sync_failed(self, err_msg):
        self.btn_sync_survey["state"] = tk.NORMAL
        self.lbl_survey_sync_status["text"] = "Đồng bộ thất bại"
        self._write_log(f"❌ Đồng bộ khảo sát thất bại: {err_msg}")
        messagebox.showerror("Lỗi Đồng Bộ", f"Không thể đồng bộ khảo sát từ Google Sheets:\n{err_msg}")

    def _start_survey_consolidate_job(self):
        if not self.survey_cache:
            return
        selected_items = self.tree_survey.selection()
        if not selected_items:
            messagebox.showwarning("Thiếu thông tin", "Vui lòng chọn ít nhất 1 dòng khảo sát từ bảng trên.")
            return
            
        selected_responses = []
        for item in selected_items:
            vals = self.tree_survey.item(item, "values")
            rid = vals[0]
            resp = self.survey_cache.get_by_id(rid)
            if resp:
                if resp.qc_status.lower() != "approved":
                    ans = messagebox.askyesno(
                        "Cảnh báo QC", 
                        f"Phản hồi hàng {resp.response_id} ({resp.store_code}) chưa được QLKD duyệt (QC_Status = '{resp.qc_status}').\nBạn có chắc chắn muốn đưa phản hồi này vào báo cáo tổng hợp không?"
                    )
                    if not ans:
                        return
                selected_responses.append(resp)
                
        stores = [r.store_code for r in selected_responses]
        rids = [r.response_id for r in selected_responses]
        asm = selected_responses[0].qlkd_asm or "ASM"
        
        self.current_job_id = str(uuid.uuid4())[:8]
        self.cancel_file_path = os.path.join(self.root_dir, f"temp/cancel_{self.current_job_id}.sig")
        os.makedirs(os.path.dirname(self.cancel_file_path), exist_ok=True)
        
        self._set_ui_blocked(True)
        self.progress_bar["value"] = 0
        self.log_text.delete("1.0", tk.END)
        self._write_log(f"Bắt đầu tổng hợp báo cáo khảo sát thị trường | Số lượng: {len(rids)} phản hồi")
        
        self.is_running = True
        self.last_event_time = time.time()
        
        worker_thread = threading.Thread(target=self._run_survey_subprocess, args=(asm, stores, rids))
        worker_thread.daemon = True
        worker_thread.start()
        
        self.watchdog_thread = threading.Thread(target=self._run_watchdog)
        self.watchdog_thread.daemon = True
        self.watchdog_thread.start()

    def _run_survey_subprocess(self, asm: str, stores: list, form_response_ids: list):
        worker_script = os.path.join(self.root_dir, "app_worker.py")
        venv_python = os.path.join(self.root_dir, ".venv/Scripts/python.exe")
        
        cmd = [
            venv_python, "-u", worker_script,
            "--job-id", self.current_job_id,
            "--asm", asm,
            "--stores", ",".join(stores),
            "--cancel-file", self.cancel_file_path,
            "--form-response-ids", ",".join(form_response_ids),
            "--schema", "market_survey"
        ]
        
        try:
            self.worker_process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            for line in iter(self.worker_process.stdout.readline, ""):
                if not self.is_running:
                    break
                line = line.strip()
                if line:
                    self._parse_ipc_line(line)
                    
            self.worker_process.wait()
            
            stderr_output = self.worker_process.stderr.read()
            if stderr_output.strip() and self.worker_process.returncode != 0:
                self.root.after(0, lambda: self._on_job_failed(f"Subprocess exit code: {self.worker_process.returncode}", stderr_output))
                
        except Exception as e:
            self.root.after(0, lambda: self._on_job_failed(str(e), traceback.format_exc()))
        finally:
            self.root.after(0, self._finalize_job)

    def _on_close(self):
        """Confirm if user wants to close when worker is active."""
        if self.is_running:
            if messagebox.askyesno("Đóng ứng dụng", "Tiến trình đang chạy. Bạn có muốn hủy bỏ job và đóng ứng dụng ngay không?"):
                self._trigger_cancel()
                self._force_kill_subprocess()
                self.root.destroy()
        else:
            self.root.destroy()

# Helper
def datetime_now_str() -> str:
    return datetime.now().strftime("%d/%m/%Y")

if __name__ == "__main__":
    root = tk.Tk()
    app = StoreVisitApp(root)
    root.mainloop()
