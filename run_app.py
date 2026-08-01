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

class StoreVisitApp:
    def __init__(self, root):
        self.root = root
        self.root.title("StoreVisit - Tự Động Hóa Báo Cáo Kiểm Tra Cụm Cửa Hàng")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        
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
            
        # Styles & Themes
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        # Palette: Navy, Soft gray, Teal
        self.style.configure(".", font=("Segoe UI", 10))
        self.style.configure("TLabel", foreground="#2C3E50")
        self.style.configure("Header.TLabel", font=("Segoe UI", 15, "bold"), foreground="#0A2342")
        self.style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), foreground="white", background="#0A2342")
        self.style.map("Accent.TButton", background=[("active", "#1A3D6C"), ("disabled", "#BDC3C7")])
        self.style.configure("Cancel.TButton", font=("Segoe UI", 10, "bold"), foreground="white", background="#C0392B")
        self.style.map("Cancel.TButton", background=[("active", "#E74C3C"), ("disabled", "#BDC3C7")])
        
        self.style.configure("Card.TFrame", background="#F8F9FA", relief="groove", borderwidth=1)
        self.style.configure("Progress.Horizontal.TProgressbar", thickness=15)
        
        self._build_ui()
        self._load_data()

    def _build_ui(self):
        # Main Layout container
        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        header = ttk.Label(main_frame, text="HỆ THỐNG TỰ ĐỘNG HÓA BÁO CÁO CÔNG TÁC CỬA HÀNG (STOREVISIT)", style="Header.TLabel")
        header.pack(pady=(0, 10), anchor="w")
        
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
        
        self.btn_exec_weekly = ttk.Button(btn_frame_core, text="🚀 BÁO CÁO TUẦN EXECUTIVE (PPTX + EXCEL)", style="Accent.TButton", command=lambda: self._start_executive_job("weekly"))
        self.btn_exec_weekly.pack(side=tk.LEFT, padx=(0, 10))
        
        self.btn_exec_monthly = ttk.Button(btn_frame_core, text="📊 BÁO CÁO THÁNG EXECUTIVE (PPTX + EXCEL)", style="Accent.TButton", command=lambda: self._start_executive_job("monthly"))
        self.btn_exec_monthly.pack(side=tk.LEFT)
        
        # --- TAB 2: Google Sync ---
        # Sync control
        sync_frame = ttk.LabelFrame(tab_google, text=" ĐỒNG BỘ DỮ LIỆU TỪ GOOGLE FORMS ", padding=10)
        sync_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.btn_sync = ttk.Button(sync_frame, text="Đồng bộ ngay", style="Accent.TButton", command=self._sync_google_forms)
        self.btn_sync.pack(side=tk.LEFT, padx=(0, 15))
        
        self.lbl_sync_status = ttk.Label(sync_frame, text="Lần sync gần nhất: Chưa đồng bộ", font=("Segoe UI", 9, "italic"))
        self.lbl_sync_status.pack(side=tk.LEFT, anchor="center")
        
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
        self.btn_run_google.pack(side=tk.LEFT)

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

        # --- SHARED PANELS (Progress & Log) ---
        # Middle Panel: Progress & Operations
        ops_frame = ttk.LabelFrame(main_frame, text=" TIẾN TRÌNH XỬ LÝ (JOB STATUS) ", padding=10)
        ops_frame.pack(fill=tk.X, pady=(0, 10))
        
        btn_frame = ttk.Frame(ops_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.btn_cancel = ttk.Button(btn_frame, text="HỦY BỎ JOB ĐANG CHẠY", style="Cancel.TButton", state=tk.DISABLED, command=self._trigger_cancel)
        self.btn_cancel.pack(side=tk.LEFT)
        
        self.progress_lbl = ttk.Label(ops_frame, text="Trạng thái: Sẵn sàng thực hiện")
        self.progress_lbl.pack(anchor="w", pady=(2, 2))
        
        self.progress_bar = ttk.Progressbar(ops_frame, style="Progress.Horizontal.TProgressbar", mode="determinate")
        self.progress_bar.pack(fill=tk.X)
        
        # Bottom Panel: Console Log output
        console_frame = ttk.LabelFrame(main_frame, text=" NHẬT KÝ CHI TIẾT (LIVE LOG) ", padding=10)
        console_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = tk.Text(console_frame, height=5, font=("Consolas", 9), background="#1E1E1E", foreground="#F8F9FA", wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        log_scroll = ttk.Scrollbar(console_frame, orient="vertical", command=self.log_text.yview)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        
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
            
            # Open output folder
            output_dir = os.path.dirname(payload.get("pptx"))
            messagebox.showinfo("Thành Công", f"Đã xuất báo cáo cụm thành công tại:\n{output_dir}")
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

    def _on_job_failed(self, error: str, tb: str):
        self.progress_lbl["text"] = "Trạng thái: Gặp lỗi trong quá trình tạo báo cáo."
        self._write_log(f"❌ THẤT BẠI: {error}")
        if tb:
            self._write_log(f"Traceback chi tiết:\n{tb}")
        messagebox.showerror("Thất Bại", f"Quá trình tạo báo cáo thất bại:\n{error}")
        self._finalize_job()

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

    def _run_watchdog(self):
        """Monitor event timeouts in background thread."""
        # Configurable watchdog timeout per stage, default to 300s
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
            JobLock().release()
        except Exception:
            pass
            
        self._finalize_job()

    def _finalize_job(self):
        self.is_running = False
        self.worker_process = None
        self.current_job_id = None
        self.cancel_file_path = None
        self._set_ui_blocked(False)
        self._refresh_treeview()

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
        
    def _on_sync_failed(self, err_msg):
        self.btn_sync["state"] = tk.NORMAL
        self.lbl_sync_status["text"] = "Đồng bộ thất bại"
        self._write_log(f"❌ Đồng bộ thất bại: {err_msg}")
        messagebox.showerror("Lỗi Đồng Bộ", f"Không thể đồng bộ dữ liệu từ Google Sheets:\n{err_msg}")

    def _refresh_treeview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        if not self.form_cache:
            return
            
        self.form_cache.load() # Reload cache files
        responses = self.form_cache.get_all()
        for r in responses:
            self.tree.insert("", tk.END, values=(
                r.response_id,
                r.store_code,
                r.report_date,
                r.asm_name,
                r.cht_name,
                r.status
            ))

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
