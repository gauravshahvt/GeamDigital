import os
import sys
import time
import threading
import subprocess
from datetime import datetime
from typing import Optional, Dict, Any

# Ensure UTF-8 output on Windows
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# High-DPI scaling on Windows
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

# Add project root to sys.path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import customtkinter as ctk
from tkinter import filedialog, messagebox

from pdf_engine.district_automator import scan_district_hierarchy, run_district_automation
from pdf_engine.ocr_engine import is_tesseract_available

# Appearance
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class DistrictBatchDesktopApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Geam Digital - District Automated Voter Excel Studio Pro (2026)")
        self.geometry("1100x780")
        self.minsize(980, 680)

        # State
        self.is_scanning = False
        self.is_running = False
        self.stop_requested = False
        self.current_scan_data: Optional[Dict[str, Any]] = None
        self.summary_file_path: Optional[str] = None
        self.start_time: Optional[float] = None

        self._build_ui()

    def _build_ui(self):
        # Grid layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # -------------------------------------------------------------
        # 1. TOP HEADER BANNER
        # -------------------------------------------------------------
        self.header_frame = ctk.CTkFrame(self, corner_radius=12, fg_color=("#1E293B", "#0B1526"))
        self.header_frame.grid(row=0, column=0, padx=16, pady=(14, 8), sticky="ew")
        self.header_frame.grid_columnconfigure(0, weight=1)

        header_content = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        header_content.grid(row=0, column=0, padx=16, pady=12, sticky="ew")
        header_content.grid_columnconfigure(0, weight=1)

        # Title & Subtitle
        title_box = ctk.CTkFrame(header_content, fg_color="transparent")
        title_box.grid(row=0, column=0, sticky="w")

        title_badge = ctk.CTkLabel(
            title_box,
            text="🏛️ GEAM DIGITAL • DISTRICT AUTOMATION STUDIO 2026",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color="#38BDF8"
        )
        title_badge.pack(anchor="w")

        main_title = ctk.CTkLabel(
            title_box,
            text="सम्पूर्ण ज़िला वोटर एक्सेल ऑटोमेशन (District Batch Studio)",
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
            text_color="#F8FAFC"
        )
        main_title.pack(anchor="w", pady=(2, 0))

        sub_title = ctk.CTkLabel(
            title_box,
            text="ज़िले का मुख्य फ़ोल्डर चुनें: पंचायत समिति ➔ ग्राम पंचायत ➔ सभी वार्ड्स की मल्टी-शीट एक्सेल (.xlsx) फाइल स्वतः बनाकर वहीं सहेजेगा!",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#94A3B8"
        )
        sub_title.pack(anchor="w", pady=(2, 0))

        # Status Pill / Badges
        badge_box = ctk.CTkFrame(header_content, fg_color="transparent")
        badge_box.grid(row=0, column=1, sticky="e", padx=(10, 0))

        ocr_ok = is_tesseract_available()
        ocr_text = "✓ Tesseract OCR Active" if ocr_ok else "ℹ️ Standard Font Extraction"
        ocr_color = ("#10B981", "#059669") if ocr_ok else ("#F59E0B", "#D97706")
        
        self.ocr_badge = ctk.CTkLabel(
            badge_box,
            text=ocr_text,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=ocr_color,
            text_color="#FFFFFF",
            corner_radius=8,
            padx=12,
            pady=4
        )
        self.ocr_badge.pack(side="right", padx=4)

        # -------------------------------------------------------------
        # 2. FOLDER SELECTION & SCAN CARD
        # -------------------------------------------------------------
        self.folder_card = ctk.CTkFrame(self, corner_radius=12, fg_color=("#F1F5F9", "#1E293B"))
        self.folder_card.grid(row=1, column=0, padx=16, pady=6, sticky="ew")
        self.folder_card.grid_columnconfigure(1, weight=1)

        folder_label = ctk.CTkLabel(
            self.folder_card,
            text="📂 ज़िला मुख्य फ़ोल्डर:",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#E2E8F0"
        )
        folder_label.grid(row=0, column=0, padx=(16, 8), pady=12, sticky="w")

        self.folder_entry = ctk.CTkEntry(
            self.folder_card,
            placeholder_text="उदा. D:\\Bhilwara_District या फ़ोल्डर ब्राउज़ करें...",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            height=36
        )
        self.folder_entry.grid(row=0, column=1, padx=6, pady=12, sticky="ew")

        self.browse_btn = ctk.CTkButton(
            self.folder_card,
            text="📁 Browse Folder",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#334155",
            hover_color="#475569",
            height=36,
            width=130,
            command=self._on_browse_folder
        )
        self.browse_btn.grid(row=0, column=2, padx=6, pady=12)

        self.scan_btn = ctk.CTkButton(
            self.folder_card,
            text="🔍 Scan District",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#0284C7",
            hover_color="#0369A1",
            height=36,
            width=130,
            command=self._on_scan_district
        )
        self.scan_btn.grid(row=0, column=3, padx=(6, 16), pady=12)

        # -------------------------------------------------------------
        # 3. MAIN WORKSPACE: STATS + TREE + CONSOLE
        # -------------------------------------------------------------
        self.workspace_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.workspace_frame.grid(row=2, column=0, padx=16, pady=6, sticky="nsew")
        self.workspace_frame.grid_columnconfigure(0, weight=1)
        self.workspace_frame.grid_rowconfigure(2, weight=1)

        # Stats Cards Row (6 metrics)
        self.stats_frame = ctk.CTkFrame(self.workspace_frame, fg_color="transparent")
        self.stats_frame.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        for col in range(6):
            self.stats_frame.grid_columnconfigure(col, weight=1)

        self.stat_cards = {}
        metric_configs = [
            ("district_name", "मुख्य ज़िला", "-", "#94A3B8"),
            ("samitis", "पंचायत समितियां", "0", "#38BDF8"),
            ("panchayats", "ग्राम पंचायतें", "0", "#818CF8"),
            ("pdfs", "कुल वार्ड PDFs", "0", "#34D399"),
            ("already", "पूर्व निर्मित (Skip)", "0", "#FBBF24"),
            ("pending", "शेष कार्य (Pending)", "0", "#F87171"),
        ]

        for idx, (key, title, default_val, color) in enumerate(metric_configs):
            card = ctk.CTkFrame(self.stats_frame, corner_radius=10, fg_color=("#F8FAFC", "#1E293B"))
            card.grid(row=0, column=idx, padx=4, pady=0, sticky="ew")

            t_lbl = ctk.CTkLabel(card, text=title, font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"), text_color="#94A3B8")
            t_lbl.pack(pady=(6, 0))

            v_lbl = ctk.CTkLabel(card, text=default_val, font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"), text_color=color)
            v_lbl.pack(pady=(0, 6))

            self.stat_cards[key] = v_lbl

        # Execution Controls & Launch Bar
        self.control_bar = ctk.CTkFrame(self.workspace_frame, corner_radius=10, fg_color=("#0F172A", "#0F172A"))
        self.control_bar.grid(row=1, column=0, sticky="ew", pady=6)
        self.control_bar.grid_columnconfigure(1, weight=1)

        self.skip_check = ctk.CTkCheckBox(
            self.control_bar,
            text="पूर्व निर्मित पंचायतों को छोड़ें (Skip Already Processed)",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="#E2E8F0"
        )
        self.skip_check.grid(row=0, column=0, padx=16, pady=10, sticky="w")
        self.skip_check.select()

        # Theme selector
        theme_box = ctk.CTkFrame(self.control_bar, fg_color="transparent")
        theme_box.grid(row=0, column=1, sticky="w", padx=10)
        theme_lbl = ctk.CTkLabel(theme_box, text="एक्सेल थीम:", font=ctk.CTkFont(family="Segoe UI", size=11), text_color="#94A3B8")
        theme_lbl.pack(side="left", padx=(0, 6))

        self.theme_dropdown = ctk.CTkOptionMenu(
            theme_box,
            values=["Geam Digital Pro (Navy)", "Corporate Navy", "Emerald Finance", "Modern Slate", "Classic Office"],
            font=ctk.CTkFont(family="Segoe UI", size=11),
            width=170,
            height=28
        )
        self.theme_dropdown.set("Geam Digital Pro (Navy)")
        self.theme_dropdown.pack(side="left")

        # Action Buttons
        btn_box = ctk.CTkFrame(self.control_bar, fg_color="transparent")
        btn_box.grid(row=0, column=2, sticky="e", padx=16, pady=8)

        self.start_btn = ctk.CTkButton(
            btn_box,
            text="🚀 स्टार्ट ज़िला ऑटोमेशन",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color="#059669",
            hover_color="#047857",
            height=36,
            width=180,
            command=self._on_start_automation
        )
        self.start_btn.pack(side="left", padx=4)

        self.stop_btn = ctk.CTkButton(
            btn_box,
            text="⏹️ रोकें (Stop)",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#DC2626",
            hover_color="#B91C1C",
            height=36,
            width=100,
            state="disabled",
            command=self._on_stop_automation
        )
        self.stop_btn.pack(side="left", padx=4)

        # Middle Content: Progress Bar + Live Log Console
        self.middle_frame = ctk.CTkFrame(self.workspace_frame, corner_radius=10, fg_color=("#F8FAFC", "#1E293B"))
        self.middle_frame.grid(row=2, column=0, sticky="nsew", pady=4)
        self.middle_frame.grid_columnconfigure(0, weight=1)
        self.middle_frame.grid_rowconfigure(2, weight=1)

        # Progress Monitor Header
        prog_header = ctk.CTkFrame(self.middle_frame, fg_color="transparent")
        prog_header.grid(row=0, column=0, padx=14, pady=(10, 4), sticky="ew")
        prog_header.grid_columnconfigure(0, weight=1)

        self.progress_status_lbl = ctk.CTkLabel(
            prog_header,
            text="स्थिति: फ़ोल्डर स्कैन करने हेतु तैयार",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="#F8FAFC"
        )
        self.progress_status_lbl.grid(row=0, column=0, sticky="w")

        self.progress_percent_lbl = ctk.CTkLabel(
            prog_header,
            text="0%",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#38BDF8"
        )
        self.progress_percent_lbl.grid(row=0, column=1, sticky="e")

        # Progress Bar
        self.prog_bar = ctk.CTkProgressBar(self.middle_frame, height=12, corner_radius=6)
        self.prog_bar.grid(row=1, column=0, padx=14, pady=(0, 8), sticky="ew")
        self.prog_bar.set(0.0)

        # Live Console Text Box
        self.console_box = ctk.CTkTextbox(
            self.middle_frame,
            font=ctk.CTkFont(family="Consolas", size=11),
            fg_color="#090D16",
            text_color="#E2E8F0",
            corner_radius=8
        )
        self.console_box.grid(row=2, column=0, padx=14, pady=(0, 10), sticky="nsew")
        self._log("GEAM DIGITAL - DISTRICT AUTOMATED VOTER EXCEL STUDIO READY.")
        self._log("निर्देश: ऊपर फ़ोल्डर चुनें और 'Scan District' दबाएं।")

        # -------------------------------------------------------------
        # 4. BOTTOM ACTION & SUMMARY BAR
        # -------------------------------------------------------------
        self.bottom_bar = ctk.CTkFrame(self, corner_radius=10, fg_color=("#F1F5F9", "#1E293B"))
        self.bottom_bar.grid(row=3, column=0, padx=16, pady=(4, 14), sticky="ew")
        self.bottom_bar.grid_columnconfigure(0, weight=1)

        self.summary_info_lbl = ctk.CTkLabel(
            self.bottom_bar,
            text="तैयार | Geam Digital Pro (2026 Edition)",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color="#94A3B8"
        )
        self.summary_info_lbl.grid(row=0, column=0, padx=14, pady=8, sticky="w")

        btn_group = ctk.CTkFrame(self.bottom_bar, fg_color="transparent")
        btn_group.grid(row=0, column=1, padx=10, pady=8, sticky="e")

        self.open_summary_btn = ctk.CTkButton(
            btn_group,
            text="📊 Open Master Summary Excel",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color="#0284C7",
            hover_color="#0369A1",
            height=30,
            state="disabled",
            command=self._on_open_summary_excel
        )
        self.open_summary_btn.pack(side="left", padx=4)

        self.open_folder_btn = ctk.CTkButton(
            btn_group,
            text="📂 Open District Folder",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color="#334155",
            hover_color="#475569",
            height=30,
            state="disabled",
            command=self._on_open_district_folder
        )
        self.open_folder_btn.pack(side="left", padx=4)

    def _log(self, text: str):
        """Appends a timestamped line to the console box."""
        ts = datetime.now().strftime("%H:%M:%S")
        self.console_box.insert("end", f"[{ts}] {text}\n")
        self.console_box.see("end")

    def _on_browse_folder(self):
        folder = filedialog.askdirectory(title="ज़िले का मुख्य फ़ोल्डर चुनें (Select District Folder)")
        if folder:
            self.folder_entry.delete(0, "end")
            self.folder_entry.insert(0, folder)
            self._on_scan_district()

    def _on_scan_district(self):
        folder = self.folder_entry.get().strip('\"\' ')
        if not folder or not os.path.exists(folder):
            messagebox.showerror("अमान्य फ़ोल्डर", "कृपया एक मान्य ज़िला फ़ोल्डर चुनें या पाथ दर्ज करें।")
            return

        self.scan_btn.configure(state="disabled", text="Scanning...")
        self._log(f"फ़ोल्डर स्कैन किया जा रहा है: {folder} ...")

        def _worker():
            try:
                data = scan_district_hierarchy(folder)
                self.after(0, lambda: self._apply_scan_results(data))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("स्कैन त्रुटि", str(e)))
                self.after(0, lambda: self._log(f"त्रुटि: {e}"))
            finally:
                self.after(0, lambda: self.scan_btn.configure(state="normal", text="🔍 Scan District"))

        threading.Thread(target=_worker, daemon=True).start()

    def _apply_scan_results(self, data: Dict[str, Any]):
        self.current_scan_data = data
        self.stat_cards["district_name"].configure(text=data["district_name"])
        self.stat_cards["samitis"].configure(text=str(data["total_samitis"]))
        self.stat_cards["panchayats"].configure(text=str(data["total_panchayats"]))
        self.stat_cards["pdfs"].configure(text=f"{data['total_pdfs']:,}")
        self.stat_cards["already"].configure(text=str(data["already_processed_count"]))
        self.stat_cards["pending"].configure(text=str(data["pending_count"]))

        self.open_folder_btn.configure(state="normal")

        self._log("-" * 60)
        self._log(f"स्कैन परिणाम: {data['total_samitis']} समितियां, {data['total_panchayats']} पंचायतें, {data['total_pdfs']} PDFs मिले।")
        self._log(f"पूर्व निर्मित एक्सेल: {data['already_processed_count']} | शेष: {data['pending_count']}")

        for samiti, plist in data["samitis"].items():
            self._log(f"  • समिति: {samiti} ({len(plist)} पंचायतें)")

        self.progress_status_lbl.configure(
            text=f"स्कैन पूरा: {data['total_panchayats']} पंचायतें तैयार हैं। 'स्टार्ट ज़िला ऑटोमेशन' दबाएँ।"
        )

    def _on_start_automation(self):
        if not self.current_scan_data or self.current_scan_data["total_panchayats"] == 0:
            messagebox.showwarning("स्कैन आवश्यक", "कृपया पहले फ़ोल्डर स्कैन करें ताकि पंचायतों की पहचान हो सके।")
            return

        folder = self.folder_entry.get().strip('\"\' ')
        skip_existing = bool(self.skip_check.get())
        theme_map = {
            "Geam Digital Pro (Navy)": "geam_digital",
            "Corporate Navy": "navy",
            "Emerald Finance": "emerald",
            "Modern Slate": "charcoal",
            "Classic Office": "classic"
        }
        theme = theme_map.get(self.theme_dropdown.get(), "geam_digital")

        tot = self.current_scan_data["total_panchayats"]
        if not messagebox.askyesno("ऑटोमेशन पुष्टि", f"क्या आप सम्पूर्ण ज़िले की {tot} पंचायतों का स्वतः एक्सेल निर्माण प्रारंभ करना चाहते हैं?"):
            return

        self.is_running = True
        self.stop_requested = False
        self.start_time = time.time()
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.browse_btn.configure(state="disabled")
        self.scan_btn.configure(state="disabled")

        self.prog_bar.set(0.0)
        self.progress_percent_lbl.configure(text="0%")

        self._log("=" * 60)
        self._log("🚀 सम्पूर्ण ज़िला ऑटोमेशन प्रारंभ हुआ...")

        def _worker():
            def _on_progress(event):
                if self.stop_requested:
                    raise KeyboardInterrupt("उपयोगकर्ता द्वारा रोका गया")
                self.after(0, lambda e=event: self._handle_progress_event(e))

            try:
                res = run_district_automation(
                    root_dir=folder,
                    skip_existing=skip_existing,
                    theme=theme,
                    progress_callback=_on_progress
                )
                self.after(0, lambda: self._on_automation_finished(res))
            except KeyboardInterrupt:
                self.after(0, lambda: self._log("⚠️ ऑटोमेशन उपयोगकर्ता द्वारा रोक दिया गया।"))
                self.after(0, self._reset_run_state)
            except Exception as e:
                self.after(0, lambda: self._log(f"❌ त्रुटि: {e}"))
                self.after(0, lambda: messagebox.showerror("ऑटोमेशन त्रुटि", str(e)))
                self.after(0, self._reset_run_state)

        threading.Thread(target=_worker, daemon=True).start()

    def _handle_progress_event(self, event: Dict[str, Any]):
        e_type = event.get("type")
        pct = event.get("percent", 0.0)
        idx = event.get("index", 0)
        tot = event.get("total_panchayats", 1)

        self.prog_bar.set(pct / 100.0)
        self.progress_percent_lbl.configure(text=f"{pct}%")

        if e_type == "start_panchayat":
            samiti = event.get("samiti", "")
            panchayat = event.get("panchayat", "")
            cnt = event.get("pdf_count", 0)
            self.progress_status_lbl.configure(
                text=f"[{idx}/{tot}] समिति: {samiti} ➔ पंचायत: {panchayat} ({cnt} वार्ड्स प्रोसेस हो रहे हैं)..."
            )
            self._log(f"[{idx}/{tot}] ⏳ {samiti} ➔ {panchayat} ({cnt} वार्ड्स)...")
        elif e_type == "complete_panchayat":
            panchayat = event.get("panchayat", "")
            voters = event.get("active_voters", 0)
            f_name = event.get("excel_file", "")
            self._log(f"[{idx}/{tot}] ✓ सफल: {panchayat} ➔ {voters:,} मतदाता सहेजे गए ({f_name})")
        elif e_type == "skip_panchayat":
            panchayat = event.get("panchayat", "")
            self._log(f"[{idx}/{tot}] ⏩ स्किप: {panchayat} (पूर्व निर्मित फ़ाइल मौजूद)")
        elif e_type == "error_panchayat":
            panchayat = event.get("panchayat", "")
            err = event.get("error", "")
            self._log(f"[{idx}/{tot}] ❌ त्रुटि: {panchayat} ➔ {err}")

    def _on_automation_finished(self, res: Dict[str, Any]):
        self._reset_run_state()
        self.summary_file_path = res.get("summary_file")
        if self.summary_file_path and os.path.exists(self.summary_file_path):
            self.open_summary_btn.configure(state="normal")

        mins = int(res["elapsed_seconds"] // 60)
        secs = int(res["elapsed_seconds"] % 60)

        self.progress_status_lbl.configure(
            text=f"🎉 ज़िला ऑटोमेशन संपन्न! {res['completed_panchayats']} पंचायतें पूर्ण ({res['total_voters']:,} मतदाता)।"
        )
        self.prog_bar.set(1.0)
        self.progress_percent_lbl.configure(text="100%")

        self._log("=" * 60)
        self._log("🎉 सम्पूर्ण ज़िला ऑटोमेशन सफलतापूर्वक संपन्न!")
        self._log(f"• पूर्ण पंचायतें: {res['completed_panchayats']} / {res['total_panchayats']}")
        self._log(f"• कुल सक्रिय मतदाता: {res['total_voters']:,}")
        self._log(f"• कुल समय लगा: {mins} मिनट {secs} सेकंड")
        self._log(f"• मास्टर रिपोर्ट: {self.summary_file_path}")
        self._log("सभी एक्सेल फाइलें संबंधित पंचायत फ़ोल्डर में सहेजी जा चुकी हैं।")

        messagebox.showinfo(
            "ज़िला ऑटोमेशन पूर्ण",
            f"बधाई! ज़िला ऑटोमेशन सफलतापूर्वक पूर्ण हुआ!\n\n"
            f"• कुल पंचायतें: {res['completed_panchayats']} / {res['total_panchayats']}\n"
            f"• कुल मतदाता: {res['total_voters']:,}\n"
            f"• समय: {mins} मिनट {secs} सेकंड\n\n"
            f"मास्टर रिपोर्ट नीचे दिए गए बटन से खोल सकते हैं।"
        )

    def _on_stop_automation(self):
        if self.is_running:
            if messagebox.askyesno("ऑटोमेशन रोकें", "क्या आप चालू ऑटोमेशन को रोकना चाहते हैं?"):
                self.stop_requested = True
                self._log("⚠️ रोकने का अनुरोध भेजा गया...")

    def _reset_run_state(self):
        self.is_running = False
        self.stop_requested = False
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.browse_btn.configure(state="normal")
        self.scan_btn.configure(state="normal")

    def _on_open_summary_excel(self):
        if self.summary_file_path and os.path.exists(self.summary_file_path):
            try:
                os.startfile(self.summary_file_path)
            except Exception as e:
                messagebox.showerror("फ़ाइल खोलने में त्रुटि", str(e))

    def _on_open_district_folder(self):
        folder = self.folder_entry.get().strip('\"\' ')
        if folder and os.path.exists(folder):
            try:
                os.startfile(folder)
            except Exception as e:
                messagebox.showerror("फ़ोल्डर खोलने में त्रुटि", str(e))

def main():
    app = DistrictBatchDesktopApp()
    app.mainloop()

if __name__ == "__main__":
    main()
