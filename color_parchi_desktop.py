import os
import sys
import time
import base64
import threading
import subprocess
from datetime import datetime
from typing import Optional, Dict, Any, List

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

import pymupdf
from pdf_engine.parchi_generator import parse_excel_for_parchi, get_browser_executable
from pdf_engine.color_parchi_generator import (
    generate_color_parchi_pdf,
    DEFAULT_HAND_DATA,
    DEFAULT_LOTUS_DATA,
    DEFAULT_CYCLE_DATA,
    DEFAULT_CANDIDATE_DATA,
    SYMBOL_PRESETS
)

# Appearance
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

def image_to_base64_data_url(file_path: str) -> str:
    """Converts a local image file into a Data URL for HTML embedding."""
    if not file_path or not os.path.exists(file_path):
        return ""
    ext = os.path.splitext(file_path)[1].lower().replace('.', '')
    if ext == 'svg':
        mime = 'image/svg+xml'
    elif ext in ['jpg', 'jpeg']:
        mime = 'image/jpeg'
    elif ext == 'webp':
        mime = 'image/webp'
    else:
        mime = 'image/png'
    try:
        with open(file_path, 'rb') as f:
            encoded = base64.b64encode(f.read()).decode('ascii')
        return f"data:{mime};base64,{encoded}"
    except Exception:
        return ""

class ColorVoterSlipDesktopApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Geam Digital - Color Voter Slip & Poster Studio Pro (2026)")
        self.geometry("1180x840")
        self.minsize(1020, 720)

        # State Variables
        self.current_excel_path: Optional[str] = None
        self.parsed_data: Optional[Dict[str, Any]] = None
        self.is_generating = False
        self.last_generated_pdf: Optional[str] = None
        self.custom_symbol_data_url = ""
        self.custom_photo_data_url = ""
        self.custom_poster_data_url = ""

        self._build_ui()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # -------------------------------------------------------------
        # 1. HEADER BANNER
        # -------------------------------------------------------------
        header_frame = ctk.CTkFrame(self, corner_radius=12, fg_color=("#1E293B", "#0B1526"))
        header_frame.grid(row=0, column=0, padx=16, pady=(12, 6), sticky="ew")
        header_frame.grid_columnconfigure(0, weight=1)

        header_box = ctk.CTkFrame(header_frame, fg_color="transparent")
        header_box.grid(row=0, column=0, padx=16, pady=10, sticky="ew")
        header_box.grid_columnconfigure(0, weight=1)

        title_badge = ctk.CTkLabel(
            header_box,
            text="🎨 GEAM DIGITAL • COLOR VOTER SLIP STUDIO 2026",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color="#F59E0B"
        )
        title_badge.pack(anchor="w")

        main_title = ctk.CTkLabel(
            header_box,
            text="कलर मतदाता पर्ची एवं उम्मीदवार पोस्टर डेस्कटॉप स्टूडियो (Color Voter Slip Pro)",
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
            text_color="#F8FAFC"
        )
        main_title.pack(anchor="w", pady=(2, 0))

        sub_title = ctk.CTkLabel(
            header_box,
            text="एक्सेल फ़ाइल चुनें ➔ प्रत्याशी विवरण, फोटो व चुनाव चिन्ह सेट करें ➔ उच्च गुणवत्ता वाली A4 कलर पर्ची PDF तुरंत बनाएं!",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#94A3B8"
        )
        sub_title.pack(anchor="w", pady=(2, 0))

        # Browser Engine Pill
        badge_box = ctk.CTkFrame(header_box, fg_color="transparent")
        badge_box.place(relx=1.0, rely=0.5, anchor="e")

        browser_path = get_browser_executable()
        b_name = "Chrome" if browser_path and "chrome" in browser_path.lower() else ("Edge" if browser_path else "Not Found")
        b_ok = bool(browser_path)
        b_text = f"✓ {b_name} Engine Active (Fast Chunked)" if b_ok else "❌ No Browser Found"
        b_color = ("#10B981", "#059669") if b_ok else ("#EF4444", "#DC2626")

        engine_badge = ctk.CTkLabel(
            badge_box,
            text=b_text,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=b_color,
            text_color="#FFFFFF",
            corner_radius=8,
            padx=12,
            pady=4
        )
        engine_badge.pack(side="right")

        # -------------------------------------------------------------
        # 2. MAIN SPLIT WORKSPACE: LEFT (INPUTS) + RIGHT (PREVIEW & LOGS)
        # -------------------------------------------------------------
        workspace = ctk.CTkFrame(self, fg_color="transparent")
        workspace.grid(row=1, column=0, padx=16, pady=6, sticky="nsew")
        workspace.grid_columnconfigure(0, weight=5)  # Left (Controls)
        workspace.grid_columnconfigure(1, weight=5)  # Right (Preview & Progress)
        workspace.grid_rowconfigure(0, weight=1)

        # =============================================================
        # LEFT COLUMN: SCROLLABLE SETTINGS & INPUTS
        # =============================================================
        left_scroll = ctk.CTkScrollableFrame(workspace, corner_radius=12, fg_color=("#F1F5F9", "#111C2E"))
        left_scroll.grid(row=0, column=0, padx=(0, 8), pady=0, sticky="nsew")
        left_scroll.grid_columnconfigure(0, weight=1)

        # -------------------------------------------------------------
        # CARD 1: EXCEL FILE & SHEET SELECTION
        # -------------------------------------------------------------
        c1 = ctk.CTkFrame(left_scroll, corner_radius=10, fg_color=("#E2E8F0", "#1E293B"))
        c1.pack(fill="x", padx=6, pady=6)
        c1.grid_columnconfigure(1, weight=1)

        c1_title = ctk.CTkLabel(
            c1,
            text="📂 1. एक्सेल फ़ाइल एवं वार्ड चयन (Select Voter Excel)",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#38BDF8"
        )
        c1_title.grid(row=0, column=0, columnspan=3, padx=12, pady=(10, 6), sticky="w")

        self.file_entry = ctk.CTkEntry(
            c1,
            placeholder_text="कृपया वोटर एक्सेल फाइल (.xlsx) चुनें...",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            height=32
        )
        self.file_entry.grid(row=1, column=0, columnspan=2, padx=(12, 6), pady=6, sticky="ew")

        browse_file_btn = ctk.CTkButton(
            c1,
            text="📁 Browse",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            width=90,
            height=32,
            fg_color="#0284C7",
            hover_color="#0369A1",
            command=self._on_browse_file
        )
        browse_file_btn.grid(row=1, column=2, padx=(0, 12), pady=6)

        # Sheet Selection Dropdown
        sheet_lbl = ctk.CTkLabel(
            c1,
            text="शीट / वार्ड चुनें:",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color="#CBD5E1"
        )
        sheet_lbl.grid(row=2, column=0, padx=(12, 6), pady=6, sticky="w")

        self.sheet_combo = ctk.CTkComboBox(
            c1,
            values=["फ़ाइल लोड करें..."],
            font=ctk.CTkFont(family="Segoe UI", size=11),
            height=32,
            command=self._on_sheet_changed
        )
        self.sheet_combo.grid(row=2, column=1, columnspan=2, padx=(0, 12), pady=6, sticky="ew")

        # Quick Stats Banner
        self.stats_label = ctk.CTkLabel(
            c1,
            text="📊 स्थिति: कोई फ़ाइल लोड नहीं है",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color="#94A3B8"
        )
        self.stats_label.grid(row=3, column=0, columnspan=3, padx=12, pady=(4, 10), sticky="w")

        # -------------------------------------------------------------
        # CARD 2: CANDIDATE & POSTER INFORMATION
        # -------------------------------------------------------------
        c2 = ctk.CTkFrame(left_scroll, corner_radius=10, fg_color=("#E2E8F0", "#1E293B"))
        c2.pack(fill="x", padx=6, pady=6)
        c2.grid_columnconfigure(1, weight=1)

        c2_title = ctk.CTkLabel(
            c2,
            text="👤 2. प्रत्याशी विवरण एवं पोस्टर सेटिंग (Candidate Info)",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#F59E0B"
        )
        c2_title.grid(row=0, column=0, columnspan=3, padx=12, pady=(10, 6), sticky="w")

        # Candidate Name
        ctk.CTkLabel(c2, text="प्रत्याशी का नाम:", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")).grid(row=1, column=0, padx=12, pady=4, sticky="w")
        self.cand_name_entry = ctk.CTkEntry(c2, font=ctk.CTkFont(family="Segoe UI", size=11), height=30)
        self.cand_name_entry.insert(0, "संजय कुमार मेवाड़ा")
        self.cand_name_entry.grid(row=1, column=1, columnspan=2, padx=(0, 12), pady=4, sticky="ew")
        self.cand_name_entry.bind("<KeyRelease>", lambda e: self._update_preview())

        # Header Prefix & Ward Badge & Suffix
        ctk.CTkLabel(c2, text="पद / हेडर शीर्षक:", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")).grid(row=2, column=0, padx=12, pady=4, sticky="w")
        h_box = ctk.CTkFrame(c2, fg_color="transparent")
        h_box.grid(row=2, column=1, columnspan=2, padx=(0, 12), pady=4, sticky="ew")
        h_box.grid_columnconfigure(0, weight=3)
        h_box.grid_columnconfigure(1, weight=1)
        h_box.grid_columnconfigure(2, weight=3)

        self.header_prefix_entry = ctk.CTkEntry(h_box, font=ctk.CTkFont(family="Segoe UI", size=11), height=30)
        self.header_prefix_entry.insert(0, "पार्षद पद हेतु वार्ड नं")
        self.header_prefix_entry.grid(row=0, column=0, padx=(0, 4), sticky="ew")
        self.header_prefix_entry.bind("<KeyRelease>", lambda e: self._update_preview())

        self.ward_badge_entry = ctk.CTkEntry(h_box, font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"), height=30, width=50)
        self.ward_badge_entry.insert(0, "03")
        self.ward_badge_entry.grid(row=0, column=1, padx=4, sticky="ew")
        self.ward_badge_entry.bind("<KeyRelease>", lambda e: self._update_preview())

        self.header_suffix_entry = ctk.CTkEntry(h_box, font=ctk.CTkFont(family="Segoe UI", size=11), height=30)
        self.header_suffix_entry.insert(0, "से लोकप्रिय प्रत्याशी")
        self.header_suffix_entry.grid(row=0, column=2, padx=(4, 0), sticky="ew")
        self.header_suffix_entry.bind("<KeyRelease>", lambda e: self._update_preview())

        # Symbol & Button No
        ctk.CTkLabel(c2, text="चुनाव चिन्ह (Symbol):", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")).grid(row=3, column=0, padx=12, pady=4, sticky="w")
        sym_box = ctk.CTkFrame(c2, fg_color="transparent")
        sym_box.grid(row=3, column=1, columnspan=2, padx=(0, 12), pady=4, sticky="ew")
        sym_box.grid_columnconfigure(0, weight=2)
        sym_box.grid_columnconfigure(1, weight=1)

        self.symbol_combo = ctk.CTkComboBox(
            sym_box,
            values=["हाथ (Hand)", "कमल (Lotus)", "साइकिल (Cycle)", "कस्टम फ़ोटो (Custom Image)"],
            font=ctk.CTkFont(family="Segoe UI", size=11),
            height=30,
            command=self._on_symbol_changed
        )
        self.symbol_combo.grid(row=0, column=0, padx=(0, 6), sticky="ew")

        self.browse_sym_btn = ctk.CTkButton(
            sym_box,
            text="चिन्ह फोटो चुनें",
            font=ctk.CTkFont(family="Segoe UI", size=10),
            height=30,
            fg_color="#334155",
            hover_color="#475569",
            command=self._on_browse_symbol
        )
        self.browse_sym_btn.grid(row=0, column=1, sticky="ew")

        # Button No
        ctk.CTkLabel(c2, text="ईवीएम बटन संख्या:", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")).grid(row=4, column=0, padx=12, pady=4, sticky="w")
        self.btn_no_entry = ctk.CTkEntry(c2, font=ctk.CTkFont(family="Segoe UI", size=11), height=30)
        self.btn_no_entry.insert(0, "2")
        self.btn_no_entry.grid(row=4, column=1, columnspan=2, padx=(0, 12), pady=4, sticky="ew")
        self.btn_no_entry.bind("<KeyRelease>", lambda e: self._update_preview())

        # Slogan
        ctk.CTkLabel(c2, text="स्लोगन (Slogan):", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")).grid(row=5, column=0, padx=12, pady=4, sticky="w")
        self.slogan_entry = ctk.CTkEntry(c2, font=ctk.CTkFont(family="Segoe UI", size=11), height=30)
        self.slogan_entry.insert(0, "को हाथ के निशान पर बटन दबा कर विजयी बनावें।")
        self.slogan_entry.grid(row=5, column=1, columnspan=2, padx=(0, 12), pady=4, sticky="ew")
        self.slogan_entry.bind("<KeyRelease>", lambda e: self._update_preview())

        # Voting Time
        ctk.CTkLabel(c2, text="मतदान समय:", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")).grid(row=6, column=0, padx=12, pady=4, sticky="w")
        self.time_entry = ctk.CTkEntry(c2, font=ctk.CTkFont(family="Segoe UI", size=11), height=30)
        self.time_entry.insert(0, "प्रातः 7 से सायं 6 बजे तक")
        self.time_entry.grid(row=6, column=1, columnspan=2, padx=(0, 12), pady=4, sticky="ew")
        self.time_entry.bind("<KeyRelease>", lambda e: self._update_preview())

        # Candidate Photo & Ready-Made Poster Browse
        ctk.CTkLabel(c2, text="फ़ोटो / पोस्टर:", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")).grid(row=7, column=0, padx=12, pady=4, sticky="w")
        photo_box = ctk.CTkFrame(c2, fg_color="transparent")
        photo_box.grid(row=7, column=1, columnspan=2, padx=(0, 12), pady=(4, 10), sticky="ew")
        photo_box.grid_columnconfigure(0, weight=1)
        photo_box.grid_columnconfigure(1, weight=1)

        self.browse_photo_btn = ctk.CTkButton(
            photo_box,
            text="📷 प्रत्याशी फोटो चुनें",
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            height=30,
            fg_color="#0F766E",
            hover_color="#115E59",
            command=self._on_browse_photo
        )
        self.browse_photo_btn.grid(row=0, column=0, padx=(0, 4), sticky="ew")

        self.browse_poster_btn = ctk.CTkButton(
            photo_box,
            text="🖼️ रेडीमेड पोस्टर चुनें",
            font=ctk.CTkFont(family="Segoe UI", size=10),
            height=30,
            fg_color="#334155",
            hover_color="#475569",
            command=self._on_browse_poster
        )
        self.browse_poster_btn.grid(row=0, column=1, padx=(4, 0), sticky="ew")

        # -------------------------------------------------------------
        # CARD 3: LAYOUT & FILTERS
        # -------------------------------------------------------------
        c3 = ctk.CTkFrame(left_scroll, corner_radius=10, fg_color=("#E2E8F0", "#1E293B"))
        c3.pack(fill="x", padx=6, pady=6)
        c3.grid_columnconfigure(1, weight=1)

        c3_title = ctk.CTkLabel(
            c3,
            text="⚙️ 3. लेआउट एवं फ़िल्टर (Layout & Print Options)",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#10B981"
        )
        c3_title.grid(row=0, column=0, columnspan=3, padx=12, pady=(10, 6), sticky="w")

        # Layout Type
        ctk.CTkLabel(c3, text="लेआउट साइज़:", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")).grid(row=1, column=0, padx=12, pady=4, sticky="w")
        self.layout_seg = ctk.CTkSegmentedButton(
            c3,
            values=["A4 - 8 Slips (2x4)", "A4 - 12 Slips (3x4)", "A4 - 4 Slips (Left Photo + Right Slip)"],
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            command=lambda v: self._update_preview()
        )
        self.layout_seg.set("A4 - 8 Slips (2x4)")
        self.layout_seg.grid(row=1, column=1, columnspan=2, padx=(0, 12), pady=4, sticky="ew")

        # Checkboxes
        chk_box = ctk.CTkFrame(c3, fg_color="transparent")
        chk_box.grid(row=2, column=0, columnspan=3, padx=12, pady=4, sticky="w")

        self.active_only_var = ctk.BooleanVar(value=True)
        self.active_chk = ctk.CTkCheckBox(
            chk_box,
            text="केवल सक्रिय मतदाता (विलोपित छोड़ें / Skip Deleted)",
            variable=self.active_only_var,
            font=ctk.CTkFont(family="Segoe UI", size=11)
        )
        self.active_chk.pack(side="left", padx=(0, 16))

        self.page_num_var = ctk.BooleanVar(value=True)
        self.page_num_chk = ctk.CTkCheckBox(
            chk_box,
            text="पृष्ठ संख्या दिखाएं (Show Page No)",
            variable=self.page_num_var,
            font=ctk.CTkFont(family="Segoe UI", size=11)
        )
        self.page_num_chk.pack(side="left")

        # Serial Range
        ctk.CTkLabel(c3, text="क्रम संख्या फ़िल्टर:", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")).grid(row=3, column=0, padx=12, pady=4, sticky="w")
        range_box = ctk.CTkFrame(c3, fg_color="transparent")
        range_box.grid(row=3, column=1, columnspan=2, padx=(0, 12), pady=(4, 10), sticky="w")

        self.all_voters_var = ctk.BooleanVar(value=True)
        self.all_voters_chk = ctk.CTkCheckBox(
            range_box,
            text="सभी मतदाता",
            variable=self.all_voters_var,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            command=self._toggle_range_inputs
        )
        self.all_voters_chk.pack(side="left", padx=(0, 12))

        self.range_from_entry = ctk.CTkEntry(range_box, placeholder_text="से", width=55, height=28, font=ctk.CTkFont(family="Segoe UI", size=11))
        self.range_from_entry.pack(side="left", padx=2)
        self.range_from_entry.configure(state="disabled")

        ctk.CTkLabel(range_box, text="-", font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=2)

        self.range_to_entry = ctk.CTkEntry(range_box, placeholder_text="तक", width=55, height=28, font=ctk.CTkFont(family="Segoe UI", size=11))
        self.range_to_entry.pack(side="left", padx=2)
        self.range_to_entry.configure(state="disabled")

        # =============================================================
        # RIGHT COLUMN: PREVIEW + GENERATE ACTION + CONSOLE
        # =============================================================
        right_frame = ctk.CTkFrame(workspace, corner_radius=12, fg_color=("#F1F5F9", "#111C2E"))
        right_frame.grid(row=0, column=1, padx=(8, 0), pady=0, sticky="nsew")
        right_frame.grid_columnconfigure(0, weight=1)
        right_frame.grid_rowconfigure(2, weight=1)

        # -------------------------------------------------------------
        # CARD 4: LIVE PREVIEW OF COLOR VOTER SLIP
        # -------------------------------------------------------------
        preview_card = ctk.CTkFrame(right_frame, corner_radius=10, fg_color=("#E2E8F0", "#1E293B"))
        preview_card.grid(row=0, column=0, padx=12, pady=(12, 6), sticky="ew")
        preview_card.grid_columnconfigure(0, weight=1)

        preview_title = ctk.CTkLabel(
            preview_card,
            text="👁️ लाइव पर्ची प्रीव्यू (Live Color Voter Slip Card Preview)",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="#38BDF8"
        )
        preview_title.pack(anchor="w", padx=12, pady=(8, 4))

        # Simulated Graphic Card Frame
        self.mock_card = ctk.CTkFrame(preview_card, corner_radius=8, fg_color="#FFFFFF", border_width=2, border_color="#000000")
        self.mock_card.pack(fill="x", padx=14, pady=(4, 10))
        self.mock_card.grid_columnconfigure(1, weight=1)

        # Mock Poster Bar (Yellow Top)
        self.mock_poster_bar = ctk.CTkFrame(self.mock_card, fg_color="#FFF033", corner_radius=0)
        self.mock_poster_bar.pack(fill="x", padx=0, pady=0)

        self.mock_ribbon_lbl = ctk.CTkLabel(
            self.mock_poster_bar,
            text="पार्षद पद हेतु वार्ड नं [ 03 ] से लोकप्रिय प्रत्याशी",
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            text_color="#000000"
        )
        self.mock_ribbon_lbl.pack(pady=(2, 0))

        mock_body = ctk.CTkFrame(self.mock_poster_bar, fg_color="transparent")
        mock_body.pack(fill="x", padx=8, pady=(2, 4))
        mock_body.grid_columnconfigure(1, weight=1)

        self.mock_sym_lbl = ctk.CTkLabel(
            mock_body,
            text="✋\nचिन्ह",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color="#000000",
            width=50
        )
        self.mock_sym_lbl.grid(row=0, column=0, padx=4)

        center_col = ctk.CTkFrame(mock_body, fg_color="transparent")
        center_col.grid(row=0, column=1)

        self.mock_cand_name_lbl = ctk.CTkLabel(
            center_col,
            text="संजय कुमार मेवाड़ा",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color="#B91C1C"
        )
        self.mock_cand_name_lbl.pack()

        self.mock_btn_lbl = ctk.CTkLabel(
            center_col,
            text="बटन नं. 2",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color="#0284C7",
            text_color="#FFFFFF",
            corner_radius=4,
            padx=10,
            pady=1
        )
        self.mock_btn_lbl.pack(pady=2)

        self.mock_slogan_lbl = ctk.CTkLabel(
            center_col,
            text="को हाथ के निशान पर बटन दबा कर विजयी बनावें।",
            font=ctk.CTkFont(family="Segoe UI", size=9),
            text_color="#1E293B"
        )
        self.mock_slogan_lbl.pack()

        self.mock_photo_lbl = ctk.CTkLabel(
            mock_body,
            text="👤\nफोटो",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color="#000000",
            width=50
        )
        self.mock_photo_lbl.grid(row=0, column=2, padx=4)

        # Mock Cut Divider
        cut_frame = ctk.CTkFrame(self.mock_card, fg_color="#F8FAFC", corner_radius=0)
        cut_frame.pack(fill="x", padx=0, pady=0)
        ctk.CTkLabel(
            cut_frame,
            text="✂ ----------------- मतदान केंद्र पर जाने से पहले ऊपर का हिस्सा फाड़ देवे ----------------- ✂",
            font=ctk.CTkFont(family="Segoe UI", size=8),
            text_color="#64748B"
        ).pack(pady=1)

        # Mock Voter Data
        v_frame = ctk.CTkFrame(self.mock_card, fg_color="#FFFFFF", corner_radius=0)
        v_frame.pack(fill="x", padx=8, pady=(2, 6))

        ctk.CTkLabel(
            v_frame,
            text="क्रम सं. : [ 1 ]   |   वार्ड नं. - 03   |   भाग सं : 03   |   Voter ID : RJ/12/345/67890",
            font=ctk.CTkFont(family="Segoe UI", size=9, weight="bold"),
            text_color="#0F172A"
        ).pack(anchor="w")

        ctk.CTkLabel(
            v_frame,
            text="नाम : रमेश कुमार   |   पिता/पति : मोहन लाल   |   मकान सं : 45   |   पुरूष 38",
            font=ctk.CTkFont(family="Segoe UI", size=9),
            text_color="#334155"
        ).pack(anchor="w")

        self.mock_time_bar = ctk.CTkLabel(
            v_frame,
            text="मतदान समय : प्रातः 7 से सायं 6 बजे तक",
            font=ctk.CTkFont(family="Segoe UI", size=9, weight="bold"),
            fg_color="#E11D48",
            text_color="#FFFFFF",
            corner_radius=2
        )
        self.mock_time_bar.pack(fill="x", pady=(4, 0))

        # -------------------------------------------------------------
        # CARD 5: GENERATE BUTTON & REAL-TIME PROGRESS
        # -------------------------------------------------------------
        action_card = ctk.CTkFrame(right_frame, corner_radius=10, fg_color=("#E2E8F0", "#1E293B"))
        action_card.grid(row=1, column=0, padx=12, pady=6, sticky="ew")
        action_card.grid_columnconfigure(0, weight=1)

        self.generate_btn = ctk.CTkButton(
            action_card,
            text="🚀 कलर वोटर पर्ची PDF बनाएं (Generate Color Voter Slips PDF)",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            fg_color="#059669",
            hover_color="#047857",
            height=46,
            command=self._on_start_generation
        )
        self.generate_btn.pack(fill="x", padx=14, pady=(12, 8))

        self.progress_bar = ctk.CTkProgressBar(action_card, height=14)
        self.progress_bar.pack(fill="x", padx=14, pady=4)
        self.progress_bar.set(0.0)

        self.status_lbl = ctk.CTkLabel(
            action_card,
            text="तैयार (Ready) - एक्सेल फ़ाइल लोड करें",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color="#38BDF8"
        )
        self.status_lbl.pack(anchor="w", padx=14, pady=(2, 6))

        # Completion Action Buttons (Initially Hidden/Disabled)
        btn_row = ctk.CTkFrame(action_card, fg_color="transparent")
        btn_row.pack(fill="x", padx=14, pady=(0, 10))
        btn_row.grid_columnconfigure(0, weight=1)
        btn_row.grid_columnconfigure(1, weight=1)

        self.open_pdf_btn = ctk.CTkButton(
            btn_row,
            text="📄 PDF फ़ाइल खोलें (Open PDF)",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            height=34,
            fg_color="#0284C7",
            hover_color="#0369A1",
            state="disabled",
            command=self._on_open_pdf
        )
        self.open_pdf_btn.grid(row=0, column=0, padx=(0, 4), sticky="ew")

        self.open_folder_btn = ctk.CTkButton(
            btn_row,
            text="📂 फ़ोल्डर खोलें (Open Folder)",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            height=34,
            fg_color="#334155",
            hover_color="#475569",
            state="disabled",
            command=self._on_open_folder
        )
        self.open_folder_btn.grid(row=0, column=1, padx=(4, 0), sticky="ew")

        # -------------------------------------------------------------
        # CARD 6: CONSOLE LOG
        # -------------------------------------------------------------
        console_card = ctk.CTkFrame(right_frame, corner_radius=10, fg_color=("#E2E8F0", "#1E293B"))
        console_card.grid(row=2, column=0, padx=12, pady=(6, 12), sticky="nsew")
        console_card.grid_columnconfigure(0, weight=1)
        console_card.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            console_card,
            text="📋 रीयल-टाइम एक्टिविटी लॉग (Live Generation Log):",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color="#94A3B8"
        ).grid(row=0, column=0, padx=12, pady=(8, 2), sticky="w")

        self.log_text = ctk.CTkTextbox(
            console_card,
            font=ctk.CTkFont(family="Consolas", size=10),
            corner_radius=6,
            fg_color=("#0F172A", "#080F1A"),
            text_color="#E2E8F0"
        )
        self.log_text.grid(row=1, column=0, padx=12, pady=(2, 10), sticky="nsew")

        self._log("Geam Digital Color Voter Slip Studio प्रारंभ हुआ।")
        self._log(f"ब्राउज़र रेंडरिंग इंजन: {b_name} ({browser_path})")

    # -------------------------------------------------------------
    # LOGGING & PREVIEW HELPERS
    # -------------------------------------------------------------
    def _log(self, msg: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{timestamp}] {msg}\n")
        self.log_text.see("end")

    def _update_preview(self):
        c_name = self.cand_name_entry.get().strip() or "प्रत्याशी नाम"
        w_badge = self.ward_badge_entry.get().strip() or "01"
        prefix = self.header_prefix_entry.get().strip()
        suffix = self.header_suffix_entry.get().strip()
        btn = self.btn_no_entry.get().strip() or "1"
        slogan = self.slogan_entry.get().strip()
        time_txt = self.time_entry.get().strip()

        self.mock_cand_name_lbl.configure(text=c_name)
        self.mock_ribbon_lbl.configure(text=f"{prefix} [ {w_badge} ] {suffix}")
        self.mock_btn_lbl.configure(text=f"बटन नं. {btn}")
        self.mock_slogan_lbl.configure(text=slogan)
        self.mock_time_bar.configure(text=f"मतदान समय : {time_txt}")

    def _toggle_range_inputs(self):
        is_all = self.all_voters_var.get()
        if is_all:
            self.range_from_entry.configure(state="disabled")
            self.range_to_entry.configure(state="disabled")
        else:
            self.range_from_entry.configure(state="normal")
            self.range_to_entry.configure(state="normal")

    def _on_symbol_changed(self, choice):
        if "हाथ" in choice:
            self.mock_sym_lbl.configure(text="✋\nहाथ")
            self.custom_symbol_data_url = ""
        elif "कमल" in choice:
            self.mock_sym_lbl.configure(text="🪷\nकमल")
            self.custom_symbol_data_url = ""
        elif "साइकिल" in choice:
            self.mock_sym_lbl.configure(text="🚲\nसाइकिल")
            self.custom_symbol_data_url = ""
        elif "कस्टम" in choice:
            self._on_browse_symbol()

    def _on_browse_symbol(self):
        path = filedialog.askopenfilename(
            title="चुनाव चिन्ह फोटो चुनें (Select Symbol Image)",
            filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.svg;*.webp"), ("All Files", "*.*")]
        )
        if path:
            self.custom_symbol_data_url = image_to_base64_data_url(path)
            self.symbol_combo.set("कस्टम फ़ोटो (Custom Image)")
            self.mock_sym_lbl.configure(text="🖼️\nकस्टम")
            self._log(f"कस्टम चुनाव चिन्ह लोड किया गया: {os.path.basename(path)}")

    def _on_browse_photo(self):
        path = filedialog.askopenfilename(
            title="उम्मीदवार की फोटो चुनें (Select Candidate Photo)",
            filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.svg;*.webp"), ("All Files", "*.*")]
        )
        if path:
            self.custom_photo_data_url = image_to_base64_data_url(path)
            self.mock_photo_lbl.configure(text="📷\nकस्टम")
            self._log(f"उम्मीदवार फोटो लोड की गई: {os.path.basename(path)}")

    def _on_browse_poster(self):
        path = filedialog.askopenfilename(
            title="रेडीमेड पोस्टर फोटो चुनें (Select Poster Image)",
            filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.webp"), ("All Files", "*.*")]
        )
        if path:
            self.custom_poster_data_url = image_to_base64_data_url(path)
            self._log(f"रेडीमेड पोस्टर लोड किया गया: {os.path.basename(path)}")
            messagebox.showinfo("पोस्टर सेट", f"रेडीमेड पोस्टर लोड हो गया है: {os.path.basename(path)}")

    # -------------------------------------------------------------
    # EXCEL FILE LOADING
    # -------------------------------------------------------------
    def _on_browse_file(self):
        path = filedialog.askopenfilename(
            title="वोटर एक्सेल फ़ाइल चुनें (Select Voter Excel)",
            filetypes=[("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")]
        )
        if not path:
            return

        self.current_excel_path = path
        self.file_entry.delete(0, "end")
        self.file_entry.insert(0, path)
        self._load_excel(path)

    def _load_excel(self, path: str):
        self._log(f"फ़ाइल पढ़ी जा रही है: {os.path.basename(path)} ...")
        try:
            parsed = parse_excel_for_parchi(path)
            self.parsed_data = parsed

            # Populate Sheet Combo
            sheets = parsed.get('sheet_names', [])
            combo_values = []
            
            # Add specific sheets with count
            for s in sheets:
                if 'सारांश' in s or 'summary' in s.lower():
                    continue
                combo_values.append(s)

            # Add Batch All Wards option if multiple ward sheets exist
            ward_sheets = [s for s in sheets if 'वार्ड_' in s or 'ward' in s.lower()]
            if len(ward_sheets) > 1:
                combo_values.append("⚡ सभी वार्ड्स की अलग-अलग पीडीएफ (Batch Export All Wards)")

            if not combo_values:
                combo_values = sheets

            self.sheet_combo.configure(values=combo_values)
            
            # Default to first ward sheet if present, else active sheet
            def_sheet = ward_sheets[0] if ward_sheets else parsed.get('active_sheet', combo_values[0])
            self.sheet_combo.set(def_sheet)

            # Auto-set Ward Badge if present in sheet name
            m_ward = re.search(r'\d+', def_sheet)
            if m_ward:
                self.ward_badge_entry.delete(0, "end")
                self.ward_badge_entry.insert(0, m_ward.group(0).zfill(2))

            # Update Stats
            total_v = parsed.get('total_voters', 0)
            act_v = parsed.get('active_count', total_v)
            del_v = parsed.get('deleted_count', 0)
            booth = parsed.get('default_booth_address', '-')
            
            self.stats_label.configure(
                text=f"📊 कुल मतदाता: {total_v} | सक्रिय: {act_v} | विलोपित: {del_v}\n🏢 बूथ: {booth[:50]}"
            )
            self._log(f"सफलता: {len(sheets)} शीट्स मिलीं, कुल {total_v} मतदाता। सक्रिय शीट: '{def_sheet}'")
            self._update_preview()
        except Exception as e:
            self._log(f"❌ फ़ाइल पढ़ने में त्रुटि: {e}")
            messagebox.showerror("त्रुटि", f"एक्सेल फ़ाइल लोड करने में विफल:\n{e}")

    def _on_sheet_changed(self, choice):
        if not self.current_excel_path or not self.parsed_data:
            return
        if "Batch Export" in choice:
            self._log("⚡ 'सभी वार्ड्स बैच एक्सपोर्ट' मोड चुना गया।")
            return

        self._log(f"शीट बदली जा रही है: {choice} ...")
        try:
            parsed = parse_excel_for_parchi(self.current_excel_path, sheet_name=choice)
            self.parsed_data = parsed
            
            m_ward = re.search(r'\d+', choice)
            if m_ward:
                self.ward_badge_entry.delete(0, "end")
                self.ward_badge_entry.insert(0, m_ward.group(0).zfill(2))

            total_v = parsed.get('total_voters', 0)
            act_v = parsed.get('active_count', total_v)
            del_v = parsed.get('deleted_count', 0)
            booth = parsed.get('default_booth_address', '-')

            self.stats_label.configure(
                text=f"📊 कुल मतदाता: {total_v} | सक्रिय: {act_v} | विलोपित: {del_v}\n🏢 बूथ: {booth[:50]}"
            )
            self._update_preview()
            self._log(f"शीट '{choice}' लोड हुई: {total_v} मतदाता।")
        except Exception as e:
            self._log(f"❌ शीट लोड त्रुटि: {e}")

    # -------------------------------------------------------------
    # GENERATION EXECUTION (THREADED)
    # -------------------------------------------------------------
    def _on_start_generation(self):
        if not self.current_excel_path or not self.parsed_data:
            messagebox.showwarning("चेतावनी", "कृपया पहले एक वैध एक्सेल फ़ाइल चुनें।")
            return

        if self.is_generating:
            return

        # Check browser
        if not get_browser_executable():
            messagebox.showerror("त्रुटि", "सिस्टम में Google Chrome या Microsoft Edge नहीं मिला, जो PDF रेंडरिंग हेतु आवश्यक है।")
            return

        choice = self.sheet_combo.get()
        if "Batch Export" in choice:
            threading.Thread(target=self._run_batch_generation, daemon=True).start()
        else:
            threading.Thread(target=self._run_single_generation, daemon=True).start()

    def _gather_candidate_info(self) -> Dict[str, Any]:
        info = {
            'candidate_name': self.cand_name_entry.get().strip() or "संजय कुमार मेवाड़ा",
            'header_prefix': self.header_prefix_entry.get().strip(),
            'ward_badge': self.ward_badge_entry.get().strip() or "01",
            'header_suffix': self.header_suffix_entry.get().strip(),
            'button_no': self.btn_no_entry.get().strip() or "1",
            'slogan': self.slogan_entry.get().strip(),
            'voting_time': self.time_entry.get().strip(),
            'poster_img_data': self.custom_poster_data_url,
            'candidate_photo_data': self.custom_photo_data_url or DEFAULT_CANDIDATE_DATA,
            'jila_parishad': self.parsed_data.get('default_jila_parishad', '') if self.parsed_data else '',
            'panchayat_samiti': self.parsed_data.get('default_panchayat_samiti', '') if self.parsed_data else ''
        }

        # Symbol
        sym_choice = self.symbol_combo.get()
        if self.custom_symbol_data_url:
            info['custom_symbol_data'] = self.custom_symbol_data_url
        elif "कमल" in sym_choice:
            info['symbol_preset'] = 'lotus'
        elif "साइकिल" in sym_choice:
            info['symbol_preset'] = 'cycle'
        else:
            info['symbol_preset'] = 'hand'

        return info

    def _run_single_generation(self):
        self.is_generating = True
        self.generate_btn.configure(state="disabled")
        self.open_pdf_btn.configure(state="disabled")
        self.open_folder_btn.configure(state="disabled")
        self.progress_bar.set(0.0)

        t_start = time.time()
        try:
            voters = self.parsed_data['voters']
            cand_info = self._gather_candidate_info()
            layout_val = self.layout_seg.get()
            if "4" in layout_val:
                layout_type = "4"
            elif "12" in layout_val:
                layout_type = "12"
            else:
                layout_type = "8"
            show_pg = self.page_num_var.get()
            act_only = self.active_only_var.get()

            r_from = None
            r_to = None
            if not self.all_voters_var.get():
                try:
                    r_from = int(self.range_from_entry.get())
                except Exception:
                    r_from = None
                try:
                    r_to = int(self.range_to_entry.get())
                except Exception:
                    r_to = None

            self.status_lbl.configure(text="🚀 कलर वोटर पर्ची PDF जनरेट हो रही है...")
            self._log(f"PDF निर्माण प्रारंभ: कुल {len(voters)} मतदाता | लेआउट: A4-{layout_type} Slips...")

            def on_progress(p_info):
                pct = p_info.get('percent', 0.0) / 100.0
                st = p_info.get('status', '')
                self.progress_bar.set(pct)
                self.status_lbl.configure(text=f"{st} ({int(pct * 100)}%)")
                self._log(st)

            pdf_bytes = generate_color_parchi_pdf(
                voters=voters,
                candidate_info=cand_info,
                layout_type=layout_type,
                show_page_number=show_pg,
                filter_active_only=act_only,
                range_from=r_from,
                range_to=r_to,
                progress_callback=on_progress
            )

            # Output filename
            sheet_clean = re.sub(r'[^\w\-]', '_', self.sheet_combo.get()).strip('_')
            cand_clean = re.sub(r'[^\w\-]', '_', cand_info['candidate_name']).strip('_')
            out_name = f"Color_Voter_Slip_{sheet_clean}_{cand_clean}_A4_{layout_type}.pdf"
            
            out_dir = os.path.dirname(self.current_excel_path)
            out_path = os.path.join(out_dir, out_name)

            with open(out_path, 'wb') as f:
                f.write(pdf_bytes)

            self.last_generated_pdf = out_path
            elapsed = time.time() - t_start

            doc = pymupdf.open(out_path)
            total_pages = len(doc)
            doc.close()

            self.progress_bar.set(1.0)
            self.status_lbl.configure(text=f"✓ PDF तैयार: {total_pages} पृष्ठ ({elapsed:.1f} सेकंड)")
            self._log(f"🎉 सफलता! PDF सहेजी गई: {out_path} ({total_pages} पृष्ठ, {len(pdf_bytes)} बाइट्स)")

            self.open_pdf_btn.configure(state="normal")
            self.open_folder_btn.configure(state="normal")

            messagebox.showinfo(
                "सफल!",
                f"कलर वोटर पर्ची PDF सफलतापूर्वक बन गई है!\n\n"
                f"• कुल पृष्ठ: {total_pages}\n"
                f"• समय: {elapsed:.1f} सेकंड\n"
                f"• फ़ाइल पाथ:\n{out_path}"
            )
        except Exception as e:
            self.status_lbl.configure(text=f"❌ त्रुटि: {e}")
            self._log(f"❌ PDF निर्माण विफल: {e}")
            messagebox.showerror("त्रुटि", f"कलर पर्ची PDF निर्माण में त्रुटि आई:\n{e}")
        finally:
            self.is_generating = False
            self.generate_btn.configure(state="normal")

    def _run_batch_generation(self):
        """Batch exports all individual ward sheets into separate PDFs."""
        self.is_generating = True
        self.generate_btn.configure(state="disabled")
        self.open_pdf_btn.configure(state="disabled")
        self.open_folder_btn.configure(state="disabled")
        self.progress_bar.set(0.0)

        t_start = time.time()
        try:
            wb = pymupdf.open()  # dummy
            wb_excel = parse_excel_for_parchi(self.current_excel_path)
            all_sheets = wb_excel.get('sheet_names', [])
            ward_sheets = [s for s in all_sheets if 'वार्ड_' in s or 'ward' in s.lower()]
            if not ward_sheets:
                ward_sheets = [s for s in all_sheets if 'सारांश' not in s and 'summary' not in s.lower()]

            total_wards = len(ward_sheets)
            self._log(f"⚡ बैच एक्सपोर्ट प्रारंभ: {total_wards} वार्ड्स...")

            out_dir = os.path.dirname(self.current_excel_path)
            cand_info_base = self._gather_candidate_info()
            layout_val = self.layout_seg.get()
            if "4" in layout_val:
                layout_type = "4"
            elif "12" in layout_val:
                layout_type = "12"
            else:
                layout_type = "8"
            show_pg = self.page_num_var.get()
            act_only = self.active_only_var.get()

            created_files = []

            for idx, w_sheet in enumerate(ward_sheets, 1):
                self._log(f"[{idx}/{total_wards}] वार्ड लोड हो रहा है: {w_sheet} ...")
                p_sheet = parse_excel_for_parchi(self.current_excel_path, sheet_name=w_sheet)
                v_list = p_sheet['voters']

                # Update ward badge dynamically
                w_info = dict(cand_info_base)
                m_num = re.search(r'\d+', w_sheet)
                if m_num:
                    w_info['ward_badge'] = m_num.group(0).zfill(2)

                def batch_cb(info):
                    p_pct = ((idx - 1) + (info.get('percent', 0) / 100.0)) / total_wards
                    self.progress_bar.set(p_pct)
                    self.status_lbl.configure(text=f"वार्ड {idx}/{total_wards}: {w_sheet} ({int(p_pct * 100)}%)")

                pdf_bytes = generate_color_parchi_pdf(
                    voters=v_list,
                    candidate_info=w_info,
                    layout_type=layout_type,
                    show_page_number=show_pg,
                    filter_active_only=act_only,
                    progress_callback=batch_cb
                )

                s_clean = re.sub(r'[^\w\-]', '_', w_sheet).strip('_')
                out_name = f"Color_Parchi_{s_clean}_A4_{layout_type}.pdf"
                out_path = os.path.join(out_dir, out_name)

                with open(out_path, 'wb') as f:
                    f.write(pdf_bytes)

                created_files.append(out_path)
                self._log(f"  ✓ पूर्ण: {out_name} ({len(v_list)} मतदाता)")

            self.last_generated_pdf = created_files[0] if created_files else None
            elapsed = time.time() - t_start
            self.progress_bar.set(1.0)
            self.status_lbl.configure(text=f"✓ बैच एक्सपोर्ट पूर्ण: {total_wards} वार्ड्स ({elapsed:.1f}s)")
            self.open_pdf_btn.configure(state="normal")
            self.open_folder_btn.configure(state="normal")

            messagebox.showinfo(
                "बैच एक्सपोर्ट सफल!",
                f"सभी {total_wards} वार्ड्स की कलर वोटर पर्ची PDF बन गई हैं!\n\n"
                f"• कुल समय: {elapsed:.1f} सेकंड\n"
                f"• फ़ोल्डर: {out_dir}"
            )
        except Exception as e:
            self._log(f"❌ बैच एक्सपोर्ट त्रुटि: {e}")
            messagebox.showerror("त्रुटि", f"बैच एक्सपोर्ट विफल:\n{e}")
        finally:
            self.is_generating = False
            self.generate_btn.configure(state="normal")

    def _on_open_pdf(self):
        if self.last_generated_pdf and os.path.exists(self.last_generated_pdf):
            try:
                os.startfile(self.last_generated_pdf)
            except Exception as e:
                messagebox.showerror("त्रुटि", f"फ़ाइल खोलने में असमर्थ: {e}")

    def _on_open_folder(self):
        path = self.last_generated_pdf or self.current_excel_path
        if path:
            folder = os.path.dirname(os.path.abspath(path))
            try:
                os.startfile(folder)
            except Exception as e:
                messagebox.showerror("त्रुटि", f"फ़ोल्डर खोलने में असमर्थ: {e}")

def main():
    app = ColorVoterSlipDesktopApp()
    app.mainloop()

if __name__ == '__main__':
    main()
