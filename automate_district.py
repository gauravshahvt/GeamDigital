import os
import sys
import time

# Ensure UTF-8 output on Windows consoles
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# Add project root to sys.path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from pdf_engine.district_automator import scan_district_hierarchy, run_district_automation

def choose_folder_gui() -> str:
    """Tries to show a native Windows Browse Folder dialog."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        folder = filedialog.askdirectory(title="ज़िले का मुख्य फ़ोल्डर चुनें (Select District Folder)")
        root.destroy()
        return folder.strip() if folder else ""
    except Exception:
        return ""

def main():
    print("=" * 80)
    print("🏛️  GEAM DIGITAL - DISTRICT AUTOMATED VOTER EXCEL STUDIO (2026)")
    print("    सम्पूर्ण ज़िला: पंचायत समिति ➔ ग्राम पंचायत ➔ सभी वार्ड्स एक्सेल ऑटोमेशन")
    print("=" * 80)
    
    # 1. Get folder path
    folder_path = ""
    if len(sys.argv) > 1:
        folder_path = sys.argv[1].strip('\"\' ')
    else:
        print("\n📂 फ़ोल्डर चयन विंडो खोली जा रही है...")
        folder_path = choose_folder_gui()
        
    if not folder_path:
        folder_path = input("\n👉 कृपया ज़िले/पंचायत समिति का फ़ोल्डर पाथ पेस्ट या टाइप करें:\n   > ").strip('\"\' ')
        
    if not folder_path or not os.path.exists(folder_path):
        print(f"\n❌ त्रुटि: फ़ोल्डर मौजूद नहीं है या अमान्य पाथ है: {folder_path}")
        input("\nEnter दबाकर बाहर निकलें...")
        return
        
    print(f"\n🔍 फ़ोल्डर स्कैन किया जा रहा है: {folder_path} ...")
    try:
        scan_info = scan_district_hierarchy(folder_path)
    except Exception as e:
        print(f"❌ स्कैनिंग में त्रुटि: {e}")
        input("\nEnter दबाकर बाहर निकलें...")
        return
        
    print("\n" + "-" * 80)
    print(f"📊 स्कैन परिणाम:")
    print(f"   • मुख्य फ़ोल्डर / ज़िला : {scan_info['district_name']}")
    print(f"   • कुल पंचायत समितियां : {scan_info['total_samitis']}")
    print(f"   • कुल ग्राम पंचायतें    : {scan_info['total_panchayats']}")
    print(f"   • कुल वार्ड पीडीएफ्स  : {scan_info['total_pdfs']}")
    print(f"   • पूर्व निर्मित एक्सेल : {scan_info['already_processed_count']}")
    print(f"   • शेष (Pending) कार्य : {scan_info['pending_count']}")
    print("-" * 80)
    
    if scan_info['total_panchayats'] == 0:
        print("⚠️ इस फ़ोल्डर या इसके सब-फ़ोल्डरों में कोई वार्ड पीडीएफ फाइलें (.pdf) नहीं मिलीं।")
        input("\nEnter दबाकर बाहर निकलें...")
        return
        
    print("\nविस्तृत संरचना:")
    for samiti, p_list in scan_info['samitis'].items():
        print(f"   📁 पंचायत समिति: {samiti} ({len(p_list)} पंचायतें)")
        for p in p_list[:5]:
            status_txt = "✓ [Excel पहले से बना है]" if p['has_excel'] else "⏳ [Pending]"
            print(f"      └─ {p['panchayat_name']}: {p['pdf_count']} वार्ड्स {status_txt}")
        if len(p_list) > 5:
            print(f"      └─ ... एवं {len(p_list) - 5} अन्य पंचायतें")
            
    print("\n" + "=" * 80)
    print("विकल्प चुनें:")
    print("  [1] ऑटोमेशन प्रारंभ करें (पूर्व निर्मित फाइलों को छोड़ें / Skip Existing) - अनुशंसित")
    print("  [2] सभी पंचायतों का पुनः एक्सेल बनाएं (पुनर्निर्माण / Overwrite All)")
    print("  [Q] बाहर निकलें (Cancel)")
    print("=" * 80)
    
    choice = input("\n👉 विकल्प दर्ज करें (1, 2 या Q) [डिफ़ॉल्ट 1]: ").strip().upper()
    if choice == 'Q':
        print("ऑटोमेशन रद्द किया गया।")
        return
        
    skip_existing = (choice != '2')
    
    print("\n🚀 ज़िला ऑटोमेशन प्रारंभ हो रहा है...\n")
    
    def on_cli_progress(event):
        e_type = event.get('type')
        idx = event.get('index', 0)
        tot = event.get('total_panchayats', 1)
        pct = event.get('percent', 0.0)
        
        if e_type == 'start_panchayat':
            samiti = event.get('samiti', '')
            panchayat = event.get('panchayat', '')
            cnt = event.get('pdf_count', 0)
            print(f"[{idx}/{tot}] ({pct}%) ⏳ {samiti} ➔ {panchayat} ({cnt} वार्ड्स प्रोसेस हो रहे हैं)...")
        elif e_type == 'ward_progress':
            w_idx = event.get('ward_idx', 1)
            t_w = event.get('total_wards', 1)
            print(f"         ├─ वार्ड {w_idx}/{t_w} एक्सट्रैक्ट किया जा रहा है...", end='\r')
        elif e_type == 'complete_panchayat':
            p_name = event.get('panchayat', '')
            v_cnt = event.get('active_voters', 0)
            f_name = event.get('excel_file', '')
            print(f"[{idx}/{tot}] ({pct}%) ✓ पूर्ण: {p_name} ➔ {v_cnt:,} सक्रिय मतदाता | फ़ाइल: {f_name}")
        elif e_type == 'skip_panchayat':
            p_name = event.get('panchayat', '')
            print(f"[{idx}/{tot}] ({pct}%) ⏩ स्किप: {p_name} (एक्सेल फाइल पहले से मौजूद)")
        elif e_type == 'error_panchayat':
            p_name = event.get('panchayat', '')
            err = event.get('error', '')
            print(f"[{idx}/{tot}] ({pct}%) ❌ त्रुटि: {p_name} ➔ {err}")

    res = run_district_automation(
        root_dir=folder_path,
        skip_existing=skip_existing,
        theme='geam_digital',
        progress_callback=on_cli_progress
    )
    
    mins = int(res['elapsed_seconds'] // 60)
    secs = int(res['elapsed_seconds'] % 60)
    
    print("\n" + "=" * 80)
    print("🎉 ज़िला ऑटोमेशन सफलतापूर्वक संपन्न!")
    print("=" * 80)
    print(f"• कुल पंचायतें प्रोसेस हुईं  : {res['completed_panchayats']} / {res['total_panchayats']}")
    print(f"• स्किप की गईं पंचायतें     : {res['skipped_panchayats']}")
    print(f"• असफल पंचायतें           : {res['failed_panchayats']}")
    print(f"• कुल सक्रिय मतदाता एक्सट्रैक्ट : {res['total_voters']:,}")
    print(f"• कुल समय लगा             : {mins} मिनट {secs} सेकंड")
    print(f"• ज़िला मास्टर सारांश रिपोर्ट : {res['summary_file']}")
    print("\nसभी एक्सेल फाइलें संबंधित पंचायत फ़ोल्डर में सहेजी (Save) जा चुकी हैं।")
    print("=" * 80)
    input("\nEnter दबाकर समाप्त करें...")

if __name__ == '__main__':
    main()
