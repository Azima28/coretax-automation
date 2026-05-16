import difflib
import tkinter as tk
import customtkinter as ctk
import pandas as pd
import openpyxl
from openpyxl.styles import NumberFormatDescriptor
import pdfplumber
import requests
import base64
import io
import re
import time
import json
import os
import threading

class MappingWindow(ctk.CTkToplevel):
    def __init__(self, parent, all_names, categories):
        super().__init__(parent)
        self.title("Batch Mapping: Tentukan Kategori Barang")
        self.geometry("750x650")
        self.result = {}
        self.all_names = sorted(list(all_names))
        self.categories = ["Abaikan"] + categories
        self.combo_vars = {}
        self.row_frames = []
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        header_frame = ctk.CTkFrame(self)
        header_frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        
        self.label_info = ctk.CTkLabel(header_frame, text=f"Ditemukan {len(self.all_names)} barang unik (MENTAH).\nKlik AUTO-MAP untuk meringkas:", font=("Arial", 12, "bold"))
        self.label_info.pack(side="left", padx=10)
        
        self.btn_auto = ctk.CTkButton(header_frame, text="AUTO-MAP (SMART)", fg_color="#E67E22", hover_color="#D35400", command=self.do_smart_action)
        self.btn_auto.pack(side="right", padx=10)

        self.scroll_frame = ctk.CTkScrollableFrame(self)
        self.scroll_frame.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        self.scroll_frame.grid_columnconfigure(0, weight=1)

        self.render_ui(self.all_names)

        self.btn_save = ctk.CTkButton(self, text="SIMPAN & PROSES SEMUA", fg_color="#27AE60", hover_color="#219150", command=self.save_mapping)
        self.btn_save.grid(row=2, column=0, padx=20, pady=20)

    def render_ui(self, names):
        for f in self.row_frames: f.destroy()
        self.row_frames = []
        self.combo_vars = {}
        for i, name in enumerate(names):
            row_f = ctk.CTkFrame(self.scroll_frame)
            row_f.grid(row=i, column=0, padx=5, pady=5, sticky="ew")
            self.row_frames.append(row_f)
            ctk.CTkLabel(row_f, text=f"Nama : {name}", wraplength=450, justify="left").pack(side="left", padx=10)
            var = ctk.StringVar(value="Abaikan")
            combo = ctk.CTkComboBox(row_f, values=self.categories, variable=var, width=200)
            combo.pack(side="right", padx=10)
            self.combo_vars[name] = var

    def do_smart_action(self):
        # 1. Grouping Logic
        def get_core(name):
            n = name.upper().strip()
            n = re.sub(r'\b(METER KUBIK|METER|KUBIK|UNIT|PCS|KG|LITER|SAK|ZAK|LTR|UNIT|BOX|ROLL|BTG|LBR)\b', ' ', n)
            n = re.sub(r'[\d.,\-()/xX*]+', ' ', n)
            words = n.split()
            if not words: return name
            anchors = ["JASA", "SERVICE", "SP", "SMN", "GT", "WL", "R", "C", "MATERIAL"]
            if words[0] in anchors:
                if words[0] in ["JASA", "SP", "SMN", "MATERIAL"]: return words[0]
                return " ".join(words[:2])
            return " ".join(words[:2])

        grouped = {}
        for n in self.all_names:
            core = get_core(n)
            if core not in grouped: grouped[core] = []
            grouped[core].append(n)
            
        # 2. Re-render UI
        display_list = []
        self.final_group_map = {} 
        for core, originals in grouped.items():
            rep = originals[0]
            count = len(originals)
            d_name = f"{rep} (+{count-1} lainnya)" if count > 1 else rep
            display_list.append(d_name)
            self.final_group_map[d_name] = originals
            
        self.render_ui(display_list)
        self.label_info.configure(text=f"Selesai! Diringkas menjadi {len(display_list)} kelompok.")
        self.auto_map_logic()

    def auto_map_logic(self):
        knowledge = {
            "SEMEN": ["SMN", "PCC", "MU", "CEMENT"],
            "BBM": ["SOLAR", "HSD", "DEX", "PERTA", "FUEL"],
            "SPAREPART": ["SP", "HOSE", "BELT", "SEAL", "FILTER", "BEARING", "GEAR", "BOLT", "NUT", "TYRE", "BAN"],
            "ASPAL": ["ASPHALT", "TACK", "PRIME", "HOTMIX"],
            "PELUMAS": ["OLI", "OIL", "LUBE", "GREASE"],
            "BIAYA": ["SERVICE", "REPAIR", "JASA", "MAINTENANCE"],
            "ANGKUT": ["TRANSPORT", "LOGISTIK", "EXPEDISI"]
        }
        for d_name, var in self.combo_vars.items():
            u_name = d_name.upper()
            best_cat = "Abaikan"; max_score = 0
            for cat in self.categories[1:]:
                u_cat = cat.upper()
                for main_key, synonyms in knowledge.items():
                    if main_key in u_cat or any(s in u_cat for s in synonyms):
                        if any(s in u_name for s in synonyms) or main_key in u_name:
                            max_score = 1.0; best_cat = cat; break
                if max_score == 1.0: break
                score = difflib.SequenceMatcher(None, u_name, u_cat).ratio()
                if score > max_score: max_score = score; best_cat = cat
            if max_score >= 0.7: var.set(best_cat)

    def save_mapping(self):
        if hasattr(self, 'final_group_map'):
            for d_name, var in self.combo_vars.items():
                for orig in self.final_group_map[d_name]:
                    self.result[orig] = var.get()
        else:
            for name, var in self.combo_vars.items():
                self.result[name] = var.get()
        self.destroy()

class CoreTaxApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("CoreTax Auto-Monitor Premium v1.3")
        self.geometry("900x700")
        ctk.set_appearance_mode("dark")
        
        # Variables
        self.file_path = ""
        self.token = ""
        self.cookie = ""
        self.dynamic_categories = {} # {Kategori: (Index_Harga, Index_QTY)}
        
        # UI
        self.setup_ui()
        
    def setup_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        
        ctk.CTkLabel(header_frame, text="CoreTax AI v1.3", font=("Arial", 28, "bold")).pack(side="left")
        self.status_label = ctk.CTkLabel(header_frame, text="Status: Terhubung (TID: 275bb07a...)", text_color="green")
        self.status_label.pack(side="right")
        
        # Session Inputs
        session_frame = ctk.CTkFrame(self)
        session_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        session_frame.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(session_frame, text="Bearer Token:").grid(row=0, column=0, padx=10, pady=5)
        self.entry_token = ctk.CTkEntry(session_frame, placeholder_text="Masukkan Bearer Token...")
        self.entry_token.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        
        ctk.CTkLabel(session_frame, text="Cookie:").grid(row=1, column=0, padx=10, pady=5)
        self.entry_cookie = ctk.CTkEntry(session_frame, placeholder_text="Masukkan Cookie (id-ID)...")
        self.entry_cookie.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        
        ctk.CTkLabel(session_frame, text="TID (ID):").grid(row=2, column=0, padx=10, pady=5)
        self.entry_tid = ctk.CTkEntry(session_frame, placeholder_text="Masukkan Taxpayer ID...")
        self.entry_tid.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        
        btn_action_frame = ctk.CTkFrame(session_frame, fg_color="transparent")
        btn_action_frame.grid(row=0, column=2, rowspan=3, padx=10, pady=5, sticky="ns")

        self.btn_login = ctk.CTkButton(btn_action_frame, text="LOGIN CORETAX", command=self.start_login, fg_color="blue", hover_color="darkblue")
        self.btn_login.pack(expand=True, fill="both", pady=2)
        
        ctk.CTkButton(btn_action_frame, text="SIMPAN SESSION", command=self.save_session, fg_color="green").pack(expand=True, fill="both", pady=2)

        # Excel Inputs
        input_frame = ctk.CTkFrame(self)
        input_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        input_frame.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(input_frame, text="Target Excel:").grid(row=0, column=0, padx=10, pady=10)
        self.entry_file = ctk.CTkEntry(input_frame, placeholder_text="Pilih file monitoring...")
        self.entry_file.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(input_frame, text="BROWSE", width=100, command=self.browse_file).grid(row=0, column=2, padx=10, pady=10)

        ctk.CTkLabel(input_frame, text="MODE PROSES:").grid(row=1, column=0, padx=10, pady=5)
        self.mode_var = ctk.StringVar(value="AUTO-PROCESS (Clear & Fill)")
        self.mode_menu = ctk.CTkOptionMenu(input_frame, values=["AUTO-PROCESS (Clear & Fill)", "ONLY-CLEAR (Hapus Saja)", "CHECK-ONLY (PDF Saja)"], variable=self.mode_var)
        self.mode_menu.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        # Row 2: QUICK CHECK
        ctk.CTkLabel(input_frame, text="CEK 1 FAKTUR:").grid(row=2, column=0, padx=10, pady=5)
        self.entry_single_faktur = ctk.CTkEntry(input_frame, placeholder_text="Masukkan No Faktur...")
        self.entry_single_faktur.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        self.btn_quick_check = ctk.CTkButton(input_frame, text="QUICK CHECK", width=100, command=self.quick_check, fg_color="orange", text_color="black")
        self.btn_quick_check.grid(row=2, column=2, padx=10, pady=5)
        
        self.log_box = ctk.CTkTextbox(self, font=("Consolas", 12))
        self.log_box.grid(row=3, column=0, padx=20, pady=10, sticky="nsew")
        
        self.btn_run = ctk.CTkButton(self, text="START PROCESS", height=50, command=self.start_process, font=("Arial", 16, "bold"))
        self.btn_run.grid(row=4, column=0, padx=20, pady=20, sticky="ew")
        
        # Default values
        self.entry_cookie.insert(0, "id-ID")
        self.entry_tid.insert(0, "275bb07a-d021-4389-943e-a740246a56e8")

    def start_login(self):
        self.btn_login.configure(state="disabled", text="BROWSER OPEN...")
        threading.Thread(target=self.playwright_login, daemon=True).start()

    def playwright_login(self):
        from playwright.sync_api import sync_playwright
        self.add_log("[*] Membuka browser untuk login...")
        try:
            with sync_playwright() as p:
                # Tambahkan user agent asli agar tidak diblokir
                user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
                browser = p.chromium.launch(headless=False)
                context = browser.new_context(user_agent=user_agent)
                page = context.new_page()
                
                found_token = False
                
                def handle_request(request):
                    nonlocal found_token
                    auth = request.headers.get("authorization")
                    if auth and "Bearer" in auth and len(auth) > 200:
                        token_val = auth.replace("Bearer ", "")
                        self.token = token_val
                        self.after(0, lambda: self.entry_token.delete(0, "end"))
                        self.after(0, lambda: self.entry_token.insert(0, token_val))
                        if not found_token:
                            self.add_log("[√] SESSION DITANGKAP!")
                            found_token = True

                page.on("request", handle_request)
                
                # Gunakan URL Login yang Anda berikan
                self.add_log("[*] Membuka halaman login...")
                page.goto("https://coretaxdjp.pajak.go.id/identityproviderportal/Account/Login", wait_until="domcontentloaded", timeout=60000)
                
                while True:
                    if not browser.is_connected(): break
                    
                    # Capture Cookies
                    cookies = context.cookies()
                    cookie_str = "; ".join([f"{c['name']}={c['value']}" for c in cookies])
                    if "sl-session" in cookie_str:
                        self.cookie = cookie_str
                        self.after(0, lambda: self.entry_cookie.delete(0, "end"))
                        self.after(0, lambda: self.entry_cookie.insert(0, self.cookie))
                    
                    # Extract TID from token
                    if hasattr(self, 'token') and self.token:
                        try:
                            payload_b64 = self.token.split('.')[1]
                            missing_padding = len(payload_b64) % 4
                            if missing_padding: payload_b64 += '=' * (4 - missing_padding)
                            payload = json.loads(base64.b64decode(payload_b64).decode('utf-8'))
                            tid = payload.get("taxpayer_id")
                            if tid:
                                self.tid = tid
                                self.after(0, lambda: self.entry_tid.delete(0, "end"))
                                self.after(0, lambda: self.entry_tid.insert(0, tid))
                        except: pass
                    
                    time.sleep(1)
        except Exception as e:
            self.add_log(f"[!] Browser Error: {e}")
        finally:
            self.after(0, lambda: self.btn_login.configure(state="normal", text="LOGIN CORETAX"))
            self.add_log("[*] Browser ditutup.")

    def save_session(self):
        self.token = self.entry_token.get()
        self.cookie = self.entry_cookie.get()
        self.tid = self.entry_tid.get()
        self.add_log("[√] SESSION DIPERBARUI & DISIMPAN!")

    def add_log(self, msg):
        self.log_box.insert("end", f"{time.strftime('%H:%M:%S')} | {msg}\n")
        self.log_box.see("end")

    def browse_file(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.file_path = path
            self.entry_file.delete(0, "end")
            self.entry_file.insert(0, path)

    def auto_detect_header(self, file_path, sheet_name='PM'):
        try:
            df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=10)
            for i, row in df_raw.iterrows():
                row_str = " ".join([str(c) for c in row.values]).upper()
                if "NOMOR FAKTUR" in row_str or "HARGA JUAL" in row_str:
                    return i
            return 0
        except: return 0

    def _get_col_idx(self, ws, name):
        for cell in ws[1]:
            if name.upper() in str(cell.value).upper():
                return cell.column
        return None

    def quick_check(self):
        faktur = self.entry_single_faktur.get().strip()
        if not faktur:
            self.add_log("[!] Masukkan nomor faktur.")
            return
        
        self.token = self.entry_token.get().strip()
        self.cookie = self.entry_cookie.get().strip()
        tid = self.entry_tid.get().strip()
        
        if not self.token or not self.cookie:
            self.add_log("[!] Sesi belum ada. Silakan Login dulu.")
            return

        def run():
            self.add_log(f"\n[QUICK CHECK] Memproses Faktur: {faktur}...")
            # 1. RAW DATA (Cari Faktur)
            
            # 2. SYSTEM PARSING
            self.add_log("[*] Mengekstrak isi PDF...")
            items, msg, inv = self._fetch_pdf_items(faktur, tid, 0)
            
            if items:
                self.add_log(f"--- [RAW DATA DARI SERVER] ---")
                self.add_log(json.dumps(inv, indent=2))
                self.add_log(f"--- [HASIL PARSING SISTEM] ---")
                for it in items:
                    self.add_log(f" > {it['name']} | Qty: {it['qty']} | Total: {it['total']:,}")
            else:
                if inv:
                    self.add_log(f"--- [RAW DATA DARI SERVER] ---")
                    self.add_log(json.dumps(inv, indent=2))
                self.add_log(f"[!] Error Quick Check: {msg}")

        threading.Thread(target=run, daemon=True).start()

    def start_process(self):
        if not self.file_path:
            self.add_log("[!] Pilih file Excel dulu.")
            return
        self.btn_run.configure(state="disabled", text="PROCESSING...")
        threading.Thread(target=self.run_logic, daemon=True).start()

    def run_logic(self):
        try:
            self.btn_run.configure(state="disabled", text="PROCESSING...")
            
            tid = self.entry_tid.get().strip()
            self.token = self.entry_token.get().strip()
            self.cookie = self.entry_cookie.get().strip()
            
            if not tid or not self.token or not self.cookie:
                self.add_log("[!] Gagal mendapatkan Taxpayer ID / Sesi. Pastikan sudah login.")
                self.btn_run.configure(state="normal", text="START PROCESS")
                return
                
            wb = openpyxl.load_workbook(self.file_path)
            
            # 1. CARI SHEET YANG RELEVAN (Prioritaskan PM jika ada)
            target_sheet_name = None
            # Cek dulu apakah ada sheet dengan nama mengandung 'PM'
            pm_sheets = [sn for sn in wb.sheetnames if "PM" in sn.upper()]
            other_sheets = [sn for sn in wb.sheetnames if "PM" not in sn.upper()]
            
            for sn in (pm_sheets + other_sheets): # Prioritaskan PM
                ws_temp = wb[sn]
                found_header = False
                for r in range(1, 15):
                    try:
                        row_vals = [str(cell.value).upper() if cell.value else "" for cell in ws_temp[r]]
                        if any("NOMOR FAKTUR" in v for v in row_vals):
                            target_sheet_name = sn
                            found_header = True
                            break
                    except: pass
                if found_header: break
            
            if not target_sheet_name:
                self.add_log("[!] ERROR: Tidak menemukan kolom 'Nomor Faktur' di sheet manapun.")
                self.btn_run.configure(state="normal", text="START PROCESS")
                return
            
            self.add_log(f"[*] Memproses Sheet: {target_sheet_name}")
            ws = wb[target_sheet_name]
            
            # Konversi ke DataFrame untuk kemudahan filter
            data_raw = list(ws.values)
            # Cari baris header yang benar
            h_idx = 0
            for i, row in enumerate(data_raw):
                row_str = [str(v).upper() if v else "" for v in row]
                if any("NOMOR FAKTUR" in v for v in row_str):
                    h_idx = i
                    break
            
            cols = data_raw[h_idx]
            df = pd.DataFrame(data_raw[h_idx+1:], columns=cols)
            # Bersihkan nama kolom dari spasi tersembunyi
            df.columns = [str(c).strip() if c else f"COL_{i}" for i, c in enumerate(df.columns)]
            headers = list(df.columns)

            self.dynamic_categories = {}
            blacklist = ["TOTAL", "DPP", "PPN", "JUMLAH PENJABARAN", "SELISIH", "QTY", "KET"]
            
            for i, col in enumerate(headers):
                col_str = str(col).strip().upper()
                if "QTY" in col_str:
                    if i + 1 < len(headers):
                        raw_cat = str(headers[i+1]).strip()
                        is_blacklisted = any(b in raw_cat.upper() for b in blacklist)
                        if not is_blacklisted and raw_cat:
                            self.dynamic_categories[raw_cat] = (i+2, i+1) # (HargaCol, QtyCol)
            
            self.add_log(f"[+] Kategori terdeteksi: {', '.join(self.dynamic_categories.keys())}")
            
            # Ambil index kolom penting
            c_faktur = next((c for c in headers if "NOMOR FAKTUR" in str(c).upper()), None)
            c_harga  = next((c for c in headers if "HARGA JUAL" in str(c).upper()), None)
            c_selisih = next((c for c in headers if "SELISIH" in str(c).upper()), None)
            
            if not c_faktur or not c_harga:
                self.add_log("[!] ERROR: Kolom 'Nomor Faktur' atau 'Harga Jual' tidak ditemukan.")
                self.btn_run.configure(state="normal", text="START PROCESS")
                return
            
            c_penjabaran = next((c for c in headers if "PENJABARAN" in str(c).upper()), None)

            def is_not_positive(val):
                if val is None: return True
                s_val = str(val).strip()
                if s_val == "" or s_val == "0" or s_val == "0.0": return True
                # Jika diawali '=' berarti RUMUS, kita anggap "Belum Ada Isinya" 
                # (nanti akan dicek lebih lanjut via has_any_real_entry)
                if s_val.startswith("="): return True
                try:
                    f_val = float(s_val.replace(',',''))
                    if f_val <= 0: return True
                except: pass
                return False

            def should_process(r):
                if pd.isna(r[c_faktur]): return False
                
                def is_truly_empty(v):
                    if v is None or pd.isna(v): return True
                    s_v = str(v).strip().lower()
                    if s_v == "" or s_v == "0" or s_v == "0.0" or s_v == "nan": return True
                    return False

                # 1. Cek Kategori Secara Mendalam
                has_any_real_entry = False
                for cat in self.dynamic_categories.keys():
                    if cat in r:
                        val = r[cat]
                        if not is_truly_empty(val) and not str(val).startswith("="):
                            has_any_real_entry = True
                            break

                # 2. ATURAN EMAS: Cek Kolom Penjabaran (AW)
                pj = r[c_penjabaran] if c_penjabaran else None
                
                # Jika Penjabaran KOSONG atau <= 0 atau RUMUS (tanpa isi kategori) -> PROSES
                if is_not_positive(pj):
                    if str(pj).startswith("=") and has_any_real_entry:
                        return False
                    return True
                
                return False
            
            targets = df[df.apply(should_process, axis=1)].copy()
            targets['excel_row'] = targets.index + h_idx + 2
            self.add_log(f"[+] Ditemukan {len(targets)} faktur untuk diproses.")
            
            if len(targets) == 0:
                self.add_log("[!] Tidak ada faktur yang perlu diproses. Selesai.")
                self.btn_run.configure(state="normal", text="START PROCESS")
                return

            # TAHAP 1: HAPUS SEMUA TARGET DI EXCEL DULU
            self.add_log(f"[*] Tahap 1: Membersihkan {len(targets)} baris di Excel & Saving...")
            wb_clear = openpyxl.load_workbook(self.file_path)
            ws_clear = wb_clear[target_sheet_name]  # Dinamis: gunakan sheet yang terdeteksi
            for _, row in targets.iterrows():
                ex_row = int(row['excel_row'])
                for cat_name, (p_idx, q_idx) in self.dynamic_categories.items():
                    ws_clear.cell(row=ex_row, column=p_idx).value = None
                    ws_clear.cell(row=ex_row, column=q_idx).value = None
            wb_clear.save(self.file_path)
            self.add_log("[√] Seluruh baris target sudah BERSIH.")

            if self.mode_var.get() == "ONLY-CLEAR (Hapus Saja)":
                self.btn_run.configure(state="normal", text="START PROCESS")
                return

            # TAHAP 2: GLOBAL SCAN (Semua Faktur)
            all_invoice_data = {} # {ex_row: [items]}
            all_unique_names = set()
            
            self.add_log("[*] Tahap 2: Men-scan SELURUH PDF (Global Scan)...")
            for index, row in targets.iterrows():
                raw_f = row[c_faktur]
                expected_total = pd.to_numeric(row[c_harga], errors='coerce')
                ex_row = int(row['excel_row'])
                faktur = "{:.0f}".format(raw_f).zfill(17) if isinstance(raw_f, (float, int)) else str(raw_f).zfill(17)
                
                self.add_log(f"    - Scanning PDF Baris {ex_row}: {faktur}")
                items, msg, raw_json = self._fetch_pdf_items(faktur, tid, expected_total)
                if items:
                    self.add_log(f"--- [RAW DATA DARI SERVER: {faktur}] ---")
                    self.add_log(json.dumps(raw_json, indent=2))
                    self.add_log(f"--- [HASIL PARSING SISTEM] ---")
                    all_invoice_data[ex_row] = items
                    for it in items: 
                        all_unique_names.add(it['name'])
                        self.add_log(f"      > Item: {it['name']} | Qty: {it['qty']} | Total: {it['total']:,}")
                else:
                    self.add_log(f"      [!] Skip: {msg}")

            if not all_invoice_data:
                self.add_log("[!] Scan selesai, tidak ada data yang ditemukan.")
                self.btn_run.configure(state="normal", text="START PROCESS")
                return

            # TAHAP 3: GLOBAL MAPPING (Smart Grouping / Reduce)
            self.add_log(f"[*] Tahap 3: Meringkas {len(all_unique_names)} barang menjadi kelompok unik...")
            
            # TAHAP 3: GLOBAL MAPPING (Smart Grouping / Reduce)
            self.add_log(f"[*] Tahap 3: Meringkas {len(all_unique_names)} barang menjadi kelompok unik...")
            
            # Algoritma Pembersihan Agresif (Anchor Grouping)
            def get_core_identity(name):
                n = str(name).upper().strip()
                # 1. Hapus Satuan & Angka Dulu
                n = re.sub(r'\b(METER KUBIK|METER|KUBIK|UNIT|PCS|KG|LITER|SAK|ZAK|LTR|UNIT|BOX|ROLL|BTG|LBR|CURAH|JB)\b', ' ', n)
                n = re.sub(r'[\d.,\-()/xX*]+', ' ', n)
                words = n.split()
                if not words: return name.upper()
                
                # 2. Logika Anchor (Jika diawali kata kunci, ambil depannya saja)
                anchors = ["JASA", "SERVICE", "SP", "SMN", "GT", "WL", "R", "C", "MATERIAL", "ULTRAPRO", "SOLAR", "BIOSOLAR", "SPLIT", "SCREENING", "ABU"]
                first_word = words[0]
                
                if first_word in anchors:
                    # Khusus kata kunci utama, kita ambil 1 kata saja agar gabung semua
                    if first_word in ["JASA", "SP", "SMN", "MATERIAL", "ULTRAPRO", "SOLAR", "BIOSOLAR", "ABU"]:
                        return first_word
                    # Untuk yang lain ambil 2 kata agar tidak terlalu umum
                    return " ".join(words[:2])
                
                # Jika tidak ada anchor, ambil 2 kata pertama sebagai identitas
                return " ".join(words[:2])

            # Buat Map: {original_name: core_name}
            name_to_core = {name: get_core_identity(name) for name in all_unique_names}
            unique_cores = sorted(list(set(name_to_core.values())))
            
            mapping_window = MappingWindow(self, unique_cores, list(self.dynamic_categories.keys()))
            self.wait_window(mapping_window)
            core_mapping = mapping_window.result # {core_name: category}
            
            if not core_mapping:
                self.add_log("[!] Mapping dibatalkan.")
                self.btn_run.configure(state="normal", text="START PROCESS")
                return

            # TAHAP 4: ISI DATA DENGAN AKUMULASI
            self.add_log("[*] Tahap 4: Mengisi data dengan Logika Akumulasi...")
            wb = openpyxl.load_workbook(self.file_path)
            ws = wb[target_sheet_name]
            success_count = 0
            
            for ex_row, items in all_invoice_data.items():
                self.add_log(f"   -> Memproses Baris {ex_row}...")
                
                # Akumulasi berdasarkan kategori hasil mapping
                category_totals = {} # {Category: {'qty': 0, 'total': 0}}
                
                for it in items:
                    orig_name = it['name']
                    core_name = name_to_core.get(orig_name)
                    cat = core_mapping.get(core_name)
                    
                    if cat and cat != "Abaikan":
                        if cat not in category_totals: category_totals[cat] = {'qty': 0, 'total': 0}
                        category_totals[cat]['qty'] += it['qty']
                        category_totals[cat]['total'] += it['total']
                
                # Tulis hasil akumulasi ke Excel
                for cat, val in category_totals.items():
                    if cat in self.dynamic_categories:
                        p_idx, q_idx = self.dynamic_categories[cat]
                        ws.cell(row=ex_row, column=p_idx).value = val['total']
                        ws.cell(row=ex_row, column=p_idx).number_format = '#,##0'
                        ws.cell(row=ex_row, column=q_idx).value = val['qty']
                        ws.cell(row=ex_row, column=q_idx).number_format = '#,##0.00'
                        self.add_log(f"      [√] {cat}: Total Harga={val['total']:,}, Total Qty={val['qty']}")

                # Update Selisih
                diff_idx = self._get_col_idx(ws, "Selisih")
                if diff_idx: ws.cell(row=ex_row, column=diff_idx).value = "-"
                success_count += 1

            wb.save(self.file_path)
            self.add_log(f"[√] SELESAI: {success_count} faktur berhasil diproses secara batch.")
            self.btn_run.configure(state="normal", text="START PROCESS")
            
        except Exception as e:
            self.add_log(f"[!] ERROR: {str(e)}")
            self.btn_run.configure(state="normal", text="START AUTO-CHECK")

    def _fetch_pdf_items(self, no_faktur, tid, expected_total):
        headers = {"authority": "coretaxdjp.pajak.go.id", "authorization": f"Bearer {self.token}", "content-type": "application/json", "cookie": self.cookie, "x-dgt-code": "7AcAAA=="}
        try:
            s_url = "https://coretaxdjp.pajak.go.id/einvoiceportal/api/inputinvoice/list"
            payload = {"BuyerTaxpayerAggregateIdentifier": tid, "TaxpayerAggregateIdentifier": tid, "First": 0, "Rows": 1, "LanguageId": "id-ID", "Filters": [{"PropertyName": "TaxInvoiceNumber", "Value": no_faktur, "MatchMode": "equals"}]}
            resp = requests.post(s_url, headers=headers, json=payload, timeout=15).json()
            data = resp.get("Payload", {}).get("Data", [])
            if not data: return None, "Faktur tidak ditemukan di server", None

            inv = data[0]
            d_url = "https://coretaxdjp.pajak.go.id/einvoiceportal/api/DownloadInvoice/download-invoice-document"
            d_payload = {"EInvoiceRecordIdentifier": inv["RecordId"], "EInvoiceAggregateIdentifier": inv["AggregateIdentifier"], "DocumentAggregateIdentifier": inv["DocumentFormAggregateIdentifier"], "TaxpayerAggregateIdentifier": tid, "LetterNumber": no_faktur, "EInvoiceMenuType": "Input", "TaxInvoiceStatus": "APPROVED"}
            d_resp = requests.post(d_url, headers=headers, json=d_payload, timeout=15).json()
            pdf_b64 = d_resp.get("Content")
            if not pdf_b64: return None, "Data PDF kosong dari server", inv

            pdf_bytes = base64.b64decode(pdf_b64)
            extracted_items = []
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if not table: continue
                    for row in table:
                        if not row or len(row) < 3: continue

                        # 1. CEK STRUKTUR: Kolom pertama harus Angka Urut (1, 2, 3...)
                        first_col = str(row[0]).strip()
                        if not first_col.isdigit(): continue

                        # 2. EKSTRAKSI NAMA (Kolom ke-3, Index 2)
                        raw_name_cell = str(row[2]) if row[2] else ""
                        # Bersihkan Nama dari teks rumus harga (seperti Rp 185 x 40.000)
                        clean_name = re.split(r"Rp|\sx\s|\n", raw_name_cell)[0].strip()
                        # NORMALISASI SPASI: Hapus spasi ganda agar nama yang mirip jadi satu
                        clean_name = re.sub(r"\s+", " ", clean_name).strip()

                        # 4. EKSTRAKSI DATA (Logika Perkalian Matematika)
                        all_nums = []
                        all_cells_text = " ".join([str(c) for c in row[1:] if c])
                        found_nums_raw = re.findall(r"[\d.]+(?:,[\d]+)?", all_cells_text)
                        for n in found_nums_raw:
                            try:
                                val = float(n.replace(".", "").replace(",", "."))
                                if val > 0: all_nums.append(val)
                            except: pass

                        if not all_nums: continue

                        # Harga Total Baris: Pasti angka TERBESAR
                        price = max(all_nums)

                        # Qty: Cari pasangan A * B = Price (Logika Matematika 100% Akurat)
                        qty = 1.0
                        potential_candidates = [n for n in all_nums if n != price]

                        found_math_match = False
                        if len(potential_candidates) >= 2:
                            for i, a in enumerate(potential_candidates):
                                for b in potential_candidates[i:]:
                                    # Cek apakah a * b mendekati price (toleransi 1% untuk pembulatan pajak)
                                    if abs((a * b) - price) < (0.01 * price):
                                        # Ambil yang terkecil sebagai Qty
                                        qty = a if a < b else b
                                        found_math_match = True
                                        break
                                if found_math_match: break

                        # Fallback jika tidak ada perkalian yang match (misal Qty tersembunyi/tunggal)
                        if not found_math_match:
                            q_match = re.search(r"([\d.,]+)\s*x", all_cells_text)
                            if q_match:
                                qty = float(q_match.group(1).replace(".", "").replace(",", "."))
                            else:
                                # Jika ada angka selain Harga, ambil yang paling kecil (asumsi Qty)
                                if potential_candidates:
                                    qty = min(potential_candidates)
                                else:
                                    qty = 1.0

                        extracted_items.append({"name": clean_name, "qty": qty, "total": price})

            if not extracted_items:
                return None, "Tidak ada item barang bernomor ditemukan di PDF", inv

            # 5. VALIDASI AKURASI (Jika bukan Quick Check)
            if expected_total > 0:
                sum_pdf = sum(it["total"] for it in extracted_items)
                if abs(sum_pdf - expected_total) > 500: # Toleransi 500 rupiah
                    return None, f"Total PDF ({sum_pdf:,}) tidak sinkron dengan Excel ({expected_total:,})", inv

            return extracted_items, "Success", inv
        except Exception as e: return None, str(e), None

if __name__ == "__main__":
    app = CoreTaxApp()
    app.mainloop()
