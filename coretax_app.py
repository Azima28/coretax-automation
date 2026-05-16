import os
import requests
import pandas as pd
import openpyxl
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import time
import re
import io
import base64
import pdfplumber
import json

# --- CONFIG ---
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class MappingWindow(ctk.CTkToplevel):
    def __init__(self, parent, unique_names, categories, initial_mapping=None):
        super().__init__(parent)
        self.title("MAPPING BARANG KE KOLOM EXCEL")
        self.geometry("600x500")
        self.result = {}
        self.categories = ["Abaikan"] + categories
        
        ctk.CTkLabel(self, text="Petakan nama barang dari PDF ke kolom Excel yang sesuai:", font=("Arial", 14, "bold")).pack(pady=10)
        
        scroll_frame = ctk.CTkScrollableFrame(self, width=550, height=350)
        scroll_frame.pack(padx=10, pady=10, fill="both", expand=True)
        
        self.menus = {}
        for name in unique_names:
            row = ctk.CTkFrame(scroll_frame)
            row.pack(fill="x", pady=2, padx=5)
            
            ctk.CTkLabel(row, text=name, width=250, anchor="w").pack(side="left", padx=5)
            
            # Default value based on initial mapping or Abaikan
            default_val = initial_mapping.get(name, "Abaikan")
            if default_val not in self.categories: default_val = "Abaikan"
            
            var = ctk.StringVar(value=default_val)
            menu = ctk.CTkOptionMenu(row, values=self.categories, variable=var, width=200)
            menu.pack(side="right", padx=5)
            self.menus[name] = var
            
        ctk.CTkButton(self, text="SIMPAN & PROSES", command=self.on_save, fg_color="green").pack(pady=10)
        self.grab_set()

    def on_save(self):
        for name, var in self.menus.items():
            self.result[name] = var.get()
        self.destroy()

class CoreTaxApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("CoreTax Auto-Monitor (Robust v1.5)")
        self.geometry("900x750")
        
        self.token = ""
        self.cookie = ""
        self.tid = ""
        self.file_path = ""
        self.dynamic_categories = {}
        self.history_file = "mapping_history.json"
        
        self.setup_ui()

    def setup_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)
        
        # --- LOGIN SECTION ---
        login_frame = ctk.CTkFrame(self)
        login_frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        
        self.btn_login = ctk.CTkButton(login_frame, text="LOGIN CORETAX", command=self.start_login, fg_color="orange", text_color="black")
        self.btn_login.pack(side="left", padx=10, pady=10)
        
        self.entry_token = ctk.CTkEntry(login_frame, placeholder_text="Token (Auto-filled)", width=200)
        self.entry_token.pack(side="left", padx=5)
        
        self.entry_cookie = ctk.CTkEntry(login_frame, placeholder_text="Cookie (Auto-filled)", width=200)
        self.entry_cookie.pack(side="left", padx=5)

        self.entry_tid = ctk.CTkEntry(login_frame, placeholder_text="TID / NPWP16", width=200)
        self.entry_tid.pack(side="left", padx=5)

        # --- INPUT SECTION ---
        input_frame = ctk.CTkFrame(self)
        input_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        input_frame.grid_columnconfigure(1, weight=1)
        
        # Row 0: Excel File
        ctk.CTkLabel(input_frame, text="Target Excel:").grid(row=0, column=0, padx=10, pady=10)
        self.entry_file = ctk.CTkEntry(input_frame, placeholder_text="Pilih file monitoring...")
        self.entry_file.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(input_frame, text="BROWSE", width=100, command=self.browse_file).grid(row=0, column=2, padx=10, pady=10)

        # Row 1: Mode Proses
        ctk.CTkLabel(input_frame, text="MODE PROSES:").grid(row=1, column=0, padx=10, pady=5)
        self.mode_var = ctk.StringVar(value="AUTO-PROCESS (Clear & Fill)")
        self.mode_menu = ctk.CTkOptionMenu(input_frame, values=["AUTO-PROCESS (Clear & Fill)", "ONLY-CLEAR (Hapus Saja)", "CHECK-ONLY (Cek PDF Saja)"], variable=self.mode_var)
        self.mode_menu.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        # Row 2: Checkbox Auto-Map
        self.check_auto_map = ctk.CTkCheckBox(input_frame, text="AUTO-MAPPING (Gunakan Memori Pintar)")
        self.check_auto_map.select()
        self.check_auto_map.grid(row=2, column=1, padx=10, pady=10, sticky="w")

        # Row 3: QUICK CHECK (Single Invoice)
        ctk.CTkLabel(input_frame, text="CEK 1 FAKTUR:").grid(row=3, column=0, padx=10, pady=5)
        self.entry_single_faktur = ctk.CTkEntry(input_frame, placeholder_text="Masukkan 16/17 digit no faktur...")
        self.entry_single_faktur.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        self.btn_quick_check = ctk.CTkButton(input_frame, text="QUICK CHECK", width=100, command=self.quick_check, fg_color="#cc7a00", text_color="white")
        self.btn_quick_check.grid(row=3, column=2, padx=10, pady=5)
        
        # Log Box (Main Window Row 3)
        self.log_box = ctk.CTkTextbox(self, font=("Consolas", 12))
        self.log_box.grid(row=3, column=0, padx=20, pady=10, sticky="nsew")
        
        # Start Button (Main Window Row 4)
        self.btn_run = ctk.CTkButton(self, text="START PROCESS", height=50, command=self.start_process, font=("Arial", 16, "bold"))
        self.btn_run.grid(row=4, column=0, padx=20, pady=20, sticky="ew")
        
        # Default values (Placeholder)
        self.entry_cookie.insert(0, "id-ID")
        self.entry_tid.insert(0, "275bb07a-d021-4389-943e-a740246a56e8")

    def add_log(self, msg):
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.update_idletasks()

    def browse_file(self):
        f = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if f:
            self.file_path = f
            self.entry_file.delete(0, "end")
            self.entry_file.insert(0, f)

    def load_mapping_history(self):
        if os.path.exists(self.history_file):
            try:
                with open(self.history_file, 'r') as f:
                    return json.load(f)
            except: return {}
        return {}

    def save_mapping_history(self, new_mapping):
        history = self.load_mapping_history()
        history.update(new_mapping)
        with open(self.history_file, 'w') as f:
            json.dump(history, f, indent=4)

    def _get_smart_cat(self, item_name, categories):
        history = self.load_mapping_history()
        if item_name in history:
            return history[item_name]
        
        # Conservative heuristic: Only auto-match core items
        name_up = item_name.upper()
        if "PUPUK" in name_up:
            target = next((c for c in categories if "PUPUK" in c.upper()), None)
            if target: return target
            
        return "Abaikan"

    def start_login(self):
        self.btn_login.configure(state="disabled", text="BROWSER OPEN...")
        threading.Thread(target=self.playwright_login, daemon=True).start()

    def playwright_login(self):
        from playwright.sync_api import sync_playwright
        self.add_log("[*] Membuka browser untuk login...")
        try:
            with sync_playwright() as p:
                user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
                browser = p.chromium.launch(headless=False)
                context = browser.new_context(user_agent=user_agent)
                page = context.new_page()
                
                found_token = False
                def handle_request(request):
                    nonlocal found_token
                    # Cek semua header secara case-insensitive
                    headers = {k.lower(): v for k, v in request.headers.items()}
                    auth = headers.get("authorization")
                    
                    if auth and "Bearer" in auth and not found_token:
                        token = auth.replace("Bearer ", "").strip()
                        self.entry_token.delete(0, "end")
                        self.entry_token.insert(0, token)
                        
                        cookies = context.cookies()
                        cookie_str = "; ".join([f"{c['name']}={c['value']}" for c in cookies])
                        self.entry_cookie.delete(0, "end")
                        self.entry_cookie.insert(0, cookie_str)
                        
                        self.add_log("[√] Sesi ditangkap secara otomatis!")
                        found_token = True

                page.on("request", handle_request)
                page.on("close", lambda: self.btn_login.configure(state="normal", text="LOGIN CORETAX"))
                
                # Gunakan portal utama agar redirect terbaca sempurna
                page.goto("https://coretaxdjp.pajak.go.id/identityproviderportal/Account/Login")
                self.add_log("[*] Silakan masukkan username & password di browser...")
                
                while True:
                    time.sleep(1)
                    if not browser.is_connected(): break
                    if found_token:
                        # Tetap biarkan browser terbuka sebentar agar user bisa melihat dashboard
                        time.sleep(2)
                        break
        except Exception as e:
            self.add_log(f"[!] Browser Error: {str(e)}")
            self.btn_login.configure(state="normal", text="LOGIN CORETAX")

    def auto_detect_header(self, file_path, sheet_name='PM'):
        try:
            df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=10)
            for i, row in df_raw.iterrows():
                row_str = " ".join([str(c) for c in row.values]).upper()
                if "NOMOR FAKTUR" in row_str or "HARGA JUAL" in row_str:
                    return i
            return 0
        except: return 0

    def _get_col_idx(self, ws, name_part):
        for cell in ws[1]:
            if name_part.upper() in str(cell.value).upper():
                return cell.column
        return None

    def _fetch_pdf_items(self, faktur, tid, expected_total):
        if not hasattr(self, 'session'): self.session = requests.Session()
        
        list_url = "https://coretaxdjp.pajak.go.id/einvoiceportal/api/inputinvoice/list"
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Cookie": self.cookie,
            "Content-Type": "application/json",
            "x-dgt-code": "7AcAAA=="
        }
        search_payload = {
            "BuyerTaxpayerAggregateIdentifier": tid, "TaxpayerAggregateIdentifier": tid,
            "First": 0, "Rows": 1, "LanguageId": "id-ID",
            "Filters": [{"PropertyName": "TaxInvoiceNumber", "Value": faktur, "MatchMode": "equals"}]
        }
        
        try:
            time.sleep(0.4)
            resp = self.session.post(list_url, json=search_payload, headers=headers, timeout=15)
            if resp.status_code == 401: return None, "401 Sesi Habis"
            if resp.status_code != 200: return None, f"HTTP {resp.status_code}"
            
            items_list = resp.json().get("Payload", {}).get("Data", [])
            if not items_list: return None, "Faktur tidak ditemukan"
            
            inv = items_list[0]
            detail_url = "https://coretaxdjp.pajak.go.id/einvoiceportal/api/DownloadInvoice/download-invoice-document"
            detail_payload = {
                "EInvoiceRecordIdentifier": inv["RecordId"],
                "EInvoiceAggregateIdentifier": inv["AggregateIdentifier"],
                "DocumentAggregateIdentifier": inv["DocumentFormAggregateIdentifier"],
                "TaxpayerAggregateIdentifier": tid,
                "LetterNumber": faktur,
                "EInvoiceMenuType": "Input",
                "TaxInvoiceStatus": "APPROVED"
            }
            d_resp = self.session.post(detail_url, json=detail_payload, headers=headers)
            pdf_b64 = d_resp.json().get("Content")
            if not pdf_b64: return None, "Konten PDF kosong"
            
            pdf_bytes = base64.b64decode(pdf_b64)
            extracted_items = []
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if not table: continue
                    for row in table:
                        if not row or len(row) < 3: continue
                        row_text = " ".join([str(c) if c is not None else "" for c in row])
                        nums = re.findall(r'[\d.]+(?:,[\d]+)?', row_text)
                        parsed_nums = []
                        for n in nums:
                            clean_n = n.replace('.', '').replace(',', '.')
                            try:
                                val = float(clean_n)
                                if val > 0: parsed_nums.append(val)
                            except: pass
                        
                        if not parsed_nums: continue
                        for n in parsed_nums:
                            # If Quick Check (expected=0), we collect all likely price lines
                            # If Normal Process, we match only specific line
                            if expected_total == 0 or abs(n - expected_total) < 100:
                                qty = 1.0
                                for q_cand in parsed_nums:
                                    if q_cand != n and q_cand < 1000000: qty = q_cand; break
                                cells = [str(c) if c is not None else "" for c in row]
                                raw_name = max(cells, key=len).strip()
                                name = re.split(r'Rp|\sx\s|\n', raw_name)[0].strip()
                                extracted_items.append({"name": name, "qty": qty, "total": n})
                                if expected_total > 0: return extracted_items, "Success"
            
            if expected_total == 0 and extracted_items: return extracted_items, "Success"
            return None, "Item tidak ditemukan di PDF"
        except Exception as e: return None, str(e)

    def quick_check(self):
        faktur = self.entry_single_faktur.get().strip()
        if not faktur: return self.add_log("[!] Masukkan nomor faktur.")
        self.token = self.entry_token.get().strip()
        self.cookie = self.entry_cookie.get().strip()
        tid = self.entry_tid.get().strip()
        if not self.token or not self.cookie: return self.add_log("[!] Silakan Login dulu.")

        def run():
            self.add_log(f"\n[QUICK CHECK] Faktur: {faktur}")
            items, msg = self._fetch_pdf_items(faktur, tid, 0)
            if items:
                self.add_log("--- [ITEM TERDETEKSI] ---")
                for it in items:
                    self.add_log(f" > {it['name']} | QTY: {it['qty']} | Harga: {it['total']:,}")
                self.add_log("-------------------------")
            else:
                self.add_log(f"[!] Gagal: {msg}")
        threading.Thread(target=run, daemon=True).start()

    def start_process(self):
        if not self.file_path: return self.add_log("[!] Pilih file Excel.")
        self.btn_run.configure(state="disabled", text="PROCESSING...")
        threading.Thread(target=self.run_logic, daemon=True).start()

    def run_logic(self):
        try:
            self.token = self.entry_token.get().strip()
            self.cookie = self.entry_cookie.get().strip()
            tid = self.entry_tid.get().strip()
            
            self.add_log(f"[+] File: {os.path.basename(self.file_path)}")
            h_idx = self.auto_detect_header(self.file_path)
            df = pd.read_excel(self.file_path, sheet_name='PM', header=h_idx)
            headers = df.columns.tolist()
            
            # Identify columns by index (Robust)
            c_faktur_idx = next((i for i, c in enumerate(headers) if "Nomor Faktur" in str(c)), None)
            c_harga_idx = next((i for i, c in enumerate(headers) if "Harga Jual" in str(c)), None)
            c_selisih_idx = next((i for i, c in enumerate(headers) if "Selisih" in str(c)), None)
            
            self.dynamic_categories = {}
            for i, col in enumerate(headers):
                if "QTY" in str(col).upper() and i+1 < len(headers):
                    cat_name = str(headers[i+1]).strip()
                    if not any(b in cat_name.upper() for b in ["NPWP", "DPP", "PPN", "QTY"]):
                        # (Col_Price_Excel, Col_QTY_Excel, Pandas_Col_Index)
                        self.dynamic_categories[cat_name] = (i+2, i+1, i+1)
            
            def should_process(r):
                for val in r:
                    if isinstance(val, (int, float)) and val in [310000, 320000, 31000, 32000]: return True
                if pd.isna(r.iloc[c_faktur_idx]): return False
                hj = pd.to_numeric(r.iloc[c_harga_idx], errors='coerce')
                if pd.isna(hj) or hj <= 0: return False
                
                cur_sum = 0
                for cat, (p, q, idx) in self.dynamic_categories.items():
                    val = pd.to_numeric(r.iloc[idx], errors='coerce')
                    if not pd.isna(val): cur_sum += val
                
                selisih = r.iloc[c_selisih_idx]
                if pd.isna(selisih) or cur_sum == 0: return True
                return False

            targets = df[df.apply(should_process, axis=1)].copy()
            targets['excel_row'] = targets.index + h_idx + 2
            self.add_log(f"[*] Ditemukan {len(targets)} faktur.")
            
            wb = openpyxl.load_workbook(self.file_path)
            ws = wb["PM"]
            
            # WIPE
            for _, row in targets.iterrows():
                ex_r = int(row['excel_row'])
                for cat, (p, q, idx) in self.dynamic_categories.items():
                    ws.cell(row=ex_r, column=p).value = None
                    ws.cell(row=ex_r, column=q).value = None
            
            if self.mode_var.get() == "ONLY-CLEAR (Hapus Saja)":
                wb.save(self.file_path); return self.add_log("[√] Clear Selesai.")

            # SCAN PDF
            all_items = {}; unique_names = set()
            for _, row in targets.iterrows():
                f = "{:.0f}".format(row.iloc[c_faktur_idx]).zfill(17) if isinstance(row.iloc[c_faktur_idx], (float, int)) else str(row.iloc[c_faktur_idx]).zfill(17)
                hj = pd.to_numeric(row.iloc[c_harga_idx], errors='coerce')
                self.add_log(f" > Scan: {f}")
                items, msg = self._fetch_pdf_items(f, tid, hj)
                if items:
                    all_items[int(row['excel_row'])] = items
                    for it in items: unique_names.add(it['name'])
                else: self.add_log(f"   [!] Gagal: {msg}")

            if not all_items: return self.add_log("[!] Tidak ada data PDF.")

            # MAPPING
            final_mapping = {}; unmapped = []
            for name in unique_names:
                guessed = self._get_smart_cat(name, list(self.dynamic_categories.keys()))
                if guessed != "Abaikan": final_mapping[name] = guessed
                else: unmapped.append(name)
            
            if not self.check_auto_map.get() or unmapped:
                win = MappingWindow(self, list(unique_names), list(self.dynamic_categories.keys()), final_mapping)
                self.wait_window(win)
                final_mapping = win.result
                if final_mapping: self.save_mapping_history(final_mapping)

            if not final_mapping: return self.add_log("[!] Mapping batal.")

            # FILL
            for r_idx, items in all_items.items():
                for it in items:
                    cat = final_mapping.get(it['name'])
                    if cat and cat in self.dynamic_categories:
                        p, q, idx = self.dynamic_categories[cat]
                        ws.cell(row=r_idx, column=p).value = it['total']
                        ws.cell(row=r_idx, column=q).value = it['qty']
                if c_selisih_idx: ws.cell(row=r_idx, column=c_selisih_idx+1).value = "-"
            
            wb.save(self.file_path)
            self.add_log("[√] SELESAI.")
        except Exception as e: self.add_log(f"[!] Error: {str(e)}")
        finally: self.btn_run.configure(state="normal", text="START PROCESS")

if __name__ == "__main__":
    app = CoreTaxApp()
    app.mainloop()
