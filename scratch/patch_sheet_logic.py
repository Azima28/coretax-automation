import re

path = r'c:\Work\project automation\coretax\coretax_app.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

new_logic = """    def run_logic(self):
        try:
            self.btn_run.configure(state="disabled", text="PROCESSING...")
            wb = openpyxl.load_workbook(self.file_path)
            
            # 1. CARI SHEET YANG RELEVAN (Cari yang ada 'Nomor Faktur')
            target_sheet_name = None
            for sn in wb.sheetnames:
                ws_temp = wb[sn]
                found_header = False
                for r in range(1, 15): # Cek 15 baris pertama
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
                return"""

# Find the start of run_logic and replace until c_selisih logic
pattern = r'    def run_logic\(self\):.*?c_selisih = next\(\(c for c in headers if "Selisih" in str\(c\)\), None\)'
new_content = re.sub(pattern, new_logic, content, flags=re.DOTALL)

with open(path, 'w', encoding='utf-8') as f:
    f.write(new_content)

print("Patch applied successfully")
