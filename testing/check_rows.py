import pandas as pd
import openpyxl
import os

FILE_NAME = 'MONITORING PK-PM RTSP 2026AAA.xlsx'
SHEET_NAME = 'PM'

def is_not_positive(val):
    if val is None: return True
    s_val = str(val).strip()
    if s_val == "" or s_val == "0" or s_val == "0.0": return True
    if s_val.startswith("="): return True
    try:
        f_val = float(s_val.replace(',',''))
        if f_val <= 0: return True
    except: pass
    return False

def check_file():
    if not os.path.exists(FILE_NAME):
        print(f"ERROR: File {FILE_NAME} tidak ditemukan!")
        return

    print(f"[*] Membuka File: {FILE_NAME}")
    wb = openpyxl.load_workbook(FILE_NAME, data_only=False)
    ws = wb[SHEET_NAME]
    data_raw = list(ws.values)
    
    # Deteksi Header
    h_idx = -1
    for i, row in enumerate(data_raw):
        if any("NOMOR FAKTUR" in str(cell).upper() for cell in row if cell):
            h_idx = i
            break
            
    if h_idx == -1:
        print("ERROR: Header 'NOMOR FAKTUR' tidak ditemukan!")
        return

    headers = [str(h).strip() if h else f"COL_{i}" for i, h in enumerate(data_raw[h_idx])]
    df = pd.DataFrame(data_raw[h_idx+1:], columns=headers)
    
    c_faktur = next((c for c in headers if "NOMOR FAKTUR" in str(c).upper()), None)
    c_penjabaran = next((c for c in headers if "PENJABARAN" in str(c).upper()), None)
    
    # Deteksi Kategori
    dynamic_categories = {}
    for i, col_str in enumerate(headers):
        if "QTY" in str(col_str).upper():
            if i + 1 < len(headers):
                raw_cat = str(headers[i+1]).strip()
                if raw_cat and not raw_cat.startswith("COL_"):
                    dynamic_categories[raw_cat] = i + 1 # Index Kolom (0-based)
    
    print(f"[+] Kategori Terdeteksi: {list(dynamic_categories.keys())}")
    
    targets = []
    skipped_reasons = []

    for idx, r in df.iterrows():
        excel_row = idx + h_idx + 2
        f_val = r[c_faktur]
        
        # 1. Cek Faktur
        if pd.isna(f_val) or str(f_val).strip() == "":
            continue
            
        # 2. Cek Isi Kategori
        has_any_real_entry = False
        reason_entry = ""
        for cat, col_idx in dynamic_categories.items():
            val = r[cat]
            if val is not None:
                s_v = str(val).strip()
                if s_v != "" and s_v != "0" and s_v != "0.0" and not s_v.startswith("="):
                    try:
                        if float(s_v.replace(',','')) != 0:
                            has_any_real_entry = True
                            reason_entry = f"Ada isi di {cat}: {s_v}"
                            break
                    except: pass

        # 3. Cek Penjabaran
        pj = r[c_penjabaran]
        if is_not_positive(pj):
            if str(pj).startswith("=") and has_any_real_entry:
                skipped_reasons.append(f"Baris {excel_row}: SKIP (Rumus tapi {reason_entry})")
            else:
                targets.append(excel_row)
        else:
            skipped_reasons.append(f"Baris {excel_row}: SKIP (Penjabaran > 0: {pj})")

    print("\n" + "="*50)
    print(f"HASIL ANALISIS:")
    print(f"Total Faktur Valid: {len(df[~pd.isna(df[c_faktur])])}")
    print(f"Total Target Ditemukan: {len(targets)}")
    print(f"Baris-baris Target: {targets[:10]} ... {targets[-5:] if len(targets)>5 else ''}")
    print("="*50 + "\n")
    
    if len(targets) < 5:
        print("[!] Mengapa yang lain dilewati? Ini beberapa alasannya:")
        for reason in skipped_reasons[:20]:
            print(f"  > {reason}")

if __name__ == "__main__":
    check_file()
