path = r"c:\Work\project automation\coretax\coretax_app.py"
with open(path, "r", encoding="utf-8") as f:
    lines = f.readlines()

start_line = -1
end_line = -1

for i, line in enumerate(lines):
    if "def _fetch_pdf_items(self, no_faktur, tid, expected_total):" in line:
        start_line = i
    if start_line != -1 and "except Exception as e: return None, str(e), None" in line:
        end_line = i
        break

if start_line != -1 and end_line != -1:
    new_func = [
        "    def _fetch_pdf_items(self, no_faktur, tid, expected_total):\n",
        "        headers = {\"authority\": \"coretaxdjp.pajak.go.id\", \"authorization\": f\"Bearer {self.token}\", \"content-type\": \"application/json\", \"cookie\": self.cookie, \"x-dgt-code\": \"7AcAAA==\"}\n",
        "        try:\n",
        "            s_url = \"https://coretaxdjp.pajak.go.id/einvoiceportal/api/inputinvoice/list\"\n",
        "            payload = {\"BuyerTaxpayerAggregateIdentifier\": tid, \"TaxpayerAggregateIdentifier\": tid, \"First\": 0, \"Rows\": 1, \"LanguageId\": \"id-ID\", \"Filters\": [{\"PropertyName\": \"TaxInvoiceNumber\", \"Value\": no_faktur, \"MatchMode\": \"equals\"}]}\n",
        "            resp = requests.post(s_url, headers=headers, json=payload, timeout=15).json()\n",
        "            data = resp.get(\"Payload\", {}).get(\"Data\", [])\n",
        "            if not data: return None, \"Faktur tidak ditemukan di server\", None\n",
        "\n",
        "            inv = data[0]\n",
        "            d_url = \"https://coretaxdjp.pajak.go.id/einvoiceportal/api/DownloadInvoice/download-invoice-document\"\n",
        "            d_payload = {\"EInvoiceRecordIdentifier\": inv[\"RecordId\"], \"EInvoiceAggregateIdentifier\": inv[\"AggregateIdentifier\"], \"DocumentAggregateIdentifier\": inv[\"DocumentFormAggregateIdentifier\"], \"TaxpayerAggregateIdentifier\": tid, \"LetterNumber\": no_faktur, \"EInvoiceMenuType\": \"Input\", \"TaxInvoiceStatus\": \"APPROVED\"}\n",
        "            d_resp = requests.post(d_url, headers=headers, json=d_payload, timeout=15).json()\n",
        "            pdf_b64 = d_resp.get(\"Content\")\n",
        "            if not pdf_b64: return None, \"Data PDF kosong dari server\", inv\n",
        "\n",
        "            pdf_bytes = base64.b64decode(pdf_b64)\n",
        "            extracted_items = []\n",
        "            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:\n",
        "                for page in pdf.pages:\n",
        '                    table = page.extract_table()\n',
        '                    if not table: continue\n',
        '                    for row in table:\n',
        '                        if not row or len(row) < 3: continue\n',
        '\n',
        '                        # 1. CEK STRUKTUR: Kolom pertama harus Angka Urut (1, 2, 3...)\n',
        '                        first_col = str(row[0]).strip()\n',
        '                        if not first_col.isdigit(): continue\n',
        '\n',
        '                        # 2. EKSTRAKSI NAMA (Kolom ke-3, Index 2)\n',
        '                        raw_name_cell = str(row[2]) if row[2] else ""\n',
        '                        # Bersihkan Nama dari teks rumus harga (seperti Rp 185 x 40.000)\n',
        '                        clean_name = re.split(r"Rp|\\sx\\s|\\n", raw_name_cell)[0].strip()\n',
        '                        # NORMALISASI SPASI: Hapus spasi ganda agar nama yang mirip jadi satu\n',
        '                        clean_name = re.sub(r"\\s+", " ", clean_name).strip()\n',
        '\n',
        '                        # 4. EKSTRAKSI DATA (Logika Perkalian Matematika)\n',
        '                        all_nums = []\n',
        '                        all_cells_text = " ".join([str(c) for c in row[1:] if c])\n',
        '                        found_nums_raw = re.findall(r"[\\d.]+(?:,[\\d]+)?", all_cells_text)\n',
        '                        for n in found_nums_raw:\n',
        '                            try:\n',
        '                                val = float(n.replace(".", "").replace(",", "."))\n',
        '                                if val > 0: all_nums.append(val)\n',
        '                            except: pass\n',
        '\n',
        '                        if not all_nums: continue\n',
        '\n',
        '                        # Harga Total Baris: Pasti angka TERBESAR\n',
        '                        price = max(all_nums)\n',
        '\n',
        '                        # Qty: Cari pasangan A * B = Price (Logika Matematika 100% Akurat)\n',
        '                        qty = 1.0\n',
        '                        potential_candidates = [n for n in all_nums if n != price]\n',
        '\n',
        '                        found_math_match = False\n',
        '                        if len(potential_candidates) >= 2:\n',
        '                            for i, a in enumerate(potential_candidates):\n',
        '                                for b in potential_candidates[i:]:\n',
        '                                    # Cek apakah a * b mendekati price (toleransi 1% untuk pembulatan pajak)\n',
        '                                    if abs((a * b) - price) < (0.01 * price):\n',
        '                                        # Ambil yang terkecil sebagai Qty\n',
        '                                        qty = a if a < b else b\n',
        '                                        found_math_match = True\n',
        '                                        break\n',
        '                                if found_math_match: break\n',
        '\n',
        '                        # Fallback jika tidak ada perkalian yang match (misal Qty tersembunyi/tunggal)\n',
        '                        if not found_math_match:\n',
        '                            q_match = re.search(r"([\\d.,]+)\\s*x", all_cells_text)\n',
        '                            if q_match:\n',
        '                                qty = float(q_match.group(1).replace(".", "").replace(",", "."))\n',
        '                            else:\n',
        '                                # Jika ada angka selain Harga, ambil yang paling kecil (asumsi Qty)\n',
        '                                if potential_candidates:\n',
        '                                    qty = min(potential_candidates)\n',
        '\n',
        '                        extracted_items.append({"name": clean_name, "qty": qty, "total": price})\n',
        '\n',
        '            if not extracted_items:\n',
        '                return None, "Tidak ada item barang bernomor ditemukan di PDF", inv\n',
        '\n',
        '            # 5. VALIDASI AKURASI (Jika bukan Quick Check)\n',
        '            if expected_total > 0:\n',
        '                sum_pdf = sum(it["total"] for it in extracted_items)\n',
        '                if abs(sum_pdf - expected_total) > 500: # Toleransi 500 rupiah\n',
        '                    return None, f"Total PDF ({sum_pdf:,}) tidak sinkron dengan Excel ({expected_total:,})", inv\n',
        '\n',
        '            return extracted_items, "Success", inv\n',
        '        except Exception as e: return None, str(e), None\n'
    ]
    
    lines[start_line:end_line+1] = new_func
    
    with open(path, "w", encoding="utf-8") as f:
        f.writelines(lines)
    print("Patch applied successfully")
else:
    print(f"Start: {start_line}, End: {end_line}")
