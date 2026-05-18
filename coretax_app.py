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
import webbrowser
from PIL import Image

class AboutWindow(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("About ZiTax Automator")
        self.geometry("400x350")
        self.transient(parent)
        self.grab_set()
        
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (400 // 2)
        y = (self.winfo_screenheight() // 2) - (350 // 2)
        self.geometry(f"+{x}+{y}")
        
        # Logo placeholder
        self.logo_label = ctk.CTkLabel(self, text="[ TEMPAT LOGO ZITAX ]", font=("Arial", 16, "bold"), width=150, height=100, fg_color="#2c3e50", corner_radius=10)
        self.logo_label.pack(pady=(20, 10))
        
        # Coba load logo jika ada
        logo_path = os.path.join(os.path.dirname(__file__), "zitax_logo.png")
        if os.path.exists(logo_path):
            try:
                logo_img = ctk.CTkImage(light_image=Image.open(logo_path), dark_image=Image.open(logo_path), size=(120, 120))
                self.logo_label.configure(text="", image=logo_img)
            except Exception:
                pass
                
        ctk.CTkLabel(self, text="ZiTax Automator", font=("Arial", 22, "bold")).pack(pady=5)
        ctk.CTkLabel(self, text="Version 1.4 Premium", font=("Arial", 12)).pack(pady=(0, 15))
        
        # Clickable links
        ig_label = ctk.CTkLabel(self, text="Instagram: @zimm.dev", font=("Arial", 14), text_color="#3498db", cursor="hand2")
        ig_label.pack(pady=5)
        ig_label.bind("<Button-1>", lambda e: webbrowser.open("https://instagram.com/zimm.dev"))
        
        email_label = ctk.CTkLabel(self, text="Email: azimarizki228@gmail.com", font=("Arial", 14), text_color="#3498db", cursor="hand2")
        email_label.pack(pady=5)
        email_label.bind("<Button-1>", lambda e: webbrowser.open("mailto:azimarizki228@gmail.com"))
        
        ctk.CTkButton(self, text="Tutup", command=self.destroy, width=100).pack(pady=20)

class iOSConfirmDialog(ctk.CTkToplevel):
    def __init__(self, parent, title, message):
        super().__init__(parent)
        self.title("")
        self.geometry("340x260")
        
        # Windows transparent corners hack
        if os.name == 'nt':
            self.configure(fg_color="#000001")
            self.attributes("-transparentcolor", "#000001")
        else:
            self.configure(fg_color="transparent")
            
        self.overrideredirect(True)
        self.transient(parent)
        self.grab_set()
        
        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (340 // 2)
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (260 // 2)
        self.geometry(f"+{x}+{y}")
        
        self.result = False
        
        main_frame = ctk.CTkFrame(self, fg_color="#1C1C1E", corner_radius=15, border_width=1, border_color="#38383A")
        main_frame.pack(fill="both", expand=True, padx=2, pady=2)
        
        ctk.CTkLabel(main_frame, text=title, font=("Segoe UI", 16, "bold"), text_color="#FFFFFF").pack(pady=(20, 5))
        ctk.CTkLabel(main_frame, text=message, font=("Segoe UI", 13), text_color="#8E8E93", justify="center", wraplength=300).pack(pady=(0, 15), padx=20, expand=True)
        
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=(0, 15))
        btn_frame.grid_columnconfigure((0, 1), weight=1)
        
        btn_no = ctk.CTkButton(btn_frame, text="No", font=("Segoe UI", 15, "bold"), fg_color="#2C2C2E", text_color="#0A84FF", hover_color="#3A3A3C", corner_radius=8, height=40, command=self.on_no)
        btn_no.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        
        btn_yes = ctk.CTkButton(btn_frame, text="Yes", font=("Segoe UI", 15, "bold"), fg_color="#FF3B30", text_color="#FFFFFF", hover_color="#FF453A", corner_radius=8, height=40, command=self.on_yes)
        btn_yes.grid(row=0, column=1, sticky="ew", padx=(5, 0))
        
        # Focus management agar selalu di atas window utama tapi tidak menimpa aplikasi lain
        def on_focus_in(e): self.attributes("-topmost", True)
        def on_focus_out(e): self.attributes("-topmost", False)
        
        self.bind("<FocusIn>", on_focus_in)
        self.bind("<FocusOut>", on_focus_out)
        parent.bind("<FocusIn>", on_focus_in, add="+")
        
    def on_yes(self):
        self.result = True
        self.grab_release()
        self.master.grab_set()
        self.destroy()
        
    def on_no(self):
        self.result = False
        self.grab_release()
        self.master.grab_set()
        self.destroy()

class MappingWindow(ctk.CTkToplevel):
    def __init__(self, parent, all_names, categories):
        super().__init__(parent)
        self.title("Batch Mapping: Tentukan Kategori Barang")
        self.geometry("800x700")
        self.result = {}
        self.categories = ["Abaikan"] + sorted(categories)
        self.original_names = sorted(list(all_names)) # Simpan nama asli
        self.current_items = self.original_names # Daftar yang sedang tampil
        
        # UI Setup
        self.label_info = ctk.CTkLabel(self, text=f"Ditemukan {len(self.original_names)} barang unik (ASLI).", font=("Arial", 14, "bold"))
        self.label_info.pack(pady=10)
        
        btn_frame = ctk.CTkFrame(self)
        btn_frame.pack(fill="x", padx=20)
        
        self.btn_auto = ctk.CTkButton(btn_frame, text="AUTO-MAP & RINGKAS (SMART)", fg_color="#d35400", hover_color="#e67e22", command=self.do_smart_action)
        self.btn_auto.pack(side="right", pady=10)

        self.scroll_frame = ctk.CTkScrollableFrame(self)
        self.scroll_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        self.combo_vars = {} # {display_name: StringVar}
        self.render_ui(self.current_items)

        self.btn_save = ctk.CTkButton(self, text="SIMPAN & PROSES SEMUA", fg_color="#27ae60", hover_color="#2ecc71", font=("Arial", 14, "bold"), command=self.save_mapping)
        self.btn_save.pack(pady=20)
        # Dibuat terpisah agar muncul di taskbar sebagai window mandiri
        self.focus_force()
        self.lift()
        self.grab_set()

    def render_ui(self, names):
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
        self.combo_vars = {}
        
        for name in names:
            row_f = ctk.CTkFrame(self.scroll_frame)
            row_f.pack(fill="x", pady=2, padx=5)
            
            ctk.CTkLabel(row_f, text=f"Nama : {name}", wraplength=450, justify="left").pack(side="left", padx=10)
            
            var = ctk.StringVar(value="Abaikan")
            combo = ctk.CTkComboBox(row_f, values=self.categories, variable=var, width=200)
            combo.pack(side="right", padx=10)
            self.combo_vars[name] = var

    def do_smart_action(self):
        # 1. Grouping Logic (Anchor & Prefix)
        def get_core(name):
            n = str(name).upper().strip()
            # Hapus Satuan
            n = re.sub(r'\b(METER KUBIK|METER|KUBIK|UNIT|PCS|KG|LITER|SAK|ZAK|LTR|UNIT|BOX|ROLL|BTG|LBR|CURAH|JB)\b', ' ', n)
            n = re.sub(r'[\d.,\-()/xX*]+', ' ', n)
            words = n.split()
            if not words: return name.upper()
            anchors = ["JASA", "SERVICE", "SP", "SMN", "GT", "WL", "R", "C", "MATERIAL", "ULTRAPRO", "SOLAR", "BIOSOLAR", "SPLIT", "SCREENING", "ABU"]
            first_word = words[0]
            if first_word in anchors:
                if first_word in ["JASA", "SP", "SMN", "MATERIAL", "ULTRAPRO", "SOLAR", "BIOSOLAR", "ABU"]:
                    return first_word
                return " ".join(words[:2])
            return " ".join(words[:2])

        self.final_group_map = {} # {core: [originals]}
        for n in self.original_names:
            core = get_core(n)
            if core not in self.final_group_map: self.final_group_map[core] = []
            self.final_group_map[core].append(n)
            
        # 2. Update UI ke versi Ringkas
        self.current_items = sorted(list(self.final_group_map.keys()))
        self.label_info.configure(text=f"Selesai! Diringkas menjadi {len(self.current_items)} kelompok.")
        self.render_ui(self.current_items)
        
        # 3. Guess Categories
        knowledge = {
            "SPAREPART": ["BAN", "TYRE", "FILTER", "BOLT", "NUT", "BEARING", "PART", "OIL", "STRAP", "GREASE", "VALVE", "HOSE", "CHAMPIRO", "MAXMILER", "WL"],
            "SEMEN": ["SEMEN", "SMN", "PCC", "ULTRAPRO", "PORTLAND", "ALUMINA"],
            "BBM": ["SOLAR", "HSD", "BBM", "BIOSOLAR", "DEXLITE"],
            "BESI": ["BESI", "STEEL", "WIRE", "MESH", "PLATE", "UNP", "CNP"],
            "JASA ANGKUTAN": ["ANGKUTAN", "TRANSPORT", "CARGO", "EKSPEDISI", "TRUCKING"],
            "SEWA": ["SEWA", "RENTAL", "RENT", "SEWA ALAT", "ALAT BERAT", "EXCAVATOR", "DUMP TRUCK"],
            "SPLIT": ["SPLIT", "BATU SPLIT"],
            "SCREENING": ["SCREENING", "ABU BATU", "ABU"],
            "BETON": ["BETON", "READY MIX"],
            "HOTMIX": ["HOTMIX", "ASPAL", "ASPHALT"]
        }
        
        for name, var in self.combo_vars.items():
            u_name = name.upper()
            best_cat = "Abaikan"; max_score = 0
            for cat in self.categories[1:]:
                u_cat = cat.upper()
                # Cek Keyword Exact Match dulu
                for k_cat, keys in knowledge.items():
                    if k_cat in u_cat or any(k in u_cat for k in keys):
                        if any(k in u_name for k in keys) or k_cat in u_name:
                            max_score = 1.0; best_cat = cat; break
                if max_score == 1.0: break
                
                # Fallback ke Fuzzy
                score = difflib.SequenceMatcher(None, u_name, u_cat).ratio()
                if score > max_score: max_score = score; best_cat = cat
                
            if max_score >= 0.7: var.set(best_cat)
            elif "JASA" in u_name:
                jasa_cat = next((c for c in self.categories if "JASA" in str(c).upper()), "Abaikan")
                var.set(jasa_cat)

    def save_mapping(self):
        abaikan_items = []
        for name, var in self.combo_vars.items():
            if var.get() == "Abaikan":
                abaikan_items.append(name)
                
        if abaikan_items:
            items_str = "\n".join(f"- {item}" for item in abaikan_items[:4])
            if len(abaikan_items) > 4:
                items_str += f"\n... dan {len(abaikan_items) - 4} lainnya."
                
            dialog = iOSConfirmDialog(
                self,
                "Abaikan Barang",
                f"Barang berikut akan diabaikan:\n{items_str}\n\nApakah Anda yakin untuk melanjutkan?"
            )
            self.wait_window(dialog)
            if not dialog.result:
                return

        # Jika sudah di-group, sebar hasil mapping ke semua nama asli
        if hasattr(self, 'final_group_map'):
            for core_name, var in self.combo_vars.items():
                cat = var.get()
                for orig in self.final_group_map.get(core_name, []):
                    self.result[orig] = cat
        else:
            # Jika belum di-group (manual satu-satu)
            for name, var in self.combo_vars.items():
                self.result[name] = var.get()
        self.destroy()

class CoreTaxApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("ZiTax Automator")
        self.geometry("900x700")
        
        # Posisikan window di tengah layar
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (900 // 2)
        y = (self.winfo_screenheight() // 2) - (700 // 2)
        self.geometry(f"900x700+{x}+{y}")
        
        ctk.set_appearance_mode("dark")
        
        # Variables
        self.file_path = ""
        self.token = ""
        self.cookie = ""
        self.dynamic_categories = {} # {Kategori: (Index_Harga, Index_QTY)}
        
        # UI
        self.setup_ui()
        
    def setup_ui(self):
        self.configure(fg_color="#000000") # Pure black background for iOS style
        self.geometry("1150x700") # Lebarkan window untuk layout 2 kolom
        
        self.grid_columnconfigure(0, weight=1) # Kiri: Pengaturan
        self.grid_columnconfigure(1, weight=1) # Kanan: Logs
        self.grid_rowconfigure(1, weight=1) 
        
        # Font settings
        font_title = ("Segoe UI", 28, "bold")
        font_label = ("Segoe UI", 14)
        font_btn = ("Segoe UI", 14, "bold")
        font_link = ("Segoe UI", 14)
        
        # Header
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, columnspan=2, padx=20, pady=(20, 10), sticky="ew")
        
        ctk.CTkLabel(header_frame, text="ZiTax Automator", font=font_title, text_color="#FFFFFF").pack(side="left")
        
        self.status_label = ctk.CTkLabel(header_frame, text="Terhubung (TID: 275...)", text_color="#8E8E93", font=font_label)
        self.status_label.pack(side="right", padx=15)
        
        # ================= KOLOM KIRI: PENGATURAN =================
        left_panel = ctk.CTkFrame(self, fg_color="transparent")
        left_panel.grid(row=1, column=0, sticky="nsew", padx=(20, 10))
        left_panel.grid_columnconfigure(0, weight=1)
        
        input_kwargs = {"font": font_label, "fg_color": "#2C2C2E", "border_width": 1, "border_color": "#38383A", "text_color": "#0A84FF", "corner_radius": 6}

        # --- GROUP 1: KONEKSI CORETAX ---
        ctk.CTkLabel(left_panel, text="KONEKSI CORETAX", font=("Segoe UI", 12), text_color="#8E8E93").pack(anchor="w", padx=15, pady=(0, 5))
        
        group1 = ctk.CTkFrame(left_panel, fg_color="#1C1C1E", corner_radius=10)
        group1.pack(fill="x", pady=(0, 15))
        
        # Row 1
        row1 = ctk.CTkFrame(group1, fg_color="transparent")
        row1.pack(fill="x", padx=15, pady=8)
        row1.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(row1, text="Bearer Token", font=font_label, text_color="#FFFFFF", width=120, anchor="w").grid(row=0, column=0)
        self.entry_token = ctk.CTkEntry(row1, placeholder_text="Token...", **input_kwargs)
        self.entry_token.grid(row=0, column=1, sticky="ew")
        
        ctk.CTkFrame(group1, height=1, fg_color="#38383A").pack(fill="x", padx=15)
        
        # Row 2
        row2 = ctk.CTkFrame(group1, fg_color="transparent")
        row2.pack(fill="x", padx=15, pady=8)
        row2.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(row2, text="Cookie", font=font_label, text_color="#FFFFFF", width=120, anchor="w").grid(row=0, column=0)
        self.entry_cookie = ctk.CTkEntry(row2, placeholder_text="id-ID...", **input_kwargs)
        self.entry_cookie.grid(row=0, column=1, sticky="ew")
        self.entry_cookie.insert(0, "id-ID")
        
        ctk.CTkFrame(group1, height=1, fg_color="#38383A").pack(fill="x", padx=15)
        
        # Row 3
        row3 = ctk.CTkFrame(group1, fg_color="transparent")
        row3.pack(fill="x", padx=15, pady=8)
        row3.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(row3, text="Taxpayer ID", font=font_label, text_color="#FFFFFF", width=120, anchor="w").grid(row=0, column=0)
        self.entry_tid = ctk.CTkEntry(row3, placeholder_text="ID...", **input_kwargs)
        self.entry_tid.grid(row=0, column=1, sticky="ew")
        self.entry_tid.insert(0, "275bb07a-d021-4389-943e-a740246a56e8")
        
        ctk.CTkFrame(group1, height=1, fg_color="#38383A").pack(fill="x", padx=15)
        
        # Row 4 (Actions)
        row4 = ctk.CTkFrame(group1, fg_color="transparent")
        row4.pack(fill="x", padx=15, pady=10)
        row4.grid_columnconfigure((0,1), weight=1)
        
        btn_kwargs = {"font": font_link, "fg_color": "transparent", "text_color": "#0A84FF", "hover_color": "#2C2C2E", "border_width": 1, "border_color": "#38383A", "corner_radius": 6}
        
        self.btn_login = ctk.CTkButton(row4, text="Login CoreTax", command=self.start_login, **btn_kwargs)
        self.btn_login.grid(row=0, column=0, sticky="ew", padx=5)
        
        btn_save = ctk.CTkButton(row4, text="Simpan Sesi", command=self.save_session, **btn_kwargs)
        btn_save.grid(row=0, column=1, sticky="ew", padx=5)

        # --- GROUP 2: KONFIGURASI PROSES ---
        ctk.CTkLabel(left_panel, text="KONFIGURASI PROSES", font=("Segoe UI", 12), text_color="#8E8E93").pack(anchor="w", padx=15, pady=(5, 5))
        
        group2 = ctk.CTkFrame(left_panel, fg_color="#1C1C1E", corner_radius=10)
        group2.pack(fill="x", pady=(0, 20))
        
        # File Row
        row_f = ctk.CTkFrame(group2, fg_color="transparent")
        row_f.pack(fill="x", padx=15, pady=8)
        row_f.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(row_f, text="File Excel", font=font_label, text_color="#FFFFFF", width=120, anchor="w").grid(row=0, column=0)
        self.entry_file = ctk.CTkEntry(row_f, placeholder_text="Pilih file...", **input_kwargs)
        self.entry_file.configure(text_color="#8E8E93")
        self.entry_file.grid(row=0, column=1, sticky="ew")
        ctk.CTkButton(row_f, text="Browse", width=60, command=self.browse_file, **btn_kwargs).grid(row=0, column=2, padx=(5,0))
        
        ctk.CTkFrame(group2, height=1, fg_color="#38383A").pack(fill="x", padx=15)
        
        # Sheet Row
        row_s = ctk.CTkFrame(group2, fg_color="transparent")
        row_s.pack(fill="x", padx=15, pady=8)
        row_s.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(row_s, text="Target Sheet", font=font_label, text_color="#FFFFFF", width=120, anchor="w").grid(row=0, column=0)
        self.sheet_var = ctk.StringVar(value="Hanya PM (Pajak Masukan)")
        self.sheet_menu = ctk.CTkOptionMenu(row_s, values=["Hanya PM (Pajak Masukan)", "Hanya PK (Pajak Keluaran)", "Keduanya (PK & PM)"], variable=self.sheet_var, font=font_label, fg_color="#2C2C2E", button_color="#2C2C2E", text_color="#0A84FF", dropdown_fg_color="#1C1C1E")
        self.sheet_menu.grid(row=0, column=1, sticky="e")
        
        ctk.CTkFrame(group2, height=1, fg_color="#38383A").pack(fill="x", padx=15)
        
        # Mode Row
        row_m = ctk.CTkFrame(group2, fg_color="transparent")
        row_m.pack(fill="x", padx=15, pady=8)
        row_m.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(row_m, text="Eksekusi", font=font_label, text_color="#FFFFFF", width=120, anchor="w").grid(row=0, column=0)
        self.mode_var = ctk.StringVar(value="AUTO-PROCESS")
        self.mode_menu = ctk.CTkOptionMenu(row_m, values=["AUTO-PROCESS", "ONLY-CLEAR", "CHECK-ONLY"], variable=self.mode_var, font=font_label, fg_color="#2C2C2E", button_color="#2C2C2E", text_color="#0A84FF", dropdown_fg_color="#1C1C1E")
        self.mode_menu.grid(row=0, column=1, sticky="e")
        
        ctk.CTkFrame(group2, height=1, fg_color="#38383A").pack(fill="x", padx=15)
        
        # Penjabaran Column Fallback Row
        row_p = ctk.CTkFrame(group2, fg_color="transparent")
        row_p.pack(fill="x", padx=15, pady=8)
        row_p.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(row_p, text="Kolom Penjabaran", font=font_label, text_color="#FFFFFF", width=120, anchor="w").grid(row=0, column=0)
        self.entry_penjabaran = ctk.CTkEntry(row_p, placeholder_text="Jumlah Penjabaran, Jumlah...", **input_kwargs)
        self.entry_penjabaran.grid(row=0, column=1, sticky="ew")
        self.entry_penjabaran.insert(0, "Jumlah Penjabaran, Jumlah")
        
        ctk.CTkFrame(group2, height=1, fg_color="#38383A").pack(fill="x", padx=15)
        
        # Quick Check Row
        row_q = ctk.CTkFrame(group2, fg_color="transparent")
        row_q.pack(fill="x", padx=15, pady=8)
        row_q.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(row_q, text="Cek 1 Faktur", font=font_label, text_color="#FFFFFF", width=120, anchor="w").grid(row=0, column=0)
        self.entry_single_faktur = ctk.CTkEntry(row_q, placeholder_text="Nomor faktur...", **input_kwargs)
        self.entry_single_faktur.grid(row=0, column=1, sticky="ew")
        
        btn_cek_kwargs = {"font": font_link, "fg_color": "transparent", "text_color": "#FF9F0A", "hover_color": "#2C2C2E", "border_width": 1, "border_color": "#38383A", "corner_radius": 6}
        ctk.CTkButton(row_q, text="Cek", width=60, command=self.quick_check, **btn_cek_kwargs).grid(row=0, column=2, padx=(5,0))
        
        # ================= KOLOM KANAN: LOG TERMINAL =================
        right_panel = ctk.CTkFrame(self, fg_color="transparent")
        right_panel.grid(row=1, column=1, sticky="nsew", padx=(10, 20), pady=(0, 20))
        right_panel.grid_columnconfigure(0, weight=1)
        right_panel.grid_rowconfigure(1, weight=1)
        
        log_header = ctk.CTkFrame(right_panel, fg_color="transparent")
        log_header.grid(row=0, column=0, sticky="ew", pady=(0, 5))
        ctk.CTkLabel(log_header, text="SYSTEM LOG", font=("Segoe UI", 12), text_color="#8E8E93").pack(side="left")
        ctk.CTkButton(log_header, text="Clear", font=("Segoe UI", 12), fg_color="transparent", text_color="#FF3B30", width=40, hover_color="#1C1C1E", command=lambda: self.log_box.delete("0.0", "end")).pack(side="right")
        
        self.log_box = ctk.CTkTextbox(right_panel, font=("Consolas", 12), fg_color="#1C1C1E", text_color="#FFFFFF", corner_radius=10, border_width=1, border_color="#38383A")
        self.log_box.grid(row=1, column=0, sticky="nsew")
        
        # --- BOTTOM FIX ACTION (Spans across both columns) ---
        bottom_frame = ctk.CTkFrame(self, fg_color="transparent")
        bottom_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=20, pady=(0, 20))
        bottom_frame.grid_columnconfigure(0, weight=1)
        
        self.btn_run = ctk.CTkButton(bottom_frame, text="Mulai Pemrosesan", height=55, corner_radius=12, command=self.start_process, font=("Segoe UI", 16, "bold"), fg_color="#0A84FF", hover_color="#0056B3")
        self.btn_run.grid(row=0, column=0, sticky="ew")
        
        font_link_underline = ("Segoe UI", 14, "underline")
        ctk.CTkButton(bottom_frame, text="About ZiTax", font=font_link_underline, fg_color="transparent", text_color="#8E8E93", hover_color="#1C1C1E", command=self.show_about).grid(row=1, column=0, pady=(10,0))

    def show_about(self):
        AboutWindow(self)

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
                
                # Gunakan Chrome asli bawaan Windows untuk bypass blokir Antivirus & WAF DJP
                launch_args = [
                    "--disable-blink-features=AutomationControlled",
                    "--ignore-certificate-errors"
                ]
                
                try:
                    browser = p.chromium.launch(headless=False, channel="chrome", args=launch_args)
                except Exception:
                    # Fallback ke Edge atau Chromium bawaan jika Chrome tidak terinstall
                    try:
                        browser = p.chromium.launch(headless=False, channel="msedge", args=launch_args)
                    except Exception:
                        browser = p.chromium.launch(headless=False, args=launch_args)
                        
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
                    if not browser.is_connected() or len(context.pages) == 0: break
                    
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
            err_msg = str(e).upper()
            if "ERR_INTERNET_DISCONNECTED" in err_msg or "ERR_NAME_NOT_RESOLVED" in err_msg:
                self.add_log("[!] KONEKSI INTERNET MATI: Pastikan komputer Anda terhubung ke jaringan internet.")
            elif "ERR_CONNECTION_TIMED_OUT" in err_msg or "TIMEOUT" in err_msg:
                self.add_log("[!] KONEKSI TIMEOUT: Server CoreTax DJP lambat merespon atau sedang down.")
            elif "ERR_CONNECTION_RESET" in err_msg:
                self.add_log("[!] KONEKSI DIPUTUS (RESET): Akses diblokir oleh Firewall/Antivirus Anda atau server DJP sedang padat.")
            else:
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
            jenis = "PK" if "PK" in self.sheet_var.get() and "PM" not in self.sheet_var.get() else "PM"
            self.add_log(f"[*] Mengekstrak isi PDF sebagai {jenis}...")
            items, msg, inv = self._fetch_pdf_items(faktur, tid, 0, jenis)
            
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
            
            sheet_mode = self.sheet_var.get()
            sheet_targets = []
            if "PM" in sheet_mode: sheet_targets.append("PM")
            if "PK" in sheet_mode: sheet_targets.append("PK")

            sheets_to_process = []
            for target_jenis in sheet_targets:
                found_sheet = None
                for sn in wb.sheetnames:
                    if sn.upper() == target_jenis:
                        found_sheet = sn
                        break
                if not found_sheet:
                    for sn in wb.sheetnames:
                        if target_jenis in sn.upper():
                            found_sheet = sn
                            break
                if found_sheet:
                    sheets_to_process.append((found_sheet, target_jenis))
                else:
                    self.add_log(f"[!] Peringatan: Sheet untuk {target_jenis} tidak ditemukan.")

            if not sheets_to_process:
                self.add_log("[!] ERROR: Tidak menemukan sheet target yang sesuai.")
                self.btn_run.configure(state="normal", text="START PROCESS")
                return

            total_success = 0
            for target_sheet_name, jenis_pajak in sheets_to_process:
                self.add_log(f"\n{'='*40}\n[*] MEMULAI PEMROSESAN SHEET: {target_sheet_name} ({jenis_pajak})\n{'='*40}")
                success = self._process_single_sheet(wb, target_sheet_name, jenis_pajak, tid)
                total_success += success
            
            if total_success > 0 or self.mode_var.get() == "ONLY-CLEAR (Hapus Saja)":
                wb.save(self.file_path)
            
            self.add_log(f"\n[√] SELESAI: {total_success} faktur berhasil diproses dari semua sheet.")
            self.btn_run.configure(state="normal", text="START PROCESS")
        except Exception as e:
            self.add_log(f"[!] ERROR: {str(e)}")
            self.btn_run.configure(state="normal", text="START AUTO-CHECK")

    def _fetch_pdf_items(self, no_faktur, tid, expected_total, jenis_pajak="PM"):
        headers = {"authority": "coretaxdjp.pajak.go.id", "authorization": f"Bearer {self.token}", "content-type": "application/json", "cookie": self.cookie, "x-dgt-code": "7AcAAA=="}
        try:
            is_pk = (jenis_pajak == "PK")
            api_endpoint = "outputinvoice/list" if is_pk else "inputinvoice/list"
            s_url = f"https://coretaxdjp.pajak.go.id/einvoiceportal/api/{api_endpoint}"
            
            payload = {
                "TaxpayerAggregateIdentifier": tid, 
                "First": 0, "Rows": 1, "LanguageId": "id-ID", 
                "Filters": [{"PropertyName": "TaxInvoiceNumber", "Value": no_faktur, "MatchMode": "equals"}]
            }
            if is_pk:
                payload["SellerTaxpayerAggregateIdentifier"] = tid
            else:
                payload["BuyerTaxpayerAggregateIdentifier"] = tid

            resp = requests.post(s_url, headers=headers, json=payload, timeout=15).json()
            data = resp.get("Payload", {}).get("Data", [])
            if not data: return None, f"Faktur tidak ditemukan di server ({jenis_pajak})", None

            inv = data[0]
            d_url = "https://coretaxdjp.pajak.go.id/einvoiceportal/api/DownloadInvoice/download-invoice-document"
            # BUGS DJP CORETAX: PDF Download API always requires "Input" even for Output invoices (PK)!
            menu_type = "Input" 
            d_payload = {"EInvoiceRecordIdentifier": inv["RecordId"], "EInvoiceAggregateIdentifier": inv["AggregateIdentifier"], "DocumentAggregateIdentifier": inv["DocumentFormAggregateIdentifier"], "TaxpayerAggregateIdentifier": tid, "LetterNumber": no_faktur, "EInvoiceMenuType": menu_type, "TaxInvoiceStatus": "APPROVED"}
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
        except requests.exceptions.ConnectionError:
            return None, "Jaringan terputus atau server DJP sedang offline.", None
        except requests.exceptions.Timeout:
            return None, "Koneksi TIMEOUT: Server DJP lambat merespon (batas 15 detik).", None
        except Exception as e:
            return None, str(e), None

    def _process_single_sheet(self, wb, target_sheet_name, jenis_pajak, tid):
        try:
            ws = wb[target_sheet_name]
            # Konversi ke DataFrame dengan menyimpan NOMOR BARIS ASLI
            data_raw = list(ws.values)
            h_idx = 0
            for i, row in enumerate(data_raw):
                row_str = [str(v).upper() if v else "" for v in row]
                if any("NOMOR FAKTUR" in v for v in row_str):
                    h_idx = i
                    break
            
            cols = data_raw[h_idx]
            # Tambahkan kolom internal untuk melacak Row Excel Asli
            rows_with_meta = []
            for i, row in enumerate(data_raw[h_idx+1:]):
                row_list = list(row)
                row_list.append(i + h_idx + 2) # Ini nomor baris Excel asli (1-based)
                rows_with_meta.append(row_list)
            
            headers = [str(c).strip() if c else f"COL_{i}" for i, c in enumerate(cols)]
            df = pd.DataFrame(rows_with_meta, columns=headers + ["_EXCEL_ROW"])
            
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
            
            self.add_log(f"[+] Kategori terdeteksi di {target_sheet_name}: {', '.join(self.dynamic_categories.keys())}")
            
            # Ambil index kolom penting
            c_faktur = next((c for c in headers if "NOMOR FAKTUR" in str(c).upper()), None)
            c_harga  = next((c for c in headers if "HARGA JUAL" in str(c).upper() or "HARGA" in str(c).upper()), None)
            
            # Cek fallback dari input GUI
            fallback_input = self.entry_penjabaran.get().strip()
            fallbacks = [f.strip().upper() for f in fallback_input.split(",") if f.strip()]
            if not fallbacks:
                fallbacks = ["JUMLAH PENJABARAN", "PENJABARAN", "JUMLAH"]
                
            c_penjabaran = None
            # 1. Coba exact match dahulu
            for fb_name in fallbacks:
                c_penjabaran = next((c for c in headers if str(c).upper().strip() == fb_name), None)
                if c_penjabaran:
                    break
            
            # 2. Coba partial/fuzzy match jika exact match tidak ditemukan
            if not c_penjabaran:
                for fb_name in fallbacks:
                    c_penjabaran = next((c for c in headers if fb_name in str(c).upper().strip()), None)
                    if c_penjabaran:
                        break
            
            if not c_faktur:
                self.add_log(f"[!] ERROR: Kolom 'Nomor Faktur' tidak ditemukan di {target_sheet_name}.")
                return 0

            if not c_penjabaran:
                self.add_log(f"[!] ERROR: Kolom Penjabaran (dari fallback: {fallback_input}) tidak ditemukan di {target_sheet_name}.")
                return 0

            def is_truly_empty(v):
                if v is None or pd.isna(v): return True
                s_v = str(v).strip().lower()
                if s_v == "" or s_v == "0" or s_v == "0.0" or s_v == "nan": return True
                return False

            def should_process(r):
                if pd.isna(r[c_faktur]): return False
                has_any_real_entry = False
                for cat in self.dynamic_categories.keys():
                    if not is_truly_empty(r[cat]):
                        has_any_real_entry = True
                        break
                pj = r[c_penjabaran]
                if has_any_real_entry: return False
                if pj is not None:
                    s_pj = str(pj).strip()
                    if s_pj.startswith("="): return True
                    try:
                        f_pj = float(s_pj.replace(",",""))
                        if f_pj > 0: return False
                    except: pass
                return True

            targets = df[df.apply(should_process, axis=1)]
            
            if targets.empty:
                self.add_log(f"[√] Seluruh baris target di {target_sheet_name} sudah BERSIH.")
                return 0
            else:
                self.add_log(f"[+] Ditemukan {len(targets)} baris baru untuk diproses di {target_sheet_name}.")

            if self.mode_var.get() != "CHECK-ONLY (PDF Saja)":
                self.add_log(f"[*] Tahap 1: Membersihkan {len(targets)} baris...")
                for _, row in targets.iterrows():
                    ex_row = int(row["_EXCEL_ROW"])
                    for cat_name, (p_idx, q_idx) in self.dynamic_categories.items():
                        ws.cell(row=ex_row, column=p_idx).value = None
                        ws.cell(row=ex_row, column=q_idx).value = None

            if self.mode_var.get() == "ONLY-CLEAR (Hapus Saja)":
                return 0

            all_invoice_data = {}
            all_unique_names = set()
            
            self.add_log(f"[*] Tahap 2: Men-scan SELURUH PDF (Global Scan) {target_sheet_name}...")
            for index, row in targets.iterrows():
                raw_f = row[c_faktur]
                expected_total = pd.to_numeric(row[c_harga], errors='coerce') if c_harga else 0
                ex_row = int(row["_EXCEL_ROW"])
                faktur = "{:.0f}".format(raw_f).zfill(17) if isinstance(raw_f, (float, int)) else str(raw_f).zfill(17)
                
                self.add_log(f"    - Scanning PDF Baris {ex_row}: {faktur}")
                items, msg, raw_json = self._fetch_pdf_items(faktur, tid, expected_total, jenis_pajak)
                if items:
                    self.add_log(f"--- [RAW DATA DARI SERVER: {faktur}] ---")
                    all_invoice_data[ex_row] = items
                    for it in items: 
                        all_unique_names.add(it['name'])
                        self.add_log(f"      > Item: {it['name']} | Qty: {it['qty']} | Total: {it['total']:,}")
                else:
                    self.add_log(f"      [!] Skip: {msg}")

            if not all_invoice_data:
                self.add_log("[!] Scan selesai, tidak ada data yang ditemukan.")
                return 0

            self.add_log(f"[?] Menunggu validasi mapping untuk {len(all_unique_names)} barang ({jenis_pajak})...")
            mapping_window = MappingWindow(self, list(all_unique_names), list(self.dynamic_categories.keys()))
            self.wait_window(mapping_window)
            
            final_mapping = mapping_window.result
            
            if not final_mapping:
                self.add_log("[!] Mapping dibatalkan.")
                return 0

            self.add_log("[*] Tahap 4: Mengisi data dengan Logika Akumulasi...")
            success_count = 0
            
            for ex_row, items in all_invoice_data.items():
                category_totals = {}
                for it in items:
                    cat = final_mapping.get(it['name'])
                    if cat and cat != "Abaikan":
                        if cat not in category_totals: category_totals[cat] = {'qty': 0, 'total': 0}
                        category_totals[cat]['qty'] += it['qty']
                        category_totals[cat]['total'] += it['total']
                
                if not category_totals:
                    self.add_log(f"      [!] Baris {ex_row} | Semua item diabaikan (Tidak masuk rekap).")
                else:
                    for cat, val in category_totals.items():
                        if cat in self.dynamic_categories:
                            p_idx, q_idx = self.dynamic_categories[cat]
                            ws.cell(row=ex_row, column=p_idx).value = val['total']
                            ws.cell(row=ex_row, column=p_idx).number_format = '#,##0'
                            ws.cell(row=ex_row, column=q_idx).value = val['qty']
                            ws.cell(row=ex_row, column=q_idx).number_format = '#,##0.00'
                            self.add_log(f"      [√] Baris {ex_row} | {cat}: Total={val['total']:,}")
                            
                cat_cols = [p for p, q in self.dynamic_categories.values()]
                if cat_cols:
                    min_l = openpyxl.utils.get_column_letter(min(cat_cols))
                    max_l = openpyxl.utils.get_column_letter(max(cat_cols))
                    pj_idx = headers.index(c_penjabaran) + 1 if c_penjabaran else None
                    if pj_idx:
                        ws.cell(row=ex_row, column=pj_idx).value = f"=SUM({min_l}{ex_row}:{max_l}{ex_row})"
                        ws.cell(row=ex_row, column=pj_idx).number_format = '#,##0'

                c_selisih = next((c for c in headers if "SELISIH" in str(c).upper()), None)
                diff_idx = headers.index(c_selisih) + 1 if c_selisih else None
                if diff_idx: ws.cell(row=ex_row, column=diff_idx).value = "-"
                success_count += 1

            return success_count
        except Exception as e:
            self.add_log(f"[!] ERROR di sheet {target_sheet_name}: {str(e)}")
            return 0


if __name__ == "__main__":
    app = CoreTaxApp()
    app.mainloop()

