import re

path = r'c:\Work\project automation\coretax\coretax_app.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Add import difflib at the top if not exists
if 'import difflib' not in content:
    content = "import difflib\n" + content

# 2. Replace the whole MappingWindow class using a more robust replacement
new_class = r"""class MappingWindow(ctk.CTkToplevel):
    def __init__(self, parent, unique_names, categories):
        super().__init__(parent)
        self.title("Batch Mapping: Tentukan Kategori Barang")
        self.geometry("750x650")
        self.result = {}
        self.unique_names = unique_names
        self.categories = ["Abaikan"] + categories
        self.combo_vars = {}
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        # Header & Tombol Cerdas
        header_frame = ctk.CTkFrame(self)
        header_frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        
        ctk.CTkLabel(header_frame, text="Ditemukan beberapa jenis barang baru.\nSilakan tentukan kolom Excel untuk setiap nama barang:", font=("Arial", 12, "bold")).pack(side="left", padx=10)
        
        self.btn_auto = ctk.CTkButton(header_frame, text="AUTO-MAP (SMART)", fg_color="#E67E22", hover_color="#D35400", command=self.auto_map)
        self.btn_auto.pack(side="right", padx=10)

        # Scrollable area
        self.scroll_frame = ctk.CTkScrollableFrame(self)
        self.scroll_frame.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        self.scroll_frame.grid_columnconfigure(0, weight=1)

        for i, name in enumerate(self.unique_names):
            row_f = ctk.CTkFrame(self.scroll_frame)
            row_f.grid(row=i, column=0, padx=5, pady=5, sticky="ew")
            
            ctk.CTkLabel(row_f, text=f"Nama : {name}", wraplength=450, justify="left").pack(side="left", padx=10)
            
            var = ctk.StringVar(value="Abaikan")
            combo = ctk.CTkComboBox(row_f, values=self.categories, variable=var, width=200)
            combo.pack(side="right", padx=10)
            self.combo_vars[name] = var

        self.btn_save = ctk.CTkButton(self, text="SIMPAN & PROSES SEMUA", fg_color="#27AE60", hover_color="#219150", command=self.save_mapping)
        self.btn_save.grid(row=2, column=0, padx=20, pady=20)

    def auto_map(self):
        # Logika Fuzzy Matching
        for name, var in self.combo_vars.items():
            real_cats = self.categories[1:]
            # 1. Cek exact match atau sangat mirip
            matches = difflib.get_close_matches(name.upper(), [c.upper() for c in real_cats], n=1, cutoff=0.5)
            
            if matches:
                idx = [c.upper() for c in real_cats].index(matches[0])
                var.set(real_cats[idx])
            else:
                # 2. Cek apakah ada kata dari Kategori yang muncul di Nama Barang
                found = False
                for cat in real_cats:
                    if cat.upper() in name.upper() or name.upper() in cat.upper():
                        var.set(cat)
                        found = True
                        break
                
                if not found:
                    # 3. Ratio SequenceMatcher
                    best_match = "Abaikan"
                    best_score = 0
                    for cat in real_cats:
                        score = difflib.SequenceMatcher(None, cat.upper(), name.upper()).ratio()
                        if score > best_score and score > 0.3:
                            best_score = score
                            best_match = cat
                    var.set(best_match)

    def save_mapping(self):
        for name, var in self.combo_vars.items():
            self.result[name] = var.get()
        self.destroy()"""

# Replacement without using re.sub for the replacement string to avoid backslash issues
start_marker = 'class MappingWindow(ctk.CTkToplevel):'
end_marker = 'self.destroy()'

start_idx = content.find(start_marker)
end_idx = content.find(end_marker, start_idx) + len(end_marker)

if start_idx != -1 and end_idx != -1:
    new_content = content[:start_idx] + new_class + content[end_idx:]
    with open(path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    print("Patch applied successfully")
else:
    print(f"Markers not found: start={start_idx}, end={end_idx}")
