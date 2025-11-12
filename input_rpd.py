import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import messagebox


class InputRPD:
    def __init__(self, root):
        self.root = root
        self.root.title("Input Data RPD")
        self.root.geometry("1600x800")  # Ukuran window lebih besar
        self.root.style = ttk.Style("flatly")

        header = ttk.Label(
            root,
            text="🧾 Input Data RPD",
            font=("Segoe UI", 24, "bold"),
            bootstyle="primary",
        )
        header.pack(pady=15)

        # FRAME TABEL
        frame = ttk.Frame(root, padding=10)
        frame.pack(fill=BOTH, expand=True)

        # SCROLLBAR
        x_scroll = ttk.Scrollbar(frame, orient=HORIZONTAL)
        y_scroll = ttk.Scrollbar(frame, orient=VERTICAL)
        x_scroll.pack(side=BOTTOM, fill=X)
        y_scroll.pack(side=RIGHT, fill=Y)

        # KOLUMNYA
        self.columns = [
            "uraian", "pagu", "realisasi_sdlalu", "sisa_pagu",
            "realisasi_ls", "gup_kkp1", "gup_kkp2", "gup_kkp3",
            "realisasi_gup1", "realisasi_gup2", "realisasi_gup3",
            "sisa_saldo", "total_realisasi",
        ]
        headings = [
            "Uraian", "PAGU", "Realisasi S/D Lalu", "Sisa PAGU S/D Lalu",
            "Realisasi LS", "Realisasi GUP KKP 1", "Realisasi GUP KKP 2", "Realisasi GUP KKP 3",
            "Realisasi GUP 1", "Realisasi GUP 2", "Realisasi GUP 3",
            "Sisa Saldo", "Total Realisasi",
        ]

        # STYLE TABLE
        style = ttk.Style()
        style.configure("Treeview", font=("Segoe UI", 11), rowheight=60, padding=8)  # Tinggi baris lebih besar
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"))  # Font header lebih besar
        style.map("Treeview",
                  background=[("selected", "#4CAF50")],
                  foreground=[("selected", "white")])

        # TREEVIEW
        self.tree = ttk.Treeview(
            frame,
            columns=self.columns,
            show="headings",
            xscrollcommand=x_scroll.set,
            yscrollcommand=y_scroll.set,
            bootstyle="info",
        )
        self.tree.pack(fill=BOTH, expand=True)
        x_scroll.config(command=self.tree.xview)
        y_scroll.config(command=self.tree.yview)

        # Set column widths - wider for "Uraian", appropriate for others
        column_widths = {
            "uraian": 350,  # Lebih lebar untuk teks deskripsi
            "pagu": 140,
            "realisasi_sdlalu": 180,  # Lebih lebar untuk header panjang
            "sisa_pagu": 180,  # Lebih lebar untuk header panjang
            "realisasi_ls": 140,
            "gup_kkp1": 180,  # Lebih lebar untuk header panjang
            "gup_kkp2": 180,  # Lebih lebar untuk header panjang
            "gup_kkp3": 180,  # Lebih lebar untuk header panjang
            "realisasi_gup1": 150,
            "realisasi_gup2": 150,
            "realisasi_gup3": 150,
            "sisa_saldo": 140,
            "total_realisasi": 150,
        }

        for col, head in zip(self.columns, headings):
            self.tree.heading(col, text=head)
            width = column_widths.get(col, 160)
            # Use LEFT anchor for "Uraian" column, CENTER for others
            anchor = "w" if col == "uraian" else CENTER
            self.tree.column(col, width=width, anchor=anchor, minwidth=120, stretch=True)  # minwidth lebih besar

        # EVENT
        self.tree.bind("<Double-1>", self.edit_cell)

        # FRAME TOMBOL
        button_frame = ttk.Frame(root, padding=10)
        button_frame.pack(fill=X, side="bottom")

        ttk.Button(button_frame, text="💾 Simpan Data", width=18,
           bootstyle="primary", command=self.save_data).pack(side=LEFT, padx=5)

        ttk.Button(button_frame, text="📂 Import Excel", width=18,
                bootstyle="info", command=self.import_excel).pack(side=LEFT, padx=5)

        ttk.Button(button_frame, text="↻ Refresh", width=18,
                bootstyle="secondary", command=self.refresh).pack(side=LEFT, padx=5)
        self.update_total_row()

    # ===== TAMBAH BARIS =====
    def add_and_edit_row(self):
        new_row = self.tree.insert("", "end", values=["" for _ in self.columns])
        self.root.after(200, lambda: self.start_edit(new_row, 0))
        self.update_total_row()

    # ===== HAPUS BARIS =====
    def delete_row(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Peringatan", "Pilih baris yang ingin dihapus!")
            return
        for sel in selected:
            self.tree.delete(sel)
        self.update_total_row()

    # ===== SIMPAN =====
    def save_data(self):
        data = [self.tree.item(row)["values"] for row in self.tree.get_children()]
        messagebox.showinfo("Sukses", f"{len(data)} baris data berhasil disimpan!")

    # ===== REFRESH =====
    def refresh(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.update_total_row()

    def import_excel(self):
    # """Impor data dari file Excel"""
        import pandas as pd
        from tkinter import filedialog

        file_path = filedialog.askopenfilename(
            title="Pilih File Excel",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)
            df.columns = [str(col).lower().strip().replace(" ", "_").replace("/","_").replace(".","") for col in df.columns]

            mapping = {
                "uraian": ["uraian", "nama_kegiatan", "item", "deskripsi"],
                "pagu": ["pagu", "anggaran", "alokasi"],
                "realisasi_sdlalu": ["realisasi_s_d_lalu", "realisasi_sd_lalu", "realisasi_sdlalu", "realisasi sd lalu", "realisasi.s.d.lalu"],
                "sisa_pagu": ["sisa_pagu_s_d_lalu", "sisa_pagu_sd_lalu", "sisa_pagu", "sisa pagu sd lalu", "sisa anggaran"],
                "realisasi_ls": ["realisasi_ls", "ls"],
                "gup_kkp1": ["gup_kkp_1", "realisasi_gup_kkp1", "gup1", "realisasi gup kkp i"],
                "gup_kkp2": ["gup_kkp_2", "realisasi_gup_kkp2", "gup2", "realisasi gup kkp ii"],
                "gup_kkp3": ["gup_kkp_3", "realisasi_gup_kkp3", "gup3", "realisasi gup kkp iii"],
                "realisasi_gup1": ["realisasi_gup_1", "realisasi_gup1", "realisasi gup i"],
                "realisasi_gup2": ["realisasi_gup_2", "realisasi_gup2", "realisasi gup ii"],
                "realisasi_gup3": ["realisasi_gup_3", "realisasi_gup3", "realisasi gup iii"],
                "sisa_saldo": ["sisa_saldo", "saldo_akhir", "saldo akhir"],
                "total_realisasi": ["total_realisasi", "total", "jumlah", "grand_total"],
            }


            # buat kolom kosong jika ditemukan
            data = {}
            for key, aliases in mapping.items():
                found = next((col for col in df.columns if col in aliases), None)
                if found:
                    data[key] = df[found]
                else:
                    data[key] = ["" for _ in range(len(df))]
            new_df = pd.DataFrame(data)

            for row in self.tree.get_children():
                self.tree.delete(row)
            
            # Tambahkan data ke tabel
            for _, row in df.iterrows():
                values = [row.get(col, "") for col in self.columns]
                self.tree.insert("", "end", values=values)

            self.update_total_row()
            messagebox.showinfo("Sukses", "Data dari Excel berhasil diimpor!")

        except Exception as e:
            messagebox.showerror("Error", f"Gagal membaca file Excel:\n{e}")


    # ===== EDIT =====
    def edit_cell(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        col_index = int(column.replace("#", "")) - 1
        self.start_edit(row_id, col_index)

    def start_edit(self, row_id, col_index):
        """Edit cell sesuai tinggi baris"""
        bbox = self.tree.bbox(row_id, f"#{col_index+1}")
        if not bbox:
            return
        x, y, width, height = bbox
        value = self.tree.set(row_id, self.columns[col_index])

        # Use LEFT justify for "Uraian" column, CENTER for others
        justify_style = "left" if self.columns[col_index] == "uraian" else "center"
        
        entry = ttk.Entry(self.tree, font=("Segoe UI", 11), justify=justify_style)  # Font lebih besar
        entry.place(x=x + 1, y=y + 1, width=width - 2, height=height - 2)
        entry.insert(0, value)
        entry.focus()

        def save_edit(event=None):
            new_value = entry.get()
            self.tree.set(row_id, self.columns[col_index], new_value)
            entry.destroy()

            # 🔥 Hitung otomatis
            self.calculate_auto_values(row_id)
            self.update_total_row()

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)

    # ===== PERHITUNGAN OTOMATIS =====
    def calculate_auto_values(self, row_id):
        try:
            pagu = float(self.tree.set(row_id, "pagu") or 0)
            realisasi_lalu = float(self.tree.set(row_id, "realisasi_sdlalu") or 0)
            realisasi_ls = float(self.tree.set(row_id, "realisasi_ls") or 0)
            gup1 = float(self.tree.set(row_id, "gup_kkp1") or 0)
            gup2 = float(self.tree.set(row_id, "gup_kkp2") or 0)
            gup3 = float(self.tree.set(row_id, "gup_kkp3") or 0)
            rg1 = float(self.tree.set(row_id, "realisasi_gup1") or 0)
            rg2 = float(self.tree.set(row_id, "realisasi_gup2") or 0)
            rg3 = float(self.tree.set(row_id, "realisasi_gup3") or 0)
        except ValueError:
            return

        sisa_pagu = pagu - realisasi_lalu
        total_realisasi = realisasi_ls + gup1 + gup2 + gup3 + rg1 + rg2 + rg3
        sisa_saldo = sisa_pagu - total_realisasi

        self.tree.set(row_id, "sisa_pagu", f"{sisa_pagu:.0f}")
        self.tree.set(row_id, "total_realisasi", f"{total_realisasi:.0f}")
        self.tree.set(row_id, "sisa_saldo", f"{sisa_saldo:.0f}")

    # ===== TOTAL =====
    def update_total_row(self):
        # hapus baris total lama
        for row in self.tree.get_children():
            values = self.tree.item(row)["values"]
            if values and str(values[0]).startswith("Total Rencana"):
                self.tree.delete(row)

        kolom_dijumlah = [
            "pagu", "realisasi_sdlalu", "sisa_pagu",
            "realisasi_ls", "gup_kkp1", "gup_kkp2", "gup_kkp3",
            "realisasi_gup1", "realisasi_gup2", "realisasi_gup3",
            "total_realisasi"
        ]
        total = {col: 0 for col in kolom_dijumlah}

        # hitung total
        for row in self.tree.get_children():
            values = self.tree.item(row)["values"]
            if not values or str(values[0]).startswith("Total Rencana"):
                continue
            for i, col in enumerate(self.columns):
                if col in total:
                    try:
                        total[col] += float(values[i])
                    except:
                        pass

        total_values = ["" for _ in self.columns]
        total_values[0] = "Total Rencana Pencairan Peralatan dan Mesin"
        for i, col in enumerate(self.columns):
            if col in total:
                total_values[i] = f"{total[col]:,.0f}"

        total_row = self.tree.insert("", "end", values=total_values, tags=("totalrow",))
        self.tree.tag_configure(
            "totalrow",
            background="#E3F2FD",
            font=("Segoe UI", 11, "bold"),  # Font lebih besar
            foreground="#000000",
        )
        self.tree.see(total_row)


if __name__ == "__main__":
    root = ttk.Window(themename="flatly")
    InputRPD(root)
    root.mainloop()
