from math import e
import tkinter as tk
from tkinter import Menu, ttk, messagebox
from tkcalendar import DateEntry
import pyodbc
import subprocess
import sys
from openpyxl import Workbook
from tkinter import filedialog
# ========================================== KẾT NỐI CƠ SỞ DỮ LIỆU ==================================
conn = pyodbc.connect(
    'DRIVER={SQL Server};'
    'SERVER=ADMIN-PC\\SQLEXPRESS;'
    'DATABASE=QLHS;'
    'Trusted_Connection=yes;'
)
cur = conn.cursor()


def center_window(win, w=700, h=600):
    ws = win.winfo_screenwidth()
    hs = win.winfo_screenheight()
    x = (ws // 2) - (w // 2)
    y = (hs // 2) - (h // 2)
    win.geometry(f'{w}x{h}+{x}+{y}')
# ========================================== TẠO CỬA SỐ CHÍNH ==================================
root = tk.Tk()
root.title("Quản lý điểm học sinh")
center_window(root, 700, 600)
root.resizable(False, False)


# ========================== TẠO CỘT ĐIỂM TRUNG BÌNH VÀ CỘT XẾP LOẠI NẾU CHƯA CÓ =============================
def them_cot(): 
    # Lấy tất cả tên cột của bảng Diem 
    cur.execute(""" SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'Diem' """) 
    columns = [row[0].lower() for row in cur.fetchall()] 
    
    # Thêm cột nếu chưa tồn tại 
    if 'diemtb' not in columns: 
        cur.execute("ALTER TABLE Diem ADD DiemTB FLOAT") 
        print("Đã thêm cột DiemTB") 
    else: 
        print("Cột DiemTB đã tồn tại") 
    if 'xeploai' not in columns: 
        cur.execute("ALTER TABLE Diem ADD XepLoai NVARCHAR(10)") 
        print("Đã thêm cột XepLoai") 
    else: 
        print("Cột XepLoai đã tồn tại") 
    conn.commit() 
    
them_cot()



# =========================================== MENU ==============================================
def mo_form_dangnhap():
    root.withdraw()
    try:
        subprocess.Popen([sys.executable, r"E:\DoAn_QLHS\QLHS\Login.py"])
        root.destroy()  # đóng form hiện tại hoàn toàn
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể mở form đăng nhập:\n{e}")

menubar = tk.Menu(root)


# ----- Menu Quản lý -----
menu_quanly = tk.Menu(menubar, tearoff=0)
menu_quanly.add_command(label="Quản lý điểm", font=("Times New Roman", 11), 
                        command=lambda: messagebox.showinfo("Thông báo", "Bạn đang ở trang Quản lý điểm"))
menu_quanly.add_command(label="Đăng xuất",font=("Times New Roman", 11), 
                        command=lambda: mo_form_dangnhap())
menubar.add_cascade(label="Quản lý", font=("Times New Roman", 11), menu=menu_quanly)


root.config(menu=menubar)
# ====== Style Treeview ======
style = ttk.Style()
style.configure("Treeview", rowheight=25, font=("Times New Roman", 11))

# ====== Label tiêu đề ======
lbl_tieu_de = tk.Label(root, text="QUẢN LÝ ĐIỂM", font=("Times New Roman", 24, "bold")  )
lbl_tieu_de.pack(pady=10)
# ====== Frame nhập thông tin ====== 
frame_tt = tk.Frame(root) 
frame_tt.pack(pady=5, padx=10, fill="x")

tk.Label(frame_tt, text="Mã lớp", font=("Times New Roman", 13),bg="#F5F5F5").grid(row=0, column=0, padx=5, pady=5, sticky="w")
entry_malop = tk.Entry(frame_tt, width=20)
entry_malop.grid(row=0, column=1, padx=5, pady=5, sticky="w")


# Đặt khoảng cách lớn giữa entry_maso và label "Lớp"
frame_tt.grid_columnconfigure(2, minsize=100)

tk.Label(frame_tt, text="Lớp", font=("Times New Roman", 13),bg="#F5F5F5").grid(row=0, column=3, padx=5, pady=5, sticky="w")
cbb_lop = ttk.Combobox(frame_tt, values=["10A1", "10A2", "10A3", "10A4", "10A5", "10A6"], width=15)
cbb_lop.grid(row=0, column=4, padx=5, pady=5, sticky="w")

lop_duoc_chon = ""

#  Bản đồ lớp - mã lớp
lop_to_malop = {
    "10A1": "L01",
    "10A2": "L02",
    "10A3": "L03",
    "10A4": "L04",
    "10A5": "L05",
    "10A6": "L06"
}

# Tạo bản đồ ngược lại (Mã lớp → Lớp)
malop_to_lop = {v: k for k, v in lop_to_malop.items()}

# Khi chọn lớp, tự điền mã lớp 
def capnhat_malop(event=None):
    ma_lop = entry_malop.get().strip().upper()
    lop = malop_to_lop.get(ma_lop, "")
    if lop:
        cbb_lop.set(lop)
        global lop_duoc_chon
        lop_duoc_chon = lop
        load_dulieu()  # Tự động lọc khi nhập mã lớp

def on_lop_duoc_chon(event=None):
    global lop_duoc_chon
    lop_duoc_chon = cbb_lop.get()
    # Cập nhật mã lớp
    ma_lop = lop_to_malop.get(lop_duoc_chon, "")
    entry_malop.delete(0, tk.END)
    entry_malop.insert(0, ma_lop)
    load_dulieu()  # Load danh sách theo lớp

# Gán sự kiện
cbb_lop.bind("<<ComboboxSelected>>", on_lop_duoc_chon)

def capnhat_so_luong():
    """Cập nhật số lượng học sinh đang hiển thị"""
    if lop_duoc_chon:
        # Đếm số học sinh theo lớp
        cur.execute("SELECT COUNT(*) FROM HocSinh WHERE MaLop = ?", (lop_to_malop[lop_duoc_chon],))
        so_luong = cur.fetchone()[0]
        label_Soluong.config(text=f"Số lượng học sinh lớp {lop_duoc_chon}: {so_luong}")
    else:
        # Đếm tổng số học sinh (khi chưa chọn lớp)
        cur.execute("SELECT COUNT(*) FROM HocSinh")
        so_luong = cur.fetchone()[0]
        label_Soluong.config(text=f"Tổng số học sinh: {so_luong}")


# ================================================= CÁC HÀM CHỨC NĂNG =============================================
def kiem_tra_diem(P):
    if P == "":  # cho phép xoá hết để nhập lại
        return True
    try:
        value = float(P)
        return 0 <= value <= 10
    except ValueError:
        return False


vcmd = (root.register(kiem_tra_diem), '%P')

def mo_form_diemmoi(maso, hoten):
    form = tk.Toplevel(root)
    form.title(f"Nhập điểm cho {hoten}")
    form.geometry("600x500")  # Tăng kích thước form để chứa nhiều môn
    form.grab_set()
    tk.Label(form, text=f"Học sinh: {hoten} ({maso})", font=("Times New Roman", 13, "bold")).pack(pady=10)
    frame = tk.Frame(form)
    frame.pack(pady=10)
    # HK1 và HK2 cho từng môn
    tk.Label(frame, text="Môn", font=("Times New Roman", 11)).grid(row=0, column=0)
    tk.Label(frame, text="HKI", font=("Times New Roman", 11)).grid(row=0, column=1)
    tk.Label(frame, text="HKII", font=("Times New Roman", 11)).grid(row=0, column=2)
    # Danh sách môn học
    subjects = [
        ("Toán", "M01"), ("Văn", "M02"), ("Anh", "M03"),
        ("Lý", "M04"), ("Hóa", "M05"), ("Sinh", "M06"),
        ("Sử", "M07"), ("Địa", "M08"), ("GDCD", "M09")
    ]
    entries = {}
    for i, (subject, _) in enumerate(subjects, start=1):
        tk.Label(frame, text=subject, font=("Times New Roman", 11)).grid(row=i, column=0)
        entries[f"{subject}_hk1"] = tk.Entry(frame, width=10, validate="key", validatecommand=vcmd)
        entries[f"{subject}_hk1"].grid(row=i, column=1)
        entries[f"{subject}_hk2"] = tk.Entry(frame, width=10, validate="key", validatecommand=vcmd)
        entries[f"{subject}_hk2"].grid(row=i, column=2)
    def save_score():
        hoi = messagebox.askyesno("Xác nhận", "Bạn có chắc muốn lưu điểm mới này ?")
        if not hoi:
            return
        try:
            # Xóa điểm cũ (nếu có)
            cur.execute("DELETE FROM Diem WHERE MaHS=?", (maso,))
            # Thêm điểm mới cho tất cả môn
            for subject, ma_mon in subjects:
                hk1 = entries[f"{subject}_hk1"].get()
                hk2 = entries[f"{subject}_hk2"].get()
                # Chỉ thêm nếu có điểm
                if hk1 and hk2:
                    cur.execute("INSERT INTO Diem (MaHS, MaMon, DiemHK1, DiemHK2) VALUES (?, ?, ?, ?)",
                                (maso, ma_mon, float(hk1), float(hk2)))
            conn.commit()
            load_dulieu()
            messagebox.showinfo("Thành công", "Cập nhật điểm thành công.")
            form.destroy()
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể lưu điểm:\n{e}")
    tk.Button(form,
              text="Lưu điểm",
              font=("Times New Roman", 11),
              width=15,
              relief="ridge",
              cursor="hand2",
              command=save_score).pack(padx=(32, 0), pady=10)
       

def sua_diem():
    selected = tree.focus()
    if not selected:
        messagebox.showwarning("Thông báo", "Vui lòng chọn học sinh cần sửa điểm!")
        return
    hoi = messagebox.askyesno("Xác nhận", "Bạn có chắc muốn sửa điểm cho học sinh này không?")
    if hoi: 
        values = tree.item(selected, "values")
        maso = values[0]
        hoten = values[1]
        mo_form_diemmoi(maso, hoten)
    else:
        return

def tinhdiem():
    for item in tree.get_children(): # Lấy tất cả item
        values = tree.item(item, "values")
        MaHS = values[0]
        try:
            # Lấy điểm HKI và HKII cho 9 môn
            HKI_scores = [float(values[i]) for i in range(2, 11) if values[i]]  # HKI: Toán, Văn, Anh, Lý, Hóa, Sinh, Sử, Địa, GDCD
            HKII_scores = [float(values[i]) for i in range(11, 20) if values[i]]  # HKII: Toán, Văn, Anh, Lý, Hóa, Sinh, Sử, Địa, GDCD
            if HKI_scores and HKII_scores:  # Chỉ tính nếu có điểm
                DiemTB = round((sum(HKI_scores) + sum(HKII_scores)) / (len(HKI_scores) + len(HKII_scores)), 2)
                if DiemTB >= 8.0:
                    XepLoai = "Giỏi"
                elif DiemTB >= 6.5:
                    XepLoai = "Khá"
                elif DiemTB >= 5.0:
                    XepLoai = "Trung Bình"
                else:
                    XepLoai = "Yếu"
                cur.execute("""
                    UPDATE Diem
                    SET DiemTB=?, XepLoai=?
                    WHERE MaHS=?
                """, (DiemTB, XepLoai, MaHS))
        except ValueError:
            continue # Nếu chưa có điểm, bỏ qua
    conn.commit()
    load_dulieu()
    messagebox.showinfo("Thông báo", "Tính Điểm Trung Bình 2 Học Kỳ và Xếp Loại thành công!")
# Xuất dữ liệu ra file Excel
def xuat_excel():
    if not lop_duoc_chon:
        messagebox.showwarning("Thông báo", "Vui lòng chọn lớp để in bảng điểm!")
        return
    hoi = messagebox.askyesno("Xác nhận", f"Bạn có chắc muốn xuất bảng điểm lớp {lop_duoc_chon} ra file Excel không?")
    if not hoi:
        return
    rows = [tree.item(item, "values") for item in tree.get_children()]
    if not rows:
        messagebox.showinfo("Thông báo", "Không có dữ liệu để xuất.")
        return
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Lưu bảng điểm",
        initialfile=f"BangDiem_{lop_duoc_chon}.xlsx"
    )
    if not file_path:
        return
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = f"Bảng điểm {lop_duoc_chon}"
        headers = ["Mã HS", "Họ Tên", "HKI_Toán", "HKI_Văn", "HKI_Anh", "HKI_Lý", "HKI_Hóa", "HKI_Sinh", "HKI_Sử", "HKI_Địa", "HKI_GDCD",
                   "HKII_Toán", "HKII_Văn", "HKII_Anh", "HKII_Lý", "HKII_Hóa", "HKII_Sinh", "HKII_Sử", "HKII_Địa", "HKII_GDCD", "Điểm TB", "Xếp Loại"]
        ws.append(headers)
        for row in rows:
            ws.append(row)
        for col in ws.columns:
            max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
            ws.column_dimensions[col[0].column_letter].width = max_length + 3
        wb.save(file_path)
        messagebox.showinfo("Thành công", f"Đã xuất bảng điểm lớp {lop_duoc_chon} ra file Excel thành công!")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể xuất file Excel:\n{e}")


def load_dulieu():
    for i in tree.get_children():
        tree.delete(i)
    if lop_duoc_chon: # Nếu chọn lớp thì lọc
        cur.execute("""
            SELECT hs.MaHS, hs.HoTen,
                   MAX(CASE WHEN d.MaMon='M01' THEN d.DiemHK1 END) AS HKI_Toan,
                   MAX(CASE WHEN d.MaMon='M02' THEN d.DiemHK1 END) AS HKI_Van,
                   MAX(CASE WHEN d.MaMon='M03' THEN d.DiemHK1 END) AS HKI_Anh,
                   MAX(CASE WHEN d.MaMon='M04' THEN d.DiemHK1 END) AS HKI_Ly,
                   MAX(CASE WHEN d.MaMon='M05' THEN d.DiemHK1 END) AS HKI_Hoa,
                   MAX(CASE WHEN d.MaMon='M06' THEN d.DiemHK1 END) AS HKI_Sinh,
                   MAX(CASE WHEN d.MaMon='M07' THEN d.DiemHK1 END) AS HKI_Su,
                   MAX(CASE WHEN d.MaMon='M08' THEN d.DiemHK1 END) AS HKI_Dia,
                   MAX(CASE WHEN d.MaMon='M09' THEN d.DiemHK1 END) AS HKI_GDCD,
                   MAX(CASE WHEN d.MaMon='M01' THEN d.DiemHK2 END) AS HKII_Toan,
                   MAX(CASE WHEN d.MaMon='M02' THEN d.DiemHK2 END) AS HKII_Van,
                   MAX(CASE WHEN d.MaMon='M03' THEN d.DiemHK2 END) AS HKII_Anh,
                   MAX(CASE WHEN d.MaMon='M04' THEN d.DiemHK2 END) AS HKII_Ly,
                   MAX(CASE WHEN d.MaMon='M05' THEN d.DiemHK2 END) AS HKII_Hoa,
                   MAX(CASE WHEN d.MaMon='M06' THEN d.DiemHK2 END) AS HKII_Sinh,
                   MAX(CASE WHEN d.MaMon='M07' THEN d.DiemHK2 END) AS HKII_Su,
                   MAX(CASE WHEN d.MaMon='M08' THEN d.DiemHK2 END) AS HKII_Dia,
                   MAX(CASE WHEN d.MaMon='M09' THEN d.DiemHK2 END) AS HKII_GDCD,
                   MAX(d.DiemTB) AS DiemTB,
                   MAX(d.XepLoai) AS XepLoai
            FROM HocSinh hs
            LEFT JOIN Diem d ON hs.MaHS = d.MaHS
            WHERE hs.MaLop = ?
            GROUP BY hs.MaHS, hs.HoTen
            ORDER BY hs.MaHS
        """, (lop_to_malop[lop_duoc_chon],))
    else: # Nếu chưa chọn lớp thì hiển thị tất cả
        cur.execute("""
            SELECT hs.MaHS, hs.HoTen,
                   MAX(CASE WHEN d.MaMon='M01' THEN d.DiemHK1 END) AS HKI_Toan,
                   MAX(CASE WHEN d.MaMon='M02' THEN d.DiemHK1 END) AS HKI_Van,
                   MAX(CASE WHEN d.MaMon='M03' THEN d.DiemHK1 END) AS HKI_Anh,
                   MAX(CASE WHEN d.MaMon='M04' THEN d.DiemHK1 END) AS HKI_Ly,
                   MAX(CASE WHEN d.MaMon='M05' THEN d.DiemHK1 END) AS HKI_Hoa,
                   MAX(CASE WHEN d.MaMon='M06' THEN d.DiemHK1 END) AS HKI_Sinh,
                   MAX(CASE WHEN d.MaMon='M07' THEN d.DiemHK1 END) AS HKI_Su,
                   MAX(CASE WHEN d.MaMon='M08' THEN d.DiemHK1 END) AS HKI_Dia,
                   MAX(CASE WHEN d.MaMon='M09' THEN d.DiemHK1 END) AS HKI_GDCD,
                   MAX(CASE WHEN d.MaMon='M01' THEN d.DiemHK2 END) AS HKII_Toan,
                   MAX(CASE WHEN d.MaMon='M02' THEN d.DiemHK2 END) AS HKII_Van,
                   MAX(CASE WHEN d.MaMon='M03' THEN d.DiemHK2 END) AS HKII_Anh,
                   MAX(CASE WHEN d.MaMon='M04' THEN d.DiemHK2 END) AS HKII_Ly,
                   MAX(CASE WHEN d.MaMon='M05' THEN d.DiemHK2 END) AS HKII_Hoa,
                   MAX(CASE WHEN d.MaMon='M06' THEN d.DiemHK2 END) AS HKII_Sinh,
                   MAX(CASE WHEN d.MaMon='M07' THEN d.DiemHK2 END) AS HKII_Su,
                   MAX(CASE WHEN d.MaMon='M08' THEN d.DiemHK2 END) AS HKII_Dia,
                   MAX(CASE WHEN d.MaMon='M09' THEN d.DiemHK2 END) AS HKII_GDCD,
                   MAX(d.DiemTB) AS DiemTB,
                   MAX(d.XepLoai) AS XepLoai
            FROM HocSinh hs
            LEFT JOIN Diem d ON hs.MaHS = d.MaHS
            GROUP BY hs.MaHS, hs.HoTen
            ORDER BY hs.MaHS
        """)
    rows = cur.fetchall()
    for row in rows:
        capnhat_so_luong()
        tree.insert("",
                   tk.END, values=(row[0], row[1],
                                  row[2], row[3], row[4],row[5],row[6],row[7],row[8],row[9],row[10],
                                 row[11], row[12], row[13], row[14],row[15],row[16],row[17],row[18],row[19],
                                 row[20], row[21]))


#============================================= NÚT BẤM =============================================
# Khoảng cách trước các nút
frame_nut = tk.Frame(frame_tt, bg="#F5F5F5")
frame_nut.grid(row=1, column=0, columnspan=5, sticky="w", pady=5)

# Nút tính điểm TB và xếp loại
btn_diem = tk.Button(
    frame_nut,
    text="Tính Điểm TB và Xếp Loại",
    font=("Times New Roman", 12),
    
    relief="ridge",
    cursor="hand2",
    command=tinhdiem)
btn_diem.grid(row=1, column=3, padx=5, pady=5, sticky="w")

# Nút nhập điểm mới
btn_nhapdiem = tk.Button(
    frame_nut, text="Nhập điểm mới",
    font=("Times New Roman", 12),
 
    relief="ridge",
    cursor="hand2",
    command=sua_diem)
btn_nhapdiem.grid(row=1, column=4, padx=15, pady=5)

# Nút xuất Excel
btn_in_excel = tk.Button(
    frame_nut,
    text="In bảng điểm",
    font=("Times New Roman", 12),
    relief="ridge",
    cursor="hand2",
    command=xuat_excel
)
btn_in_excel.grid(row=1, column=5, padx=15, pady=5)

label_Soluong = tk.Label(
    frame_nut,
    text="(Số lượng học sinh)",
    font=("Times New Roman", 11, "bold"),
    bg="#F5F5F5"

)
label_Soluong.grid(row=0, column=0, columnspan=4, sticky="w", padx=15, pady=5)


#============================================= FRAME TREE VIEW =============================================
frame_tree = tk.Frame(root)
frame_tree.pack(fill="both", expand=True, padx=10, pady=5)
# Thanh cuộn dọc
thanhcuon_doc = tk.Scrollbar(frame_tree, orient="vertical")
thanhcuon_doc.pack(side="right", fill="y")
# Thanh cuộn ngang
thanhcuon_ngang = tk.Scrollbar(frame_tree, orient="horizontal")
thanhcuon_ngang.pack(side="bottom", fill="x")
# Cấu hình Treeview với cả hai thanh cuộn
columns = (
    "MaHS", "HoTen",
    "I_Toan", "I_Van", "I_Anh", "I_Ly", "I_Hoa", "I_Sinh", "I_Su", "I_Dia", "I_GDCD",
    "II_Toan", "II_Van", "II_Anh", "II_Ly", "II_Hoa", "II_Sinh", "II_Su", "II_Dia", "II_GDCD",
    "DiemTB", "XepLoai"
)
tree = ttk.Treeview(frame_tree, columns=columns, show="headings", height=15)
tree.pack(fill="both", expand=True)
tree.configure(yscrollcommand=thanhcuon_doc.set, xscrollcommand=thanhcuon_ngang.set)

thanhcuon_doc.config(command=tree.yview)
thanhcuon_ngang.config(command=tree.xview)

# Đặt chiều rộng các cột Treeview
for col in columns:
    tree.heading(col, text=col)

tree.column("MaHS", width=60, anchor="center")
tree.column("HoTen", width=200, anchor="w")
for col in columns[2:]:
    tree.column(col, width=60, anchor="center")
tree.column("DiemTB", width=80, anchor="center")
tree.column("XepLoai", width=80, anchor="center")

tree.pack(fill="both", expand=True)


#============================================= MAU NEN =============================================

root.configure(bg="#F5F5F5")
frame_tt.configure(bg="#F5F5F5")


lbl_tieu_de.configure(bg="#F5F5F5", fg="#0000FF")



style = ttk.Style()
style.configure("Treeview",
                font=("Times New Roman", 11),
                rowheight=25,
                background="#FFFFFF",
                fieldbackground="#FFFFFF",
                foreground="#1E293B")

style.configure("Treeview.Heading",
               font=("Times New Roman", 10, "bold"),
               background="#64748B")

style.map("Treeview", background=[("selected", "#003366")])

load_dulieu()
root.mainloop()
