from math import e
import tkinter as tk
from tkinter import Menu, ttk, messagebox
from tkcalendar import DateEntry
import pyodbc
import subprocess
import sys
from openpyxl import Workbook
from tkinter import filedialog
# ====== Kết nối cơ sở dữ liệu ======
conn = pyodbc.connect(
            'DRIVER={SQL Server};'
            'SERVER=ADMIN-PC\\SQLEXPRESS;'
            'DATABASE=QLHS;'
            'Trusted_Connection=yes;'
        )

cur = conn.cursor()


# ====== Hàm canh giữa cửa sổ ====== 
def center_window(win, w=700, h=600):
    ws = win.winfo_screenwidth()
    hs = win.winfo_screenheight()
    x = (ws // 2) - (w // 2)
    y = (hs // 2) - (h // 2)
    win.geometry(f'{w}x{h}+{x}+{y}')
# ====== Cửa sổ chính ======
root = tk.Tk()
root.title("Quản lý học sinh")
center_window(root, 700, 600)
root.resizable(False, False)



# ================================= MENU ===================================

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
menu_quanly.add_command(label="Quản lý học sinh",font=("Times New Roman", 11), 
                        command=lambda: messagebox.showinfo("Thông báo", "Bạn đang ở trang Quản lý học sinh"))
menu_quanly.add_command(label="Đăng xuất",font=("Times New Roman", 11), command=lambda: mo_form_dangnhap())
menubar.add_cascade(label="Menu",font=("Times New Roman", 11),menu=menu_quanly)

root.config(menu=menubar)


# ========================================== TIÊU ĐỀ ==================================================
lbl_tieu_de = tk.Label(root, text="QUẢN LÝ HỌC SINH", font=("Times New Roman", 18, "bold")  )
lbl_tieu_de.pack(pady=10)
# ====== Frame nhập thông tin ====== 
frame_tt = tk.Frame(root) 
frame_tt.pack(pady=5, padx=10, fill="x")

tk.Label(frame_tt, text="Mã số", font=("Times New Roman", 13),bg="#F5F5F5").grid(row=0, column=0, padx=5, pady=5, sticky="w")
entry_maso = tk.Entry(frame_tt, width=20)
entry_maso.grid(row=0, column=1, padx=5, pady=5, sticky="w")

# Đặt khoảng cách lớn giữa entry_maso và label "Lớp"
frame_tt.grid_columnconfigure(2, minsize=100)

tk.Label(frame_tt, text="Lớp", font=("Times New Roman", 13),bg="#F5F5F5").grid(row=0, column=3, padx=5, pady=5, sticky="w")
cbb_lop = ttk.Combobox(frame_tt, values=["10A1", "10A2", "10A3", "10A4", "10A5", "10A6"], width=15)
cbb_lop.grid(row=0, column=4, padx=5, pady=5, sticky="w")

tk.Label(frame_tt, text="Họ tên", font=("Times New Roman", 13),bg="#F5F5F5").grid(row=1,	column=0, padx=5, pady=5, sticky="w")
entry_hoten = tk.Entry(frame_tt, width=35) 
entry_hoten.grid(row=1, column=1, padx=5, pady=5, sticky="w")

tk.Label(frame_tt, text="Giới tính", font=("Times New Roman", 13),bg="#F5F5F5").grid(row=2, column=0, padx=5, pady=5, sticky="w") 
gender_var = tk.StringVar(value="Nam")
tk.Radiobutton(frame_tt,	text="Nam",font=("Times New Roman", 13),bg="#F5F5F5",
               variable=gender_var, value="Nam").grid(row=2, column=1, padx=5, sticky="w")
tk.Radiobutton(frame_tt, text="Nữ",font=("Times New Roman", 13),bg="#F5F5F5",
               variable=gender_var, value="Nữ").grid(row=2, column=1, padx=70, sticky="w")

tk.Label(frame_tt, text="Mã lớp", font=("Times New Roman", 13),bg="#F5F5F5").grid(row=1, column=3, padx=5, pady=5, sticky="w")
entry_malop = tk.Entry(frame_tt, width=18)
entry_malop.grid(row=1, column=4, padx=5, pady=5, sticky="w")


tk.Label(frame_tt, text="Ngày sinh", font=("Times New Roman", 13),bg="#F5F5F5").grid(row=2, column=3, padx=5, pady=5, sticky="w")
date_entry = DateEntry(frame_tt, width=13, background="darkblue", 
                       foreground="white", date_pattern="yyyy-mm-dd", font=("Times New Roman", 11))
date_entry.grid(row=2, column=4, padx=5, pady=5, sticky="w")

#========================================= CÁC HÀM CHỨC NĂNG ================================================
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
        load_du_lieu()  # Tự động lọc khi nhập mã lớp

def on_lop_duoc_chon(event=None):
    global lop_duoc_chon
    lop_duoc_chon = cbb_lop.get()
    # Cập nhật mã lớp
    ma_lop = lop_to_malop.get(lop_duoc_chon, "")
    entry_malop.delete(0, tk.END)
    entry_malop.insert(0, ma_lop)
    load_du_lieu()  # Load danh sách theo lớp

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

def load_du_lieu():
    for i in tree.get_children():
        tree.delete(i)

    if lop_duoc_chon:
        cur.execute("SELECT MaHS, HoTen, NgaySinh, GioiTinh, MaLop FROM HocSinh WHERE MaLop=?", (lop_to_malop[lop_duoc_chon],))
    else:
        cur.execute("SELECT MaHS, HoTen, NgaySinh, GioiTinh, MaLop FROM HocSinh")
    rows = cur.fetchall()
    for row in rows:
        capnhat_so_luong()
        tree.insert("", tk.END, values=(row[0], row[1], row[2], row[3], row[4]))


def Huy_tt():
    entry_maso.delete(0, tk.END)
    entry_hoten.delete(0, tk.END)
    entry_malop.delete(0, tk.END)
    gender_var.set("Nam")
    date_entry.set_date("2000-01-01")
    cbb_lop.set("")


def them_hs():
    maso = entry_maso.get().strip()
    hoten = entry_hoten.get().strip()
    ngaysinh = date_entry.get_date()
    phai = gender_var.get()
    malop = entry_malop.get().strip()

    if not maso or not hoten or not malop:
        messagebox.showwarning("Cảnh báo", "Vui lòng điền đầy đủ thông tin.")
        return

    try:
        # Kiểm tra mã học sinh đã tồn tại chưa
        cur.execute("SELECT COUNT(*) FROM HocSinh WHERE MaHS = ?", (maso,))
        count_mahs = cur.fetchone()[0]

        if count_mahs > 0:
            messagebox.showerror("Lỗi", "Mã học sinh đã tồn tại. Vui lòng nhập mã khác.")
            return

        # Kiểm tra lớp đã tồn tại chưa
        cur.execute("SELECT COUNT(*) FROM Lop WHERE MaLop = ?", (malop,))
        count_lop = cur.fetchone()[0]

        # Nếu chưa có lớp thì thêm vào bảng Lop
        if count_lop == 0:
            # Ở đây bạn có thể yêu cầu người dùng nhập thêm TenLop (nếu có Textbox riêng)
            # hoặc tạm thời dùng MaLop làm TenLop
            cur.execute("INSERT INTO Lop (MaLop, TenLop) VALUES (?, ?)", (malop, malop))
            conn.commit()

        # Thêm học sinh vào bảng HocSinh
        cur.execute("""
            INSERT INTO HocSinh (MaHS, HoTen, NgaySinh, GioiTinh, MaLop)
            VALUES (?, ?, ?, ?, ?)
        """, (maso, hoten, ngaysinh.strftime("%Y-%m-%d"), phai, malop))
        conn.commit()

        messagebox.showinfo("Thành công", "Thêm học sinh thành công.")
        load_du_lieu()
        Huy_tt()

    except Exception as e:
        messagebox.showerror("Lỗi", f"Có lỗi xảy ra: {str(e)}")

def xoa_hs():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Cảnh báo", "Vui lòng chọn học sinh để xóa.")
        return
    hoi = messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn xóa học sinh này không?")
    if hoi:
        maso = tree.item(selected[0])['values'][0]
        cur.execute("DELETE FROM Diem WHERE MaHS=?", (maso,))
        cur.execute("DELETE FROM HocSinh WHERE MaHS = ?", (maso,))
    
        conn.commit()
        messagebox.showinfo("Thành công","Xóa học sinh thành công.")
        load_du_lieu()
        Huy_tt()
    else:
        return

# global selected_maso dùng chung
selected_maso = None

def sua_hs():
    global selected_maso
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Cảnh báo", "Vui lòng chọn học sinh để sửa.")
        return
    # lưu mã cũ ngay khi bấm sửa
    selected_maso = tree.item(selected[0])["values"][0]
    values = tree.item(selected[0])["values"]
    # điền form
    entry_maso.delete(0, tk.END)
    entry_maso.insert(0, values[0])

    entry_hoten.delete(0, tk.END)
    entry_hoten.insert(0, values[1])

    # nếu values[2] là chuỗi 'YYYY-MM-DD' hoặc dạng date, set_date sẽ chấp nhận
    try:
        date_entry.set_date(values[2])
    except Exception:
        # fallback: nếu là datetime.date object
        date_entry.set_date(values[2])

    gender_var.set(values[3])
    cbb_lop.set(values[4])

    # Nếu bạn có cả entry_malop (text) và combobox, giữ đồng bộ
    entry_malop.delete(0, tk.END)
    entry_malop.insert(0, values[4])


def luu_sua():
    global selected_maso
    maso_moi = entry_maso.get().strip()
    hoten = entry_hoten.get().strip()
    ngaysinh = date_entry.get_date()
    phai = gender_var.get()
    malop = entry_malop.get().strip()

    if not maso_moi or not hoten or not malop:
        messagebox.showwarning("Cảnh báo", "Vui lòng điền đầy đủ thông tin.")
        return

    if not selected_maso:
        messagebox.showwarning("Cảnh báo", "Không xác định học sinh đang sửa.")
        return
    hoi = messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn lưu thay đổi không?")
    if not hoi:
        return
    try:
        # Nếu người dùng đổi mã -> kiểm tra trùng
        if maso_moi != selected_maso:
            cur.execute("SELECT COUNT(*) FROM HocSinh WHERE MaHS = ?", (maso_moi,))
            if cur.fetchone()[0] > 0:
                messagebox.showerror("Lỗi", "Mã học sinh mới đã tồn tại. Vui lòng nhập mã khác.")
                return

        # Thực hiện cập nhật (bao gồm đổi MaHS nếu cần)
        cur.execute("""
            UPDATE HocSinh
            SET MaHS=?, HoTen=?, NgaySinh=?, GioiTinh=?, MaLop=?
            WHERE MaHS=?
        """, (maso_moi, hoten, ngaysinh.strftime("%Y-%m-%d"), phai, malop, selected_maso))

        conn.commit()
        messagebox.showinfo("Thành công", "Cập nhật học sinh thành công.")

        # load_data() có thể reload tree; nếu reload thì restore selection theo MaHS mới
        load_du_lieu()
        try:
            # sau khi load lại, tìm row có MaHS == maso_moi và set selection
            for iid in tree.get_children():
                vals = tree.item(iid, "values")
                if vals and vals[0] == maso_moi:
                    tree.selection_set(iid)
                    tree.focus(iid)
                    break
        except Exception:
            pass

        Huy_tt()
        selected_maso = None

    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể cập nhật học sinh:\n{e}")


def Huy_hs():
    hoi = messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn hủy thao tác không?")
    if not hoi:
        return
    Huy_tt()
    messagebox.showinfo("Hủy", "Hủy thao tác thành công.")
    load_du_lieu()

def Thoat():
    if messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn thoát không?"):
        try:
            conn.close()
        except:
            pass
        root.destroy()

def xuat_excel():
    if not lop_duoc_chon:
        messagebox.showwarning("Thông báo", "Vui lòng chọn lớp để xuất danh sách học sinh!")
        return
    hoi = messagebox.askyesno("Xác nhận", f"Bạn có chắc chắn muốn xuất danh sách học sinh lớp {lop_duoc_chon} ra file Excel không?")
    if not hoi:
        return
    # Lấy dữ liệu hiện tại từ Treeview
    rows = [tree.item(item, "values") for item in tree.get_children()]
    if not rows:
        messagebox.showinfo("Thông báo", "Không có dữ liệu để xuất.")
        return

    # Chọn nơi lưu file
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Lưu danh sách học sinh",
        initialfile=f"DanhSachHS_{lop_duoc_chon}.xlsx"
    )
    if not file_path:
        return

    try:
        wb = Workbook()
        ws = wb.active
        ws.title = f"Lớp {lop_duoc_chon}"

        # Ghi tiêu đề theo Treeview
        headers = ["Mã HS", "Họ tên", "Ngày sinh", "Giới tính", "Mã lớp"]
        ws.append(headers)

        # Ghi dữ liệu
        for row in rows:
            ws.append(row)

        # Căn chỉnh độ rộng cột tự động
        for col in ws.columns:
            max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
            ws.column_dimensions[col[0].column_letter].width = max_length + 3

        wb.save(file_path)
        messagebox.showinfo("Thành công", f"Đã xuất danh sách học sinh lớp {lop_duoc_chon} ra file Excel thành công!")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể xuất file Excel:\n{e}")




# ==================================================== NÚT BẤM =================================================
frame_nut = tk.Frame(root) 
frame_nut.pack(pady=10, anchor="w", padx=10) 
 
btn_them = tk.Button(
    frame_nut, text="Thêm",
    font=("Times New Roman", 11),
    width=11,
    relief="ridge",
    cursor="hand2",
    command= them_hs
    )
btn_them.grid(row=0, column=0, padx=5) 

btn_xoa = tk.Button(
    frame_nut, 
    text="Xóa", 
    font=("Times New Roman", 11), 
    width=11, 
    relief="ridge",
    cursor="hand2",
    command= xoa_hs
    )
btn_xoa.grid(row=0, column=1, padx=5)


btn_sua = tk.Button(
    frame_nut, 
    text="Sửa", 
    font=("Times New Roman", 11), 
    width=11, 
    relief="ridge",
    cursor="hand2",
    command= sua_hs
    )
btn_sua.grid(row=0, column=2, padx=5)


btn_luu = tk.Button(
    frame_nut, 
    text="Lưu", 
    font=("Times New Roman", 11),
    relief="ridge",
    cursor="hand2",
    width=11, 
    command= luu_sua
    )
btn_luu.grid(row=0, column=3, padx=5)


btn_huy = tk.Button(
    frame_nut, text="Hủy", 
    font=("Times New Roman", 11), 
    width=11, 
    relief="ridge",
    cursor="hand2",
    command= Huy_hs
    )
btn_huy.grid(row=0, column=4, padx=5)


btn_thoat = tk.Button(
    frame_nut, 
    text="Thoát",
    font=("Times New Roman", 11), 
    width=11, 
    relief="ridge",
    cursor="hand2",
    command= Thoat
    )
btn_thoat.grid(row=0, column=5, padx=5)

btn_in_excel = tk.Button(
    frame_nut,
    text="In Danh Sách",
    font=("Times New Roman", 12),
    relief="ridge",
    cursor="hand2",
    command=xuat_excel
)
btn_in_excel.grid(row=1, column=5, padx=5, pady=5)

label_Soluong = tk.Label(
    frame_nut,
    text="(Số lượng học sinh)",
    font=("Times New Roman", 11, "bold"),
    bg="#F5F5F5"
)
label_Soluong.grid(row=1, column=0, columnspan=4, sticky="w", padx=5, pady=5)

#================================= FRAME TREEVIEW ==================================
lbl_ds = tk.Label(root, text="Danh sách học sinh", font=("Times New Roman", 13, "bold"))
lbl_ds.pack(pady=5, anchor="w", padx=10)
#====== Treeview hiển thị danh sách học sinh ======


# ====== Frame chứa Treeview và thanh cuộn ======
frame_tree = tk.Frame(root)
frame_tree.pack(padx=10, pady=5, fill="both", expand=True)

# Thanh cuộn
thanhcuon_doc = tk.Scrollbar(frame_tree, orient="vertical")
thanhcuon_doc.pack(side="right", fill="y")

# Cấu hình Treeview
columns = ("Mã số", "Họ tên", "Ngày sinh", "Giới tính", "Mã lớp")
tree = ttk.Treeview(frame_tree, columns=columns, show="headings", height=10, yscrollcommand=thanhcuon_doc.set)
thanhcuon_doc.config(command=tree.yview)

# Đặt tiêu đề cột 
for col in columns:
    tree.heading(col, text=col)

# Đặt kích thước cột 
tree.column("Mã số", width=80, anchor="center")
tree.column("Họ tên", width=200)
tree.column("Ngày sinh", width=120, anchor="center")
tree.column("Giới tính", width=80, anchor="center")
tree.column("Mã lớp", width=100, anchor="center")


tree.pack(fill="both", expand=True)
style = ttk.Style()
style.configure("Treeview",
                font=("Times New Roman", 11),
                rowheight=25)
style.configure("Treeview.Heading",
               font=("Times New Roman", 12, "bold"))



#=========================================== MÀU NỀN ========================================
# Màu nền form - Xám kem tinh tế
root.configure(bg="#F5F5F5")
frame_tt.configure(bg="#F5F5F5")
frame_nut.configure(bg="#F5F5F5")

# Màu tiêu đề - Slate đậm
lbl_tieu_de.configure(bg="#F5F5F5", fg="#0000FF")
lbl_ds.configure(bg="#F5F5F5",fg="#0000FF")

style = ttk.Style()
style.configure("Treeview",
                font=("Times New Roman", 11),
                rowheight=25,
                background="#FFFFFF",
                fieldbackground="#FFFFFF",
                foreground="#1E293B")

style.configure("Treeview.Heading",
               font=("Times New Roman", 12, "bold"),
               background="#64748B")

style.map("Treeview", background=[("selected", "#003366")])


# ====== Tải dữ liệu ban đầu ======
load_du_lieu()
root.mainloop()


