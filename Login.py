from math import e
import tkinter as tk
from tkinter import Menu, ttk, messagebox
from tkcalendar import DateEntry
import pyodbc
import subprocess
import sys

# ====== Kết nối cơ sở dữ liệu ======
conn = pyodbc.connect(
            'DRIVER={SQL Server};'
            'SERVER=ADMIN-PC\\SQLEXPRESS;'
            'DATABASE=QLHS;'
            'Trusted_Connection=yes;'
        )

cur = conn.cursor()


# ====== Hàm canh giữa cửa sổ ======
def center_window(win, w=700, h=500):
    ws = win.winfo_screenwidth()
    hs = win.winfo_screenheight()
    x = (ws // 2) - (w // 2)
    y = (hs // 2) - (h // 2)
    win.geometry(f'{w}x{h}+{x}+{y}')

# ====== Cửa sổ chính ======
root = tk.Tk()
root.title("Đăng nhập hệ thống")
center_window(root, 700, 500)
root.resizable(False, False)
root.configure(bg="#F5F5F5")



# ====== Tiêu đề ======
lbl_title = tk.Label(
    root, 
    text="ĐĂNG NHẬP \nHỆ THỐNG QUẢN LÝ HỌC SINH", 
    font=("Times New Roman", 24, "bold"), 
    bg="#F5F5F5", 
    fg="#0000FF"
)
lbl_title.pack(pady=15)

# ====== Frame thông tin ======
frame_tt = tk.Frame(root, bg="#F5F5F5")
frame_tt.pack(pady=10, padx=20)

tk.Label(frame_tt, text="Tên đăng nhập:", font=("Times New Roman", 16, "bold"), bg="#F5F5F5").grid(row=0, column=0, padx=5, pady=20, sticky="w")
entry_username = tk.Entry(frame_tt, width=25, font=("Times New Roman", 16))
entry_username.grid(row=0, column=1, padx=5, pady=20)

tk.Label(frame_tt, text="Mật khẩu:", font=("Times New Roman", 16, "bold"), bg="#F5F5F5").grid(row=2, column=0, padx=5, pady=20, sticky="w")
entry_password = tk.Entry(frame_tt, width=25, font=("Times New Roman", 16), show="*")
entry_password.grid(row=2, column=1, padx=5, pady=20)

tk.Label(frame_tt, text="Phân quyền:", font=("Times New Roman", 16), bg="#F5F5F5").grid(row=3, column=0, padx=5, pady=20, sticky="w")
cbb_role = ttk.Combobox(frame_tt, values=["Admin", "Giáo Viên"], font=("Times New Roman", 16), width=23)
cbb_role.grid(row=3, column=1, padx=5, pady=20)
cbb_role.current(0)


# ================================== CÁC HÀM CHỨC NĂNG ==================================
def mo_form_diem():
    root.withdraw()
    try:
        subprocess.Popen([sys.executable, r"E:\DoAn_QLHS\QLHS\QuanLyDiem.py"])
        root.destroy()  # đóng form hiện tại hoàn toàn
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể mở form tính điểm:\n{e}")
menubar = tk.Menu(root)

def mo_form_quanlyhs():
    root.withdraw()
    try:
        subprocess.Popen([sys.executable, r"E:\DoAn_QLHS\QLHS\QuanLyHocSinh.py"])
        root.destroy()  # đóng form hiện tại hoàn toàn
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể mở form quản lý học sinh:\n{e}")


def DangNhap():

    username = entry_username.get().strip()
    password = entry_password.get().strip()
    phanquyen = cbb_role.get().strip()
    if not username or not password:
        messagebox.showwarning("Cảnh báo", "Vui lòng nhập đầy đủ thông tin đăng nhập.")
        return
    query = "SELECT * FROM TaiKhoan WHERE UserName=? AND PassWord=? AND PhanQuyen=?"
    cur.execute(query, (username, password, phanquyen))
    result = cur.fetchone()
    if result:
        messagebox.showinfo("Thành công", f"Đăng nhập thành công với vai trò: {phanquyen}")
        if phanquyen == "Admin":
            mo_form_quanlyhs()
        else:
            mo_form_diem()
    else:
        messagebox.showerror("Lỗi", "Tên đăng nhập, mật khẩu  hoặc chọn phân quyền không đúng.")
        entry_username.delete(0, tk.END)
        entry_password.delete(0, tk.END)
        entry_username.focus()
# ====================================== NÚT BẤM ======================================
frame_nut = tk.Frame(root, bg="#F5F5F5")
frame_nut.pack(pady=15)

btn_dangnhap = tk.Button(
    frame_nut, 
    text="Đăng nhập", 
    font=("Times New Roman", 14, "bold"), 
    bg="#0066cc", 
    fg="white", 
    width=14, 
    relief="ridge", 
    cursor="hand2", 
    command=lambda: DangNhap()
)
btn_dangnhap.grid(row=0, column=0, padx=10)

btn_thoat = tk.Button(
    frame_nut, 
    text="Thoát", 
    font=("Times New Roman", 14, "bold"), 
    bg="#cc0000", 
    fg="white", 
    width=14, 
    relief="ridge", 
    cursor="hand2",
    command=root.destroy
)
btn_thoat.grid(row=0, column=1, padx=10)
# ====== Footer ======
lbl_footer = tk.Label(
    root,
    text="© 2025 - Nhóm Quản Lý Học Sinh \n Liên hệ: truc_dth235800@student.agu.edu.vn hoặc tuyen_dth235810@student.agu.edu.vn",
    font=("Times New Roman", 11, "italic"),
    bg="#F5F5F5",
    fg="gray"
)
lbl_footer.pack(side="bottom", pady=15)

root.mainloop()

