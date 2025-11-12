# DoAn_Python_Nhom13_Lop.DH24TH3_NhomLT3.ToTH1
Ứng dụng quản lý học sinh gồm có 3 form đó là: form đăng nhập, form quản lý học sinh và form quản lý điểm.
Ở form đăng nhập:
- Khi người dùng đăng nhập đúng tên đăng nhập và mật khẩu ở phân quyền admin. Khi nhấn nút đăng nhập. Form đăng nhập sẽ được đóng và tiến hành mở form quản lý học sinh 
- Khi người dùng đăng nhập đúng tên đăng nhập và mật khẩu ở phân quyền giáo viên. Khi nhấn nút đăng nhập. Form đăng nhập sẽ được đóng và tiến hành mở form quản lý điểm học sinh.
- Phân quyền ở form đăng nhập cho nhầm để giới hạn lại quyền sửa chữa của từng vị trí với từng quyền riêng như sau:
  + Admin: chỉ được quyền thêm xóa sửa học sinh, không được vào form quản lý điểm
  + Giáo viên: chỉ được quyền thêm, sửa điểm học sinh và tiến hành tính điểm cũng như xếp loại học sinh qua điểm từng môn ở 2 học kỳ. Giáo viên không được vào form quản lý học sinh.

