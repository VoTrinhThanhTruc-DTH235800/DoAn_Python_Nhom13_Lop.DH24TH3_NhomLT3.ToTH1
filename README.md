# DoAn_Python_Nhom13_Lop.DH24TH3_NhomLT3.ToTH1
Ứng dụng quản lý học sinh gồm có 3 form đó là: form đăng nhập, form quản lý học sinh và form quản lý điểm.
Ở form đăng nhập:
- Khi người dùng đăng nhập đúng tên đăng nhập và mật khẩu ở phân quyền admin. Khi nhấn nút đăng nhập. Form đăng nhập sẽ được đóng và tiến hành mở form quản lý học sinh 
- Khi người dùng đăng nhập đúng tên đăng nhập và mật khẩu ở phân quyền giáo viên. Khi nhấn nút đăng nhập. Form đăng nhập sẽ được đóng và tiến hành mở form quản lý điểm học sinh.
- Phân quyền ở form đăng nhập cho nhầm để giới hạn lại quyền sửa chữa của từng vị trí với từng quyền riêng như sau:
  + Admin: chỉ được quyền thêm xóa sửa học sinh, không được vào form quản lý điểm
  + Giáo viên: chỉ được quyền thêm, sửa điểm học sinh và tiến hành tính điểm cũng như xếp loại học sinh qua điểm từng môn ở 2 học kỳ. Giáo viên không được vào form quản lý học sinh.
Ở form quản lý học sinh:
- Ở form này người dùng admin có khả năng thêm, xóa, sửa, hủy, lưu, và in danh sách sinh viên theo lớp:
- Vừa vào form ta sẽ thấy dữ liệu được load từ cơ sở dữ liệu lên bảng danh sách sinh viên.
- Khi thêm 1 học sinh vào thì admin phải đảm bảo nhập đầy đủ các thông tin cần có. Nếu nhập thiếu 1 trong số các thông tin sẽ không lưu được và hệ thống sẽ hiện messagebox báo cần nhập đẩy đủ thông tin học sinh trước khi thêm.
- Khi xóa học sinh thì admin chỉ cần nháy chọn học sinh cần xóa và nhấn nút xóa. Hệ thống sẽ có 1 messagebox cảnh báo là có chắc chắn là xóa học sinh đó không.
- Khi sửa học sinh thì admin sẽ tiến hành chọn học sinh sau đó nhấn nút sửa. Khi sửa xong và nhấn nút lưu, hệ thống sẽ hỏi có chắc lưu thông tin mới sửa hay không.
- Khi admin chọn 1 học sinh mà sau đó muốn hủy tháo tác đó. Admin cần nhấn nút Hủy sẽ hủy được thao tác vừa thực hiện.
- Khi admin muốn in danh sách học sinh theo lớp. Admin chỉ cần chọn lớp và nhấn in danh sách. Danh sách học sinh sẽ được lưu dưới định dạng file excel theo lớp.
Ở form quản lý điểm:
- Ở form này giáo viên có khả năng sửa hoặc thêm điểm và tính điểm, xếp loại học sinh.
- Khi bên form quản lý học sinh có thêm 1 học sinh mới không có trong cơ sở dữ liệu thì điểm ở từng môn ở mỗi học kì là none nên giáo viên sẽ tiến hành nhập điểm mới.
- Nếu ở file này giáo viên muốn in bảng điểm theo lớp. Giáo viên chỉ cần chọn lớp và nhấn nút in danh sách. Danh sách sẽ được in theo định dạng excel.
