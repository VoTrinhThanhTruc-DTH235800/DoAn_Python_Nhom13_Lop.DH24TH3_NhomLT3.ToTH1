--DROP DATABASE QLHS;
CREATE DATABASE QLHS
ON (
NAME = QLHS_mdf, 
    FILENAME = 'E:\CSDL_HS\QLHS.mdf',  
    SIZE = 15, 
    MAXSIZE = 50, 
    FILEGROWTH = 5
)
LOG ON
(
    NAME = QLSV_log, 
    FILENAME = 'E:\CSDL_HS\QLHS.ldf',  
    SIZE = 15, 
    MAXSIZE = 50, 
    FILEGROWTH = 5
);
USE QLHS;
CREATE TABLE Lop (
    MaLop CHAR(3) PRIMARY KEY,
    TenLop NVARCHAR(20) NOT NULL
);
CREATE TABLE HocSinh (
    MaHS CHAR(4) PRIMARY KEY,
    HoTen NVARCHAR(50) NOT NULL,
    NgaySinh DATE,
    GioiTinh NVARCHAR(5),
    MaLop CHAR(3),
    FOREIGN KEY (MaLop) REFERENCES Lop(MaLop)
);
CREATE TABLE MonHoc (
    MaMon CHAR(3) PRIMARY KEY,
    TenMon NVARCHAR(30) NOT NULL
);
CREATE TABLE Diem (
    MaHS CHAR(4),
    MaMon CHAR(3),
    DiemHK1 FLOAT CHECK (DiemHK1 BETWEEN 0 AND 10),
    DiemHK2 FLOAT CHECK (DiemHK2 BETWEEN 0 AND 10),
    PRIMARY KEY (MaHS, MaMon),
    FOREIGN KEY (MaHS) REFERENCES HocSinh(MaHS),
    FOREIGN KEY (MaMon) REFERENCES MonHoc(MaMon)
);
CREATE TABLE TaiKhoan (
    UserName VARCHAR(10) PRIMARY KEY,       -- Mã người dùng: GV01, CB01, v.v.
    PassWord VARCHAR(20) NOT NULL,          -- Mật khẩu đăng nhập
    PhanQuyen NVARCHAR(20) NOT NULL         -- Quyền: Giáo Viên hoặc Admin
);

-- Bảng Lop
INSERT INTO Lop (MaLop, TenLop) VALUES
('L01', N'10A1'),
('L02', N'10A2');

-- Bảng HocSinh
INSERT INTO HocSinh (MaHS, HoTen, NgaySinh, GioiTinh, MaLop) VALUES
('HS01', N'Võ Bình An', '2008-07-17', N'Nam', 'L01'),
('HS02', N'Trần Linh Chi', '2008-06-22', N'Nam', 'L01'),
('HS03', N'Võ Văn Tần', '2006-09-07', N'Nam', 'L01'),
('HS04', N'Trần Minh Hiếu', '2006-05-16', N'Nam', 'L02'),
('HS05', N'Phạm Thùy Chi', '2007-04-08', N'Nữ', 'L02'),
('HS06', N'Lê Minh An', '2007-10-09', N'Nữ', 'L02'); 

drop table TaiKhoan

-- Bảng MonHoc
-- Bảng MonHoc
INSERT INTO MonHoc (MaMon, TenMon) VALUES
('M01', N'Toán'),
('M02', N'Ngữ Văn'),
('M03', N'Tiếng Anh'),
('M04', N'Vật Lý'),
('M05', N'Hóa Học'),
('M06', N'Sinh Học'),
('M07', N'Lịch Sử'),
('M08', N'Địa Lý'),
('M09', N'GDCD');

-- Bảng Diem
INSERT INTO Diem (MaHS, MaMon, DiemHK1, DiemHK2) VALUES
-- HS01
('HS01', 'M01', 5.2, 6.6),
('HS01', 'M02', 5.9, 6.1),
('HS01', 'M03', 7.3, 5.6),
('HS01', 'M04', 6.5, 7.0),
('HS01', 'M05', 5.8, 6.3),
('HS01', 'M06', 7.1, 6.8),
('HS01', 'M07', 6.0, 5.5),
('HS01', 'M08', 6.7, 7.2),
('HS01', 'M09', 8.0, 7.5),
-- HS02
('HS02', 'M01', 6.7, 7.4),
('HS02', 'M02', 6.1, 7.2),
('HS02', 'M03', 8.5, 8.5),
('HS02', 'M04', 7.3, 7.8),
('HS02', 'M05', 6.9, 7.1),
('HS02', 'M06', 8.0, 7.5),
('HS02', 'M07', 7.2, 6.8),
('HS02', 'M08', 7.5, 7.0),
('HS02', 'M09', 8.2, 8.0),
-- HS03
('HS03', 'M01', 7.6, 7.5),
('HS03', 'M02', 9.5, 7.5),
('HS03', 'M03', 8.1, 9.1),
('HS03', 'M04', 7.8, 8.0),
('HS03', 'M05', 8.2, 7.7),
('HS03', 'M06', 7.5, 7.9),
('HS03', 'M07', 8.0, 7.3),
('HS03', 'M08', 7.7, 8.1),
('HS03', 'M09', 8.5, 8.3),
-- HS04
('HS04', 'M01', 9.3, 6.0),
('HS04', 'M02', 8.3, 5.3),
('HS04', 'M03', 7.8, 9.1),
('HS04', 'M04', 8.0, 7.5),
('HS04', 'M05', 7.4, 6.8),
('HS04', 'M06', 7.9, 7.2),
('HS04', 'M07', 6.5, 7.0),
('HS04', 'M08', 7.1, 6.9),
('HS04', 'M09', 8.3, 8.0),
-- HS05
('HS05', 'M01', 7.7, 8.1),
('HS05', 'M02', 9.8, 9.1),
('HS05', 'M03', 8.4, 7.6),
('HS05', 'M04', 7.9, 8.2),
('HS05', 'M05', 8.1, 7.8),
('HS05', 'M06', 8.3, 7.5),
('HS05', 'M07', 7.6, 7.2),
('HS05', 'M08', 8.0, 7.9),
('HS05', 'M09', 8.7, 8.5),
-- HS06
('HS06', 'M01', 8.6, 6.0),
('HS06', 'M02', 7.2, 8.2),
('HS06', 'M03', 8.1, 5.3),
('HS06', 'M04', 7.5, 6.8),
('HS06', 'M05', 7.0, 7.3),
('HS06', 'M06', 7.8, 7.1),
('HS06', 'M07', 6.9, 6.5),
('HS06', 'M08', 7.4, 7.0),
('HS06', 'M09', 8.0, 7.7);
Select* from monhoc
Select* from taikhoan
Select* from lop
Select* from hocsinh
Select* from diem

INSERT INTO TaiKhoan (UserName, PassWord, PhanQuyen)
VALUES 
('GV01', 'L0101', N'Giáo Viên'),
('GV02', 'L0102', N'Giáo Viên'),
('GV03', 'L0203', N'Giáo Viên'),
('CB01', 'CBQL01', N'Admin');