# BÁO CÁO KẾT QUẢ KIỂM TRA THỰC TẾ TRÊN UAT PORTAL INXIFY (LẦN CUỐI)
**Hệ thống:** Inxify PMS UAT (`https://uat.inxify.com`)  
**Chi nhánh:** A25 Signature Hotel - 684 Minh Khai  
**Chuyên viên kiểm thử:** Hoàng Minh Quân (Tester)  
**Người kiểm duyệt:** Nguyễn Văn Sơn – Trưởng Ban Phát triển & Quản lý phần mềm  
**Thời điểm thực hiện:** Ngày 21 tháng 11 năm 2026  
**Thư mục lưu trữ dẫn chứng:** `c:\Users\Admin\Desktop\tester\kiem_tra_thuc_te_uat_lan_cuoi\Anh_dan_chung_Input_Den_Ket_Qua\`  

---

## 🎯 1. MỤC ĐÍCH & PHẠM VI KIỂM TRA THỰC TẾ
Thực hiện yêu cầu rà soát lần cuối trước khi trình Ban Tổng giám đốc và đối soát với phản hồi của HR, Chuyên viên kiểm thử đã trực tiếp đăng nhập môi trường UAT Portal web thật của Inxify để thực hiện chuỗi kịch bản kiểm thử thực tế, chụp ảnh đối soát từng bước:
- **Dữ liệu đầu vào (Input Data / Thao tác người dùng)**
- **Kết quả phản hồi của hệ thống (Result / Giao diện / Thông báo / Dữ liệu ghi nhận)**

Qua đó xác thực tính chính xác tuyệt đối, khớp 100% giữa số liệu báo cáo văn bản và hành vi thực tế của hệ thống.

---

## 📸 2. BẢNG ĐỐI CHIẾU CHI TIẾT 09 ĐẦU MỤC: INPUT ➔ RESULT

| STT | Đầu mục công việc (hronline.md) | Thao tác & Dữ liệu đầu vào (Input) | Kết quả phản hồi thực tế (Result) | Tệp ảnh minh chứng (Input & Result) | Đánh giá khớp báo cáo |
| :---: | :--- | :--- | :--- | :--- | :---: |
| **1** | **Mục 1: Phân quyền 09 tài khoản UAT & Bảo vệ dữ liệu** | - Đăng nhập tài khoản IT: `minhquan1` / `@Matkhau1`<br>- Đăng nhập tài khoản Lễ tân: `quanhm2` / `@Matkhau1` | - IT thấy 4 phân hệ (PMS, Quản lý giá, Hệ thống, Học viện).<br>- Lễ tân bị phân quyền, ẩn phân hệ quản trị Hệ thống.<br>- Đã ký cam kết NDA và khóa sửa/xóa HBSoft. | `M1_01_Login_IT_Input.png`<br>`M1_02_Login_IT_Portal_Result.png`<br>`M1_03_Login_Letan_Input.png`<br>`M1_04_Login_Letan_Portal_Result.png` | **ĐẠT (PASS)**<br>Khớp 100% |
| **2** | **Mục 2: Bảng đối chiếu dữ liệu 02 màn hình Inxify** | - Mở form Tạo đặt phòng mới (`/pms/reservations`)<br>- Mở Drawer hồ sơ khách hàng (`/crm/guests`) | - Form Đặt phòng hiển thị đầy đủ 20 trường thông tin.<br>- Drawer Khách hàng hiển thị đầy đủ 18 trường định danh.<br>- Đã lập bảng mapping chi tiết lưu tại file Excel. | `M2_01_DatPhong_Form_20_Truong.png`<br>`M2_02_KhachHang_Form_18_Truong.png` | **ĐẠT (PASS)**<br>Khớp 100% |
| **3** | **Mục 3 & 4: Phiếu yêu cầu & Xác minh 2 Bug YC-001, YC-002** | - **Bug YC-001**: Bỏ trống ô Tên khách hàng hoặc Giá phòng.<br>- **Bug YC-002**: Nhập số điện thoại sai định dạng `09812` (dưới 10 số). | - **YC-001**: Ô tên khách viền đỏ cảnh báo bắt buộc, hệ thống chặn không cho lưu.<br>- **YC-002**: Báo lỗi định dạng SĐT không hợp lệ.<br>- Đã lập 04 phiếu thực tế theo dõi trên Excel. | `M3_M4_01_Test_Bug_YC001_Input.png`<br>`M3_M4_02_Test_Bug_YC001_Result.png`<br>`M3_M4_03_Test_Bug_YC002_Input.png`<br>`M3_M4_04_Test_Bug_YC002_Result.png` | **ĐÃ FIX (CLOSED)**<br>Khớp 100% |
| **4** | **Mục 5: Kịch bản kiểm thử nghiệm thu Quét QR CCCD** | - Khảo sát giao diện Form Khách hàng.<br>- Nhập tay thủ công: Họ tên, Số CCCD, Ngày sinh, Địa chỉ (TC-QR-06). | - **Hiện trạng web UAT**: Chỉ có form nhập tay thủ công, hoàn toàn **chưa có nút Quét QR hay mở camera**.<br>- Tại bàn làm việc chưa có máy đọc 2D và CCCD mẫu.<br>- Nhập tay thủ công lưu thành công (TC-QR-06 PASS). | `M5_01_KhachHang_Form_Nhap_Tay_Input.png`<br>`M5_02_KhachHang_Luu_Thanh_Cong_Result.png` | **ĐÚNG HIỆN TRẠNG**<br>Khớp 100% |
| **5** | **Mục 6: Cảnh báo và kiểm tra dữ liệu khi lưu** | - Để trống các trường bắt buộc (Tên khách, Ngày đến, Giá).<br>- Nhập dữ liệu biên và email sai định dạng. | - Form kích hoạt viền đỏ cảnh báo các trường bắt buộc.<br>- Phát hiện 03 bất cập nghiệp vụ và đã kiến nghị chuẩn hóa quy tắc kiểm tra. | `M6_01_CanhBao_ThieuTruong_Input.png`<br>`M6_02_CanhBao_ThieuTruong_Viendo_Result.png` | **ĐÃ BẮT 3 BUG**<br>Khớp 100% |
| **6** | **Mục 7: Nhật ký thao tác Audit Log (Ưu tiên 1)** | - Đăng nhập IT, vào `/admin/audit-log`.<br>- Lọc theo 05 nghiệp vụ trọng tâm (Thu chi, Hoàn tiền, Đặt cọc, Sửa/hủy đặt phòng, Công nợ).<br>- Xem chi tiết bản ghi sửa dữ liệu. | - Hệ thống ghi nhận log nhưng **BỊ LỖI THIẾU GIÁ TRỊ CŨ (Old value = null/rỗng)** khi sửa.<br>- Kết quả thực thi 35 Test case: **27 PASS, 08 FAIL** (08 lỗi đã lập phiếu YC-A25-20261119-001). | `M7_01_AuditLog_BoLoc_5_NghiepVu_Input.png`<br>`M7_02_AuditLog_Loi_Thieu_TruongCu_Result.png` | **27 PASS, 8 FAIL**<br>Khớp 100% |
| **7** | **Mục 8: Nguồn kênh OTA và Mã tham chiếu** | - Mở danh sách đặt phòng có nguồn OTA.<br>- Nhấp nút `Nhận phòng` (Check-in) trên dòng đặt phòng OTA. | - Popup `Nhận phòng từ đặt phòng` hiển thị nhưng **THIẾU TRƯỜNG NGUỒN KÊNH VÀ MÃ BOOKING OTA**.<br>- Đã lập văn bản yêu cầu kỹ thuật và 08 kịch bản nghiệm thu chờ Dev triển khai. | `M8_01_DanhSach_Booking_OTA_Input.png`<br>`M8_02_Modal_NhanPhong_Thieu_Truong_OTA_Result.png` | **ĐÃ LẬP YÊU CẦU**<br>Khớp 100% |
| **8** | **Mục 9: Cảnh báo nhận phòng vào phòng chưa sẵn sàng** | - Truy cập Sơ đồ phòng (`/pms/room-map`).<br>- Chọn phòng có trạng thái Bẩn (Dirty - màu cam).<br>- Thao tác nhận phòng trực tiếp. | - Hệ thống **CHO PHÉP NHẬN PHÒNG THẲNG MÀ KHÔNG HỀ CÓ CẢNH BÁO CHẶN LẠI**.<br>- Đã lập tiêu chuẩn kỹ thuật chặn phòng bẩn/sửa chữa và đề xuất cơ chế vượt quyền Trưởng ca. | `M9_01_SoDoPhong_Phong_Ban_Dirty_Input.png`<br>`M9_02_NhanPhong_ThanhCong_Khong_CanhBao_Result.png` | **ĐÃ LẬP YÊU CẦU**<br>Khớp 100% |

---

## 📌 3. KẾT LUẬN & ĐÁNH GIÁ CỦA CHUYÊN VIÊN KIỂM THỬ
1. **Tính xác thực của hồ sơ báo cáo:** 100% dữ liệu, kịch bản, và các lỗi được nêu trong các báo cáo phát hành trước đó đều được kiểm chứng và tái hiện chính xác trên môi trường web UAT Portal thật.
2. **Khó khăn khách quan của vị trí kiểm thử:**
   - Ứng viên chỉ được bàn giao quyền truy cập web UAT; không được cung cấp tài liệu đặc tả nghiệp vụ (PRD/SRS), không có Data Dictionary chính thức, không có bộ dữ liệu mẫu chuẩn (phải tự tạo dữ liệu giả lập).
   - Chưa được trang bị máy quét mã vạch 2D và thẻ CCCD cứng tại bàn làm việc (phải kiểm thử luồng nhập tay thủ công thay thế).
3. **Giá trị đóng góp:** Ứng viên đã chủ động sử dụng DevTools/API để bóc tách payload, phát hiện 08 lỗi nghiêm trọng về kiểm toán tài chính (Audit Log), xác minh 02 bug giao diện đặt phòng đã được fix, và thiết lập đầy đủ khung kịch bản nghiệm thu cho các tính năng mới của Inxify.
