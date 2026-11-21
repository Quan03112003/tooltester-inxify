# BÁO CÁO TỔNG HỢP TIẾN ĐỘ THỰC HIỆN, CÁC VẤN ĐỀ ĐÃ XỬ LÝ, TỒN ĐỌNG VÀ PHƯƠNG ÁN GIẢI QUYẾT
**Dự án: Kiểm thử phần mềm Quản lý Khách sạn Inxify PMS & Chuẩn hóa Hệ sinh thái số A25 Hotel**  
**Vị trí thực hiện:** Hoàng Minh Quân – Chuyên viên kiểm thử phần mềm (Tester)  
**Đơn vị:** Ban Phát triển và Quản lý phần mềm – Phòng Công nghệ thông tin  
**Người kiểm duyệt & Phụ trách số liệu:** Nguyễn Văn Sơn – Trưởng ban  
**Thời điểm tổng hợp:** Ngày 21 tháng 11 năm 2026  

---

> ### ⚠️ ĐẶC THÙ & ĐIỀU KIỆN THỰC TẾ CỦA TESTER TRONG DỰ ÁN
> Hiện tại ở vị trí thử việc, chuyên viên kiểm thử **CHỈ ĐƯỢC CUNG CẤP DUY NHẤT MÔI TRƯỜNG UAT PORTAL (`https://uat.inxify.com`)**, hoàn toàn:
> - **CHƯA CÓ tài liệu đặc tả yêu cầu nghiệp vụ (PRD / SRS / BRD)** chính thức được phê duyệt;
> - **CHƯA CÓ Từ điển dữ liệu (Data Dictionary)** hay tài liệu mô tả trường (Data Mapping);
> - **CHƯA CÓ Bộ dữ liệu kiểm thử chuẩn (Standard Test Data)** mô phỏng đầy đủ dữ liệu vận hành chuỗi A25;
> - **CHƯA CÓ thiết bị phần cứng kiểm thử** (Đầu đọc quét mã vạch 2D CCCD cổng USB, thẻ CCCD cứng thật).
>
> Toàn bộ quá trình kiểm thử, lập kịch bản, đối chiếu trường dữ liệu và phát hiện lỗi đều dựa trên **sự chủ động khảo sát trực tiếp giao diện web UAT, kiểm tra mã nguồn giao diện / API payload qua DevTools (F12) và vận dụng tư duy logic nghiệp vụ kiểm thử phần mềm** của Tester kết hợp các buổi xác nhận quy tắc với Mentor/BA.

---

## PHẦN I. DANH MỤC TOÀN BỘ CÁC ĐẦU MỤC CÔNG VIỆC TRONG DỰ ÁN

Toàn bộ công việc của chuyên viên kiểm thử trong giai đoạn vừa qua được phân thành 03 nhóm trọng tâm:

1. **Nhóm I: Kiểm thử Nghiệm thu Hệ thống Inxify PMS (Theo chỉ đạo HROnline tháng 11/2026):**
   * *Đầu mục 1 (Mục 1 HROnline):* Thiết lập môi trường UAT Inxify độc lập & Khóa quyền sửa/xóa dữ liệu thật trên HBSoft.
   * *Đầu mục 2 (Mục 2 HROnline):* Lập bảng đối chiếu dữ liệu theo từng trường cho 02 màn hình Đặt phòng PMS và Khách hàng CRM.
   * *Đầu mục 3 (Mục 5 HROnline):* Xây dựng kịch bản kiểm thử nghiệm thu tính năng Quét mã QR Căn cước công dân (CCCD).
   * *Đầu mục 4 (Mục 6 HROnline):* Kiểm thử nghiệm thu tính năng Cảnh báo kiểm tra dữ liệu khi lưu (thiếu trường, sai định dạng, trùng lặp).
   * *Đầu mục 5 (Mục 7 HROnline - Ưu tiên số 1):* Kiểm thử nghiệm thu chức năng Nhật ký thao tác theo trường (Audit Log - "Ai sửa, sửa gì, lúc nào") cho 05 nghiệp vụ trọng tâm.
   * *Đầu mục 6 (Mục 8 HROnline):* Văn bản tổng hợp yêu cầu kỹ thuật & Kịch bản nghiệm thu tính năng Hiển thị nguồn kênh đặt phòng và Mã tham chiếu OTA.
   * *Đầu mục 7 (Mục 9 HROnline):* Văn bản tổng hợp yêu cầu kỹ thuật & Kịch bản nghiệm thu tính năng Cảnh báo và kiểm soát khi nhận phòng chưa sẵn sàng.

2. **Nhóm II: Chuẩn hóa Quy trình Tiếp nhận, Kiểm soát & Quản trị Phiếu yêu cầu phần mềm:**
   * *Đầu mục 8 (Mục 3 HROnline):* Thiết lập cơ chế kiểm soát đầu vào, kiên quyết từ chối tiếp nhận và trả lại phiếu thiếu trường bắt buộc, không chọn loại phản ánh. Báo cáo định kỳ thứ Hai hằng tuần.
   * *Đầu mục 9 (Mục 4 HROnline):* Lập và theo dõi vòng đời tối thiểu 03 phiếu yêu cầu phần mềm chuẩn hóa trên Sổ theo dõi hằng ngày từ 18/11/2026.

3. **Nhóm III: Chuẩn hóa Môi trường Làm việc Số & Học liệu Đào tạo Microsoft 365 (Nhiệm vụ phối hợp):**
   * *Đầu mục 10:* Lập Tờ trình & Báo cáo phương án chuẩn hóa Microsoft 365 (Outlook, Teams, OneDrive, Entra ID), phân định ranh giới độc tôn HROnline và lộ trình 7 ngày đóng nhóm Zalo văn phòng.
   * *Đầu mục 11:* Xây dựng Bộ học liệu đa phương tiện: Video đào tạo thực tế MP4 HD 720p (02:05) giọng đọc Neural Microsoft (`vi-VN-NamMinhNeural`), Trình chiếu tương tác HTML5 và hướng dẫn thao tác từng bước.

---

## PHẦN II. CHI TIẾT CÁC VẤN ĐỀ ĐÃ XỬ LÝ (KÈM KẾT QUẢ ĐỊNH LƯỢNG & BẰNG CHỨNG)

| STT | Đầu mục công việc | Kết quả đã xử lý thực tế | Số liệu định lượng & Sản phẩm bàn giao | Tình trạng |
| :---: | :--- | :--- | :--- | :---: |
| **1** | **Môi trường UAT & Bảo vệ dữ liệu thật (Mục 1)** | - Tiếp nhận bàn giao 09 tài khoản phân quyền trên UAT Inxify (`https://uat.inxify.com/login`) ngày 19/11/2026.<br>- Ký cam kết bảo mật thông tin nội bộ ngày 18/11/2026.<br>- Khóa quyền sửa/xóa tài khoản dữ liệu thật trên HBSoft, chuyển sang chế độ Chỉ xem (Read-only). | - 09 tài khoản phân quyền: IT, Employee, Accountant, Housekeeping, Secretary, Warehouse, OTA Manager, Inspector, Receptionist.<br>- Biên bản cam kết bảo mật nộp Phòng TCNS. | **ĐÃ HOÀN THÀNH 100%** |
| **2** | **Bảng đối chiếu 02 màn hình Inxify (Mục 2)** | - Hoàn thành trước hạn 02/12/2026.<br>- Tự khảo sát mã nguồn web, soi cấu trúc API F12 để phân tích chi tiết từng trường dữ liệu (tên trường, giá trị nhập, giá trị kỳ vọng, giá trị thực tế, đạt/không đạt). | - Bảng đối chiếu 20 trường màn hình Đặt phòng PMS (`DOI_CHIEU_MH1`).<br>- Bảng đối chiếu 18 trường màn hình Khách hàng CRM (`DOI_CHIEU_MH2`).<br>- Lưu tại file `Mau_Excel_QC_Inxify_DoiChieu_va_TheoDoiPhieu_Hoan_Chinh.xlsx`. | **ĐÃ HOÀN THÀNH 100%** |
| **3** | **Kiểm soát tiếp nhận phiếu & Trả lại phiếu sai (Mục 3)** | - Áp dụng quy tắc từ chối phiếu từ 18/11/2026.<br>- Phát hiện và lập biên bản trả lại 01 phiếu yêu cầu không chọn loại phản ánh, không nêu bước tái hiện của Phòng Kinh doanh.<br>- Thiết lập công thức báo cáo tuần tự động trên Excel. | - Biên bản trả lại phiếu `YC-A25-20261118-004`.<br>- Bảng thống kê tuần tại sheet `BAO_CAO_TUAN` (Nhận: 4, Trả lại: 1, Đã đóng: 2, Thời gian xử lý: 1 ngày/phiếu). | **ĐÃ HOÀN THÀNH 100%** |
| **4** | **Lập và theo dõi vòng đời phiếu thực tế (Mục 4)** | - Lập và theo dõi 04 phiếu thực tế (vượt chỉ tiêu tối thiểu 3 phiếu của Sếp).<br>- Theo dõi trạng thái từng ngày từ 18/11/2026 đến khi đóng phiếu. | - 04 phiếu trên sổ `THEO_DOI_PHIEU`:<br>  + YC-001 (Lỗi form đặt phòng) -> ĐÃ ĐÓNG.<br>  + YC-002 (Lỗi validate SĐT) -> ĐÃ ĐÓNG.<br>  + YC-003 (Yêu cầu hiển thị mã OTA) -> Đang xử lý.<br>  + YC-004 (Phiếu kinh doanh) -> Đã trả lại. | **ĐÃ HOÀN THÀNH 100%** |
| **5** | **Kịch bản Quét QR CCCD (Mục 5 - Phần Kịch bản & Nhập tay)** | - Nghiên cứu cấu trúc chuỗi QR CCCD 7 trường chuẩn Bộ Công an.<br>- Lập bộ kịch bản kiểm thử nghiệm thu chi tiết.<br>- Kiểm tra xác nhận luồng nhập liệu thủ công bằng tay (Manual Input) hiện tại trên Inxify hoạt động bình thường (ĐÃ PASS). | - Tài liệu [Tai_lieu_Kich_ban_Kiem_thu_Nghiem_thu_Quet_QR_CCCD_Inxify_Muc_5.docx](file:///c:/Users/Admin/Desktop/tester/Tai_lieu_Kich_ban_Kiem_thu_Nghiem_thu_Quet_QR_CCCD_Inxify_Muc_5.docx).<br>- 06 Kịch bản kiểm thử nghiệm thu (TC-QR-01 đến TC-QR-06) kèm bộ dữ liệu giả lập (Mock Data). | **ĐÃ HOÀN THÀNH KỊCH BẢN** |
| **6** | **Cảnh báo lưu dữ liệu khi lưu (Mục 6)** | - Thiết kế 10 kịch bản kiểm thử nghiệm thu.<br>- Thực thi 100% trên môi trường UAT Inxify.<br>- Phát hiện 03 lỗi nghiệp vụ/validate dữ liệu.<br>- Thu thập 07 hình ảnh dẫn chứng chụp trực tiếp từ UAT. | - Báo cáo nghiệm thu [Bao_cao_Kiem_thu_Canh_bao_Luu_Du_lieu_Inxify_Muc_6_Co_Anh_Dan_Chung.docx](file:///c:/Users/Admin/Desktop/tester/Bao_cao_Kiem_thu_Canh_bao_Luu_Du_lieu_Inxify_Muc_6_Co_Anh_Dan_Chung.docx).<br>- 10 Test Cases (7 Pass, 3 Fail).<br>- 07 ảnh dẫn chứng UAT (M6_01 đến M6_07). | **ĐÃ HOÀN THÀNH 100%** |
| **7** | **Nhật ký thao tác Audit Log (Mục 7 - Ưu tiên số 1)** | - Tiếp nhận Đặc tả BA v0.3 từ BA Châu Tuấn.<br>- Xây dựng Ma trận trường dữ liệu và Bộ kịch bản cho 05 nghiệp vụ trọng tâm (Thu chi, Hoàn tiền, Đặt cọc, Sửa/Hủy đặt phòng, Công nợ).<br>- Thực thi 100% trên UAT Inxify, phát hiện 08 lỗi ghi thiếu trường cũ/mới.<br>- Báo cáo Ban Giám đốc đúng hạn 20/11/2026. | - Báo cáo trình Ban Giám đốc [Bao_cao_Trinh_Ban_Giam_Doc_Ket_Qua_Nghiem_Thu_Audit_Log_Inxify_Muc_7.docx](file:///c:/Users/Admin/Desktop/tester/Bao_cao_Trinh_Ban_Giam_Doc_Ket_Qua_Nghiem_Thu_Audit_Log_Inxify_Muc_7.docx).<br>- File Excel [Template_TestCase_Audit_Log_Inxify_Thang_11_2026_Hoan_chinh.xlsx](file:///c:/Users/Admin/Desktop/tester/Template_TestCase_Audit_Log_Inxify_Thang_11_2026_Hoan_chinh.xlsx).<br>- 35 Test Cases (27 Pass, 8 Fail). | **ĐÃ HOÀN THÀNH 100%** |
| **8 & 9** | **Nguồn kênh OTA & Cảnh báo phòng chưa sẵn sàng (Mục 8 & 9)** | - Khảo sát trực tiếp giao diện Inxify UAT.<br>- Quy chuẩn hóa 05 yêu cầu kỹ thuật chi tiết (`REQ-M8-01` đến `REQ-M9-02`).<br>- Xây dựng 08 kịch bản kiểm thử nghiệm thu đón đầu cho Dev. | - Văn bản kỹ thuật [Van_ban_Tong_hop_Yeu_cau_Ky_thuat_Nghiem_thu_Inxify_Muc_8_va_Muc_9_Co_Anh_Dan_Chung.docx](file:///c:/Users/Admin/Desktop/tester/Van_ban_Tong_hop_Yeu_cau_Ky_thuat_Nghiem_thu_Inxify_Muc_8_va_Muc_9_Co_Anh_Dan_Chung.docx).<br>- 05 Tiêu chuẩn kỹ thuật + 08 Test Cases. | **ĐÃ HOÀN THÀNH 100%** |
| **10 & 11** | **Chuẩn hóa Microsoft 365 & Học liệu Video Đào tạo** | - Hoàn thiện Tờ trình phương án M365 (dự toán 8.6 triệu/tháng, ranh giới HROnline, quy tắc 7 ngày đóng Zalo).<br>- Sản xuất Video bài giảng thực hành MP4 HD 720p (02:05) thuyết minh giọng Neural Microsoft chuẩn xác, thao tác liên tục và phân Role rõ ràng.<br>- Xây dựng Trình duyệt video tương tác HTML5. | - Báo cáo Word [Bao_cao_Phuong_an_Chuan_hoa_Cong_cu_Microsoft_365_A25_Hotel.docx](file:///c:/Users/Admin/Desktop/tester/Bao_cao_Phuong_an_Chuan_hoa_Cong_cu_Microsoft_365_A25_Hotel.docx) (4.03 MB).<br>- Video [Video_Huong_Dan_Chuan_Hoa_Microsoft_365_A25_Hotel.mp4](file:///c:/Users/Admin/Desktop/tester/Video_Huong_Dan_Chuan_Hoa_Microsoft_365_A25_Hotel.mp4).<br>- Trình phát [Trung_Tam_Video_Huong_Dan_M365_A25_Hotel.html](file:///c:/Users/Admin/Desktop/tester/Trung_Tam_Video_Huong_Dan_M365_A25_Hotel.html). | **ĐÃ HOÀN THÀNH 100%** |

---

## PHẦN III. CHI TIẾT CÁC VẤN ĐỀ CÒN TỒN ĐỌNG & PHÂN TÍCH NGUYÊN NHÂN

Đứng trên vai trò của **một Tester thực chiến chỉ được cấp môi trường UAT Portal**, chuyên viên kiểm thử xin chỉ rõ 06 vấn đề tồn đọng cốt lõi đang ảnh hưởng trực tiếp đến chất lượng và tiến độ dự án:

### 1. Tồn đọng 1: Thiếu tài liệu đặc tả nghiệp vụ chính thức (PRD / SRS / Data Dictionary)
* **Hiện trạng thực tế:** Ngoài tài liệu BA v0.3 (mới ở mức dự thảo sơ bộ cho chức năng Audit Log), toàn bộ các phân hệ còn lại của Inxify (Đặt phòng, Quản lý buồng phòng, Lễ tân check-in/check-out, Thu ngân, Công nợ, Khách hàng) đều **hoàn toàn chưa có tài liệu PRD/SRS chính thức**.
* **Nguyên nhân:** Dự án Inxify đang trong giai đoạn chuyển đổi gấp rút từ hệ thống cũ, đội ngũ BA chưa kịp chuẩn hóa và ban hành hệ thống tài liệu yêu cầu nghiệp vụ.
* **Tác động đến công tác kiểm thử:** Tester phải tự "mò mẫm" trên giao diện web UAT, kiểm tra mã nguồn và payload API qua DevTools để tự suy đoán logic nghiệp vụ và tự định nghĩa Expected Result. Điều này tiềm ẩn nguy cơ sai lệch giữa cách hiểu của Tester với thực tế vận hành tại quầy của khách sạn, dễ dẫn đến tranh cãi giữa Tester và Dev khi bắt lỗi.

### 2. Tồn đọng 2: Môi trường UAT thiếu Bộ dữ liệu kiểm thử chuẩn (Standard Test Data)
* **Hiện trạng thực tế:** Cơ sở dữ liệu trên môi trường UAT Inxify còn rất mới, dữ liệu chắp vá hoặc rỗng. Thiếu hẳn các bộ dữ liệu mẫu mang tính đại diện cho chuỗi A25 (ví dụ: các hạng phòng Standard/Deluxe/Suite, bảng giá theo khung giờ/ngày đêm, danh sách khách đoàn, đối tác lữ hành, công nợ OTA, lịch sử lưu trú cũ).
* **Nguyên nhân:** Chưa có quy trình trích xuất và đồng bộ dữ liệu mẫu (đã ẩn danh) từ môi trường vận hành thật sang môi trường kiểm thử.
* **Tác động đến công tác kiểm thử:** Tester phải mất rất nhiều thời gian tự nhập tay tạo dữ liệu mồi thủ công từng bước; dữ liệu không đồng nhất khiến việc kiểm thử các kịch bản liên hoàn (End-to-End) bị gián đoạn và không thể thực hiện kiểm thử hiệu năng/chịu tải.

### 3. Tồn đọng 3: Tính năng Quét QR CCCD (Mục 5) chưa thể kiểm thử phần cứng và tự động hóa
* **Hiện trạng thực tế:** 
  * Tại bàn làm việc thử việc của Tester **chưa có đầu đọc quét mã vạch 2D cầm tay (kết nối cổng USB)**;
  * Tester **không có thẻ CCCD cứng thật** của khách hàng (nhằm tuân thủ nghiêm ngặt quy định bảo vệ dữ liệu cá nhân theo Nghị định 13/2023/NĐ-CP);
  * Phần mềm Inxify **chưa có tính năng/ứng dụng liên kết camera điện thoại**;
  * Trên giao diện web Inxify hiện tại **mới chỉ có duy nhất form nhập liệu thủ công bằng tay (Manual Input)**, Dev Inxify chưa ra mắt tính năng nhận diện chuỗi quét QR.
* **Nguyên nhân:** Thiết bị máy quét 2D chuyên dụng đang được trang bị tập trung tại các quầy lễ tân chi nhánh đang kinh doanh; phía Dev Inxify chưa bàn giao phiên bản có tích hợp thư viện quét mã lên UAT.
* **Tác động đến công tác kiểm thử:** Hiện tại Tester mới chỉ kiểm thử được luồng nhập tay thủ công (TC-QR-06) và viết bộ kịch bản nghiệm thu đón đầu bằng chuỗi dữ liệu giả lập (Mock Data); chưa thể kiểm thử tốc độ nhận diện phần cứng, chưa test được trường hợp thẻ mờ/xước/lóa sáng ngoài thực tế.

### 4. Tồn đọng 4: 02 tính năng Hiển thị nguồn OTA (Mục 8) và Cảnh báo nhận phòng chưa sẵn sàng (Mục 9) chưa được Dev lập trình lên UAT
* **Hiện trạng thực tế:** Mặc dù Tester đã hoàn thành và nộp Văn bản tổng hợp yêu cầu kỹ thuật ngày 19/11/2026, nhưng trên môi trường UAT Inxify hiện tại:
  * Modal Nhận phòng từ đặt phòng (`CheckInFromReservationModal`) vẫn chưa được bổ sung 02 trường: Nguồn kênh đặt phòng và Mã booking OTA;
  * Thao tác nhận phòng vào phòng bẩn (Dirty) hoặc phòng đang sửa chữa (Out of Order) vẫn diễn ra bình thường, hệ thống chưa hề có cảnh báo màu vàng hay chặn chuyển trạng thái.
* **Nguyên nhân:** Đội ngũ Lập trình viên (Dev) Inxify đang tập trung xử lý các module khác, chưa cập nhật mã nguồn của Mục 8 và Mục 9 lên môi trường UAT.
* **Tác động đến công tác kiểm thử:** 08 kịch bản kiểm thử nghiệm thu mới chỉ nằm ở dạng kịch bản thiết kế trên tài liệu; Tester chưa thể bấm chạy thực tế để kết luận Đạt/Không đạt.

### 5. Tồn đọng 5: 11 lỗi hệ thống đã phát hiện trên UAT Inxify đang chờ Dev khắc phục triệt để
* **Hiện trạng thực tế:** Trong các đợt kiểm thử vừa qua, Tester đã phát hiện và ghi nhận chính thức 11 lỗi hệ thống:
  * **08 lỗi chức năng Audit Log (Mục 7):** Thiếu thông tin trường cũ/trường mới khi sửa dữ liệu, sai định dạng thời gian ISO, thiếu log thao tác hủy cọc phòng, thiếu ghi nhận IP thiết bị ở một số phân hệ.
  * **03 lỗi Validate dữ liệu khi lưu (Mục 6):** Bỏ trống tên khách hàng vẫn lưu được, bỏ trống giá phòng vẫn tạo được đặt phòng, nhập số điện thoại sai định dạng (< 10 số) hệ thống không chặn.
* **Nguyên nhân:** Mã nguồn phía Backend/Frontend của đối tác phát triển Inxify chưa được kiểm soát chặt chẽ ở tầng Data Validation và Event Listener.
* **Tác động đến công tác kiểm thử:** Nguy cơ sai lệch số liệu và không đáp ứng tiêu chuẩn nghiệm thu "Ưu tiên số 1: Ai sửa – Sửa gì – Lúc nào" của Ban Giám đốc nếu không được fix sớm.

### 6. Tồn đọng 6: Quy trình xử lý phiếu yêu cầu phần mềm còn mang tính thủ công, thói quen phòng ban chưa thay đổi
* **Hiện trạng thực tế:** Việc tiếp nhận, ghi chép và theo dõi vòng đời phiếu yêu cầu hiện vẫn phải quản lý thủ công qua file Excel (`THEO_DOI_PHIEU`); nhân sự các phòng ban (Kinh doanh, Lễ tân) vẫn giữ thói quen cũ là nhắn tin Zalo hoặc phản ánh miệng, khi được yêu cầu điền form phiếu thì còn lúng túng hoặc điền thiếu thông tin.
* **Nguyên nhân:** Hệ thống Inxify/HROnline chưa có phân hệ Helpdesk/Ticketing điện tử tập trung; quy chế sử dụng công cụ mới đang trong giai đoạn triển khai thí điểm.

---

## PHẦN IV. PHƯƠNG ÁN GIẢI QUYẾT CHI TIẾT CHO CÁC VẤN ĐỀ TỒN ĐỌNG

Để giải quyết dứt điểm các vướng mắc trên và bảo đảm chất lượng nghiệm thu hệ thống trước mốc thời gian **16h00 ngày 02/12/2026**, chuyên viên kiểm thử phối hợp cùng Trưởng ban Nguyễn Văn Sơn đề xuất các phương án giải quyết cụ thể như sau:

```mermaid
flowchart TD
    A["VẤN ĐỀ TỒN ĐỌNG THỰC TẾ"] --> B["1. Thiếu PRD & Data Dictionary"]
    A --> C["2. Thiếu dữ liệu test chuẩn UAT"]
    A --> D["3. Chưa có máy quét 2D & thẻ CCCD"]
    A --> E["4. Mục 8, 9 chưa code & 11 Bug chưa fix"]
    A --> F["5. Quy trình phiếu còn thủ công"]

    B --> B1["Lấy Bảng đối chiếu trường làm Baseline; Họp chốt quy tắc với BA & Mentor"]
    C --> C1["Đề xuất Masking DB từ Production nạp vào UAT; Duy trì Mock Data giả lập"]
    D --> D1["Mượn 01 máy quét 2D từ Lễ tân A25 Minh Khai; Test 30 phút bằng CCCD công vụ"]
    E --> E1["Họp kỹ thuật với Dev chốt Release trước 27/11; QA chạy Regression Test"]
    F --> F1["Kiên quyết trả lại phiếu Zalo/nói miệng; Đề xuất module Ticketing trên Inxify"]
```

### 1. Phương án giải quyết Tồn đọng 1 (Khắc phục việc thiếu tài liệu PRD/SRS):
* **Giải pháp ngắn hạn (Đang thực hiện):** Tester sử dụng chính các Bảng đối chiếu trường dữ liệu (20 trường Đặt phòng, 18 trường Khách hàng) và Ma trận trường Audit Log đã lập làm tài liệu kỹ thuật cơ sở (Baseline Specs); gửi trực tiếp cho BA Châu Tuấn và Trưởng ban Sơn xác nhận tính chính xác của các trường dữ liệu và quy tắc nghiệp vụ.
* **Giải pháp dài hạn:** Kiến nghị Ban Giám đốc chỉ đạo bộ phận BA hoàn thiện và ban hành chính thức Bộ tài liệu đặc tả chức năng (SRS) chuẩn hóa cho toàn bộ hệ thống Inxify PMS trước khi nghiệm thu diện rộng.

### 2. Phương án giải quyết Tồn đọng 2 (Khắc phục việc thiếu Bộ dữ liệu kiểm thử chuẩn):
* **Đề xuất kỹ thuật:** Phòng CNTT phối hợp cùng bộ phận Quản trị dữ liệu trích xuất một bản sao cơ sở dữ liệu (Database Snapshot) từ hệ thống vận hành thật; thực hiện quy trình **Ẩn danh hóa dữ liệu (Data Masking)** để xóa/mã hóa toàn bộ thông tin nhạy cảm của khách hàng (Số CCCD, SĐT, Email, Tên thật) rồi nạp vào môi trường UAT Inxify.
* **Thời hạn đề xuất hoàn thành:** Trước ngày **26/11/2026** để Tester có đủ dữ liệu chạy kiểm thử hồi quy và đối chiếu số liệu báo cáo kế toán.

### 3. Phương án giải quyết Tồn đọng 3 (Phương án kiểm thử phần cứng Quét QR CCCD - Mục 5):
* **Phương án phối hợp thực tế:** Khi đội ngũ Dev Inxify thông báo hoàn thành tính năng nhận diện chuỗi quét QR trên giao diện UAT:
  1. Phòng CNTT sẽ làm thủ tục mượn 01 đầu đọc mã vạch 2D chuyên dụng (cổng USB) từ quầy Lễ tân chi nhánh **A25 Hotel (cơ sở 684 Minh Khai)** kết nối trực tiếp vào máy tính văn phòng;
  2. Sử dụng thẻ CCCD công vụ của nội bộ nhân sự Phòng CNTT (được sự đồng thuận của nhân sự) để tiến hành quét thử nghiệm thực tế;
  3. Thời gian thực hiện: Gói gọn trong **01 buổi làm việc (khoảng 30 – 45 phút)** tại văn phòng để chạy đủ 06 kịch bản nghiệm thu đã lập, ghi nhận biên bản nghiệm thu phần cứng và bàn giao trả lại thiết bị cho lễ tân ngay trong ngày.

### 4. Phương án giải quyết Tồn đọng 4 & 5 (Xử lý 11 lỗi hệ thống & Đẩy nhanh tiến độ Mục 8, 9):
* **Lập danh sách bàn giao Dev:** Tester tổng hợp toàn bộ 11 lỗi (kèm mã phiếu yêu cầu, bước tái hiện, ảnh chụp màn hình dẫn chứng và mã lỗi JSON) gửi trực tiếp cho Trưởng ban Sơn và đội ngũ Lập trình viên Inxify.
* **Chốt mốc thời gian (Deadline) với Dev:**
  * **Trước ngày 25/11/2026:** Dev hoàn thành sửa chữa 11 lỗi hệ thống (ưu tiên lỗi ghi thiếu trường Audit Log và lỗi bỏ trống tên/giá vẫn lưu được).
  * **Trước ngày 27/11/2026:** Dev triển khai bản cập nhật mã nguồn của Mục 8 (Hiển thị mã OTA) và Mục 9 (Chặn nhận phòng bẩn) lên môi trường UAT.
  * **Từ ngày 28/11 đến 30/11/2026:** Tester thực thi kiểm thử hồi quy (Regression Testing) toàn bộ 45 kịch bản cũ và chạy nghiệm thu thực tế 14 kịch bản mới, đảm bảo 100% kịch bản đều có kết quả PASS/FAIL rõ ràng trước ngày 01/12/2026.

### 5. Phương án giải quyết Tồn đọng 6 (Quy chuẩn tiếp nhận phiếu & Quản trị tập trung):
* **Kỷ luật tiếp nhận:** Phòng CNTT tiếp tục thực hiện nghiêm túc chỉ đạo của Ban Giám đốc tại Mục 3: Kiên quyết từ chối tiếp nhận và lập biên bản trả lại 100% các yêu cầu gửi qua Zalo hoặc phản ánh miệng; chỉ xử lý các phiếu được lập đúng Biểu mẫu yêu cầu phần mềm đã ban hành.
* **Tích hợp phần mềm:** Về lâu dài, kiến nghị tích hợp thẳng tính năng gửi phiếu hỗ trợ (Helpdesk Ticket) ngay trên cổng Portal Inxify hoặc hệ thống HROnline để người dùng gửi trực tiếp từ phần mềm, chấm dứt việc nhập liệu thủ công qua file Excel.

---

## PHẦN V. KẾ HOẠCH HÀNH ĐỘNG VÀ CAM KẾT MỐC BÁO CÁO KẾT THÚC THỬ VIỆC

| STT | Nội dung công việc trọng tâm | Đầu mối thực hiện | Thời hạn hoàn thành | Kết quả đầu ra dự kiến |
| :---: | :--- | :--- | :---: | :--- |
| 1 | Bàn giao danh mục 11 lỗi và tiêu chuẩn kỹ thuật Mục 8, 9 cho đội ngũ Dev Inxify | Tester Hoàng Minh Quân phối hợp Trưởng ban Sơn | **22/11/2026** | Biên bản bàn giao lỗi & Kế hoạch fix bug |
| 2 | Nạp dữ liệu mẫu chuẩn (đã Masking) lên môi trường UAT Inxify | Phòng CNTT phối hợp Quản trị hệ thống | **25/11/2026** | Môi trường UAT có đầy đủ dữ liệu mẫu A25 |
| 3 | Dev bàn giao bản fix 11 lỗi và mã nguồn tính năng Mục 8, Mục 9 lên UAT | Đội ngũ Lập trình viên Inxify | **27/11/2026** | UAT Portal được cập nhật phiên bản mới |
| 4 | Mượn máy quét 2D tại Lễ tân Minh Khai, nghiệm thu tính năng Quét QR CCCD | Tester Hoàng Minh Quân | **28/11/2026** | Biên bản kiểm thử thiết bị phần cứng thực tế |
| 5 | Thực thi kiểm thử hồi quy (Regression Test) toàn bộ 59 Test Cases trên UAT | Tester Hoàng Minh Quân | **29/11 – 30/11/2026** | Bảng kết quả kiểm thử hoàn chỉnh 100% |
| 6 | Hoàn thiện Báo cáo kết thúc thử việc chính thức (đầy đủ 4 chỉ số định lượng) | Tester Hoàng Minh Quân & Trưởng ban Sơn | **Trước 16h00 ngày 02/12/2026** | Báo cáo chính thức trình Ban Giám đốc & HR |

**CAM KẾT CỦA ỨNG VIÊN & ĐƠN VỊ:**  
Trưởng ban Nguyễn Văn Sơn và ứng viên Hoàng Minh Quân cam kết chịu trách nhiệm toàn diện về tính trung thực, chính xác của toàn bộ số liệu báo cáo, khắc phục triệt để các tồn đọng và nộp hồ sơ kết thúc thử việc hoàn chỉnh **trước 16h00 ngày 02/12/2026 (tuyệt đối không để sang ngày 03/12/2026)** đúng như chỉ đạo của Sếp.
