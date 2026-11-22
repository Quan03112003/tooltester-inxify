# BÁO CÁO ĐÁNH GIÁ CHUYÊN SÂU & KẾT LUẬN CUỐI CÙNG VỀ PHẦN MỀM QUẢN LÝ KHÁCH SẠN INXIFY PMS
## PHÂN TÍCH TÍNH SẴN SÀNG TRIỂN KHAI TOÀN CHUỖI (GO-LIVE) VÀ PHƯƠNG ÁN THÍ ĐIỂM CÓ ĐIỀU KIỆN TẠI 01 CHI NHÁNH

---

**Dự án:** Chuyển đổi và Nâng cấp Hệ thống Quản lý Khách sạn Inxify PMS – Tập đoàn Khách sạn A25  
**Vị trí thực hiện:** Hoàng Minh Quân – Chuyên viên kiểm thử phần mềm (Tester)  
**Đơn vị công tác:** Ban Phát triển & Quản lý phần mềm – Phòng Công nghệ thông tin  
**Người kiểm duyệt:** Nguyễn Văn Sơn – Trưởng ban Phát triển & Quản lý phần mềm  
**Kính trình:** Sếp, Ban Giám đốc Tập đoàn Khách sạn A25, Phòng Thanh tra giám sát & Ban Kiểm soát nội bộ  
**Thời điểm đánh giá:** Ngày 22 tháng 11 năm 2026  
**Môi trường kiểm thử đối soát:** UAT Portal (`https://uat.inxify.com`) & Backend API (`https://api-uat.inxify.com`)  

---

## LỜI MỞ ĐẦU & CĂN CỨ ĐÁNH GIÁ THỰC TẾ

Báo cáo này được lập dựa trên kết quả kiểm thử thực tế độc lập, khách quan của Chuyên viên kiểm thử Hoàng Minh Quân xuyên suốt giai đoạn thử việc tháng 11/2026 tại Phòng CNTT Tập đoàn Khách sạn A25. Quá trình kiểm thử bao gồm:
1. **Kiểm thử chức năng và nghiệp vụ:** Trực tiếp khảo sát 09 đầu mục nghiệp vụ trọng tâm, thiết kế và thực thi **59 Test Cases** trên môi trường UAT Portal (gồm 35 Test Cases chuyên sâu về phân hệ Audit Log - "Ưu tiên số 1" của Ban Giám đốc);
2. **Kiểm thử giao diện & Trải nghiệm người dùng (UI/UX):** Bóc tách đối chiếu 20 trường thông tin màn hình Đặt phòng PMS và 18 trường thông tin màn hình Khách hàng CRM;
3. **Kiểm thử hiệu năng và khả năng chịu tải (Performance Test):** Đo đạc tự động 15 API cốt lõi qua công cụ k6 v2.4.0 với 05 cấp độ tải (1 VU, 7 VUs, 100 VUs, 300 VUs và 500 VUs);
4. **Kiểm tra tính toàn vẹn dữ liệu và kiểm toán thất thoát:** Phân tích mã nguồn mạng (DevTools F12 Network/Payload) để xác thực độ tin cậy của dữ liệu giao dịch tài chính.

---

## PHẦN I. TÓM TẮT ĐIỀU HÀNH & KẾT LUẬN TRỰC DIỆN (EXECUTIVE SUMMARY)

Đứng trên góc độ chuyên môn kiểm thử phần mềm độc lập và bảo vệ an toàn vận hành, an toàn tài chính tối đa cho Tập đoàn Khách sạn A25, chuyên viên kiểm thử xin trả lời trực diện 02 câu hỏi trọng tâm của Ban Giám đốc như sau:

```
┌────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│                                 KẾT LUẬN ĐIỀU HÀNH TỔNG THỂ CỦA TESTER                                 │
├───────────────────────────────────┬───────────────────────────────────┬────────────────────────────────┤
│      CÂU HỎI ĐẶT RA CỦA SẾP       │        KẾT LUẬN TRỰC DIỆN         │       MỨC ĐỘ RỦI RO / KHUYẾN NGHỊ│
├───────────────────────────────────┼───────────────────────────────────┼────────────────────────────────┤
│ 1. Phần mềm Inxify PMS đã có thể  │       TUYỆT ĐỐI CHƯA NÊN          │ 🔴 RỦI RO BÁO ĐỘNG ĐỎ          │
│    đưa vào sử dụng ngay (Go-Live) │       (KHÔNG ĐỦ ĐIỀU KIỆN)        │ Thất thoát doanh thu, mất dấu  │
│    toàn bộ 30 khách sạn A25 chưa? │             [NO-GO]               │ kiểm toán, sập server giờ đỉnh.│
├───────────────────────────────────┼───────────────────────────────────┼────────────────────────────────┤
│ 2. Đã có thể triển khai thí điểm  │      CHƯA THỂ BẬT NGAY HÔM NAY    │ 🟡 RỦI RO CÓ THỂ KIỂM SOÁT     │
│    tại 01 chi nhánh (Pilot Run)   │  NHƯNG ĐỦ ĐIỀU KIỆN THÍ ĐIỂM      │ Chỉ thí điểm chạy song song    │
│    được chưa?                     │  CÓ ĐIỀU KIỆN TỪ NGÀY 01/12/2026  │ (Parallel Run) sau khi Dev fix │
│                                   │       [CONDITIONAL PILOT]         │ xong 11 bug trước 27/11/2026.  │
└───────────────────────────────────┴───────────────────────────────────┴────────────────────────────────┘
```

### 1. Về việc "Đưa vào sử dụng ngay lập tức cho toàn bộ chuỗi A25 (Go-Live)":
* **KẾT LUẬN: TUYỆT ĐỐI CHƯA ĐỦ ĐIỀU KIỆN (KHUYẾN NGHỊ: NO-GO TOÀN CHUỖI).**
* **Căn cứ cốt lõi:**
  - Hệ thống đang tồn tại **11 lỗi nghiêm trọng**, trong đó có lỗi chí mạng tại phân hệ **Audit Log (Mục 7) không lưu giá trị cũ (`oldValues = null`)** khi nhân viên sửa giá phòng, sửa tiền cọc, hoặc hủy booking. Điều này làm mất hoàn toàn công cụ kiểm toán, tạo lỗ hổng cho gian lận nội bộ và thất thoát doanh thu.
  - **Mục 8 (Hiển thị nguồn OTA & Mã booking)** và **Mục 9 (Chặn nhận phòng bẩn Dirty/OOO)** chưa được Dev lập trình lên môi trường UAT (14 Test Cases chưa được kiểm thử thực tế). Lễ tân có thể nhận khách vào phòng chưa dọn vệ sinh, gây thảm họa dịch vụ khách hàng.
  - Đo kiểm thử hiệu năng k6 chứng minh: Khi tải đạt **300 - 500 người dùng đồng thời (VUs)** (tương đương giờ cao điểm giao ca chiều toàn bộ 30 khách sạn), cơ sở dữ liệu Inxify bị cạn kiệt Connection Pool, tỷ lệ lỗi vọt lên **18.2% - 46.8%**, API bị treo tới 25 giây (Request Timeout). Nếu Go-Live ngay, hệ thống sẽ sập cục bộ trên toàn chuỗi.

### 2. Về việc "Triển khai thí điểm tại 01 chi nhánh (Pilot Run)":
* **KẾT LUẬN: CHƯA NÊN BẬT NGAY TRONG HÔM NAY (22/11/2026), NHƯNG ĐỦ ĐIỀU KIỆN THÍ ĐIỂM CÓ ĐIỀU KIỆN TỪ NGÀY 01/12/2026.**
* **Phương án triển khai:**
  - **Chi nhánh đề xuất thí điểm:** **A25 Hotel – Chi nhánh 684 Minh Khai, Hà Nội** (Quy mô vừa phải ~45 phòng, khách đa dạng, Lễ tân trẻ thao tác nhanh, có sẵn đầu đọc mã vạch 2D CCCD tại quầy, gần trung tâm thuận tiện cho đội IT/Tester hỗ trợ on-site).
  - **Cơ chế vận hành:** Chạy song song hai hệ thống (**Parallel Run / Dual Run**) giữa phần mềm HBSoft cũ và Inxify PMS mới trong **14 ngày (01/12/2026 – 14/12/2026)**.
  - **Điều kiện tiên quyết (Quality Gate):** Phía đối tác phát triển Inxify phải bàn giao bản sửa lỗi 11 bug và triển khai xong Mục 8, Mục 9 lên UAT trước ngày 27/11/2026 để Tester kiểm thử hồi quy (Regression Test) đạt tỷ lệ Pass 100% trong hai ngày 28-29/11/2026.

---

## PHẦN II. MA TRẬN ĐÁNH GIÁ 5 TIÊU CHÍ CHẤT LƯỢNG (QUALITY DIMENSIONS)
### TỔNG ĐIỂM HỆ THỐNG: 6.45 / 10 – XẾP LOẠI: KHÁ (CHƯA ĐẠT CHUẨN PRODUCTION)

Để Ban Giám đốc có cơ sở kỹ thuật định lượng minh bạch, hệ thống Inxify PMS được chấm điểm theo 05 trục tiêu chí chất lượng quốc tế (Quality Dimensions) kèm căn cứ thực tế chi tiết:

```
                      BIỂU ĐỒ RADAR 5 TRỤC CHẤT LƯỢNG INXIFY PMS
                             Chức năng (Functionality)
                                      [7.0]
                                     /     \
                                    /       \
             Độ ổn định [7.5]      /         \      [6.5] Trải nghiệm (UI/UX)
                   \              /     •     \              /
                    \            /    (6.45)   \            /
                     \          /               \          /
                      Kiểm toán [5.5] ───────── [6.0] Hiệu năng (Performance)
```

| STT | Trục tiêu chí đánh giá | Trọng số | Điểm (Thang 10) | Điểm quy đổi | Phân tích căn cứ chấm điểm chi tiết (Log bằng chứng thực tế) |
| :---: | :--- | :---: | :---: | :---: | :--- |
| **01** | **Chức năng & Nghiệp vụ**<br>*(Functionality & Business Fit)* | **25%** | **7.0 / 10**<br>*(Khá)* | **1.75** | • **Điểm đạt (+7.0):** Đã vận hành được 07/09 phân hệ nghiệp vụ chính (Sơ đồ phòng, Đặt phòng, Hồ sơ khách hàng, Buồng phòng, Quản lý kho, Thu ngân). Luồng nhập tay thông tin CCCD ĐÃ PASS (TC-QR-06).<br>• **Điểm trừ (-3.0):** Mục 8 (Nguồn OTA & Mã booking) và Mục 9 (Chặn nhận phòng bẩn Dirty/OOO) chưa được Dev lập trình lên UAT; Tính năng Quét mã QR CCCD tự động bằng phần cứng chưa có trên web; 14/59 Test Cases còn ở dạng kịch bản chờ kiểm thử. |
| **02** | **Trải nghiệm người dùng**<br>*(UI/UX & Operational Usability)* | **20%** | **6.5 / 10**<br>*(Trung bình)* | **1.30** | • **Điểm đạt (+6.5):** Giao diện thiết kế theo phong cách hiện đại (React Ant Design), sơ đồ phòng phân màu trực quan rõ ràng (Xanh lá - Phòng trống, Xanh dương - Đang ở, Cam - Phòng bẩn, Tím - Phòng hỏng).<br>• **Điểm trừ (-3.5):** Khi bấm nút "Lưu đặt phòng" mà thiếu trường bắt buộc, hệ thống KHÔNG HIỆN VIỀN ĐỎ ở ô bị thiếu khiến Lễ tân lúng túng; Màn hình danh sách Khách hàng (`/crm/guests`) không có nút "+ Thêm mới" trực tiếp mà phải mở Drawer; Thiếu phím tắt nhanh cho thao tác quầy; Lớp phủ giới thiệu Ant Tour che khuất các nút bấm chính. |
| **03** | **Hiệu năng & Khả năng chịu tải**<br>*(Performance & Scalability)* | **20%** | **6.0 / 10**<br>*(Trung bình)* | **1.20** | • **Điểm đạt (+6.0):** Đo k6 thực tế 15 API cốt lõi. Tải nhẹ (1-7 VUs) phản hồi siêu tốc (< 80ms cho 7/15 API); Khả năng tự phục hồi (Resilience) rất tốt: sau khi chịu đợt sốc tải 1,000 VUs, hệ thống hạ tải và thời gian phản hồi quay về bình thường 990ms trong 10 giây.<br>• **Điểm trừ (-4.0):** API Sơ đồ phòng (`/rooms/map`) chứa dữ liệu thô 115 KB chưa nén Gzip/Brotli, chưa cấu hình Redis Cache khiến p95 lên tới 4,653ms ở 100 VUs; Báo cáo điều hành mất 883ms; Tải 300 VUs lỗi 18.2%, 500 VUs lỗi 46.8% do cạn Database Connection Pool (Request Timeout 25s - BUG-003). |
| **04** | **Toàn vẹn dữ liệu & Kiểm toán**<br>*(Data Integrity & Audit Log)* | **20%** | **5.5 / 10**<br>*(Rủi ro cao)* | **1.10** | • **Điểm đạt (+5.5):** Đã xây dựng riêng phân hệ Nhật ký thao tác (Audit Log), ghi nhận được sự kiện "Ai sửa - Vào lúc nào" cho 05 nghiệp vụ trọng điểm.<br>• **Điểm trừ (-4.5):** **LỖ HỔNG LỚN NHẤT HỆ THỐNG**: 8/35 Test Case Audit Log Fail do **KHÔNG GHI NHẬN GIÁ TRỊ CŨ (`oldValues = null`)**, không biết nhân viên sửa từ giá bao nhiêu sang giá bao nhiêu; Form Đặt phòng chấp nhận lưu SĐT 5 số (09812), Email sai cú pháp; CRM cho phép tạo trùng lặp CMND/CCCD mà không cảnh báo gộp hồ sơ. |
| **05** | **Độ ổn định & Hạ tầng kỹ thuật**<br>*(Stability & Infrastructure)* | **15%** | **7.5 / 10**<br>*(Khá tốt)* | **1.10** | • **Điểm đạt (+7.5):** Hạ tầng máy chủ Cloud duy trì trạng thái trực tuyến tốt, không bị sập dịch vụ bất thường; Đã cấu hình Rate Limiting (HTTP 429) chống tấn công từ chối dịch vụ (DDoS) và spam brute-force mật khẩu.<br>• **Điểm trừ (-2.5):** Cấu hình Database Connection Pool của server UAT quá thấp (nghẽn tại 300 VUs); Chưa có cơ chế hàng đợi bất đồng bộ (Message Queue) khi xử lý các batch job lớn (như tính toán báo cáo đêm Night Audit). |
| 🎯 | **TỔNG ĐIỂM TRUNG BÌNH CÓ TRỌNG SỐ** | **100%** | **6.45 / 10** | **6.45** | **KẾT LUẬN TIÊU CHÍ:** Đạt ngưỡng cho phép tiếp tục hoàn thiện UAT (yêu cầu >= 6.0), nhưng **CÒN THẤP HƠN NHIỀU SO VỚI NGƯỠNG ĐỦ ĐIỀU KIỆN GO-LIVE CHÍNH THỨC (YÊU CẦU >= 8.0/10)**. |

---

## PHẦN III. 04 LÝ DO "TỬ HUYỆT" KHẲNG ĐỊNH TUYỆT ĐỐI CHƯA THỂ GO-LIVE TOÀN CHUỖI

Để Ban Giám đốc hình dung rõ ràng những hiểm họa tiềm ẩn nếu vội vã đưa Inxify PMS vào vận hành trên toàn bộ 30 khách sạn A25 trong thời điểm hiện tại, Tester xin phân tích chi tiết 04 "tử huyệt" kỹ thuật và vận hành:

```
┌───────────────────────────────────────────────────────────────────────────────────────────────────┐
│                           04 "TỬ HUYỆT" KHIẾN HỆ THỐNG CHƯA THỂ GO-LIVE                           │
├───────────────────────────────────────────────────────────────────────────────────────────────────┤
│ 1. NGUY CƠ THẤT THOÁT TÀI CHÍNH & MẤT DẤU VẾT KIỂM TOÁN (AUDIT LOG THIẾU 'OLDVALUES' - BUG-001)   │
│ 2. NGUY CƠ XUNG ĐỘT QUẦY LỄ TÂN & VỠ DỊCH VỤ KHÁCH HÀNG (CHƯA CHẶN NHẬN PHÒNG BẨN - MỤC 9)        │
│ 3. NGUY CƠ DỮ LIỆU RÁC & PHÂN MẢNH LỊCH SỬ KHÁCH HÀNG (VALIDATE SĐT 5 SỐ, TRÙNG CCCD)            │
│ 4. NGUY CƠ TÊ LIỆT HỆ THỐNG TOÀN CHUỖI GIỜ CAO ĐIỂM GIAO CA (DATABASE TIMEOUT 25S Ở 300 VUS)     │
└───────────────────────────────────────────────────────────────────────────────────────────────────┘
```

### 1. Tử huyệt 1: Lỗ hổng kiểm toán tài chính – Audit Log bị mất giá trị cũ (BUG-001)
* **Thực trạng phát hiện:** Trong đợt kiểm thử chuyên sâu phân hệ Audit Log (35 Test Cases hoàn thành ngày 20/11/2026), Tester phát hiện **08 ca kiểm thử bị FAIL** do API `/api/v1/audit-logs` trả về giá trị trường `oldValues` hoàn toàn là `null` hoặc chuỗi rỗng khi thực hiện thao tác Cập nhật (UPDATE).
* **Rủi ro thực tế đối với A25 Hotel:**
  - Khách sạn A25 vận hành theo mô hình tiền mặt và chuyển khoản liên tục tại quầy. Nếu một Lễ tân hoặc Thu ngân sửa giá phòng từ 1.200.000đ xuống 600.000đ, hoặc sửa tiền đặt cọc của khách đoàn từ 5.000.000đ thành 1.000.000đ rồi bỏ túi phần chênh lệch, hệ thống chỉ ghi lại giá trị mới là 600.000đ và hoàn toàn không lưu giá trị cũ là 1.200.000đ.
  - Khi Phòng Thanh tra giám sát hoặc Kiểm toán nội bộ vào cuộc đối soát, **hoàn toàn không có bằng chứng kỹ thuật để chứng minh nhân viên đã sửa từ bao nhiêu tiền sang bao nhiêu tiền**.
  - **Khẳng định:** Ban Giám đốc đã xác định "Mục 7 Audit Log là Ưu tiên số 1: Ai sửa – Sửa gì – Lúc nào". Khi tiêu chí này chưa đạt, Go-Live sẽ mở toang cánh cửa cho thất thoát tài chính.

### 2. Tử huyệt 2: Chưa có tính năng Chặn nhận phòng bẩn (Mục 9) và Thiếu nguồn kênh OTA (Mục 8)
* **Thực trạng phát hiện:** Hiện tại trên môi trường UAT, Dev Inxify chưa tích hợp mã nguồn của Văn bản tổng hợp yêu cầu kỹ thuật ngày 19/11/2026:
  - Thao tác check-in vào phòng bẩn (Dirty) hoặc phòng hỏng (Out of Order) vẫn diễn ra bình thường, hệ thống **không hề có bất kỳ cảnh báo màu vàng hay chặn chuyển trạng thái**;
  - Modal nhận phòng từ đặt phòng (`CheckInFromReservationModal`) chưa hiển thị Nguồn kênh đặt phòng và Mã booking tham chiếu OTA (Agoda, Booking.com, Traveloka).
* **Rủi ro thực tế đối với A25 Hotel:**
  - Nếu áp dụng ngay tại 30 chi nhánh với hàng trăm lượt check-in mỗi ngày, Lễ tân trong lúc vội giờ cao điểm sẽ vô tình gán khách VIP vào phòng vừa trả chưa dọn dẹp vệ sinh. Khách mở cửa thấy phòng bẩn sẽ lập tức khiếu nại, đánh giá 1 sao trên mạng xã hội, làm tổn hại nghiêm trọng uy tín thương hiệu A25 Hotel.
  - Việc không hiển thị mã OTA sẽ khiến Lễ tân không thể đối chiếu mã đặt phòng của khách với hệ thống kênh bán, dẫn đến tình trạng trùng phòng (Overbooking) hoặc không thu được tiền hoa hồng OTA.

### 3. Tử huyệt 3: Validate dữ liệu lỏng lẻo gây "ô nhiễm dữ liệu" toàn hệ thống
* **Thực trạng phát hiện:** Kết quả kiểm thử Mục 6 (Cảnh báo kiểm tra dữ liệu khi lưu) ghi nhận 03 lỗi nghiêm trọng:
  - Form Đặt phòng cho phép nhập số điện thoại chỉ có 5 chữ số (ví dụ: `09812`) và địa chỉ Email sai cú pháp chuẩn (ví dụ: `test@`) mà hệ thống vẫn ấn nút Lưu thành công;
  - Màn hình CRM Khách hàng cho phép tạo vô hạn các hồ sơ khách hàng mới có cùng số CCCD và trùng cả Họ tên, Ngày sinh mà không có cơ chế cảnh báo hoặc hợp nhất dữ liệu (Deduplication).
* **Rủi ro thực tế đối với A25 Hotel:**
  - Nếu đưa vào sử dụng ngay cho toàn chuỗi, chỉ sau 1 tháng vận hành, cơ sở dữ liệu A25 sẽ ngập tràn "dữ liệu rác".
  - Bộ phận Chăm sóc khách hàng và Kinh doanh không thể gửi tin nhắn SMS/Zalo chăm sóc khách quen, lịch sử tích điểm và thói quen lưu trú của khách VIP bị chia nhỏ thành nhiều hồ sơ khác nhau, làm phá sản chiến lược khai thác dữ liệu khách hàng số.

### 4. Tử huyệt 4: Nguy cơ sập hệ thống cục bộ vào giờ cao điểm (Nghẽn Database Connection Pool)
* **Thực trạng phát hiện:** Báo cáo kiểm thử tải k6 v2.4.0 ghi nhận kết quả thực tế trên 15 API:
  - Ở mức tải 100 VUs, API Sơ đồ phòng (`/rooms/map`) có độ trễ p95 lên đến **4,653 ms (hơn 4.6 giây)** vì trả về gói JSON thô nặng tới 115 KB do chưa bật nén HTTP Gzip/Brotli và chưa cấu hình bộ nhớ đệm Redis Cache;
  - Khi tăng tải lên mức **300 VUs và 500 VUs**, hệ thống xuất hiện hiện tượng cạn kiệt kết nối cơ sở dữ liệu (Connection Pool Exhaustion). Tỷ lệ yêu cầu bị lỗi vọt lên **18.20% (ở 300 VUs)** và **46.80% (ở 500 VUs)**, kèm theo hàng loạt lỗi **Request Timeout 25s (BUG-003)**.
* **Rủi ro thực tế đối với A25 Hotel:**
  - Tập đoàn A25 có 30 chi nhánh với gần 1.000 phòng hoạt động 24/7. Vào các khung giờ cao điểm như 14h00 (nhận phòng đồng loạt) và 12h00 trưa (trả phòng giao ca), số lượng nhân viên Lễ tân, Thu ngân, Buồng phòng và Quản lý truy cập đồng thời chắc chắn vượt ngưỡng 200 - 300 phiên kết nối.
  - Nếu hệ thống bị đơ 25 giây hoặc văng lỗi kết nối, quầy lễ tân toàn chuỗi sẽ bị tê liệt, khách hàng phải đứng chờ tại sảnh, tạo ra cuộc khủng hoảng vận hành trên quy mô toàn tập đoàn.

---

## PHẦN IV. ĐỀ ÁN THÍ ĐIỂM CÓ ĐIỀU KIỆN (CONDITIONAL PILOT) TẠI 01 CHI NHÁNH

Mặc dù việc Go-Live toàn chuỗi là **TUYỆT ĐỐI KHÔNG THỂ**, nhưng việc triển khai **THÍ ĐIỂM CÓ ĐIỀU KIỆN TẠI 01 CHI NHÁNH** là hoàn toàn khả thi và vô cùng cần thiết. Điều này giúp dự án vừa kiểm soát được rủi ro trong phạm vi hẹp, vừa thu thập được phản hồi thực tế từ quầy lễ tân để hoàn thiện phần mềm trước khi nhân rộng.

```
┌───────────────────────────────────────────────────────────────────────────────────────────────────┐
│                        MÔ HÌNH VẬN HÀNH THÍ ĐIỂM: CHẠY SONG SONG (PARALLEL RUN)                   │
├───────────────────────────────────────────────────────────────────────────────────────────────────┤
│                                                                                                   │
│                     KHÁCH HÀNG ĐẾN GIAO DỊCH TẠI QUẦY LỄ TÂN MINH KHAI                            │
│                                           │                                                       │
│                     ┌─────────────────────┴─────────────────────┐                                 │
│                     ▼                                           ▼                                 │
│          ┌───────────────────────┐                   ┌───────────────────────┐                    │
│          │     HỆ THỐNG CHÍNH    │                   │   HỆ THỐNG THÍ ĐIỂM   │                    │
│          │      HBSOFT (CŨ)      │                   │    INXIFY PMS (MỚI)   │                    │
│          ├───────────────────────┤                   ├───────────────────────┤                    │
│          │ • Cơ sở pháp lý chính │                   │ • Đánh giá tính năng  │                    │
│          │ • Quản lý tiền phòng  │                   │ • Test máy quét 2D    │                    │
│          │ • Xuất hóa đơn VAT    │                   │ • Đo tốc độ thao tác  │                    │
│          └──────────┬────────────┘                   └──────────┬────────────┘                    │
│                     │                                           │                                 │
│                     └─────────────────────┬─────────────────────┘                                 │
│                                           ▼                                                       │
│                           ĐỐI SOÁT CUỐI CA & BÁO CÁO NGÀY (DAILY RECONCILIATION)                  │
│                           • Doanh thu lệch = 0 đồng                                              │
│                           • Trạng thái phòng trùng khớp 100%                                     │
│                                                                                                   │
└───────────────────────────────────────────────────────────────────────────────────────────────────┘
```

### 1. Chi nhánh đề xuất lựa chọn thí điểm: **A25 Hotel – Chi nhánh 684 Minh Khai, Hà Nội**
* **Căn cứ lựa chọn:**
  - **Quy mô phù hợp:** Số lượng phòng ở mức trung bình (~45-50 phòng), khối lượng giao dịch vừa phải, không gây quá tải cho nhân sự;
  - **Đầy đủ phân khúc khách:** Có sự pha trộn đồng đều giữa khách lẻ theo giờ, khách lưu trú qua đêm và khách đặt qua kênh OTA;
  - **Nhân sự năng động:** Đội ngũ Lễ tân trẻ tuổi, tiếp thu phần mềm mới nhanh, tinh thần hợp tác cao;
  - **Cơ sở vật chất sẵn có:** Quầy lễ tân đã được trang bị sẵn **đầu đọc quét mã vạch 2D CCCD kết nối cổng USB**, đáp ứng ngay điều kiện kiểm thử phần cứng Mục 5;
  - **Khoảng cách địa lý:** Nằm gần trung tâm Hà Nội, giúp Chuyên viên kiểm thử và đội ngũ kỹ thuật Phòng CNTT có thể trực tiếp có mặt tại quầy để hỗ trợ on-site ngay khi có sự cố.

### 2. Cổng chất lượng tiên quyết (Quality Gates - Điều kiện bắt buộc để BẬT thí điểm):
Chi nhánh Minh Khai **CHỈ ĐƯỢC PHÉP BẬT THÍ ĐIỂM** khi và chỉ khi vượt qua 03 Cổng nghiệm thu kỹ thuật sau đây:

* **Cổng 1 (Phía Đội ngũ Dev Inxify – Hạn chót: 17h00 ngày 27/11/2026):**
  - Khắc phục triệt để lỗi BUG-001 của Audit Log: Trường `oldValues` bắt buộc phải trả về dữ liệu cũ trước khi sửa đổi;
  - Khắc phục các lỗi Validate của Mục 6: Chặn số điện thoại < 10 số, chặn email sai cú pháp, chặn tạo trùng lặp CCCD trên CRM;
  - Triển khai mã nguồn Mục 8 (Nguồn OTA) và Mục 9 (Chặn nhận phòng bẩn Dirty/OOO) lên môi trường UAT;
  - Cấu hình Redis Cache cho API sơ đồ phòng và tăng kích thước Database Connection Pool.

* **Cổng 2 (Phía Chuyên viên kiểm thử QC – Hạn chót: 17h00 ngày 30/11/2026):**
  - Tester Hoàng Minh Quân thực hiện kiểm thử hồi quy (Regression Testing) toàn bộ 59 Test Cases và 35 Test Cases Audit Log;
  - Tỷ lệ ca kiểm thử đạt (Pass Rate) phải đạt **100% đối với các lỗi mức Nghiêm trọng (Critical/High)** và đạt tối thiểu **95% trên tổng số ca test**.

* **Cổng 3 (Nghiệm thu phần cứng Quét QR CCCD – Ngày 28/11/2026):**
  - Tester trực tiếp xuống quầy lễ tân A25 Minh Khai, cắm đầu đọc 2D USB vào máy tính quầy và quét thử thẻ CCCD thật;
  - Đảm bảo thời gian nhận diện chuỗi 7 trường thông tin của Bộ Công an diễn ra dưới **1.0 giây** và tự động điền chính xác vào form khách hàng mà không bị lỗi font Tiếng Việt có dấu.

### 3. Phương thức vận hành thí điểm: Chạy song song (Parallel Run / Dual Run) trong 14 ngày
* **Thời gian thí điểm:** Từ ngày **01/12/2026 đến hết ngày 14/12/2026**.
* **Nguyên tắc vận hành bất di bất dịch:**
  - **Hệ thống HBSoft cũ giữ vai trò là CĂN CỨ PHÁP LÝ DUY NHẤT:** Toàn bộ việc thu tiền phòng, in hóa đơn VAT, đối soát tiền mặt với Kế toán và xuất báo cáo doanh thu tài chính vẫn được chốt 100% dựa trên HBSoft;
  - **Inxify PMS giữ vai trò là HỆ THỐNG THÍ ĐIỂM:** Nhân viên Lễ tân sau khi thao tác trên HBSoft sẽ thực hiện lại thao tác đó trên Inxify (Check-in, Check-out, gán phòng, đổi phòng, tạo dịch vụ giặt ủi/nước uống).
* **Quy trình đối soát cuối ca (Daily Reconciliation):**
  - Vào cuối mỗi ca làm việc (6h00, 14h00, 22h00), Trưởng ca Lễ tân chi nhánh Minh Khai xuất báo cáo doanh thu và danh sách phòng ở từ cả 2 phần mềm để đối chiếu chéo;
  - Nếu xuất hiện sai lệch giữa 2 hệ thống, Lễ tân ghi nhận ngay vào sổ nhật ký thí điểm và thông báo cho Tester/IT trong vòng 15 phút.

### 4. Ma trận rủi ro & Kế hoạch Rollback khẩn cấp (Khôi phục nguyên trạng):

| Tình huống rủi ro trong thí điểm | Mức độ rủi ro | Biện pháp xử lý ngay lập tức | Kế hoạch Rollback (Dừng thí điểm) |
| :--- | :---: | :--- | :--- |
| **Hệ thống Inxify bị sập hoặc treo > 15 phút** | Cao | Lễ tân dừng thao tác trên Inxify, tập trung 100% phục vụ khách trên HBSoft. IT kiểm tra log máy chủ. | Nếu sự cố kéo dài quá 2 giờ trong ngày, tạm dừng thí điểm 24h để Dev vá lỗi hạ tầng. |
| **Sai lệch số liệu doanh thu giữa 2 bên** | Rất cao | Giữ nguyên số liệu của HBSoft để thu tiền khách. Tester bóc tách log giao dịch trên Inxify để tìm nguyên nhân. | Nếu sai lệch do lỗi tính toán logic lõi của Inxify, dừng ngay thí điểm để Dev audit lại mã nguồn tính giá. |
| **Lễ tân bị quá tải vì phải nhập 2 hệ thống** | Trung bình | Phòng CNTT cử Chuyên viên kiểm thử trực tiếp ngồi tại quầy hỗ trợ nhập liệu Inxify trong 3 ngày đầu. | Bố trí thêm 01 nhân sự thực tập sinh/hỗ trợ tại quầy trong khung giờ cao điểm (12h - 14h). |

---

## PHẦN V. LỘ TRÌNH HÀNH ĐỘNG CHI TIẾT (ACTION ROADMAP 23/11 – 15/12/2026)

Để đảm bảo tiến độ công việc chuẩn xác và hoàn thành hồ sơ thử việc trước mốc cam kết **16h00 ngày 02/12/2026**, lộ trình hành động được thiết lập chi tiết theo từng ngày:

```
┌───────────────────────────────────────────────────────────────────────────────────────────────────┐
│                              LỘ TRÌNH TRIỂN KHAI VÀ NGHIỆM THU CHI TIẾT                           │
├────────────────────┬──────────────────────────────────────────────────────────────┬───────────────┤
│    KHOẢNG THỜI GIAN│                      NỘI DUNG CÔNG VIỆC TRỌNG TÂM            │  TRÁCH NHIỆM  │
├────────────────────┼──────────────────────────────────────────────────────────────┼───────────────┤
│ 23/11 – 26/11/2026 │ • Dev Inxify fix 11 bug hệ thống (đặc biệt oldValues Audit)  │ Dev Inxify    │
│                    │ • Tích hợp Mục 8 (OTA) và Mục 9 (Chặn phòng bẩn) lên UAT     │               │
│                    │ • Cấu hình Redis Cache cho sơ đồ phòng & tối ưu DB Pool      │               │
├────────────────────┼──────────────────────────────────────────────────────────────┼───────────────┤
│ 27/11/2026         │ • Deploy phiên bản hoàn chỉnh lên môi trường UAT Portal      │ Dev Inxify    │
│                    │ • Bàn giao Changelog và Release Notes cho Tester             │               │
├────────────────────┼──────────────────────────────────────────────────────────────┼───────────────┤
│ 28/11/2026         │ • Tester xuống Lễ tân A25 Minh Khai, mượn máy quét 2D USB    │ Tester Quân & │
│                    │ • Thực thi kiểm thử phần cứng quét mã QR CCCD thật (Mục 5)    │ Lễ tân MK     │
├────────────────────┼──────────────────────────────────────────────────────────────┼───────────────┤
│ 29/11 – 30/11/2026 │ • Thực thi kiểm thử hồi quy (Regression Testing) toàn bộ 59 TC│ Tester Quân & │
│                    │ • Đo lại hiệu năng k6 sau khi tối ưu Redis Cache             │ Trưởng ban Sơn│
│                    │ • Họp đánh giá Cổng chất lượng (Quality Gate) duyệt thí điểm │               │
├────────────────────┼──────────────────────────────────────────────────────────────┼───────────────┤
│ 01/12/2026         │ • Đào tạo nghiệp vụ 2 tiếng cho Lễ tân chi nhánh Minh Khai   │ Tester Quân & │
│                    │ • CHÍNH THỨC BẬT THÍ ĐIỂM CHẠY SONG SONG TẠI A25 MINH KHAI   │ IT A25        │
├────────────────────┼──────────────────────────────────────────────────────────────┼───────────────┤
│ TRƯỚC 16H00        │ • BÀN GIAO TOÀN BỘ HỒ SƠ BÁO CÁO KẾT THÚC THỬ VIỆC CHÍNH THỨC│ Tester Quân & │
│ NGÀY 02/12/2026    │ • Trình Ban Giám đốc & Phòng TCNS (đầy đủ 4 chỉ số định lượng│ Trưởng ban Sơn│
├────────────────────┼──────────────────────────────────────────────────────────────┼───────────────┤
│ 01/12 – 14/12/2026 │ • Vận hành thí điểm 14 ngày tại A25 Minh Khai                │ Đội dự án     │
│                    │ • Theo dõi log lỗi hằng ngày, đối soát doanh thu cuối ca     │ Inxify & A25  │
├────────────────────┼──────────────────────────────────────────────────────────────┼───────────────┤
│ 15/12/2026         │ • Họp tổng kết giai đoạn thí điểm chi nhánh Minh Khai         │ Ban Giám đốc  │
│                    │ • Quyết định phương án mở rộng ra 05 chi nhánh tiếp theo     │ & Các Trưởng  │
└────────────────────┴──────────────────────────────────────────────────────────────┴───────────────┘
```

---

## PHẦN VI. KIẾN NGHỊ VÀ ĐỀ XUẤT CỦA TESTER GỬI BAN GIÁM ĐỐC

Từ các phân tích chuyên sâu nêu trên, Chuyên viên kiểm thử Hoàng Minh Quân kính đề xuất Ban Giám đốc xem xét phê duyệt 03 nội dung mang tính quyết định:

1. **VỀ VIỆC GO-LIVE TOÀN CHUỖI:**
   * Kính đề xuất Ban Giám đốc **CHƯA PHÊ DUYỆT GO-LIVE TOÀN CHUỖI** trong tháng 11 và tháng 12/2026;
   * Chỉ xem xét Go-Live toàn chuỗi khi và chỉ khi điểm đánh giá chất lượng hệ thống đạt tối thiểu **8.0 / 10 điểm**, và giai đoạn thí điểm tại chi nhánh Minh Khai vận hành ổn định không phát sinh lỗi tài chính trong 14 ngày liên tục.

2. **VỀ VIỆC THÍ ĐIỂM TẠI 01 CHI NHÁNH:**
   * Kính đề xuất Ban Giám đốc **PHÊ DUYỆT CHỦ TRƯƠNG THÍ ĐIỂM CHẠY SONG SONG** tại **A25 Hotel – 684 Minh Khai** bắt đầu từ ngày **01/12/2026**;
   * Giao Phòng CNTT chủ trì, phối hợp cùng Quản lý chi nhánh Minh Khai tổ chức chạy song song với HBSoft và thực hiện đối soát doanh thu hằng ngày.

3. **VỀ PHÍA ĐỐI TÁC PHÁT TRIỂN PHẦN MỀM INXIFY:**
   * Yêu cầu đối tác ký cam kết bàn giao bản vá lỗi 11 bug trước ngày 27/11/2026;
   * Yêu cầu đối tác cử tối thiểu **01 Kỹ sư Lập trình Backend và 01 Chuyên viên Hỗ trợ trực chiến on-call 24/7** trong suốt 14 ngày thí điểm tại chi nhánh Minh Khai để xử lý tức thời các sự cố phát sinh.

---

## BẢNG KÝ DUYỆT CÁC CẤP

| Đơn vị / Vị trí ký duyệt | Họ tên & Chữ ký xác nhận | Ý kiến chỉ đạo & Thời gian duyệt |
| :--- | :--- | :--- |
| **Người lập báo cáo** | **Hoàng Minh Quân**<br>Chuyên viên kiểm thử phần mềm<br>*(Đã ký)* | *Cam kết số liệu kiểm thử trung thực, khách quan và chịu trách nhiệm về các phát hiện lỗi.*<br>Ngày 22/11/2026 |
| **Ban Phát triển & Quản lý phần mềm** | **Nguyễn Văn Sơn**<br>Trưởng ban<br>*(Đã ký)* | *Thống nhất với đánh giá của Tester: Chưa thể Go-Live ngay; Đồng ý đề xuất thí điểm Minh Khai.*<br>Ngày 22/11/2026 |
| **Phòng Công nghệ thông tin** | **Lãnh đạo Phòng CNTT**<br>*(Ký và ghi rõ họ tên)* | ............................................................................<br>Ngày ...../...../2026 |
| **Phòng Thanh tra giám sát** | **Trưởng phòng Thanh tra**<br>*(Ký và ghi rõ họ tên)* | ............................................................................<br>Ngày ...../...../2026 |
| **Ban Giám đốc Tập đoàn** | **Tổng Giám đốc / Sếp**<br>*(Phê duyệt chủ trương)* | ............................................................................<br>Ngày ...../...../2026 |

---
*Tài liệu nội bộ thuộc Ban Quản lý & Phát triển phần mềm – Phòng CNTT Tập đoàn Khách sạn A25. Nghiêm cấm sao chép, phát tán ra ngoài tổ chức khi chưa được phê duyệt.*
