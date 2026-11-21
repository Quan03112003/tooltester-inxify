import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import datetime

FILE_PATH = r"c:\Users\Admin\Desktop\tester\Mau_Excel_QC_Inxify_DoiChieu_va_TheoDoiPhieu.xlsx"
FILE_BACKUP = r"c:\Users\Admin\Desktop\tester\Mau_Excel_QC_Inxify_DoiChieu_va_TheoDoiPhieu_Hoan_Chinh.xlsx"

def apply_resolution_updates():
    wb = openpyxl.load_workbook(FILE_PATH)

    # Styles
    font_arial_10 = Font(name="Arial", size=10)
    font_arial_10_bold = Font(name="Arial", size=10, bold=True)
    font_arial_11_bold = Font(name="Arial", size=11, bold=True)
    font_arial_14_bold = Font(name="Arial", size=14, bold=True, color="1E3A8A")
    font_arial_9_italic = Font(name="Arial", size=9, italic=True, color="4B5563")

    thin_side = Side(border_style="thin", color="D1D5DB")
    border_thin = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    fill_header_blue = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    fill_section_blue = PatternFill(start_color="E0E7FF", end_color="E0E7FF", fill_type="solid")
    fill_sub_header = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    fill_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_wrap_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # =========================================================================
    # 1. CẬP NHẬT SHEET 'THEO_DOI_PHIEU': THÊM CỘT 15 (HƯỚNG GIẢI QUYẾT)
    # =========================================================================
    ws_td = wb["THEO_DOI_PHIEU"]

    # Header for Col 15 (Column O)
    ws_td.cell(3, 15, "Hướng giải quyết / Phương án xử lý").font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    ws_td.cell(3, 15).fill = fill_header_blue
    ws_td.cell(3, 15).border = border_thin
    ws_td.cell(3, 15).alignment = align_center

    resolutions_td = {
        4: "• Kỹ thuật (IT/Dev): Cài đặt thuộc tính required: true và rule validation viền đỏ trên Ant Design; Backend bổ sung DTO middleware trả HTTP 400.\n• Vận hành (Lễ tân): Lễ tân bắt buộc nhập đủ Tên khách và Giá/đêm trước khi lưu; hệ thống tự khóa lưu để chống sót dữ liệu giao ca.",
        5: "• Kỹ thuật (IT/Dev): Tích hợp biểu thức chính quy (Regex) chuẩn SĐT di động VN (^0\\d{9}$) và Email RFC 5322; cảnh báo đỏ real-time khi gõ.\n• Vận hành (Lễ tân): Hướng dẫn Lễ tân quy chuẩn nhập 10 chữ số để phục vụ tự động gửi tin nhắn Zalo/SMS và tích điểm thân thiết CRM.",
        6: "• Kỹ thuật (IT/Dev): Lập Văn bản yêu cầu kỹ thuật (theo Mục 8 HROnline), phối hợp Dev bổ sung trường Nguồn OTA & Mã tham chiếu lên Popup nhận phòng.\n• Vận hành tạm thời (Lễ tân): Lễ tân đối soát mã voucher tại cột 'Nguồn' và 'Mã ĐL/OTA' trên danh sách đặt phòng trước khi bấm nhận phòng.",
        7: "• Phương án xử lý (Phòng ban): Phòng CNTT gửi thông báo trả lại phiếu kèm Biểu mẫu chuẩn; hướng dẫn chuyên viên Kinh doanh chọn loại phản ánh và mô tả rõ màn hình/lỗi trước khi gửi lại theo đúng Mục 3 HROnline."
    }

    for row_idx, res_text in resolutions_td.items():
        cell = ws_td.cell(row_idx, 15, res_text)
        cell.font = font_arial_10
        cell.border = border_thin
        cell.alignment = align_wrap_left
        cell.fill = fill_zebra if row_idx % 2 == 1 else fill_white

    ws_td.column_dimensions["O"].width = 50

    # =========================================================================
    # 2. CẬP NHẬT SHEET 'BIEU_MAU_PHIEU_YC': BỔ SUNG MỤC IV (HƯỚNG GIẢI QUYẾT)
    # =========================================================================
    sheet_name_bm = "BIEU_MAU_PHIEU_YC"
    if sheet_name_bm in wb.sheetnames:
        del wb[sheet_name_bm]
    ws_bm = wb.create_sheet(title=sheet_name_bm)

    def style_range(ws, start_col, start_row, end_col, end_row, font=None, fill=None, border=None, alignment=None):
        for r in range(start_row, end_row + 1):
            for c in range(start_col, end_col + 1):
                cell = ws.cell(r, c)
                if font: cell.font = font
                if fill: cell.fill = fill
                if border: cell.border = border
                if alignment: cell.alignment = alignment

    # Header
    ws_bm.merge_cells("A1:G1")
    ws_bm["A1"] = "BỘ BIỂU MẪU YÊU CẦU PHẦN MỀM CHUẨN HÓA (MỤC 4 HRONLINE)"
    ws_bm["A1"].font = font_arial_14_bold
    ws_bm["A1"].alignment = align_center

    ws_bm.merge_cells("A2:G2")
    ws_bm["A2"] = "Thực hiện nhiệm vụ Mục 4: Ứng viên Hoàng Minh Quân lập các phiếu yêu cầu thực tế (kèm hướng giải quyết cho phòng ban và kỹ thuật) theo Biểu mẫu chuẩn ban hành."
    ws_bm["A2"].font = font_arial_9_italic
    ws_bm["A2"].alignment = align_center

    forms_full_data = [
        {
            "num": "01",
            "ma": "YC-INX-20261118-001",
            "ngay_lap": "18/11/2026",
            "don_vi": "Bộ phận Lễ tân – Khách sạn A25 Signature (684 Minh Khai)",
            "nguoi_yc": "Hoàng Minh Quân (Chuyên viên Kiểm thử phần mềm)",
            "chuc_vu": "Chuyên viên QA / Ban PT & QLPM",
            "lien_he": "quanhm.cntt@a25hotel.com | 0989.xxx.xxx",
            "uu_tien": "Cao (Ảnh hưởng dữ liệu đặt phòng & vận hành bàn giao ca lễ tân)",
            "loai": "[x] Báo lỗi phần mềm (Bug)      [ ] Yêu cầu cải tiến (Change Request)      [ ] Hỗ trợ nghiệp vụ",
            "phan_he": "Quản lý Đặt phòng PMS (Màn hình: https://uat.inxify.com/pms/reservations)",
            "tieu_de": "Lỗi chặn lưu khi bỏ trống các trường bắt buộc (Họ tên khách, Giá phòng) trên modal Đặt phòng mới",
            "hien_trang": "Khi lễ tân mở modal 'Tạo đặt phòng mới', nếu bỏ trống trường Họ tên khách (guestName) hoặc Giá phòng/đêm (ratePerNight), hệ thống cần có cơ chế cảnh báo trực quan viền đỏ và thông báo cụ thể để nhân viên không lưu thiếu thông tin.",
            "buoc_tai_hien": "1. Truy cập https://uat.inxify.com/pms/reservations -> Bấm nút '+ Đặt phòng mới'.\n2. Nhập các thông tin khác nhưng cố tình để trống ô 'Tên khách' hoặc ô 'Giá/đêm'.\n3. Bấm nút 'Lưu' hoặc di chuyển chuột ra ngoài trường nhập liệu (Blur/Tab).",
            "du_lieu_test": "- Tên khách: [Để trống]\n- Giá/đêm: [Để trống]\n- Loại phòng: Deluxe (P.1407)\n- Thời gian: 20/11/2026 - 22/11/2026",
            "ky_vong": "Hệ thống chặn submit, hiển thị viền đỏ bao quanh ô Tên khách và Giá/đêm; xuất hiện dòng chữ cảnh báo màu đỏ: 'Vui lòng nhập tên khách' và 'Vui lòng nhập giá'.",
            "thuc_te_uat": "ĐÃ ĐẠT (PASS): Giao diện Ant Design trên UAT Inxify kích hoạt đúng rule validation, viền đỏ ô nhập và hiện cảnh báo đỏ chuẩn xác. Người dùng không thể gửi form khi thiếu trường bắt buộc.",
            "kiem_tra_cntt": "Đủ trường bắt buộc: [x] ĐẠT (Có)     |     Phân loại phản ánh: [x] ĐẠT (Có chọn 'Bug')     |     Kết luận tiếp nhận: [x] HỢP LỆ, TIẾP NHẬN",
            "phan_cong": "Châu Tuấn (Dev) xử lý đồng bộ Backend DTO; Hoàng Minh Quân (QA) thực thi kiểm thử nghiệm thu.",
            "ngay_nhan": "18/11/2026",
            "ngay_dong": "19/11/2026 (Thời gian xử lý: 01 ngày)",
            "trang_thai": "ĐÃ ĐÓNG (Nghiệm thu đạt 100%)",
            "ket_luan": "Đã nghiệm thu đạt trên UAT Inxify. Đính kèm ảnh minh chứng: M6_01_DatPhong_CanhBao_Thieu_Truong_Bat_Buoc.png.",
            "huong_giai_quyet": [
                ("1. Phương án kỹ thuật (IT/Dev):", "Cài đặt thuộc tính required: true và rules: [{ required: true, message: 'Vui lòng nhập...' }] trực tiếp trên form Ant Design. Tầng Backend bổ sung DTO Validation Middleware từ chối lưu và trả mã HTTP 400 Bad Request nếu thiếu trường guestName hoặc ratePerNight."),
                ("2. Hướng giải quyết cho Lễ tân:", "Lễ tân bắt buộc phải nhập đầy đủ Tên khách và Giá phòng trước khi lưu. Khi thiếu thông tin, form sẽ tự động viền đỏ và báo lỗi ngay tại ô nhập liệu, giúp nhân viên không bị sót thông tin khách và tránh chênh lệch giá khi bàn giao ca."),
                ("3. Kế hoạch kiểm thử & Đóng phiếu:", "Chuyên viên QA (Hoàng Minh Quân) thực thi kiểm thử biên trên UAT Inxify ngày 19/11/2026, xác nhận tính năng chặn lỗi hoạt động hoàn hảo. Đủ điều kiện đóng phiếu và bàn giao vận hành.")
            ]
        },
        {
            "num": "02",
            "ma": "YC-INX-20261119-002",
            "ngay_lap": "19/11/2026",
            "don_vi": "Bộ phận Lễ tân – Khách sạn A25 Signature (684 Minh Khai)",
            "nguoi_yc": "Hoàng Minh Quân (Chuyên viên Kiểm thử phần mềm)",
            "chuc_vu": "Chuyên viên QA / Ban PT & QLPM",
            "lien_he": "quanhm.cntt@a25hotel.com | 0989.xxx.xxx",
            "uu_tien": "Cao (Kiểm soát chất lượng dữ liệu khách hàng & liên lạc)",
            "loai": "[x] Báo lỗi phần mềm (Bug)      [ ] Yêu cầu cải tiến (Change Request)      [ ] Hỗ trợ nghiệp vụ",
            "phan_he": "Quản lý Đặt phòng PMS (Màn hình: https://uat.inxify.com/pms/reservations)",
            "tieu_de": "Kiểm tra và cảnh báo sai định dạng dữ liệu Số điện thoại và Email trên form Đặt phòng",
            "hien_trang": "Lễ tân nhập nhanh thông tin khách hàng có thể gõ sai định dạng Số điện thoại (dưới 10 số, không bắt đầu bằng số 0 như '09812') hoặc Email sai cấu trúc (như 'email_sai_dinh_dang@@'). Cần cơ chế regex chặn ngay tại giao diện.",
            "buoc_tai_hien": "1. Mở modal 'Tạo đặt phòng mới' tại /pms/reservations.\n2. Tại ô Số điện thoại: Nhập '09812' (chỉ có 5 số).\n3. Tại ô Email: Nhập 'email_sai_dinh_dang@@' (thiếu domain chuẩn).\n4. Bấm Tab hoặc bấm Lưu.",
            "du_lieu_test": "- SĐT test lỗi: '09812' hoặc '12345'\n- Email test lỗi: 'email_sai_dinh_dang@@' hoặc 'notanemail'\n- CCCD test: '123'",
            "ky_vong": "Giao diện chặn lưu, xuất hiện thông báo đỏ cảnh báo rõ ràng:\n+ 'SĐT 10 số, bắt đầu bằng 0'\n+ 'Email không đúng định dạng'\n+ Chú thích CCCD cần đủ 12 chữ số.",
            "thuc_te_uat": "ĐÃ ĐẠT (PASS trên Frontend): Form hiển thị cảnh báo đỏ chính xác như kỳ vọng. (Lưu ý QA: Đã lưu ý bộ phận Dev đồng bộ kiểm tra DTO phía Backend API để chống lọt dữ liệu qua tool ngoài).",
            "kiem_tra_cntt": "Đủ trường bắt buộc: [x] ĐẠT (Có)     |     Phân loại phản ánh: [x] ĐẠT (Có chọn 'Bug')     |     Kết luận tiếp nhận: [x] HỢP LỆ, TIẾP NHẬN",
            "phan_cong": "Châu Tuấn (Dev) tối ưu biểu thức chính quy (Regex); Hoàng Minh Quân (QA) thực thi kiểm thử biên.",
            "ngay_nhan": "19/11/2026",
            "ngay_dong": "20/11/2026 (Thời gian xử lý: 01 ngày)",
            "trang_thai": "ĐÃ ĐÓNG (Nghiệm thu đạt 100%)",
            "ket_luan": "Đã nghiệm thu đạt trên giao diện UAT Inxify. Đính kèm ảnh minh chứng: M6_02_DatPhong_CanhBao_Sai_Dinh_Dang_SDT_Email.png.",
            "huong_giai_quyet": [
                ("1. Phương án kỹ thuật (IT/Dev):", "Cấu hình Regex biểu thức chính quy kiểm tra SĐT Việt Nam: ^0\\d{9}$ (bắt buộc đúng 10 chữ số, ký tự đầu tiên là chữ số 0). Cấu hình Regex chuẩn RFC 5322 kiểm tra địa chỉ Email có đủ ký tự @ và tên miền hợp lệ."),
                ("2. Hướng giải quyết cho Lễ tân:", "Hướng dẫn Lễ tân quy chuẩn nhập số điện thoại 10 số của khách hàng ngay khi đặt phòng. Dữ liệu chuẩn giúp hệ thống gửi tin nhắn xác nhận tự động (SMS/Zalo ZNS) và tích lũy lịch sử lưu trú, điểm thưởng hạng thẻ chính xác trên CRM."),
                ("3. Kế hoạch kiểm thử & Đóng phiếu:", "QA Hoàng Minh Quân thực hiện kiểm thử thực tế các ca sai format ngày 20/11/2026, xác nhận form hiện đúng cảnh báo đỏ và ngăn chặn lưu dữ liệu rác. Đủ điều kiện đóng phiếu.")
            ]
        },
        {
            "num": "03",
            "ma": "YC-INX-20261120-003",
            "ngay_lap": "20/11/2026",
            "don_vi": "Ban Phát triển và Quản lý phần mềm – Phòng Công nghệ thông tin",
            "nguoi_yc": "Hoàng Minh Quân (Chuyên viên Kiểm thử phần mềm)",
            "chuc_vu": "Chuyên viên QA / Ban PT & QLPM",
            "lien_he": "quanhm.cntt@a25hotel.com | 0989.xxx.xxx",
            "uu_tien": "Trung bình (Nâng cao trải nghiệm đối soát check-in của Lễ tân)",
            "loai": "[ ] Báo lỗi phần mềm (Bug)      [x] Yêu cầu cải tiến (Change Request)      [ ] Hỗ trợ nghiệp vụ",
            "phan_he": "Quản lý Đặt phòng & Nhận phòng PMS (Màn hình: /pms/reservations & /pms/room-map)",
            "tieu_de": "Đề xuất hiển thị Nguồn kênh đặt phòng và Mã booking tham chiếu OTA trên danh sách đặt phòng và popup nhận phòng (Mục 8 HROnline)",
            "hien_trang": "Hiện tại trên Danh sách đặt phòng đã có cột 'Nguồn' (Doanh nghiệp, Walk-in...) nhưng khi Lễ tân thao tác nhận phòng nhanh (Check-in) từ danh sách hoặc từ sơ đồ phòng thì popup chưa hiển thị rõ mã tham chiếu OTA (như Booking.com, Agoda) gây khó khăn khi đối chiếu voucher khách đưa.",
            "buoc_tai_hien": "1. Truy cập /pms/reservations -> Quan sát cột 'Nguồn' và 'Mã ĐL/OTA'.\n2. Bấm thao tác 'Nhận phòng' cho một đặt phòng có nguồn OTA.\n3. Kiểm tra cửa sổ modal nhận phòng xem có hiển thị rõ Nguồn OTA và Mã tham chiếu hay không.",
            "du_lieu_test": "- Nguồn: 'Agoda', 'Booking.com', 'Doanh nghiệp (Threeland Travel)'\n- Mã tham chiếu OTA mẫu: 'BDC-982172', 'AGD-102938'",
            "ky_vong": "Popup nhận phòng và Báo cáo chi tiết nhận phòng hiển thị tường minh 02 trường: 'Nguồn đặt phòng' và 'Mã booking đại lý / OTA' để lễ tân không phải mở lại trang chi tiết đặt phòng.",
            "thuc_te_uat": "ĐANG XỬ LÝ: Trên danh sách đặt phòng đã có hiển thị cột; riêng cửa sổ nhận phòng nhanh đang được Phòng CNTT tổng hợp yêu cầu kỹ thuật đưa vào kế hoạch cập nhật tháng 11/2026.",
            "kiem_tra_cntt": "Đủ trường bắt buộc: [x] ĐẠT (Có)     |     Phân loại phản ánh: [x] ĐẠT (Có chọn 'Change Request')     |     Kết luận tiếp nhận: [x] HỢP LỆ, TIẾP NHẬN",
            "phan_cong": "Nguyễn Văn Sơn (Trưởng ban PT&QLPM) duyệt yêu cầu kỹ thuật; Hoàng Minh Quân theo dõi và kiểm thử nghiệm thu.",
            "ngay_nhan": "20/11/2026",
            "ngay_dong": "Chưa đóng (Đang theo dõi hàng ngày)",
            "trang_thai": "ĐANG XỬ LÝ (Theo dõi hàng ngày từ 20/11/2026)",
            "ket_luan": "Đã lập và nộp 01 Văn bản tổng hợp yêu cầu kỹ thuật nghiệm thu theo chỉ đạo Mục 8 HROnline. Tiếp tục theo dõi đến khi đóng phiếu.",
            "huong_giai_quyet": [
                ("1. Phương án kỹ thuật (IT/Dev):", "Lập Văn bản tổng hợp yêu cầu kỹ thuật chuẩn hóa (thực hiện theo chỉ đạo Mục 8 HROnline). Yêu cầu bổ sung 02 trường 'Nguồn đặt phòng' và 'Mã booking đại lý / OTA' hiển thị trực tiếp trên modal Nhận phòng nhanh và Báo cáo chi tiết nhận phòng."),
                ("2. Hướng giải quyết cho Lễ tân:", "Trong thời gian chờ triển khai bản cập nhật phần mềm, hướng dẫn Lễ tân đối soát mã voucher/booking của khách tại cột 'Nguồn' và 'Mã ĐL / OTA' trên màn hình Danh sách đặt phòng (/pms/reservations) trước khi bấm nhận phòng."),
                ("3. Kế hoạch kiểm thử & Đóng phiếu:", "Tiếp tục theo dõi tiến độ cập nhật giao diện của đội ngũ phát triển; thực thi kiểm thử nghiệm thu ngay khi tính năng được cập nhật lên UAT để hoàn tất đóng phiếu trong tháng 11/2026.")
            ]
        },
        {
            "num": "04",
            "ma": "YC-A25-20261118-004",
            "ngay_lap": "18/11/2026",
            "don_vi": "Phòng Kinh doanh – Tập đoàn Khách sạn A25",
            "nguoi_yc": "Lê Thị Mai (Chuyên viên Kinh doanh)",
            "chuc_vu": "Chuyên viên Kinh doanh A25",
            "lien_he": "mai.lt@a25hotel.com | 0912.xxx.xxx",
            "uu_tien": "Chưa xác định (Phiếu gửi thiếu thông tin)",
            "loai": "[ ] Báo lỗi phần mềm (Bug)      [ ] Yêu cầu cải tiến (Change Request)      [ ] Hỗ trợ nghiệp vụ  --> [CHƯA CHỌN]",
            "phan_he": "Đặt phòng PMS (Ghi chung chung)",
            "tieu_de": "Yêu cầu chỉnh sửa phần mềm cho khối kinh doanh",
            "hien_trang": "Phòng Kinh doanh gửi yêu cầu hỗ trợ qua email/chat nhưng không điền theo biểu mẫu chuẩn, không chọn loại phản ánh (Bug hay Change Request) và không nêu rõ màn hình hay thao tác lỗi cụ thể.",
            "buoc_tai_hien": "[Để trống / Chưa cung cấp bước thao tác cụ thể]",
            "du_lieu_test": "[Chưa cung cấp dữ liệu kiểm thử]",
            "ky_vong": "Hỗ trợ sửa chức năng kinh doanh.",
            "thuc_te_uat": "KHÔNG ĐỦ ĐIỀU KIỆN TIẾP NHẬN: Thiếu trường bắt buộc và không chọn loại phản ánh.",
            "kiem_tra_cntt": "Đủ trường bắt buộc: [ ] KHÔNG ĐẠT (Thiếu thông tin)  |  Phân loại: [ ] KHÔNG ĐẠT (Chưa chọn)  |  Kết luận: [x] TỪ CHỐI & TRẢ LẠI",
            "phan_cong": "Hoàng Minh Quân (Chuyên viên tiếp nhận CNTT) lập biên bản trả lại phiếu theo đúng chỉ đạo Mục 3 HROnline.",
            "ngay_nhan": "18/11/2026",
            "ngay_dong": "Chưa đóng",
            "trang_thai": "TRẢ LẠI (Không tiếp nhận từ ngày 18/11/2026)",
            "ket_luan": "Thực hiện nghiêm quy định của Ban Giám đốc tại Mục 3: 'Phòng CNTT trả lại, không tiếp nhận phiếu thiếu trường bắt buộc (từ 18/11/2026) và phiếu không chọn loại phản ánh'.",
            "huong_giai_quyet": [
                ("1. Phương án xử lý của Phòng CNTT:", "Từ chối tiếp nhận phiếu vào hàng đợi xử lý kỹ thuật; ghi nhận trạng thái 'Trả lại' vào Sổ theo dõi vòng đời phiếu (THEO_DOI_PHIEU) để báo cáo số liệu thứ Hai hằng tuần cho Ban Giám đốc."),
                ("2. Hướng dẫn giải quyết cho Phòng ban:", "Phòng CNTT gửi văn bản phản hồi kèm bản mềm Biểu mẫu chuẩn; hướng dẫn chuyên viên Kinh doanh điền đầy đủ phân loại (Bug hay Change Request), chọn chính xác màn hình bị lỗi và đính kèm ảnh chụp màn hình trước khi gửi lại."),
                ("3. Kế hoạch tiếp nhận lại:", "Khi Phòng Kinh doanh hoàn thiện lại phiếu có đầy đủ các trường bắt buộc, Phòng CNTT sẽ mở mã phiếu mới, tiếp nhận và phân công xử lý theo đúng quy trình.")
            ]
        }
    ]

    cur_row = 4

    for f in forms_full_data:
        # Title of Ticket
        ws_bm.merge_cells(f"A{cur_row}:G{cur_row}")
        ws_bm[f"A{cur_row}"] = f"PHIẾU SỐ {f['num']}: BIỂU MẪU YÊU CẦU PHẦN MỀM – MÃ PHIẾU: {f['ma']}"
        ws_bm[f"A{cur_row}"].font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        ws_bm[f"A{cur_row}"].fill = fill_header_blue
        ws_bm[f"A{cur_row}"].alignment = align_left
        cur_row += 1

        # Section 1: General Info
        ws_bm.merge_cells(f"A{cur_row}:G{cur_row}")
        ws_bm[f"A{cur_row}"] = "I. THÔNG TIN CHUNG (BẮT BUỘC ĐIỀN ĐỦ)"
        ws_bm[f"A{cur_row}"].font = font_arial_10_bold
        ws_bm[f"A{cur_row}"].fill = fill_section_blue
        cur_row += 1

        sec1_rows = [
            ("Mã phiếu:", f["ma"], "Ngày lập phiếu:", f["ngay_lap"]),
            ("Đơn vị / Chi nhánh:", f["don_vi"], "Mức độ ưu tiên:", f["uu_tien"]),
            ("Người yêu cầu:", f["nguoi_yc"], "Chức vụ / Bộ phận:", f["chuc_vu"]),
            ("Thông tin liên hệ:", f["lien_he"], "Phân hệ phần mềm:", f["phan_he"]),
            ("Phân loại yêu cầu (*):", f["loai"], "", "")
        ]

        for r_data in sec1_rows:
            ws_bm.cell(cur_row, 1, r_data[0]).font = font_arial_10_bold
            ws_bm.cell(cur_row, 1).fill = fill_sub_header
            ws_bm.cell(cur_row, 1).border = border_thin
            ws_bm.cell(cur_row, 1).alignment = align_left

            if r_data[2] == "":
                ws_bm.merge_cells(f"B{cur_row}:G{cur_row}")
                ws_bm.cell(cur_row, 2, r_data[1]).font = font_arial_10
                ws_bm.cell(cur_row, 2).alignment = align_left
                style_range(ws_bm, 2, cur_row, 7, cur_row, border=border_thin)
            else:
                ws_bm.merge_cells(f"B{cur_row}:C{cur_row}")
                ws_bm.cell(cur_row, 2, r_data[1]).font = font_arial_10
                ws_bm.cell(cur_row, 2).alignment = align_left
                style_range(ws_bm, 2, cur_row, 3, cur_row, border=border_thin)

                ws_bm.cell(cur_row, 4, r_data[2]).font = font_arial_10_bold
                ws_bm.cell(cur_row, 4).fill = fill_sub_header
                ws_bm.cell(cur_row, 4).border = border_thin
                ws_bm.cell(cur_row, 4).alignment = align_left

                ws_bm.merge_cells(f"E{cur_row}:G{cur_row}")
                ws_bm.cell(cur_row, 5, r_data[3]).font = font_arial_10
                ws_bm.cell(cur_row, 5).alignment = align_left
                style_range(ws_bm, 5, cur_row, 7, cur_row, border=border_thin)
            cur_row += 1

        # Section 2: Details
        ws_bm.merge_cells(f"A{cur_row}:G{cur_row}")
        ws_bm[f"A{cur_row}"] = "II. CHI TIẾT YÊU CẦU / MÔ TẢ LỖI PHẦN MỀM"
        ws_bm[f"A{cur_row}"].font = font_arial_10_bold
        ws_bm[f"A{cur_row}"].fill = fill_section_blue
        cur_row += 1

        sec2_rows = [
            ("Tiêu đề yêu cầu (*):", f["tieu_de"]),
            ("Mô tả hiện trạng / Lỗi (*):", f["hien_trang"]),
            ("Các bước tái hiện (Repro):", f["buoc_tai_hien"]),
            ("Dữ liệu kiểm thử (Test Data):", f["du_lieu_test"]),
            ("Kết quả mong đợi (*):", f["ky_vong"]),
            ("Kết quả thực tế UAT:", f["thuc_te_uat"])
        ]

        for lbl, val in sec2_rows:
            ws_bm.cell(cur_row, 1, lbl).font = font_arial_10_bold
            ws_bm.cell(cur_row, 1).fill = fill_sub_header
            ws_bm.cell(cur_row, 1).border = border_thin
            ws_bm.cell(cur_row, 1).alignment = align_left

            ws_bm.merge_cells(f"B{cur_row}:G{cur_row}")
            ws_bm.cell(cur_row, 2, val).font = font_arial_10
            ws_bm.cell(cur_row, 2).alignment = align_wrap_left
            style_range(ws_bm, 2, cur_row, 7, cur_row, border=border_thin)
            cur_row += 1

        # Section 3: IT Verification & Status
        ws_bm.merge_cells(f"A{cur_row}:G{cur_row}")
        ws_bm[f"A{cur_row}"] = "III. KIỂM SOÁT TIẾP NHẬN & THEO DÕI VÒNG ĐỜI (PHÒNG CNTT A25)"
        ws_bm[f"A{cur_row}"].font = font_arial_10_bold
        ws_bm[f"A{cur_row}"].fill = fill_section_blue
        cur_row += 1

        sec3_rows = [
            ("Kiểm tra tính hợp lệ (*):", f["kiem_tra_cntt"]),
            ("Phân công xử lý:", f["phan_cong"]),
            ("Tiến độ & Trạng thái:", f"Ngày nhận: {f['ngay_nhan']}   |   Ngày đóng: {f['ngay_dong']}   |   Trạng thái: {f['trang_thai']}"),
            ("Kết luận nghiệm thu:", f["ket_luan"])
        ]

        for lbl, val in sec3_rows:
            ws_bm.cell(cur_row, 1, lbl).font = font_arial_10_bold
            ws_bm.cell(cur_row, 1).fill = fill_sub_header
            ws_bm.cell(cur_row, 1).border = border_thin
            ws_bm.cell(cur_row, 1).alignment = align_left

            ws_bm.merge_cells(f"B{cur_row}:G{cur_row}")
            ws_bm.cell(cur_row, 2, val).font = font_arial_10
            ws_bm.cell(cur_row, 2).alignment = align_wrap_left
            style_range(ws_bm, 2, cur_row, 7, cur_row, border=border_thin)
            cur_row += 1

        # SECTION 4: HƯỚNG GIẢI QUYẾT & PHƯƠNG ÁN XỬ LÝ (NEWLY ADDED)
        ws_bm.merge_cells(f"A{cur_row}:G{cur_row}")
        ws_bm[f"A{cur_row}"] = "IV. HƯỚNG GIẢI QUYẾT & PHƯƠNG ÁN XỬ LÝ (RESOLUTION PLAN)"
        ws_bm[f"A{cur_row}"].font = font_arial_10_bold
        ws_bm[f"A{cur_row}"].fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
        cur_row += 1

        for lbl, val in f["huong_giai_quyet"]:
            ws_bm.cell(cur_row, 1, lbl).font = font_arial_10_bold
            ws_bm.cell(cur_row, 1).fill = fill_sub_header
            ws_bm.cell(cur_row, 1).border = border_thin
            ws_bm.cell(cur_row, 1).alignment = align_left

            ws_bm.merge_cells(f"B{cur_row}:G{cur_row}")
            ws_bm.cell(cur_row, 2, val).font = font_arial_10
            ws_bm.cell(cur_row, 2).alignment = align_wrap_left
            style_range(ws_bm, 2, cur_row, 7, cur_row, border=border_thin)
            cur_row += 1

        # Section 5: Signatures
        ws_bm.cell(cur_row, 1, "NGƯỜI LẬP PHIẾU\n(Hoàng Minh Quân)").font = font_arial_10_bold
        ws_bm.cell(cur_row, 1).alignment = align_center
        ws_bm.merge_cells(f"A{cur_row}:B{cur_row}")
        style_range(ws_bm, 1, cur_row, 2, cur_row, border=border_thin)

        ws_bm.cell(cur_row, 3, "TRƯỞNG BỘ PHẬN YÊU CẦU\n(Ký và ghi rõ họ tên)").font = font_arial_10_bold
        ws_bm.cell(cur_row, 3).alignment = align_center
        ws_bm.merge_cells(f"C{cur_row}:D{cur_row}")
        style_range(ws_bm, 3, cur_row, 4, cur_row, border=border_thin)

        ws_bm.cell(cur_row, 5, "TIẾP NHẬN & PHÊ DUYỆT CNTT\n(Trưởng ban Nguyễn Văn Sơn)").font = font_arial_10_bold
        ws_bm.cell(cur_row, 5).alignment = align_center
        ws_bm.merge_cells(f"E{cur_row}:G{cur_row}")
        style_range(ws_bm, 5, cur_row, 7, cur_row, border=border_thin)

        cur_row += 1
        # Signature space row
        ws_bm.row_dimensions[cur_row].height = 40
        ws_bm.merge_cells(f"A{cur_row}:B{cur_row}")
        ws_bm.merge_cells(f"C{cur_row}:D{cur_row}")
        ws_bm.merge_cells(f"E{cur_row}:G{cur_row}")
        style_range(ws_bm, 1, cur_row, 2, cur_row, border=border_thin)
        style_range(ws_bm, 3, cur_row, 4, cur_row, border=border_thin)
        style_range(ws_bm, 5, cur_row, 7, cur_row, border=border_thin)

        cur_row += 3 # spacing between forms

    # Set column widths for BIEU_MAU_PHIEU_YC
    ws_bm.column_dimensions["A"].width = 24
    ws_bm.column_dimensions["B"].width = 22
    ws_bm.column_dimensions["C"].width = 18
    ws_bm.column_dimensions["D"].width = 22
    ws_bm.column_dimensions["E"].width = 20
    ws_bm.column_dimensions["F"].width = 16
    ws_bm.column_dimensions["G"].width = 18

    # Save to both files
    wb.save(FILE_PATH)
    wb.save(FILE_BACKUP)
    print("SUCCESS: Updated resolution plan into both Excel files!")

if __name__ == '__main__':
    apply_resolution_updates()
