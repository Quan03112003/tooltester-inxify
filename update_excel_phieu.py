import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import datetime

FILE_PATH = r"c:\Users\Admin\Desktop\tester\Mau_Excel_QC_Inxify_DoiChieu_va_TheoDoiPhieu.xlsx"

def update_excel():
    wb = openpyxl.load_workbook(FILE_PATH)
    
    # Fonts & Styles
    font_arial_10 = Font(name="Arial", size=10)
    font_arial_10_bold = Font(name="Arial", size=10, bold=True)
    font_arial_11_bold = Font(name="Arial", size=11, bold=True)
    font_arial_12_bold = Font(name="Arial", size=12, bold=True)
    font_arial_14_bold = Font(name="Arial", size=14, bold=True, color="1E3A8A")
    font_arial_9_italic = Font(name="Arial", size=9, italic=True, color="555555")
    
    thin_side = Side(border_style="thin", color="D1D5DB")
    border_thin = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    fill_header_blue = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    fill_section_blue = PatternFill(start_color="E0E7FF", end_color="E0E7FF", fill_type="solid")
    fill_sub_header = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    fill_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_green_tag = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid") # Đã đóng
    fill_amber_tag = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid") # Đang xử lý
    fill_red_tag = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")   # Trả lại

    font_green = Font(name="Arial", size=10, bold=True, color="166534")
    font_amber = Font(name="Arial", size=10, bold=True, color="92400E")
    font_red = Font(name="Arial", size=10, bold=True, color="991B1B")

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_wrap_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # =========================================================================
    # 1. CẬP NHẬT SHEET 'THEO_DOI_PHIEU'
    # =========================================================================
    ws_td = wb["THEO_DOI_PHIEU"]
    
    # Update title & note
    ws_td["A1"].value = "THEO DÕI VÒNG ĐỜI PHIẾU YÊU CẦU PHẦN MỀM INXIFY (MỤC 3 & MỤC 4 HRONLINE)"
    ws_td["A2"].value = "Sổ theo dõi thực tế các phiếu yêu cầu phần mềm Inxify – Áp dụng quy chuẩn Mục 3 (chặn phiếu thiếu trường) & Mục 4 (theo dõi hàng ngày từ 18/11/2026)."
    ws_td["A2"].font = Font(name="Arial", size=9.5, italic=True, color="4B5563")

    # Data for the tickets: 3 main tickets for Reservations + 1 demonstration rejected ticket for Rule 3
    tickets_data = [
        # Row 4: Phiếu 1
        {
            "row": 4,
            "ma": "YC-INX-20261118-001",
            "ngay_nhan": datetime.datetime(2026, 11, 18, 0, 0),
            "phong_ban": "Bộ phận Lễ tân (A25 684 Minh Khai)",
            "nguoi_yc": "Hoàng Minh Quân (QA/Tester)",
            "phan_loai": "Bug",
            "phan_he": "Đặt phòng (PMS)",
            "mo_ta": "Cảnh báo thiếu trường bắt buộc khi tạo đặt phòng (/pms/reservations): Không cho phép lưu khi bỏ trống trường Tên khách (guestName) hoặc Giá/đêm (ratePerNight). Yêu cầu form viền đỏ và báo lỗi chi tiết.",
            "du_truong": "Có",
            "co_loai": "Có",
            "trang_thai": "Đã đóng",
            "nguoi_xl": "Châu Tuấn (Dev) / Hoàng Minh Quân (QA)",
            "ngay_dong": datetime.datetime(2026, 11, 19, 0, 0),
            "so_ngay_formula": '=IF(OR(B4="",L4=""),"",L4-B4)',
            "ghi_chu": "Đã kiểm thử UAT đạt (Pass): Form tự động viền đỏ và hiển thị cảnh báo đỏ 'Vui lòng nhập tên khách', 'Vui lòng nhập giá'. Nghiệm thu đạt, đóng phiếu."
        },
        # Row 5: Phiếu 2
        {
            "row": 5,
            "ma": "YC-INX-20261119-002",
            "ngay_nhan": datetime.datetime(2026, 11, 19, 0, 0),
            "phong_ban": "Bộ phận Lễ tân (A25 684 Minh Khai)",
            "nguoi_yc": "Hoàng Minh Quân (QA/Tester)",
            "phan_loai": "Bug",
            "phan_he": "Đặt phòng (PMS)",
            "mo_ta": "Cảnh báo sai định dạng dữ liệu Số điện thoại và Email trên form Đặt phòng: Nhập SĐT không đủ 10 số (VD: 09812) hoặc email sai cú pháp (VD: email_sai@@) phải chặn lưu tại giao diện.",
            "du_truong": "Có",
            "co_loai": "Có",
            "trang_thai": "Đã đóng",
            "nguoi_xl": "Châu Tuấn (Dev) / Hoàng Minh Quân (QA)",
            "ngay_dong": datetime.datetime(2026, 11, 20, 0, 0),
            "so_ngay_formula": '=IF(OR(B5="",L5=""),"",L5-B5)',
            "ghi_chu": "Đã kiểm thử UAT đạt (Pass): Giao diện hiện cảnh báo đỏ 'SĐT 10 số, bắt đầu bằng 0' và 'Email không đúng định dạng'. Nghiệm thu đạt, đóng phiếu."
        },
        # Row 6: Phiếu 3
        {
            "row": 6,
            "ma": "YC-INX-20261120-003",
            "ngay_nhan": datetime.datetime(2026, 11, 20, 0, 0),
            "phong_ban": "Ban Phát triển & Quản lý PM",
            "nguoi_yc": "Hoàng Minh Quân (QA/Tester)",
            "phan_loai": "Change Request",
            "phan_he": "Đặt phòng (PMS)",
            "mo_ta": "Hiển thị trường Nguồn kênh đặt phòng (Agoda, Booking, Doanh nghiệp) và Mã booking OTA trên Danh sách đặt phòng và Cửa sổ nhận phòng nhanh (theo Mục 8 HROnline).",
            "du_truong": "Có",
            "co_loai": "Có",
            "trang_thai": "Đang xử lý",
            "nguoi_xl": "Nguyễn Văn Sơn (Trưởng ban PT&QLPM)",
            "ngay_dong": None,
            "so_ngay_formula": '=IF(OR(B6="",L6=""),"",L6-B6)',
            "ghi_chu": "Đã nộp văn bản tổng hợp yêu cầu kỹ thuật ngày 20/11/2026 theo chỉ đạo BGĐ. Đang phối hợp Dev để hiển thị trên màn hình nhận phòng."
        },
        # Row 7: Phiếu 4 (Minh họa áp dụng quy tắc trả lại phiếu theo Mục 3)
        {
            "row": 7,
            "ma": "YC-A25-20261118-004",
            "ngay_nhan": datetime.datetime(2026, 11, 18, 0, 0),
            "phong_ban": "Phòng Kinh doanh A25",
            "nguoi_yc": "Lê Thị Mai (Kinh doanh)",
            "phan_loai": None,
            "phan_he": "Đặt phòng (PMS)",
            "mo_ta": "Yêu cầu chỉnh sửa phần mềm cho khối kinh doanh nhưng không chọn loại phản ánh, không nêu rõ màn hình hay thao tác lỗi cụ thể.",
            "du_truong": "Không",
            "co_loai": "Không",
            "trang_thai": "Trả lại",
            "nguoi_xl": "Hoàng Minh Quân (Tiếp nhận CNTT)",
            "ngay_dong": None,
            "so_ngay_formula": '=IF(OR(B7="",L7=""),"",L7-B7)',
            "ghi_chu": "Áp dụng Mục 3 HROnline (áp dụng từ 18/11/2026): Phòng CNTT từ chối tiếp nhận và trả lại phiếu do thiếu trường bắt buộc và không chọn loại phản ánh."
        }
    ]

    for item in tickets_data:
        r = item["row"]
        ws_td.cell(r, 1, item["ma"])
        ws_td.cell(r, 2, item["ngay_nhan"])
        ws_td.cell(r, 3, item["phong_ban"])
        ws_td.cell(r, 4, item["nguoi_yc"])
        ws_td.cell(r, 5, item["phan_loai"])
        ws_td.cell(r, 6, item["phan_he"])
        ws_td.cell(r, 7, item["mo_ta"])
        ws_td.cell(r, 8, item["du_truong"])
        ws_td.cell(r, 9, item["co_loai"])
        ws_td.cell(r, 10, item["trang_thai"])
        ws_td.cell(r, 11, item["nguoi_xl"])
        ws_td.cell(r, 12, item["ngay_dong"])
        ws_td.cell(r, 13, item["so_ngay_formula"])
        ws_td.cell(r, 14, item["ghi_chu"])

        # Formatting
        ws_td.cell(r, 1).alignment = align_center
        ws_td.cell(r, 2).alignment = align_center
        ws_td.cell(r, 2).number_format = "dd/mm/yyyy"
        ws_td.cell(r, 3).alignment = align_left
        ws_td.cell(r, 4).alignment = align_left
        ws_td.cell(r, 5).alignment = align_center
        ws_td.cell(r, 6).alignment = align_center
        ws_td.cell(r, 7).alignment = align_wrap_left
        ws_td.cell(r, 8).alignment = align_center
        ws_td.cell(r, 9).alignment = align_center
        ws_td.cell(r, 10).alignment = align_center
        ws_td.cell(r, 11).alignment = align_left
        ws_td.cell(r, 12).alignment = align_center
        ws_td.cell(r, 12).number_format = "dd/mm/yyyy"
        ws_td.cell(r, 13).alignment = align_center
        ws_td.cell(r, 13).number_format = "0"
        ws_td.cell(r, 14).alignment = align_wrap_left

        for c in range(1, 15):
            cell = ws_td.cell(r, c)
            cell.font = font_arial_10
            cell.border = border_thin
            if c == 10: # Status tag
                if item["trang_thai"] == "Đã đóng":
                    cell.fill = fill_green_tag
                    cell.font = font_green
                elif item["trang_thai"] == "Đang xử lý":
                    cell.fill = fill_amber_tag
                    cell.font = font_amber
                elif item["trang_thai"] == "Trả lại":
                    cell.fill = fill_red_tag
                    cell.font = font_red
            else:
                cell.fill = fill_zebra if r % 2 == 1 else fill_white

    # Clear rows 8 to 15 if any stale data
    for r in range(8, 20):
        for c in range(1, 15):
            cell = ws_td.cell(r, c)
            if cell.value is not None:
                cell.value = None
                cell.fill = fill_white
                cell.border = Border()

    # Ensure column widths in THEO_DOI_PHIEU
    col_widths_td = {
        1: 22,  # Mã phiếu
        2: 13,  # Ngày nhận
        3: 25,  # Phòng ban
        4: 25,  # Người yêu cầu
        5: 16,  # Phân loại
        6: 18,  # Phân hệ
        7: 45,  # Mô tả
        8: 18,  # Đủ trường?
        9: 16,  # Có chọn loại?
        10: 16, # Trạng thái
        11: 30, # Người xử lý
        12: 13, # Ngày đóng
        13: 15, # Số ngày
        14: 45  # Ghi chú
    }
    for col_idx, width in col_widths_td.items():
        ws_td.column_dimensions[get_column_letter(col_idx)].width = width

    # =========================================================================
    # 2. CẬP NHẬT SHEET 'HUONG_DAN'
    # =========================================================================
    ws_hd = wb["HUONG_DAN"]
    ws_hd["B8"].value = "Theo dõi từng phiếu từ lúc nhận đến khi đóng theo chỉ đạo Mục 3 & Mục 4 HROnline. Ghi nhận 03 phiếu yêu cầu thực tế cho tính năng Đặt phòng Inxify PMS từ 18/11/2026 kèm 01 phiếu minh họa quy tắc trả lại."
    ws_hd["B4"].value = "Gộp các đầu việc thử việc A25: (1) Bảng đối chiếu dữ liệu 02 màn hình Inxify (Mục 2); (2) Sổ theo dõi phiếu yêu cầu phần mềm và Báo cáo tuần (Mục 3); (3) Lập 03 phiếu yêu cầu thực tế theo biểu mẫu (Mục 4)."

    # =========================================================================
    # 3. TẠO MỚI SHEET 'BIEU_MAU_PHIEU_YC' (Chi tiết 3 phiếu theo biểu mẫu chuẩn)
    # =========================================================================
    sheet_name_bm = "BIEU_MAU_PHIEU_YC"
    if sheet_name_bm in wb.sheetnames:
        del wb[sheet_name_bm]
    ws_bm = wb.create_sheet(title=sheet_name_bm)

    # Style helper for form
    def style_range(ws, start_col, start_row, end_col, end_row, font=None, fill=None, border=None, alignment=None):
        for r in range(start_row, end_row + 1):
            for c in range(start_col, end_col + 1):
                cell = ws.cell(r, c)
                if font: cell.font = font
                if fill: cell.fill = fill
                if border: cell.border = border
                if alignment: cell.alignment = alignment

    cur_row = 1
    # Sheet Header
    ws_bm.merge_cells("A1:G1")
    ws_bm["A1"] = "BỘ BIỂU MẪU YÊU CẦU PHẦN MỀM CHUẨN HÓA (MỤC 4 HRONLINE)"
    ws_bm["A1"].font = font_arial_14_bold
    ws_bm["A1"].alignment = align_center

    ws_bm.merge_cells("A2:G2")
    ws_bm["A2"] = "Thực hiện nhiệm vụ Mục 4: Ứng viên Hoàng Minh Quân lập 03 phiếu yêu cầu thực tế trên phần mềm Inxify PMS (Tính năng Đặt phòng) theo Biểu mẫu yêu cầu phần mềm ban hành."
    ws_bm["A2"].font = font_arial_9_italic
    ws_bm["A2"].alignment = align_center

    cur_row = 4

    forms_data = [
        {
            "num": "01",
            "ma": "YC-INX-20261118-001",
            "ngay_lap": "18/11/2026",
            "don_vi": "Bộ phận Lễ tân – Khách sạn A25 Signature (684 Minh Khai)",
            "nguoi_yc": "Hoàng Minh Quân (Chuyên viên Kiểm thử phần mềm)",
            "chuc_vu": "Chuyên viên QA / Ban PT & QLPM",
            "lien_he": "quanhm.cntt@a25hotel.com | 0989.xxx.xxx",
            "uu_tien": "Cao (Ảnh hưởng dữ liệu đặt phòng & vận hành lễ tân)",
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
            "ket_luan": "Đã nghiệm thu đạt trên UAT Inxify. Đính kèm ảnh minh chứng: M6_01_DatPhong_CanhBao_Thieu_Truong_Bat_Buoc.png."
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
            "ket_luan": "Đã nghiệm thu đạt trên giao diện UAT Inxify. Đính kèm ảnh minh chứng: M6_02_DatPhong_CanhBao_Sai_Dinh_Dang_SDT_Email.png."
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
            "ket_luan": "Đã lập và nộp 01 Văn bản tổng hợp yêu cầu kỹ thuật nghiệm thu theo chỉ đạo Mục 8 HROnline. Tiếp tục theo dõi đến khi đóng phiếu."
        }
    ]

    for f in forms_data:
        # Title of Ticket
        ws_bm.merge_cells(f"A{cur_row}:G{cur_row}")
        ws_bm[f"A{cur_row}"] = f"PHIẾU SỐ {f['num']}: BIỂU MẪU YÊU CẦU PHẦN MỀM – MÃ PHIẾU: {f['ma']}"
        ws_bm[f"A{cur_row}"].font = font_arial_11_bold
        ws_bm[f"A{cur_row}"].fill = fill_header_blue
        ws_bm[f"A{cur_row}"].alignment = align_left
        ws_bm[f"A{cur_row}"].font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
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

            if r_data[2] == "": # Span full width
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

        # Signatures
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

    # Save
    wb.save(FILE_PATH)
    print("SUCCESS: Updated Excel workbook at", FILE_PATH)

if __name__ == '__main__':
    update_excel()
