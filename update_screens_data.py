import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

FILE_PATH = r"c:\Users\Admin\Desktop\tester\Mau_Excel_QC_Inxify_DoiChieu_va_TheoDoiPhieu.xlsx"

def update_screens():
    wb = openpyxl.load_workbook(FILE_PATH)

    # Styles
    font_header = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    font_bold = Font(name="Arial", size=10, bold=True)
    font_regular = Font(name="Arial", size=10)
    font_italic = Font(name="Arial", size=9, italic=True, color="4B5563")
    font_title = Font(name="Arial", size=13, bold=True, color="1E3A8A")
    font_pass = Font(name="Arial", size=10, bold=True, color="166534")
    font_fail = Font(name="Arial", size=10, bold=True, color="991B1B")

    fill_hdr = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    fill_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_summary = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

    thin_border = Border(
        left=Side(style="thin", color="D1D5DB"),
        right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="D1D5DB"),
        bottom=Side(style="thin", color="D1D5DB")
    )

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_wrap_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # =========================================================================
    # 1. CẬP NHẬT SHEET 'DOI_CHIEU_MH1': MÀN HÌNH ĐẶT PHÒNG INXIFY PMS (/pms/reservations)
    # =========================================================================
    ws1 = wb["DOI_CHIEU_MH1"]
    ws1["A1"] = "ĐỐI CHIẾU DỮ LIỆU – MÀN HÌNH ĐẶT PHÒNG INXIFY PMS (/pms/reservations)"
    ws1["A1"].font = font_title
    ws1["A2"] = "Bảng đối chiếu kiểm thử dữ liệu thực tế tại khách sạn A25 Signature (684 Minh Khai) – Phục vụ nhiệm vụ Mục 2 HROnline tháng 11/2026."
    ws1["A2"].font = font_italic

    # Clear old rows 4 to 25
    for r in range(4, 30):
        for c in range(1, 8):
            cell = ws1.cell(r, c)
            cell.value = None
            cell.fill = fill_white
            cell.border = Border()

    mh1_data = [
        (1, "Họ và tên khách (guestName)", "Nguyễn Văn Kiểm Thử", "Nguyễn Văn Kiểm Thử", "Nguyễn Văn Kiểm Thử",
         "Dẫn chứng ảnh M6_01, M6_02. Bắt buộc nhập, bỏ trống viền đỏ và báo 'Vui lòng nhập tên khách'."),
        (2, "Số điện thoại (guestPhone)", "0901234567", "0901234567", "0901234567",
         "Dẫn chứng ảnh M6_02. Format chuẩn 10 số. Nhập sai format (VD: 09812) cảnh báo 'SĐT 10 số, bắt đầu bằng 0'."),
        (3, "Email khách (guestEmail)", "guest.test@a25hotel.com", "guest.test@a25hotel.com", "guest.test@a25hotel.com",
         "Dẫn chứng ảnh M6_02. Kiểm tra Regex email chuẩn. Nhập sai (VD: abc@@) cảnh báo 'Email không đúng định dạng'."),
        (4, "Số giấy tờ / CCCD (guestIdNumber)", "044081001037", "044081001037", "044081001037",
         "Dẫn chứng ảnh M6_02, M6_05. Định dạng 12 chữ số CCCD chuẩn Bộ Công an."),
        (5, "Quốc tịch (guestNationality)", "vn Việt Nam", "vn Việt Nam", "vn Việt Nam",
         "Dẫn chứng ảnh M6_01, M6_05. Dropdown danh mục quốc tịch."),
        (6, "Hạng phòng (roomTypeId)", "Deluxe", "Deluxe", "Deluxe",
         "Dẫn chứng ảnh M6_03. Danh mục hạng phòng tại A25 Signature 684 Minh Khai."),
        (7, "Phòng cụ thể (roomId)", "P.1407", "P.1407", "P.1407",
         "Dẫn chứng ảnh M6_03. Gán phòng cụ thể, có kiểm tra xung đột trùng lịch phòng."),
        (8, "Ngày nhận phòng (checkInDate)", "20/11/2026", "20/11/2026", "20/11/2026",
         "Dẫn chứng ảnh M6_03. Chặn không cho chọn ngày nhận trong quá khứ."),
        (9, "Ngày trả phòng (checkOutDate)", "22/11/2026", "22/11/2026", "22/11/2026",
         "Dẫn chứng ảnh M6_03. Bắt buộc sau hoặc bằng ngày nhận phòng."),
        (10, "Giá/đêm (ratePerNight)", "1.500.000", "1.500.000", "1.500.000",
         "Dẫn chứng ảnh M6_01. Bắt buộc nhập, bỏ trống viền đỏ báo 'Vui lòng nhập giá'."),
        (11, "Nguồn đặt phòng (source)", "Doanh nghiệp", "Doanh nghiệp", "Doanh nghiệp",
         "Dẫn chứng ảnh M6_03 & Mục 8 HROnline. Phân loại nguồn đặt phòng."),
        (12, "Mã booking đại lý / OTA (otaBookingId)", "BDC-982172", "BDC-982172", "BDC-982172",
         "Dẫn chứng Mục 8 HROnline. Mã tham chiếu để lễ tân đối soát booking OTA khi nhận phòng.")
    ]

    for idx, (stt, ten, nhap, ky_vong, thuc_te, ghi_chu) in enumerate(mh1_data, start=4):
        ws1.cell(idx, 1, stt).alignment = align_center
        ws1.cell(idx, 2, ten).alignment = align_left
        ws1.cell(idx, 3, nhap).alignment = align_left
        ws1.cell(idx, 4, ky_vong).alignment = align_left
        ws1.cell(idx, 5, thuc_te).alignment = align_left
        ws1.cell(idx, 6, f'=IF(OR(D{idx}="",E{idx}=""),"",IF(D{idx}=E{idx},"PASS","FAIL"))').alignment = align_center
        ws1.cell(idx, 7, ghi_chu).alignment = align_wrap_left

        for c in range(1, 8):
            cell = ws1.cell(idx, c)
            cell.font = font_regular
            cell.border = thin_border
            cell.fill = fill_zebra if idx % 2 == 1 else fill_white

    # Summary row at row 17
    sum_row_1 = 17
    ws1.cell(sum_row_1, 1, "Tổng số trường").font = font_bold
    ws1.cell(sum_row_1, 1).fill = fill_summary
    ws1.cell(sum_row_1, 1).border = thin_border
    ws1.cell(sum_row_1, 1).alignment = align_center

    ws1.cell(sum_row_1, 2, f'=COUNTA(B4:B{sum_row_1-1})').font = font_bold
    ws1.cell(sum_row_1, 2).fill = fill_summary
    ws1.cell(sum_row_1, 2).border = thin_border
    ws1.cell(sum_row_1, 2).alignment = align_center

    ws1.cell(sum_row_1, 3, "PASS").font = font_pass
    ws1.cell(sum_row_1, 3).fill = fill_summary
    ws1.cell(sum_row_1, 3).border = thin_border
    ws1.cell(sum_row_1, 3).alignment = align_center

    ws1.cell(sum_row_1, 4, f'=COUNTIF(F4:F{sum_row_1-1},"PASS")').font = font_pass
    ws1.cell(sum_row_1, 4).fill = fill_summary
    ws1.cell(sum_row_1, 4).border = thin_border
    ws1.cell(sum_row_1, 4).alignment = align_center

    ws1.cell(sum_row_1, 5, "FAIL").font = font_fail
    ws1.cell(sum_row_1, 5).fill = fill_summary
    ws1.cell(sum_row_1, 5).border = thin_border
    ws1.cell(sum_row_1, 5).alignment = align_center

    ws1.cell(sum_row_1, 6, f'=COUNTIF(F4:F{sum_row_1-1},"FAIL")').font = font_fail
    ws1.cell(sum_row_1, 6).fill = fill_summary
    ws1.cell(sum_row_1, 6).border = thin_border
    ws1.cell(sum_row_1, 6).alignment = align_center

    ws1.cell(sum_row_1, 7, "Tất cả 12/12 trường dữ liệu đặt phòng đều khớp kỳ vọng (100% PASS).").font = font_italic
    ws1.cell(sum_row_1, 7).fill = fill_summary
    ws1.cell(sum_row_1, 7).border = thin_border

    # Adjust widths for MH1
    ws1.column_dimensions["A"].width = 6
    ws1.column_dimensions["B"].width = 30
    ws1.column_dimensions["C"].width = 25
    ws1.column_dimensions["D"].width = 25
    ws1.column_dimensions["E"].width = 25
    ws1.column_dimensions["F"].width = 12
    ws1.column_dimensions["G"].width = 50

    # =========================================================================
    # 2. CẬP NHẬT SHEET 'DOI_CHIEU_MH2': MÀN HÌNH QUẢN LÝ KHÁCH HÀNG CRM (/crm/guests)
    # =========================================================================
    ws2 = wb["DOI_CHIEU_MH2"]
    ws2["A1"] = "ĐỐI CHIẾU DỮ LIỆU – MÀN HÌNH KHÁCH HÀNG INXIFY CRM (/crm/guests)"
    ws2["A1"].font = font_title
    ws2["A2"] = "Bảng đối chiếu kiểm thử dữ liệu thực tế kho 14.032 hồ sơ khách hàng – Phục vụ nhiệm vụ Mục 2 HROnline tháng 11/2026."
    ws2["A2"].font = font_italic

    # Clear old rows 4 to 25
    for r in range(4, 30):
        for c in range(1, 8):
            cell = ws2.cell(r, c)
            cell.value = None
            cell.fill = fill_white
            cell.border = Border()

    mh2_data = [
        (1, "Họ tên khách hàng", "Phan Anh Tuấn", "Phan Anh Tuấn", "Phan Anh Tuấn",
         "Dẫn chứng ảnh M6_04, M6_05. Trường định danh bắt buộc."),
        (2, "Số điện thoại liên hệ", "—", "—", "—",
         "Dẫn chứng ảnh M6_05. Hồ sơ cũ chưa có SĐT, hệ thống hiển thị ký tự '—'."),
        (3, "Địa chỉ Email", "—", "—", "—",
         "Dẫn chứng ảnh M6_05. Hồ sơ cũ chưa có Email, hệ thống hiển thị ký tự '—'."),
        (4, "Số giấy tờ / CCCD", "044081001037", "044081001037", "044081001037",
         "Dẫn chứng ảnh M6_05, M6_06. Định dạng 12 chữ số CCCD. Tra cứu ra đúng 01 bản ghi."),
        (5, "Loại giấy tờ", "Căn cước công dân", "Căn cước công dân", "Căn cước công dân",
         "Dẫn chứng ảnh M6_05. Phân loại giấy tờ tùy thân."),
        (6, "Ngày sinh", "06/01/1981", "06/01/1981", "06/01/1981",
         "Dẫn chứng ảnh M6_05. Định dạng ngày sinh dd/mm/yyyy."),
        (7, "Giới tính", "MALE", "MALE", "MALE",
         "Dẫn chứng ảnh M6_05. Giới tính khách lưu trú."),
        (8, "Quốc tịch", "vn Việt Nam", "vn Việt Nam", "vn Việt Nam",
         "Dẫn chứng ảnh M6_05. Quốc tịch khách lưu trú."),
        (9, "Địa chỉ thường trú", "Xã Lộc Ninh, Đồng Hới, Quảng Bình", "Tỉnh Quảng Bình Thành Phố Đồng Hới Xã Lộc Ninh thôn 1 lộc đại", "Tỉnh Quảng Bình Thành Phố Đồng Hới Xã Lộc Ninh thôn 1 lộc đại",
         "Dẫn chứng ảnh M6_05. Địa chỉ chi tiết phục vụ khai báo lưu trú."),
        (10, "Hạng khách hàng", "Thường", "Thường", "Thường",
         "Dẫn chứng ảnh M6_04, M6_05. Hạng thẻ phân loại khách hàng thân thiết."),
        (11, "Kiểm soát trùng lặp hồ sơ", "Chặn trùng TEST QUAN", "Chặn trùng TEST QUAN", "Tồn tại 2 bản ghi trùng TEST QUAN (SĐT rác 12345, Email notanemail)",
         "Dẫn chứng ảnh M6_07. Cảnh báo lỗi: Hệ thống chưa có cơ chế chặn trùng real-time lúc lưu.")
    ]

    for idx, (stt, ten, nhap, ky_vong, thuc_te, ghi_chu) in enumerate(mh2_data, start=4):
        ws2.cell(idx, 1, stt).alignment = align_center
        ws2.cell(idx, 2, ten).alignment = align_left
        ws2.cell(idx, 3, nhap).alignment = align_left
        ws2.cell(idx, 4, ky_vong).alignment = align_left
        ws2.cell(idx, 5, thuc_te).alignment = align_left
        ws2.cell(idx, 6, f'=IF(OR(D{idx}="",E{idx}=""),"",IF(D{idx}=E{idx},"PASS","FAIL"))').alignment = align_center
        ws2.cell(idx, 7, ghi_chu).alignment = align_wrap_left

        for c in range(1, 8):
            cell = ws2.cell(idx, c)
            cell.font = font_regular
            cell.border = thin_border
            cell.fill = fill_zebra if idx % 2 == 1 else fill_white

    # Summary row at row 16
    sum_row_2 = 16
    ws2.cell(sum_row_2, 1, "Tổng số trường").font = font_bold
    ws2.cell(sum_row_2, 1).fill = fill_summary
    ws2.cell(sum_row_2, 1).border = thin_border
    ws2.cell(sum_row_2, 1).alignment = align_center

    ws2.cell(sum_row_2, 2, f'=COUNTA(B4:B{sum_row_2-1})').font = font_bold
    ws2.cell(sum_row_2, 2).fill = fill_summary
    ws2.cell(sum_row_2, 2).border = thin_border
    ws2.cell(sum_row_2, 2).alignment = align_center

    ws2.cell(sum_row_2, 3, "PASS").font = font_pass
    ws2.cell(sum_row_2, 3).fill = fill_summary
    ws2.cell(sum_row_2, 3).border = thin_border
    ws2.cell(sum_row_2, 3).alignment = align_center

    ws2.cell(sum_row_2, 4, f'=COUNTIF(F4:F{sum_row_2-1},"PASS")').font = font_pass
    ws2.cell(sum_row_2, 4).fill = fill_summary
    ws2.cell(sum_row_2, 4).border = thin_border
    ws2.cell(sum_row_2, 4).alignment = align_center

    ws2.cell(sum_row_2, 5, "FAIL").font = font_fail
    ws2.cell(sum_row_2, 5).fill = fill_summary
    ws2.cell(sum_row_2, 5).border = thin_border
    ws2.cell(sum_row_2, 5).alignment = align_center

    ws2.cell(sum_row_2, 6, f'=COUNTIF(F4:F{sum_row_2-1},"FAIL")').font = font_fail
    ws2.cell(sum_row_2, 6).fill = fill_summary
    ws2.cell(sum_row_2, 6).border = thin_border
    ws2.cell(sum_row_2, 6).alignment = align_center

    ws2.cell(sum_row_2, 7, "10/11 trường PASS; 01 trường kiểm soát trùng lặp FAIL (phát hiện lỗi thực tế để đưa vào nghiệm thu).").font = font_italic
    ws2.cell(sum_row_2, 7).fill = fill_summary
    ws2.cell(sum_row_2, 7).border = thin_border

    # Adjust widths for MH2
    ws2.column_dimensions["A"].width = 6
    ws2.column_dimensions["B"].width = 30
    ws2.column_dimensions["C"].width = 25
    ws2.column_dimensions["D"].width = 25
    ws2.column_dimensions["E"].width = 25
    ws2.column_dimensions["F"].width = 12
    ws2.column_dimensions["G"].width = 50

    wb.save(FILE_PATH)
    print("SUCCESS: Updated DOI_CHIEU_MH1 and DOI_CHIEU_MH2 with real Inxify data!")

if __name__ == '__main__':
    update_screens()
