# -*- coding: utf-8 -*-
"""
Script tạo Báo cáo mẫu chuẩn PCNTT (Số: 2111/26/BC-PCNTT)
Bổ sung đầy đủ:
- Bộ 04 con số định lượng
- 06 Tồn đọng & Phương án giải quyết
- Mục V: Đánh giá chất lượng hệ thống Inxify PMS theo các tiêu chí (Dimensions) & Kết quả đo Hiệu năng (Performance Testing - k6)
- Mục VI: Danh mục hồ sơ, tài liệu & minh chứng đính kèm đồng bộ OneDrive (Chuẩn hóa layout bảng & đường dẫn chính xác)
- Bảng ký duyệt 4 cấp hành chính
"""

import os
import shutil
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

BASE_DIR = r"c:\Users\Admin\Desktop\tester"
TARGET_FILE = os.path.join(BASE_DIR, "PCNTT_MauBC.docx")
NAMED_FILE = os.path.join(BASE_DIR, "Bao_cao_PCNTT_Tong_hop_Tien_do_Xu_ly_Ton_dong_va_Phuong_an_Giai_quyet.docx")
ONEDRIVE_FILE = os.path.join(BASE_DIR, "PCNTT_Báo cáo kiểm thử pms inxify.docx")
BACKUP_DIR_FILE = os.path.join(BASE_DIR, "dan_chung_bao_cao", "08_Bao_Cao_Tong_Hop_PCNTT_Va_Ho_So_Thu_Viec", "PCNTT_MauBC_Da_dien_chuan.docx")
BACKUP_ROOT_FILE = os.path.join(BASE_DIR, "dan_chung_bao_cao", "PCNTT_MauBC.docx")
BACKUP_OD_FILE = os.path.join(BASE_DIR, "dan_chung_bao_cao", "PCNTT_Báo cáo kiểm thử pms inxify.docx")

def set_cell_background(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=35, bottom=35, left=35, right=35):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_cell_border(cell, **kwargs):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            val, sz, space, color = edge_data
            tag = f'<w:{edge} {nsdecls("w")} w:val="{val}" w:sz="{sz}" w:space="{space}" w:color="{color}"/>'
            tcBorders.append(parse_xml(tag))
        else:
            tag = f'<w:{edge} {nsdecls("w")} w:val="none"/>'
            tcBorders.append(parse_xml(tag))
    tcPr.append(tcBorders)

def add_heading_styled(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.bold = True
    run.font.name = "Times New Roman"
    if level == 1:
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(0, 0, 0)
    elif level == 2:
        p.paragraph_format.space_before = Pt(9)
        p.paragraph_format.space_after = Pt(3)
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(20, 50, 120)
    elif level == 3:
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(50, 50, 50)
    return p

def create_full_report():
    doc = docx.Document()

    # Page Margins (TCVN standard administrative)
    for s in doc.sections:
        s.top_margin = Inches(0.6)
        s.bottom_margin = Inches(0.6)
        s.left_margin = Inches(0.7)
        s.right_margin = Inches(0.6)

    # 1. TABLE 0: Header (Phòng CNTT & Ngày tháng)
    t_hdr = doc.add_table(rows=1, cols=2)
    t_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
    c0 = t_hdr.rows[0].cells[0]
    c1 = t_hdr.rows[0].cells[1]
    c0.width = Inches(3.6)
    c1.width = Inches(3.5)

    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r0_1 = p0.add_run("PHÒNG CÔNG NGHỆ THÔNG TIN\n")
    r0_1.bold = True
    r0_1.font.name = "Times New Roman"
    r0_1.font.size = Pt(12)
    r0_2 = p0.add_run("Số: 2111/26/BC-PCNTT\n")
    r0_2.font.name = "Times New Roman"
    r0_2.font.size = Pt(11)
    r0_2.bold = True
    r0_3 = p0.add_run("Mã công việc: INX-QC-202611")
    r0_3.font.name = "Times New Roman"
    r0_3.font.size = Pt(10)
    r0_3.italic = True

    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r1_1 = p1.add_run("Hà Nội, ngày 21 tháng 11 năm 2026")
    r1_1.font.name = "Times New Roman"
    r1_1.font.size = Pt(12)
    r1_1.italic = True

    for c in (c0, c1):
        set_cell_border(c)
        set_cell_margins(c, 25, 25, 25, 25)

    # Spacing
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(4)
    p_sp.paragraph_format.space_after = Pt(4)

    # 2. Main Title
    p_t = doc.add_paragraph()
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t.paragraph_format.space_before = Pt(2)
    p_t.paragraph_format.space_after = Pt(4)
    r_t = p_t.add_run("BÁO CÁO")
    r_t.bold = True
    r_t.font.name = "Times New Roman"
    r_t.font.size = Pt(16)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(8)
    r_s = p_sub.add_run("V/v: Tổng hợp tiến độ thực hiện, các vấn đề đã xử lý, tồn đọng và phương án giải quyết trong công tác kiểm thử phần mềm Inxify PMS (09 đầu mục thử việc)")
    r_s.bold = True
    r_s.font.name = "Times New Roman"
    r_s.font.size = Pt(12)
    r_s.italic = True

    # I. CĂN CỨ
    add_heading_styled(doc, "I. CĂN CỨ", 1)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    p.add_run(
        "• Căn cứ Kết luận chỉ đạo của Ban Giám đốc và Lãnh đạo Phòng Công nghệ thông tin tại văn bản hronline.md (tháng 11/2026) về 09 đầu mục nhiệm vụ trọng tâm kiểm thử phần mềm Inxify PMS;\n"
        "• Căn cứ Kết luận đánh giá Báo cáo sơ bộ sau 07 ngày thử việc của ứng viên Hoàng Minh Quân – vị trí Chuyên viên kiểm thử phần mềm (Ban Phát triển và Quản lý phần mềm – Phòng CNTT);\n"
        "• Căn cứ Kế hoạch thử việc và phân công nhiệm vụ kiểm thử hệ thống phần mềm Quản lý khách sạn Inxify PMS từ ngày 03/11/2026 đến ngày 02/12/2026;\n"
        "• Căn cứ Hiện trạng điều kiện kiểm thử thực tế tại đơn vị: Chuyên viên kiểm thử hiện tại CHỈ ĐƯỢC CUNG CẤP DUY NHẤT MÔI TRƯỜNG UAT PORTAL (https://uat.inxify.com), hoàn toàn chưa có tài liệu đặc tả yêu cầu nghiệp vụ chính thức (PRD/SRS), chưa có từ điển dữ liệu (Data Dictionary), chưa có bộ dữ liệu kiểm thử chuẩn (Standard Test Data) và chưa có thiết bị phần cứng kiểm thử chuyên dụng (đầu đọc mã vạch 2D CCCD, thẻ CCCD cứng thật)."
    )

    # II. MỤC ĐÍCH
    add_heading_styled(doc, "II. MỤC ĐÍCH", 1)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    p.add_run(
        "• Báo cáo tổng thể toàn diện tiến độ triển khai toàn bộ 09 đầu mục công việc kiểm thử phần mềm Inxify PMS được giao tính đến ngày 21/11/2026;\n"
        "• Báo cáo chi tiết các vấn đề ĐÃ XỬ LÝ gắn liền với kết quả định lượng cụ thể (Bộ 04 con số định lượng Sếp yêu cầu), sản phẩm hồ sơ bàn giao, số lượng kịch bản kiểm thử, lỗi phát hiện và bằng chứng thực tế;\n"
        "• Báo cáo trung thực, minh bạch các vấn đề CÒN TỒN ĐỌNG và nguyên nhân khách quan xuất phát từ điều kiện thực tế (chỉ có UAT Portal, thiếu tài liệu và dữ liệu chuẩn, tính năng CCCD chưa có phần cứng và màn hình mới có form nhập tay);\n"
        "• Đề xuất PHƯƠNG ÁN GIẢI QUYẾT khả thi, chuyên nghiệp và kế hoạch hành động cụ thể để khắc phục toàn bộ các tồn đọng, bảo đảm hoàn thành hồ sơ kết thúc thử việc chất lượng cao trình Ban Giám đốc trước 16h00 ngày 02/12/2026."
    )

    # III. NỘI DUNG BÁO CÁO
    add_heading_styled(doc, "III. NỘI DUNG BÁO CÁO", 1)

    # 1. Danh mục các đầu mục
    add_heading_styled(doc, "1. Tổng hợp hiện trạng 09 đầu mục công việc kiểm thử Inxify PMS", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    p.add_run(
        "Trong giai đoạn thử việc vừa qua, chuyên viên kiểm thử Hoàng Minh Quân được giao phụ trách thực hiện trọn vẹn 09 đầu mục công việc kiểm thử Inxify PMS theo chỉ đạo tại hronline.md, phân chia thành 02 nhóm nhiệm vụ nghiệp vụ cụ thể:\n"
        "• Nhóm I (Kiểm thử chức năng hệ thống Inxify PMS - 07 đầu mục): Gồm các Mục 1, 2, 5, 6, 7, 8, 9 (Môi trường UAT Inxify, Đối chiếu 02 màn hình, Quét QR CCCD, Cảnh báo lưu dữ liệu, Nhật ký Audit Log, Nguồn kênh OTA, Cảnh báo nhận phòng chưa sẵn sàng).\n"
        "• Nhóm II (Chuẩn hóa quy trình phiếu yêu cầu phần mềm - 02 đầu mục): Gồm Mục 3 và Mục 4 (Kiểm soát từ chối phiếu thiếu trường, lập và theo dõi vòng đời tối thiểu 03 phiếu thực tế hằng ngày từ ngày 18/11)."
    )

    # 2. Chi tiết các vấn đề đã xử lý
    add_heading_styled(doc, "2. Chi tiết các vấn đề ĐÃ XỬ LÝ (Kèm kết quả định lượng và bằng chứng)", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    p.add_run(
        "Mặc dù gặp nhiều khó khăn về điều kiện kỹ thuật đầu vào (chỉ được cung cấp web UAT, không có PRD/SRS hay bộ dữ liệu chuẩn), chuyên viên kiểm thử đã chủ động vận dụng tư duy logic kiểm thử, khảo sát mã nguồn giao diện web và cấu trúc API qua DevTools để hoàn thành khối lượng lớn công việc cụ thể:"
    )

    # Bảng tổng hợp đã xử lý
    tbl_done = doc.add_table(rows=8, cols=4)
    tbl_done.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_done.autofit = False
    d_headers = ["Đầu mục công việc", "Nhiệm vụ theo chỉ đạo", "Kết quả đã xử lý thực tế", "Sản phẩm / Bằng chứng bàn giao"]
    d_w = [Inches(1.5), Inches(1.8), Inches(2.2), Inches(1.6)]

    for c_idx, h in enumerate(d_headers):
        cell = tbl_done.rows[0].cells[c_idx]
        cell.width = d_w[c_idx]
        cell.text = h
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.runs[0]
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 35, 35, 35, 35)
        set_cell_border(cell, top=('single', '4', '0', 'CBD5E1'), bottom=('single', '4', '0', 'CBD5E1'), left=('single', '4', '0', 'CBD5E1'), right=('single', '4', '0', 'CBD5E1'))

    done_rows = [
        ("Mục 1:\nMôi trường UAT", "Cấp tài khoản UAT Inxify; ký cam kết bảo mật; khóa quyền sửa HBSoft.", "Đã tiếp nhận 09 tài khoản phân quyền ngày 19/11; ký cam kết bảo mật và khóa tài khoản dữ liệu thật ngày 18/11.", "09 tài khoản phân quyền UAT (taikhoan.md); Biên bản cam kết bảo mật."),
        ("Mục 2:\nĐối chiếu 02 màn hình", "Lập bảng đối chiếu trường dữ liệu cho 02 màn hình trên Inxify (Hạn 02/12).", "Hoàn thành trước hạn: Lập bảng đối chiếu 20 trường Đặt phòng và 18 trường Khách hàng qua khảo sát web/API.", "File Excel Mau_Excel_QC_Inxify_DoiChieu_va_TheoDoiPhieu_Hoan_Chinh.xlsx."),
        ("Mục 3 & 4:\nKiểm soát phiếu yêu cầu", "Trả lại phiếu thiếu trường; lập và theo dõi tối thiểu 03 phiếu hằng ngày từ 18/11.", "Đã trả lại 01 phiếu sai chuẩn của Phòng Kinh doanh; lập và theo dõi 04 phiếu thực tế (đóng 2, đang xử lý 1, trả lại 1).", "Biên bản trả lại phiếu YC-A25-20261118-004; Sổ THEO_DOI_PHIEU và BAO_CAO_TUAN."),
        ("Mục 5:\nQuét QR CCCD", "Viết kịch bản kiểm thử nghiệm thu tính năng trên Inxify.", "Đã lập 06 kịch bản nghiệm thu + Mock Data 7 trường BCA; kiểm thử xác nhận luồng nhập tay thủ công ĐÃ PASS (TC-QR-06).", "Tài liệu Tai_lieu_Kich_ban_Kiem_thu_Nghiem_thu_Quet_QR_CCCD_Inxify_Muc_5.docx."),
        ("Mục 6:\nCảnh báo lưu dữ liệu", "Viết kịch bản và kiểm thử cảnh báo khi lưu cho Đặt phòng và Khách hàng.", "Thiết kế 10 kịch bản; chạy 100% trên UAT (7 Pass, 3 Fail); phát hiện 03 lỗi (thiếu viền đỏ, bỏ trống tên/giá, SĐT sai).", "Báo cáo nghiệm thu Bao_cao_Kiem_thu_Canh_bao_Luu_Du_lieu_Inxify_Muc_6_Co_Anh_Dan_Chung.docx (07 ảnh UAT)."),
        ("Mục 7:\nNhật ký Audit Log (Ưu tiên số 1)", "Kiểm thử nhật ký theo trường cho 05 nghiệp vụ trọng tâm (Hạn 20/11).", "Hoàn thành đúng hạn 20/11: Thiết kế Ma trận 35 TC; chạy 100% trên UAT (27 Pass, 8 Fail); phát hiện 08 lỗi ghi thiếu trường.", "Báo cáo trình BGD Bao_cao_Trinh_Ban_Giam_Doc_Ket_Qua_Nghiem_Thu_Audit_Log_Inxify_Muc_7.docx & Excel."),
        ("Mục 8 & 9:\nNguồn OTA & Phòng chưa sẵn sàng", "Lập 01 văn bản tổng hợp yêu cầu kỹ thuật đưa vào danh mục nghiệm thu.", "Đã quy chuẩn hóa 05 tiêu chuẩn kỹ thuật (REQ-M8-01 -> REQ-M9-02) và 08 kịch bản nghiệm thu đón đầu cho Dev.", "Văn bản Van_ban_Tong_hop_Yeu_cau_Ky_thuat_Nghiem_thu_Inxify_Muc_8_va_Muc_9_Co_Anh_Dan_Chung.docx.")
    ]

    for r_idx, row in enumerate(done_rows, start=1):
        for c_idx, val in enumerate(row):
            cell = tbl_done.rows[r_idx].cells[c_idx]
            cell.width = d_w[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            r = p.runs[0]
            r.font.name = "Times New Roman"
            r.font.size = Pt(8.5)
            if c_idx == 0:
                r.bold = True
            set_cell_background(cell, "F9FAFB" if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, 30, 30, 30, 30)
            set_cell_border(cell, top=('single', '4', '0', 'D1D5DB'), bottom=('single', '4', '0', 'D1D5DB'), left=('single', '4', '0', 'D1D5DB'), right=('single', '4', '0', 'D1D5DB'))

    # Bảng Bộ 04 con số định lượng
    p_q = doc.add_paragraph()
    p_q.paragraph_format.space_before = Pt(8)
    p_q.paragraph_format.space_after = Pt(4)
    r_q = p_q.add_run("Bảng tổng hợp Bộ 04 chỉ số định lượng chấm kết thúc thử việc (Căn cứ chấm đạt theo yêu cầu của Sếp):")
    r_q.bold = True
    r_q.font.name = "Times New Roman"
    r_q.font.size = Pt(10.5)

    tbl_num = doc.add_table(rows=5, cols=3)
    tbl_num.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_num.autofit = False
    n_w = [Inches(2.5), Inches(1.5), Inches(3.1)]
    n_headers = ["Chỉ số định lượng Sếp yêu cầu", "Kết quả thực tế đạt được", "Căn cứ hồ sơ & Diễn giải chi tiết"]

    for c_idx, h in enumerate(n_headers):
        cell = tbl_num.rows[0].cells[c_idx]
        cell.width = n_w[c_idx]
        cell.text = h
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.runs[0]
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 35, 35, 35, 35)
        set_cell_border(cell, top=('single', '4', '0', 'CBD5E1'), bottom=('single', '4', '0', 'CBD5E1'), left=('single', '4', '0', 'CBD5E1'), right=('single', '4', '0', 'CBD5E1'))

    num_rows = [
        ("1. Số chức năng đã có kịch bản được duyệt trên danh sách duyệt 20/11", "05 chức năng lớn\n(59 Test Cases)", "Audit Log Mục 7 (35 TC); Cảnh báo lưu dữ liệu Mục 6 (10 TC); Quét QR CCCD Mục 5 (06 TC); Nguồn kênh OTA Mục 8 (04 TC); Cảnh báo phòng chưa sẵn sàng Mục 9 (04 TC)."),
        ("2. Số kịch bản đã chạy thực tế trên Inxify", "45 kịch bản\n(45/59 TC đã chạy)", "35/35 TC Audit Log đã chạy (27 Pass, 8 Fail); 10/10 TC Cảnh báo lưu dữ liệu đã chạy (7 Pass, 3 Fail). 14 TC còn lại (Mục 5, 8, 9) là kịch bản nghiệm thu đón đầu chờ Dev release code."),
        ("3. Số lỗi tìm được\n(Mỗi lỗi 01 phiếu yêu cầu phần mềm)", "11 lỗi thực tế\n(04 phiếu trên sổ)", "08 lỗi Audit Log ghi thiếu trường cũ/mới; 03 lỗi validate form Đặt phòng. Đã lập 04 phiếu yêu cầu (YC-001, YC-002 đã đóng; YC-003 đang xử lý; YC-004 trả lại theo Mục 3)."),
        ("4. Số lần hỏi lại người hướng dẫn (Mentor)", "06 lần chốt quy tắc", "Lần 1 (04/11: quyền dữ liệu thật); Lần 2 (08/11: cọc/hoàn tiền OTA); Lần 3 (18/11: UAT Inxify); Lần 4 (19/11: 9 tài khoản UAT); Lần 5 (19/11: chuẩn phiếu); Lần 6 (20/11: kết quả Audit log).")
    ]

    for r_idx, row in enumerate(num_rows, start=1):
        for c_idx, val in enumerate(row):
            cell = tbl_num.rows[r_idx].cells[c_idx]
            cell.width = n_w[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 1 else WD_ALIGN_PARAGRAPH.LEFT
            r = p.runs[0]
            r.font.name = "Times New Roman"
            r.font.size = Pt(9)
            if c_idx == 0:
                r.bold = True
            elif c_idx == 1:
                r.bold = True
                r.font.color.rgb = RGBColor(180, 20, 20)
            set_cell_background(cell, "F9FAFB" if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, 35, 35, 35, 35)
            set_cell_border(cell, top=('single', '4', '0', 'D1D5DB'), bottom=('single', '4', '0', 'D1D5DB'), left=('single', '4', '0', 'D1D5DB'), right=('single', '4', '0', 'D1D5DB'))

    # 3. Tồn đọng
    add_heading_styled(doc, "3. Chi tiết các vấn đề CÒN TỒN ĐỌNG & PHÂN TÍCH NGUYÊN NHÂN", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    p.add_run(
        "Đứng trên vai trò của một Tester thực chiến chỉ được cấp môi trường UAT Portal, chuyên viên kiểm thử xin báo cáo trung thực 06 vấn đề tồn đọng cốt lõi và nguyên nhân khách quan xuất phát từ điều kiện thực tế tại đơn vị:\n"
        "a. Tồn đọng 1: Thiếu tài liệu đặc tả yêu cầu nghiệp vụ chính thức (PRD / SRS / Data Dictionary)\n"
        "• Thực trạng: Ngoại trừ tài liệu BA v0.3 (mới ở mức dự thảo sơ bộ cho chức năng Audit Log), toàn bộ các phân hệ còn lại của Inxify (Đặt phòng, Buồng phòng, Check-in/Check-out, Thu ngân, Công nợ, Khách hàng) đều hoàn toàn chưa có tài liệu PRD/SRS chính thức.\n"
        "• Nguyên nhân: Dự án Inxify đang trong giai đoạn chuyển đổi gấp rút từ hệ thống cũ; đội ngũ BA chưa kịp chuẩn hóa và ban hành hệ thống tài liệu yêu cầu nghiệp vụ.\n"
        "• Tác động: Tester phải tự khảo sát giao diện web UAT, soi F12/API để suy đoán logic và tự định nghĩa Expected Result; tiềm ẩn nguy cơ sai lệch so với vận hành thực tế tại khách sạn và dễ dẫn đến tranh cãi khi bắt lỗi.\n\n"
        "b. Tồn đọng 2: Môi trường UAT thiếu Bộ dữ liệu kiểm thử chuẩn (Standard Test Data)\n"
        "• Thực trạng: Cơ sở dữ liệu UAT Inxify còn mới, chắp vá hoặc rỗng. Thiếu hẳn các bộ dữ liệu mẫu điển hình của chuỗi A25 (hạng phòng Standard/Deluxe/VIP, bảng giá giờ/ngày/đêm, đối tác lữ hành, công nợ OTA lũy kế, lịch sử phòng).\n"
        "• Nguyên nhân: Chưa có quy trình trích xuất và đồng bộ dữ liệu mẫu (đã ẩn danh) từ môi trường vận hành thật sang UAT.\n"
        "• Tác động: Tester phải mất nhiều thời gian tự nhập tay dữ liệu mồi thủ công từng bước; không kiểm thử được các kịch bản phức tạp liên hoàn (End-to-End) hay kiểm thử tải/áp lực.\n\n"
        "c. Tồn đọng 3: Tính năng Quét QR CCCD (Mục 5) chưa thể kiểm thử phần cứng thực tế\n"
        "• Thực trạng: Tại bàn làm việc của Tester hoàn toàn CHƯA CÓ đầu đọc mã vạch 2D cầm tay (cổng USB); KHÔNG CÓ thẻ CCCD cứng thật của khách hàng (tuân thủ quy định bảo vệ dữ liệu cá nhân); phần mềm Inxify CHƯA CÓ tính năng quét qua camera điện thoại; trên giao diện web Inxify hiện tại MỚI CHỈ CÓ FORM NHẬP TAY THỦ CÔNG (Manual Input), Dev chưa ra mắt tính năng nhận diện chuỗi quét QR.\n"
        "• Nguyên nhân: Thiết bị máy quét 2D đang được trang bị tại các quầy lễ tân chi nhánh đang kinh doanh; phía Dev Inxify chưa bàn giao phiên bản có tích hợp thư viện quét mã lên UAT.\n"
        "• Tác động: Tester mới chỉ kiểm thử được luồng nhập tay thủ công (TC-QR-06 ĐÃ PASS) và lập bộ 06 kịch bản nghiệm thu đón đầu bằng chuỗi dữ liệu giả lập (Mock Data); chưa thể kiểm thử tốc độ nhận diện phần cứng ngoài thực tế quầy lễ tân.\n\n"
        "d. Tồn đọng 4: 02 tính năng Hiển thị nguồn OTA (Mục 8) và Cảnh báo nhận phòng chưa sẵn sàng (Mục 9) chưa được Dev lập trình lên UAT\n"
        "• Thực trạng: Modal Nhận phòng từ đặt phòng (CheckInFromReservationModal) vẫn chưa được bổ sung 02 trường nguồn kênh và mã booking OTA; thao tác nhận phòng vào phòng bẩn (Dirty) hoặc phòng sửa chữa (OOO) vẫn diễn ra bình thường, chưa có cảnh báo hay chặn chuyển trạng thái.\n"
        "• Nguyên nhân: Đội ngũ Lập trình viên (Dev) Inxify chưa cập nhật mã nguồn của Mục 8 và Mục 9 lên môi trường UAT.\n"
        "• Tác động: 08 kịch bản nghiệm thu mới chỉ nằm ở dạng thiết kế đón đầu; Tester chưa thể bấm chạy thực tế để kết luận Đạt/Không đạt.\n\n"
        "e. Tồn đọng 5: 11 lỗi hệ thống đã phát hiện trên UAT Inxify đang chờ Dev khắc phục triệt để\n"
        "• Thực trạng: Gồm 08 lỗi Audit Log (ghi thiếu trường cũ/mới, sai định dạng thời gian, thiếu log hủy cọc) và 03 lỗi validate form Đặt phòng (bỏ trống tên/giá vẫn lưu được, SĐT sai định dạng vẫn nhận).\n"
        "• Nguyên nhân: Mã nguồn phía Backend/Frontend của đối tác phát triển Inxify chưa kiểm soát chặt chẽ ở tầng Data Validation và Event Listener.\n"
        "• Tác động: Ảnh hưởng trực tiếp đến tính toàn vẹn dữ liệu và tiêu chuẩn nghiệm thu 'Ưu tiên số 1: Ai sửa – Sửa gì – Lúc nào'.\n\n"
        "f. Tồn đọng 6: Quy trình xử lý phiếu yêu cầu phần mềm còn mang tính thủ công, thói quen phòng ban chưa thay đổi\n"
        "• Thực trạng: Việc theo dõi phiếu hiện vẫn phải quản lý thủ công qua file Excel; một số nhân sự phòng ban ngoài (Kinh doanh, Lễ tân) vẫn nhắn tin Zalo hoặc yêu cầu miệng, khi được yêu cầu điền form phiếu thì còn lúng túng hoặc điền thiếu thông tin.\n"
        "• Nguyên nhân: Hệ thống Inxify/HROnline chưa có phân hệ Helpdesk/Ticketing điện tử tập trung."
    )

    # 4. Phương án giải quyết
    add_heading_styled(doc, "4. Phương án giải quyết chi tiết cho từng vấn đề tồn đọng", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    p.add_run(
        "Nhằm giải quyết dứt điểm các vướng mắc trên và bảo đảm chất lượng nghiệm thu hệ thống trước mốc thời gian 16h00 ngày 02/12/2026, chuyên viên kiểm thử phối hợp cùng Trưởng ban Nguyễn Văn Sơn đề xuất các phương án hành động cụ thể:\n"
        "• Phương án cho Tồn đọng 1 (Tài liệu nghiệp vụ): Sử dụng chính Bảng đối chiếu trường dữ liệu (20 trường Đặt phòng, 18 trường Khách hàng) và Ma trận trường Audit Log đã lập làm tài liệu kỹ thuật cơ sở (Baseline Specs); gửi BA Châu Tuấn và Trưởng ban Sơn xác nhận làm căn cứ kiểm thử chính thức.\n"
        "• Phương án cho Tồn đọng 2 (Dữ liệu kiểm thử UAT): Đề xuất Phòng CNTT phối hợp Quản trị hệ thống trích xuất một bản sao dữ liệu (Database Snapshot) từ hệ thống vận hành; thực hiện Ẩn danh hóa thông tin nhạy cảm (Data Masking) để xóa/mã hóa số CCCD, SĐT, Tên thật rồi nạp vào UAT Inxify trước ngày 25/11/2026.\n"
        "• Phương án cho Tồn đọng 3 (Thiết bị Quét QR CCCD - Mục 5): Khi Dev Inxify thông báo hoàn thành tính năng nhận diện chuỗi quét QR trên UAT, Phòng CNTT làm thủ tục mượn 01 đầu đọc mã vạch 2D từ quầy Lễ tân A25 Hotel (chi nhánh 684 Minh Khai) kết nối máy tính văn phòng trong 01 buổi (khoảng 30 phút). Sử dụng thẻ CCCD công vụ của nhân sự CNTT để chạy kiểm thử nghiệm thu phần cứng thực tế theo đúng 06 kịch bản đã xây dựng.\n"
        "• Phương án cho Tồn đọng 4 & 5 (Xử lý 11 lỗi & Đẩy nhanh Mục 8, 9): Bàn giao danh mục Bug Tracking chi tiết cho đội ngũ Dev Inxify. Chốt thời hạn: Dev hoàn thành fix 11 lỗi trước ngày 25/11/2026 và release tính năng Mục 8, 9 lên UAT trước ngày 27/11/2026. Tester thực thi kiểm thử hồi quy (Regression Test) toàn bộ kịch bản từ ngày 28/11 đến 30/11/2026.\n"
        "• Phương án cho Tồn đọng 6 (Quy trình tiếp nhận phiếu): Kiên quyết từ chối tiếp nhận và lập biên bản trả lại 100% các yêu cầu gửi qua Zalo hoặc phản ánh miệng; hướng dẫn nhân sự điền Biểu mẫu phiếu yêu cầu phần mềm đã ban hành. Kiến nghị tích hợp phân hệ Helpdesk trực tiếp trên Inxify/HROnline trong giai đoạn tiếp theo."
    )

    # IV. ĐỀ XUẤT VÀ TỔ CHỨC THỰC HIỆN
    add_heading_styled(doc, "IV. ĐỀ XUẤT VÀ TỔ CHỨC THỰC HIỆN", 1)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    p.add_run(
        "Kính trình Ban Tổng Giám đốc và Lãnh đạo Phòng Công nghệ thông tin phê duyệt kế hoạch tổ chức thực hiện các nội dung trọng tâm từ nay đến mốc báo cáo kết thúc thử việc (02/12/2026):"
    )

    tbl_plan = doc.add_table(rows=7, cols=4)
    tbl_plan.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_plan.autofit = False
    pl_w = [Inches(0.6), Inches(3.2), Inches(2.2), Inches(1.1)]
    pl_headers = ["STT", "Nội dung công việc trọng tâm", "Đơn vị / Đầu mối thực hiện", "Thời hạn hoàn thành"]

    for c_idx, h in enumerate(pl_headers):
        cell = tbl_plan.rows[0].cells[c_idx]
        cell.width = pl_w[c_idx]
        cell.text = h
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.runs[0]
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 35, 35, 35, 35)
        set_cell_border(cell, top=('single', '4', '0', 'CBD5E1'), bottom=('single', '4', '0', 'CBD5E1'), left=('single', '4', '0', 'CBD5E1'), right=('single', '4', '0', 'CBD5E1'))

    plan_rows = [
        ("01", "Bàn giao danh mục 11 lỗi hệ thống và văn bản tiêu chuẩn kỹ thuật Mục 8, 9 cho đội ngũ Lập trình viên Inxify.", "Tester Hoàng Minh Quân phối hợp Trưởng ban Sơn", "22/11/2026"),
        ("02", "Trích xuất, thực hiện Ẩn danh hóa (Data Masking) và nạp bộ dữ liệu mẫu chuẩn của chuỗi A25 lên môi trường UAT Inxify.", "Phòng CNTT phối hợp Bộ phận Quản trị hệ thống", "25/11/2026"),
        ("03", "Đội ngũ Dev Inxify hoàn thành sửa chữa 11 lỗi hệ thống và cập nhật mã nguồn tính năng Mục 8, 9 lên môi trường UAT.", "Đội ngũ Lập trình viên đối tác Inxify", "27/11/2026"),
        ("04", "Làm thủ tục mượn đầu đọc 2D tại Lễ tân A25 Minh Khai, thực hiện kiểm thử nghiệm thu phần cứng tính năng Quét QR CCCD.", "Tester Hoàng Minh Quân phối hợp Quầy Lễ tân", "28/11/2026"),
        ("05", "Thực thi kiểm thử hồi quy (Regression Test) toàn diện 59 Test Cases trên môi trường UAT Inxify; chốt kết quả PASS/FAIL.", "Tester Hoàng Minh Quân", "29/11 – 30/11/2026"),
        ("06", "Tổng hợp toàn bộ hồ sơ, dữ liệu định lượng và hoàn thiện Báo cáo kết thúc thử việc chính thức trình Ban Giám đốc và HR.", "Tester Hoàng Minh Quân & Trưởng ban Nguyễn Văn Sơn", "Trước 16h00 ngày 02/12/2026")
    ]

    for r_idx, row in enumerate(plan_rows, start=1):
        for c_idx, val in enumerate(row):
            cell = tbl_plan.rows[r_idx].cells[c_idx]
            cell.width = pl_w[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx in (0, 3) else WD_ALIGN_PARAGRAPH.LEFT
            r = p.runs[0]
            r.font.name = "Times New Roman"
            r.font.size = Pt(9)
            if c_idx == 0:
                r.bold = True
            elif c_idx == 3:
                r.bold = True
                r.font.color.rgb = RGBColor(180, 20, 20)
            set_cell_background(cell, "F9FAFB" if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, 35, 35, 35, 35)
            set_cell_border(cell, top=('single', '4', '0', 'D1D5DB'), bottom=('single', '4', '0', 'D1D5DB'), left=('single', '4', '0', 'D1D5DB'), right=('single', '4', '0', 'D1D5DB'))

    # V. ĐÁNH GIÁ CHẤT LƯỢNG HỆ THỐNG INXIFY THEO CÁC TIÊU CHÍ (DIMENSIONS) & HIỆU NĂNG (PERFORMANCE)
    add_heading_styled(doc, "V. ĐÁNH GIÁ CHẤT LƯỢNG HỆ THỐNG INXIFY PMS THEO CÁC TIÊU CHÍ (DIMENSIONS) & HIỆU NĂNG (PERFORMANCE)", 1)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    p.add_run(
        "Nhằm cung cấp cho Ban Giám đốc cái nhìn khách quan, đa chiều và có cơ sở kỹ thuật định lượng chuẩn xác về chất lượng của phần mềm Inxify PMS trước khi đưa vào vận hành thực tế toàn chuỗi, chuyên viên kiểm thử xây dựng Khung đánh giá chất lượng hệ thống theo 05 trục tiêu chí (Quality Dimensions) theo thang điểm 10, kết hợp số liệu đo kiểm thử tải và hiệu năng thực tế bằng công cụ chuyên dụng k6:"
    )

    # V.1 Bảng chấm điểm 5 tiêu chí Dimension
    add_heading_styled(doc, "1. Bảng chấm điểm chất lượng hệ thống Inxify PMS theo 05 tiêu chí (Thang điểm 10) & Log chi tiết", 2)
    
    tbl_dim = doc.add_table(rows=7, cols=5)
    tbl_dim.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_dim.autofit = False
    dim_w = [Inches(0.45), Inches(1.85), Inches(0.7), Inches(0.8), Inches(3.3)]
    dim_headers = ["STT", "Tiêu chí đánh giá (Dimension)", "Trọng số", "Điểm (Thang 10)", "Log chi tiết căn cứ chấm điểm & Bằng chứng thực tế"]

    for c_idx, h in enumerate(dim_headers):
        cell = tbl_dim.rows[0].cells[c_idx]
        cell.width = dim_w[c_idx]
        cell.text = h
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.runs[0]
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 35, 35, 35, 35)
        set_cell_border(cell, top=('single', '4', '0', 'CBD5E1'), bottom=('single', '4', '0', 'CBD5E1'), left=('single', '4', '0', 'CBD5E1'), right=('single', '4', '0', 'CBD5E1'))

    dim_rows = [
        ("01", "Chức năng & Nghiệp vụ\n(Functionality)", "25%", "7.0 / 10\n(Khá)", "• Điểm cộng: 07/09 phân hệ cốt lõi đã có trên UAT (Sơ đồ phòng, Đặt phòng, Khách hàng, Buồng phòng, Kho, Thu ngân). Luồng nhập tay thông tin CCCD ĐÃ PASS (TC-QR-06).\n• Điểm trừ (-3.0): Mục 8 (Nguồn OTA) và Mục 9 (Chặn nhận phòng bẩn Dirty/OOO) chưa được Dev lập trình lên UAT; Tính năng Quét QR CCCD chưa có thư viện nhận diện chuỗi; 14/59 TC mới ở dạng kịch bản đón đầu."),
        ("02", "Trải nghiệm người dùng\n(UI/UX Usability)", "20%", "6.5 / 10\n(Trung bình)", "• Điểm cộng: Giao diện React Ant Design hiện đại, sơ đồ phòng phân màu trực quan (Xanh - Trống, Dương - Ở, Cam - Bẩn, Tím - Vấn đề).\n• Điểm trừ (-3.5): Form Đặt phòng khi bấm Lưu thiếu trường bắt buộc KHÔNG HIỆN VIỀN ĐỎ cảnh báo ô lỗi; Màn hình CRM Khách hàng (/crm/guests) không có nút '+ Thêm khách hàng' trực tiếp; Lớp phủ Ant Tour che nút; Thiếu phím tắt thao tác nhanh cho Lễ tân."),
        ("03", "Hiệu năng & Tải\n(Performance)", "20%", "6.0 / 10\n(Trung bình)", "• Điểm cộng: Đo k6 thực tế 15 API cốt lõi. Tải nhẹ (1-7 VU) rất mượt, 7/15 API phản hồi < 80ms; Throughput đỉnh 22.58 req/s; Tự phục hồi cực nhanh (Resilience: về 990ms sau khi hạ tải 1000 VU).\n• Điểm trừ (-4.0): API Sơ đồ phòng (/rooms/map) nặng 115 KB thô, p95 vọt lên 4,653ms ở 100 VU; Báo cáo điều hành mất 883ms; Tải vượt 300 VU lỗi 18.2%, 500 VU lỗi 46.8% do cạn kiệt Database Connection Pool (Timeout 25s - BUG-003)."),
        ("04", "Toàn vẹn & Kiểm toán\n(Data Integrity & Audit)", "20%", "5.5 / 10\n(Rủi ro cao)", "• Điểm cộng: Có phân hệ Audit Log riêng biệt, ghi nhận được hành vi 'Ai sửa - Lúc nào' cho 5 nghiệp vụ.\n• Điểm trừ (-4.5): Rủi ro kiểm toán lớn nhất: Audit Log fail 8/35 TC vì KHÔNG LƯU GIÁ TRỊ CŨ (oldValues null), không biết sửa từ giá nào sang giá nào; Form Đặt phòng cho phép lưu SĐT chỉ 5 số (09812), Email sai format; CRM Khách hàng cho phép tạo hồ sơ trùng lặp họ tên, ngày sinh, thiếu cảnh báo gộp hồ sơ."),
        ("05", "Độ ổn định & Hạ tầng\n(Stability & Resilience)", "15%", "7.5 / 10\n(Khá tốt)", "• Điểm cộng: Hạ tầng Cloud duy trì kết nối ổn định, không bị sập nguồn đột ngột; có cơ chế Rate Limiting (HTTP 429) chống tấn công DDoS và spam đăng nhập dồn dập.\n• Điểm trừ (-2.5): Chưa cấu hình Redis Caching cho API dữ liệu lớn, chưa bật nén Gzip/Brotli phía Web Server; cấu hình Connection Pool của Database server UAT còn thấp (chạm trần ở 300 VU)."),
        ("⭐", "ĐIỂM TRUNG BÌNH CÓ TRỌNG SỐ", "100%", "6.45 / 10\n(XẾP LOẠI: KHÁ)", "KẾT LUẬN CHẤT LƯỢNG TỔNG THỂ: Hệ thống Inxify PMS đạt mức ĐỦ ĐIỀU KIỆN TIẾP TỤC KIỂM THỬ UAT CÓ ĐIỀU KIỆN. Tuy nhiên, CHƯA ĐỦ ĐIỀU KIỆN GO-LIVE TOÀN DIỆN CHO CHUỖI KHÁCH SẠN A25 nếu phía đối tác phát triển chưa hoàn thành khắc phục 11 lỗi hệ thống và tối ưu hóa 2 nút thắt hiệu năng cốt lõi.")
    ]

    for r_idx, row in enumerate(dim_rows, start=1):
        for c_idx, val in enumerate(row):
            cell = tbl_dim.rows[r_idx].cells[c_idx]
            cell.width = dim_w[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx in (0, 2, 3) else WD_ALIGN_PARAGRAPH.LEFT
            r = p.runs[0]
            r.font.name = "Times New Roman"
            r.font.size = Pt(8.5)
            if r_idx == 6: # Total row
                r.bold = True
                set_cell_background(cell, "EFF6FF")
                if c_idx == 3:
                    r.font.color.rgb = RGBColor(180, 20, 20)
                    r.font.size = Pt(9.5)
            else:
                if c_idx in (0, 1, 3):
                    r.bold = True
                if c_idx == 3:
                    r.font.color.rgb = RGBColor(20, 50, 120)
                set_cell_background(cell, "F9FAFB" if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, 30, 30, 30, 30)
            set_cell_border(cell, top=('single', '4', '0', 'D1D5DB'), bottom=('single', '4', '0', 'D1D5DB'), left=('single', '4', '0', 'D1D5DB'), right=('single', '4', '0', 'D1D5DB'))

    # V.2 Kết quả đo kiểm thử hiệu năng k6 thực tế
    add_heading_styled(doc, "2. Báo cáo kết quả kiểm thử tải và hiệu năng thực tế (Performance Test bằng k6 v2.4.0)", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    p.add_run(
        "Nhằm chủ động dự báo khả năng đáp ứng của phần mềm khi triển khai đồng loạt trên 196 phòng thuộc hệ thống khách sạn A25, Tester đã độc lập thiết lập và thực thi bài kiểm thử hiệu năng tự động trên môi trường UAT (https://api-uat.inxify.com) thông qua công cụ k6 v2.4.0. Toàn bộ hồ sơ kịch bản và dữ liệu chi tiết được lưu trữ tại thư mục 'performance/':"
    )

    # Bảng phân nhóm 15 API
    tbl_api = doc.add_table(rows=4, cols=4)
    tbl_api.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_api.autofit = False
    api_w = [Inches(1.8), Inches(1.3), Inches(2.3), Inches(1.7)]
    api_headers = ["Phân nhóm API", "Tốc độ phản hồi", "Danh sách API nghiệp vụ Inxify", "Đánh giá & Khuyến nghị"]

    for c_idx, h in enumerate(api_headers):
        cell = tbl_api.rows[0].cells[c_idx]
        cell.width = api_w[c_idx]
        cell.text = h
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.runs[0]
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 30, 30, 30, 30)
        set_cell_border(cell, top=('single', '4', '0', 'CBD5E1'), bottom=('single', '4', '0', 'CBD5E1'), left=('single', '4', '0', 'CBD5E1'), right=('single', '4', '0', 'CBD5E1'))

    api_rows = [
        ("Nhóm Siêu tốc\n(7 API cốt lõi)", "< 80 ms\n(Rất mượt)", "• /properties/my (41.7ms)\n• /housekeeping (41.8ms)\n• /cash-shifts/current (40.8ms)\n• /audit-logs (45.6ms)\n• /rooms/stats (50.0ms)\n• /room-types (43.7ms)\n• /warehouse/units (40.7ms)", "PASS - Tối ưu rất tốt, truy vấn bảng đơn hoặc bảng danh mục ít dữ liệu."),
        ("Nhóm Nghiệp vụ\n(4 API thường dùng)", "100 – 350 ms\n(Đạt yêu cầu)", "• /auth/login (108.6ms)\n• /reservations (116.5ms)\n• /guests (135.0ms)\n• /inventory/items (266.5ms cho 474 vật tư kho)", "PASS - Phục vụ trải nghiệm Lễ tân/Thư ký tốt khi thao tác hằng ngày."),
        ("Nhóm Nặng cần tối ưu\n(2 API nút thắt)", "> 350 ms\n(Cần xử lý gấp)", "• /rooms/map (Sơ đồ 196 phòng: 368ms thô, p95 458ms)\n• /reports/executive (Báo cáo điều hành tổng hợp: 883ms thô, p95 963ms)", "WARN - Nút thắt hiệu năng: Gom dữ liệu lớn, tính toán phức tạp; cần bật Cache và nén dữ liệu.")
    ]

    for r_idx, row in enumerate(api_rows, start=1):
        for c_idx, val in enumerate(row):
            cell = tbl_api.rows[r_idx].cells[c_idx]
            cell.width = api_w[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx in (0, 1) else WD_ALIGN_PARAGRAPH.LEFT
            r = p.runs[0]
            r.font.name = "Times New Roman"
            r.font.size = Pt(8.5)
            if c_idx == 0:
                r.bold = True
            elif c_idx == 1 and r_idx == 3:
                r.bold = True
                r.font.color.rgb = RGBColor(180, 20, 20)
            set_cell_background(cell, "F9FAFB" if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, 30, 30, 30, 30)
            set_cell_border(cell, top=('single', '4', '0', 'D1D5DB'), bottom=('single', '4', '0', 'D1D5DB'), left=('single', '4', '0', 'D1D5DB'), right=('single', '4', '0', 'D1D5DB'))

    # Bảng đo tải các cấp độ
    p_load = doc.add_paragraph()
    p_load.paragraph_format.space_before = Pt(6)
    p_load.paragraph_format.space_after = Pt(3)
    r_l = p_load.add_run("Bảng kết quả đo kiểm thử tải theo từng cấp độ người dùng đồng thời (k6 Execution Metrics):")
    r_l.bold = True
    r_l.font.name = "Times New Roman"
    r_l.font.size = Pt(9.5)

    tbl_load = doc.add_table(rows=6, cols=6)
    tbl_load.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_load.autofit = False
    ld_w = [Inches(1.2), Inches(0.9), Inches(1.1), Inches(1.0), Inches(1.1), Inches(1.8)]
    ld_headers = ["Cấp độ tải", "Số Request", "Thông lượng (req/s)", "Tỷ lệ lỗi (%)", "p95 Sơ đồ phòng", "Đánh giá trạng thái hệ thống"]

    for c_idx, h in enumerate(ld_headers):
        cell = tbl_load.rows[0].cells[c_idx]
        cell.width = ld_w[c_idx]
        cell.text = h
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.runs[0]
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 30, 30, 30, 30)
        set_cell_border(cell, top=('single', '4', '0', 'CBD5E1'), bottom=('single', '4', '0', 'CBD5E1'), left=('single', '4', '0', 'CBD5E1'), right=('single', '4', '0', 'CBD5E1'))

    load_rows = [
        ("1 VU (Baseline)", "30 reqs", "0.64 req/s", "0.00%", "534.4 ms", "PASS - Tốc độ rất nhanh, ổn định."),
        ("7 VUs (Đa Role)", "147 reqs", "4.31 req/s", "0.00%", "912.7 ms", "PASS - Tăng tải tuyến tính tốt."),
        ("100 VUs (Tải cao)", "1,876 reqs", "22.58 req/s", "19.08%", "4,653.8 ms", "WARN - Đạt đỉnh thông lượng xử lý của UAT."),
        ("300 VUs (Mô hình thật)", "1,745 reqs", "27.70 req/s", "18.20%", "6,203.8 ms", "WARN - Người dùng có nghỉ 4s, bắt đầu nghẽn DB pool."),
        ("500 VUs (Mô hình thật)", "2,367 reqs", "35.00 req/s", "46.80%", "9,539.5 ms", "FAIL - Database cạn Connection Pool, timeout 25s.")
    ]

    for r_idx, row in enumerate(load_rows, start=1):
        for c_idx, val in enumerate(row):
            cell = tbl_load.rows[r_idx].cells[c_idx]
            cell.width = ld_w[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx in (0, 1, 2, 3, 4) else WD_ALIGN_PARAGRAPH.LEFT
            r = p.runs[0]
            r.font.name = "Times New Roman"
            r.font.size = Pt(8.5)
            if c_idx == 0:
                r.bold = True
            elif c_idx == 5 and "FAIL" in val:
                r.bold = True
                r.font.color.rgb = RGBColor(180, 20, 20)
            set_cell_background(cell, "F9FAFB" if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, 30, 30, 30, 30)
            set_cell_border(cell, top=('single', '4', '0', 'D1D5DB'), bottom=('single', '4', '0', 'D1D5DB'), left=('single', '4', '0', 'D1D5DB'), right=('single', '4', '0', 'D1D5DB'))

    # Khuyến nghị kỹ thuật tối ưu
    p_rec = doc.add_paragraph()
    p_rec.paragraph_format.space_before = Pt(6)
    p_rec.paragraph_format.space_after = Pt(4)
    p_rec.paragraph_format.line_spacing = 1.15
    p_rec.add_run(
        "• 03 Khuyến nghị kỹ thuật then chốt gửi đội ngũ Dev Inxify:\n"
        "  1. Bật cơ chế Caching Redis cho API Sơ đồ phòng (/rooms/map): Đặt thời gian sống cache (TTL) từ 15 - 30 giây theo từng mã khách sạn (propertyId) để giảm 75% số lần truy vấn JOIN nhiều bảng trong DB.\n"
        "  2. Bật nén dữ liệu HTTP Gzip / Brotli phía Nginx/Backend: Hiện tại response JSON thô của sơ đồ 196 phòng có kích thước ~115 KB. Bật nén sẽ giảm 80% dung lượng truyền tải mạng xuống còn ~20 KB.\n"
        "  3. Nâng dung lượng Database Connection Pool và cấu hình hàng đợi (Queue): Khắc phục triệt để lỗi Request Timeout 25s khi khách sạn vào ca cao điểm có từ 300 đến 500 nhân viên thao tác đồng thời."
    )

    # VI. DANH MỤC HỒ SƠ, TÀI LIỆU VÀ MINH CHỨNG ĐÍNH KÈM (ĐỒNG BỘ ONEDRIVE)
    add_heading_styled(doc, "VI. DANH MỤC HỒ SƠ, TÀI LIỆU VÀ MINH CHỨNG ĐÍNH KÈM (ĐỒNG BỘ ONEDRIVE)", 1)
    
    # Hướng dẫn Sếp mở OneDrive 3 bước
    p_hd = doc.add_paragraph()
    p_hd.paragraph_format.space_before = Pt(3)
    p_hd.paragraph_format.space_after = Pt(4)
    p_hd.paragraph_format.line_spacing = 1.15
    r_hd_t = p_hd.add_run("💡 HƯỚNG DẪN TRA CỨU HỒ SƠ ĐỐI SOÁT NHANH TRÊN ONEDRIVE (DÀNH CHO LÃNH ĐẠO & CÁC PHÒNG BAN):\n")
    r_hd_t.bold = True
    r_hd_t.font.name = "Times New Roman"
    r_hd_t.font.size = Pt(10)
    r_hd_t.font.color.rgb = RGBColor(20, 50, 120)
    r_hd_b = p_hd.add_run(
        "Tại thư mục chia sẻ 'My files > A25_QC_Inxify' trên OneDrive của Phòng Công nghệ thông tin:\n"
        "• Bước 1: Mở thư mục 'dan_chung_bao_cao' để xem toàn bộ hồ sơ nghiệm thu từ Mục 1 đến Mục 8 và tài liệu Performance (gồm file Excel ma trận, Word đặc tả, JSON DevTools và thư mục ảnh dẫn chứng tương ứng từng mục).\n"
        "• Bước 2: Mở thư mục 'kiem_tra_thuc_te_uat_lan_cuoi' để xem Báo cáo kiểm tra thực tế trên web UAT Portal và trọn bộ 20 ảnh đối soát chi tiết Input ➔ Result.\n"
        "• Bước 3: Mở file Word 'PCNTT_Báo cáo kiểm thử pms inxify.docx' (hoặc 'PCNTT_MauBC.docx') tại thư mục gốc để theo dõi báo cáo tổng hợp có chữ ký các cấp."
    )
    r_hd_b.font.name = "Times New Roman"
    r_hd_b.font.size = Pt(9.5)

    # TABLE ATTACHMENTS (STT, Tên tài liệu, Định dạng, Vị trí thư mục & Tệp tin OneDrive, Mục đích nghiệm thu)
    att_headers = ["STT", "Hồ sơ / Tài liệu minh chứng", "Định dạng", "Vị trí thư mục OneDrive & Tệp tin cần mở", "Mục đích phục vụ nghiệm thu"]
    att_w = [Inches(0.45), Inches(2.2), Inches(0.75), Inches(2.5), Inches(1.2)]

    attachments_data = [
        ("01", "Mục 1: Quyết định phân quyền 09 tài khoản UAT & Cam kết NDA", "Markdown / Scan PDF", "📁 dan_chung_bao_cao / 01_Muc_1_UAT_...\n📄 M1_Danh_sach_09_tai_khoan_UAT_Inxify.md", "Minh chứng nhận tài khoản UAT ngày 19/11 & bảo vệ dữ liệu HBSoft."),
        ("02", "Mục 2: Bảng đối chiếu 20 trường Đặt phòng & 18 trường Khách hàng", "Excel (.xlsx)", "📁 dan_chung_bao_cao / 02_Muc_2_Doi_Chieu_...\n📄 M2_Bang_doi_chieu_20_truong_DatPhong...", "Đối chiếu chi tiết 38 trường dữ liệu giữa giao diện web & cấu trúc API."),
        ("03", "Mục 2: Cặp ảnh dẫn chứng Form Đặt phòng & Drawer Khách hàng", "Hình ảnh (.png)", "📁 dan_chung_bao_cao / 02_Muc_2_Doi_Chieu_...\n🖼️ M2_01_DatPhong... & M2_02_KhachHang...", "Minh chứng trực quan 20 trường form đặt phòng & 18 trường Drawer CRM."),
        ("04", "Mục 3 & 4: Sổ theo dõi vòng đời 04 phiếu thực tế & Báo cáo tuần", "Excel (.xlsx)", "📁 dan_chung_bao_cao / 03_Muc_3_Va_4_...\n📄 M3_M4_So_theo_doi_04_phieu_thuc_te...", "Quản lý vòng đời phiếu YC-001 -> YC-004 và công thức thống kê tự động."),
        ("05", "Mục 3: Biên bản trả lại phiếu yêu cầu thiếu thông tin", "Biên bản / Excel", "📁 dan_chung_bao_cao / 03_Muc_3_Va_4_...\n📄 Bien_ban_tra_lai_phieu_YC-004.xlsx", "Căn cứ từ chối tiếp nhận yêu cầu sai quy chuẩn của Phòng Kinh doanh."),
        ("06", "Mục 5: Kịch bản kiểm thử nghiệm thu Quét QR CCCD & Mock Data", "Word (.docx)", "📁 dan_chung_bao_cao / 04_Muc_5_Kich_Ban_...\n📄 M5_Tai_lieu_Kich_ban_nghiem_thu_Quet_QR...", "Bộ 06 kịch bản nghiệm thu đón đầu + Mock Data 7 trường theo chuẩn BCA."),
        ("07", "Mục 5: Ảnh dẫn chứng luồng nhập tay thủ công CCCD ĐÃ PASS", "Hình ảnh (.png)", "📁 dan_chung_bao_cao / 04_Muc_5_Kich_Ban_...\n🖼️ M5_01_Input.png & M5_02_Result.png", "Chứng minh tính năng web hiện có: Nhập tay thủ công đạt chuẩn TC-QR-06."),
        ("08", "Mục 6: Báo cáo kiểm thử cảnh báo dữ liệu khi lưu (kèm 07 ảnh)", "Word (.docx)", "📁 dan_chung_bao_cao / 05_Muc_6_Canh_Bao_...\n📄 M6_Bao_cao_kiem_thu_Canh_bao_luu...", "Báo cáo thực thi 10 Test cases, phát hiện 03 bất cập dữ liệu kèm 07 ảnh UAT."),
        ("09", "Mục 6: Thư mục 07 ảnh dẫn chứng UAT kiểm tra dữ liệu", "Thư mục ảnh PNG", "📁 dan_chung_bao_cao / 05_Muc_6_Canh_Bao_...\n📁 Anh_dan_chung_UAT_Muc_6 /", "Dẫn chứng viền đỏ, lỗi SĐT/Email, trùng lặp hồ sơ TEST QUAN trên CRM."),
        ("10", "Mục 7: Báo cáo trình BGD kết quả nghiệm thu Audit Log (Ưu tiên 1)", "Word (.docx)", "📁 dan_chung_bao_cao / 06_Muc_7_Nhat_Ky_...\n📄 M7_Bao_cao_Trinh_Ban_Giam_Doc_Audit_Log...", "Báo cáo trọng tâm hoàn thành 20/11: 35 TC (27 Pass, 8 Fail), bắt 08 lỗi kiểm toán."),
        ("11", "Mục 7: Ma trận 35 Test Case kiểm thử Audit Log 05 nghiệp vụ", "Excel (.xlsx)", "📁 dan_chung_bao_cao / 06_Muc_7_Nhat_Ky_...\n📄 M7_Ma_tran_35_TestCase_Audit_Log...", "Ma trận trường dữ liệu, bước thực thi và kết quả chi tiết từng ca test."),
        ("12", "Mục 7: Dữ liệu bóc tách DevTools Network API & 13 ảnh UAT", "JSON & PNG", "📁 dan_chung_bao_cao / 06_Muc_7_Nhat_Ky_...\n📁 Du_lieu_DevTools_Payload_JSON /", "Minh chứng kỹ thuật bóc tách payload JSON chứng minh lỗi thiếu oldValues."),
        ("13", "Mục 8 & 9: Văn bản tổng hợp yêu cầu kỹ thuật OTA & Phòng chưa sẵn sàng", "Word (.docx)", "📁 dan_chung_bao_cao / 07_Muc_8_Va_9_...\n📄 M8_M9_Van_ban_Tong_hop_Yeu_cau_ky_thuat...", "Quy chuẩn 05 tiêu chuẩn kỹ thuật & 08 kịch bản nghiệm thu chờ Dev release."),
        ("14", "Kiểm tra thực tế: Báo cáo kết quả kiểm tra thực tế UAT Portal", "Word (.docx)", "📁 kiem_tra_thuc_te_uat_lan_cuoi /\n📄 Bao_cao_Kiem_tra_Thuc_te_UAT_Portal...", "Báo cáo tổng hợp đối soát trực tiếp trên web UAT Portal cho toàn bộ 09 mục."),
        ("15", "Bộ minh chứng: Thư mục 20 ảnh đối soát Input ➔ Result toàn diện", "Thư mục ảnh PNG", "📁 kiem_tra_thuc_te_uat_lan_cuoi /\n📁 Anh_dan_chung_Input_Den_Ket_Qua /", "Đầy đủ 20 tệp ảnh đối soát chi tiết từ dữ liệu nhập đến kết quả phản hồi."),
        ("16", "Hiệu năng: Kế hoạch kiểm thử Performance k6 & File Excel 9 sheet", "Word & Excel", "📁 performance / (hoặc dan_chung_bao_cao/performance)\n📄 Ke_hoach_Performance.docx & Template.xlsx", "Hồ sơ kiểm thử hiệu năng k6 v2.4.0 đo 15 API & kiểm thử tải 1000 VU."),
        ("17", "Mã nguồn kiểm thử tự động & Kho lưu trữ trực tuyến GitHub", "Python & Git Repo", "https://github.com/Quan03112003/tooltester-inxify\n(Kho sao lưu dự phòng trực tuyến)", "Bộ kịch bản Playwright E2E, Network Interception và kho backup online.")
    ]

    tbl_att = doc.add_table(rows=len(attachments_data) + 1, cols=5)
    tbl_att.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_att.autofit = False

    for c_idx, h in enumerate(att_headers):
        cell = tbl_att.rows[0].cells[c_idx]
        cell.width = att_w[c_idx]
        cell.text = h
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.runs[0]
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 35, 35, 35, 35)
        set_cell_border(cell, top=('single', '4', '0', 'CBD5E1'), bottom=('single', '4', '0', 'CBD5E1'), left=('single', '4', '0', 'CBD5E1'), right=('single', '4', '0', 'CBD5E1'))

    for r_idx, row in enumerate(attachments_data, start=1):
        for c_idx, val in enumerate(row):
            cell = tbl_att.rows[r_idx].cells[c_idx]
            cell.width = att_w[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx in (0, 2) else WD_ALIGN_PARAGRAPH.LEFT
            r = p.runs[0]
            r.font.name = "Times New Roman"
            r.font.size = Pt(8)
            if c_idx == 0:
                r.bold = True
            elif c_idx == 1:
                r.bold = True
                r.font.color.rgb = RGBColor(20, 50, 120)
            elif c_idx == 3:
                r.font.name = "Consolas"
                r.font.size = Pt(7.5)
                r.font.color.rgb = RGBColor(15, 23, 42)
            set_cell_background(cell, "F9FAFB" if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, 28, 28, 28, 28)
            set_cell_border(cell, top=('single', '4', '0', 'D1D5DB'), bottom=('single', '4', '0', 'D1D5DB'), left=('single', '4', '0', 'D1D5DB'), right=('single', '4', '0', 'D1D5DB'))

    p_note_od = doc.add_paragraph()
    p_note_od.paragraph_format.space_before = Pt(6)
    p_note_od.paragraph_format.space_after = Pt(4)
    r_no = p_note_od.add_run("(*) Ghi chú lưu trữ đám mây: Toàn bộ thư mục cục bộ 'dan_chung_bao_cao/', 'kiem_tra_thuc_te_uat_lan_cuoi/' và 'performance/' đã được cấu trúc đồng bộ sẵn sàng khớp hoàn toàn 1:1 với thư mục chia sẻ OneDrive của Phòng Công nghệ thông tin.")
    r_no.italic = True
    r_no.font.name = "Times New Roman"
    r_no.font.size = Pt(9)

    # Spacing before Table Sign
    p_sp2 = doc.add_paragraph()
    p_sp2.paragraph_format.space_before = Pt(6)
    p_sp2.paragraph_format.space_after = Pt(4)

    # TABLE: Ký duyệt các cấp (rows=6, cols=2)
    tbl_sign = doc.add_table(rows=6, cols=2)
    tbl_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_sign.autofit = False
    s_w = [Inches(3.2), Inches(3.9)]

    sign_roles = [
        ("Người lập", "Hoàng Minh Quân – Chuyên viên kiểm thử phần mềm\n(Ký và ghi rõ họ tên)\n\n\n"),
        ("Phòng Công nghệ thông tin", "Nguyễn Văn Sơn – Trưởng Ban Phát triển & QLPM\n(Ký và ghi rõ họ tên)\n\n\n"),
        ("Phòng Thanh tra giám sát", "(Ký và ghi rõ họ tên)\n\n\n"),
        ("Giám đốc Kinh doanh", "(Ký và ghi rõ họ tên)\n\n\n"),
        ("Phó Tổng giám đốc", "(Ký và ghi rõ họ tên)\n\n\n"),
        ("Ý kiến Sếp", "(Phê duyệt chủ trương và chỉ đạo tổ chức thực hiện)\n\n\n")
    ]

    for r_idx, (role, sign_box) in enumerate(sign_roles):
        c_lbl = tbl_sign.rows[r_idx].cells[0]
        c_val = tbl_sign.rows[r_idx].cells[1]
        c_lbl.width = s_w[0]
        c_val.width = s_w[1]
        c_lbl.text = role
        c_val.text = sign_box
        
        p_l = c_lbl.paragraphs[0]
        r_l = p_l.runs[0]
        r_l.bold = True
        r_l.font.name = "Times New Roman"
        r_l.font.size = Pt(11)
        
        p_v = c_val.paragraphs[0]
        r_v = p_v.runs[0]
        r_v.font.name = "Times New Roman"
        r_v.font.size = Pt(10)
        
        set_cell_background(c_lbl, "F1F5F9")
        set_cell_background(c_val, "FFFFFF")
        set_cell_margins(c_lbl, 35, 35, 35, 35)
        set_cell_margins(c_val, 35, 35, 35, 35)
        set_cell_border(c_lbl, top=('single', '4', '0', 'CBD5E1'), bottom=('single', '4', '0', 'CBD5E1'), left=('single', '4', '0', 'CBD5E1'), right=('single', '4', '0', 'CBD5E1'))
        set_cell_border(c_val, top=('single', '4', '0', 'CBD5E1'), bottom=('single', '4', '0', 'CBD5E1'), left=('single', '4', '0', 'CBD5E1'), right=('single', '4', '0', 'CBD5E1'))

    # Footer: Nơi nhận / Nơi lưu
    p_nn = doc.add_paragraph()
    p_nn.paragraph_format.space_before = Pt(6)
    p_nn.paragraph_format.space_after = Pt(2)
    r_nn_t = p_nn.add_run("Nơi nhận: ")
    r_nn_t.bold = True
    r_nn_t.font.name = "Times New Roman"
    r_nn_t.font.size = Pt(10)
    r_nn_v = p_nn.add_run("Sếp; Ban Tổng giám đốc; Phòng Thanh tra giám sát; Ban Kiểm soát nội bộ.")
    r_nn_v.font.name = "Times New Roman"
    r_nn_v.font.size = Pt(10)

    p_nl = doc.add_paragraph()
    p_nl.paragraph_format.space_after = Pt(4)
    r_nl_t = p_nl.add_run("Nơi lưu: ")
    r_nl_t.bold = True
    r_nl_t.font.name = "Times New Roman"
    r_nl_t.font.size = Pt(10)
    r_nl_v = p_nl.add_run("Phòng Công nghệ thông tin; Ban Thư ký Tổng hợp.")
    r_nl_v.font.name = "Times New Roman"
    r_nl_v.font.size = Pt(10)

    # Save to all target files
    doc.save(TARGET_FILE)
    doc.save(NAMED_FILE)
    doc.save(ONEDRIVE_FILE)
    doc.save(BACKUP_ROOT_FILE)
    doc.save(BACKUP_OD_FILE)
    if os.path.exists(os.path.dirname(BACKUP_DIR_FILE)):
        doc.save(BACKUP_DIR_FILE)

    # Sync performance directory into dan_chung_bao_cao/performance
    src_perf = os.path.join(BASE_DIR, "performance")
    dst_perf = os.path.join(BASE_DIR, "dan_chung_bao_cao", "performance")
    if os.path.exists(src_perf):
        if not os.path.exists(dst_perf):
            os.makedirs(dst_perf, exist_ok=True)
        for item in os.listdir(src_perf):
            s = os.path.join(src_perf, item)
            d = os.path.join(dst_perf, item)
            if os.path.isfile(s):
                shutil.copy2(s, d)

    print("SUCCESS: Successfully regenerated PCNTT_MauBC.docx with Dimensions, Performance and Clean OneDrive Table!")
    print("SUCCESS: Also generated:", ONEDRIVE_FILE)
    print("SUCCESS: Also synced performance folder into dan_chung_bao_cao/performance")

if __name__ == "__main__":
    create_full_report()
