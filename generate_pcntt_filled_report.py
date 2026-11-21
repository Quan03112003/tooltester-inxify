import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import os
import shutil

TARGET_FILE = r"c:\Users\Admin\Desktop\tester\PCNTT_MauBC.docx"
BACKUP_FILE = r"c:\Users\Admin\Desktop\tester\PCNTT_MauBC_backup.docx"
NAMED_FILE = r"c:\Users\Admin\Desktop\tester\Bao_cao_PCNTT_Tong_hop_Tien_do_Xu_ly_Ton_dong_va_Phuong_an_Giai_quyet.docx"

if os.path.exists(TARGET_FILE) and not os.path.exists(BACKUP_FILE):
    shutil.copy2(TARGET_FILE, BACKUP_FILE)
    print("Created backup of PCNTT_MauBC.docx")

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=70, bottom=70, left=90, right=90):
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
        run.font.color.rgb = RGBColor(0, 0, 0) # Administrative black
    elif level == 2:
        p.paragraph_format.space_before = Pt(9)
        p.paragraph_format.space_after = Pt(3)
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(20, 50, 120) # Subtle deep blue
    elif level == 3:
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        run.font.size = Pt(12.5)
        run.font.color.rgb = RGBColor(50, 50, 50)
    return p

def create_full_report():
    doc = docx.Document()

    # Page Margins (TCVN standard administrative)
    for s in doc.sections:
        s.top_margin = Inches(0.6)
        s.bottom_margin = Inches(0.7)
        s.left_margin = Inches(0.75)
        s.right_margin = Inches(0.65)

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
    r0_3.font.size = Pt(10.5)
    r0_3.italic = True

    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r1_1 = p1.add_run("Hà Nội, ngày 21 tháng 11 năm 2026")
    r1_1.font.name = "Times New Roman"
    r1_1.font.size = Pt(12)
    r1_1.italic = True

    for c in (c0, c1):
        set_cell_border(c)
        set_cell_margins(c, 30, 30, 30, 30)

    # Spacing
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(4)
    p_sp.paragraph_format.space_after = Pt(6)

    # 2. Main Title
    p_t = doc.add_paragraph()
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t.paragraph_format.space_before = Pt(4)
    p_t.paragraph_format.space_after = Pt(4)
    r_t = p_t.add_run("BÁO CÁO")
    r_t.bold = True
    r_t.font.name = "Times New Roman"
    r_t.font.size = Pt(16)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(10)
    r_s = p_sub.add_run("V/v: Tổng hợp tiến độ thực hiện, các vấn đề đã xử lý, tồn đọng và phương án giải quyết trong công tác kiểm thử phần mềm Inxify PMS & Hệ thống số A25 Hotel")
    r_s.bold = True
    r_s.font.name = "Times New Roman"
    r_s.font.size = Pt(12.5)
    r_s.italic = True

    # I. CĂN CỨ
    add_heading_styled(doc, "I. CĂN CỨ", 1)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    p.add_run(
        "• Căn cứ Kết luận chỉ đạo của Ban Giám đốc và Lãnh đạo Phòng Công nghệ thông tin tại văn bản hronline.md (tháng 11/2026) về 09 đầu mục nhiệm vụ trọng tâm;\n"
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
        "• Báo cáo tổng thể toàn diện tiến độ triển khai toàn bộ các đầu mục công việc được giao trong dự án tính đến ngày 21/11/2026;\n"
        "• Báo cáo chi tiết các vấn đề ĐÃ XỬ LÝ gắn liền với kết quả định lượng cụ thể, sản phẩm hồ sơ bàn giao, số lượng kịch bản kiểm thử, lỗi phát hiện và bằng chứng thực tế;\n"
        "• Báo cáo trung thực, minh bạch các vấn đề CÒN TỒN ĐỌNG và nguyên nhân khách quan xuất phát từ điều kiện thực tế (chỉ có UAT Portal, thiếu tài liệu và dữ liệu chuẩn, tính năng CCCD chưa có phần cứng và màn hình mới có form nhập tay);\n"
        "• Đề xuất PHƯƠNG ÁN GIẢI QUYẾT khả thi, chuyên nghiệp và kế hoạch hành động cụ thể để khắc phục toàn bộ các tồn đọng, bảo đảm hoàn thành hồ sơ kết thúc thử việc chất lượng cao trình Ban Giám đốc trước 16h00 ngày 02/12/2026."
    )

    # III. NỘI DUNG BÁO CÁO
    add_heading_styled(doc, "III. NỘI DUNG BÁO CÁO", 1)

    # 1. Danh mục các đầu mục
    add_heading_styled(doc, "1. Tổng hợp hiện trạng toàn bộ các đầu mục công việc trong dự án", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    p.add_run(
        "Trong giai đoạn thử việc vừa qua, chuyên viên kiểm thử Hoàng Minh Quân được giao phụ trách và phối hợp thực hiện 03 nhóm nhiệm vụ trọng tâm gồm 11 đầu mục công việc cụ thể:\n"
        "• Nhóm I (Kiểm thử hệ thống Inxify PMS - 07 đầu mục): Gồm các Mục 1, 2, 5, 6, 7, 8, 9 theo chỉ đạo hronline.md (UAT Inxify, Đối chiếu 02 màn hình, Quét QR CCCD, Cảnh báo lưu dữ liệu, Nhật ký Audit Log, Nguồn kênh OTA, Cảnh báo nhận phòng chưa sẵn sàng).\n"
        "• Nhóm II (Chuẩn hóa quy trình phiếu yêu cầu - 02 đầu mục): Gồm Mục 3 và Mục 4 (Kiểm soát từ chối phiếu thiếu trường, lập và theo dõi vòng đời tối thiểu 03 phiếu thực tế hằng ngày từ 18/11).\n"
        "• Nhóm III (Chuẩn hóa hệ sinh thái số & Đào tạo M365 - 02 đầu mục): Xây dựng Báo cáo phương án chuẩn hóa Microsoft 365 và Sản xuất bộ học liệu đa phương tiện (Video HD 720p giọng Neural Microsoft 02:05, Trình phát tương tác HTML5)."
    )

    # 2. Chi tiết các vấn đề đã xử lý
    add_heading_styled(doc, "2. Chi tiết các vấn đề ĐÃ XỬ LÝ (Kèm kết quả định lượng và bằng chứng)", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    p.add_run(
        "Mặc dù gặp nhiều khó khăn về điều kiện kỹ thuật đầu vào, chuyên viên kiểm thử đã chủ động vận dụng tư duy logic kiểm thử, khảo sát mã nguồn giao diện web và cấu trúc API qua DevTools để hoàn thành khối lượng lớn công việc cụ thể:"
    )

    # Bảng tổng hợp đã xử lý
    tbl_done = doc.add_table(rows=8, cols=4)
    tbl_done.alignment = WD_TABLE_ALIGNMENT.CENTER
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
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 40, 40, 40, 40)
        set_cell_border(cell, top=('single', '4', '0', 'CBD5E1'), bottom=('single', '4', '0', 'CBD5E1'), left=('single', '4', '0', 'CBD5E1'), right=('single', '4', '0', 'CBD5E1'))

    done_rows = [
        ("Mục 1:\nMôi trường UAT", "Cấp tài khoản UAT Inxify; ký cam kết bảo mật; khóa quyền sửa HBSoft.", "Đã tiếp nhận 09 tài khoản phân quyền ngày 19/11; ký cam kết bảo mật và khóa tài khoản dữ liệu thật ngày 18/11.", "09 tài khoản phân quyền UAT (taikhoan.md); Biên bản cam kết bảo mật."),
        ("Mục 2:\nĐối chiếu 02 màn hình", "Lập bảng đối chiếu trường dữ liệu cho 02 màn hình trên Inxify (Hạn 02/12).", "Hoàn thành trước hạn: Lập bảng đối chiếu 20 trường Đặt phòng và 18 trường Khách hàng qua khảo sát web/API.", "File Excel Mau_Excel_QC_Inxify_DoiChieu_va_TheoDoiPhieu_Hoan_Chinh.xlsx."),
        ("Mục 3 & 4:\nKiểm soát phiếu yêu cầu", "Trả lại phiếu thiếu trường; lập và theo dõi tối thiểu 03 phiếu hằng ngày từ 18/11.", "Đã trả lại 01 phiếu sai chuẩn của Phòng Kinh doanh; lập và theo dõi 04 phiếu thực tế (đóng 2, đang xử lý 1, trả lại 1).", "Biên bản trả lại phiếu YC-A25-20261118-004; Sổ THEO_DOI_PHIEU và BAO_CAO_TUAN."),
        ("Mục 5:\nQuét QR CCCD", "Viết kịch bản kiểm thử nghiệm thu tính năng trên Inxify.", "Đã lập 06 kịch bản nghiệm thu + Mock Data 7 trường BCA; kiểm thử xác nhận luồng nhập tay thủ công ĐÃ PASS.", "Tài liệu Tai_lieu_Kich_ban_Kiem_thu_Nghiem_thu_Quet_QR_CCCD_Inxify_Muc_5.docx."),
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
            r.font.size = Pt(9.5)
            if c_idx == 0:
                r.bold = True
                r.font.color.rgb = RGBColor(30, 58, 138)
            set_cell_background(cell, "F9FAFB" if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, 35, 35, 35, 35)
            set_cell_border(cell, top=('single', '4', '0', 'D1D5DB'), bottom=('single', '4', '0', 'D1D5DB'), left=('single', '4', '0', 'D1D5DB'), right=('single', '4', '0', 'D1D5DB'))

    # Bảng 04 chỉ số định lượng
    p_num = doc.add_paragraph()
    p_num.paragraph_format.space_before = Pt(8)
    p_num.paragraph_format.space_after = Pt(4)
    r_n_t = p_num.add_run("Bảng tổng hợp Bộ 04 chỉ số định lượng chấm kết thúc thử việc (Căn cứ chấm đạt theo yêu cầu của Sếp):")
    r_n_t.bold = True
    r_n_t.font.name = "Times New Roman"
    r_n_t.font.size = Pt(11)

    tbl_num = doc.add_table(rows=5, cols=3)
    tbl_num.alignment = WD_TABLE_ALIGNMENT.CENTER
    num_headers = ["Chỉ số định lượng Sếp yêu cầu", "Kết quả thực tế đạt được", "Căn cứ hồ sơ & Diễn giải chi tiết"]
    num_w = [Inches(2.5), Inches(1.5), Inches(3.1)]

    for c_idx, h in enumerate(num_headers):
        cell = tbl_num.rows[0].cells[c_idx]
        cell.width = num_w[c_idx]
        cell.text = h
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.runs[0]
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 40, 40, 40, 40)
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
            cell.width = num_w[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 1 else WD_ALIGN_PARAGRAPH.LEFT
            r = p.runs[0]
            r.font.name = "Times New Roman"
            r.font.size = Pt(9.5)
            if c_idx == 0:
                r.bold = True
            elif c_idx == 1:
                r.bold = True
                r.font.color.rgb = RGBColor(180, 20, 20)
            set_cell_background(cell, "F9FAFB" if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, 35, 35, 35, 35)
            set_cell_border(cell, top=('single', '4', '0', 'D1D5DB'), bottom=('single', '4', '0', 'D1D5DB'), left=('single', '4', '0', 'D1D5DB'), right=('single', '4', '0', 'D1D5DB'))

    # 3. Chi tiết các vấn đề còn tồn đọng & phân tích nguyên nhân
    add_heading_styled(doc, "3. Chi tiết các vấn đề CÒN TỒN ĐỌNG & PHÂN TÍCH NGUYÊN NHÂN", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    p.add_run(
        "Đứng trên vai trò của một Tester thực chiến chỉ được cấp môi trường UAT Portal, chuyên viên kiểm thử xin báo cáo trung thực 06 vấn đề tồn đọng cốt lõi và nguyên nhân khách quan xuất phát từ điều kiện thực tế tại đơn vị:\n\n"
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

    # 4. Phương án giải quyết chi tiết
    add_heading_styled(doc, "4. Phương án giải quyết chi tiết cho từng vấn đề tồn đọng", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    p.add_run(
        "Nhằm giải quyết dứt điểm các vướng mắc trên và bảo đảm chất lượng nghiệm thu hệ thống trước mốc thời gian 16h00 ngày 02/12/2026, chuyên viên kiểm thử phối hợp cùng Trưởng ban Nguyễn Văn Sơn đề xuất các phương án hành động cụ thể:\n\n"
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
    p.add_run("Kính trình Ban Tổng Giám đốc và Lãnh đạo Phòng Công nghệ thông tin phê duyệt kế hoạch tổ chức thực hiện các nội dung trọng tâm từ nay đến mốc báo cáo kết thúc thử việc (02/12/2026):")

    # TABLE 1: Phân công thực hiện (STT, Nội dung, Đơn vị thực hiện, Thời hạn)
    tbl_plan = doc.add_table(rows=7, cols=4)
    tbl_plan.alignment = WD_TABLE_ALIGNMENT.CENTER
    pl_headers = ["STT", "Nội dung công việc trọng tâm", "Đơn vị / Đầu mối thực hiện", "Thời hạn hoàn thành"]
    pl_w = [Inches(0.6), Inches(3.3), Inches(2.1), Inches(1.2)]

    for c_idx, h in enumerate(pl_headers):
        cell = tbl_plan.rows[0].cells[c_idx]
        cell.width = pl_w[c_idx]
        cell.text = h
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.runs[0]
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 40, 40, 40, 40)
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
            r.font.size = Pt(9.5)
            if c_idx == 0:
                r.bold = True
            elif c_idx == 3:
                r.bold = True
                r.font.color.rgb = RGBColor(180, 20, 20)
            set_cell_background(cell, "F9FAFB" if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, 35, 35, 35, 35)
            set_cell_border(cell, top=('single', '4', '0', 'D1D5DB'), bottom=('single', '4', '0', 'D1D5DB'), left=('single', '4', '0', 'D1D5DB'), right=('single', '4', '0', 'D1D5DB'))

    # Spacing before Table 2
    p_sp2 = doc.add_paragraph()
    p_sp2.paragraph_format.space_before = Pt(8)
    p_sp2.paragraph_format.space_after = Pt(4)

    # TABLE 2: Ký duyệt các cấp (rows=6, cols=2)
    tbl_sign = doc.add_table(rows=6, cols=2)
    tbl_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    s_w = [Inches(3.2), Inches(4.0)]

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
        r_v.font.size = Pt(10.5)
        
        set_cell_background(c_lbl, "F1F5F9")
        set_cell_background(c_val, "FFFFFF")
        set_cell_margins(c_lbl, 40, 40, 40, 40)
        set_cell_margins(c_val, 40, 40, 40, 40)
        set_cell_border(c_lbl, top=('single', '4', '0', 'CBD5E1'), bottom=('single', '4', '0', 'CBD5E1'), left=('single', '4', '0', 'CBD5E1'), right=('single', '4', '0', 'CBD5E1'))
        set_cell_border(c_val, top=('single', '4', '0', 'CBD5E1'), bottom=('single', '4', '0', 'CBD5E1'), left=('single', '4', '0', 'CBD5E1'), right=('single', '4', '0', 'CBD5E1'))

    # Footer: Nơi nhận / Nơi lưu
    p_nn = doc.add_paragraph()
    p_nn.paragraph_format.space_before = Pt(8)
    p_nn.paragraph_format.space_after = Pt(2)
    r_nn_t = p_nn.add_run("Nơi nhận: ")
    r_nn_t.bold = True
    r_nn_t.font.name = "Times New Roman"
    r_nn_t.font.size = Pt(10.5)
    r_nn_v = p_nn.add_run("Sếp; Ban Tổng giám đốc; Phòng Thanh tra giám sát; Ban Kiểm soát nội bộ.")
    r_nn_v.font.name = "Times New Roman"
    r_nn_v.font.size = Pt(10.5)

    p_nl = doc.add_paragraph()
    p_nl.paragraph_format.space_after = Pt(4)
    r_nl_t = p_nl.add_run("Nơi lưu: ")
    r_nl_t.bold = True
    r_nl_t.font.name = "Times New Roman"
    r_nl_t.font.size = Pt(10.5)
    r_nl_v = p_nl.add_run("Phòng Công nghệ thông tin; Ban Thư ký Tổng hợp.")
    r_nl_v.font.name = "Times New Roman"
    r_nl_v.font.size = Pt(10.5)

    # Save to TARGET_FILE (PCNTT_MauBC.docx) and NAMED_FILE
    doc.save(TARGET_FILE)
    doc.save(NAMED_FILE)
    print("SUCCESS: Filled report into:", TARGET_FILE, f"({os.path.getsize(TARGET_FILE)} bytes)")
    print("SUCCESS: Also saved copy to:", NAMED_FILE, f"({os.path.getsize(NAMED_FILE)} bytes)")

if __name__ == "__main__":
    create_full_report()
