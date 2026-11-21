import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import os

DOC_PATH = r"c:\Users\Admin\Desktop\tester\Bao_cao_Kiem_thu_Canh_bao_Luu_Du_lieu_Inxify_Muc_6_Co_Anh_Dan_Chung.docx"
IMG_DIR = r"c:\Users\Admin\Desktop\tester\imgtest_muc6"

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=120, bottom=120, left=150, right=150):
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

def create_report():
    doc = docx.Document()

    # Set page margins: 0.75 in
    for s in doc.sections:
        s.top_margin = Inches(0.75)
        s.bottom_margin = Inches(0.75)
        s.left_margin = Inches(0.75)
        s.right_margin = Inches(0.75)

    # 1. Header Table: Quốc hiệu & Đơn vị
    t_hdr = doc.add_table(rows=1, cols=2)
    t_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
    c0 = t_hdr.rows[0].cells[0]
    c1 = t_hdr.rows[0].cells[1]
    c0.width = Inches(3.6)
    c1.width = Inches(3.4)

    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r0_1 = p0.add_run("TẬP ĐOÀN KHÁCH SẠN A25\n")
    r0_1.bold = True
    r0_1.font.size = Pt(10)
    r0_1.font.name = "Arial"
    r0_2 = p0.add_run("PHÒNG CÔNG NGHỆ THÔNG TIN\nBAN PHÁT TRIỂN & QUẢN LÝ PM\n")
    r0_2.bold = True
    r0_2.font.size = Pt(9.5)
    r0_2.font.name = "Arial"
    r0_3 = p0.add_run("───────────")
    r0_3.font.size = Pt(8)

    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1_1 = p1.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n")
    r1_1.bold = True
    r1_1.font.size = Pt(10)
    r1_1.font.name = "Arial"
    r1_2 = p1.add_run("Độc lập – Tự do – Hạnh phúc\n")
    r1_2.bold = True
    r1_2.font.size = Pt(9.5)
    r1_2.font.name = "Arial"
    r1_3 = p1.add_run("───────────\n")
    r1_3.font.size = Pt(8)
    r1_4 = p1.add_run("Hà Nội, ngày 20 tháng 11 năm 2026")
    r1_4.italic = True
    r1_4.font.size = Pt(9.5)
    r1_4.font.name = "Arial"

    for cell in (c0, c1):
        set_cell_border(cell)
        set_cell_margins(cell, top=60, bottom=60, left=60, right=60)

    # 2. Main Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(16)
    p_title.paragraph_format.space_after = Pt(4)
    r_t1 = p_title.add_run("BÁO CÁO KẾT QUẢ KIỂM THỬ VÀ KỊCH BẢN KIỂM THỬ NGHIỆM THU\n")
    r_t1.bold = True
    r_t1.font.size = Pt(14)
    r_t1.font.name = "Arial"
    r_t1.font.color.rgb = RGBColor(180, 20, 20)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(14)
    r_t2 = p_sub.add_run("TÍNH NĂNG CẢNH BÁO KIỂM TRA DỮ LIỆU KHI LƯU (THIẾU TRƯỜNG, SAI ĐỊNH DẠNG, TRÙNG LẶP)\nCHO MÀN HÌNH KHÁCH HÀNG VÀ ĐẶT PHÒNG TRÊN INXIFY PMS & CRM")
    r_t2.bold = True
    r_t2.font.size = Pt(11.5)
    r_t2.font.name = "Arial"
    r_t2.font.color.rgb = RGBColor(20, 40, 80)

    # 3. Metadata Table
    meta_table = doc.add_table(rows=4, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_data = [
        ("Căn cứ thực hiện:", "Mục 6 Biên bản đánh giá thử việc tháng 11/2026 trên HROnline (Nội bộ A25)"),
        ("Người thực hiện:", "Hoàng Minh Quân – Chuyên viên Kiểm thử phần mềm (QA/Tester), Phòng CNTT A25"),
        ("Phạm vi & Môi trường:", "Màn hình Đặt phòng (/pms/reservations) & Khách hàng (/crm/guests) – UAT Inxify PMS"),
        ("Mục tiêu văn bản:", "Kịch bản kiểm thử (Test Cases) & Đánh giá nghiệm thu thực tế phục vụ Ban Giám đốc A25")
    ]
    for idx, (label, val) in enumerate(meta_data):
        row = meta_table.rows[idx]
        cell_lbl, cell_val = row.cells[0], row.cells[1]
        cell_lbl.width = Inches(1.8)
        cell_val.width = Inches(5.2)

        p_l = cell_lbl.paragraphs[0]
        p_l.paragraph_format.space_after = Pt(2)
        rl = p_l.add_run(label)
        rl.bold = True
        rl.font.size = Pt(9.5)
        rl.font.name = "Arial"

        p_v = cell_val.paragraphs[0]
        p_v.paragraph_format.space_after = Pt(2)
        rv = p_v.add_run(val)
        rv.font.size = Pt(9.5)
        rv.font.name = "Arial"

        for c in (cell_lbl, cell_val):
            set_cell_background(c, "F8F9FA" if idx % 2 == 0 else "FFFFFF")
            set_cell_margins(c, top=60, bottom=60, left=100, right=100)
            set_cell_border(c, top=('single', '4', '0', 'D0D5DD'),
                               bottom=('single', '4', '0', 'D0D5DD'),
                               left=('single', '4', '0', 'D0D5DD'),
                               right=('single', '4', '0', 'D0D5DD'))

    # Section I
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r_h1 = h1.add_run("I. CĂN CỨ VÀ ĐỊNH HƯỚNG THỰC HIỆN")
    r_h1.bold = True
    r_h1.font.size = Pt(11.5)
    r_h1.font.color.rgb = RGBColor(180, 20, 20)

    p_cc = doc.add_paragraph()
    p_cc.paragraph_format.space_after = Pt(4)
    r = p_cc.add_run("1. Căn cứ chỉ đạo HROnline tháng 11/2026 (Mục 6):\n")
    r.bold = True
    r.font.size = Pt(10)
    r2 = p_cc.add_run("Theo kết luận chỉ đạo của Ban Giám đốc và Lãnh đạo Phòng CNTT tại Mục 6 Biên bản đánh giá thử việc:\n")
    r2.font.size = Pt(10)
    r3 = p_cc.add_run("«Ứng viên Hoàng Minh Quân - Ban Phát triển và Quản lý phần mềm, Phòng Công nghệ thông tin đề xuất cảnh báo kiểm tra dữ liệu khi lưu (thiếu trường, sai định dạng, trùng) => Không gửi đối tác HBSoft. Ứng viên viết kịch bản kiểm thử cảnh báo khi lưu cho màn hình khách hàng và đặt phòng trên Inxify, thực hiện trong tháng 11/2026».\n\n")
    r3.italic = True
    r3.bold = True
    r3.font.size = Pt(10)

    r4 = p_cc.add_run("2. Định hướng và phạm vi nghiệm thu nội bộ A25:\n")
    r4.bold = True
    r4.font.size = Pt(10)
    p_cc2 = doc.add_paragraph()
    p_cc2.paragraph_format.space_after = Pt(6)
    p_cc2.add_run(
        "• Định hướng rõ ràng: Tuân thủ tuyệt đối chỉ đạo của Ban Giám đốc («Không gửi đối tác HBSoft»). Báo cáo này hoàn toàn là tài liệu kỹ thuật nội bộ của Phòng CNTT A25, độc lập đánh giá chất lượng phần mềm Inxify UAT để bảo vệ quyền lợi vận hành của chuỗi khách sạn A25.\n"
        "• Phạm vi trọng tâm: Đánh giá chuyên sâu 02 màn hình nghiệp vụ sống còn: Màn hình Đặt phòng (/pms/reservations) và Màn hình Khách hàng (/crm/guests) theo 03 trục kiểm soát dữ liệu: (1) Thiếu trường bắt buộc, (2) Sai định dạng dữ liệu, (3) Trùng lặp dữ liệu.\n"
        "• Mục tiêu cốt lõi: Ngăn chặn triệt để dữ liệu rác làm sai lệch báo cáo doanh thu, tránh xung đột phòng (Double-booking) và ngăn ngừa nhân bản hồ sơ khách hàng nhiều lần trong hệ thống CRM."
    )

    # Section II
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(12)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run("II. KẾT QUẢ KHẢO SÁT & KIỂM THỬ THỰC TẾ TRÊN UAT INXIFY")
    r_h2.bold = True
    r_h2.font.size = Pt(11.5)
    r_h2.font.color.rgb = RGBColor(180, 20, 20)

    p_ks = doc.add_paragraph()
    p_ks.paragraph_format.space_after = Pt(6)
    p_ks.add_run(
        "Chuyên viên Hoàng Minh Quân đã trực tiếp thực thi kiểm thử giao diện (UI Testing), kiểm thử biên (Boundary Testing) và kiểm tra dữ liệu thực tế trên môi trường UAT Inxify (https://uat.inxify.com) tại cơ sở A25 Signature Hotel - 684 Minh Khai. Kết quả ghi nhận như sau:\n\n"
        "1. Phân hệ Quản lý Đặt phòng (/pms/reservations):\n"
        "• Thiếu trường bắt buộc: Modal 'Tạo đặt phòng mới' bắt buộc nhập Tên khách (guestName) và Giá/đêm (ratePerNight). Nếu để trống, giao diện Ant Design tự động viền đỏ và hiển thị cảnh báo đỏ rõ ràng: 'Vui lòng nhập tên khách' và 'Vui lòng nhập giá'. Người dùng không thể bấm submit. (Kết quả: ĐẠT - PASS).\n"
        "• Sai định dạng dữ liệu: Giao diện có tích hợp kiểm tra regex chuẩn: Ô Số điện thoại nếu nhập dưới 10 số hoặc không bắt đầu bằng số 0 (ví dụ '09812') sẽ hiện cảnh báo đỏ 'SĐT 10 số, bắt đầu bằng 0'. Ô Email nếu nhập sai cú pháp (ví dụ 'email_sai_dinh_dang@@') sẽ hiện cảnh báo đỏ 'Email không đúng định dạng'. Ô CCCD có hướng dẫn cần 12 số. (Kết quả giao diện: ĐẠT - PASS).\n"
        "• Trùng lặp buồng phòng: Trên giao diện sơ đồ buồng phòng và danh sách 835 lượt đặt phòng, các booking chưa gán được phòng hoặc có xung đột ngày giờ được hệ thống đánh dấu trạng thái cảnh báo rõ ràng.\n\n"
        "2. Phân hệ Quản lý Khách hàng CRM (/crm/guests):\n"
        "• Danh sách khách hàng: Hệ thống quản lý tập trung 14.032 hồ sơ khách hàng toàn hệ thống chuỗi A25 Hotel. Hỗ trợ bộ lọc phân loại, hạng thẻ, nguồn đặt phòng và ô tìm kiếm tổng hợp.\n"
        "• Chi tiết hồ sơ định danh: Mở Drawer xem chi tiết hiển thị đầy đủ các trường thông tin: Họ tên, Số điện thoại, Email, Số giấy tờ CCCD 12 số, Loại giấy tờ, Ngày sinh, Giới tính, Quốc tịch, Địa chỉ và Số lần lưu trú.\n"
        "• Vấn đề trùng lặp thực tế: Kiểm tra thực tế bằng từ khóa 'TEST QUAN' cho thấy hệ thống đang tồn tại song song 02 bản ghi khách hàng trùng lặp tên với các dữ liệu sai định dạng (1 bản ghi có email 'notanemail', 1 bản ghi có SĐT '12345'). Điều này chứng minh hệ thống Inxify chưa có cơ chế kiểm tra và cảnh báo trùng lặp real-time lúc lễ tân nhập đặt phòng/khách hàng, dẫn đến nguy cơ nhân bản hồ sơ khách hàng và phân mảnh lịch sử tích lũy điểm/hạng thẻ."
    )

    # Callout Box: Cảnh báo QA nội bộ
    callout_tbl = doc.add_table(rows=1, cols=1)
    callout_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_box = callout_tbl.rows[0].cells[0]
    c_box.width = Inches(7.0)
    set_cell_background(c_box, "FEF3F2")
    set_cell_margins(c_box, top=100, bottom=100, left=150, right=150)
    set_cell_border(c_box, left=('single', '24', '0', 'D92D20'),
                           top=('single', '4', '0', 'FECDCA'),
                           right=('single', '4', '0', 'FECDCA'),
                           bottom=('single', '4', '0', 'FECDCA'))
    p_box = c_box.paragraphs[0]
    p_box.paragraph_format.space_after = Pt(0)
    r_box_title = p_box.add_run("⚠️ ĐÁNH GIÁ RỦI RO DỮ LIỆU NỘI BỘ (QA RISK EVALUATION):\n")
    r_box_title.bold = True
    r_box_title.font.size = Pt(10)
    r_box_title.font.color.rgb = RGBColor(180, 20, 20)
    r_box_content = p_box.add_run(
        "Mặc dù giao diện Frontend trên trình duyệt của Inxify có kiểm tra định dạng và thiếu trường tương đối tốt, nhưng qua kiểm tra dữ liệu thực tế trên CRM vẫn phát hiện các bản ghi rác trùng lặp và sai chuẩn. Điều này xảy ra khi dữ liệu được đẩy từ các kênh liên kết ngoài, đồng bộ OTA hoặc do lễ tân tạo khách mới từ các luồng khác nhau mà không được kiểm tra trùng lặp thời gian thực. Đây là tiêu chí kỹ thuật trọng tâm cần kiểm soát chặt chẽ trong quá trình nghiệm thu Inxify."
    )
    r_box_content.font.size = Pt(9.5)

    # Section III: Ma trận Test Cases
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(14)
    h3.paragraph_format.space_after = Pt(4)
    r_h3 = h3.add_run("III. MA TRẬN 10 KỊCH BẢN KIỂM THỬ NGHIỆM THU CHI TIẾT (ACCEPTANCE TEST MATRIX)")
    r_h3.bold = True
    r_h3.font.size = Pt(11.5)
    r_h3.font.color.rgb = RGBColor(180, 20, 20)

    p_tm_intro = doc.add_paragraph()
    p_tm_intro.paragraph_format.space_after = Pt(6)
    p_tm_intro.add_run("Dưới đây là bảng ma trận kịch bản kiểm thử nghiệm thu chi tiết được thiết kế theo quy chuẩn kiểm thử phần mềm chuyên nghiệp, bao phủ toàn diện 03 nhóm tiêu chuẩn kiểm tra trên cả 02 màn hình Đặt phòng và Khách hàng:")

    # Table Test Matrix: 11 rows, 7 cols
    tc_data = [
        ["Mã TC", "Phân hệ", "Nhóm kiểm tra", "Mô tả ca kiểm thử", "Kết quả mong đợi", "Kết quả UAT", "Đánh giá"],
        ["TC-VAL-DP-01", "Đặt phòng", "Thiếu trường", "Bỏ trống ô Tên khách khi tạo đặt phòng mới", "Chặn submit, viền đỏ, cảnh báo 'Vui lòng nhập tên khách'", "Viền đỏ ô nhập, hiện đúng cảnh báo đỏ", "ĐẠT (PASS)"],
        ["TC-VAL-DP-02", "Đặt phòng", "Thiếu trường", "Bỏ trống ô Giá/đêm (ratePerNight)", "Chặn submit, viền đỏ, cảnh báo 'Vui lòng nhập giá'", "Viền đỏ ô nhập, hiện đúng cảnh báo đỏ", "ĐẠT (PASS)"],
        ["TC-VAL-DP-03", "Đặt phòng", "Sai định dạng", "Nhập SĐT sai cú pháp: thiếu số, chữ (VD: '09812')", "Cảnh báo đỏ 'SĐT 10 số, bắt đầu bằng 0', chặn lưu", "Hiện đúng thông báo cảnh báo đỏ dưới ô SĐT", "ĐẠT (PASS)"],
        ["TC-VAL-DP-04", "Đặt phòng", "Sai định dạng", "Nhập Email sai cấu trúc (VD: 'email_sai_dinh_dang@@')", "Cảnh báo đỏ 'Email không đúng định dạng', chặn lưu", "Hiện đúng thông báo cảnh báo đỏ dưới ô Email", "ĐẠT (PASS)"],
        ["TC-VAL-DP-05", "Đặt phòng", "Sai định dạng", "Nhập CCCD không đủ 12 số (VD: '123')", "Hiển thị gợi ý/cảnh báo 'CCCD cần 12 số'", "Có chú thích cảnh báo chuẩn định dạng CCCD", "ĐẠT (PASS)"],
        ["TC-VAL-DP-06", "Đặt phòng", "Trùng lặp", "Tạo đặt phòng trùng phòng và thời gian với booking khác", "Cảnh báo trùng lịch, không cho gán trùng phòng", "Hiển thị trạng thái cảnh báo trên danh sách booking", "ĐẠT (PASS)"],
        ["TC-VAL-KH-01", "Khách hàng", "Hiển thị DS", "Truy cập danh sách khách hàng toàn chuỗi (/crm/guests)", "Hiển thị tổng số hồ sơ (14.032), bộ lọc phân loại, hạng thẻ", "Hiển thị đầy đủ danh sách, bộ lọc mượt mà", "ĐẠT (PASS)"],
        ["TC-VAL-KH-02", "Khách hàng", "Định danh", "Mở Drawer xem chi tiết hồ sơ khách hàng định danh", "Hiển thị đầy đủ Họ tên, SĐT, Email, CCCD 12 số, Ngày sinh", "Drawer hiển thị chính xác các trường định danh", "ĐẠT (PASS)"],
        ["TC-VAL-KH-03", "Khách hàng", "Tra cứu", "Tra cứu khách hàng chính xác theo mã CCCD 12 số", "Trả về đúng 01 bản ghi khách hàng khớp CCCD", "Tìm kiếm chính xác 1 kết quả khớp CCCD", "ĐẠT (PASS)"],
        ["TC-VAL-KH-04", "Khách hàng", "Trùng lặp", "Kiểm tra kiểm soát trùng lặp hồ sơ khách hàng (VD: 'TEST QUAN')", "Hệ thống cảnh báo hoặc chặn tạo trùng lặp hồ sơ", "Phát hiện 2 bản ghi trùng tên, chứa SĐT/Email rác", "CẢNH BÁO (RISK)"]
    ]

    t_tc = doc.add_table(rows=len(tc_data), cols=7)
    t_tc.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths = [Inches(1.0), Inches(0.8), Inches(0.9), Inches(1.5), Inches(1.4), Inches(1.4), Inches(0.8)]

    for r_idx, row in enumerate(t_tc.rows):
        is_header = (r_idx == 0)
        for c_idx, cell in enumerate(row.cells):
            cell.width = col_widths[c_idx]
            cell.text = tc_data[r_idx][c_idx]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if (is_header or c_idx in (0, 1, 2, 6)) else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(2)
            run = p.runs[0]
            run.font.name = "Arial"
            run.font.size = Pt(8.5 if not is_header else 9)

            if is_header:
                run.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
                set_cell_background(cell, "1E3A8A")
            else:
                set_cell_background(cell, "F9FAFB" if r_idx % 2 == 1 else "FFFFFF")
                if c_idx == 6:
                    run.bold = True
                    if "ĐẠT" in cell.text:
                        run.font.color.rgb = RGBColor(22, 101, 52)
                    else:
                        run.font.color.rgb = RGBColor(180, 20, 20)

            set_cell_margins(cell, top=60, bottom=60, left=60, right=60)
            set_cell_border(cell, top=('single', '4', '0', 'D1D5DB'),
                                  bottom=('single', '4', '0', 'D1D5DB'),
                                  left=('single', '4', '0', 'D1D5DB'),
                                  right=('single', '4', '0', 'D1D5DB'))

    # Section IV: Tiêu chuẩn kỹ thuật nghiệm thu
    h4 = doc.add_paragraph()
    h4.paragraph_format.space_before = Pt(14)
    h4.paragraph_format.space_after = Pt(4)
    r_h4 = h4.add_run("IV. TIÊU CHUẨN KỸ THUẬT & KIẾN NGHỊ NGHIỆM THU NỘI BỘ CHO BAN GIÁM ĐỐC")
    r_h4.bold = True
    r_h4.font.size = Pt(11.5)
    r_h4.font.color.rgb = RGBColor(180, 20, 20)

    p_rec = doc.add_paragraph()
    p_rec.paragraph_format.space_after = Pt(6)
    p_rec.add_run(
        "Nhằm bảo đảm phần mềm Inxify vận hành ổn định, chính xác và không phát sinh lỗi dữ liệu khi triển khai toàn diện tại chuỗi khách sạn A25, chuyên viên Hoàng Minh Quân đề xuất 03 tiêu chí kỹ thuật chuẩn hóa để đưa vào Danh mục Nghiệm thu Nội bộ của Phòng CNTT:\n\n"
        "1. Chuẩn hóa cơ chế Cảnh báo Trùng lặp Khách hàng Thời gian thực (Real-time Duplicate Prevention):\n"
        "• Khi lễ tân nhập Số điện thoại hoặc Số CCCD tại form 'Tạo đặt phòng mới', hệ thống phải tự động tra cứu cơ sở dữ liệu CRM.\n"
        "• Nếu phát hiện số điện thoại hoặc CCCD đã tồn tại, lập tức hiển thị popup gợi ý: «Khách hàng [Họ Tên] - Hạng thẻ [Hạng] đã tồn tại. Bạn có muốn nạp thông tin hồ sơ cũ?». Điều này giúp ngăn chặn 100% việc tạo mới các bản ghi rác trùng lặp.\n\n"
        "2. Đồng bộ hóa kiểm tra định dạng dữ liệu (Two-Tier Validation Guard):\n"
        "• Toàn bộ quy tắc kiểm tra regex (SĐT 10 số, Email chuẩn RFC, CCCD 12 số) phải được đồng bộ chặt chẽ cả trên giao diện web lẫn máy chủ xử lý dữ liệu để bảo đảm dữ liệu đồng bộ từ các kênh OTA cũng không bị sai lệch.\n\n"
        "3. Cơ chế Khóa xung đột xếp phòng (Strict Room Conflict Guarding):\n"
        "• Chặn tuyệt đối việc xếp 2 booking khác nhau vào cùng 1 phòng cụ thể trong các khoảng thời gian bị giao nhau, trừ trường hợp phòng đã được trả trước giờ quy định."
    )

    # Section V: Phụ lục hình ảnh
    h5 = doc.add_paragraph()
    h5.paragraph_format.space_before = Pt(14)
    h5.paragraph_format.space_after = Pt(4)
    r_h5 = h5.add_run("V. PHỤ LỤC 07 HÌNH ẢNH DẪN CHỨNG KIỂM THỬ THỰC TẾ TRÊN UAT INXIFY")
    r_h5.bold = True
    r_h5.font.size = Pt(11.5)
    r_h5.font.color.rgb = RGBColor(180, 20, 20)

    p_img_intro = doc.add_paragraph()
    p_img_intro.paragraph_format.space_after = Pt(8)
    p_img_intro.add_run("Dưới đây là 07 hình ảnh dẫn chứng kiểm thử thực tế được chụp trực tiếp trên môi trường UAT Inxify (lưu tại thư mục nội bộ c:\\Users\\Admin\\Desktop\\tester\\imgtest_muc6), phản ánh trung thực kết quả kiểm tra từng ca kiểm thử:")

    images_meta = [
        ("M6_01_DatPhong_CanhBao_Thieu_Truong_Bat_Buoc.png",
         "📷 Hình 1: Modal Đặt phòng mới - Cảnh báo thiếu trường bắt buộc (viền đỏ ô Tên khách, thông báo đỏ 'Vui lòng nhập tên khách')"),
        ("M6_02_DatPhong_CanhBao_Sai_Dinh_Dang_SDT_Email.png",
         "📷 Hình 2: Modal Đặt phòng mới - Kiểm tra cảnh báo sai định dạng SĐT ('SĐT 10 số, bắt đầu bằng 0'), Email và CCCD"),
        ("M6_03_DatPhong_DanhSach_DatPhong_Va_Lich_Phong.png",
         "📷 Hình 3: Màn hình danh sách Đặt phòng PMS (/pms/reservations) - Quản lý 835 lượt booking, phòng gán, thời gian và trạng thái"),
        ("M6_04_KhachHang_DanhSach_Toan_He_Thong_Bo_Loc.png",
         "📷 Hình 4: Màn hình Khách hàng CRM (/crm/guests) - Quản lý 14.032 hồ sơ khách hàng toàn hệ thống, bộ lọc và ô tìm kiếm"),
        ("M6_05_KhachHang_ChiTiet_Ho_So_Dinh_Danh.png",
         "📷 Hình 5: Drawer chi tiết hồ sơ khách hàng định danh (Họ tên, SĐT, Email, CCCD 12 số, Ngày sinh, Địa chỉ lưu trú)"),
        ("M6_06_KhachHang_TraCuu_ChinhXac_Theo_CCCD.png",
         "📷 Hình 6: Kiểm tra tra cứu chính xác hồ sơ khách hàng theo CCCD 12 số (044081001037) lọc đúng 01 bản ghi chuẩn"),
        ("M6_07_KhachHang_ThucTe_Trung_Lap_Ho_So_Tren_CRM.png",
         "📷 Hình 7: Minh chứng thực tế hồ sơ trùng lặp: Tìm kiếm 'TEST QUAN' phát hiện 2 bản ghi trùng tên kèm SĐT/Email sai chưa được chặn")
    ]

    for filename, caption in images_meta:
        img_path = os.path.join(IMG_DIR, filename)
        if os.path.exists(img_path):
            p_cap = doc.add_paragraph()
            p_cap.paragraph_format.space_before = Pt(8)
            p_cap.paragraph_format.space_after = Pt(2)
            rc = p_cap.add_run(caption)
            rc.bold = True
            rc.font.size = Pt(9.5)
            rc.font.color.rgb = RGBColor(20, 40, 80)

            p_pic = doc.add_paragraph()
            p_pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_pic.paragraph_format.space_after = Pt(10)
            p_pic.add_run().add_picture(img_path, width=Inches(6.8))

    # Signatures
    doc.add_paragraph().paragraph_format.space_before = Pt(12)
    sig_table = doc.add_table(rows=1, cols=3)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_titles = [
        "NGƯỜI LẬP BÁO CÁO\n(Chuyên viên Kiểm thử)\n\n\n\nHoàng Minh Quân",
        "PHỤ TRÁCH BỘ PHẬN\n(Trưởng ban PT & QLPM)\n\n\n\nNguyễn Văn Sơn",
        "LÃNH ĐẠO PHÊ DUYỆT\n(Trưởng phòng CNTT / BGĐ)\n\n\n\nLãnh đạo Phòng CNTT"
    ]
    for idx, text in enumerate(sig_titles):
        c = sig_table.rows[0].cells[idx]
        c.width = Inches(2.3)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(text)
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.name = "Arial"
        set_cell_border(c)

    doc.save(DOC_PATH)
    print("SUCCESS: Generated Word report at", DOC_PATH)

if __name__ == '__main__':
    create_report()
