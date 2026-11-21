import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import os

DOC_PATH = r"c:\Users\Admin\Desktop\tester\Tai_lieu_Kich_ban_Kiem_thu_Nghiem_thu_Quet_QR_CCCD_Inxify_Muc_5.docx"

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=80, bottom=80, left=100, right=100):
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
    run.font.name = "Arial"
    if level == 1:
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(180, 20, 20)
    elif level == 2:
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(3)
        run.font.size = Pt(10.5)
        run.font.color.rgb = RGBColor(30, 58, 138)
    return p

def create_cccd_test_doc():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(0.75)
        s.bottom_margin = Inches(0.75)
        s.left_margin = Inches(0.75)
        s.right_margin = Inches(0.75)

    # Header
    t_hdr = doc.add_table(rows=1, cols=2)
    t_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
    c0, c1 = t_hdr.rows[0].cells
    c0.width, c1.width = Inches(3.6), Inches(3.4)

    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r0 = p0.add_run("CÔNG TY CỔ PHẦN A25 HOTEL\nPHÒNG CÔNG NGHỆ THÔNG TIN\nBAN PHÁT TRIỂN & QLPM")
    r0.bold = True
    r0.font.size = Pt(9.5)
    r0.font.name = "Arial"

    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p1.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập – Tự do – Hạnh phúc\n-------------------")
    r1.bold = True
    r1.font.size = Pt(9.5)
    r1.font.name = "Arial"

    for c in (c0, c1):
        set_cell_border(c)
        set_cell_margins(c, 40, 40, 40, 40)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(12)
    p_title.paragraph_format.space_after = Pt(4)
    rt = p_title.add_run("TÀI LIỆU KỊCH BẢN KIỂM THỬ NGHIỆM THU (ACCEPTANCE TEST PLAN)\nTÍNH NĂNG QUÉT MÃ QR CĂN CƯỚC CÔNG DÂN (CCCD) TRÊN INXIFY PMS")
    rt.bold = True
    rt.font.size = Pt(13)
    rt.font.color.rgb = RGBColor(180, 20, 20)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(10)
    rs = p_sub.add_run("(Thực hiện nhiệm vụ Mục 5 theo Kết luận chỉ đạo HROnline tháng 11/2026)")
    rs.italic = True
    rs.font.size = Pt(10)

    # Metadata
    tbl_m = doc.add_table(rows=4, cols=2)
    tbl_m.alignment = WD_TABLE_ALIGNMENT.CENTER
    m_data = [
        ("Căn cứ thực hiện:", "Chỉ đạo Mục 5: 'Ứng viên viết kịch bản kiểm thử nghiệm thu tính năng này trên Inxify, thực hiện trong tháng 11/2026'"),
        ("Người lập kịch bản:", "Hoàng Minh Quân – Chuyên viên kiểm thử phần mềm (Ứng viên thử việc)"),
        ("Người kiểm duyệt:", "Nguyễn Văn Sơn – Trưởng Ban Phát triển & Quản lý phần mềm"),
        ("Phạm vi & Phương pháp:", "Kiểm thử Logic phân tích chuỗi dữ liệu QR (Parser Logic) bằng Dữ liệu giả lập (Mockup Data) khi chưa có thiết bị quét phần cứng 2D và thẻ cứng tại văn phòng")
    ]
    for idx, (lbl, val) in enumerate(m_data):
        c_l, c_v = tbl_m.rows[idx].cells
        c_l.width, c_v.width = Inches(2.0), Inches(5.0)
        c_l.text, c_v.text = lbl, val
        c_l.paragraphs[0].runs[0].bold = True
        set_cell_background(c_l, "F1F5F9")
        set_cell_margins(c_l, 40, 40, 40, 40)
        set_cell_margins(c_v, 40, 40, 40, 40)
        set_cell_border(c_l, top=('single', '4', '0', 'CBD5E1'), bottom=('single', '4', '0', 'CBD5E1'), left=('single', '4', '0', 'CBD5E1'), right=('single', '4', '0', 'CBD5E1'))
        set_cell_border(c_v, top=('single', '4', '0', 'CBD5E1'), bottom=('single', '4', '0', 'CBD5E1'), left=('single', '4', '0', 'CBD5E1'), right=('single', '4', '0', 'CBD5E1'))

    # Section I
    add_heading_styled(doc, "I. BẢN CHẤT KỸ THUẬT & PHƯƠNG PHÁP KIỂM THỬ KHI CHƯA CÓ THIẾT BỊ CỨNG", 1)
    p = doc.add_paragraph()
    p.add_run(
        "1. Bản chất cơ chế quét QR CCCD trên phần mềm khách sạn:\n"
        "• Các đầu đọc mã vạch 2D chuyên dụng tại quầy lễ tân (Honeywell, Zebra, Sunmi kết nối cổng USB) hoạt động theo cơ chế Bàn phím ảo (Keyboard Wedge / HID Device).\n"
        "• Khi quét mã QR trên thẻ CCCD gắn chip của công dân Việt Nam, thiết bị sẽ giải mã hình ảnh QR thành một Chuỗi văn bản thô (Raw String) kết thúc bằng ký tự Enter (CR/LF) và truyền trực tiếp vào ô nhập liệu đang được con trỏ chuột kích hoạt (Focus).\n"
        "• Chuỗi văn bản QR CCCD chuẩn do Bộ Công an ban hành gồm 07 trường thông tin cố định phân tách bởi ký tự gạch đứng (|):\n"
        "   Cấu trúc: [Số CCCD]|[Số CMND cũ]|[Họ và tên]|[Ngày sinh (DDMMYYYY)]|[Giới tính]|[Địa chỉ thường trú]|[Ngày cấp (DDMMYYYY)]\n"
        "   Ví dụ thực tế: 001099012345||Nguyễn Văn An|15081995|Nam|Số 10 Tràng Tiền, Phường Tràng Tiền, Quận Hoàn Kiếm, Hà Nội|25042021\n\n"
        "2. Giải pháp kiểm thử chuyên môn khi chưa được cấp máy quét và thẻ cứng:\n"
        "• Hiện tại ở bàn làm việc thử việc của Ứng viên chưa được trang bị đầu đọc quét 2D cầm tay và không có thẻ CCCD thật của khách hàng (nhằm bảo đảm an toàn thông tin cá nhân theo quy định pháp luật).\n"
        "• Để đảm bảo tiến độ nghiệm thu theo đúng Mục 5, chuyên viên kiểm thử sử dụng phương pháp Chuỗi dữ liệu giả lập (Mock / Simulated QR Stream) nhập hoặc dán trực tiếp vào ô nhận diện quét mã để kiểm thử toàn diện Logic bóc tách (Parser Logic) và tự động điền dữ liệu (Auto-fill) của phần mềm Inxify."
    )

    # Section II
    add_heading_styled(doc, "II. MA TRẬN 06 KỊCH BẢN KIỂM THỬ NGHIỆM THU (ACCEPTANCE TEST MATRIX)", 1)
    tbl_tc = doc.add_table(rows=7, cols=5)
    tbl_tc.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["Mã TC", "Tên kịch bản & Mục tiêu", "Dữ liệu đầu vào giả lập (Mock Input)", "Kết quả kỳ vọng (Expected Result)", "Trạng thái UAT"]
    w_tc = [Inches(1.0), Inches(1.8), Inches(2.2), Inches(1.5), Inches(0.8)]

    for c_idx, h in enumerate(headers):
        cell = tbl_tc.rows[0].cells[c_idx]
        cell.width = w_tc[c_idx]
        cell.text = h
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.runs[0]
        r.bold = True
        r.font.name = "Arial"
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 40, 40, 40, 40)
        set_cell_border(cell, top=('single', '4', '0', 'CBD5E1'), bottom=('single', '4', '0', 'CBD5E1'), left=('single', '4', '0', 'CBD5E1'), right=('single', '4', '0', 'CBD5E1'))

    tc_data = [
        ("TC-QR-01", "Bóc tách chuỗi QR CCCD chuẩn 7 trường hợp lệ",
         "001099012345||Nguyễn Văn An|15081995|Nam|Số 10 Tràng Tiền, Hoàn Kiếm, Hà Nội|25042021",
         "Hệ thống tự động điền đúng 05 trường trên form: Số CCCD=001099012345; Tên=Nguyễn Văn An; Ngày sinh=15/08/1995; Giới tính=Nam; Địa chỉ=Số 10 Tràng Tiền...",
         "Sẵn sàng test (Ready)"),
        ("TC-QR-02", "Hỗ trợ ký tự tiếng Việt có dấu chuẩn Unicode (UTF-8)",
         "038095001122|168452136|Trần Thị Kiều Diễm|02091998|Nữ|Thôn 5, Xã Hòa Bình, TP Thanh Hóa|10122022",
         "Họ tên và Địa chỉ hiển thị đúng tiếng Việt đầy đủ dấu, không bị vỡ font, không biến thành ký tự lạ (?) hay lỗi mã hóa.",
         "Sẵn sàng test (Ready)"),
        ("TC-QR-03", "Xử lý chuỗi QR sai định dạng hoặc mã QR không phải CCCD",
         "https://a25hotel.com/promotion hoặc chuỗi chỉ có 3 trường: ABC|123|XYZ",
         "Hệ thống không điền bừa vào form, hiển thị thông báo lỗi màu vàng: 'Mã QR không đúng định dạng CCCD của Bộ Công an. Vui lòng kiểm tra lại'.",
         "Sẵn sàng test (Ready)"),
        ("TC-QR-04", "Kiểm tra cảnh báo trùng khách khi CCCD quét đã tồn tại",
         "Quét chuỗi có Số CCCD: 001099012345 (Đã có hồ sơ lưu trong CRM A25)",
         "Hệ thống cảnh báo: 'Khách hàng đã có trong hệ thống (Mã KH: KH-001)'. Gợi ý nút 'Cập nhật thông tin khách cũ' thay vì tạo mới trùng rác.",
         "Sẵn sàng test (Ready)"),
        ("TC-QR-05", "Cho phép Lễ tân chỉnh sửa, bổ sung thông tin sau khi quét",
         "Sau khi Auto-fill từ QR, Lễ tân nhập thêm Số điện thoại di động và Email",
         "Cho phép sửa/thêm trường mà không làm mất các trường đã được điền tự động từ QR. Nút 'Lưu' hoạt động chính xác.",
         "Sẵn sàng test (Ready)"),
        ("TC-QR-06", "Cơ chế dự phòng (Fallback) khi không có thiết bị quét",
         "Quầy lễ tân không cắm máy quét hoặc thiết bị quét hỏng",
         "Lễ tân vẫn thao tác gõ tay thủ công bình thường vào form Khách hàng, hệ thống không bắt buộc phải có bước quét QR.",
         "ĐÃ PASS (Hiện tại)")
    ]

    for r_idx, row in enumerate(tc_data, start=1):
        for c_idx, val in enumerate(row):
            cell = tbl_tc.rows[r_idx].cells[c_idx]
            cell.width = w_tc[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx in (0, 4) else WD_ALIGN_PARAGRAPH.LEFT
            r = p.runs[0]
            r.font.name = "Arial"
            r.font.size = Pt(8)
            if c_idx == 0:
                r.bold = True
                r.font.color.rgb = RGBColor(30, 58, 138)
            elif c_idx == 4:
                r.bold = True
                r.font.color.rgb = RGBColor(22, 101, 52) if "ĐÃ PASS" in val else RGBColor(180, 83, 9)

            set_cell_background(cell, "F9FAFB" if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, 35, 35, 35, 35)
            set_cell_border(cell, top=('single', '4', '0', 'D1D5DB'), bottom=('single', '4', '0', 'D1D5DB'), left=('single', '4', '0', 'D1D5DB'), right=('single', '4', '0', 'D1D5DB'))

    # Section III
    add_heading_styled(doc, "III. KẾ HOẠCH NGHIỆM THU PHẦN CỨNG THỰC TẾ KHI DEV HOÀN TẤT TÍNH NĂNG", 1)
    p = doc.add_paragraph()
    p.add_run(
        "Nhằm bảo đảm tính khả thi và nghiệm thu thực tế 100% trước khi triển khai rộng rãi cho chuỗi khách sạn A25, Phòng CNTT đề xuất phương án phối hợp như sau:\n"
        "1. Giai đoạn 1 (Hiện tại - Tháng 11/2026): Chuyên viên QA hoàn thiện kịch bản nghiệm thu và bộ dữ liệu test giả lập, bàn giao cho đội ngũ Lập trình viên (Dev) làm căn cứ code tính năng bắt sự kiện quét mã.\n"
        "2. Giai đoạn 2 (Khi Dev Inxify đưa bản cập nhật lên UAT): Phòng CNTT sẽ mượn 01 đầu đọc mã vạch 2D chuyên dụng từ quầy lễ tân chi nhánh (ví dụ: A25 Hotel 684 Minh Khai) kết nối máy tính văn phòng, đồng thời sử dụng thẻ CCCD công vụ của nội bộ nhân sự CNTT để thực thi quét thực tế trong 01 buổi làm việc.\n"
        "3. Đánh giá tính năng liên kết điện thoại: Đối với đề xuất dùng camera điện thoại quét CCCD, giải pháp này đòi hỏi phát triển module PWA hoặc App Mobile riêng biệt có tích hợp thư viện quét mã (như ZXing / Html5-QRCode). Phòng CNTT kiến nghị ưu tiên triển khai đầu đọc 2D cắm USB tại quầy lễ tân trước vì chi phí thấp, tốc độ quét tức thì (< 0.5 giây) và không phụ thuộc vào pin hay camera điện thoại của nhân viên."
    )

    # Signatures
    p_s = doc.add_paragraph()
    p_s.paragraph_format.space_before = Pt(12)
    tbl_s = doc.add_table(rows=2, cols=2)
    tbl_s.alignment = WD_TABLE_ALIGNMENT.CENTER
    s0, s1 = tbl_s.rows[0].cells
    s0.width, s1.width = Inches(3.5), Inches(3.5)
    s0.text = "NGƯỜI LẬP KỊCH BẢN\n(Chuyên viên kiểm thử phần mềm)"
    s1.text = "TRƯỞNG BAN KIỂM DUYỆT\n(Ban Phát triển & Quản lý phần mềm)"
    s0.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    s1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for cell in (s0, s1):
        cell.paragraphs[0].runs[0].bold = True
        set_cell_border(cell)

    s0_b, s1_b = tbl_s.rows[1].cells
    s0_b.width, s1_b.width = Inches(3.5), Inches(3.5)
    s0_b.text = "\n\n\nHoàng Minh Quân"
    s1_b.text = "\n\n\nNguyễn Văn Sơn"
    s0_b.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    s1_b.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for cell in (s0_b, s1_b):
        cell.paragraphs[0].runs[0].bold = True
        set_cell_border(cell)

    doc.save(DOC_PATH)
    print("SUCCESS: Generated CCCD Acceptance Test Plan at:", DOC_PATH)
    print("File size:", os.path.getsize(DOC_PATH), "bytes")

if __name__ == "__main__":
    create_cccd_test_doc()
