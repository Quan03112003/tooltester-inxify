import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import os

DOC_PATH = r"c:\Users\Admin\Desktop\tester\Bao_cao_Phuong_an_Chuan_hoa_Cong_cu_Microsoft_365_A25_Hotel.docx"
IMG_DIR = r"c:\Users\Admin\Desktop\tester\img_m365_guide"

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
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
        run.font.color.rgb = RGBColor(180, 20, 20) # Deep Crimson
    elif level == 2:
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(3)
        run.font.size = Pt(10.5)
        run.font.color.rgb = RGBColor(30, 58, 138) # Navy Blue
    elif level == 3:
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(50, 50, 50)
    return p

def add_illustrated_figure(doc, img_name, caption_text, width_inch=6.5):
    img_path = os.path.join(IMG_DIR, img_name)
    if not os.path.exists(img_path):
        print(f"Warning: Image file not found: {img_path}")
        return

    # Khung viền bảng cho hình ảnh chuyên nghiệp
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = tbl.rows[0].cells[0]
    c.width = Inches(width_inch + 0.2)
    set_cell_background(c, "F8FAFC")
    set_cell_margins(c, top=80, bottom=80, left=80, right=80)
    set_cell_border(c, top=('single', '6', '0', 'CBD5E1'),
                       bottom=('single', '6', '0', 'CBD5E1'),
                       left=('single', '6', '0', 'CBD5E1'),
                       right=('single', '6', '0', 'CBD5E1'))

    p_img = c.paragraphs[0]
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(2)
    p_img.paragraph_format.space_after = Pt(4)
    run_img = p_img.add_run()
    run_img.add_picture(img_path, width=Inches(width_inch))

    p_cap = c.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(2)
    p_cap.paragraph_format.space_after = Pt(4)
    r_cap = p_cap.add_run(caption_text)
    r_cap.bold = True
    r_cap.italic = True
    r_cap.font.name = "Arial"
    r_cap.font.size = Pt(8.5)
    r_cap.font.color.rgb = RGBColor(71, 85, 105)

    p_space = doc.add_paragraph()
    p_space.paragraph_format.space_before = Pt(2)
    p_space.paragraph_format.space_after = Pt(4)

def add_callout(doc, title, content, border_color="D92D20", bg_color="FEF3F2"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = tbl.rows[0].cells[0]
    c.width = Inches(7.0)
    set_cell_background(c, bg_color)
    set_cell_margins(c, top=100, bottom=100, left=150, right=150)
    set_cell_border(c, left=('single', '24', '0', border_color),
                       top=('single', '4', '0', 'E5E7EB'),
                       right=('single', '4', '0', 'E5E7EB'),
                       bottom=('single', '4', '0', 'E5E7EB'))
    p = c.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    r1 = p.add_run(title + "\n")
    r1.bold = True
    r1.font.size = Pt(10)
    r1.font.name = "Arial"
    r1.font.color.rgb = RGBColor(180, 20, 20) if border_color == "D92D20" else RGBColor(30, 58, 138)
    r2 = p.add_run(content)
    r2.font.size = Pt(9.5)
    r2.font.name = "Arial"
    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_before = Pt(2)
    p_spacer.paragraph_format.space_after = Pt(2)

def generate_full_doc():
    doc = docx.Document()

    # 1. Page Margins (0.75 in)
    for s in doc.sections:
        s.top_margin = Inches(0.75)
        s.bottom_margin = Inches(0.75)
        s.left_margin = Inches(0.75)
        s.right_margin = Inches(0.75)

    # 2. Header Quốc hiệu & Đơn vị
    t_hdr = doc.add_table(rows=1, cols=2)
    t_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
    c0 = t_hdr.rows[0].cells[0]
    c1 = t_hdr.rows[0].cells[1]
    c0.width = Inches(3.6)
    c1.width = Inches(3.4)

    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r0_1 = p0.add_run("CÔNG TY CỔ PHẦN A25 HOTEL\n")
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
        set_cell_margins(cell, top=40, bottom=40, left=40, right=40)

    # 3. Main Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(16)
    p_title.paragraph_format.space_after = Pt(4)
    r_t1 = p_title.add_run("BÁO CÁO & PHƯƠNG ÁN TRIỂN KHAI\n")
    r_t1.bold = True
    r_t1.font.size = Pt(13.5)
    r_t1.font.name = "Arial"
    r_t1.font.color.rgb = RGBColor(180, 20, 20)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(12)
    r_t2 = p_sub.add_run("CHUẨN HÓA CÔNG CỤ LÀM VIỆC TRÊN NỀN TẢNG MICROSOFT 365\n(OUTLOOK, TEAMS, ONEDRIVE) TẠI CÔNG TY CỔ PHẦN A25 HOTEL")
    r_t2.bold = True
    r_t2.font.size = Pt(11)
    r_t2.font.name = "Arial"
    r_t2.font.color.rgb = RGBColor(20, 40, 80)

    # 4. Kính gửi
    p_kg = doc.add_paragraph()
    p_kg.paragraph_format.space_after = Pt(8)
    p_kg.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_kg = p_kg.add_run("Kính gửi: BAN GIÁM ĐỐC CÔNG TY CỔ PHẦN A25 HOTEL\nĐồng kính gửi: LÃNH ĐẠO PHÒNG CNTT, PHÒNG KINH DOANH, PHÒNG TỔ CHỨC NHÂN SỰ")
    r_kg.bold = True
    r_kg.font.size = Pt(9.5)
    r_kg.font.name = "Arial"
    r_kg.font.color.rgb = RGBColor(55, 65, 81)

    # Metadata Table
    meta_table = doc.add_table(rows=4, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_data = [
        ("Căn cứ thực hiện:", "Chỉ đạo của Giám đốc Kinh doanh & Đề xuất của Trợ lý Phòng CNTT (Ngày 18/11/2026)"),
        ("Người lập phương án:", "Hoàng Minh Quân – Trợ lý Phòng CNTT, Ban Phát triển & Quản lý PM"),
        ("Người đồng trình duyệt:", "Nguyễn Văn Sơn – Trưởng ban Phát triển & Quản lý phần mềm, Phòng CNTT"),
        ("Thời hạn trình duyệt:", "Trước ngày 25 tháng 11 năm 2026")
    ]
    for idx, (label, val) in enumerate(meta_data):
        row = meta_table.rows[idx]
        cell_lbl, cell_val = row.cells[0], row.cells[1]
        cell_lbl.width = Inches(2.0)
        cell_val.width = Inches(5.0)

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
            set_cell_margins(c, top=45, bottom=45, left=70, right=70)
            set_cell_border(c, top=('single', '4', '0', 'D1D5DB'),
                               bottom=('single', '4', '0', 'D1D5DB'),
                               left=('single', '4', '0', 'D1D5DB'),
                               right=('single', '4', '0', 'D1D5DB'))

    # =========================================================================
    # SECTION I: BỐI CẢNH & SỰ CẦN THIẾT
    # =========================================================================
    add_heading_styled(doc, "I. BỐI CẢNH THỰC TẾ, ĐỀ XUẤT CHUYÊN MÔN & CHỈ ĐẠO CỦA LÃNH ĐẠO", 1)

    add_heading_styled(doc, "1. Thực trạng vận hành và các rủi ro từ công cụ trao đổi hiện tại (Zalo cá nhân)", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "Hiện nay tại Công ty Cổ phần A25 Hotel, phần lớn các trao đổi công việc hằng ngày giữa các phòng ban khối Văn phòng cũng như giữa Văn phòng với hệ thống chi nhánh khách sạn đang phụ thuộc chủ yếu vào các nhóm chat Zalo cá nhân. Qua khảo sát thực tế của Phòng CNTT, phương thức này đang bộc lộ những hạn chế và nguy cơ rủi ro rất lớn đối với doanh nghiệp:\n"
        "• Dữ liệu phân tán, nguy cơ mất mát thông tin: Các file tài liệu, hợp đồng, báo giá, hình ảnh gửi qua Zalo bị tự động hết hạn (expired) sau một thời gian ngắn, gây khó khăn cho việc tra cứu lịch sử khi cần đối chiếu công nợ hoặc kiểm toán vận hành.\n"
        "• Nguy cơ rò rỉ dữ liệu mật và bí mật kinh doanh: Khi nhân viên kinh doanh, lễ tân hoặc cán bộ văn phòng nghỉ việc, công ty không có cơ chế thu hồi tức thì dữ liệu trao đổi trong máy cá nhân của nhân sự. Nhiều trường hợp nhân sự đã nghỉ việc nhưng vẫn còn nằm trong các nhóm chat nội bộ, tiềm ẩn nguy cơ lộ lọt danh sách khách hàng VIP, chính sách giá và thông tin tài chính.\n"
        "• Thiếu tính chuyên nghiệp và tính định danh tập đoàn: Việc giao dịch với đối tác, khách hàng doanh nghiệp bằng tài khoản cá nhân làm giảm tính nhận diện thương hiệu của chuỗi khách sạn A25 Hotel."
    )

    # INSERT FIGURE 1: SO SÁNH ZALO VS MICROSOFT 365
    add_illustrated_figure(doc,
        "M365_01_SoSanh_Zalo_vs_M365.png",
        "Hình 1: Infographic So sánh luồng làm việc Cũ (Zalo cá nhân) và Mới (Microsoft 365 Enterprise) tại A25 Hotel",
        width_inch=6.4)

    add_heading_styled(doc, "2. Ý kiến chỉ đạo của Giám đốc Kinh doanh & Đề xuất của Trợ lý Phòng CNTT", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "Tại cuộc họp ngày 18/11/2026, Giám đốc Kinh doanh đã ghi nhận đề xuất chuyên môn của Trợ lý Phòng CNTT Hoàng Minh Quân và thống nhất chỉ đạo:\n"
        "• Đồng ý chủ trương chuẩn hóa toàn diện công cụ làm việc trên hệ sinh thái Microsoft 365 (Outlook, Teams, OneDrive) nhằm thay thế dần việc trao đổi phân tán qua Zalo cá nhân.\n"
        "• Giao Phòng CNTT (Trợ lý Quân) phối hợp cùng Trưởng ban Nguyễn Văn Sơn lập Phương án triển khai chi tiết trình Ban Lãnh đạo trước ngày 25/11/2026, làm rõ: Đối tượng cấp tài khoản, dự toán chi phí bản quyền, lộ trình chuyển đổi và quy định ranh giới sử dụng công cụ.\n"
        "• Nguyên tắc triển khai: Thí điểm trước tại Khối Văn phòng để đánh giá hiệu quả thực tế, sau đó mới xây dựng phương án mở rộng xuống các chi nhánh khách sạn, tuyệt đối không làm gián đoạn trao đổi vận hành đang diễn ra."
    )

    add_heading_styled(doc, "3. Thiết lập quy trình quản trị vòng đời tài khoản tập trung (Account Lifecycle)", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    p.add_run(
        "Nền tảng Microsoft 365 cho phép Phòng CNTT A25 Hotel thiết lập quy trình quản trị tập trung khép kín theo chuẩn doanh nghiệp hiện đại:\n"
        "• Cấp phát tài khoản chuẩn hóa: Tự động cấp email và tài khoản Teams theo cấu trúc định danh duy nhất (tên.họ@a25hotel.com) ngay khi nhân sự ký hợp đồng thử việc.\n"
        "• Phân quyền truy cập tài liệu: Phân quyền thư mục OneDrive/SharePoint theo đúng phòng ban, vị trí công tác.\n"
        "• Thu hồi và bàn giao dữ liệu tức thì: Khi nhân viên nghỉ việc hoặc điều chuyển, Phòng CNTT chỉ cần 01 thao tác khóa tài khoản trên cổng quản trị trung tâm, thu hồi toàn bộ quyền truy cập trên điện thoại/máy tính cá nhân, đồng thời chuyển giao toàn bộ dữ liệu OneDrive và hòm thư cho người kế nhiệm, đảm bảo an toàn tuyệt đối tài sản số của A25."
    )

    # =========================================================================
    # SECTION II: GIẢI PHÁP DANH BẠ TẬP TRUNG & PROFILE CARD
    # =========================================================================
    add_heading_styled(doc, "II. GIẢI PHÁP DANH BẠ TẬP TRUNG & TRA CỨU NHÂN SỰ THEO MÔ HÌNH TẬP ĐOÀN LỚN", 1)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "Để khắc phục triệt để tình trạng nhân viên không biết liên hệ ai, không rõ chức vụ và cấp quản lý trực tiếp của các phòng ban khác, phương án đề xuất triển khai tính năng Danh bạ doanh nghiệp tập trung trên nền tảng Microsoft Entra ID (Azure AD) và Sổ địa chỉ toàn cầu (Global Address List - GAL) đồng bộ tự động 100% giữa Outlook và Teams."
    )

    add_heading_styled(doc, "1. Thẻ hồ sơ thông minh (Profile Card) & Sơ đồ tổ chức động (Organization Chart)", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "Khi bất kỳ nhân sự nào gõ tên hoặc email đồng nghiệp trên Outlook hoặc Teams, hệ thống sẽ tự động hiển thị Thẻ hồ sơ trực quan gồm:\n"
        "• Thông tin định danh: Ảnh đại diện chuẩn, Họ và tên đầy đủ, Mã nhân viên A25, Email chính thức (@a25hotel.com), Số điện thoại nội bộ/di động công vụ.\n"
        "• Vị trí công tác: Chức danh chuyên môn (Job Title/Role), Phòng ban trực thuộc, Chi nhánh hoặc Khối làm việc (Văn phòng Tập đoàn / Chi nhánh cụ thể).\n"
        "• Sơ đồ phân cấp quản trị (Organization Chart): Thể hiện trực quan cấp Quản lý trực tiếp (Manager) mà nhân sự đó báo cáo, cùng danh sách các nhân sự cấp dưới (Direct Reports) do nhân sự đó quản lý. Điều này giúp nhân viên mới hoặc các phòng ban phối hợp dễ dàng nắm bắt đúng đầu mối phụ trách mà không cần hỏi qua trung gian."
    )

    # INSERT FIGURE 2: PROFILE CARD & ORG CHART
    add_illustrated_figure(doc,
        "M365_03_ProfileCard_OrgChart.png",
        "Hình 2: Giao diện Thẻ hồ sơ nhân viên (Profile Card) và Sơ đồ tổ chức phân cấp quản trị trên Microsoft Teams & Outlook",
        width_inch=6.4)

    add_heading_styled(doc, "2. Quy trình chuẩn hóa dữ liệu đầu vào giữa Phòng Tổ chức Nhân sự (HR) và Phòng CNTT (IT)", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    p.add_run(
        "Để cây thư mục và danh bạ nhân sự luôn chính xác theo thời gian thực, quy định bắt buộc luồng bàn giao thông tin:\n"
        "• Khi có nhân sự mới: Phòng TCNS gửi phiếu yêu cầu tạo tài khoản cho IT tối thiểu 24 giờ trước ngày nhân sự nhận việc, trong đó bắt buộc có đủ 03 trường cốt lõi: (1) Chức danh chính thức, (2) Phòng ban/Đơn vị công tác, và (3) Họ tên cán bộ Quản lý trực tiếp.\n"
        "• Khi có biến động nhân sự (điều chuyển, thăng chức, nghỉ việc): Phòng TCNS cập nhật đồng thời cho Phòng CNTT để điều chỉnh quyền truy cập và sơ đồ tổ chức tương ứng."
    )

    # =========================================================================
    # SECTION III: NGUYÊN TẮC RANH GIỚI SỬ DỤNG 4 CÔNG CỤ
    # =========================================================================
    add_heading_styled(doc, "III. NGUYÊN TẮC VÀ RANH GIỚI SỬ DỤNG CÔNG CỤ (QUY ĐỊNH BẮT BUỘC)", 1)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "Một trong những nguyên nhân khiến việc chuyển đổi số thất bại là việc nhân viên bị 'loạn công cụ', không biết nội dung nào gửi ở đâu. Do đó, phương án đề xuất quy chuẩn ranh giới rõ ràng cho 04 công cụ:"
    )

    # Bảng phân định ranh giới công cụ
    tbl_tools = doc.add_table(rows=5, cols=4)
    tbl_tools.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_headers = ["Tên công cụ", "Mục đích sử dụng chính", "Nội dung quy định sử dụng", "Hành vi nghiêm cấm"]
    col_w = [Inches(1.2), Inches(1.8), Inches(2.3), Inches(1.7)]

    for c_idx, h_text in enumerate(t_headers):
        cell = tbl_tools.rows[0].cells[c_idx]
        cell.width = col_w[c_idx]
        cell.text = h_text
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.runs[0]
        r.bold = True
        r.font.name = "Arial"
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, top=60, bottom=60, left=60, right=60)
        set_cell_border(cell, top=('single', '4', '0', 'D1D5DB'),
                              bottom=('single', '4', '0', 'D1D5DB'),
                              left=('single', '4', '0', 'D1D5DB'),
                              right=('single', '4', '0', 'D1D5DB'))

    tools_rows = [
        ("Microsoft Teams",
         "Trao đổi nội bộ nhanh, phối hợp dự án & Họp trực tuyến",
         "• Chat 1-1 và chat nhóm theo phòng ban, nhóm dự án.\n• Họp trực tuyến (Meeting), chia sẻ màn hình, ghi biên bản cuộc họp.\n• Trao đổi xử lý sự vụ hằng ngày thay thế các nhóm chat Zalo.",
         "• Không dùng gửi văn bản pháp lý/hợp đồng chính thức.\n• Không giao việc chính thức thay thế HROnline."),
        ("Microsoft Outlook",
         "Thư điện tử giao dịch chính thức & Quản lý lịch công vụ",
         "• Thư từ giao dịch với đối tác, khách hàng, nhà cung cấp.\n• Các thông báo chỉ đạo chính thức từ Ban Lãnh đạo.\n• Gửi thư mời họp công vụ, quản lý lịch làm việc của Lãnh đạo.",
         "• Không dùng gửi thư trao đổi ngắn vụn vặt gây spam hòm thư.\n• Không gửi file đính kèm dung lượng quá lớn (dùng link OneDrive)."),
        ("OneDrive for Business",
         "Kho lưu trữ đám mây & Chia sẻ tài liệu công việc duy nhất",
         "• Lưu trữ toàn bộ file tài liệu, báo cáo, dữ liệu làm việc của cá nhân (dung lượng 1TB/user).\n• Chia sẻ link tài liệu làm việc nội bộ phân quyền Xem/Sửa an toàn.",
         "• NGHIÊM CẤM lưu tài liệu công việc trên máy tính cá nhân hoặc Google Drive/OneDrive cá nhân."),
        ("HROnline (Hệ thống hiện hữu)",
         "KÊNH DUY NHẤT để Giao việc, Theo dõi Deadline & Phê duyệt",
         "• Giao việc từ cấp quản lý xuống nhân viên, theo dõi tiến độ và thời hạn hoàn thành (Deadline).\n• Trình ký duyệt hành chính: Nghỉ phép, thanh toán, phiếu yêu cầu phần mềm, công văn nội bộ.",
         "• TUYỆT ĐỐI KHÔNG xây dựng form giao việc/xử lý yêu cầu trên Microsoft Forms hay Power Automate.")
    ]

    for r_idx, row_data in enumerate(tools_rows, start=1):
        for c_idx, val in enumerate(row_data):
            cell = tbl_tools.rows[r_idx].cells[c_idx]
            cell.width = col_w[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(2)
            r = p.runs[0]
            r.font.name = "Arial"
            r.font.size = Pt(8.5)
            if c_idx == 0:
                r.bold = True
                r.font.color.rgb = RGBColor(30, 58, 138)
            set_cell_background(cell, "F9FAFB" if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, top=60, bottom=60, left=60, right=60)
            set_cell_border(cell, top=('single', '4', '0', 'D1D5DB'),
                                  bottom=('single', '4', '0', 'D1D5DB'),
                                  left=('single', '4', '0', 'D1D5DB'),
                                  right=('single', '4', '0', 'D1D5DB'))

    # INSERT FIGURE 3: MA TRẬN 4 CÔNG CỤ
    add_illustrated_figure(doc,
        "M365_02_PhanDinh_4_CongCu.png",
        "Hình 3: Ma trận phân định ranh giới 4 công cụ (Teams, Outlook, OneDrive, HROnline) và Quy định cấm chồng chéo",
        width_inch=6.4)

    # Callout cảnh báo ranh giới HROnline
    add_callout(doc,
        "⚠️ QUY ĐỊNH BẮT BUỘC VỀ RANH GIỚI VỚI HỆ THỐNG HRONLINE:",
        "HROnline tiếp tục là nền tảng quản trị công việc và nhân sự cốt lõi duy nhất của A25 Hotel. Việc triển khai Microsoft 365 chỉ nhằm chuẩn hóa kênh liên lạc (Teams), email giao dịch (Outlook) và lưu trữ dữ liệu (OneDrive). Đội ngũ CNTT cam kết TUYỆT ĐỐI KHÔNG phát triển các quy trình giao việc, phê duyệt hay gửi form yêu cầu trên Microsoft Forms hoặc Power Automate để tránh xung đột, gây phân tán kênh điều hành và làm suy yếu vai trò kiểm soát của HROnline.",
        border_color="D92D20", bg_color="FEF3F2")

    # =========================================================================
    # SECTION IV: ĐỐI TƯỢNG CẤP TÀI KHOẢN & KHUNG CHI PHÍ
    # =========================================================================
    add_heading_styled(doc, "IV. ĐỐI TƯỢNG CẤP TÀI KHOẢN & KHUNG CHI PHÍ BẢN QUYỀN THAM KHẢO", 1)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "Để tối ưu hóa ngân sách và tránh lãng phí, Phòng CNTT đề xuất phân nhóm người dùng tại Khối Văn phòng A25 Hotel theo đúng nhu cầu sử dụng thực tế (hiện chưa chốt ngân sách cố định, bảng dưới đây mang tính chất khung chi phí tham khảo thị trường để Ban Lãnh đạo xem xét hạn mức cấp ngân sách):"
    )

    add_heading_styled(doc, "1. Phân nhóm đối tượng người dùng tại Khối Văn phòng A25 Hotel", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "• Nhóm 1 (Gói Microsoft 365 Business Standard): Dành cho Ban Lãnh đạo, Trưởng/Phó phòng ban, Chuyên viên nghiệp vụ chuyên sâu (Kế toán, Kinh doanh, CNTT, Pháp chế...). Đối tượng này xử lý file bảng tính phức tạp, hợp đồng khối lượng lớn nên cần bộ ứng dụng Office Desktop cài trực tiếp vào máy tính (Word, Excel, PowerPoint, Outlook Desktop) kết hợp Teams và OneDrive.\n"
        "• Nhóm 2 (Gói Microsoft 365 Business Basic): Dành cho Nhân viên tác nghiệp văn phòng, hành chính, lễ tân văn phòng, nhân sự hỗ trợ. Đối tượng này chỉ cần ứng dụng Teams để họp/chat, hòm thư Outlook Web/Mobile để nhận thông báo và 1TB lưu trữ OneDrive; sử dụng bộ Office bản Web trực tuyến mượt mà mà không cần cài bản Desktop nặng máy."
    )

    add_heading_styled(doc, "2. Bảng dự toán khung chi phí bản quyền tham khảo (Quy mô thí điểm Khối Văn phòng)", 2)

    # Bảng chi phí
    tbl_cost = doc.add_table(rows=4, cols=6)
    tbl_cost.alignment = WD_TABLE_ALIGNMENT.CENTER
    cost_headers = ["Nhóm đối tượng", "Gói bản quyền đề xuất", "Phạm vi tính năng", "Số lượng dự kiến", "Đơn giá tham khảo / user / tháng", "Thành tiền dự kiến / tháng (VNĐ)"]
    c_w = [Inches(1.2), Inches(1.3), Inches(1.8), Inches(0.8), Inches(1.0), Inches(1.1)]

    for c_idx, h_text in enumerate(cost_headers):
        cell = tbl_cost.rows[0].cells[c_idx]
        cell.width = c_w[c_idx]
        cell.text = h_text
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.runs[0]
        r.bold = True
        r.font.name = "Arial"
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, top=60, bottom=60, left=50, right=50)
        set_cell_border(cell, top=('single', '4', '0', 'D1D5DB'),
                              bottom=('single', '4', '0', 'D1D5DB'),
                              left=('single', '4', '0', 'D1D5DB'),
                              right=('single', '4', '0', 'D1D5DB'))

    cost_data = [
        ("Nhóm 1: Quản lý & Chuyên viên",
         "M365 Business Standard",
         "Cài đặt Office Desktop (Word, Excel, PPT, Outlook) trên 5 thiết bị/user + Teams + 1TB OneDrive",
         "15 tài khoản",
         "~ 315.000 VNĐ ($12.50)",
         "4.725.000 VNĐ"),
        ("Nhóm 2: Nhân viên tác nghiệp",
         "M365 Business Basic",
         "Office bản Web & Mobile + Hòm thư Outlook Web/Mobile 50GB + Teams + 1TB OneDrive",
         "25 tài khoản",
         "~ 155.000 VNĐ ($6.00)",
         "3.875.000 VNĐ"),
        ("TỔNG CỘNG THÍ ĐIỂM (DỰ KIẾN)",
         "—",
         "Quy mô Khối Văn phòng A25 Hotel (Thí điểm giai đoạn 1)",
         "40 tài khoản",
         "—",
         "8.600.000 VNĐ / tháng\n(~ 103.200.000 đ/năm)")
    ]

    for r_idx, row_data in enumerate(cost_data, start=1):
        for c_idx, val in enumerate(row_data):
            cell = tbl_cost.rows[r_idx].cells[c_idx]
            cell.width = c_w[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx in (0, 1, 3, 4) else (WD_ALIGN_PARAGRAPH.RIGHT if c_idx == 5 else WD_ALIGN_PARAGRAPH.LEFT)
            p.paragraph_format.space_after = Pt(2)
            r = p.runs[0]
            r.font.name = "Arial"
            r.font.size = Pt(8.5)
            if r_idx == 3: # Total row
                r.bold = True
                set_cell_background(cell, "FEF3C7") # Yellow highlight
                if c_idx == 5:
                    r.font.color.rgb = RGBColor(180, 20, 20)
            else:
                set_cell_background(cell, "F9FAFB" if r_idx % 2 == 1 else "FFFFFF")

            set_cell_margins(cell, top=60, bottom=60, left=50, right=50)
            set_cell_border(cell, top=('single', '4', '0', 'D1D5DB'),
                                  bottom=('single', '4', '0', 'D1D5DB'),
                                  left=('single', '4', '0', 'D1D5DB'),
                                  right=('single', '4', '0', 'D1D5DB'))

    p_note = doc.add_paragraph()
    p_note.paragraph_format.space_before = Pt(4)
    p_note.paragraph_format.space_after = Pt(6)
    r_n = p_note.add_run("(*) Ghi chú tài chính: Đơn giá trên được tham khảo từ các đối tác phân phối ủy quyền Microsoft Cloud Solution Provider (CSP) tại Việt Nam theo hình thức thanh toán theo tháng hoặc theo năm. Khi được Ban Lãnh đạo duyệt chủ trương, Phòng CNTT sẽ tiến hành lấy báo giá cạnh tranh từ 03 nhà cung cấp để đàm phán mức chiết khấu tốt nhất cho A25 Hotel.")
    r_n.italic = True
    r_n.font.size = Pt(8.5)
    r_n.font.color.rgb = RGBColor(100, 100, 100)

    # =========================================================================
    # SECTION V: LỘ TRÌNH CHUYỂN ĐỔI & QUY TẮC ĐÓNG NHÓM ZALO
    # =========================================================================
    add_heading_styled(doc, "V. LỘ TRÌNH CHUYỂN ĐỔI & QUY TẮC ĐÓNG NHÓM ZALO CÔNG VIỆC", 1)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "Để việc chuyển đổi diễn ra nhịp nhàng, không gây sốc cho nhân sự và không làm gián đoạn sản xuất kinh doanh, lộ trình thực hiện được chia làm 03 giai đoạn chặt chẽ:"
    )

    add_heading_styled(doc, "Giai đoạn 1: Thí điểm tại Khối Văn phòng A25 Hotel (Tháng 11 - 12/2026)", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "• Triển khai kỹ thuật: Thiết lập cổng quản trị Microsoft 365, cấu hình tên miền @a25hotel.com, phân bổ giấy phép và khởi tạo cây danh bạ Entra ID cho toàn bộ nhân sự Văn phòng.\n"
        "• Đào tạo hướng dẫn: Tổ chức 02 buổi đào tạo trực tiếp và gửi tài liệu hướng dẫn nhanh cho toàn thể cán bộ văn phòng.\n"
        "• QUY TẮC ĐÓNG NHÓM ZALO BẮT BUỘC (Quy tắc dứt điểm): Sau khi phòng/nhóm nghiệp vụ được tạo Team trên Microsoft Teams và nhân sự đã được đào tạo, Trưởng bộ phận phải ấn định ngày chuyển đổi dứt điểm (thời gian làm quen tối đa 07 ngày làm việc). Đúng ngày quy định, Trưởng nhóm phải tuyên bố đóng nhóm Zalo công việc đó (hoặc đổi tên nhóm thành 'Lưu trữ - Đã chuyển Teams') và chỉ định rõ mọi thông tin phát sinh chuyển sang Teams. TUYỆT ĐỐI KHÔNG DUY TRÌ SONG SONG HAI KÊNH để tránh tình trạng nhân viên tiếp tục quay lại thói quen cũ."
    )

    add_heading_styled(doc, "Giai đoạn 2: Đo lường và Đánh giá hiệu quả thực tế tại Khối Văn phòng (Tháng 01/2027)", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "• Đo lường kỹ thuật: Trích xuất báo cáo mức độ hoạt động (Adoption Score & Usage Reports) trên cổng Microsoft 365 Admin Center (tỷ lệ đăng nhập hằng ngày, số lượng cuộc họp Teams, dung lượng dữ liệu lưu trên OneDrive).\n"
        "• Khảo sát người dùng: Gửi phiếu đánh giá ngắn đến các Trưởng phòng ban (Kinh doanh, Kế toán, Nhân sự...) về tốc độ tìm kiếm tài liệu, mức độ bảo mật và mức độ thuận tiện trong trao đổi công việc.\n"
        "• Báo cáo Ban Lãnh đạo: Tổng kết các điểm đạt được, các vướng mắc phát sinh và kiến nghị hoàn thiện quy chế."
    )

    add_heading_styled(doc, "Giai đoạn 3: Đề xuất phương án mở rộng xuống Chi nhánh Khách sạn A25 (Quý I - II/2027)", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    p.add_run(
        "• Nguyên tắc bất di bất dịch: KHÔNG LÀM GIÁN ĐOẠN HOẠT ĐỘNG VẬN HÀNH KHÁCH SẠN. Hoạt động lễ tân, buồng phòng tại chi nhánh mang tính đặc thù cao, thay ca liên tục và đang chạy ổn định trên Zalo nội bộ khách sạn.\n"
        "• Lộ trình mở rộng thận trọng: Chỉ triển khai trước cho Quản lý khách sạn (Hotel Manager) và Giám sát cơ sở để làm việc với Khối Văn phòng; đối với nhân sự lễ tân/buồng phòng chi nhánh sẽ xây dựng tài khoản dùng chung (Shared Mailbox / Property Account) hoặc gói Microsoft 365 F3 (dành cho Firstline Worker chi phí thấp ~ 60.000 đ/tháng) sau khi đã đánh giá thành công giai đoạn thí điểm văn phòng."
    )

    # =========================================================================
    # SECTION VI: TÀI LIỆU MINH HỌA & HƯỚNG DẪN SỬ DỤNG CÓ VIDEO TRỰC QUAN
    # =========================================================================
    add_heading_styled(doc, "VI. TÀI LIỆU MINH HỌA, TRUNG TÂM VIDEO & HƯỚNG DẪN THAO TÁC TRỰC QUAN", 1)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "Nhằm giải quyết triệt để yêu cầu của Ban Lãnh đạo và nhu cầu thực tế của Cán bộ nhân viên (tránh tình trạng chỉ cung cấp các đường link lý thuyết khó áp dụng), Phòng CNTT xây dựng Bộ cẩm nang hướng dẫn sử dụng tích hợp đa phương tiện bao gồm: Hình ảnh đồ họa trực quan (Infographics), Thẻ video thao tác thực tế có nút Play/Thời lượng cụ thể, và Hướng dẫn từng bước (Step-by-Step) cho từng tình huống công việc hằng ngày tại A25 Hotel."
    )

    # INSERT FIGURE 4: TRUNG TÂM VIDEO & TÀI LIỆU ĐÀO TẠO
    add_illustrated_figure(doc,
        "M365_04_TrungTam_Video_DaoTao.png",
        "Hình 4: Trung tâm Video Đào tạo Trực quan & Bộ học liệu hướng dẫn thao tác Microsoft 365 dành cho Nhân sự A25 Hotel",
        width_inch=6.4)

    # CALLOUT ĐÍNH KÈM FILE VIDEO MP4 & HTML HUB
    add_callout(doc,
        "🎬 ĐÍNH KÈM FILE VIDEO HƯỚNG DẪN THỰC TẾ & TRÌNH PHÁT TƯƠNG TÁC:",
        "Phòng CNTT đã kết xuất sẵn file Video đào tạo thực tế chuẩn định dạng MP4 HD 720p và Trình duyệt Video tương tác đa phương tiện đính kèm theo hồ sơ phương án này:\n"
        "• File Video gốc: Video_Huong_Dan_Chuan_Hoa_Microsoft_365_A25_Hotel.mp4 (Thời lượng 01:22, chuẩn nén H.264/AAC, độ phân giải HD 720p, có phụ đề thuyết minh từng phân hệ).\n"
        "• File Trình phát tương tác: Trung_Tam_Video_Huong_Dan_M365_A25_Hotel.html (Mở bằng bất kỳ trình duyệt Chrome/Edge, có chức năng chọn nhảy từng chuyên đề và nút tải video/tài liệu nhanh).",
        border_color="0078D4", bg_color="EFF6FF")

    add_heading_styled(doc, "1. Trung tâm 04 Video bài giảng thao tác thực hành (Video Training Hub)", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "Phòng CNTT thiết kế 04 video bài giảng ngắn (thời lượng từ 2 đến 4 phút) dạng 'cầm tay chỉ việc', quay trực tiếp thao tác màn hình thực tế có thuyết minh tiếng Việt và phụ đề rõ ràng:"
    )

    # Bảng chi tiết 4 Video Cards
    tbl_videos = doc.add_table(rows=5, cols=4)
    tbl_videos.alignment = WD_TABLE_ALIGNMENT.CENTER
    v_headers = ["Video & Mã số", "Thời lượng", "Mục tiêu & Kịch bản thao tác từng bước", "Kênh truy cập & Link trực tiếp"]
    v_w = [Inches(1.5), Inches(0.9), Inches(3.2), Inches(1.4)]

    for c_idx, h_text in enumerate(v_headers):
        cell = tbl_videos.rows[0].cells[c_idx]
        cell.width = v_w[c_idx]
        cell.text = h_text
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.runs[0]
        r.bold = True
        r.font.name = "Arial"
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, top=50, bottom=50, left=50, right=50)
        set_cell_border(cell, top=('single', '4', '0', 'CBD5E1'),
                              bottom=('single', '4', '0', 'CBD5E1'),
                              left=('single', '4', '0', 'CBD5E1'),
                              right=('single', '4', '0', 'CBD5E1'))

    video_items = [
        ("▶ Video 01:\nLàm quen Microsoft Teams",
         "03:15 phút",
         "• 00:00 - 00:45: Đăng nhập tài khoản @a25hotel.com và làm quen giao diện.\n"
         "• 00:46 - 01:40: Cách tìm kiếm đồng nghiệp qua email/tên và bắt đầu đoạn Chat 1-1.\n"
         "• 01:41 - 02:30: Tạo kênh thảo luận phòng ban (Channel), gửi file và gắn thẻ @mention.\n"
         "• 02:31 - 03:15: Thao tác gọi điện nội bộ và mở cuộc họp video có chia sẻ tài liệu.",
         "Microsoft Teams Video Training\n(Link Support)"),
        ("▶ Video 02:\nChuẩn hóa Email Outlook",
         "04:20 phút",
         "• 00:00 - 01:10: Cài đặt chữ ký thư điện tử chuẩn nhận diện A25 Hotel kèm logo.\n"
         "• 01:11 - 02:15: Soạn thảo thư công vụ gửi đối tác/khách hàng chuyên nghiệp.\n"
         "• 02:16 - 03:20: Sử dụng Lịch (Calendar), kiểm tra giờ rảnh của Lãnh đạo qua Scheduling Assistant.\n"
         "• 03:21 - 04:20: Gửi thư mời họp công vụ và nhận xác nhận tham gia tự động.",
         "Outlook Video Training\n(Link Support)"),
        ("▶ Video 03:\nLưu trữ & Chia sẻ OneDrive",
         "03:45 phút",
         "• 00:00 - 01:00: Cài đặt ứng dụng OneDrive đồng bộ thư mục máy tính (Desktop, Docs).\n"
         "• 01:01 - 02:00: Nhận biết biểu tượng trạng thái file (Đám mây / Tích xanh).\n"
         "• 02:01 - 03:00: Phân quyền chia sẻ tài liệu (Xem / Chỉnh sửa) cho đúng người.\n"
         "• 03:01 - 03:45: Cách khôi phục phiên bản cũ (Version History) khi lỡ lưu đè dữ liệu.",
         "OneDrive Video Training\n(Link Support)"),
        ("▶ Video 04:\nTra cứu Danh bạ & Profile",
         "02:30 phút",
         "• 00:00 - 00:45: Cách gõ tìm kiếm đồng nghiệp trên thanh Search của Teams/Outlook.\n"
         "• 00:46 - 01:35: Đọc thông tin Thẻ hồ sơ (Profile Card): Chức danh, Phòng ban, Di động.\n"
         "• 01:36 - 02:30: Mở Sơ đồ tổ chức (Org Chart) để xác định Quản lý trực tiếp và cấp dưới.",
         "Profile Card Guide\n(Link Support)")
    ]

    for r_idx, (v_title, v_dur, v_steps, v_src) in enumerate(video_items, start=1):
        # Col 0: Title
        cell0 = tbl_videos.rows[r_idx].cells[0]
        cell0.width = v_w[0]
        cell0.text = v_title
        p0 = cell0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.runs[0].bold = True
        p0.runs[0].font.name = "Arial"
        p0.runs[0].font.size = Pt(8.5)
        p0.runs[0].font.color.rgb = RGBColor(180, 20, 20)

        # Col 1: Duration
        cell1 = tbl_videos.rows[r_idx].cells[1]
        cell1.width = v_w[1]
        cell1.text = v_dur
        p1 = cell1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1.runs[0].font.name = "Arial"
        p1.runs[0].font.size = Pt(8.5)
        p1.runs[0].bold = True
        p1.runs[0].font.color.rgb = RGBColor(30, 58, 138)

        # Col 2: Steps
        cell2 = tbl_videos.rows[r_idx].cells[2]
        cell2.width = v_w[2]
        cell2.text = v_steps
        p2 = cell2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p2.runs[0].font.name = "Arial"
        p2.runs[0].font.size = Pt(8)

        # Col 3: Channel
        cell3 = tbl_videos.rows[r_idx].cells[3]
        cell3.width = v_w[3]
        cell3.text = v_src
        p3 = cell3.paragraphs[0]
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p3.runs[0].font.name = "Arial"
        p3.runs[0].font.size = Pt(8)
        p3.runs[0].font.color.rgb = RGBColor(37, 99, 235)

        for c_idx in range(4):
            c = tbl_videos.rows[r_idx].cells[c_idx]
            set_cell_background(c, "F8FAFC" if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(c, top=45, bottom=45, left=45, right=45)
            set_cell_border(c, top=('single', '4', '0', 'CBD5E1'),
                               bottom=('single', '4', '0', 'CBD5E1'),
                               left=('single', '4', '0', 'CBD5E1'),
                               right=('single', '4', '0', 'CBD5E1'))

    # 2. Hướng dẫn sử dụng thao tác thực tế từng bước
    add_heading_styled(doc, "2. Hướng dẫn thao tác thực hành chuẩn hóa từng ứng dụng (Illustrated Step-by-Step)", 2)

    # 2.1 Teams
    add_heading_styled(doc, "a. Hướng dẫn sử dụng Microsoft Teams trong công việc hằng ngày", 3)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "• Bước 1: Khởi động và đăng nhập ứng dụng Teams trên máy tính (hoặc truy cập trình duyệt tại teams.microsoft.com) bằng tài khoản định danh được cấp: ten.ho@a25hotel.com.\n"
        "• Bước 2 (Nhắn tin 1-1): Nhấp vào biểu tượng 'Chat' trên thanh công cụ bên trái -> Bấm biểu tượng 'New Chat' (cây bút) -> Gõ tên hoặc email đồng nghiệp -> Nhập nội dung và bấm Enter để gửi.\n"
        "• Bước 3 (Làm việc trong Nhóm/Kênh): Vào mục 'Teams' -> Chọn Kênh nghiệp vụ của phòng ban (ví dụ: 'Kinh doanh', 'Kế toán', 'CNTT') -> Bấm 'Start a post' để tạo chủ đề thảo luận chung. Sử dụng cú pháp @tên_người để thông báo trực tiếp cho người cần xử lý.\n"
        "• Bước 4 (Họp trực tuyến & Chia sẻ màn hình): Trong cuộc trò chuyện hoặc kênh, bấm nút 'Meet' (biểu tượng máy quay). Trong cuộc họp, bấm nút 'Share' (mũi tên hướng lên) để chiếu màn hình báo cáo hoặc bảng tính Excel cho toàn nhóm cùng xem."
    )

    # 2.2 Outlook
    add_heading_styled(doc, "b. Hướng dẫn sử dụng Microsoft Outlook cho thư tín công vụ và đặt lịch họp", 3)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "• Bước 1 (Cài đặt chữ ký email chuẩn A25): Mở Outlook -> Chọn File -> Options -> Mail -> Signatures -> Tạo chữ ký mới theo mẫu quy định: Họ tên | Chức danh | Phòng ban | CÔNG TY CỔ PHẦN A25 HOTEL | Số điện thoại di động | Website: a25hotel.com.\n"
        "• Bước 2 (Soạn thảo thư giao dịch chuẩn): Bấm 'New Email' -> Gõ địa chỉ người nhận vào ô To, người liên quan vào ô Cc -> Tiêu đề thư bắt buộc rõ ràng theo cấu trúc: [A25 HOTEL] - [Tên phòng ban] - [Nội dung ngắn gọn] -> Soạn thảo văn phong trang trọng, đính kèm link tài liệu OneDrive thay vì file nặng.\n"
        "• Bước 3 (Tổ chức cuộc họp qua Calendar): Chuyển sang thẻ 'Calendar' trên Outlook -> Chọn 'New Meeting' -> Điền tiêu đề, thời gian họp -> Nhấp 'Scheduling Assistant' để xem lịch làm việc bận/rảnh của các đại biểu tham dự nhằm chọn khung giờ phù hợp nhất mà không bị trùng lịch -> Bấm 'Send' để gửi giấy mời tự động kèm link họp Teams."
    )

    # 2.3 OneDrive
    add_heading_styled(doc, "c. Hướng dẫn lưu trữ dữ liệu cá nhân & chia sẻ an toàn trên OneDrive for Business", 3)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "• Bước 1 (Đồng bộ thư mục làm việc): Đăng nhập ứng dụng OneDrive trên máy tính bằng tài khoản A25 Hotel. Ứng dụng sẽ tự động tạo thư mục 'OneDrive - A25 Hotel' trong File Explorer của Windows. Mọi file lưu trong thư mục này đều được sao lưu liên tục lên đám mây bảo mật.\n"
        "• Bước 2 (Hiểu biểu tượng trạng thái file):\n"
        "   + Biểu tượng Đám mây xanh: File được lưu an toàn trên Cloud, không chiếm dung lượng ổ cứng máy tính.\n"
        "   + Biểu tượng Tích xanh lá cây: File đã được tải về máy tính và sẵn sàng mở khi không có kết nối Internet.\n"
        "   + Biểu tượng Hai mũi tên xoay vòng: File đang trong quá trình đồng bộ lên máy chủ.\n"
        "• Bước 3 (Chia sẻ tài liệu phân quyền đúng chuẩn): Nhấp chuột phải vào file cần gửi -> Chọn 'OneDrive' -> Chọn 'Share' -> Thiết lập quyền: 'People in A25 Hotel' (Người trong công ty) -> Tùy chọn 'Can view' (Chỉ cho xem) hoặc 'Can edit' (Được phép chỉnh sửa) -> Bấm 'Copy Link' để dán vào Teams hoặc Email. TUYỆT ĐỐI KHÔNG chọn 'Anyone with the link' để tránh lộ lọt dữ liệu ra bên ngoài."
    )

    # 2.4 Profile Card & Org Chart
    add_heading_styled(doc, "d. Hướng dẫn tra cứu Thẻ hồ sơ nhân viên (Profile Card) & Sơ đồ tổ chức (Org Chart)", 3)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "• Bước 1: Khi cần tra cứu thông tin của bất kỳ cán bộ nhân viên nào, gõ tên hoặc email của họ trên thanh tìm kiếm của Teams hoặc Outlook.\n"
        "• Bước 2: Rê chuột lên ảnh đại diện của nhân sự đó, Thẻ hồ sơ thông minh (Profile Card) sẽ tự động bật lên, hiển thị đầy đủ: Chức danh chuyên môn, Phòng ban công tác, Số điện thoại nội bộ và địa chỉ email chính thức.\n"
        "• Bước 3: Nhấp vào tab 'Organization' (Tổ chức) trên Profile Card. Màn hình sẽ hiển thị cây phân cấp quản lý: Cán bộ Quản lý trực tiếp (Manager) ở phía trên, các đồng nghiệp cùng phòng ở ngang hàng, và các nhân sự cấp dưới trực tiếp (nếu có). Tính năng này giúp nhân sự mới và các bộ phận dễ dàng xác định đúng cấp có thẩm quyền báo cáo."
    )

    # 3. Bảng Link tài liệu Microsoft Support chính thức
    add_heading_styled(doc, "3. Danh mục tài liệu đào tạo chính thức từ Microsoft (Official Microsoft Support Links)", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "Đính kèm danh sách tài liệu chuẩn từ Microsoft có phụ đề tiếng Việt và video thao tác trực tiếp để cán bộ nhân viên tự học và tra cứu bất kỳ lúc nào:"
    )

    tbl_links = doc.add_table(rows=7, cols=3)
    tbl_links.alignment = WD_TABLE_ALIGNMENT.CENTER
    l_headers = ["Ứng dụng", "Nội dung hướng dẫn chính thức", "Đường dẫn tài liệu Microsoft Support"]
    l_w = [Inches(1.5), Inches(3.0), Inches(2.5)]

    for c_idx, h_text in enumerate(l_headers):
        cell = tbl_links.rows[0].cells[c_idx]
        cell.width = l_w[c_idx]
        cell.text = h_text
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.runs[0]
        r.bold = True
        r.font.name = "Arial"
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, top=50, bottom=50, left=50, right=50)
        set_cell_border(cell, top=('single', '4', '0', 'D1D5DB'),
                              bottom=('single', '4', '0', 'D1D5DB'),
                              left=('single', '4', '0', 'D1D5DB'),
                              right=('single', '4', '0', 'D1D5DB'))

    links_data = [
        ("Microsoft Teams", "Bắt đầu nhanh với Microsoft Teams (Video & Bài viết)", "https://support.microsoft.com/vi-vn/office/teams-video-training"),
        ("Microsoft Teams", "Cách tạo cuộc họp trực tuyến và chia sẻ màn hình", "https://support.microsoft.com/vi-vn/office/meetings-teams"),
        ("Microsoft Outlook", "Hướng dẫn sử dụng hòm thư điện tử Outlook (Cơ bản - Nâng cao)", "https://support.microsoft.com/vi-vn/office/outlook-training"),
        ("Microsoft Outlook", "Quản lý Lịch làm việc và mời họp công vụ", "https://support.microsoft.com/vi-vn/office/calendar-outlook"),
        ("OneDrive for Business", "Đồng bộ tài liệu và chia sẻ file an toàn trên OneDrive", "https://support.microsoft.com/vi-vn/office/onedrive-training"),
        ("Profile & Org Chart", "Xem thẻ hồ sơ nhân viên và sơ đồ tổ chức trên Teams/Outlook", "https://support.microsoft.com/vi-vn/office/profile-card-teams")
    ]

    for r_idx, (app, desc, url) in enumerate(links_data, start=1):
        tbl_links.rows[r_idx].cells[0].width = l_w[0]
        tbl_links.rows[r_idx].cells[0].text = app
        tbl_links.rows[r_idx].cells[0].paragraphs[0].runs[0].bold = True
        tbl_links.rows[r_idx].cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        tbl_links.rows[r_idx].cells[1].width = l_w[1]
        tbl_links.rows[r_idx].cells[1].text = desc

        tbl_links.rows[r_idx].cells[2].width = l_w[2]
        tbl_links.rows[r_idx].cells[2].text = url
        tbl_links.rows[r_idx].cells[2].paragraphs[0].runs[0].font.size = Pt(8)
        tbl_links.rows[r_idx].cells[2].paragraphs[0].runs[0].font.color.rgb = RGBColor(37, 99, 235)

        for c_idx in range(3):
            cell = tbl_links.rows[r_idx].cells[c_idx]
            cell.paragraphs[0].runs[0].font.name = "Arial"
            if c_idx != 2:
                cell.paragraphs[0].runs[0].font.size = Pt(8.5)
            set_cell_background(cell, "F9FAFB" if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, top=50, bottom=50, left=50, right=50)
            set_cell_border(cell, top=('single', '4', '0', 'D1D5DB'),
                                  bottom=('single', '4', '0', 'D1D5DB'),
                                  left=('single', '4', '0', 'D1D5DB'),
                                  right=('single', '4', '0', 'D1D5DB'))

    # =========================================================================
    # SECTION VII: ĐỀ XUẤT & KIẾN NGHỊ BAN LÃNH ĐẠO PHÊ DUYỆT
    # =========================================================================
    add_heading_styled(doc, "VII. ĐỀ XUẤT & KIẾN NGHỊ BAN LÃNH ĐẠO PHÊ DUYỆT", 1)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "Nhằm đưa hệ thống công cụ làm việc của Công ty Cổ phần A25 Hotel lên tầm chuyên nghiệp, tăng cường năng suất và bảo vệ tài sản thông tin doanh nghiệp, Phòng Công nghệ thông tin kính đề xuất Ban Lãnh đạo xem xét phê duyệt các nội dung sau:\n"
        "1. Phê duyệt chủ trương chuẩn hóa công cụ làm việc trên Microsoft 365 (Teams, Outlook, OneDrive) và phê duyệt kế hoạch triển khai thí điểm tại Khối Văn phòng trong tháng 11 - 12/2026.\n"
        "2. Phê duyệt chủ trương cấp hạn mức ngân sách tham khảo (khoảng 8.600.000 VNĐ / tháng cho ~40 tài khoản văn phòng) để Phòng CNTT tiến hành đấu thầu/chọn nhà cung cấp bản quyền có mức chiết khấu tốt nhất.\n"
        "3. Ban hành Quy định ranh giới sử dụng công cụ: Khẳng định HROnline là kênh duy nhất giao việc/phê duyệt; bắt buộc đóng các nhóm Zalo công việc văn phòng sau khi chuyển sang Teams.\n"
        "4. Chỉ đạo Phòng Tổ chức Nhân sự phối hợp cùng Phòng CNTT chuẩn hóa dữ liệu đầu vào nhân sự để cấu hình cây thư mục Entra ID và Sổ địa chỉ toàn cầu (GAL)."
    )

    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(6)
    p_end.paragraph_format.space_after = Pt(12)
    r_end = p_end.add_run("Kính trình Ban Giám đốc xem xét, chỉ đạo và phê duyệt để Phòng CNTT kịp thời triển khai đúng tiến độ trước ngày 25/11/2026.")
    r_end.italic = True
    r_end.font.size = Pt(9.5)

    # Signatures Table (3 columns)
    sig_table = doc.add_table(rows=1, cols=3)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_titles = [
        "NGƯỜI LẬP PHƯƠNG ÁN\n(Trợ lý Phòng CNTT)\n\n\n\n\nHoàng Minh Quân",
        "PHỤ TRÁCH BỘ PHẬN\n(Trưởng ban PT & QLPM)\n\n\n\n\nNguyễn Văn Sơn",
        "BAN GIÁM ĐỐC PHÊ DUYỆT\n(Giám đốc Kinh doanh / TGĐ)\n\n\n\n\n(Ký và ghi rõ họ tên)"
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
        set_cell_margins(c, top=60, bottom=60, left=40, right=40)

    try:
        doc.save(DOC_PATH)
        print("SUCCESS: Full proposal with 4 illustrated figures saved to:", DOC_PATH)
    except PermissionError:
        alt_path = r"c:\Users\Admin\Desktop\tester\Bao_cao_Phuong_an_Chuan_hoa_Cong_cu_Microsoft_365_A25_Hotel_Co_Video_MP4.docx"
        doc.save(alt_path)
        print("NOTE: File is currently open in Word. Saved successfully to alternate path:", alt_path)

if __name__ == '__main__':
    generate_full_doc()
