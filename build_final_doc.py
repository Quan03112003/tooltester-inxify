import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import os

DOC_PATH = r"c:\Users\Admin\Desktop\tester\Bao_cao_Phuong_an_Chuan_hoa_Cong_cu_Microsoft_365_A25_Hotel.docx"
IMG_GUIDE_DIR = r"c:\Users\Admin\Desktop\tester\img_m365_guide"
IMG_REAL_DIR = r"c:\Users\Admin\Desktop\tester\img_real_m365"

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

def resolve_image_path(img_name):
    # Try img_m365_guide first
    p1 = os.path.join(IMG_GUIDE_DIR, img_name)
    if os.path.exists(p1):
        return p1
    # Try img_real_m365
    p2 = os.path.join(IMG_REAL_DIR, img_name)
    if os.path.exists(p2):
        return p2
    # Try current directory
    p3 = os.path.join(r"c:\Users\Admin\Desktop\tester", img_name)
    if os.path.exists(p3):
        return p3
    return None

def add_illustrated_figure(doc, img_name, caption_text, width_inch=6.5):
    img_path = resolve_image_path(img_name)
    if not img_path:
        print(f"Warning: Image file not found: {img_name}")
        return

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

def generate_final_m365_doc():
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
    r0_3 = p0.add_run("Số: 18-11/2026/TTr-CNTT-A25")
    r0_3.font.size = Pt(9)
    r0_3.font.name = "Arial"
    r0_3.italic = True

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
    r1_3 = p1.add_run("-------------------\n")
    r1_3.font.size = Pt(8)
    r1_4 = p1.add_run("Hà Nội, ngày 20 tháng 11 năm 2026")
    r1_4.font.size = Pt(9)
    r1_4.font.name = "Arial"
    r1_4.italic = True

    for c in (c0, c1):
        set_cell_border(c)
        set_cell_margins(c, top=40, bottom=40, left=40, right=40)

    # Divider line
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(6)
    p_div.paragraph_format.space_after = Pt(12)

    # 3. Main Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(6)
    p_title.paragraph_format.space_after = Pt(4)
    run_t = p_title.add_run("TỜ TRÌNH PHƯƠNG ÁN CHUẨN HÓA CÔNG CỤ LÀM VIỆC SỐ\nTRÊN HỆ SINH THÁI MICROSOFT 365 TẠI A25 HOTEL")
    run_t.bold = True
    run_t.font.name = "Arial"
    run_t.font.size = Pt(15)
    run_t.font.color.rgb = RGBColor(180, 20, 20)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(12)
    run_sub = p_sub.add_run("(V/v: Chuẩn hóa hòm thư Outlook, Nhóm làm việc Teams, Lưu trữ đám mây OneDrive,\nĐồng bộ Danh bạ Entra ID, Phân Role Owner/Member & Lộ trình đóng nhóm Zalo văn phòng)")
    run_sub.font.name = "Arial"
    run_sub.font.size = Pt(10.5)
    run_sub.font.color.rgb = RGBColor(30, 58, 138)
    run_sub.bold = True

    # 4. Meta Table (Kính gửi, Người lập, Ngày trình)
    tbl_meta = doc.add_table(rows=4, cols=2)
    tbl_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    m_w = [Inches(1.8), Inches(5.2)]
    meta_rows = [
        ("Kính gửi:", "Ban Tổng Giám đốc Công ty Cổ phần A25 Hotel\nGiám đốc Khối Vận hành & Giám đốc Kinh doanh"),
        ("Người lập tờ trình:", "Hoàng Minh Quân – Trợ lý Phòng CNTT (Ban Phát triển & Quản lý phần mềm)"),
        ("Người kiểm duyệt:", "Nguyễn Văn Sơn – Trưởng Ban Phát triển & Quản lý phần mềm"),
        ("Thời hạn báo cáo:", "Hoàn thành và trình duyệt trước ngày 25/11/2026 (Theo chỉ đạo cuộc họp 18/11/2026)")
    ]
    for idx, (label, val) in enumerate(meta_rows):
        c_lbl = tbl_meta.rows[idx].cells[0]
        c_val = tbl_meta.rows[idx].cells[1]
        c_lbl.width = m_w[0]
        c_val.width = m_w[1]
        c_lbl.text = label
        c_val.text = val
        p_l = c_lbl.paragraphs[0]
        p_l.runs[0].bold = True
        p_l.runs[0].font.name = "Arial"
        p_l.runs[0].font.size = Pt(9.5)
        p_v = c_val.paragraphs[0]
        p_v.runs[0].font.name = "Arial"
        p_v.runs[0].font.size = Pt(9.5)
        set_cell_background(c_lbl, "F1F5F9")
        set_cell_background(c_val, "FFFFFF")
        set_cell_margins(c_lbl, top=45, bottom=45, left=50, right=50)
        set_cell_margins(c_val, top=45, bottom=45, left=50, right=50)
        set_cell_border(c_lbl, top=('single', '4', '0', 'D1D5DB'),
                               bottom=('single', '4', '0', 'D1D5DB'),
                               left=('single', '4', '0', 'D1D5DB'),
                               right=('single', '4', '0', 'D1D5DB'))
        set_cell_border(c_val, top=('single', '4', '0', 'D1D5DB'),
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

    # FIGURE 1: SO SÁNH ZALO VS M365
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
    # SECTION II: GIẢI PHÁP DANH BẠ TẬP TRUNG, PHÂN ROLE & LIÊN KẾT TEAMS - OUTLOOK
    # =========================================================================
    add_heading_styled(doc, "II. GIẢI PHÁP DANH BẠ TẬP TRUNG, PHÂN QUYỀN ROLE & LIÊN KẾT TEAMS - OUTLOOK", 1)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "Để giải quyết dứt điểm các vướng mắc của nhân viên về việc 'không biết tìm kiếm ai', 'không rõ cấp trên là ai', và 'thiếu tính liên kết giữa các công cụ', phương án xây dựng một hệ thống danh bạ đồng bộ hoàn chỉnh trên Microsoft Entra ID (Azure AD), tích hợp sâu giữa Microsoft Teams và Microsoft Outlook:"
    )

    add_heading_styled(doc, "1. Thẻ hồ sơ thông minh (Profile Card) & Sơ đồ tổ chức động (Organization Chart)", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "Khi bất kỳ nhân sự nào gõ tên hoặc email đồng nghiệp trên Outlook hoặc Teams, hệ thống sẽ tự động hiển thị Thẻ hồ sơ trực quan gồm:\n"
        "• Thông tin định danh: Ảnh đại diện chuẩn, Họ và tên đầy đủ, Mã nhân viên A25, Email chính thức (@a25hotel.com), Số điện thoại nội bộ/di động công vụ.\n"
        "• Vị trí công tác & Role: Chức danh chuyên môn (Job Title/Role), Phòng ban trực thuộc, Chi nhánh hoặc Khối làm việc (Văn phòng Tập đoàn / Khối Vận hành).\n"
        "• Sơ đồ phân cấp quản trị (Organization Chart): Thể hiện trực quan cấp Quản lý trực tiếp (Manager) mà nhân sự đó báo cáo, cùng danh sách các nhân sự cấp dưới (Direct Reports) do nhân sự đó quản lý. Điều này giúp nhân viên mới hoặc các phòng ban phối hợp dễ dàng nắm bắt đúng đầu mối phụ trách mà không cần hỏi qua trung gian."
    )

    # FIGURE 2: PROFILE CARD & ORG CHART
    add_illustrated_figure(doc,
        "M365_03_ProfileCard_OrgChart.png",
        "Hình 2: Giao diện Thẻ hồ sơ nhân viên (Profile Card) và Sơ đồ tổ chức phân cấp quản trị trên Microsoft Teams & Outlook",
        width_inch=6.4)

    add_heading_styled(doc, "2. Cơ chế Phân quyền vai trò (Role Assignment) trong Microsoft Teams", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "Tại giao diện quản trị nhóm (Manage Team), quyền hạn của nhân sự được phân định rõ ràng thành 02 cấp bậc chuẩn hóa:\n"
        "• Chủ sở hữu nhóm (Owner): Dành riêng cho Cán bộ Quản lý, Trưởng/Phó phòng ban hoặc Trưởng dự án (được cấp giấy phép M365 Business Standard). Chủ sở hữu có toàn quyền: Duyệt hoặc gỡ bỏ thành viên, tạo và quản lý các kênh chuyên môn (Standard/Private Channels), thiết lập quyền gửi bài, ghim tin nhắn quan trọng và kiểm duyệt nội dung trao đổi.\n"
        "• Thành viên nhóm (Member): Dành cho toàn bộ Nhân viên chuyên môn, Chuyên viên tác nghiệp (được cấp giấy phép M365 Business Basic). Thành viên có quyền: Trò chuyện trong các kênh chung, gửi file tài liệu, cộng tác trực tiếp trên Word/Excel, mở cuộc họp video nội bộ, gắn thẻ @mention đồng nghiệp nhưng KHÔNG có quyền tự ý xóa kênh, thêm khách ngoài hoặc gỡ thành viên khác.\n"
        "• Khách (Guest): Nghiêm cấm cấp quyền Khách ngoài tổ chức vào các nhóm nội bộ của A25 Hotel khi chưa có phê duyệt bằng văn bản của Trưởng phòng CNTT."
    )

    # FIGURE 3: REAL TEAMS ROLES
    add_illustrated_figure(doc,
        "real_teams_overview_2.png",
        "Hình 3: Giao diện Quản trị Phân quyền Nhóm thực tế trên Microsoft Teams (Phân cấp Chủ sở hữu - Owner và Thành viên - Member)",
        width_inch=6.4)

    add_heading_styled(doc, "3. Liên kết đồng bộ trực tiếp giữa Microsoft Teams và Microsoft Outlook Calendar", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "Điểm đột phá của hệ sinh thái Microsoft 365 là tính năng tích hợp 2 chiều khép kín giữa Outlook và Teams, loại bỏ hoàn toàn tình trạng phải copy paste link thủ công:\n"
        "• Tích hợp nút 'New Teams Meeting' trên Outlook: Khi nhân viên mở ứng dụng Outlook và chuyển sang màn hình Lịch (Calendar), trên thanh công cụ Ribbon xuất hiện nút 'New Teams Meeting' (Cuộc họp Teams mới). Chỉ với 01 thao tác bấm, hệ thống tự động sinh phòng họp Teams trực tuyến với đường link an toàn, Meeting ID và Passcode đính kèm ngay vào nội dung thư mời.\n"
        "• Trợ lý xếp lịch thông minh (Scheduling Assistant): Khi mời nhiều thành viên tham gia (ví dụ: Ban Giám đốc, Trưởng phòng, các Chuyên viên), tính năng Trợ lý xếp lịch sẽ quét thời gian biểu thời gian thực của tất cả mọi người từ hệ thống Entra ID. Màn hình hiển thị dạng lưới phân biệt rõ khung giờ 'Rảnh' (Free), 'Bận' (Busy) hoặc 'Tạm thời' (Tentative). Người tạo cuộc họp chỉ cần nhấp chuột chọn khung giờ xanh không trùng lịch (ví dụ: 14:00 - 15:00), bấm Send là thư mời và lịch họp sẽ tự động đồng bộ tức thì vào Calendar của cả Outlook và Teams của tất cả thành viên."
    )

    # FIGURE 4: REAL OUTLOOK SCHEDULING & TEAMS MEETING
    add_illustrated_figure(doc,
        "real_outlook_sched_2.png",
        "Hình 4: Giao diện Trợ lý xếp lịch (Scheduling Assistant) và Nút tạo Teams Meeting trực tiếp trên Outlook Calendar",
        width_inch=6.4)

    add_heading_styled(doc, "4. Quy trình chuẩn hóa dữ liệu đầu vào giữa Phòng Tổ chức Nhân sự (HR) và Phòng CNTT (IT)", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    p.add_run(
        "Để cây thư mục và danh bạ nhân sự luôn chính xác theo thời gian thực, quy định bắt buộc luồng bàn giao thông tin:\n"
        "• Khi có nhân sự mới: Phòng TCNS gửi phiếu yêu cầu tạo tài khoản cho IT tối thiểu 24 giờ trước ngày nhân sự nhận việc, trong đó bắt buộc có đủ 04 trường cốt lõi: (1) Chức danh chính thức, (2) Phòng ban/Đơn vị công tác, (3) Họ tên cán bộ Quản lý trực tiếp (Manager) và (4) Vai trò cấp phép (Standard cho Lãnh đạo, Basic cho Nhân viên).\n"
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
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, "1E3A8A") # Navy Blue
        set_cell_margins(cell, top=50, bottom=50, left=50, right=50)
        set_cell_border(cell, top=('single', '4', '0', 'D1D5DB'),
                              bottom=('single', '4', '0', 'D1D5DB'),
                              left=('single', '4', '0', 'D1D5DB'),
                              right=('single', '4', '0', 'D1D5DB'))

    tools_data = [
        ("HROnline\n(Hệ thống cốt lõi)",
         "Giao việc, Báo cáo tiến độ, Chấm công & Trình ký phê duyệt",
         "• Giao việc có thời hạn (Deadline), phân công người chủ trì/phối hợp.\n"
         "• Trình duyệt đề xuất thanh toán, xin nghỉ phép, duyệt chi hoa hồng OTA.\n"
         "• Cập nhật tỷ lệ % hoàn thành công việc và đánh giá KPI.",
         "KHÔNG dùng Teams hay Email để giao việc chính thức hoặc duyệt chi tiền."),
        ("Microsoft Teams\n(Cộng tác tức thì)",
         "Trao đổi nội bộ nhanh, Họp trực tuyến & Làm việc nhóm",
         "• Chat 1-1, chat nhóm phòng ban giải quyết sự vụ phát sinh.\n"
         "• Tạo Kênh chuyên môn theo dự án (Setup khách sạn mới, Bảo trì cơ sở).\n"
         "• Mở cuộc họp trực tuyến có chia sẻ màn hình tài liệu.",
         "KHÔNG dùng Teams để thay thế việc giao nhiệm vụ chính thức trên HROnline."),
        ("Microsoft Outlook\n(Thư điện tử công vụ)",
         "Giao dịch văn bản chính thức, Đối tác ngoài & Đặt lịch họp",
         "• Gửi thư báo giá, hợp đồng với khách đoàn, đối tác lữ hành, nhà cung cấp.\n"
         "• Ban hành thông báo, quy chế, văn bản điều hành tập đoàn.\n"
         "• Sử dụng Calendar đặt lịch làm việc với Lãnh đạo qua nút New Teams Meeting.",
         "KHÔNG dùng email cá nhân (Gmail, Yahoo) để giao dịch công việc của A25."),
        ("Microsoft OneDrive\n(Lưu trữ đám mây 1TB)",
         "Lưu trữ dữ liệu cá nhân & Chia sẻ tài liệu an toàn",
         "• Đồng bộ dữ liệu máy tính cá nhân lên đám mây dung lượng 1TB.\n"
         "• Chia sẻ link tài liệu phân quyền nội bộ thay vì đính kèm file nặng.\n"
         "• Khôi phục phiên bản lịch sử (Version History) khi chỉnh sửa nhầm.",
         "KHÔNG chọn chế độ 'Anyone with link' (Bất kỳ ai có link) để tránh lộ dữ liệu mật.")
    ]

    for r_idx, (t_name, t_goal, t_use, t_ban) in enumerate(tools_data, start=1):
        row = tbl_tools.rows[r_idx]
        for c_idx, text_val in enumerate((t_name, t_goal, t_use, t_ban)):
            c = row.cells[c_idx]
            c.width = col_w[c_idx]
            c.text = text_val
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            r = p.runs[0]
            r.font.name = "Arial"
            r.font.size = Pt(8.5)
            if c_idx == 0:
                r.bold = True
                if "HROnline" in text_val:
                    r.font.color.rgb = RGBColor(180, 20, 20)
                else:
                    r.font.color.rgb = RGBColor(30, 58, 138)
            elif c_idx == 3:
                r.font.color.rgb = RGBColor(220, 38, 38)
                r.bold = True

            set_cell_background(c, "F9FAFB" if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(c, top=45, bottom=45, left=45, right=45)
            set_cell_border(c, top=('single', '4', '0', 'D1D5DB'),
                               bottom=('single', '4', '0', 'D1D5DB'),
                               left=('single', '4', '0', 'D1D5DB'),
                               right=('single', '4', '0', 'D1D5DB'))

    # FIGURE 5: INFOGRAPHIC RANH GIỚI
    add_illustrated_figure(doc,
        "M365_02_PhanDinh_4_CongCu.png",
        "Hình 5: Ma trận Ranh giới sử dụng 04 công cụ: HROnline, Microsoft Teams, Outlook và OneDrive tại A25 Hotel",
        width_inch=6.4)

    # FIGURE 6: REAL ONEDRIVE SYNC & SHARE
    add_illustrated_figure(doc,
        "real_onedrive_sync_share_2.png",
        "Hình 6: Giao diện Lưu trữ & Chia sẻ tệp an toàn trên OneDrive for Business tích hợp trong Windows File Explorer",
        width_inch=6.4)

    # =========================================================================
    # SECTION IV: DỰ TOÁN BẢN QUYỀN
    # =========================================================================
    add_heading_styled(doc, "IV. DỰ TOÁN CHI PHÍ BẢN QUYỀN MICROSOFT 365 (GIAI ĐOẠN THÍ ĐIỂM VĂN PHÒNG)", 1)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "Nhằm tối ưu hóa chi phí đầu tư công nghệ thông tin cho doanh nghiệp, Phòng CNTT đề xuất phương án phân tầng bản quyền (Tiered Licensing) theo đúng nhu cầu sử dụng thực tế của từng nhóm nhân sự khối Văn phòng:"
    )

    tbl_cost = doc.add_table(rows=4, cols=6)
    tbl_cost.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_headers = ["Nhóm đối tượng", "Gói bản quyền đề xuất", "Quyền lợi và ứng dụng bao gồm", "Số lượng", "Đơn giá dự kiến\n(VNĐ/user/tháng)", "Thành tiền dự kiến\n(VNĐ/tháng)"]
    c_w = [Inches(1.4), Inches(1.3), Inches(2.2), Inches(0.7), Inches(1.2), Inches(1.2)]

    for c_idx, h_text in enumerate(c_headers):
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
        set_cell_margins(cell, top=50, bottom=50, left=45, right=45)
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
        "• Đào tạo & Chuyển giao: Tổ chức 02 buổi đào tạo trực tiếp kết hợp cung cấp Bộ video bài giảng thực tế cho nhân sự Khối Văn phòng.\n"
        "• Áp dụng quy tắc Đóng nhóm Zalo công việc (Nguyên tắc 7 ngày):\n"
        "   + Ngày 1 - 3: Tạo nhóm Teams tương ứng cho từng phòng ban (Kinh doanh, Kế toán, CNTT, HCNS). Gửi thông báo trên nhóm Zalo cũ về kế hoạch chuyển đổi kèm link bài giảng.\n"
        "   + Ngày 4 - 6: Duy trì song song nhưng toàn bộ file tài liệu, kế hoạch, thông báo mới BẮT BUỘC gửi qua Teams/Outlook. Trên Zalo chỉ dùng để nhắc nhở đồng nghiệp kiểm tra Teams.\n"
        "   + Ngày 7: Trưởng phòng ban đăng thông báo chính thức 'Nhóm Zalo này hoàn thành nhiệm vụ và đóng lại từ hôm nay' -> Thu hồi quyền đăng bài hoặc Giải tán nhóm Zalo văn phòng. Mọi phát sinh chuyển toàn bộ sang Microsoft Teams."
    )

    add_heading_styled(doc, "Giai đoạn 2: Đánh giá hiệu quả & Hoàn thiện quy trình (Tháng 01/2027)", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "• Đánh giá mức độ hài lòng của CBNV và tỷ lệ tương tác trên Teams/Outlook.\n"
        "• Tinh chỉnh cấu hình bảo mật, phân quyền truy cập OneDrive/SharePoint và đo lường sự suy giảm rủi ro thất thoát dữ liệu."
    )

    add_heading_styled(doc, "Giai đoạn 3: Nhân rộng toàn hệ thống Chi nhánh Khách sạn (Từ Quý II/2027)", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    p.add_run(
        "• Cấp phát tài khoản Microsoft 365 (gói Basic/F1) cho toàn bộ Quản lý chi nhánh, Giám sát lễ tân, Kế toán trưởng chi nhánh trên toàn quốc.\n"
        "• Đưa Teams thành kênh kết nối chính thức giữa Khối Văn phòng Tập đoàn và từng cơ sở khách sạn thuộc hệ thống A25."
    )

    # =========================================================================
    # SECTION VI: BỘ HỌC LIỆU ĐA PHƯƠNG TIỆN, VIDEO MP4 & THAO TÁC THỰC TẾ
    # =========================================================================
    add_heading_styled(doc, "VI. BỘ HỌC LIỆU ĐA PHƯƠNG TIỆN, VIDEO MP4 ĐÀO TẠO & HƯỚNG DẪN THỰC HÀNH", 1)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "Nhằm giải quyết triệt để yêu cầu của Ban Lãnh đạo và nhu cầu thực tế của Cán bộ nhân viên (tránh tình trạng chỉ cung cấp các đường link lý thuyết khó áp dụng), Phòng CNTT xây dựng Bộ cẩm nang hướng dẫn sử dụng tích hợp đa phương tiện bao gồm: Hình ảnh minh họa chuẩn xác thực từ Microsoft, Video đào tạo HD 720p hoàn chỉnh có thuyết minh giọng AI Neural tự nhiên của Microsoft (vi-VN-NamMinhNeural), Trình phát đa phương tiện HTML5 tương tác chọn chuyên đề, và Hướng dẫn thao tác từng bước (Step-by-Step)."
    )

    # FIGURE 7: TRUNG TÂM VIDEO & TÀI LIỆU ĐÀO TẠO
    add_illustrated_figure(doc,
        "M365_04_TrungTam_Video_DaoTao.png",
        "Hình 7: Trung tâm Video Đào tạo Trực quan & Bộ học liệu hướng dẫn thao tác Microsoft 365 dành cho Nhân sự A25 Hotel",
        width_inch=6.4)

    # CALLOUT ĐÍNH KÈM FILE VIDEO MP4 & HTML HUB
    add_callout(doc,
        "🎬 ĐÍNH KÈM FILE VIDEO HƯỚNG DẪN THỰC TẾ & TRÌNH PHÁT TƯƠNG TÁC (NEURAL AUDIO & PHÂN ROLE):",
        "Phòng CNTT đã kết xuất sẵn file Video đào tạo thực tế chuẩn định dạng MP4 HD 720p và Trình duyệt Video tương tác đa phương tiện đính kèm theo hồ sơ phương án này:\n"
        "• File Video gốc: Video_Huong_Dan_Chuan_Hoa_Microsoft_365_A25_Hotel.mp4 (Thời lượng: 02:05, chuẩn nén H.264/AAC, độ phân giải HD 720p, giọng thuyết minh Microsoft Neural vi-VN-NamMinhNeural chuẩn phát âm tự nhiên, có phụ đề đồng bộ và thao tác trực quan: Liên kết Teams-Outlook, Tra cứu Profile Card & Sơ đồ tổ chức, Phân quyền Owner/Member, Tạo Teams Meeting từ Calendar và Trợ lý xếp lịch Scheduling Assistant).\n"
        "• File Trình phát tương tác: Trung_Tam_Video_Huong_Dan_M365_A25_Hotel.html (Mở bằng bất kỳ trình duyệt Chrome/Edge, có chức năng chọn nhảy từng phân đoạn bài giảng và nút tải video MP4 / tài liệu Word nhanh).",
        border_color="0078D4", bg_color="EFF6FF")

    add_heading_styled(doc, "1. Trung tâm Video bài giảng thao tác thực hành (Video Training Hub)", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "Kịch bản bài giảng trong video được chia làm 05 module logic, dẫn dắt cán bộ nhân viên thực hành trôi chảy qua các tình huống công vụ hằng ngày:"
    )

    # Bảng chi tiết 5 Module Video
    tbl_videos = doc.add_table(rows=6, cols=4)
    tbl_videos.alignment = WD_TABLE_ALIGNMENT.CENTER
    v_headers = ["Module Video", "Thời lượng", "Kịch bản & Thao tác tương tác chi tiết", "Ý nghĩa nghiệp vụ A25"]
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
        ("Module 1:\nGiới thiệu Tổng quan",
         "00:00 - 00:18\n(18.5 giây)",
         "• Nêu bật bối cảnh chuyển đổi công cụ làm việc tại A25 Hotel.\n"
         "• Điểm mặt các rủi ro mất mát dữ liệu và trôi file trên Zalo cá nhân.\n"
         "• Mục tiêu thiết lập môi trường số bảo mật, chuyên nghiệp trên Microsoft 365.",
         "Định hướng nhận thức CBNV về chuyển đổi số"),
        ("Module 2:\nTeams & Phân Role (Entra ID)",
         "00:18 - 00:54\n(35.6 giây)",
         "• Gõ tìm kiếm đồng nghiệp trên thanh Search của Teams.\n"
         "• Bật Thẻ hồ sơ (Profile Card) kiểm tra Chức danh (Role) và Sơ đồ tổ chức (Manager).\n"
         "• Mở mục Quản trị nhóm (Manage Team): Thao tác phân quyền Chủ sở hữu (Owner) cho Lãnh đạo và Thành viên (Member) cho nhân sự tác nghiệp.",
         "Chuẩn hóa cây nhân sự và quyền quản trị nhóm"),
        ("Module 3:\nOutlook: Teams Meeting & Lịch",
         "00:54 - 01:17\n(23.6 giây)",
         "• Chuyển sang thẻ Calendar trên Microsoft Outlook.\n"
         "• Bấm nút 'New Teams Meeting' tích hợp sẵn trên thanh Ribbon để tự động sinh link họp trực tuyến.\n"
         "• Mở 'Scheduling Assistant' (Trợ lý xếp lịch) kiểm tra giờ rảnh của Lãnh đạo và kéo chọn khung giờ tối ưu (14:00 - 15:00) không bị trùng.",
         "Liên kết đồng bộ Teams - Outlook Calendar"),
        ("Module 4:\nLưu trữ OneDrive 1TB",
         "01:17 - 01:40\n(22.7 giây)",
         "• Nhận biết trạng thái đồng bộ trong Windows File Explorer (Đám mây xanh / Tích xanh).\n"
         "• Nhấp chuột phải chọn OneDrive -> Share -> Thiết lập quyền chia sẻ nội bộ 'People in A25 Hotel' (Can view/Can edit).\n"
         "• Khôi phục phiên bản lịch sử (Version History) bảo vệ dữ liệu chống ghi đè.",
         "An toàn dữ liệu, chống mất mát file khi nghỉ việc"),
        ("Module 5:\nRanh giới HROnline & Đóng Zalo",
         "01:40 - 02:05\n(25.0 giây)",
         "• Khẳng định vai trò độc tôn của HROnline trong giao việc, chấm công và duyệt chi.\n"
         "• Teams và Outlook tuyệt đối không chồng chéo với chức năng của HROnline.\n"
         "• Kích hoạt quy tắc 7 ngày chuyển tiếp để đóng vĩnh viễn các nhóm chat Zalo văn phòng.",
         "Kỷ luật số và tránh loạn công cụ công việc")
    ]

    for r_idx, (v_title, v_dur, v_steps, v_src) in enumerate(video_items, start=1):
        cell0 = tbl_videos.rows[r_idx].cells[0]
        cell0.width = v_w[0]
        cell0.text = v_title
        p0 = cell0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.runs[0].bold = True
        p0.runs[0].font.name = "Arial"
        p0.runs[0].font.size = Pt(8.5)
        p0.runs[0].font.color.rgb = RGBColor(180, 20, 20)

        cell1 = tbl_videos.rows[r_idx].cells[1]
        cell1.width = v_w[1]
        cell1.text = v_dur
        p1 = cell1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1.runs[0].font.name = "Arial"
        p1.runs[0].font.size = Pt(8)
        p1.runs[0].bold = True
        p1.runs[0].font.color.rgb = RGBColor(30, 58, 138)

        cell2 = tbl_videos.rows[r_idx].cells[2]
        cell2.width = v_w[2]
        cell2.text = v_steps
        p2 = cell2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p2.runs[0].font.name = "Arial"
        p2.runs[0].font.size = Pt(8)

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
    add_heading_styled(doc, "a. Hướng dẫn sử dụng Microsoft Teams: Tra cứu Role, Chat và Quản trị Nhóm", 3)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "• Bước 1: Khởi động và đăng nhập ứng dụng Teams bằng tài khoản định danh: ten.ho@a25hotel.com.\n"
        "• Bước 2 (Tra cứu Profile Card & Phân Role): Gõ tên đồng nghiệp trên thanh tìm kiếm phía trên cùng -> Rê chuột vào ảnh đại diện để mở Thẻ hồ sơ thông minh -> Xem rõ Chức danh chuyên môn (Role), Phòng ban trực thuộc -> Nhấp vào tab 'Organization' để xem sơ đồ cấp Quản lý trực tiếp (Manager).\n"
        "• Bước 3 (Quản trị phân quyền Nhóm - Manage Team): Nhấp chuột phải vào tên Kênh/Nhóm -> Chọn 'Manage Team' -> Trong danh sách thành viên, kiểm tra vai trò: Cán bộ quản lý phòng ban được gán quyền 'Owner', nhân sự chuyên môn được gán quyền 'Member'.\n"
        "• Bước 4 (Thảo luận Kênh & Họp nhanh): Vào Kênh nghiệp vụ phòng ban -> Bấm 'Start a post' để tạo chủ đề thảo luận chung. Sử dụng cú pháp @tên_người để thông báo trực tiếp. Khi cần trao đổi trực tiếp, bấm nút 'Meet' ở góc trên bên phải để bắt đầu cuộc họp video có chia sẻ tài liệu."
    )

    # 2.2 Outlook
    add_heading_styled(doc, "b. Hướng dẫn sử dụng Microsoft Outlook: Tạo Teams Meeting & Trợ lý xếp lịch", 3)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "• Bước 1 (Cài đặt chữ ký email chuẩn A25): Mở Outlook -> Chọn File -> Options -> Mail -> Signatures -> Tạo chữ ký mới theo mẫu: Họ tên | Chức danh | Phòng ban | CÔNG TY CỔ PHẦN A25 HOTEL | Số điện thoại di động | Website: a25hotel.com.\n"
        "• Bước 2 (Tạo Teams Meeting trực tiếp từ Calendar): Chuyển sang màn hình Lịch (Calendar) trên Outlook -> Nhấp vào nút '+ New Teams Meeting' màu tím trên thanh công cụ Ribbon -> Màn hình tự động nhúng đường link tham gia phòng họp Teams trực tuyến kèm mã truy cập vào nội dung thư mời.\n"
        "• Bước 3 (Sử dụng Scheduling Assistant chọn giờ không trùng): Nhấp vào thẻ 'Scheduling Assistant' trên thanh công cụ -> Nhập danh sách đại biểu bắt buộc (Required) và tùy chọn (Optional) -> Hệ thống quét lưới lịch làm việc của tất cả mọi người -> Kéo chọn dải thời gian rảnh không bị vướng lịch họp khác (khung giờ khuyến nghị 14:00 - 15:00) -> Bấm 'Send' để gửi giấy mời tự động."
    )

    # 2.3 OneDrive
    add_heading_styled(doc, "c. Hướng dẫn lưu trữ dữ liệu cá nhân & chia sẻ an toàn trên OneDrive for Business", 3)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "• Bước 1 (Đồng bộ thư mục làm việc): Đăng nhập ứng dụng OneDrive trên máy tính. Ứng dụng tự động tạo thư mục 'OneDrive - A25 Hotel' trong File Explorer của Windows. Mọi file lưu trong thư mục này đều được sao lưu liên tục lên đám mây 1TB.\n"
        "• Bước 2 (Nhận biết trạng thái file):\n"
        "   + Biểu tượng Đám mây xanh: File lưu an toàn trên Cloud, không chiếm dung lượng ổ cứng máy tính.\n"
        "   + Biểu tượng Tích xanh lá cây: File đã tải về máy tính và sẵn sàng mở khi offline.\n"
        "   + Biểu tượng Hai mũi tên xoay vòng: File đang trong quá trình đồng bộ lên máy chủ.\n"
        "• Bước 3 (Chia sẻ tài liệu phân quyền đúng chuẩn): Nhấp chuột phải vào file cần gửi -> Chọn 'OneDrive' -> Chọn 'Share' -> Thiết lập quyền: 'People in A25 Hotel' (Người trong công ty) -> Tùy chọn 'Can view' (Chỉ cho xem) hoặc 'Can edit' (Được phép chỉnh sửa) -> Bấm 'Copy Link' để dán vào Teams hoặc Email. TUYỆT ĐỐI KHÔNG chọn 'Anyone with the link' để tránh lộ lọt dữ liệu ra bên ngoài."
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
        ("Microsoft Teams", "Cách quản trị vai trò thành viên nhóm (Manage Team Roles)", "https://support.microsoft.com/vi-vn/office/assign-team-roles-in-microsoft-teams"),
        ("Microsoft Outlook", "Tạo cuộc họp Teams trực tiếp từ Outlook Calendar", "https://support.microsoft.com/vi-vn/office/schedule-a-teams-meeting-from-outlook"),
        ("Microsoft Outlook", "Sử dụng Trợ lý xếp lịch (Scheduling Assistant) trong Outlook", "https://support.microsoft.com/vi-vn/office/use-the-scheduling-assistant-in-outlook"),
        ("Microsoft OneDrive", "Đồng bộ hóa tệp với OneDrive trong Windows File Explorer", "https://support.microsoft.com/vi-vn/office/sync-files-with-onedrive-in-windows"),
        ("Entra ID & Profile", "Tra cứu Thẻ hồ sơ (Profile Card) và Sơ đồ tổ chức", "https://support.microsoft.com/vi-vn/office/profile-cards-in-microsoft-365")
    ]

    for r_idx, (app, desc, url) in enumerate(links_data, start=1):
        c0 = tbl_links.rows[r_idx].cells[0]
        c1 = tbl_links.rows[r_idx].cells[1]
        c2 = tbl_links.rows[r_idx].cells[2]
        c0.width = l_w[0]
        c1.width = l_w[1]
        c2.width = l_w[2]
        c0.text = app
        c1.text = desc
        c2.text = url

        for c_idx, cell in enumerate((c0, c1, c2)):
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            r = p.runs[0]
            r.font.name = "Arial"
            r.font.size = Pt(8.5)
            if c_idx == 0:
                r.bold = True
                r.font.color.rgb = RGBColor(30, 58, 138)
            elif c_idx == 2:
                r.font.color.rgb = RGBColor(37, 99, 235)
                r.underline = True

            set_cell_background(cell, "F8FAFC" if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, top=45, bottom=45, left=45, right=45)
            set_cell_border(cell, top=('single', '4', '0', 'CBD5E1'),
                                  bottom=('single', '4', '0', 'CBD5E1'),
                                  left=('single', '4', '0', 'CBD5E1'),
                                  right=('single', '4', '0', 'CBD5E1'))

    # =========================================================================
    # SECTION VII: KIẾN NGHỊ & ĐỀ XUẤT BAN LÃNH ĐẠO
    # =========================================================================
    add_heading_styled(doc, "VII. KIẾN NGHỊ & ĐỀ XUẤT PHÊ DUYỆT CỦA PHÒNG CÔNG NGHỆ THÔNG TIN", 1)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        "Kính trình Ban Tổng Giám đốc và các Giám đốc Khối xem xét, phê duyệt các nội dung trọng tâm sau:\n"
        "1. Phê duyệt chủ trương chuẩn hóa: Đồng ý cho Phòng CNTT triển khai thí điểm phương án chuyển đổi số công cụ làm việc sang hệ sinh thái Microsoft 365 đối với 40 tài khoản Khối Văn phòng A25 Hotel từ tháng 11/2026.\n"
        "2. Phê duyệt ngân sách dự toán: Chấp thuận hạn mức kinh phí bản quyền dự kiến khoảng 8.600.000 VNĐ/tháng (~ 103.200.000 VNĐ/năm). Giao Phòng CNTT phối hợp Ban Mua sắm/Tài chính thương thảo với các đối tác CSP để có mức giá ưu đãi nhất.\n"
        "3. Ban hành Quy chế sử dụng công cụ: Ban hành Quy định ranh giới sử dụng 04 công cụ (HROnline, Teams, Outlook, OneDrive) và Quy tắc đóng nhóm Zalo văn phòng trong 7 ngày theo hình thức bắt buộc đối với toàn bộ cán bộ nhân viên Khối Văn phòng.\n"
        "4. Giao nhiệm vụ tổ chức triển khai: Giao Ban Phát triển & Quản lý phần mềm (Trưởng ban Nguyễn Văn Sơn, Trợ lý Hoàng Minh Quân) chịu trách nhiệm cấu hình kỹ thuật, phân bổ tài khoản, đồng bộ sơ đồ Entra ID và tổ chức các buổi đào tạo thực hành theo đúng tiến độ đề ra."
    )

    # Chữ ký xác nhận
    p_sign_intro = doc.add_paragraph()
    p_sign_intro.paragraph_format.space_before = Pt(8)
    p_sign_intro.paragraph_format.space_after = Pt(8)

    tbl_sign = doc.add_table(rows=2, cols=3)
    tbl_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    s_w = [Inches(2.3), Inches(2.3), Inches(2.4)]
    s_headers = [
        "NGƯỜI LẬP TỜ TRÌNH\n(Trợ lý Phòng CNTT)",
        "TRƯỞNG BAN PHÁT TRIỂN & QLPM\n(Kiểm duyệt nội dung)",
        "BAN TỔNG GIÁM ĐỐC\n(Phê duyệt chủ trương)"
    ]

    for c_idx, s_text in enumerate(s_headers):
        c = tbl_sign.rows[0].cells[c_idx]
        c.width = s_w[c_idx]
        c.text = s_text
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.bold = True
            r.font.name = "Arial"
            r.font.size = Pt(9.5)
            r.font.color.rgb = RGBColor(30, 58, 138)
        set_cell_border(c)
        set_cell_margins(c, top=40, bottom=40, left=40, right=40)

    sign_names = [
        "\n\n\n\nHoàng Minh Quân",
        "\n\n\n\nNguyễn Văn Sơn",
        "\n\n\n\n(Ký và ghi rõ họ tên)"
    ]
    for c_idx, s_name in enumerate(sign_names):
        c = tbl_sign.rows[1].cells[c_idx]
        c.width = s_w[c_idx]
        c.text = s_name
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.bold = True
            r.font.name = "Arial"
            r.font.size = Pt(10)
        set_cell_border(c)
        set_cell_margins(c, top=40, bottom=40, left=40, right=40)

    # Save document
    doc.save(DOC_PATH)
    print("SUCCESS: Generated full M365 Word Document at:", DOC_PATH)
    print("File size:", os.path.getsize(DOC_PATH), "bytes")

if __name__ == "__main__":
    generate_final_m365_doc()
