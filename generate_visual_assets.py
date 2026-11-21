import asyncio
import os
from playwright.async_api import async_playwright

OUT_DIR = r"c:\Users\Admin\Desktop\tester\img_m365_guide"
os.makedirs(OUT_DIR, exist_ok=True)

# 1. HTML Infographic: So sánh Zalo vs M365
HTML_01 = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  * { box-sizing: border-box; font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 0; }
  body { background: #f8fafc; padding: 24px; width: 1000px; }
  .card { background: #fff; border-radius: 12px; border: 1px solid #e2e8f0; box-shadow: 0 4px 16px rgba(0,0,0,0.06); padding: 28px; }
  .header { text-align: center; margin-bottom: 24px; border-bottom: 2px solid #e2e8f0; padding-bottom: 16px; }
  .badge { background: #dc2626; color: #fff; font-size: 11px; font-weight: 700; padding: 4px 10px; border-radius: 20px; text-transform: uppercase; letter-spacing: 0.5px; }
  .title { font-size: 22px; font-weight: 800; color: #1e3a8a; margin: 8px 0 4px; }
  .sub { font-size: 13px; color: #64748b; }
  .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-top: 16px; }
  .col { border-radius: 10px; padding: 20px; }
  .col-old { background: #fef2f2; border: 1.5px solid #fecaca; }
  .col-new { background: #eff6ff; border: 1.5px solid #bfdbfe; }
  .col-title { font-size: 16px; font-weight: 700; display: flex; align-items: center; gap: 8px; margin-bottom: 14px; }
  .col-old .col-title { color: #b91c1c; }
  .col-new .col-title { color: #1d4ed8; }
  .item { display: flex; align-items: flex-start; gap: 10px; margin-bottom: 12px; font-size: 13px; line-height: 1.5; }
  .item:last-child { margin-bottom: 0; }
  .icon-bad { color: #dc2626; font-size: 16px; font-weight: bold; flex-shrink: 0; }
  .icon-good { color: #16a34a; font-size: 16px; font-weight: bold; flex-shrink: 0; }
  .footer { margin-top: 20px; background: #f1f5f9; padding: 12px 16px; border-radius: 8px; font-size: 12px; color: #334155; text-align: center; }
</style>
</head>
<body>
<div class="card">
  <div class="header">
    <span class="badge">Infographic So Sánh Quy Trình</span>
    <h1 class="title">CHUYỂN ĐỔI SỐ NỘI BỘ TẠI A25 HOTEL: TỪ ZALO SANG MICROSOFT 365</h1>
    <p class="sub">Giải pháp bảo vệ tài sản số, chuẩn hóa thông tin và nâng cao hiệu suất làm việc toàn chuỗi</p>
  </div>
  <div class="grid">
    <div class="col col-old">
      <div class="col-title">❌ CŨ: NHÓM CHAT ZALO CÁ NHÂN (RỦI RO CAO)</div>
      <div class="item"><span class="icon-bad">✕</span><div><b>File & hình ảnh tự động hết hạn:</b> Sau 1-2 tuần không thể tải lại, mất dấu vết chứng từ và lịch sử trao đổi.</div></div>
      <div class="item"><span class="icon-bad">✕</span><div><b>Nguy cơ lộ lọt dữ liệu mật:</b> Nhân sự nghỉ việc vẫn ở trong nhóm chat; mang theo toàn bộ danh sách khách VIP và giá phòng.</div></div>
      <div class="item"><span class="icon-bad">✕</span><div><b>Phân tán, trôi tin nhắn:</b> Khó tra cứu khi cần đối soát; nhân viên bị quá tải vì tin nhắn cá nhân trộn lẫn công việc.</div></div>
      <div class="item"><span class="icon-bad">✕</span><div><b>Không có quản trị tập trung:</b> Công ty không kiểm soát được tài khoản; không thể thu hồi dữ liệu khi nhân sự nghỉ đột ngột.</div></div>
      <div class="item"><span class="icon-bad">✕</span><div><b>Thiếu tính định danh doanh nghiệp:</b> Giao dịch với khách hàng bằng số điện thoại cá nhân, giảm độ uy tín thương hiệu A25.</div></div>
    </div>
    <div class="col col-new">
      <div class="col-title">✅ MỚI: MICROSOFT 365 DOANH NGHIỆP (AN TOÀN - CHUẨN HÓA)</div>
      <div class="item"><span class="icon-good">✓</span><div><b>Lưu trữ vĩnh viễn trên OneDrive 1TB:</b> Tài liệu, hợp đồng, báo cáo được lưu trên đám mây công ty, không bao giờ bị mất hay hết hạn.</div></div>
      <div class="item"><span class="icon-good">✓</span><div><b>Quản trị vòng đời tài khoản tập trung:</b> Nhân sự nghỉ việc chỉ cần 1 cú click của IT để thu hồi toàn bộ quyền và chuyển giao dữ liệu.</div></div>
      <div class="item"><span class="icon-good">✓</span><div><b>Teams phân kênh theo phòng ban/dự án:</b> Trao đổi mạch lạc theo chủ đề, họp trực tuyến có ghi âm và biên bản tự động.</div></div>
      <div class="item"><span class="icon-good">✓</span><div><b>Danh bạ & Sơ đồ tổ chức động (GAL):</b> Gõ tên là biết ngay chức vụ, phòng ban và cấp quản lý trực tiếp của đồng nghiệp.</div></div>
      <div class="item"><span class="icon-good">✓</span><div><b>Định danh chuyên nghiệp @a25hotel.com:</b> Hòm thư Outlook chuẩn hóa, khẳng định đẳng cấp tập đoàn khách sạn hàng đầu.</div></div>
    </div>
  </div>
  <div class="footer">
    🎯 <b>Mục tiêu chuyển đổi:</b> Chuẩn hóa trước tại Khối Văn phòng trong tháng 11 - 12/2026 $\rightarrow$ Đóng nhóm Zalo công việc dứt điểm $\rightarrow$ Bảo vệ dữ liệu A25 Hotel.
  </div>
</div>
</body>
</html>
"""

# 2. HTML Infographic: Ma trận phân định 4 công cụ
HTML_02 = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  * { box-sizing: border-box; font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 0; }
  body { background: #f8fafc; padding: 24px; width: 1000px; }
  .card { background: #fff; border-radius: 12px; border: 1px solid #e2e8f0; box-shadow: 0 4px 16px rgba(0,0,0,0.06); padding: 28px; }
  .header { text-align: center; margin-bottom: 24px; border-bottom: 2px solid #e2e8f0; padding-bottom: 14px; }
  .badge { background: #1e3a8a; color: #fff; font-size: 11px; font-weight: 700; padding: 4px 10px; border-radius: 20px; text-transform: uppercase; }
  .title { font-size: 22px; font-weight: 800; color: #1e3a8a; margin: 8px 0 4px; }
  .sub { font-size: 13px; color: #64748b; }
  .tools-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-top: 16px; }
  .tool-box { border-radius: 10px; padding: 18px 14px; display: flex; flex-direction: column; border: 1.5px solid #e2e8f0; }
  .tool-teams { background: #f5f3ff; border-color: #ddd6fe; }
  .tool-outlook { background: #eff6ff; border-color: #bfdbfe; }
  .tool-onedrive { background: #f0fdf4; border-color: #bbf7d0; }
  .tool-hronline { background: #fffbeb; border-color: #fde68a; }
  .tool-icon { font-size: 32px; text-align: center; margin-bottom: 8px; }
  .tool-name { font-size: 15px; font-weight: 800; text-align: center; margin-bottom: 6px; }
  .tool-role { font-size: 11px; font-weight: 700; text-align: center; padding: 3px 8px; border-radius: 12px; margin-bottom: 12px; text-transform: uppercase; }
  .tool-teams .tool-name { color: #5b21b6; }
  .tool-teams .tool-role { background: #ede9fe; color: #6d28d9; }
  .tool-outlook .tool-name { color: #1e40af; }
  .tool-outlook .tool-role { background: #dbeafe; color: #1d4ed8; }
  .tool-onedrive .tool-name { color: #166534; }
  .tool-onedrive .tool-role { background: #dcfce7; color: #15803d; }
  .tool-hronline .tool-name { color: #92400e; }
  .tool-hronline .tool-role { background: #fef3c7; color: #b45309; }
  .dos, .donts { font-size: 11.5px; line-height: 1.5; margin-bottom: 8px; }
  .sec-tag { font-weight: 700; font-size: 11px; margin-bottom: 4px; display: block; }
  .dos-tag { color: #15803d; }
  .donts-tag { color: #b91c1c; }
  .banner-prohibit { margin-top: 18px; background: #fef2f2; border: 1.5px solid #f87171; border-radius: 8px; padding: 12px 16px; display: flex; align-items: center; gap: 12px; }
  .prohibit-text { font-size: 12.5px; color: #991b1b; line-height: 1.5; }
</style>
</head>
<body>
<div class="card">
  <div class="header">
    <span class="badge">Ma Trận Ranh Giới Công Cụ</span>
    <h1 class="title">NGUYÊN TẮC 4 CỬA SỔ LÀM VIỆC CHUẨN HÓA TẠI A25 HOTEL</h1>
    <p class="sub">Phân định rõ ràng: Nội dung nào dùng công cụ nào – Chống chồng chéo, chống phân tán điều hành</p>
  </div>
  <div class="tools-grid">
    <div class="tool-box tool-teams">
      <div class="tool-icon">💬</div>
      <div class="tool-name">MICROSOFT TEAMS</div>
      <div class="tool-role">Trao đổi nội bộ nhanh</div>
      <div class="dos">
        <span class="sec-tag dos-tag">✓ ĐƯỢC PHÉP:</span>
        • Chat 1-1 và chat nhóm phòng ban/dự án.<br>
        • Họp trực tuyến (Meeting), chia sẻ màn hình.<br>
        • Thảo luận giải quyết sự vụ hằng ngày.
      </div>
      <div class="donts">
        <span class="sec-tag donts-tag">✕ CẤM:</span>
        • Không gửi văn bản pháp lý/hợp đồng chính thức.<br>
        • Không giao việc chính thức thay HROnline.
      </div>
    </div>
    <div class="tool-box tool-outlook">
      <div class="tool-icon">✉️</div>
      <div class="tool-name">MICROSOFT OUTLOOK</div>
      <div class="tool-role">Thư từ & Lịch công vụ</div>
      <div class="dos">
        <span class="sec-tag dos-tag">✓ ĐƯỢC PHÉP:</span>
        • Giao dịch với đối tác, khách hàng, nhà cung cấp.<br>
        • Thông báo điều hành chính thức từ Ban Lãnh đạo.<br>
        • Mời họp công vụ qua Calendar.
      </div>
      <div class="donts">
        <span class="sec-tag donts-tag">✕ CẤM:</span>
        • Không chat vụn vặt gây spam hòm thư.<br>
        • Không đính kèm file dung lượng lớn (dùng link OneDrive).
      </div>
    </div>
    <div class="tool-box tool-onedrive">
      <div class="tool-icon">☁️</div>
      <div class="tool-name">ONEDRIVE 1TB</div>
      <div class="tool-role">Kho dữ liệu đám mây duy nhất</div>
      <div class="dos">
        <span class="sec-tag dos-tag">✓ ĐƯỢC PHÉP:</span>
        • Lưu trữ toàn bộ file làm việc, báo cáo, mẫu biểu.<br>
        • Chia sẻ link tài liệu phân quyền Xem/Sửa an toàn.<br>
        • Đồng bộ tự động trên máy tính và điện thoại.
      </div>
      <div class="donts">
        <span class="sec-tag donts-tag">✕ CẤM:</span>
        • NGHIÊM CẤM lưu file mật trên ổ cứng cá nhân.<br>
        • Không dùng Google Drive / tài khoản cá nhân.
      </div>
    </div>
    <div class="tool-box tool-hronline">
      <div class="tool-icon">📋</div>
      <div class="tool-name">HRONLINE (HIỆN HỮU)</div>
      <div class="tool-role">Kênh giao việc & Phê duyệt duy nhất</div>
      <div class="dos">
        <span class="sec-tag dos-tag">✓ ĐƯỢC PHÉP:</span>
        • Giao nhiệm vụ, theo dõi tiến độ và deadline.<br>
        • Phê duyệt nghỉ phép, đề xuất, công văn nội bộ.<br>
        • Đánh giá thử việc và nghiệm thu công việc.
      </div>
      <div class="donts">
        <span class="sec-tag donts-tag">✕ CẤM:</span>
        • Không nhận giao việc qua chat Teams hoặc Zalo.<br>
        • Không tự duyệt ngoài hệ thống HROnline.
      </div>
    </div>
  </div>
  <div class="banner-prohibit">
    <div style="font-size: 26px;">⛔</div>
    <div class="prohibit-text">
      <b>ĐIỀU KHOẢN CẤM QUAN TRỌNG:</b> Tuyệt đối không xây dựng chức năng gửi/xử lý yêu cầu công việc trên Microsoft Forms hoặc Power Automate. Toàn bộ quy trình giao việc, phê duyệt và quản lý vòng đời phiếu phải tập trung 100% trên <b>HROnline</b>.
    </div>
  </div>
</div>
</body>
</html>
"""

# 3. HTML Mockup: Profile Card & Org Chart
HTML_03 = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  * { box-sizing: border-box; font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 0; }
  body { background: #f8fafc; padding: 24px; width: 1000px; }
  .card { background: #fff; border-radius: 12px; border: 1px solid #e2e8f0; box-shadow: 0 4px 16px rgba(0,0,0,0.06); padding: 28px; }
  .header { text-align: center; margin-bottom: 20px; border-bottom: 2px solid #e2e8f0; padding-bottom: 12px; }
  .badge { background: #16a34a; color: #fff; font-size: 11px; font-weight: 700; padding: 4px 10px; border-radius: 20px; text-transform: uppercase; }
  .title { font-size: 22px; font-weight: 800; color: #1e3a8a; margin: 8px 0 4px; }
  .sub { font-size: 13px; color: #64748b; }
  .mockup-container { display: grid; grid-template-columns: 360px 1fr; gap: 24px; margin-top: 16px; }
  
  /* Left: Profile Card */
  .profile-card { background: #fff; border: 1px solid #d1d5db; border-radius: 12px; box-shadow: 0 6px 20px rgba(0,0,0,0.08); overflow: hidden; }
  .profile-cover { height: 75px; background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%); }
  .profile-body { padding: 0 18px 18px; position: relative; }
  .avatar { width: 68px; height: 68px; border-radius: 50%; background: #0284c7; color: #fff; font-size: 24px; font-weight: bold; display: flex; align-items: center; justify-content: center; border: 3px solid #fff; position: absolute; top: -34px; box-shadow: 0 2px 8px rgba(0,0,0,0.15); }
  .status-dot { width: 14px; height: 14px; background: #22c55e; border: 2px solid #fff; border-radius: 50%; position: absolute; left: 62px; top: 16px; }
  .emp-name { font-size: 17px; font-weight: 800; color: #111827; margin-top: 42px; }
  .emp-title { font-size: 13px; color: #4b5563; font-weight: 600; margin-bottom: 4px; }
  .emp-dept { font-size: 12px; color: #1e3a8a; font-weight: 600; background: #eff6ff; display: inline-block; padding: 2px 8px; border-radius: 6px; margin-bottom: 12px; }
  .profile-tabs { display: flex; gap: 14px; border-bottom: 1px solid #e5e7eb; padding-bottom: 8px; margin-bottom: 12px; font-size: 12px; font-weight: 600; }
  .tab-active { color: #2563eb; border-bottom: 2px solid #2563eb; padding-bottom: 8px; }
  .info-row { display: flex; align-items: center; gap: 8px; font-size: 12px; color: #374151; margin-bottom: 8px; }
  .info-icon { font-size: 14px; color: #6b7280; width: 16px; text-align: center; }
  .action-btns { display: flex; gap: 8px; margin-top: 14px; }
  .btn-act { flex: 1; padding: 6px; border-radius: 6px; border: 1px solid #d1d5db; background: #f9fafb; font-size: 12px; font-weight: 600; text-align: center; cursor: pointer; }
  .btn-primary { background: #2563eb; color: #fff; border-color: #2563eb; }

  /* Right: Org Chart */
  .org-container { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px; padding: 18px; }
  .org-title { font-size: 14px; font-weight: 700; color: #1e3a8a; margin-bottom: 14px; display: flex; align-items: center; gap: 6px; }
  .node-box { background: #fff; border: 1.5px solid #cbd5e1; border-radius: 8px; padding: 10px 14px; width: 280px; box-shadow: 0 2px 6px rgba(0,0,0,0.04); margin-bottom: 12px; }
  .node-manager { border-color: #3b82f6; background: #eff6ff; }
  .node-current { border-color: #10b981; background: #ecfdf5; box-shadow: 0 0 0 2px #a7f3d0; }
  .node-sub { border-color: #94a3b8; }
  .node-label { font-size: 10px; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px; margin-bottom: 2px; }
  .manager-lbl { color: #1d4ed8; }
  .current-lbl { color: #047857; }
  .sub-lbl { color: #64748b; }
  .node-name { font-size: 13.5px; font-weight: 700; color: #0f172a; }
  .node-role { font-size: 11.5px; color: #475569; }
  .tree-line { width: 2px; height: 16px; background: #cbd5e1; margin-left: 24px; margin-top: -12px; margin-bottom: -2px; }
  .tree-branch { margin-left: 24px; padding-left: 16px; border-left: 2px solid #cbd5e1; }
  .explain-box { margin-top: 14px; font-size: 11.5px; color: #334155; line-height: 1.5; background: #fff; padding: 10px 12px; border-radius: 6px; border: 1px solid #e2e8f0; }
</style>
</head>
<body>
<div class="card">
  <div class="header">
    <span class="badge">Mô Hình Danh Bạ Tập Đoàn</span>
    <h1 class="title">THẺ HỒ SƠ (PROFILE CARD) & CÂY SƠ ĐỒ TỔ CHỨC TRÊN TEAMS & OUTLOOK</h1>
    <p class="sub">Tự động đồng bộ từ Microsoft Entra ID – Tra cứu tức thì danh tính, chức danh và người quản lý</p>
  </div>
  <div class="mockup-container">
    <!-- Left: Profile Card -->
    <div class="profile-card">
      <div class="profile-cover"></div>
      <div class="profile-body">
        <div class="avatar">HQ</div>
        <div class="status-dot"></div>
        <div class="emp-name">Hoàng Minh Quân</div>
        <div class="emp-title">Trợ lý CNTT / Chuyên viên QA</div>
        <div class="emp-dept">Phòng Công nghệ thông tin - Khối VP</div>
        <div class="profile-tabs">
          <div class="tab-active">Tổng quan</div>
          <div style="color:#6b7280">Tổ chức</div>
          <div style="color:#6b7280">Tập tin</div>
        </div>
        <div class="info-row"><span class="info-icon">🆔</span> Mã nhân viên: <b>A25-IT-089</b></div>
        <div class="info-row"><span class="info-icon">✉️</span> Email: <b>quanhm.cntt@a25hotel.com</b></div>
        <div class="info-row"><span class="info-icon">📱</span> Di động: <b>0989.xxx.xxx</b></div>
        <div class="info-row"><span class="info-icon">🏢</span> Đơn vị: <b>Văn phòng Tập đoàn A25 Hotel</b></div>
        <div class="info-row"><span class="info-icon">👤</span> Quản lý trực tiếp: <b>Nguyễn Văn Sơn</b></div>
        <div class="action-btns">
          <div class="btn-act btn-primary">💬 Nhắn tin</div>
          <div class="btn-act">✉️ Gửi Email</div>
          <div class="btn-act">📅 Lịch họp</div>
        </div>
      </div>
    </div>
    <!-- Right: Org Chart -->
    <div class="org-container">
      <div class="org-title">🏢 CÂY SƠ ĐỒ TỔ CHỨC ĐỘNG (ORGANIZATION CHART)</div>
      
      <!-- Manager -->
      <div class="node-box node-manager">
        <div class="node-label manager-lbl">CẤP QUẢN LÝ TRỰC TIẾP (MANAGER)</div>
        <div class="node-name">Nguyễn Văn Sơn</div>
        <div class="node-role">Trưởng ban Phát triển & Quản lý PM</div>
      </div>
      <div class="tree-line"></div>
      
      <!-- Current User -->
      <div class="node-box node-current">
        <div class="node-label current-lbl">NHÂN SỰ ĐANG TRA CỨU</div>
        <div class="node-name">Hoàng Minh Quân</div>
        <div class="node-role">Trợ lý Phòng CNTT / Chuyên viên QA</div>
      </div>
      <div class="tree-line"></div>

      <!-- Subordinates / Peers -->
      <div class="tree-branch">
        <div class="node-box node-sub" style="margin-top: 4px;">
          <div class="node-label sub-lbl">NHÂN SỰ BÁO CÁO / PHỐI HỢP CÙNG NHÓM</div>
          <div class="node-name">Châu Tuấn</div>
          <div class="node-role">Chuyên viên Phát triển phần mềm (Dev)</div>
        </div>
      </div>

      <div class="explain-box">
        💡 <b>Lợi ích vượt trội:</b> Bất kỳ nhân sự nào trong tập đoàn chỉ cần rê chuột vào tên đồng nghiệp là thấy ngay sơ đồ phân cấp quản lý. Nhân viên mới không còn tình trạng không biết ai là sếp, liên hệ việc gì với ai!
      </div>
    </div>
  </div>
</div>
</body>
</html>
"""

# 4. HTML Video Training Cards Hub
HTML_04 = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  * { box-sizing: border-box; font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 0; }
  body { background: #f8fafc; padding: 24px; width: 1000px; }
  .card { background: #fff; border-radius: 12px; border: 1px solid #e2e8f0; box-shadow: 0 4px 16px rgba(0,0,0,0.06); padding: 28px; }
  .header { text-align: center; margin-bottom: 24px; border-bottom: 2px solid #e2e8f0; padding-bottom: 12px; }
  .badge { background: #dc2626; color: #fff; font-size: 11px; font-weight: 700; padding: 4px 10px; border-radius: 20px; text-transform: uppercase; }
  .title { font-size: 22px; font-weight: 800; color: #1e3a8a; margin: 8px 0 4px; }
  .sub { font-size: 13px; color: #64748b; }
  .video-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; margin-top: 14px; }
  .video-item { border: 1px solid #e2e8f0; border-radius: 10px; overflow: hidden; background: #fff; box-shadow: 0 2px 8px rgba(0,0,0,0.04); display: flex; flex-direction: column; }
  .thumb-box { height: 130px; position: relative; display: flex; align-items: center; justify-content: center; color: #fff; font-weight: bold; }
  .th-teams { background: linear-gradient(135deg, #4f46e5 0%, #7c3aed 100%); }
  .th-outlook { background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%); }
  .th-onedrive { background: linear-gradient(135deg, #059669 0%, #047857 100%); }
  .th-gal { background: linear-gradient(135deg, #d97706 0%, #b45309 100%); }
  .play-btn { width: 48px; height: 48px; border-radius: 50%; background: rgba(255,255,255,0.9); color: #dc2626; font-size: 20px; display: flex; align-items: center; justify-content: center; box-shadow: 0 4px 12px rgba(0,0,0,0.25); cursor: pointer; }
  .duration { position: absolute; bottom: 8px; right: 10px; background: rgba(0,0,0,0.7); font-size: 11px; padding: 2px 6px; border-radius: 4px; font-weight: 600; }
  .video-info { padding: 14px 16px; flex: 1; display: flex; flex-direction: column; justify-content: space-between; }
  .v-title { font-size: 14px; font-weight: 700; color: #0f172a; margin-bottom: 4px; line-height: 1.4; }
  .v-desc { font-size: 12px; color: #64748b; line-height: 1.4; margin-bottom: 10px; }
  .v-link { font-size: 11.5px; color: #2563eb; font-weight: 600; display: flex; align-items: center; gap: 4px; text-decoration: none; }
  .notice { margin-top: 18px; background: #eff6ff; border: 1px solid #bfdbfe; border-radius: 8px; padding: 12px 16px; font-size: 12px; color: #1e40af; text-align: center; }
</style>
</head>
<body>
<div class="card">
  <div class="header">
    <span class="badge">Trung Tâm Đào Tạo Trực Quan</span>
    <h1 class="title">BỘ VIDEO & TÀI LIỆU HƯỚNG DẪN THỰC HÀNH NHANH (MICROSOFT 365)</h1>
    <p class="sub">Tài liệu chuẩn hóa chính thức từ Microsoft – Thời lượng 2-4 phút/video – Dễ học, thao tác ngay</p>
  </div>
  <div class="video-grid">
    <!-- Video 1: Teams -->
    <div class="video-item">
      <div class="thumb-box th-teams">
        <div class="play-btn">▶</div>
        <div class="duration">03:15</div>
      </div>
      <div class="video-info">
        <div class="v-title">Video 1: Làm quen Microsoft Teams & Họp trực tuyến</div>
        <div class="v-desc">Hướng dẫn tạo nhóm chat, kênh phòng ban, gọi video, chia sẻ màn hình và đặt lịch họp Teams.</div>
        <a class="v-link" href="#">🔗 Bấm để xem Video hướng dẫn Teams (Microsoft Support) &rarr;</a>
      </div>
    </div>
    <!-- Video 2: Outlook -->
    <div class="video-item">
      <div class="thumb-box th-outlook">
        <div class="play-btn">▶</div>
        <div class="duration">04:20</div>
      </div>
      <div class="video-info">
        <div class="v-title">Video 2: Quản lý Email & Lịch công vụ trên Outlook</div>
        <div class="v-desc">Hướng dẫn gửi thư chuẩn chữ ký A25, quản lý thư mục hòm thư, đặt lịch mời họp phòng ban.</div>
        <a class="v-link" href="#">🔗 Bấm để xem Video hướng dẫn Outlook (Microsoft Support) &rarr;</a>
      </div>
    </div>
    <!-- Video 3: OneDrive -->
    <div class="video-item">
      <div class="thumb-box th-onedrive">
        <div class="play-btn">▶</div>
        <div class="duration">03:45</div>
      </div>
      <div class="video-info">
        <div class="v-title">Video 3: Lưu trữ & Chia sẻ tài liệu an toàn OneDrive 1TB</div>
        <div class="v-desc">Hướng dẫn đồng bộ thư mục làm việc, cấp quyền xem/sửa file cho đồng nghiệp, không sợ mất dữ liệu.</div>
        <a class="v-link" href="#">🔗 Bấm để xem Video hướng dẫn OneDrive (Microsoft Support) &rarr;</a>
      </div>
    </div>
    <!-- Video 4: Profile & GAL -->
    <div class="video-item">
      <div class="thumb-box th-gal">
        <div class="play-btn">▶</div>
        <div class="duration">02:30</div>
      </div>
      <div class="video-info">
        <div class="v-title">Video 4: Tra cứu Danh bạ & Sơ đồ tổ chức trên Teams</div>
        <div class="v-desc">Hướng dẫn tìm kiếm nhanh đồng nghiệp, xem Profile Card, kiểm tra chức danh và cán bộ quản lý.</div>
        <a class="v-link" href="#">🔗 Bấm để xem Video tra cứu Profile Card & Org Chart &rarr;</a>
      </div>
    </div>
  </div>
  <div class="notice">
    🎥 <b>Phương thức đào tạo:</b> Toàn bộ video và slide hướng dẫn thao tác từng bước sẽ được đính kèm vào email bàn giao tài khoản và đăng tải trên kênh Teams chung của Khối Văn phòng để nhân sự tra cứu 24/7.
  </div>
</div>
</body>
</html>
"""

async def render_all_images():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page(viewport={'width': 1050, 'height': 800}, device_scale_factor=2)

        items = [
            ("M365_01_SoSanh_Zalo_vs_M365.png", HTML_01),
            ("M365_02_PhanDinh_4_CongCu.png", HTML_02),
            ("M365_03_ProfileCard_OrgChart.png", HTML_03),
            ("M365_04_TrungTam_Video_DaoTao.png", HTML_04)
        ]

        for filename, html in items:
            await page.set_content(html)
            await page.wait_for_timeout(500)
            card = page.locator('.card')
            out_path = os.path.join(OUT_DIR, filename)
            await card.screenshot(path=out_path)
            print("Rendered:", out_path)

        await browser.close()
        print("ALL VISUAL ASSETS RENDERED SUCCESSFULLY!")

if __name__ == '__main__':
    asyncio.run(render_all_images())
