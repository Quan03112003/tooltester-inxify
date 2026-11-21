import subprocess
import os
import time

AUDIO_DIR = r"c:\Users\Admin\Desktop\tester\audio_narration_edge"

missing = {
    "02_teams_outlook_role": (
        "Điểm mấu chốt đầu tiên là liên kết đồng bộ giữa Microsoft Teams và Microsoft Outlook. "
        "Toàn bộ nhân sự A25 Hotel được quản lý danh bạ tập trung trên Entra ID. "
        "Khi mở Thẻ hồ sơ Profile Card trên Teams hoặc Outlook, hệ thống hiển thị chính xác chức danh nghiệp vụ và vai trò. "
        "Bấm vào tab Sơ đồ tổ chức Organization để thấy rõ cán bộ Quản lý trực tiếp Manager và nhân sự trực thuộc. "
        "Trong Microsoft Teams, hệ thống phân định rõ vai trò Chủ sở hữu Owner cho cấp Quản lý và vai trò Thành viên Member cho nhân viên, "
        "đảm bảo an toàn bảo mật và kiểm soát phân quyền dữ liệu chặt chẽ."
    ),
    "05_matrix_zalo": (
        "Cuối cùng là nguyên tắc ranh giới bắt buộc: "
        "HROnline tiếp tục là kênh duy nhất để giao việc, theo dõi hạn chót và phê duyệt hành chính. "
        "Nghiêm cấm tạo quy trình giao việc trên Microsoft 365. "
        "Đồng thời, sau bảy ngày làm quen trên Teams, Trưởng bộ phận bắt buộc ấn định ngày đóng dứt điểm nhóm Zalo công việc văn phòng. "
        "Phòng Công nghệ thông tin chúc toàn thể Cán bộ nhân viên A25 Hotel áp dụng thành công!"
    )
}

for k, text in missing.items():
    mp3 = os.path.join(AUDIO_DIR, f"{k}.mp3")
    for attempt in range(4):
        print(f"Generating {k} (attempt {attempt+1})...")
        cmd = [
            "python", "-m", "edge_tts",
            "--voice", "vi-VN-NamMinhNeural",
            "--text", text,
            "--write-media", mp3
        ]
        res = subprocess.run(cmd, capture_output=True, text=True)
        if res.returncode == 0 and os.path.exists(mp3) and os.path.getsize(mp3) > 1000:
            print(f"SUCCESS for {k}: {os.path.getsize(mp3)} bytes")
            break
        else:
            print(f"Retry {k} in 2 seconds... (err: {res.stderr[:100]})")
            time.sleep(2)
