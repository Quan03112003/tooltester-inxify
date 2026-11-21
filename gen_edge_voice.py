import subprocess
import os
import imageio_ffmpeg

ffmpeg_exe = imageio_ffmpeg.get_ffmpeg_exe()
AUDIO_DIR = r"c:\Users\Admin\Desktop\tester\audio_narration_edge"
os.makedirs(AUDIO_DIR, exist_ok=True)

scripts = {
    "01_intro": (
        "Kính chào toàn thể Cán bộ nhân viên Công ty Cổ phần A25 Hotel. "
        "Thực hiện chỉ đạo của Ban Giám đốc, Phòng Công nghệ thông tin triển khai phương án "
        "chuẩn hóa công cụ làm việc trên hệ sinh thái Microsoft 365, "
        "nhằm thay thế việc trao đổi phân tán qua Zalo cá nhân, bảo vệ dữ liệu doanh nghiệp và nâng cao hiệu suất công việc."
    ),
    "02_teams_outlook_role": (
        "Điểm mấu chốt đầu tiên là liên kết đồng bộ giữa Microsoft Teams và Microsoft Outlook. "
        "Toàn bộ nhân sự A25 Hotel được quản trị danh bạ tập trung trên Entra ID. "
        "Khi quý vị mở Thẻ hồ sơ Profile Card trên Teams hoặc Outlook, hệ thống hiển thị chính xác chức danh nghiệp vụ và vai trò. "
        "Bấm vào tab Sơ đồ tổ chức Organization để thấy rõ cán bộ Quản lý trực tiếp Manager và nhân sự trực thuộc. "
        "Trong Microsoft Teams, hệ thống phân định rõ vai trò Chủ sở hữu Owner cho cấp Quản lý và vai trò Thành viên Member cho nhân viên, "
        "đảm bảo bảo mật và kiểm soát phân quyền chặt chẽ."
    ),
    "03_outlook_teams_meeting": (
        "Tiếp theo là tính năng liên kết lịch họp trực tiếp giữa Outlook và Teams. "
        "Ngay trên thanh công cụ của Outlook Calendar, quý vị bấm nút Teams Meeting để tạo cuộc họp trực tuyến tức thì. "
        "Đồng thời, sử dụng tính năng Trợ lý xếp lịch Scheduling Assistant để kiểm tra trước dải giờ bận rảnh của Ban Giám đốc, "
        "tự động chọn khung giờ tối ưu mười bốn giờ đến mười lăm giờ để gửi giấy mời họp mà không lo trùng lịch công tác."
    ),
    "04_onedrive_share": (
        "Phân hệ thứ ba là kho lưu trữ đám mây OneDrive dung lượng một terabyte cho mỗi nhân viên. "
        "Dữ liệu máy tính được tự động đồng bộ an toàn. "
        "Khi chia sẻ tài liệu nội bộ, quý vị nhấp chuột phải, chọn phân quyền chỉ cho người trong A25 Hotel được xem hoặc chỉnh sửa. "
        "Tính năng Lịch sử phiên bản Version History cho phép khôi phục lại file gốc khi bị sửa đè hoặc lưu nhầm."
    ),
    "05_matrix_zalo": (
        "Cuối cùng là nguyên tắc ranh giới bắt buộc: "
        "HROnline tiếp tục là kênh duy nhất để giao việc, theo dõi hạn chót và phê duyệt hành chính. "
        "Nghiêm cấm tạo quy trình giao việc trên Microsoft 365. "
        "Đồng thời, sau bảy ngày làm quen trên Teams, Trưởng bộ phận bắt buộc ấn định ngày đóng dứt điểm nhóm Zalo công việc văn phòng. "
        "Phòng Công nghệ thông tin chúc toàn thể Cán bộ nhân viên A25 Hotel áp dụng thành công!"
    )
}

print("Generating high-quality Neural Vietnamese voice clips via edge-tts...")
durations = {}
for key, text in scripts.items():
    mp3_path = os.path.join(AUDIO_DIR, f"{key}.mp3")
    cmd = [
        "python", "-m", "edge_tts",
        "--voice", "vi-VN-NamMinhNeural",
        "--text", text,
        "--write-media", mp3_path
    ]
    res = subprocess.run(cmd, capture_output=True, text=True)
    if res.returncode != 0:
        print(f"Error for {key}:", res.stderr)
        continue
    
    # get duration
    res_dur = subprocess.run([ffmpeg_exe, "-i", mp3_path], stderr=subprocess.PIPE, text=True)
    dur = 0.0
    for line in res_dur.stderr.splitlines():
        if "Duration:" in line:
            part = line.split("Duration:")[1].split(",")[0].strip()
            h, m, s = part.split(":")
            dur = float(h)*3600 + float(m)*60 + float(s)
            break
    durations[key] = dur
    print(f" [Generated] {key}: {dur:.2f}s -> {mp3_path}")

# Concatenate master audio
list_file = os.path.join(AUDIO_DIR, "files.txt")
with open(list_file, "w", encoding="utf-8") as fp:
    for k in ["01_intro", "02_teams_outlook_role", "03_outlook_teams_meeting", "04_onedrive_share", "05_matrix_zalo"]:
        fp.write(f"file '{k}.mp3'\n")

master_m4a = os.path.join(AUDIO_DIR, "master_narration_edge.m4a")
cmd_cat = [
    ffmpeg_exe, "-y",
    "-f", "concat",
    "-safe", "0",
    "-i", list_file,
    "-c:a", "aac",
    "-b:a", "192k",
    master_m4a
]
res_cat = subprocess.run(cmd_cat, cwd=AUDIO_DIR, capture_output=True, text=True)
print("Concat return code:", res_cat.returncode)

res_tot = subprocess.run([ffmpeg_exe, "-i", master_m4a], capture_output=True, text=True)
for line in res_tot.stderr.splitlines():
    if "Duration:" in line:
        print("Total Master Narration Duration:", line.strip())

print("Durations dictionary:")
print(durations)
