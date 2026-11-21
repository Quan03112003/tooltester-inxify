import os
import subprocess
from gtts import gTTS
import imageio_ffmpeg

ffmpeg_exe = imageio_ffmpeg.get_ffmpeg_exe()
AUDIO_DIR = r"c:\Users\Admin\Desktop\tester\audio_narration"
os.makedirs(AUDIO_DIR, exist_ok=True)

scripts = {
    "01_intro": (
        "Kính chào toàn thể Cán bộ nhân viên Công ty Cổ phần A25 Hotel. "
        "Thực hiện chỉ đạo của Ban Giám đốc, Phòng Công nghệ thông tin triển khai phương án "
        "chuẩn hóa công cụ làm việc trên hệ sinh thái Microsoft 365, "
        "nhằm thay thế việc trao đổi phân tán qua Zalo cá nhân, bảo vệ dữ liệu doanh nghiệp và nâng cao tính chuyên nghiệp."
    ),
    "02_teams": (
        "Đầu tiên là Microsoft Teams, kênh trao đổi nội bộ và họp trực tuyến. "
        "Toàn bộ nhân sự A25 Hotel được đồng bộ trên Sổ địa chỉ toàn cầu. "
        "Quý vị chỉ cần gõ tên đồng nghiệp trên thanh tìm kiếm để tra cứu Thẻ hồ sơ thông minh. "
        "Bấm vào tab Sơ đồ tổ chức để thấy rõ cán bộ Quản lý trực tiếp và đồng nghiệp cùng nhóm. "
        "Mọi trao đổi công việc, gửi tài liệu và họp trực tuyến đều được lưu trữ vĩnh viễn, không lo bị trôi file."
    ),
    "03_outlook": (
        "Tiếp theo là Microsoft Outlook, kênh thư điện tử giao dịch chính thức duy nhất với Khách hàng và Đối tác. "
        "Email gửi đi được chuẩn hóa chữ ký nhận diện thương hiệu A25 Hotel kèm hotline và website. "
        "Đặc biệt, khi cần đặt lịch họp với Ban Lãnh đạo, quý vị mở Lịch công vụ, "
        "sử dụng tính năng Trợ lý xếp lịch Scheduling Assistant để kiểm tra trước khung giờ bận rảnh, "
        "đảm bảo không bị trùng lịch họp."
    ),
    "04_onedrive": (
        "Phân hệ thứ ba là kho lưu trữ đám mây OneDrive dung lượng một terabyte cho mỗi cán bộ nhân viên. "
        "Dữ liệu máy tính được tự động sao lưu an toàn. "
        "Khi chia sẻ tài liệu nội bộ, quý vị nhấp chuột phải, chọn phân quyền chỉ cho người trong A25 Hotel được xem hoặc chỉnh sửa. "
        "Tính năng Lịch sử phiên bản cho phép khôi phục lại file gốc khi bị sửa đè hoặc lưu nhầm."
    ),
    "05_matrix": (
        "Cuối cùng là nguyên tắc ranh giới bắt buộc: "
        "HROnline tiếp tục là kênh duy nhất để giao việc, theo dõi hạn chót và phê duyệt hành chính. "
        "Nghiêm cấm tạo quy trình giao việc trên Microsoft 365. "
        "Đồng thời, sau bảy ngày làm quen trên Teams, Trưởng bộ phận bắt buộc ấn định ngày đóng dứt điểm nhóm Zalo công việc văn phòng. "
        "Phòng Công nghệ thông tin chúc toàn thể Cán bộ nhân viên A25 Hotel áp dụng thành công!"
    )
}

durations = {}
print("Generating audio narration clips via gTTS...")
for key, text in scripts.items():
    mp3_path = os.path.join(AUDIO_DIR, f"{key}.mp3")
    tts = gTTS(text, lang='vi')
    tts.save(mp3_path)
    
    # get duration
    res = subprocess.run([ffmpeg_exe, '-i', mp3_path], stderr=subprocess.PIPE, text=True)
    dur = 0.0
    for line in res.stderr.splitlines():
        if "Duration:" in line:
            # Duration: 00:00:15.20,
            part = line.split("Duration:")[1].split(",")[0].strip()
            h, m, s = part.split(":")
            dur = float(h)*3600 + float(m)*60 + float(s)
            break
    durations[key] = dur
    print(f" - {key}: {dur:.2f} seconds ({mp3_path})")

print("All audio files generated successfully:")
print(durations)
