import subprocess
import os
import imageio_ffmpeg

ffmpeg_exe = imageio_ffmpeg.get_ffmpeg_exe()
OUT_DIR = r"c:\Users\Admin\Desktop\tester\youtube_extracted_frames"
os.makedirs(OUT_DIR, exist_ok=True)

# 1. Download short section from Outlook Scheduling Assistant tutorial: n_ah404ttYA
video_id = "n_ah404ttYA"
video_path = os.path.join(OUT_DIR, "outlook_sched.mp4")

cmd_dl = [
    "python", "-m", "yt_dlp",
    "-f", "bestvideo[height<=720]+bestaudio/best[height<=720]",
    "--download-sections", "*00:00:40-00:01:10",
    "--ffmpeg-location", ffmpeg_exe,
    "-o", video_path,
    "--force-overwrites",
    f"https://www.youtube.com/watch?v={video_id}"
]

print("Downloading Outlook Scheduling Assistant clip from YouTube...")
res = subprocess.run(cmd_dl, capture_output=True, text=True)
print("Return code:", res.returncode)

if os.path.exists(video_path):
    print("Downloaded video size:", os.path.getsize(video_path))
    # Extract 3 high-res frames
    frame1 = os.path.join(OUT_DIR, "real_outlook_sched_frame1.png")
    subprocess.run([ffmpeg_exe, "-y", "-ss", "00:00:05", "-i", video_path, "-vframes", "1", frame1])
    frame2 = os.path.join(OUT_DIR, "real_outlook_sched_frame2.png")
    subprocess.run([ffmpeg_exe, "-y", "-ss", "00:00:15", "-i", video_path, "-vframes", "1", frame2])
    print("Extracted frames:", os.path.exists(frame1), os.path.exists(frame2))
else:
    print("Stderr:", res.stderr[:500])
