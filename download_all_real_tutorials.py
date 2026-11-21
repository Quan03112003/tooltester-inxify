import subprocess
import os
import glob
import imageio_ffmpeg

ffmpeg_exe = imageio_ffmpeg.get_ffmpeg_exe()
OUT_DIR = r"c:\Users\Admin\Desktop\tester\youtube_extracted_frames"
os.makedirs(OUT_DIR, exist_ok=True)

targets = [
    {
        "name": "outlook_inbox_compose",
        "query": "ytsearch1:Microsoft Outlook beginners tutorial compose email inbox",
        "start": "00:01:30",
        "end": "00:02:00",
        "ss1": "00:00:08",
        "ss2": "00:00:22"
    },
    {
        "name": "teams_overview",
        "query": "ytsearch1:Microsoft Teams beginners tutorial chat interface",
        "start": "00:01:00",
        "end": "00:01:30",
        "ss1": "00:00:08",
        "ss2": "00:00:20"
    },
    {
        "name": "onedrive_sync_share",
        "query": "ytsearch1:OneDrive beginners tutorial share file file explorer",
        "start": "00:01:30",
        "end": "00:02:00",
        "ss1": "00:00:08",
        "ss2": "00:00:20"
    }
]

for t in targets:
    name = t["name"]
    webm_out = os.path.join(OUT_DIR, f"{name}.webm")
    print(f"\nSearching & Downloading for: {name} ({t['query']})...")
    
    cmd_dl = [
        "python", "-m", "yt_dlp",
        "-f", "bestvideo[height<=720]+bestaudio/best[height<=720]",
        "--download-sections", f"*{t['start']}-{t['end']}",
        "--ffmpeg-location", ffmpeg_exe,
        "-o", webm_out,
        "--force-overwrites",
        t["query"]
    ]
    res = subprocess.run(cmd_dl, capture_output=True, text=True)
    print("Download return code:", res.returncode)
    
    # Locate actual downloaded file
    matching = glob.glob(os.path.join(OUT_DIR, f"{name}*"))
    video_file = None
    for m in matching:
        if m.endswith((".webm", ".mp4", ".mkv")):
            video_file = m
            break
            
    if video_file and os.path.exists(video_file):
        print(f"Downloaded video file: {video_file} ({os.path.getsize(video_file)} bytes)")
        f1 = os.path.join(OUT_DIR, f"real_{name}_1.png")
        f2 = os.path.join(OUT_DIR, f"real_{name}_2.png")
        subprocess.run([ffmpeg_exe, "-y", "-ss", t["ss1"], "-i", video_file, "-vframes", "1", f1], capture_output=True)
        subprocess.run([ffmpeg_exe, "-y", "-ss", t["ss2"], "-i", video_file, "-vframes", "1", f2], capture_output=True)
        print("Extracted frames:", os.path.exists(f1), os.path.exists(f2))
    else:
        print("Warning: Could not find video file for", name)

print("\nAll extractions completed!")
