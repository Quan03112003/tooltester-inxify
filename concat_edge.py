import subprocess
import os
import imageio_ffmpeg

ffmpeg_exe = imageio_ffmpeg.get_ffmpeg_exe()
AUDIO_DIR = r"c:\Users\Admin\Desktop\tester\audio_narration_edge"

order = ["01_intro", "02_teams_outlook_role", "03_outlook_teams_meeting", "04_onedrive_share", "05_matrix_zalo"]
durations = {}

for k in order:
    p = os.path.join(AUDIO_DIR, f"{k}.mp3")
    res = subprocess.run([ffmpeg_exe, "-i", p], stderr=subprocess.PIPE, text=True)
    dur = 0.0
    for line in res.stderr.splitlines():
        if "Duration:" in line:
            part = line.split("Duration:")[1].split(",")[0].strip()
            h, m, s = part.split(":")
            dur = float(h)*3600 + float(m)*60 + float(s)
            break
    durations[k] = dur
    print(k, "->", dur, "seconds")

list_file = os.path.join(AUDIO_DIR, "files.txt")
with open(list_file, "w", encoding="utf-8") as fp:
    for k in order:
        fp.write(f"file '{k}.mp3'\n")

master_m4a = os.path.join(AUDIO_DIR, "master_narration_edge.m4a")
cmd = [
    ffmpeg_exe, "-y",
    "-f", "concat",
    "-safe", "0",
    "-i", list_file,
    "-c:a", "aac",
    "-b:a", "192k",
    master_m4a
]
subprocess.run(cmd, cwd=AUDIO_DIR, capture_output=True)

res_tot = subprocess.run([ffmpeg_exe, "-i", master_m4a], stderr=subprocess.PIPE, text=True)
for line in res_tot.stderr.splitlines():
    if "Duration:" in line:
        print("Total Master Audio:", line.strip())

print("Durations dict:")
print(durations)
