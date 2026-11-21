import subprocess
import imageio_ffmpeg

ffmpeg_exe = imageio_ffmpeg.get_ffmpeg_exe()

list_file = r"c:\Users\Admin\Desktop\tester\audio_narration\files.txt"
with open(list_file, "w", encoding="utf-8") as f:
    for name in ["01_intro", "02_teams", "03_outlook", "04_onedrive", "05_matrix"]:
        f.write(f"file '{name}.mp3'\n")

out_m4a = r"c:\Users\Admin\Desktop\tester\audio_narration\master_narration.m4a"
cmd = [
    ffmpeg_exe, "-y",
    "-f", "concat",
    "-safe", "0",
    "-i", list_file,
    "-c:a", "aac",
    "-b:a", "192k",
    out_m4a
]
res = subprocess.run(cmd, cwd=r"c:\Users\Admin\Desktop\tester\audio_narration", capture_output=True, text=True)
print("Concat return code:", res.returncode)

res2 = subprocess.run([ffmpeg_exe, "-i", out_m4a], capture_output=True, text=True)
for line in res2.stderr.splitlines():
    if "Duration:" in line:
        print("Master audio:", line.strip())
