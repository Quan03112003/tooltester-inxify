# -*- coding: utf-8 -*-
import os
import subprocess

base_dir = r"c:\Users\Admin\Desktop\tester\Báo cáo kiểm thử inxify"
docx_path = os.path.join(base_dir, "PCNTT_Báo cáo kiểm thử pms inxify.docx")
pdf_path = os.path.join(base_dir, "PCNTT_Báo cáo kiểm thử pms inxify.pdf")

ps_script = f'''
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {{
    $doc = $word.Documents.Open('{docx_path}')
    $doc.SaveAs([ref]'{pdf_path}', [ref]17)
    $doc.Close([ref]$false)
    Write-Host 'Exported PDF successfully!'
}} catch {{
    Write-Host 'Error:' $_
}} finally {{
    $word.Quit()
}}
'''

ps_file = os.path.join(r"c:\Users\Admin\Desktop\tester", "run_export.ps1")
with open(ps_file, "w", encoding="utf-8") as f:
    f.write(ps_script)

res = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", ps_file], capture_output=True, text=True)
print("STDOUT:", res.stdout)
print("STDERR:", res.stderr)

if os.path.exists(pdf_path):
    print("SUCCESS: PDF created at", pdf_path, "Size:", os.path.getsize(pdf_path))
    # Copy to tester root
    import shutil
    shutil.copy2(pdf_path, os.path.join(r"c:\Users\Admin\Desktop\tester", "PCNTT_Báo cáo kiểm thử pms inxify.pdf"))
