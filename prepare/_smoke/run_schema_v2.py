"""Correct driver: mirror run_diffperf.py exactly (Import a .bas, Run a plain name).

The earlier failures were caused by opening the target workbook and then calling
Run("ZZRUNSCHEMA"): the active workbook became the TARGET, so Excel looked for
the macro there. Also AddFromString modules were unreliable. Both are avoided
here: the target path is passed as an argument, and the module is Imported from
a file exactly like the proven run_diffperf.py driver.
"""
import os, time, subprocess, shutil
import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
BAS = os.path.join(ROOT, "collect", "11-test-schema.bas")
WORK = os.path.join(HERE, "schemaworkv2"); XL = os.path.join(WORK, "h.xlsm")
TARGET = os.environ.get("ZZ_TARGET", os.path.join(HERE, "schema_probe.xlsx"))
REPORT = os.path.join(os.environ.get("TEMP", r"C:\Temp"), "mddt_schema.txt")

if os.path.isdir(WORK):
    for _ in range(5):
        try: shutil.rmtree(WORK); break
        except Exception: time.sleep(1)
os.makedirs(WORK, exist_ok=True)
if os.path.exists(REPORT): os.remove(REPORT)

subprocess.run(["powershell","-NoProfile","-Command",
                "Get-Process EXCEL -ErrorAction SilentlyContinue|Stop-Process -Force"],
               capture_output=True)
time.sleep(1)

app = win32.DispatchEx("Excel.Application")
app.Visible = False; app.DisplayAlerts = False
app.EnableEvents = False; app.ScreenUpdating = False; app.AutomationSecurity = 1

try:
    wb = app.Workbooks.Add()
    wb.SaveAs(XL, 52)
    app.Calculation = -4135
    wb.VBProject.VBComponents.Import(BAS)
    print("imported:", [c.Name for c in wb.VBProject.VBComponents])

    # compile probe: a plain name is fine here because the harness is active
    app.Run("ZZCOMPILEPROBE")
    print("compile OK")

    # pass the target as an argument so the harness stays the active workbook
    app.Run("ZZRUNSCHEMA", TARGET)
    print("run returned")
finally:
    for _ in range(3):
        try: app.Quit(); break
        except Exception: time.sleep(0.5)
time.sleep(0.5)
subprocess.run(["powershell","-NoProfile","-Command",
                "Get-Process EXCEL -ErrorAction SilentlyContinue|Stop-Process -Force"],
               capture_output=True)

print("=" * 22, "REPORT", "=" * 22)
print(open(REPORT, encoding="utf-8", errors="replace").read() if os.path.exists(REPORT) else "NO REPORT")
