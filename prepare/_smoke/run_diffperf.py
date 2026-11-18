"""Compile-probe and run driver for 8-test-diffperf.bas.

Strategy
--------
1. Open a fresh Excel instance (hidden).
2. SaveAs a macro-enabled workbook so VBProject is writable.
3. Import the .bas module.
4. Compile it (Application.Run of a compile-only probe is not reliable, so we
   use VBProject-level compile via the host: calling any Public Sub forces
   compilation of the whole module; a syntax error surfaces as a VBA dialog
   unless DisplayAlerts is off, so we read the error via Err in a wrapper).
5. Run ZZRUNALL and dump the report.

Run:  python _smoke/run_diffperf.py
"""

import os
import sys
import time
import shutil
import subprocess

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
BAS = os.environ.get("BAS_FILE", os.path.join(ROOT, "collect", "8-test-diffperf.bas"))
WORK = os.path.join(HERE, "diffperf")
XL = os.path.join(WORK, "harness.xlsm")
REPORT = os.path.join(os.environ.get("TEMP", r"C:\Temp"), os.environ.get("REPORT_FILE", "mddt_t89.txt"))


def log(msg):
    print(msg, flush=True)


def clean():
    if os.path.isdir(WORK):
        for _ in range(5):
            try:
                shutil.rmtree(WORK)
                break
            except Exception:
                time.sleep(1)
    os.makedirs(WORK, exist_ok=True)


def kill_excel():
    subprocess.run(
        ["powershell", "-NoProfile", "-Command",
         "Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force"],
        capture_output=True,
    )


def main():
    clean()
    kill_excel()
    time.sleep(1)

    if os.path.exists(REPORT):
        os.remove(REPORT)

    log("starting Excel ...")
    app = win32.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    app.EnableEvents = False
    app.ScreenUpdating = False
    app.AutomationSecurity = 1
    # Calculation cannot be set before any workbook exists; set it after Add.

    wb = None
    try:
        wb = app.Workbooks.Add()
        wb.SaveAs(XL, 52)  # xlOpenXMLWorkbookMacroEnabled
        log("harness workbook: %s" % XL)

        try:
            app.Calculation = -4135
        except Exception as e:
            log("could not set Calculation: %s" % e)

        comps = wb.VBProject.VBComponents
        mod = comps.Import(BAS)
        log("imported module: %s" % mod.Name)

        # Compile probe: running any Public Sub forces a full module compile.
        # A syntax error surfaces as a COM exception instead of a modal dialog,
        # because DisplayAlerts is off and the call comes from outside Excel.
        log("compile probe ...")
        try:
            app.Run("ZZCOMPILEPROBE")
            log("compile OK")
        except Exception as e:
            log("COMPILE FAILED: %s" % e)
            raise

        entry = os.environ.get("ENTRY", "ZZRUNALL")
        log("running %s ..." % entry)
        t0 = time.time()
        app.Run(entry)
        log("%s returned in %.1f s" % (entry, time.time() - t0))
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass
        del app
        kill_excel()
        time.sleep(1)

    if os.path.exists(REPORT):
        log("\n================ REPORT ================")
        with open(REPORT, "r", encoding="utf-8", errors="replace") as f:
            sys.stdout.write(f.read())
    else:
        log("NO REPORT produced at %s" % REPORT)


if __name__ == "__main__":
    main()