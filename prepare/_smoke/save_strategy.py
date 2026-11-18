"""
_smoke/save_strategy.py  --  how does Excel actually write the file?

Watches the probe directory at high frequency while a Save() is in flight and
records every directory mutation (create / rename / delete).

Answers:
  * is the file written in place, or written to a temp file and renamed?
  * is there ever a moment where NO valid file exists?
  * what temp files appear, and are they cleaned up?
"""
import os, sys, time, threading, subprocess, shutil, glob

PROBE_DIR = r"G:\projects\mddt_lock_issue\_smoke\_probe"
SRC       = r"G:\projects\mddt_lock_issue\_smoke\big.xlsx"
TARGET    = os.path.join(PROBE_DIR, "big_probe.xlsx")

os.makedirs(PROBE_DIR, exist_ok=True)

# --- snapshot the directory constantly in a background thread
events = []
stop = threading.Event()

def watch():
    prev = {}
    while not stop.is_set():
        try:
            cur = {}
            for e in os.scandir(PROBE_DIR):
                try:
                    st = e.stat()
                    cur[e.name] = (st.st_size, st.st_mtime_ns)
                except OSError:
                    cur[e.name] = (-1, -1)
            for n, v in cur.items():
                if n not in prev:
                    events.append((time.time(), "CREATE", n, v[0]))
                elif prev[n] != v:
                    events.append((time.time(), "MODIFY", n, v[0]))
            for n in prev:
                if n not in cur:
                    events.append((time.time(), "DELETE", n, None))
            prev = cur
        except OSError:
            pass

t = threading.Thread(target=watch, daemon=True)
t.start()

# --- do the save
shutil.copy2(SRC, TARGET)
time.sleep(0.7)

import win32com.client as w32
events.append((time.time(), "==SAVE_START==", "", None))

xl = w32.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.EnableEvents = False
xl.ScreenUpdating = False
t0 = time.time()
wb = xl.Workbooks.Open(TARGET, 0, False)
t_open = time.time()
wb.Worksheets(1).Range("A1").Value2 = "STRATEGY_PROBE"
t_dirty = time.time()
wb.Save()
t_save = time.time()
events.append((time.time(), "==SAVE_END==", "", None))
wb.Close(False)
xl.Quit()
time.sleep(0.5)

stop.set()
t.join(timeout=2)

print("open   : %.3f s" % (t_open - t0))
print("dirty  : %.3f s" % (t_dirty - t_open))
print("SAVE   : %.3f s" % (t_save - t_dirty))
print()
base = None
for ts, kind, name, size in events:
    if base is None:
        base = ts
    print("%8.3f  %-14s %-34s %s" % (ts - base, kind, name, size))
print()
print("distinct names touched during save:")
sv0 = [i for i, e in enumerate(events) if e[1] == "==SAVE_START=="]
sv1 = [i for i, e in enumerate(events) if e[1] == "==SAVE_END=="]
if sv0 and sv1:
    win = events[sv0[0]:sv1[0]]
    names = sorted({e[2] for e in win if e[2]})
    for n in names:
        print("  ", n)
    print("count:", len(names))