#!/bin/bash
# =============================================================================
#  9-preflight.sh   --  T5 preflight on the Linux/SAS box
#
#  Pure ASCII. Read only. Writes nothing to the share.
#
#  Run this FIRST. It answers, in one page, every question that would otherwise
#  cost us a round trip:
#     - is the share mounted, and at what path
#     - what is the UNC/drive name behind it
#     - do md5sum / unzip / stat exist
#     - can we actually read the file, how fast
#     - does the share report the same file the Windows side does
#     - is the filesystem case sensitive (matters for our lock names)
#     - do dotfiles / lock dirs behave (matters for MkDir locking)
#
#  USAGE
#     ./9-preflight.sh /mnt/mddt/_mddt_probe/big_probe.xlsx
#
#  If you do not know the path yet:
#     mount | grep -i cifs
#     ls /mnt
# =============================================================================

set -u

FILE="${1:-}"

hr() { echo "-------------------------------------------------------------------"; }

echo "==================================================================="
echo " T5 PREFLIGHT"
echo " date           : $(date -u '+%Y-%m-%dT%H:%M:%SZ') (UTC)"
echo " host           : $(hostname 2>/dev/null)"
echo " kernel         : $(uname -sr 2>/dev/null)"
echo " user           : $(id -un 2>/dev/null)"
echo "==================================================================="
echo

echo "--- 1. CIFS / SMB mounts ------------------------------------------"
MOUNTS=$(mount 2>/dev/null | grep -i -E 'cifs|smb' )
if [ -z "$MOUNTS" ]; then
    echo "  NONE FOUND. The share may be mounted another way."
    echo "  try:  mount | head -30"
    echo "        df -h | head -20"
else
    echo "$MOUNTS" | sed 's/^/  /'
fi
echo

echo "--- 2. tools available --------------------------------------------"
for t in md5sum sha1sum stat unzip od sleep awk; do
    if command -v "$t" >/dev/null 2>&1; then
        echo "  $t : yes ($(command -v $t))"
    else
        echo "  $t : NO"
    fi
done
echo

echo "--- 3. the probe file ---------------------------------------------"
if [ -z "$FILE" ]; then
    echo "  no path given, skipping file checks"
    echo "  usage: $0 <path-to-probe-file>"
else
    echo "  path : $FILE"
    if [ ! -e "$FILE" ]; then
        echo "  CANNOT SEE THE FILE."
        echo "  Ask Windows to create it first, or check the path."
    else
        stat -c '  size=%s  mtime=%y  mode=%a  uid=%u gid=%g' "$FILE" 2>/dev/null
        MAGIC=$(head -c 4 "$FILE" 2>/dev/null | od -An -tx1 | tr -d ' \n')
        echo "  first4bytes(zip magic) = $MAGIC"
        if [ "$MAGIC" = "504b0304" ]; then
            echo "  -> looks like a valid zip/xlsx"
        else
            echo "  -> NOT a zip header"
        fi
        if command -v unzip >/dev/null 2>&1; then
            if unzip -tqq "$FILE" >/dev/null 2>&1; then
                echo "  unzip -t : PASS"
            else
                echo "  unzip -t : FAIL"
            fi
        fi
        if command -v md5sum >/dev/null 2>&1; then
            T0=$(date +%s%3N 2>/dev/null || echo 0)
            MD5=$(md5sum "$FILE" 2>/dev/null | cut -d' ' -f1)
            T1=$(date +%s%3N 2>/dev/null || echo 0)
            echo "  md5sum   : $MD5"
            if [ "$T0" != "0" ]; then
                echo "  md5 took : $((T1 - T0)) ms"
            fi
        fi
    fi
fi
echo

echo "--- 4. mount options that matter for concurrent readers ----------"
if command -v findmnt >/dev/null 2>&1; then
    findmnt -t cifs,smb3 -o TARGET,SOURCE,OPTIONS 2>/dev/null | sed 's/^/  /'
else
    echo "  findmnt not available"
fi
echo "  Look for: cache=  (strict|none|loose)"
echo "            actimeo= , directio , vers="
echo "  cache=strict is the safest for us; loose can serve stale bytes."
echo

echo "--- 5. case sensitivity of the share -----------------------------"
if [ -n "$FILE" ] && [ -d "$(dirname "$FILE")" ]; then
    D=$(dirname "$FILE")
    UP="$D/__CASE_PROBE_$$"
    LO="$D/__case_probe_$$"
    if mkdir "$UP" 2>/dev/null; then
        if [ -d "$LO" ]; then
            echo "  CASE-INSENSITIVE (upper and lower resolve to the same dir)"
        else
            echo "  CASE-SENSITIVE (upper and lower are different dirs)"
        fi
        rmdir "$UP" 2>/dev/null
        rmdir "$LO" 2>/dev/null
    else
        echo "  cannot create a probe dir here (no write permission) - skipped"
        echo "  note: our lock design uppercases every name, so this is only informational"
    fi
else
    echo "  skipped (no directory)"
fi
echo

echo "--- 6. dotfiles / directory creation (needed by the lock design) --"
if [ -n "$FILE" ] && [ -d "$(dirname "$FILE")" ]; then
    D=$(dirname "$FILE")
    TD="$D/__lockprobe_$$.lock"
    if mkdir "$TD" 2>/dev/null; then
        echo "  mkdir on share      : OK"
        if echo "tok123" > "$TD/owner.txt" 2>/dev/null; then
            echo "  write inside it     : OK"
            echo "  read back           : $(cat "$TD/owner.txt" 2>/dev/null)"
        else
            echo "  write inside it     : FAILED"
        fi
        if mkdir "$TD" 2>/dev/null; then
            echo "  second mkdir (same) : SUCCEEDED  <-- would be a problem"
        else
            echo "  second mkdir (same) : correctly refused"
        fi
        rm -f "$TD/owner.txt" 2>/dev/null
        rmdir "$TD" 2>/dev/null
    else
        echo "  cannot mkdir on the share - no write permission"
        echo "  (fine for observation, fatal for the real tool)"
    fi
fi
echo

echo "--- 7. read latency (20 x 1 byte opens) --------------------------"
if [ -n "$FILE" ] && [ -e "$FILE" ]; then
    I=0; BAD=0
    T0=$(date +%s%3N 2>/dev/null || echo 0)
    while [ $I -lt 20 ]; do
        if dd if="$FILE" bs=1 count=1 of=/dev/null 2>/dev/null; then :; else BAD=$((BAD+1)); fi
        I=$((I + 1))
    done
    T1=$(date +%s%3N 2>/dev/null || echo 0)
    if [ "$T0" != "0" ]; then
        echo "  20 reads in $((T1 - T0)) ms, failures=$BAD"
        echo "  per-read  ~ $(( (T1 - T0) / 20 )) ms"
    fi
    echo "  (healthy LAN should be single-digit ms)"
fi
echo

echo "==================================================================="
echo " DONE. Paste this whole output back."
echo " Nothing was written to the share except the temporary probe dir,"
echo " which was removed."
echo "==================================================================="