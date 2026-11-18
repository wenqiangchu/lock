#!/bin/bash
# =============================================================================
#  9-linux-observer.sh   --   T5 / T3  Linux-side observer
#
#  Pure ASCII. Runs on the SAS box (or anything with bash + coreutils).
#
#  WHAT IT DOES
#    Samples a file on the SMB share as fast as it can for a fixed duration and
#    records one CSV row per sample:
#
#      epoch_ms,size,magic_ok,md5,unzip_ok,phase,lockfiles
#
#    It never writes to the observed file. Read only.
#
#  WHY IT MATTERS
#    Only the reader can answer these:
#      T5  while Excel saves, does a reader ever see a torn / truncated file?
#      T3  after Excel is killed mid-save, is the surviving file still valid?
#
#    Interpretation:
#      md5 only ever = OLD or NEW            -> save is atomic for readers  (good)
#      size = MISSING appears                -> rename-swap, reader must retry
#      magic_ok = 0 or unzip_ok = 0          -> TORN READ, reader sees garbage
#      a third md5 value appears             -> TORN READ, worst case
#
#  USAGE
#    ./9-linux-observer.sh <file> [seconds] [interval_ms] [phase_file] [out_csv]
#
#  EXAMPLE
#    ./9-linux-observer.sh /mnt/mddt/_probe/book.xlsx 180 200 \
#        /mnt/mddt/_probe/_phase.txt /tmp/observe.csv
#
#  NOTE
#    The Windows driver writes the current phase name into <phase_file> so the
#    two logs can be aligned without trusting that the clocks agree.
# =============================================================================

set -u

FILE="${1:-}"
DUR="${2:-180}"
IVL="${3:-200}"
PHASE="${4:-}"
OUT="${5:-/tmp/observe_$$.csv}"

if [ -z "$FILE" ]; then
    echo "usage: $0 <file> [seconds] [interval_ms] [phase_file] [out_csv]" >&2
    exit 2
fi

DIR=$(dirname "$FILE")

# interval in seconds, for sleep
IVL_S=$(awk -v ms="$IVL" 'BEGIN { printf "%.3f", ms/1000.0 }')

HAVE_UNZIP=0
if command -v unzip >/dev/null 2>&1; then
    HAVE_UNZIP=1
fi

{
    echo "epoch_ms,size,magic_ok,md5,unzip_ok,phase,lockfiles"
} > "$OUT"

echo "observer: file=$FILE"
echo "observer: duration=${DUR}s interval=${IVL}ms"
echo "observer: phase_file=${PHASE:-<none>}"
echo "observer: out=$OUT"
echo "observer: unzip available=$HAVE_UNZIP"
echo "observer: hit Ctrl-C to stop early"
echo

BAD_MAGIC=0
BAD_UNZIP=0
MISSING=0
SAMPLES=0
T0=$(date +%s)
DEADLINE=$((T0 + DUR))

while :; do
    NOW=$(date +%s)
    if [ "$NOW" -ge "$DEADLINE" ]; then
        break
    fi

    TS=$(date +%s%3N 2>/dev/null)
    if [ -z "$TS" ]; then
        TS=$(( $(date +%s) * 1000 ))
    fi

    PHR="idle"
    if [ -n "$PHASE" ] && [ -f "$PHASE" ]; then
        PHR=$(head -n 1 "$PHASE" 2>/dev/null | tr -d '\r\n' | tr ',' ';')
        [ -z "$PHR" ] && PHR="idle"
    fi

    LK=0
    if [ -d "$DIR" ]; then
        LK=$(ls -1a "$DIR" 2>/dev/null | grep -c '^~\$')
    fi

    if [ ! -e "$FILE" ]; then
        MISSING=$((MISSING + 1))
        SAMPLES=$((SAMPLES + 1))
        echo "$TS,MISSING,,NO_FILE,NO_FILE,$PHR,$LK" >> "$OUT"
        sleep "$IVL_S"
        continue
    fi

    SIZE=$(stat -c %s "$FILE" 2>/dev/null)
    [ -z "$SIZE" ] && SIZE="ERR"

    MAGIC=$(head -c 4 "$FILE" 2>/dev/null | od -An -tx1 2>/dev/null | tr -d ' \n')
    if [ "$MAGIC" = "504b0304" ]; then
        MAGIC_OK=1
    else
        MAGIC_OK=0
        BAD_MAGIC=$((BAD_MAGIC + 1))
    fi

    MD5=$(md5sum "$FILE" 2>/dev/null | cut -d' ' -f1)
    [ -z "$MD5" ] && MD5="ERR"

    UNZ="skip"
    if [ "$HAVE_UNZIP" = "1" ]; then
        if [ $((SAMPLES % 5)) -eq 0 ]; then
            if unzip -tqq "$FILE" >/dev/null 2>&1; then
                UNZ=1
            else
                UNZ=0
                BAD_UNZIP=$((BAD_UNZIP + 1))
            fi
        fi
    else
        UNZ=NA
    fi

    SAMPLES=$((SAMPLES + 1))
    echo "$TS,$SIZE,$MAGIC_OK,$MD5,$UNZ,$PHR,$LK" >> "$OUT"

    sleep "$IVL_S"
done

echo
echo "================= OBSERVER SUMMARY ================="
echo "samples            : $SAMPLES"
echo "missing-file reads : $MISSING"
echo "bad zip magic      : $BAD_MAGIC"
echo "unzip failures     : $BAD_UNZIP"
echo

echo "distinct md5 values seen (value, count):"
tail -n +2 "$OUT" | awk -F, '$4 != "ERR" && $4 != "" { c[$4]++ } END { for (k in c) print "  " k "  " c[k] " " }'
echo

echo "distinct sizes seen:"
tail -n +2 "$OUT" | awk -F, '{ c[$2]++ } END { for (k in c) print "  " k "  " c[k] " " }'
echo

echo "phases seen:"
tail -n +2 "$OUT" | awk -F, '{ c[$6]++ } END { for (k in c) print "  " k "  " c[k] " " }'
echo

echo "---- VERDICT --------------------------------------"
if [ "$BAD_MAGIC" -gt 0 ] || [ "$BAD_UNZIP" -gt 0 ]; then
    echo "TORN READ DETECTED."
    echo "A reader saw a structurally invalid file."
    echo "The tool MUST protect external readers (SAS) or they will fail."
elif [ "$MISSING" -gt 0 ]; then
    echo "File briefly disappeared (rename-swap style save)."
    echo "Readers must retry. Tell the SAS owners."
else
    NMD5=$(tail -n +2 "$OUT" | awk -F, '$4 != "ERR" && $4 != "" { print $4 }' | sort -u | wc -l)
    if [ "$NMD5" -le 2 ]; then
        echo "CLEAN: reader only ever saw whole, valid files"
        echo "($NMD5 distinct content hash(es) = old and/or new)."
        echo "External readers are safe during a normal save."
    else
        echo "SUSPICIOUS: $NMD5 distinct content hashes."
        echo "More than old+new means intermediate states were visible."
    fi
fi
echo "==================================================="
echo
echo "csv: $OUT"