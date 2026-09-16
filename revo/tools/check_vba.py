#!/usr/bin/env python3
"""Static checks for the REVO VBA modules.

These exist because three separate classes of compile error shipped to a live
workbook before anyone could run them: a statement over VBA's line-continuation
limit, an early-bound reference to a form that does not exist at import time,
and identifiers colliding with reserved words. None of them need Excel to
detect. Run this before handing over modules.

    python3 revo/tools/check_vba.py revo/vba

Exit code is non-zero if anything fails, so it can gate a commit.
"""
import re
import sys
import glob
import os

MAX_CONTINUATIONS = 24  # VBA hard limit per logical statement

# The official VBA reserved words. None of these may be used as an identifier,
# and VBA is case-insensitive, so "dO" collides with "Do" and "eNum" with "Enum".
RESERVED = set(w.lower() for w in """
And As Boolean ByRef Byte ByVal Call Case CBool CByte CCur CDate CDbl CDec CInt CLng Const CSng
CStr Currency CVar CVErr Date Debug Decimal Declare Dim Do Double Each Else ElseIf Empty End
EndIf Enum Eqv Erase Error Event Exit False For Friend Function Get GoSub GoTo If Imp Implements
In Input Integer Is Len Let Like Lock Long Loop LSet Me Mod New Next Not Nothing Null Object On
Open Option Optional Or ParamArray Preserve Print Private Property Public RaiseEvent ReDim Rem
Resume Return RSet Seek Select Set Single Static Stop String Sub Then To True Type TypeOf Until
Variant Wend While With WithEvents Write Xor
""".split())

DECL_KEYWORDS = {'dim', 'static', 'const', 'public', 'private', 'type', 'sub',
                 'function', 'property', 'end', 'friend', 'byval', 'byref',
                 'optional', 'paramarray'}

PROC_START = re.compile(
    r'^(?:public\s+|private\s+|friend\s+|static\s+)*'
    r'(sub|function|property\s+(?:get|let|set))\s+([A-Za-z_]\w*)', re.I)
PROC_END = re.compile(r'^end\s+(sub|function|property)\b', re.I)


def logical_lines(text):
    """Join VBA line continuations into single logical lines."""
    out, buf = [], ''
    for line in text.split('\n'):
        stripped = line.rstrip()
        if stripped.endswith(' _'):
            buf += stripped[:-1]
            continue
        out.append(buf + line)
        buf = ''
    if buf:
        out.append(buf)
    return out


def split_top_level(body):
    """Split on commas that are not inside parentheses."""
    parts, depth, cur = [], 0, ''
    for ch in body:
        if ch == '(':
            depth += 1
        elif ch == ')':
            depth -= 1
        if ch == ',' and depth == 0:
            parts.append(cur)
            cur = ''
        else:
            cur += ch
    parts.append(cur)
    return parts


def check_continuations(path, text, fail):
    lines = text.split('\n')
    i = 0
    while i < len(lines):
        if lines[i].rstrip().endswith(' _'):
            start, n = i, 0
            while i < len(lines) and lines[i].rstrip().endswith(' _'):
                n += 1
                i += 1
            if n > MAX_CONTINUATIONS:
                fail(f"{path}:{start+1}: {n} line continuations, VBA allows "
                     f"{MAX_CONTINUATIONS} - the module will refuse to import")
        i += 1


def check_reserved_words(path, text, fail):
    for i, raw in enumerate(text.split('\n'), 1):
        s = raw.strip()
        if not s or s.startswith("'"):
            continue
        names = []

        m = (re.match(r'^(?:Dim|Static|Const)\s+(.*)$', s, re.I) or
             re.match(r'^(?:Public|Private)\s+'
                      r'(?!Sub|Function|Property|Type|Const|Enum|Declare)(.*)$', s, re.I))
        if m:
            for part in split_top_level(m.group(1).split("'")[0]):
                dm = re.match(r'\s*([A-Za-z_]\w*)', part)
                if dm:
                    names.append(dm.group(1))

        for dm in re.finditer(r'(?:ByVal|ByRef|Optional|ParamArray)\s+([A-Za-z_]\w*)', s, re.I):
            names.append(dm.group(1))

        um = re.match(r'^([A-Za-z_]\w*)\s+As\s+\w', s)   # UDT member
        if um and um.group(1).lower() not in DECL_KEYWORDS:
            names.append(um.group(1))

        for n in names:
            if n.lower() in RESERVED and n.lower() not in DECL_KEYWORDS:
                fail(f"{path}:{i}: '{n}' collides with the VBA reserved word "
                     f"'{n.lower()}' - identifiers are case-insensitive")


def check_blocks_and_handlers(path, text, fail):
    lines = logical_lines(text)
    raw_lines = text.split('\n')
    proc = None
    depth = dict(If=0, For=0, With=0, Do=0, Select=0)
    labels = []

    for i, raw in enumerate(lines, 1):
        s = raw.strip()
        low = s.lower()
        if not s or s.startswith("'"):
            continue

        m = PROC_START.match(low)
        if m and not low.startswith('end '):
            proc = m.group(2)
            depth = dict(If=0, For=0, With=0, Do=0, Select=0)
            labels = []
            continue

        if PROC_END.match(low):
            for k, v in depth.items():
                if v:
                    fail(f"{path}:{i}: {proc}: unbalanced {k} ({v:+d})")
            proc = None
            continue

        if not proc:
            continue

        lm = re.match(r'^([A-Za-z_]\w*):\s*$', s)
        if lm:
            if lm.group(1).lower() in labels:
                fail(f"{path}:{i}: duplicate label '{lm.group(1)}' in {proc}")
            labels.append(lm.group(1).lower())

        if re.match(r'^if\b', low) and ' then' in low and not low.split(' then', 1)[1].strip():
            depth['If'] += 1
        if re.match(r'^end\s+if\b', low):
            depth['If'] -= 1
        if re.match(r'^for\b', low):
            depth['For'] += 1
        if re.match(r'^next\b', low):
            depth['For'] -= 1
        if re.match(r'^with\b', low):
            depth['With'] += 1
        if re.match(r'^end\s+with\b', low):
            depth['With'] -= 1
        if re.match(r'^do\b', low):
            depth['Do'] += 1
        if re.match(r'^loop\b', low):
            depth['Do'] -= 1
        if re.match(r'^select\s+case\b', low):
            depth['Select'] += 1
        if re.match(r'^end\s+select\b', low):
            depth['Select'] -= 1

    # Error handlers: Err must be captured first, and the handler must Resume.
    body = '\n'.join(raw_lines)
    for i, ln in enumerate(raw_lines):
        lm = re.match(r'^([A-Za-z_]\w*):\s*$', ln.strip())
        if not lm:
            continue
        label = lm.group(1)
        if not re.search(r'On Error GoTo\s+' + label + r'\b', body):
            continue
        j, seg = i + 1, []
        while j < len(raw_lines) and not PROC_END.match(raw_lines[j].strip().lower()):
            seg.append(raw_lines[j].strip())
            j += 1
        segment = '\n'.join(seg)
        first = next((x for x in seg if x and not x.startswith("'")), '')
        if 'Err.' in segment and 'Err.' not in first and not first.startswith('Dim'):
            fail(f"{path}:{i+1}: handler '{label}' reads Err after calling something "
                 f"- anything with 'On Error GoTo 0' in it clears Err first")
        if not re.search(r'\bResume\b', segment):
            fail(f"{path}:{i+1}: handler '{label}' falls off End without Resume "
                 f"- the error stays pending and fires in the caller's handler")


def check_early_binding(path, text, fail):
    """A .bas module must not early-bind to a UserForm: the form is created at
    install time, so the reference fails to compile before it exists."""
    if not path.endswith('.bas'):
        return
    for i, ln in enumerate(text.split('\n'), 1):
        if ln.strip().startswith("'"):
            continue
        if re.search(r'\b(?:As|New)\s+frm[A-Za-z_]\w*', ln):
            fail(f"{path}:{i}: early-bound UserForm reference - use "
                 f"VBA.UserForms.Add(\"<name>\") and declare As Object")


def main(argv):
    target = argv[1] if len(argv) > 1 else os.path.join('revo', 'vba')
    files = sorted(glob.glob(os.path.join(target, '*.bas'))) + \
            sorted(glob.glob(os.path.join(target, '*.vb')))
    if not files:
        print(f"no VBA sources under {target}", file=sys.stderr)
        return 2

    failures = []

    def fail(msg):
        failures.append(msg)

    for path in files:
        text = open(path, encoding='utf-8', errors='replace').read()
        check_continuations(path, text, fail)
        check_reserved_words(path, text, fail)
        check_blocks_and_handlers(path, text, fail)
        check_early_binding(path, text, fail)

    if failures:
        print(f"FAIL  {len(failures)} problem(s) across {len(files)} file(s):\n")
        for f in failures:
            print(f"  {f}")
        return 1

    print(f"PASS  {len(files)} file(s) clean: continuation limit, reserved words, "
          f"block balance, error handlers, early binding.")
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
