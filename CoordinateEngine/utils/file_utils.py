import os


def ensure_dir(path):
    os.makedirs(path, exist_ok=True)


def col_letter(n):
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def col_num(s):
    n = 0
    for c in s.upper().strip():
        n = n * 26 + (ord(c) - 64)
    return n


def char_to_pt(cw, mdw=7.33):
    return (cw * mdw + 5.0) * 72.0 / 96.0


def parse_range(rng):
    import re
    m = re.match(r"\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)", rng)
    if m:
        return col_num(m.group(1)), int(m.group(2)), col_num(m.group(3)), int(m.group(4))
    m = re.match(r"\$?([A-Z]+)\$?(\d+)$", rng)
    if m:
        c = col_num(m.group(1))
        r = int(m.group(2))
        return c, r, c, r
    return None
