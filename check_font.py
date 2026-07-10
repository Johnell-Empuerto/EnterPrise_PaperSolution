import zipfile, re
xl = "old_form.xlsx"
z = zipfile.ZipFile(xl)
c = z.read("xl/styles.xml").decode()
fonts = re.findall(r"<font>.*?</font>", c, re.DOTALL)
for i, f in enumerate(fonts[:4]):
    nm = re.search(r"<name[^>]*val=\"([^\"]+)\"", f)
    sz = re.search(r"<sz[^>]*val=\"([^\"]+)\"", f)
    n = nm.group(1) if nm else "?"
    s = sz.group(1) if sz else "?"
    print(f"Font {i}: name={n} sz={s}")
dcw = re.search(r"defaultColWidth=\"([^\"]+)\"", c)
print(f"defaultColWidth: {dcw.group(1) if dcw else 'NOT SET'}")
