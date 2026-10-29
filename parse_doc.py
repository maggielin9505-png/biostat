import re

with open('1022.doc', 'rb') as f:
    blob = f.read()

res = bytearray()
i = 0
n = len(blob)
while i < n:
    if blob[i:i+2] == b"\\'" and i+4 <= n:
        try:
            hex_val = int(blob[i+2:i+4], 16)
            res.append(hex_val)
            i += 4
            continue
        except ValueError:
            pass
    res.append(blob[i])
    i += 1

text = res.decode('cp950', errors='replace')
# remove fonttbl, stylesheet, etc.
text = re.sub(r'\{\\\*?[^{}]+?\}', '', text)
text = re.sub(r'\\[a-zA-Z]+-?\d* ?', ' ', text)
text = re.sub(r'[{}]', '', text)

lines = [l.strip() for l in text.splitlines() if l.strip()]
with open('cleaned_1022_doc.txt', 'w', encoding='utf-8') as out:
    for l in lines:
        out.write(l + '\n')
print(f'Done! Wrote {len(lines)} lines to cleaned_1022_doc.txt')
