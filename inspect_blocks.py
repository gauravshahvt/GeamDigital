import sys, pymupdf
sys.stdout.reconfigure(encoding='utf-8')

doc = pymupdf.open('uploads/b9707b19-9ff5-425e-bd9d-f0973deb1ea6_BAGMALI-Ward_No-001.pdf')
p = doc[2] # page 3
blocks = p.get_text('blocks')
print(f"Total blocks: {len(blocks)}")
for i, b in enumerate(blocks[:35]):
    txt = b[4].strip().replace('\n', ' // ')
    print(f"Block {i:02d} [{b[0]:.1f}, {b[1]:.1f}, {b[2]:.1f}, {b[3]:.1f}]: {txt}")
