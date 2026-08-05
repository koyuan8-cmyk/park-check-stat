#!/usr/bin/env python3
"""验证 xlsx 是否含 VML 页眉 logo（规格：3.74×1.59cm 左上角，所有 sheet）。"""
import sys, zipfile, re, os

W_PT, H_PT = 106.03, 45.08  # 3.74cm × 1.59cm

def verify_header(xlsx_path):
    ok = True
    with zipfile.ZipFile(xlsx_path) as z:
        names = set(z.namelist())
        # 1. logo 媒体存在
        if not any(n.startswith('xl/media/') and n.endswith(('.jpg', '.jpeg', '.png')) for n in names):
            print(f"  ✗ 缺少 xl/media/logo 图片"); ok = False
        # 2. VML drawing 存在且尺寸正确
        vmls = [n for n in names if re.match(r'xl/drawings/vmlDrawing\d+\.vml$', n)]
        if not vmls:
            print(f"  ✗ 缺少 vmlDrawing 部件"); ok = False
        for v in vmls:
            data = z.read(v).decode('utf-8', 'replace')
            m = re.search(r'width:([\d.]+)pt;height:([\d.]+)pt', data)
            if not m:
                print(f"  ✗ {v} 无尺寸"); ok = False
            else:
                w, h = float(m.group(1)), float(m.group(2))
                if abs(w - W_PT) > 0.1 or abs(h - H_PT) > 0.1:
                    print(f"  ✗ {v} 尺寸 {w}×{h}pt ≠ {W_PT}×{H_PT}pt"); ok = False
                else:
                    print(f"  ✓ {v} 尺寸 {w}×{h}pt")
            if 'margin-left:0;margin-top:0' not in data:
                print(f"  ✗ {v} 非左上角"); ok = False
            if '<x:ClientData' in data:
                print(f"  ✗ {v} 含 ClientData 锚点（与规格不一致）"); ok = False
        # 3. 每个 sheet 都有 legacyDrawingHF 引用
        sheets = [n for n in names if re.match(r'xl/worksheets/sheet\d+\.xml$', n)]
        if not sheets:
            print(f"  ✗ 无 worksheet"); ok = False
        for s in sheets:
            xml = z.read(s).decode('utf-8', 'replace')
            if 'legacyDrawingHF' not in xml:
                print(f"  ✗ {s} 缺 legacyDrawingHF"); ok = False
            else:
                print(f"  ✓ {s} 有 legacyDrawingHF")
        # 4. Content_Types 注册了 vmlDrawing
        ct = z.read('[Content_Types].xml').decode('utf-8', 'replace')
        if 'vmlDrawing' not in ct:
            print(f"  ✗ [Content_Types].xml 未注册 vmlDrawing"); ok = False
    return ok

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python3 verify_header.py <xlsx...>"); sys.exit(1)
    all_ok = True
    for p in sys.argv[1:]:
        print(f"== {os.path.basename(p)} ==")
        all_ok = verify_header(p) and all_ok
    print("✅ 全部通过" if all_ok else "❌ 存在失败")
    sys.exit(0 if all_ok else 1)
