#!/usr/bin/env python3
"""Đóng gói skill thành dist/van-ban-hanh-chinh.zip để tải lên Claude (Cowork, claude.ai: Skills > Upload).

Trong zip là thư mục van-ban-hanh-chinh/ chứa SKILL.md, scripts/, references/, examples/.
Trang tải skill của Claude giới hạn mô tả 200 ký tự nên SKILL.md trong zip dùng mô tả ngắn SHORT.
"""
import os
import re
import sys
import zipfile

SHORT = ("Soạn, kiểm tra, sửa văn bản hành chính (.docx) đúng thể thức Nghị định 30/2020/NĐ-CP - công văn, "
         "quyết định, tờ trình, báo cáo, giấy mời... Dùng khi cần văn bản chuẩn NĐ30.")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
NAME = "van-ban-hanh-chinh"
SRC = os.path.join(ROOT, "skills", NAME)


def build(out_dir=os.path.join(ROOT, "dist")):
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, NAME + ".zip")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for base, dirs, files in os.walk(SRC):
            dirs[:] = sorted(d for d in dirs if d not in ("__pycache__", ".claude-plugin"))
            for f in sorted(files):
                if f.endswith(".pyc") or f == ".DS_Store":
                    continue
                path = os.path.join(base, f)
                arc = NAME + "/" + os.path.relpath(path, SRC).replace(os.sep, "/")
                if f == "SKILL.md":
                    with open(path, encoding="utf-8") as fh:
                        text = re.sub(r"(?m)^description: .*$", "description: " + SHORT, fh.read(), count=1)
                    z.writestr(arc, text)
                else:
                    z.write(path, arc)
    return out


assert len(SHORT) <= 200, len(SHORT)

if __name__ == "__main__":
    print(build(*sys.argv[1:]))
