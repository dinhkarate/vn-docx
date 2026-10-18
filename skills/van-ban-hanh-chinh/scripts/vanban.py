#!/usr/bin/env python3
"""vanban.py - soạn và kiểm tra văn bản hành chính theo Nghị định 30/2020/NĐ-CP.

Chỉ dùng thư viện chuẩn Python 3.9+, không cần cài thêm gì.

  python3 vanban.py tao noi-dung.md van-ban.docx [--pdf]   # soạn .docx từ nội dung
  python3 vanban.py kiem-tra van-ban.docx [--json]        # kiểm tra thể thức
  python3 vanban.py sua vao.docx ra.docx                  # sửa lỗi khổ giấy, lề, phông, số trang
  python3 vanban.py mau cong-van                          # in nội dung mẫu của một loại
  python3 vanban.py bai-tap dung.docx sai.docx            # cài sẵn lỗi thường gặp (bài tập cho học viên)
  python3 vanban.py loai                                  # bảng tên loại và chữ viết tắt

Cú pháp noi-dung.md: references/cu-phap.md. Quy định gốc: references/the-thuc-nd30.md.
"""
import argparse
import datetime
import json
import os
import pathlib
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
import zipfile
from xml.sax.saxutils import escape as _esc

VERSION = "0.1.0"
HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(HERE)

# ---------------------------------------------------------------- dữ liệu NĐ 30

# Điều 7 và Phụ lục III NĐ 30/2020: mã -> (tên loại in hoa, chữ viết tắt)
LOAI = {
    "nghi-quyet": ("NGHỊ QUYẾT", "NQ"),
    "quyet-dinh": ("QUYẾT ĐỊNH", "QĐ"),
    "chi-thi": ("CHỈ THỊ", "CT"),
    "quy-che": ("QUY CHẾ", "QC"),
    "quy-dinh": ("QUY ĐỊNH", "QyĐ"),
    "thong-cao": ("THÔNG CÁO", "TC"),
    "thong-bao": ("THÔNG BÁO", "TB"),
    "huong-dan": ("HƯỚNG DẪN", "HD"),
    "chuong-trinh": ("CHƯƠNG TRÌNH", "CTr"),
    "ke-hoach": ("KẾ HOẠCH", "KH"),
    "phuong-an": ("PHƯƠNG ÁN", "PA"),
    "de-an": ("ĐỀ ÁN", "ĐA"),
    "du-an": ("DỰ ÁN", "DA"),
    "bao-cao": ("BÁO CÁO", "BC"),
    "bien-ban": ("BIÊN BẢN", "BB"),
    "to-trinh": ("TỜ TRÌNH", "TTr"),
    "hop-dong": ("HỢP ĐỒNG", "HĐ"),
    "cong-van": ("", ""),
    "cong-dien": ("CÔNG ĐIỆN", "CĐ"),
    "ban-ghi-nho": ("BẢN GHI NHỚ", "BGN"),
    "ban-thoa-thuan": ("BẢN THỎA THUẬN", "BTT"),
    "giay-uy-quyen": ("GIẤY ỦY QUYỀN", "GUQ"),
    "giay-moi": ("GIẤY MỜI", "GM"),
    "giay-gioi-thieu": ("GIẤY GIỚI THIỆU", "GGT"),
    "giay-nghi-phep": ("GIẤY NGHỈ PHÉP", "GNP"),
    "phieu-gui": ("PHIẾU GỬI", "PG"),
    "phieu-chuyen": ("PHIẾU CHUYỂN", "PC"),
    "phieu-bao": ("PHIẾU BÁO", "PB"),
    "thu-cong": ("THƯ CÔNG", ""),
}
QUOC_HIEU = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"
TIEU_NGU = "Độc lập - Tự do - Hạnh phúc"
MO_DAU_DIEU = {"quyet-dinh": "QUYẾT ĐỊNH:", "nghi-quyet": "QUYẾT NGHỊ:"}

# Độ rộng ký tự Times New Roman (phần nghìn em), lấy từ bảng hmtx của font, dùng để kẻ
# đường dưới Tiêu ngữ dài đúng bằng dòng chữ.
_ASCII = "".join(chr(c) for c in range(32, 127)) + "ĐđƠơƯư"
_W = {
    "R": "250 333 408 500 500 833 778 180 333 333 500 564 250 333 250 278 500 500 500 500 500 500 500 500 500 500 "
         "278 278 564 564 564 444 921 722 667 667 722 611 556 722 722 333 389 722 611 889 722 722 556 722 667 556 "
         "611 722 722 944 722 722 611 333 278 333 469 500 333 444 500 444 500 444 333 500 500 278 278 500 278 778 "
         "500 500 500 500 333 389 278 500 500 722 500 500 444 480 200 480 541 722 500 722 529 776 542",
    "B": "250 333 555 500 500 1000 833 278 333 333 500 570 250 333 250 278 500 500 500 500 500 500 500 500 500 500 "
         "333 333 570 570 570 500 930 722 667 722 722 667 611 778 778 389 500 778 667 944 722 778 611 778 722 556 "
         "667 722 722 1000 722 722 667 333 278 333 581 500 333 500 556 444 556 444 333 500 556 278 333 556 278 833 "
         "556 500 556 556 444 389 333 556 500 722 500 500 444 394 220 394 520 722 556 778 554 796 600",
    "I": "250 333 420 500 500 833 778 214 333 333 500 675 250 333 250 278 500 500 500 500 500 500 500 500 500 500 "
         "333 333 675 675 675 500 920 611 611 667 722 611 611 722 722 333 444 667 556 833 667 722 611 722 611 500 "
         "556 722 611 833 611 556 556 389 278 389 422 500 333 500 500 444 500 444 278 500 500 278 278 444 278 722 "
         "500 500 500 500 389 389 278 500 444 667 444 444 389 400 275 400 541 722 500 722 534 754 545",
}
WIDTH = {k: dict(zip(_ASCII, map(int, v.split()))) for k, v in _W.items()}
HORN = {"o": "ơ", "O": "Ơ", "u": "ư", "U": "Ư"}


def text_width(text, size, style="R"):
    """Độ rộng dòng chữ (twip) với phông Times New Roman cỡ `size` pt."""
    table, total = WIDTH[style], 0
    for ch in text:
        w = table.get(ch)
        if w is None:
            nfd = unicodedata.normalize("NFD", ch)
            base = HORN.get(nfd[0], nfd[0]) if "\u031b" in nfd else nfd[0]
            w = table.get(base, 500)
        total += w
    return total * size * 20 / 1000


# ---------------------------------------------------------------- đơn vị, XML

def tw_mm(mm):
    return int(round(mm * 1440 / 25.4))


A4_W, A4_H = 11906, 16838
FONT = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>'
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')


def esc(s):
    return _esc(s, {'"': "&quot;"})


def rpr(size=None, b=False, i=False, u=False):
    out = ""
    if b:
        out += "<w:b/><w:bCs/>"
    if i:
        out += "<w:i/><w:iCs/>"
    if u:
        out += '<w:u w:val="single"/>'
    if size:
        out += f'<w:sz w:val="{int(size * 2)}"/><w:szCs w:val="{int(size * 2)}"/>'
    return out


def run(text, size=None, b=False, i=False, u=False):
    props = rpr(size, b, i, u)
    parts = []
    for k, piece in enumerate(text.split("\n")):
        if k:
            parts.append("<w:br/>")
        for m, seg in enumerate(piece.split("\t")):
            if m:
                parts.append("<w:tab/>")
            if seg:
                parts.append(f'<w:t xml:space="preserve">{esc(seg)}</w:t>')
    return f"<w:r><w:rPr>{props}</w:rPr>{''.join(parts)}</w:r>"


INLINE = re.compile(r"(\*\*[^*]+\*\*|\*[^*\s][^*]*\*)")


def inline(text, size, b=False, i=False):
    """**đậm**, *nghiêng* trong một đoạn."""
    out = []
    for tok in INLINE.split(text):
        if not tok:
            continue
        if tok.startswith("**") and tok.endswith("**") and len(tok) > 4:
            out.append(run(tok[2:-2], size, True, i))
        elif tok.startswith("*") and tok.endswith("*") and len(tok) > 2:
            out.append(run(tok[1:-1], size, b, True))
        else:
            out.append(run(tok, size, b, i))
    return "".join(out)


def para(runs="", jc="both", first=0, left=0, right=0, hanging=0, before=0, after=0, line=None,
         rule="auto", keep_next=False, page_break=False, border=None, size=None):
    ppr = ""
    if keep_next:
        ppr += "<w:keepNext/>"
    if page_break:
        ppr += "<w:pageBreakBefore/>"
    if border:
        ppr += f"<w:pBdr>{border}</w:pBdr>"
    sp = f'<w:spacing w:before="{before}" w:after="{after}"'
    if line:
        sp += f' w:line="{line}" w:lineRule="{rule}"'
    ppr += sp + "/>"
    ind = f'<w:ind w:left="{left}" w:right="{right}"'
    if hanging:
        ind += f' w:hanging="{hanging}"'
    elif first:
        ind += f' w:firstLine="{first}"'
    ppr += ind + "/>"
    ppr += f'<w:jc w:val="{jc}"/>'
    if size:
        ppr += f"<w:rPr>{rpr(size)}</w:rPr>"
    return f"<w:p><w:pPr>{ppr}</w:pPr>{runs}</w:p>"


def hline(length, box):
    """Đường kẻ ngang nét liền dài `length` twip, đặt cân đối trong khoảng rộng `box`."""
    ind = max(0, int((box - length) / 2))
    border = '<w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
    return para("", "center", left=ind, right=ind, before=60, line=160, rule="exact", border=border, size=1)


def cell(content, width, valign="top"):
    return (f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/><w:vAlign w:val="{valign}"/></w:tcPr>'
            f'{content or para()}</w:tc>')


def table(rows, widths, borders=False, layout="fixed"):
    b = "single" if borders else "nil"
    bd = "".join(f'<w:{s} w:val="{b}" w:sz="4" w:space="0" w:color="000000"/>'
                 for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    trs = "".join("<w:tr><w:trPr><w:cantSplit/></w:trPr>" + "".join(cell(c, w) for c, w in zip(r, widths)) + "</w:tr>"
                  for r in rows)
    return (f'<w:tbl><w:tblPr><w:tblW w:w="{sum(widths)}" w:type="dxa"/><w:tblBorders>{bd}</w:tblBorders>'
            f'<w:tblLayout w:type="{layout}"/><w:tblCellMar><w:left w:w="57" w:type="dxa"/>'
            f'<w:right w:w="57" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>{grid}</w:tblGrid>{trs}</w:tbl>')


# ---------------------------------------------------------------- đọc noi-dung.md

def parse_source(text):
    """Front matter dạng `khoa: gia tri` / danh sách `- muc`, rồi phần nội dung."""
    lines = text.lstrip("\ufeff").splitlines()
    meta, key = {}, None
    if not lines or lines[0].strip() != "---":
        raise SystemExit("noi-dung.md phải bắt đầu bằng khối thông tin giữa hai dòng ---")
    try:
        end = next(i for i in range(1, len(lines)) if lines[i].strip() == "---")
    except StopIteration:
        raise SystemExit("thiếu dòng --- đóng khối thông tin")
    for ln in lines[1:end]:
        s = ln.strip()
        if not s or s.startswith("#"):
            continue
        m = re.match(r"^-\s*(.*)$", s)
        if m and key is not None and (ln[:1] in " \t-"):
            if not isinstance(meta.get(key), list):
                meta[key] = [] if not meta.get(key) else [meta[key]]
            meta[key].append(unquote(m.group(1)))
            continue
        k, sep, v = ln.partition(":")
        if not sep:
            raise SystemExit(f"dòng không hợp lệ trong khối thông tin: {ln!r}")
        key = k.strip().lower().replace(" ", "_")
        meta[key] = unquote(v.strip())
    return meta, lines[end + 1:]


def unquote(v):
    if len(v) >= 2 and v[0] == v[-1] and v[0] in "\"'":
        return v[1:-1]
    return v


def as_list(v):
    if v is None or v == "":
        return []
    return v if isinstance(v, list) else [v]


def ngay_thang(meta):
    """'Hà Nội, ngày 05 tháng 3 năm 2026' (PL I mục II.4.b: ngày < 10 và tháng 1, 2 thêm số 0)."""
    place = meta.get("dia_danh") or "…"
    raw = str(meta.get("ngay") or "").strip()
    if not raw:
        year = meta.get("nam") or datetime.date.today().year
        return f"{place}, ngày      tháng      năm {year}"
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", raw) or None
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    else:
        m = re.match(r"^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$", raw)
        if not m:
            raise SystemExit(f"ngay không hợp lệ: {raw!r} (dùng 2026-10-17 hoặc 17/10/2026)")
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    datetime.date(y, mo, d)  # báo lỗi nếu ngày không tồn tại
    day = f"{d:02d}"
    month = f"{mo:02d}" if mo <= 2 else str(mo)
    return f"{place}, ngày {day} tháng {month} năm {y}"


def so_ky_hieu(meta, loai):
    """'Số: 05/QĐ-UBND' (PL I mục II.3.c: số nhỏ hơn 10 thêm số 0)."""
    abbr = LOAI[loai][1]
    kh = str(meta.get("ky_hieu") or "").strip()
    if not kh:
        cq = str(meta.get("viet_tat_co_quan") or "").strip()
        dv = str(meta.get("viet_tat_don_vi") or "").strip()
        if loai == "cong-van":
            kh = "-".join(x for x in (cq, dv) if x) or "…"
        else:
            kh = "-".join(x for x in (abbr, cq) if x) or "…"
    so = str(meta.get("so") or "").strip()
    if so.isdigit():
        so = f"{int(so):02d}"
    return f"Số: {so or '      '}/{kh}"


def end_mark(text):
    """Nội dung văn bản kết thúc bằng './.' như các mẫu Phụ lục III."""
    t = text.rstrip()
    if t.endswith("/."):
        return t
    if t.endswith((".", ";", ":", ",")):
        t = t[:-1]
    return t + "./."


def punct_list(items, last="."):
    """Mỗi dòng kết thúc bằng ';', dòng cuối kết thúc bằng '.'."""
    out = []
    for k, it in enumerate(items):
        it = it.rstrip().rstrip(";.,")
        out.append(it + (last if k == len(items) - 1 else ";"))
    return out


# ---------------------------------------------------------------- dựng văn bản

class Builder:
    def __init__(self, meta, body):
        self.meta, self.body = meta, body
        loai = str(meta.get("loai") or "").strip().lower()
        if loai not in LOAI:
            raise SystemExit(f"loai {loai!r} không có trong danh mục; chạy `vanban.py loai` để xem")
        self.loai = loai
        size = int(str(meta.get("co_chu") or 14))
        if size not in (13, 14):
            raise SystemExit("co_chu phải là 13 hoặc 14 (PL I mục II.6.e)")
        self.s = size                       # nội dung, tên loại, trích yếu, chữ ký, địa danh
        self.s_qh = size - 1                # Quốc hiệu 12-13, đi cặp với Tiêu ngữ 13-14 (chú thích PL I mục V)
        self.s_cq = 13 if size == 14 else 12
        top, bottom, left, right = (float(x) for x in str(meta.get("le") or "20 20 30 15").split())
        if not (20 <= top <= 25 and 20 <= bottom <= 25 and 30 <= left <= 35 and 15 <= right <= 20):
            raise SystemExit("le (trên dưới trái phải, mm) phải trong 20-25, 20-25, 30-35, 15-20 (PL I mục I.3)")
        self.margins = [tw_mm(x) for x in (top, bottom, left, right)]
        self.text_w = A4_W - self.margins[2] - self.margins[3]
        spacing = float(str(meta.get("gian_dong") or 1.15))
        if not 1.0 <= spacing <= 1.5:
            raise SystemExit("gian_dong phải từ 1.0 đến 1.5 (PL I mục II.6.e)")
        self.line = int(round(240 * spacing))
        self.first = tw_mm(float(str(meta.get("lui_dau_dong") or 12.7)))
        right_w = max(5670, int(text_width(QUOC_HIEU, self.s_qh, "B")) + 250)
        self.cols = [self.text_w - right_w, right_w]
        self.sections = []

    # -- phần đầu: ô 1 đến ô 5
    def header(self):
        m, s, cq = self.meta, self.s, self.s_cq
        left, box = "", self.cols[0] - 114
        for name in as_list(m.get("co_quan_chu_quan")):
            for line in split_name(name.upper(), cq, False, box):
                left += para(run(line, cq), "center")
        names = []
        for name in as_list(m.get("co_quan")) or ["TÊN CƠ QUAN, TỔ CHỨC"]:
            names += split_name(name.upper(), cq, True, box)
        for name in names:
            left += para(run(name, cq, b=True), "center")
        longest = max(text_width(n.upper(), cq, "B") for n in names)
        left += hline(min(longest, self.cols[0]) * 0.4, self.cols[0])
        right = para(run(QUOC_HIEU, self.s_qh, b=True), "center")
        right += para(run(TIEU_NGU, s, b=True), "center")
        right += hline(text_width(TIEU_NGU, s, "B"), self.cols[1])
        rows = [[left, right]]
        so = para(run(so_ky_hieu(m, self.loai), 13), "center")
        if self.loai == "cong-van" and m.get("trich_yeu"):
            so += para(run("V/v " + strip_vv(m["trich_yeu"]), cq), "center", before=120)
        for key in ("do_mat", "do_khan", "pham_vi_luu_hanh"):
            if m.get(key):
                box = "".join(f'<w:{x} w:val="single" w:sz="6" w:space="1" w:color="000000"/>'
                              for x in ("top", "left", "bottom", "right"))
                txt = str(m[key]).upper()
                pad = max(0, int((self.cols[0] - text_width(txt, s, "B") - 300) / 2))
                so += para(run(txt, s, b=True), "center", left=pad, right=pad, before=120, border=box)
        place = para(run(ngay_thang(m), s, i=True), "center")
        rows.append([so, place])
        return table(rows, self.cols)

    def title(self):
        m, s, out = self.meta, self.s, ""
        name = m.get("ten_loai") or LOAI[self.loai][0]
        if self.loai == "cong-van":
            return ""
        out += para(run(name.upper(), s, b=True), "center", before=360, keep_next=True)
        ty = as_list(m.get("trich_yeu"))
        for line in ty:
            out += para(run(line, s, b=True), "center", keep_next=True)
        if ty:
            longest = max(text_width(t, s, "B") for t in ty)
            out += hline(min(longest, self.text_w) * 0.4, self.text_w)
        elif m.get("ke_duoi_ten_loai", "co") != "khong":
            out += hline(text_width(name.upper(), s, "B") * 0.6, self.text_w)
        return out

    def recipients(self):
        """Kính gửi (PL I mục II.9.d, ô 9a) hoặc '... điện:' của công điện."""
        m, s = self.meta, self.s
        items = as_list(m.get("kinh_gui"))
        if not items:
            return ""
        label = f"{m['chuc_danh_dien']} điện:" if self.loai == "cong-dien" and m.get("chuc_danh_dien") else "Kính gửi:"
        before = 240
        if len(items) == 1:
            return para(run(f"{label} {items[0].rstrip('.;')}.", s), "center", before=before, after=240)
        lab_w = int(text_width(label + " ", s))
        items = punct_list(items)
        if lab_w > self.text_w * 0.35:  # nhãn dài (công điện): nhãn một dòng, danh sách bên dưới
            out = para(run(label, s), "center", before=before, keep_next=True)
            widest = max(text_width("- " + it, s) for it in items)
            left = max(self.first, int((self.text_w - widest) / 2))
            for k, it in enumerate(items):
                out += para(run("- " + it, s), "left", left=left, hanging=int(text_width("- ", s)),
                            after=240 if k == len(items) - 1 else 0)
            return out
        widest = max(text_width("- " + it, s) for it in items)
        left = max(tw_mm(15), int((self.text_w - lab_w - widest) / 2))
        left = min(left, tw_mm(55))
        out = ""
        for k, it in enumerate(items):
            text = (label + " " if k == 0 else "") + "- " + it
            out += para(run(text, s), "left", left=left + lab_w, hanging=lab_w if k == 0 else 0,
                        before=before if k == 0 else 0, after=240 if k == len(items) - 1 else 0)
        return out

    def preamble(self):
        """Thẩm quyền ban hành, căn cứ (in nghiêng), 'QUYẾT ĐỊNH:' (Mẫu 1.1-1.3)."""
        m, s, out = self.meta, self.s, ""
        if m.get("tham_quyen"):
            out += para(run(str(m["tham_quyen"]).upper(), s, b=True), "center", before=240, after=120)
        cc = as_list(m.get("can_cu"))
        for line in punct_list(cc):
            out += para(run(line, s, i=True), "both", first=self.first, after=120, line=self.line)
        opener = m.get("mo_dau_dieu", MO_DAU_DIEU.get(self.loai, "") if (cc or m.get("tham_quyen")) else "")
        if opener:
            out += para(run(opener, s, b=True), "center", before=120, after=120, keep_next=True)
        return out

    def body_paragraph(self, text, **kw):
        opts = dict(jc="both", first=self.first, after=120, line=self.line)
        opts.update(kw)
        return para(inline(text, self.s, opts.pop("b", False), opts.pop("i", False)), **opts)

    def content(self):
        lines = [ln.rstrip() for ln in self.body]
        # bỏ dòng trống đầu, cuối; đánh dấu đoạn cuối để thêm './.'
        idx = [k for k, ln in enumerate(lines) if ln.strip() and not ln.strip().startswith("[")
               and not ln.strip().startswith("|")]
        stops = [k for k, ln in enumerate(lines) if ln.strip().lower() == "[phụ lục]"] + [len(lines)]
        lasts = set()
        if self.meta.get("ket_thuc", "co") != "khong":
            start = 0
            for stop in stops:  # đoạn cuối của văn bản chính và của từng phụ lục
                part = [k for k in idx if start <= k < stop]
                if part:
                    lasts.add(part[-1])
                start = stop
        out, k, s = [], 0, self.s
        while k < len(lines):
            raw = lines[k]
            t = raw.strip()
            if k in lasts and not re.match(r"^(Phụ lục|PHỤ LỤC)\b", t):
                t = end_mark(t)
            if not t:
                k += 1
                continue
            low = t.lower()
            if low == "[ngắt trang]":
                out.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
            elif low == "[phụ lục]":
                out.append(self.section_break())
                k = self.appendix_head(lines, k + 1, out)
                continue
            elif t.startswith("|"):
                rows = []
                while k < len(lines) and lines[k].strip().startswith("|"):
                    row = [c.strip() for c in lines[k].strip().strip("|").split("|")]
                    if not all(re.fullmatch(r":?-{2,}:?", c) for c in row):
                        rows.append(row)
                    k += 1
                out.append(self.data_table(rows))
                continue
            elif re.match(r"^(Phần|Chương|Mục|Tiểu mục)\s+(thứ\s+\S+|[IVXLCDM]+|\d+)\b", t):
                m = re.match(r"^((?:Phần|Chương|Mục|Tiểu mục)\s+(?:thứ\s+\S+|[IVXLCDM]+|\d+))[.:]?\s*(.*)$", t)
                head, rest = m.group(1), m.group(2)
                out.append(para(run(head, s, b=True), "center", before=240, keep_next=True, line=self.line))
                if not rest and k + 1 < len(lines) and lines[k + 1].strip() and is_upper(lines[k + 1].strip()):
                    k += 1
                    rest = lines[k].strip()
                if rest:
                    out.append(para(run(rest.upper(), s, b=True), "center", after=120, keep_next=True, line=self.line))
            elif re.match(r"^Điều\s+\d+[a-z]?\.", t):
                m = re.match(r"^(Điều\s+\d+[a-z]?\.)\s*(.*)$", t)
                rest = m.group(2)
                if is_heading_text(rest):
                    out.append(self.body_paragraph(t, b=True, keep_next=True))
                else:
                    out.append(para(run(m.group(1) + " ", s, b=True) + inline(rest, s), "both", first=self.first,
                                    after=120, line=self.line))
            elif re.match(r"^[IVXLCDM]+\.\s", t):
                out.append(self.body_paragraph(t, b=True, keep_next=True))
            elif re.match(r"^\d+(\.\d+)*\.\s", t) and is_heading_text(re.sub(r"^\d+(\.\d+)*\.\s*", "", t)):
                out.append(self.body_paragraph(t, b=True, keep_next=True))
            elif is_upper(t) and len(t) <= 60:
                out.append(para(run(t, s, b=True), "center", before=120, after=120, keep_next=True))
            else:
                out.append(self.body_paragraph(t))
            k += 1
        return "".join(out)

    def appendix_head(self, lines, k, out):
        """Phụ lục I / TÊN PHỤ LỤC / (Kèm theo ...) - PL I mục III.1.b, c."""
        s, head = self.s, []
        while k < len(lines) and not lines[k].strip():
            k += 1
        while k < len(lines) and lines[k].strip() and len(head) < 3:
            head.append(lines[k].strip())
            k += 1
        for n, h in enumerate(head):
            if n == 0:
                out.append(para(run(h, 14, b=True), "center", keep_next=True))
            elif h.startswith("("):
                out.append(para(run(h, s, i=True), "center", keep_next=True))
            else:
                out.append(para(run(h, s, b=True), "center", keep_next=True))
        if head:
            out.append(hline(tw_mm(30), self.text_w))
        return k

    def data_table(self, rows):
        n = max(len(r) for r in rows)
        lens = [max(min(max(len(r[i]) if i < len(r) else 0, 4), 50) for r in rows) for i in range(n)]
        widths = [int(self.text_w * x / sum(lens)) for x in lens]
        widths[-1] += self.text_w - sum(widths)
        size = self.s - 1
        body = []
        for ri, r in enumerate(rows):
            r = (r + [""] * n)[:n]
            body.append([para(inline(c, size, b=(ri == 0)), "center" if ri == 0 else "left", before=40, after=40)
                         for c in r])
        return table(body, widths, borders=True) + para(after=120)

    # -- chữ ký và nơi nhận: ô 7, 9b, 12
    def signature(self, prefix=""):
        m, s = self.meta, self.s
        out = ""
        for line in as_list(m.get(prefix + "quyen_han")):
            out += para(run(line.upper(), s, b=True), "center", keep_next=True)
        if m.get(prefix + "chuc_vu"):
            out += para(run(str(m[prefix + "chuc_vu"]).upper(), s, b=True), "center", keep_next=True)
        gap = int(str(m.get("dong_cho_ky") or 4))
        out += "".join(para("", "center", keep_next=True, size=s) for _ in range(gap))
        if m.get(prefix + "nguoi_ky"):
            out += para(run(str(m[prefix + "nguoi_ky"]), s, b=True), "center")
        return out

    def receivers(self):
        m = self.meta
        items = as_list(m.get("noi_nhan"))
        if not items:
            return ""
        out = para(run("Nơi nhận:", 12, b=True, i=True), "left")
        for it in punct_list([x.lstrip("- ").strip() for x in items]):
            out += para(run("- " + it, 11), "left")
        if m.get("ky_hieu_nguoi_soan"):
            out += para(run(str(m["ky_hieu_nguoi_soan"]), 11), "left")
        return out

    def closing(self):
        m = self.meta
        rows = []
        if m.get("trai_chuc_vu") or m.get("trai_nguoi_ky"):  # biên bản: thư ký (trái), chủ trì (phải)
            rows.append([self.signature("trai_"), self.signature()])
            tail = self.receivers()
            return (para(after=120) + table(rows, [self.text_w // 2, self.text_w - self.text_w // 2])
                    + (para(after=120) + tail if tail else ""))
        rows.append([self.receivers(), self.signature()])
        return para(after=120) + table(rows, self.cols)

    def section_break(self, last=False):
        top, bottom, left, right = self.margins
        refs = ('<w:headerReference w:type="default" r:id="rIdHeader"/>'
                '<w:headerReference w:type="first" r:id="rIdHeaderFirst"/>'
                f'<w:footerReference w:type="first" r:id="{"rIdFooterFirst" if not self.sections else "rIdFooterEmpty"}"/>')
        sect = (f'<w:sectPr>{refs}<w:pgSz w:w="{A4_W}" w:h="{A4_H}"/>'
                f'<w:pgMar w:top="{top}" w:right="{right}" w:bottom="{bottom}" w:left="{left}" '
                f'w:header="567" w:footer="567" w:gutter="0"/><w:pgNumType w:start="1"/><w:titlePg/></w:sectPr>')
        self.sections.append(sect)
        return sect if last else f"<w:p><w:pPr>{sect}</w:pPr></w:p>"

    def document(self):
        parts = [self.header(), self.title(), self.recipients(), self.preamble()]
        content = self.content()  # có thể chứa ngắt section của phụ lục
        first_break = content.find("<w:sectPr>")
        if first_break == -1:
            parts += [content, self.closing()]
        else:
            start = content.rfind("<w:p>", 0, first_break)
            parts += [content[:start], self.closing(), content[start:]]
        parts.append(self.section_break(last=True))
        return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{"".join(parts)}</w:body></w:document>'

    def footer_first(self):
        lines = as_list(self.meta.get("dia_chi"))
        if not lines:
            return ftr("")
        border = '<w:top w:val="single" w:sz="6" w:space="1" w:color="000000"/>'
        body = "".join(para(run(t, 11), "left", border=border if k == 0 else None) for k, t in enumerate(lines))
        return ftr(body)


UNIT_WORDS = ("TỈNH", "THÀNH PHỐ", "XÃ", "PHƯỜNG", "ĐẶC KHU", "HUYỆN", "QUẬN", "THỊ XÃ", "THỊ TRẤN")


def split_name(name, size, bold, box):
    """Tên cơ quan dài hơn ô thì xuống dòng; ưu tiên ngắt trước tên đơn vị hành chính
    (ỦY BAN NHÂN DÂN / XÃ AN BÌNH), nếu không thì ngắt cho hai dòng cân nhau."""
    style = "B" if bold else "R"
    if text_width(name, size, style) <= box:
        return [name]
    words = name.split()
    best = None
    for k in range(1, len(words)):
        a, b = " ".join(words[:k]), " ".join(words[k:])
        wa, wb = text_width(a, size, style), text_width(b, size, style)
        if max(wa, wb) > box:
            continue
        unit = any(b.startswith(u + " ") for u in UNIT_WORDS)
        score = (0 if unit else 1, abs(wa - wb))
        if best is None or score < best[0]:
            best = (score, [a, b])
    return best[1] if best else [name]


def strip_vv(text):
    return re.sub(r"^(V/v|Về việc)\s*", "", str(text).strip(), flags=re.I)


def is_upper(t):
    letters = [c for c in t if c.isalpha()]
    return bool(letters) and all(c.isupper() for c in letters)


def is_heading_text(rest):
    """Tiêu đề (điều, khoản) ngắn, không kết thúc bằng dấu câu."""
    rest = rest.strip()
    return bool(rest) and len(rest) <= 120 and not rest.endswith((".", ";", ":", ",", "/."))


def hdr(body):
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {NS}>{body or para()}</w:hdr>'


def ftr(body):
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr {NS}>{body or para()}</w:ftr>'


def page_number_header(size):
    fld = (f'<w:r><w:rPr>{rpr(size)}</w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
           f'<w:r><w:rPr>{rpr(size)}</w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
           f'<w:r><w:rPr>{rpr(size)}</w:rPr><w:fldChar w:fldCharType="separate"/></w:r>'
           f'<w:r><w:rPr>{rpr(size)}</w:rPr><w:t>2</w:t></w:r>'
           f'<w:r><w:rPr>{rpr(size)}</w:rPr><w:fldChar w:fldCharType="end"/></w:r>')
    return hdr(para(fld, "center"))


def styles_xml(size):
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles {NS}><w:docDefaults><w:rPrDefault><w:rPr>{FONT}<w:color w:val="000000"/><w:sz w:val="{size * 2}"/><w:szCs w:val="{size * 2}"/><w:lang w:val="vi-VN" w:eastAsia="vi-VN" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:widowControl/></w:pPr><w:rPr>{FONT}<w:sz w:val="{size * 2}"/><w:szCs w:val="{size * 2}"/></w:rPr></w:style>
<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:semiHidden/></w:style>
</w:styles>'''


SETTINGS = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings {NS}><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat><w:themeFontLang w:val="vi-VN"/></w:settings>'''

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>'''

ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'''

DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rIdHeaderFirst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/><Relationship Id="rIdFooterFirst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/><Relationship Id="rIdFooterEmpty" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/></Relationships>'''


def core_xml(title):
    now = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties '
            'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f'<dc:title>{esc(title)}</dc:title><dc:creator>vn-docx</dc:creator><dc:language>vi-VN</dc:language>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified></cp:coreProperties>')


APP_XML = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties '
           'xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
           f'<Application>vn-docx {VERSION}</Application></Properties>')


def build(src, dst):
    with open(src, encoding="utf-8") as f:
        meta, body = parse_source(f.read())
    b = Builder(meta, body)
    document = b.document()
    title = " ".join(as_list(meta.get("trich_yeu"))) or LOAI[b.loai][0]
    files = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": ROOT_RELS,
        "word/_rels/document.xml.rels": DOC_RELS,
        "word/document.xml": document,
        "word/styles.xml": styles_xml(b.s),
        "word/settings.xml": SETTINGS,
        "word/header1.xml": page_number_header(b.s),
        "word/header2.xml": hdr(""),
        "word/footer1.xml": b.footer_first(),
        "word/footer2.xml": ftr(""),
        "docProps/core.xml": core_xml(title),
        "docProps/app.xml": APP_XML,
    }
    os.makedirs(os.path.dirname(os.path.abspath(dst)), exist_ok=True)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)
    return dst


# ---------------------------------------------------------------- đọc .docx để kiểm tra

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _xml(z, name):
    import xml.etree.ElementTree as ET
    try:
        return ET.fromstring(z.read(name))
    except KeyError:
        return None


def _val(el, tag, attr="val"):
    if el is None:
        return None
    x = el.find(W + tag)
    return None if x is None else x.get(W + attr)


def _on(el, tag):
    """Thuộc tính bật/tắt (w:b, w:i): None nếu không đặt."""
    if el is None:
        return None
    x = el.find(W + tag)
    if x is None:
        return None
    return x.get(W + "val") not in ("0", "false", "off")


class Styles:
    def __init__(self, root):
        self.by_id, self.default_para = {}, None
        self.doc_rpr = None
        if root is None:
            return
        d = root.find(f"{W}docDefaults/{W}rPrDefault/{W}rPr")
        self.doc_rpr = d
        for st in root.findall(W + "style"):
            sid = st.get(W + "styleId")
            self.by_id[sid] = st
            if st.get(W + "type") == "paragraph" and st.get(W + "default") == "1":
                self.default_para = sid

    def chain(self, sid):
        seen = []
        while sid and sid in self.by_id and sid not in seen:
            seen.append(sid)
            base = self.by_id[sid].find(W + "basedOn")
            sid = base.get(W + "val") if base is not None else None
        return [self.by_id[s] for s in seen]

    def rpr_chain(self, para_style, run_style=None):
        out = []
        for st in self.chain(run_style):
            out.append(st.find(W + "rPr"))
        for st in self.chain(para_style or self.default_para):
            out.append(st.find(W + "rPr"))
        out.append(self.doc_rpr)
        return [x for x in out if x is not None]

    def ppr_chain(self, para_style):
        return [st.find(W + "pPr") for st in self.chain(para_style or self.default_para)
                if st.find(W + "pPr") is not None]


def _first(chain, fn):
    for el in chain:
        v = fn(el)
        if v is not None:
            return v
    return None


class Para:
    """Một đoạn với định dạng đã tính theo style (đủ cho việc kiểm tra thể thức)."""

    def __init__(self, p, styles, where):
        self.where = where  # ("body", None) hoặc ("table", (bảng, hàng, cột))
        ppr = p.find(W + "pPr")
        pstyle = _val(ppr, "pStyle")
        pchain = ([ppr] if ppr is not None else []) + styles.ppr_chain(pstyle)
        self.jc = _first(pchain, lambda e: _val(e, "jc")) or "left"
        ind = _first(pchain, lambda e: e.find(W + "ind"))
        self.first_line = int(ind.get(W + "firstLine") or 0) if ind is not None else 0
        if ind is not None and ind.get(W + "hanging"):
            self.first_line = -int(ind.get(W + "hanging"))
        sp = _first(pchain, lambda e: e.find(W + "spacing"))
        self.before = int(sp.get(W + "before") or 0) if sp is not None else 0
        self.after = int(sp.get(W + "after") or 0) if sp is not None else 0
        self.line = int(sp.get(W + "line") or 240) if sp is not None and sp.get(W + "line") else 240
        self.line_rule = (sp.get(W + "lineRule") if sp is not None else None) or "auto"
        self.border = ppr is not None and ppr.find(W + "pBdr") is not None
        self.section_end = ppr is not None and ppr.find(W + "sectPr") is not None
        self.runs, texts = [], []
        for r in p.iter(W + "r"):
            t = "".join((x.text or "") if x.tag == W + "t" else ("\t" if x.tag == W + "tab" else "")
                        for x in r if x.tag in (W + "t", W + "tab"))
            if not t:
                continue
            rp = r.find(W + "rPr")
            chain = ([rp] if rp is not None else []) + styles.rpr_chain(pstyle, _val(rp, "rStyle"))
            fonts = _first(chain, lambda e: e.find(W + "rFonts"))
            font = None
            if fonts is not None:
                font = fonts.get(W + "ascii") or fonts.get(W + "hAnsi")
                if not font and (fonts.get(W + "asciiTheme") or fonts.get(W + "hAnsiTheme")):
                    font = "(font theme)"
            size = _first(chain, lambda e: _val(e, "sz"))
            color = _first(chain, lambda e: _val(e, "color"))
            caps = _first(chain, lambda e: _on(e, "caps"))
            self.runs.append({
                "text": t, "font": font, "size": int(size) / 2 if size else 10,
                "b": bool(_first(chain, lambda e: _on(e, "b"))), "i": bool(_first(chain, lambda e: _on(e, "i"))),
                "u": _first(chain, lambda e: _val(e, "u")) not in (None, "none"), "color": color, "caps": bool(caps),
            })
            texts.append(t.upper() if caps else t)
        self.text = "".join(texts)
        self.has_drawing = any(True for _ in p.iter("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline")) or \
            any(True for _ in p.iter("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}anchor")) or \
            any(True for _ in p.iter("{urn:schemas-microsoft-com:vml}line")) or \
            any(True for _ in p.iter("{urn:schemas-microsoft-com:vml}shape"))

    def visible_runs(self):
        return [r for r in self.runs if r["text"].strip()]

    def all(self, key):
        rs = self.visible_runs()
        return bool(rs) and all(r[key] for r in rs)

    def sizes(self):
        return sorted({r["size"] for r in self.visible_runs()})


def read_docx(path):
    with zipfile.ZipFile(path) as z:
        doc = _xml(z, "word/document.xml")
        if doc is None:
            raise SystemExit(f"{path}: không phải file .docx hợp lệ")
        styles = Styles(_xml(z, "word/styles.xml"))
        rels = _xml(z, "word/_rels/document.xml.rels")
        targets = {}
        if rels is not None:
            for r in rels:
                targets[r.get("Id")] = r.get("Target")
        parts = {name: z.read(name).decode("utf-8", "replace") for name in z.namelist()
                 if re.match(r"word/(header|footer)[^/]*\.xml$", name)}
    body = doc.find(W + "body")
    paras, tables = [], 0
    for el in body:
        if el.tag == W + "p":
            paras.append(Para(el, styles, ("body", None)))
        elif el.tag == W + "tbl":
            tables += 1
            for ri, tr in enumerate(el.findall(W + "tr")):
                for ci, tc in enumerate(tr.findall(W + "tc")):
                    for p in tc.iter(W + "p"):
                        paras.append(Para(p, styles, ("table", (tables, ri, ci))))
    sects = list(body.iter(W + "sectPr"))
    return paras, sects, targets, parts


# ---------------------------------------------------------------- kiểm tra thể thức

def norm(t):
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", t)).strip()


def mm(tw):
    return round(int(tw) * 25.4 / 1440, 1)


class Report:
    def __init__(self):
        self.items = []

    def add(self, level, code, msg, ref, where=None):
        self.items.append({"muc_do": level, "ma": code, "noi_dung": msg, "can_cu": ref, "vi_tri": where})

    def ok(self, code, msg, ref):
        self.add("dat", code, msg, ref)

    def err(self, code, msg, ref, where=None):
        self.add("loi", code, msg, ref, where)

    def warn(self, code, msg, ref, where=None):
        self.add("canh_bao", code, msg, ref, where)


def in_range(sizes, lo, hi):
    return bool(sizes) and all(lo <= s <= hi for s in sizes)


def fmt_sizes(sizes):
    return "/".join(str(int(s) if float(s).is_integer() else s) for s in sizes) or "?"


def check(path):
    paras, sects, targets, parts = read_docx(path)
    rep = Report()
    P = "PL I NĐ 30/2020"

    # 1. Khổ giấy, lề (PL I mục I.1-3)
    sect = sects[-1] if sects else None
    pg = sect.find(W + "pgSz") if sect is not None else None
    if pg is None:
        rep.err("kho-giay", "Không thấy khai báo khổ giấy (mặc định của Word có thể là Letter).", f"{P}, mục I.1")
    else:
        w, h = int(pg.get(W + "w")), int(pg.get(W + "h"))
        if abs(w - A4_W) <= 30 and abs(h - A4_H) <= 30:
            rep.ok("kho-giay", "Khổ giấy A4, trình bày theo chiều dài.", f"{P}, mục I.1-2")
        elif abs(h - A4_W) <= 30 and abs(w - A4_H) <= 30:
            rep.warn("kho-giay", "Khổ A4 nhưng xoay ngang: chỉ dùng khi văn bản có bảng biểu.", f"{P}, mục I.2")
        else:
            rep.err("kho-giay", f"Khổ giấy {mm(w)} x {mm(h)} mm, không phải A4 (210 x 297 mm).", f"{P}, mục I.1")
    mar = sect.find(W + "pgMar") if sect is not None else None
    if mar is not None:
        limits = {"top": ("trên", 20, 25), "bottom": ("dưới", 20, 25), "left": ("trái", 30, 35), "right": ("phải", 15, 20)}
        bad = []
        for k, (name, lo, hi) in limits.items():
            v = mm(abs(int(mar.get(W + k) or 0)))
            if not lo - 0.2 <= v <= hi + 0.2:
                bad.append(f"lề {name} {v} mm (cần {lo}-{hi})")
        if bad:
            rep.err("le-trang", "Định lề sai: " + "; ".join(bad) + ".", f"{P}, mục I.3")
        else:
            rep.ok("le-trang", "Định lề trang đúng quy định.", f"{P}, mục I.3")

    # 2. Phông chữ, màu chữ (PL I mục I.4)
    fonts, colors = {}, set()
    for p in paras:
        for r in p.visible_runs():
            fonts.setdefault(r["font"] or "(mặc định)", []).append(r["text"])
            if r["color"] not in (None, "auto", "000000"):
                colors.add(r["color"])
    other = {f: t for f, t in fonts.items() if f != "Times New Roman"}
    if other:
        sample = "; ".join(f"{f}: “{norm(t[0])[:30]}”" for f, t in list(other.items())[:3])
        rep.err("phong-chu", f"Có chữ không dùng phông Times New Roman ({sample}).", f"{P}, mục I.4")
    else:
        rep.ok("phong-chu", "Toàn bộ văn bản dùng phông Times New Roman.", f"{P}, mục I.4")
    if colors:
        rep.warn("mau-chu", f"Có chữ màu khác màu đen ({', '.join(sorted(colors))}).", f"{P}, mục I.4")

    # 3. Số trang (PL I mục I.7)
    hdr_ids = {}
    if sect is not None:
        for ref in sect.findall(W + "headerReference"):
            hdr_ids[ref.get(W + "type")] = targets.get(ref.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"))
    default_hdr = parts.get("word/" + (hdr_ids.get("default") or ""), "")
    first_hdr = parts.get("word/" + (hdr_ids.get("first") or ""), "")
    title_pg = sect is not None and _on(sect, "titlePg")
    if "PAGE" not in default_hdr:
        rep.warn("so-trang", "Chưa đánh số trang ở phần lề trên (nếu văn bản dài hơn một trang).", f"{P}, mục I.7")
    elif not title_pg or "PAGE" in first_hdr:
        rep.err("so-trang", "Trang thứ nhất đang hiển thị số trang; phải ẩn số trang đầu.", f"{P}, mục I.7")
    elif 'w:val="center"' not in default_hdr:
        rep.warn("so-trang", "Số trang chưa canh giữa theo chiều ngang.", f"{P}, mục I.7")
    else:
        rep.ok("so-trang", "Số trang canh giữa lề trên, không hiện ở trang đầu.", f"{P}, mục I.7")

    texts = [norm(p.text) for p in paras]

    def find(pred):
        return next((k for k, t in enumerate(texts) if t and pred(t)), None)

    # 4. Quốc hiệu, Tiêu ngữ (PL I mục II.1)
    qh = find(lambda t: t.upper().replace("HOÀ", "HÒA") == QUOC_HIEU)
    if qh is None:
        rep.err("quoc-hieu", f"Không thấy Quốc hiệu “{QUOC_HIEU}”.", f"{P}, mục II.1.a")
    else:
        p = paras[qh]
        prob = []
        if texts[qh] != QUOC_HIEU and texts[qh].upper() == QUOC_HIEU.replace("HÒA", "HOÀ"):
            rep.warn("quoc-hieu", "Quốc hiệu viết “HOÀ”; mẫu NĐ 30 viết “HÒA”.", f"{P}, mục II.1.a")
        if texts[qh] != texts[qh].upper():
            prob.append("chưa in hoa")
        if not p.all("b"):
            prob.append("chưa in đậm")
        if p.all("i"):
            prob.append("đang in nghiêng")
        if not in_range(p.sizes(), 12, 13):
            prob.append(f"cỡ chữ {fmt_sizes(p.sizes())} (cần 12-13)")
        if prob:
            rep.err("quoc-hieu", "Quốc hiệu " + ", ".join(prob) + ".", f"{P}, mục II.1.a")
        else:
            rep.ok("quoc-hieu", "Quốc hiệu đúng: in hoa, đứng, đậm, cỡ " + fmt_sizes(p.sizes()) + ".", f"{P}, mục II.1.a")
    tn = find(lambda t: "Độc lập" in t and "Hạnh phúc" in t)
    if tn is None:
        rep.err("tieu-ngu", f"Không thấy Tiêu ngữ “{TIEU_NGU}”.", f"{P}, mục II.1.b")
    else:
        p, prob = paras[tn], []
        if texts[tn] != TIEU_NGU:
            prob.append(f"viết là “{texts[tn]}” (đúng: “{TIEU_NGU}”, gạch nối có cách chữ)")
        if not p.all("b"):
            prob.append("chưa in đậm")
        if not in_range(p.sizes(), 13, 14):
            prob.append(f"cỡ chữ {fmt_sizes(p.sizes())} (cần 13-14)")
        if prob:
            rep.err("tieu-ngu", "Tiêu ngữ " + "; ".join(prob) + ".", f"{P}, mục II.1.b")
        else:
            rep.ok("tieu-ngu", "Tiêu ngữ đúng chữ, đứng, đậm, cỡ " + fmt_sizes(p.sizes()) + ".", f"{P}, mục II.1.b")
        near = paras[tn:tn + 3]
        if not (p.all("u") or any(x.border or x.has_drawing for x in near)):
            rep.warn("tieu-ngu", "Chưa thấy đường kẻ ngang dưới Tiêu ngữ.", f"{P}, mục II.1.b")
        if qh is not None and paras[qh].sizes() and p.sizes() and paras[qh].sizes()[0] + 1 != p.sizes()[0]:
            rep.warn("co-chu", "Cỡ chữ Quốc hiệu và Tiêu ngữ chưa đi cặp (13 với 14, hoặc 12 với 13).",
                     f"{P}, mục V, chú thích 1")

    # 5. Số, ký hiệu (PL I mục II.3)
    so = find(lambda t: re.match(r"^Số\s*:?", t) and "/" in t)
    loai_ten = None
    if so is None:
        rep.err("so-ky-hieu", "Không thấy dòng “Số: .../...”.", f"{P}, mục II.3")
    else:
        t, p, prob = texts[so], paras[so], []
        m = re.match(r"^Số\s*(:?)\s*([0-9]*)\s*/\s*(\S+)", t)
        if not m or not m.group(1):
            prob.append("thiếu dấu hai chấm sau từ “Số”")
        if m and m.group(2) and len(m.group(2)) == 1:
            prob.append(f"số {m.group(2)} nhỏ hơn 10 phải ghi thêm số 0 ({int(m.group(2)):02d})")
        if m and re.search(r"\s-|-\s", t.split("/", 1)[1]):
            prob.append("có khoảng trắng quanh dấu gạch nối của ký hiệu")
        if m:
            kh = m.group(3)
            allowed = {a for _, a in LOAI.values() if a}
            groups = kh.split("-")
            if groups and groups[0] in allowed:
                loai_ten = next(n for n, a in LOAI.values() if a == groups[0])
            for g in groups:
                if g and g not in allowed and g != g.upper():
                    prob.append(f"nhóm ký hiệu “{g}” chưa in hoa")
        if not in_range(p.sizes(), 13, 13):
            prob.append(f"cỡ chữ {fmt_sizes(p.sizes())} (cần 13)")
        if any(r["b"] or r["i"] for r in p.visible_runs()):
            prob.append("phải là chữ đứng, không đậm")
        if prob:
            rep.err("so-ky-hieu", "Số, ký hiệu: " + "; ".join(prob) + ".", f"{P}, mục II.3.c", t)
        else:
            rep.ok("so-ky-hieu", f"Số, ký hiệu đúng: “{t}”.", f"{P}, mục II.3.c")
        # Tên cơ quan ban hành: các dòng in hoa ngay trên dòng Số trong cùng ô
        k = so - 1
        org = []
        while k >= 0 and paras[k].where[0] == paras[so].where[0]:
            if texts[k] and texts[k] not in (QUOC_HIEU,) and is_upper(texts[k]):
                org.insert(0, k)
            elif texts[k]:
                break
            k -= 1
        if org:
            issuer, prob = paras[org[-1]], []
            if not issuer.all("b"):
                prob.append("tên cơ quan ban hành chưa in đậm")
            for idx in org:
                if not in_range(paras[idx].sizes(), 12, 13):
                    prob.append(f"“{texts[idx][:30]}” cỡ {fmt_sizes(paras[idx].sizes())} (cần 12-13)")
            if len(org) > 1 and paras[org[0]].all("b"):
                prob.append("tên cơ quan chủ quản không in đậm")
            if prob:
                rep.err("co-quan", "Tên cơ quan: " + "; ".join(prob) + ".", f"{P}, mục II.2.b")
            else:
                rep.ok("co-quan", "Tên cơ quan in hoa, cơ quan ban hành in đậm, cỡ 12-13.", f"{P}, mục II.2.b")

    # 6. Địa danh, thời gian (PL I mục II.4)
    dd = find(lambda t: re.search(r",\s*ngày\b.*\btháng\b.*\bnăm\b", t))
    if dd is None:
        rep.err("dia-danh", "Không thấy dòng “Địa danh, ngày ... tháng ... năm ...”.", f"{P}, mục II.4")
    else:
        t, p, prob = texts[dd], paras[dd], []
        m = re.search(r"ngày\s+(\d+)\s+tháng\s+(\d+)\s+năm\s+(\d+)", t)
        if m:
            d, mo = m.group(1), m.group(2)
            if len(d) == 1:
                prob.append(f"ngày {d} phải viết {int(d):02d}")
            if int(mo) in (1, 2) and len(mo) == 1:
                prob.append(f"tháng {mo} phải viết {int(mo):02d}")
            if 3 <= int(mo) <= 9 and len(mo) == 2:
                rep.warn("dia-danh", f"Tháng {mo}: NĐ 30 chỉ thêm số 0 với tháng 1, 2 (mẫu viết “tháng {int(mo)}”).",
                         f"{P}, mục II.4.b")
        if not p.all("i"):
            prob.append("chưa in nghiêng")
        if any(r["b"] for r in p.visible_runs()):
            prob.append("không in đậm")
        if not in_range(p.sizes(), 13, 14):
            prob.append(f"cỡ chữ {fmt_sizes(p.sizes())} (cần 13-14)")
        place = t.split(",")[0]
        if place and place[0].islower():
            prob.append("chữ cái đầu địa danh chưa viết hoa")
        if prob:
            rep.err("dia-danh", "Địa danh, thời gian: " + "; ".join(prob) + ".", f"{P}, mục II.4", t)
        else:
            rep.ok("dia-danh", "Địa danh, thời gian đúng: in thường, nghiêng, cỡ " + fmt_sizes(p.sizes()) + ".", f"{P}, mục II.4")

    # 7. Tên loại, trích yếu (PL I mục II.5)
    names = {n for n, _ in LOAI.values() if n}
    tl = find(lambda t: t.upper() in names and t == t.upper())
    vv = find(lambda t: t.startswith("V/v"))
    body_start = None
    if tl is not None:
        p, prob = paras[tl], []
        if not p.all("b"):
            prob.append("chưa in đậm")
        if not in_range(p.sizes(), 13, 14):
            prob.append(f"cỡ chữ {fmt_sizes(p.sizes())} (cần 13-14)")
        if p.jc != "center":
            prob.append("chưa canh giữa")
        if prob:
            rep.err("ten-loai", f"Tên loại “{texts[tl]}”: " + "; ".join(prob) + ".", f"{P}, mục II.5.b")
        else:
            rep.ok("ten-loai", f"Tên loại “{texts[tl]}” in hoa, đứng, đậm, canh giữa.", f"{P}, mục II.5.b")
        ty = tl + 1
        while ty < len(paras) and not texts[ty]:
            ty += 1
        if ty < len(paras) and paras[ty].jc == "center" and not texts[ty].startswith(("Kính gửi", "Căn cứ")):
            p, prob = paras[ty], []
            if not p.all("b"):
                prob.append("chưa in đậm")
            if not in_range(p.sizes(), 13, 14):
                prob.append(f"cỡ chữ {fmt_sizes(p.sizes())} (cần 13-14)")
            if prob:
                rep.err("trich-yeu", "Trích yếu: " + "; ".join(prob) + ".", f"{P}, mục II.5.b", texts[ty])
            else:
                rep.ok("trich-yeu", "Trích yếu in thường, đứng, đậm.", f"{P}, mục II.5.b")
            if not any(x.border or x.has_drawing for x in paras[ty:ty + 3]):
                rep.warn("trich-yeu", "Chưa thấy đường kẻ ngang dưới trích yếu.", f"{P}, mục II.5.b")
            body_start = ty + 1
        else:
            body_start = tl + 1
    elif vv is not None:
        p, prob = paras[vv], []
        if not in_range(p.sizes(), 12, 13):
            prob.append(f"cỡ chữ {fmt_sizes(p.sizes())} (cần 12-13)")
        if any(r["b"] or r["i"] for r in p.visible_runs()):
            prob.append("phải là chữ đứng, không đậm")
        if prob:
            rep.err("trich-yeu", "Trích yếu công văn: " + "; ".join(prob) + ".", f"{P}, mục II.5.b", texts[vv])
        else:
            rep.ok("trich-yeu", "Trích yếu công văn “V/v ...” cỡ 12-13, chữ đứng.", f"{P}, mục II.5.b")
        body_start = vv + 1
    else:
        rep.warn("ten-loai", "Không nhận ra tên loại văn bản hoặc trích yếu “V/v” của công văn.", f"{P}, mục II.5")
    if loai_ten and tl is not None and texts[tl] != loai_ten:
        rep.warn("ten-loai", f"Ký hiệu cho thấy loại “{loai_ten}” nhưng tên loại là “{texts[tl]}”.", "PL III NĐ 30/2020")

    # 8. Kính gửi (PL I mục II.9.d)
    kg = find(lambda t: t.startswith("Kính gửi"))
    if kg is not None:
        t, p, prob = texts[kg], paras[kg], []
        if not t.startswith("Kính gửi:"):
            prob.append("thiếu dấu hai chấm sau “Kính gửi”")
        if not in_range(p.sizes(), 13, 14):
            prob.append(f"cỡ chữ {fmt_sizes(p.sizes())} (cần 13-14)")
        if p.all("b") or p.all("i"):
            prob.append("phải là chữ đứng, không đậm")
        if prob:
            rep.err("kinh-gui", "Kính gửi: " + "; ".join(prob) + ".", f"{P}, mục II.9.d", t)
        else:
            rep.ok("kinh-gui", "Phần “Kính gửi” đúng cỡ chữ và dấu câu.", f"{P}, mục II.9.d")

    # 9. Nơi nhận (PL I mục II.9.d, ô 9b)
    nn = find(lambda t: t.startswith("Nơi nhận"))
    if nn is None:
        rep.warn("noi-nhan", "Không thấy phần “Nơi nhận:”.", f"{P}, mục II.9")
    else:
        p, prob = paras[nn], []
        if not texts[nn].startswith("Nơi nhận:"):
            prob.append("thiếu dấu hai chấm")
        if not (p.all("b") and p.all("i")):
            prob.append("từ “Nơi nhận” phải nghiêng, đậm")
        if not in_range(p.sizes(), 12, 12):
            prob.append(f"cỡ chữ {fmt_sizes(p.sizes())} (cần 12)")
        items = []
        k = nn + 1
        while k < len(paras) and paras[k].where == paras[nn].where and texts[k].startswith(("-", "–")):
            items.append(k)
            k += 1
        if k < len(paras) and paras[k].where == paras[nn].where and texts[k].startswith("-") is False and items:
            pass
        for idx in items:
            if not in_range(paras[idx].sizes(), 11, 11):
                prob.append(f"dòng “{texts[idx][:25]}” cỡ {fmt_sizes(paras[idx].sizes())} (cần 11)")
                break
        if items:
            for idx in items[:-1]:
                if not texts[idx].endswith(";"):
                    prob.append(f"dòng “{texts[idx][:25]}” phải kết thúc bằng “;”")
                    break
            last = texts[items[-1]]
            if not re.match(r"^-\s*Lưu\s*:\s*VT\b", last):
                prob.append("dòng cuối phải là “- Lưu: VT, ...”")
            if not last.endswith("."):
                prob.append("dòng “Lưu” phải kết thúc bằng dấu chấm")
        else:
            prob.append("không có dòng nơi nhận nào bắt đầu bằng “-”")
        if prob:
            rep.err("noi-nhan", "Nơi nhận: " + "; ".join(prob) + ".", f"{P}, mục II.9.d")
        else:
            rep.ok("noi-nhan", "Nơi nhận đúng: nhãn cỡ 12 nghiêng đậm, danh sách cỡ 11, kết thúc “Lưu: VT”.", f"{P}, mục II.9.d")

    # 10. Chữ ký: quyền hạn, chức vụ, họ tên (PL I mục II.7)
    sig = [k for k, t in enumerate(texts) if re.match(r"^(TM|KT|Q|TL|TUQ)\s*\.", t)]
    if sig:
        k = sig[0]
        t, prob = texts[k], []
        if not re.match(r"^(TM|KT|Q|TL|TUQ)\. \S", t):
            prob.append(f"viết tắt quyền hạn “{t.split()[0]}” phải có dấu chấm và một khoảng trắng (KT. GIÁM ĐỐC)")
        if t != t.upper() or not paras[k].all("b"):
            prob.append("quyền hạn, chức vụ phải in hoa, đậm")
        if not in_range(paras[k].sizes(), 13, 14):
            prob.append(f"cỡ chữ {fmt_sizes(paras[k].sizes())} (cần 13-14)")
        if prob:
            rep.err("chu-ky", "Chữ ký: " + "; ".join(prob) + ".", f"{P}, mục II.7.b, đ", t)
        else:
            rep.ok("chu-ky", f"Quyền hạn ký “{t}” đúng.", f"{P}, mục II.7")
    hoc_vi = find(lambda t: re.match(r"^(GS|PGS|TS|ThS|KS|BS|Th\.S|T\.S)\.?\s*(TS\.?\s*)?[A-ZĐ][a-zà-ỹ]+ ", t))
    if hoc_vi is not None and paras[hoc_vi].where[0] == "table":
        rep.warn("chu-ky", f"Trước họ tên người ký không ghi học hàm, học vị: “{texts[hoc_vi]}”.", f"{P}, mục II.7.c")

    # 11. Nội dung (PL I mục II.6): từ sau trích yếu đến khối chữ ký, cộng nội dung các phụ lục
    end = next((k for k in range(body_start or 0, len(paras))
                if texts[k].startswith("Nơi nhận") or (sig and k == sig[0])
                or re.match(r"^(THƯ KÝ|CHỦ TỌA|CHỦ TRÌ)$", texts[k])), len(paras))
    appendix = next((k + 1 for k in range(end, len(paras)) if paras[k].section_end), len(paras))
    skip, k = set(), body_start or 0
    while k < end:  # khối "Kính gửi" hoặc "... điện:" ngay sau phần đầu
        if texts[k].startswith("Kính gửi") or texts[k].endswith("điện:"):
            skip.add(k)
            k += 1
            while k < end and texts[k].startswith("-") and paras[k].jc != "both":
                skip.add(k)
                k += 1
            break
        if texts[k] and paras[k].where[0] == "body":
            break
        k += 1
    body = [k for k in list(range(body_start or 0, end)) + list(range(appendix, len(paras)))
            if k not in skip and paras[k].where[0] == "body" and texts[k]
            and paras[k].jc not in ("center", "right")]
    if body_start is None:
        body = []
    if body:
        n = len(body)
        bad_size = [k for k in body if not in_range(paras[k].sizes(), 13, 14)]
        bad_jc = [k for k in body if paras[k].jc not in ("both", "distribute")]
        bad_ind = [k for k in body if not 540 <= paras[k].first_line <= 740]
        bad_line = [k for k in body if (paras[k].line_rule == "auto" and not 240 <= paras[k].line <= 360)
                    or (paras[k].line_rule == "exact" and paras[k].line < 20 * max(paras[k].sizes() or [13]))]
        for code, lst, msg in (
            ("noi-dung-co-chu", bad_size, "cỡ chữ ngoài 13-14"),
            ("noi-dung-canh-deu", bad_jc, "chưa canh đều hai lề"),
            ("noi-dung-lui-dau-dong", bad_ind, "lùi đầu dòng không phải 1 cm hoặc 1,27 cm"),
            ("noi-dung-gian-dong", bad_line, "giãn dòng ngoài khoảng dòng đơn đến 1,5 lines"),
        ):
            if lst:
                (rep.err if len(lst) > n / 2 else rep.warn)(
                    code, f"{len(lst)}/{n} đoạn nội dung {msg}.", f"{P}, mục II.6.e", texts[lst[0]][:60])
        gaps = [max(paras[a].after, paras[b].before) for a, b in zip(body, body[1:]) if b == a + 1]
        if gaps and sum(1 for g in gaps if g < 120) > len(gaps) / 2:
            rep.warn("noi-dung-cach-doan", "Khoảng cách giữa các đoạn nhỏ hơn 6pt.", f"{P}, mục II.6.e")
        if not any((bad_size, bad_jc, bad_ind, bad_line)):
            rep.ok("noi-dung", f"{n} đoạn nội dung đúng cỡ chữ, canh đều, lùi đầu dòng, giãn dòng.", f"{P}, mục II.6.e")
        # Căn cứ ban hành: các dòng trước "QUYẾT ĐỊNH:" / "QUYẾT NGHỊ:" (Mẫu 1.1-1.3)
        opener = next((k for k in range(body_start or 0, end) if re.match(r"^QUYẾT (ĐỊNH|NGHỊ)\s*:$", texts[k])), None)
        if opener is not None:
            cc = [k for k in range(body_start or 0, opener)
                  if texts[k].startswith(("Căn cứ", "Theo đề nghị", "Xét đề nghị"))]
            if cc:
                if not all(paras[k].all("i") for k in cc):
                    rep.err("can-cu", "Căn cứ ban hành phải in nghiêng.", f"{P}, mục II.6.a")
                if not all(texts[k].endswith(";") for k in cc[:-1]) or not texts[cc[-1]].endswith("."):
                    rep.warn("can-cu", "Mỗi căn cứ kết thúc bằng “;”, căn cứ cuối kết thúc bằng “.”.", f"{P}, mục II.6.a")
                if all(paras[k].all("i") for k in cc):
                    rep.ok("can-cu", "Căn cứ ban hành in nghiêng.", f"{P}, mục II.6.a")
        last_body = [k for k in body if k < end]
        if last_body and not texts[last_body[-1]].endswith("/."):
            rep.warn("ket-thuc", "Đoạn nội dung cuối thường kết thúc bằng “./.” như các mẫu.", "PL III NĐ 30/2020",
                     texts[last_body[-1]][-60:])
    return rep


LEVEL_ICON = {"dat": "✔", "loi": "✖", "canh_bao": "⚠"}


def print_report(path, rep, as_json=False):
    if as_json:
        print(json.dumps({"file": path, "ket_qua": rep.items,
                          "so_loi": sum(1 for x in rep.items if x["muc_do"] == "loi"),
                          "so_canh_bao": sum(1 for x in rep.items if x["muc_do"] == "canh_bao")},
                         ensure_ascii=False, indent=2))
        return
    errs = sum(1 for x in rep.items if x["muc_do"] == "loi")
    warns = sum(1 for x in rep.items if x["muc_do"] == "canh_bao")
    print(f"Kiểm tra thể thức: {path}")
    for level in ("loi", "canh_bao", "dat"):
        for x in rep.items:
            if x["muc_do"] != level:
                continue
            extra = f"\n      tại: “{x['vi_tri']}”" if x.get("vi_tri") else ""
            print(f"  {LEVEL_ICON[level]} {x['noi_dung']}  [{x['can_cu']}]{extra}")
    print(f"Kết quả: {errs} lỗi, {warns} cảnh báo, {sum(1 for x in rep.items if x['muc_do'] == 'dat')} mục đạt.")


# ---------------------------------------------------------------- sửa lỗi cấu trúc

def fix(src, dst):
    """Sửa những lỗi sửa tự động được an toàn: khổ A4, lề, phông Times New Roman, số trang."""
    changes = []
    with zipfile.ZipFile(src) as z:
        files = {n: z.read(n) for n in z.namelist()}
    doc = files["word/document.xml"].decode("utf-8")

    def fix_sect(m):
        s = m.group(0)
        pg = re.search(r"<w:pgSz\b[^>]*/>", s)
        landscape = bool(pg and 'w:orient="landscape"' in pg.group(0))
        new_pg = (f'<w:pgSz w:w="{A4_H}" w:h="{A4_W}" w:orient="landscape"/>' if landscape
                  else f'<w:pgSz w:w="{A4_W}" w:h="{A4_H}"/>')
        if pg:
            if pg.group(0) != new_pg:
                s = s.replace(pg.group(0), new_pg)
                changes.append("đặt khổ giấy A4")
        else:
            s = s.replace("</w:sectPr>", new_pg + "</w:sectPr>")
            changes.append("đặt khổ giấy A4")
        mar = re.search(r"<w:pgMar\b[^>]*/>", s)
        want = {"top": (20, 25, 20), "bottom": (20, 25, 20), "left": (30, 35, 30), "right": (15, 20, 15)}
        attrs = dict(re.findall(r'w:(\w+)="(-?\d+)"', mar.group(0))) if mar else {}
        fixed = dict(attrs)
        for k, (lo, hi, default) in want.items():
            v = mm(abs(int(attrs.get(k, 0))))
            if not lo - 0.2 <= v <= hi + 0.2:
                fixed[k] = str(tw_mm(min(max(v, lo), hi)) if v else tw_mm(default))
        fixed.setdefault("header", "567")
        fixed.setdefault("footer", "567")
        fixed.setdefault("gutter", "0")
        new_mar = "<w:pgMar " + " ".join(f'w:{k}="{v}"' for k, v in fixed.items()) + "/>"
        if fixed != attrs:
            changes.append("chỉnh lề trang về khoảng cho phép")
            s = s.replace(mar.group(0), new_mar) if mar else s.replace("</w:sectPr>", new_mar + "</w:sectPr>")
        return s

    doc = re.sub(r"<w:sectPr\b.*?</w:sectPr>", fix_sect, doc, flags=re.S)

    def tnr(xml):
        n = 0

        def rep_font(m):
            nonlocal n
            tag = m.group(0)
            if re.search(r'"(Symbol|Wingdings[^"]*|Webdings)"', tag):
                return tag
            if tag != FONT:
                n += 1
            return FONT
        out = re.sub(r"<w:rFonts\b[^>]*/>", rep_font, xml)
        return out, n

    total = 0
    for name in list(files):
        if re.match(r"word/(document|styles|header[^/]*|footer[^/]*|footnotes|endnotes)\.xml$", name):
            xml = doc if name == "word/document.xml" else files[name].decode("utf-8")
            xml, n = tnr(xml)
            if name == "word/styles.xml" and "<w:rPrDefault>" in xml:
                xml = re.sub(r"<w:rPrDefault>\s*<w:rPr>(?!.*?<w:rFonts)", "<w:rPrDefault><w:rPr>" + FONT, xml, count=1, flags=re.S)
            if name == "word/styles.xml":
                xml = re.sub(r"<w:rPrDefault/>", f"<w:rPrDefault><w:rPr>{FONT}</w:rPr></w:rPrDefault>", xml)
            total += n
            if name == "word/document.xml":
                doc = xml
            else:
                files[name] = xml.encode("utf-8")
    if total:
        changes.append(f"đổi {total} chỗ khai báo phông sang Times New Roman")
    if "word/theme/theme1.xml" in files:
        th = files["word/theme/theme1.xml"].decode("utf-8")
        th2 = re.sub(r'(<a:(?:major|minor)Font>\s*<a:latin typeface=")[^"]*(")', r"\1Times New Roman\2", th)
        if th2 != th:
            files["word/theme/theme1.xml"] = th2.encode("utf-8")
            changes.append("đặt phông theme sang Times New Roman")

    # số trang: thêm khi chưa có trường PAGE trong header
    has_page = any(b"PAGE" in data for n, data in files.items() if re.match(r"word/header[^/]*\.xml$", n))
    last_sect = list(re.finditer(r"<w:sectPr\b.*?</w:sectPr>", doc, flags=re.S))
    if not has_page and last_sect and "headerReference" not in last_sect[-1].group(0):
        size = 14
        files["word/header_vn_docx.xml"] = page_number_header(size).encode("utf-8")
        files["word/header_vn_docx_first.xml"] = hdr("").encode("utf-8")
        rels = files["word/_rels/document.xml.rels"].decode("utf-8")
        rels = rels.replace("</Relationships>",
                            '<Relationship Id="rIdVnHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header_vn_docx.xml"/>'
                            '<Relationship Id="rIdVnHdrFirst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header_vn_docx_first.xml"/></Relationships>')
        files["word/_rels/document.xml.rels"] = rels.encode("utf-8")
        ct = files["[Content_Types].xml"].decode("utf-8")
        for part in ("header_vn_docx.xml", "header_vn_docx_first.xml"):
            ct = ct.replace("</Types>", f'<Override PartName="/word/{part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>')
        files["[Content_Types].xml"] = ct.encode("utf-8")
        s = last_sect[-1].group(0)
        open_tag = re.match(r"<w:sectPr\b[^>]*>", s).group(0)
        s2 = s.replace(open_tag, open_tag + '<w:headerReference w:type="default" r:id="rIdVnHdr"/>'
                       '<w:headerReference w:type="first" r:id="rIdVnHdrFirst"/>', 1)
        if "<w:titlePg" not in s2:
            s2 = s2.replace("</w:sectPr>", "<w:titlePg/></w:sectPr>")
        doc = doc[:last_sect[-1].start()] + s2 + doc[last_sect[-1].end():]
        if "xmlns:r=" not in doc[:2000]:
            doc = doc.replace("<w:document ", '<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
        changes.append("thêm số trang canh giữa lề trên, ẩn ở trang đầu")
    files["word/document.xml"] = doc.encode("utf-8")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)
    return changes


# ---------------------------------------------------------------- PDF (tùy chọn)

def find_soffice():
    for c in ("soffice", "libreoffice"):
        p = shutil.which(c)
        if p:
            return p
    for p in ("/Applications/LibreOffice.app/Contents/MacOS/soffice",
              r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.exists(p):
            return p
    return None


def to_pdf(docx):
    exe = find_soffice()
    if not exe:
        print("Không tìm thấy LibreOffice (soffice), bỏ qua bước xuất PDF.", file=sys.stderr)
        return None
    out_dir = os.path.dirname(os.path.abspath(docx))
    with tempfile.TemporaryDirectory() as profile:
        try:
            subprocess.run([exe, "-env:UserInstallation=" + pathlib.Path(profile).as_uri(), "--headless", "--convert-to", "pdf",
                            "--outdir", out_dir, docx], check=True, stdout=subprocess.DEVNULL,
                           stderr=subprocess.DEVNULL, timeout=180)
        except (subprocess.SubprocessError, OSError) as e:
            print(f"LibreOffice không xuất được PDF ({e}); file .docx vẫn dùng được.", file=sys.stderr)
            return None
    pdf = os.path.splitext(os.path.abspath(docx))[0] + ".pdf"
    return pdf if os.path.exists(pdf) else None


# ---------------------------------------------------------------- dòng lệnh

def cmd_types():
    print(f"{'mã loai':<17}{'tên loại':<18}viết tắt")
    for code, (name, abbr) in LOAI.items():
        print(f"{code:<17}{name or '(công văn)':<18}{abbr or '-'}")


def cmd_example(kind):
    path = os.path.join(SKILL_DIR, "examples", f"{kind}.md")
    if not os.path.exists(path):
        have = sorted(f[:-3] for f in os.listdir(os.path.join(SKILL_DIR, "examples")) if f.endswith(".md"))
        raise SystemExit(f"chưa có mẫu {kind!r}; các mẫu có sẵn: {', '.join(have)}")
    with open(path, encoding="utf-8") as f:
        sys.stdout.write(f.read())


def main(argv=None):
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(prog="vanban", description="Soạn và kiểm tra văn bản hành chính theo NĐ 30/2020/NĐ-CP")
    ap.add_argument("--version", action="version", version=f"vn-docx {VERSION}")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("tao", aliases=["build"], help="soạn .docx từ noi-dung.md")
    p.add_argument("src")
    p.add_argument("dst")
    p.add_argument("--pdf", action="store_true", help="xuất thêm PDF (cần LibreOffice)")
    p = sub.add_parser("kiem-tra", aliases=["check"], help="kiểm tra thể thức một file .docx")
    p.add_argument("files", nargs="+")
    p.add_argument("--json", action="store_true")
    p = sub.add_parser("sua", aliases=["fix"], help="sửa khổ giấy, lề, phông, số trang")
    p.add_argument("src")
    p.add_argument("dst")
    p = sub.add_parser("bai-tap", aliases=["exercise"], help="cài sẵn lỗi thể thức vào văn bản đúng (bài tập)")
    p.add_argument("src")
    p.add_argument("dst")
    p.add_argument("--loi", nargs="+", help="chỉ cài các lỗi này (mặc định: tất cả)")
    p = sub.add_parser("mau", aliases=["example"], help="in nội dung mẫu của một loại văn bản")
    p.add_argument("kind")
    sub.add_parser("loai", aliases=["types"], help="bảng tên loại và chữ viết tắt (PL III)")
    a = ap.parse_args(argv)
    if a.cmd in ("tao", "build"):
        out = build(a.src, a.dst)
        print(f"Đã tạo {out}")
        if a.pdf:
            pdf = to_pdf(out)
            if pdf:
                print(f"Đã tạo {pdf}")
        return 0
    if a.cmd in ("kiem-tra", "check"):
        code = 0
        for f in a.files:
            rep = check(f)
            print_report(f, rep, a.json)
            if any(x["muc_do"] == "loi" for x in rep.items):
                code = 1
        return code
    if a.cmd in ("sua", "fix"):
        changes = fix(a.src, a.dst)
        print(f"Đã ghi {a.dst}")
        for c in changes or ["không có gì cần sửa tự động"]:
            print(f"  - {c}")
        return 0
    if a.cmd in ("bai-tap", "exercise"):
        sys.path.insert(0, HERE)
        import loi_mau
        bad = [k for k in (a.loi or []) if k not in loi_mau.LOI]
        if bad:
            raise SystemExit(f"không có lỗi {', '.join(bad)}; chọn trong: {', '.join(loi_mau.LOI)}")
        applied = loi_mau.make(a.src, a.dst, a.loi)
        print(f"Đã ghi {a.dst} với các lỗi:")
        for k in applied:
            print(f"  - {loi_mau.LOI[k][0]}")
        skipped = [k for k in a.loi or loi_mau.LOI if k not in applied]
        if skipped:
            print(f"Không cài được (văn bản không có chỗ phù hợp): {', '.join(skipped)}")
        return 0
    if a.cmd in ("mau", "example"):
        cmd_example(a.kind)
        return 0
    cmd_types()
    return 0


if __name__ == "__main__":
    sys.exit(main())
