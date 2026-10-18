"""Cài sẵn lỗi thể thức thường gặp vào một văn bản đúng: bài tập "tìm và sửa lỗi" cho học viên, và cho test.

  python3 vanban.py bai-tap dung.docx sai.docx [--loi kho-giay tieu-ngu ...]   # mặc định: tất cả
"""
import re
import zipfile

FONT_TNR = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>'
FONT_ARIAL = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/>'


def _run_with(doc, text):
    """Vị trí (start, end) của run chứa `text`."""
    i = doc.index(text)
    start = doc.rfind("<w:r>", 0, i)
    end = doc.index("</w:r>", i) + len("</w:r>")
    return start, end


def _edit_run(doc, text, fn):
    a, b = _run_with(doc, text)
    return doc[:a] + fn(doc[a:b]) + doc[b:]


# mã lỗi -> (mô tả, hàm sửa document.xml, hàm sửa styles.xml)
LOI = {
    "kho-giay": ("Khổ giấy Letter thay vì A4",
                 lambda d: re.sub(r'<w:pgSz w:w="\d+" w:h="\d+"/>', '<w:pgSz w:w="12240" w:h="15840"/>', d), None),
    "le-trang": ("Lề mặc định của Word (2,54 cm mỗi bên)",
                 lambda d: re.sub(r'w:top="\d+" w:right="\d+" w:bottom="\d+" w:left="\d+"',
                                  'w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"', d), None),
    "phong-chu": ("Phông Arial", None, lambda s: s.replace(FONT_TNR, FONT_ARIAL)),
    "quoc-hieu": ("Quốc hiệu không in đậm",
                  lambda d: _edit_run(d, "CỘNG HÒA XÃ HỘI", lambda r: r.replace("<w:b/><w:bCs/>", "")), None),
    "tieu-ngu": ("Tiêu ngữ dùng gạch ngang dài (–)",
                 lambda d: d.replace("Độc lập - Tự do - Hạnh phúc", "Độc lập – Tự do – Hạnh phúc"), None),
    "so-ky-hieu": ("“Số” thiếu dấu hai chấm, bỏ số 0 trước số nhỏ hơn 10",
                   lambda d: re.sub(r"Số: 0?(\d+)/", r"Số \1/", d, count=1), None),
    "dia-danh": ("Ngày tháng không in nghiêng, bỏ số 0 trước ngày, tháng nhỏ hơn 10",
                 lambda d: _edit_run(re.sub(r"(ngày|tháng) 0(\d)", r"\1 \2", d), ", ngày ",
                                     lambda r: r.replace("<w:i/><w:iCs/>", "")), None),
    "noi-nhan": ("“Nơi nhận” không in nghiêng, dòng Lưu thiếu dấu chấm",
                 lambda d: re.sub(r"(- Lưu: [^<]*)\.</w:t>", r"\1</w:t>",
                                  _edit_run(d, "Nơi nhận:", lambda r: r.replace("<w:i/><w:iCs/>", ""))), None),
    "so-trang": ("Hiện số trang ở trang đầu", lambda d: d.replace("<w:titlePg/>", ""), None),
    "noi-dung": ("Nội dung căn trái, không lùi đầu dòng",
                 lambda d: re.sub(r'<w:ind w:left="0" w:right="0" w:firstLine="\d+"/><w:jc w:val="both"/>',
                                  '<w:ind w:left="0" w:right="0"/><w:jc w:val="left"/>', d), None),
}


def make(src, dst, kinds=None):
    """Ghi `dst` từ `src` với các lỗi `kinds` (mặc định: tất cả); trả về các lỗi thực sự cài được."""
    with zipfile.ZipFile(src) as z:
        files = {n: z.read(n) for n in z.namelist()}
    doc = files["word/document.xml"].decode("utf-8")
    styles = files["word/styles.xml"].decode("utf-8")
    applied = []
    for k in kinds or LOI:
        _, fd, fs = LOI[k]
        try:
            new_doc = fd(doc) if fd else doc
        except ValueError:  # văn bản không có thành phần này
            continue
        new_styles = fs(styles) if fs else styles
        if (new_doc, new_styles) != (doc, styles):
            doc, styles = new_doc, new_styles
            applied.append(k)
    files["word/document.xml"] = doc.encode("utf-8")
    files["word/styles.xml"] = styles.encode("utf-8")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in files.items():
            z.writestr(n, data)
    return applied
