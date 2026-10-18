"""Kiểm thử vanban.py: python3 -m unittest discover -s tests"""
import glob
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from contextlib import redirect_stdout

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SKILL = os.path.join(ROOT, "skills", "van-ban-hanh-chinh")
sys.path.insert(0, os.path.join(SKILL, "scripts"))

import vanban  # noqa: E402
import loi_mau  # noqa: E402

EXAMPLES = sorted(glob.glob(os.path.join(SKILL, "examples", "*.md")))


def levels(rep, code=None):
    return [x["muc_do"] for x in rep.items if code is None or x["ma"] == code]


class TmpDir(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmp)

    def build(self, example="cong-van"):
        out = os.path.join(self.tmp, example + ".docx")
        vanban.build(os.path.join(SKILL, "examples", example + ".md"), out)
        return out

    def write_md(self, text):
        path = os.path.join(self.tmp, "nd.md")
        with open(path, "w", encoding="utf-8") as f:
            f.write(text)
        return path


class Formatting(unittest.TestCase):
    def test_ngay_thang_them_so_0_cho_ngay_nho_hon_10_va_thang_1_2(self):
        f = lambda d: vanban.ngay_thang({"dia_danh": "Hà Nội", "ngay": d})
        self.assertEqual(f("2026-01-05"), "Hà Nội, ngày 05 tháng 01 năm 2026")
        self.assertEqual(f("2026-02-09"), "Hà Nội, ngày 09 tháng 02 năm 2026")
        self.assertEqual(f("2026-03-05"), "Hà Nội, ngày 05 tháng 3 năm 2026")
        self.assertEqual(f("29/6/2019"), "Hà Nội, ngày 29 tháng 6 năm 2019")
        self.assertEqual(f("2026-12-15"), "Hà Nội, ngày 15 tháng 12 năm 2026")
        with self.assertRaises(ValueError):
            f("2026-02-30")

    def test_so_ky_hieu(self):
        self.assertEqual(vanban.so_ky_hieu({"so": "5", "ky_hieu": "QĐ-UBND"}, "quyet-dinh"), "Số: 05/QĐ-UBND")
        self.assertEqual(vanban.so_ky_hieu({"so": "15", "viet_tat_co_quan": "BNV"}, "quyet-dinh"), "Số: 15/QĐ-BNV")
        self.assertEqual(vanban.so_ky_hieu({"so": "12", "viet_tat_co_quan": "UBND", "viet_tat_don_vi": "VX"},
                                           "cong-van"), "Số: 12/UBND-VX")
        self.assertEqual(vanban.so_ky_hieu({"viet_tat_co_quan": "SNV"}, "quy-dinh"), "Số:       /QyĐ-SNV")

    def test_ket_thuc_bang_dau_cham_gach_cheo(self):
        self.assertEqual(vanban.end_mark("thi hành Quyết định này."), "thi hành Quyết định này./.")
        self.assertEqual(vanban.end_mark("đúng giờ"), "đúng giờ./.")
        self.assertEqual(vanban.end_mark("xong./."), "xong./.")

    def test_dau_cau_danh_sach(self):
        self.assertEqual(vanban.punct_list(["Như trên", "Lưu: VT, VP"]), ["Như trên;", "Lưu: VT, VP."])
        self.assertEqual(vanban.punct_list(["A;", "B."]), ["A;", "B."])

    def test_ngat_ten_co_quan_truoc_cap_hanh_chinh(self):
        lines = vanban.split_name("ỦY BAN NHÂN DÂN XÃ AN BÌNH", 13, True, 3400)
        self.assertEqual(lines, ["ỦY BAN NHÂN DÂN", "XÃ AN BÌNH"])
        self.assertEqual(vanban.split_name("SỞ NỘI VỤ", 13, True, 3400), ["SỞ NỘI VỤ"])

    def test_do_rong_chu(self):
        # Quốc hiệu đậm cỡ 13 rộng khoảng 9,5 cm: phải vừa cột phải của bảng đầu văn bản
        w_cm = vanban.text_width(vanban.QUOC_HIEU, 13, "B") / 567
        self.assertTrue(9.0 < w_cm < 10.0, w_cm)

    def test_bang_viet_tat_theo_phu_luc_III(self):
        abbr = {k: v[1] for k, v in vanban.LOAI.items()}
        self.assertEqual(abbr["quy-dinh"], "QyĐ")
        self.assertEqual(abbr["chuong-trinh"], "CTr")
        self.assertEqual(abbr["to-trinh"], "TTr")
        self.assertEqual(abbr["ban-ghi-nho"], "BGN")
        self.assertEqual(abbr["giay-nghi-phep"], "GNP")
        self.assertEqual(abbr["cong-van"], "")
        self.assertEqual(len(vanban.LOAI), 29)  # Điều 7 NĐ 30/2020


class Examples(TmpDir):
    def test_moi_mau_dung_the_thuc(self):
        self.assertGreaterEqual(len(EXAMPLES), 10)
        for path in EXAMPLES:
            name = os.path.basename(path)[:-3]
            with self.subTest(name):
                out = self.build(name)
                rep = vanban.check(out)
                bad = [x for x in rep.items if x["muc_do"] != "dat"]
                self.assertEqual(bad, [], name)

    def test_docx_hop_le(self):
        out = self.build("quyet-dinh")
        with zipfile.ZipFile(out) as z:
            names = set(z.namelist())
            self.assertIn("word/document.xml", names)
            import xml.etree.ElementTree as ET
            for n in names:
                if n.endswith(".xml") or n.endswith(".rels"):
                    ET.fromstring(z.read(n))  # mọi phần đều là XML hợp lệ
            doc = z.read("word/document.xml").decode("utf-8")
        self.assertEqual(doc.count("<w:sectPr>"), 2)  # văn bản chính + phụ lục (đánh số trang riêng)
        self.assertIn("Quyết định này./.", doc)

    def test_cong_van_cap_xa_mot_dong_ten_co_quan(self):
        out = self.build("cong-van-cap-xa")
        paras, *_ = vanban.read_docx(out)
        texts = [p.text for p in paras]
        self.assertIn("ỦY BAN NHÂN DÂN", texts)
        self.assertIn("XÃ AN BÌNH", texts)
        self.assertFalse(any(t.startswith("UBND TỈNH") for t in texts))


class Mistakes(TmpDir):
    def test_bat_duoc_tung_loi(self):
        for example in ("cong-van", "to-trinh", "quyet-dinh", "bien-ban", "cong-dien"):
            good = self.build(example)
            for kind in loi_mau.LOI:
                with self.subTest(example=example, kind=kind):
                    bad = os.path.join(self.tmp, kind + ".docx")
                    self.assertEqual(loi_mau.make(good, bad, [kind]), [kind])
                    rep = vanban.check(bad)
                    code = {"noi-dung": "noi-dung-canh-deu"}.get(kind, kind)
                    self.assertIn("loi", levels(rep, code), [x for x in rep.items if x["muc_do"] != "dat"])

    def test_sua_loi_cau_truc(self):
        good = self.build("cong-van")
        bad = os.path.join(self.tmp, "bad.docx")
        loi_mau.make(good, bad, ["kho-giay", "le-trang", "phong-chu"])
        fixed = os.path.join(self.tmp, "fixed.docx")
        changes = vanban.fix(bad, fixed)
        self.assertTrue(changes)
        rep = vanban.check(fixed)
        for code in ("kho-giay", "le-trang", "phong-chu"):
            self.assertEqual(levels(rep, code), ["dat"], code)

    def test_sua_them_so_trang(self):
        good = self.build("cong-van")
        with zipfile.ZipFile(good) as z:
            files = {n: z.read(n) for n in z.namelist()}
        doc = files["word/document.xml"].decode("utf-8")
        import re
        doc = re.sub(r"<w:headerReference[^>]*/>", "", doc).replace("<w:titlePg/>", "")
        files["word/document.xml"] = doc.encode("utf-8")
        for n in [n for n in files if n.startswith("word/header")]:
            del files[n]
        plain = os.path.join(self.tmp, "plain.docx")
        with zipfile.ZipFile(plain, "w") as z:
            for n, d in files.items():
                z.writestr(n, d)
        self.assertNotEqual(levels(vanban.check(plain), "so-trang"), ["dat"])
        fixed = os.path.join(self.tmp, "fixed.docx")
        vanban.fix(plain, fixed)
        self.assertEqual(levels(vanban.check(fixed), "so-trang"), ["dat"])


class Source(TmpDir):
    def test_thieu_front_matter(self):
        with self.assertRaises(SystemExit):
            vanban.parse_source("không có khối thông tin")

    def test_loai_khong_hop_le(self):
        path = self.write_md("---\nloai: don-xin-viec\n---\nNội dung.\n")
        with self.assertRaises(SystemExit):
            vanban.build(path, os.path.join(self.tmp, "x.docx"))

    def test_le_ngoai_khoang(self):
        path = self.write_md("---\nloai: thong-bao\nle: 10 10 20 10\n---\nNội dung.\n")
        with self.assertRaises(SystemExit):
            vanban.build(path, os.path.join(self.tmp, "x.docx"))

    def test_toi_thieu(self):
        path = self.write_md("---\nloai: thong-bao\nco_quan: CÔNG TY TNHH A\nviet_tat_co_quan: CTA\n"
                             "dia_danh: Hà Nội\nngay: 2026-10-17\ntrich_yeu: Lịch nghỉ lễ\n---\nNội dung.\n")
        out = vanban.build(path, os.path.join(self.tmp, "x.docx"))
        paras, *_ = vanban.read_docx(out)
        texts = [p.text for p in paras]
        self.assertIn("Số:       /TB-CTA", texts)
        self.assertIn("Nội dung./.", texts)

    def test_co_chu_13(self):
        path = self.write_md("---\nloai: thong-bao\nco_chu: 13\nco_quan: A\ndia_danh: Huế\nngay: 2026-10-17\n"
                             "trich_yeu: B\nnoi_nhan:\n  - Lưu: VT\n---\nNội dung.\n")
        out = vanban.build(path, os.path.join(self.tmp, "x.docx"))
        rep = vanban.check(out)
        self.assertEqual([x for x in rep.items if x["muc_do"] == "loi"], [])
        paras, *_ = vanban.read_docx(out)
        qh = next(p for p in paras if p.text == vanban.QUOC_HIEU)
        self.assertEqual(qh.sizes(), [12.0])


class Cli(TmpDir):
    def run_cli(self, *args):
        buf = io.StringIO()
        with redirect_stdout(buf):
            code = vanban.main(list(args))
        return code, buf.getvalue()

    def test_loai_va_mau(self):
        code, out = self.run_cli("loai")
        self.assertEqual(code, 0)
        self.assertIn("QyĐ", out)
        code, out = self.run_cli("mau", "cong-van")
        self.assertIn("loai: cong-van", out)

    def test_kiem_tra_json_va_ma_thoat(self):
        good = self.build("cong-van")
        code, out = self.run_cli("kiem-tra", "--json", good)
        self.assertEqual(code, 0)
        self.assertEqual(json.loads(out)["so_loi"], 0)
        bad = os.path.join(self.tmp, "bad.docx")
        loi_mau.make(good, bad)
        code, out = self.run_cli("kiem-tra", bad)
        self.assertEqual(code, 1)
        self.assertIn("✖", out)

    def test_bai_tap(self):
        good = self.build("cong-van")
        bad = os.path.join(self.tmp, "sai.docx")
        code, out = self.run_cli("bai-tap", good, bad, "--loi", "tieu-ngu", "dia-danh")
        self.assertEqual(code, 0)
        self.assertIn("gạch ngang dài", out)
        rep = vanban.check(bad)
        self.assertIn("loi", levels(rep, "tieu-ngu"))
        self.assertEqual(levels(rep, "kho-giay"), ["dat"])

    def test_cai_dat(self):
        home = os.path.join(self.tmp, "home")
        claude = os.path.join(home, ".claude", "skills", "van-ban-hanh-chinh")
        agents = os.path.join(home, ".agents", "skills", "van-ban-hanh-chinh")
        # bản cũ: Claude Code trỏ symlink sang ~/.agents (kiểu npx skills), ~/.agents có file thừa
        os.makedirs(agents)
        open(os.path.join(agents, "cu.txt"), "w").close()
        os.makedirs(os.path.dirname(claude))
        if hasattr(os, "symlink") and os.name != "nt":
            os.symlink(agents, claude)
        for _ in range(2):  # chạy lại là cập nhật
            done = vanban.install(home)
            self.assertEqual([p for _, p in done], [claude, agents])
            for d in (claude, agents):
                self.assertFalse(os.path.islink(d))
                self.assertTrue(os.path.isfile(os.path.join(d, "SKILL.md")))
                self.assertTrue(os.path.isfile(os.path.join(d, "scripts", "vanban.py")))
                self.assertFalse(os.path.exists(os.path.join(d, "cu.txt")))
                self.assertFalse(os.path.exists(os.path.join(d, ".claude-plugin")))
        # chạy từ bản đã cài: không tự xóa chính mình
        r = subprocess.run([sys.executable, os.path.join(agents, "scripts", "vanban.py"), "install"],
                           capture_output=True, text=True, encoding="utf-8",
                           env={**os.environ, "HOME": home, "USERPROFILE": home, "PYTHONIOENCODING": "utf-8"})
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertIn(claude, r.stdout)
        self.assertTrue(os.path.isfile(os.path.join(agents, "SKILL.md")))

    def test_khong_lenh_in_huong_dan(self):
        buf = io.StringIO()
        with redirect_stdout(buf):
            code = vanban.main([])
        self.assertEqual(code, 2)
        self.assertIn("cai-dat", buf.getvalue())

    def test_npm_wrapper(self):
        node = shutil.which("node")
        if not node:
            self.skipTest("không có node")
        r = subprocess.run([node, os.path.join(ROOT, "bin", "vn-docx.js"), "--version"], capture_output=True, text=True)
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertIn(vanban.VERSION, r.stdout)


def read(*parts):
    with open(os.path.join(*parts), encoding="utf-8") as f:
        return f.read()


class Release(TmpDir):
    def test_mot_phien_ban(self):
        pkg = json.loads(read(ROOT, "package.json"))
        plugin = json.loads(read(SKILL, ".claude-plugin", "plugin.json"))
        skill = read(SKILL, "SKILL.md")
        changelog = read(ROOT, "CHANGELOG.md")
        v = vanban.VERSION
        self.assertEqual(pkg["version"], v)
        self.assertEqual(plugin["version"], v)
        self.assertIn(f"version: {v}", skill)
        self.assertIn(f"## {v}", changelog)

    def test_goi_zip(self):
        sys.path.insert(0, os.path.join(ROOT, "tools"))
        import build_zip
        with zipfile.ZipFile(build_zip.build(self.tmp)) as z:
            names = z.namelist()
            skill = z.read("van-ban-hanh-chinh/SKILL.md").decode("utf-8")
        self.assertTrue(all(n.startswith("van-ban-hanh-chinh/") for n in names), names)
        self.assertIn("van-ban-hanh-chinh/scripts/vanban.py", names)
        self.assertIn("van-ban-hanh-chinh/scripts/loi_mau.py", names)
        self.assertFalse([n for n in names if ".claude-plugin" in n or n.endswith(".pyc")])
        desc = re.search(r"(?m)^description: (.*)$", skill).group(1)
        self.assertEqual(desc, build_zip.SHORT)
        self.assertLessEqual(len(desc), 200)
        self.assertNotIn(": ", desc)  # YAML không có ngoặc kép
        self.assertNotIn(" #", desc)


if __name__ == "__main__":
    unittest.main()
