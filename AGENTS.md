# Ghi chú cho agent làm việc trên repo này

- Sản phẩm là skill `skills/vn-docx/` (SKILL.md, `scripts/vanban.py`, `scripts/loi_mau.py`,
  `references/`, `examples/`). `bin/vn-docx.js` chỉ gọi lại `vanban.py` bằng Python của máy.
- `vanban.py` chỉ dùng thư viện chuẩn Python 3.9+. Không thêm thư viện ngoài.
- Test: `python3 -m unittest discover -s tests`. Mọi mẫu trong `examples/` phải ra 0 lỗi, 0 cảnh báo khi
  `kiem-tra`; mỗi lỗi trong `loi_mau.LOI` phải bị phát hiện.
- Sửa cách trình bày: dựng lại các mẫu và xem bằng mắt (LibreOffice: `vanban.py tao ... --pdf`), không chỉ dựa vào test.
- Quy định ghi trong `references/the-thuc-nd30.md` phải kèm trích dẫn Phụ lục I NĐ 30/2020 hoặc văn bản hướng dẫn.
- Phiên bản ở 4 chỗ phải khớp: `package.json`, `skills/vn-docx/.claude-plugin/plugin.json`,
  `SKILL.md`, `VERSION` trong `vanban.py`; thêm mục `CHANGELOG.md`. Phát hành bằng tag `vX.Y.Z` (xem CONTRIBUTING.md).
- Mô tả SKILL.md dài; gói zip cho Claude Cowork dùng mô tả ngắn `SHORT` trong `tools/build_zip.py` (tối đa 200 ký tự,
  không có `: `).
- Tài liệu cho người dùng viết bằng tiếng Việt; `README.en.md` là bản tóm tắt tiếng Anh.
