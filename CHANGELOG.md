# Changelog

## 0.1.0 - 2026-10-17

Bản đầu tiên.

- Skill `van-ban-hanh-chinh` cho Claude Code, Claude Cowork, Codex, Antigravity, OpenCode và các agent đọc Agent Skill.
- `vanban.py tao`: dựng .docx đúng Phụ lục I NĐ 30/2020 từ một file Markdown đơn giản: bảng phần đầu hai cột,
  đường kẻ dưới Tiêu ngữ dài bằng dòng chữ, số và ngày tự thêm số 0, Kính gửi, căn cứ, chữ ký, nơi nhận,
  độ khẩn, địa chỉ chân trang, số trang ẩn ở trang đầu, phụ lục đánh số trang riêng, kết thúc `./.`.
- `vanban.py kiem-tra`: kiểm tra khổ giấy, lề, phông, số trang, Quốc hiệu, Tiêu ngữ, tên cơ quan, số ký hiệu,
  địa danh và ngày, tên loại, trích yếu, Kính gửi, căn cứ, nội dung, chữ ký, nơi nhận; báo cáo kèm căn cứ.
- `vanban.py sua`: sửa khổ A4, lề, phông Times New Roman, thêm số trang cho file Word có sẵn.
- 12 mẫu, có mẫu cấp xã theo mô hình chính quyền địa phương 2 cấp (Công văn 4168/BNV-CQĐP).
- npm `vn-docx`, plugin Claude Code, gói zip cho Claude Cowork.
