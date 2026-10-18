# Changelog

## 0.3.1 - 2026-10-18

- Windows: `cai-dat` không gỡ được skill tên cũ `van-ban-hanh-chinh` (xóa khi file còn mở). Đã sửa.

## 0.3.0 - 2026-10-18

- Đổi tên skill `van-ban-hanh-chinh` thành `vn-docx`: cài và gọi cùng một tên (`/vn-docx` trong Claude Code,
  `$vn-docx` trong Codex). Plugin Claude Code: `/plugin install vn-docx@vn-docx`. File zip cho Cowork: `vn-docx.zip`.
- `vn-docx cai-dat` tự gỡ skill tên cũ `van-ban-hanh-chinh` (nếu do vn-docx cài) để không bị trùng.

## 0.2.0 - 2026-10-18

- `vn-docx cai-dat` (bí danh `install`): cài skill cho Claude Code (`~/.claude/skills`) và Codex, các agent khác
  (`~/.agents/skills`) bằng một lệnh `npx vn-docx cai-dat`; chạy lại để cập nhật. Học viên chỉ cần nói "cài vn-docx".
- Gọi `vn-docx` không kèm lệnh: in hướng dẫn, nhắc lệnh cài skill.
- `bai-tap`: chỉ báo những lỗi thực sự cài được vào văn bản (trước đây báo cả lỗi không áp dụng).
- Gói zip cho Cowork dựng đúng đường dẫn khi đóng gói trên Windows.

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
