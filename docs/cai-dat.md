# Cài đặt vn-docx

Skill tên **`van-ban-hanh-chinh`**, nằm trong repo GitHub **`dinhkarate/vn-docx`**. Chọn đúng nơi bạn dùng AI:

| Bạn dùng | Cách nhanh nhất |
|---|---|
| Claude Code, Codex, Antigravity, OpenCode, Cursor, Gemini CLI | [Một lệnh `npx skills`](#1-một-lệnh-cho-mọi-agent) |
| Claude desktop, chế độ **Cowork** | [Thêm plugin từ GitHub](#2-claude-cowork) hoặc tải file zip |
| claude.ai trên trình duyệt | [Tải file zip lên](#3-claudeai-trên-trình-duyệt) |
| Giảng viên cài cho cả lớp | [Cài cho tổ chức](#4-giảng-viên-cài-cho-cả-lớp) |
| Không dùng AI, chỉ muốn kiểm tra file Word | [Dòng lệnh `npx vn-docx`](#5-không-dùng-ai-dòng-lệnh) |

Yêu cầu chung: **Python 3.9 trở lên** trên máy chạy skill. macOS, Linux có sẵn; Windows cài từ
[python.org](https://www.python.org/downloads/) và tích **Add python.exe to PATH**. Lệnh `npx` cần
[Node.js](https://nodejs.org/) 16+. Claude Cowork, claude.ai chạy skill trong máy ảo của Claude nên không cần cài gì.

## 1. Một lệnh cho mọi agent

```bash
npx skills add dinhkarate/vn-docx -g
```

- `-g` cài cho người dùng (mọi dự án). Bỏ `-g` để chỉ cài vào thư mục dự án đang mở.
- Chọn agent bằng `-a`, ví dụ `-a claude-code -a codex`; `-a '*'` cho tất cả; thêm `-y` để không hỏi.
- Skill được chép vào `~/.claude/skills/` (Claude Code) và `~/.agents/skills/` (Codex và các agent theo chuẩn Agent Skills).
- Cập nhật: `npx skills update -g`. Gỡ: `npx skills remove van-ban-hanh-chinh -g`.

**Nhờ chính agent cài.** Mở Claude Code hoặc Codex (bản dòng lệnh hay desktop) rồi gõ:

> Cài giúp tôi skill soạn văn bản hành chính: chạy `npx skills add dinhkarate/vn-docx -g -y`, xong thì kiểm tra
> bằng `python3 ~/.agents/skills/van-ban-hanh-chinh/scripts/vanban.py loai`.

Mở phiên mới để agent nạp skill. Thử: *"Soạn công văn của UBND xã An Bình mời họp giao ban tuần, người ký Chủ tịch
Lê Văn C"*.

**Claude Code, cách plugin** (thay cho `npx skills`, chọn một trong hai):

```text
/plugin marketplace add dinhkarate/vn-docx
/plugin install van-ban-hanh-chinh@vn-docx
```

Kiểm tra bằng `claude plugin list`; cập nhật bằng `claude plugin update van-ban-hanh-chinh@vn-docx`.

## 2. Claude Cowork

Bật **Settings > Capabilities > Code execution and file creation** trước.

**Cách A, plugin từ GitHub** (gói Pro, Max, Team, Enterprise):

1. **Customize > Plugins > Add > Add marketplace > Add from a repository**.
2. Nhập `dinhkarate/vn-docx`, xác nhận.
3. Chọn plugin **van-ban-hanh-chinh**, bấm cài.

**Cách B, tải file zip** (mọi gói có Skills):

1. Tải [`van-ban-hanh-chinh.zip`](https://github.com/dinhkarate/vn-docx/releases/latest/download/van-ban-hanh-chinh.zip)
   (không giải nén).
2. **Customize > Skills > + > Create skill > Upload a skill**, chọn file zip.
3. Bật công tắc của skill **van-ban-hanh-chinh**.

Trong Cowork, chọn thư mục làm việc rồi yêu cầu: *"Soạn tờ trình xin kinh phí số hóa tài liệu, xuất file Word"*,
hoặc *"Kiểm tra file `cong-van.docx` có đúng Nghị định 30 không"*.

> Tên menu có thể đổi theo phiên bản ứng dụng. Hướng dẫn gốc:
> [Use plugins in Claude](https://support.claude.com/en/articles/13837440-use-plugins-in-claude),
> [Use skills in Claude](https://support.claude.com/en/articles/12512180-using-skills-in-claude).

## 3. claude.ai trên trình duyệt

Giống cách B của Cowork: bật **Code execution and file creation**, rồi **Customize > Skills > Upload a skill** với
file zip. Skill chỉ thuộc tài khoản của bạn.

## 4. Giảng viên cài cho cả lớp

- **Claude Team / Enterprise**: chủ tổ chức vào **Organization settings > Plugins & skills**, tải
  `van-ban-hanh-chinh.zip` lên. Skill hiện trong danh sách Skills của mọi thành viên (có biểu tượng nhóm), học viên
  tự bật. Enterprise: bật **Code execution and file creation** và **Skills** ở tab Policy.
- **Học viên dùng tài khoản cá nhân** (Free, Pro): gửi link file zip và mục 2 hoặc 3 ở trên; hoặc học viên dùng
  Claude Code / Codex thì gửi một dòng `npx skills add dinhkarate/vn-docx -g`.
- Phòng máy: cài một lần cho mọi agent bằng `npx skills add dinhkarate/vn-docx -g -a '*' -y`.

## 5. Không dùng AI: dòng lệnh

```bash
npx vn-docx kiem-tra van-ban.docx          # kiểm tra thể thức
npx vn-docx sua van-ban.docx da-sua.docx   # sửa khổ giấy, lề, phông, số trang
npx vn-docx mau cong-van > noi-dung.md     # lấy khung mẫu
npx vn-docx tao noi-dung.md van-ban.docx   # dựng văn bản
```

Không có Node: tải repo rồi chạy `python3 skills/van-ban-hanh-chinh/scripts/vanban.py <lệnh>`
(Windows: `py -3 ...`).

## Sự cố thường gặp

| Hiện tượng | Cách xử lý |
|---|---|
| `python3: command not found` (Windows) | Cài Python từ python.org, tích Add to PATH; hoặc dùng `py -3` |
| Agent không dùng skill | Mở phiên mới; gọi thẳng `/van-ban-hanh-chinh` (Claude Code) hoặc `$van-ban-hanh-chinh` (Codex); kiểm tra thư mục skill có `SKILL.md` |
| Cowork báo không chạy được code | Bật Code execution and file creation; mở cuộc trò chuyện mới |
| Không có file PDF | Cần LibreOffice; hoặc mở .docx bằng Word, File > Save As > PDF |
| Kiểm tra báo lỗi phông trên file Word | Word dùng phông theme (Calibri, Aptos); chạy `sua` hoặc chọn toàn bộ văn bản, đặt Times New Roman |
