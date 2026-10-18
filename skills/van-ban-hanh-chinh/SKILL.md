---
name: van-ban-hanh-chinh
description: Soạn, kiểm tra và sửa văn bản hành chính Việt Nam (.docx Word) đúng thể thức, kỹ thuật trình bày theo Nghị định 30/2020/NĐ-CP - công văn, quyết định, tờ trình, báo cáo, kế hoạch, thông báo, giấy mời, giấy giới thiệu, biên bản, công điện và đủ 29 loại văn bản hành chính, cập nhật mô hình chính quyền 2 cấp. Dùng khi người dùng muốn soạn/viết/làm một văn bản hành chính, công văn, quyết định, tờ trình..., trình bày lại hoặc chuẩn hóa văn bản "theo Nghị định 30", "đúng thể thức", "chuẩn NĐ30", kiểm tra/rà soát/chấm lỗi thể thức một file Word, hoặc học, dạy cách trình bày văn bản hành chính. Also for Vietnamese government/administrative document drafting and format checking (Decree 30/2020).
version: 0.2.0
tags: [vietnamese, nghi-dinh-30, van-ban-hanh-chinh, docx, word, government-documents]
---

# Văn bản hành chính theo Nghị định 30/2020/NĐ-CP

Bộ công cụ `scripts/vanban.py` chỉ dùng thư viện chuẩn Python 3.9+, chạy offline, không gửi dữ liệu đi đâu:

```bash
V="<thư mục chứa file SKILL.md này>/scripts/vanban.py"
python3 "$V" mau cong-van > noi-dung.md        # lấy khung mẫu của một loại
python3 "$V" tao noi-dung.md Ten-Van-Ban.docx   # dựng .docx đúng thể thức (thêm --pdf nếu có LibreOffice)
python3 "$V" kiem-tra van-ban.docx              # kiểm tra thể thức, mã thoát 1 nếu có lỗi (--json cho máy đọc)
python3 "$V" sua vao.docx ra.docx               # sửa khổ giấy, lề, phông, số trang của file có sẵn
python3 "$V" bai-tap dung.docx sai.docx         # cài sẵn lỗi thường gặp: bài tập cho học viên
python3 "$V" loai                               # 29 loại văn bản và chữ viết tắt
```

Windows: dùng `py -3` hoặc `python` thay cho `python3`. Có Node: `npx vn-docx <lệnh>` chạy cùng chương trình.

| File | Đọc khi |
|---|---|
| `references/cu-phap.md` | Trước khi viết `noi-dung.md` lần đầu: mọi trường thông tin và cách viết nội dung |
| `references/the-thuc-nd30.md` | Cần trích quy định, giải thích lỗi, văn bản cấp xã (chính quyền 2 cấp), phân biệt với VBQPPL, văn bản Đảng |
| `references/viet-hoa.md` | Rà viết hoa (Phụ lục II) |
| `examples/<loai>.md` | Mẫu đầy đủ: cong-van, cong-van-cap-xa, quyet-dinh, quyet-dinh-ban-hanh, to-trinh, bao-cao, ke-hoach, thong-bao, giay-moi, giay-gioi-thieu, bien-ban, cong-dien |

## Không dùng skill này cho

- Văn bản quy phạm pháp luật (nghị quyết, quyết định QPPL có số dạng `12/2026/QĐ-UBND`): theo NĐ 78/2025/NĐ-CP.
- Văn bản của Đảng: theo Hướng dẫn 05-HD/VPTW ngày 27/5/2026.
- Đơn từ cá nhân, hợp đồng thương mại mẫu riêng, luận văn: không có thể thức NĐ 30.

Gặp các trường hợp trên thì nói rõ cho người dùng; chỉ dùng skill nếu họ vẫn muốn trình bày gần giống NĐ 30.

## 1. Soạn văn bản mới

1. **Chốt thông tin.** Cần: loại văn bản, cơ quan ban hành (và cơ quan chủ quản), người nhận, nội dung chính,
   người ký và quyền hạn ký, nơi nhận. Yêu cầu đã đủ ý thì làm luôn. Thiếu thông tin quan trọng thì hỏi
   **một lần**, gom các câu hỏi lại. Không bịa số văn bản, ngày ký, họ tên người ký, số liệu: để trống
   (chương trình tự để chỗ trống `Số:      /...`, `ngày   tháng   năm`) và báo lại cho người dùng.
2. **Lấy khung**: `python3 "$V" mau <loai>`; loại không có mẫu thì lấy mẫu gần nhất (`to-trinh` cho văn bản
   có tên loại, `cong-van` cho công văn) và đổi `loai`.
3. **Viết `noi-dung.md`** theo `references/cu-phap.md`: mỗi dòng một đoạn; cấu trúc `I.`, `1.`, `a)`,
   `Điều 1.`, `Chương I` được nhận ra tự động. Văn phong hành chính: câu ngắn, rõ, một nghĩa, không khẩu ngữ;
   viết hoa theo `references/viet-hoa.md`; viện dẫn lần đầu ghi đủ tên loại, số ký hiệu, ngày, cơ quan,
   trích yếu. Cơ quan cấp xã: một dòng `ỦY BAN NHÂN DÂN XÃ ...`, không có dòng tên tỉnh.
4. **Dựng**: `python3 "$V" tao noi-dung.md <Ten-File>.docx` (tên file không dấu, nối bằng `-`, đặt cạnh file
   nguồn hoặc trong thư mục làm việc của người dùng).
5. **Kiểm tra**: `python3 "$V" kiem-tra <Ten-File>.docx` phải ra `0 lỗi`. Có lỗi: sửa `noi-dung.md`, dựng lại.
6. **Trả kết quả** ngắn gọn: đường dẫn file, các chỗ còn để trống cần điền, điều chưa kiểm được
   (căn cứ còn hiệu lực, thẩm quyền ký). Không chép lại toàn bộ văn bản vào câu trả lời.

## 2. Kiểm tra file Word có sẵn

1. `python3 "$V" kiem-tra <file>.docx` (file `.doc` cũ: nhờ người dùng lưu thành `.docx`, hoặc
   `soffice --headless --convert-to docx` nếu máy có LibreOffice).
2. Báo lại theo nhóm **Lỗi** (✖, sai quy định) và **Cảnh báo** (⚠, nên xem lại), mỗi mục một dòng bằng lời
   dễ hiểu, kèm căn cứ trong ngoặc vuông của báo cáo và cách sửa trong Word (ví dụ: Layout > Margins >
   Custom Margins, Left 3 cm).
3. Rà thêm phần chương trình không kiểm được (cuối `references/the-thuc-nd30.md`): viết hoa, văn phong,
   viện dẫn, quyền hạn ký, tên cơ quan và địa danh sau sắp xếp đơn vị hành chính.
4. Đề nghị cách sửa: `sua` cho lỗi khổ giấy, lề, phông, số trang (giữ nguyên nội dung và định dạng khác);
   lỗi phần đầu, chữ ký, nơi nhận thì trích nội dung sang `noi-dung.md` và dựng lại bằng `tao`.
   Không ghi đè file gốc: luôn ghi ra file mới.

## 3. Chuẩn hóa văn bản thô

Người dùng dán chữ hoặc gửi file Word trình bày lộn xộn: giữ nguyên ý và số liệu, tách thông tin vào khối
`---` (cơ quan, số, ngày, trích yếu, người ký, nơi nhận), phần thân vào nội dung, sửa chính tả và viết hoa
rõ ràng, rồi làm như mục 1 từ bước 4. Báo lại những chỗ đã sửa câu chữ.

## 4. Dạy và học

Khi người dùng là học viên đang học (hỏi "vì sao", "sai ở đâu", nộp bài để chấm):

- Chấm bài bằng `kiem-tra`, giải thích từng lỗi: quy định là gì, trích dẫn mục của Phụ lục I, cách sửa
  trong Word. Không sửa hộ ngay; gợi ý để học viên tự sửa rồi chấm lại.
- Bài tập "tìm lỗi": dựng văn bản đúng từ `examples/`, rồi `python3 "$V" bai-tap dung.docx sai.docx`
  (chọn lỗi bằng `--loi kho-giay le-trang phong-chu quoc-hieu tieu-ngu so-ky-hieu dia-danh noi-nhan so-trang noi-dung`).
  Giao `sai.docx` cho học viên; chấm bài nộp lại bằng `kiem-tra`.
- Tóm tắt nhanh một thành phần bằng bảng trong `references/the-thuc-nd30.md`, không đọc thuộc lòng cả nghị định.

## Lỗi hay gặp khi dùng

- Giá trị trong khối `---` bắt đầu bằng `#`, hay có `: ` ở giữa câu thì đặt trong ngoặc kép.
- Một đoạn bị xuống dòng giữa chừng trong `noi-dung.md` sẽ thành hai đoạn: viết mỗi đoạn trên một dòng.
- `Điều 2. Tổ công tác có nhiệm vụ:` có dấu câu cuối nên chỉ "Điều 2." đậm; tiêu đề điều không có dấu câu cuối.
- `kiem-tra` báo lỗi phông trên file do Word tạo: thường do phông theme (Calibri, Aptos); chạy `sua`.
- Không có LibreOffice thì bỏ qua `--pdf`; người dùng tự xuất PDF từ Word (File > Save As > PDF).
