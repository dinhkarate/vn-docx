# Giáo trình: Trình bày văn bản hành chính theo Nghị định 30 với AI

Khóa ngắn 6 bài, mỗi bài 45-60 phút, cho công chức, viên chức, nhân viên văn phòng, sinh viên ngành Lưu trữ học,
Quản trị văn phòng. Học viên vừa nắm quy định trình bày, vừa biết giao việc soạn thảo cho AI và **tự kiểm chứng**
kết quả.

| Bài | Nội dung | Thực hành |
|---|---|---|
| 0 | Chuẩn bị | Cài skill, chạy thử |
| 1 | Thể thức văn bản là gì | Đọc và kiểm tra văn bản mẫu |
| 2 | Trang giấy: khổ, lề, phông, số trang | Bài tập tìm lỗi, sửa trong Word, sửa bằng `sua` |
| 3 | Phần đầu văn bản | Bài tập tìm lỗi Quốc hiệu, Tiêu ngữ, số, ngày tháng |
| 4 | Nội dung và phần kết | Căn cứ, bố cục, chữ ký, nơi nhận |
| 5 | Soạn văn bản bằng AI | Viết yêu cầu, rà kết quả, chuẩn hóa văn bản thô |
| 6 | Bài cuối khóa | Soạn bộ văn bản cho một tình huống, chấm tự động |

Tài liệu nguồn cho giảng viên: [the-thuc-nd30.md](../../skills/van-ban-hanh-chinh/references/the-thuc-nd30.md)
(bảng quy định kèm trích dẫn), [viet-hoa.md](../../skills/van-ban-hanh-chinh/references/viet-hoa.md), và toàn văn
Nghị định 30/2020/NĐ-CP (Phụ lục I, II, III).

## Bài 0. Chuẩn bị

- Học viên có Claude (Cowork hoặc claude.ai, gói có Skills) hoặc Claude Code / Codex. Không có AI vẫn học được
  bài 1-4 bằng `npx vn-docx`.
- Cài theo [docs/cai-dat.md](../cai-dat.md). Lớp dùng Claude Team: giảng viên tải skill cho cả tổ chức (mục 4).
- Chạy thử: yêu cầu agent *"Liệt kê các loại văn bản hành chính và chữ viết tắt"*, hoặc `npx vn-docx loai`.

## Bài 1. Thể thức văn bản là gì

**Mục tiêu**: kể được các thành phần thể thức (Điều 8 NĐ 30) và chỉ ra vị trí trên trang (Mục IV Phụ lục I).

1. Giảng viên dựng văn bản mẫu: `npx vn-docx mau cong-van > cong-van.md`, `npx vn-docx tao cong-van.md cong-van.docx`.
2. Học viên mở file, đánh số từng thành phần: Quốc hiệu và Tiêu ngữ; tên cơ quan; số, ký hiệu; địa danh, thời gian;
   tên loại, trích yếu; nội dung; chức vụ, họ tên người ký; dấu; nơi nhận.
3. Chạy `npx vn-docx kiem-tra cong-van.docx`, đọc từng dòng ✔ và đối chiếu căn cứ trong ngoặc vuông.

**Câu hỏi**: Công văn khác văn bản có tên loại ở điểm nào? (trích yếu "V/v" dưới số, ký hiệu; có "Kính gửi").
Văn bản của UBND xã từ 01/7/2025 có ghi tên tỉnh không? (không, xem `examples/cong-van-cap-xa.md`).

## Bài 2. Trang giấy: khổ, lề, phông, số trang

**Mục tiêu**: đặt đúng khổ A4, lề trên, dưới 20-25 mm, trái 30-35 mm, phải 15-20 mm; Times New Roman; số trang
giữa lề trên, không hiện ở trang 1 (Mục I Phụ lục I).

```bash
npx vn-docx tao cong-van.md dung.docx
npx vn-docx bai-tap dung.docx bai-2.docx --loi kho-giay le-trang phong-chu so-trang
```

1. Học viên mở `bai-2.docx`, tự tìm 4 lỗi, sửa trong Word:
   Layout > Size > A4; Layout > Margins > Custom Margins; Home > Font; Insert > Page Number > Top of Page >
   Plain Number 2 và tích *Different First Page*.
2. Nộp bài, giảng viên chấm: `npx vn-docx kiem-tra bai-2-da-sua.docx`.
3. So sánh với `npx vn-docx sua bai-2.docx bai-2-may-sua.docx`: máy sửa được lỗi nào, lỗi nào phải tự sửa.

**Lỗi hay gặp ngoài thực tế**: file mẫu cũ khổ Letter; lề 2,54 cm mặc định; phông Calibri, Aptos của theme Word.

## Bài 3. Phần đầu văn bản

**Mục tiêu**: trình bày đúng Quốc hiệu, Tiêu ngữ (cỡ chữ, đậm, đường kẻ), tên cơ quan, "Số:", địa danh và ngày
tháng (Mục II.1-II.4 Phụ lục I).

```bash
npx vn-docx bai-tap dung.docx bai-3.docx --loi quoc-hieu tieu-ngu so-ky-hieu dia-danh
```

Điểm cần nhấn mạnh:

- Tiêu ngữ dùng **gạch nối có cách chữ** `Độc lập - Tự do - Hạnh phúc`, không dùng gạch dài `–`.
- Đường kẻ dưới Tiêu ngữ dài bằng dòng chữ; dưới tên cơ quan ban hành dài 1/3 đến 1/2 dòng chữ.
- `Số: 05/QĐ-UBND`: có dấu hai chấm, số nhỏ hơn 10 thêm số 0.
- `Hà Nội, ngày 05 tháng 02 năm 2026`: in nghiêng; ngày nhỏ hơn 10 và tháng 1, 2 thêm số 0 (tháng 3 trở đi không).
- Địa danh là tên gọi chính thức của đơn vị hành chính nơi cơ quan đóng trụ sở; kiểm tra sau sắp xếp đơn vị hành
  chính 2025.

## Bài 4. Nội dung và phần kết

**Mục tiêu**: bố cục phần, chương, mục, điều; căn cứ ban hành; quyền hạn ký (TM., KT., Q., TL., TUQ.); nơi nhận.

```bash
npx vn-docx mau quyet-dinh > quyet-dinh.md
npx vn-docx tao quyet-dinh.md qd.docx
npx vn-docx bai-tap qd.docx bai-4.docx --loi noi-dung noi-nhan
```

Thảo luận trên văn bản mẫu:

- Căn cứ in nghiêng, mỗi căn cứ kết thúc bằng `;`, căn cứ cuối bằng `.`; "QUYẾT ĐỊNH:" đứng riêng.
- Nội dung canh đều hai lề, lùi đầu dòng 1 cm hoặc 1,27 cm, cách đoạn tối thiểu 6 pt, giãn dòng 1-1,5.
- Kết thúc nội dung bằng `./.`.
- Phó giám đốc ký thay: `KT. GIÁM ĐỐC` trên, `PHÓ GIÁM ĐỐC` dưới. Ký thay mặt tập thể: `TM. ỦY BAN NHÂN DÂN`.
- "Nơi nhận:" nghiêng đậm cỡ 12; từng dòng cỡ 11, bắt đầu bằng `-`, dòng cuối `- Lưu: VT, ...`.

Bài tập viết hoa (Phụ lục II): lấy các ví dụ sai trong bảng lỗi thường gặp của `viet-hoa.md`, học viên sửa, rồi hỏi agent
*"Kiểm tra viết hoa đoạn này theo Phụ lục II Nghị định 30"* để đối chiếu.

## Bài 5. Soạn văn bản bằng AI

**Mục tiêu**: giao việc cho AI có đủ thông tin; biết agent làm gì, không làm gì; rà lại kết quả.

Một yêu cầu tốt có: **loại văn bản, cơ quan ban hành (và cơ quan cấp trên), gửi ai, việc gì, người ký và quyền
hạn ký, nơi nhận**. Ví dụ:

> Soạn quyết định của Giám đốc Sở Nội vụ tỉnh Đồng Nai thành lập Tổ kiểm tra công tác văn thư năm 2026, gồm 5
> người (để trống họ tên), thời gian hoạt động đến 31/12/2026. Phó Giám đốc Nguyễn Văn A ký thay.

> Đây là nội dung thông báo tôi viết vội trên Zalo [dán]. Trình bày lại thành thông báo của Văn phòng UBND phường
> Tân Định, giữ nguyên ý, sửa chính tả và viết hoa, báo tôi biết chỗ nào đã sửa.

> Kiểm tra file `to-trinh.docx` đính kèm có đúng Nghị định 30 không. Giải thích từng lỗi và cách sửa trong Word,
> đừng sửa hộ.

Học viên rà kết quả theo danh sách:

- Agent có tự điền số văn bản, ngày ký, họ tên không có trong yêu cầu không? (không được phép)
- Căn cứ pháp lý agent đưa vào còn hiệu lực không? Tên cơ quan, địa danh có đúng sau sắp xếp không?
- `kiem-tra` có ra `0 lỗi` không?
- Văn phong: câu ngắn, một nghĩa, không khẩu ngữ.

**Bảo mật**: không đưa văn bản mật, thông tin cá nhân nhạy cảm cho AI bên ngoài. Kiểm tra thể thức file loại này
bằng `npx vn-docx kiem-tra` (chạy trên máy, không gửi đi đâu).

## Bài 6. Bài cuối khóa

Tình huống mẫu: *UBND phường X tổ chức Hội nghị tổng kết công tác văn thư, lưu trữ năm 2026.* Học viên soạn (có
thể dùng AI) bộ 4 văn bản:

1. Tờ trình của Văn phòng xin chủ trương tổ chức hội nghị.
2. Kế hoạch tổ chức hội nghị.
3. Giấy mời đại biểu.
4. Quyết định khen thưởng tập thể, cá nhân (danh sách ở phụ lục).

Thang điểm gợi ý (10 điểm):

| Tiêu chí | Điểm | Cách chấm |
|---|---|---|
| Thể thức | 4 | `kiem-tra`: mỗi file 0 lỗi được 1 điểm; còn lỗi trừ 0,25 mỗi lỗi |
| Thẩm quyền, quyền hạn ký, nơi nhận hợp lý | 2 | Giảng viên |
| Nội dung, văn phong, viết hoa | 3 | Giảng viên |
| Không có thông tin bịa (số, tên, số liệu) | 1 | Giảng viên |

## Dành cho giảng viên

**Ra đề không trùng nhau.** Mỗi học viên một văn bản gốc khác (12 mẫu trong `examples/`) và một tổ hợp lỗi khác:

```bash
npx vn-docx mau to-trinh > to-trinh.md && npx vn-docx tao to-trinh.md goc.docx
npx vn-docx bai-tap goc.docx de-an.docx --loi kho-giay tieu-ngu noi-nhan
npx vn-docx bai-tap goc.docx de-binh.docx --loi le-trang so-ky-hieu noi-dung
```

`bai-tap` in ra các lỗi đã cài: lưu lại làm đáp án. Danh sách mã lỗi: `kho-giay le-trang phong-chu quoc-hieu
tieu-ngu so-ky-hieu dia-danh noi-nhan so-trang noi-dung`.

**Chấm cả thư mục bài nộp.** Nhờ agent: *"Chấm tất cả file .docx trong thư mục `bai-nop` bằng skill văn bản hành
chính, lập bảng: tên file, số lỗi, số cảnh báo, các lỗi chính"*. Hoặc tự chạy (macOS, Linux):

```bash
for f in bai-nop/*.docx; do
  printf '%s\t' "$f"; npx -y vn-docx kiem-tra "$f" | tail -1
done
```

Windows PowerShell:

```powershell
Get-ChildItem bai-nop\*.docx | ForEach-Object { "$($_.Name)`t" + (npx -y vn-docx kiem-tra $_.FullName | Select-Object -Last 1) }
```

`npx vn-docx kiem-tra --json file.docx` cho kết quả máy đọc được (`so_loi`, `so_canh_bao`, `ket_qua`), tiện nhập
vào bảng tính.

**Giới hạn cần nói rõ với học viên**: chương trình kiểm thể thức, không kiểm nội dung, thẩm quyền, viết hoa trong
câu, căn cứ còn hiệu lực. Những phần đó là việc của người soạn và người duyệt.
