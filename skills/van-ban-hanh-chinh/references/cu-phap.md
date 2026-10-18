# Cú pháp `noi-dung.md`

Một file gồm hai phần: **khối thông tin** giữa hai dòng `---` và **nội dung** phía dưới.
Xem file mẫu đầy đủ bằng `python3 scripts/vanban.py mau <loai>` (danh sách trong `examples/`).

```markdown
---
loai: cong-van
co_quan_chu_quan: UBND TỈNH ĐỒNG NAI
co_quan: SỞ NỘI VỤ
so: 5
ky_hieu: SNV-VP
dia_danh: Đồng Nai
ngay: 2026-02-03
trich_yeu: V/v báo cáo kết quả công tác văn thư năm 2025
kinh_gui:
  - Ủy ban nhân dân các xã, phường
  - Các cơ quan chuyên môn thuộc Ủy ban nhân dân tỉnh
quyen_han: KT. GIÁM ĐỐC
chuc_vu: PHÓ GIÁM ĐỐC
nguoi_ky: Nguyễn Văn A
noi_nhan:
  - Như trên
  - Lưu: VT, VP.
---
Nội dung văn bản, mỗi dòng là một đoạn.
```

## Khối thông tin

- `khoa: gia tri` trên một dòng. Danh sách: để trống sau dấu hai chấm, mỗi mục một dòng `  - mục`.
- Dòng bắt đầu bằng `#` là ghi chú. Giá trị có dấu `:` hay `#` ở đầu thì đặt trong ngoặc kép.
- Không có trường nào bắt buộc ngoài `loai`. Trường bỏ trống được để chỗ trống cho người dùng điền,
  **không bịa** số văn bản, tên người ký, ngày ký.

| Trường | Ý nghĩa | Ví dụ, ghi chú |
|---|---|---|
| `loai` | Mã loại văn bản (29 loại, `vanban.py loai`) | `cong-van`, `quyet-dinh`, `to-trinh`, `bao-cao`, `ke-hoach`, `thong-bao`, `giay-moi`, `bien-ban`, `cong-dien`... |
| `co_quan_chu_quan` | Tên cơ quan chủ quản trực tiếp (ô 2), in hoa, không đậm. Có thể là danh sách nhiều dòng | `UBND TỈNH ĐỒNG NAI`, `BỘ NỘI VỤ`. UBND cấp xã: **bỏ trống** |
| `co_quan` | Tên cơ quan ban hành (ô 2), in hoa, đậm. Tên dài tự xuống dòng trước cấp hành chính; muốn tự ngắt thì viết danh sách | `SỞ NỘI VỤ`; `ỦY BAN NHÂN DÂN XÃ AN BÌNH` |
| `so` | Số văn bản. Số nhỏ hơn 10 tự thêm số 0 | `5` → `05`. Để trống khi chưa cấp số |
| `ky_hieu` | Ký hiệu đầy đủ sau dấu `/` | Văn bản có tên loại `QĐ-UBND`; công văn `UBND-VP` |
| `viet_tat_co_quan`, `viet_tat_don_vi` | Thay cho `ky_hieu`: tự ghép theo PL III | `quyet-dinh` + `UBND` → `QĐ-UBND`; `cong-van` + `UBND` + `VX` → `UBND-VX` |
| `dia_danh` | Địa danh (ô 4) | `Hà Nội`, `Thành phố Hồ Chí Minh`, tên xã/phường nơi đóng trụ sở |
| `ngay` | Ngày ban hành, `2026-10-17` hoặc `17/10/2026`. Tự thêm số 0 cho ngày < 10, tháng 1 và 2 | Để trống: `ngày   tháng   năm 2026` |
| `trich_yeu` | Trích yếu. Công văn: tự thêm `V/v`. Có thể là danh sách nhiều dòng | `Về việc thành lập Tổ công tác` |
| `ten_loai` | Ghi đè tên loại (hiếm dùng) | `QUY CHẾ PHỐI HỢP` |
| `ke_duoi_ten_loai` | `khong` để bỏ đường kẻ khi văn bản không có trích yếu | giấy giới thiệu |
| `kinh_gui` | Nơi nhận ô 9a: công văn, tờ trình, báo cáo gửi cấp trên. Một nơi: cùng dòng; nhiều nơi: danh sách | Tự thêm `;` và `.` |
| `chuc_danh_dien` | Công điện: chức danh trước chữ `điện:` | `CHỦ TỊCH ỦY BAN NHÂN DÂN TỈNH` (danh sách nhận điện ghi ở `kinh_gui`) |
| `tham_quyen` | Thẩm quyền ban hành (nghị quyết, quyết định), in hoa đậm canh giữa | `CHỦ TỊCH ỦY BAN NHÂN DÂN XÃ AN BÌNH` |
| `can_cu` | Danh sách căn cứ, in nghiêng; tự thêm `;` cuối mỗi dòng và `.` dòng cuối | Dòng cuối thường là `Theo đề nghị của ...` |
| `mo_dau_dieu` | Dòng trước Điều 1 | Tự có `QUYẾT ĐỊNH:` (quyết định), `QUYẾT NGHỊ:` (nghị quyết) |
| `do_mat`, `do_khan`, `pham_vi_luu_hanh` | Dấu độ mật (ô 10a), độ khẩn (ô 10b), chỉ dẫn lưu hành (ô 11), có khung | `MẬT`; `HỎA TỐC`, `THƯỢNG KHẨN`, `KHẨN`; `LƯU HÀNH NỘI BỘ` |
| `quyen_han` | Quyền hạn ký (ô 7a). Có thể là danh sách | `TM. ỦY BAN NHÂN DÂN`, `KT. GIÁM ĐỐC`, `TL. BỘ TRƯỞNG`, `TUQ. CHỦ TỊCH`, `Q. GIÁM ĐỐC` |
| `chuc_vu` | Chức vụ người ký (ô 7b) | `PHÓ GIÁM ĐỐC`, `CHỦ TỊCH` |
| `nguoi_ky` | Họ tên người ký, **không** ghi học hàm, học vị | `Nguyễn Văn A` |
| `trai_quyen_han`, `trai_chuc_vu`, `trai_nguoi_ky` | Chữ ký thứ hai bên trái (biên bản: thư ký) | Nơi nhận khi đó đặt dưới hai chữ ký |
| `dong_cho_ky` | Số dòng trống để ký (mặc định 4) | |
| `noi_nhan` | Nơi nhận ô 9b, dòng cuối `Lưu: VT, <đơn vị soạn thảo>` | Tự thêm `- ` đầu dòng, `;` và `.` |
| `ky_hieu_nguoi_soan` | Ký hiệu người soạn thảo, số lượng bản phát hành (ô 12) | `NVA.05` |
| `dia_chi` | Địa chỉ, điện thoại, thư điện tử (ô 13), chân trang đầu. Danh sách | |
| `co_chu` | `14` (mặc định) hoặc `13`: cỡ chữ nội dung; Quốc hiệu, tên cơ quan đi cặp tương ứng | `13` → Quốc hiệu 12, Tiêu ngữ 13 |
| `le` | Lề `trên dưới trái phải` (mm) | Mặc định `20 20 30 15` |
| `gian_dong` | Giãn dòng nội dung, 1.0 đến 1.5 | Mặc định `1.15` |
| `lui_dau_dong` | Lùi đầu dòng (mm), 10 hoặc 12.7 | Mặc định `12.7` |
| `ket_thuc` | `khong` để không tự thêm `./.` cuối nội dung | |

## Nội dung

**Mỗi dòng là một đoạn.** Dòng trống bị bỏ qua. Không cần định dạng tay: chương trình nhận ra cấu trúc
theo đầu dòng và trình bày theo Phụ lục I NĐ 30.

| Viết | Kết quả |
|---|---|
| `Phần I`, `Chương II` (dòng tiếp theo là tiêu đề in hoa) | Canh giữa, đậm; tiêu đề in hoa đậm |
| `Mục 1`, `Tiểu mục 2` + tiêu đề | Như trên |
| `Điều 1. Phạm vi điều chỉnh` (tiêu đề ngắn, không dấu chấm cuối) | Cả dòng đậm, lùi đầu dòng |
| `Điều 1. Thành lập ... .` (câu nội dung) | Chỉ `Điều 1.` đậm |
| `I. TÊN MỤC` | Đậm |
| `1. Tiêu đề khoản` (không dấu câu cuối) | Đậm |
| `1. Nội dung khoản.`, `a) ...`, `- ...`, `+ ...` | Chữ thường, lùi đầu dòng, canh đều |
| Dòng toàn chữ in hoa ngắn, ví dụ `QUYẾT ĐỊNH:` | Canh giữa, đậm |
| `**chữ đậm**`, `*chữ nghiêng*` | Đậm, nghiêng trong đoạn |
| Bảng Markdown `\| a \| b \|` | Bảng có khung, dòng đầu đậm |
| `[ngắt trang]` | Sang trang |
| `[phụ lục]` rồi 2-3 dòng: `Phụ lục I` / `TÊN PHỤ LỤC` / `(Kèm theo ...)` | Trang mới, số trang riêng, tiêu đề phụ lục đúng PL I mục III.1 |

Đoạn cuối của văn bản chính và của mỗi phụ lục được thêm `./.` (như các mẫu ở Phụ lục III).
Văn bản ban hành kèm theo quyết định (quy chế, quy định...) cũng đặt sau `[phụ lục]`, dòng đầu là tên loại,
ví dụ `QUY CHẾ`.
