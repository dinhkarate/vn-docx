# Đóng góp cho vn-docx

Cảm ơn bạn. Đóng góp có giá trị nhất là **chỉ ra chỗ trình bày chưa đúng quy định**, kèm căn cứ.

## Báo lỗi thể thức, đề xuất mẫu

Mở [issue](https://github.com/dinhkarate/vn-docx/issues/new/choose) theo mẫu có sẵn. Ghi rõ:

- Văn bản nào, thành phần nào sai (đính kèm file `.docx` đã xóa thông tin cá nhân, hoặc file `noi-dung.md`).
- Căn cứ: điều, khoản, mục của Nghị định 30/2020/NĐ-CP hoặc văn bản hướng dẫn (số, ngày, cơ quan ban hành).
- Cách trình bày đúng theo bạn.

Không gửi văn bản mật, văn bản có thông tin cá nhân lên issue.

## Sửa mã

```bash
git clone https://github.com/dinhkarate/vn-docx && cd vn-docx
python3 -m unittest discover -s tests       # phải chạy được trên Python 3.9
python3 skills/van-ban-hanh-chinh/scripts/vanban.py tao skills/van-ban-hanh-chinh/examples/cong-van.md /tmp/cv.docx
```

Nguyên tắc:

- `vanban.py` **chỉ dùng thư viện chuẩn Python**: không thêm `python-docx`, `lxml`... để skill chạy được ở mọi nơi,
  kể cả Claude Cowork và máy cơ quan không cài được thư viện.
- Quy định mới trong `the-thuc-nd30.md` phải có trích dẫn (mục của Phụ lục I, hoặc số văn bản).
- Thêm mẫu: đặt file vào `skills/van-ban-hanh-chinh/examples/`, ghi chú mẫu số mấy của Phụ lục III ở dòng đầu khối
  `---`. Test tự kiểm tra mọi mẫu phải ra 0 lỗi, 0 cảnh báo.
- Thêm kiểu kiểm tra: thêm test có file sai (xem `loi_mau.py`) và file đúng (các mẫu) để tránh báo nhầm.
- Mở file kết quả bằng Word hoặc LibreOffice và xem bằng mắt trước khi gửi pull request.

## Phát hành (người duy trì)

1. Tăng phiên bản ở `package.json`, `skills/van-ban-hanh-chinh/.claude-plugin/plugin.json`, `SKILL.md`
   (`version:`), `VERSION` trong `vanban.py`; thêm mục mới ở đầu `CHANGELOG.md`. Test `Release` kiểm tra bốn chỗ
   này khớp nhau.
2. Gắn thẻ `git tag v0.2.0 && git push --tags`. GitHub Actions chạy test, đăng npm (trusted publishing) và đính
   `van-ban-hanh-chinh.zip` vào GitHub Release.

Đã cấu hình sẵn (bản 0.1.0 đăng tay, sau đó bật trusted publishing). Nếu phải làm lại, ví dụ khi chuyển repo:

```bash
npm trust github vn-docx --file release.yml --repo dinhkarate/vn-docx --env npm --allow-publish
gh variable set NPM_PUBLISH --body true   # workflow chỉ đăng npm khi biến này là true
```
