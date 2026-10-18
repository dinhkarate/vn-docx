# Phần mềm, công cụ hỗ trợ trình bày văn bản theo Nghị định 30 (tham khảo)

Khảo sát ngày 17/10/2026. Tính năng ghi theo **tuyên bố của nhà phát triển** hoặc trang giới thiệu, chưa kiểm thử
độc lập. Giá và phiên bản có thể đã đổi; mở link để xem bản mới nhất. Không công cụ nào dưới đây là công cụ
**chính thức** của Bộ Nội vụ hay Cục Văn thư và Lưu trữ nhà nước: đến thời điểm khảo sát chưa thấy add-in hay
mẫu `.dotx` chính thức nào do cơ quan nhà nước phát hành.

> Văn bản mật, văn bản có thông tin cá nhân: **không** tải lên công cụ web hay gửi qua AI bên ngoài. Dùng công cụ
> chạy trên máy (add-in offline, `vn-docx`) hoặc hệ thống của cơ quan.

## 1. Add-in Word, phần mềm cài trên máy

| Công cụ | Chức năng chính | Giá | Nền tảng | Nhận xét |
|---|---|---|---|---|
| **V-FORMAT** (TAFINEX) [tafinex.com](https://tafinex.com/san-pham-chi-tiet?slug=vformat), [ngdaitai.com](https://ngdaitai.com/) | Chuẩn hóa một chạm: lề, phông, giãn dòng, khối tiêu đề, chữ ký, nơi nhận, số trang; chèn thành phần thể thức; thư viện mẫu; chuyển TCVN3/VNI sang Unicode. Có hồ sơ NĐ 30, HD 05 (Đảng), một số quy chuẩn ngành | Trả phí: khoảng 149.000 đ/6 tháng, 229.000 đ/năm, 499.000 đ trọn đời; có dùng thử | Windows 10/11, Word 2013+ / Microsoft 365, chạy offline | Có pháp nhân, lịch sử phiên bản (v2.2.9 ngày 15/10/2026). Một số thông số trên trang chưa thống nhất; tự nhận không thay thế thẩm định |
| **OfficeKit** [officekit.wpspeedy.net](https://officekit.wpspeedy.net/) | "Check & Fix" hơn 20 lỗi NĐ 30, đọc số thành chữ, chuyển phông, thư viện mẫu; "AI Audit" qua API AI bên ngoài | Dùng thử 7 ngày; 249.000-499.000 đ/năm | Word, Excel (theo FAQ: Windows, Mac, web) | Bản Beta, trang không nêu đơn vị phát triển. AI Audit gửi nội dung ra ngoài: **không dùng cho văn bản mật** |
| Add-in "Chuẩn hóa thể thức" (ngoctien.id.vn) | Kiểm tra thể thức, chính tả, chuyển phông, đặt lề, số trang; hồ sơ NĐ 30, HD 05 | Bài giới thiệu ghi miễn phí | Word, Windows | Tác giả không rõ danh tính; bài hướng dẫn trên một trang tin khuyên loại trừ thư mục khỏi Windows Defender: **rủi ro bảo mật, cân nhắc kỹ** |
| UniKey Toolkit [unikey.org](https://www.unikey.org/) | Chuyển mã TCVN3 (.VnTime), VNI sang Unicode (TCVN 6909:2001) | Miễn phí | Windows | Chuyển qua clipboard nên mất định dạng; phải chọn đúng bảng mã nguồn |

## 2. Công cụ web

| Công cụ | Chức năng | Giá | Nhận xét |
|---|---|---|---|
| **Chuẩn Word** [chuanword.vn](https://www.chuanword.vn/) | Tải lên .doc/.docx, kiểm tra khoảng 22-26 tiêu chí NĐ 30, tự sửa, chấm điểm; AI kiểm tra chính tả, viết hoa | Miễn phí (thời điểm khảo sát) | Không nêu đơn vị phát triển, chính sách lưu file chưa rõ: **không dùng cho văn bản mật, nhạy cảm** |
| mauvanban.vn | Bán mẫu Word theo NĐ 30 | 12.000-35.000 đ/mẫu | Tự nhận "tuân thủ 100%", chưa kiểm chứng |
| Tiện ích "Căn chỉnh NĐ30" (giaoviendoimoi) | Tải file lên, nhận file đã căn chỉnh | Từng miễn phí, sau chuyển trả phí | Thông tin cũ, có thể đã ngừng |

## 3. Hệ thống quản lý văn bản và điều hành (e-office)

VNPT iOffice, Viettel VOffice / mOffice và các hệ thống tương tự lo **cấp số, trình ký, ký số, chèn hình chữ ký
và dấu** đúng vị trí NĐ 30 (ô 7c, 8; ký số góc trên phải cho phụ lục). Chưa thấy tài liệu công khai cho biết các
hệ thống này tự định dạng nội dung .docx theo NĐ 30: phần trình bày dự thảo vẫn do người soạn làm trong Word.

## 4. Mã nguồn mở, AI skill

Phần lớn ra đời năm 2026, ít người dùng. Đọc mã trước khi chạy.

| Dự án | Mô tả | Giấy phép | Nhận xét |
|---|---|---|---|
| [kanazawahere/nd30](https://github.com/kanazawahere/nd30) | Phân loại văn bản, hỏi dữ liệu, sinh .docx, có script kiểm tra và test | MIT | Đáng tham khảo; cần Python và thư viện |
| [TheDevio/ND30_Skills](https://github.com/TheDevio/ND30_Skills) | Skill: thông số, viết hoa, thư viện mẫu, `validate_docx.py` (python-docx), checklist | MIT | Đang hoạt động |
| [bGiaHuy/vietnamese-docs-style](https://github.com/bGiaHuy/vietnamese-docs-style) | Điền mẫu `{{token}}`, kiểm tra theo hồ sơ, dựng bằng LibreOffice | Chưa có giấy phép | Thiết kế cẩn thận; chưa có giấy phép nên chưa được phép tái sử dụng mã |
| [biencuong/vbhc](https://github.com/biencuong/vbhc) | Skill + MCP: phân loại, điền mẫu, kiểm tra, gợi ý nơi nhận | Unlicense | Phụ thuộc máy chủ bên ngoài, cần khóa truy cập |
| [yinlin49/nd30-skills](https://github.com/yinlin49/nd30-skills) | 17 skill theo các mẫu 1.1-3.2 Phụ lục III | MIT | Mới, chưa có người dùng |
| [hoanglong149/office-bidding-skills](https://github.com/hoanglong149/office-bidding-skills) | Skill `nd30-forms` (10 mẫu) và hồ sơ dự thầu | MIT | Mới |

Không tìm thấy GPT công khai (ChatGPT) đủ tin cậy cho NĐ 30. Nhiều skill chỉ "nhớ" NĐ 30, chưa cập nhật mô hình
chính quyền 2 cấp (từ 01/7/2025) và Hướng dẫn 05-HD/VPTW (văn bản Đảng, 27/5/2026).

## 5. vn-docx khác gì

| | vn-docx | Add-in Word | Công cụ web | Skill khác |
|---|---|---|---|---|
| Soạn văn bản mới từ yêu cầu bằng lời | Có (qua AI agent) | Không | Không | Một số |
| Kiểm tra file Word có sẵn, chỉ rõ căn cứ | Có | Có | Có | Một số |
| Tự sửa lề, phông, số trang | Có | Có | Có | Ít |
| Chạy trên máy, không gửi văn bản đi đâu | Có (AI chỉ đọc khi bạn dùng agent) | Có | Không | Tùy |
| Cài thêm thư viện | Không (Python chuẩn) | Cài add-in | Không | Thường cần python-docx |
| Windows, macOS, Linux, Claude Cowork | Có | Thường chỉ Windows | Trình duyệt | Tùy |
| Chính quyền 2 cấp, mẫu cấp xã | Có | Tùy bản | Chưa rõ | Ít |
| Mã nguồn mở, miễn phí | MIT | Không | Không | Phần lớn |
| Bài tập "tìm lỗi" cho lớp học | Có (`bai-tap`) | Không | Không | Không |

Giới hạn của vn-docx: chưa dựng văn bản quy phạm pháp luật (NĐ 78/2025) và văn bản của Đảng (HD 05-HD/VPTW);
chưa chèn hình chữ ký số, con dấu (việc của hệ thống ký số); không thay thế khâu thẩm định nội dung, thẩm quyền.
