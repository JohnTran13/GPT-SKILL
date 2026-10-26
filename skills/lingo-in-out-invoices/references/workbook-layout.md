# Mẫu sổ hóa đơn LINGO

Đọc file đang dùng trước khi nhập. Các vị trí dưới đây mô tả mẫu gốc, không phải giới hạn dòng cố định cho các năm hoặc bản mẫu khác.

## Cấu trúc đã quan sát

`Sheet1` có hai bảng: mua vào ở phía trên và bán ra ở phía dưới. Mẫu gốc có tiêu đề mua vào tại dòng 2 và bán ra tại dòng 150. `Sheet2` là danh mục mã đối tác, tên đầy đủ, loại hàng/dịch vụ. Các bảng Excel dùng structured references và công thức VLOOKUP; giữ native table và công thức chưa cần sửa.

| Cột | Tiêu đề | Nội dung |
| --- | --- | --- |
| B | STT | Đánh số trong từng nhóm mua/bán; theo quy ước hoặc công thức sẵn có |
| C | LOẠI H.ĐƠN | `H.ĐƠN MUA VÀO` hoặc `H.ĐƠN BÁN RA` |
| D | NGÀY H.ĐƠN | Ngày lập trên XML/PDF, kiểu ngày Excel và giữ format của ô |
| E | K.HIỆU H.ĐƠN | Ký hiệu đầy đủ, thí dụ ghép KHMSHDon + KHHDon nếu XML tách hai phần |
| F | SỐ H.ĐƠN | Định danh; chuẩn hóa số 0 đầu chỉ để ghép khóa, giữ cách hiển thị của mẫu |
| G | GIÁ TRỊ | Tiền hàng trước VAT khi chứng từ có tách; hóa đơn bán hàng không tách thuế ghi giá trị chứng từ và chú thích |
| H | VAT | Thuế suất phần trăm dạng số, hoặc trạng thái như KCT; không ghi tiền thuế vào đây |
| I | THÀNH TIỀN | Tổng thanh toán khớp chứng từ, gồm cách làm tròn của chứng từ |
| J | NCC/K.HÀNG V.TẮT | Mã đối tác có trong danh mục; mua dùng bên bán, bán dùng bên mua |
| K | NCC ĐẦY ĐỦ | Tên đối tác, thường lấy bằng công thức theo mã J |
| L | LOẠI HÀNG HÓA /DV (mua) | Loại nghiệp vụ theo danh mục; giữ giá trị riêng đã có |
| L | HÓA ĐƠN ĐẦU VÀO (bán) | Tham chiếu có căn cứ hoặc kết quả khớp hàng được ghi rõ là suy luận |
| M | GHI CHÚ | Giữ ghi chú cũ; bổ sung thiếu nguồn, sai lệch thuế suất, PXK/PO liên quan |

Danh mục `Sheet2`: C = mã viết tắt, D = tên đầy đủ, E = loại hàng/dịch vụ. Chỉ thêm đối tác mới khi cần, sao chép kiểu của dòng liền kề và bảo đảm phạm vi tra cứu bao gồm dòng mới. Không tự đổi mã đối tác đã dùng.

Mẫu không có cột MST. Để tái dựng khóa dòng cũ, đối chiếu ký hiệu/số với bên phát hành trong nguồn, đồng thời kiểm tra mã/tên đối tác của dòng. Nếu cùng ký hiệu/số thuộc nhiều MST hoặc tên đối tác chưa khớp duy nhất, đánh dấu cần kiểm tra thay vì tự coi là cùng hóa đơn. Hóa đơn bán ra dùng MST bên bán LINGO cho khóa; mã khách hàng ở J là bên mua.

## Tiền thuế và tổng thanh toán

- Giữ tiền thuế/tiền tổng nguyên gốc trong dữ liệu đối chiếu tạm. Mẫu không có cột tiền thuế riêng.
- Nếu phép tính `G * (1 + H)` cho kết quả lẻ khác tiền tổng do làm tròn, dùng cách tính/công thức phản ánh được chứng từ. Ví dụ `=G<row>+<tiền_thuế_gốc>` có thể giữ tổng gốc và tính động theo G. Không khẳng định đây là thuế tính lại theo H khi H thay đổi.
- Khi có thuế suất in khác tiền thuế thực tế, không sửa chứng từ hoặc chọn tỷ lệ theo phỏng đoán. Ghi rõ hai giá trị cần kiểm tra.
- Thuế suất `0%`, `KCT`, `KKKNT`, thuế suất khác và không tách VAT là các trạng thái khác nhau. Không đồng loạt đổi thành số 0.
- Kiểm tra `tiền hàng + VAT = thanh toán` với Decimal và sai số do chứng từ quy định. Đọc phí/chiết khấu/điều chỉnh trước khi kết luận có lỗi.

## Nguồn trong IN

- `01_INCOME ...`: hóa đơn mua vào và tài liệu liên quan. Bên mua thường là LINGO, nhưng phải kiểm tra nội dung.
- `02_OUT COME ...`: hóa đơn bán ra và tài liệu giao hàng. Bên bán thường là LINGO. Phân nhóm thư mục không chứng minh bản chất của PDF.
- Một hóa đơn có thể được giữ dạng XML trực tiếp, XML trong ZIP, ZIP lồng ZIP và PDF. Các bản này là một chứng từ khi cùng khóa bên phát hành/ký hiệu/số.
- Email `.msg`, kho `.rar` và bảng kê Excel được giữ nguyên. Khi chúng có thể chứa hóa đơn chưa có trong các nguồn đã đọc, dùng công cụ phù hợp để kiểm tra và nêu rõ giới hạn còn lại; đừng đánh dấu toàn bộ IN đã nhập đủ chỉ dựa trên số XML.
