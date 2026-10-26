---
name: lingo-in-out-invoices
description: Điền và cập nhật sổ Excel hóa đơn mua vào, bán ra của LINGO từ chứng từ trong IN theo mẫu có sẵn trong OUT, giữ nguyên tệp nguồn và định dạng. Dùng cho yêu cầu cập nhật chứng từ mới, đối chiếu hóa đơn hoặc tiếp tục quy trình IN–OUT của LINGO.
---

# LINGO IN–OUT

Chuyển chứng từ vào sổ hóa đơn theo mẫu người dùng đã chọn. Bao gồm lần nhập đầu và các lần bổ sung. Mặc định tìm `IN` và `OUT` trong workspace hiện tại; ưu tiên đường dẫn người dùng nêu. Nếu chưa xác định được thư mục, hỏi đường dẫn trước khi xử lý.

## Ràng buộc của quy trình

- Chỉ đọc `IN`. Giữ nguyên mọi chứng từ, tệp nén, mẫu gốc và bản Excel của lần trước. Lưu bản kết quả mới trong `OUT`; chỉ cập nhật đè khi người dùng yêu cầu rõ.
- Giữ tên/thứ tự sheet, màu, font, kiểu bảng, độ rộng cột, chiều cao dòng, định dạng ngày/số, công thức và ghi chú hiện có ngoài phạm vi sửa. Không đổi tên sheet hoặc thêm dashboard để phục vụ việc nhập liệu.
- Nhập dữ liệu từ nội dung chứng từ. Tên thư mục và tên tệp chỉ giúp tìm kiếm, có thể ghi sai ngày, ký hiệu hoặc số hóa đơn.
- Không tính một hóa đơn hai lần khi xuất hiện dưới dạng PDF, XML, ZIP hoặc bản sao. Khóa đối chiếu gồm **MST bên phát hành + ký hiệu + số hóa đơn chuẩn hóa**. Không dùng riêng số hóa đơn.
- Phiếu xuất kho kiêm vận chuyển nội bộ, ủy nhiệm chi, email thông báo và bảng kê không phải một dòng doanh thu/chi phí hóa đơn độc lập. Liên kết chúng vào ghi chú của hóa đơn có căn cứ phù hợp.
- Thiếu số liệu thì để trống hoặc ghi rõ ở cột ghi chú sẵn có. Không suy ra thuế suất, hóa đơn đầu vào hoặc tình trạng thanh toán chỉ từ tên tệp hay số tiền bằng nhau.

## Thực hiện

1. **Chọn bản nền.** Liệt kê `IN`, các Excel trong `OUT` và trạng thái lần trước nếu có. Đọc workbook để xác nhận bản đầy đủ mới nhất; không chỉ chọn theo thời gian sửa tệp. Trong hồ sơ ban đầu, mẫu là `LINGO 2026_HDMV HDBR.xlsx`, các bản sau có hậu tố `_DA_DIEN`, `_CAP_NHAT`. Không dùng tên hoặc số lượng này làm giới hạn cho lần chạy mới.
2. **Đọc mẫu trước khi sửa.** Áp dụng kỹ năng Spreadsheets hiện có, tải runtime qua `load_workspace_dependencies`, đọc công thức và render vùng sẽ sửa. Đọc [references/workbook-layout.md](references/workbook-layout.md) cho cách ánh xạ cột của mẫu LINGO. Nếu mẫu thực tế khác, theo mẫu thực tế.
3. **Quét chứng từ.** Chạy helper sau bằng Python của runtime, ghi kết quả vào thư mục tạm của workspace:

   ```text
   python <skill>/scripts/scan_invoices.py --input <workspace>/IN --output <tmp>/scan.json
   ```

   Có trạng thái lần trước thì thêm `--previous <previous-scan.json>`. Đọc `--help` và JSON thực tế của helper. Kiểm tra lỗi, PDF chưa phân tích, tệp thay đổi và các khóa trùng; không coi quét xong là đã nhập đủ. Helper không ghi vào `IN` và không giải nén trực tiếp lên thư mục nguồn. Khi cần đọc PDF, áp dụng kỹ năng PDF.
4. **Đối chiếu và lập kế hoạch sửa.** So sánh khóa chứng từ với các dòng workbook. Dữ liệu mới được thêm vào vùng trống đúng loại hóa đơn; chứng từ đổi nội dung thì đối chiếu trước khi sửa dòng hiện có. Giữ ghi chú nghiệp vụ cũ. XML chuẩn có độ ưu tiên cao hơn PDF của cùng hóa đơn, nhưng chênh lệch giữa các bản phải được nêu rõ. Đối chiếu bản PDF cần dùng toàn bộ trang, nhất là hóa đơn nhiều trang.
5. **Điền dữ liệu.** Tác giả workbook bằng công cụ và runtime của kỹ năng Spreadsheets. Ngày/số tiền/thuế suất phải là dữ liệu có kiểu, mã số là định danh. Dùng tiền thanh toán và tiền thuế trên hóa đơn làm số kiểm soát; không để công thức tỷ lệ làm mất sai số làm tròn của chứng từ. Cột VAT chứa thuế suất, không phải tiền thuế. Với `KCT`, hóa đơn bán hàng không tách VAT hoặc nhiều thuế suất, giữ đúng ý nghĩa và ghi chú gọn trong mẫu.
6. **Giữ định dạng khi lưu.** Khi sửa nội dung trong các dòng có sẵn, sau khi xuất bản `authored.xlsx`, có thể chạy `scripts/preserve_xlsx.py` để chỉ chuyển nội dung các ô đã được cho phép về cấu trúc gốc. Đọc [references/preservation.md](references/preservation.md) trước khi dùng. Nếu hết dòng dự phòng, mở rộng bảng theo kiểu sẵn có bằng công cụ hỗ trợ native table; cập nhật công thức, bộ lọc, vùng in có liên quan và kiểm tra lại. Không cắt bớt hóa đơn để vừa mẫu. Helper bảo toàn ô không thay thế cho thao tác mở rộng cấu trúc.
7. **Kiểm tra rồi bàn giao.** Đối chiếu số hóa đơn và tổng tiền mua/bán riêng, số tiền từng dòng, ngày/ký hiệu/MST, công thức đầu–giữa–cuối vùng mới và các liên kết phụ thuộc. Render vùng sửa, mở lại bản đã lưu, kiểm tra không có lỗi công thức mới và so sánh định dạng/native features với bản nền. Xác nhận hash chứng từ nguồn không thay đổi. Nêu số dòng bổ sung và các vấn đề còn cần người dùng kiểm tra, kèm liên kết bản Excel kết quả.

## Cập nhật lặp lại

- Manifest hash chỉ giúp phát hiện tệp mới/đổi/xóa; workbook và khóa hóa đơn vẫn là căn cứ chống nhập trùng. Một hóa đơn cũ được gửi lại dưới tên mới không tạo dòng mới.
- Nếu có tệp bị xóa khỏi `IN`, không tự xóa dòng trong sổ. Báo sự khác biệt nếu ảnh hưởng đối chiếu.
- Với hóa đơn điều chỉnh/thay thế, giữ quan hệ với hóa đơn gốc theo nội dung chứng từ; không tự xóa hoặc cộng đè như hóa đơn thông thường.
- Có thể lưu scan đã đối chiếu vào `.lingo-in-out/` của workspace sau khi bản kết quả đã kiểm tra thành công. Không ghi đè trạng thái tốt trước khi cập nhật hoàn tất. Không giữ cache nội dung tài chính trong thư mục Skill.
- Khi dữ liệu không thay đổi, báo không phát sinh cập nhật; không xuất thêm bản Excel giống hệt.

## Ghép chứng từ liên quan

Theo số PO/hợp đồng hoặc số hóa đơn được dẫn rõ trước. Nếu chỉ khớp hàng hóa, quy cách, số lượng và giá, ghi **“Khớp hàng: HĐ …”** để phân biệt với quan hệ được chứng từ xác nhận. Một hóa đơn mua có thể khớp nhiều hóa đơn bán; không yêu cầu tổng tiền hai bên bằng nhau khi có chênh lệch giá bán. Nếu chưa đủ căn cứ, ghi “Chưa xác định HĐ đầu vào”.

Các helper hỗ trợ quy trình làm việc với mẫu thực tế. Chứng từ chưa phân tích được vẫn cần kiểm tra riêng.
