# Bảo toàn cấu trúc Excel khi chỉ sửa ô

`scripts/preserve_xlsx.py` là bước đóng gói sau khi đã tác giả và xuất workbook bằng công cụ Spreadsheets. Nó không thay thế công cụ tác giả hoặc mở rộng bảng.

## Đầu vào

```text
python <skill>/scripts/preserve_xlsx.py --base <base.xlsx> --authored <authored.xlsx> --plan <edits.json> --output <new.xlsx>
```

Đọc `--help` của phiên bản script trong Skill. Plan liệt kê chính xác các ô được phép đổi, theo sheet có sẵn. Mỗi ô chứa `value` hoặc `formula`, có thể chứa `cache` là kết quả đã đối chiếu độc lập. Giữ dữ liệu ngày dưới dạng serial Excel hoặc ngày được công cụ tác giả xuất đúng kiểu.

Ví dụ cấu trúc với số liệu giả:

```json
{
  "Sheet1": {
    "F160": {"value": 10},
    "G160": {"value": 1000000},
    "I160": {"formula": "=G160+80000", "cache": 1080000}
  },
  "Sheet2": {}
}
```

Plan phải khớp với các ô trong bản authored, không chỉ đúng trong JSON. Với ô chưa tồn tại trong XML của bản nền, thêm `"style_from": "G159"` (cùng sheet, chọn đúng cột và kiểu của dòng mẫu) vào entry đích. Không đưa ô chưa sửa vào plan để lách kiểm tra khác biệt. Không ghi vào file base, authored hoặc chứng từ nguồn.

## Giới hạn có chủ ý

Helper sao chép nguyên các phần của gói XLSX và chỉ thay nội dung ô được cho phép. Font, fill, style ID, kích thước, table range, validation, vùng in và sheet views theo base. Các sheet/dòng phải đã tồn tại; ô mới trong dòng có sẵn cần nguồn style rõ ràng.

Không dùng helper khi đã thêm/xóa/sắp xếp sheet, chèn dòng, mở rộng bảng, đổi merge, đổi validation hoặc thay layout. Khi đó dùng công cụ hỗ trợ thao tác native cần thiết, rồi kiểm tra toàn bộ phần cấu trúc có liên quan. Không loại bỏ hoặc bỏ qua kiểm tra của helper chỉ để có file kết quả.

Nếu helper từ chối shared formula hoặc tính năng không hỗ trợ trong vùng sửa, giữ nguyên bản nền và dùng cách chỉnh phù hợp với tính năng đó. Đừng làm phẳng toàn bộ workbook để sửa vài ô.

## Kiểm tra bản kết quả

1. Recalculate bản tác giả và kiểm tra các ô, số tiền nguồn trước khi chạy helper.
2. Helper kiểm tra plan với authored, kiểu ô, cache, sự toàn vẹn ZIP/XML và vùng không sửa.
3. Sau khi đóng gói, nhập lại **file kết quả**, kiểm tra các ô mới và render vùng thay đổi. Xác nhận tổng tiền từng hóa đơn và tổng riêng mua/bán.
4. Đối chiếu hash các tệp nguồn với manifest. Giữ nguyên bản nền cho đến sau khi bàn giao; quy trình này mặc định tạo phiên bản mới.
