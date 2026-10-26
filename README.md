# GPT-SKILL
Gpt Skill for my job

## Skills

### [LINGO IN–OUT](skills/lingo-in-out-invoices/SKILL.md)

Đọc chứng từ XML/PDF và ZIP lồng nhau trong `IN`, đối chiếu rồi cập nhật sổ hóa đơn Excel theo mẫu trong `OUT`. Hỗ trợ nhập lần đầu và bổ sung chứng từ, chống nhập trùng, phân biệt hóa đơn với phiếu xuất kho/ủy nhiệm chi, giữ tệp nguồn và định dạng mẫu.

**Thư mục Skill:** `skills/lingo-in-out-invoices`

Để cài thủ công, sao chép toàn bộ thư mục `lingo-in-out-invoices` vào thư mục skills cá nhân của Codex, mặc định `~/.codex/skills/`.

Ví dụ sử dụng:

> Dùng $lingo-in-out-invoices kiểm tra chứng từ mới trong IN và cập nhật Excel trong OUT, giữ nguyên nguồn và định dạng.

Skill sử dụng các kỹ năng Spreadsheets/PDF và runtime của Codex. Hai helper đã được kiểm thử với Python 3.12; đọc PDF cần thư viện `pypdf`.

Kho này chứa mã nguồn và hướng dẫn của Skill. Các tệp chứng từ và sổ Excel được xử lý tại workspace người dùng.
