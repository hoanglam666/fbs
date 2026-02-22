# In hàng loạt theo mẫu (Google Apps Script)

File `print_batch.gs` hỗ trợ:
- Tạo trang in hàng loạt từ dữ liệu `Table_FBS`.
- Xuất PDF để tải về theo khổ **A6 ngang** (landscape).

## Logic in
- Mỗi dòng SKU có `Ma_kien = N` sẽ sinh ra `N` tem.
- Ô `Mã Thùng` hiển thị `trang hiện tại / tổng` (ví dụ: `2 / 5`).
- Số lượng tem mỗi kiện được chia từ `So_luong` (kiện cuối nhận phần dư).

## Cách dùng
1. Mở Google Sheet → **Extensions → Apps Script**.
2. Dán nội dung `print_batch.gs`.
3. Chỉnh `CONFIG` nếu tên sheet/vị trí ô khác thực tế.
4. Reload Google Sheet.
5. Vào menu **🖨️ In tem**:
   - **Tạo sheet in hàng loạt**: tạo sheet `IN_HANG_LOAT` để kiểm tra.
   - **Tạo PDF A6 ngang (download)**: sinh PDF và hiển thị link tải.

## Gán nút bấm trên Google Sheet
- Insert → Drawing (hoặc Image) → gán script:
  - `inHangLoatVaXuatPdf` (khuyên dùng: vừa tạo sheet vừa xuất PDF)
  - hoặc `inHangLoatTheoMau`

## Lưu ý
- File PDF sẽ được lưu vào Google Drive của tài khoản chạy script.
- Script tạo 1 spreadsheet tạm để gom từng tem thành từng trang PDF, sau đó tự xóa file tạm.
- Nếu cột `Ma_don`, `Supplier_name` không có thì các ô tương ứng để trống.
