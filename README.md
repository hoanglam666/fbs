# In hàng loạt theo mẫu (Google Apps Script)

File `print_batch.gs` hỗ trợ:
- Tạo trang in hàng loạt từ dữ liệu `Table_FBS`.
- Xuất PDF để tải về theo khổ **A6 ngang** (landscape).

## Đã chỉnh theo mẫu thực tế
Script đang map theo đúng layout mẫu bạn gửi:
- Vùng template: `A1:E11`
- `Mã Đơn` -> ô `B3`
- `Supplier name` -> ô `B4`
- `Mã Thùng` -> ô `C5` (dạng `trang hiện tại / tổng`)
- Dòng dữ liệu -> `A8:E8`
- Tổng -> `E9`
- Người in -> `B11` (tự động lấy tài khoản đang chạy script)
- Ngày in -> `D11` (tự động lấy thời gian hiện tại)

## Cột dữ liệu đầu vào
Sheet `Table_FBS` cần có các cột (hỗ trợ nhiều biến thể tên cột):
- `Ma_sku`
- `Ten_sp`
- `Phan_loai`
- `So_luong`
- `Ma_kien`
- `Ma_don`  ✅
- `Supplier_name` ✅

## Cách dùng
1. Mở Google Sheet → **Extensions → Apps Script**.
2. Dán nội dung `print_batch.gs`.
3. Reload Google Sheet.
4. Vào menu **🖨️ In tem**:
   - **Tạo sheet in hàng loạt**: tạo sheet `IN_HANG_LOAT` để kiểm tra.
   - **Tạo PDF A6 ngang (download)**: sinh PDF và hiển thị link tải.

## Gán nút bấm trên Google Sheet
- Insert → Drawing (hoặc Image) → gán script:
  - `inHangLoatVaXuatPdf` (khuyên dùng: vừa tạo sheet vừa xuất PDF)
  - hoặc `inHangLoatTheoMau`


## Sửa lỗi đã gặp
- Đã sửa lỗi `Bạn không thể xóa tất cả trang tính trong tài liệu` khi xuất PDF bằng cách tạo trước sheet từ `Mau_in`, sau đó mới xóa sheet mặc định.
- Đã sửa lỗi `Dải ô đích và dải ô nguồn phải trên cùng một bảng tính` bằng cách không copy `Range` giữa hai spreadsheet, thay vào đó copy cả sheet template (`copyTo`).
