## Phần 1. Phân tích: Xác định Record và Field trong bảng Excel giao dịch

Trong bảng cơ sở dữ liệu Excel đối soát giao dịch của ví điện tử **PayKX**, các khái niệm được phân định rõ ràng như sau:

*   **Field (Trường dữ liệu / Cột):** Là các thuộc tính đơn lẻ cấu thành nên thông tin của một giao dịch. 
*   **Record (Bản ghi / Dòng dữ liệu):** Là tập hợp toàn bộ thông tin chi tiết của **một giao dịch độc lập** nằm trên một hàng ngang.
## Phần 3. Kiểm thử: Quy trình 4 điểm tìm nguyên nhân lỗi `#N/A` khi dùng `VLOOKUP`
Khi hàm `VLOOKUP` trả về lỗi `#N/A` (không tìm thấy giá trị tra cứu trong bảng đối soát), thực hiện kiểm tra tuần tự theo quy trình 4 điểm sau:
1.  **Kiểm tra giá trị tìm kiếm (`lookup_value`):** Xác thực xem giá trị cần tìm có thực sự tồn tại trong cột đầu tiên của vùng bảng tra cứu (`table_array`) hay không.
2.  **Kiểm tra khoảng trắng thừa (Leading/Trailing Spaces):** Dữ liệu thường bị lệch do lỗi gõ phím hoặc import từ hệ thống ngoài tạo ra khoảng trắng ẩn (ví dụ: `"TX1001 "` khác với `"TX1001"`). Cần sử dụng hàm `TRIM` để làm sạch dữ liệu ở cả cột nguồn và cột tra cứu.
3.  **Kiểm tra tính cố định vùng tham chiếu (`table_array`):** Khi kéo công thức xuống các ô bên dưới, nếu quên khóa vùng tra cứu bằng dấu đô la (ví dụ quên viết thành `$A$2:$B$100`), vùng dữ liệu sẽ bị trôi, khiến hàm không quét đúng phạm vi bảng.
4.  **Kiểm tra tham số tìm kiếm chính xác (`Range_lookup`):** Đảm bảo tham số cuối cùng của hàm `VLOOKUP` được thiết lập là `0` để bắt buộc tìm kiếm chính xác tuyệt đối. Nếu để `TRUE` hoặc bỏ trống, hàm sẽ trả về kết quả sai hoặc `#N/A` khi dữ liệu chưa được sắp xếp tăng dần.
