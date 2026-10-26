# BÀI 5: CHUẨN HÓA QUY TRÌNH BÁO CÁO SỰ CỐ AN NINH MẠNG NETGUARD

## 1. Mục tiêu

Bài thực hành xây dựng bộ tài liệu báo cáo sự cố an ninh mạng cho đơn vị NetGuard, gồm:

- Soạn thảo báo cáo sự cố bằng Microsoft Word, sử dụng Heading Styles và mục lục tự động.
- Quản lý nhật ký hệ thống bằng Microsoft Excel, sử dụng VLOOKUP kết hợp TRIM để tra cứu mã mã độc.
- Thiết kế sơ đồ phân cấp phòng thủ an ninh mạng 3 tầng bằng Microsoft PowerPoint.
- Chuẩn hóa tài liệu và dữ liệu để hỗ trợ quá trình phát hiện, xử lý và phòng ngừa sự cố.

## 2. Phần 1: Phân tích

### Câu 1: Có thể xem AutoFilter là “Google Search của bảng tính” không? Vì sao?

Có thể xem AutoFilter là một công cụ tìm kiếm và lọc dữ liệu trong bảng tính. Khi bật AutoFilter, người dùng có thể tìm kiếm theo từ khóa hoặc chọn điều kiện lọc để chỉ hiển thị những dòng dữ liệu phù hợp.

Ví dụ: Trong bảng nhật ký hệ thống NetGuard, có thể lọc các cảnh báo có mức độ nguy hiểm là “Nghiêm trọng” hoặc tìm các dòng có mã mã độc `VIR002`.

Tuy nhiên, AutoFilter không hoàn toàn giống Google Search. Google tìm kiếm thông tin trên nhiều trang web, còn AutoFilter chỉ lọc dữ liệu trong bảng tính hiện tại. Các dòng không phù hợp sẽ tạm thời bị ẩn, không bị xóa khỏi dữ liệu gốc.

### Câu 2: Vì sao cần chuẩn hóa dữ liệu trước khi tra cứu?

Dữ liệu không đồng nhất, chẳng hạn mã `VIR001` có khoảng trắng ở đầu hoặc cuối, có thể khiến công thức tra cứu không tìm thấy kết quả chính xác.

Vì vậy, cần sử dụng hàm `TRIM` để loại bỏ khoảng trắng thừa trước khi tra cứu. Việc chuẩn hóa giúp dữ liệu nhất quán, hạn chế lỗi và tăng độ chính xác của kết quả.

### Câu 3: Vai trò của VLOOKUP và IFERROR là gì?

- `VLOOKUP`: Tìm mã mã độc trong cơ sở dữ liệu và trả về thông tin tương ứng.
- `TRIM`: Loại bỏ khoảng trắng thừa trong mã cần tra cứu.
- `FALSE`: Yêu cầu tìm kiếm chính xác mã cần tìm.
- `IFERROR`: Hiển thị thông báo dễ hiểu khi công thức gặp lỗi, thay vì để lỗi như `#N/A`.

Ví dụ công thức tra cứu tên mã độc:

```excel
=IFERROR(VLOOKUP(TRIM(D2),CoSoDuLieuVirus!$A$2:$D$9,2,FALSE),"Chưa xác định")
```

Công thức trên lấy mã tại ô D2, loại bỏ khoảng trắng thừa, tìm mã trong bảng cơ sở dữ liệu virus và trả về tên mã độc. Nếu không tìm thấy mã, Excel hiển thị “Chưa xác định”.

### Câu 4: Vì sao cần xây dựng chiến lược phòng thủ an ninh mạng theo nhiều tầng?

Phòng thủ nhiều tầng giúp giảm nguy cơ một điểm yếu khiến toàn bộ hệ thống bị xâm nhập. Nếu một lớp phòng thủ không ngăn chặn được sự cố, các lớp khác vẫn có thể hỗ trợ phát hiện, hạn chế thiệt hại và xử lý sự cố.

Mô hình NetGuard gồm:

1. **Tầng 1 – Phòng thủ mạng:** Tường lửa, IDS/IPS và giám sát lưu lượng.
2. **Tầng 2 – Thiết bị và hệ thống:** Chống mã độc, cập nhật bản vá và kiểm soát truy cập.
3. **Tầng 3 – Giám sát và ứng phó:** Phân tích nhật ký, điều tra và khôi phục hệ thống.

## 3. Phần 2: Triển khai

### File 1: BaoCaoSuCo.docx

**Mục đích:** Xây dựng báo cáo sự cố an ninh mạng có cấu trúc rõ ràng, dễ đọc và dễ cập nhật.

**Nội dung chính:**
- Tổng quan và thông tin sự cố.
- Phân tích nhật ký hệ thống.
- Phương án xử lý sự cố.
- Chiến lược phòng thủ an ninh mạng.
- Khuyến nghị bảo mật và kết luận.

**Kỹ thuật thực hiện:**
- Sử dụng Heading 1 và Heading 2 để phân cấp tiêu đề.
- Bật Navigation Pane để di chuyển nhanh giữa các phần.
- Tạo mục lục tự động bằng References → Table of Contents.
- Cập nhật mục lục sau khi chỉnh sửa nội dung.

### File 2: DuLieuSuCO.xlsx

**Mục đích:** Quản lý dữ liệu mã độc và hỗ trợ tra cứu thông tin từ nhật ký hệ thống.

**Các trang tính:**
- `CoSoDuLieuVirus`: Lưu mã mã độc, tên mã độc, mức độ nguy hiểm và biện pháp xử lý.
- `NhatKyHeThong`: Lưu mã cảnh báo, thời gian, địa chỉ IP, mã mã độc, kết quả tra cứu và trạng thái xử lý.
- `HuongDan`: Giải thích công thức và cách sử dụng bảng dữ liệu.

**Kỹ thuật thực hiện:**
- Chuyển vùng dữ liệu thành bảng bằng Format as Table.
- Sử dụng AutoFilter để tìm kiếm và lọc cảnh báo.
- Dùng `TRIM` kết hợp `VLOOKUP` để tra cứu chính xác mã mã độc.
- Dùng `IFERROR` để xử lý trường hợp không tìm thấy mã.
- Cố định vùng tra cứu bằng `$A$2:$D$9` để công thức không bị lệch khi sao chép xuống các dòng dưới.

**Kiểm thử:** Các mã có khoảng trắng thừa như ` VIR001 ` vẫn có thể được tra cứu sau khi chuẩn hóa. Mã `VIR999` không có trong cơ sở dữ liệu sẽ trả về “Chưa xác định”.

### File 3: ChienLuocPhongThu.pptx

**Mục đích:** Trình bày trực quan mô hình phòng thủ an ninh mạng 3 tầng và quy trình xử lý sự cố.

**Nội dung chính:**
- Slide 1: Giới thiệu chiến lược phòng thủ an ninh mạng NetGuard.
- Slide 2: Sơ đồ phân cấp phòng thủ 3 tầng.


**Yêu cầu thiết kế:**
- Sử dụng font Montserrat hoặc Inter, không sử dụng Times New Roman.
- Dùng nền xanh đậm, chữ sáng và màu sắc khác nhau để phân biệt các tầng.
- Sử dụng hình khối, mũi tên và biểu tượng phù hợp.
- Trình bày ngắn gọn, dễ đọc, ưu tiên sơ đồ thay vì đoạn văn dài.

## 4. Kết quả đạt được

Sau khi hoàn thành bài thực hành, bộ tài liệu NetGuard có thể hỗ trợ:

- **Word:** Chuẩn hóa báo cáo sự cố và tạo mục lục tự động.
- **Excel:** Quản lý nhật ký, tra cứu thông tin mã độc, lọc cảnh báo và xử lý lỗi dữ liệu.
- **PowerPoint:** Trình bày mô hình phòng thủ nhiều tầng và quy trình ứng phó sự cố.
- **README.md:** Tổng hợp mục tiêu, phân tích, kỹ thuật triển khai và kết quả của toàn bộ bài thực hành.

## 5. Kết luận

Bài thực hành giúp vận dụng Microsoft Word, Excel và PowerPoint vào một tình huống an ninh mạng cụ thể. Việc kết hợp tài liệu có cấu trúc, dữ liệu được chuẩn hóa và sơ đồ trực quan giúp thông tin sự cố dễ theo dõi hơn, hỗ trợ quá trình tra cứu và nâng cao hiệu quả phối hợp xử lý sự cố tại NetGuard.
