PHẦN 1. PHÂN TÍCH: BẢNG EXCEL "XẤU" (VI PHẠM QUY TẮC VÀNG) VS. BẢNG "ĐÚNG" (DATABASE MINDSET)
Để quản lý danh sách nhà đất và tự động tính toán hoa hồng môi giới chính xác, bảng dữ liệu Excel cần được chuẩn hóa theo tư duy Database (Database Mindset).

Tiêu chí	Bảng Excel "Xấu" (Vi phạm Quy tắc Vàng)	Bảng Excel "Đúng" (Database Mindset chuẩn)
Gộp ô (Merge Cells)	Gộp các ô tiêu đề hoặc ô dữ liệu trùng lặp (ví dụ: gộp các ô Biệt thự, Căn hộ). Dẫn đến lỗi không thể sắp xếp (Sort), không lọc được dữ liệu (Filter) và làm hỏng công thức VLOOKUP/INDEX-MATCH.	
Không Merge Cell trong toàn bộ vùng dữ liệu. Mỗi ô chứa đúng 1 giá trị độc lập. 

Dòng Header & Hàng trống	Có nhiều dòng tiêu đề chồng chéo, chèn các hàng trống ngắt quãng hoặc chèn dòng "Tổng cộng" ở giữa bảng.	
Duy nhất 1 Header Row ở dòng đầu tiên của bảng dữ liệu. Dữ liệu liên tục, không chứa dòng/cột trống. 

Định dạng dữ liệu (Data Types)	Trộn lẫn số, văn bản và ký tự tiền tệ thủ công trong cùng một ô (ví dụ: gõ "15 tỷ VND" làm số không thể tính toán).	
Phân định rõ kiểu dữ liệu: Cột tiền tệ để định dạng Currency/Number (15,000,000,000), Cột mã định dạng Text. 
Cấu trúc Record & Field	Dữ liệu trình bày theo dạng ma trận ngẫu nhiên, tên thuộc tính ghi rải rác.	
Cấu trúc Field (Cột) cố định thuộc tính, Record (Dòng) đại diện cho đúng 1 bất động sản duy nhất. 

PHẦN 2. TRIỂN KHAI: TOÀN BỘ HỌC LIỆU DỰ ÁN SMART-HOME
Tôi sẽ khởi tạo file báo cáo triển khai đầy đủ cho bộ 3 học liệu Word (SRS Hồ sơ giới thiệu), Excel (Bảng tính quản lý & VLOOKUP) và PowerPoint (Pitch Deckเรียก vốn).

smart_home_project_package.md [generated with file_gen tool]

PHẦN 3. KIỂM THỬ: SỬ DỤNG DẤU $ (DOLLAR SIGN) ĐỂ KHẮC PHỤC LỖI TRÔI VÙNG DỮ LIỆU KHI KÉO CÔNG THỨC
1. Bản chất sự cố trôi vùng dữ liệu
Khi viết công thức VLOOKUP với tham chiếu tương đối (ví dụ: =VLOOKUP(E2, H4:J7, 3, FALSE)), khi ta sao chép hoặc kéo công thức từ dòng 2 xuống dòng 3, 4, 5..., Excel sẽ tự động tăng số dòng của vùng bảng tra cứu thành H5:J8, H6:J9, H7:J10...
Dẫn đến việc các dòng phía dưới không còn tìm thấy thông tin trong bảng tra cứu môi giới và trả về lỗi #N/A hoặc kết quả bị sai lệch.

2. Nguyên lý và quy trình khắc phục bằng dấu $ (Tham chiếu tuyệt đối)
Nguyên lý:
Ký tự $ trong Excel có tác dụng cố định (khóa) cột hoặc dòng phía sau nó:

$H$4: Khóa tuyệt đối cả cột H và dòng 4.

$J$7: Khóa tuyệt đối cả cột J và dòng 7.

Quy trình áp dụng chuẩn cho dự án Smart-Home:   

Bước 1: Tại ô đầu tiên tính Tỷ lệ hoa hồng (ô G2), viết công thức VLOOKUP.   

Bước 2: Bôi đen vùng bảng tra cứu H4:J7 trong công thức và nhấn phím F4 trên bàn phím.

Bước 3: Công thức tự động chuyển thành tham chiếu tuyệt đối chuẩn:

=VLOOKUP(E2, $H$4:$J$7, 3, FALSE)

(Trong đó: E2 là Mã MG cần tra cứu, $H$4:$J$7 là Bảng tra cứu cố định, 3 là Cột Tỷ lệ hoa hồng, FALSE là Khớp chính xác).   
JPG

Bước 4: Bấm Enter và thực hiện kéo sao chép công thức xuống toàn bộ danh sách. Bảng tra cứu $H$4:$J$7 sẽ giữ nguyên vị trí cố định 100%, tính toán chính xác số tiền hoa hồng cho mọi bất động sản.   
JPG
