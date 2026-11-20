# TÀI LIỆU QUY ĐỊNH RULE RÁP MÃ CÔNG NỢ TRÊN EXCEL

## 1. Thông tin chung

| Nội dung | Thông tin |
|---|---|
| Tên quy trình | Ráp mã công nợ |
| Bộ phận phụ trách | ___ |
| Người thực hiện | ___ |
| Người kiểm tra | ___ |
| Ngày áp dụng | ___ |
| Phiên bản | 1.1 |
| File Excel sử dụng | File nguồn: `test_dbf88f.xlsx` (sheet `BR` hóa đơn bán ra, `MV` hóa đơn mua vào, `VND112` sao kê). File kết quả: `Matching_cong_no_v4.xlsx`. Script tạo file: `matching_cong_no.py` |

## 2. Mục đích

- Chuẩn hóa quy tắc ráp mã công nợ giữa các nguồn dữ liệu.
- Đảm bảo các giao dịch được ghép đúng mã công nợ.
- Hạn chế sai sót khi xử lý dữ liệu thủ công.
- Tạo cơ chế kiểm tra và xử lý các trường hợp không thể tự động matching.

## 3. Nguyên tắc ráp mã công nợ

- Ưu tiên matching theo các trường thông tin có độ chính xác cao.
- Không tự động ghép nếu dữ liệu không đáp ứng điều kiện của Rule.
- Không ghi đè mã công nợ đã được xác nhận nếu chưa có phê duyệt.
- Các trường hợp không xác định được mã phải chuyển sang danh sách kiểm tra thủ công.
- Lưu lại kết quả xử lý để phục vụ đối soát và truy vết.

**Quy ước dùng xuyên suốt tài liệu**

- **Cột Mã hàng (cột `D` của sheet `VND112`) ban đầu luôn trống.** Quy trình này là bước điền cột đó: công thức điền Mã hàng = mã khách hàng (MST) tìm được, hoặc mã khách lẻ. Không có rule nào phụ thuộc vào việc người dùng nhập tay cột `D`.
- **Mã khách hàng** = MST của đối tượng, tương ứng cột `Mã KH` trong BR/MV. MST 13 số (mã chi nhánh) được gộp về 10 số của công ty. MST 10 số và MST 12 số (mã số thuế cá nhân theo CCCD) giữ nguyên.
- **Mã công nợ** = `131-<mã khách hàng>` khi thu tiền của khách hàng, `331-<mã khách hàng>` khi chi tiền cho nhà cung cấp. Đây là quy ước do người lập đề xuất, cần bộ phận kế toán xác nhận. Nếu công ty dùng quy ước khác, chỉ cần sửa công thức cột `AC` của sheet `VND112`.
- **Khách lẻ**: hóa đơn bán ra có tên người mua là "khách lẻ", "người tiêu dùng", "người mua"... hoặc là tên cá nhân. Hóa đơn khách lẻ không có MST thì dùng mã khách lẻ `KHACHLE`, mã công nợ là `131-KHACHLE`.
- **Chiều giao dịch**: Thu = Có TK 131 (cột `M`); Chi = Nợ TK 331 (cột `L`).
- Giao dịch thu được đối chiếu với **BR**, giao dịch chi được đối chiếu với **MV**.

## 4. Danh sách Business Rules

| Rule ID | Tên Rule | Điều kiện matching | Kết quả xử lý | Ưu tiên | Ghi chú |
|---|---|---|---|---|---|
| RULE_001 | Matching chính xác | MST + số hóa đơn (trích từ nội dung sao kê) tìm thấy đúng 1 lần trong sổ **và** số tiền giao dịch = tổng TT các hóa đơn (chênh lệch ≤ dung sai, mặc định 0) | Gán mã công nợ chính thức. Trạng thái MATCHED | 1 | Hiển thị "Đã ghép - Rule 1". Giao dịch trả nhiều hóa đơn thì cộng tổng TT các hóa đơn |
| RULE_002 | Matching theo mã khách hàng | Tìm tên công ty trong nội dung sao kê (kể cả viết tắt, không dấu), so với BR (thu) hoặc MV (chi). Điểm tin cậy ≥ 0,6 và hơn ứng viên thứ 2 ≥ 0,05 | MST dùng để ghép (cột `U`) và điền vào cột `D`. Không xác định được thì chuyển RULE_010 hoặc RULE_006 | 2 | Chạy trước RULE_001, 003, 004 vì là đầu vào bắt buộc |
| RULE_003 | Matching theo số hóa đơn | Trích số hóa đơn từ nội dung sao kê (sau chữ HD / HOA DON), tra khóa `MST\|Số HĐ` trong BR hoặc MV | Tìm thấy: chuyển so tiền (RULE_001 nếu khớp, EX_004 nếu lệch). Không tìm thấy: EX_001. Trùng: RULE_007 | 3 | Số hóa đơn là giá trị tính lúc chạy script |
| RULE_004 | Matching theo số tiền | Có MST nhưng nội dung không có số hóa đơn. Trong sổ có đúng 1 hóa đơn cùng MST và cùng tổng TT, đồng thời chỉ có 1 giao dịch có MST + số tiền đó | MATCHED, ghi rõ số hóa đơn ghép được (cột `AM`) | 4 | Hiển thị "Đã ghép - Rule 2". Độ tin cậy thấp hơn RULE_001, cần soát lại |
| RULE_005 | Matching theo ngày giao dịch | Dùng ngày thanh toán (cột `C`) so với ngày hóa đơn (BR cột `E`) trong RULE_010 (khách lẻ). Ngoài ra nhận diện ngày ghi trong nội dung có năm 2025 để loại trừ hóa đơn kỳ trước | Khách lẻ: chọn hóa đơn có ngày gần nhất. Hóa đơn kỳ trước: không ghép, chuyển kiểm tra | 5 | Với khách hàng có MST, ngày giao dịch không dùng làm điều kiện ghép |
| RULE_006 | Xử lý dữ liệu thiếu | Thiếu định khoản đối ứng; định khoản ngược chiều; không xác định được mã khách hàng; không có số hóa đơn và không khớp MST + số tiền; số hóa đơn không có trong sổ | NOT MATCHED hoặc PENDING REVIEW, đưa vào sheet `Kiểm tra thủ công` | - | Vẫn ghi mã công nợ đề xuất và tên công ty tìm được (nếu có) để hỗ trợ kiểm tra |
| RULE_007 | Xử lý dữ liệu trùng | Khóa `MST\|Số HĐ` xuất hiện nhiều lần trong sổ; nhiều hóa đơn hoặc nhiều giao dịch cùng MST + số tiền; nhiều hóa đơn khách lẻ hoặc nhiều giao dịch cùng số tiền | MULTIPLE MATCH, đưa vào `Kiểm tra thủ công` | - | Không tự chọn một trong các kết quả |
| RULE_008 | Phân loại giao dịch | Thu (Có 131), Chi (Nợ 331), không thuộc công nợ (lương, BHXH, thuế, lãi...), ngược chiều (Nợ 131 hoặc Có 331) | Chỉ giao dịch Thu/Chi mới đi tiếp. Các loại còn lại xử lý theo EX_003, EX_007 hoặc bỏ qua | - | Chạy đầu tiên |
| RULE_009 | Khóa mã đã xác nhận | Cột `AG` (Mã công nợ đã xác nhận, nhập tay) có giá trị | Mã chính thức (cột `AH`) lấy theo `AG`, không bị công thức ghi đè | - | Khác mã đề xuất thì có cảnh báo ở cột `AI` và phải ghi Nhật ký phê duyệt |
| RULE_010 | Matching khách lẻ | Tiền thu, không xác định được MST từ tên (tên người chuyển có thể là cá nhân). Trong BR có đúng 1 hóa đơn khách lẻ cùng tổng TT với số tiền thanh toán, ngày hóa đơn gần ngày thanh toán (cửa sổ ± 7 ngày, ưu tiên ± 2 ngày), và chỉ có 1 giao dịch như vậy | MATCHED, mã công nợ `131-KHACHLE` (hoặc `131-<MST>` nếu hóa đơn có MST). Ghi rõ số hóa đơn và số ngày lệch | 1* | Không so tên người chuyển với tên trên hóa đơn. Độ tin cậy thấp hơn RULE_001, cần soát lại. *Chạy khi RULE_002 không xác định được MST |

**Thứ tự thực thi thực tế trên Excel** (cột `AD`, dừng ở điều kiện đầu tiên thỏa):

| Bước | Điều kiện | Kết quả (cột AD) | Rule |
|---|---|---|---|
| 1 | Cột L hoặc M trống | Cần kiểm tra - chưa phân loại TK | RULE_008, 006 |
| 2 | Nợ 131 hoặc Có 331 | Cần kiểm tra - bút toán ngược chiều | RULE_008, 006 |
| 3 | Không phải Thu 131 hoặc Chi 331 | Không thuộc công nợ | RULE_008 |
| 4 | Chưa có MST (RULE_002 không xác định được), tiền thu, có đúng 1 hóa đơn khách lẻ cùng số tiền trong cửa sổ ngày và chỉ 1 giao dịch như vậy | **Đã ghép - Rule khách lẻ** | RULE_010, 005 |
| 5 | Chưa có MST, tiền thu, có nhiều hóa đơn khách lẻ hoặc nhiều giao dịch cùng số tiền | Cần kiểm tra - nhiều HĐ/GD khách lẻ cùng số tiền | RULE_007 |
| 6 | Chưa có MST và không có hóa đơn khách lẻ phù hợp | Cần kiểm tra - thiếu MST | RULE_006 |
| 7 | Có số HĐ trích được và nội dung có ngày 2025 | Cần kiểm tra - HĐ kỳ trước (2025) | RULE_005 |
| 8 | Có số HĐ nhưng MST + số HĐ không có trong sổ | Cần kiểm tra - không tìm thấy HĐ | RULE_003, 006 |
| 9 | Khóa MST + số HĐ trùng trong sổ | Cần kiểm tra - trùng khóa HĐ | RULE_007 |
| 10 | Tiền khớp tổng TT hóa đơn | **Đã ghép - Rule 1** | RULE_001 |
| 11 | Tiền không khớp tổng TT hóa đơn | Cần kiểm tra - lệch tiền | RULE_003 |
| 12 | Không có số HĐ, đúng 1 hóa đơn và 1 giao dịch cùng MST + tiền | **Đã ghép - Rule 2** | RULE_004 |
| 13 | Không có số HĐ, nhiều hóa đơn hoặc giao dịch cùng MST + tiền | Cần kiểm tra - nhiều HĐ/GD cùng MST+tiền | RULE_007 |
| 14 | Không có số HĐ, không có hóa đơn nào cùng MST + tiền | Cần kiểm tra - không khớp MST+tiền (không có số HĐ) | RULE_006 |

**Tham số** (sheet `Rule`, ô nền vàng, có thể chỉnh):

| Ô | Tham số | Mặc định | Ý nghĩa |
|---|---|---|---|
| `Rule!C5` | Dung sai chênh lệch tiền (VND) | 0 | 0 nghĩa là phải khớp tuyệt đối. Lệch dù 1 đồng vẫn chuyển kiểm tra thủ công |
| `Rule!C6` | Ngưỡng tin cậy tên công ty | 0,6 | MST lấy từ tên chỉ được dùng để ghép khi điểm đạt ngưỡng này |
| `Rule!C7` | Biên độ so với ứng viên thứ 2 | 0,05 | Tránh chọn nhầm khi hai công ty có tên gần giống nhau |
| `Rule!C8` | Cửa sổ rộng khách lẻ (± ngày) | 7 | Ngày thanh toán lệch tối đa bao nhiêu ngày so với ngày hóa đơn khách lẻ |
| `Rule!C9` | Cửa sổ hẹp khách lẻ (± ngày) | 2 | Nếu có đúng 1 hóa đơn trong cửa sổ hẹp thì chọn hóa đơn đó (gần ngày thanh toán nhất) |
| `Rule!C10` | Mã khách lẻ | KHACHLE | Mã khách hàng khi hóa đơn khách lẻ không có MST |

## 5. Chi tiết từng Rule

### Rule ID: RULE_001

- Tên Rule: Matching chính xác
- Mô tả: Ghép giao dịch sao kê với hóa đơn khi cả ba thông tin MST, số hóa đơn và số tiền cùng khớp. Đây là rule có độ chính xác cao nhất nên được ưu tiên số 1.
- Điều kiện đầu vào:
  - Cột dữ liệu 1: `VND112!U` (MST dùng để ghép, kết quả của RULE_002)
  - Cột dữ liệu 2: sheet `Chi tiết HĐ` cột B (số hóa đơn trích từ nội dung, kết quả của RULE_003)
  - Cột dữ liệu 3: `VND112!Y` (số tiền giao dịch = thu `H` + chi `I`) và `BR/MV!U` (Tổng TT)
- Điều kiện matching: khóa `MST|Số HĐ` xuất hiện đúng 1 lần trong sổ tương ứng, không có hóa đơn trong nhóm bị thiếu hoặc trùng, và `ABS(Số tiền GD - Tổng TT các HĐ) <= Rule!C5`.
- Logic xử lý trên Excel:
  ```
  X4 (Tổng TT HĐ trong sổ) = SUMIFS('Chi tiết HĐ'!$H$2:$H$600, 'Chi tiết HĐ'!$A$2:$A$600, 4)
  Y4 (Số tiền giao dịch)   = N($H4) + N($I4)
  Z4 (Chênh lệch)          = IF($V4=0, "", $Y4 - $X4)
  AD4 (Kết quả)            = ... IF(ABS($Z4) <= Rule!$C$5, "Đã ghép - Rule 1", "Cần kiểm tra - lệch tiền")
  AC4 (Mã đề xuất)         = IF($U4<>"", IF($AV4="Thu","131","331") & "-" & $U4, $BC4)
  AH4 (Mã chính thức)      = IF($F4="","", IF($AG4<>"", $AG4, IF(LEFT($AD4,7)="Đã ghép", $AC4, "")))
  ```
  Trong đó `4` là số dòng của giao dịch trong `VND112`. Công thức đầy đủ của cột `AD` nằm trong file kết quả.
- Kết quả mong đợi: cột `AD` = "Đã ghép - Rule 1", cột `AH` có mã công nợ chính thức.
- Trường hợp ngoại lệ: giao dịch trả nhiều hóa đơn (ví dụ HĐ 1, 7, 12, 16) thì cộng tổng TT các hóa đơn rồi so với số tiền. Trả từng phần hoặc trả thừa thì không ghép (EX_004). Hóa đơn bị thay thế trong BR/MV có Tổng TT = 0 nên cũng báo lệch tiền.

### Rule ID: RULE_002

- Tên Rule: Matching theo mã khách hàng (tìm MST từ tên công ty)
- Mô tả: Cột Mã hàng ban đầu trống nên MST phải được suy ra từ nội dung sao kê. Rule tìm tên công ty ghi trong nội dung, kể cả khi viết tắt hoặc không dấu, rồi lấy MST của công ty tương ứng trong BR (tiền thu) hoặc MV (tiền chi).
- Điều kiện đầu vào:
  - Cột dữ liệu 1: `VND112!F` (nội dung sao kê)
  - Cột dữ liệu 2: chiều giao dịch (thu hoặc chi) suy ra từ `VND112!H` / `I`
  - Cột dữ liệu 3: `BR!F`, `BR!G` (MST, tên khách hàng) hoặc `MV!F`, `MV!G` (MST, tên nhà cung cấp)
- Điều kiện matching: điểm tin cậy ≥ `Rule!C6` và hơn ứng viên thứ 2 ≥ `Rule!C7`. Nếu không thỏa thì để trống MST (chỉ hiển thị ứng viên để tham khảo).
- Logic xử lý trên Excel:
  ```
  U4 (MST dùng để ghép) = IF($F4="","", IF($AQ4="","",
                            IF(AND($AQ4>=Rule!$C$6, $AQ4-N($AR4)>=Rule!$C$7), $AP4, "")))
  D4 (Mã hàng)          = IF($F4="","", IF($U4<>"", $U4, IF($BC4<>"", MID($BC4,5,50), "")))
  AU4 (Cách xác định MST) = "Tên công ty (tin cậy xx%)" | "Khách lẻ (số tiền + ngày gần HĐ)" | "Chưa có - gợi ý theo số tiền" | "Chưa xác định"
  BR/MV!AB = IF(LEN(TRIM(F2&""))=13, LEFT(TRIM(F2&""),10), TRIM(F2&""))   (MST chuẩn hóa)
  ```
  Cột `AO:AS` (tên công ty tìm được, MST theo tên, điểm tin cậy, điểm và tên ứng viên thứ 2) do script tính sẵn bằng thuật toán so khớp tên bên dưới, ghi dưới dạng giá trị (chữ xanh).

  **Thuật toán so khớp tên** (hàm `Directory.score` trong script):
  1. Chuẩn hóa nội dung: bỏ dấu, chữ hoa, bỏ ký tự đặc biệt; bỏ mã giao dịch (chuỗi chữ lẫn số dài, số từ 6 chữ số trở lên) và bỏ các số rời.
  2. Bỏ tên công ty mình (`THUONG MAI THUAN THANH PHAT`, `THUAN THANH PHAT`, `THANH PHAT`). Với giao dịch thu, bỏ thêm cụm `CHO CTY THUAN PHAT` (cách gọi tắt công ty mình).
  3. Bỏ từ pháp lý (CONG TY, CTY, TNHH, CP, CO PHAN, MTV...) và cụm từ nhiễu (THANH TOAN, TRA TIEN, CHUYEN KHOAN, HOA DON, HD, SO, NGAY, CHO...).
  4. Chọn danh bạ: tiền thu so với tên khách hàng trong BR, tiền chi so với tên nhà cung cấp trong MV.
  5. Chấm điểm cho từng công ty: tách tên công ty thành từ "riêng" (có tính phân biệt) và từ "chung" (sản xuất, thương mại, hóa chất...). Điểm = độ phủ của các từ riêng, có trọng số theo độ hiếm của từ trong danh bạ, nhân hệ số phụ theo độ phủ của từ chung (kể cả viết tắt như TM, SX, KT, XNK) và độ chính xác của nội dung so với tên.
  6. Các tín hiệu cộng thêm: tên riêng ở cuối tên (ví dụ "Bien xanh"), từ riêng hiếm có lỗi chính tả nhẹ (ví dụ SMALLFORTUNR), chữ viết tắt từ ba từ liên tiếp trở lên (ví dụ GTVT), tên viết dính nhau (TIMSEN), nội dung chứa trọn tên đã chuẩn hóa (điểm 1,0).
  7. Chọn công ty điểm cao nhất; bỏ qua ứng viên cùng MST khi tìm ứng viên thứ 2.
- Kết quả mong đợi: cột `U` có MST, cột `D` được điền, cột `AU` ghi rõ nguồn. Cột `AT` là gợi ý MST theo số tiền duy nhất (chỉ tham khảo, không dùng để tự ghép vì sẽ tự khớp vòng tròn).
- Trường hợp ngoại lệ: nội dung không nêu tên công ty (ví dụ chỉ có mã giao dịch), công ty chưa có trong BR/MV kỳ này (khách hàng cũ năm 2025), hoặc hai công ty tên gần giống nhau (biên độ không đạt) thì không tự điền MST. Các dòng này sang RULE_010 (nếu là tiền thu) hoặc RULE_006.

### Rule ID: RULE_003

- Tên Rule: Matching theo số hóa đơn
- Mô tả: Trích số hóa đơn từ nội dung sao kê và tra trong sổ hóa đơn.
- Điều kiện đầu vào:
  - Cột dữ liệu 1: `VND112!F` (nội dung sao kê)
  - Cột dữ liệu 2: `BR!D` / `MV!D` (số hóa đơn)
  - Cột dữ liệu 3: `BR!AC` / `MV!AC` (khóa `MST|Số HĐ`, cột trợ giúp do script thêm)
- Điều kiện matching: tồn tại khóa `MST|Số HĐ` trong sổ tương ứng (BR nếu thu, MV nếu chi).
- Logic xử lý trên Excel:
  ```
  BR/MV!AC = AB2 & "|" & TRIM(D2&"")                    (khóa MST|Số HĐ)
  'Chi tiết HĐ'!E = MST & "|" & SốHĐ
  'Chi tiết HĐ'!F = COUNTIF(BR!$AC$2:$AC$1100, E2)      (số lần trong sổ; Chi thì dùng MV)
  'Chi tiết HĐ'!G = IF(F2>0,1,0)
  VND112!AA4 = COUNTIFS('Chi tiết HĐ'!A:A, 4, 'Chi tiết HĐ'!G:G, 0)   (số HĐ không có trong sổ)
  ```
  **Quy tắc trích số hóa đơn** (hàm `extract` trong script, ghi vào sheet `Chi tiết HĐ` cột A-B và `VND112!AN`):
  - Lấy dãy số 1 đến 5 chữ số sau từ khóa `HD` hoặc `HOA DON` (có hoặc không có `SO`), ví dụ `HD 01.07.12.16`, `HD016`, `HOA DON SO 30`.
  - Các số trong dãy được tách bằng dấu chấm, phẩy, chấm phẩy, gạch ngang, dấu `&`, chữ `VA` hoặc khoảng trắng. Bỏ số 0 đứng đầu.
  - Bỏ qua ngày tháng đi kèm (`NGAY dd/mm/yyyy`, `(dd/mm/yy)`) và các mã giao dịch dài (`HD-280126-...`).
  - Ngày có năm 2025 trong nội dung thì đánh dấu hóa đơn kỳ trước (cột `W`).
- Kết quả mong đợi: giao dịch có số hóa đơn được tra khóa; kết quả chuyển tiếp cho RULE_001 (khớp tiền) hoặc EX_004 (lệch tiền).
- Trường hợp ngoại lệ: số hóa đơn không có trong sổ (EX_001); khóa trùng (RULE_007); nội dung có nhiều số bị cắt dở hoặc dạng `HD 32,9` có thể tách thừa số (đã kiểm thử, xem Phụ lục D). Số hóa đơn trích được là **giá trị cố định**, khi thêm dòng sao kê mới cần chạy lại script.

### Rule ID: RULE_004

- Tên Rule: Matching theo số tiền
- Mô tả: Dành cho giao dịch đã có MST nhưng nội dung không nêu số hóa đơn. Ghép theo cặp MST và số tiền khi cặp này xác định duy nhất một hóa đơn.
- Điều kiện đầu vào:
  - Cột dữ liệu 1: `VND112!U` (MST dùng để ghép)
  - Cột dữ liệu 2: `VND112!Y` (số tiền giao dịch)
  - Cột dữ liệu 3: `BR!AD` / `MV!AD` (khóa `MST|Tổng TT`)
- Điều kiện matching: `VND112!V = 0` (không có số hóa đơn) và `AK = 1` (đúng 1 hóa đơn cùng MST + tiền) và `AL = 1` (chỉ 1 giao dịch có MST + tiền đó).
- Logic xử lý trên Excel:
  ```
  AJ4 (khóa MST|Số tiền)   = $U4 & "|" & $Y4
  AK4 (số HĐ trong sổ)     = COUNTIF(BR!$AD$2:$AD$1100, $AJ4)     (Chi thì dùng MV)
  AL4 (số GD cùng MST+tiền)= COUNTIFS($AJ$4:$AJ$310, $AJ4, $T$4:$T$310, $T4)
  AM4 (số HĐ ghép được)    = INDEX(BR!$D$2:$D$1100, MATCH($AJ4, BR!$AD$2:$AD$1100, 0))
  AD4                      = IF($AK4=0, "Cần kiểm tra - không khớp MST+tiền (không có số HĐ)",
                               IF(AND($AK4=1,$AL4=1), "Đã ghép - Rule 2",
                                  "Cần kiểm tra - nhiều HĐ/GD cùng MST+tiền"))
  ```
- Kết quả mong đợi: "Đã ghép - Rule 2", mã công nợ chính thức được gán, cột `AM` và cột lý do `AE` nêu số hóa đơn ghép được để người dùng soát lại.
- Trường hợp ngoại lệ: nhiều hóa đơn hoặc nhiều giao dịch cùng MST + số tiền (MULTIPLE MATCH); không có hóa đơn nào cùng MST + số tiền (NOT MATCHED).

### Rule ID: RULE_005

- Tên Rule: Matching theo ngày giao dịch
- Mô tả: Ngày giao dịch được dùng ở hai nơi. (1) RULE_010: so ngày thanh toán với ngày hóa đơn khách lẻ để chọn hóa đơn gần nhất. (2) Nhận diện ngày ghi trong nội dung có năm 2025 để loại trừ hóa đơn kỳ trước, tránh ghép nhầm với hóa đơn cùng số của kỳ này.
- Điều kiện đầu vào:
  - Cột dữ liệu 1: `VND112!C` (ngày thanh toán) và `BR!E` (ngày hóa đơn), dùng cho khách lẻ
  - Cột dữ liệu 2: `VND112!F` (nội dung sao kê), dùng nhận diện năm 2025
  - Cột dữ liệu 3: `VND112!V` (số hóa đơn trích được)
- Điều kiện matching: (1) ngày hóa đơn nằm trong ± `Rule!C8` ngày quanh ngày thanh toán (xem RULE_010). (2) nội dung chứa `/2025`, `-2025` hoặc `.2025` (cột `W` = 1) và có số hóa đơn trích được.
- Logic xử lý trên Excel:
  ```
  W4  = IF(OR(ISNUMBER(SEARCH("/2025",$F4&"")), ISNUMBER(SEARCH("-2025",$F4&"")), ISNUMBER(SEARCH(".2025",$F4&""))), 1, 0)
  AD4 = ... IF($W4=1, "Cần kiểm tra - HĐ kỳ trước (2025)", ...)
  BB4 (Chênh lệch ngày, khách lẻ) = $C4 - $BA4
  ```
- Kết quả mong đợi: khách lẻ chọn được hóa đơn có ngày gần nhất; hóa đơn kỳ trước không ghép mà chuyển kiểm tra để đối chiếu số dư đầu kỳ (PENDING REVIEW).
- Trường hợp ngoại lệ: với khách hàng có MST, ngày giao dịch không tham gia điều kiện ghép. Nội dung chỉ ghi số hóa đơn mà không ghi ngày (ví dụ "HD 15") thì không nhận diện được kỳ trước, giao dịch sẽ rơi vào "không tìm thấy HĐ" (EX_001) khi hóa đơn nằm ngoài sổ kỳ này.

### Rule ID: RULE_006

- Tên Rule: Xử lý dữ liệu thiếu
- Mô tả: Mọi giao dịch không đủ thông tin để ghép tự động đều chuyển danh sách kiểm tra thủ công, kèm lý do và mã đề xuất (nếu có).
- Điều kiện đầu vào:
  - Cột dữ liệu 1: `VND112!L`, `M` (định khoản)
  - Cột dữ liệu 2: `VND112!U` (MST)
  - Cột dữ liệu 3: `VND112!V` (số hóa đơn trích được)
- Điều kiện matching: một trong các trường hợp ở bước 1, 2, 6, 8, 14 của bảng thứ tự thực thi.
- Logic xử lý trên Excel:
  ```
  AF4 (STT kiểm tra) = IF(LEFT($AD4,4)="Cần ", MAX($AF$3:$AF3)+1, "")
  Sheet 'Kiểm tra thủ công': B = MATCH(STT, VND112!$AF$1:$AF$400, 0), các cột còn lại = INDEX(VND112!<cột>, B)
  ```
  Sheet `Kiểm tra thủ công` có sẵn 260 dòng công thức và tự điền theo kết quả. Mỗi dòng gồm: ngày, MST, nội dung, số hóa đơn trích được, chiều, số tiền, tổng TT hóa đơn, chênh lệch, kết quả, lý do, mã đề xuất, tên công ty tìm được, gợi ý MST theo số tiền.
- Kết quả mong đợi: dòng xuất hiện trong `Kiểm tra thủ công` với lý do cụ thể (cột `AE`).
- Trường hợp ngoại lệ: dòng chưa có định khoản đối ứng nhưng tìm được tên công ty vẫn có mã đề xuất (ví dụ `131-0314251773`) để kế toán xác nhận nhanh, nhưng chưa được coi là đã ghép.

### Rule ID: RULE_007

- Tên Rule: Xử lý dữ liệu trùng
- Mô tả: Khi một giao dịch có thể ứng với nhiều hóa đơn hoặc ngược lại, không tự chọn mà chuyển kiểm tra.
- Điều kiện đầu vào:
  - Cột dữ liệu 1: `VND112!AB` (số hóa đơn có khóa trùng trong sổ)
  - Cột dữ liệu 2: `VND112!AK`, `AL` (số hóa đơn, số giao dịch cùng MST + tiền)
  - Cột dữ liệu 3: `VND112!AX`, `AY` (số hóa đơn khách lẻ trong cửa sổ hẹp, số giao dịch thu cùng số tiền)
- Điều kiện matching: `AB > 0` hoặc `AK > 1` hoặc `AL > 1` hoặc (khách lẻ) `AX > 1`, `AW > 1` khi không có hóa đơn trong cửa sổ hẹp, hoặc `AY > 1`.
- Logic xử lý trên Excel:
  ```
  AB4 = COUNTIFS('Chi tiết HĐ'!$A$2:$A$600, 4, 'Chi tiết HĐ'!$F$2:$F$600, ">1")
  AD4 = ... IF($AB4>0, "Cần kiểm tra - trùng khóa HĐ", ...)    (có số hóa đơn)
        ... "Cần kiểm tra - nhiều HĐ/GD cùng MST+tiền"          (không có số hóa đơn)
        ... "Cần kiểm tra - nhiều HĐ/GD khách lẻ cùng số tiền"  (khách lẻ)
  ```
- Kết quả mong đợi: MULTIPLE MATCH, chuyển `Kiểm tra thủ công`.
- Trường hợp ngoại lệ: một hóa đơn được thanh toán bằng nhiều giao dịch (trả từng phần) hoặc trả trùng được phát hiện gián tiếp qua "lệch tiền" (EX_004). Hiện chưa có cột riêng đối chiếu tổng đã thanh toán trên mỗi hóa đơn.

### Rule ID: RULE_008

- Tên Rule: Phân loại giao dịch
- Mô tả: Xác định giao dịch có thuộc công nợ phải thu (131) hoặc phải trả (331) hay không, và đối chiếu với sổ nào.
- Điều kiện đầu vào:
  - Cột dữ liệu 1: `VND112!L` (Nợ)
  - Cột dữ liệu 2: `VND112!M` (Có)
  - Cột dữ liệu 3: `VND112!H`, `I` (thu, chi)
- Điều kiện matching: Thu khi `M = 131`; Chi khi `L = 331`.
- Logic xử lý trên Excel:
  ```
  T4  (Chiều)             = IF(($M4&"")="131","Thu", IF(($L4&"")="331","Chi",""))
  AV4 (Chiều suy từ tiền) = IF($T4<>"",$T4, IF(N($H4)>0,"Thu", IF(N($I4)>0,"Chi","")))
  ```
  `AV` dùng để đề xuất mã công nợ cho dòng chưa có định khoản. Giao dịch lương, BHXH, thuế TNCN, lãi, chuyển nhầm... không được tìm tên công ty.
- Kết quả mong đợi: Thu (đối chiếu BR), Chi (đối chiếu MV), "Không thuộc công nợ", hoặc chuyển kiểm tra khi thiếu định khoản hoặc ngược chiều.
- Trường hợp ngoại lệ: hoàn tiền, điều chỉnh (Nợ 131, Có 331) chuyển kiểm tra thủ công (EX_007).

### Rule ID: RULE_009

- Tên Rule: Khóa mã đã xác nhận
- Mô tả: Bảo đảm nguyên tắc không ghi đè mã công nợ đã được xác nhận nếu chưa có phê duyệt. Đây cũng là nơi kế toán nhập mã cho các dòng mà tự động không xác định được.
- Điều kiện đầu vào:
  - Cột dữ liệu 1: `VND112!AG` (Mã công nợ đã xác nhận, ô nền vàng, nhập tay)
  - Cột dữ liệu 2: `VND112!AC` (mã đề xuất)
  - Cột dữ liệu 3: sheet `Nhật ký`
- Điều kiện matching: `AG` có giá trị.
- Logic xử lý trên Excel:
  ```
  AH4 (Mã chính thức) = IF($F4="","", IF($AG4<>"", $AG4, IF(LEFT($AD4,7)="Đã ghép", $AC4, "")))
  AI4 (Cảnh báo)      = IF(AND($AG4<>"",$AC4<>"",$AG4<>$AC4), "Khác mã đề xuất - cần phê duyệt và ghi Nhật ký", "")
  ```
- Kết quả mong đợi: mã chính thức luôn ưu tiên mã đã xác nhận. Nếu khác mã đề xuất, cột `AI` cảnh báo và người duyệt phải ghi vào sheet `Nhật ký` (thời điểm, dòng, mã cũ, mã mới, người duyệt, căn cứ).
- Trường hợp ngoại lệ: chưa có cột trạng thái APPROVED/REJECTED tự động. Quy ước thao tác: nhập mã vào `AG` và ghi `Nhật ký` là APPROVED; nếu người duyệt bác mã đề xuất thì để trống `AG`, ghi `Nhật ký` với lý do REJECTED.

### Rule ID: RULE_010

- Tên Rule: Matching khách lẻ (số tiền + ngày thanh toán gần ngày hóa đơn)
- Mô tả: Khách lẻ chuyển tiền thanh toán thường dùng tên cá nhân, tên này có thể khác tên trên hóa đơn hoặc hóa đơn chỉ ghi "khách lẻ", "bán cho người tiêu dùng". Vì vậy rule không so tên người chuyển mà ghép theo tổng số tiền thanh toán và ngày thanh toán gần ngày hóa đơn.
- Điều kiện đầu vào:
  - Cột dữ liệu 1: `BR!G` (tên người mua) để nhận diện hóa đơn khách lẻ, kết quả ở `BR!AE` (1 = khách lẻ, chữ xanh, có thể sửa tay)
  - Cột dữ liệu 2: `VND112!Y` (số tiền thanh toán) và `BR!U` (Tổng TT)
  - Cột dữ liệu 3: `VND112!C` (ngày thanh toán) và `BR!E` (ngày hóa đơn)
- Điều kiện matching:
  1. Giao dịch là tiền thu (`AV` = "Thu") và RULE_002 chưa xác định được MST (`U` trống).
  2. Hóa đơn BR được nhận diện là khách lẻ khi tên người mua chứa "khách lẻ", "khách vãng lai", "người tiêu dùng", "người mua", "bán lẻ", "không lấy hóa đơn", "cá nhân"; hoặc là tên cá nhân (2 đến 5 từ, không chứa số, không có dấu hiệu tổ chức như công ty, TNHH, CP, doanh nghiệp tư nhân, hộ kinh doanh, trường, nhà hàng, khách sạn...).
  3. Hóa đơn có Tổng TT bằng đúng số tiền thanh toán và ngày hóa đơn trong ± `Rule!C8` ngày (mặc định 7) quanh ngày thanh toán.
  4. Chọn hóa đơn: nếu đúng 1 hóa đơn trong cửa sổ hẹp ± `Rule!C9` ngày (mặc định 2) thì chọn hóa đơn đó (gần ngày nhất); nếu không có hóa đơn nào trong cửa sổ hẹp và chỉ có đúng 1 hóa đơn trong cửa sổ rộng thì chọn hóa đơn đó; các trường hợp còn lại (nhiều hóa đơn) chuyển kiểm tra.
  5. Chỉ có đúng 1 giao dịch thu chưa xác định MST có cùng số tiền trong cửa sổ rộng. Nếu có nhiều giao dịch thì chuyển kiểm tra.
- Logic xử lý trên Excel:
  ```
  BR!AE   = Khách lẻ (1/0)                                  (script nhận diện, sửa tay được)
  BR!AF   = IF(AE2=1, U2&"|"&E2, "")                         (khóa Tổng TT|Ngày HĐ)
  AW4 (số HĐ lẻ, cửa sổ rộng) = IF(AND($F4<>"",$AV4="Thu",$U4=""),
        COUNTIFS(BR!$AE$2:$AE$1100,1, BR!$U$2:$U$1100,$Y4,
                 BR!$E$2:$E$1100,">="&($C4-Rule!$C$8), BR!$E$2:$E$1100,"<="&($C4+Rule!$C$8)), "")
  AX4 (số HĐ lẻ, cửa sổ hẹp)  = (như AW4 nhưng dùng Rule!$C$9)
  AY4 (số GD thu cùng số tiền)= COUNTIFS($AW$4:$AW$310,">=0", $Y$4:$Y$310,$Y4, $C$4:$C$310,">="&($C4-Rule!$C$8), $C$4:$C$310,"<="&($C4+Rule!$C$8))
  AZ4 (số HĐ ghép được)      = IF($AX4=1, SUMIFS(BR!$D..., cửa sổ hẹp), IF(AND($AX4=0,$AW4=1), SUMIFS(BR!$D..., cửa sổ rộng), ""))
  BA4 (ngày HĐ), BB4 (lệch ngày = $C4-$BA4)
  BC4 (mã công nợ khách lẻ)   = IF(OR($BA4="",N($AY4)<>1),"","131-" & IF(MST của HĐ<>"", MST của HĐ, Rule!$C$10))
  AD4 = IF($U4="", IF($AZ4<>"", IF($AY4=1,"Đã ghép - Rule khách lẻ","Cần kiểm tra - nhiều HĐ/GD khách lẻ cùng số tiền"),
                      IF(OR(N($AW4)>1,N($AX4)>1),"Cần kiểm tra - nhiều HĐ/GD khách lẻ cùng số tiền","Cần kiểm tra - thiếu MST")), ...)
  ```
- Kết quả mong đợi: "Đã ghép - Rule khách lẻ", cột `D` (Mã hàng) = MST của hóa đơn hoặc `KHACHLE`, mã công nợ chính thức `131-KHACHLE`. Cột `AE` nêu số hóa đơn, ngày hóa đơn và số ngày lệch để người dùng soát lại.
- Trường hợp ngoại lệ: có nhiều hóa đơn khách lẻ cùng số tiền và ngày gần nhau; nhiều khách chuyển cùng một số tiền trong vài ngày; số tiền chuyển khác tổng TT hóa đơn (trả từng phần, gộp nhiều hóa đơn); ngày thanh toán xa ngày hóa đơn hơn cửa sổ. Các trường hợp này chuyển kiểm tra thủ công. Vì không dựa vào tên người chuyển, rule có độ tin cậy thấp hơn RULE_001 nên kết quả cần được soát lại.

## 6. Quy định xử lý ngoại lệ

| Mã ngoại lệ | Trường hợp | Cách xử lý | Trạng thái |
|---|---|---|---|
| EX_001 | Không tìm thấy mã công nợ (không tìm thấy HĐ; không khớp MST + tiền; thiếu MST) | Kiểm tra thủ công tại sheet `Kiểm tra thủ công`. Xem lý do ở cột `AE`, tên công ty gợi ý và gợi ý MST theo số tiền | Pending |
| EX_002 | Trùng nhiều mã công nợ (trùng khóa HĐ; nhiều HĐ/GD cùng MST + tiền) | Chuyển kiểm tra, người xử lý chọn đúng hóa đơn rồi nhập mã vào `AG` | Pending |
| EX_003 | Thiếu thông tin đối chiếu (chưa phân loại TK; thiếu MST; nội dung không có số HĐ) | Bổ sung định khoản trong sổ, hoặc nhập mã công nợ xác nhận vào `AG` | Pending |
| EX_004 | Số tiền không khớp (trả từng phần, trả thừa, sai số tiền, hóa đơn bị thay thế) | Kiểm tra sai lệch. Cột `Z` cho biết chênh lệch, cột `AE` mô tả thiếu hay thừa | Pending |
| EX_005 | Mã đã được sử dụng (nhiều giao dịch cùng MST + số tiền; một hóa đơn bị thanh toán nhiều lần) | Xác minh lại giao dịch. Hiện phát hiện qua `AL > 1` hoặc qua "lệch tiền" | Pending |
| EX_006 | Hóa đơn kỳ trước (nội dung có ngày 2025) | Đối chiếu số dư đầu kỳ, không ghép với hóa đơn cùng số của kỳ này | Pending |
| EX_007 | Bút toán ngược chiều (Nợ 131 hoặc Có 331: hoàn tiền, điều chỉnh) | Kế toán xem lại định khoản | Pending |
| EX_008 | Mã đã xác nhận (`AG`) khác mã đề xuất (`AC`) | Phê duyệt và ghi sheet `Nhật ký` | Pending |
| EX_009 | Tên công ty không đạt ngưỡng tin cậy hoặc hai ứng viên sát điểm | MST không tự dùng. Kiểm tra ứng viên ở `AO`, `AS`, nếu xác nhận thì nhập mã công nợ vào `AG` | Pending |
| EX_010 | Khách lẻ: nhiều hóa đơn cùng số tiền và ngày gần nhau, nhiều giao dịch cùng số tiền, hoặc số tiền khác tổng TT hóa đơn | Đối chiếu với bên bán hàng (người lập hóa đơn) để xác định hóa đơn, nhập mã vào `AG` | Pending |

## 7. Kết quả sau khi xử lý

Quy ước cột: "Mã giao dịch" ghi số dòng trong sheet `VND112`. "Mã công nợ gốc" là giá trị cột `D` ban đầu, luôn trống. "Mã công nợ sau ráp" là mã chính thức (cột `AH`); mã trong ngoặc là mã đề xuất chưa được coi là đã ghép.

**Ví dụ kết quả thực tế** (từ `test_dbf88f.xlsx`):

| STT | Mã giao dịch | Mã khách hàng | Số hóa đơn | Số tiền | Mã công nợ gốc | Mã công nợ sau ráp | Rule áp dụng | Trạng thái |
|---|---|---|---|---|---|---|---|---|
| 1 | VND112!8 | 0304244618 | 1, 7, 12, 16 | 62.280.003 | (trống) | 331-0304244618 | RULE_002, 001 | MATCHED |
| 2 | VND112!9 | 3600630513 | 1018 | 14.399.999 | (trống) | 331-3600630513 | RULE_002, 001 | MATCHED |
| 3 | VND112!7 | 6001684988 | 2 (theo số tiền) | 1.673.784 | (trống) | 131-6001684988 | RULE_002, 004 | MATCHED |
| 4 | VND112!28 | 3600863821 | 14 (theo số tiền) | 8.128.080 | (trống) | 131-3600863821 | RULE_002, 004 | MATCHED |
| 5 | VND112!34 | 0304244618 | 66 | 11.374.844 (HĐ 11.374.884) | (trống) | (331-0304244618) | RULE_003, EX_004 | PENDING REVIEW |
| 6 | VND112!51 | 0304409475 | 9 | 80.000.000 (HĐ 155.347.200) | (trống) | (131-0304409475) | RULE_003, EX_004 | PENDING REVIEW |
| 7 | VND112!13 | 0304409475 | 15 | 118.098.000 | (trống) | (131-0304409475) | RULE_003, EX_001 | NOT MATCHED |
| 8 | VND112!6 | 0316242189 | 30 | 24.377.760 | (trống) | (131-0316242189) | RULE_005, EX_006 | PENDING REVIEW |
| 9 | VND112!56 | 0306168465 | (không có) | 2.160.000 | (trống) | (331-0306168465) | RULE_007, EX_002 | MULTIPLE MATCH |
| 10 | VND112!5 | (chưa xác định) | (không có) | 8.272.800 | (trống) | (trống) | RULE_006, EX_003 | NOT MATCHED |
| 11 | VND112!150 | 0314251773 | 117 | 3.402.000 | (trống) | (131-0314251773) | RULE_006, EX_003 | PENDING REVIEW |

STT 5 và 6 cho thấy rule hoạt động đúng nguyên tắc: giao dịch VND112!34 chỉ lệch 40 đồng vẫn không tự ghép mà chuyển kiểm tra, đồng thời vẫn đưa ra mã đề xuất. STT 10 (VND112!5) chưa xác định được MST từ tên, nhưng có gợi ý MST `0314842173` theo số tiền duy nhất (cột `AT`). Nếu chấp nhận, kế toán nhập mã công nợ vào cột `AG`.

**Ví dụ khách lẻ** (dữ liệu thử có chủ đích, vì BR trong file mẫu không có hóa đơn khách lẻ nào, xem Phụ lục D):

| STT | Mã giao dịch | Mã khách hàng | Số hóa đơn | Số tiền | Mã công nợ gốc | Mã công nợ sau ráp | Rule áp dụng | Trạng thái |
|---|---|---|---|---|---|---|---|---|
| 1 | VND112!258 | KHACHLE | 901 (lệch 1 ngày) | 1.500.000 | (trống) | 131-KHACHLE | RULE_010, 005 | MATCHED |
| 2 | VND112!259 | KHACHLE | 903 (lệch 0 ngày) | 2.000.000 | (trống) | 131-KHACHLE | RULE_010, 005 | MATCHED |
| 3 | VND112!260 | (chưa xác định) | 904 hoặc 905 | 3.000.000 | (trống) | (trống) | RULE_010, 007, EX_010 | MULTIPLE MATCH |
| 4 | VND112!261 | (chưa xác định) | (không có) | 4.000.000 | (trống) | (trống) | RULE_006, EX_003 | NOT MATCHED |

**Tổng hợp trên 249 dòng sao kê của file mẫu:**

| Trạng thái trên Excel (cột AD) | Số dòng | Trạng thái xử lý |
|---|---|---|
| Đã ghép - Rule 1 | 57 | MATCHED |
| Đã ghép - Rule 2 | 13 | MATCHED |
| Đã ghép - Rule khách lẻ | 0 | MATCHED |
| Cần kiểm tra - không tìm thấy HĐ | 13 | NOT MATCHED |
| Cần kiểm tra - thiếu MST | 22 | NOT MATCHED |
| Cần kiểm tra - không khớp MST+tiền (không có số HĐ) | 5 | NOT MATCHED |
| Cần kiểm tra - nhiều HĐ/GD cùng MST+tiền | 6 | MULTIPLE MATCH |
| Cần kiểm tra - nhiều HĐ/GD khách lẻ cùng số tiền | 0 | MULTIPLE MATCH |
| Cần kiểm tra - trùng khóa HĐ | 0 | MULTIPLE MATCH |
| Cần kiểm tra - chưa phân loại TK | 95 | PENDING REVIEW |
| Cần kiểm tra - lệch tiền | 7 | PENDING REVIEW |
| Cần kiểm tra - HĐ kỳ trước (2025) | 2 | PENDING REVIEW |
| Cần kiểm tra - bút toán ngược chiều | 1 | PENDING REVIEW |
| Không thuộc công nợ | 28 | Không áp dụng |
| **Tổng** | **249** | |

Tổng số dòng "Cần kiểm tra" là 151. Dòng "Tổng cộng" trong sheet `Rule` được đối chiếu với số dòng sao kê (chênh lệch phải bằng 0). Các dòng khách lẻ bằng 0 vì BR của file mẫu không có hóa đơn khách lẻ.

## 8. Trạng thái xử lý

- MATCHED: Đã ráp mã thành công. Trên Excel: "Đã ghép - Rule 1", "Đã ghép - Rule 2", "Đã ghép - Rule khách lẻ".
- NOT MATCHED: Không tìm thấy dữ liệu phù hợp. Trên Excel: "không tìm thấy HĐ", "thiếu MST", "không khớp MST+tiền (không có số HĐ)".
- MULTIPLE MATCH: Có nhiều kết quả phù hợp. Trên Excel: "trùng khóa HĐ", "nhiều HĐ/GD cùng MST+tiền", "nhiều HĐ/GD khách lẻ cùng số tiền".
- PENDING REVIEW: Chờ kiểm tra thủ công. Trên Excel: "chưa phân loại TK", "lệch tiền", "HĐ kỳ trước (2025)", "bút toán ngược chiều". Mọi dòng NOT MATCHED, MULTIPLE MATCH và PENDING REVIEW đều xuất hiện ở sheet `Kiểm tra thủ công`.
- APPROVED: Đã xác nhận kết quả. Người duyệt nhập mã vào cột `AG` và ghi sheet `Nhật ký`. Mã chính thức (`AH`) lấy theo `AG` và không bị ghi đè.
- REJECTED: Kết quả ráp mã không hợp lệ. Người duyệt bác mã đề xuất, để trống `AG` và ghi `Nhật ký` với lý do.

---

## Phụ lục A. Cách chạy script

```bash
pip install openpyxl
python matching_cong_no.py FILE_GOC.xlsx FILE_KET_QUA.xlsx
# tùy chọn:
#   --tolerance 0          dung sai lệch tiền (VND)
#   --name-threshold 0.6   ngưỡng tin cậy tên công ty
#   --name-margin 0.05     biên độ so với ứng viên thứ 2
#   --retail-window 7      cửa sổ rộng khách lẻ (± ngày)
#   --retail-tight 2       cửa sổ hẹp khách lẻ (± ngày)
#   --retail-code KHACHLE  mã khách lẻ khi hóa đơn không có MST
```

Sau khi chạy, mở file kết quả bằng Excel (hoặc LibreOffice) để công thức được tính lại. Script tự bỏ các liên kết ngoài bị hỏng của file gốc (tạo bản tạm) và không sửa file gốc. Giá trị cột `D` của file gốc không được sao chép, vì cột này được điền bằng công thức.

**Yêu cầu cấu trúc file gốc**

| Sheet | Cột bắt buộc |
|---|---|
| `VND112` (sao kê) | `C` ngày, `F` nội dung, `H` thu, `I` chi, `L` Nợ, `M` Có. Cột `D` (Mã hàng) để trống, script tự điền. Dữ liệu bắt đầu từ dòng 4 |
| `BR` (bán ra) | `D` số hóa đơn, `E` ngày hóa đơn, `F` Mã KH (MST), `G` tên khách hàng, `U` Tổng TT |
| `MV` (mua vào) | `D` số hóa đơn, `F` Mã KH (MST), `G` tên nhà cung cấp, `U` Tổng TT |

**Giới hạn hiện tại** (sửa hằng số đầu script nếu file lớn hơn): `VND112` tối đa 310 dòng, `BR`/`MV` tối đa 1100 dòng, tối đa khoảng 600 cặp (dòng sao kê, số hóa đơn).

## Phụ lục B. Cấu trúc file kết quả

| Sheet | Nội dung |
|---|---|
| `Rule` | Tham số (ô vàng), bảng rule theo thứ tự đánh giá, bảng tổng hợp kết quả có dòng tự kiểm tra, giả định |
| `VND112` | Dữ liệu sao kê gốc, cột `D` (Mã hàng) được điền bằng công thức, và các cột mới `T:BC`: chiều, MST dùng để ghép, số hóa đơn, chênh lệch, mã đề xuất, kết quả, lý do, mã xác nhận (nhập tay), mã chính thức, tên công ty tìm được, gợi ý theo số tiền, các cột khách lẻ (`AW:BC`) |
| `Kiểm tra thủ công` | Danh sách các dòng cần kiểm tra, tự động lấy từ `VND112` |
| `Chi tiết HĐ` | Mỗi cặp (dòng sao kê, số hóa đơn trích được) một dòng, kèm số lần xuất hiện trong sổ và tổng TT |
| `Danh bạ KH-NCC` | Danh sách MST và tên công ty từ BR, MV dùng để so khớp tên |
| `Nhật ký` | Bảng ghi phê duyệt (có 1 dòng ví dụ) và thống kê lần chạy |
| `BR`, `MV` | Bản sao giá trị từ file gốc, thêm cột trợ giúp `AB`, `AC`, `AD` (MST chuẩn hóa, khóa MST\|Số HĐ, khóa MST\|Tổng TT). `BR` thêm `AE` (khách lẻ 1/0) và `AF` (khóa Tổng TT\|Ngày HĐ) |

Chú giải màu: nền vàng hoặc chữ xanh = ô nhập tay hoặc giá trị do script tính sẵn; xanh lá = đã ghép; cam = cần kiểm tra; xám = không thuộc công nợ.

## Phụ lục C. Giả định và lưu ý

1. Quy ước mã công nợ `131-` hoặc `331-` cộng mã khách hàng là giả định của người lập, cần bộ phận kế toán xác nhận.
2. Sheet `BR`, `MV`, `VND112` trong file kết quả là bản sao giá trị. Công thức gốc và liên kết ngoài bị hỏng của file gốc không được mang sang.
3. Kết quả tìm tên công ty (`AO:AS`, `AT`), số hóa đơn trích được (`AN`, sheet `Chi tiết HĐ`) và nhận diện khách lẻ (`BR!AE`) là giá trị tính khi chạy script, không tự cập nhật khi thêm dòng mới. Cần chạy lại script khi có dữ liệu mới.
4. Gợi ý MST theo số tiền (`AT`) chỉ để tham khảo, không dùng để tự ghép.
5. RULE_010 chỉ áp dụng khi RULE_002 không xác định được MST. Nếu tên trong nội dung khớp đủ tin cậy với một công ty thì đi theo luồng khách hàng có MST. Giả định này cần xác nhận nếu muốn ưu tiên khách lẻ trước.
6. Nhận diện khách lẻ dựa vào từ khóa và hình dạng tên cá nhân, có thể nhận nhầm tên thương hiệu không có dấu hiệu tổ chức. Cần soát cột `BR!AE` và sửa tay nếu sai.
7. Hóa đơn khách lẻ có MST (ví dụ MST cá nhân 12 số) thì dùng MST đó làm mã khách hàng. Không có MST thì dùng mã khách lẻ `KHACHLE`.
8. RULE_010 không đối chiếu tên người chuyển tiền. Nếu muốn dùng thêm tên (ví dụ ưu tiên hóa đơn có tên trùng người chuyển) cần bổ sung rule.
9. Chưa có cột trạng thái APPROVED/REJECTED tự động, thao tác theo quy ước ở RULE_009.

## Phụ lục D. Kết quả kiểm thử

Đối chiếu với cột Mã hàng và cột Số hóa đơn mà kế toán đã nhập tay ở các file mẫu trước (chỉ dùng làm đáp án tham khảo, không phải đầu vào của quy trình):

| Hạng mục kiểm thử | Kết quả |
|---|---|
| Tách số hóa đơn từ nội dung thô (126 dòng có đáp án) | 89 dòng đúng hoàn toàn, 32 dòng nội dung không có số hóa đơn, 5 dòng sai hoặc thiếu (ví dụ nội dung ghi HĐ 29 nhưng kế toán nhập 30) |
| Tìm MST từ tên công ty, ngưỡng 0,6 và biên độ 0,05 (126 dòng có đáp án) | 103 dòng đúng, 0 dòng sai, 23 dòng bỏ qua do nội dung không có tên công ty |
| Gợi ý MST theo số tiền duy nhất (dòng bỏ qua có đáp án) | 17/17 gợi ý đúng |
| Ghép hóa đơn tự động bằng MST tìm từ tên so với khi MST nhập tay | Ghép được 70 dòng so với 89 dòng. 19 dòng không ghép được vì thiếu tên công ty, không có dòng nào ghép nhầm |
| Script gộp `matching_cong_no.py` chạy trên file mẫu | Công thức không lỗi, kết quả trùng 100% với các lần dựng file trước |

**Kiểm thử RULE_010 (khách lẻ).** BR của file mẫu không có hóa đơn khách lẻ nào (toàn bộ là công ty, hộ kinh doanh, trường học), nên rule được kiểm thử bằng dữ liệu thử có chủ đích (file `test_khach_le_input.xlsx`, kết quả ở `test_khach_le_ket_qua.xlsx`): thêm 9 hóa đơn khách lẻ vào BR và 9 giao dịch thu vào VND112. Kết quả đúng mong đợi ở cả 9 trường hợp:

| Tình huống thử | Kết quả |
|---|---|
| 1 hóa đơn tên cá nhân, thanh toán sau 1 ngày, số tiền khớp | Đã ghép - Rule khách lẻ (HĐ 901, lệch 1 ngày) |
| 2 hóa đơn "Khách lẻ" cùng số tiền (lệch 4 ngày và 0 ngày) | Đã ghép HĐ gần nhất (HĐ 903) nhờ cửa sổ hẹp |
| 2 hóa đơn cùng số tiền, cùng nằm trong ± 2 ngày | Nhiều HĐ/GD khách lẻ cùng số tiền (chuyển kiểm tra) |
| Hóa đơn "bán cho người tiêu dùng", thanh toán sau 50 ngày | Không ghép, thiếu MST (ngoài cửa sổ ngày) |
| Tên cá nhân, thanh toán sau 1 ngày | Đã ghép - Rule khách lẻ (HĐ 907) |
| 2 giao dịch cùng số tiền cho 1 hóa đơn | Cả 2 giao dịch chuyển kiểm tra (nhiều HĐ/GD khách lẻ) |
| Cá nhân có MST 12 số, tên khớp nội dung | Đi theo RULE_002 (tên) và RULE_004, giữ nguyên MST 12 số |
| Giao dịch tên công ty trùng số tiền với hóa đơn khách lẻ | Không áp dụng RULE_010, đi theo luồng khách hàng có MST |

Lưu ý: RULE_010 chưa được kiểm chứng trên dữ liệu thật vì mẫu không có hóa đơn khách lẻ. Nên chạy thử trên một kỳ có hóa đơn khách lẻ và soát lại kết quả. Ngoài ra, 72 dòng thuộc nhóm "chưa phân loại TK" có MST do tìm từ tên nhưng không có đáp án để đối chiếu, nên độ chính xác ở nhóm này chưa được kiểm chứng và vẫn cần kế toán soát lại.

## Phụ lục E. Lịch sử phiên bản

| Phiên bản | Thay đổi |
|---|---|
| 1.0 | Bản đầu: 9 rule, có nhánh "Mã hàng nhập tay luôn được ưu tiên" |
| 1.1 | Bỏ nhánh Mã hàng nhập tay vì cột Mã hàng ban đầu luôn trống: cột `D` giờ được điền bằng công thức, RULE_002 chỉ còn tìm MST từ tên. Thêm RULE_010 (khách lẻ: số tiền + ngày thanh toán gần ngày hóa đơn), EX_010 và 3 tham số khách lẻ. Giữ nguyên MST 12 số (cá nhân), chỉ gộp MST 13 số về 10 số |
