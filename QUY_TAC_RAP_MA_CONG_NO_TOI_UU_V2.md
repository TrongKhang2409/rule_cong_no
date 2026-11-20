# QUY TẮC RÁP MÃ CÔNG NỢ — ĐẶC TẢ CHUẨN HÓA V2.0

> **Mục tiêu:** Từ sổ bán ra (BR), sổ mua vào (MV) và sao kê ngân hàng, xác định **đối tượng công nợ**, **chứng từ liên quan** và **mức độ chắc chắn** cho từng giao dịch. Chỉ tự động ghi mã khi đủ chứng cứ; phần không đủ chứng cứ phải có đề xuất và lý do rõ ràng.
>
> **Phạm vi:** Đây là đặc tả xử lý, **không phải kết quả ráp dữ liệu**. Excel `LOOP1(1).xlsx` chỉ được sử dụng để kiểm tra cấu trúc và nhận diện tình huống thực tế. Những quy tắc mới bổ sung được ghi chú **[CẢI TIẾN]**; chưa coi là chính sách kế toán đã được phê duyệt.

## 1. Định nghĩa đầu ra và nguyên tắc bất biến

Phải tách biệt **mã đối tượng**, **tài khoản kế toán**, **mã công nợ**, **liên kết hóa đơn** và **trạng thái ghép**. Một giao dịch có thể tìm đúng khách hàng nhưng chưa tìm đủ bằng chứng để đối chiếu hóa đơn.

| Trường | Ý nghĩa | Ví dụ |
|---|---|---|
| `partyCode` | Mã đối tượng theo danh mục kế toán; có thể là MST hoặc mã nội bộ | `0306168465`, `KL`, `VIMO` |
| `taxIdRaw` / `taxIdNormalized` | MST gốc / MST được chuẩn hóa để tìm kiếm | `0300951119001` / `0300951119` (mã tra cứu của công ty mẹ) |
| `account` | Tài khoản công nợ đã phân loại, không mặc định chỉ có 131/331 | `131`, `331` |
| `proposedDebtCode` | Mã công nợ **đề xuất** theo cấu hình | `331-0306168465` |
| `finalDebtCode` | Mã được phép ghi chính thức | Chỉ khi tự động đủ điều kiện hoặc đã được duyệt |
| `matchedInvoiceIds` | Tập ID hóa đơn đã liên kết, không chỉ là chuỗi số HĐ | Danh sách khóa duy nhất |
| `status` / `ruleId` / `reasonCodes` | Kết luận, rule thắng và các lý do | `MATCHED`, `P03`, `AMOUNT_MISMATCH` |
| `evidence` | Chứng cứ dẫn tới quyết định | Nội dung gốc, từ khóa, giá trị so sánh, ứng viên khác |

**Bất biến bắt buộc:**

1. Không dùng thông tin ở các cột **đáp án/kiểm soát nhập tay** làm bằng chứng tự động. Trong file mẫu, các cột `P: R` của sheet `ACB 1121-01` chứa mã/tên/HĐ tham khảo và **không được sử dụng trong thuật toán**; chỉ dùng để kiểm thử độc lập nếu đã xác thực nguồn gốc.
2. Không suy diễn rằng **mã hàng = MST**. `KL`, `VIMO` và các mã nội bộ hợp lệ khác phải được lưu nguyên dạng; không chuyển thành MST giả.
3. Không mặc định `131-<MST>`/`331-<MST>` là mã chính thức của doanh nghiệp. Đây là **quy ước đề xuất từ tài liệu cũ**, chỉ sử dụng khi được cấu hình/phê duyệt; ưu tiên mapping mã kế toán chính thức nếu có.
4. Một hóa đơn không được ghép toàn phần lặp lại với nhiều giao dịch. Một giao dịch không được phân bổ số tiền vượt giá trị giao dịch. Mọi phân bổ cần sổ theo dõi (allocation ledger).
5. Không tự động ghép khi có nhiều ứng viên hợp lệ ngang cấp, thiếu bằng chứng cốt lõi, hoặc chứng cứ mâu thuẫn. Không được dùng độ khớp tên cao để bỏ qua sai lệch số tiền/hóa đơn.
6. Không chỉnh sửa ba sheet nguồn; ghi kết quả ở file/sheet riêng và lưu dấu vết quyết định. Quyết định chạy lại phải có tính xác định (cùng input, config và dữ liệu đối chiếu → cùng output).

## 2. Cấu trúc dữ liệu nguồn và mapping cột

### 2.1. Đối chiếu trực tiếp với Excel mẫu

| Vai trò | Sheet thực tế | Dòng tiêu đề / dữ liệu | Trường sử dụng |
|---|---|---|---|
| Bán ra | `BR` | Tiêu đề dòng 1; dữ liệu từ dòng 3 | `C` ký hiệu; `D` số HĐ; `E` ngày; `F` Mã KH; `G` tên; `U` Tổng TT; `M` TK Nợ |
| Mua vào | `MV` | Tiêu đề dòng 1; dữ liệu từ dòng 3 | `C` ký hiệu; `D` số HĐ; `E` ngày; `F` Mã KH; `G` tên; `U` Tổng TT; `N` TK Có |
| Sao kê | `ACB 1121-01` | Nhãn cột dòng 2–3; dữ liệu từ dòng 4 | `C` ngày; `D` Mã hàng (**đầu vào trống**); `E` đối tượng hiển thị bởi ngân hàng; `F` nội dung; `G` số HĐ (**đầu vào có thể trống**); `H` thu; `I` chi; `L` TK Nợ; `M` TK Có |
| Đáp án tham khảo | Cùng sheet sao kê | Các dòng dữ liệu | `P` mã tham khảo; `Q` tên/nội dung phân giải; `R` hóa đơn tham khảo — **không phải đầu vào thuật toán** |

**Phát hiện quan trọng từ file mẫu:** `BR` chứa nhiều người mua ghi là **“Khách hàng không lấy hóa đơn”** với `Mã KH` là `KL`, `VIMO` hoặc để trống. Vì vậy không thể dựa vào tên người mua để mặc định tất cả hóa đơn này thuộc một mã `KHACHLE`. Trong `MV` tồn tại MST chi nhánh 13 chữ số; hóa đơn có số HĐ trùng nhau giữa nhiều nhà cung cấp. `BR`/`MV` có dòng không thuộc công nợ 131/331 và có nghiệp vụ hóa đơn điều chỉnh. Tất cả phải được phân biệt khi lọc ứng viên.

### 2.2. Mapping linh hoạt [CẢI TIẾN]

- Cấu hình `sheets.sales`, `sheets.purchases`, `sheets.bank` theo tên hiện có, **không hard-code** `VND112`.
- Cấu hình `headerRow`, `dataStartRow`, `columns` và tài khoản ngân hàng (`1121-01`) riêng theo loại file; kiểm tra bằng nhãn cột, không chỉ vị trí.
- Bỏ qua dòng trống, dòng tổng cộng, dòng mẫu có giá trị kiểm thử `x`, dòng không phải giao dịch, và vùng định dạng dư. Không lấy `max_row` làm số lượng bản ghi thực.
- Đọc MST, mã đối tượng, ký hiệu và số HĐ như **chuỗi**, bảo toàn số 0 đầu. Tiền dùng số nguyên VND hoặc decimal chính xác, tuyệt đối không so sánh float xấp xỉ.
- Phân biệt `missing` (thiếu), `invalid` (sai định dạng), `zero` (giá trị bằng 0) và `not_applicable` (không áp dụng).
- Mỗi dòng có `sourceRowId = file/sheet/row` và `sourceFingerprint` để phát hiện thay đổi dữ liệu khi chạy lại.

## 3. Chuẩn hóa và xây dựng danh mục

### N01 — Chuẩn hóa nội dung và tên

1. Unicode normalize, chuyển chữ hoa, bỏ dấu để **so khớp**, giữ nguyên chuỗi gốc để hiển thị.
2. Chuẩn hóa khoảng trắng, dấu ngăn, biến thể `CTY/CONG TY`, `TNHH`, `CP`, `TM`, `SX`, `XNK` nhưng không xóa thông tin nhận dạng phân biệt.
3. Tách rõ thành phần: người chuyển, tên đối tác, MST, nội dung thanh toán, số HĐ, ngày chứng từ, kỳ dịch vụ, mã giao dịch ngân hàng. Không coi mọi chuỗi số là số HĐ.
4. Loại nhiễu định danh giao dịch, số thẻ che dấu, chuỗi thời gian, số tài khoản và tên doanh nghiệp sở hữu tài khoản; việc loại bỏ phải có regex rõ ràng và kiểm thử.
5. Tạo aliases từ danh bạ KH/NCC có xác nhận. Tên viết tắt, viết liền, lỗi chính tả có thể dùng tìm ứng viên, **không đủ để quyết định nếu đứng một mình**.

### N02 — MST, mã nội bộ, đối tượng

- `taxIdRaw` giữ nguyên. MST doanh nghiệp 13 số có thể tra cứu quan hệ với mã 10 số **nhưng giữ cả mã chi nhánh và mã mẹ**, không tự ghi đè mã đối tượng/đối tượng kế toán của chi nhánh.
- MST 10 số và mã cá nhân 12 số giữ nguyên theo tài liệu cũ; không biến một mã nội bộ thành MST chỉ vì trùng độ dài.
- Mã `KL`, `VIMO`, mã số khác (ví dụ mã đối tác dùng riêng) thuộc namespace `internalPartyCode`. Không gộp `KL` và `VIMO` trừ khi có **bảng ánh xạ chính thức**.
- Nhận diện một pháp nhân theo `taxId`, chi nhánh, aliases; một **mã hạch toán** theo danh mục kế toán. Quan hệ hai loại không mặc định 1:1.

### N03 — Khóa hóa đơn và số tiền

- Khóa hóa đơn ưu tiên: `sourceBook | partyCode hoặc taxIdRaw | invoiceSeries | invoiceNo | invoiceDate`. Nếu thiếu thành phần, tạo **khóa tra cứu không đầy đủ**, không coi là ID duy nhất.
- Nếu số HĐ trong nội dung chỉ có `HD 12`, không thể giả định đó là HĐ duy nhất trên cả năm; cần kiểm tra kỳ, ký hiệu, đối tượng và sự duy nhất.
- `invoiceAmount` lấy từ `U` (Tổng TT) nhưng kiểm tra hợp lệ, không âm không lý do, so với trạng thái thay thế/điều chỉnh và định khoản. Không tự coi HĐ đã thanh toán nếu dòng hóa đơn không thuộc phạm vi công nợ đang xét.
- `paymentAmount = H` (thu) hoặc `I` (chi) theo chiều hợp lệ; không cộng `H+I` để che lỗi khi cả hai cùng có tiền.

## 4. Quy tắc tiền kiểm — PHẢI CHẠY TRƯỚC MỌI MATCH

Các gate này không phải điểm ưu tiên cạnh tranh; một gate lỗi sẽ **dừng tự động ghi mã**, nhưng có thể sinh gợi ý kiểm tra.

| ID | Điều kiện kiểm tra | Kết quả nếu lỗi |
|---|---|---|
| `G00` | Dòng có giao dịch thực, ngày và nội dung/căn cứ tối thiểu, số tiền hợp lệ | `INVALID_DATA` |
| `G01` | Một dòng chỉ có **một chiều tiền**: `H > 0 XOR I > 0` | `INVALID_DIRECTION` |
| `G02` | Kiểm tra TK Nợ/Có và chiều tiền: thu thông thường `Nợ 112*/Có 131*`; chi thông thường `Nợ 331*/Có 112*` | `PENDING_ACCOUNT_REVIEW` khi thiếu, đảo chiều hoặc không nhất quán |
| `G03` | Loại giao dịch không thuộc đối chiếu 131/331: phí ngân hàng, lương, BHXH, thuế, rút tiền, chuyển nội bộ, doanh thu/chi phí khác | `OUT_OF_SCOPE`; **không dò tên hoặc số tiền để ép vào công nợ** |
| `G04` | Loại trừ invoice không hợp lệ, trùng chứng từ thực, hóa đơn điều chỉnh/thay thế chưa xử lý, dòng không thuộc nghiệp vụ công nợ phù hợp | `PENDING_INVOICE_REVIEW` |
| `G05` | Kiểm tra hóa đơn/giao dịch đã được phân bổ hoặc đã được duyệt, khóa không ghi đè | `ALREADY_ALLOCATED` / `LOCKED` / review |

**Thứ tự `G02/G03`:** trước tiên phân loại theo cặp **TK + chiều tiền + tài khoản ngân hàng**, sau đó xác định `OUT_OF_SCOPE` hoặc `PENDING_ACCOUNT_REVIEW`. Ví dụ thu vào 1121-01 nhưng Có 515 là thu lãi, không phải thu công nợ; chi từ ngân hàng với Nợ 1111 là rút tiền, không phải chi công nợ. Nếu chỉ thiếu tài khoản trong một trường hợp có vẻ phù hợp, không tự điền mã chính thức.

**Ngoại lệ đảo chiều:** hoàn tiền khách/NCC, bù trừ hoặc trả lại tiền không được ép sang `131`/`331` bằng dấu tiền; chuyển `PENDING_ACCOUNT_REVIEW` với nhãn nghiệp vụ và yêu cầu chứng cứ liên quan.

## 5. Thứ tự ưu tiên tìm kiếm và chấp nhận kết quả

**Nguyên tắc:** Tìm ứng viên theo nhiều nguồn trước, đánh giá bằng chứng sau, rồi mới quyết định. Không áp dụng kiểu “rule nào chạy trước cũng auto ghi mã” khi rule yếu mâu thuẫn với rule mạnh.

| Ưu tiên | Rule | Nguồn chứng cứ tối thiểu | Điều kiện tự động | Kết quả |
|---:|---|---|---|---|
| **P00** | `R00` — mã đã xác nhận | Bản ghi duyệt hợp lệ kèm người duyệt, lịch sử và khóa giao dịch | Quyền duyệt hợp lệ, không thay đổi chứng cứ nguồn | `APPROVED`, giữ mã |
| **P01** | `R01` — MST/mã đối tượng xác thực + HĐ + tiền | Đối tượng xác định chắc chắn, khóa HĐ đơn trị, đủ HĐ, tổng tiền bằng GD | Không xung đột, chưa phân bổ, phù hợp kỳ và chiều | `MATCHED` |
| **P02** | `R02` — số HĐ trích rõ + đối tượng xác định độc lập + tiền | Số HĐ rõ ràng; đối tượng từ danh bạ/MST/tên đủ điều kiện; tổng khớp | Khóa HĐ đầy đủ và duy nhất trong nhóm đối tượng | `MATCHED` |
| **P03** | `R03` — đối tượng chắc chắn + tiền đơn trị | Không có HĐ tin cậy; **1 GD ↔ 1 HĐ** theo đối tượng và số tiền | Duy nhất trên **toàn bộ dữ liệu chưa phân bổ**, không có chỉ dấu HĐ khác | `MATCHED_LOW_CONFIDENCE` hoặc `REVIEW` theo cấu hình |
| **P04** | `R04` — khách lẻ/đối tượng nội bộ có định danh | Mã `KL`/`VIMO` hợp lệ theo danh mục, có ràng buộc HĐ và tổng tiền | Mã nội bộ xác định độc lập, HĐ duy nhất, tiền khớp | `MATCHED` hoặc `REVIEW` theo chất lượng bằng chứng |
| **P05** | `R05` — khách lẻ bằng tiền + ngày | Không xác định đối tượng; HĐ bán lẻ trong cửa sổ thời gian | **Mặc định KHÔNG tự ghi mã**; chỉ gợi ý tập ứng viên | `PENDING_REVIEW` |
| **P06** | `R06` — gợi ý tên tương tự / MST theo số tiền | Chỉ là tín hiệu phụ để khoanh vùng | Không auto nếu đứng độc lập | `PENDING_REVIEW` |

**Giải thích thứ tự:** `R02` không có nghĩa là luôn thấp hơn `R01`: cả hai là nhóm bằng chứng mạnh về đối tượng + HĐ + tiền, khác nhau ở cách thu thập. Nếu cùng một giao dịch cho hai đối tượng/HĐ khác nhau, **không lấy rule có thứ hạng cao hơn để ghi đè**, mà chuyển `CONFLICT_REVIEW`. Ngưỡng tên không được “nâng hạng” một ghép thiếu hóa đơn hoặc lệch tiền.

### R01/R02 — Hóa đơn được nêu cụ thể

1. Trích tập số HĐ từ nội dung `F`, hỗ trợ `HĐ`, `HD`, `HOA DON`, `HĐ SỐ`, danh sách cách nhau bởi `, . ; / & VA`; không tự trích số từ ngày, kỳ thanh toán, mã giao dịch và chuỗi `HD-<mã giao dịch>`.
2. Xác định `partyCode`/MST **độc lập** bằng mã/MST xuất hiện trong nội dung hoặc tên đối tác trùng danh mục đã chuẩn hóa. Một MST của chính công ty hoặc của bên ngân hàng không phải MST đối tác.
3. Lọc đúng sổ: thu → BR; chi → MV; lọc chính xác đối tượng, ký hiệu (nếu có), số HĐ, kỳ, tình trạng hóa đơn.
4. Với nhiều HĐ trong nội dung, tất cả phải được tìm thấy đơn trị; so tổng số tiền hóa đơn hợp lệ với giao dịch. Nếu thiếu một HĐ thì **không** chấp nhận chỉ những HĐ còn lại.
5. Chênh lệch tuyệt đối `≤ amountTolerance` (mặc định **0 VND**) mới đủ điều kiện. Lệch 1 đồng cũng đưa vào kiểm tra. Với thanh toán gộp, tổng kiểm tra chính xác bằng decimal/số nguyên.
6. Không ghép cùng HĐ hai lần, trừ khi sử dụng module phân bổ thanh toán từng phần được kế toán chấp thuận.

### R03 — Đối tượng + số tiền duy nhất

- Chỉ dùng khi **không có số HĐ đủ tin cậy**; đã có đối tượng tin cậy và đã qua mọi gate.
- HĐ ứng viên đúng **một** với `partyCode`/MST và `invoiceAmount = paymentAmount`; GD ứng viên cùng đối tượng + số tiền đúng **một**; HĐ không bị phân bổ.
- Nếu cùng số tiền trùng ở nhiều HĐ hoặc GD, không dùng “ngày gần nhất” để tự quyết. Thời gian chỉ có thể hỗ trợ giải thích hoặc loại ứng viên không hợp lý.
- Mặc định có thể tự ghi mã khi được kiểm chứng theo quy trình; khuyến nghị ban đầu `autoApproveAmountOnly = false` để kế toán đánh giá trên dữ liệu thật. `MATCHED_LOW_CONFIDENCE` không đồng nghĩa `APPROVED`.

### R04/R05 — Khách lẻ [CẢI TIẾN QUAN TRỌNG]

1. **Không đồng nhất** các chuỗi `KL`, `VIMO`, `KHACHLE` và MST. `KL`, `VIMO` là mã đối tượng đang xuất hiện trong BR, muốn chuyển mã phải có bảng mapping được duyệt.
2. Không mặc định mọi người có tên cá nhân đều là khách lẻ; không mặc định tên chung “Khách hàng không lấy hóa đơn” xác định được một khách cụ thể.
3. Nếu tìm được mã nội bộ **từ nguồn độc lập hợp lệ** (danh mục, mã giao dịch POS/đối tác được xác thực), ghép theo mã + HĐ + tiền như R04.
4. Khi chỉ có số tiền + ngày ±7 (ưu tiên thông tin trong ±2), đưa danh sách ứng viên kèm ngày lệch và lý do. **Không tự gán `131-KHACHLE`** chỉ từ tiền + ngày; đây là thay đổi an toàn so với RULE_010 cũ.
5. Không tự ghép một hóa đơn bán lẻ với giao dịch thanh toán có nội dung không liên quan; trong file mẫu có thanh toán qua cổng/đơn vị trung gian và giao dịch ghi chú cá nhân, nên cần danh mục nguồn thu và mã tham chiếu nếu muốn tự động hóa nhiều hơn.

### R06 — Xác định đối tượng từ tên [CẢI TIẾN]

- Thu thập ứng viên tên bằng normalize, aliases, token hiếm, tên viết tắt, sai chính tả nhẹ; loại tên công ty sở hữu tài khoản.
- Điểm tham khảo kế thừa tài liệu cũ: `nameScore ≥ 0.60`, cách ứng viên thứ hai `≥ 0.05`. Đây là **ngưỡng tạo ứng viên**, không bảo đảm độ chính xác chung; cần hiệu chỉnh sau kiểm thử.
- Nếu tên khớp nhưng MST/đối tượng mâu thuẫn với HĐ → `CONFLICT_REVIEW`.
- Gợi ý đối tượng từ **số tiền duy nhất** được hiển thị ở trường `suggestedPartyCode`, không được đưa ngược làm đầu vào để tự chứng minh chính phép ghép (tránh vòng lặp suy luận).

## 6. Ngày, kỳ trước và hóa đơn điều chỉnh

- **Không hard-code năm 2025** làm điều kiện loại trừ. Hãy suy ra `statementPeriod`, `invoicePeriod`, `openingBalancePeriod` từ dữ liệu/cấu hình.
- Phân biệt `ngày thanh toán`, `ngày lập hóa đơn`, `ngày thực hiện dịch vụ` và `ngày ghi trong nội dung`. Ví dụ `KỲ (THÁNG) 12/2025` có thể chỉ là **kỳ tiền điện**, không có nghĩa đó là hóa đơn năm 2025.
- Giao dịch liên quan HĐ kỳ trước, ngoài phạm vi dữ liệu BR/MV hiện tại → `PRIOR_PERIOD_REVIEW`; nếu có sổ công nợ chuyển kỳ đã xác minh, dùng nguồn đó trước khi xét không tìm thấy hóa đơn.
- HĐ bị điều chỉnh/thay thế, ghi chú điều chỉnh giảm, tổng tiền bằng 0 hoặc âm → `ADJUSTMENT_REVIEW` nếu thiếu chuỗi hóa đơn gốc–điều chỉnh; không tùy tiện cộng trừ các số HĐ trùng.
- Không loại HĐ chỉ vì ngày thanh toán xa ngày HĐ; nhiều khoản công nợ thanh toán muộn. Ngày là **tín hiệu hỗ trợ**, không phải điều kiện bắt buộc của R01–R03.

## 7. Ràng buộc chống ghép sai và đối soát nhiều-một

### 7.1. Chống trùng và xung đột

- `invoiceKey` có >1 bản ghi: phân biệt **trùng dữ liệu nhập** với **hai chứng từ hợp lệ khác kỳ/ký hiệu**; chưa rõ thì `DUPLICATE_INVOICE_KEY`.
- Có >1 đối tượng/HĐ/tập HĐ đạt cùng điều kiện: `AMBIGUOUS_CANDIDATES`; không lấy ứng viên đầu hoặc ngày gần nhất.
- Một HĐ đã ghép với giao dịch khác: `ALREADY_ALLOCATED`; không auto ghép lại.
- Đối tượng xác định theo tên A nhưng hóa đơn nêu B: `PARTY_INVOICE_CONFLICT`.
- Số HĐ trích được khác với một mã tham chiếu rõ ràng khác: `INVOICE_REFERENCE_CONFLICT`.

### 7.2. Thanh toán từng phần, gộp khoản và thừa tiền

- **Mặc định V2.0:** R01/R02 chỉ auto khi toàn bộ tập HĐ và tiền bằng nhau. Thanh toán một phần, ứng trước, thu gộp nhiều kỳ, phí trừ tại nguồn, hoàn tiền, bù trừ → `PARTIAL_OR_COMPOSITE_REVIEW`.
- **Mở rộng tùy chọn** khi có bảng công nợ tồn đầu kỳ + phân bổ: duy trì `invoiceOpenAmount`; phân bổ tối đa `min(paymentRemaining, invoiceOpenAmount)`, có chứng cứ về mục đích thanh toán và phê duyệt. Không tự suy rằng tổng tiền khớp một tổ hợp hóa đơn bất kỳ là đủ, vì có thể có nhiều tổ hợp trùng tổng.
- Quy tắc chính: `SUM(allocations theo giao dịch) ≤ paymentAmount` và `SUM(allocations theo HĐ) ≤ invoiceOpenAmount ban đầu` (xét chứng từ hoàn/đảo riêng). Không ghi trạng thái `FULLY_SETTLED` nếu còn chênh lệch.

## 8. Luồng quyết định chuẩn — dùng làm thuật toán triển khai

```text
FOR mỗi dòng sao kê thực tế:
  1. Đọc dữ liệu gốc; chuẩn hóa, xác định ID nguồn.
  2. Nếu có bản ghi đã phê duyệt hợp lệ và nguồn không đổi: khóa kết quả (R00).
  3. Kiểm tra G00..G05; không đủ điều kiện thì ghi đúng trạng thái, không auto.
  4. Chọn sổ BR/MV theo TK + chiều tiền; lập ứng viên đối tượng độc lập.
  5. Tách số HĐ và loại số ngày/mã giao dịch; tạo ứng viên HĐ.
  6. Kiểm tra R01/R02 (đối tượng + HĐ + tổng tiền + tính duy nhất).
  7. Nếu không có HĐ tin cậy: thử R03 (đối tượng + tiền duy nhất).
  8. Nếu đối tượng là mã nội bộ đã được xác thực: thử R04.
  9. Nếu chưa đủ căn cứ: R05 (khách lẻ) và R06 (tên/tiền) chỉ để gợi ý.
 10. Soát xung đột toàn cục giữa các giao dịch và hóa đơn chưa phân bổ.
 11. Chốt status, ruleId, amountDifference, evidence, proposedDebtCode.
 12. Ghi finalDebtCode chỉ nếu đạt chính sách auto hoặc đã được duyệt.
END
```

**Lưu ý về bước 6–9:** Các ứng viên R01/R02/R03/R04 cần được đánh giá trên cùng snapshot; nếu một phép ghép mức thấp chiếm mất HĐ của một phép ghép mức cao, giải quyết **toàn cục**: giữ phương án có bằng chứng cao hơn, không phụ thuộc thứ tự dòng Excel. Hai phương án ngang cấp và đều hợp lý → review, không chọn tùy ý. Ghi nhật ký mọi thay đổi nếu tính toán lại.

## 9. Ma trận trạng thái cuối và điều kiện ghi mã

| Status | Khi sử dụng | `proposedDebtCode` | `finalDebtCode` |
|---|---|---|---|
| `APPROVED` | Được người có quyền phê duyệt | Có | Theo phê duyệt |
| `MATCHED` | R01/R02/R04 đủ điều kiện và không xung đột, chính sách cho phép | Có | Có |
| `MATCHED_LOW_CONFIDENCE` | R03 phù hợp nhưng cần soát theo cấu hình | Có | Mặc định trống |
| `PENDING_REVIEW` | Thiếu bằng chứng; khách lẻ tiền + ngày; hóa đơn kỳ trước; điều chỉnh | Nếu xác định đối tượng thì có | Trống |
| `AMBIGUOUS` | Nhiều ứng viên hoặc trùng | Có thể có nhiều gợi ý | Trống |
| `CONFLICT` | Chứng cứ mâu thuẫn/đã phân bổ | Gợi ý và lý do | Trống |
| `NOT_MATCHED` | Không tìm thấy đối tượng hoặc HĐ đủ điều kiện | Nếu có đối tượng | Trống |
| `OUT_OF_SCOPE` | Giao dịch không thuộc công nợ đối chiếu | Trống | Trống |
| `INVALID_DATA` | Dòng nguồn lỗi/sai chiều/sai tiền | Trống | Trống |
| `REJECTED` | Người duyệt bác kết quả | Giữ lịch sử đề xuất | Trống hoặc theo nghiệp vụ đã duyệt |

**Thứ tự khóa mã:** `APPROVED` của người dùng có quyền > mã tự động đủ điều kiện > không có mã chính thức. Nếu mã duyệt khác đề xuất phải có lý do, người duyệt và timestamp. Không để việc “nhập mã vào ô” tự động đồng nghĩa đã phê duyệt nếu thiếu audit.

## 10. Danh mục mã lý do (reason codes)

| Nhóm | Mã lý do | Diễn giải |
|---|---|---|
| Nguồn | `MISSING_AMOUNT`, `INVALID_DIRECTION`, `MISSING_ACCOUNT` | Thiếu/sai tiền, chiều hoặc định khoản |
| Phân loại | `NON_DEBT_TRANSACTION`, `REVERSED_ENTRY` | Ngoài công nợ hoặc bút toán đảo chiều |
| Đối tượng | `PARTY_NOT_FOUND`, `PARTY_AMBIGUOUS`, `INTERNAL_CODE_UNMAPPED` | Thiếu/chồng chéo danh mục |
| Hóa đơn | `INVOICE_NOT_FOUND`, `INVOICE_AMBIGUOUS`, `INVOICE_PARSE_UNCERTAIN` | Không tìm thấy, nhiều kết quả, trích số không chắc chắn |
| Tiền | `AMOUNT_MISMATCH`, `PARTIAL_OR_COMPOSITE`, `POSSIBLE_FEE` | Lệch tiền, thanh toán từng phần/gộp, nghi có phí |
| Kỳ | `PRIOR_PERIOD`, `ADJUSTMENT_UNRESOLVED` | Chứng từ ngoài kỳ hoặc điều chỉnh |
| Xung đột | `ALREADY_ALLOCATED`, `DUPLICATE_INVOICE_KEY`, `PARTY_INVOICE_CONFLICT` | Trùng/đã ghép/mâu thuẫn |
| Phê duyệt | `OVERRIDE_REQUIRES_APPROVAL`, `SOURCE_CHANGED_AFTER_APPROVAL` | Yêu cầu duyệt lại |

Một dòng được phép có **nhiều reason codes** nhưng chỉ có một `status` kết luận; `primaryReason` theo mức độ nghiêm trọng: sai nguồn/chiều > xung đột > trùng > lệch tiền > thiếu đối tượng/HĐ > thiếu bằng chứng mềm.

## 11. Tham số cấu hình đề xuất

```yaml
version: '2.0'
sheets:
  sales: BR
  purchases: MV
  bank: 'ACB 1121-01'
rows:
  salesStart: 3
  purchasesStart: 3
  bankStart: 4
bank:
  accountCode: '1121-01'
  moneyUnit: VND
mapping:
  sales: {invoiceSeries: C, invoiceNo: D, invoiceDate: E, partyCode: F, partyName: G, amount: U, debitAccount: M}
  purchases: {invoiceSeries: C, invoiceNo: D, invoiceDate: E, partyCode: F, partyName: G, amount: U, creditAccount: N}
  bank: {date: C, content: F, receive: H, pay: I, debitAccount: L, creditAccount: M}
  # Các cột P/Q/R của sheet bank chỉ có thể sử dụng trong bộ đánh giá, tuyệt đối không sử dụng khi matching.
match:
  amountToleranceVnd: 0
  nameCandidateThreshold: 0.60
  nameCandidateMargin: 0.05
  retailWideDays: 7
  retailTightDays: 2
  autoApproveExactInvoice: true
  autoApproveAmountOnly: false
  autoApproveRetailByAmountAndDate: false
  allowPartialPaymentAuto: false
  preserveInvoiceSeries: true
  requireNoGlobalConflict: true
party:
  preserveInternalCodes: true
  # Chỉ nhập các ánh xạ đã được bộ phận kế toán xác nhận.
  aliases: {}
  internalCodeMapping: {}
debtCode:
  # Dùng danh mục kế toán làm nguồn chuẩn; nếu không có, chỉ xuất mã đề xuất.
  source: accounting_master_if_available
  proposedFormat: '{account}-{partyCode}'
```

## 12. Định dạng output khuyến nghị

Tạo **4 bảng/sheet kết quả**, không thêm hàng chục cột công thức rải rác trên dữ liệu gốc:

1. **`Ket_qua`**: `transactionId`, dòng nguồn, ngày, nội dung, thu/chi, TK Nợ/Có, đối tượng, mã công nợ đề xuất, mã chính thức, trạng thái, rule thắng, reason, số HĐ liên kết, tổng tiền HĐ, lệch tiền, độ tin cậy, tóm tắt bằng chứng.
2. **`Kiem_tra_thu_cong`**: chỉ các dòng chưa chắc chắn, sắp theo mức độ ảnh hưởng và lý do; kèm tối đa vài ứng viên, bằng chứng và trường nhập quyết định.
3. **`Lien_ket_hoa_don`**: bảng nhiều-nhiều: giao dịch ↔ hóa đơn ↔ số tiền phân bổ ↔ nguồn chứng cứ; dùng để chống trùng và đối soát.
4. **`Nhat_ky`**: phiên bản rule, hash dữ liệu nguồn, thời gian chạy, người duyệt, mã trước/sau, lý do duyệt/bác và lịch sử sửa.

Các bảng danh mục chuẩn hóa đối tác và hóa đơn có thể xuất riêng để kiểm toán khi cần. Không điền công thức hay kết quả vào cột `D` của sao kê gốc, trừ khi có yêu cầu xuất bản sao đã được đối soát.

## 13. Kiểm thử và tiêu chuẩn nghiệm thu

**Test tối thiểu cần có:**

- 1 HĐ–1 GD đầy đủ MST + HĐ + tiền; nhiều HĐ–1 GD đủ tổng; thiếu 1 HĐ; số HĐ trùng giữa hai công ty; trùng số HĐ khác ký hiệu/kỳ; trùng nhiều GD cùng tiền.
- Khác số tiền 1 VND; thanh toán một phần; thanh toán quá; nhiều giao dịch thanh toán một hóa đơn; hoàn tiền; thu/chi ngược chiều; giao dịch ngoài công nợ.
- MST chi nhánh 13 số; MST/mã 12 số; mã nội bộ `KL` và `VIMO` khác nhau; người mua ghi “Khách hàng không lấy hóa đơn” nhưng thiếu mã; thanh toán qua bên thứ ba.
- Nội dung `HD 01.07.12.16`; ngày `02/01/2026`; `KỲ 12/2025`; ký tự viết liền; mã giao dịch dài có tiền tố `HD`; tên công ty trùng từ chung.
- Hóa đơn điều chỉnh/thay thế; tổng tiền 0; hóa đơn bán ra/mua vào không thuộc 131/331; kỳ trước nhưng chưa nhập số dư đầu kỳ.
- Chạy lại đúng một input → cùng kết quả; đổi thứ tự dòng → không đổi matching; dữ liệu nguồn đổi sau APPROVED → phải cảnh báo; cột P/Q/R thay đổi → **không được làm thay đổi kết quả thuật toán**.

**Tiêu chuẩn xuất kết quả:**

- Không có giao dịch bị ghép toàn phần hai lần; không vượt tiền nguồn.
- Mọi `MATCHED` có evidence đầy đủ, số tiền và đối tượng phù hợp; không có mã chính thức cho dòng thiếu bằng chứng.
- Tổng số dòng phân loại bằng tổng dòng giao dịch hợp lệ + dòng nguồn lỗi đã ghi nhận (không thất lạc dòng).
- Không lấy tỷ lệ ghép cao làm KPI duy nhất; theo dõi **precision (tỷ lệ tự ghép đúng)**, coverage (tỷ lệ tự ghép), số trường hợp cần kiểm tra và số xung đột. Ưu tiên precision trước coverage.
- Thống kê nghiệm thu **mới** phải được tính bằng bộ test có nhãn độc lập; các con số trong tài liệu 1.1 gắn với file mẫu cũ nên **không áp dụng cho `LOOP1(1).xlsx`**.

## 14. Những điểm cần kế toán xác nhận trước khi triển khai ghi mã chính thức

| Quyết định | Giá trị mặc định an toàn | Lý do cần xác nhận |
|---|---|---|
| Cấu trúc mã công nợ | Chỉ `proposedDebtCode`; lấy danh mục kế toán làm chuẩn | `131-<mã>`, `331-<mã>` mới là giả định |
| Quan hệ `KL` – `VIMO` – khách không lấy HĐ | Giữ **tách biệt** | File mẫu có cả các mã này |
| Mã công nợ của chi nhánh 13 số | Giữ mã gốc, MST mẹ chỉ để tra cứu | Tránh gộp sai đối tượng kế toán |
| HĐ BR/MV thanh toán tiền mặt 1111 | Không tự ghép với ngân hàng 112* nếu không có căn cứ | Phải xem đúng dòng công nợ/nghiệp vụ |
| Cho phép R03 tự ghi mã | **Tắt ban đầu** | Cùng tiền không luôn là cùng hóa đơn |
| Cho phép khách lẻ tiền + ngày tự ghi mã | **Tắt** | Rủi ro cao khi nhiều HĐ lẻ trùng tiền |
| Thanh toán từng phần, trả thừa, phí thu hộ | Review | Cần nghiệp vụ phân bổ và số dư đầu kỳ |
| Danh mục bí danh, mapping POS/cổng thanh toán | Chỉ dùng bản đã duyệt | Tăng tỷ lệ ghép mà không đánh đổi độ chính xác |

## 15. So sánh nhanh với tài liệu cũ (v1.1 → v2.0)

| Điểm trong v1.1 | Hướng xử lý v2.0 |
|---|---|
| Hard-code `VND112`, số dòng 310, BR/MV 1100 | Mapping động, không giới hạn nhân tạo; sheet thật `ACB 1121-01` |
| `Mã hàng = MST` hoặc `KHACHLE` | Tách mã nội bộ/MST/mã hạch toán; giữ `KL`, `VIMO` |
| Gộp MST chi nhánh 13 số thành 10 số | Giữ mã chi nhánh, dùng MST mẹ như thuộc tính tra cứu |
| Số HĐ chủ yếu bằng MST + số HĐ | Dùng thêm ký hiệu, kỳ, ngày, mã nội bộ và ID nguồn |
| Khách lẻ tiền + ngày có thể MATCHED | Mặc định chỉ đề xuất, không auto |
| Kỳ trước xác định bằng chuỗi năm 2025 | Xác định kỳ linh hoạt, phân biệt kỳ dịch vụ và ngày hóa đơn |
| Thiếu cơ chế kiểm tra phân bổ toàn cục | Sổ phân bổ, chống sử dụng HĐ lặp và xung đột |
| `MATCHED` và `APPROVED` dễ lẫn | Quy định quyền ghi mã chính thức, nhật ký và khóa duyệt |
| Dùng độ tin cậy tên ≥0,6 để suy ra MST | Chỉ là tín hiệu ứng viên, phải xác minh bằng chứng độc lập |

---

**Nguồn kế thừa:** `QUY_TAC_RAP_MA_CONG_NO(2).md` (bản 1.1): nguyên tắc kiểm tra, các RULE_001–010, danh mục ngoại lệ, ngưỡng tham số và logic gợi ý. **Nguồn tham khảo cấu trúc:** `LOOP1(1).xlsx` (`BR`, `MV`, `ACB 1121-01`). **Các quy định bổ sung [CẢI TIẾN]** là thiết kế đề xuất nhằm giảm rủi ro ghép nhầm, chưa phải xác nhận nghiệp vụ từ doanh nghiệp.
