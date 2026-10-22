<div align="center">

## Website phân tích GPA (Grade Point Average) áp dụng tại Đại học Cần Thơ
[![Live Demo](https://img.shields.io/badge/Live%20Demo-GitHub%20Pages-1E88E5?style=for-the-badge&logo=github&logoColor=white)](https://huynhthanhphong231004it.github.io/WebPhanTichGPA_CTU/)

**Sinh viên thực hiện:** Huỳnh Thanh Phong  - Khoa học máy tính, Trường Công nghệ thông tin và Truyền thông, Đại học Cần Thơ

[![Facebook](https://img.shields.io/badge/Facebook-Huỳnh%20Thanh%20Phong-E53935?style=for-the-badge&logo=facebook&logoColor=white&labelColor=333333)](https://www.facebook.com/huynh.thanh.phong.561667)
[![TikTok](https://img.shields.io/badge/TikTok-@reorioll-E53935?style=for-the-badge&logo=tiktok&logoColor=white&labelColor=333333)](https://www.tiktok.com/@reorioll)
[![YouTube](https://img.shields.io/badge/YouTube-ReoRioll-E53935?style=for-the-badge&logo=youtube&logoColor=white&labelColor=333333)](https://www.youtube.com/@ReoRioll-2304CICTCTU)


</div>

---

## 1. Tóm tắt

Kết quả học tập thường chỉ được thông báo dưới dạng một con số cuối cùng, trong khi sinh viên cần hiểu *con số đó hình thành thế nào*, *cần đạt bao nhiêu ở các kỳ còn lại*, và cán bộ lớp cần *xếp loại cả danh sách* một cách nhất quán. Dự án giải quyết ba nhu cầu này bằng ba module độc lập trong một giao diện duy nhất.

| Module | Câu hỏi giải quyết | Đầu vào | Đầu ra |
|---|---|---|---|
| **M1 – Tính GPA** | GPA học kỳ của tôi là bao nhiêu và được tính ra sao? | Điểm hệ 10, số tín chỉ từng môn | GPA hệ 4.0 + bảng chi tiết từng môn + công thức |
| **M2 – Ước lượng tốt nghiệp** | Tôi cần đạt GPA bao nhiêu ở phần còn lại để tới mục tiêu? | GPA hiện tại, tín chỉ đã tích lũy, tổng tín chỉ ngành, GPA mục tiêu | GPA cần đạt + gợi ý phân phối điểm chữ |
| **M3 – Xếp loại học kỳ** | Cả danh sách này được xếp loại và xếp hạng thế nào? | File Excel (GPA, điểm rèn luyện, tín chỉ đăng ký) | Bảng đã xếp loại và sắp xếp thứ hạng |

| Thông tin | Chi tiết |
|---|---|
| Kiến trúc | Client-side thuần, không máy chủ, không cơ sở dữ liệu |
| Công nghệ | HTML, JavaScript, Bootstrap 5.3, SheetJS (đọc Excel), MathJax 3 (hiển thị công thức) |
| Cách chạy | Mở `index.html` bằng trình duyệt |
| Cấu trúc mã | `index.html` (giao diện + hướng dẫn) · `web.js` (logic 3 module) · `Demo.png` (mẫu file Excel) |

---

## 2. Phương pháp

### 2.1 Kiến trúc hệ thống

```mermaid
flowchart LR
    %% Định nghĩa phong cách màu sắc chuẩn Academic (Nhiều màu, chuyên nghiệp)
    classDef userNode fill:#E3F2FD,stroke:#1565C0,stroke-width:2px,color:#0D47A1,font-weight:bold;
    classDef decisionNode fill:#FFF3E0,stroke:#EF6C00,stroke-width:2px,color:#E65100,font-weight:bold;
    classDef moduleNode fill:#E8F5E9,stroke:#2E7D32,stroke-width:1.5px,color:#1B5E20,font-weight:bold;
    classDef dataNode fill:#FFFDE7,stroke:#FBC02D,stroke-width:1.5px,color:#F57F17,font-weight:bold;
    classDef resultNode fill:#F3E5F5,stroke:#7B1FA2,stroke-width:1.5px,color:#4A148C,font-weight:bold;
    classDef mathNode fill:#FFEBEE,stroke:#C62828,stroke-width:2px,color:#B71C1C,font-weight:bold;

    %% Các nút trong sơ đồ
    U([Sinh viên / Cán bộ lớp]):::userNode
    S{Chọn chức năng}:::decisionNode
    
    M1["M1: Tính GPA"]:::moduleNode
    M2["M2: Ước lượng<br/>tốt nghiệp"]:::moduleNode
    M3["M3: Xếp loại<br/>học kỳ"]:::moduleNode
    
    X[(File Excel<br/>.xlsx / .xls)]:::dataNode
    
    R["Kết quả & Chi tiết<br/>tính toán"]:::resultNode
    MJ["MathJax<br/>Hiển thị công thức"]:::mathNode

    %% Luồng liên kết
    U --> S
    
    S -->|gpa| M1
    S -->|estimate| M2
    S -->|rank| M3
    
    X -->|FileReader + SheetJS| M3
    
    M1 --> R
    M2 --> R
    M3 --> R
    
    R --> MJ
```

Toàn bộ xử lý diễn ra trong trình duyệt; file Excel được đọc cục bộ bằng `FileReader`, không gửi đi đâu.

### 2.2 Quy đổi điểm (dùng chung cho M1 và M2)

Bảng quy đổi điểm và công thức tính GPA được áp dụng chuẩn xác theo Quy định công tác học vụ dành cho sinh viên trình độ đại học hệ chính quy (ban hành kèm theo Quyết định số 3266/QĐ-ĐHCT ngày 15/08/2024 của Hiệu trưởng Trường Đại học Cần Thơ, áp dụng từ Học kỳ 2, năm học 2024 – 2025), dựa trên Quy chế đào tạo trình độ đại học (Thông tư số 08/2021/TT-BGDĐT của Bộ Giáo dục và Đào tạo).Việc quy đổi từ điểm hệ 10 sang điểm chữ (A, B+, B, C+, C, D+, D, F) và chuyển tiếp sang thang điểm 4.0 giúp phản ánh chính xác phân mức học lực, đảm bảo tính minh bạch trong việc xếp loại học kỳ, đánh giá tiến độ tích lũy và xét tốt nghiệp theo đúng hệ thống đào tạo theo tín chỉ.

> [![Thang A](https://img.shields.io/badge/%E2%89%A5_9.0-Điểm_A_|_Hệ_4.0-800000?style=for-the-badge&labelColor=333333)](#)
> [![Thang B+](https://img.shields.io/badge/%E2%89%A5_8.0-Điểm_B+_|_Hệ_3.5-A30000?style=for-the-badge&labelColor=333333)](#)
> [![Thang B](https://img.shields.io/badge/%E2%89%A5_7.0-Điểm_B_|_Hệ_3.0-C62828?style=for-the-badge&labelColor=333333)](#)
> [![Thang C+](https://img.shields.io/badge/%E2%89%A5_6.5-Điểm_C+_|_Hệ_2.5-D32F2F?style=for-the-badge&labelColor=333333)](#)
> [![Thang C](https://img.shields.io/badge/%E2%89%A5_5.5-Điểm_C_|_Hệ_2.0-E53935?style=for-the-badge&labelColor=333333)](#)
> [![Thang D+](https://img.shields.io/badge/%E2%89%A5_5.0-Điểm_D+_|_Hệ_1.5-EF5350?style=for-the-badge&labelColor=333333)](#)
> [![Thang D](https://img.shields.io/badge/%E2%89%A5_4.0-Điểm_D_|_Hệ_1.0-E57373?style=for-the-badge&labelColor=333333)](#)
> [![Thang F](https://img.shields.io/badge/<_4.0-Điểm_F_|_Hệ_0.0-EF9A9A?style=for-the-badge&labelColor=333333)](#)
> 
### 2.3 M1 – Tính GPA học kỳ

$$
\text{GPA} = \frac{\sum_{i=1}^{n} (x_i \times c_i)}{\sum_{i=1}^{n} c_i}
$$

với $x_i$ là điểm hệ 4.0 và $c_i$ là số tín chỉ của môn $i$. Môn điểm **F** bị loại khỏi cả tử số và mẫu số theo quy ước của công cụ.

### 2.4 M2 – Ước lượng GPA tốt nghiệp (bài toán nghịch)

$$
\text{GPA}_{cần} = \frac{G_{mt} \times C_{tổng} - G_{ht} \times C_{ht}}{C_{tổng} - C_{ht}}
$$

với $G_{mt}$ là GPA mục tiêu, $G_{ht}$ và $C_{ht}$ là GPA và tín chỉ hiện tại, $C_{tổng}$ là tổng tín chỉ của ngành.

```mermaid
flowchart TD
    %% Định nghĩa phong cách màu sắc chuẩn Academic (Nhiều màu, chuyên nghiệp)
    classDef inputNode fill:#E3F2FD,stroke:#1565C0,stroke-width:2px,color:#0D47A1,font-weight:bold;
    classDef processNode fill:#E8F5E9,stroke:#2E7D32,stroke-width:1.5px,color:#1B5E20,font-weight:bold;
    classDef decisionNode fill:#FFF3E0,stroke:#EF6C00,stroke-width:2px,color:#E65100,font-weight:bold;
    classDef warningNode fill:#FFEBEE,stroke:#C62828,stroke-width:2px,color:#B71C1C,font-weight:bold;
    classDef algoNode fill:#F3E5F5,stroke:#7B1FA2,stroke-width:1.5px,color:#4A148C,font-weight:bold;
    classDef successNode fill:#E0F2F1,stroke:#00695C,stroke-width:2px,color:#004D40,font-weight:bold;

    %% Các nút trong sơ đồ
    A["Nhập dữ liệu:<br/>GPA hiện tại · TC hiện tại<br/>Tổng TC ngành · GPA mục tiêu"]:::inputNode
    
    B["Tính số tín chỉ (TC) còn lại:<br/>TC còn lại = Tổng TC - TC hiện tại"]:::processNode
    C["Tính GPA cần đạt cho các môn còn lại"]:::processNode
    
    D{"GPA cần đạt > 4.0?"}:::decisionNode
    
    E[" Cảnh báo:<br/>Mục tiêu không khả thi!"]:::warningNode
    
    F["Quy đổi ước tính số môn:<br/>Số môn = TC còn lại / 3"]:::processNode
    G["Chạy thuật toán Backtracking<br/>trên tập điểm [A, B+, B, C+, C]"]:::algoNode
    
    H{"Có tổ hợp có TB ≥ GPA cần đạt?"}:::decisionNode
    
    I[" Chọn tổ hợp tối ưu<br/>có trung bình gần ngưỡng nhất"]:::successNode
    J[" Thông báo:<br/>Không thể phân phối phù hợp"]:::warningNode

    %% Luồng liên kết thuật toán
    A --> B
    B --> C
    C --> D
    
    D -->|Có| E
    D -->|Không| F
    
    F --> G
    G --> H
    
    H -->|Có| I
    H -->|Không| J
```

| Tham số của bước gợi ý | Giá trị |
|---|---|
| Tín chỉ mỗi môn giả định | 3 |
| Tập điểm chữ được xét | A, B+, B, C+, C |
| Tiêu chí chọn | Trung bình ≥ GPA cần đạt và sát ngưỡng nhất |

### 2.5 M3 – Xếp loại và xếp hạng học kỳ

**Định dạng file Excel đầu vào** (xem `Demo.png`):

| Cột bắt buộc | Kiểu |
|---|---|
| `GPA học kỳ` | Số, thang 4.0 |
| `Điểm rèn luyện` | Số, thang 100 |
| `Số tín chỉ đăng ký` | Số nguyên |

**Ngưỡng xếp loại:**

> [![Xuất sắc](https://img.shields.io/badge/Mức_Xuất_Sắc-GPA_%E2%89%A5_3.6_%20%7C%20ĐRL_%E2%89%A5_90-800000?style=for-the-badge&labelColor=333333)](#)
> [![Giỏi](https://img.shields.io/badge/Mức_Giỏi-GPA_%E2%89%A5_3.2_%20%7C%20ĐRL_%E2%89%A5_80-A30000?style=for-the-badge&labelColor=333333)](#)
> [![Khá](https://img.shields.io/badge/Mức_Khá-GPA_%E2%89%A5_2.5_%20%7C%20ĐRL_%E2%89%A5_70-C62828?style=for-the-badge&labelColor=333333)](#)
> [![Trung bình](https://img.shields.io/badge/Mức_Trung_Bình-GPA_%E2%89%A5_2.0_%20%7C%20ĐRL_%E2%89%A5_60-E53935?style=for-the-badge&labelColor=333333)](#)
> [![Yếu](https://img.shields.io/badge/Mức_Yếu-GPA_%E2%89%A5_1.0_%20%7C%20ĐRL_%E2%89%A5_50-EF5350?style=for-the-badge&labelColor=333333)](#)
> [![Không xét](https://img.shields.io/badge/Mức_Không_Xét-GPA_<_1.0_%20%7C%20ĐRL_<_50-EF9A9A?style=for-the-badge&labelColor=333333)](#)

**Xếp loại cuối** là mức thấp hơn giữa hai tiêu chí:

$$
\text{Xếp loại} = \min(\text{Loại}_{GPA},\ \text{Loại}_{ĐRL})
$$

```mermaid
flowchart LR
    %% Định nghĩa dải màu sắc chuyên nghiệp (Paper-Ready Academic Style)
    classDef dataNode fill:#E3F2FD,stroke:#1565C0,stroke-width:2px,color:#0D47A1,font-weight:bold;
    classDef processNode fill:#E8F5E9,stroke:#2E7D32,stroke-width:1.5px,color:#1B5E20,font-weight:bold;
    classDef evalNode fill:#FFF3E0,stroke:#EF6C00,stroke-width:2px,color:#E65100,font-weight:bold;
    classDef sortNode fill:#F3E5F5,stroke:#7B1FA2,stroke-width:1.5px,color:#4A148C,font-weight:bold;
    classDef outputNode fill:#E0F2F1,stroke:#00695C,stroke-width:2px,color:#004D40,font-weight:bold;

    %% Các nút trong sơ đồ
    A["Đọc từng dòng<br/>dữ liệu Excel"]:::dataNode
    
    B["Xếp loại theo GPA<br/>(Điểm học tập)"]:::processNode
    C["Xếp loại theo ĐRL<br/>(Điểm rèn luyện)"]:::processNode
    
    D["Đánh giá chung:<br/>min(GPA, ĐRL)"]:::evalNode
    
    E["Sắp xếp giảm dần theo:<br/>1. Xếp loại cuối cùng<br/>2. GPA học kỳ<br/>3. Điểm rèn luyện<br/>4. Số tín chỉ đăng ký"]:::sortNode
    
    F["Xuất Bảng xếp hạng<br/>Học kỳ hoàn chỉnh"]:::outputNode

    %% Luồng liên kết
    A --> B
    A --> C
    
    B --> D
    C --> D
    
    D --> E
    E --> F
```

---

## 3. Đóng góp

```mermaid
graph LR
    subgraph K["Khoảng trống thực tế"]
        K1["GPA chỉ là con số,<br/>khó kiểm chứng"]
        K2["Mục tiêu tốt nghiệp<br/>thiếu định lượng"]
        K3["Xếp loại danh sách<br/>thủ công, dễ lệch"]
        K4["Dữ liệu điểm<br/>mang tính cá nhân"]
    end

    subgraph H["Hướng tiếp cận"]
        H1["Phơi bày toàn bộ<br/>bước trung gian"]
        H2["Giải bài toán nghịch +<br/>kiểm tra khả thi"]
        H3["Quy tắc min +<br/>sắp xếp xác định"]
        H4["Xử lý hoàn toàn<br/>phía Client"]
    end

    subgraph G["Giá trị mang lại"]
        G1["Tự kiểm chứng,<br/>hiểu tác động từng môn"]
        G2["Lập kế hoạch học tập<br/>dựa trên số liệu"]
        G3["Kết quả nhất quán,<br/>tái lập được"]
        G4["Riêng tư, triển khai<br/>không tốn hạ tầng"]
    end

    K1 --> H1 --> G1
    K2 --> H2 --> G2
    K3 --> H3 --> G3
    K4 --> H4 --> G4

    %% Style màu sắc ma trận
    style K fill:#FFEBEE,stroke:#C62828,color:#B71C1C
    style H fill:#E3F2FD,stroke:#1565C0,color:#0D47A1
    style G fill:#E8F5E9,stroke:#2E7D32,color:#1B5E20
```

\- Từ kết quả sang sự hiểu: minh bạch hóa quá trình tính GPA: Người học thường tiếp nhận GPA như một kết quả đóng, nên khó nhận ra môn nào kéo điểm xuống hay vì sao một môn điểm F lại làm thay đổi mẫu số. M1 không chỉ xuất ra GPA mà trình bày chuỗi quy đổi điểm hệ 10, điểm chữ, hệ 4.0, tích với tín chỉ và tổng tín chỉ hiệu dụng, kèm công thức. Nhờ đó công cụ đóng vai trò *thiết bị kiểm chứng*: sinh viên đối chiếu được con số với cổng thông tin và hình dung được môn nào đáng ưu tiên cải thiện.

\- Từ mô tả sang hoạch định: đảo chiều bài toán GPA: Hầu hết công cụ tính điểm chỉ trả lời "đã đạt bao nhiêu". M2 trả lời câu hỏi mang tính quyết định hơn: "cần đạt bao nhiêu". Bằng cách giải ngược GPA cần đạt trên số tín chỉ còn lại, công cụ cho biết ngay mục tiêu có khả thi hay không (vượt 4.0 thì cảnh báo), tránh việc sinh viên theo đuổi một kỳ vọng không thể đạt. Với mục tiêu khả thi, bước tìm kiếm tổ hợp chuyển con số trừu tượng thành kế hoạch cụ thể, ví dụ cần bao nhiêu môn A, B+, B để đạt ngưỡng. Đây là bước đưa kết quả từ "đánh giá" sang "hành động".

\- Từ thao tác thủ công sang quy trình nhất quán: chuẩn hóa xếp loại hàng loạt: Việc xếp loại cho cả lớp theo hai tiêu chí độc lập (GPA và điểm rèn luyện) dễ phát sinh sai lệch khi làm tay, nhất là khi phải phân xử các trường hợp bằng điểm. M3 mã hóa quy tắc lấy mức thấp hơn giữa hai tiêu chí và áp dụng thứ tự ưu tiên bốn khóa (xếp loại, GPA, điểm rèn luyện, số tín chỉ đăng ký). Vì cùng đầu vào luôn cho cùng thứ hạng, kết quả có tính tái lập và có thể đối chiếu, phù hợp cho cán bộ lớp phục vụ xét học bổng hoặc khen thưởng.

\- Giảm rào cản triển khai và bảo vệ dữ liệu: Do toàn bộ tính toán diễn ra phía client, dữ liệu điểm và file danh sách không rời khỏi thiết bị người dùng, đồng thời hệ thống không cần máy chủ hay cơ sở dữ liệu. Các thang điểm và ngưỡng được tách thành bảng hằng số trong `web.js`, nên việc điều chỉnh theo quy chế mới chỉ là thay đổi dữ liệu thay vì viết lại logic.

---

## 4. Kết luận

Dự án chứng minh rằng ba nhu cầu phân tích GPA điển hình của sinh viên (kiểm chứng, hoạch định, xếp loại hàng loạt) có thể được đáp ứng bằng một ứng dụng web nhẹ, không phụ thuộc hạ tầng và không đưa dữ liệu ra khỏi máy người dùng. Điểm khác biệt nằm ở việc chuyển công cụ từ *máy tính điểm* sang *công cụ hỗ trợ ra quyết định*.

**Hạn chế và hướng phát triển**

| Hạn chế hiện tại | Hướng khắc phục |
|---|---|
| M2 giả định mọi môn 3 tín chỉ và chỉ xét điểm A đến C | Cho phép nhập tín chỉ từng môn, mở rộng tập điểm chữ |
| Backtracking tăng nhanh theo số môn còn lại | Thay bằng quy hoạch động hoặc cắt nhánh |
| Chưa xử lý ngoại lệ (tất cả môn F, tín chỉ còn lại ≤ 0, thiếu cột Excel) | Bổ sung kiểm tra đầu vào và thông báo lỗi |
| Bảng hướng dẫn ghi mức Khá của điểm rèn luyện là 65–79, trong khi mã dùng ngưỡng ≥ 70 | Thống nhất giữa tài liệu và mã nguồn |
| Thang điểm, ngưỡng được cài cứng | Đối chiếu quy chế hiện hành, đưa ra file cấu hình |
| Chưa có kiểm thử tự động và đánh giá thực nghiệm | Bổ sung bộ test cho các ca biên |

---

<sub>Tác giả: huynhthanhphong231004IT · Đại học Cần Thơ</sub>
