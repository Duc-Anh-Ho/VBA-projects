# Kế hoạch chuẩn hóa & thương mại hóa Danh-Tools

> Tài liệu này đánh giá tính khả thi của việc chuyển Danh-Tools sang VSTO để bán thương mại,
> so sánh các lựa chọn công nghệ, và đề ra lộ trình chuẩn hóa project.
> Người viết: phân tích dựa trên toàn bộ source hiện có (2026-05-30).

---

## 1. Hiện trạng project (tóm tắt sau khi đọc toàn bộ repo)

Repo `D:\repos\VBA-projects` đang chứa **nhiều thứ trộn lẫn**, không phải một sản phẩm duy nhất:

| Thư mục | Vai trò | Nên làm gì |
| :--- | :--- | :--- |
| `final-installation/` | Bản Danh-Tools "phát hành" (v2.3.4) — 35 file VBA | Bản cũ. Dùng làm tham chiếu ổn định. |
| `xlwings/DANH-project/` | **Bản đang phát triển, mới nhất** — 47 file VBA, có git riêng, docs riêng, kiến trúc tiến hóa (Utils_*, Interface_*, DI) | **Nguồn chính** để chuẩn hóa. |
| `part-tools/` | Các tool lẻ (auto-backup, auto-addin, invert-color, combine-sheets…) | Tách thành thư viện/ví dụ riêng. |
| `draft-and-reference/` | Bản cũ, file phá pass, backup | Đưa vào archive, **bỏ khỏi sản phẩm**. |
| `malware-investigation/` | File mẫu virus để nghiên cứu | **Tách hẳn ra repo khác** — không được nằm trong repo thương mại. |

### Vấn đề cấu trúc cần xử lý ngay
1. **Hai phiên bản song song** (`final-installation` và `xlwings/DANH-project`) → phải chốt 1 nguồn sự thật (single source of truth).
2. **Git repo lồng nhau**: `xlwings/DANH-project/.git` là repo độc lập, **không phải submodule, chưa có remote** → outer repo không quản được. Phải quyết: gộp (subtree/merge) hay tách thành submodule có remote.
3. **Binary trong git**: `.xlsb/.xlam/.frx/.zip/.rar` được commit trực tiếp → cần Git LFS hoặc chỉ track text export.
4. **Trộn sản phẩm với tool cá nhân**: các lớp `PJ1_*`, `Utils_Test` xử lý log UT (sfcmaplg/sfcmerlg, encoding `x-euc-jp`, đường dẫn `D:\share\ASS\Tasks\...`) rõ ràng là **tool công việc cá nhân**, không phải sản phẩm bán. Phải tách.
5. **Tính năng nhạy cảm**: auto-send wifi qua email (`Developer.autoSendWifi`, `EmailCDO`, `InternetConnector`) — chính README cảnh báo bị antivirus nhận diện là virus. **Bắt buộc loại bỏ** khỏi sản phẩm thương mại (rủi ro pháp lý + SmartScreen/AV).

### Điểm mạnh sẵn có (tài sản để tận dụng)
- Kiến trúc **controller-based** rõ ràng, đã có sẵn tài liệu API đầy đủ (`docs/README.md` trong DANH-project liệt kê từng method).
- Đã có **kế hoạch refactor SOLID/DI** (`ARCHITECTURE_REFACTORING_PLAN.md`, `REFACTORING_PLAN.md`), **coding conventions** (`CODING_CONVENTIONS.md`), và đã áp dụng thành công mẫu DI cho `Utils_Address`/`Interface_Address`.
- Đã có khung **test** (`Utils_Test`) và workflow edit VBA bằng text qua **xlwings CLI** (`xlwings vba edit -f <addin>`).

> 👉 Nghĩa là: phần "đặc tả sản phẩm" và "kiến trúc sạch" đã được làm khá tốt. Việc còn thiếu chủ yếu là **dọn repo + chốt công nghệ + thương mại hóa**.

---

## 2. Đánh giá: VSTO có khả quan không?

**Trả lời ngắn:** Khả thi về mặt kỹ thuật, **nhưng "chuyển sang VSTO" thực chất là viết lại từ đầu (VBA → C#/.NET)** — không có công cụ convert tự động. Và VSTO **chưa chắc là lựa chọn tốt nhất năm 2026**.

### 2.1. Bản đồ tính năng ↔ khả năng từng nền tảng

Tính năng hiện tại của Danh-Tools **phụ thuộc rất nặng vào tích hợp hệ thống Windows**. Đây là yếu tố quyết định:

| Tính năng | VBA (nay) | VSTO (.NET, Win desktop) | Excel-DNA (.NET XLL, Win) | Office.js (Web add-in, đa nền tảng) |
| :--- | :---: | :---: | :---: | :---: |
| Ribbon UI | ✅ | ✅ | ✅ | ✅ |
| Sheets/Charts/Pivot/Range (COM) | ✅ | ✅ | ✅ | ⚠️ một phần |
| Phím tắt toàn cục (OnKey) | ✅ | ✅ | ✅ | ❌ |
| Chụp màn hình qua ShareX (gọi .exe) | ✅ | ✅ | ✅ | ❌ |
| Chạy PowerShell | ✅ | ✅ | ✅ | ❌ |
| Đọc/ghi file tùy ý trên ổ đĩa | ✅ | ✅ | ✅ | ❌ (sandbox) |
| **Import/Export code VBA (VBE)** | ✅ | ✅ | ✅ | ❌ (không thể) |
| Win32 API (CopyMemory…) | ✅ | ✅ | ✅ | ❌ |
| Bán trên **AppSource** (marketplace MS) | – | ❌ | ❌ | ✅ |
| Chạy **Mac / Excel Online / iPad** | ❌ | ❌ | ❌ | ✅ |
| Triển khai (deploy) | copy .xlam | ClickOnce/MSI (nặng) | 1 file .xll (nhẹ) | manifest web (tập trung) |

**Kết luận từ bảng:**
- Nếu **giữ bộ tính năng hiện tại** (snip, PowerShell, import/export VBA, phím tắt toàn cục) → **bắt buộc ở lại Windows desktop .NET** (VSTO hoặc Excel-DNA). **Office.js không làm được** ~một nửa tính năng.
- Office.js chỉ hợp nếu **đa nền tảng là bắt buộc** và bạn **chấp nhận bỏ** các tính năng hệ thống — không khuyến nghị cho sản phẩm này ở dạng hiện tại.

### 2.2. VSTO vs Excel-DNA (cùng là .NET trên Windows desktop)

| Tiêu chí | VSTO | Excel-DNA ⭐ |
| :--- | :--- | :--- |
| Trạng thái MS | Maintenance mode (MS đẩy sang Office.js) | Mã nguồn mở, cộng đồng active |
| Deploy | ClickOnce/MSI, phụ thuộc .NET runtime, hay lỗi | **1 file .xll**, gọn, dễ bundle runtime |
| Tải sản phẩm thương mại nổi tiếng | Có | Rất nhiều add-in bán thương mại dùng Excel-DNA |
| Học/độ phức tạp | Cao hơn | Nhẹ hơn, tập trung vào ribbon + UDF + COM |
| Khi nào nên chọn | Cần document-level customization đặc thù của VSTO | **Hầu hết add-in COM + ribbon → chọn cái này** |

### 2.3. Khuyến nghị công nghệ

> **Đề xuất chính:** Viết lại trên **.NET (C#) cho Windows desktop**, và **ưu tiên Excel-DNA hơn VSTO** (deploy gọn hơn, hiện đại hơn, hợp gu add-in thương mại), trừ khi phát hiện nhu cầu đặc thù chỉ VSTO mới có.
>
> - **Office.js**: chỉ chọn nếu sau này muốn ra Mac/Web và chịu cắt tính năng hệ thống. Có thể để dành cho "phiên bản Lite" đa nền tảng sau này.
> - **VSTO**: dùng được, nhưng đừng mặc định chọn chỉ vì quen tên — cân nhắc Excel-DNA trước.

### 2.4. Tính khả thi thương mại
- Thị trường add-in Excel Windows trả phí **vẫn sống tốt** (Kutools, ASAP Utilities…). → Khả quan.
- Nhưng cần đầu tư **ngoài code**: chứng chỉ ký số (code signing — tránh cảnh báo SmartScreen/AV), hệ thống license/trial, installer + auto-update, trang bán hàng, hỗ trợ khách, chính sách quyền riêng tư (đặc biệt vì có chụp màn hình/đọc wifi/gọi shell).
- **Cảnh báo lớn:** mọi tính năng gọi PowerShell / đọc wifi / chụp màn hình rất dễ bị AV/SmartScreen gắn cờ. Phải xử lý kỹ (ký số, giải trình, bỏ tính năng email-wifi).

---

## 3. Lộ trình chuẩn hóa (phân theo giai đoạn)

Nguyên tắc: **dọn dẹp & chuẩn hóa trước (làm được ngay, giá trị bất kể chọn công nghệ nào), rồi mới prototype, rồi mới port dần.** Giữ bản VBA vẫn chạy trong suốt quá trình.

### Giai đoạn 0 — Quyết định & phạm vi (1–2 ngày)
- [ ] Chốt **nền tảng + công nghệ** dựa trên mục 2 (đề xuất: Excel-DNA/.NET Windows).
- [ ] Tách rõ **sản phẩm thương mại** = bộ tiện ích Danh-Tools (Sheets/Charts/Pivot/Ranges/Pictures/Shortcuts) **+** tính năng Import/Export VBA. Phần **PJ1_*/UT-log tách ra repo riêng cá nhân** (quyết định giữ/bỏ sau, nhưng không nằm trong sản phẩm bán).
- [ ] Định nghĩa SKU & mô hình giá (free/pro, trial, license vĩnh viễn vs subscription).

### Giai đoạn 1 — Dọn & chuẩn hóa repo (giá trị ngay, ít rủi ro)
- [ ] **Chốt 1 nguồn sự thật**: lấy `xlwings/DANH-project` làm bản chính; đánh dấu `final-installation` là legacy/tham chiếu.
- [ ] **Xử lý git lồng nhau**: chuyển `DANH-project` thành submodule có remote riêng, **hoặc** gộp lịch sử bằng `git subtree`. Không để repo-trong-repo không quản lý.
- [ ] **Tách rác khỏi sản phẩm**: đưa `malware-investigation/` và `draft-and-reference/` sang repo/archive riêng.
- [ ] Chuẩn hóa cây thư mục mục tiêu, ví dụ:
  ```
  /src        (code VBA text export — nguồn sự thật)
  /build      (workbook .xlsb/.xlam được build ra; gitignore hoặc LFS)
  /assets     (Images, CustomUI XML, icon)
  /docs       (plan, conventions, architecture, API)
  /tests      (test VBA / sau này test .NET)
  /tools      (script build/export/import)
  ```
- [ ] **Binary**: bật Git LFS cho `.xlsb/.xlam/.frx` hoặc chỉ commit text export + build tự động.
- [ ] **Loại bỏ tính năng nhạy cảm** (auto-send wifi/email) khỏi nhánh sản phẩm.
- [ ] Chuẩn hóa **workflow export/import VBA ↔ text** thành script (dựa trên xlwings CLI sẵn có) để versioning ổn định.

### Giai đoạn 2 — Hoàn thiện refactor & đặc tả trên VBA (nền cho rewrite)
> Hoàn thiện kiến trúc sạch trên VBA trước giúp việc port sang .NET map 1-1 dễ dàng, và bản VBA vẫn bán/chạy được trong lúc rewrite.
- [ ] Tiếp tục theo `REFACTORING_PLAN.md`: tách `SystemUpdate` (God class) → các service nhỏ; áp DI rộng rãi; `ConfigurationService` thay hardcode (đường dẫn ShareX, link ping…).
- [ ] Áp `CODING_CONVENTIONS.md` toàn bộ (`Left$`, fully-qualified calls, access modifier).
- [ ] Viết **test** cho các service (khung `Utils_Test` đã có) → có lưới an toàn trước khi port.
- [ ] **Đặc tả từng tính năng** thành spec ngắn (input/output/hành vi) — phần lớn đã có trong `docs/README.md`, bổ sung cho đủ. Đây là "hợp đồng" để viết lại trên .NET.

### Giai đoạn 3 — Prototype công nghệ đã chọn (spike, xác nhận khả thi)
- [ ] Dựng skeleton .NET (Excel-DNA hoặc VSTO): Ribbon + **2 tính năng đại diện** — 1 thuần COM (vd Sheets controller) + 1 tích hợp hệ thống (vd snip qua ShareX / chạy shell) — để chứng minh khả thi.
- [ ] Prototype **deployment**: installer, ký số (code signing cert), thử trên máy sạch (kiểm tra SmartScreen/AV).
- [ ] Prototype **license/trial** tối thiểu.

### Giai đoạn 4 — Port theo module (tăng dần)
- [ ] Map từng `*Controller` VBA → service .NET, giữ kiến trúc layered đã định (UI / Application / Domain-Service / Infrastructure).
- [ ] Ưu tiên port tính năng giá trị cao trước; mỗi module có test.
- [ ] Đảm bảo tương đương hành vi với bản VBA (dựa trên spec giai đoạn 2).

### Giai đoạn 5 — Thương mại hóa
- [ ] License + trial + activation; installer + auto-update; ký số chính thức.
- [ ] Trang bán hàng/landing, tài liệu người dùng, kênh hỗ trợ.
- [ ] **Chính sách quyền riêng tư + bảo mật** (bắt buộc vì có chụp màn hình/shell): nêu rõ tool làm gì, không gửi dữ liệu đi đâu.
- [ ] QA + chương trình beta.

---

## 4. Quyết định còn mở (cần bạn chốt)
1. **Công nghệ cuối cùng**: Excel-DNA (đề xuất) vs VSTO vs (sau này) Office.js Lite? 
2. **Xử lý git lồng nhau**: submodule hay gộp subtree?
3. **Mô hình giá**: một lần vs subscription; có bản free không?
4. **PJ1/UT-log**: tách repo cá nhân ngay (đề xuất) — đồng ý không?
5. **Phạm vi "bản 1.0 thương mại"**: bán nguyên bộ hay chỉ nhóm tính năng mạnh nhất (Sheets + Ranges + Pictures + Shortcuts)?

---

## 5. Khuyến nghị hành động tiếp theo (gọn)
1. Chốt mục 4.1 và 4.2.
2. Làm **Giai đoạn 1** (dọn repo) — giá trị ngay, không rủi ro, độc lập với lựa chọn công nghệ.
3. Làm **spike Excel-DNA** (Giai đoạn 3) song song để có số liệu thực tế về deploy/AV trước khi cam kết rewrite toàn bộ.
