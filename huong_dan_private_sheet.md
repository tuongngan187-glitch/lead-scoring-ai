# HƯỚNG DẪN KẾT NỐI GOOGLE SHEETS Ở CHẾ ĐỘ PRIVATE (HẠN CHẾ)
Tài liệu này hướng dẫn chi tiết cách cấu hình ứng dụng Streamlit kết nối bảo mật với Google Sheets ở chế độ **Private (Hạn chế)** thông qua tài khoản dịch vụ **Google Cloud Service Account** và thư viện `st-gsheets-connection`.

---

## 📌 Tổng Quan Quy Trình
Quy trình thiết lập kết nối bảo mật gồm 4 bước chính:
```mermaid
graph TD
    A[1. Lấy chìa khóa JSON từ Google Cloud] --> B[2. Chia sẻ Google Sheet với email Service Account]
    B --> C[3. Cấu hình Secrets trên Streamlit Cloud / File cục bộ]
    C --> D[4. Cập nhật mã nguồn sử dụng st-gsheets-connection]
```

---

## 🔑 Bước 1: Lấy chìa khóa từ Google Cloud (Service Account)
Để ứng dụng Streamlit có thể truy cập Google Sheets mà không cần đăng nhập tài khoản cá nhân, ta cần tạo một **Service Account** (Tài khoản dịch vụ) đóng vai trò như một "Robot" truy cập dữ liệu:

1. **Truy cập Google Cloud Console**:
   * Truy cập vào [Google Cloud Console](https://console.cloud.google.com/).
   * Đăng nhập bằng tài khoản Google của bạn.
2. **Tạo Project mới (Nếu chưa có)**:
   * Bấm vào menu chọn dự án ở góc trên bên trái $\rightarrow$ Chọn **New Project** (Dự án mới).
   * Nhập tên dự án (ví dụ: `Lead-Scoring-AI`) $\rightarrow$ Bấm **Create**.
3. **Kích hoạt các API dịch vụ**:
   * Bấm vào thanh tìm kiếm ở trên cùng $\rightarrow$ Tìm kiếm **Google Sheets API** $\rightarrow$ Chọn Sheets API và bấm **Enable** (Kích hoạt).
   * Tiếp tục tìm kiếm **Google Drive API** $\rightarrow$ Bấm **Enable** (Kích hoạt) (cần thiết để cấp quyền truy cập tệp tin).
4. **Tạo tài khoản dịch vụ (Service Account)**:
   * Mở Menu góc trái $\rightarrow$ Chọn **IAM & Admin** $\rightarrow$ Chọn **Service Accounts**.
   * Bấm nút **Create Service Account** ở trên cùng.
   * Nhập tên tài khoản dịch vụ (ví dụ: `streamlit-sheets-bot`) $\rightarrow$ Bấm **Create and Continue**.
   * Phần gán quyền (Select a role): Bạn có thể bỏ qua $\rightarrow$ Bấm **Done** để hoàn tất.
5. **Tải file chìa khóa JSON**:
   * Tại danh sách Service Accounts vừa tạo, bấm chọn vào email của Service Account bạn mới tạo.
   * Chuyển sang tab **Keys** (Chìa khóa) ở phía trên.
   * Bấm **Add Key** $\rightarrow$ Chọn **Create new key** (Tạo khóa mới).
   * Chọn định dạng khóa là **JSON** $\rightarrow$ Bấm **Create**.
   * File khóa định dạng `.json` sẽ tự động được tải về máy của bạn. **Hãy bảo mật file này và không chia sẻ công khai.**

---

## 🤖 Bước 2: Kết nối Sheet với "Robot" (Chia sẻ quyền truy cập)
Hiện tại Google Sheet của bạn đang ở chế độ **Private (Hạn chế)**, chỉ tài khoản của bạn mới xem được. Ta cần cấp quyền cho tài khoản dịch vụ vừa tạo:

1. Mở file `.json` vừa tải về bằng bất kỳ công cụ soạn thảo văn bản nào (như Notepad, VS Code).
2. Tìm dòng chứa trường `"client_email"` và copy địa chỉ email đó. Địa chỉ email sẽ có dạng:
   `streamlit-sheets-bot@your-project-id.iam.gserviceaccount.com`
3. Mở bảng tính Google Sheet mà bạn muốn kết nối.
4. Bấm nút **Chia sẻ (Share)** ở góc trên bên phải bảng tính.
5. Dán địa chỉ email vừa copy vào ô thêm người dùng.
6. Thiết lập quyền hạn là **Viewer (Người xem)** (hoặc *Editor* nếu ứng dụng của bạn cần ghi thêm dữ liệu).
7. Bỏ chọn ô "Gửi thông báo cho mọi người" $\rightarrow$ Bấm **Chia sẻ (Share)** hoặc **Gửi**.

---

## 🔒 Bước 3: Cấu hình "Hộp đen bảo mật" (Secrets)
Bạn cần dán thông tin từ file `.json` vào mục quản lý mã bảo mật (Secrets) của Streamlit để ứng dụng đọc được khóa.

### Trường hợp 1: Chạy cục bộ dưới máy (Local Development)
1. Trong thư mục dự án của bạn, tạo một thư mục ẩn tên là `.streamlit` (nếu chưa có).
2. Tạo file tên là `secrets.toml` bên trong thư mục đó (đường dẫn: `.streamlit/secrets.toml`).
3. Chuyển đổi định dạng tệp tin từ `JSON` sang cấu trúc `TOML` của Streamlit như sau:

```toml
# .streamlit/secrets.toml
[connections.gsheets]
type = "service_account"
project_id = "tên-project-id-trong-file-json"
private_key_id = "mã-private-key-id-trong-file-json"
private_key = "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQC...\n-----END PRIVATE KEY-----\n"
client_email = "email-service-account-trong-file-json"
client_id = "mã-client-id-trong-file-json"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.google.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "đường-dẫn-cert-url-trong-file-json"
```

> [!IMPORTANT]
> Trường `private_key` phải giữ nguyên các ký tự xuống dòng dạng `\n` và được bọc trong dấu ngoặc kép `"`, hoặc viết ở dạng chuỗi nhiều dòng bằng ba dấu nháy kép `"""` trong TOML.

### Trường hợp 2: Khi triển khai trên Streamlit Community Cloud
1. Truy cập vào trang quản lý [Streamlit Share Dashboard](https://share.streamlit.io/).
2. Tìm đến ứng dụng của bạn $\rightarrow$ Bấm vào biểu tượng **3 chấm** ở góc bên phải $\rightarrow$ Chọn **Settings**.
3. Chọn thẻ **Secrets** ở thanh menu bên trái.
4. Dán toàn bộ nội dung cấu hình TOML bên trên (đã thay thế các thông số thật từ file `.json`) vào khung nhập liệu.
5. Bấm **Save** để lưu lại.

---

## 🛠️ Bước 4: Sửa Code để dùng Chìa khóa
Chúng ta sẽ sử dụng thư viện `st-gsheets-connection` của Streamlit để tự động nhận diện tài khoản dịch vụ trong Secrets và đọc dữ liệu từ Private Google Sheet.

### 1. Cập nhật thư viện
Thêm dòng sau vào file `requirements.txt` của dự án để Streamlit Cloud tự động cài đặt:
```text
st-gsheets-connection>=0.0.3
```

### 2. Sửa mã nguồn trong `streamlit_app.py`
Thay thế đoạn mã tải dữ liệu Google Sheets bằng kết nối của `st-gsheets-connection`:

```python
import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection

# 1. Khởi tạo kết nối Google Sheets bảo mật
# Connection sẽ tự động đọc credentials từ mục [connections.gsheets] trong Secrets
conn = st.connection("gsheets", type=GSheetsConnection)

# 2. Đọc dữ liệu từ Google Sheet
# Nhận vào link Google Sheets đầy đủ (ở chế độ Private)
sheet_url_input = st.text_input("Đường dẫn Google Sheets (Private)", value="https://docs.google.com/spreadsheets/d/your-spreadsheet-id/edit?usp=sharing")

if st.button("📥 Tải dữ liệu & Chấm điểm bảo mật", type="primary"):
    try:
        # Gọi phương thức đọc dữ liệu bảo mật bằng kết nối tài khoản dịch vụ
        df = conn.read(
            spreadsheet=sheet_url_input,
            ttl="5m"  # Thời gian cache dữ liệu để tránh gọi API liên tục (5 phút)
        )
        
        # Chuẩn hóa dữ liệu và tiến hành chấm điểm
        normalized = normalize_dataframe(df)
        
        for lead in normalized:
            res_score = rule_based_scorer(lead.get('requirement', ''))
            lead['score'] = res_score['score']
            lead['classification'] = res_score['classification']
            lead['reason'] = res_score['reason']
            
        st.session_state.leads = normalized
        st.session_state.processed_count = len(normalized)
        st.success(f"Đã đồng bộ bảo mật thành công {st.session_state.processed_count} khách hàng!")
        st.rerun()
        
    except Exception as e:
        st.error(f"Không thể kết nối đến Google Sheets bảo mật: {str(e)}")
        st.info("Mẹo: Hãy đảm bảo rằng bạn đã chia sẻ quyền Viewer của bảng tính cho email Service Account và cấu hình đúng Secrets.")
```

---
*Tài liệu hướng dẫn được thiết lập theo chuẩn tích hợp bảo mật tốt nhất của Streamlit. Nếu có bất kỳ thắc mắc nào trong quá trình thực hiện, vui lòng kiểm tra lại nhật ký lỗi (Log) của ứng dụng.*
