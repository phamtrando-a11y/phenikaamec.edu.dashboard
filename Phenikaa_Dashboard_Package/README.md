# Hướng dẫn Cài đặt và Chạy ứng dụng (Dashboard CME)

Chào bạn, đây là hướng dẫn để thiết lập lại ứng dụng này ở một máy tính hoặc môi trường mới.

## 1. Yêu cầu hệ thống
- Máy tính đã cài đặt Python (phiên bản 3.9 trở lên). Bạn có thể tải tại [python.org](https://www.python.org/downloads/).

## 2. Các bước cài đặt

**Bước 1: Mở Terminal (MacOS/Linux) hoặc Command Prompt / PowerShell (Windows)** tại đúng thư mục chứa source code này.

**Bước 2: Tạo và kích hoạt môi trường ảo (Virtual Environment)**
Việc này giúp các thư viện của dự án không bị xung đột với các ứng dụng Python khác trên máy.
- Trên MacOS/Linux:
  ```bash
  python3 -m venv .venv
  source .venv/bin/activate
  ```
- Trên Windows:
  ```bash
  python -m venv .venv
  .venv\Scripts\activate
  ```

**Bước 3: Cài đặt các thư viện phụ thuộc**
Đảm bảo bạn vẫn đang ở trong môi trường ảo (có chữ `(.venv)` ở đầu dòng lệnh), chạy lệnh sau:
```bash
pip install -r requirements.txt
```

## 3. Khởi động ứng dụng
Sau khi cài đặt xong, bạn có thể chạy ứng dụng bằng lệnh:
```bash
streamlit run app.py
```
Trình duyệt sẽ tự động mở lên tại địa chỉ `http://localhost:8501`.

## 4. Đối với phiên bản Netlify
Nếu bạn muốn đưa lên Netlify, toàn bộ mã nguồn tĩnh đã được đặt trong thư mục `netlify-dashboard`. Bạn chỉ cần nén lại hoặc kéo thả trực tiếp thư mục `netlify-dashboard` này vào giao diện deploy của Netlify là xong.
