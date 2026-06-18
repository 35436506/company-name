# Company Name Matcher

App Streamlit để so khớp tên công ty (fuzzy matching) giữa danh sách **Indue** (chuẩn so sánh) và một danh sách khác. Hỗ trợ upload:

- **1 file Excel có nhiều tab** (ví dụ tab `IndueD` và tab `KEV`), hoặc
- **2 file Excel riêng biệt**

App tự động nhận diện tab/file nào là "Indue" dựa trên tên (chứa chữ `indue`, không phân biệt hoa thường), tự dò cột chứa tên công ty, và xuất ra file kết quả gồm 3 tab: `Result`, `Source - <Indue>`, `Source - <khác>`.

## Chạy thử trên máy (local)

```bash
pip install -r requirements.txt
streamlit run app.py
```

Sau đó mở `http://localhost:8501` trên trình duyệt.

## Deploy lên Streamlit Community Cloud (miễn phí)

### Bước 1 — Đưa code lên GitHub

1. Tạo một repository mới trên GitHub (ví dụ tên `company-name-matcher`), để **Public** hoặc **Private** đều được.
2. Trên máy, vào thư mục chứa các file này (`app.py`, `requirements.txt`, `README.md`) và chạy:

```bash
git init
git add .
git commit -m "Initial commit: company name matcher app"
git branch -M main
git remote add origin https://github.com/<TÊN_GITHUB_CỦA_BẠN>/company-name-matcher.git
git push -u origin main
```

> Thay `<TÊN_GITHUB_CỦA_BẠN>` bằng username GitHub của bạn. Nếu chưa cấu hình Git, chạy thêm:
> `git config --global user.email "ban@example.com"` và `git config --global user.name "Tên Bạn"`.

Nếu bạn không quen dùng dòng lệnh Git, cách đơn giản hơn:
1. Vào [github.com/new](https://github.com/new), tạo repo mới, **không** tick "Add a README".
2. Trên trang repo vừa tạo, bấm **"uploading an existing file"**.
3. Kéo thả 3 file `app.py`, `requirements.txt`, `README.md` vào, rồi bấm **Commit changes**.

### Bước 2 — Deploy trên Streamlit Cloud

1. Truy cập [share.streamlit.io](https://share.streamlit.io) và đăng nhập bằng tài khoản GitHub.
2. Bấm **"New app"**.
3. Chọn:
   - **Repository**: repo bạn vừa tạo (ví dụ `<tên-bạn>/company-name-matcher`)
   - **Branch**: `main`
   - **Main file path**: `app.py`
4. Bấm **"Deploy!"**.

Sau khoảng 1–2 phút, Streamlit sẽ cấp cho bạn một link dạng:

```
https://<tên-app>-<random>.streamlit.app
```

Mỗi lần bạn push code mới lên GitHub, app trên Streamlit Cloud sẽ tự động cập nhật theo.

## Cấu trúc file

```
.
├── app.py              # Mã nguồn chính của app Streamlit
├── requirements.txt    # Các thư viện Python cần thiết
└── README.md            # Hướng dẫn này
```

## Lưu ý

- Ngưỡng điểm matching (mặc định 86/100) có thể điều chỉnh trực tiếp trên giao diện app bằng thanh trượt.
- Nếu app không tự nhận diện đúng tab/file "Indue" (ví dụ tên file/tab không chứa chữ "Indue"), bạn có thể chọn tay trong giao diện.
- File tải về không ghi đè file gốc — luôn là một file Excel mới.
