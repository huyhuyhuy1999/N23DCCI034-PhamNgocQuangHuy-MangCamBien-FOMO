# Phát hiện chai nhựa tái chế theo size bằng FOMO

Sinh viên: Phạm Ngọc Quang Huy  
MSSV: N23DCCI034  
Lớp: D23CQCI01-N  
Môn: Mạng Cảm Biến  

## Đề tài

Phát hiện chai nhựa tái chế theo size `500ml`, `1L`, `1.5L` bằng Edge Impulse FOMO/object detection.  
Project Edge Impulse: `https://studio.edgeimpulse.com/studio/1021052`

## Nội dung repo

- `report/PhamNgocQuangHuy_final_cuoiky.docx`: báo cáo Word cuối kỳ đã chỉnh theo project của đề tài.
- `data/synthetic_bottle_dataset_edge_impulse.zip`: dataset prototype đã đóng gói theo định dạng Edge Impulse.
- `screenshots/`: ảnh minh chứng từ Edge Impulse project, gồm data acquisition, feature generation, training, model testing, deployment build và browser realtime.
- `obs_test_assets/`: bộ ảnh 16:9 và slideshow HTML để test realtime browser qua OBS Virtual Camera.
- `scripts/make_synthetic_bottle_dataset.py`: tạo dataset prototype 3 class.
- `scripts/upload_synthetic_dataset_to_edge_impulse.py`: upload dataset bằng Edge Impulse ingestion API, đọc API key từ biến môi trường `EI_API_KEY`.
- `scripts/edge_impulse_upload_bridge.py`: bridge hỗ trợ upload từ browser/session.

## Kết quả thực nghiệm

- Dataset prototype: 120 ảnh, 90 training / 30 testing.
- Nhãn: `500ml`, `1L`, `1.5L`.
- Mô hình: FOMO MobileNetV2 0.35, 40 cycles.
- Validation F1: 86.5%.
- Test accuracy: 86.67%.
- Test precision / recall / F1 non-background: 0.90 / 0.93 / 0.92.
- Deployment: build thành công `v12 (C++ library)` lúc 21:40:07 ngày 05/06/2026.
- Browser realtime: client đã tải model và chạy trạng thái `Inferencing`; vùng camera trong screenshot đã được làm mờ để bảo vệ riêng tư.

## Test realtime browser bằng OBS

Repo có sẵn bộ ảnh test tại `obs_test_assets/` để bạn đưa vào OBS và kiểm tra realtime browser của Edge Impulse.

1. Mở `obs_test_assets/slideshow.html` bằng trình duyệt.
2. Trong OBS, thêm nguồn `Window Capture` hoặc `Browser Source` trỏ tới slideshow này.
3. Bật `Start Virtual Camera` trong OBS.
4. Mở realtime browser/classifier của Edge Impulse và chọn camera là `OBS Virtual Camera`.
5. Quan sát nhãn dự đoán cho các class `500ml`, `1L`, `1.5L`.

Các ảnh test là prototype/synthetic để kiểm tra pipeline realtime, không thay thế bộ ảnh nghiệm thu bằng chai thật ngoài hiện trường. Khi demo chính thức, nên test thêm với chai thật ở nhiều nền, ánh sáng và khoảng cách.

## Lưu ý bảo mật

- Không commit API key Edge Impulse, token GitHub hoặc file `.env`.
- Không đưa ảnh webcam thô/chưa che thông tin riêng tư lên repo.
- Chỉ dùng screenshot đã redact khi cần minh chứng browser realtime.
