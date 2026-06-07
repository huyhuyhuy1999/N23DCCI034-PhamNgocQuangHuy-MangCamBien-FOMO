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

## Lưu ý bảo mật

- Không commit API key Edge Impulse, token GitHub hoặc file `.env`.
- Không đưa ảnh webcam thô/chưa che thông tin riêng tư lên repo.
- Chỉ dùng screenshot đã redact khi cần minh chứng browser realtime.
