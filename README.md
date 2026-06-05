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
- `data/synthetic_bottle_dataset_edge_impulse.zip`: dataset prototype/synthetic ban đầu.
- `data/real_web_bottle_dataset_ingestion_format.zip`: dataset real-web seed đã dùng để retrain trên Edge Impulse, gồm 90 training / 30 testing.
- `data/real_web_bottle_source_manifest.csv`: nguồn ảnh thật, tác giả và license.
- `data/real_web_bottle_attribution.md`: attribution rút gọn cho ảnh nguồn.
- `data/real_web_dataset_preview.jpg`: preview một số ảnh testing real-web sau augmentation.
- `screenshots/`: ảnh minh chứng từ Edge Impulse project, gồm dataset, training, model testing, deployment build, browser realtime và kết quả retrain real-web.
- `obs_test_assets/`: bộ ảnh 16:9 và slideshow HTML để test realtime browser qua OBS Virtual Camera.
- `scripts/make_synthetic_bottle_dataset.py`: tạo dataset prototype 3 class.
- `scripts/upload_synthetic_dataset_to_edge_impulse.py`: upload dataset bằng Edge Impulse ingestion API, đọc API key từ biến môi trường `EI_API_KEY`.
- `scripts/edge_impulse_upload_bridge.py`: bridge hỗ trợ upload từ browser/session.

## Kết quả hiện tại sau khi retrain bằng dữ liệu thật từ web

- Dataset trong Edge Impulse: 120 ảnh real-web seed, 90 training / 30 testing.
- Nguồn ảnh seed: 15 ảnh thật từ Wikimedia Commons, có metadata/license trong `data/real_web_bottle_source_manifest.csv`.
- Nhãn: `500ml`, `1L`, `1.5L`.
- Mô hình: FOMO MobileNetV2 0.35, 40 cycles, data augmentation bật.
- Validation F1: 75.0%.
- Validation precision / recall / F1 non-background: 0.92 / 0.63 / 0.75.
- Test accuracy: 56.67%.
- Test precision / recall / F1 non-background: 0.72 / 0.74 / 0.73.
- Browser realtime sau retrain: client đã build lại project và dừng ở màn hình yêu cầu cấp quyền camera; khi demo, chọn `OBS Virtual Camera` hoặc webcam thật.

Kết quả real-web thấp hơn dataset synthetic cũ vì ảnh web ít, khác nền, khác kiểu chai và có vài ảnh không hoàn toàn cùng domain với chai demo. Tuy vậy, đây là kết quả trung thực hơn để chuẩn bị nhận diện realtime bằng chai thật.

## Test realtime browser bằng OBS

Repo có sẵn bộ ảnh test tại `obs_test_assets/` để bạn đưa vào OBS và kiểm tra realtime browser của Edge Impulse.

1. Mở `obs_test_assets/slideshow.html` bằng trình duyệt.
2. Trong OBS, thêm nguồn `Window Capture` hoặc `Browser Source` trỏ tới slideshow này.
3. Bật `Start Virtual Camera` trong OBS.
4. Mở realtime browser/classifier của Edge Impulse và chọn camera là `OBS Virtual Camera`.
5. Quan sát nhãn dự đoán cho các class `500ml`, `1L`, `1.5L`.

Các ảnh OBS hiện là prototype/synthetic để kiểm tra pipeline realtime. Sau khi đã retrain bằng real-web seed, nên bổ sung thêm ảnh chai thật do bạn tự chụp bằng webcam/điện thoại để model bám tốt hơn vào điều kiện demo thực tế.

## Lưu ý bảo mật

- Không commit API key Edge Impulse, token GitHub hoặc file `.env`.
- Không đưa ảnh webcam thô/chưa che thông tin riêng tư lên repo.
- Chỉ dùng screenshot đã redact khi cần minh chứng browser realtime.
