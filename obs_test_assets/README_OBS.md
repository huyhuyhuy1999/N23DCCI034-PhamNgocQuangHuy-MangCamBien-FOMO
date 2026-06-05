# Hướng dẫn test realtime browser bằng OBS

Thư mục này chứa ảnh test 16:9 cho ba nhãn `500ml`, `1L`, `1.5L` và file `slideshow.html` để phát luân phiên các ảnh đó.

## Cách dùng nhanh

1. Mở file `slideshow.html` bằng trình duyệt.
2. Mở OBS, thêm nguồn `Window Capture` để bắt cửa sổ slideshow, hoặc thêm `Browser Source` và trỏ tới đường dẫn file `slideshow.html`.
3. Trong OBS, bấm `Start Virtual Camera`.
4. Mở realtime browser/classifier của Edge Impulse.
5. Chọn camera là `OBS Virtual Camera`.
6. Quan sát kết quả dự đoán theo từng ảnh.

## Phím tắt trong slideshow

- `Space`: dừng hoặc chạy tự động.
- `ArrowRight`: chuyển sang ảnh tiếp theo.
- `ArrowLeft`: quay lại ảnh trước.
- `F`: bật hoặc tắt fullscreen.

## Ghi chú

Các ảnh này được dựng từ dataset prototype của project để kiểm tra luồng realtime. Khi bảo vệ hoặc nghiệm thu thực tế, nên bổ sung test bằng chai thật trước camera với nhiều điều kiện ánh sáng và khoảng cách.
