# Báo cáo thống kê hợp đồng đối tác Referral

Hệ thống gửi email báo cáo tự động về tình trạng hợp đồng đối tác referral, tập trung vào các hợp đồng đã hết hạn và sắp hết hạn.

## Tính năng chính

- **Phân loại thông minh**: Tự động phân loại đối tác theo trạng thái hợp đồng
- **Sắp xếp ưu tiên**: Hiển thị theo mức độ khẩn cấp (hết hạn lâu nhất → sắp hết hạn nhất)
- **Báo cáo theo nhân viên**: Nhóm đối tác theo người phụ trách
- **Giao diện chuyên nghiệp**: Email HTML responsive với màu sắc phân biệt
- **Tự động hóa**: Hỗ trợ trigger gửi báo cáo định kỳ

## Cấu trúc dự án

```
├── ReferralMail.js     # Module báo cáo đối tác referral
├── Mail.js             # Module báo cáo chính
├── Update.js           # Module cập nhật dữ liệu
├── appsscript.json     # Cấu hình Apps Script
└── .clasp.json         # Cấu hình CLASP
```

## Cài đặt

### 1. Cấu hình Google Sheets

- Tạo Google Sheets với sheet tên "referral"
- Cập nhật `SPREADSHEET_ID` trong `ReferralMail.js`
- Đảm bảo các cột cần thiết:
  - Tên đối tác
  - Hạng mục
  - Nơi công tác
  - Loại hình hợp tác
  - Ngày hết hạn hợp đồng
  - Tên nhân viên
  - Tên chuyên khoa
  - Hiệu lực hợp đồng
  - Số ngày đến hạn
  - Trạng thái hợp đồng

### 2. Cấu hình email

```javascript
const CONFIG = {
  spreadsheetId: 'YOUR_SPREADSHEET_ID',
  sheetName: 'referral',
  emailRecipients: ['email1@domain.com', 'email2@domain.com'],
  expiringSoonDays: 30,
  debugMode: false
};
```

## Sử dụng

### Gửi báo cáo thử nghiệm

```javascript
// Mở Apps Script Editor
// Chạy function:
sendReferralReportTest();
```

### Gửi báo cáo thực tế

```javascript
sendReferralReport();
```

### Tạo trigger tự động

```javascript
// Gửi báo cáo hàng tuần vào thứ 2
createWeeklyReferralTrigger();
```

### Xóa trigger

```javascript
deleteReferralTriggers();
```

### Kiểm tra dữ liệu

```javascript
testReferralData();
```

## Cấu hình nâng cao

### Tùy chỉnh ngưỡng cảnh báo

```javascript
const CONFIG = {
  expiringSoonDays: 30, // Thay đổi số ngày cảnh báo
};
```

### Bật chế độ debug

```javascript
const CONFIG = {
  debugMode: true, // Hiển thị log chi tiết
};
```

## Cấu trúc email

### Tổng quan
- Số lượng hợp đồng đã hết hạn
- Số lượng hợp đồng sắp hết hạn
- Tổng số hợp đồng cần chú ý

### Chi tiết theo nhân viên
- Danh sách đối tác được nhóm theo người phụ trách
- Thông tin chi tiết: tên đối tác, hạng mục, nơi công tác, ngày hết hạn
- Màu sắc phân biệt theo trạng thái

## Màu sắc và trạng thái

- **Đỏ**: Hợp đồng đã hết hạn
- **Cam**: Hợp đồng sắp hết hạn (≤30 ngày)

## Troubleshooting

### Lỗi thường gặp

1. **Không tìm thấy sheet**: Kiểm tra tên sheet "referral"
2. **Lỗi quyền truy cập**: Đảm bảo Apps Script có quyền truy cập Sheets và Gmail
3. **Email không gửi được**: Kiểm tra danh sách email nhận

### Debug

```javascript
// Bật debug mode
const CONFIG = { debugMode: true };

// Kiểm tra log trong Apps Script Editor
// View > Logs
```

## Yêu cầu hệ thống

- Google Apps Script
- Google Sheets API
- Gmail API
- Quyền truy cập Google Workspace

## Phiên bản

- **v1.0**: Phiên bản cơ bản với báo cáo email
- **v1.1**: Thêm sắp xếp theo mức độ ưu tiên
- **v1.2**: Loại bỏ hợp đồng còn hiệu lực, tối ưu giao diện

## Hỗ trợ

Nếu gặp vấn đề, vui lòng kiểm tra:
1. Cấu hình SPREADSHEET_ID
2. Quyền truy cập Google APIs
3. Cấu trúc dữ liệu trong sheet
4. Logs trong Apps Script Editor