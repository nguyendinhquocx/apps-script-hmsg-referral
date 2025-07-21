/**
 * HMSG Referral Partner Report - Email Notification System
 * 
 * Gửi email báo cáo đối tác theo nhân viên với phân loại:
 * - Đã hết hạn (màu đỏ)
 * - Sắp hết hạn (màu vàng) 
 * - Còn hiệu lực (màu xanh)
 * 
 * @version 1.0
 * @author Nguyen Dinh Quoc
 * @updated 2025-01-01
 */
function sendWeeklyReferralReport() {
  const CONFIG = {
    spreadsheetId: '14A_0CcRRJtfEKdWs93UsY6L6idDQCtIdZyFfTvt7Le8',
    sheetName: 'referral',
    
    // Email configuration
    emailTo: 'quoc.nguyen3@hoanmy.com', // Thay đổi khi deploy production
    // emailTo: 'luan.tran@hoanmy.com, khanh.tran@hoanmy.com, hong.le@hoanmy.com, quynh.bui@hoanmy.com, thuy.pham@hoanmy.com, anh.ngo@hoanmy.com, truc.nguyen3@hoanmy.com, trang.nguyen9@hoanmy.com, tram.mai@hoanmy.com, vuong.duong@hoanmy.com, phi.tran@hoanmy.com, quoc.nguyen3@hoanmy.com',
    
    // Thresholds for contract status
    expiringSoonDays: 30, // Sắp hết hạn trong 30 ngày
    
    // Icons
    expiredIcon: 'https://cdn-icons-png.flaticon.com/128/17694/17694317.png',
    expiringSoonIcon: 'https://cdn-icons-png.flaticon.com/128/7046/7046053.png', 
    activeIcon: 'https://cdn-icons-png.flaticon.com/128/10995/10995390.png',
    calendarIcon: 'https://cdn-icons-png.flaticon.com/128/3239/3239948.png',
    partnerIcon: 'https://cdn-icons-png.flaticon.com/128/3135/3135715.png',
    
    // Debug mode
    debugMode: false
  };

  try {
    // Mở spreadsheet
    const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    const sheet = ss.getSheetByName(CONFIG.sheetName);
    
    if (!sheet) {
      Logger.log(`❌ Sheet '${CONFIG.sheetName}' không tồn tại`);
      return;
    }

    // Lấy dữ liệu từ sheet (bỏ qua header row)
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];
    const data = values.slice(1);
    
    if (CONFIG.debugMode) {
      Logger.log(`📊 Tổng số dòng dữ liệu: ${data.length}`);
      Logger.log(`📋 Headers: ${headers.join(', ')}`);
    }

    // Tìm index các cột cần thiết
    const columnIndexes = {
      tenDoiTac: headers.indexOf('ten doi tac'),
      hangMuc: headers.indexOf('hang muc'),
      noiCongTac: headers.indexOf('noi cong tac'),
      loaiHinhHopTac: headers.indexOf('loai hinh hop tac'),
      ngayHopTac: headers.indexOf('ngay hop tac'),
      tenNhanVien: headers.indexOf('ten nhan vien'),
      tenChuyenKhoa: headers.indexOf('ten chuyen khoa'),
      hieuLucHd: headers.indexOf('hieu luc hd'),
      ngayHetHanHd: headers.indexOf('ngay het han hd'),
      soNgayDenHanHd: headers.indexOf('so ngay den han hd'),
      trangThaiHd: headers.indexOf('trang thai hd'),
      id: headers.indexOf('ID')
    };

    // Kiểm tra các cột bắt buộc
    const requiredColumns = ['tenDoiTac', 'tenNhanVien', 'soNgayDenHanHd', 'trangThaiHd'];
    for (let col of requiredColumns) {
      if (columnIndexes[col] === -1) {
        Logger.log(`❌ Không tìm thấy cột: ${col}`);
        return;
      }
    }

    // Xử lý dữ liệu và nhóm theo nhân viên
    const employeeData = {};
    const today = new Date();
    
    data.forEach((row, index) => {
      try {
        const tenNhanVien = row[columnIndexes.tenNhanVien];
        const tenDoiTac = row[columnIndexes.tenDoiTac];
        const soNgayDenHan = row[columnIndexes.soNgayDenHanHd];
        const trangThaiHd = row[columnIndexes.trangThaiHd];
        
        if (!tenNhanVien || !tenDoiTac) return;
        
        // Chỉ xử lý hợp đồng đã hết hạn hoặc sắp hết hạn
        let contractStatus = null;
        if (soNgayDenHan < 0 || trangThaiHd === 'Đã hết hạn') {
          contractStatus = 'expired';
        } else if (soNgayDenHan <= CONFIG.expiringSoonDays) {
          contractStatus = 'expiring_soon';
        }
        
        // Bỏ qua hợp đồng còn hiệu lực
        if (!contractStatus) return;
        
        // Khởi tạo dữ liệu nhân viên nếu chưa có
        if (!employeeData[tenNhanVien]) {
          employeeData[tenNhanVien] = {
            expired: [],
            expiring_soon: []
          };
        }
        
        // Thêm đối tác vào danh sách tương ứng
        const partnerInfo = {
          tenDoiTac: tenDoiTac,
          hangMuc: row[columnIndexes.hangMuc] || '',
          noiCongTac: row[columnIndexes.noiCongTac] || '',
          loaiHinhHopTac: row[columnIndexes.loaiHinhHopTac] || '',
          tenChuyenKhoa: row[columnIndexes.tenChuyenKhoa] || '',
          ngayHetHanHd: row[columnIndexes.ngayHetHanHd] || '',
          soNgayDenHan: soNgayDenHan,
          trangThaiHd: trangThaiHd,
          id: row[columnIndexes.id] || ''
        };
        
        employeeData[tenNhanVien][contractStatus].push(partnerInfo);
        
      } catch (error) {
        Logger.log(`⚠️ Lỗi xử lý dòng ${index + 2}: ${error.message}`);
      }
    });

    // Sắp xếp đối tác theo số ngày hết hạn (từ hết hạn nhiều nhất đến ngắn nhất)
    Object.keys(employeeData).forEach(empName => {
      const empData = employeeData[empName];
      // Sắp xếp expired: số âm lớn nhất trước (hết hạn lâu nhất)
      empData.expired.sort((a, b) => a.soNgayDenHan - b.soNgayDenHan);
      // Sắp xếp expiring_soon: số nhỏ nhất trước (sắp hết hạn nhất)
      empData.expiring_soon.sort((a, b) => a.soNgayDenHan - b.soNgayDenHan);
    });
    
    if (CONFIG.debugMode) {
      Logger.log(`👥 Số nhân viên: ${Object.keys(employeeData).length}`);
      Object.keys(employeeData).forEach(emp => {
        const data = employeeData[emp];
        Logger.log(`${emp}: Hết hạn(${data.expired.length}), Sắp hết hạn(${data.expiring_soon.length})`);
      });
    }

    // Tạo email content
    const emailContent = buildEmailContent(employeeData, CONFIG, today);
    
    // Gửi email
    const subject = `HMSG | P.KD - BÁO CÁO THỐNG KÊ HỢP ĐỒNG ĐỐI TÁC REFERRAL TÁC TUẦN ${Utilities.formatDate(today, ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy')}`;
    
    sendEmailWithRetry({
      to: CONFIG.emailTo,
      subject: subject,
      htmlBody: emailContent
    });

    Logger.log(`✅ Email báo cáo đối tác đã được gửi thành công`);

  } catch (error) {
    Logger.log(`❌ Lỗi khi gửi email báo cáo đối tác: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
  }
}

/**
 * Xây dựng nội dung email HTML
 */
function buildEmailContent(employeeData, CONFIG, today) {
  const dayNames = ['Chủ nhật', 'Thứ hai', 'Thứ ba', 'Thứ tư', 'Thứ năm', 'Thứ sáu', 'Thứ bảy'];
  const dayOfWeek = dayNames[today.getDay()];
  const detailedDate = `${dayOfWeek}, ngày ${today.getDate()} tháng ${today.getMonth() + 1} năm ${today.getFullYear()}`;
  
  // Tính tổng số đối tác theo trạng thái
  let totalExpired = 0, totalExpiringSoon = 0;
  Object.values(employeeData).forEach(emp => {
    totalExpired += emp.expired.length;
    totalExpiringSoon += emp.expiring_soon.length;
  });
  
  const totalPartners = totalExpired + totalExpiringSoon;
  
  // Color scheme
  const colors = {
    border: '#000000',
    headerTitle: '#1a1a1a',
    headerSubtitle: '#8e8e93',
    dateText: '#495057',
    sectionTitle: '#1a1a1a',
    expiredTitle: '#dc3545',
    expiringSoonTitle: '#f59e0b',
    activeTitle: '#22c55e',
    namesList: '#1a1a1a',
    footerName: '#8e8e93',
    footerLabel: '#1a1a1a',
    disclaimerColor: '#8e8e93'
  };
  
  // Badge styles
  const getBadgeStyle = (type) => {
    switch(type) {
      case 'expired': return 'background: #fef2f2; color: #dc2626; border: 1px solid #fecaca;';
      case 'expiring_soon': return 'background: #fffbeb; color: #d97706; border: 1px solid #fed7aa;';
      case 'active': return 'background: #f0fdf4; color: #16a34a; border: 1px solid #bbf7d0;';
      default: return 'background: #f9fafb; color: #4b5563; border: 1px solid #e5e7eb;';
    }
  };
  
  // Xây dựng summary dashboard
  const summaryDashboard = `
    <div style="margin-bottom: 32px; background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 12px; padding: 24px;">
      <h2 style="margin: 0 0 20px; font-size: 18px; font-weight: 500; color: ${colors.sectionTitle};">
        Tổng quan đối tác
      </h2>
      <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 16px;">
        <div style="text-align: center; padding: 16px; background: white; border-radius: 8px; border: 1px solid #e2e8f0;">
          <div style="font-size: 24px; font-weight: 600; color: ${colors.expiredTitle}; margin-bottom: 4px;">${totalExpired}</div>
          <div style="font-size: 12px; color: ${colors.expiredTitle}; font-weight: 500;">Đã hết hạn</div>
        </div>
        <div style="text-align: center; padding: 16px; background: white; border-radius: 8px; border: 1px solid #e2e8f0;">
          <div style="font-size: 24px; font-weight: 600; color: ${colors.expiringSoonTitle}; margin-bottom: 4px;">${totalExpiringSoon}</div>
          <div style="font-size: 12px; color: ${colors.expiringSoonTitle}; font-weight: 500;">Sắp hết hạn</div>
        </div>
        <div style="text-align: center; padding: 16px; background: white; border-radius: 8px; border: 1px solid #e2e8f0;">
          <div style="font-size: 24px; font-weight: 600; color: ${colors.sectionTitle}; margin-bottom: 4px;">${totalPartners}</div>
          <div style="font-size: 12px; color: ${colors.sectionTitle}; font-weight: 500;">Tổng cộng</div>
        </div>
      </div>
    </div>
  `;
  
  // Xây dựng danh sách theo nhân viên
  let employeeSections = '';
  const sortedEmployees = Object.keys(employeeData).sort();
  
  sortedEmployees.forEach(employeeName => {
    const empData = employeeData[employeeName];
    const empTotal = empData.expired.length + empData.expiring_soon.length;
    
    if (empTotal === 0) return; // Bỏ qua nhân viên không có đối tác
    
    employeeSections += `
      <div style="margin-bottom: 32px; background-color: #ffffff; border: 1px solid #e9ecef; border-radius: 12px; overflow: hidden;">
        <div style="padding: 20px 24px; background-color: #f8fafc; border-bottom: 1px solid #e2e8f0;">
          <h3 style="margin: 0; font-size: 16px; font-weight: 500; color: ${colors.sectionTitle};">${employeeName}</h3>
          <div style="margin-top: 8px; display: flex; gap: 8px; flex-wrap: wrap;">
            ${empData.expired.length > 0 ? `<span style="${getBadgeStyle('expired')} padding: 4px 8px; border-radius: 8px; font-size: 11px; font-weight: 500;">Hết hạn: ${empData.expired.length}</span>` : ''}
            ${empData.expiring_soon.length > 0 ? `<span style="${getBadgeStyle('expiring_soon')} padding: 4px 8px; border-radius: 8px; font-size: 11px; font-weight: 500;">Sắp hết hạn: ${empData.expiring_soon.length}</span>` : ''}
          </div>
        </div>
    `;
    
    // Đã hết hạn
    if (empData.expired.length > 0) {
      employeeSections += buildPartnerSection('Đã hết hạn', empData.expired, CONFIG.expiredIcon, colors.expiredTitle, 'expired');
    }
    
    // Sắp hết hạn
    if (empData.expiring_soon.length > 0) {
      employeeSections += buildPartnerSection('Sắp hết hạn', empData.expiring_soon, CONFIG.expiringSoonIcon, colors.expiringSoonTitle, 'expiring_soon');
    }
    
    employeeSections += '</div>';
  });
  
  // HTML Email Template
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Báo cáo đối tác tuần ${Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yyyy')}</title>
    </head>
    <body style="margin: 0; padding: 0; background-color: #ffffff; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;">
      
      <!-- Main Container -->
      <div style="max-width: 800px; margin: 40px auto; padding: 40px; border: 1px solid ${colors.border}; border-radius: 12px;">
        
        <!-- Header -->
        <div style="text-align: center; margin-bottom: 48px;">
          <h1 style="margin: 0; font-size: 28px; font-weight: 300; color: ${colors.headerTitle}; letter-spacing: -0.5px;">
            Báo cáo thống kê hợp đồng đối tác Referral
          </h1>
          <p style="margin: 8px 0 0; font-size: 16px; font-weight: 400; color: ${colors.headerSubtitle};">
            Phòng Kinh Doanh HMSG
          </p>
        </div>

        <!-- Date -->
        <div style="margin-bottom: 32px;">
          <span style="font-size: 14px; font-weight: 500; color: ${colors.dateText};">
            <img src="${CONFIG.calendarIcon}" width="16" height="16" style="vertical-align: middle; margin-right: 8px;" alt="Calendar">
            ${detailedDate}
          </span>
        </div>

        <!-- Summary Dashboard -->
        ${summaryDashboard}

        <!-- Employee Sections -->
        ${employeeSections}

        <!-- Footer -->
        <div style="text-align: center; padding-top: 32px; border-top: 1px solid #f5f5f5;">
          <p style="margin: 0 0 6px; font-size: 12px; font-weight: 400; color: ${colors.footerLabel};">
            Trân trọng
          </p>
          <p style="margin: 0; font-size: 12px; font-weight: 500; color: ${colors.footerName};">
            Nguyen Dinh Quoc
          </p>
        </div>

        <!-- Disclaimer -->
        <div style="margin-top: 40px; text-align: center;">
          <p style="margin: 0; font-size: 12px; color: ${colors.disclaimerColor}; line-height: 1.4; font-style: italic;">
            Báo cáo đối tác tự động hàng tuần. Vui lòng không trả lời email này.<br>
            Liên hệ: quoc.nguyen3@hoanmy.com
          </p>
        </div>

      </div>
      
    </body>
    </html>
  `;
}

/**
 * Xây dựng section cho từng loại đối tác
 */
function buildPartnerSection(title, partners, icon, titleColor, type) {
  const formatDate = (dateValue) => {
    if (!dateValue) return 'N/A';
    if (dateValue instanceof Date) {
      return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }
    return dateValue.toString();
  };
  
  const formatDaysRemaining = (days) => {
    if (days < 0) return `Quá hạn ${Math.abs(days)} ngày`;
    if (days === 0) return 'Hết hạn hôm nay';
    return `Còn ${days} ngày`;
  };
  
  const getUrgencyColor = (days, type) => {
    if (type === 'expired') return '#dc3545';
    if (type === 'expiring_soon') {
      if (days <= 7) return '#dc3545';
      if (days <= 15) return '#f59e0b';
      return '#f59e0b';
    }
    return '#22c55e';
  };
  
  let partnersHtml = partners.map(partner => {
    const urgencyColor = getUrgencyColor(partner.soNgayDenHan, type);
    
    return `
      <div style="padding: 16px 0; border-bottom: 1px solid #f5f5f5; display: flex; justify-content: space-between; align-items: flex-start; gap: 16px;">
        <div style="flex: 1; min-width: 0;">
          <div style="font-size: 15px; font-weight: 500; color: #1a1a1a; margin-bottom: 4px;">${partner.tenDoiTac}</div>
          <div style="font-size: 13px; color: #6b7280; margin-bottom: 2px;">
            <strong>Nơi CT:</strong> ${partner.noiCongTac}
          </div>
          <div style="font-size: 13px; color: #6b7280; margin-bottom: 2px;">
            <strong>Loại HT:</strong> ${partner.loaiHinhHopTac}
          </div>
          <div style="font-size: 13px; color: #6b7280;">
            <strong>Chuyên khoa:</strong> ${partner.tenChuyenKhoa}
          </div>
        </div>
        <div style="text-align: right; flex-shrink: 0;">
          <div style="font-size: 12px; color: ${urgencyColor}; font-weight: 500; margin-bottom: 2px;">
            ${formatDaysRemaining(partner.soNgayDenHan)}
          </div>
          <div style="font-size: 11px; color: #8e8e93;">
            HH: ${formatDate(partner.ngayHetHanHd)}
          </div>
          ${partner.id ? `<div style="font-size: 10px; color: #a1a1aa; margin-top: 2px;">${partner.id}</div>` : ''}
        </div>
      </div>
    `;
  }).join('');
  
  return `
    <div style="padding: 0 24px;">
      <div style="padding: 16px 0 8px; border-bottom: 1px solid #e2e8f0; margin-bottom: 8px;">
        <h4 style="margin: 0; font-size: 14px; font-weight: 500; color: ${titleColor}; display: flex; align-items: center;">
          <img src="${icon}" width="16" height="16" style="margin-right: 8px;" alt="${title}">
          ${title} (${partners.length})
        </h4>
      </div>
      ${partnersHtml}
    </div>
  `;
}

/**
 * Gửi email với retry mechanism
 */
function sendEmailWithRetry(emailConfig, maxRetries = 3) {
  for (let i = 0; i < maxRetries; i++) {
    try {
      MailApp.sendEmail(emailConfig);
      Logger.log(`✅ Email sent successfully on attempt ${i + 1}`);
      return true;
    } catch (error) {
      Logger.log(`❌ Email attempt ${i + 1} failed: ${error.message}`);
      if (i === maxRetries - 1) throw error;
      Utilities.sleep(1000 * (i + 1)); // Exponential backoff
    }
  }
  return false;
}

/**
 * Test function để kiểm tra dữ liệu
 */
function testReferralData() {
  const CONFIG = {
    spreadsheetId: '14A_0CcRRJtfEKdWs93UsY6L6idDQCtIdZyFfTvt7Le8',
    sheetName: 'referral',
    debugMode: true
  };
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    const sheet = ss.getSheetByName(CONFIG.sheetName);
    
    if (!sheet) {
      Logger.log(`❌ Sheet '${CONFIG.sheetName}' không tồn tại`);
      return;
    }
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    Logger.log(`📊 Tổng số dòng: ${values.length}`);
    Logger.log(`📋 Headers: ${values[0].join(', ')}`);
    
    // Hiển thị 5 dòng đầu tiên
    for (let i = 1; i <= Math.min(5, values.length - 1); i++) {
      Logger.log(`Dòng ${i}: ${values[i].join(' | ')}`);
    }
    
  } catch (error) {
    Logger.log(`❌ Lỗi test: ${error.message}`);
  }
}

/**
 * Tạo trigger để gửi email hàng tuần (Thứ 2 lúc 8:00 AM)
 */
function createWeeklyTrigger() {
  // Xóa trigger cũ nếu có
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendWeeklyReferralReport') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Tạo trigger mới - Thứ 2 hàng tuần lúc 8:00 AM
  ScriptApp.newTrigger('sendWeeklyReferralReport')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
    
  Logger.log('✅ Đã tạo trigger gửi email hàng tuần (Thứ 2, 8:00 AM)');
}

/**
 * Xóa trigger
 */
function deleteWeeklyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendWeeklyReferralReport') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('✅ Đã xóa trigger gửi email hàng tuần');
    }
  });
}