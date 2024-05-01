const fetch = require('node-fetch');
const cron = require('node-cron');
const xlsx = require('xlsx-populate');

// Đường dẫn file Excel
const filePath = `/Users/son/Documents/project_sample/cron/output-data-2023.xlsx`;

// Hàm lấy dữ liệu từ API
async function fetchData(date) {
  const url = `https://www.nldc.evn.vn/api/services/app/Dashboard/GetBieuDoTuongQuanPT?day=${date}`;
  try {
    const response = await fetch(url);
    const data = await response.json();
    return data.result.data.phuTai;
  } catch (error) {
    console.error('Error fetching data:', error);
  }
}

// Hàm lưu dữ liệu vào sheet trong file Excel
async function saveToExcel(data, date) {
  let workbook;
  try {
    // Cố gắng mở workbook hiện có
    workbook = await xlsx.fromFileAsync(filePath);
  } catch (error) {
    // Nếu không có, tạo mới
    workbook = await xlsx.fromBlankAsync();
  }

  const monthYear = date.split('-');
  const sheetName = `${monthYear[1]}-${monthYear[2]}`;
  let sheet = workbook.sheet(sheetName);

  // Nếu sheet không tồn tại, tạo mới
  if (!sheet) {
    sheet = workbook.addSheet(sheetName);
    // Tạo header cho sheet mới
    sheet.cell('A1').value('Thời Gian');
    sheet.cell('B1').value('Công Suất');
    sheet.cell('C1').value('Giá Bán');
  }

  // Tính dòng bắt đầu để thêm dữ liệu mới
  const startRow = sheet.usedRange() ? sheet.usedRange().endCell().row() + 1 : 2;

  // Thêm dữ liệu vào sheet
  data.forEach((item, index) => {
    const rowIndex = startRow + index;
    sheet.cell(`A${rowIndex}`).value(item.thoiGian);
    sheet.cell(`B${rowIndex}`).value(item.congSuat);
    sheet.cell(`C${rowIndex}`).value(item.giaBan);
  });

  // Lưu workbook
  await workbook.toFileAsync(filePath);
}

// Lập lịch lấy dữ liệu cho toàn bộ năm 2023
cron.schedule('* * * * *', async () => {
  const year = 2023;
  for (let month = 1; month <= 12; month++) {
    const daysInMonth = new Date(year, month, 0).getDate(); // Số ngày trong tháng
    for (let day = 1; day <= daysInMonth; day++) {
      const date = `${day}/${month}/${year}`;
      const data = await fetchData(date);
      if (data) {
        saveToExcel(data, `${day}-${month}-${year}`);
      }
    }
  }
}, {
  scheduled: true,
  timezone: "Asia/Ho_Chi_Minh"
});

console.log('Scheduled task set for fetching data for the entire year 2023');
