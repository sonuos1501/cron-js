const fetch = require('node-fetch');
const cron = require('node-cron');
const Papa = require('papaparse');
const fs = require('fs');
const xlsx = require('xlsx-populate');

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
  const workbook = await xlsx.fromBlankAsync();
  const sheet = workbook.sheet(0);
  const monthYear = date.split('-');
  sheet.name(`${monthYear[1]}-${monthYear[2]}`);

  // Tạo header cho sheet
  sheet.cell('A1').value('Thời Gian');
  sheet.cell('B1').value('Công Suất');
  sheet.cell('C1').value('Giá Bán');

  // Thêm dữ liệu từ mảng vào sheet
  data.forEach((item, index) => {
    const rowIndex = index + 2;
    sheet.cell(`A${rowIndex}`).value(item.thoiGian);
    sheet.cell(`B${rowIndex}`).value(item.congSuat);
    sheet.cell(`C${rowIndex}`).value(item.giaBan);
  });

  // Lưu workbook vào file
  await workbook.toFileAsync(`/Users/son/Documents/project_sample/cron/output-${monthYear[1]}-${monthYear[2]}.xlsx`);
}

// Lập lịch lấy dữ liệu cho toàn bộ năm 2023
cron.schedule('* * * * *', async () => {
  const year = 2023;
  for (let month = 1; month <= 12; month++) {
    for (let day = 1; day <= 31; day++) {
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

console.log('Scheduled task set for fetching data daily for the year 2023');
