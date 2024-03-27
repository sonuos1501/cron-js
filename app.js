const cron = require('node-cron');
const fetch = require('node-fetch');
const fs = require('fs');

// Hàm chuyển đổi dữ liệu JSON thành chuỗi CSV

// Hàm ghi dữ liệu CSV vào tệp

// Hàm lấy dữ liệu từ API và ghi vào tệp CSV cho một ngày cụ thể
async function fetchDataAndWriteToCSVForDate(date) {
  try {
    const formattedDate = date.toISOString().slice(0, 10).split('-').reverse().join('%2F');
    const apiUrl = `https://www.nldc.evn.vn/api/services/app/Dashboard/GetBieuDoTuongQuanPT?day=${formattedDate}`;
    const response = await fetch(apiUrl);
    const data = await response.json();

    const csvData = convertToCSV(data);

    const filePath = `output_${date.toISOString().slice(0, 10)}.csv`;
    writeCSVToFile(csvData, filePath);

    console.log(`CSV file created for ${date.toISOString().slice(0, 10)}`);
  } catch (error) {
    console.error('Error fetching data:', error);
  }
}

// Hàm lặp qua tất cả các ngày trong năm và gọi fetchDataAndWriteToCSVForDate cho mỗi ngày
async function fetchDataAndWriteToCSVForYear(year) {
  const startDate = new Date(year, 0, 1);
  const endDate = new Date(year, 11, 31);

  for (let date = new Date(startDate); date <= endDate; date.setDate(date.getDate() + 1)) {
    await fetchDataAndWriteToCSVForDate(date);
  }
}

// Lập lịch chạy hàm fetchDataAndWriteToCSVForYear mỗi ngày lúc 12:00 AM
cron.schedule('0 0 * * *', () => {
  fetchDataAndWriteToCSVForYear(2024);
});
