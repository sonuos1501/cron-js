const fs = require('fs');
const fetch = require('node-fetch');
const xlsx = require('xlsx-populate');

const year = 2023

// Đường dẫn file Excel
const filePath = `/Users/son/Documents/project_sample/cron/output-data-${year}.xlsx`;

// Kiểm tra và tạo file nếu chưa tồn tại
function ensureFileExists() {
  if (!fs.existsSync(filePath)) {
    const workbook = xlsx.fromBlankAsync();
    return workbook.then(wb => wb.toFileAsync(filePath));
  }
  return Promise.resolve();
}

// Hàm lấy dữ liệu từ API
async function fetchData(date) {
  const url = `https://www.nldc.evn.vn/api/services/app/Dashboard/GetBieuDoTuongQuanPT?day=${date}`;
  try {
    const response = await fetch(url);

    if (!response.ok) {
      throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const data = await response.json();

    // Check if the data structure is as expected
    if (!data || !data.result || !data.result.data || !data.result.data.phuTai) {
      console.log(url);
      console.error('Data is missing or the format is incorrect:', data);
      return null;  // Return null or appropriate value when data is incorrect
    }

    console.log(data.result.data.phuTai);

    return data.result.data.phuTai;
  } catch (error) {
    console.error('Error fetching data:', error);
    return null;  // Return null to handle errors gracefully
  }
}

// Hàm lưu dữ liệu vào file Excel
async function saveToExcel(data, date) {
  await ensureFileExists();  // Đảm bảo file tồn tại trước khi thực hiện lưu

  let workbook = await xlsx.fromFileAsync(filePath);
  const monthYear = date.split('-');
  const sheetName = `${monthYear[1]}-${monthYear[2]}`;
  let sheet = workbook.sheet(sheetName);

  if (!sheet) {
    sheet = workbook.addSheet(sheetName);
    sheet.cell('A1').value('Thời Gian');
    sheet.cell('B1').value('Công Suất');
    sheet.cell('C1').value('Giá Bán');
  }

  if (workbook.sheet('Sheet1')) {
    workbook.deleteSheet(workbook.sheet('Sheet1'))
  }

  data.forEach((item, index) => {
    const rowIndex = sheet.usedRange() ? sheet.usedRange()._numRows + 1 : 2;
    sheet.cell(`A${rowIndex}`).value(item.thoiGian);
    sheet.cell(`B${rowIndex}`).value(item.congSuat);
    sheet.cell(`C${rowIndex}`).value(item.giaBan);
  });

  await workbook.toFileAsync(filePath);
}

// Lập lịch lấy dữ liệu cho toàn bộ năm 2023
async function main() {
  for (let month = 1; month <= 12; month++) {
    const monthFormatted = month.toString().padStart(2, '0');  // Định dạng tháng với 2 chữ số
    const daysInMonth = new Date(year, month, 0).getDate(); // Số ngày trong tháng
    for (let day = 1; day <= daysInMonth; day++) {
      const dayFormatted = day.toString().padStart(2, '0');  // Định dạng ngày với 2 chữ số
      const date = `${dayFormatted}/${monthFormatted}/${year}`;
      const data = await fetchData(date);
      if (data) {
        await saveToExcel(data, `${dayFormatted}-${monthFormatted}-${year}`);
      }
    }
  }
}

console.log('Scheduled task set for fetching and saving data for the entire year 2023');
main()
