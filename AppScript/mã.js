///Lệnh deploy, phần ngoặc kép để tên html nào thì khi deploy sẽ ra trang đó
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Form');
}

///Form Nhập báo cáo
function submitForm(data) {
  const sheetBaoCao = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Báo cáo');
  const sheetNhipKD = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Nhịp KD');

  const PRODUCTS_NEED_MULTIPLY = ['Casa', 'Bond', 'FD', 'Lending', 'NFI'];

  function insertRowAtTop(sheet, rowData) {
    sheet.insertRowBefore(2); // luôn insert trước dòng 2 (giữ dòng 1 là tiêu đề)
    sheet.getRange(2, 1, 1, rowData.length).setValues([rowData]);
  }

  const timestamp = "'" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss');

  // ====== Xử lý Báo cáo kế hoạch ======
  if (data.planDate) {
    const planDateObj = new Date(data.planDate);
    const dateStr = Utilities.formatDate(planDateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const week = 'Tuần ' + Math.ceil(planDateObj.getDate() / 7);
    const month = 'Tháng ' + (planDateObj.getMonth() + 1);

    if (data.planDetails?.length) {
      data.planDetails.forEach(entry => {
        // let value = parseFloat(entry.value);
        let value = Number(entry.value) || 0
        // if (entry.product && entry.detail && !isNaN(value) && value > 0) {
          if (entry.product && entry.detail && !isNaN(value)) {
          insertRowAtTop(sheetBaoCao, [
            dateStr, week, data.name, data.title, "BC đầu ngày",
            entry.product, entry.detail, value, month
          ]);
        }
      });
    }

    insertRowAtTop(sheetNhipKD, [
      timestamp, dateStr, week, month, data.name, data.title, "BC đầu ngày",
      data.planTasks || "",
      data.planGoi || "", data.planGoiDetail || "",
      data.planHen || "", data.planHenDetail || "",
      data.planGap || "", data.planGapDetail || "",
      data.planOtherTasks || "", data.planNotes || ""
    ]);
  }

  // ====== Xử lý Báo cáo kết quả ======
  if (data.resultDate) {
    const resultDateObj = new Date(data.resultDate);
    const dateStr = Utilities.formatDate(resultDateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const week = 'Tuần ' + Math.ceil(resultDateObj.getDate() / 7);
    const month = 'Tháng ' + (resultDateObj.getMonth() + 1);

    if (data.resultDetails?.length) {
      data.resultDetails.forEach(entry => {
        // let value = parseFloat(entry.value);
        let value = Number(entry.value) || 0

        // if (entry.product && entry.detail && !isNaN(value) && value > 0) {
          if (entry.product && entry.detail && !isNaN(value)) {
          insertRowAtTop(sheetBaoCao, [
            dateStr, week, data.name, data.title, "BC cuối ngày",
            entry.product, entry.detail, value, month
          ]);
        }
      });
    }

    insertRowAtTop(sheetNhipKD, [
      timestamp, dateStr, week, month, data.name, data.title, "BC cuối ngày",
      data.resultTasks || "",
      data.resultGoi || "", data.resultGoiDetail || "",
      data.resultHen || "", data.resultHenDetail || "",
      data.resultGap || "", data.resultGapDetail || "",
      data.resultOtherTasks || "", data.resultNotes || ""
    ]);
  }
}

///--- BẢO CÁO CHECK NHẬP REPORT

function getCheckReport(mode, param) {
  const userList = [
  "KSV1", "GDV1", "SRM1", "Trưởng phòng tín dụng",
  "Tín dụng 1", "Tín dụng 2", "Tín dụng 3", "Tín dụng 4"
  ];

  // Khởi tạo danh sách kết quả mặc định
  const result = userList.map(name => ({ name, am: false, pm: false }));

  const targetDate = new Date(param.date);
  const checkDate = (rawDate) => {
    if (!rawDate) return false;
    const d = new Date(rawDate);
    return d.toDateString() === targetDate.toDateString();
  };

  // Hàm xử lý dữ liệu từng sheet
  function processSheet(sheetName, nameColLabel, typeColLabel, dateColLabel) {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    const header = data[0];
    const rows = data.slice(1);

    const nameCol = header.indexOf(nameColLabel);
    const typeCol = header.indexOf(typeColLabel);
    const dateCol = header.indexOf(dateColLabel);

    for (let row of rows) {
      const rawDate = row[dateCol];
      if (!checkDate(rawDate)) continue;

      const name = row[nameCol];
      const type = row[typeCol];
      if (!name || !type) continue;

      const match = result.find(r => r.name === name);
      if (!match) continue;

      if (type === "BC đầu ngày") match.am = true;
      if (type === "BC cuối ngày") match.pm = true;
    }
  }

  // Xử lý cả hai sheet
  processSheet("Báo cáo", "Họ và tên", "Loại báo cáo", "Ngày nhập");
  processSheet("Nhịp KD", "Họ và tên", "Loại báo cáo", "Ngày");

  return result;
}

// Giao diện
function showCheckReportSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("CheckInputReport")
    .setTitle("Check Nhập Báo Cáo")
    .setWidth(400);
  SpreadsheetApp.getUi().showModalDialog(html, "Check Nhập Báo Cáo");
}

function showSummaryReportSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("SummaryReport")
    .setTitle("Báo Cáo Tổng Hợp")
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getCheckReportData(mode, param) {
  return getCheckReport(mode, param);
}

function openForm() {
  const html = HtmlService.createHtmlOutputFromFile("Form")
    .setWidth(900)  // Tăng kích thước cho modal dialog
    .setHeight(700)
    .setTitle("📝 Nhập Báo Cáo Kinh Doanh");
  SpreadsheetApp.getUi().showModalDialog(html, "📝 NHẬP BÁO CÁO KINH DOANH");
}

 
/// --> XÂY DỰNG BÁO CÁO REPORT KẾT QUẢ KINH DOANH ***


function getSummaryReportByDate(mode, param) {
  const data = getSheetData('Báo cáo');
  const headers = data[0];
  const rows = data.slice(1);

  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const resultsMap = new Map();
  const targetDate = new Date(param.date);
  const targetDateStr = targetDate.toDateString();

  // Danh sách nhân viên cần loại trừ
  const excludeStaff = ["KSV1", "GDV1", "SRM1"];

  rows.forEach(row => {
    const dateRaw = row[idx["Ngày nhập"]];
    const weekStr = row[idx["Tuần"]];
    const name = row[idx["Họ và tên"]];
    const type = row[idx["Loại báo cáo"]];
    const productRaw = row[idx["Sản phẩm"]];
    const detailRaw = row[idx["Chi tiết sản phẩm"]];
    const value = row[idx["Giá trị"]];

    if (!dateRaw || !name || !type || !productRaw || !detailRaw || value === undefined || value === '') return;
    
    // Chỉ loại trừ khi không lọc theo từng nhân viên
    if (!param.staff && excludeStaff.includes(name)) return;
    
    // Kiểm tra nếu đang xem theo nhân viên cụ thể
    if (param.staff && name !== param.staff) return;

    const reportDate = new Date(dateRaw);
    let matchDate = false;

    switch(mode) {
      case "date":
        matchDate = reportDate.toDateString() === targetDateStr;
        break;
      case "month":
        matchDate = reportDate.getMonth() === targetDate.getMonth() && 
                   reportDate.getFullYear() === targetDate.getFullYear();
        break;
      case "week":
        matchDate = extractWeekNumber(weekStr) === param.week;
        break;
      case "year":
        matchDate = reportDate.getFullYear() === param.year;
        break;
    }

    if (!matchDate) return;

    const product = productRaw.toString().trim().replace(/\s+/g, ' ');
    const detail = detailRaw.toString().trim().replace(/\s+/g, ' ');

    let dateKey = Utilities.formatDate(reportDate, "GMT+7", "yyyy-MM-dd");
    if (mode === "week") dateKey = `W${param.week}`;
    if (mode === "month") dateKey = Utilities.formatDate(reportDate, "GMT+7", "yyyy-MM");
    if (mode === "year") dateKey = reportDate.getFullYear().toString();

    const key = [product, detail, dateKey, param.staff || "All"].join("|");

    if (!resultsMap.has(key)) {
      resultsMap.set(key, {
        product,
        detail,
        begin: 0,
        end: 0,
        staff: param.staff || "Toàn chi nhánh"
      });
    }

    const entry = resultsMap.get(key);
    if (type === "BC đầu ngày") entry.begin += Number(value);
    if (type === "BC cuối ngày") entry.end += Number(value);
  });

  return Array.from(resultsMap.values());
}

function extractWeekNumber(weekStr) {
  const match = weekStr.match(/\d+/);
  return match ? parseInt(match[0]) : null;
}





///SỬA GIAO DIỆN CHO BÁO CÁO NHẬP, BÁO CÁO TỔNG HỌP, CHECK NHẬP BÁO CÁO
function openForm() {
  const html = HtmlService.createHtmlOutputFromFile("Form")
    .setWidth(900)  // Tăng kích thước cho modal dialog
    .setHeight(700)
    .setTitle("📝 Nhập Báo Cáo Kinh Doanh");
  SpreadsheetApp.getUi().showModalDialog(html, "📝 NHẬP BÁO CÁO KINH DOANH");
}

function openReport() {
  const html = HtmlService.createHtmlOutputFromFile("Report")
    .setWidth(1100)
    .setHeight(750)
    .setTitle("📊 Báo Cáo Tổng Hợp Đơn Vị");
  SpreadsheetApp.getUi().showModalDialog(html, "📊 BÁO CÁO TỔNG HỢP ĐƠN VỊ");
}

function showCheckReportSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("CheckInputReport")
    .setWidth(1000)
    .setHeight(700)
    .setTitle("✅ Check Nhập Báo Cáo Ngày");
  SpreadsheetApp.getUi().showModalDialog(html, "✅ CHECK NHẬP BÁO CÁO");
}

//// TẠO DASHBOARD TỔNG HỢP BÁO CÁO
function getEmployeeDailyReportRaw(param) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Báo cáo");
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1);

  const colDate = header.indexOf("Ngày nhập");
  const colName = header.indexOf("Họ và tên");
  const colTitle = header.indexOf("Chức danh");
  const colType = header.indexOf("Loại báo cáo");
  const colProduct = header.indexOf("Sản phẩm");
  const colDetail = header.indexOf("Chi tiết sản phẩm");
  const colValue = header.indexOf("Giá trị");

  const resultMap = {};
  const headersSet = new Set();

  rows.forEach(row => {
    const rawDate = row[colDate];
    const type = row[colType];
    const name = row[colName];
    const title = row[colTitle];
    const product = row[colProduct];
    const detail = row[colDetail];
    // const value = parseFloat(row[colValue]) || "";

    // const value = parseFloat(row[colValue])
    // const safeValue = isNaN(value) ? 0 : value
    const value = Number(row[colValue]) || 0



    if (!rawDate || !type || !name || !product || !detail) return;

    const reportDate = new Date(rawDate);
    const viewDate = new Date(param.date);
    let include = false;

    if (param.mode === "date") {
      include = reportDate.toDateString() === viewDate.toDateString();
    } else if (param.mode === "month") {
      include = reportDate.getMonth() === viewDate.getMonth() &&
                reportDate.getFullYear() === viewDate.getFullYear();
    } else if (param.mode === "year") {
      include = reportDate.getFullYear() === viewDate.getFullYear();
    }

    if (!include) return;

    const key = name + "||" + title;
    const colKey = `${product} - ${detail}`;
    headersSet.add(colKey);

    if (!resultMap[key]) {
      resultMap[key] = {
        name,
        title,
        begin: {},
        end: {}
      };
    }

    if (type === "BC đầu ngày") {
      resultMap[key].begin[colKey] = value;
    } else if (type === "BC cuối ngày") {
      resultMap[key].end[colKey] = value;
    }
  });

  const headers = Array.from(headersSet).sort();
  const rowsFormatted = Object.values(resultMap).map(row => {
    // Đảm bảo đủ tất cả colKey và mặc định là 0 nếu thiếu
    headers.forEach(colKey => {
      if (!row.begin.hasOwnProperty(colKey)) row.begin[colKey] = 0;
      if (!row.end.hasOwnProperty(colKey)) row.end[colKey] = 0;
    });
    return row;
  });

  const result = {
    headers,
    rows: rowsFormatted
  };
  Logger.log(JSON.stringify(result, null, 2)); // THÊM DÒNG NÀY VÀO
  return result
}

  
function openEmployeeDailyReport() {
  const html = HtmlService.createHtmlOutputFromFile("EmployeeDailyReport")
    .setWidth(1400)
    .setHeight(800)
    .setTitle("📘 Báo Cáo Nhân Viên");
  SpreadsheetApp.getUi().showModalDialog(html, "📘 BÁO CÁO CHI TIẾT NHÂN VIÊN");
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📊 Báo Cáo")
    .addItem("📤 Nhập Báo Cáo", "openForm")
    .addItem("✅ Check Báo Cáo Đầu - Cuối", "showCheckReportSidebar")
    .addItem("📊 Xem KQKD", "showKQKDSidebar")
    .addItem("📈 Báo Cáo LD Team", "openLDteamReport")
    .addItem("📝 Nhập cam kết tuần", "openWeeklyReport")
    .addItem("📋 Xem Cam kết tuần", "openCKKDReport")
    .addItem("📈 Xem KPI", "openKPIReport")
    .addToUi();
}

/// BỔ SUNG GIAO DIỆN CHO BÁO CÁO NHÂN VIÊN
// Gọi giao diện EmployeeDailyReport từ sidebar
function openEmployeeDailyReportSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("EmployeeDailyReport")
    .setTitle("Báo Cáo Chi Tiết Nhân Viên")
    .setWidth(1300);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Xử lý dữ liệu báo cáo theo nhiều chế độ: ngày/tháng/tuần/năm
function getEmployeeDailyReport(param) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Báo cáo");
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1);

  const colDate = header.indexOf("Ngày nhập");
  const colName = header.indexOf("Họ và tên");
  const colTitle = header.indexOf("Chức danh");
  const colType = header.indexOf("Loại báo cáo");
  const colProduct = header.indexOf("Sản phẩm");
  const colDetail = header.indexOf("Chi tiết sản phẩm");
  const colValue = header.indexOf("Giá trị");

  const resultMap = {};
  const headersSet = new Set();

  rows.forEach(row => {
    const rawDate = row[colDate];
    const type = row[colType];
    const name = row[colName];
    const title = row[colTitle];
    const product = row[colProduct];
    const detail = row[colDetail];
    const value = row[colValue];

    if (!rawDate || !type || !name || !product || !detail || value === undefined || value === '') return;

    const reportDate = new Date(rawDate);
    const key = name + "||" + title;
    const colKey = `${product} - ${detail}`;
    headersSet.add(colKey);

    let include = false;
    const viewDate = new Date(param.date);

    if (param.mode === "date") {
      include = reportDate.toDateString() === viewDate.toDateString();
    } else if (param.mode === "month") {
      include = reportDate.getMonth() + 1 === Number(param.month) && reportDate.getFullYear() === viewDate.getFullYear();
    } else if (param.mode === "year") {
      include = reportDate.getFullYear() === viewDate.getFullYear();
    } else if (param.mode === "week") {
      const reportMonth = reportDate.getMonth() + 1;
      const reportWeek = Math.ceil(reportDate.getDate() / 7);
      include = reportMonth === Number(param.month) && reportWeek === Number(param.week);
    }

    if (!include) return;
    if (param.names && param.names.length > 0 && !param.names.includes(name)) return;

    if (!resultMap[key]) resultMap[key] = [];

    resultMap[key].push({
      name,
      title,
      reportType: type,
      values: { [colKey]: Number(value) }
    });
  });

  const headers = Array.from(headersSet).sort();
  const finalRows = [];

  Object.entries(resultMap).forEach(([key, entries]) => {
    const grouped = { begin: {}, end: {} };
    const [name, title] = key.split("||");

    entries.forEach(entry => {
      const type = entry.reportType;
      const values = entry.values;
      const store = type === "BC đầu ngày" ? grouped.begin : grouped.end;
      for (let k in values) store[k] = values[k];
    });

    finalRows.push({ name, title, reportType: "BC đầu ngày", values: grouped.begin });
    finalRows.push({ name, title, reportType: "BC cuối ngày", values: grouped.end });
  });

  return { headers, rows: finalRows };
}

function openEmployeeDailyReport() {
  const html = HtmlService.createHtmlOutputFromFile("EmployeeDailyReport")
    .setWidth(1400)
    .setHeight(800)
    .setTitle("📘 Báo Cáo Nhân Viên");
  SpreadsheetApp.getUi().showModalDialog(html, "📘 BÁO CÁO CHI TIẾT NHÂN VIÊN");
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📊 Báo Cáo")
    .addItem("📤 Nhập Báo Cáo", "openForm")
    .addItem("✅ Check Báo Cáo Đầu - Cuối", "showCheckReportSidebar")
    .addItem("📊 Xem KQKD", "showKQKDSidebar")
    .addItem("📈 Báo Cáo LD Team", "openLDteamReport")
    .addItem("📝 Nhập cam kết tuần", "openWeeklyReport")
    .addItem("📋 Xem Cam kết tuần", "openCKKDReport")
    .addItem("📈 Xem KPI", "openKPIReport")
    .addToUi();
}

/// --> TẠO KQKD MSB 
function getKQKDFilteredData(filter) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Báo cáo");
  const data = sheet.getDataRange().getValues();
  const header = data[0];

  const colDate = header.indexOf("Ngày nhập");
  const colWeek = header.indexOf("Tuần");
  const colLoaiBC = header.indexOf("Loại báo cáo");
  const colChiTiet = header.indexOf("Chi tiết sản phẩm");
  const colGiaTri = header.indexOf("Giá trị");
  const colNhanVien = header.indexOf("Họ và tên");

  const result = {};

  const allChiTiet = [
    "Casa BQ tăng net", "CCQ BQ tăng net",
    "Bond mới", "Bond gia hạn",
    "FD mới", "FD gia hạn",
    "GN - Thế chấp", "GN - Tín chấp", "GN - Ứng vốn",
    "Chi tiêu",
    "TKSĐ", "Banca NT", "Banca PNT", "FX", "Khác",
    "Tự mở", "CTV", "Activation"
  ];

  allChiTiet.forEach(ct => {
    result[ct] = {
      "Tuần 1": 0,
      "Tuần 2": 0,
      "Tuần 3": 0,
      "Tuần 4": 0,
      "Tuần 5": 0,
      "Tổng": 0
    };
  });

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const loaiBC = row[colLoaiBC];
    const ctsp = row[colChiTiet]?.toString().trim();
    const giatri = parseFloat(row[colGiaTri]) || 0;

    if (loaiBC !== "BC cuối ngày") continue;
    if (!ctsp || !result[ctsp]) continue;

    const date = new Date(row[colDate]);
    const month = date.getMonth() + 1;
    const year = date.getFullYear();
    const week = row[colWeek];
    const nhanvien = row[colNhanVien];

    if (filter.year && parseInt(filter.year) !== year) continue;
    if (filter.month && parseInt(filter.month) !== month) continue;

    // So sánh nhiều nhân viên
    if (Array.isArray(filter.nhanvien) && !filter.nhanvien.includes(nhanvien)) continue;

    if (week && result[ctsp][week] !== undefined) {
      result[ctsp][week] += giatri;
    }

    result[ctsp]["Tổng"] += giatri;
  }

  return result;
}


function getFilterOptions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Báo cáo");
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const colDate = header.indexOf("Ngày nhập");
  const colNhanVien = header.indexOf("Họ và tên");

  const months = new Set();
  const years = new Set();
  const employees = new Set();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const date = new Date(row[colDate]);
    if (isNaN(date)) continue;

    months.add(date.getMonth() + 1);
    years.add(date.getFullYear());
    employees.add(row[colNhanVien]);
  }

  return {
    months: Array.from(months).sort((a, b) => a - b),
    years: Array.from(years).sort((a, b) => b - a),
    employees: ["Toàn chi nhánh", ...Array.from(employees).sort()]
  };
}

function showReportWithFilter() {
  const html = HtmlService.createHtmlOutputFromFile("ReportFilter")
    .setWidth(1400)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, "Báo cáo KQKD toàn màn hình");
}

function showKQKDSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("ReportFilter")
    .setWidth(1400)
    .setHeight(800)
    .setTitle("📊 Báo Cáo KQKD");
  SpreadsheetApp.getUi().showModalDialog(html, "📊 BÁO CÁO KQKD");
}

function openReportFullScreen() {
  const html = HtmlService.createHtmlOutputFromFile("ReportFilter")
    .setWidth(1600)
    .setHeight(900) // bạn có thể tăng thêm nếu muốn full lớn
    .setTitle("📊 KQKD Tuần Sản Phẩm");

  SpreadsheetApp.getUi().showModalDialog(html, "📊 BÁO CÁO KQKD");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 REPORT")
    .addItem("📝 REPORT Input Daily", "openForm")
    .addItem("📘 REPORT Kết quả KD Đầu ngày- Cuối ngày", "openReport")
    .addItem("📊 REPORT Kết quả KD Weekly", "showModalReportFilter")  // Replaced sidebar function
    .addItem("✅ CHECK Nhân sự Báo cáo", "showCheckReportSidebar")
    .addItem("🖥️ REPORT Nhịp KD Fullscreen", "openNhipKDDialog")
    .addItem("📋 Xem Cam kết tuần", "openCKKDReport")
    .addItem("📈 Xem KPI", "openKPIReport")
    .addToUi();
}

function showReportWithFilter() {
  const html = HtmlService.createHtmlOutputFromFile("ReportFilter")
    .setWidth(1600)  // Đặt chiều rộng của cửa sổ popup lớn hơn
    .setHeight(900)  // Đặt chiều cao của cửa sổ popup phù hợp với màn hình lớn
    .setTitle("Báo cáo KQKD tuần theo sản phẩm");

  SpreadsheetApp.getUi().showModalDialog(html, "Báo cáo KQKD");
}

function showModalReportFilter() {
  const html = HtmlService.createHtmlOutputFromFile("ReportFilter")
    .setWidth(1200)  // Set width for the modal dialog
    .setHeight(800)  // Set height for the modal dialog
    .setTitle("📊 Báo cáo KQKD tuần theo sản phẩm");

  SpreadsheetApp.getUi().showModalDialog(html, "📊 Báo cáo KQKD tuần theo sản phẩm");
}



/// XÂY DỰNG REPORT NHỊP KINH DOANH
// File Code.gs hoàn chỉnh cho Nhịp KD

function getGopBaoCaoNhipKD() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Nhịp KD");
  if (!sheet) throw new Error("Không tìm thấy sheet Nhịp KD");

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const idx = {};
  headers.forEach((h, i) => idx[h.trim()] = i);

  const grouped = new Map();

  for (let row of rows) {
    const dateRaw = row[idx["Ngày"]];
    const name = row[idx["Họ và tên"]];
    const title = row[idx["Chức danh"]];
    const type = row[idx["Loại báo cáo"]];
    const week = row[idx["Tuần"]];
    const month = row[idx["Tháng"]];

    if (!dateRaw || !name || !type) continue;

    const date = Utilities.formatDate(new Date(dateRaw), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const key = `${name}_${date}_${type}`;

    grouped.set(key, {
      date: date,
      week: week,
      month: month,
      name: name,
      title: title,
      type: type,
      task: row[idx["CV triển khai trong ngày"]] || "",
      goi: row[idx["Gọi- Số cuộc gọi"]] || "",
      goiDetail: row[idx["Chi tiết cuộc gọi"]] || "",
      hen: row[idx["Hẹn- Số cuộc hẹn"]] || "",
      henDetail: row[idx["Chi tiết cuộc hẹn"]] || "",
      gap: row[idx["Gặp- Số cuộc gặp"]] || "",
      gapDetail: row[idx["Chi tiết cuộc gặp"]] || "",
      other: row[idx["Công việc khác"]] || "",
      note: row[idx["Ghi chú"]] || ""
    });
  }

  return Array.from(grouped.values());
}

function openNhipKDSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Nhịp KD")
    .setTitle("📋 Báo cáo Nhịp KD");
  SpreadsheetApp.getUi().showSidebar(html);
}

function openNhipKDDialog() {
  const html = HtmlService.createHtmlOutputFromFile("Nhịp KD")
    .setWidth(1400)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, "📋 Báo cáo Nhịp KD");
}

// Cache for spreadsheet data
let sheetCache = {
  baoCao: null,
  nhipKD: null,
  lastUpdate: null
};

// Cache timeout in milliseconds (5 minutes)
const CACHE_TIMEOUT = 5 * 60 * 1000;

// Get sheet data with caching
function getSheetData(sheetName) {
  const now = new Date().getTime();
  if (!sheetCache[sheetName] || !sheetCache.lastUpdate || (now - sheetCache.lastUpdate > CACHE_TIMEOUT)) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) throw new Error(`Không tìm thấy sheet ${sheetName}`);
    sheetCache[sheetName] = sheet.getDataRange().getValues();
    sheetCache.lastUpdate = now;
  }
  return sheetCache[sheetName];
}

// Optimized submitForm function
function submitForm(data) {
  const sheetBaoCao = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Báo cáo');
  const sheetNhipKD = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Nhịp KD');

  const timestamp = "'" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss');
  const baoCaoData = [];
  const nhipKDData = [];

  // Process plan data
  if (data.planDate) {
    const planDateObj = new Date(data.planDate);
    const dateStr = Utilities.formatDate(planDateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const week = `Tuần ${Math.ceil(planDateObj.getDate() / 7)}`;
    const month = `Tháng ${planDateObj.getMonth() + 1}`;

    // Prepare plan details data for Báo cáo sheet
    if (data.planDetails?.length) {
      data.planDetails.forEach(entry => {
        // Chỉ xử lý nếu có product, detail và value được nhập
        if (entry.product && entry.detail && entry.value !== undefined && entry.value !== '' && entry.value !== null) {
          const customerName = entry.customerName || "";
          baoCaoData.push([
            dateStr, week, data.name, data.title, "BC đầu ngày",
            entry.product, entry.detail, entry.value, month,
            customerName // Add customer name as 10th column
          ]);
        }
      });
    }

    // Chỉ thêm vào nhipKDData nếu có ít nhất một trường được nhập
    if (data.planTasks || data.planGoi || data.planGoiDetail || data.planHen || 
        data.planHenDetail || data.planGap || data.planGapDetail || 
        data.planOtherTasks || data.planNotes) {
      nhipKDData.push([
        timestamp, dateStr, week, month, data.name, data.title, "BC đầu ngày",
        data.planTasks || "",
        data.planGoi || "", data.planGoiDetail || "",
        data.planHen || "", data.planHenDetail || "",
        data.planGap || "", data.planGapDetail || "",
        data.planOtherTasks || "", data.planNotes || ""
      ]);
    }
  }

  // Process result data
  if (data.resultDate) {
    const resultDateObj = new Date(data.resultDate);
    const dateStr = Utilities.formatDate(resultDateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const week = `Tuần ${Math.ceil(resultDateObj.getDate() / 7)}`;
    const month = `Tháng ${resultDateObj.getMonth() + 1}`;

    // Prepare result details data for Báo cáo sheet
    if (data.resultDetails?.length) {
      data.resultDetails.forEach(entry => {
        // Chỉ xử lý nếu có product, detail và value được nhập
        if (entry.product && entry.detail && entry.value !== undefined && entry.value !== '' && entry.value !== null) {
          const customerName = entry.customerName || "";
          baoCaoData.push([
            dateStr, week, data.name, data.title, "BC cuối ngày",
            entry.product, entry.detail, entry.value, month,
            customerName // Add customer name as 10th column
          ]);
        }
      });
    }

    // Chỉ thêm vào nhipKDData nếu có ít nhất một trường được nhập
    if (data.resultTasks || data.resultGoi || data.resultGoiDetail || data.resultHen || 
        data.resultHenDetail || data.resultGap || data.resultGapDetail || 
        data.resultOtherTasks || data.resultNotes) {
      nhipKDData.push([
        timestamp, dateStr, week, month, data.name, data.title, "BC cuối ngày",
        data.resultTasks || "",
        data.resultGoi || "", data.resultGoiDetail || "",
        data.resultHen || "", data.resultHenDetail || "",
        data.resultGap || "", data.resultGapDetail || "",
        data.resultOtherTasks || "", data.resultNotes || ""
      ]);
    }
  }

  // Batch write to sheets only if there is data to write
  if (baoCaoData.length > 0) {
    sheetBaoCao.insertRowsBefore(2, baoCaoData.length);
    sheetBaoCao.getRange(2, 1, baoCaoData.length, 10).setValues(baoCaoData); // Update to 10 columns
  }

  if (nhipKDData.length > 0) {
    sheetNhipKD.insertRowsBefore(2, nhipKDData.length);
    sheetNhipKD.getRange(2, 1, nhipKDData.length, 16).setValues(nhipKDData);
  }
}

// Optimized getCheckReport function
function getCheckReport(mode, param) {
  const userList = [
  "KSV1", "GDV1", "SRM1", "Trưởng phòng tín dụng",
  "Tín dụng 1", "Tín dụng 2", "Tín dụng 3", "Tín dụng 4"
  ];

  const result = userList.map(name => ({ name, am: false, pm: false }));
  const targetDate = new Date(param.date);
  const targetDateStr = targetDate.toDateString();

  // Get data from cache
  const baoCaoData = getSheetData('Báo cáo');
  const nhipKDData = getSheetData('Nhịp KD');

  // Process both sheets
  [baoCaoData, nhipKDData].forEach(data => {
    const header = data[0];
    const rows = data.slice(1);

    const nameCol = header.indexOf("Họ và tên");
    const typeCol = header.indexOf("Loại báo cáo");
    const dateCol = header.indexOf("Ngày nhập");

    if (nameCol === -1 || typeCol === -1 || dateCol === -1) return;

    rows.forEach(row => {
      const rowDate = new Date(row[dateCol]);
      if (rowDate.toDateString() !== targetDateStr) return;

      const name = row[nameCol];
      const type = row[typeCol];
      if (!name || !type) return;

      const match = result.find(r => r.name === name);
      if (!match) return;

      if (type === "BC đầu ngày") match.am = true;
      if (type === "BC cuối ngày") match.pm = true;
    });
  });

  return result;
}

// Optimized getSummaryReportByDate function
function getSummaryReportByDate(mode, param) {
  const data = getSheetData('Báo cáo');
  const headers = data[0];
  const rows = data.slice(1);

  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const resultsMap = new Map();
  const targetDate = new Date(param.date);
  const targetDateStr = targetDate.toDateString();

  // Danh sách nhân viên cần loại trừ
  const excludeStaff = ["KSV1", "GDV1", "SRM1"];

  rows.forEach(row => {
    const dateRaw = row[idx["Ngày nhập"]];
    const weekStr = row[idx["Tuần"]];
    const name = row[idx["Họ và tên"]];
    const type = row[idx["Loại báo cáo"]];
    const productRaw = row[idx["Sản phẩm"]];
    const detailRaw = row[idx["Chi tiết sản phẩm"]];
    const value = row[idx["Giá trị"]];

    if (!dateRaw || !name || !type || !productRaw || !detailRaw || value === undefined || value === '') return;
    
    // Chỉ loại trừ khi không lọc theo từng nhân viên
    if (!param.staff && excludeStaff.includes(name)) return;
    
    // Kiểm tra nếu đang xem theo nhân viên cụ thể
    if (param.staff && name !== param.staff) return;

    const reportDate = new Date(dateRaw);
    let matchDate = false;

    switch(mode) {
      case "date":
        matchDate = reportDate.toDateString() === targetDateStr;
        break;
      case "month":
        matchDate = reportDate.getMonth() === targetDate.getMonth() && 
                   reportDate.getFullYear() === targetDate.getFullYear();
        break;
      case "week":
        matchDate = extractWeekNumber(weekStr) === param.week;
        break;
      case "year":
        matchDate = reportDate.getFullYear() === param.year;
        break;
    }

    if (!matchDate) return;

    const product = productRaw.toString().trim().replace(/\s+/g, ' ');
    const detail = detailRaw.toString().trim().replace(/\s+/g, ' ');

    let dateKey = Utilities.formatDate(reportDate, "GMT+7", "yyyy-MM-dd");
    if (mode === "week") dateKey = `W${param.week}`;
    if (mode === "month") dateKey = Utilities.formatDate(reportDate, "GMT+7", "yyyy-MM");
    if (mode === "year") dateKey = reportDate.getFullYear().toString();

    const key = [product, detail, dateKey, param.staff || "All"].join("|");

    if (!resultsMap.has(key)) {
      resultsMap.set(key, {
        product,
        detail,
        begin: 0,
        end: 0,
        staff: param.staff || "Toàn chi nhánh"
      });
    }

    const entry = resultsMap.get(key);
    if (type === "BC đầu ngày") entry.begin += Number(value);
    if (type === "BC cuối ngày") entry.end += Number(value);
  });

  return Array.from(resultsMap.values());
}

// Helper function to extract week number
function extractWeekNumber(weekStr) {
  const match = weekStr.match(/Tuần (\d+)/);
  return match ? parseInt(match[1]) : 0;
}

// UI Functions
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📊 Báo Cáo")
    .addItem("📤 Nhập Báo Cáo", "openForm")
    .addItem("✅ Check Báo Cáo Đầu - Cuối", "showCheckReportSidebar")
    .addItem("📊 Xem KQKD", "showKQKDSidebar")
    .addItem("📈 Báo Cáo LD Team", "openLDteamReport")
    .addItem("📝 Nhập cam kết tuần", "openWeeklyReport")
    .addItem("📋 Xem Cam kết tuần", "openCKKDReport")
    .addItem("📈 Xem KPI", "openKPIReport")
    .addToUi();
}

function openForm() {
  const html = HtmlService.createHtmlOutputFromFile("Form")
    .setWidth(900)  // Tăng kích thước cho modal dialog
    .setHeight(700)
    .setTitle("📝 Nhập Báo Cáo Kinh Doanh");
  SpreadsheetApp.getUi().showModalDialog(html, "📝 NHẬP BÁO CÁO KINH DOANH");
}

function showCheckReportSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("CheckInputReport")
    .setWidth(1000)
    .setHeight(700)
    .setTitle("✅ Check Nhập Báo Cáo Ngày");
  SpreadsheetApp.getUi().showModalDialog(html, "✅ CHECK NHẬP BÁO CÁO");
}

function openEmployeeDailyReport() {
  const html = HtmlService.createHtmlOutputFromFile("EmployeeDailyReport")
    .setWidth(1400)
    .setHeight(800)
    .setTitle("📘 Báo Cáo Nhân Viên");
  SpreadsheetApp.getUi().showModalDialog(html, "📘 BÁO CÁO CHI TIẾT NHÂN VIÊN");
}

function showKQKDSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("ReportFilter")
    .setWidth(1400)
    .setHeight(800)
    .setTitle("📊 Báo Cáo KQKD");
  SpreadsheetApp.getUi().showModalDialog(html, "📊 BÁO CÁO KQKD");
}

function openLDteamReport() {
  const html = HtmlService.createHtmlOutputFromFile("LDteam")
    .setWidth(1400)
    .setHeight(800)
    .setTitle("📈 Báo Cáo LD Team");
  SpreadsheetApp.getUi().showModalDialog(html, "📈 BÁO CÁO LD TEAM");
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📊 Báo Cáo")
    .addItem("📤 Nhập Báo Cáo", "openForm")
    .addItem("✅ Check Báo Cáo Đầu - Cuối", "showCheckReportSidebar")
    .addItem("📊 Xem KQKD", "showKQKDSidebar")
    .addItem("📈 Báo Cáo LD Team", "openLDteamReport")
    .addItem("📝 Nhập cam kết tuần", "openWeeklyReport")
    .addItem("📋 Xem Cam kết tuần", "openCKKDReport")
    .addItem("📈 Xem KPI", "openKPIReport")
    .addToUi();
}

function getLDteamData(params) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Báo cáo");
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1);

  // Danh sách nhân viên hiển thị (đúng thứ tự, đủ 6 người)
  const staffList = [
    "KSV1",
    "GDV1",
    "SRM1",
    "Trưởng phòng tín dụng",
    "Tín dụng 1",
    "Tín dụng 2",
    "Tín dụng 3",
    "Tín dụng 4"
  ];

  // Khởi tạo productDetails dạng mảng có thứ tự
  const productDetails = [
    { product: 'Lending', details: ['Trình', 'Phê duyệt', 'GN - Tín chấp', 'GN - Thế chấp', 'GN - Ứng vốn'] },
    { product: 'Thẻ tín dụng', details: ['Trình', 'Phê duyệt', 'Active', 'Chi tiêu'] },
    { product: 'Casa', details: ['CASA BQ tăng net'] },
    { product: 'NFI', details: ['TKSĐ', 'FX', 'Banca NT', 'Banca PNT', 'Khác'] },
    { product: 'TKMM', details: ['Tự mở', 'CTV', 'Activation'] }
  ];

  // Khởi tạo kết quả với cấu trúc cố định
  let result = {};
  productDetails.forEach(pd => {
    result[pd.product] = {};
    pd.details.forEach(detail => {
      if (params.dataType === 'personal') {
        result[pd.product][detail] = {};
        staffList.forEach(staff => {
          result[pd.product][detail][staff] = 0;
        });
      } else {
        result[pd.product][detail] = { begin: 0, end: 0 };
      }
    });
  });

  // Map sản phẩm và chi tiết
  const productMapping = {
    'Lending': {
      'Trình': 'Lending',
      'Phê duyệt': 'Lending',
      'GN - Tín chấp': 'Lending',
      'GN - Thế chấp': 'Lending',
      'GN - Ứng vốn': 'Lending'
    },
    'Thẻ tín dụng': {
      'Trình': 'Thẻ tín dụng',
      'Phê duyệt': 'Thẻ tín dụng',
      'Active': 'Thẻ tín dụng',
      'Chi tiêu': 'Thẻ tín dụng'
    },
    'Casa': {
      'CASA BQ tăng net': 'Casa'
    },
    'NFI': {
      'TKSĐ': 'NFI',
      'FX': 'NFI',
      'Banca NT': 'NFI',
      'Banca PNT': 'NFI',
      'Khác': 'NFI'
    },
    'TKMM': {
      'Tự mở': 'TKMM',
      'CTV': 'TKMM',
      'Activation': 'TKMM'
    }
  };

  // Xử lý dữ liệu theo thời gian
  rows.forEach(row => {
    const date = new Date(row[header.indexOf("Ngày nhập")]);
    const type = row[header.indexOf("Loại báo cáo")];
    const product = row[header.indexOf("Sản phẩm")]?.toString().trim();
    const detail = row[header.indexOf("Chi tiết sản phẩm")]?.toString().trim();
    const value = Number(row[header.indexOf("Giá trị")]) || 0;
    const weekStr = row[header.indexOf("Tuần")]?.toString().trim();
    const name = row[header.indexOf("Họ và tên")]?.toString().trim();

    // Loại trừ nhân viên không thuộc staffList
    if (!staffList.includes(name)) return;
    if (!date || !type || !product || !detail) return;

    // Kiểm tra ngày được chọn
    let include = false;
    switch(params.reportType) {
      case 'date':
        const targetDate = new Date(params.filterValue);
        include = date.toDateString() === targetDate.toDateString();
        break;
      case 'week':
        const [weekYear, weekMonth, weekNumber] = params.filterValue.split('-');
        const targetWeekMonth = Number(weekMonth);
        const targetWeekYear = Number('20' + weekYear);
        const targetWeek = Number(weekNumber);
        const weekMatch = weekStr && weekStr.match(/Tuần (\d+)/);
        const rowWeek = weekMatch ? parseInt(weekMatch[1]) : 0;
        include = date.getFullYear() === targetWeekYear && (date.getMonth() + 1) === targetWeekMonth && rowWeek === targetWeek;
        break;
      case 'month':
        const [monthYear, monthNumber] = params.filterValue.split('-');
        const targetMonthNumber = Number(monthNumber);
        const targetYearMonth = Number('20' + monthYear);
        include = date.getFullYear() === targetYearMonth && (date.getMonth() + 1) === targetMonthNumber;
        break;
      case 'year':
        const targetYear = Number(params.filterValue);
        include = date.getFullYear() === targetYear;
        break;
    }
    if (!include) return;

    // Tìm sản phẩm chính dựa trên mapping
    let mainProduct = null;
    for (const [mainProd, details] of Object.entries(productMapping)) {
      if (details[detail] === product) {
        mainProduct = mainProd;
        break;
      }
    }
    if (mainProduct && result[mainProduct]?.[detail]) {
      if (params.dataType === 'team') {
        if (type === "BC đầu ngày") {
          result[mainProduct][detail].begin += value;
        } else if (type === "BC cuối ngày") {
          result[mainProduct][detail].end += value;
        }
      } else {
        if (type === "BC cuối ngày" && staffList.includes(name)) {
          result[mainProduct][detail][name] += value;
        }
      }
    }
  });

  return result;
}

// Lấy dữ liệu KPI cho bảng KPI.html
function getKPIData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KPI');
  if (!sheet) return [];
  // Lấy đúng vùng P6:AC14
  const data = sheet.getRange(6, 16, 9, 14).getValues(); // P=16, 9 dòng, 14 cột (P-Q-R...AC)
  const result = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const name = row[0]; // cột P
    if (!name || typeof name !== 'string' || name.trim() === '') continue;
    // months: các cột R-AC, tức là row[2] đến row[13] (bỏ Q là row[1])
    const months = [];
    for (let j = 2; j < row.length; j++) {
      months.push(row[j]);
    }
    result.push({ name: name.trim(), months });
  }
  return result;
}

// Hàm trả dữ liệu cho CKKDreport.html
function getCKKDrawData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CKKDraw');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1);
  // Trả về mảng object, mỗi object là 1 dòng dữ liệu gốc
  return rows.map(row => {
    const obj = {};
    header.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

// Hàm xử lý server side cho CKKDreport.html
function getCKKDataWithFilter(params) {
  const sheetCKKDraw = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CKKDraw');
  const sheetBaoCao = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Báo cáo');
  const sheetCCKD = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CCKD');
  
  if (!sheetCKKDraw || !sheetBaoCao || !sheetCCKD) return [];
  
  const ckkDrawData = sheetCKKDraw.getDataRange().getValues();
  const baoCaoData = sheetBaoCao.getDataRange().getValues();
  
  const ckkDrawHeader = ckkDrawData[0];
  const baoCaoHeader = baoCaoData[0];
  const ckkDrawRows = ckkDrawData.slice(1);
  const baoCaoRows = baoCaoData.slice(1);

  // Định nghĩa tất cả sản phẩm và chi tiết cần hiển thị
  const productDetails = {
    'Lending': ['Trình', 'Phê duyệt', 'Giải ngân'],
    'NFI': ['TKSĐ', 'FX', 'Banca NT', 'Banca PNT', 'Khác'],
    'Thẻ tín dụng': ['Trình', 'Chi tiêu', 'Active'],
    'TKMM': ['Tự mở', 'CTV', 'Activation']
  };

  // Chuẩn bị dữ liệu note từ sheet CCKD
  // Map tuần -> cột
  const weekNoteCol = {
    '1': 9,   // I
    '2': 13,  // M
    '3': 18,  // R
    '4': 23,  // W
    '5': 28   // AB
  };
  // Dòng 19-31 (1-based) => index 18-30 (0-based)
  let notesArr = [];
  let week = params.week && params.week !== '' ? params.week : null;
  if (week && weekNoteCol[week]) {
    try {
      notesArr = sheetCCKD.getRange(19, weekNoteCol[week], 13, 1).getValues().map(r => r[0]);
    } catch (e) {
      notesArr = [];
    }
  }

  const result = [];
  // Xử lý từng sản phẩm và chi tiết
  let rowIdx = 0;
  Object.entries(productDetails).forEach(([product, details]) => {
    details.forEach(detail => {
      // === LỌC DỮ LIỆU KẾ HOẠCH TỪ SHEET CKKDraw ===
      const filteredKeHoach = ckkDrawRows.filter(row => {
        // Loại trừ BM khỏi tất cả báo cáo
        if (row[2] === 'BM') return false;
        // Loại trừ thêm cho Team LD
  if (params.employee === 'Team LD' && (row[2] === 'SRM1' || row[2] === 'GDV1' || row[2] === 'KSV1')) return false;
        // Lọc tháng (cột A)
        if (params.month && params.month !== '') {
          const rowMonth = String(row[0]).replace(/[^0-9]/g, '');
          if (rowMonth !== params.month) return false;
        }
        // Lọc tuần (cột B)
        if (params.week && params.week !== '') {
          const rowWeek = String(row[1]).replace(/[^0-9]/g, '');
          if (rowWeek !== params.week) return false;
        }
        // Lọc nhân viên (cột C)
        if (params.employee && params.employee !== '' && params.employee !== 'Team LD') {
          if (row[2] !== params.employee) return false;
        }
        // Lọc sản phẩm (cột D)
        if (row[3] !== product) return false;
        // Lọc chi tiết sản phẩm (cột E)
        if (row[4] !== detail) return false;
        return true;
      });

      // Tính tổng kế hoạch từ cột "Số cam kết" (cột F)
      const keHoach = filteredKeHoach.reduce((sum, row) => {
        const value = Number(row[5]) || 0;
        return sum + value;
      }, 0);

      // Lấy danh sách người cam kết
      const nguoiCamKetList = filteredKeHoach
        .map(row => row[2]) // Lấy tên nhân viên
        .filter((name, index, arr) => arr.indexOf(name) === index); // Loại bỏ trùng lặp
      const nguoiCamKet = nguoiCamKetList.length > 0 ? nguoiCamKetList.join(', ') : '';

      // Nếu lọc theo nhân viên cụ thể, lấy kế hoạch cam kết từ cột G (index 6) của dòng tương ứng, có thể là text
      let keHoachCamKet = '';
      if (params.employee && params.employee !== '' && params.employee !== 'Team LD') {
        const keHoachRow = ckkDrawRows.find(row =>
          row[2] === params.employee &&
          row[3] === product &&
          row[4] === detail &&
          (!params.month || String(row[0]).replace(/[^0-9]/g, '') === params.month) &&
          (!params.week || String(row[1]).replace(/[^0-9]/g, '') === params.week)
        );
        if (keHoachRow) {
          keHoachCamKet = keHoachRow[6] !== undefined ? keHoachRow[6] : '';
        }
      }

      // === LỌC DỮ LIỆU THỰC HIỆN TỪ SHEET BÁO CÁO ===
      const filteredThucHien = baoCaoRows.filter(row => {
        // Loại trừ BM khỏi tất cả báo cáo
        if (row[2] === 'BM') return false;
        // Loại trừ thêm cho Team LD
  if (params.employee === 'Team LD' && (row[2] === 'SRM1' || row[2] === 'GDV1' || row[2] === 'KSV1')) return false;
        // Lọc tháng (cột I - Month)
        if (params.month && params.month !== '') {
          const rowMonth = String(row[8]).replace(/[^0-9]/g, '');
          if (rowMonth !== params.month) return false;
        }
        // Lọc tuần (cột B) - chỉ cho dữ liệu tuần hiện tại
        if (params.week && params.week !== '') {
          const rowWeek = String(row[1]).replace(/[^0-9]/g, '');
          if (rowWeek !== params.week) return false;
        }
        // Lọc nhân viên (cột C)
        if (params.employee && params.employee !== '' && params.employee !== 'Team LD') {
          if (row[2] !== params.employee) return false;
        }
        // Lọc sản phẩm (cột F)
        if (row[5] !== product) return false;
        // Lọc chi tiết sản phẩm (cột G) với mapping cho Lending
        let rowDetail = row[6];
        if (product === 'Lending' && detail === 'Giải ngân') {
          // Map các chi tiết giải ngân về "Giải ngân"
          if (rowDetail !== 'GN - Thế chấp' && rowDetail !== 'GN - Tín chấp' && rowDetail !== 'GN - Ứng vốn') {
            return false;
          }
        } else {
          // Các trường hợp khác so sánh chính xác
          if (rowDetail !== detail) return false;
        }
        // Chỉ lấy BC cuối ngày (cột E)
        if (row[4] !== 'BC cuối ngày') return false;
        return true;
      });

      // Tính tổng thực hiện từ cột "Giá trị" (cột H)
      const thucHien = filteredThucHien.reduce((sum, row) => {
        const value = Number(row[7]) || 0;
        return sum + value;
      }, 0);

      // === TÍNH DỮ LIỆU LŨY KẾ (từ đầu tháng, không lọc theo tuần) ===
      let luyKe = 0;
      if (params.includeCumulative) {
        const filteredLuyKe = baoCaoRows.filter(row => {
          // Loại trừ BM khỏi tất cả báo cáo
          if (row[2] === 'BM') return false;
          // Loại trừ thêm cho Team LD
          if (params.employee === 'Team LD' && (row[2] === 'SRM1' || row[2] === 'GDV1' || row[2] === 'KSV1')) return false;
          // Lọc tháng (cột I - Month)
          if (params.month && params.month !== '') {
            const rowMonth = String(row[8]).replace(/[^0-9]/g, '');
            if (rowMonth !== params.month) return false;
          }
          // KHÔNG lọc tuần cho dữ liệu lũy kế - lấy tất cả từ đầu tháng
          // Lọc nhân viên (cột C)
          if (params.employee && params.employee !== '' && params.employee !== 'Team LD') {
            if (row[2] !== params.employee) return false;
          }
          // Lọc sản phẩm (cột F)
          if (row[5] !== product) return false;
          // Lọc chi tiết sản phẩm (cột G) với mapping cho Lending
          let rowDetail = row[6];
          if (product === 'Lending' && detail === 'Giải ngân') {
            // Map các chi tiết giải ngân về "Giải ngân"
            if (rowDetail !== 'GN - Thế chấp' && rowDetail !== 'GN - Tín chấp' && rowDetail !== 'GN - Ứng vốn') {
              return false;
            }
          } else {
            // Các trường hợp khác so sánh chính xác
            if (rowDetail !== detail) return false;
          }
          // Chỉ lấy BC cuối ngày (cột E)
          if (row[4] !== 'BC cuối ngày') return false;
          return true;
        });

        // Tính tổng lũy kế từ cột "Giá trị" (cột H)
        luyKe = filteredLuyKe.reduce((sum, row) => {
          const value = Number(row[7]) || 0;
          return sum + value;
        }, 0);
      }

      // === LẤY DỮ LIỆU KẾ HOẠCH CỦA BM TỪ SHEET CKKDraw ===
      let bmKeHoach = 0;
      if (params.includeCumulative) {
        const filteredBMKeHoach = ckkDrawRows.filter(row => {
          // Chỉ lấy dữ liệu của BM
          if (row[2] !== 'BM') return false;
          // Lọc tháng (cột A)
          if (params.month && params.month !== '') {
            const rowMonth = String(row[0]).replace(/[^0-9]/g, '');
            if (rowMonth !== params.month) return false;
          }
          // Lọc sản phẩm (cột D)
          if (row[3] !== product) return false;
          // Lọc chi tiết sản phẩm (cột E)
          if (row[4] !== detail) return false;
          return true;
        });

        // Tính tổng kế hoạch của BM từ cột "Kế hoạch cam kết" (cột I)
        bmKeHoach = filteredBMKeHoach.reduce((sum, row) => {
          const value = Number(row[8]) || 0;
          return sum + value;
        }, 0);
      }

      // Lấy note tương ứng nếu có
      let note = '';
      if (notesArr && notesArr.length > rowIdx) {
        note = notesArr[rowIdx] || '';
      }
      
      result.push({
        product: product,
        detail: detail,
        keHoach: keHoach,
        thucHien: thucHien,
        luyKe: luyKe,
        bmKeHoach: bmKeHoach,
        note: note,
        nguoiCamKet: nguoiCamKet,
        keHoachCamKet: keHoachCamKet
      });
      rowIdx++;
    });
  });

  return result;
}

// Hàm lưu dữ liệu từ form báo cáo cam kết hàng tuần
function saveCKKDraw(dataArray) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CKKDraw');
    if (!sheet) {
      throw new Error('Không tìm thấy sheet CKKDraw');
    }

    // Chuẩn bị dữ liệu để insert
    const rowsToInsert = dataArray.map(item => [
      item.month,           // Cột A: Tháng
      item.week,            // Cột B: Tuần
      item.name,            // Cột C: Họ và tên
      item.product,         // Cột D: Sản phẩm
      item.detail,          // Cột E: Chi tiết sản phầm
      item.soCamKet || '',  // Cột F: Số cam kết
      item.keHoachCamKet || '', // Cột G: Kế hoạch cam kết (di chuyển từ I sang G)
      '',                   // Cột H: Để trống
      ''                    // Cột I: Để trống
    ]);

    // Insert dữ liệu vào đầu sheet (sau header)
    if (rowsToInsert.length > 0) {
      sheet.insertRowsBefore(2, rowsToInsert.length);
      sheet.getRange(2, 1, rowsToInsert.length, 9).setValues(rowsToInsert);
    }

    return { success: true, message: `Đã lưu ${rowsToInsert.length} dòng dữ liệu thành công` };
  } catch (error) {
    console.error('Lỗi khi lưu dữ liệu:', error);
    throw new Error(`Lỗi khi lưu dữ liệu: ${error.message}`);
  }
}

function openWeeklyReport() {
  const html = HtmlService.createHtmlOutputFromFile("weekly_report")
    .setWidth(800)
    .setHeight(700)
    .setTitle("📋 Báo Cáo Cam Kết Hàng Tuần");
  SpreadsheetApp.getUi().showModalDialog(html, "📋 BÁO CÁO CAM KẾT HÀNG TUẦN");
}

// Thêm hàm mở CKKDreport.html
function openCKKDReport() {
  const html = HtmlService.createHtmlOutputFromFile("CKKDreport")
    .setWidth(1400)
    .setHeight(800)
    .setTitle("📋 Cam kết tuần");
  SpreadsheetApp.getUi().showModalDialog(html, "📋 Cam kết tuần");
}

// Thêm hàm mở KPI.html
function openKPIReport() {
  const html = HtmlService.createHtmlOutputFromFile("KPI")
    .setWidth(1200)
    .setHeight(800)
    .setTitle("📈 KPI");
  SpreadsheetApp.getUi().showModalDialog(html, "📈 KPI");
}


/**
 * Nhận dữ liệu từ SCAN.html và ghi vào sheet 'SCAN' (A-Q)
 */
function submitSCANForm(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SCAN');
  if (!sheet) {
    SpreadsheetApp.getActiveSpreadsheet().insertSheet('SCAN');
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SCAN');
  }
  var values = sheet.getDataRange().getValues();
  var foundRow = -1;
  var colCIF = 0; // Cột A
  var colName = 1; // Cột B
  var cif = data.cif ? data.cif.trim() : '';
  var hoten = data.hoten ? data.hoten.trim().toLowerCase() : '';
  // Ưu tiên tìm trùng cả CIF và Họ và Tên
  for (var i = 1; i < values.length; i++) {
    var rowCIF = values[i][colCIF] ? values[i][colCIF].toString().trim() : '';
    var rowName = values[i][colName] ? values[i][colName].toString().trim().toLowerCase() : '';
    if (cif && hoten && rowCIF === cif && rowName === hoten) {
      foundRow = i + 1;
      break;
    }
  }
  // Nếu không có, tìm trùng CIF
  if (foundRow === -1 && cif) {
    for (var i = 1; i < values.length; i++) {
      var rowCIF = values[i][colCIF] ? values[i][colCIF].toString().trim() : '';
      if (rowCIF === cif) {
        foundRow = i + 1;
        break;
      }
    }
  }
  // Nếu không có, tìm trùng Họ và Tên
  if (foundRow === -1 && hoten) {
    var nameMatches = [];
    for (var i = 1; i < values.length; i++) {
      var rowName = values[i][colName] ? values[i][colName].toString().trim().toLowerCase() : '';
      if (rowName === hoten) {
        nameMatches.push(i + 1);
      }
    }
    if (nameMatches.length === 1) {
      foundRow = nameMatches[0];
    } else if (nameMatches.length > 1) {
      throw new Error('Có nhiều khách hàng trùng tên, vui lòng nhập thêm CIF để xác định chính xác.');
    }
  }
  var rowData = [
    data.cif || '',
    data.hoten || '',
    data.chuyenvien || '',
    data.thang || '',
    data.tuan || '',
    data.lead_trongtam || '',
    data.trangthai || '',
    data.ghichu_trangthai || '',
    data.nhankhau || '',
    data.taichinh || '',
    data.hanhvi_tc || '',
    data.hanhvi_cuocsong || '',
    data.du_dinh || '',
    data.kichban || '',
    data.taisao_mua || '',
    data.case_size || '',
    data.laytien || '',
    data.taisao_now || '',
    data.nextstep || '',
    data.ketqua || '',
    data.ghichu || ''
  ];
  if (foundRow > 0) {
    // Lấy dữ liệu dòng cũ
    var oldRow = sheet.getRange(foundRow, 1, 1, rowData.length).getValues()[0];
    // Chỉ cập nhật ô có dữ liệu mới, giữ nguyên ô cũ nếu dữ liệu mới rỗng
    for (var j = 0; j < rowData.length; j++) {
      if (rowData[j] === '' || rowData[j] === null || rowData[j] === undefined) {
        rowData[j] = oldRow[j];
      }
    }
    sheet.getRange(foundRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
}

// === API cho CheckInputReport.html: Kiểm tra nhập cam kết tuần ===
function getWeeklyCommitCheck(month) {
  // Lấy danh sách nhân viên từ sheet CKKDraw (cột C, loại bỏ BM)
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CKKDraw');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1);
  // Lấy danh sách nhân viên duy nhất (không BM)
  const staffSet = new Set();
  rows.forEach(row => {
    const name = row[2];
    if (name && name !== 'BM') staffSet.add(name);
  });
  const staffList = Array.from(staffSet);
  // Chuẩn bị kết quả
  const result = staffList.map(name => ({
    name,
    week1: false,
    week2: false,
    week3: false,
    week4: false,
    week5: false
  }));
  // Duyệt từng dòng, đánh dấu tuần đã nhập cho từng nhân viên trong tháng
  rows.forEach(row => {
    const rowMonth = String(row[0]).replace(/[^0-9]/g, '');
    const rowWeek = String(row[1]).replace(/[^0-9]/g, '');
    const name = row[2];
    if (!name || name === 'BM') return;
    if (rowMonth !== String(month)) return;
    const weekNum = parseInt(rowWeek);
    if (weekNum >= 1 && weekNum <= 5) {
      const staff = result.find(r => r.name === name);
      if (staff) staff['week' + weekNum] = true;
    }
  });
  return result;
}
