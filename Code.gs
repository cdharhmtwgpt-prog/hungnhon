/**
 * SỔ THU CHI — PHÒNG BẤM HUYỆT HƯNG NHƠN
 * Backend Google Apps Script: đọc & ghi dữ liệu vào Google Sheet.
 *
 * ============ CÀI ĐẶT (làm 1 lần) ============
 * 1. Mở Google Sheet dùng làm sổ sách (Sheet đang dùng, hoặc tạo Sheet mới).
 * 2. Menu Extensions > Apps Script.
 * 3. Xoá code mẫu (myFunction rỗng), dán TOÀN BỘ nội dung file này vào.
 * 4. Ở thanh công cụ, chọn hàm "setupSpreadsheet" trong dropdown > bấm Run.
 *    - Lần đầu chạy, Google sẽ hỏi cấp quyền (Authorize) — chọn tài khoản,
 *      bấm "Advanced" > "Go to ... (unsafe)" nếu hiện cảnh báo, rồi Allow.
 *    - Hàm này tạo tab tháng hiện tại nếu chưa có, KHÔNG đụng vào tab cũ.
 * 5. Deploy > New deployment > bánh răng chọn "Web app":
 *    - Description: tuỳ ý
 *    - Execute as: Me (tài khoản của em)
 *    - Who has access: Anyone
 *    Bấm Deploy > copy "Web app URL" (dạng https://script.google.com/.../exec)
 *    → dán vào biến APPS_SCRIPT_URL trong file index.html.
 * 6. Mỗi khi SỬA lại code này, phải Deploy > Manage deployments > bấm nút sửa
 *    (bút chì) > chọn "New version" > Deploy lại thì URL mới nhận code mới
 *    (URL giữ nguyên, không đổi).
 *
 * ============ QUY ƯỚC DỮ LIỆU (đọc kỹ trước khi dùng thật) ============
 * - Mỗi THÁNG là 1 tab riêng, đặt tên theo mẫu "MM-YYYY", ví dụ "08-2026".
 *   → Nếu Sheet hiện tại của em đang đặt tên tab kiểu khác (vd "Thang 8",
 *     "8/2026"...), sửa lại hàm getSheetNameForDate() bên dưới cho khớp,
 *     hoặc đổi tên tab cũ theo đúng mẫu "MM-YYYY".
 * - Dòng 1: ô tổng (Tổng doanh thu / Tổng chi tiêu / Lãi-Lỗ) — dùng công
 *   thức SUM tự động cộng dồn, không cần tính tay.
 * - Dòng 2: tiêu đề cột. Dữ liệu bắt đầu từ dòng 3.
 * - Cột A = số ngày trong tháng (1–31), B = Doanh thu, C = Chi tiêu,
 *   D = Ghi chú (nội dung chi, hoặc "Lương <Tên>").
 * - MỖI LẦN LƯU TRÊN APP = 1 DÒNG MỚI (1 giao dịch), kể cả khi trùng ngày.
 *   → Khác với cách gộp 1-dòng-1-ngày trước đây, nhưng an toàn hơn: không
 *     bao giờ mất dữ liệu khi 1 ngày có nhiều khoản chi khác nhau, và Tổng
 *     ở dòng 1 vẫn đúng vì dùng công thức SUM cộng hết các dòng.
 *   → Nếu em muốn giữ đúng kiểu "1 dòng 1 ngày" như file cũ, báo lại để anh
 *     sửa doPost() thành tìm-dòng-của-ngày-đó rồi cộng dồn thay vì thêm dòng.
 */

const STAFF_LIST = ['Chị Vân', 'Phát', 'Phương']; // đổi nếu đổi nhân sự

const SUMMARY_ROW = 1;
const HEADER_ROW = 2;
const FIRST_DATA_ROW = 3;

// ---------- Helpers ----------

function pad2_(n) {
  return (n < 10 ? '0' : '') + n;
}

// Tên tab cho 1 ngày, dựa vào chuỗi 'YYYY-MM-DD' (không dùng đối tượng Date
// để tránh lệch múi giờ giữa máy chủ Apps Script và VN).
function getSheetNameForDateStr_(ymd) {
  const parts = String(ymd).split('-');
  const year = parts[0];
  const month = parts[1];
  return month + '-' + year;
}

function getSheetNameForDate(dateObj) {
  return pad2_(dateObj.getMonth() + 1) + '-' + dateObj.getFullYear();
}

function getOrCreateMonthSheet_(ss, sheetName) {
  let sh = ss.getSheetByName(sheetName);
  if (sh) return sh;
  sh = ss.insertSheet(sheetName);
  sh.getRange(SUMMARY_ROW, 1, 1, 6).setValues([[
    'Tổng doanh thu', '=SUM(B' + FIRST_DATA_ROW + ':B9999)',
    'Tổng chi tiêu', '=SUM(C' + FIRST_DATA_ROW + ':C9999)',
    'Lãi/Lỗ', '=B1-D1'
  ]]);
  sh.getRange(HEADER_ROW, 1, 1, 4).setValues([['Ngày', 'Doanh thu', 'Chi tiêu', 'Lãi/Lỗ hoặc Ghi chú']]);
  sh.getRange(SUMMARY_ROW, 1, 1, 6).setFontWeight('bold');
  sh.getRange(HEADER_ROW, 1, 1, 4).setFontWeight('bold');
  sh.setFrozenRows(HEADER_ROW);
  sh.setColumnWidths(1, 4, 140);
  return sh;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- Chạy tay 1 lần lúc cài đặt ----------

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  getOrCreateMonthSheet_(ss, getSheetNameForDate(new Date()));
}

// ---------- API cho app ----------

// GET .../exec?months=08-2026,07-2026
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const monthsParam = (e.parameter.months || '').split(',').map(s => s.trim()).filter(Boolean);
  const targetMonths = monthsParam.length ? monthsParam : [getSheetNameForDate(new Date())];

  const result = {};
  targetMonths.forEach(function (m) {
    const sh = ss.getSheetByName(m);
    if (!sh) { result[m] = { exists: false, rows: [] }; return; }
    const lastRow = sh.getLastRow();
    const rows = [];
    if (lastRow >= FIRST_DATA_ROW) {
      const values = sh.getRange(FIRST_DATA_ROW, 1, lastRow - FIRST_DATA_ROW + 1, 4).getValues();
      values.forEach(function (r) {
        if (r[0] === '' || r[0] === null) return;
        rows.push({ day: r[0], thu: Number(r[1]) || 0, chi: Number(r[2]) || 0, note: r[3] || '' });
      });
    }
    result[m] = { exists: true, rows: rows };
  });

  return jsonOut_({ ok: true, months: result, staff: STAFF_LIST, serverMonth: getSheetNameForDate(new Date()) });
}

// POST body (text/plain, nội dung là JSON — xem ghi chú CORS trong index.html):
// { date: 'YYYY-MM-DD', type: 'thu'|'chi', amount: number, note?: string, staff?: string }
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (!body.date || !body.type || !body.amount || Number(body.amount) <= 0) {
      return jsonOut_({ ok: false, error: 'Thiếu dữ liệu hoặc số tiền không hợp lệ' });
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = getSheetNameForDateStr_(body.date);
    const sh = getOrCreateMonthSheet_(ss, sheetName);

    const day = Number(String(body.date).split('-')[2]);
    const thu = body.type === 'thu' ? Number(body.amount) : '';
    const chi = body.type === 'chi' ? Number(body.amount) : '';
    const note = body.type === 'chi' ? (body.note || '') : '';

    sh.appendRow([day, thu, chi, note]);

    return jsonOut_({ ok: true, sheet: sheetName });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}
