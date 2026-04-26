# BAH.github.io
// ═══════════════════════════════════════════════════════════════════════════════
// BÌNH AN HOUSE – Auth Backend (Google Apps Script)
// Deploy: Extensions → Apps Script → Deploy → New deployment → Web app
//         Execute as: Me | Who has access: Anyone
// ═══════════════════════════════════════════════════════════════════════════════

// ── CONFIG ────────────────────────────────────────────────────────────────────
const ADMIN_EMAIL   = 'binhanhome8686@gmail.com';
const SHEET_NAME    = 'Accounts';         // Tab tên Accounts trong Google Sheet
const OTP_SHEET     = 'OTP_Log';          // Tab lưu OTP tạm
const SPREADSHEET_ID = '1hqmhXXet6kbdidk3sCqH9Bl6X8VktfgLdP8FCTar_3g';  // Thay bằng ID Google Sheet của bạn
const OTP_EXPIRE_MS  = 5 * 60 * 1000;    // 5 phút

// ── ENTRY POINT ───────────────────────────────────────────────────────────────
function doPost(e) {
  const cors = ContentService.createTextOutput();
  cors.setMimeType(ContentService.MimeType.JSON);
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;
    if      (action === 'sendOTP')    result = handleSendOTP(body.email);
    else if (action === 'verifyOTP')  result = handleVerifyOTP(body.email, body.otp);
    else if (action === 'register')   result = handleRegister(body);
    else if (action === 'approveUser')result = handleApproveUser(body);
    else if (action === 'rejectUser') result = handleRejectUser(body);
    else if (action === 'addUser')    result = handleAddUser(body);
    else if (action === 'listUsers')  result = handleListUsers(body.requester);
    else if (action === 'removeUser') result = handleRemoveUser(body);
    else result = {ok: false, msg: 'Unknown action'};
    cors.setContent(JSON.stringify(result));
  } catch(err) {
    cors.setContent(JSON.stringify({ok: false, msg: err.message}));
  }
  return cors;
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ok:true,msg:'BAH Auth API running'}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── SEND OTP ──────────────────────────────────────────────────────────────────
function handleSendOTP(email) {
  if (!email) return {ok: false, msg: 'Email không hợp lệ'};

  // Admin luôn được phép
  const isAdmin = email.toLowerCase() === ADMIN_EMAIL.toLowerCase();
  if (!isAdmin) {
    // Kiểm tra email có trong danh sách tài khoản không
    const accounts = getAccounts();
    const found = accounts.find(a => a.email.toLowerCase() === email.toLowerCase());
    if (!found) return {ok: false, msg: 'Email chưa được cấp quyền. Liên hệ admin.'};
    if (found.status !== 'active') return {ok: false, msg: 'Tài khoản đã bị vô hiệu hoá.'};
  }

  // Tạo OTP 6 chữ số
  const otp = String(Math.floor(100000 + Math.random() * 900000));
  const expiry = new Date(Date.now() + OTP_EXPIRE_MS).toISOString();

  // Lưu OTP vào sheet
  saveOTP(email, otp, expiry);

  // Gửi email
  const subject = '🔐 Mã OTP đăng nhập – Bình An House';
  const body = `
Xin chào,

Mã OTP đăng nhập vào hệ thống Bình An House của bạn là:

━━━━━━━━━━━━━━━━
  ${otp}
━━━━━━━━━━━━━━━━

Mã có hiệu lực trong 5 phút.
Nếu bạn không yêu cầu đăng nhập, hãy bỏ qua email này.

Bình An House
136 Trương Định, Sơn Trà, Đà Nẵng
`;
  GmailApp.sendEmail(email, subject, body, {
    from: ADMIN_EMAIL,
    name: 'Bình An House – System'
  });

  return {ok: true, msg: 'OTP đã được gửi'};
}

// ── VERIFY OTP ────────────────────────────────────────────────────────────────
function handleVerifyOTP(email, otp) {
  if (!email || !otp) return {ok: false, msg: 'Thiếu thông tin'};
  const record = getOTP(email);
  if (!record) return {ok: false, msg: 'Không tìm thấy OTP. Vui lòng yêu cầu lại.'};
  if (new Date() > new Date(record.expiry)) return {ok: false, msg: 'Mã OTP đã hết hạn.'};
  if (record.otp !== otp.trim()) return {ok: false, msg: 'Mã OTP không đúng.'};

  // Xoá OTP sau khi dùng
  deleteOTP(email);

  // Trả về role
  const isAdmin = email.toLowerCase() === ADMIN_EMAIL.toLowerCase();
  if (isAdmin) return {ok: true, role: 'admin', name: 'Admin'};

  const accounts = getAccounts();
  const user = accounts.find(a => a.email.toLowerCase() === email.toLowerCase());
  return {ok: true, role: 'staff', name: user ? user.name : email};
}

// ── ADD USER (admin only) ─────────────────────────────────────────────────────
function handleAddUser(body) {
  if (body.requester?.toLowerCase() !== ADMIN_EMAIL.toLowerCase())
    return {ok: false, msg: 'Chỉ admin mới có quyền thêm tài khoản'};

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Email','Name','Role','Status','CreatedAt']);
  }

  // Kiểm tra trùng
  const accounts = getAccounts();
  if (accounts.find(a => a.email.toLowerCase() === body.email.toLowerCase()))
    return {ok: false, msg: 'Email này đã tồn tại'};

  sheet.appendRow([
    body.email,
    body.name || '',
    'staff',
    'active',
    new Date().toISOString()
  ]);

  // Gửi email thông báo cho nhân viên
  GmailApp.sendEmail(body.email,
    'Bạn đã được cấp quyền truy cập – Bình An House',
    `Xin chào ${body.name || ''},\n\nBạn đã được Admin cấp quyền truy cập hệ thống hoá đơn Bình An House.\nĐăng nhập tại link: ${body.appUrl || ''}\n\nBình An House`,
    {from: ADMIN_EMAIL, name: 'Bình An House'}
  );

  return {ok: true, msg: 'Đã thêm tài khoản và gửi email thông báo'};
}

// ── LIST USERS (admin only) ───────────────────────────────────────────────────
function handleListUsers(requester) {
  if (requester?.toLowerCase() !== ADMIN_EMAIL.toLowerCase())
    return {ok: false, msg: 'Chỉ admin mới có quyền xem danh sách'};
  return {ok: true, users: getAccounts()};
}

// ── REMOVE USER (admin only) ──────────────────────────────────────────────────
function handleRemoveUser(body) {
  if (body.requester?.toLowerCase() !== ADMIN_EMAIL.toLowerCase())
    return {ok: false, msg: 'Chỉ admin mới có quyền xoá tài khoản'};

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return {ok: false, msg: 'Sheet không tồn tại'};

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toLowerCase() === body.email.toLowerCase()) {
      sheet.deleteRow(i + 1);
      return {ok: true, msg: 'Đã xoá tài khoản'};
    }
  }
  return {ok: false, msg: 'Không tìm thấy tài khoản'};
}

// ── HELPERS: Accounts ─────────────────────────────────────────────────────────
function getAccounts() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return [];
  const [, ...rows] = sheet.getDataRange().getValues();
  return rows.map(r => ({email: r[0], name: r[1], role: r[2], status: r[3]}));
}

// ── HELPERS: OTP ──────────────────────────────────────────────────────────────
function saveOTP(email, otp, expiry) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet   = ss.getSheetByName(OTP_SHEET);
  if (!sheet) { sheet = ss.insertSheet(OTP_SHEET); sheet.appendRow(['Email','OTP','Expiry']); }
  // Xoá OTP cũ của email này trước
  deleteOTP(email, sheet);
  sheet.appendRow([email, otp, expiry]);
}

function getOTP(email) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(OTP_SHEET);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toLowerCase() === email.toLowerCase())
      return {otp: String(data[i][1]), expiry: data[i][2]};
  }
  return null;
}

function deleteOTP(email, sheet) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  sheet = sheet || ss.getSheetByName(OTP_SHEET);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0].toLowerCase() === email.toLowerCase()) sheet.deleteRow(i + 1);
  }
}

// ── REGISTER (nhân viên tự đăng ký, chờ admin duyệt) ─────────────────────────
function handleRegister(body) {
  const { name, email, phone, jobRole } = body;
  if (!email || !name) return {ok: false, msg: 'Thiếu thông tin bắt buộc'};

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Email','Name','Role','Status','Phone','JobRole','CreatedAt','ApprovedAt']);
  }

  // Kiểm tra email đã tồn tại chưa
  const accounts = getAccounts();
  const existing = accounts.find(a => a.email.toLowerCase() === email.toLowerCase());
  if (existing) {
    if (existing.status === 'pending')
      return {ok: false, msg: 'Email này đã gửi yêu cầu, đang chờ admin duyệt'};
    if (existing.status === 'active')
      return {ok: false, msg: 'Email này đã có tài khoản. Vui lòng đăng nhập'};
    return {ok: false, msg: 'Email đã tồn tại trong hệ thống'};
  }

  // Lưu với status = pending
  sheet.appendRow([
    email, name, 'staff', 'pending',
    phone || '', jobRole || '',
    new Date().toISOString(), ''
  ]);

  // Tạo link duyệt và từ chối cho admin
  const approveLink = `${ScriptApp.getService().getUrl()}?action=approveUser&email=${encodeURIComponent(email)}&adminKey=BAH2026`;
  const rejectLink  = `${ScriptApp.getService().getUrl()}?action=rejectUser&email=${encodeURIComponent(email)}&adminKey=BAH2026`;

  // Email thông báo cho admin
  GmailApp.sendEmail(ADMIN_EMAIL,
    `[Bình An House] Yêu cầu đăng ký mới – ${name}`,
    `Có yêu cầu đăng ký tài khoản mới:\n\nHọ tên:   ${name}\nEmail:    ${email}\nSĐT:      ${phone || '—'}\nChức vụ:  ${jobRole || '—'}\nThời gian: ${new Date().toLocaleString('vi-VN')}\n\n✅ DUYỆT: ${approveLink}\n\n❌ TỪ CHỐI: ${rejectLink}\n\nHoặc vào bảng Admin trên ứng dụng để quản lý.\n\nBình An House System`,
    { name: 'Bình An House – System' }
  );

  return {ok: true, msg: 'Đã gửi yêu cầu. Admin sẽ xem xét và phản hồi qua email.'};
}

// ── APPROVE USER ──────────────────────────────────────────────────────────────
function handleApproveUser(body) {
  // Hỗ trợ cả GET (click từ email) và POST (từ admin panel)
  const requester = body.requester || body.adminKey;
  const isAdmin   = requester === ADMIN_EMAIL || requester === 'BAH2026';
  if (!isAdmin) return {ok: false, msg: 'Chỉ admin mới có quyền duyệt'};

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return {ok: false, msg: 'Sheet không tồn tại'};

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toLowerCase() === body.email.toLowerCase()) {
      sheet.getRange(i + 1, 4).setValue('active');        // Status → active
      sheet.getRange(i + 1, 8).setValue(new Date().toISOString()); // ApprovedAt
      // Gửi email thông báo cho nhân viên
      GmailApp.sendEmail(body.email,
        '✅ Tài khoản đã được duyệt – Bình An House',
        `Xin chào ${data[i][1]},\n\nTài khoản của bạn đã được Admin phê duyệt.\nBạn có thể đăng nhập ngay bằng email này.\n\nChúc bạn làm việc hiệu quả!\n\nBình An House\n136 Trương Định, Sơn Trà, Đà Nẵng`,
        { from: ADMIN_EMAIL, name: 'Bình An House' }
      );
      return {ok: true, msg: `Đã duyệt tài khoản ${body.email}`};
    }
  }
  return {ok: false, msg: 'Không tìm thấy tài khoản'};
}

// ── REJECT USER ───────────────────────────────────────────────────────────────
function handleRejectUser(body) {
  const requester = body.requester || body.adminKey;
  const isAdmin   = requester === ADMIN_EMAIL || requester === 'BAH2026';
  if (!isAdmin) return {ok: false, msg: 'Chỉ admin mới có quyền từ chối'};

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return {ok: false, msg: 'Sheet không tồn tại'};

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toLowerCase() === body.email.toLowerCase()) {
      sheet.deleteRow(i + 1);
      GmailApp.sendEmail(body.email,
        'Yêu cầu đăng ký – Bình An House',
        `Xin chào ${data[i][1]},\n\nRất tiếc, yêu cầu đăng ký của bạn chưa được chấp thuận lần này.\nNếu có thắc mắc, vui lòng liên hệ: ${ADMIN_EMAIL}\n\nBình An House`,
        { from: ADMIN_EMAIL, name: 'Bình An House' }
      );
      return {ok: true, msg: `Đã từ chối và xoá yêu cầu của ${body.email}`};
    }
  }
  return {ok: false, msg: 'Không tìm thấy tài khoản'};
}

// ── doGet: xử lý click link duyệt/từ chối từ email ───────────────────────────
function doGet(e) {
  const action = e.parameter.action;
  const email  = e.parameter.email;
  const key    = e.parameter.adminKey;

  if (action === 'approveUser' && email) {
    const result = handleApproveUser({email, adminKey: key});
    return HtmlService.createHtmlOutput(
      `<h2 style="font-family:Arial">${result.ok ? '✅ ' : '❌ '}${result.msg}</h2>
       <p style="font-family:Arial;color:#666">Bạn có thể đóng tab này.</p>`
    );
  }
  if (action === 'rejectUser' && email) {
    const result = handleRejectUser({email, adminKey: key});
    return HtmlService.createHtmlOutput(
      `<h2 style="font-family:Arial">${result.ok ? '✅ ' : '❌ '}${result.msg}</h2>
       <p style="font-family:Arial;color:#666">Bạn có thể đóng tab này.</p>`
    );
  }

  return ContentService.createTextOutput(JSON.stringify({ok:true, msg:'BAH Auth API running'}))
    .setMimeType(ContentService.MimeType.JSON);
}
