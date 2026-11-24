/**
 * KONFIGURASI DEFAULT APLIKASI CETAK MMT
 */
const DEFAULT_CONFIG = {
  ADMIN_EMAIL: '', // Kosongkan agar otomatis memakai email pemilik script, atau isi email admin
  ADMIN_PIN: '123456',
  BCA_ACCOUNT: '0153829104',
  BCA_NAME: 'SOLOPRINT DIGITAL',
  QRIS_IMAGE_URL: '',
  AUTO_APPROVE_USER: 'YA',
  APP_NAME: 'SoloPrint',
  SHEET_ORDERS: 'Data_Pesanan',
  SHEET_SETTINGS: 'Pengaturan',
  SHEET_PRODUCTS: 'Produk_Bahan',
  SHEET_USERS: 'Data_Pelanggan',
  SHEET_CHATS: 'Chat_CS',
  DRIVE_FOLDER_NAME: 'File_Desain_Cetak_MMT',
  MAX_EMAIL_ATTACHMENT_MB: 20
};

/**
 * Entry Point Web App Google Apps Script
 * - Buka URL utama -> menampilkan halaman pelanggan (index.html)
 * - Buka URL + ?page=admin -> menampilkan panel admin khusus (panel_admin.html)
 */
function doGet(e) {
  try {
    ensureDatabaseInitialized();
  } catch (err) {}

  const page = (e && e.parameter && e.parameter.page) ? String(e.parameter.page).toLowerCase() : '';
  const templateFile = (page === 'admin') ? 'panel_admin' : 'index';
  let appName = DEFAULT_CONFIG.APP_NAME;
  try {
    appName = String(getAppSettings().APP_NAME || DEFAULT_CONFIG.APP_NAME).trim() || DEFAULT_CONFIG.APP_NAME;
  } catch (brandErr) {}
  const pageTitle = (page === 'admin')
    ? 'Admin Panel - ' + appName
    : appName + ' - Jasa Print MMT, Banner & Baliho Online';

  let bootPayload = {};
  try {
    if (page === 'admin') {
      // Jangan mengirim data admin/PIN ke browser sebelum login.
      // Semua data privat baru dimuat setelah adminGetAllData memvalidasi PIN.
      bootPayload = { success: true, requiresLogin: true, appName: appName };
    } else {
      bootPayload = getPublicPaymentConfig();
    }
  } catch (bootErr) {
    bootPayload = { bootError: String(bootErr && bootErr.message ? bootErr.message : bootErr) };
  }

  const tmpl = HtmlService.createTemplateFromFile(templateFile);
  tmpl.bootDataJson = JSON.stringify(bootPayload).replace(/<\//g, '<\\/');

  return tmpl.evaluate()
    .setTitle(pageTitle)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * Jalur cadangan HTTP POST agar semua tombol tetap berjalan meski google.script.run terhambat cookie browser HP
 */
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = String(body.fn || '');
    const args = Array.isArray(body.args) ? body.args : [];
    const allowedFns = {
      pingServer: pingServer,
      getPublicPaymentConfig: getPublicPaymentConfig,
      registerCustomer: registerCustomer,
      loginCustomer: loginCustomer,
      submitOrderToGAS: submitOrderToGAS,
      submitPaymentProof: submitPaymentProof,
      trackCustomerOrder: trackCustomerOrder,
      getCustomerChatHistory: getCustomerChatHistory,
      sendCustomerChat: sendCustomerChat,
      adminGetAllData: adminGetAllData,
      adminUpdateOrderStatus: adminUpdateOrderStatus,
      adminDeleteOrder: adminDeleteOrder,
      adminSaveUser: adminSaveUser,
      adminDeleteUser: adminDeleteUser,
      adminSaveProduct: adminSaveProduct,
      adminDeleteProduct: adminDeleteProduct,
      adminUpdateSettings: adminUpdateSettings,
      adminGetAllChats: adminGetAllChats,
      adminReplyChat: adminReplyChat,
      adminMarkChatRead: adminMarkChatRead
    };

    if (!allowedFns[fn]) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Fungsi tidak ditemukan: ' + fn }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const result = allowedFns[fn].apply(null, args);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Server error: ' + err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Mendapatkan Spreadsheet utama (baik terikat ke Google Sheets maupun Standalone Script)
 */
function getMainSpreadsheet() {
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (ss) return ss;

  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('MMT_SPREADSHEET_ID');
  if (savedId) {
    try {
      ss = SpreadsheetApp.openById(savedId);
      if (ss) return ss;
    } catch (e) {}
  }

  ss = SpreadsheetApp.create('Database_Order_Cetak_MMT');
  props.setProperty('MMT_SPREADSHEET_ID', ss.getId());
  return ss;
}

/**
 * Tes koneksi ringan antara frontend dan server GAS
 */
function pingServer() {
  const ss = ensureDatabaseInitialized();
  return {
    ok: true,
    spreadsheetUrl: ss.getUrl(),
    serverTime: Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd MMM yyyy HH:mm:ss')
  };
}

/**
 * Memastikan seluruh tab database & header sudah tersedia
 */
function ensureDatabaseInitialized() {
  const ss = getMainSpreadsheet();

  // 1. Sheet Data Pesanan
  let orderSheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_ORDERS);
  if (!orderSheet) {
    orderSheet = ss.insertSheet(DEFAULT_CONFIG.SHEET_ORDERS);
    const headers = [
      'Kode Invoice', 'Tanggal Order', 'Cabang', 'Nama Pemesan', 'No. WhatsApp',
      'Pengiriman', 'Metode Bayar', 'Total Tagihan', 'Status Pesanan',
      'Link File Desain (Drive)', 'Detail Pesanan (JSON)', 'Ringkasan Item',
      'Bukti Pembayaran (Drive)', 'ID Pelanggan'
    ];
    orderSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    orderSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#FEE2E2');
    orderSheet.setFrozenRows(1);
  }

  // 2. Sheet Pengaturan
  let settingSheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_SETTINGS);
  if (!settingSheet) {
    settingSheet = ss.insertSheet(DEFAULT_CONFIG.SHEET_SETTINGS);
    let defaultEmail = DEFAULT_CONFIG.ADMIN_EMAIL;
    try {
      if (!defaultEmail) defaultEmail = Session.getActiveUser().getEmail() || 'admin@percetakan.com';
    } catch (e) {
      defaultEmail = 'admin@percetakan.com';
    }

    const defaultRows = [
      ['Kunci Pengaturan', 'Nilai'],
      ['ADMIN_EMAIL', defaultEmail],
      ['ADMIN_PIN', DEFAULT_CONFIG.ADMIN_PIN],
      ['BCA_ACCOUNT', DEFAULT_CONFIG.BCA_ACCOUNT],
      ['BCA_NAME', DEFAULT_CONFIG.BCA_NAME],
      ['QRIS_IMAGE_URL', DEFAULT_CONFIG.QRIS_IMAGE_URL],
      ['AUTO_APPROVE_USER', DEFAULT_CONFIG.AUTO_APPROVE_USER],
      ['APP_NAME', DEFAULT_CONFIG.APP_NAME]
    ];
    settingSheet.getRange(1, 1, defaultRows.length, 2).setValues(defaultRows);
    settingSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#F3F4F6');
  }

  // 3. Sheet Produk & Harga Bahan
  let prodSheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_PRODUCTS);
  if (!prodSheet) {
    prodSheet = ss.insertSheet(DEFAULT_CONFIG.SHEET_PRODUCTS);
    const prodRows = [
      ['ID Produk', 'Kategori', 'Nama Bahan / Produk', 'Harga per m2', 'Status'],
      ['PROD-001', 'outdoor', 'Flexy 380 matte', 46667, 'AKTIF'],
      ['PROD-002', 'outdoor', 'Flexy 280 gr (Ekonomis)', 25000, 'AKTIF'],
      ['PROD-003', 'outdoor', 'Flexy 340 gr (Standar Tebal)', 35000, 'AKTIF'],
      ['PROD-004', 'outdoor', 'Flexy Korea 440 gr High-Res', 60000, 'AKTIF'],
      ['PROD-005', 'outdoor', 'Backlite Outdoor (Neon Box)', 85000, 'AKTIF'],
      ['PROD-006', 'indoor', 'Albatros + Laminasi', 75000, 'AKTIF'],
      ['PROD-007', 'indoor', 'Luster Photo Paper', 80000, 'AKTIF'],
      ['PROD-008', 'indoor', 'Sticker Vinyl Indoor', 70000, 'AKTIF'],
      ['PROD-009', 'indoor', 'Art Canvas Indoor', 135000, 'AKTIF'],
      ['PROD-010', 'uv', 'UV Flexy Korea Super Tajam', 95000, 'AKTIF'],
      ['PROD-011', 'uv', 'UV Backlite Film Double Side', 140000, 'AKTIF'],
      ['PROD-012', 'uv', 'UV Sticker Transparan / Ritrama', 120000, 'AKTIF']
    ];
    prodSheet.getRange(1, 1, prodRows.length, 5).setValues(prodRows);
    prodSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#FEF3C7');
    prodSheet.setFrozenRows(1);
  }

  // 4. Sheet Data Pelanggan
  let userSheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_USERS);
  if (!userSheet) {
    userSheet = ss.insertSheet(DEFAULT_CONFIG.SHEET_USERS);
    const userHeaders = ['ID Pelanggan', 'Tanggal Daftar', 'Nama Lengkap / Usaha', 'No. WhatsApp', 'Email', 'Alamat', 'PIN Login', 'Status Akun'];
    userSheet.getRange(1, 1, 1, userHeaders.length).setValues([userHeaders]);
    userSheet.getRange(1, 1, 1, userHeaders.length).setFontWeight('bold').setBackground('#DBEAFE');
    userSheet.setFrozenRows(1);
  }

  // 5. Sheet Live Chat CS
  let chatSheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_CHATS);
  if (!chatSheet) {
    chatSheet = ss.insertSheet(DEFAULT_CONFIG.SHEET_CHATS);
    const chatHeaders = ['ID Pesan', 'Waktu', 'Session / No. WA', 'Nama Pelanggan', 'Pengirim', 'Isi Pesan', 'Status Baca'];
    chatSheet.getRange(1, 1, 1, chatHeaders.length).setValues([chatHeaders]);
    chatSheet.getRange(1, 1, 1, chatHeaders.length).setFontWeight('bold').setBackground('#DCFCE7');
    chatSheet.setFrozenRows(1);
  }

  return ss;
}

/**
 * Mengambil seluruh pengaturan toko dari Sheet Pengaturan
 */
function getAppSettings() {
  const ss = ensureDatabaseInitialized();
  const sheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_SETTINGS);
  const data = sheet.getDataRange().getValues();
  const map = {
    ADMIN_EMAIL: DEFAULT_CONFIG.ADMIN_EMAIL,
    ADMIN_PIN: DEFAULT_CONFIG.ADMIN_PIN,
    BCA_ACCOUNT: DEFAULT_CONFIG.BCA_ACCOUNT,
    BCA_NAME: DEFAULT_CONFIG.BCA_NAME,
    QRIS_IMAGE_URL: DEFAULT_CONFIG.QRIS_IMAGE_URL,
    AUTO_APPROVE_USER: DEFAULT_CONFIG.AUTO_APPROVE_USER,
    APP_NAME: DEFAULT_CONFIG.APP_NAME
  };

  for (let i = 1; i < data.length; i++) {
    const k = String(data[i][0] || '').trim();
    const v = String(data[i][1] || '').replace(/^'/, '').trim();
    if (k) map[k] = v;
  }

  if (!map.ADMIN_EMAIL) {
    try {
      map.ADMIN_EMAIL = Session.getActiveUser().getEmail() || '';
    } catch (e) {}
  }
  return map;
}

/**
 * Mengambil katalog bahan aktif yang dikelompokkan per kategori cetak
 */
function getCatalogFromSheet() {
  const products = getAllProductsList();
  const catalog = {
    outdoor: { label: 'Print Outdoor', materials: [] },
    indoor: { label: 'Print Indoor', materials: [] },
    uv: { label: 'Print UV', materials: [] }
  };

  products.forEach(function(p) {
    if (p.active === 'AKTIF' && catalog[p.category]) {
      catalog[p.category].materials.push({
        id: p.id,
        name: p.name,
        ratePerSqm: Number(p.ratePerSqm) || 35000
      });
    }
  });

  return catalog;
}

/**
 * Mengambil konfigurasi publik untuk halaman pelanggan
 */
function getPublicPaymentConfig() {
  const settings = getAppSettings();
  let webAppUrl = '';
  try {
    webAppUrl = ScriptApp.getService().getUrl() || '';
  } catch (e) {}

  return {
    bcaAccount: String(settings.BCA_ACCOUNT || DEFAULT_CONFIG.BCA_ACCOUNT),
    bcaName: String(settings.BCA_NAME || DEFAULT_CONFIG.BCA_NAME),
    qrisImageUrl: String(settings.QRIS_IMAGE_URL || ''),
    appName: String(settings.APP_NAME || DEFAULT_CONFIG.APP_NAME),
    catalog: getCatalogFromSheet(),
    webAppUrl: webAppUrl
  };
}

/**
 * Menyusun data awal Admin saat halaman admin.html dimuat
 */
function buildAdminBootPayload() {
  const settings = getAppSettings();
  let webAppUrl = '';
  try {
    webAppUrl = ScriptApp.getService().getUrl() || '';
  } catch (e) {}

  return {
    success: true,
    orders: fetchAllOrdersFromSheet(),
    usersList: getAllUsersList(),
    productsList: getAllProductsList(),
    chats: getGroupedChatsList(),
    catalog: getCatalogFromSheet(),
    webAppUrl: webAppUrl,
    settings: {
      adminEmail: String(settings.ADMIN_EMAIL || ''),
      bcaAccount: String(settings.BCA_ACCOUNT || ''),
      bcaName: String(settings.BCA_NAME || ''),
      qrisImageUrl: String(settings.QRIS_IMAGE_URL || ''),
      autoApproveUser: String(settings.AUTO_APPROVE_USER || DEFAULT_CONFIG.AUTO_APPROVE_USER),
      appName: String(settings.APP_NAME || DEFAULT_CONFIG.APP_NAME)
    }
  };
}

/**
 * Normalisasi nomor WhatsApp menjadi format 08...
 */
function normalizeWaServer(wa) {
  let clean = String(wa || '').replace(/[^0-9]/g, '');
  if (clean.startsWith('62')) clean = '0' + clean.slice(2);
  return clean;
}

/**
 * Registrasi pelanggan baru
 */
function registerCustomer(userData) {
  const ss = ensureDatabaseInitialized();
  const sheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_USERS);
  const settings = getAppSettings();

  const name = String(userData.name || '').trim();
  const wa = normalizeWaServer(userData.wa);
  const email = String(userData.email || '').trim();
  const address = String(userData.address || '').trim();
  const pin = String(userData.pin || '').trim();

  if (!name || wa.length < 9 || pin.length < 4) {
    return { success: false, message: 'Nama, Nomor WA valid, & PIN minimal 4 digit wajib diisi!' };
  }

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const existingWa = normalizeWaServer(String(data[i][3] || '').replace(/^'/, ''));
    if (existingWa === wa) {
      return { success: false, message: 'Nomor WhatsApp sudah terdaftar! Silakan masuk menggunakan PIN Anda.' };
    }
  }

  const userId = 'USR-' + new Date().getTime().toString().slice(-6);
  const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd MMM yyyy HH:mm');
  const status = (settings.AUTO_APPROVE_USER === 'TIDAK') ? 'MENUNGGU PERSETUJUAN' : 'AKTIF';

  sheet.appendRow([userId, nowStr, name, "'" + wa, email, address, "'" + pin, status]);
  SpreadsheetApp.flush();

  if (settings.ADMIN_EMAIL && settings.ADMIN_EMAIL.indexOf('@') > -1) {
    try {
      MailApp.sendEmail({
        to: settings.ADMIN_EMAIL,
        subject: '[PELANGGAN BARU] ' + name + ' (' + wa + ')',
        body: 'Pendaftaran pelanggan baru di Aplikasi Cetak MMT:\n\n' +
              'ID Pelanggan : ' + userId + '\n' +
              'Nama         : ' + name + '\n' +
              'No. WhatsApp : ' + wa + '\n' +
              'Email        : ' + (email || '-') + '\n' +
              'Alamat       : ' + (address || '-') + '\n' +
              'Status Akun  : ' + status + '\n\n' +
              'Kelola akun pelanggan melalui Panel Kontrol Admin.'
      });
    } catch (e) {}
  }

  return {
    success: true,
    user: { id: userId, createdAt: nowStr, name: name, wa: wa, email: email, address: address, status: status },
    message: status === 'AKTIF'
      ? 'Registrasi berhasil! Akun Anda langsung aktif dan siap memesan.'
      : 'Registrasi berhasil! Menunggu persetujuan (ACC) dari Admin.'
  };
}

/**
 * Login pelanggan terdaftar
 */
function loginCustomer(waInput, pinInput) {
  const ss = ensureDatabaseInitialized();
  const sheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_USERS);
  const wa = normalizeWaServer(waInput);
  const pin = String(pinInput || '').trim();

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowWa = normalizeWaServer(String(data[i][3] || '').replace(/^'/, ''));
    const rowPin = String(data[i][6] || '').replace(/^'/, '').trim();
    if (rowWa === wa && rowPin === pin) {
      const status = String(data[i][7] || 'AKTIF').trim();
      if (status !== 'AKTIF') {
        return { success: false, message: 'Akun Anda berstatus ' + status + '. Silakan hubungi Admin.' };
      }
      const userObj = {
        id: String(data[i][0]),
        createdAt: String(data[i][1]),
        name: String(data[i][2]),
        wa: rowWa,
        email: String(data[i][4] || ''),
        address: String(data[i][5] || ''),
        status: status
      };
      const allOrders = fetchAllOrdersFromSheet();
      const userOrders = allOrders.filter(function(o) { return normalizeWaServer(o.custWa) === wa; });
      return { success: true, user: userObj, orders: userOrders };
    }
  }

  return { success: false, message: 'Nomor WhatsApp atau PIN salah!' };
}

/**
 * Menyimpan file Base64 ke folder Google Drive aplikasi
 */
function saveBase64FileToDrive(base64Data, mimeType, fileName) {
  const folders = DriveApp.getFoldersByName(DEFAULT_CONFIG.DRIVE_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(DEFAULT_CONFIG.DRIVE_FOLDER_NAME);
  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, mimeType || 'application/octet-stream', fileName);
  const file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}
  return {
    url: file.getUrl(),
    blob: blob
  };
}

/**
 * Menerima pesanan baru dari pelanggan, menyimpan ke Google Sheets, Drive & mengirim Email ke Admin
 */
function submitOrderToGAS(payload) {
  const ss = ensureDatabaseInitialized();
  const orderSheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_ORDERS);
  const settings = getAppSettings();

  const custName = String(payload.custName || '').trim();
  const custWa = normalizeWaServer(payload.custWa);
  if (!custName || !custWa) {
    return { success: false, message: 'Nama dan Nomor WhatsApp wajib diisi!' };
  }

  // Validasi status pelanggan di database
  const userSheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_USERS);
  const userData = userSheet.getDataRange().getValues();
  let customerId = 'GUEST';
  for (let i = 1; i < userData.length; i++) {
    const rWa = normalizeWaServer(String(userData[i][3] || '').replace(/^'/, ''));
    if (rWa === custWa) {
      const st = String(userData[i][7] || 'AKTIF').trim();
      if (st !== 'AKTIF') {
        return { success: false, message: 'Akun Anda berstatus ' + st + '. Hubungi Admin untuk aktivasi.' };
      }
      customerId = String(userData[i][0] || 'USER');
      break;
    }
  }

  const now = new Date();
  const invCode = 'INV-' + Utilities.formatDate(now, 'Asia/Jakarta', 'yyyyMMdd') + '-' + Math.floor(100 + Math.random() * 900);
  const dateStr = Utilities.formatDate(now, 'Asia/Jakarta', 'dd MMM yyyy HH:mm');
  const branch = String(payload.branch || 'Pusat Percetakan');
  const deliveryLabel = String(payload.deliveryLabel || 'Ambil di Toko');
  const payMethod = String(payload.payMethod || 'Transfer BCA');
  const grandTotal = Number(payload.grandTotal) || 0;
  const items = Array.isArray(payload.items) ? payload.items : [];

  let designUrl = String(payload.designLink || '').trim();
  let emailAttachments = [];

  if (payload.fileBase64 && payload.fileName) {
    try {
      const saved = saveBase64FileToDrive(payload.fileBase64, payload.fileMimeType, invCode + '_' + payload.fileName);
      designUrl = saved.url;
      if (saved.blob.getBytes().length <= DEFAULT_CONFIG.MAX_EMAIL_ATTACHMENT_MB * 1024 * 1024) {
        emailAttachments.push(saved.blob);
      }
    } catch (driveErr) {
      designUrl = 'File: ' + payload.fileName + ' (Gagal upload Drive: ' + driveErr.message + ')';
    }
  }

  if (!designUrl) designUrl = 'Belum lampirkan file';

  const summaryText = items.map(function(i) {
    return i.materialName + ' (' + i.panjang + 'x' + i.lebar + 'cm x' + i.jumlah + ')';
  }).join('; ');

  orderSheet.appendRow([
    invCode,
    dateStr,
    branch,
    custName,
    "'" + custWa,
    deliveryLabel,
    payMethod,
    grandTotal,
    'Menunggu Pembayaran',
    designUrl,
    JSON.stringify(items),
    summaryText,
    '',
    customerId
  ]);
  SpreadsheetApp.flush();

  // Kirim Email Notifikasi Otomatis ke Admin
  if (settings.ADMIN_EMAIL && settings.ADMIN_EMAIL.indexOf('@') > -1) {
    try {
      const detailLines = items.map(function(it, idx) {
        return (idx + 1) + '. ' + it.materialName + ' [' + it.categoryLabel + ']\n' +
               '   Ukuran    : ' + it.panjang + ' x ' + it.lebar + ' cm (' + Number(it.areaSqm || 0).toFixed(2) + ' m2)\n' +
               '   Jumlah    : ' + it.jumlah + ' pcs\n' +
               '   Produksi  : ' + it.waktuLabel + '\n' +
               '   Finishing : ' + it.finishing + '\n' +
               '   Catatan   : ' + (it.catatan || '-') + '\n' +
               '   Subtotal  : Rp ' + Math.round(it.subtotal).toLocaleString('id-ID');
      }).join('\n\n');

      const mailOpts = {
        to: settings.ADMIN_EMAIL,
        subject: '[ORDER BARU] ' + invCode + ' - ' + custName + ' (Rp ' + Math.round(grandTotal).toLocaleString('id-ID') + ')',
        body: 'Pesanan cetak MMT baru masuk:\n\n' +
              'Kode Invoice : ' + invCode + '\n' +
              'Tanggal      : ' + dateStr + '\n' +
              'Pemesan      : ' + custName + ' (' + custWa + ')\n' +
              'Pengiriman   : ' + deliveryLabel + '\n' +
              'Pembayaran   : ' + payMethod + '\n' +
              'Total Tagihan: Rp ' + Math.round(grandTotal).toLocaleString('id-ID') + '\n' +
              'File Desain  : ' + designUrl + '\n\n' +
              'RINCIAN ITEM:\n' + detailLines
      };
      if (emailAttachments.length > 0) {
        mailOpts.attachments = emailAttachments;
      }
      MailApp.sendEmail(mailOpts);
    } catch (mailErr) {}
  }

  return {
    success: true,
    order: {
      code: invCode,
      date: dateStr,
      branch: branch,
      custName: custName,
      custWa: custWa,
      deliveryLabel: deliveryLabel,
      payMethod: payMethod,
      grandTotal: grandTotal,
      status: 'Menunggu Pembayaran',
      designUrl: designUrl,
      paymentProofUrl: '',
      items: items
    }
  };
}

/**
 * Mengunggah bukti transfer pembayaran dari pelanggan
 */
function submitPaymentProof(invoiceCode, base64Data, mimeType, fileName) {
  const ss = ensureDatabaseInitialized();
  const sheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_ORDERS);
  const settings = getAppSettings();
  const code = String(invoiceCode || '').trim();

  let proofUrl = '';
  try {
    const saved = saveBase64FileToDrive(base64Data, mimeType, 'BUKTI_' + code + '_' + fileName);
    proofUrl = saved.url;
  } catch (e) {
    return { success: false, message: 'Gagal menyimpan file bukti ke Drive: ' + e.message };
  }

  const data = sheet.getDataRange().getValues();
  let custName = '';
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === code) {
      custName = String(data[i][3] || '');
      sheet.getRange(i + 1, 9).setValue('Verifikasi Pembayaran');
      sheet.getRange(i + 1, 13).setValue(proofUrl);
      break;
    }
  }
  SpreadsheetApp.flush();

  if (settings.ADMIN_EMAIL && settings.ADMIN_EMAIL.indexOf('@') > -1) {
    try {
      MailApp.sendEmail({
        to: settings.ADMIN_EMAIL,
        subject: '[BUKTI BAYAR] ' + code + ' - ' + custName,
        body: 'Pelanggan telah mengunggah bukti pembayaran untuk invoice ' + code + '.\n\n' +
              'Link Bukti Transfer: ' + proofUrl + '\n\n' +
              'Silakan cek mutasi dan ubah status pesanan di Panel Admin.'
      });
    } catch (e) {}
  }

  return {
    success: true,
    proofUrl: proofUrl,
    status: 'Verifikasi Pembayaran',
    message: 'Bukti pembayaran berhasil dikirim!'
  };
}

/**
 * Melacak pesanan pelanggan berdasarkan Invoice atau Nomor WA
 */
function trackCustomerOrder(query) {
  const q = String(query || '').toLowerCase().trim();
  if (!q) return [];
  const all = fetchAllOrdersFromSheet();
  return all.filter(function(o) {
    return String(o.code).toLowerCase().indexOf(q) > -1 ||
           normalizeWaServer(o.custWa).indexOf(normalizeWaServer(q)) > -1;
  });
}

/**
 * Mengambil seluruh daftar pesanan dari Sheet Data_Pesanan
 */
function fetchAllOrdersFromSheet() {
  const ss = ensureDatabaseInitialized();
  const sheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_ORDERS);
  const data = sheet.getDataRange().getValues();
  const list = [];

  for (let i = data.length - 1; i >= 1; i--) {
    const r = data[i];
    if (!r[0]) continue;
    let parsedItems = [];
    try {
      parsedItems = JSON.parse(String(r[10] || '[]'));
    } catch (e) {}

    const dateFormatted = (r[1] instanceof Date)
      ? Utilities.formatDate(r[1], 'Asia/Jakarta', 'dd MMM yyyy HH:mm')
      : String(r[1] || '');

    list.push({
      code: String(r[0]).trim(),
      date: dateFormatted,
      branch: String(r[2] || ''),
      custName: String(r[3] || ''),
      custWa: normalizeWaServer(String(r[4] || '').replace(/^'/, '')),
      deliveryLabel: String(r[5] || ''),
      payMethod: String(r[6] || ''),
      grandTotal: Number(r[7]) || 0,
      status: String(r[8] || 'Menunggu Pembayaran'),
      designUrl: String(r[9] || ''),
      items: parsedItems,
      summary: String(r[11] || ''),
      paymentProofUrl: String(r[12] || ''),
      customerId: String(r[13] || '')
    });
  }
  return list;
}

/**
 * Mengambil seluruh daftar pelanggan dari Sheet Data_Pelanggan
 */
function getAllUsersList() {
  const ss = ensureDatabaseInitialized();
  const sheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  const list = [];

  for (let i = data.length - 1; i >= 1; i--) {
    const r = data[i];
    if (!r[0]) continue;
    const dateFormatted = (r[1] instanceof Date)
      ? Utilities.formatDate(r[1], 'Asia/Jakarta', 'dd MMM yyyy HH:mm')
      : String(r[1] || '');
    list.push({
      id: String(r[0]).trim(),
      createdAt: dateFormatted,
      name: String(r[2] || '').trim(),
      wa: normalizeWaServer(String(r[3] || '').replace(/^'/, '')),
      email: String(r[4] || '').trim(),
      address: String(r[5] || '').trim(),
      pin: String(r[6] || '').replace(/^'/, '').trim(),
      status: String(r[7] || 'AKTIF').trim()
    });
  }
  return list;
}

/**
 * Mengambil seluruh produk bahan dari Sheet Produk_Bahan
 */
function getAllProductsList() {
  const ss = ensureDatabaseInitialized();
  const sheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_PRODUCTS);
  const data = sheet.getDataRange().getValues();
  const list = [];

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    list.push({
      id: String(r[0]).trim(),
      category: String(r[1] || 'outdoor').trim(),
      name: String(r[2] || '').trim(),
      ratePerSqm: Number(r[3]) || 35000,
      active: String(r[4] || 'AKTIF').trim()
    });
  }
  return list;
}

/**
 * Admin menambah atau mengedit data pelanggan
 */
function adminSaveUser(pin, userObj) {
  const settings = getAppSettings();
  if (String(pin).trim() !== String(settings.ADMIN_PIN).trim()) {
    return { success: false, message: 'PIN Admin tidak valid!' };
  }

  const ss = ensureDatabaseInitialized();
  const sheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();

  const id = String(userObj.id || '').trim();
  const name = String(userObj.name || '').trim();
  const wa = normalizeWaServer(userObj.wa);
  const userPin = String(userObj.pin || '123456').trim();
  const email = String(userObj.email || '').trim();
  const address = String(userObj.address || '').trim();
  const status = String(userObj.status || 'AKTIF').trim();

  if (!name || !wa) {
    return { success: false, message: 'Nama dan Nomor WhatsApp wajib diisi!' };
  }

  if (id) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === id) {
        sheet.getRange(i + 1, 3, 1, 6).setValues([[name, "'" + wa, email, address, "'" + userPin, status]]);
        SpreadsheetApp.flush();
        return { success: true, usersList: getAllUsersList(), message: 'Data pelanggan diperbarui!' };
      }
    }
  }

  for (let i = 1; i < data.length; i++) {
    const existingWa = normalizeWaServer(String(data[i][3] || '').replace(/^'/, ''));
    if (existingWa === wa) {
      return { success: false, message: 'Nomor WhatsApp sudah dipakai oleh akun lain!' };
    }
  }

  const newId = 'USR-' + new Date().getTime().toString().slice(-6);
  const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd MMM yyyy HH:mm');
  sheet.appendRow([newId, nowStr, name, "'" + wa, email, address, "'" + userPin, status]);
  SpreadsheetApp.flush();

  return { success: true, usersList: getAllUsersList(), message: 'Pelanggan baru berhasil ditambahkan!' };
}

/**
 * Admin menghapus akun pelanggan
 */
function adminDeleteUser(pin, userId) {
  const settings = getAppSettings();
  if (String(pin).trim() !== String(settings.ADMIN_PIN).trim()) {
    return { success: false, message: 'PIN Admin tidak valid!' };
  }

  const ss = ensureDatabaseInitialized();
  const sheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(userId).trim()) {
      sheet.deleteRow(i + 1);
      SpreadsheetApp.flush();
      return { success: true, usersList: getAllUsersList(), message: 'Akun pelanggan dihapus!' };
    }
  }
  return { success: false, message: 'Akun tidak ditemukan.' };
}

/**
 * Admin mengubah status pesanan
 */
function adminUpdateOrderStatus(pin, invoiceCode, newStatus) {
  const settings = getAppSettings();
  if (String(pin).trim() !== String(settings.ADMIN_PIN).trim()) {
    return { success: false, message: 'PIN Admin tidak valid!' };
  }

  const ss = ensureDatabaseInitialized();
  const sheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_ORDERS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(invoiceCode).trim()) {
      sheet.getRange(i + 1, 9).setValue(String(newStatus).trim());
      SpreadsheetApp.flush();
      return { success: true, orders: fetchAllOrdersFromSheet(), message: 'Status ' + invoiceCode + ' diubah ke ' + newStatus };
    }
  }
  return { success: false, message: 'Pesanan tidak ditemukan.' };
}

/**
 * Admin menghapus pesanan
 */
function adminDeleteOrder(pin, invoiceCode) {
  const settings = getAppSettings();
  if (String(pin).trim() !== String(settings.ADMIN_PIN).trim()) {
    return { success: false, message: 'PIN Admin tidak valid!' };
  }

  const ss = ensureDatabaseInitialized();
  const sheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_ORDERS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(invoiceCode).trim()) {
      sheet.deleteRow(i + 1);
      SpreadsheetApp.flush();
      return { success: true, orders: fetchAllOrdersFromSheet(), message: 'Pesanan ' + invoiceCode + ' dihapus.' };
    }
  }
  return { success: false, message: 'Pesanan tidak ditemukan.' };
}

/**
 * Admin menambah atau memperbarui produk bahan & harga per m2
 */
function adminSaveProduct(pin, prodObj) {
  const settings = getAppSettings();
  if (String(pin).trim() !== String(settings.ADMIN_PIN).trim()) {
    return { success: false, message: 'PIN Admin tidak valid!' };
  }

  const ss = ensureDatabaseInitialized();
  const sheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_PRODUCTS);
  const data = sheet.getDataRange().getValues();

  const id = String(prodObj.id || '').trim();
  const category = String(prodObj.category || 'outdoor').trim();
  const name = String(prodObj.name || '').trim();
  const ratePerSqm = Number(prodObj.ratePerSqm) || 35000;
  const active = String(prodObj.active || 'AKTIF').trim();

  if (id) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === id) {
        sheet.getRange(i + 1, 2, 1, 4).setValues([[category, name, ratePerSqm, active]]);
        SpreadsheetApp.flush();
        return { success: true, productsList: getAllProductsList(), message: 'Produk & harga diperbarui!' };
      }
    }
  }

  const newId = 'PROD-' + new Date().getTime().toString().slice(-6);
  sheet.appendRow([newId, category, name, ratePerSqm, active]);
  SpreadsheetApp.flush();
  return { success: true, productsList: getAllProductsList(), message: 'Produk baru berhasil ditambahkan!' };
}

/**
 * Admin menghapus produk bahan
 */
function adminDeleteProduct(pin, prodId) {
  const settings = getAppSettings();
  if (String(pin).trim() !== String(settings.ADMIN_PIN).trim()) {
    return { success: false, message: 'PIN Admin tidak valid!' };
  }

  const ss = ensureDatabaseInitialized();
  const sheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_PRODUCTS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(prodId).trim()) {
      sheet.deleteRow(i + 1);
      SpreadsheetApp.flush();
      return { success: true, productsList: getAllProductsList(), message: 'Produk berhasil dihapus!' };
    }
  }
  return { success: false, message: 'Produk tidak ditemukan.' };
}

/**
 * Admin menyimpan pengaturan pembayaran, email, PIN, QRIS, & kebijakan akun
 */
function adminUpdateSettings(pin, newEmail, newPin, bcaAcc, bcaName, qrisBase64, qrisMime, qrisName, autoApproveUser, appName) {
  const settings = getAppSettings();
  if (String(pin).trim() !== String(settings.ADMIN_PIN).trim()) {
    return { success: false, message: 'PIN Admin tidak valid!' };
  }

  let qrisUrl = settings.QRIS_IMAGE_URL || '';
  if (qrisBase64 && qrisName) {
    try {
      const saved = saveBase64FileToDrive(qrisBase64, qrisMime, 'QRIS_TOKO_' + qrisName);
      qrisUrl = saved.url;
    } catch (e) {}
  }

  const ss = ensureDatabaseInitialized();
  const sheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_SETTINGS);
  const updates = {
    ADMIN_EMAIL: String(newEmail || '').trim(),
    ADMIN_PIN: "'" + (String(newPin || '').trim() || String(settings.ADMIN_PIN || '123456').trim()),
    BCA_ACCOUNT: "'" + String(bcaAcc || '').trim(),
    BCA_NAME: String(bcaName || '').trim(),
    QRIS_IMAGE_URL: qrisUrl,
    AUTO_APPROVE_USER: String(autoApproveUser || 'YA').trim(),
    APP_NAME: String(appName || settings.APP_NAME || DEFAULT_CONFIG.APP_NAME).trim().slice(0, 60) || DEFAULT_CONFIG.APP_NAME
  };

  const data = sheet.getDataRange().getValues();
  Object.keys(updates).forEach(function(k) {
    let foundRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === k) {
        foundRow = i + 1;
        break;
      }
    }
    if (foundRow > 0) {
      sheet.getRange(foundRow, 2).setValue(updates[k]);
    } else {
      sheet.appendRow([k, updates[k]]);
    }
  });
  SpreadsheetApp.flush();

  return { success: true, message: 'Pengaturan Admin berhasil disimpan!' };
}

/**
 * Pelanggan mengirim pesan chat CS
 */
function sendCustomerChat(sessionId, custName, custWa, text) {
  const ss = ensureDatabaseInitialized();
  const sheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_CHATS);
  const sid = String(sessionId || custWa || 'TAMU').trim();
  const name = String(custName || 'Pelanggan').trim();
  const msg = String(text || '').trim();
  if (!msg) return { success: false, message: 'Pesan kosong' };

  const msgId = 'MSG-' + new Date().getTime().toString().slice(-7);
  const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd MMM HH:mm');
  sheet.appendRow([msgId, nowStr, "'" + sid, name, 'CUSTOMER', msg, 'UNREAD']);
  SpreadsheetApp.flush();

  return {
    success: true,
    messages: getCustomerChatHistory(sid)
  };
}

/**
 * Mengambil seluruh percakapan CS yang dikelompokkan per pelanggan
 */
function getGroupedChatsList() {
  const ss = ensureDatabaseInitialized();
  const sheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_CHATS);
  const data = sheet.getDataRange().getValues();
  const threadMap = {};

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    const msgId = String(r[0]).trim();
    const timeStr = (r[1] instanceof Date)
      ? Utilities.formatDate(r[1], 'Asia/Jakarta', 'dd MMM HH:mm')
      : String(r[1] || '').trim();
    const sessionId = String(r[2] || '').replace(/^'/, '').trim();
    const custName = String(r[3] || 'Pelanggan').trim();
    const sender = String(r[4] || 'CUSTOMER').trim().toUpperCase();
    const text = String(r[5] || '').trim();
    const readStatus = String(r[6] || 'UNREAD').trim().toUpperCase();

    if (!sessionId) continue;
    if (!threadMap[sessionId]) {
      threadMap[sessionId] = {
        sessionId: sessionId,
        customerName: custName,
        unreadCount: 0,
        lastMessage: text,
        lastTime: timeStr,
        lastRowIndex: i,
        messages: []
      };
    }

    if (custName && custName !== 'Tamu' && custName !== 'Pelanggan') {
      threadMap[sessionId].customerName = custName;
    }
    if (sender === 'CUSTOMER' && readStatus === 'UNREAD') {
      threadMap[sessionId].unreadCount += 1;
    }
    threadMap[sessionId].lastMessage = text;
    threadMap[sessionId].lastTime = timeStr;
    threadMap[sessionId].lastRowIndex = i;
    threadMap[sessionId].messages.push({
      id: msgId,
      time: timeStr,
      sender: sender,
      senderName: sender === 'ADMIN' ? 'Admin & Operator' : threadMap[sessionId].customerName,
      text: text,
      readStatus: readStatus
    });
  }

  return Object.keys(threadMap)
    .map(function(k) { return threadMap[k]; })
    .sort(function(a, b) { return b.lastRowIndex - a.lastRowIndex; });
}

/**
 * Mengambil riwayat chat untuk 1 pelanggan berdasarkan sessionId / No. WA
 */
function getCustomerChatHistory(sessionId) {
  const sid = String(sessionId || '').trim();
  if (!sid) return [];
  const allThreads = getGroupedChatsList();
  const found = allThreads.find(function(t) { return t.sessionId === sid; });
  return found ? found.messages : [];
}

/**
 * Mengambil seluruh data sekaligus untuk halaman khusus admin.html
 */
function adminGetAllData(pin) {
  try {
    const settings = getAppSettings();
    if (String(pin || '').trim() !== String(settings.ADMIN_PIN || '123456').trim()) {
      return { success: false, message: 'PIN Admin salah! (Default: 123456)' };
    }
    return buildAdminBootPayload();
  } catch (err) {
    return {
      success: false,
      message: 'Gagal memuat database: ' + (err && err.message ? err.message : String(err))
    };
  }
}

/**
 * Admin mengambil semua daftar chat pelanggan
 */
function adminGetAllChats(pin) {
  const settings = getAppSettings();
  if (String(pin).trim() !== String(settings.ADMIN_PIN).trim()) {
    return { success: false, message: 'Sesi Admin tidak valid!' };
  }
  return {
    success: true,
    chats: getGroupedChatsList()
  };
}

/**
 * Admin membalas pesan chat pelanggan & menandai pesan pelanggan sudah dibaca
 */
function adminReplyChat(pin, sessionId, replyText) {
  const settings = getAppSettings();
  if (String(pin).trim() !== String(settings.ADMIN_PIN).trim()) {
    return { success: false, message: 'Sesi Admin tidak valid!' };
  }

  const sid = String(sessionId || '').trim();
  const text = String(replyText || '').trim();
  if (!sid || !text) {
    return { success: false, message: 'Balasan tidak boleh kosong.' };
  }

  const ss = ensureDatabaseInitialized();
  const sheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_CHATS);
  const data = sheet.getDataRange().getValues();
  let custName = 'Pelanggan';

  for (let i = 1; i < data.length; i++) {
    const rowSid = String(data[i][2] || '').replace(/^'/, '').trim();
    if (rowSid === sid) {
      if (data[i][3]) custName = String(data[i][3]).trim();
      if (String(data[i][4]).trim().toUpperCase() === 'CUSTOMER') {
        sheet.getRange(i + 1, 7).setValue('READ');
      }
    }
  }

  const msgId = 'MSG-' + new Date().getTime().toString().slice(-7);
  const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd MMM HH:mm');
  sheet.appendRow([msgId, nowStr, "'" + sid, custName, 'ADMIN', text, 'READ']);
  SpreadsheetApp.flush();

  return {
    success: true,
    chats: getGroupedChatsList()
  };
}

/**
 * Menandai chat pelanggan sudah dibaca saat dibuka Admin
 */
function adminMarkChatRead(pin, sessionId) {
  const settings = getAppSettings();
  if (String(pin).trim() !== String(settings.ADMIN_PIN).trim()) {
    return { success: false };
  }
  const sid = String(sessionId || '').trim();
  const ss = ensureDatabaseInitialized();
  const sheet = ss.getSheetByName(DEFAULT_CONFIG.SHEET_CHATS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const rowSid = String(data[i][2] || '').replace(/^'/, '').trim();
    if (rowSid === sid && String(data[i][4]).trim().toUpperCase() === 'CUSTOMER') {
      sheet.getRange(i + 1, 7).setValue('READ');
    }
  }
  SpreadsheetApp.flush();
  return { success: true, chats: getGroupedChatsList() };
}