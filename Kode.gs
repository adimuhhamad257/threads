/**
 * ThreadChain Manager - Enterprise Backend Core
 * Single Spreadsheet Connection Architecture
 */

const SPREADSHEET_ID = '1fbG5lrBAAIfYG0np1P7kqSpOclsm9WN1G4E1wBiJ1Uw';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ThreadChain Manager - Meta Threads Automation')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Shared Helper: Serializes any data safely for google.script.run
 */
function serializeForClient(obj) {
  if (obj === null || obj === undefined) return null;
  if (obj instanceof Date) {
    return obj.toISOString();
  }
  if (Array.isArray(obj)) {
    return obj.map(function(item) { return serializeForClient(item); });
  }
  if (typeof obj === 'object') {
    var copy = {};
    for (var key in obj) {
      if (Object.prototype.hasOwnProperty.call(obj, key)) {
        copy[key] = serializeForClient(obj[key]);
      }
    }
    return copy;
  }
  return obj;
}

/**
 * Shared Helper: Normalisasi nilai tanggal menjadi string ISO 8601 UTC murni
 */
function toIsoUtcString(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return isNaN(val.getTime()) ? '' : val.toISOString();
  }
  var sVal = String(val).trim();
  if (!sVal) return '';
  // Jika format datetime-local tanpa offset timezone "YYYY-MM-DDTHH:mm", asumsikan WIB (UTC+7)
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(?::\d{2})?$/.test(sVal)) {
    sVal += '+07:00';
  }
  var d = new Date(sVal);
  return isNaN(d.getTime()) ? '' : d.toISOString();
}

/**
 * Normalizes rows into JSON objects using header row 1
 */
function readSheetAsJson(sheetName) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  var headers = values[0];
  var result = [];
  var dateCols = ['scheduled_at', 'created_at', 'view_check_until', 'published_at', 'token_expires_at', 'last_synced_at'];

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var item = {};
    var hasContent = false;
    for (var j = 0; j < headers.length; j++) {
      var val = row[j];
      var headerName = headers[j];
      if (val instanceof Date) {
        if (headerName === 'generation_time') {
          val = Utilities.formatDate(val, 'Asia/Jakarta', 'HH:mm');
        } else {
          val = val.toISOString();
        }
      }
      if (dateCols.indexOf(headerName) !== -1 && val) {
        val = toIsoUtcString(val);
      }
      if (headerName === 'generation_time' && val) {
        var mGenTime = String(val).match(/(\d{1,2}):(\d{2})/);
        if (mGenTime) {
          val = ('0' + mGenTime[1]).slice(-2) + ':' + mGenTime[2];
        } else {
          val = '05:00';
        }
      }
      if (headerName === 'is_active') {
        val = (val === true || String(val).toUpperCase() === 'TRUE') ? 'TRUE' : 'FALSE';
      }
      item[headerName] = (val === '' ? null : val);
      if (val !== '' && val !== null) hasContent = true;
    }
    if (hasContent) result.push(item);
  }
  return result;
}

/**
 * Writes or appends an object array efficiently in batch
 */
function appendRowData(sheetName, rowObj) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var newRow = [];
  for (var i = 0; i < headers.length; i++) {
    var key = headers[i];
    var val = rowObj[key];
    if (val === undefined || val === null) {
      newRow.push('');
    } else {
      newRow.push(val);
    }
  }
  sheet.appendRow(newRow);
}

/**
 * DATABASE INITIALIZER
 * TAB_NAME definitions & safety re-runs with header overwrites
 */
function setupDatabase() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  var schemas = [
    {
      name: 'Accounts',
      headers: ['account_id', 'meta_user_id', 'username', 'name', 'profile_picture_url', 'access_token', 'token_expires_at', 'is_active', 'category', 'account_persona', 'created_at']
    },
    {
      name: 'Products',
      headers: ['product_id', 'title', 'category', 'price', 'description', 'affiliate_link', 'dashboard_url', 'commission_pct', 'commission_rp', 'cover_url', 'assigned_account_ids', 'created_at']
    },
    {
      name: 'Media_Batches',
      headers: ['batch_id', 'product_id', 'batch_number', 'created_at']
    },
    {
      name: 'Media_Files',
      headers: ['media_id', 'batch_id', 'product_id', 'media_type', 'cloudinary_url', 'thumbnail_url', 'cloudinary_public_id', 'created_at']
    },
    {
      name: 'Queue_Posts',
      headers: ['post_id', 'account_id', 'product_id', 'publishing_mode', 'interval_minutes', 'target_views', 'status', 'scheduled_at', 'view_check_until', 'is_evergreen', 'error_log', 'created_at']
    },
    {
      name: 'Published_Posts',
      headers: ['post_id', 'account_id', 'product_id', 'publishing_mode', 'interval_minutes', 'target_views', 'status', 'scheduled_at', 'view_check_until', 'is_evergreen', 'error_log', 'created_at']
    },
    {
      name: 'Post_Chains',
      headers: ['chain_id', 'post_id', 'sequence_order', 'content_text', 'media_ids', 'temp_container_id', 'threads_post_id', 'permalink', 'published_at']
    },
    {
      name: 'Post_Analytics',
      headers: ['analytics_id', 'chain_id', 'views_count', 'likes_count', 'reposts_count', 'replies_count', 'last_synced_at']
    },
    {
      name: 'Post_Comments',
      headers: ['comment_id', 'chain_id', 'threads_comment_id', 'from_username', 'text', 'created_at', 'status', 'reply_text']
    },
    {
      name: 'Settings',
      headers: ['setting_key', 'setting_value']
    },
    {
      name: 'Viral_Library',
      headers: ['viral_id', 'title', 'source_url', 'chains_json', 'product_ids', 'created_at']
    }
  ];

  var oldAutoSheet = ss.getSheetByName('Autopilot_Rules');
  if (oldAutoSheet) {
    ss.deleteSheet(oldAutoSheet);
  }

  // Migrasi otomatis data lama dari tabel Posts ke Queue_Posts & Published_Posts
  var oldPostsSheet = ss.getSheetByName('Posts');
  if (oldPostsSheet) {
    var oldValues = oldPostsSheet.getDataRange().getValues();
    if (oldValues.length > 1) {
      var oldHeaders = oldValues[0];
      var statusCol = oldHeaders.indexOf('status');
      var qSheet = ss.getSheetByName('Queue_Posts') || ss.insertSheet('Queue_Posts');
      var pSheet = ss.getSheetByName('Published_Posts') || ss.insertSheet('Published_Posts');
      for (var r = 1; r < oldValues.length; r++) {
        var row = oldValues[r];
        var st = statusCol >= 0 ? String(row[statusCol]) : '';
        if (st === 'Published' || st === 'Failed') {
          pSheet.appendRow(row);
        } else {
          qSheet.appendRow(row);
        }
      }
    }
    ss.deleteSheet(oldPostsSheet);
  }

  for (var i = 0; i < schemas.length; i++) {
    var tabInfo = schemas[i];
    var sheet = ss.getSheetByName(tabInfo.name);
    if (!sheet) {
      sheet = ss.insertSheet(tabInfo.name);
    }
    sheet.getRange(1, 1, 1, tabInfo.headers.length).setValues([tabInfo.headers]);
  }

  // Seed default settings if empty
  var settingsSheet = ss.getSheetByName('Settings');
  if (settingsSheet.getLastRow() <= 1) {
    var defaults = [
      ['timezone', 'Asia/Jakarta'],
      ['cloudinary_cloud_name', ''],
      ['cloudinary_upload_preset', ''],
      ['cloudinary_api_key', ''],
      ['cloudinary_api_secret', ''],
      ['default_fallback_views_hours', '24'],
      ['triggers_installed', 'FALSE'],
      ['best_time_pagi_start', '08:00'],
      ['best_time_pagi_end', '10:00'],
      ['best_time_siang_start', '12:00'],
      ['best_time_siang_end', '14:00'],
      ['best_time_sore_start', '16:00'],
      ['best_time_sore_end', '18:00'],
      ['best_time_malam_start', '19:00'],
      ['best_time_malam_end', '22:00']
    ];
    settingsSheet.getRange(2, 1, defaults.length, 2).setValues(defaults);
  }

  var deadSheets = ['AI_Prompts', 'Voice_DNA', 'Thread_Formats', 'AI_Training_Memory', 'Autopilot_Rules'];
  for (var ds = 0; ds < deadSheets.length; ds++) {
    var dSheet = ss.getSheetByName(deadSheets[ds]);
    if (dSheet) {
      ss.deleteSheet(dSheet);
    }
  }

  ensureRepliedColumn(ss);

  return { ok: true, message: 'Database schema successfully initialized/updated.' };
}

/**
 * Pemeliharaan skema & normalisasi kolom komentar Post_Comments secara terjadwal/startup
 */
function ensureRepliedColumn(ss) {
  var sheet = ss.getSheetByName('Post_Comments');
  if (!sheet) return;
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var statusCol = headers.indexOf('status');
  var replyCol = headers.indexOf('reply_text');

  if (statusCol === -1) {
    sheet.getRange(1, headers.length + 1).setValue('status');
    headers.push('status');
  }
  if (replyCol === -1) {
    sheet.getRange(1, headers.length + 1).setValue('reply_text');
    headers.push('reply_text');
  }

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var dateColIdx = headers.indexOf('created_at');
    if (dateColIdx !== -1) {
      var dateRange = sheet.getRange(2, dateColIdx + 1, lastRow - 1, 1);
      var dateVals = dateRange.getValues();
      var changed = false;
      var nowIso = new Date().toISOString();
      for (var r = 0; r < dateVals.length; r++) {
        var dVal = dateVals[r][0];
        if (typeof dVal === 'number' && dVal > 2000000000000) {
          dateVals[r][0] = nowIso;
          changed = true;
        } else if (dVal instanceof Date) {
          dateVals[r][0] = dVal.toISOString();
          changed = true;
        }
      }
      if (changed) {
        dateRange.setValues(dateVals);
      }
    }
  }
}

/**
 * Helper: Membaca seluruh data postingan gabungan (Antrean Queue + Arsip Terbit)
 */
function getAllPostsAsJson() {
  var q = readSheetAsJson('Queue_Posts');
  var p = readSheetAsJson('Published_Posts');
  var legacy = readSheetAsJson('Posts');
  return [].concat(q, p, legacy);
}

/**
 * Helper: Mencari lokasi spesifik baris postingan di spreadsheet
 */
function findPostRowLocation(ss, postId) {
  var targetSheets = ['Queue_Posts', 'Published_Posts', 'Posts'];
  for (var s = 0; s < targetSheets.length; s++) {
    var sheet = ss.getSheetByName(targetSheets[s]);
    if (!sheet) continue;
    var values = sheet.getDataRange().getValues();
    for (var r = 1; r < values.length; r++) {
      if (String(values[r][0]) === String(postId)) {
        return {
          sheet: sheet,
          sheetName: targetSheets[s],
          rowIndex: r + 1,
          headers: values[0],
          rowValues: values[r]
        };
      }
    }
  }
  return null;
}

/**
 * Read Initial State for Frontend
 */
function getInitialData() {
  try {
    var settingsRows = readSheetAsJson('Settings');
    var settingsMap = {};
    for (var s = 0; s < settingsRows.length; s++) {
      settingsMap[settingsRows[s].setting_key] = settingsRows[s].setting_value;
    }

    var accounts = readSheetAsJson('Accounts');
    var products = readSheetAsJson('Products');
    var queuePosts = readSheetAsJson('Queue_Posts');
    var publishedPosts = readSheetAsJson('Published_Posts');
    var posts = getAllPostsAsJson();

    accounts.sort(function(a, b) { return new Date(b.created_at || 0) - new Date(a.created_at || 0); });
    products.sort(function(a, b) { return new Date(b.created_at || 0) - new Date(a.created_at || 0); });
    posts.sort(function(a, b) {
      var getPri = function(item) {
        if (item.status === 'Waiting Views') return 5;
        if (item.status === 'Processing') return 4;
        if (item.status === 'Scheduled') return 3;
        if (item.status === 'Published') return 2;
        return 1;
      };
      var pA = getPri(a);
      var pB = getPri(b);
      if (pA !== pB) return pB - pA;
      return new Date(b.scheduled_at || b.created_at || 0) - new Date(a.scheduled_at || a.created_at || 0);
    });

    var viralLibrary = readSheetAsJson('Viral_Library');
    viralLibrary.sort(function(a, b) { return new Date(b.created_at || 0) - new Date(a.created_at || 0); });

    var payload = {
      accounts: accounts,
      products: products,
      mediaBatches: readSheetAsJson('Media_Batches'),
      mediaFiles: readSheetAsJson('Media_Files'),
      posts: posts,
      queuePosts: queuePosts,
      publishedPosts: publishedPosts,
      postChains: readSheetAsJson('Post_Chains'),
      viralLibrary: viralLibrary,
      postAnalytics: readSheetAsJson('Post_Analytics'),
      postComments: readSheetAsJson('Post_Comments'),
      settings: settingsMap
    };

    return { ok: true, data: serializeForClient(payload) };
  } catch (err) {
    return { ok: false, message: 'Gagal memuat data: ' + err.message };
  }
}

/**
 * Meta Threads Graph API Account Verification & Sync
 */
function addOrSyncAccount(accessToken, category, accountPersona) {
  try {
    var url = 'https://graph.threads.net/v1.0/me?fields=id,username,name,threads_profile_picture_url&access_token=' + accessToken;
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var data = JSON.parse(resp.getContentText());

    if (data.error) {
      return { ok: false, message: 'Meta Graph API Error: ' + data.error.message };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Accounts');
    var rows = sheet.getDataRange().getValues();
    var nowIso = new Date().toISOString();
    var expiryDate = new Date();
    expiryDate.setDate(expiryDate.getDate() + 60); // 60 days default

    var accountId = 'acc_' + Utilities.getUuid().slice(0, 8);
    var rowIdxToUpdate = -1;
    var defaultCategory = category || 'General';
    var defaultPersona = accountPersona || '';

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === String(data.id)) {
        rowIdxToUpdate = i + 1;
        accountId = rows[i][0];
        if (!category && rows[i][8]) defaultCategory = rows[i][8];
        if (!accountPersona && rows[i][9]) defaultPersona = rows[i][9];
        break;
      }
    }

    var rowValues = [
      accountId,
      String(data.id),
      data.username || '',
      data.name || data.username || '',
      data.threads_profile_picture_url || '',
      accessToken,
      expiryDate.toISOString(),
      'TRUE',
      defaultCategory,
      defaultPersona,
      nowIso
    ];

    if (rowIdxToUpdate > 0) {
      sheet.getRange(rowIdxToUpdate, 1, 1, rowValues.length).setValues([rowValues]);
      return { ok: true, message: 'Akun @' + data.username + ' sudah terdaftar & token berhasil diperbarui!', data: { account_id: accountId, username: data.username } };
    } else {
      sheet.appendRow(rowValues);
      return { ok: true, message: 'Akun @' + data.username + ' berhasil ditambahkan!', data: { account_id: accountId, username: data.username } };
    }
  } catch (err) {
    return { ok: false, message: 'Exception: ' + err.message };
  }
}

function updateAccountDetails(accountData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Accounts');
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(accountData.account_id)) {
        var row = i + 1;
        if (accountData.category !== undefined) sheet.getRange(row, 9).setValue(accountData.category);
        if (accountData.account_persona !== undefined) sheet.getRange(row, 10).setValue(accountData.account_persona);
        return { ok: true, message: 'Persona dan kategori akun berhasil diperbarui' };
      }
    }
    return { ok: false, message: 'Akun tidak ditemukan' };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

function deleteAccount(accountId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Accounts');
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(accountId)) {
        sheet.deleteRow(i + 1);
        return { ok: true };
      }
    }
    return { ok: false, message: 'Akun tidak ditemukan' };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

/**
 * Product & Media CRUD (Supports Add, Edit, Delete)
 */
function saveProduct(productData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Products');
    var pId = productData.product_id;
    var commissionRp = Math.round((Number(productData.price || 0) * Number(productData.commission_pct || 0)) / 100);
    var nowIso = new Date().toISOString();

    var assignedAccStr = '';
    if (Array.isArray(productData.assigned_account_ids)) {
      assignedAccStr = productData.assigned_account_ids.join(',');
    } else if (productData.assigned_account_ids) {
      assignedAccStr = String(productData.assigned_account_ids).trim();
    }

    if (pId) {
      // Edit existing product
      var pValues = sheet.getDataRange().getValues();
      var headers = pValues[0];
      var pIdColIdx = headers.indexOf('product_id');
      if (pIdColIdx === -1) pIdColIdx = 0;
      var rowIdx = -1;
      for (var i = 1; i < pValues.length; i++) {
        if (String(pValues[i][pIdColIdx]) === String(pId)) {
          rowIdx = i + 1;
          break;
        }
      }
      if (rowIdx > 0) {
        var existingRow = pValues[rowIdx - 1];
        var coverColIdx = headers.indexOf('cover_url');
        var createdColIdx = headers.indexOf('created_at');
        var existingCover = coverColIdx >= 0 ? existingRow[coverColIdx] : '';
        var existingCreatedAt = createdColIdx >= 0 ? existingRow[createdColIdx] : nowIso;

        var rowObj = {
          product_id: pId,
          title: productData.title,
          category: productData.category || 'General',
          price: productData.price || 0,
          description: productData.description || '',
          affiliate_link: productData.affiliate_link,
          dashboard_url: productData.dashboard_url || '',
          commission_pct: productData.commission_pct || 0,
          commission_rp: commissionRp,
          cover_url: productData.cover_url !== undefined ? productData.cover_url : existingCover,
          assigned_account_ids: assignedAccStr,
          created_at: existingCreatedAt || nowIso
        };
        var updateRow = [];
        for (var h = 0; h < headers.length; h++) {
          var colKey = headers[h];
          if (rowObj[colKey] !== undefined && rowObj[colKey] !== null) {
            updateRow.push(rowObj[colKey]);
          } else {
            updateRow.push(existingRow[h] !== undefined ? existingRow[h] : '');
          }
        }
        sheet.getRange(rowIdx, 1, 1, updateRow.length).setValues([updateRow]);
        return { ok: true, data: rowObj };
      }
    }

    // New Product
    var productId = 'prod_' + Utilities.getUuid().slice(0, 8);
    var newObj = {
      product_id: productId,
      title: productData.title,
      category: productData.category || 'General',
      price: productData.price || 0,
      description: productData.description || '',
      affiliate_link: productData.affiliate_link,
      dashboard_url: productData.dashboard_url || '',
      commission_pct: productData.commission_pct || 0,
      commission_rp: commissionRp,
      cover_url: productData.cover_url || '',
      assigned_account_ids: assignedAccStr,
      created_at: nowIso
    };
    appendRowData('Products', newObj);
    return { ok: true, data: newObj };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

function bulkAssignAccountsToProducts(productIds, accountIds, mode) {
  try {
    if (!Array.isArray(productIds) || productIds.length === 0) {
      return { ok: false, message: 'Pilih setidaknya satu produk' };
    }
    var targetAccIds = Array.isArray(accountIds) ? accountIds : (accountIds ? String(accountIds).split(',') : []);
    targetAccIds = targetAccIds.map(function(id) { return String(id).trim(); }).filter(Boolean);

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Products');
    var values = sheet.getDataRange().getValues();
    if (values.length <= 1) return { ok: false, message: 'Katalog produk kosong' };

    var headers = values[0];
    var pIdCol = headers.indexOf('product_id');
    var assignedCol = headers.indexOf('assigned_account_ids');
    if (pIdCol === -1 || assignedCol === -1) {
      return { ok: false, message: 'Header kolom Products tidak valid' };
    }

    var updatedCount = 0;
    var assignMode = mode || 'replace';

    for (var i = 1; i < values.length; i++) {
      var currentPid = String(values[i][pIdCol]);
      if (productIds.indexOf(currentPid) !== -1) {
        var currentAssignedStr = String(values[i][assignedCol] || '');
        var currentAccs = currentAssignedStr ? currentAssignedStr.split(',').map(function(s) { return s.trim(); }).filter(Boolean) : [];

        var newAccs = [];
        if (assignMode === 'replace') {
          newAccs = [].concat(targetAccIds);
        } else if (assignMode === 'add') {
          newAccs = [].concat(currentAccs);
          for (var a = 0; a < targetAccIds.length; a++) {
            if (newAccs.indexOf(targetAccIds[a]) === -1) {
              newAccs.push(targetAccIds[a]);
            }
          }
        } else if (assignMode === 'remove') {
          newAccs = currentAccs.filter(function(acc) {
            return targetAccIds.indexOf(acc) === -1;
          });
        }

        var newAssignedStr = newAccs.join(',');
        sheet.getRange(i + 1, assignedCol + 1).setValue(newAssignedStr);
        updatedCount++;
      }
    }

    return {
      ok: true,
      message: 'Hak kelola akun berhasil diperbarui pada ' + updatedCount + ' produk!'
    };
  } catch (err) {
    return { ok: false, message: 'Gagal menetapkan akun: ' + err.message };
  }
}

function setProductCover(productId, coverUrl) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Products');
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var coverCol = headers.indexOf('cover_url') + 1;
    if (coverCol <= 0) coverCol = 10;
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(productId)) {
        sheet.getRange(i + 1, coverCol).setValue(coverUrl);
        return { ok: true, message: 'Cover produk berhasil disetel!' };
      }
    }
    return { ok: false, message: 'Produk tidak ditemukan' };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

function deleteProduct(productId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var pSheet = ss.getSheetByName('Products');
    var pValues = pSheet.getDataRange().getValues();
    for (var i = 1; i < pValues.length; i++) {
      if (String(pValues[i][0]) === String(productId)) {
        pSheet.deleteRow(i + 1);
        break;
      }
    }
    // Cleanup associated media files
    var mSheet = ss.getSheetByName('Media_Files');
    var mValues = mSheet.getDataRange().getValues();
    for (var j = mValues.length - 1; j >= 1; j--) {
      if (String(mValues[j][2]) === String(productId)) {
        mSheet.deleteRow(j + 1);
      }
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

/**
 * Helper: Hapus Aset Fisik di Server Cloudinary secara Permanen via Admin/Destroy API (HTTP Basic Auth)
 */
function destroyCloudinaryAsset(publicId, resourceType) {
  if (!publicId) return;
  try {
    var settingsRows = readSheetAsJson('Settings');
    var cloudName = '', apiKey = '', apiSecret = '';
    for (var i = 0; i < settingsRows.length; i++) {
      if (settingsRows[i].setting_key === 'cloudinary_cloud_name') cloudName = settingsRows[i].setting_value;
      if (settingsRows[i].setting_key === 'cloudinary_api_key') apiKey = settingsRows[i].setting_value;
      if (settingsRows[i].setting_key === 'cloudinary_api_secret') apiSecret = settingsRows[i].setting_value;
    }

    if (!cloudName || !apiKey || !apiSecret) {
      Logger.log('Penghapusan Cloudinary dilewati: cloud_name, api_key, atau api_secret belum disetel di Settings.');
      return;
    }

    var resType = (resourceType && String(resourceType).toUpperCase() === 'VIDEO') ? 'video' : 'image';
    var deleteUrl = 'https://api.cloudinary.com/v1_1/' + cloudName + '/resources/' + resType + '/upload?public_ids[]=' + encodeURIComponent(publicId);
    var authHeader = 'Basic ' + Utilities.base64Encode(apiKey + ':' + apiSecret);

    var resp = UrlFetchApp.fetch(deleteUrl, {
      method: 'delete',
      headers: {
        'Authorization': authHeader
      },
      muteHttpExceptions: true
    });

    Logger.log('Cloudinary Destroy Status [' + resp.getResponseCode() + '] Asset ID ' + publicId + ': ' + resp.getContentText());
  } catch (e) {
    Logger.log('Exception destroyCloudinaryAsset: ' + e.message);
  }
}

function deleteMediaFile(mediaId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Media_Files');
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var pubIdCol = headers.indexOf('cloudinary_public_id');
    var typeCol = headers.indexOf('media_type');

    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(mediaId)) {
        var pubId = pubIdCol >= 0 ? values[i][pubIdCol] : '';
        var mType = typeCol >= 0 ? values[i][typeCol] : 'IMAGE';
        if (pubId) {
          destroyCloudinaryAsset(pubId, mType);
        }
        sheet.deleteRow(i + 1);
        return { ok: true, message: 'Media berhasil dihapus dari Cloudinary & database' };
      }
    }
    return { ok: false, message: 'Media tidak ditemukan' };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

function deleteMediaBatch(batchId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var bSheet = ss.getSheetByName('Media_Batches');
    var bValues = bSheet.getDataRange().getValues();
    for (var i = 1; i < bValues.length; i++) {
      if (String(bValues[i][0]) === String(batchId)) {
        bSheet.deleteRow(i + 1);
        break;
      }
    }

    var mSheet = ss.getSheetByName('Media_Files');
    var mValues = mSheet.getDataRange().getValues();
    var headers = mValues[0];
    var pubIdCol = headers.indexOf('cloudinary_public_id');
    var typeCol = headers.indexOf('media_type');

    for (var j = mValues.length - 1; j >= 1; j--) {
      if (String(mValues[j][1]) === String(batchId)) {
        var pubId = pubIdCol >= 0 ? mValues[j][pubIdCol] : '';
        var mType = typeCol >= 0 ? mValues[j][typeCol] : 'IMAGE';
        if (pubId) {
          destroyCloudinaryAsset(pubId, mType);
        }
        mSheet.deleteRow(j + 1);
      }
    }
    return { ok: true, message: 'Batch media dan file Cloudinary berhasil dihapus' };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

/**
 * Mengubah dan menyimpan urutan media tanpa merusak relasi database
 */
function reorderMediaBatch(batchId, direction) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Media_Batches');
    var values = sheet.getDataRange().getValues();
    var curRowIdx = -1;
    var productId = '';

    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(batchId)) {
        curRowIdx = i;
        productId = String(values[i][1]);
        break;
      }
    }
    if (curRowIdx === -1) return { ok: false, message: 'Kelompok media tidak ditemukan' };

    var targetRowIdx = -1;
    if (direction < 0) {
      for (var u = curRowIdx - 1; u >= 1; u--) {
        if (String(values[u][1]) === productId) {
          targetRowIdx = u;
          break;
        }
      }
    } else {
      for (var d = curRowIdx + 1; d < values.length; d++) {
        if (String(values[d][1]) === productId) {
          targetRowIdx = d;
          break;
        }
      }
    }

    if (targetRowIdx === -1) {
      return { ok: false, message: 'Kelompok media sudah di posisi paling ' + (direction < 0 ? 'atas' : 'bawah') };
    }

    var rowA = values[curRowIdx];
    var rowB = values[targetRowIdx];
    sheet.getRange(curRowIdx + 1, 1, 1, rowA.length).setValues([rowB]);
    sheet.getRange(targetRowIdx + 1, 1, 1, rowB.length).setValues([rowA]);

    return { ok: true, message: 'Urutan kelompok media berhasil dipindahkan' };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

function reorderMediaFile(mediaId, direction) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Media_Files');
    var values = sheet.getDataRange().getValues();
    var idx = -1;
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(mediaId)) {
        idx = i;
        break;
      }
    }
    if (idx === -1) return { ok: false, message: 'Media tidak ditemukan' };

    var targetIdx = idx + direction;
    if (targetIdx < 1 || targetIdx >= values.length) {
      return { ok: false, message: 'Posisi media sudah berada di ujung' };
    }

    var rowA = values[idx];
    var rowB = values[targetIdx];
    sheet.getRange(idx + 1, 1, 1, rowA.length).setValues([rowB]);
    sheet.getRange(targetIdx + 1, 1, 1, rowB.length).setValues([rowA]);

    return { ok: true, message: 'Urutan media berhasil diperbarui' };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

function saveMediaBatch(productId, mediaFiles) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var batchId = 'batch_' + Utilities.getUuid().slice(0, 8);
    var nowIso = new Date().toISOString();

    var batchSheet = ss.getSheetByName('Media_Batches');
    var batchNumber = batchSheet.getLastRow();
    batchSheet.appendRow([batchId, productId, batchNumber, nowIso]);

    var mediaSheet = ss.getSheetByName('Media_Files');
    for (var i = 0; i < mediaFiles.length; i++) {
      var m = mediaFiles[i];
      var mediaId = 'med_' + Utilities.getUuid().slice(0, 8);
      var cloudUrl = m.cloudinary_url || m.url || '';
      var thumbUrl = m.thumbnail_url || cloudUrl;

      // Otomatis ganti ekstensi video ke .jpg untuk thumbnail preview Cloudinary
      if ((m.media_type === 'VIDEO' || m.type === 'VIDEO') && cloudUrl) {
        thumbUrl = cloudUrl.replace(/\.[^/.]+$/, '.jpg');
      }

      // Pastikan cloudinary_public_id diekstrak dari URL jika tidak dikirim eksplisit
      var pubId = m.cloudinary_public_id || m.public_id || '';
      if (!pubId && cloudUrl) {
        var match = String(cloudUrl).match(/\/upload\/(?:v\d+\/)?([^\.\?#]+)/);
        if (match && match[1]) {
          pubId = decodeURIComponent(match[1]);
        }
      }
      if (!pubId) {
        var dateStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'ddMMyyyy');
        pubId = 'media_' + dateStr + '_' + Utilities.getUuid().slice(0, 4) + '_' + (i + 1);
      }

      mediaSheet.appendRow([
        mediaId,
        batchId,
        productId,
        m.media_type || m.type || 'IMAGE',
        cloudUrl,
        thumbUrl,
        pubId,
        nowIso
      ]);
    }
    return { ok: true, batch_id: batchId };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

/**
 * Post & Multi-Chain Saving with Strict Account & Duplicate Prevention
 */
function savePostWithChains(postData, chains) {
  try {
    if (!Array.isArray(chains) || chains.length === 0) {
      return { ok: false, message: 'Rantai postingan tidak boleh kosong' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var chainsSheet = ss.getSheetByName('Post_Chains');
    var mediaSheet = ss.getSheetByName('Media_Files');
    var mediaRows = mediaSheet ? mediaSheet.getDataRange().getValues() : [];
    var nowIso = new Date().toISOString();

    // Validasi daftar akun resmi yang terdaftar
    var accountsSheet = ss.getSheetByName('Accounts');
    var accValues = accountsSheet ? accountsSheet.getDataRange().getValues() : [];
    var validAccounts = {};
    for (var v = 1; v < accValues.length; v++) {
      var regId = String(accValues[v][0] || '').trim();
      if (regId && regId !== 'undefined' && regId !== 'null') {
        validAccounts[regId] = true;
      }
    }

    var isInterval = postData.publishing_mode === 'interval';
    var viewCheckUntil = '';
    if (postData.publishing_mode === 'target_views') {
      var fallbackHours = 24;
      try {
        var sRows = readSheetAsJson('Settings');
        for (var si = 0; si < sRows.length; si++) {
          if (sRows[si].setting_key === 'default_fallback_views_hours') {
            fallbackHours = Number(sRows[si].setting_value) || 24;
            break;
          }
        }
      } catch (e) {}
      var expiryDate = new Date();
      expiryDate.setHours(expiryDate.getHours() + fallbackHours);
      viewCheckUntil = expiryDate.toISOString();
    }

    // Kasus 1: Mengedit postingan tunggal yang sudah ada
    if (postData.post_id) {
      var postId = String(postData.post_id).trim();
      var editAccId = String(postData.account_id || (postData.account_ids && postData.account_ids[0]) || '').trim();
      
      if (!editAccId || editAccId.toLowerCase() === 'undefined' || editAccId.toLowerCase() === 'null') {
        return { ok: false, message: 'Akun Threads tidak valid atau belum dipilih' };
      }
      if (Object.keys(validAccounts).length > 0 && !validAccounts[editAccId]) {
        return { ok: false, message: 'Akun Threads tidak terdaftar di sistem' };
      }

      var loc = findPostRowLocation(ss, postId);
      var existingStatus = loc && loc.rowValues ? String(loc.rowValues[6] || '') : '';
      var existingCheckUntil = loc && loc.rowValues ? String(loc.rowValues[8] || '') : '';

      // Anti-downgrade: jika status eksisting sudah Waiting Views, pertahankan status tersebut
      var targetStatus = postData.status || 'Scheduled';
      if (existingStatus === 'Waiting Views') {
        targetStatus = 'Waiting Views';
        if (existingCheckUntil) {
          viewCheckUntil = existingCheckUntil;
        }
      }

      var postRow = {
        post_id: postId,
        account_id: editAccId,
        product_id: postData.product_id || '',
        publishing_mode: postData.publishing_mode,
        interval_minutes: postData.interval_minutes || 3,
        target_views: isInterval ? '' : (postData.target_views || 100),
        status: targetStatus,
        scheduled_at: toIsoUtcString(postData.scheduled_at) || nowIso,
        view_check_until: toIsoUtcString(viewCheckUntil),
        is_evergreen: postData.is_evergreen || 'FALSE',
        error_log: '',
        created_at: toIsoUtcString(loc ? loc.rowValues[11] : nowIso) || nowIso
      };

      var isTargetPublished = (postRow.status === 'Published' || postRow.status === 'Failed');
      var targetSheetName = isTargetPublished ? 'Published_Posts' : 'Queue_Posts';

      if (loc) {
        if (loc.sheetName === targetSheetName) {
          var headers = loc.headers;
          var updateRow = [];
          for (var h = 0; h < headers.length; h++) {
            updateRow.push(postRow[headers[h]] !== undefined ? postRow[headers[h]] : '');
          }
          loc.sheet.getRange(loc.rowIndex, 1, 1, updateRow.length).setValues([updateRow]);
        } else {
          loc.sheet.deleteRow(loc.rowIndex);
          appendRowData(targetSheetName, postRow);
        }
      } else {
        appendRowData(targetSheetName, postRow);
      }

      // Simpan metadata rantai yang sudah terbit sebelum pembaruan teks
      var existingChainsMap = {};
      var cValues = chainsSheet.getDataRange().getValues();
      for (var j = cValues.length - 1; j >= 1; j--) {
        if (String(cValues[j][1]) === postId) {
          var seq = Number(cValues[j][2]);
          existingChainsMap[seq] = {
            chain_id: cValues[j][0],
            temp_container_id: cValues[j][5] || '',
            threads_post_id: cValues[j][6] || '',
            permalink: cValues[j][7] || '',
            published_at: cValues[j][8] || ''
          };
          chainsSheet.deleteRow(j + 1);
        }
      }

      for (var i = 0; i < chains.length; i++) {
        var c = chains[i];
        var seqNum = i + 1;
        var prevMeta = existingChainsMap[seqNum] || {};
        var chainId = c.chain_id || prevMeta.chain_id || ('chn_' + Utilities.getUuid().slice(0, 8));
        var snapshotList = createMediaSnapshot(c.media_ids || [], mediaRows);
        var mediaIdsJson = JSON.stringify(snapshotList);
        chainsSheet.appendRow([
          chainId,
          postId,
          seqNum,
          c.content_text,
          mediaIdsJson,
          c.temp_container_id || prevMeta.temp_container_id || '',
          c.threads_post_id || prevMeta.threads_post_id || '',
          c.permalink || prevMeta.permalink || '',
          toIsoUtcString(c.published_at || prevMeta.published_at)
        ]);
      }

      return { ok: true, post_id: postId, post_ids: [postId] };
    }

    // Kasus 2: Multi-Akun posting baru (Sanitasi ketat & deduplikasi anti double-post)
    var rawCandidates = [];
    if (Array.isArray(postData.account_ids)) {
      rawCandidates = rawCandidates.concat(postData.account_ids);
    }
    if (postData.account_id) {
      rawCandidates.push(postData.account_id);
    }

    var targetAccounts = [];
    var seenAccounts = {};
    for (var a = 0; a < rawCandidates.length; a++) {
      var candidate = String(rawCandidates[a] || '').trim();
      if (candidate && candidate.toLowerCase() !== 'undefined' && candidate.toLowerCase() !== 'null') {
        if (!seenAccounts[candidate] && targetAccounts.indexOf(candidate) === -1) {
          if (Object.keys(validAccounts).length === 0 || validAccounts[candidate]) {
            seenAccounts[candidate] = true;
            targetAccounts.push(candidate);
          }
        }
      }
    }

    if (targetAccounts.length === 0) {
      return { ok: false, message: 'Pilih setidaknya satu akun Threads yang valid' };
    }

    var createdPostIds = [];
    for (var k = 0; k < targetAccounts.length; k++) {
      var currentAccId = targetAccounts[k];
      var newPostId = 'post_' + Utilities.getUuid().slice(0, 8);
      createdPostIds.push(newPostId);

      var newPostRow = {
        post_id: newPostId,
        account_id: currentAccId,
        product_id: postData.product_id || '',
        publishing_mode: postData.publishing_mode,
        interval_minutes: postData.interval_minutes || 3,
        target_views: isInterval ? '' : (postData.target_views || 100),
        status: postData.status || 'Scheduled',
        scheduled_at: toIsoUtcString(postData.scheduled_at) || nowIso,
        view_check_until: toIsoUtcString(viewCheckUntil),
        is_evergreen: postData.is_evergreen || 'FALSE',
        error_log: '',
        created_at: nowIso
      };
      appendRowData('Queue_Posts', newPostRow);

      for (var ch = 0; ch < chains.length; ch++) {
        var chainItem = chains[ch];
        var newChainId = 'chn_' + Utilities.getUuid().slice(0, 8);
        var snapshotList = createMediaSnapshot(chainItem.media_ids || [], mediaRows);
        var mIdsJson = JSON.stringify(snapshotList);
        chainsSheet.appendRow([
          newChainId,
          newPostId,
          ch + 1,
          chainItem.content_text,
          mIdsJson,
          '',
          '',
          '',
          ''
        ]);
      }
    }

    return { ok: true, post_id: createdPostIds[0], post_ids: createdPostIds };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

function deletePost(postId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var loc = findPostRowLocation(ss, postId);
    if (loc) {
      loc.sheet.deleteRow(loc.rowIndex);
    }

    var chainSheet = ss.getSheetByName('Post_Chains');
    if (chainSheet) {
      var cValues = chainSheet.getDataRange().getValues();
      for (var j = cValues.length - 1; j >= 1; j--) {
        if (String(cValues[j][1]) === String(postId)) {
          chainSheet.deleteRow(j + 1);
        }
      }
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

/**
 * Smart Best-Time Posting Algorithm with Jitter & Collision Avoidance
 */
function suggestBestPostingTime(accountId) {
  try {
    var posts = getAllPostsAsJson();
    var sRows = readSheetAsJson('Settings');
    var settingsMap = {};
    for (var s = 0; s < sRows.length; s++) {
      settingsMap[sRows[s].setting_key] = sRows[s].setting_value;
    }
    var tz = settingsMap['timezone'] || 'Asia/Jakarta';

    // Jendela Waktu Rekomendasi (Pagi, Siang, Sore, Malam)
    var windows = [
      { name: 'Pagi', start: settingsMap['best_time_pagi_start'] || '08:00', end: settingsMap['best_time_pagi_end'] || '10:00' },
      { name: 'Siang', start: settingsMap['best_time_siang_start'] || '12:00', end: settingsMap['best_time_siang_end'] || '14:00' },
      { name: 'Sore', start: settingsMap['best_time_sore_start'] || '16:00', end: settingsMap['best_time_sore_end'] || '18:00' },
      { name: 'Malam', start: settingsMap['best_time_malam_start'] || '19:00', end: settingsMap['best_time_malam_end'] || '22:00' }
    ];

    var now = new Date();
    var nowFormatted = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm');
    var timeParts = nowFormatted.split(' ')[1].split(':');
    var nowMinutes = Number(timeParts[0]) * 60 + Number(timeParts[1]);

    var selectedWindow = null;
    var targetDateOffsetDays = 0;
    var targetMinutes = 0;

    // Cari jendela waktu hari ini yang masih di masa depan (minimal +15 menit)
    for (var w = 0; w < windows.length; w++) {
      var wItem = windows[w];
      var sParts = (wItem.start || '00:00').split(':');
      var eParts = (wItem.end || '23:59').split(':');
      var sMin = Number(sParts[0]) * 60 + Number(sParts[1]);
      var eMin = Number(eParts[0]) * 60 + Number(eParts[1]);

      if (nowMinutes + 15 <= sMin) {
        selectedWindow = wItem;
        targetMinutes = sMin;
        break;
      } else if (nowMinutes + 15 <= eMin) {
        selectedWindow = wItem;
        targetMinutes = nowMinutes + 15;
        break;
      }
    }

    // Jika seluruh jendela hari ini sudah terlewat, pilih jendela pertama besok pagi
    if (!selectedWindow) {
      targetDateOffsetDays = 1;
      selectedWindow = windows[0];
      var firstStart = (selectedWindow.start || '08:00').split(':');
      targetMinutes = Number(firstStart[0]) * 60 + Number(firstStart[1]);
    }

    // Tambah variasi jitter anti-fingerprinting (2 sampai 7 menit)
    var jitter = Math.floor(Math.random() * 6) + 2;
    targetMinutes += jitter;

    var targetBase = new Date(now.getTime() + targetDateOffsetDays * 24 * 60 * 60 * 1000);
    var baseDateStr = Utilities.formatDate(targetBase, tz, 'yyyy/MM/dd');
    var targetHour = Math.floor(targetMinutes / 60);
    var targetMinute = targetMinutes % 60;
    var candidateDate = new Date(baseDateStr + ' ' + ('0' + targetHour).slice(-2) + ':' + ('0' + targetMinute).slice(-2) + ':00');

    // Validasi mutlak: waktu rekomendasi harus di masa depan
    if (candidateDate.getTime() <= now.getTime()) {
      candidateDate = new Date(now.getTime() + 20 * 60 * 1000);
    }

    // Pemeriksaan tabrakan jadwal antrean untuk akun ini (buffer aman 60 menit)
    var collisionFound = true;
    var loops = 0;
    while (collisionFound && loops < 12) {
      collisionFound = false;
      loops++;
      for (var p = 0; p < posts.length; p++) {
        var postItem = posts[p];
        if (String(postItem.account_id) === String(accountId) && postItem.status === 'Scheduled' && postItem.scheduled_at) {
          var schedTime = new Date(postItem.scheduled_at).getTime();
          var diffMin = Math.abs(schedTime - candidateDate.getTime()) / (60 * 1000);
          if (diffMin < 60) {
            candidateDate.setMinutes(candidateDate.getMinutes() + 65);
            collisionFound = true;
            break;
          }
        }
      }
    }

    var finalFormatted = Utilities.formatDate(candidateDate, tz, "yyyy-MM-dd'T'HH:mm");
    return {
      ok: true,
      data: {
        suggested_time: finalFormatted,
        reason: 'Jendela ' + selectedWindow.name + ' (' + selectedWindow.start + '-' + selectedWindow.end + ') + jitter ' + jitter + ' mnt'
      }
    };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

/**
 * Evergreen Post Recycler
 */
function clonePostToAccount(sourcePostId, targetAccountId, rewriteWithAi) {
  try {
    var posts = getAllPostsAsJson();
    var chains = readSheetAsJson('Post_Chains');

    var sourcePost = null;
    for (var i = 0; i < posts.length; i++) {
      if (posts[i].post_id === sourcePostId) {
        sourcePost = posts[i];
        break;
      }
    }
    if (!sourcePost) return { ok: false, message: 'Sumber post tidak ditemukan' };

    var sourceChains = [];
    for (var j = 0; j < chains.length; j++) {
      if (chains[j].post_id === sourcePostId) {
        sourceChains.push(chains[j]);
      }
    }
    sourceChains.sort(function(a, b) { return a.sequence_order - b.sequence_order; });

    var newChains = [];
    for (var k = 0; k < sourceChains.length; k++) {
      var item = sourceChains[k];
      var text = item.content_text;
      var mediaList = [];
      try { mediaList = JSON.parse(item.media_ids || '[]'); } catch (e) { mediaList = []; }

      newChains.push({
        content_text: text,
        media_ids: mediaList
      });
    }

    var schedTime = new Date();
    schedTime.setHours(schedTime.getHours() + 2);

    var newPostData = {
      account_id: targetAccountId,
      product_id: sourcePost.product_id,
      publishing_mode: sourcePost.publishing_mode,
      interval_minutes: sourcePost.interval_minutes,
      target_views: sourcePost.target_views,
      status: 'Scheduled',
      scheduled_at: schedTime.toISOString()
    };

    return savePostWithChains(newPostData, newChains);
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

/**
 * Settings Persistence
 */
function saveSettings(settingsObj) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Settings');
    var values = sheet.getDataRange().getValues();

    for (var key in settingsObj) {
      var found = false;
      for (var i = 1; i < values.length; i++) {
        if (values[i][0] === key) {
          sheet.getRange(i + 1, 2).setValue(settingsObj[key]);
          found = true;
          break;
        }
      }
      if (!found) {
        sheet.appendRow([key, settingsObj[key]]);
      }
    }
    return { ok: true, message: 'Pengaturan tersimpan' };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

/**
 * CRUD Pustaka Konten Viral (Bank Inspirasi)
 */
function saveViralContent(viralData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Viral_Library');
    var vId = viralData.viral_id;
    var nowIso = new Date().toISOString();
    var chainsJson = JSON.stringify(viralData.chains || []);

    var autoTitle = viralData.title ? viralData.title.trim() : '';
    if (!autoTitle) {
      var chainsArr = viralData.chains || [];
      if (chainsArr.length > 0 && chainsArr[0].content_text) {
        var t = chainsArr[0].content_text.trim();
        autoTitle = t.slice(0, 30) + (t.length > 30 ? '...' : '');
      } else {
        autoTitle = 'Inspirasi Utas';
      }
    }

    var targetProductIds = '';
    if (Array.isArray(viralData.product_ids)) {
      targetProductIds = viralData.product_ids.map(function(id) { return String(id).trim(); }).filter(Boolean).join(',');
    } else if (viralData.product_ids) {
      targetProductIds = String(viralData.product_ids).trim();
    } else if (viralData.product_id) {
      targetProductIds = String(viralData.product_id).trim();
    }

    var values = sheet.getDataRange().getValues();
    var headers = values[0];

    var colMap = {
      viral_id: headers.indexOf('viral_id'),
      title: headers.indexOf('title'),
      source_url: headers.indexOf('source_url'),
      chains_json: headers.indexOf('chains_json'),
      product_ids: headers.indexOf('product_ids') !== -1 ? headers.indexOf('product_ids') : headers.indexOf('product_id'),
      created_at: headers.indexOf('created_at')
    };

    var cleanSourceUrl = (viralData.source_url || '').trim();
    if (cleanSourceUrl) {
      var srcCol = colMap.source_url >= 0 ? colMap.source_url : 2;
      var vIdCol = colMap.viral_id >= 0 ? colMap.viral_id : 0;
      var titleCol = colMap.title >= 0 ? colMap.title : 1;
      for (var s = 1; s < values.length; s++) {
        var rowSrc = String(values[s][srcCol] || '').trim();
        var rowVid = String(values[s][vIdCol] || '').trim();
        if (rowSrc && rowSrc.toLowerCase() === cleanSourceUrl.toLowerCase()) {
          if (!vId || rowVid !== String(vId)) {
            var existingTitle = values[s][titleCol] || 'Referensi';
            return { ok: false, message: 'Tautan sumber rujukan ini sudah tersimpan: "' + existingTitle + '"' };
          }
        }
      }
    }

    if (vId) {
      var rowIdx = -1;
      var idCol = colMap.viral_id >= 0 ? colMap.viral_id : 0;
      for (var i = 1; i < values.length; i++) {
        if (String(values[i][idCol]) === String(vId)) {
          rowIdx = i + 1;
          break;
        }
      }

      if (rowIdx > 0) {
        if (colMap.title >= 0) sheet.getRange(rowIdx, colMap.title + 1).setValue(autoTitle);
        if (colMap.source_url >= 0) sheet.getRange(rowIdx, colMap.source_url + 1).setValue(viralData.source_url || '');
        if (colMap.chains_json >= 0) sheet.getRange(rowIdx, colMap.chains_json + 1).setValue(chainsJson);
        if (colMap.product_ids >= 0) sheet.getRange(rowIdx, colMap.product_ids + 1).setValue(targetProductIds);
        return { ok: true, message: 'Konten viral berhasil diperbarui' };
      }
    }

    var newId = 'vir_' + Utilities.getUuid().slice(0, 8);
    var rowObj = {
      viral_id: newId,
      title: autoTitle,
      source_url: viralData.source_url || '',
      chains_json: chainsJson,
      product_ids: targetProductIds,
      created_at: nowIso
    };

    var newRow = [];
    for (var h = 0; h < headers.length; h++) {
      var hName = headers[h];
      var cellVal = rowObj[hName];
      if (cellVal === undefined && (hName === 'product_id' || hName === 'product_ids')) {
        cellVal = targetProductIds;
      }
      newRow.push(cellVal !== undefined ? cellVal : '');
    }
    sheet.appendRow(newRow);

    return { ok: true, message: 'Konten viral berhasil disimpan ke pustaka' };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

function deleteViralContent(viralId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Viral_Library');
    var values = sheet.getDataRange().getValues();
    if (values.length <= 1) return { ok: false, message: 'Data tidak ditemukan' };

    var headers = values[0];
    var idCol = headers.indexOf('viral_id');
    var chainsCol = headers.indexOf('chains_json');
    if (idCol === -1) idCol = 0;

    for (var i = 1; i < values.length; i++) {
      if (String(values[i][idCol]) === String(viralId)) {
        // Ekstrak dan musnahkan aset fisik di Cloudinary
        if (chainsCol !== -1 && values[i][chainsCol]) {
          try {
            var chains = JSON.parse(values[i][chainsCol]);
            if (Array.isArray(chains)) {
              for (var c = 0; c < chains.length; c++) {
                var ch = chains[c];
                var mList = [];
                if (Array.isArray(ch.media_urls)) {
                  mList = ch.media_urls;
                } else if (ch.media_url) {
                  mList = [ch.media_url];
                }

                for (var m = 0; m < mList.length; m++) {
                  var mItem = mList[m];
                  var urlStr = '';
                  var pubId = '';
                  var resType = 'image';

                  if (typeof mItem === 'object' && mItem !== null) {
                    pubId = mItem.public_id || mItem.cloudinary_public_id || '';
                    urlStr = mItem.url || mItem.secure_url || mItem.cloudinary_url || '';
                    resType = (mItem.resource_type || mItem.media_type || '').toLowerCase() === 'video' ? 'video' : 'image';
                  } else if (typeof mItem === 'string') {
                    urlStr = mItem;
                  }

                  if (!pubId && urlStr) {
                    var isVid = /\.(mp4|mov|webm)$/i.test(urlStr) || urlStr.indexOf('/video/') !== -1;
                    resType = isVid ? 'video' : 'image';
                    var match = urlStr.match(/\/upload\/(?:v\d+\/)?([^\.\?#]+)/);
                    if (match && match[1]) {
                      pubId = decodeURIComponent(match[1]);
                    }
                  }

                  if (pubId) {
                    destroyCloudinaryAsset(pubId, resType);
                  }
                }
              }
            }
          } catch (parseErr) {
            Logger.log('Gagal mengurai chains_json saat menghapus aset fisik: ' + parseErr.message);
          }
        }

        sheet.deleteRow(i + 1);
        return { ok: true, message: 'Konten viral dan aset fisik Cloudinary berhasil dihapus dari pustaka' };
      }
    }
    return { ok: false, message: 'Data tidak ditemukan' };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

/**
 * Sistem Trigger Modular 2-Mode: Cek Status Terpasang & Interval
 */
function getTriggerStatus() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    var publisherTrigger = null;
    var insightsTrigger = null;

    for (var i = 0; i < triggers.length; i++) {
      var handler = triggers[i].getHandlerFunction();
      if (handler === 'processPublishingQueue') {
        publisherTrigger = triggers[i];
      } else if (handler === 'syncPostAnalytics') {
        insightsTrigger = triggers[i];
      }
    }

    var settingsRows = readSheetAsJson('Settings');
    var pubInterval = 1;
    var insInterval = 5;
    for (var s = 0; s < settingsRows.length; s++) {
      if (settingsRows[s].setting_key === 'trigger_publisher_interval') {
        pubInterval = Number(settingsRows[s].setting_value) || 1;
      }
      if (settingsRows[s].setting_key === 'trigger_insights_interval') {
        insInterval = Number(settingsRows[s].setting_value) || 5;
      }
    }

    return {
      ok: true,
      data: {
        publisher: {
          enabled: publisherTrigger !== null,
          interval: pubInterval
        },
        insights: {
          enabled: insightsTrigger !== null,
          interval: insInterval
        }
      }
    };
  } catch (err) {
    return { ok: false, message: 'Gagal membaca status trigger: ' + err.message };
  }
}

/**
 * Memasang atau mencopot trigger secara modular per jenis cron
 */
function updateTriggerConfig(triggerType, intervalMinutes, enable) {
  try {
    var targetHandler = '';
    if (triggerType === 'publisher') {
      targetHandler = 'processPublishingQueue';
    } else if (triggerType === 'insights') {
      targetHandler = 'syncPostAnalytics';
    } else {
      return { ok: false, message: 'Tipe trigger tidak dikenal' };
    }

    var triggers = ScriptApp.getProjectTriggers();

    // Hapus trigger sejenis yang sudah aktif sebelumnya
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === targetHandler) {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }

    var intervalNum = Number(intervalMinutes);
    var settingUpdates = {};

    if (enable) {
      if (triggerType === 'publisher') {
        if ([1, 5, 10].indexOf(intervalNum) === -1) intervalNum = 1;
        ScriptApp.newTrigger(targetHandler)
          .timeBased()
          .everyMinutes(intervalNum)
          .create();
        settingUpdates['trigger_publisher_interval'] = String(intervalNum);
      } else if (triggerType === 'insights') {
        if (intervalNum >= 60) {
          var hours = Math.round(intervalNum / 60);
          if (hours !== 1 && hours !== 6) hours = 1;
          ScriptApp.newTrigger(targetHandler)
            .timeBased()
            .everyHours(hours)
            .create();
        } else {
          var validMin = [5, 10, 15, 30];
          if (validMin.indexOf(intervalNum) === -1) intervalNum = 5;
          ScriptApp.newTrigger(targetHandler)
            .timeBased()
            .everyMinutes(intervalNum)
            .create();
        }
        settingUpdates['trigger_insights_interval'] = String(intervalNum);
      }
    }

    // Perbarui status umum triggers_installed jika salah satu cron masih terpasang
    var remaining = ScriptApp.getProjectTriggers();
    var anyActive = false;
    for (var r = 0; r < remaining.length; r++) {
      var h = remaining[r].getHandlerFunction();
      if (h === 'processPublishingQueue' || h === 'syncPostAnalytics') {
        anyActive = true;
        break;
      }
    }
    settingUpdates['triggers_installed'] = anyActive ? 'TRUE' : 'FALSE';
    saveSettings(settingUpdates);

    var label = triggerType === 'publisher' ? 'Publisher Antrean' : 'Sinkronisasi Insights';
    var msg = enable ? ('Trigger ' + label + ' berhasil diaktifkan!') : ('Trigger ' + label + ' berhasil dihentikan.');
    return { ok: true, message: msg };
  } catch (err) {
    return { ok: false, message: 'Gagal memperbarui trigger: ' + err.message };
  }
}

function toggleAutomationTriggers(enable) {
  var pRes = updateTriggerConfig('publisher', 1, enable);
  var iRes = updateTriggerConfig('insights', 5, enable);
  return { ok: pRes.ok && iRes.ok, message: enable ? 'Seluruh trigger otomatis berhasil diaktifkan!' : 'Semua triggers berhasil dimatikan.' };
}

/**
 * PUBLISH NOW & QUEUE ENGINE
 */
function publishPostNow(postId) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    return { ok: false, message: 'Server sedang memproses antrean lain. Silakan coba kembali.' };
  }
  try {
    return executeSinglePostPublish(postId, false);
  } finally {
    lock.releaseLock();
  }
}

function forceReleasePost(postId) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    return { ok: false, message: 'Server sedang memproses antrean lain. Silakan coba kembali.' };
  }
  try {
    return executeSinglePostPublish(postId, true);
  } finally {
    lock.releaseLock();
  }
}

function saveAndPublishNow(postPayload, chains) {
  var saveRes = savePostWithChains(postPayload, chains);
  if (!saveRes.ok) return saveRes;
  var postIds = saveRes.post_ids || [saveRes.post_id];
  var lastResult = null;
  for (var i = 0; i < postIds.length; i++) {
    lastResult = publishPostNow(postIds[i]);
  }
  return lastResult || saveRes;
}

function syncAllInsightsNow() {
  syncPostAnalytics();
  return { ok: true, message: 'Metrik insights berhasil disinkronkan langsung dari Threads API!' };
}

function syncSinglePostInsights(postId) {
  try {
    var chains = readSheetAsJson('Post_Chains');
    var posts = getAllPostsAsJson();
    var post = null;
    for (var p = 0; p < posts.length; p++) {
      if (String(posts[p].post_id) === String(postId)) {
        post = posts[p];
        break;
      }
    }
    if (!post) return { ok: false, message: 'Postingan tidak ditemukan' };

    var accounts = readSheetAsJson('Accounts');
    var account = null;
    for (var a = 0; a < accounts.length; a++) {
      if (String(accounts[a].account_id) === String(post.account_id)) {
        account = accounts[a];
        break;
      }
    }
    if (!account || !account.access_token) {
      return { ok: false, message: 'Token akses akun Threads tidak ditemukan' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var analyticsSheet = ss.getSheetByName('Post_Analytics');
    var anaRows = analyticsSheet.getDataRange().getValues();
    var syncedCount = 0;

    for (var i = 0; i < chains.length; i++) {
      if (String(chains[i].post_id) === String(postId) && chains[i].threads_post_id) {
        var thId = String(chains[i].threads_post_id).trim();
        var cId = chains[i].chain_id;
        try {
          var insightUrl = 'https://graph.threads.net/v1.0/' + encodeURIComponent(thId) + '/insights?metric=views,likes,replies,reposts&access_token=' + encodeURIComponent(account.access_token);
          var resp = UrlFetchApp.fetch(insightUrl, { muteHttpExceptions: true });
          var data = JSON.parse(resp.getContentText());
          if (data && data.data && Array.isArray(data.data)) {
            var views = 0, likes = 0, replies = 0, reposts = 0;
            for (var m = 0; m < data.data.length; m++) {
              var item = data.data[m];
              var rawVal = item.value !== undefined ? item.value : (item.total_value ? item.total_value.value : (item.values && item.values.length > 0 ? item.values[0].value : 0));
              var val = Number(rawVal || 0);
              if (item.name === 'views') views = val;
              if (item.name === 'likes') likes = val;
              if (item.name === 'replies') replies = val;
              if (item.name === 'reposts') reposts = val;
            }
            var rowIdx = -1;
            for (var r = 1; r < anaRows.length; r++) {
              if (String(anaRows[r][1]) === String(cId)) {
                rowIdx = r + 1;
                break;
              }
            }
            var nowIso = new Date().toISOString();
            if (rowIdx > 0) {
              analyticsSheet.getRange(rowIdx, 3, 1, 5).setValues([[views, likes, reposts, replies, nowIso]]);
            } else {
              var newAnaId = 'ana_' + Utilities.getUuid().slice(0, 8);
              analyticsSheet.appendRow([newAnaId, cId, views, likes, reposts, replies, nowIso]);
              anaRows.push([newAnaId, cId, views, likes, reposts, replies, nowIso]);
            }
            syncedCount++;
          }
        } catch (callErr) {}
      }
    }
    SpreadsheetApp.flush();
    // Sinkronkan juga komentar untuk postingan ini
    try {
      syncPostComments(postId);
    } catch (eComm) {}

    return { ok: true, message: syncedCount > 0 ? 'Metrik dan komentar postingan ini berhasil disinkronkan!' : 'Belum ada utas terbit pada postingan ini.' };
  } catch (err) {
    return { ok: false, message: 'Gagal sinkronisasi metrik: ' + err.message };
  }
}

/**
 * Engine Penarikan Komentar Meta Graph API (GET /{threads-post-id}/replies)
 */
function syncPostComments(postId) {
  try {
    var chains = readSheetAsJson('Post_Chains');
    var posts = getAllPostsAsJson();
    var accounts = readSheetAsJson('Accounts');
    var post = null;
    for (var p = 0; p < posts.length; p++) {
      if (String(posts[p].post_id) === String(postId)) {
        post = posts[p];
        break;
      }
    }
    if (!post) return { ok: false, message: 'Postingan tidak ditemukan' };

    var account = null;
    for (var a = 0; a < accounts.length; a++) {
      if (String(accounts[a].account_id) === String(post.account_id)) {
        account = accounts[a];
        break;
      }
    }
    if (!account || !account.access_token) {
      return { ok: false, message: 'Token akses akun Threads tidak ditemukan' };
    }

    var authorUsername = String(account.username || '').toLowerCase().trim();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var commSheet = ss.getSheetByName('Post_Comments');
    if (!commSheet) {
      commSheet = ss.insertSheet('Post_Comments');
      commSheet.appendRow(['comment_id', 'chain_id', 'threads_comment_id', 'from_username', 'text', 'created_at']);
    }
    var existingRows = commSheet.getDataRange().getValues();
    var existingIds = {};
    for (var r = 1; r < existingRows.length; r++) {
      var existThId = String(existingRows[r][2] || '').trim();
      if (existThId) existingIds[existThId] = true;
    }

    var newCommentsCount = 0;
    for (var i = 0; i < chains.length; i++) {
      if (String(chains[i].post_id) === String(postId) && chains[i].threads_post_id) {
        var thId = String(chains[i].threads_post_id).trim();
        var cId = chains[i].chain_id;
        try {
          var repUrl = 'https://graph.threads.net/v1.0/' + encodeURIComponent(thId) + '/replies?fields=id,text,username,timestamp&access_token=' + encodeURIComponent(account.access_token);
          var repResp = UrlFetchApp.fetch(repUrl, { muteHttpExceptions: true });
          var repData = JSON.parse(repResp.getContentText());
          if (repData && repData.data && Array.isArray(repData.data)) {
            for (var m = 0; m < repData.data.length; m++) {
              var rep = repData.data[m];
              var commentThId = String(rep.id || '').trim();
              if (commentThId && !existingIds[commentThId]) {
                var fromUser = rep.username || 'user';
                var isAuthor = String(fromUser).toLowerCase().trim() === authorUsername;
                var textContent = rep.text || '';
                var createdAt = rep.timestamp || new Date().toISOString();
                var newCommId = 'com_' + Utilities.getUuid().slice(0, 8);

                commSheet.appendRow([
                  newCommId,
                  cId,
                  commentThId,
                  (isAuthor ? (fromUser + ' (Author)') : fromUser),
                  textContent,
                  createdAt,
                  (isAuthor ? 'Resolved' : 'Pending'),
                  ''
                ]);
                existingIds[commentThId] = true;
                newCommentsCount++;
              }
            }
          }
        } catch (callErr) {
          Logger.log('Gagal ambil komentar Threads: ' + callErr.message);
        }
      }
    }
    SpreadsheetApp.flush();
    var allComments = readSheetAsJson('Post_Comments');
    return {
      ok: true,
      data: serializeForClient(allComments),
      message: newCommentsCount > 0 ? (newCommentsCount + ' komentar baru berhasil disinkronkan!') : 'Komentar sudah mutakhir.'
    };
  } catch (err) {
    return { ok: false, message: 'Gagal sinkronisasi komentar: ' + err.message };
  }
}

/**
 * Membalas komentar langsung di Threads via Meta Graph API
 */
function replyToThreadsComment(commentId, replyText) {
  try {
    if (!commentId || !replyText || !String(replyText).trim()) {
      return { ok: false, message: 'ID Komentar dan teks balasan wajib diisi' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var commSheet = ss.getSheetByName('Post_Comments');
    if (!commSheet) return { ok: false, message: 'Tabel Post_Comments belum ada' };

    var commRows = commSheet.getDataRange().getValues();
    var headers = commRows[0];
    var cIdCol = headers.indexOf('comment_id');
    var chainIdCol = headers.indexOf('chain_id');
    var thCommIdCol = headers.indexOf('threads_comment_id');
    var statusCol = headers.indexOf('status');
    var replyTextCol = headers.indexOf('reply_text');

    var targetRowIdx = -1;
    var chainId = '';
    var threadsCommentId = '';

    for (var i = 1; i < commRows.length; i++) {
      if (String(commRows[i][cIdCol >= 0 ? cIdCol : 0]) === String(commentId)) {
        targetRowIdx = i + 1;
        chainId = String(commRows[i][chainIdCol >= 0 ? chainIdCol : 1]);
        threadsCommentId = String(commRows[i][thCommIdCol >= 0 ? thCommIdCol : 2]);
        break;
      }
    }

    if (targetRowIdx === -1) {
      return { ok: false, message: 'Komentar tidak ditemukan' };
    }

    var chains = readSheetAsJson('Post_Chains');
    var matchedChain = chains.find(function(c) { return c.chain_id === chainId; });
    if (!matchedChain) return { ok: false, message: 'Rantai postingan tidak ditemukan' };

    var posts = readSheetAsJson('Posts');
    var matchedPost = posts.find(function(p) { return p.post_id === matchedChain.post_id; });
    if (!matchedPost) return { ok: false, message: 'Postingan induk tidak ditemukan' };

    var accounts = readSheetAsJson('Accounts');
    var matchedAcc = accounts.find(function(a) { return a.account_id === matchedPost.account_id; });
    if (!matchedAcc || !matchedAcc.access_token) {
      return { ok: false, message: 'Akses token akun Threads tidak ditemukan' };
    }

    var cleanReply = String(replyText).trim();
    var payloadObj = {
      media_type: 'TEXT',
      text: cleanReply,
      reply_to_id: threadsCommentId
    };

    var containerId = createMetaContainer(matchedAcc.meta_user_id, matchedAcc.access_token, payloadObj);
    var publishedPostId = publishMetaContainer(matchedAcc.meta_user_id, containerId, matchedAcc.access_token);

    if (!publishedPostId) {
      return { ok: false, message: 'Meta Threads tidak mengembalikan ID publikasi balasan' };
    }

    if (statusCol >= 0) commSheet.getRange(targetRowIdx, statusCol + 1).setValue('Replied');
    if (replyTextCol >= 0) commSheet.getRange(targetRowIdx, replyTextCol + 1).setValue(cleanReply);

    return {
      ok: true,
      message: 'Balasan berhasil dikirim langsung ke komentar Threads!',
      data: {
        comment_id: commentId,
        status: 'Replied',
        reply_text: cleanReply
      }
    };
  } catch (err) {
    return { ok: false, message: 'Gagal mengirim balasan: ' + err.message };
  }
}

/**
 * Memperbarui status komentar (Pending / Resolved)
 */
function updateCommentStatus(commentId, newStatus) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var commSheet = ss.getSheetByName('Post_Comments');
    if (!commSheet) return { ok: false, message: 'Tabel Post_Comments tidak ditemukan' };

    var commRows = commSheet.getDataRange().getValues();
    var headers = commRows[0];
    var cIdCol = headers.indexOf('comment_id');
    var statusCol = headers.indexOf('status');

    if (statusCol === -1) {
      statusCol = 6;
    }

    for (var i = 1; i < commRows.length; i++) {
      if (String(commRows[i][cIdCol >= 0 ? cIdCol : 0]) === String(commentId)) {
        commSheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
        return { ok: true, message: 'Status komentar diperbarui menjadi ' + newStatus };
      }
    }
    return { ok: false, message: 'Komentar tidak ditemukan' };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

/**
 * Sinkronisasi seluruh komentar untuk postingan yang berstatus Published/Waiting Views
 */
function syncAllCommentsNow() {
  try {
    var posts = getAllPostsAsJson();
    var count = 0;
    for (var i = 0; i < posts.length; i++) {
      if (posts[i].status === 'Published' || posts[i].status === 'Waiting Views') {
        try {
          var res = syncPostComments(posts[i].post_id);
          if (res && res.ok) count++;
        } catch (e) {}
      }
    }
    var allComments = readSheetAsJson('Post_Comments');
    return {
      ok: true,
      data: serializeForClient(allComments),
      message: 'Komentar dari seluruh postingan berhasil disinkronkan!'
    };
  } catch (err) {
    return { ok: false, message: 'Gagal sinkron komentar: ' + err.message };
  }
}

/**
 * MODUL 1: Pemeriksaan Kesiapan Media (Real-Time Polling Meta Container)
 */
function checkMetaContainerStatus(containerId, token, maxRetries) {
  var retries = maxRetries || 15;
  var checkUrl = 'https://graph.threads.net/v1.0/' + containerId + '?fields=status,error_message&access_token=' + encodeURIComponent(token);
  
  for (var c = 0; c < retries; c++) {
    try {
      var resp = UrlFetchApp.fetch(checkUrl, { muteHttpExceptions: true });
      var data = JSON.parse(resp.getContentText());
      if (data.status === 'FINISHED') {
        return true;
      }
      if (data.status === 'ERROR') {
        throw new Error('Meta Container ' + containerId + ' Error: ' + (data.error_message || 'Gagal diproses Meta'));
      }
    } catch (e) {
      if (e.message.indexOf('Meta Container') !== -1) throw e;
    }
    Utilities.sleep(2000);
  }
  throw new Error('Meta Container ' + containerId + ' batas waktu habis (timeout) menunggu status FINISHED.');
}

/**
 * MODUL 2: Format Pengiriman Data ke Meta (access_token di URL & Body)
 */
function createMetaContainer(metaUserId, token, payloadObj) {
  var url = 'https://graph.threads.net/v1.0/' + metaUserId + '/threads?access_token=' + encodeURIComponent(token);
  payloadObj.access_token = token; // Masukkan access_token di 2 tempat: parameter URL & body

  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    payload: payloadObj,
    muteHttpExceptions: true
  });
  var data = JSON.parse(resp.getContentText());
  if (data.error) {
    throw new Error('Meta Buat Container Gagal: ' + (data.error.message || JSON.stringify(data.error)));
  }
  if (!data.id) {
    throw new Error('Meta tidak mengembalikan Container ID');
  }
  return data.id;
}

function publishMetaContainer(metaUserId, creationId, token) {
  // Titik B: Kontainer wajib berstatus FINISHED sebelum threads_publish
  checkMetaContainerStatus(creationId, token, 15);

  var pUrl = 'https://graph.threads.net/v1.0/' + metaUserId + '/threads_publish?access_token=' + encodeURIComponent(token);
  var payload = {
    creation_id: creationId,
    access_token: token // access_token di 2 tempat
  };
  var res = UrlFetchApp.fetch(pUrl, {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  });
  var data = JSON.parse(res.getContentText());
  if (data.error) {
    throw new Error('Meta Publish Gagal: ' + (data.error.message || JSON.stringify(data.error)));
  }
  if (!data.id) {
    throw new Error('Meta Publish tidak mengembalikan Post ID');
  }
  return data.id;
}

function getOfficialMetaPermalink(publishedPostId, token, username) {
  var permalink = '';
  try {
    var pMetaUrl = 'https://graph.threads.net/v1.0/' + publishedPostId + '?fields=permalink&access_token=' + encodeURIComponent(token);
    var pMetaResp = UrlFetchApp.fetch(pMetaUrl, { muteHttpExceptions: true });
    var pMetaData = JSON.parse(pMetaResp.getContentText());
    if (pMetaData && pMetaData.permalink) {
      permalink = pMetaData.permalink;
    }
  } catch (pErr) {
    // Fallback permalink jika Meta sedang lambat
  }
  if (!permalink) {
    permalink = 'https://www.threads.net/@' + username + '/post/' + publishedPostId;
  }
  return permalink;
}

/**
 * MODUL 3: Mesin Rantai Utas Bersambung (Pemisahan Eksekusi Meta & Penulisan Sheets)
 */
function executeSinglePostPublish(postId, isWaitingViewsCheck) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var chainsSheet = ss.getSheetByName('Post_Chains');
  var mediaSheet = ss.getSheetByName('Media_Files');
  var accounts = readSheetAsJson('Accounts');

  var postLoc = findPostRowLocation(ss, postId);
  if (!postLoc) return { ok: false, message: 'Postingan tidak ditemukan' };

  var currentSheet = postLoc.sheet;
  var pIdx = postLoc.rowIndex;
  var postRow = postLoc.rowValues;

  var post = {
    post_id: postRow[0],
    account_id: postRow[1],
    product_id: postRow[2],
    publishing_mode: postRow[3],
    interval_minutes: postRow[4],
    target_views: Number(postRow[5] || 0),
    status: postRow[6],
    scheduled_at: postRow[7],
    view_check_until: postRow[8]
  };

  var chainRows = chainsSheet.getDataRange().getValues();
  var mediaRows = mediaSheet.getDataRange().getValues();

  var account = null;
  for (var a = 0; a < accounts.length; a++) {
    if (accounts[a].account_id === post.account_id) {
      account = accounts[a];
      break;
    }
  }
  if (!account || !account.access_token) {
    currentSheet.getRange(pIdx, 11).setValue('Akun atau Access Token tidak valid');
    return { ok: false, message: 'Akun atau Access Token tidak valid' };
  }

  var postChains = [];
  for (var cIdx = 1; cIdx < chainRows.length; cIdx++) {
    if (String(chainRows[cIdx][1]) === String(postId)) {
      postChains.push({
        rowIndex: cIdx + 1,
        chain_id: chainRows[cIdx][0],
        sequence_order: Number(chainRows[cIdx][2]),
        content_text: chainRows[cIdx][3],
        media_ids: chainRows[cIdx][4],
        threads_post_id: chainRows[cIdx][6],
        permalink: chainRows[cIdx][7]
      });
    }
  }
  // Urutan rantai dipastikan berurutan dari Utas 1, Utas 2, dst.
  postChains.sort(function(a, b) { return a.sequence_order - b.sequence_order; });

  var executedChainUpdates = [];
  var lastThreadsPostId = '';

  try {
    for (var k = 0; k < postChains.length; k++) {
      var chain = postChains[k];
      var chainNumber = k + 1;

      // Jika utas sudah terbit sebelumnya, simpan ID-nya untuk balasan utas berikutnya
      if (chain.threads_post_id) {
        lastThreadsPostId = chain.threads_post_id;
        continue;
      }

      // VALIDASI MUTLAK: Utas #2 DILARANG terbit jika Utas #1 belum sukses!
      if (k > 0 && !lastThreadsPostId) {
        throw new Error('Utas #' + chainNumber + ' dibatalkan karena Utas induk (#1) belum terbit di Threads.');
      }

      // Jika mode target views dan Utas #1 sudah selesai, tunda Utas sambungan
      if (k > 0 && post.publishing_mode === 'target_views' && !isWaitingViewsCheck) {
        for (var u1 = 0; u1 < executedChainUpdates.length; u1++) {
          var upd1 = executedChainUpdates[u1];
          chainsSheet.getRange(upd1.rowIndex, 7).setValue(upd1.publishedPostId);
          chainsSheet.getRange(upd1.rowIndex, 8).setValue(upd1.permalink);
          chainsSheet.getRange(upd1.rowIndex, 9).setValue(upd1.publishedAt);
        }
        currentSheet.getRange(pIdx, 7).setValue('Waiting Views');
        currentSheet.getRange(pIdx, 11).setValue('');

        var currentCheckUntil = currentSheet.getRange(pIdx, 9).getValue();
        if (!currentCheckUntil || String(currentCheckUntil).trim() === '') {
          var defaultCheckUntil = new Date(new Date().getTime() + 48 * 60 * 60 * 1000).toISOString();
          currentSheet.getRange(pIdx, 9).setValue(defaultCheckUntil);
        }

        return { ok: true, status: 'Waiting Views', message: 'Utas #1 berhasil terbit! Utas #2 menunggu target views tercapai.' };
      }

      var mediaList = [];
      try { mediaList = JSON.parse(chain.media_ids || '[]'); } catch (e) { mediaList = []; }

      var publishedPostId = '';

      // Alur Publikasi Meta Threads Container
      if (mediaList.length > 1) {
        var childContainerIds = [];
        for (var m = 0; m < mediaList.length; m++) {
          var mObj = findMediaById(mediaList[m], mediaRows);
          if (mObj) {
            var cPayload = {
              media_type: mObj.media_type === 'VIDEO' ? 'VIDEO' : 'IMAGE',
              is_carousel_item: 'true'
            };
            if (mObj.media_type === 'VIDEO') cPayload.video_url = mObj.cloudinary_url;
            else cPayload.image_url = mObj.cloudinary_url;

            // Titik A: Buat kontainer anak dan tunggu FINISHED
            var childContainerId = createMetaContainer(account.meta_user_id, account.access_token, cPayload);
            checkMetaContainerStatus(childContainerId, account.access_token, 15);
            childContainerIds.push(childContainerId);
          }
        }

        if (childContainerIds.length < 2) {
          throw new Error('Karusel Utas #' + chainNumber + ' membutuhkan minimal 2 media yang valid.');
        }

        // Titik B: Buat kontainer karusel induk
        var pCarouselPayload = {
          media_type: 'CAROUSEL',
          children: childContainerIds.join(',')
        };
        if (chain.content_text && String(chain.content_text).trim().length > 0) {
          pCarouselPayload.text = String(chain.content_text).trim();
        }
        if (lastThreadsPostId) {
          pCarouselPayload.reply_to_id = lastThreadsPostId;
        }

        var parentCarouselId = createMetaContainer(account.meta_user_id, account.access_token, pCarouselPayload);
        publishedPostId = publishMetaContainer(account.meta_user_id, parentCarouselId, account.access_token);

      } else if (mediaList.length === 1) {
        var singleMedia = findMediaById(mediaList[0], mediaRows);
        var sPayload = {
          media_type: (singleMedia && singleMedia.media_type === 'VIDEO') ? 'VIDEO' : 'IMAGE'
        };
        if (singleMedia && singleMedia.media_type === 'VIDEO') sPayload.video_url = singleMedia.cloudinary_url;
        else if (singleMedia) sPayload.image_url = singleMedia.cloudinary_url;

        if (chain.content_text && String(chain.content_text).trim().length > 0) {
          sPayload.text = String(chain.content_text).trim();
        }
        if (lastThreadsPostId) {
          sPayload.reply_to_id = lastThreadsPostId;
        }

        var singleContainerId = createMetaContainer(account.meta_user_id, account.access_token, sPayload);
        publishedPostId = publishMetaContainer(account.meta_user_id, singleContainerId, account.access_token);

      } else {
        var tPayload = {
          media_type: 'TEXT',
          text: String(chain.content_text || '').trim()
        };
        if (lastThreadsPostId) {
          tPayload.reply_to_id = lastThreadsPostId;
        }

        var textContainerId = createMetaContainer(account.meta_user_id, account.access_token, tPayload);
        publishedPostId = publishMetaContainer(account.meta_user_id, textContainerId, account.access_token);
      }

      if (!publishedPostId) {
        throw new Error('Utas #' + chainNumber + ' gagal diterbitkan oleh Meta Threads.');
      }

      var permalink = getOfficialMetaPermalink(publishedPostId, account.access_token, account.username);

      // Simpan hasil ke memori antrean; jangan tulis ke spreadsheet sebelum loop selesai
      executedChainUpdates.push({
        rowIndex: chain.rowIndex,
        publishedPostId: publishedPostId,
        permalink: permalink,
        publishedAt: new Date().toISOString()
      });

      lastThreadsPostId = publishedPostId;
      chain.threads_post_id = publishedPostId;
      chain.permalink = permalink;

      // Jeda waktu propagasi 2.5 detik sebelum memproses balasan Utas #2
      if (k < postChains.length - 1) {
        Utilities.sleep(2500);
      }
    }

    // Tulis data ke Google Sheets HANYA SETELAH seluruh proses Meta berhasil
    for (var w = 0; w < executedChainUpdates.length; w++) {
      var itemUpdate = executedChainUpdates[w];
      chainsSheet.getRange(itemUpdate.rowIndex, 7).setValue(itemUpdate.publishedPostId);
      chainsSheet.getRange(itemUpdate.rowIndex, 8).setValue(itemUpdate.permalink);
      chainsSheet.getRange(itemUpdate.rowIndex, 9).setValue(itemUpdate.publishedAt);
    }

    var allDone = true;
    for (var chk = 0; chk < postChains.length; chk++) {
      if (!postChains[chk].threads_post_id) allDone = false;
    }

    var finalStatus = allDone ? 'Published' : (post.publishing_mode === 'target_views' ? 'Waiting Views' : 'Scheduled');
    currentSheet.getRange(pIdx, 7).setValue(finalStatus);
    currentSheet.getRange(pIdx, 11).setValue('');

    // Migrasi otomatis ke Published_Posts jika telah selesai terbit
    if (finalStatus === 'Published') {
      var finishedRow = currentSheet.getRange(pIdx, 1, 1, currentSheet.getLastColumn()).getValues()[0];
      var pubSheet = ss.getSheetByName('Published_Posts');
      if (!pubSheet) pubSheet = ss.insertSheet('Published_Posts');
      pubSheet.appendRow(finishedRow);
      currentSheet.deleteRow(pIdx);
    }

    return { ok: true, status: finalStatus, message: allDone ? 'Seluruh rantai utas berhasil terbit berurutan ke Threads!' : 'Utas berhasil diproses.' };

  } catch (err) {
    for (var ew = 0; ew < executedChainUpdates.length; ew++) {
      var exItem = executedChainUpdates[ew];
      chainsSheet.getRange(exItem.rowIndex, 7).setValue(exItem.publishedPostId);
      chainsSheet.getRange(exItem.rowIndex, 8).setValue(exItem.permalink);
      chainsSheet.getRange(exItem.rowIndex, 9).setValue(exItem.publishedAt);
    }
    currentSheet.getRange(pIdx, 7).setValue('Failed');
    currentSheet.getRange(pIdx, 11).setValue(err.message);
    return { ok: false, message: 'Gagal mempublikasikan utas: ' + err.message };
  }
}

/**
 * CRITICAL WORKER 1: Queue Publisher Engine (Atomic Locks & Rate Limiting)
 */
function processPublishingQueue() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('processPublishingQueue dilewati: lock tidak tersedia setelah menunggu 30 detik.');
    return;
  }

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var queueSheet = ss.getSheetByName('Queue_Posts') || ss.getSheetByName('Posts');
    if (!queueSheet) return;
    var postRows = queueSheet.getDataRange().getValues();
    if (postRows.length <= 1) return;

    var now = new Date();
    var processedCount = 0;

    for (var pIdx = 1; pIdx < postRows.length; pIdx++) {
      if (processedCount >= 5) break;

      var pId = String(postRows[pIdx][0] || '').trim();
      var status = postRows[pIdx][6];
      var schedAtStr = postRows[pIdx][7];
      var checkUntilStr = postRows[pIdx][8];

      var schedDate = schedAtStr ? new Date(schedAtStr) : null;
      var checkUntilDate = checkUntilStr ? new Date(checkUntilStr) : null;

      var shouldPublish = false;
      var isWaitingViewsCheck = false;

      if (status === 'Scheduled' && schedDate && schedDate.getTime() <= now.getTime()) {
        shouldPublish = true;
      } else if (status === 'Waiting Views') {
        var targetViews = Number(postRows[pIdx][5] || 0);
        var chains = readSheetAsJson('Post_Chains');
        var analytics = readSheetAsJson('Post_Analytics');
        var rootChain = null;
        for (var c = 0; c < chains.length; c++) {
          if (chains[c].post_id === pId && Number(chains[c].sequence_order) === 1) {
            rootChain = chains[c];
            break;
          }
        }
        var curViews = 0;
        if (rootChain) {
          for (var a = 0; a < analytics.length; a++) {
            if (analytics[a].chain_id === rootChain.chain_id) {
              curViews = Number(analytics[a].views_count || 0);
              break;
            }
          }
        }

        // Batas waktu Waiting Views: jika kosong, isi default now + 48 jam
        if (!checkUntilDate) {
          checkUntilDate = new Date(now.getTime() + 48 * 60 * 60 * 1000);
          checkUntilStr = checkUntilDate.toISOString();
          queueSheet.getRange(pIdx + 1, 9).setValue(checkUntilStr);
        }

        var isTimeout = checkUntilDate && now.getTime() >= checkUntilDate.getTime();
        if (curViews >= targetViews || isTimeout) {
          shouldPublish = true;
          isWaitingViewsCheck = true;
        }
      }

      if (shouldPublish) {
        executeSinglePostPublish(pId, isWaitingViewsCheck);
        processedCount++;
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function publishContainer(metaUserId, creationId, token) {
  return publishMetaContainer(metaUserId, creationId, token);
}

function createMediaSnapshot(rawMediaList, mediaRows) {
  if (!Array.isArray(rawMediaList)) return [];
  var snapshots = [];
  for (var i = 0; i < rawMediaList.length; i++) {
    var item = rawMediaList[i];
    if (!item) continue;
    if (typeof item === 'object' && item !== null && (item.url || item.cloudinary_url)) {
      snapshots.push({
        media_id: item.media_id || '',
        url: item.url || item.cloudinary_url,
        type: item.type || item.media_type || 'IMAGE'
      });
      continue;
    }
    if (typeof item === 'string') {
      if (item.indexOf('http://') === 0 || item.indexOf('https://') === 0) {
        var isVid = /\.(mp4|mov|webm)$/i.test(item);
        snapshots.push({
          media_id: '',
          url: item,
          type: isVid ? 'VIDEO' : 'IMAGE'
        });
        continue;
      }
      var found = null;
      if (mediaRows) {
        for (var m = 1; m < mediaRows.length; m++) {
          if (String(mediaRows[m][0]) === String(item)) {
            found = {
              media_id: String(mediaRows[m][0]),
              url: String(mediaRows[m][4] || ''),
              type: String(mediaRows[m][3] || 'IMAGE')
            };
            break;
          }
        }
      }
      if (found) {
        snapshots.push(found);
      } else {
        snapshots.push({ media_id: item, url: '', type: 'IMAGE' });
      }
    }
  }
  return snapshots;
}

function findMediaById(mediaId, mediaRows) {
  if (!mediaId) return null;
  // Jika berupa objek snapshot permanen
  if (typeof mediaId === 'object' && mediaId !== null) {
    var mUrl = mediaId.url || mediaId.cloudinary_url || '';
    var isVid = (mediaId.type && String(mediaId.type).toUpperCase() === 'VIDEO') ||
                (mediaId.media_type && String(mediaId.media_type).toUpperCase() === 'VIDEO') ||
                /\.(mp4|mov|webm)$/i.test(mUrl) ||
                (mUrl.indexOf('/video/') !== -1);
    return {
      media_type: isVid ? 'VIDEO' : 'IMAGE',
      cloudinary_url: mUrl
    };
  }
  // Jika berupa URL langsung (misal dari referensi Pustaka Viral)
  if (typeof mediaId === 'string' && (mediaId.indexOf('http://') === 0 || mediaId.indexOf('https://') === 0)) {
    var isVideo = /\.(mp4|mov|webm)$/i.test(mediaId);
    return {
      media_type: isVideo ? 'VIDEO' : 'IMAGE',
      cloudinary_url: mediaId
    };
  }
  for (var i = 1; i < mediaRows.length; i++) {
    if (String(mediaRows[i][0]) === String(mediaId)) {
      return {
        media_type: mediaRows[i][3],
        cloudinary_url: mediaRows[i][4]
      };
    }
  }
  return null;
}

/**
 * =========================================================================
 * MODUL DIAGNOSTIK & TES KONEKSI API (META GRAPH & CLOUDINARY)
 * =========================================================================
 */

/**
 * Validasi Token Akses Meta Graph API untuk 1 Akun
 */
function checkMetaAccountStatus(accountId) {
  try {
    var accounts = readSheetAsJson('Accounts');
    var targetAcc = null;
    for (var a = 0; a < accounts.length; a++) {
      if (accounts[a].account_id === accountId) {
        targetAcc = accounts[a];
        break;
      }
    }
    if (!targetAcc) return { ok: false, message: 'Akun tidak ditemukan' };
    if (!targetAcc.access_token) return { ok: false, message: 'Access token akun ini kosong' };

    var url = 'https://graph.threads.net/v1.0/me?fields=id,username,name,threads_profile_picture_url&access_token=' + encodeURIComponent(targetAcc.access_token);
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var data = JSON.parse(resp.getContentText());

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Accounts');
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var idCol = headers.indexOf('account_id');
    var activeCol = headers.indexOf('is_active');
    var rowIdx = -1;

    for (var i = 1; i < values.length; i++) {
      if (String(values[i][idCol >= 0 ? idCol : 0]) === String(accountId)) {
        rowIdx = i + 1;
        break;
      }
    }

    if (data.error || resp.getResponseCode() !== 200) {
      var errMsg = data.error ? (data.error.message || JSON.stringify(data.error)) : ('HTTP ' + resp.getResponseCode());
      if (rowIdx > 0 && activeCol >= 0) {
        sheet.getRange(rowIdx, activeCol + 1).setValue('FALSE');
      }
      return {
        ok: false,
        data: {
          account_id: accountId,
          is_active: 'FALSE'
        },
        message: 'Token Meta tidak valid: ' + errMsg
      };
    }

    if (rowIdx > 0 && activeCol >= 0) {
      sheet.getRange(rowIdx, activeCol + 1).setValue('TRUE');
      SpreadsheetApp.flush();
    }

    return {
      ok: true,
      data: {
        account_id: accountId,
        is_active: 'TRUE',
        username: data.username || targetAcc.username,
        name: data.name || targetAcc.name,
        meta_user_id: data.id
      },
      message: 'Akun @' + (data.username || targetAcc.username) + ' aktif dan token valid!'
    };
  } catch (err) {
    return { ok: false, message: 'Pengecekan akun Meta gagal: ' + err.message };
  }
}

/**
 * Validasi Massal Seluruh Armada Akun Meta Threads
 */
function checkAllMetaAccounts() {
  try {
    var accounts = readSheetAsJson('Accounts');
    var results = [];
    var activeCount = 0;
    for (var i = 0; i < accounts.length; i++) {
      var acc = accounts[i];
      var res = checkMetaAccountStatus(acc.account_id);
      results.push({
        account_id: acc.account_id,
        username: acc.username,
        ok: res.ok,
        is_active: res.ok ? 'TRUE' : 'FALSE',
        message: res.message
      });
      if (res.ok) activeCount++;
    }
    return {
      ok: true,
      data: {
        total: accounts.length,
        active: activeCount,
        results: results
      },
      message: activeCount + ' dari ' + accounts.length + ' akun Meta Threads terverifikasi aktif.'
    };
  } catch (err) {
    return { ok: false, message: 'Gagal mengecek semua akun: ' + err.message };
  }
}

/**
 * Cek Kapasitas & Statistik Penyimpanan Cloudinary via Admin API
 */
function getCloudinaryUsageStats() {
  try {
    var settingsRows = readSheetAsJson('Settings');
    var cloudName = '', apiKey = '', apiSecret = '';
    for (var i = 0; i < settingsRows.length; i++) {
      if (settingsRows[i].setting_key === 'cloudinary_cloud_name') cloudName = settingsRows[i].setting_value;
      if (settingsRows[i].setting_key === 'cloudinary_api_key') apiKey = settingsRows[i].setting_value;
      if (settingsRows[i].setting_key === 'cloudinary_api_secret') apiSecret = settingsRows[i].setting_value;
    }

    if (!cloudName || !apiKey || !apiSecret) {
      return { ok: false, message: 'Cloudinary cloud_name, api_key, atau api_secret belum disetel di Settings.' };
    }

    var usageUrl = 'https://api.cloudinary.com/v1_1/' + encodeURIComponent(cloudName) + '/usage';
    var authHeader = 'Basic ' + Utilities.base64Encode(apiKey + ':' + apiSecret);

    var resp = UrlFetchApp.fetch(usageUrl, {
      method: 'get',
      headers: { 'Authorization': authHeader },
      muteHttpExceptions: true
    });

    var data = JSON.parse(resp.getContentText());
    if (resp.getResponseCode() !== 200 || data.error) {
      var errTxt = (data.error && data.error.message) ? data.error.message : ('HTTP ' + resp.getResponseCode());
      return { ok: false, message: 'Cloudinary Admin API Error: ' + errTxt };
    }

    var storageBytes = (data.storage && data.storage.usage) ? Number(data.storage.usage) : 0;
    var storageLimit = (data.storage && data.storage.limit) ? Number(data.storage.limit) : 0;
    var storageUsedMB = (storageBytes / (1024 * 1024)).toFixed(2);
    var storageLimitMB = storageLimit > 0 ? (storageLimit / (1024 * 1024)).toFixed(2) : 'Unlimited';

    var creditsUsage = (data.credits && data.credits.usage) ? data.credits.usage : 0;
    var creditsLimit = (data.credits && data.credits.limit) ? data.credits.limit : 0;

    return {
      ok: true,
      data: {
        plan: data.plan || 'Free',
        storage_used_bytes: storageBytes,
        storage_used_mb: storageUsedMB,
        storage_limit_mb: storageLimitMB,
        storage_pct: (data.storage && data.storage.used_percent) ? data.storage.used_percent : 0,
        credits_used: creditsUsage,
        credits_limit: creditsLimit,
        media_count: (data.objects && data.objects.usage) ? data.objects.usage : 0
      },
      message: 'Statistik penggunaan Cloudinary berhasil dimuat!'
    };
  } catch (err) {
    return { ok: false, message: 'Gagal membaca statistik Cloudinary: ' + err.message };
  }
}

/**
 * Post Insights Synchronizer (Real-Time Meta Graph API)
 */
function syncPostAnalytics() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('Sinkronisasi analitik dibatalkan: Sedang dikunci proses lain.');
    return;
  }

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var chainsSheet = ss.getSheetByName('Post_Chains');
    var analyticsSheet = ss.getSheetByName('Post_Analytics');
    var accounts = readSheetAsJson('Accounts');
    var posts = readSheetAsJson('Posts');

    var chainRows = chainsSheet.getDataRange().getValues();
    if (chainRows.length <= 1) return;

    var allPosts = getAllPostsAsJson();
    var chainHeaders = chainRows[0];
    var cIdCol = chainHeaders.indexOf('chain_id');
    var pIdCol = chainHeaders.indexOf('post_id');
    var thPostIdCol = chainHeaders.indexOf('threads_post_id');

    var anaRows = analyticsSheet.getDataRange().getValues();

    for (var i = 1; i < chainRows.length; i++) {
      var chainId = chainRows[i][cIdCol >= 0 ? cIdCol : 0];
      var postId = chainRows[i][pIdCol >= 0 ? pIdCol : 1];
      var rawThId = chainRows[i][thPostIdCol >= 0 ? thPostIdCol : 6];
      var threadsPostId = rawThId ? String(rawThId).trim() : '';

      if (!threadsPostId) continue;

      var postObj = null;
      for (var p = 0; p < allPosts.length; p++) {
        if (String(allPosts[p].post_id) === String(postId)) {
          postObj = allPosts[p];
          break;
        }
      }
      if (!postObj || !postObj.account_id) continue;

      var token = '';
      for (var a = 0; a < accounts.length; a++) {
        if (String(accounts[a].account_id) === String(postObj.account_id)) {
          token = accounts[a].access_token;
          break;
        }
      }
      if (!token) {
        Logger.log('Token tidak ditemukan untuk Akun ID: ' + postObj.account_id + ' (Chain ' + chainId + ')');
        continue;
      }

      try {
        var insightUrl = 'https://graph.threads.net/v1.0/' + encodeURIComponent(threadsPostId) + '/insights?metric=views,likes,replies,reposts&access_token=' + encodeURIComponent(token);
        var resp = UrlFetchApp.fetch(insightUrl, { muteHttpExceptions: true });
        var data = JSON.parse(resp.getContentText());

        if (data.error) {
          Logger.log('Meta API Insights Error [Chain ' + chainId + ', Post ' + threadsPostId + ']: ' + (data.error.message || JSON.stringify(data.error)));
          continue;
        }

        var views = 0, likes = 0, replies = 0, reposts = 0;
        if (data.data && Array.isArray(data.data)) {
          for (var m = 0; m < data.data.length; m++) {
            var item = data.data[m];
            var rawVal = item.value !== undefined ? item.value : (item.total_value ? item.total_value.value : (item.values && item.values.length > 0 ? item.values[0].value : 0));
            var val = Number(rawVal || 0);
            if (item.name === 'views') views = val;
            if (item.name === 'likes') likes = val;
            if (item.name === 'replies') replies = val;
            if (item.name === 'reposts') reposts = val;
          }
        }

        var anaRowIdx = -1;
        for (var r = 1; r < anaRows.length; r++) {
          if (String(anaRows[r][1]) === String(chainId)) {
            anaRowIdx = r + 1;
            break;
          }
        }

        var nowIso = new Date().toISOString();
        if (anaRowIdx > 0) {
          analyticsSheet.getRange(anaRowIdx, 3, 1, 5).setValues([[views, likes, reposts, replies, nowIso]]);
        } else {
          var anaId = 'ana_' + Utilities.getUuid().slice(0, 8);
          analyticsSheet.appendRow([anaId, chainId, views, likes, reposts, replies, nowIso]);
          anaRows.push([anaId, chainId, views, likes, reposts, replies, nowIso]);
        }

        Logger.log('Meta Insights Synced [Chain ' + chainId + ', Post ' + threadsPostId + ']: views=' + views + ', likes=' + likes + ', reposts=' + reposts + ', replies=' + replies);
      } catch (err) {
        Logger.log('Gagal sinkronisasi analitik [Chain ' + chainId + ', Post ' + threadsPostId + ']: ' + err.message);
      }
    }
  } catch (outerErr) {
    Logger.log('Exception in syncPostAnalytics: ' + outerErr.message);
  } finally {
    lock.releaseLock();
  }
}
