/**
 * ThreadChain Manager - Enterprise Backend Core
 * Single Spreadsheet Connection Architecture
 */

const SPREADSHEET_ID = '1fbG5lrBAAIfYG0np1P7kqSpOclsm9WN1G4E1wBiJ1Uw';

function doGet(e) {
  if (e && e.parameter && (e.parameter.api_key || e.parameter.action)) {
    return handleApiGatewayGet(e);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ThreadChain Manager - Meta Threads Automation')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  return handleApiGatewayPost(e);
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
 * Shared Helper: Sanitasi string jam HH:mm untuk mencegah serial date 1899 di spreadsheet
 */
function cleanTimeString(val, defaultVal) {
  if (!val) return defaultVal || '08:00';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Jakarta', 'HH:mm');
  }
  var s = String(val).trim();
  var m = s.match(/(\d{1,2}):(\d{2})/);
  if (m) {
    return ('0' + m[1]).slice(-2) + ':' + m[2];
  }
  return defaultVal || '08:00';
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
  var dateCols = ['scheduled_at', 'created_at', 'view_check_until', 'published_at', 'token_expires_at', 'last_synced_at', 'last_checked_at', 'original_posted_date'];

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var item = {};
    var hasContent = false;
    for (var j = 0; j < headers.length; j++) {
      var val = row[j];
      var headerName = headers[j];
      if (val instanceof Date) {
        if (headerName === 'generation_time' || headerName === 'window_start' || headerName === 'window_end') {
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
      if ((headerName === 'window_start' || headerName === 'window_end') && val) {
        var mWinTime = String(val).match(/(\d{1,2}):(\d{2})/);
        if (mWinTime) {
          val = ('0' + mWinTime[1]).slice(-2) + ':' + mWinTime[2];
        } else {
          val = (headerName === 'window_start' ? '08:00' : '22:00');
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
      headers: ['account_id', 'meta_user_id', 'username', 'name', 'profile_picture_url', 'access_token', 'token_expires_at', 'is_active', 'category', 'account_persona', 'posting_windows', 'daily_max_posts', 'created_at']
    },
    {
      name: 'Products',
      headers: ['product_id', 'title', 'category', 'price', 'description', 'affiliate_link', 'dashboard_url', 'commission_pct', 'shopee_commission_pct', 'commission_rp', 'cover_url', 'assigned_account_ids', 'created_at']
    },
    {
      name: 'Media_Batches',
      headers: ['batch_id', 'product_id', 'created_at']
    },
    {
      name: 'Media_Files',
      headers: ['media_id', 'sequence_order', 'batch_id', 'product_id', 'media_type', 'cloudinary_url', 'thumbnail_url', 'cloudinary_public_id', 'created_at']
    },
    {
      name: 'Queue_Posts',
      headers: ['post_id', 'chain_group_id', 'sequence_order', 'account_id', 'product_id', 'publishing_mode', 'interval_minutes', 'target_views', 'status', 'scheduled_at', 'view_check_until', 'reply_to_meta_id', 'content_text', 'media_ids', 'threads_post_id', 'permalink', 'published_at', 'is_evergreen', 'error_log', 'created_at']
    },
    {
      name: 'Published_Posts',
      headers: ['post_id', 'chain_group_id', 'sequence_order', 'account_id', 'product_id', 'publishing_mode', 'interval_minutes', 'target_views', 'status', 'scheduled_at', 'view_check_until', 'reply_to_meta_id', 'content_text', 'media_ids', 'threads_post_id', 'permalink', 'published_at', 'is_evergreen', 'error_log', 'created_at']
    },
    {
      name: 'Post_Analytics',
      headers: ['analytics_id', 'threads_post_id', 'views_count', 'likes_count', 'reposts_count', 'replies_count', 'last_synced_at']
    },
    {
      name: 'Post_Comments',
      headers: ['comment_id', 'threads_post_id', 'threads_comment_id', 'from_username', 'text', 'created_at', 'status', 'reply_text']
    },
    {
      name: 'Settings',
      headers: ['setting_key', 'setting_value']
    },
    {
      name: 'Viral_Library',
      headers: ['viral_id', 'title', 'source_url', 'chains_json', 'product_ids', 'extracted_username', 'views_count', 'last_checked_at', 'original_posted_date', 'created_at']
    },
    {
      name: 'Draft_Acc',
      headers: ['draft_id', 'account_id', 'title', 'product_id', 'viral_id', 'chains_json', 'approval_status', 'revision_notes', 'created_at', 'updated_at']
    },
    {
      name: 'API_Keys',
      headers: ['key_id', 'key_name', 'api_key', 'status', 'created_at', 'last_used_at']
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
      ['target_views_timeout_action', 'auto_release'],
      ['triggers_installed', 'FALSE'],
      ['auto_stagger_minutes', '3']
    ];
    settingsSheet.getRange(2, 1, defaults.length, 2).setValues(defaults);
  }

  var deadSheets = ['AI_Prompts', 'Voice_DNA', 'Thread_Formats', 'AI_Training_Memory', 'Autopilot_Rules', 'Post_Chains'];
  for (var ds = 0; ds < deadSheets.length; ds++) {
    var dSheet = ss.getSheetByName(deadSheets[ds]);
    if (dSheet) {
      ss.deleteSheet(dSheet);
    }
  }

  ensureRepliedColumn(ss);
  ensureNormalizedQueueColumns(ss);
  ensureAccountColumns(ss);

  return { ok: true, message: 'Database schema successfully initialized/updated.' };
}

/**
 * Memastikan skema tabel Accounts termutakhirkan dengan posting_windows dan daily_max_posts
 */
function ensureAccountColumns(ss) {
  if (!ss) ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Accounts');
  if (!sheet) return;
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var wsIdx = headers.indexOf('window_start');
  var weIdx = headers.indexOf('window_end');
  var pwIdx = headers.indexOf('posting_windows');
  var dmpIdx = headers.indexOf('daily_max_posts');

  var targetHeaders = ['account_id', 'meta_user_id', 'username', 'name', 'profile_picture_url', 'access_token', 'token_expires_at', 'is_active', 'category', 'account_persona', 'posting_windows', 'daily_max_posts', 'created_at'];

  if (pwIdx === -1 || dmpIdx === -1 || wsIdx !== -1 || weIdx !== -1) {
    var data = sheet.getDataRange().getValues();
    var migrated = [targetHeaders];
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var rowObj = {};
      for (var h = 0; h < headers.length; h++) {
        rowObj[headers[h]] = row[h];
      }
      var defaultWin = [{ start: cleanTimeString(rowObj['window_start'], '08:00'), end: cleanTimeString(rowObj['window_end'], '22:00') }];
      var pw = rowObj['posting_windows'] || JSON.stringify(defaultWin);
      var dmp = Number(rowObj['daily_max_posts'] || 30);
      migrated.push([
        rowObj['account_id'] || '',
        rowObj['meta_user_id'] || '',
        rowObj['username'] || '',
        rowObj['name'] || '',
        rowObj['profile_picture_url'] || '',
        rowObj['access_token'] || '',
        rowObj['token_expires_at'] || '',
        rowObj['is_active'] || 'TRUE',
        rowObj['category'] || 'General',
        rowObj['account_persona'] || '',
        typeof pw === 'string' ? pw : JSON.stringify(pw),
        dmp,
        rowObj['created_at'] || new Date().toISOString()
      ]);
    }
    sheet.clearContents();
    sheet.getRange(1, 1, migrated.length, targetHeaders.length).setValues(migrated);
  }
}

/**
 * Memastikan kolom normalisasi multi-utas tersedia di Queue_Posts dan Published_Posts
 */
function ensureNormalizedQueueColumns(ss) {
  if (!ss) ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var targetSheets = ['Queue_Posts', 'Published_Posts'];
  var reqHeaders = ['post_id', 'chain_group_id', 'sequence_order', 'account_id', 'product_id', 'publishing_mode', 'interval_minutes', 'target_views', 'status', 'scheduled_at', 'view_check_until', 'reply_to_meta_id', 'content_text', 'media_ids', 'threads_post_id', 'permalink', 'published_at', 'is_evergreen', 'error_log', 'created_at'];

  for (var i = 0; i < targetSheets.length; i++) {
    var sheet = ss.getSheetByName(targetSheets[i]);
    if (!sheet) continue;
    var lastCol = sheet.getLastColumn();
    if (lastCol === 0) {
      sheet.getRange(1, 1, 1, reqHeaders.length).setValues([reqHeaders]);
      continue;
    }
    var currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (currentHeaders.indexOf('chain_group_id') === -1 || currentHeaders.indexOf('reply_to_meta_id') === -1) {
      sheet.getRange(1, 1, 1, reqHeaders.length).setValues([reqHeaders]);
    }
  }
}

/**
 * Pemeliharaan skema & normalisasi kolom komentar Post_Comments secara terjadwal/startup
 */
function ensureRepliedColumn(ss) {
  var sheet = ss.getSheetByName('Post_Comments');
  if (sheet) {
    var lastCol = sheet.getLastColumn();
    if (lastCol > 0) {
      var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      var chainIdCol = headers.indexOf('chain_id');
      if (chainIdCol !== -1 && headers.indexOf('threads_post_id') === -1) {
        sheet.getRange(1, chainIdCol + 1).setValue('threads_post_id');
        headers[chainIdCol] = 'threads_post_id';
      }
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
    }
  }

  var aSheet = ss.getSheetByName('Post_Analytics');
  if (aSheet) {
    var aLastCol = aSheet.getLastColumn();
    if (aLastCol > 0) {
      var aHeaders = aSheet.getRange(1, 1, 1, aLastCol).getValues()[0];
      var aChainIdCol = aHeaders.indexOf('chain_id');
      if (aChainIdCol !== -1 && aHeaders.indexOf('threads_post_id') === -1) {
        aSheet.getRange(1, aChainIdCol + 1).setValue('threads_post_id');
      }
    }
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
    if (values.length <= 1) continue;
    var headers = values[0];
    var pIdCol = headers.indexOf('post_id');
    var gIdCol = headers.indexOf('chain_group_id');
    if (pIdCol === -1) pIdCol = 0;

    for (var r = 1; r < values.length; r++) {
      var rowPId = String(values[r][pIdCol]);
      var rowGId = gIdCol >= 0 ? String(values[r][gIdCol]) : '';
      if (rowPId === String(postId) || (rowGId && rowGId === String(postId))) {
        return {
          sheet: sheet,
          sheetName: targetSheets[s],
          rowIndex: r + 1,
          headers: headers,
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
/**
 * Watchdog: Membebaskan pos Processing yang tersangkut > 15 menit dan memindahkannya ke Failed
 */
function resetStalledProcessingPosts(ss) {
  try {
    if (!ss) ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var now = new Date().getTime();
    var timeoutMs = 15 * 60 * 1000;
    var targetSheets = ['Queue_Posts', 'Posts'];
    for (var s = 0; s < targetSheets.length; s++) {
      var sheet = ss.getSheetByName(targetSheets[s]);
      if (!sheet) continue;
      var values = sheet.getDataRange().getValues();
      if (values.length <= 1) continue;
      var headers = values[0];
      var stCol = headers.indexOf('status');
      var errCol = headers.indexOf('error_log');
      var schedCol = headers.indexOf('scheduled_at');
      if (stCol === -1) stCol = 8;
      if (errCol === -1) errCol = 18;
      if (schedCol === -1) schedCol = 9;

      for (var r = values.length - 1; r >= 1; r--) {
        var status = String(values[r][stCol] || '');
        if (status === 'Processing') {
          var dateVal = values[r][schedCol];
          var timeMs = dateVal ? new Date(dateVal).getTime() : 0;
          if (!timeMs || (now - timeMs > timeoutMs)) {
            sheet.getRange(r + 1, stCol + 1).setValue('Failed');
            sheet.getRange(r + 1, errCol + 1).setValue('Watchdog: Status Processing terhenti >15 menit. Otomatis ditandai Failed.');
            if (sheet.getName() === 'Queue_Posts') {
              var failRow = sheet.getRange(r + 1, 1, 1, sheet.getLastColumn()).getValues()[0];
              failRow[stCol] = 'Failed';
              failRow[errCol] = 'Watchdog: Status Processing terhenti >15 menit. Otomatis ditandai Failed.';
              var pSheet = ss.getSheetByName('Published_Posts');
              if (!pSheet) pSheet = ss.insertSheet('Published_Posts');
              pSheet.appendRow(failRow);
              sheet.deleteRow(r + 1);
            }
            Logger.log('Post ' + values[r][0] + ' direset dari Processing ke Failed karena watchdog timeout.');
          }
        }
      }
    }
  } catch (err) {
    Logger.log('resetStalledProcessingPosts error: ' + err.message);
  }
}

function getInitialData() {
  try {
    resetStalledProcessingPosts();
    var settingsRows = readSheetAsJson('Settings');
    var settingsMap = {};
    for (var s = 0; s < settingsRows.length; s++) {
      settingsMap[settingsRows[s].setting_key] = settingsRows[s].setting_value;
    }

    var accounts = readSheetAsJson('Accounts');
    var products = readSheetAsJson('Products');
    var queuePosts = readSheetAsJson('Queue_Posts');
    var publishedPosts = readSheetAsJson('Published_Posts');
    var allThreads = [].concat(queuePosts, publishedPosts);

    accounts.sort(function(a, b) { return new Date(b.created_at || 0) - new Date(a.created_at || 0); });
    products.sort(function(a, b) { return new Date(b.created_at || 0) - new Date(a.created_at || 0); });

    // Grouping baris multi-utas berdasarkan chain_group_id
    var groupsMap = {};
    var groupOrder = [];
    var postChainsList = [];

    for (var i = 0; i < allThreads.length; i++) {
      var t = allThreads[i];
      var gId = String(t.chain_group_id || t.post_id || '').trim();
      if (!gId) continue;

      if (!groupsMap[gId]) {
        groupsMap[gId] = [];
        groupOrder.push(gId);
      }
      groupsMap[gId].push(t);

      var parsedMedia = [];
      try {
        parsedMedia = typeof t.media_ids === 'string' ? JSON.parse(t.media_ids || '[]') : (t.media_ids || []);
      } catch (e) {
        parsedMedia = [];
      }

      postChainsList.push({
        chain_id: t.post_id,
        post_id: gId,
        chain_group_id: gId,
        sequence_order: Number(t.sequence_order || 1),
        content_text: t.content_text || '',
        media_ids: parsedMedia,
        reply_to_meta_id: t.reply_to_meta_id || '',
        threads_post_id: t.threads_post_id || '',
        permalink: t.permalink || '',
        published_at: t.published_at || '',
        status: t.status || (t.threads_post_id ? 'Published' : 'Scheduled')
      });
    }

    // Sintesis data post grup terpadu untuk tampilan tabel dan kartu dashboard
    var posts = [];
    for (var g = 0; g < groupOrder.length; g++) {
      var curGId = groupOrder[g];
      var threads = groupsMap[curGId];
      threads.sort(function(x, y) { return Number(x.sequence_order || 1) - Number(y.sequence_order || 1); });
      var root = threads[0];

      var allPublished = threads.every(function(th) { return th.status === 'Published'; });
      var anyProcessing = threads.some(function(th) { return th.status === 'Processing'; });
      var anyWaiting = threads.some(function(th) { return th.status === 'Waiting Views'; });
      var anyFailed = threads.some(function(th) { return String(th.status || '').indexOf('Failed') !== -1; });
      var anyClosed = threads.some(function(th) { return th.status === 'Closed'; });

      var overallStatus = root.status || 'Scheduled';
      var combinedErrorLog = '';

      for (var thErr = 0; thErr < threads.length; thErr++) {
        if (threads[thErr].error_log) {
          combinedErrorLog = threads[thErr].error_log;
          break;
        }
      }

      if (allPublished) {
        overallStatus = 'Published';
      } else if (anyProcessing) {
        overallStatus = 'Processing';
      } else if (anyWaiting) {
        overallStatus = 'Waiting Views';
      } else if (anyFailed && !anyWaiting) {
        var firstFailed = threads.find(function(th) { return String(th.status || '').indexOf('Failed') !== -1; });
        overallStatus = firstFailed ? firstFailed.status : 'Failed - API Error';
      } else if (anyClosed && !anyWaiting) {
        overallStatus = 'Closed';
      }

      var postObj = {
        post_id: curGId,
        chain_group_id: curGId,
        account_id: root.account_id || '',
        product_id: root.product_id || '',
        publishing_mode: root.publishing_mode || 'interval',
        interval_minutes: root.interval_minutes || 3,
        target_views: root.target_views || '',
        status: overallStatus,
        scheduled_at: root.scheduled_at || '',
        view_check_until: root.view_check_until || (threads[1] ? threads[1].view_check_until : ''),
        root_status: root.status || (root.threads_post_id ? 'Published' : 'Scheduled'),
        is_evergreen: root.is_evergreen || 'FALSE',
        error_log: combinedErrorLog,
        created_at: root.created_at || '',
        chains: threads.map(function(th) {
          var pMed = [];
          try {
            pMed = typeof th.media_ids === 'string' ? JSON.parse(th.media_ids || '[]') : (th.media_ids || []);
          } catch(e) { pMed = []; }
          return {
            chain_id: th.post_id,
            post_id: curGId,
            chain_group_id: curGId,
            sequence_order: Number(th.sequence_order || 1),
            content_text: th.content_text || '',
            media_ids: pMed,
            reply_to_meta_id: th.reply_to_meta_id || '',
            threads_post_id: th.threads_post_id || '',
            permalink: th.permalink || '',
            published_at: th.published_at || '',
            status: th.status || (th.threads_post_id ? 'Published' : 'Scheduled')
          };
        })
      };
      posts.push(postObj);
    }

    posts.sort(function(a, b) {
      var tA = Math.max(new Date(a.created_at || 0).getTime(), new Date(a.scheduled_at || 0).getTime());
      var tB = Math.max(new Date(b.created_at || 0).getTime(), new Date(b.scheduled_at || 0).getTime());
      return tB - tA;
    });

    var viralLibrary = readSheetAsJson('Viral_Library');
    viralLibrary.sort(function(a, b) {
      var tA = Math.max(new Date(a.last_checked_at || 0).getTime(), new Date(a.created_at || 0).getTime());
      var tB = Math.max(new Date(b.last_checked_at || 0).getTime(), new Date(b.created_at || 0).getTime());
      return tB - tA;
    });

    var draftAcc = readSheetAsJson('Draft_Acc');
    draftAcc.sort(function(a, b) {
      return new Date(b.created_at || 0) - new Date(a.created_at || 0);
    });

    var apiKeys = readSheetAsJson('API_Keys');
    apiKeys.sort(function(a, b) {
      return new Date(b.created_at || 0) - new Date(a.created_at || 0);
    });

    var appBaseUrl = settingsMap['web_app_base_url'] || '';
    if (!appBaseUrl) {
      try {
        appBaseUrl = ScriptApp.getService().getUrl() || '';
      } catch (errUrl) {}
    }

    var payload = {
      accounts: accounts,
      products: products,
      mediaBatches: readSheetAsJson('Media_Batches'),
      mediaFiles: readSheetAsJson('Media_Files'),
      posts: posts,
      queuePosts: queuePosts,
      publishedPosts: publishedPosts,
      postChains: postChainsList,
      viralLibrary: viralLibrary,
      postAnalytics: readSheetAsJson('Post_Analytics'),
      postComments: readSheetAsJson('Post_Comments'),
      draftAcc: draftAcc,
      apiKeys: apiKeys,
      baseUrl: appBaseUrl,
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
function addOrSyncAccount(accessToken, category, accountPersona, postingWindows, dailyMaxPosts) {
  try {
    var url = 'https://graph.threads.net/v1.0/me?fields=id,username,name,threads_profile_picture_url&access_token=' + accessToken;
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var data = JSON.parse(resp.getContentText());

    if (data.error) {
      return { ok: false, message: 'Meta Graph API Error: ' + data.error.message };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureAccountColumns(ss);
    var sheet = ss.getSheetByName('Accounts');
    var rows = sheet.getDataRange().getValues();
    var headers = rows[0];

    var nowIso = new Date().toISOString();
    var expiryDate = new Date();
    expiryDate.setDate(expiryDate.getDate() + 60);

    var accountId = 'acc_' + Utilities.getUuid().slice(0, 8);
    var rowIdxToUpdate = -1;
    var defaultCategory = category || 'General';
    var defaultPersona = accountPersona || '';

    var windowsArr = [];
    if (typeof postingWindows === 'string') {
      try { windowsArr = JSON.parse(postingWindows); } catch (e) { windowsArr = []; }
    } else if (Array.isArray(postingWindows)) {
      windowsArr = postingWindows;
    }
    if (!windowsArr || windowsArr.length === 0) {
      windowsArr = [{ start: '08:00', end: '22:00' }];
    }
    var windowsJson = JSON.stringify(windowsArr);
    var maxPosts = Number(dailyMaxPosts) || 30;

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === String(data.id)) {
        rowIdxToUpdate = i + 1;
        accountId = rows[i][0];
        if (!category && rows[i][8]) defaultCategory = rows[i][8];
        if (!accountPersona && rows[i][9]) defaultPersona = rows[i][9];
        var pwCol = headers.indexOf('posting_windows');
        var dmpCol = headers.indexOf('daily_max_posts');
        if (!postingWindows && pwCol >= 0 && rows[i][pwCol]) windowsJson = String(rows[i][pwCol]);
        if (!dailyMaxPosts && dmpCol >= 0 && rows[i][dmpCol]) maxPosts = Number(rows[i][dmpCol]) || 30;
        break;
      }
    }

    var rowObj = {
      account_id: accountId,
      meta_user_id: String(data.id),
      username: data.username || '',
      name: data.name || data.username || '',
      profile_picture_url: data.threads_profile_picture_url || '',
      access_token: accessToken,
      token_expires_at: expiryDate.toISOString(),
      is_active: 'TRUE',
      category: defaultCategory,
      account_persona: defaultPersona,
      posting_windows: windowsJson,
      daily_max_posts: maxPosts,
      created_at: nowIso
    };

    var updateRow = [];
    for (var h = 0; h < headers.length; h++) {
      var key = headers[h];
      updateRow.push(rowObj[key] !== undefined ? rowObj[key] : '');
    }

    if (rowIdxToUpdate > 0) {
      sheet.getRange(rowIdxToUpdate, 1, 1, updateRow.length).setValues([updateRow]);
      return { ok: true, message: 'Akun @' + data.username + ' sudah terdaftar & token berhasil diperbarui!', data: { account_id: accountId, username: data.username } };
    } else {
      sheet.appendRow(updateRow);
      return { ok: true, message: 'Akun @' + data.username + ' berhasil ditambahkan!', data: { account_id: accountId, username: data.username } };
    }
  } catch (err) {
    return { ok: false, message: 'Exception: ' + err.message };
  }
}

function updateAccountDetails(accountData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureAccountColumns(ss);
    var sheet = ss.getSheetByName('Accounts');
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var idCol = headers.indexOf('account_id');
    var catCol = headers.indexOf('category');
    var perCol = headers.indexOf('account_persona');
    var pwCol = headers.indexOf('posting_windows');
    var dmpCol = headers.indexOf('daily_max_posts');

    for (var i = 1; i < values.length; i++) {
      if (String(values[i][idCol >= 0 ? idCol : 0]) === String(accountData.account_id)) {
        var row = i + 1;
        if (accountData.category !== undefined && catCol >= 0) sheet.getRange(row, catCol + 1).setValue(accountData.category);
        if (accountData.account_persona !== undefined && perCol >= 0) sheet.getRange(row, perCol + 1).setValue(accountData.account_persona);
        if (accountData.posting_windows !== undefined && pwCol >= 0) {
          var pwVal = typeof accountData.posting_windows === 'string' ? accountData.posting_windows : JSON.stringify(accountData.posting_windows || []);
          sheet.getRange(row, pwCol + 1).setValue(pwVal);
        }
        if (accountData.daily_max_posts !== undefined && dmpCol >= 0) {
          sheet.getRange(row, dmpCol + 1).setValue(Number(accountData.daily_max_posts) || 30);
        }
        return { ok: true, message: 'Pengaturan akun, jendela posting, dan batas harian berhasil diperbarui' };
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
    var xtraPct = Number(productData.commission_pct || 0);
    var shopeePct = Number(productData.shopee_commission_pct || 0);
    var totalPct = xtraPct + shopeePct;
    var commissionRp = Math.round((Number(productData.price || 0) * totalPct) / 100);
    var nowIso = new Date().toISOString();

    var assignedAccStr = '';
    if (Array.isArray(productData.assigned_account_ids)) {
      assignedAccStr = productData.assigned_account_ids.join(',');
    } else if (productData.assigned_account_ids) {
      assignedAccStr = String(productData.assigned_account_ids).trim();
    }

    var pValues = sheet.getDataRange().getValues();
    var headers = pValues[0];
    if (headers.indexOf('shopee_commission_pct') === -1) {
      sheet.getRange(1, headers.length + 1).setValue('shopee_commission_pct');
      headers.push('shopee_commission_pct');
    }

    if (pId) {
      // Edit existing product
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
          commission_pct: xtraPct,
          shopee_commission_pct: shopeePct,
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
      commission_pct: xtraPct,
      shopee_commission_pct: shopeePct,
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
    if (!sheet) return { ok: false, message: 'Sheet Media_Files tidak ditemukan' };
    var values = sheet.getDataRange().getValues();
    if (values.length <= 1) return { ok: false, message: 'Media tidak ditemukan' };

    var headers = values[0];
    var idCol = headers.indexOf('media_id');
    var seqCol = headers.indexOf('sequence_order');
    var batchCol = headers.indexOf('batch_id');

    if (idCol === -1) idCol = 0;
    if (batchCol === -1) batchCol = 2;

    if (seqCol === -1) {
      sheet.getRange(1, headers.length + 1).setValue('sequence_order');
      headers.push('sequence_order');
      seqCol = headers.length - 1;
      values = sheet.getDataRange().getValues();
    }

    // 1. Cari baris milik mediaId yang dituju. Dapatkan batch_id dan sequence_order saat ini (Posisi A)
    var rowAIndex = -1;
    var targetBatchId = '';
    var posA = 0;

    for (var i = 1; i < values.length; i++) {
      if (String(values[i][idCol]) === String(mediaId)) {
        rowAIndex = i + 1;
        targetBatchId = String(values[i][batchCol]);
        posA = Number(values[i][seqCol]);
        break;
      }
    }

    if (rowAIndex === -1) return { ok: false, message: 'Media tidak ditemukan' };

    // Kumpulkan semua media di dalam batch_id yang sama untuk validasi urutan
    var batchItems = [];
    for (var j = 1; j < values.length; j++) {
      if (String(values[j][batchCol]) === targetBatchId) {
        batchItems.push({
          row: j + 1,
          id: String(values[j][idCol]),
          seq: Number(values[j][seqCol])
        });
      }
    }

    // Normalisasi jika urutan dalam batch belum 1..N (misal karena residu nomor baris sebelumnya)
    batchItems.sort(function(a, b) {
      if (isNaN(a.seq) || a.seq <= 0) return 1;
      if (isNaN(b.seq) || b.seq <= 0) return -1;
      return a.seq - b.seq;
    });

    var needsRepair = false;
    for (var n = 0; n < batchItems.length; n++) {
      if (batchItems[n].seq !== n + 1) {
        needsRepair = true;
        break;
      }
    }

    if (needsRepair) {
      for (var r = 0; r < batchItems.length; r++) {
        batchItems[r].seq = r + 1;
        sheet.getRange(batchItems[r].row, seqCol + 1).setValue(r + 1);
        if (batchItems[r].id === String(mediaId)) {
          posA = r + 1;
        }
      }
    }

    // 2. Hitung target urutan: target_order = Posisi A + direction
    var targetOrder = posA + Number(direction);
    if (targetOrder < 1 || targetOrder > batchItems.length) {
      return { ok: false, message: 'Posisi media sudah berada di ujung' };
    }

    // 3. Cari baris media lain dalam batch_id yang sama dengan sequence_order persis sama dengan target_order (Posisi B)
    var rowBIndex = -1;
    for (var k = 0; k < batchItems.length; k++) {
      if (batchItems[k].seq === targetOrder) {
        rowBIndex = batchItems[k].row;
        break;
      }
    }

    if (rowBIndex === -1) {
      return { ok: false, message: 'Media tetangga untuk pertukaran tidak ditemukan' };
    }

    // 4. Lakukan penimpaan (SWAP): Ubah sequence_order baris A -> target_order, baris B -> Posisi A
    sheet.getRange(rowAIndex, seqCol + 1).setValue(targetOrder);
    sheet.getRange(rowBIndex, seqCol + 1).setValue(posA);

    SpreadsheetApp.flush();
    return { ok: true, message: 'Urutan media berhasil diperbarui' };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

function saveMediaBatch(productId, mediaFiles, targetBatchId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var batchId = targetBatchId || ('batch_' + Utilities.getUuid().slice(0, 8));
    var nowIso = new Date().toISOString();

    var batchSheet = ss.getSheetByName('Media_Batches');
    if (!targetBatchId && batchSheet) {
      var bHeaders = batchSheet.getRange(1, 1, 1, batchSheet.getLastColumn()).getValues()[0];
      var bRow = [];
      for (var bh = 0; bh < bHeaders.length; bh++) {
        var bKey = bHeaders[bh];
        if (bKey === 'batch_id') bRow.push(batchId);
        else if (bKey === 'product_id') bRow.push(productId);
        else if (bKey === 'created_at') bRow.push(nowIso);
        else bRow.push('');
      }
      batchSheet.appendRow(bRow);
    }

    var mediaSheet = ss.getSheetByName('Media_Files');
    var mData = mediaSheet.getDataRange().getValues();
    var mHeaders = mData[0];
    if (mHeaders.indexOf('sequence_order') === -1) {
      mediaSheet.getRange(1, mHeaders.length + 1).setValue('sequence_order');
      mHeaders.push('sequence_order');
      mData = mediaSheet.getDataRange().getValues();
    }

    var bColIdx = mHeaders.indexOf('batch_id');
    if (bColIdx === -1) bColIdx = 2;

    // 1. Hitung jumlah baris media yang SUDAH ADA dengan batch_id yang sama
    var existingInBatchCount = 0;
    for (var r = 1; r < mData.length; r++) {
      if (String(mData[r][bColIdx]) === String(batchId)) {
        existingInBatchCount++;
      }
    }

    // 2. Setiap media baru di batch ini berurutan mulai dari (existingInBatchCount + 1)
    for (var i = 0; i < mediaFiles.length; i++) {
      var m = mediaFiles[i];
      var mediaId = 'med_' + Utilities.getUuid().slice(0, 8);
      var cloudUrl = m.cloudinary_url || m.url || '';
      var thumbUrl = m.thumbnail_url || cloudUrl;

      if ((m.media_type === 'VIDEO' || m.type === 'VIDEO') && cloudUrl) {
        thumbUrl = cloudUrl.replace(/\.[^/.]+$/, '.jpg');
      }

      var pubId = m.cloudinary_public_id || m.public_id || '';
      if (!pubId && cloudUrl) {
        var match = String(cloudUrl).match(/\/upload\/(?:v\d+\/)?([^\.\?#]+)/);
        if (match && match[1]) {
          pubId = decodeURIComponent(match[1]);
        }
      }
      if (!pubId) {
        var dateStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'ddMMyyyy');
        pubId = 'media_' + dateStr + '_' + Utilities.getUuid().slice(0, 4) + '_' + (existingInBatchCount + i + 1);
      }

      var rowObj = {
        media_id: mediaId,
        sequence_order: existingInBatchCount + i + 1,
        batch_id: batchId,
        product_id: productId,
        media_type: m.media_type || m.type || 'IMAGE',
        cloudinary_url: cloudUrl,
        thumbnail_url: thumbUrl,
        cloudinary_public_id: pubId,
        created_at: nowIso
      };

      var newRow = [];
      for (var h = 0; h < mHeaders.length; h++) {
        var key = mHeaders[h];
        newRow.push(rowObj[key] !== undefined ? rowObj[key] : '');
      }
      mediaSheet.appendRow(newRow);
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
    ensureNormalizedQueueColumns(ss);

    var mediaSheet = ss.getSheetByName('Media_Files');
    var mediaRows = mediaSheet ? mediaSheet.getDataRange().getValues() : [];
    var nowIso = new Date().toISOString();

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
    var intervalMin = Number(postData.interval_minutes) || 3;
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

    // KASUS 1: Edit Pos / Grup yang sudah ada
    if (postData.post_id) {
      var targetId = String(postData.post_id).trim();

      // Cek apakah sudah pernah terbit
      var qSheet = ss.getSheetByName('Queue_Posts');
      var pSheet = ss.getSheetByName('Published_Posts');
      var pValues = pSheet ? pSheet.getDataRange().getValues() : [];

      var pHeaders = pValues.length > 0 ? pValues[0] : [];
      var pThIdCol = pHeaders.indexOf('threads_post_id');
      var pStCol = pHeaders.indexOf('status');

      for (var pv = 1; pv < pValues.length; pv++) {
        var rowPId = String(pValues[pv][0] || '');
        var rowGId = String(pValues[pv][1] || '');
        if (rowPId === targetId || rowGId === targetId) {
          var rowThId = pThIdCol >= 0 ? String(pValues[pv][pThIdCol] || '').trim() : '';
          var rowSt = pStCol >= 0 ? String(pValues[pv][pStCol] || '').trim() : '';
          if (rowThId || rowSt === 'Published') {
            return { ok: false, message: 'Postingan yang sudah terbit di Threads telah dikunci dan tidak dapat diedit.' };
          }
        }
      }

      // Hapus baris lama di Queue_Posts yang belum terbit
      if (qSheet) {
        var qValues = qSheet.getDataRange().getValues();
        for (var qv = qValues.length - 1; qv >= 1; qv--) {
          var qPId = String(qValues[qv][0] || '');
          var qGId = String(qValues[qv][1] || '');
          if (qPId === targetId || qGId === targetId) {
            qSheet.deleteRow(qv + 1);
          }
        }
      }

      // Hapus baris arsip gagal di Published_Posts (agar bersih saat dijadwalkan ulang)
      if (pSheet) {
        var curPData = pSheet.getDataRange().getValues();
        for (var pdel = curPData.length - 1; pdel >= 1; pdel--) {
          var pRowPId = String(curPData[pdel][0] || '');
          var pRowGId = String(curPData[pdel][1] || '');
          if (pRowPId === targetId || pRowGId === targetId) {
            pSheet.deleteRow(pdel + 1);
          }
        }
      }

      var editAccId = String(postData.account_id || (postData.account_ids && postData.account_ids[0]) || '').trim();
      var editGroupId = String(postData.draft_id || postData.chain_group_id || targetId).trim();
      if (editGroupId.indexOf('grp_') !== 0 && editGroupId.indexOf('dacc_') !== 0) {
        editGroupId = 'grp_' + Utilities.getUuid().slice(0, 8);
      }
      var baseSchedTimeMs = postData.scheduled_at ? new Date(toIsoUtcString(postData.scheduled_at)).getTime() : Date.now();
      if (isNaN(baseSchedTimeMs)) baseSchedTimeMs = Date.now();

      for (var c = 0; c < chains.length; c++) {
        var seqOrder = c + 1;
        var threadPostId = 'post_' + Utilities.getUuid().slice(0, 8);
        var snapshotMedia = createMediaSnapshot(chains[c].media_ids || [], mediaRows);
        var mJson = JSON.stringify(snapshotMedia);

        var rowStatus = postData.status || 'Scheduled';
        var rowSchedIso = '';
        var rowViewCheckUntil = '';

        if (postData.status === 'Draft') {
          rowStatus = 'Draft';
          rowSchedIso = toIsoUtcString(postData.scheduled_at) || nowIso;
        } else if (postData.publishing_mode === 'target_views') {
          if (c === 0) {
            rowStatus = 'Scheduled';
            rowSchedIso = new Date(baseSchedTimeMs).toISOString();
          } else {
            rowStatus = 'Waiting Views';
            rowSchedIso = '';
            rowViewCheckUntil = new Date(baseSchedTimeMs + fallbackHours * 3600000).toISOString();
          }
        } else {
          rowStatus = 'Scheduled';
          rowSchedIso = new Date(baseSchedTimeMs + (c * intervalMin * 60 * 1000)).toISOString();
        }

        var normalizedRow = {
          post_id: threadPostId,
          chain_group_id: editGroupId,
          sequence_order: seqOrder,
          account_id: editAccId,
          product_id: postData.product_id || '',
          publishing_mode: postData.publishing_mode,
          interval_minutes: intervalMin,
          target_views: isInterval ? '' : (postData.target_views || 100),
          status: rowStatus,
          scheduled_at: toIsoUtcString(rowSchedIso),
          view_check_until: toIsoUtcString(rowViewCheckUntil),
          reply_to_meta_id: '',
          content_text: chains[c].content_text || '',
          media_ids: mJson,
          threads_post_id: '',
          permalink: '',
          published_at: '',
          is_evergreen: postData.is_evergreen || 'FALSE',
          error_log: '',
          created_at: nowIso
        };
        appendRowData('Queue_Posts', normalizedRow);
      }

      return { ok: true, post_id: editGroupId, chain_group_id: editGroupId };
    }

    // KASUS 2: Submit Pembuatan Baru (Mendukung Multi-Akun dengan Stagger)
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
      if (postData.status === 'Draft') {
        targetAccounts = [''];
      } else {
        return { ok: false, message: 'Pilih setidaknya satu akun Threads yang valid' };
      }
    }

    var staggerMin = 3;
    try {
      var sRows2 = readSheetAsJson('Settings');
      for (var si2 = 0; si2 < sRows2.length; si2++) {
        if (sRows2[si2].setting_key === 'auto_stagger_minutes') {
          staggerMin = Number(sRows2[si2].setting_value) || 3;
          break;
        }
      }
    } catch (e) {}

    var baseSchedMs = postData.scheduled_at ? new Date(toIsoUtcString(postData.scheduled_at)).getTime() : Date.now();
    if (isNaN(baseSchedMs)) baseSchedMs = Date.now();
    var createdGroupIds = [];

    for (var k = 0; k < targetAccounts.length; k++) {
      var currentAccId = targetAccounts[k];
      var newGroupId = String(postData.draft_id || postData.chain_group_id || '').trim();
      if (!newGroupId || (k > 0 && targetAccounts.length > 1)) {
        newGroupId = 'grp_' + Utilities.getUuid().slice(0, 8);
      }
      createdGroupIds.push(newGroupId);

      var accountBaseTimeMs = baseSchedMs + (k * staggerMin * 60 * 1000);

      for (var ci = 0; ci < chains.length; ci++) {
        var seqNum = ci + 1;
        var newPostId = 'post_' + Utilities.getUuid().slice(0, 8);
        var snapshotList = createMediaSnapshot(chains[ci].media_ids || [], mediaRows);
        var mediaIdsJson = JSON.stringify(snapshotList);

        var chainStatus = postData.status || 'Scheduled';
        var chainSchedIso = '';
        var chainViewCheckUntil = '';

        if (postData.status === 'Draft') {
          chainStatus = 'Draft';
          chainSchedIso = toIsoUtcString(postData.scheduled_at) || nowIso;
        } else if (postData.publishing_mode === 'target_views') {
          if (ci === 0) {
            chainStatus = 'Scheduled';
            chainSchedIso = new Date(accountBaseTimeMs).toISOString();
          } else {
            chainStatus = 'Waiting Views';
            chainSchedIso = '';
            chainViewCheckUntil = new Date(accountBaseTimeMs + fallbackHours * 3600000).toISOString();
          }
        } else {
          chainStatus = 'Scheduled';
          chainSchedIso = new Date(accountBaseTimeMs + (ci * intervalMin * 60 * 1000)).toISOString();
        }

        var normalizedRowObj = {
          post_id: newPostId,
          chain_group_id: newGroupId,
          sequence_order: seqNum,
          account_id: currentAccId,
          product_id: postData.product_id || '',
          publishing_mode: postData.publishing_mode,
          interval_minutes: intervalMin,
          target_views: isInterval ? '' : (postData.target_views || 100),
          status: chainStatus,
          scheduled_at: toIsoUtcString(chainSchedIso),
          view_check_until: toIsoUtcString(chainViewCheckUntil),
          reply_to_meta_id: '',
          content_text: chains[ci].content_text || '',
          media_ids: mediaIdsJson,
          threads_post_id: '',
          permalink: '',
          published_at: '',
          is_evergreen: postData.is_evergreen || 'FALSE',
          error_log: '',
          created_at: nowIso
        };
        appendRowData('Queue_Posts', normalizedRowObj);
      }
    }

    return { ok: true, post_id: createdGroupIds[0], chain_group_ids: createdGroupIds };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

function deletePost(postId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var targetSheets = ['Queue_Posts', 'Published_Posts', 'Post_Chains'];
    var targetId = String(postId).trim();

    for (var s = 0; s < targetSheets.length; s++) {
      var sheet = ss.getSheetByName(targetSheets[s]);
      if (!sheet) continue;
      var values = sheet.getDataRange().getValues();
      if (values.length <= 1) continue;
      var headers = values[0];
      var pIdCol = headers.indexOf('post_id');
      var gIdCol = headers.indexOf('chain_group_id');
      if (pIdCol === -1) pIdCol = 0;

      for (var r = values.length - 1; r >= 1; r--) {
        var rowPId = String(values[r][pIdCol] || '');
        var rowGId = gIdCol >= 0 ? String(values[r][gIdCol] || '') : '';
        if (rowPId === targetId || rowGId === targetId) {
          sheet.deleteRow(r + 1);
        }
      }
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

/**
 * Smart Best-Time Posting Algorithm with Per-Account Window & Auto-Staggering
 */
function isCurrentTimeInWindows(windowsJson, nowTz) {
  var tz = nowTz || 'Asia/Jakarta';
  var windows = [];
  try {
    windows = typeof windowsJson === 'string' ? JSON.parse(windowsJson) : (windowsJson || []);
  } catch (e) {
    windows = [];
  }
  if (!Array.isArray(windows) || windows.length === 0) {
    windows = [{ start: '08:00', end: '22:00' }];
  }

  var nowStr = Utilities.formatDate(new Date(), tz, 'HH:mm');
  var nowParts = nowStr.split(':');
  var nowMinutes = Number(nowParts[0]) * 60 + Number(nowParts[1]);

  for (var i = 0; i < windows.length; i++) {
    var w = windows[i];
    if (!w || !w.start || !w.end) continue;
    var sParts = cleanTimeString(w.start, '08:00').split(':');
    var eParts = cleanTimeString(w.end, '22:00').split(':');
    var sMin = Number(sParts[0]) * 60 + Number(sParts[1]);
    var eMin = Number(eParts[0]) * 60 + Number(eParts[1]);

    if (eMin > sMin) {
      if (nowMinutes >= sMin && nowMinutes <= eMin) return true;
    } else if (eMin < sMin) {
      if (nowMinutes >= sMin || nowMinutes <= eMin) return true;
    } else {
      return true;
    }
  }
  return false;
}

function getAccountTodayPublishedRootCount(accountId, pData, pHeaders, tz) {
  var targetTz = tz || 'Asia/Jakarta';
  var todayStr = Utilities.formatDate(new Date(), targetTz, 'yyyy-MM-dd');
  var count = 0;

  var accCol = pHeaders.indexOf('account_id');
  var seqCol = pHeaders.indexOf('sequence_order');
  var pubCol = pHeaders.indexOf('published_at');
  var stCol = pHeaders.indexOf('status');

  for (var r = 1; r < pData.length; r++) {
    var row = pData[r];
    var rowAcc = String(row[accCol] || '').trim();
    var rowSeq = Number(row[seqCol] || 1);
    var rowSt = String(row[stCol] || '').trim();
    var rowPubAt = row[pubCol];

    if (rowAcc === accountId && rowSeq === 1 && (rowSt === 'Published' || !rowSt) && rowPubAt) {
      var d = rowPubAt instanceof Date ? rowPubAt : new Date(rowPubAt);
      if (!isNaN(d.getTime())) {
        var dStr = Utilities.formatDate(d, targetTz, 'yyyy-MM-dd');
        if (dStr === todayStr) {
          count++;
        }
      }
    }
  }
  return count;
}

function suggestBestPostingTime(accountId) {
  try {
    var posts = getAllPostsAsJson();
    var sRows = readSheetAsJson('Settings');
    var settingsMap = {};
    for (var s = 0; s < sRows.length; s++) {
      settingsMap[sRows[s].setting_key] = sRows[s].setting_value;
    }
    var tz = settingsMap['timezone'] || 'Asia/Jakarta';
    var staggerMin = Number(settingsMap['auto_stagger_minutes']) || 3;

    var accounts = readSheetAsJson('Accounts');
    var acc = null;
    for (var a = 0; a < accounts.length; a++) {
      if (String(accounts[a].account_id) === String(accountId)) {
        acc = accounts[a];
        break;
      }
    }

    var windows = [];
    if (acc && acc.posting_windows) {
      try {
        windows = typeof acc.posting_windows === 'string' ? JSON.parse(acc.posting_windows) : acc.posting_windows;
      } catch (e) { windows = []; }
    }
    if (!Array.isArray(windows) || windows.length === 0) {
      windows = [{ start: '08:00', end: '22:00' }];
    }

    var now = new Date();
    var nowFormatted = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm');
    var timeParts = nowFormatted.split(' ')[1].split(':');
    var nowMinutes = Number(timeParts[0]) * 60 + Number(timeParts[1]);

    var matchedWindow = null;
    for (var w = 0; w < windows.length; w++) {
      var sP = cleanTimeString(windows[w].start, '08:00').split(':');
      var eP = cleanTimeString(windows[w].end, '22:00').split(':');
      var sMin = Number(sP[0]) * 60 + Number(sP[1]);
      var eMin = Number(eP[0]) * 60 + Number(eP[1]);

      if (nowMinutes + 15 < sMin) {
        matchedWindow = { startMin: sMin, endMin: eMin, offsetDays: 0, label: windows[w].start + '-' + windows[w].end };
        break;
      } else if (nowMinutes + 15 <= eMin) {
        matchedWindow = { startMin: nowMinutes + 15, endMin: eMin, offsetDays: 0, label: windows[w].start + '-' + windows[w].end };
        break;
      }
    }

    if (!matchedWindow) {
      var firstW = windows[0];
      var fSP = cleanTimeString(firstW.start, '08:00').split(':');
      var fMin = Number(fSP[0]) * 60 + Number(fSP[1]);
      matchedWindow = { startMin: fMin, endMin: fMin + 60, offsetDays: 1, label: firstW.start + '-' + firstW.end };
    }

    var targetMinutes = matchedWindow.startMin + Math.floor(Math.random() * 5) + 1;
    var targetBase = new Date(now.getTime() + matchedWindow.offsetDays * 24 * 60 * 60 * 1000);
    var baseDateStr = Utilities.formatDate(targetBase, tz, 'yyyy/MM/dd');
    var targetHour = Math.floor(targetMinutes / 60);
    var targetMinute = targetMinutes % 60;
    var candidateDate = new Date(baseDateStr + ' ' + ('0' + targetHour).slice(-2) + ':' + ('0' + targetMinute).slice(-2) + ':00');

    if (candidateDate.getTime() <= now.getTime()) {
      candidateDate = new Date(now.getTime() + 15 * 60 * 1000);
    }

    var collisionFound = true;
    var loops = 0;
    while (collisionFound && loops < 30) {
      collisionFound = false;
      loops++;
      for (var p = 0; p < posts.length; p++) {
        var postItem = posts[p];
        if (postItem.status === 'Scheduled' && postItem.scheduled_at) {
          var schedTime = new Date(postItem.scheduled_at).getTime();
          var diffMin = Math.abs(schedTime - candidateDate.getTime()) / (60 * 1000);
          if (diffMin < staggerMin) {
            candidateDate.setMinutes(candidateDate.getMinutes() + staggerMin);
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
        reason: 'Jendela @' + (acc ? acc.username : 'akun') + ' (' + matchedWindow.label + ') + jeda ' + staggerMin + ' mnt'
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
    var sourcePost = null;
    for (var i = 0; i < posts.length; i++) {
      if (String(posts[i].post_id) === String(sourcePostId) || String(posts[i].chain_group_id) === String(sourcePostId)) {
        sourcePost = posts[i];
        break;
      }
    }
    if (!sourcePost) return { ok: false, message: 'Sumber post tidak ditemukan' };

    var allThreads = [].concat(readSheetAsJson('Queue_Posts'), readSheetAsJson('Published_Posts'));
    var sourceChains = allThreads.filter(function(t) {
      return String(t.chain_group_id) === String(sourcePostId) || String(t.post_id) === String(sourcePostId);
    });
    sourceChains.sort(function(a, b) { return Number(a.sequence_order || 1) - Number(b.sequence_order || 1); });

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
 * Dynamic API Key & Web App Base URL Management
 */
function generateApiKey(keyName) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('API_Keys');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('API_Keys');
    }

    var keyId = 'key_' + Utilities.getUuid().slice(0, 8);
    var randomBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Date.now().toString());
    var hexToken = randomBytes.map(function(byte) {
      var v = (byte < 0 ? byte + 256 : byte).toString(16);
      return v.length === 1 ? '0' + v : v;
    }).join('').slice(0, 32);

    var apiKey = 'tc_live_' + hexToken;
    var nowIso = new Date().toISOString();
    var cleanName = (keyName || 'Default AI Key').trim();

    var rowObj = {
      key_id: keyId,
      key_name: cleanName,
      api_key: apiKey,
      status: 'Active',
      created_at: nowIso,
      last_used_at: ''
    };

    appendRowData('API_Keys', rowObj);

    var cache = CacheService.getScriptCache();
    cache.put('apk_' + apiKey, 'Active', 21600);

    return {
      ok: true,
      message: 'API Key berhasil dibuat!',
      data: serializeForClient(rowObj)
    };
  } catch (err) {
    return { ok: false, message: 'Gagal membuat API key: ' + err.message };
  }
}

function revokeApiKey(keyId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('API_Keys');
    if (!sheet) return { ok: false, message: 'Tabel API_Keys tidak ditemukan' };

    var values = sheet.getDataRange().getValues();
    if (values.length <= 1) return { ok: false, message: 'Data kunci kosong' };

    var headers = values[0];
    var idCol = headers.indexOf('key_id');
    var keyCol = headers.indexOf('api_key');
    var stCol = headers.indexOf('status');
    if (idCol === -1) idCol = 0;
    if (keyCol === -1) keyCol = 2;
    if (stCol === -1) stCol = 3;

    for (var i = 1; i < values.length; i++) {
      if (String(values[i][idCol]) === String(keyId)) {
        sheet.getRange(i + 1, stCol + 1).setValue('Revoked');
        var revokedToken = String(values[i][keyCol]).trim();
        if (revokedToken) {
          CacheService.getScriptCache().remove('apk_' + revokedToken);
        }
        return { ok: true, message: 'API Key berhasil dinonaktifkan (Revoked).' };
      }
    }
    return { ok: false, message: 'Kunci tidak ditemukan' };
  } catch (err) {
    return { ok: false, message: 'Gagal mencabut kunci: ' + err.message };
  }
}

function saveBaseUrl(url) {
  return saveSettings({ web_app_base_url: (url || '').trim() });
}

function getBaseUrl() {
  try {
    var settingsRows = readSheetAsJson('Settings');
    for (var i = 0; i < settingsRows.length; i++) {
      if (settingsRows[i].setting_key === 'web_app_base_url' && settingsRows[i].setting_value) {
        return settingsRows[i].setting_value;
      }
    }
    return ScriptApp.getService().getUrl() || '';
  } catch (e) {
    return '';
  }
}

/**
 * Universal JSON Output Response Builder for API Gateway
 */
function createJsonResponse(dataObj, statusCode) {
  var output = ContentService.createTextOutput(JSON.stringify(serializeForClient(dataObj)));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

/**
 * Fast API Key Authentication via CacheService
 */
function authenticateApiKey(apiKey) {
  if (!apiKey) return false;
  var token = String(apiKey).trim();
  if (!token) return false;

  var cache = CacheService.getScriptCache();
  var cached = cache.get('apk_' + token);
  if (cached === 'Active') {
    return true;
  }
  if (cached === 'Revoked') {
    return false;
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('API_Keys');
  if (!sheet) {
    setupDatabase();
    sheet = ss.getSheetByName('API_Keys');
  }

  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return false;

  var headers = values[0];
  var keyCol = headers.indexOf('api_key');
  var stCol = headers.indexOf('status');
  var usedCol = headers.indexOf('last_used_at');
  if (keyCol === -1) keyCol = 2;
  if (stCol === -1) stCol = 3;
  if (usedCol === -1) usedCol = 5;

  for (var i = 1; i < values.length; i++) {
    if (String(values[i][keyCol]).trim() === token) {
      var status = String(values[i][stCol] || '').trim();
      if (status === 'Active') {
        cache.put('apk_' + token, 'Active', 21600);
        try {
          sheet.getRange(i + 1, usedCol + 1).setValue(new Date().toISOString());
        } catch (eTrack) {}
        return true;
      } else {
        cache.put('apk_' + token, 'Revoked', 21600);
        return false;
      }
    }
  }
  return false;
}

/**
 * Universal API Gateway: GET Engine (Read & Analytics)
 */
function handleApiGatewayGet(e) {
  try {
    var params = e ? e.parameter : {};
    var apiKey = params.api_key || '';

    if (!authenticateApiKey(apiKey)) {
      return createJsonResponse({
        ok: false,
        status: 401,
        message: 'Unauthorized: API Key tidak valid atau telah dicabut.'
      });
    }

    var action = (params.action || '').trim().toLowerCase();

    // Custom Action: Alokasi Target Harian Akun (Dynamic Slot Allocation)
    if (action === 'get_daily_targets') {
      var accounts = readSheetAsJson('Accounts');
      var pSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Published_Posts');
      var pValues = pSheet ? pSheet.getDataRange().getValues() : [];
      var pHeaders = pValues.length > 0 ? pValues[0] : [];
      var targets = [];

      for (var a = 0; a < accounts.length; a++) {
        var acc = accounts[a];
        if (acc.is_active !== 'TRUE') continue;

        var dailyMax = Number(acc.daily_max_posts) || 30;
        var publishedToday = getAccountTodayPublishedRootCount(acc.account_id, pValues, pHeaders, 'Asia/Jakarta');
        var remainingQuota = Math.max(0, dailyMax - publishedToday);

        var windows = [];
        try {
          windows = typeof acc.posting_windows === 'string' ? JSON.parse(acc.posting_windows) : (acc.posting_windows || []);
        } catch (eWin) {
          windows = [{ start: '08:00', end: '22:00' }];
        }

        var inCurrentWindow = isCurrentTimeInWindows(acc.posting_windows, 'Asia/Jakarta');

        targets.push({
          account_id: acc.account_id,
          username: acc.username,
          category: acc.category || 'General',
          daily_max_posts: dailyMax,
          published_today: publishedToday,
          remaining_quota: remainingQuota,
          in_active_window: inCurrentWindow,
          posting_windows: windows
        });
      }

      return createJsonResponse({
        ok: true,
        action: 'get_daily_targets',
        count: targets.length,
        data: targets
      });
    }

    // Universal Action: Baca Tabel Database Bebas (read_table)
    if (action === 'read_table') {
      var tableName = params.table_name || '';
      if (!tableName) {
        return createJsonResponse({ ok: false, message: 'Parameter table_name wajib disertakan.' });
      }

      var rows = readSheetAsJson(tableName);
      if (params.filter_col && params.filter_val !== undefined) {
        var fCol = String(params.filter_col).trim();
        var fVal = String(params.filter_val).trim();
        rows = rows.filter(function(r) {
          return String(r[fCol] || '').trim() === fVal;
        });
      }

      return createJsonResponse({
        ok: true,
        action: 'read_table',
        table: tableName,
        count: rows.length,
        data: rows
      });
    }

    return createJsonResponse({
      ok: false,
      message: 'Action GET tidak dikenal. Gunakan action=read_table atau action=get_daily_targets.'
    });

  } catch (err) {
    return createJsonResponse({ ok: false, error: err.message });
  }
}

/**
 * Universal API Gateway: POST Engine (Create, Update, Delete, Reschedule)
 */
function handleApiGatewayPost(e) {
  try {
    var rawContents = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    var payload = {};
    try {
      payload = JSON.parse(rawContents);
    } catch (parseErr) {
      payload = (e && e.parameter) ? e.parameter : {};
    }

    var apiKey = payload.api_key || (e && e.parameter ? e.parameter.api_key : '');
    if (!authenticateApiKey(apiKey)) {
      return createJsonResponse({
        ok: false,
        status: 401,
        message: 'Unauthorized: API Key tidak valid atau telah dicabut.'
      });
    }

    var action = (payload.action || (e && e.parameter ? e.parameter.action : '') || '').trim().toLowerCase();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Universal Action: Read Table in POST
    if (action === 'read_table') {
      var tableName = payload.table_name || (e && e.parameter ? e.parameter.table_name : '') || '';
      if (!tableName) {
        return createJsonResponse({ ok: false, message: 'Parameter table_name wajib disertakan.' });
      }

      var targetSheet = ss.getSheetByName(tableName);
      if (!targetSheet) {
        return createJsonResponse({
          ok: false,
          status: 404,
          message: 'Tabel ' + tableName + ' tidak ditemukan (Sheet not found).'
        });
      }

      var rows = readSheetAsJson(tableName);

      // Optional column filtering
      var filterCol = payload.filter_col || (e && e.parameter ? e.parameter.filter_col : '');
      var filterVal = (payload.filter_val !== undefined) ? payload.filter_val : (e && e.parameter ? e.parameter.filter_val : undefined);
      if (filterCol && filterVal !== undefined) {
        var fColStr = String(filterCol).trim();
        var fValStr = String(filterVal).trim();
        rows = rows.filter(function(r) {
          return String(r[fColStr] || '').trim() === fValStr;
        });
      }

      // Optional limit
      var limit = payload.limit !== undefined ? payload.limit : (e && e.parameter ? e.parameter.limit : null);
      if (limit !== null && limit !== undefined && !isNaN(Number(limit))) {
        var limitNum = Math.max(0, parseInt(limit, 10));
        rows = rows.slice(0, limitNum);
      }

      return createJsonResponse({
        ok: true,
        action: 'read_table',
        table: tableName,
        count: rows.length,
        data: rows
      });
    }

    // Universal CRUD: Create Row
    if (action === 'create_row') {
      var createTable = payload.table_name || '';
      var rowData = payload.data || {};
      if (!createTable) return createJsonResponse({ ok: false, message: 'Parameter table_name wajib diisi.' });

      appendRowData(createTable, rowData);
      return createJsonResponse({
        ok: true,
        message: 'Baris berhasil ditambahkan ke ' + createTable,
        data: rowData
      });
    }

    // Universal CRUD: Update Row
    if (action === 'update_row') {
      var updateTable = payload.table_name || '';
      var idColName = payload.id_column || '';
      var idVal = String(payload.id_value || '').trim();
      var updateFields = payload.update_data || {};

      if (!updateTable || !idColName || !idVal) {
        return createJsonResponse({ ok: false, message: 'table_name, id_column, dan id_value wajib diisi.' });
      }

      var uSheet = ss.getSheetByName(updateTable);
      if (!uSheet) return createJsonResponse({ ok: false, message: 'Tabel ' + updateTable + ' tidak ditemukan.' });

      var uValues = uSheet.getDataRange().getValues();
      if (uValues.length <= 1) return createJsonResponse({ ok: false, message: 'Tabel kosong.' });

      var uHeaders = uValues[0];
      var targetColIdx = uHeaders.indexOf(idColName);
      if (targetColIdx === -1) return createJsonResponse({ ok: false, message: 'Kolom ' + idColName + ' tidak ditemukan.' });

      var matchRow = -1;
      for (var r = 1; r < uValues.length; r++) {
        if (String(uValues[r][targetColIdx]).trim() === idVal) {
          matchRow = r + 1;
          break;
        }
      }

      if (matchRow === -1) {
        return createJsonResponse({ ok: false, message: 'Data baris dengan ' + idColName + '=' + idVal + ' tidak ditemukan.' });
      }

      for (var k in updateFields) {
        var fIdx = uHeaders.indexOf(k);
        if (fIdx >= 0) {
          var valToWrite = updateFields[k];
          if (typeof valToWrite === 'object' && valToWrite !== null) {
            valToWrite = JSON.stringify(valToWrite);
          }
          uSheet.getRange(matchRow, fIdx + 1).setValue(valToWrite);
        }
      }

      SpreadsheetApp.flush();
      return createJsonResponse({
        ok: true,
        message: 'Baris ' + idColName + '=' + idVal + ' pada ' + updateTable + ' berhasil diperbarui.'
      });
    }

    // Universal CRUD: Delete Row
    if (action === 'delete_row') {
      var delTable = payload.table_name || '';
      var delIdCol = payload.id_column || '';
      var delIdVal = String(payload.id_value || '').trim();

      if (!delTable || !delIdCol || !delIdVal) {
        return createJsonResponse({ ok: false, message: 'table_name, id_column, dan id_value wajib diisi.' });
      }

      var dSheet = ss.getSheetByName(delTable);
      if (!dSheet) return createJsonResponse({ ok: false, message: 'Tabel ' + delTable + ' tidak ditemukan.' });

      var dValues = dSheet.getDataRange().getValues();
      if (dValues.length <= 1) return createJsonResponse({ ok: false, message: 'Tabel kosong.' });

      var dHeaders = dValues[0];
      var dColIdx = dHeaders.indexOf(delIdCol);
      if (dColIdx === -1) return createJsonResponse({ ok: false, message: 'Kolom ' + delIdCol + ' tidak ditemukan.' });

      var deletedCount = 0;
      for (var dr = dValues.length - 1; dr >= 1; dr--) {
        if (String(dValues[dr][dColIdx]).trim() === delIdVal) {
          dSheet.deleteRow(dr + 1);
          deletedCount++;
        }
      }

      SpreadsheetApp.flush();
      return createJsonResponse({
        ok: true,
        message: 'Berhasil menghapus ' + deletedCount + ' baris dari ' + delTable
      });
    }

    // Custom Action: Otomatis Menjadwalkan Ulang Seluruh Antrean Gagal (reschedule_failed)
    if (action === 'reschedule_failed') {
      var pSheetFailed = ss.getSheetByName('Published_Posts');
      var qSheetQueue = ss.getSheetByName('Queue_Posts');
      if (!pSheetFailed || !qSheetQueue) {
        return createJsonResponse({ ok: false, message: 'Tabel antrean tidak lengkap.' });
      }

      var pfValues = pSheetFailed.getDataRange().getValues();
      if (pfValues.length <= 1) {
        return createJsonResponse({ ok: true, message: 'Tidak ada postingan gagal.', count: 0 });
      }

      var pfHeaders = pfValues[0];
      var stColIdx = pfHeaders.indexOf('status');
      var errColIdx = pfHeaders.indexOf('error_log');
      var schedColIdx = pfHeaders.indexOf('scheduled_at');
      var thPostIdColIdx = pfHeaders.indexOf('threads_post_id');

      var rescheduledCount = 0;
      var nowMs = Date.now();

      for (var pr = pfValues.length - 1; pr >= 1; pr--) {
        var rowStatus = String(pfValues[pr][stColIdx] || '');
        var threadsPostId = String(pfValues[pr][thPostIdColIdx] || '').trim();

        if (rowStatus.indexOf('Failed') !== -1 && !threadsPostId) {
          var rowDataCopy = [].concat(pfValues[pr]);
          rowDataCopy[stColIdx] = 'Scheduled';
          rowDataCopy[errColIdx] = '';
          rowDataCopy[schedColIdx] = new Date(nowMs + (rescheduledCount + 1) * 3 * 60 * 1000).toISOString();

          qSheetQueue.appendRow(rowDataCopy);
          pSheetFailed.deleteRow(pr + 1);
          rescheduledCount++;
        }
      }

      SpreadsheetApp.flush();
      return createJsonResponse({
        ok: true,
        action: 'reschedule_failed',
        message: 'Berhasil menjadwalkan ulang ' + rescheduledCount + ' postingan gagal ke antrean aktif.',
        count: rescheduledCount
      });
    }

    return createJsonResponse({
      ok: false,
      message: 'Action POST tidak dikenal. Pilihan: read_table, create_row, update_row, delete_row, reschedule_failed.'
    });

  } catch (err) {
    return createJsonResponse({ ok: false, error: err.message });
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

    var cleanSourceUrl = (viralData.source_url || '').trim();
    var extractedUsername = (viralData.extracted_username || '').trim();
    if (!extractedUsername && cleanSourceUrl) {
      var uMatch = cleanSourceUrl.match(/@([a-zA-Z0-9._]+)/);
      if (uMatch && uMatch[1]) {
        extractedUsername = uMatch[1];
      }
    }

    var viewsCount = Number(viralData.views_count || 0);
    var origDate = viralData.original_posted_date ? String(viralData.original_posted_date).trim() : '';
    var lastCheckedAt = viralData.last_checked_at || nowIso;

    var values = sheet.getDataRange().getValues();
    var headers = values[0];

    // Ensure new columns exist in sheet row 1 if missing
    var requiredCols = ['extracted_username', 'views_count', 'last_checked_at', 'original_posted_date'];
    var addedCol = false;
    for (var rc = 0; rc < requiredCols.length; rc++) {
      if (headers.indexOf(requiredCols[rc]) === -1) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue(requiredCols[rc]);
        addedCol = true;
      }
    }
    if (addedCol) {
      values = sheet.getDataRange().getValues();
      headers = values[0];
    }

    var colMap = {
      viral_id: headers.indexOf('viral_id'),
      title: headers.indexOf('title'),
      source_url: headers.indexOf('source_url'),
      chains_json: headers.indexOf('chains_json'),
      product_ids: headers.indexOf('product_ids') !== -1 ? headers.indexOf('product_ids') : headers.indexOf('product_id'),
      extracted_username: headers.indexOf('extracted_username'),
      views_count: headers.indexOf('views_count'),
      last_checked_at: headers.indexOf('last_checked_at'),
      original_posted_date: headers.indexOf('original_posted_date'),
      created_at: headers.indexOf('created_at')
    };

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
        if (colMap.source_url >= 0) sheet.getRange(rowIdx, colMap.source_url + 1).setValue(cleanSourceUrl);
        if (colMap.chains_json >= 0) sheet.getRange(rowIdx, colMap.chains_json + 1).setValue(chainsJson);
        if (colMap.product_ids >= 0) sheet.getRange(rowIdx, colMap.product_ids + 1).setValue(targetProductIds);
        if (colMap.extracted_username >= 0) sheet.getRange(rowIdx, colMap.extracted_username + 1).setValue(extractedUsername);
        if (colMap.views_count >= 0) sheet.getRange(rowIdx, colMap.views_count + 1).setValue(viewsCount);
        if (colMap.last_checked_at >= 0) sheet.getRange(rowIdx, colMap.last_checked_at + 1).setValue(lastCheckedAt);
        if (colMap.original_posted_date >= 0) sheet.getRange(rowIdx, colMap.original_posted_date + 1).setValue(origDate);
        return { ok: true, message: 'Konten viral berhasil diperbarui' };
      }
    }

    var newId = 'vir_' + Utilities.getUuid().slice(0, 8);
    var rowObj = {
      viral_id: newId,
      title: autoTitle,
      source_url: cleanSourceUrl,
      chains_json: chainsJson,
      product_ids: targetProductIds,
      extracted_username: extractedUsername,
      views_count: viewsCount,
      last_checked_at: lastCheckedAt,
      original_posted_date: origDate,
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

function updateViralMetrics(viralId, viewsCount, originalPostedDate) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Viral_Library');
    if (!sheet) return { ok: false, message: 'Sheet Viral_Library tidak ditemukan' };
    var values = sheet.getDataRange().getValues();
    if (values.length <= 1) return { ok: false, message: 'Data tidak ditemukan' };
    var headers = values[0];

    var requiredCols = ['extracted_username', 'views_count', 'last_checked_at', 'original_posted_date'];
    var addedCol = false;
    for (var rc = 0; rc < requiredCols.length; rc++) {
      if (headers.indexOf(requiredCols[rc]) === -1) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue(requiredCols[rc]);
        addedCol = true;
      }
    }
    if (addedCol) {
      values = sheet.getDataRange().getValues();
      headers = values[0];
    }

    var idCol = headers.indexOf('viral_id');
    var viewsCol = headers.indexOf('views_count');
    var checkCol = headers.indexOf('last_checked_at');
    var dateCol = headers.indexOf('original_posted_date');
    if (idCol === -1) idCol = 0;

    var nowIso = new Date().toISOString();
    var cleanViews = Number(viewsCount || 0);
    var cleanDate = originalPostedDate ? String(originalPostedDate).trim() : '';

    for (var i = 1; i < values.length; i++) {
      if (String(values[i][idCol]) === String(viralId)) {
        var row = i + 1;
        if (viewsCol >= 0) sheet.getRange(row, viewsCol + 1).setValue(cleanViews);
        if (checkCol >= 0) sheet.getRange(row, checkCol + 1).setValue(nowIso);
        if (dateCol >= 0 && cleanDate) sheet.getRange(row, dateCol + 1).setValue(cleanDate);
        return {
          ok: true,
          message: 'Metrik views dan waktu pengecekan berhasil diperbarui!',
          data: {
            viral_id: viralId,
            views_count: cleanViews,
            last_checked_at: nowIso,
            original_posted_date: cleanDate
          }
        };
      }
    }
    return { ok: false, message: 'Konten viral tidak ditemukan' };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

/**
 * CRUD & APPROVAL ENGINE: Draft Mode ACC
 */
function saveDraftAcc(draftData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Draft_Acc');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('Draft_Acc');
    }

    var nowIso = new Date().toISOString();
    var draftId = draftData.draft_id ? String(draftData.draft_id).trim() : '';
    var accountId = draftData.account_id ? String(draftData.account_id).trim() : '';
    var title = draftData.title ? String(draftData.title).trim() : 'Draf Konten Baru';
    var productId = draftData.product_id ? String(draftData.product_id).trim() : '';
    var viralId = draftData.viral_id ? String(draftData.viral_id).trim() : '';
    var approvalStatus = draftData.approval_status || 'Pending';
    var revisionNotes = draftData.revision_notes ? String(draftData.revision_notes).trim() : '';
    var chainsJson = JSON.stringify(draftData.chains || []);

    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    if (headers.indexOf('account_id') === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue('account_id');
      values = sheet.getDataRange().getValues();
      headers = values[0];
    }
    var idCol = headers.indexOf('draft_id');
    if (idCol === -1) idCol = 0;

    if (draftId) {
      for (var i = 1; i < values.length; i++) {
        if (String(values[i][idCol]) === draftId) {
          var rowIdx = i + 1;
          var aCol = headers.indexOf('account_id') + 1;
          var tCol = headers.indexOf('title') + 1;
          var pCol = headers.indexOf('product_id') + 1;
          var vCol = headers.indexOf('viral_id') + 1;
          var cCol = headers.indexOf('chains_json') + 1;
          var sCol = headers.indexOf('approval_status') + 1;
          var rCol = headers.indexOf('revision_notes') + 1;
          var uCol = headers.indexOf('updated_at') + 1;

          if (aCol > 0) sheet.getRange(rowIdx, aCol).setValue(accountId);
          if (tCol > 0) sheet.getRange(rowIdx, tCol).setValue(title);
          if (pCol > 0) sheet.getRange(rowIdx, pCol).setValue(productId);
          if (vCol > 0) sheet.getRange(rowIdx, vCol).setValue(viralId);
          if (cCol > 0) sheet.getRange(rowIdx, cCol).setValue(chainsJson);
          if (sCol > 0) sheet.getRange(rowIdx, sCol).setValue(approvalStatus);
          if (rCol > 0) sheet.getRange(rowIdx, rCol).setValue(revisionNotes);
          if (uCol > 0) sheet.getRange(rowIdx, uCol).setValue(nowIso);

          return { ok: true, message: 'Draf ACC berhasil diperbarui.', draft_id: draftId };
        }
      }
    }

    var newDraftId = 'dacc_' + Utilities.getUuid().slice(0, 8);
    var newRowObj = {
      draft_id: newDraftId,
      account_id: accountId,
      title: title,
      product_id: productId,
      viral_id: viralId,
      chains_json: chainsJson,
      approval_status: approvalStatus,
      revision_notes: revisionNotes,
      created_at: nowIso,
      updated_at: nowIso
    };

    appendRowData('Draft_Acc', newRowObj);
    return { ok: true, message: 'Draf ACC baru berhasil diajukan untuk peninjauan.', draft_id: newDraftId };
  } catch (err) {
    return { ok: false, message: 'Gagal menyimpan draf: ' + err.message };
  }
}

function updateDraftApprovalStatus(draftId, status, revisionNotes) {
  try {
    if (!draftId) return { ok: false, message: 'ID Draf tidak valid.' };
    var validStatuses = ['Pending', 'Approved', 'Rejected', 'Revision'];
    if (validStatuses.indexOf(status) === -1) {
      return { ok: false, message: 'Status persetujuan tidak valid.' };
    }

    var cleanNotes = revisionNotes ? String(revisionNotes).trim() : '';
    if (status === 'Revision' && !cleanNotes) {
      return { ok: false, message: 'Catatan revisi wajib diisi saat meminta revisi konten!' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Draft_Acc');
    if (!sheet) return { ok: false, message: 'Tabel Draft_Acc tidak ditemukan.' };

    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var idCol = headers.indexOf('draft_id');
    var sCol = headers.indexOf('approval_status') + 1;
    var rCol = headers.indexOf('revision_notes') + 1;
    var uCol = headers.indexOf('updated_at') + 1;
    if (idCol === -1) idCol = 0;

    var nowIso = new Date().toISOString();

    for (var i = 1; i < values.length; i++) {
      if (String(values[i][idCol]) === String(draftId)) {
        var row = i + 1;
        if (sCol > 0) sheet.getRange(row, sCol).setValue(status);
        if (rCol > 0) sheet.getRange(row, rCol).setValue(cleanNotes);
        if (uCol > 0) sheet.getRange(row, uCol).setValue(nowIso);

        var msg = 'Status draf diubah menjadi ' + status + '.';
        if (status === 'Approved') msg = 'Draf berhasil disetujui (Approved)! Tim dapat segera menjadwalkannya.';
        if (status === 'Revision') msg = 'Permintaan revisi berhasil dikirim ke tim pembuat konten.';
        if (status === 'Rejected') msg = 'Draf konten telah ditolak (Rejected).';

        return { ok: true, message: msg, data: { draft_id: draftId, approval_status: status, revision_notes: cleanNotes } };
      }
    }
    return { ok: false, message: 'Draf tidak ditemukan di database.' };
  } catch (err) {
    return { ok: false, message: 'Gagal memperbarui status draf: ' + err.message };
  }
}

function deleteDraftAcc(draftId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Draft_Acc');
    if (!sheet) return { ok: false, message: 'Sheet Draft_Acc tidak ditemukan' };
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var idCol = headers.indexOf('draft_id');
    if (idCol === -1) idCol = 0;

    for (var i = 1; i < values.length; i++) {
      if (String(values[i][idCol]) === String(draftId)) {
        sheet.deleteRow(i + 1);
        return { ok: true, message: 'Draf ACC berhasil dihapus.' };
      }
    }
    return { ok: false, message: 'Draf tidak ditemukan.' };
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
  var acquired = false;
  try {
    acquired = lock.tryLock(20000);
  } catch (lockErr) {
    acquired = false;
  }
  if (!acquired) {
    return { ok: false, message: 'Server sedang sibuk memproses antrean lain. Silakan coba lagi dalam beberapa detik.' };
  }
  try {
    resetStalledProcessingPosts();
    var allPosts = getAllPostsAsJson();
    var target = allPosts.find(function(p) { return String(p.post_id) === String(postId) || String(p.chain_group_id) === String(postId); });
    if (target && String(target.status || '').indexOf('Failed') !== -1) {
      return retryFailedPost(postId, true);
    }
    return publishChainGroupNow(postId, false);
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      Logger.log('Lock release warning (publishPostNow): ' + e.message);
    }
  }
}

function forceReleasePost(postId) {
  var lock = LockService.getScriptLock();
  var acquired = false;
  try {
    acquired = lock.tryLock(20000);
  } catch (lockErr) {
    acquired = false;
  }
  if (!acquired) {
    return { ok: false, message: 'Server sedang sibuk memproses antrean lain. Silakan coba lagi dalam beberapa detik.' };
  }
  try {
    resetStalledProcessingPosts();
    var allPosts = getAllPostsAsJson();
    var target = allPosts.find(function(p) { return String(p.post_id) === String(postId) || String(p.chain_group_id) === String(postId); });
    if (target && String(target.status || '').indexOf('Failed') !== -1) {
      return retryFailedPost(postId, true);
    }
    return publishChainGroupNow(postId, true);
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      Logger.log('Lock release warning (forceReleasePost): ' + e.message);
    }
  }
}

/**
 * Mekanisme Retry / Kloning Arsip Gagal Kembali ke Queue_Posts
 */
function retryFailedPost(postId, publishNow) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var pSheet = ss.getSheetByName('Published_Posts');
    var qSheet = ss.getSheetByName('Queue_Posts');
    if (!qSheet) qSheet = ss.insertSheet('Queue_Posts');

    var targetId = String(postId).trim();
    var allThreads = [].concat(readSheetAsJson('Published_Posts'), readSheetAsJson('Queue_Posts'));
    var groupThreads = allThreads.filter(function(t) {
      return String(t.chain_group_id) === targetId || String(t.post_id) === targetId;
    });

    if (groupThreads.length === 0) {
      return { ok: false, message: 'Postingan tidak ditemukan' };
    }

    var isAnyPublished = groupThreads.some(function(t) { return !!t.threads_post_id; });
    if (isAnyPublished) {
      return { ok: false, message: 'Sebagian atau seluruh postingan sudah terbit di Threads dan tidak dapat di-retry.' };
    }

    groupThreads.sort(function(a, b) { return Number(a.sequence_order || 1) - Number(b.sequence_order || 1); });
    var root = groupThreads[0];

    // Hapus baris lama yang gagal dari Published_Posts dan Queue_Posts
    if (pSheet) {
      var pVals = pSheet.getDataRange().getValues();
      for (var p = pVals.length - 1; p >= 1; p--) {
        if (String(pVals[p][0]) === targetId || String(pVals[p][1]) === targetId) {
          pSheet.deleteRow(p + 1);
        }
      }
    }
    if (qSheet) {
      var qVals = qSheet.getDataRange().getValues();
      for (var q = qVals.length - 1; q >= 1; q--) {
        if (String(qVals[q][0]) === targetId || String(qVals[q][1]) === targetId) {
          qSheet.deleteRow(q + 1);
        }
      }
    }

    // Masukkan KEMBALI ke Queue_Posts sebagai baris aktif baru
    var nowIso = new Date().toISOString();
    var baseSchedMs = Date.now();
    var intervalMin = Number(root.interval_minutes) || 3;
    var isInterval = root.publishing_mode === 'interval';

    for (var c = 0; c < groupThreads.length; c++) {
      var th = groupThreads[c];
      var seq = c + 1;
      var newStatus = 'Scheduled';
      var newSchedIso = new Date(baseSchedMs + (c * intervalMin * 60 * 1000)).toISOString();
      var viewCheckUntil = '';

      if (root.publishing_mode === 'target_views') {
        if (c === 0) {
          newStatus = 'Scheduled';
          newSchedIso = nowIso;
        } else {
          newStatus = 'Waiting Views';
          newSchedIso = '';
          viewCheckUntil = new Date(baseSchedMs + 24 * 3600000).toISOString();
        }
      }

      var activeRow = {
        post_id: th.post_id || ('post_' + Utilities.getUuid().slice(0, 8)),
        chain_group_id: root.chain_group_id || targetId,
        sequence_order: seq,
        account_id: root.account_id,
        product_id: root.product_id || '',
        publishing_mode: root.publishing_mode,
        interval_minutes: intervalMin,
        target_views: isInterval ? '' : (root.target_views || 100),
        status: newStatus,
        scheduled_at: toIsoUtcString(newSchedIso),
        view_check_until: toIsoUtcString(viewCheckUntil),
        reply_to_meta_id: '',
        content_text: th.content_text || '',
        media_ids: typeof th.media_ids === 'string' ? th.media_ids : JSON.stringify(th.media_ids || []),
        threads_post_id: '',
        permalink: '',
        published_at: '',
        is_evergreen: root.is_evergreen || 'FALSE',
        error_log: '',
        created_at: nowIso
      };
      appendRowData('Queue_Posts', activeRow);
    }

    if (publishNow) {
      return publishChainGroupNow(targetId, false);
    }

    return { ok: true, message: 'Postingan berhasil dimasukkan kembali ke antrean aktif!' };
  } catch (err) {
    return { ok: false, message: 'Gagal retry: ' + err.message };
  }
}

function saveAndPublishNow(postPayload, chains) {
  var saveRes = savePostWithChains(postPayload, chains);
  if (!saveRes.ok) return saveRes;
  var groupIds = saveRes.chain_group_ids || [saveRes.post_id];
  var lastResult = null;
  for (var i = 0; i < groupIds.length; i++) {
    lastResult = publishChainGroupNow(groupIds[i], false);
    if (i < groupIds.length - 1) {
      Utilities.sleep(2000);
    }
  }
  return lastResult || saveRes;
}

function syncAllInsightsNow() {
  syncPostAnalytics();
  return { ok: true, message: 'Metrik insights & komentar terbaru pada pos aktif 48 jam berhasil disinkronkan!' };
}

function syncSinglePostInsights(postId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var allThreads = [].concat(readSheetAsJson('Queue_Posts'), readSheetAsJson('Published_Posts'));
    var targetThreads = allThreads.filter(function(t) {
      return String(t.chain_group_id) === String(postId) || String(t.post_id) === String(postId) || String(t.threads_post_id) === String(postId);
    });
    if (targetThreads.length === 0) return { ok: false, message: 'Postingan tidak ditemukan' };

    targetThreads.sort(function(x, y) { return Number(x.sequence_order || 1) - Number(y.sequence_order || 1); });
    var rootThread = targetThreads[0];

    var accounts = readSheetAsJson('Accounts');
    var account = accounts.find(function(a) { return String(a.account_id) === String(rootThread.account_id); });
    if (!account || !account.access_token) {
      return { ok: false, message: 'Token akses akun Threads tidak ditemukan' };
    }

    var analyticsSheet = ss.getSheetByName('Post_Analytics');
    if (!analyticsSheet) {
      analyticsSheet = ss.insertSheet('Post_Analytics');
      analyticsSheet.appendRow(['analytics_id', 'threads_post_id', 'views_count', 'likes_count', 'reposts_count', 'replies_count', 'last_synced_at']);
    }
    var anaRows = analyticsSheet.getDataRange().getValues();
    var anaHeaders = anaRows[0] || ['analytics_id', 'threads_post_id', 'views_count', 'likes_count', 'reposts_count', 'replies_count', 'last_synced_at'];
    var thColIdx = anaHeaders.indexOf('threads_post_id');
    if (thColIdx === -1) thColIdx = 1;

    var syncedCount = 0;
    for (var i = 0; i < targetThreads.length; i++) {
      var status = String(targetThreads[i].status || '').trim();
      var thId = String(targetThreads[i].threads_post_id || '').trim();
      if (status !== 'Published' || !thId) continue;

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
            if (String(anaRows[r][thColIdx]) === thId) {
              rowIdx = r + 1;
              break;
            }
          }
          var nowIso = new Date().toISOString();
          if (rowIdx > 0) {
            analyticsSheet.getRange(rowIdx, 3, 1, 5).setValues([[views, likes, reposts, replies, nowIso]]);
          } else {
            var newAnaId = 'ana_' + Utilities.getUuid().slice(0, 8);
            analyticsSheet.appendRow([newAnaId, thId, views, likes, reposts, replies, nowIso]);
            anaRows.push([newAnaId, thId, views, likes, reposts, replies, nowIso]);
          }
          syncedCount++;
        }
      } catch (callErr) {}
    }
    SpreadsheetApp.flush();
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
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var allThreads = [].concat(readSheetAsJson('Queue_Posts'), readSheetAsJson('Published_Posts'));
    var allSystemThreadIds = {};
    for (var at = 0; at < allThreads.length; at++) {
      var tpid = String(allThreads[at].threads_post_id || '').trim();
      if (tpid) allSystemThreadIds[tpid] = true;
    }

    var targetThreads = allThreads.filter(function(t) {
      return String(t.chain_group_id) === String(postId) || String(t.post_id) === String(postId) || String(t.threads_post_id) === String(postId);
    });
    if (targetThreads.length === 0) return { ok: false, message: 'Postingan tidak ditemukan' };

    targetThreads.sort(function(x, y) { return Number(x.sequence_order || 1) - Number(y.sequence_order || 1); });
    var rootThread = targetThreads[0];

    var accounts = readSheetAsJson('Accounts');
    var account = accounts.find(function(a) { return String(a.account_id) === String(rootThread.account_id); });
    if (!account || !account.access_token) {
      return { ok: false, message: 'Token akses akun Threads tidak ditemukan' };
    }

    var authorUsername = String(account.username || '').toLowerCase().trim();
    var commSheet = ss.getSheetByName('Post_Comments');
    if (!commSheet) {
      commSheet = ss.insertSheet('Post_Comments');
      commSheet.appendRow(['comment_id', 'threads_post_id', 'threads_comment_id', 'from_username', 'text', 'created_at', 'status', 'reply_text']);
    }
    var existingRows = commSheet.getDataRange().getValues();
    var existingHeaders = existingRows[0] || [];
    var thCommIdColIdx = existingHeaders.indexOf('threads_comment_id');
    if (thCommIdColIdx === -1) thCommIdColIdx = 2;

    var rowsToDelete = [];
    var existingIds = {};
    for (var r = 1; r < existingRows.length; r++) {
      var existThId = String(existingRows[r][thCommIdColIdx] || '').trim();
      if (existThId) {
        if (allSystemThreadIds[existThId]) {
          rowsToDelete.push(r + 1);
        } else {
          existingIds[existThId] = true;
        }
      }
    }
    if (rowsToDelete.length > 0) {
      for (var d = rowsToDelete.length - 1; d >= 0; d--) {
        commSheet.deleteRow(rowsToDelete[d]);
      }
    }

    var newCommentsCount = 0;
    for (var i = 0; i < targetThreads.length; i++) {
      var status = String(targetThreads[i].status || '').trim();
      var thId = String(targetThreads[i].threads_post_id || '').trim();
      if (status !== 'Published' || !thId) continue;

      try {
        var repUrl = 'https://graph.threads.net/v1.0/' + encodeURIComponent(thId) + '/replies?fields=id,text,username,timestamp&access_token=' + encodeURIComponent(account.access_token);
        var repResp = UrlFetchApp.fetch(repUrl, { muteHttpExceptions: true });
        var repData = JSON.parse(repResp.getContentText());
        if (repData && repData.data && Array.isArray(repData.data)) {
          for (var m = 0; m < repData.data.length; m++) {
            var rep = repData.data[m];
            var commentThId = String(rep.id || '').trim();
            if (allSystemThreadIds[commentThId]) continue;
            if (commentThId && !existingIds[commentThId]) {
              var fromUser = rep.username || 'user';
              var isAuthor = String(fromUser).toLowerCase().trim() === authorUsername;
              var textContent = rep.text || '';
              var createdAt = rep.timestamp || new Date().toISOString();
              var newCommId = 'com_' + Utilities.getUuid().slice(0, 8);

              commSheet.appendRow([
                newCommId,
                thId,
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
    var thPostIdCol = headers.indexOf('threads_post_id');
    if (thPostIdCol === -1) thPostIdCol = headers.indexOf('chain_id');
    var thCommIdCol = headers.indexOf('threads_comment_id');
    var statusCol = headers.indexOf('status');
    var replyTextCol = headers.indexOf('reply_text');

    var targetRowIdx = -1;
    var threadsPostId = '';
    var threadsCommentId = '';

    for (var i = 1; i < commRows.length; i++) {
      if (String(commRows[i][cIdCol >= 0 ? cIdCol : 0]) === String(commentId)) {
        targetRowIdx = i + 1;
        threadsPostId = String(commRows[i][thPostIdCol >= 0 ? thPostIdCol : 1]);
        threadsCommentId = String(commRows[i][thCommIdCol >= 0 ? thCommIdCol : 2]);
        break;
      }
    }

    if (targetRowIdx === -1) {
      return { ok: false, message: 'Komentar tidak ditemukan' };
    }

    var allThreads = [].concat(readSheetAsJson('Published_Posts'), readSheetAsJson('Queue_Posts'));
    var matchedThread = allThreads.find(function(th) {
      return String(th.threads_post_id) === threadsPostId || String(th.post_id) === threadsPostId;
    });
    if (!matchedThread) return { ok: false, message: 'Utas induk tidak ditemukan' };

    var accounts = readSheetAsJson('Accounts');
    var matchedAcc = accounts.find(function(a) { return String(a.account_id) === String(matchedThread.account_id); });
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
    var resp = UrlFetchApp.fetch(checkUrl, { muteHttpExceptions: true });
    var data = JSON.parse(resp.getContentText());

    if (data.error) {
      throw new Error('Meta Container ' + containerId + ' Error: ' + (data.error.message || JSON.stringify(data.error)));
    }

    var st = String(data.status || '').toUpperCase();
    if (st === 'FINISHED') {
      return true;
    }
    if (st === 'ERROR' || st === 'EXPIRED') {
      throw new Error('Meta Container ' + containerId + ' Gagal (' + st + '): ' + (data.error_message || 'Media gagal diproses server Meta'));
    }
    if (!data.status) {
      return true;
    }

    Utilities.sleep(2000);
  }

  throw new Error('Meta Container ' + containerId + ' Timeout: Media belum selesai diproses setelah ' + retries + ' kali pengecekan.');
}

/**
 * MODUL 2: Format Pengiriman Data ke Meta (access_token di URL & Body)
 */
function createMetaContainer(metaUserId, token, payloadObj) {
  var url = 'https://graph.threads.net/v1.0/me/threads?access_token=' + encodeURIComponent(token);
  payloadObj.access_token = token;

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

function publishMetaContainer(metaUserId, creationId, token, isMedia) {
  if (isMedia) {
    checkMetaContainerStatus(creationId, token, 15);
  }

  var pUrl = 'https://graph.threads.net/v1.0/me/threads_publish?access_token=' + encodeURIComponent(token);
  var payload = {
    creation_id: String(creationId).trim(),
    access_token: token
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
 * MODUL 3: Mesin Eksekusi Rantai Utas Berurutan (Database Normalization & Reply Chaining)
 */
function publishChainGroupNow(targetIdentifier, forceAllChains) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureNormalizedQueueColumns(ss);

  var accounts = readSheetAsJson('Accounts');
  var mediaSheet = ss.getSheetByName('Media_Files');
  var mediaRows = mediaSheet ? mediaSheet.getDataRange().getValues() : [];
  var targetId = String(targetIdentifier).trim();

  // Temukan chain_group_id target
  var qSheet = ss.getSheetByName('Queue_Posts');
  var pSheet = ss.getSheetByName('Published_Posts');
  if (!pSheet) pSheet = ss.insertSheet('Published_Posts');

  var qData = qSheet ? qSheet.getDataRange().getValues() : [];
  var qHeaders = qData.length > 0 ? qData[0] : [];
  var pData = pSheet.getDataRange().getValues();
  var pHeaders = pData.length > 0 ? pData[0] : [];

  var qPIdCol = qHeaders.indexOf('post_id');
  var qGIdCol = qHeaders.indexOf('chain_group_id');
  var pPIdCol = pHeaders.indexOf('post_id');
  var pGIdCol = pHeaders.indexOf('chain_group_id');

  var matchedGroupId = '';
  for (var q = 1; q < qData.length; q++) {
    if (String(qData[q][qPIdCol]) === targetId || String(qData[q][qGIdCol]) === targetId) {
      matchedGroupId = String(qData[q][qGIdCol] || qData[q][qPIdCol]);
      break;
    }
  }
  if (!matchedGroupId) {
    for (var p = 1; p < pData.length; p++) {
      if (String(pData[p][pPIdCol]) === targetId || String(pData[p][pGIdCol]) === targetId) {
        matchedGroupId = String(pData[p][pGIdCol] || pData[p][pPIdCol]);
        break;
      }
    }
  }
  if (!matchedGroupId) matchedGroupId = targetId;

  // Tarik seluruh utas di Queue dan Published untuk group ini
  var queueThreads = readSheetAsJson('Queue_Posts').filter(function(r) {
    return String(r.chain_group_id) === matchedGroupId || String(r.post_id) === matchedGroupId;
  });
  var pubThreads = readSheetAsJson('Published_Posts').filter(function(r) {
    return String(r.chain_group_id) === matchedGroupId || String(r.post_id) === matchedGroupId;
  });

  var allGroupThreads = [].concat(pubThreads, queueThreads);
  allGroupThreads.sort(function(a, b) { return Number(a.sequence_order || 1) - Number(b.sequence_order || 1); });

  if (allGroupThreads.length === 0) {
    return { ok: false, message: 'Data utas dalam grup tidak ditemukan' };
  }

  var rootThread = allGroupThreads[0];
  var account = accounts.find(function(a) { return String(a.account_id) === String(rootThread.account_id); });
  if (!account || !account.access_token) {
    return { ok: false, message: 'Akun Threads atau Access Token tidak valid' };
  }

  var isTargetViews = (rootThread.publishing_mode === 'target_views');
  var shouldHoldChain2 = (isTargetViews && !forceAllChains);

  try {
    for (var k = 0; k < allGroupThreads.length; k++) {
      var curThread = allGroupThreads[k];
      var seq = Number(curThread.sequence_order || (k + 1));

      // Jika utas sudah memiliki threads_post_id Meta, lewati
      if (curThread.threads_post_id) {
        continue;
      }

      // Pastikan reply_to_meta_id tersedia untuk Utas #2 dan seterusnya
      var parentMetaId = curThread.reply_to_meta_id ? String(curThread.reply_to_meta_id).trim() : '';
      if (seq > 1 && !parentMetaId) {
        var prevSeqThread = allGroupThreads.find(function(th) { return Number(th.sequence_order) === seq - 1; });
        if (prevSeqThread && prevSeqThread.threads_post_id) {
          parentMetaId = String(prevSeqThread.threads_post_id).trim();
          curThread.reply_to_meta_id = parentMetaId;
        } else {
          throw new Error('Utas #' + seq + ' gagal rilis karena Utas #' + (seq - 1) + ' belum memiliki Meta Threads ID.');
        }
      }

      // Tahan Utas #2+ jika mode target views dan bukan eksekusi paksa
      if (seq > 1 && shouldHoldChain2) {
        // Pastikan reply_to_meta_id tersimpan di Queue_Posts
        updateThreadFieldInQueue(ss, curThread.post_id, 'reply_to_meta_id', parentMetaId);
        updateThreadFieldInQueue(ss, curThread.post_id, 'status', 'Waiting Views');
        return { ok: true, status: 'Waiting Views', message: 'Utas #1 berhasil terbit! Utas #2 kini berstatus Waiting Views dan menunggu target views.' };
      }

      // Eksekusi pengiriman utas ke Meta Threads
      var pubResult = executeSingleMetaThreadPublish(curThread, account, parentMetaId, mediaRows);
      if (!pubResult.ok) {
        throw new Error('Utas #' + seq + ' gagal: ' + pubResult.message);
      }

      curThread.threads_post_id = pubResult.threads_post_id;
      curThread.permalink = pubResult.permalink;
      curThread.published_at = pubResult.published_at;
      curThread.status = 'Published';

      // Pindahkan baris utas ini dari Queue_Posts ke Published_Posts
      moveThreadToPublished(ss, curThread.post_id, pubResult.threads_post_id, pubResult.permalink, pubResult.published_at, parentMetaId);

      // SUNTIKKAN threads_post_id ke reply_to_meta_id utas berikutnya di Queue_Posts
      var nextSeqThread = allGroupThreads.find(function(th) { return Number(th.sequence_order) === seq + 1; });
      if (nextSeqThread) {
        nextSeqThread.reply_to_meta_id = pubResult.threads_post_id;
        updateThreadFieldInQueue(ss, nextSeqThread.post_id, 'reply_to_meta_id', pubResult.threads_post_id);
      }

      // Jeda waktu propagasi 5 detik antar utas agar Meta cluster mengindeks
      if (k < allGroupThreads.length - 1 && (!isTargetViews || forceAllChains)) {
        Utilities.sleep(5000);
      }
    }

    return { ok: true, message: 'Seluruh rantai utas berhasil diterbitkan berurutan ke Meta Threads!' };

  } catch (err) {
    var errMsg = err.message || String(err);
    if (curThread && curThread.post_id) {
      archiveSingleThreadFromQueue(ss, curThread.post_id, 'Failed - API Error', errMsg);
    }
    // Chain Abort: gagalkan seluruh utas sambungan berikutnya dan arsipkan ke Published_Posts
    if (allGroupThreads && curThread) {
      var failedSeqNum = Number(curThread.sequence_order || 1);
      for (var abIdx = 0; abIdx < allGroupThreads.length; abIdx++) {
        var targetAbTh = allGroupThreads[abIdx];
        if (Number(targetAbTh.sequence_order) > failedSeqNum && !targetAbTh.threads_post_id) {
          archiveSingleThreadFromQueue(ss, targetAbTh.post_id, 'Failed - Parent Error', 'Dibatalkan: Utas #' + failedSeqNum + ' gagal terbit (' + errMsg + ')');
        }
      }
    }
    return { ok: false, message: errMsg };
  }
}

/**
 * Helper: Memindahkan 1 baris gagal dari Queue_Posts ke Published_Posts
 */
function archiveSingleThreadFromQueue(ss, postId, status, errorMsg) {
  try {
    var qSheet = ss.getSheetByName('Queue_Posts');
    var pSheet = ss.getSheetByName('Published_Posts');
    if (!qSheet || !pSheet) return;

    var data = qSheet.getDataRange().getValues();
    if (data.length <= 1) return;
    var headers = data[0];
    var pIdCol = headers.indexOf('post_id');
    var stCol = headers.indexOf('status');
    var errCol = headers.indexOf('error_log');
    if (pIdCol === -1) pIdCol = 0;

    for (var r = 1; r < data.length; r++) {
      if (String(data[r][pIdCol]) === String(postId)) {
        var row = [].concat(data[r]);
        if (stCol >= 0) row[stCol] = status;
        if (errCol >= 0) row[errCol] = errorMsg;
        pSheet.appendRow(row);
        qSheet.deleteRow(r + 1);
        SpreadsheetApp.flush();
        return;
      }
    }
  } catch (e) {
    Logger.log('archiveSingleThreadFromQueue error: ' + e.message);
  }
}

/**
 * Helper: Eksekusi API Meta Threads untuk 1 Utas (Single Thread Container & Publish)
 */
function executeSingleMetaThreadPublish(threadObj, account, parentMetaId, mediaRows) {
  try {
    // 1. Ekstrak hanya nilai ID media atau URL langsung dari data mentah
    var rawMedia = threadObj.media_ids;
    if (typeof rawMedia === 'string' && rawMedia.trim()) {
      try {
        rawMedia = JSON.parse(rawMedia);
      } catch (e) {
        rawMedia = rawMedia.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
      }
    }
    if (!Array.isArray(rawMedia)) {
      rawMedia = rawMedia ? [rawMedia] : [];
    }

    var extractedIdsOrUrls = [];
    for (var mi = 0; mi < rawMedia.length; mi++) {
      var item = rawMedia[mi];
      if (!item) continue;
      if (typeof item === 'object' && item !== null) {
        if (item.media_id && String(item.media_id).trim()) {
          extractedIdsOrUrls.push(String(item.media_id).trim());
        } else if (item.url && /^https?:\/\//i.test(String(item.url).trim())) {
          extractedIdsOrUrls.push(String(item.url).trim());
        } else if (item.cloudinary_url && /^https?:\/\//i.test(String(item.cloudinary_url).trim())) {
          extractedIdsOrUrls.push(String(item.cloudinary_url).trim());
        }
      } else if (typeof item === 'string') {
        var strVal = item.trim();
        if (strVal) extractedIdsOrUrls.push(strVal);
      }
    }

    // 2. Integritas Data Mutlak: Ambil data Media_Files langsung dari database
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var mSheet = ss.getSheetByName('Media_Files');
    var mData = mSheet ? mSheet.getDataRange().getValues() : [];
    var mHeaders = mData.length > 0 ? mData[0] : [];

    var idCol = mHeaders.indexOf('media_id');
    var typeCol = mHeaders.indexOf('media_type');
    var urlCol = mHeaders.indexOf('cloudinary_url');
    if (idCol === -1) idCol = 0;
    if (typeCol === -1) typeCol = 4;
    if (urlCol === -1) urlCol = 5;

    // Build lookup map dari database: media_id -> { cloudinary_url, media_type }
    var dbMediaMap = {};
    for (var r = 1; r < mData.length; r++) {
      var rowId = String(mData[r][idCol] || '').trim();
      var rowUrl = String(mData[r][urlCol] || '').trim();
      var rowType = String(mData[r][typeCol] || 'IMAGE').trim().toUpperCase();
      if (rowId && /^https?:\/\//i.test(rowUrl)) {
        dbMediaMap[rowId] = {
          media_id: rowId,
          cloudinary_url: rowUrl,
          media_type: rowType === 'VIDEO' ? 'VIDEO' : 'IMAGE'
        };
      }
    }

    // 3. Resolusikan media: Hanya gunakan URL valid asli dari database
    var validMediaList = [];
    for (var k = 0; k < extractedIdsOrUrls.length; k++) {
      var cand = extractedIdsOrUrls[k];
      if (dbMediaMap[cand]) {
        validMediaList.push(dbMediaMap[cand]);
      } else if (/^https?:\/\//i.test(cand)) {
        var isVid = /\.(mp4|mov|webm)$/i.test(cand) || cand.indexOf('/video/') !== -1;
        validMediaList.push({
          media_id: '',
          cloudinary_url: cand,
          media_type: isVid ? 'VIDEO' : 'IMAGE'
        });
      }
    }

    var publishedPostId = '';
    var textContent = String(threadObj.content_text || '').trim();
    var cleanParentId = parentMetaId ? String(parentMetaId).trim() : '';

    if (validMediaList.length > 1) {
      // Karusel: reply_to_id HANYA dipasang di carouselPayload (Parent), DILARANG di cPayload (Child)
      var childContainerIds = [];
      for (var m = 0; m < validMediaList.length; m++) {
        var mObj = validMediaList[m];
        var isVid = (mObj.media_type === 'VIDEO');
        var cPayload = {
          media_type: isVid ? 'VIDEO' : 'IMAGE',
          is_carousel_item: 'true'
        };
        if (isVid) cPayload.video_url = mObj.cloudinary_url;
        else cPayload.image_url = mObj.cloudinary_url;

        var childId = createMetaContainer(account.meta_user_id, account.access_token, cPayload);
        checkMetaContainerStatus(childId, account.access_token, 15);
        childContainerIds.push(childId);
      }

      if (childContainerIds.length < 2) {
        throw new Error('Karusel membutuhkan minimal 2 media yang valid.');
      }

      var carouselPayload = {
        media_type: 'CAROUSEL',
        children: childContainerIds.join(',')
      };
      if (textContent) carouselPayload.text = textContent;
      if (cleanParentId) {
        carouselPayload.reply_to_id = cleanParentId;
      }

      var parentContainerId = createMetaContainer(account.meta_user_id, account.access_token, carouselPayload);
      checkMetaContainerStatus(parentContainerId, account.access_token, 15);
      publishedPostId = publishMetaContainer(account.meta_user_id, parentContainerId, account.access_token, false);

    } else if (validMediaList.length === 1) {
      // Single Image / Video: HANYA sPayload.reply_to_id jika ada parentMetaId
      var singleMedia = validMediaList[0];
      var isSingleVid = (singleMedia.media_type === 'VIDEO');
      var sPayload = {
        media_type: isSingleVid ? 'VIDEO' : 'IMAGE'
      };
      if (isSingleVid) sPayload.video_url = singleMedia.cloudinary_url;
      else sPayload.image_url = singleMedia.cloudinary_url;

      if (textContent) sPayload.text = textContent;
      if (cleanParentId) {
        sPayload.reply_to_id = cleanParentId;
      }

      var singleContainerId = createMetaContainer(account.meta_user_id, account.access_token, sPayload);
      publishedPostId = publishMetaContainer(account.meta_user_id, singleContainerId, account.access_token, true);

    } else {
      // Teks Murni: HANYA tPayload.reply_to_id jika ada parentMetaId
      var tPayload = {
        media_type: 'TEXT',
        text: textContent
      };
      if (cleanParentId) {
        tPayload.reply_to_id = cleanParentId;
      }

      var textContainerId = createMetaContainer(account.meta_user_id, account.access_token, tPayload);
      publishedPostId = publishMetaContainer(account.meta_user_id, textContainerId, account.access_token, false);
    }

    if (!publishedPostId) {
      throw new Error('Meta Threads tidak mengembalikan ID pos.');
    }

    var permalink = getOfficialMetaPermalink(publishedPostId, account.access_token, account.username);
    return {
      ok: true,
      threads_post_id: publishedPostId,
      permalink: permalink,
      published_at: new Date().toISOString()
    };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

/**
 * Helper: Update nilai 1 kolom di Queue_Posts
 */
function updateThreadFieldInQueue(ss, postId, fieldName, value) {
  try {
    var qSheet = ss.getSheetByName('Queue_Posts');
    if (!qSheet) return;
    var data = qSheet.getDataRange().getValues();
    if (data.length <= 1) return;
    var headers = data[0];
    var pIdCol = headers.indexOf('post_id');
    var targetCol = headers.indexOf(fieldName);
    if (pIdCol === -1 || targetCol === -1) return;

    for (var r = 1; r < data.length; r++) {
      if (String(data[r][pIdCol]) === String(postId)) {
        qSheet.getRange(r + 1, targetCol + 1).setValue(value);
        break;
      }
    }
    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log('updateThreadFieldInQueue error: ' + e.message);
  }
}

/**
 * Helper: Pindahkan baris utas dari Queue_Posts ke Published_Posts
 */
function moveThreadToPublished(ss, postId, threadsPostId, permalink, publishedAt, replyToMetaId) {
  try {
    var qSheet = ss.getSheetByName('Queue_Posts');
    var pSheet = ss.getSheetByName('Published_Posts');
    if (!pSheet) pSheet = ss.insertSheet('Published_Posts');

    if (qSheet) {
      var data = qSheet.getDataRange().getValues();
      var headers = data[0];
      var pIdCol = headers.indexOf('post_id');
      if (pIdCol === -1) pIdCol = 0;

      for (var r = 1; r < data.length; r++) {
        if (String(data[r][pIdCol]) === String(postId)) {
          var rowData = data[r];
          var stCol = headers.indexOf('status');
          var thCol = headers.indexOf('threads_post_id');
          var lkCol = headers.indexOf('permalink');
          var pubCol = headers.indexOf('published_at');
          var repCol = headers.indexOf('reply_to_meta_id');
          var errCol = headers.indexOf('error_log');

          if (stCol >= 0) rowData[stCol] = 'Published';
          if (thCol >= 0) rowData[thCol] = threadsPostId;
          if (lkCol >= 0) rowData[lkCol] = permalink;
          if (pubCol >= 0) rowData[pubCol] = publishedAt;
          if (repCol >= 0 && replyToMetaId) rowData[repCol] = replyToMetaId;
          if (errCol >= 0) rowData[errCol] = '';

          pSheet.appendRow(rowData);
          qSheet.deleteRow(r + 1);
          SpreadsheetApp.flush();
          return;
        }
      }
    }

    // Jika baris ternyata sudah ada di Published_Posts (misal pembaruan dari arsip)
    if (pSheet) {
      var pData = pSheet.getDataRange().getValues();
      var pHeaders = pData[0];
      var pIdCol2 = pHeaders.indexOf('post_id');
      if (pIdCol2 === -1) pIdCol2 = 0;
      var stCol2 = pHeaders.indexOf('status');
      var thCol2 = pHeaders.indexOf('threads_post_id');
      var lkCol2 = pHeaders.indexOf('permalink');
      var pubCol2 = pHeaders.indexOf('published_at');
      var repCol2 = pHeaders.indexOf('reply_to_meta_id');
      var errCol2 = pHeaders.indexOf('error_log');

      for (var pr = 1; pr < pData.length; pr++) {
        if (String(pData[pr][pIdCol2]) === String(postId)) {
          if (stCol2 >= 0) pSheet.getRange(pr + 1, stCol2 + 1).setValue('Published');
          if (thCol2 >= 0) pSheet.getRange(pr + 1, thCol2 + 1).setValue(threadsPostId);
          if (lkCol2 >= 0) pSheet.getRange(pr + 1, lkCol2 + 1).setValue(permalink);
          if (pubCol2 >= 0) pSheet.getRange(pr + 1, pubCol2 + 1).setValue(publishedAt);
          if (repCol2 >= 0 && replyToMetaId) pSheet.getRange(pr + 1, repCol2 + 1).setValue(replyToMetaId);
          if (errCol2 >= 0) pSheet.getRange(pr + 1, errCol2 + 1).setValue('');
          SpreadsheetApp.flush();
          return;
        }
      }
    }
  } catch (e) {
    Logger.log('moveThreadToPublished error: ' + e.message);
  }
}

/**
 * Sanitasi & Pembersihan Otomatis Baris Terminal di Queue_Posts
 */
function cleanupTerminalQueuePosts(ss) {
  try {
    if (!ss) ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var qSheet = ss.getSheetByName('Queue_Posts');
    if (!qSheet || qSheet.getLastRow() <= 1) return;
    var pSheet = ss.getSheetByName('Published_Posts');
    if (!pSheet) pSheet = ss.insertSheet('Published_Posts');

    var qData = qSheet.getDataRange().getValues();
    if (qData.length <= 1) return;
    var headers = qData[0];
    var stCol = headers.indexOf('status');
    if (stCol === -1) stCol = 8;
    var pIdCol = headers.indexOf('post_id');
    if (pIdCol === -1) pIdCol = 0;

    var pData = pSheet.getDataRange().getValues();
    var existingPIds = {};
    for (var p = 1; p < pData.length; p++) {
      var existPid = String(pData[p][0] || '');
      if (existPid) existingPIds[existPid] = true;
    }

    var rowsToDelete = [];
    for (var r = 1; r < qData.length; r++) {
      var st = String(qData[r][stCol] || '').trim();
      var curPId = String(qData[r][pIdCol] || '');
      if (st === 'Published' || st === 'Closed' || st.indexOf('Failed') === 0) {
        if (!existingPIds[curPId]) {
          pSheet.appendRow(qData[r]);
          existingPIds[curPId] = true;
        }
        rowsToDelete.push(r + 1);
      }
    }
    if (rowsToDelete.length > 0) {
      rowsToDelete.sort(function(a, b) { return b - a; });
      for (var d = 0; d < rowsToDelete.length; d++) {
        qSheet.deleteRow(rowsToDelete[d]);
      }
      SpreadsheetApp.flush();
    }
  } catch (e) {
    Logger.log('cleanupTerminalQueuePosts error: ' + e.message);
  }
}

/**
 * CRITICAL WORKER: Queue Publisher Engine (Cronjob Pengirim Antrean Multi-Utas)
 * Golden Standard: Jendela Toleransi 5 Menit, Looping Bebas Hambatan,
 * Dinamika Reply & Chain Abort Protocol, Safe Reverse Row Deletion.
 */
function processPublishingQueue() {
  var lock = LockService.getScriptLock();
  var acquired = false;
  try {
    acquired = lock.tryLock(25000);
  } catch (e) {
    acquired = false;
  }
  if (!acquired) {
    Logger.log('processPublishingQueue dilewati: lock tidak tersedia setelah 25 detik.');
    return;
  }

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureNormalizedQueueColumns(ss);
    cleanupTerminalQueuePosts(ss);
    resetStalledProcessingPosts(ss);

    var qSheet = ss.getSheetByName('Queue_Posts');
    if (!qSheet || qSheet.getLastRow() <= 1) return;

    var pSheet = ss.getSheetByName('Published_Posts');
    if (!pSheet) pSheet = ss.insertSheet('Published_Posts');

    var qData = qSheet.getDataRange().getValues();
    if (qData.length <= 1) return;
    var headers = qData[0];

    var pIdCol = headers.indexOf('post_id');
    var gIdCol = headers.indexOf('chain_group_id');
    var seqCol = headers.indexOf('sequence_order');
    var accCol = headers.indexOf('account_id');
    var prodCol = headers.indexOf('product_id');
    var modeCol = headers.indexOf('publishing_mode');
    var intCol = headers.indexOf('interval_minutes');
    var tvCol = headers.indexOf('target_views');
    var stCol = headers.indexOf('status');
    var schedCol = headers.indexOf('scheduled_at');
    var checkCol = headers.indexOf('view_check_until');
    var replyCol = headers.indexOf('reply_to_meta_id');
    var textCol = headers.indexOf('content_text');
    var mediaCol = headers.indexOf('media_ids');
    var thIdCol = headers.indexOf('threads_post_id');
    var linkCol = headers.indexOf('permalink');
    var pubAtCol = headers.indexOf('published_at');
    var everCol = headers.indexOf('is_evergreen');
    var errCol = headers.indexOf('error_log');

    var accounts = readSheetAsJson('Accounts');
    var mediaSheet = ss.getSheetByName('Media_Files');
    var mediaRows = mediaSheet ? mediaSheet.getDataRange().getValues() : [];
    var analytics = readSheetAsJson('Post_Analytics');
    var sRows = readSheetAsJson('Settings');

    var fallbackHours = 24;
    var timeoutAction = 'auto_release';
    for (var si = 0; si < sRows.length; si++) {
      if (sRows[si].setting_key === 'default_fallback_views_hours') {
        fallbackHours = Number(sRows[si].setting_value) || 24;
      }
      if (sRows[si].setting_key === 'target_views_timeout_action') {
        timeoutAction = sRows[si].setting_value || 'auto_release';
      }
    }

    var now = new Date();
    var nowMs = now.getTime();

    // Lacak kelompok yang terputus (aborted) dalam siklus eksekusi ini
    var abortedGroupIds = {};
    // Lacak baris sukses yang akan dipindahkan dan dihapus dari Queue_Posts (1-based index)
    var rowsToDelete = [];

    // Looping bebas hambatan melintasi seluruh baris antrean
    for (var r = 1; r < qData.length; r++) {
      var rowPId = String(qData[r][pIdCol] || '').trim();
      var rowGId = String(qData[r][gIdCol] || rowPId).trim();
      var seq = Number(qData[r][seqCol] || 1);
      var status = String(qData[r][stCol] || '').trim();
      var schedAtStr = qData[r][schedCol];
      var checkUntilStr = qData[r][checkCol];
      var pubMode = String(qData[r][modeCol] || '').trim();
      var replyMetaId = String(qData[r][replyCol] || '').trim();
      var accId = String(qData[r][accCol] || '').trim();

      // Jika grup ini sudah diputus rantainya akibat kegagalan utas induk pada siklus ini:
      if (abortedGroupIds[rowGId]) {
        if (status === 'Scheduled' || status === 'Waiting Views') {
          qSheet.getRange(r + 1, stCol + 1).setValue('Failed - Parent Error');
          qSheet.getRange(r + 1, errCol + 1).setValue(abortedGroupIds[rowGId]);
          qData[r][stCol] = 'Failed - Parent Error';
          qData[r][errCol] = abortedGroupIds[rowGId];
        }
        continue;
      }

      // Hanya proses baris berstatus Scheduled atau Waiting Views
      if (status !== 'Scheduled' && status !== 'Waiting Views') {
        continue;
      }

      var account = null;
      for (var ac = 0; ac < accounts.length; ac++) {
        if (String(accounts[ac].account_id) === accId) {
          account = accounts[ac];
          break;
        }
      }

      var shouldPublish = false;

      // 1. EVALUASI JADWAL WAKTU TIBA DI PALING ATAS
      if (status === 'Scheduled') {
        var hasExplicitSchedule = false;
        var schedMs = 0;

        if (schedAtStr) {
          var sDate = (schedAtStr instanceof Date) ? schedAtStr : new Date(schedAtStr);
          if (!isNaN(sDate.getTime())) {
            hasExplicitSchedule = true;
            schedMs = sDate.getTime();
          }
        }

        if (hasExplicitSchedule) {
          // Jika waktu jadwal belum tiba, lewati
          if (nowMs < schedMs) {
            continue;
          }

          // WAKTU JADWAL SUDAH TIBA (nowMs >= schedMs): Prioritas Mutlak!
          // ABAIKAN Batas Kuota Harian & Jendela Waktu Posting.
          // Untuk Utas #2 dan seterusnya, periksa ketersediaan reply_to_meta_id
          if (seq > 1 && !replyMetaId) {
            var pubRows = readSheetAsJson('Published_Posts');
            var prevPub = null;
            var prevFailed = false;
            for (var pbIdx = pubRows.length - 1; pbIdx >= 0; pbIdx--) {
              var pb = pubRows[pbIdx];
              if (String(pb.chain_group_id || pb.post_id) === rowGId && Number(pb.sequence_order) === seq - 1) {
                if (pb.threads_post_id) {
                  prevPub = pb;
                  break;
                } else if (String(pb.status || '').indexOf('Failed') !== -1) {
                  prevFailed = true;
                  break;
                }
              }
            }

            if (prevPub && prevPub.threads_post_id) {
              replyMetaId = String(prevPub.threads_post_id).trim();
              qSheet.getRange(r + 1, replyCol + 1).setValue(replyMetaId);
              qData[r][replyCol] = replyMetaId;
              SpreadsheetApp.flush();
              shouldPublish = true;
            } else if (prevFailed) {
              var parentMissingMsg = 'Dibatalkan otomatis: Utas #' + (seq - 1) + ' gagal terbit di Threads.';
              qData[r][stCol] = 'Failed - Parent Error';
              qData[r][errCol] = parentMissingMsg;
              pSheet.appendRow(qData[r]);
              rowsToDelete.push(r + 1);
              abortedGroupIds[rowGId] = parentMissingMsg;
              archiveSubsequentChainsToPublished(qSheet, pSheet, qData, rowGId, seq, parentMissingMsg, stCol, errCol, gIdCol, seqCol, rowsToDelete);
              continue;
            } else {
              // Utas induk belum selesai rilis -> tunda ke eksekusi cron berikutnya
              continue;
            }
          } else {
            shouldPublish = true;
          }

        } else {
          // 2. THROTTLING HANYA BERLAKU JIKA TANPA JADWAL EKSPLISIT
          if (seq === 1 && account) {
            var dailyMax = Number(account.daily_max_posts) || 30;
            var pValuesLatest = pSheet.getDataRange().getValues();
            var todayPublishedRootCount = getAccountTodayPublishedRootCount(accId, pValuesLatest, pHeaders, 'Asia/Jakarta');
            if (todayPublishedRootCount >= dailyMax) {
              Logger.log('Akun @' + (account.username || accId) + ' mencapai limit harian (' + todayPublishedRootCount + '/' + dailyMax + '). Menunda Utas #1.');
              continue;
            }

            var inWindow = isCurrentTimeInWindows(account.posting_windows, 'Asia/Jakarta');
            if (!inWindow) {
              continue;
            }
          }
          shouldPublish = true;
        }

      // 3. Evaluasi Status 'Waiting Views' (Target Views)
      } else if (status === 'Waiting Views' && pubMode === 'target_views') {
        var pubRows2 = readSheetAsJson('Published_Posts');
        var rootPub = pubRows2.find(function(pb) {
          return String(pb.chain_group_id || pb.post_id) === rowGId && Number(pb.sequence_order) === 1;
        });

        if (!rootPub || !rootPub.threads_post_id) {
          continue; // Menunggu Utas 1 selesai
        }

        // Pastikan reply_to_meta_id menunjuk ke utas tepat di atasnya
        if (!replyMetaId) {
          var prevPub2 = pubRows2.find(function(pb) {
            return String(pb.chain_group_id || pb.post_id) === rowGId && Number(pb.sequence_order) === seq - 1;
          });
          if (prevPub2 && prevPub2.threads_post_id) {
            replyMetaId = String(prevPub2.threads_post_id).trim();
            qSheet.getRange(r + 1, replyCol + 1).setValue(replyMetaId);
            qData[r][replyCol] = replyMetaId;
          } else {
            replyMetaId = String(rootPub.threads_post_id).trim();
            qSheet.getRange(r + 1, replyCol + 1).setValue(replyMetaId);
            qData[r][replyCol] = replyMetaId;
          }
        }

        var targetViews = Number(qData[r][tvCol] || 100);
        var curViews = 0;
        for (var a = 0; a < analytics.length; a++) {
          if (String(analytics[a].threads_post_id) === String(rootPub.threads_post_id) || String(analytics[a].chain_id) === String(rootPub.threads_post_id)) {
            curViews = Number(analytics[a].views_count || 0);
            break;
          }
        }

        var checkUntilDate = checkUntilStr ? new Date(checkUntilStr) : null;
        if (!checkUntilDate || isNaN(checkUntilDate.getTime())) {
          checkUntilDate = new Date(nowMs + fallbackHours * 3600000);
          qSheet.getRange(r + 1, checkCol + 1).setValue(checkUntilDate.toISOString());
          qData[r][checkCol] = checkUntilDate.toISOString();
        }

        var isTimeout = checkUntilDate && nowMs >= checkUntilDate.getTime();
        if (curViews >= targetViews || (isTimeout && timeoutAction === 'auto_release')) {
          shouldPublish = true;
        } else if (isTimeout && timeoutAction !== 'auto_release') {
          var closedMsg = 'Closed: Batas waktu target views berakhir, Utas #' + seq + ' batal rilis.';
          qSheet.getRange(r + 1, stCol + 1).setValue('Closed');
          qSheet.getRange(r + 1, errCol + 1).setValue(closedMsg);
          qData[r][stCol] = 'Closed';
          qData[r][errCol] = closedMsg;

          // Pindahkan baris Closed ke Published_Posts dan hapus dari antrean
          var closedRow = [].concat(qData[r]);
          closedRow[stCol] = 'Closed';
          closedRow[errCol] = closedMsg;
          pSheet.appendRow(closedRow);
          rowsToDelete.push(r + 1);
          continue;
        }
      }

      // 3. Eksekusi Pengiriman ke Meta Threads API
      if (shouldPublish) {
        // Pre-Flight Locking: Tandai status Processing di Sheets & memori lalu flush seketika
        qSheet.getRange(r + 1, stCol + 1).setValue('Processing');
        qData[r][stCol] = 'Processing';
        SpreadsheetApp.flush();

        var account = null;
        for (var ac = 0; ac < accounts.length; ac++) {
          if (String(accounts[ac].account_id) === accId) {
            account = accounts[ac];
            break;
          }
        }

        var threadObj = {
          post_id: rowPId,
          chain_group_id: rowGId,
          sequence_order: seq,
          account_id: accId,
          content_text: qData[r][textCol] || '',
          media_ids: qData[r][mediaCol] || '[]',
          reply_to_meta_id: replyMetaId
        };

        try {
          if (!account || !account.access_token) {
            throw new Error('Akun Threads atau Access Token tidak valid untuk akun ID: ' + accId);
          }

          var pubResult = executeSingleMetaThreadPublish(threadObj, account, replyMetaId, mediaRows);
          if (!pubResult || !pubResult.ok || !pubResult.threads_post_id) {
            throw new Error((pubResult && pubResult.message) ? pubResult.message : 'Meta API tidak mengembalikan ID postingan.');
          }

          var newThreadsPostId = pubResult.threads_post_id;
          var newPermalink = pubResult.permalink;
          var newPubAt = pubResult.published_at || new Date().toISOString();

          // SUNTIKKAN threads_post_id ke reply_to_meta_id utas berikutnya (seq + 1) di Queue_Posts seketika dan flush ke sheet
          for (var n = 1; n < qData.length; n++) {
            if (String(qData[n][gIdCol] || qData[n][pIdCol]) === rowGId && Number(qData[n][seqCol]) === seq + 1) {
              var barisUtasDua = n + 1;
              qSheet.getRange(barisUtasDua, replyCol + 1).setValue(newThreadsPostId);
              qData[n][replyCol] = newThreadsPostId;
              SpreadsheetApp.flush();
              break;
            }
          }

          // Catat data sukses & pindahkan ke Published_Posts
          var publishedRowData = [].concat(qData[r]);
          publishedRowData[stCol] = 'Published';
          publishedRowData[thIdCol] = newThreadsPostId;
          publishedRowData[linkCol] = newPermalink;
          publishedRowData[pubAtCol] = newPubAt;
          publishedRowData[errCol] = '';
          if (replyMetaId) publishedRowData[replyCol] = replyMetaId;

          pSheet.appendRow(publishedRowData);
          rowsToDelete.push(r + 1);

          Logger.log('Sukses terbit Utas #' + seq + ' (' + rowPId + ') -> Meta Post ID: ' + newThreadsPostId);

          // Jeda propagasi server 5 detik agar Meta mengindeks utas sebelum balasan berikutnya
          Utilities.sleep(5000);

        } catch (apiErr) {
          // Tangani Error API Meta: Catat error asli, arsipkan baris gagal ke Published_Posts
          var errMsg = apiErr.message || String(apiErr);
          qData[r][stCol] = 'Failed - API Error';
          qData[r][errCol] = errMsg;
          pSheet.appendRow(qData[r]);
          rowsToDelete.push(r + 1);

          Logger.log('Gagal terbit Utas #' + seq + ' (' + rowPId + '): ' + errMsg);

          // Chain Abort: gagalkan seluruh utas anak berikutnya dalam grup yang sama dan arsipkan
          var childAbortMsg = 'Dibatalkan otomatis: Utas #' + seq + ' gagal terbit (' + errMsg + ')';
          abortedGroupIds[rowGId] = childAbortMsg;
          archiveSubsequentChainsToPublished(qSheet, pSheet, qData, rowGId, seq, childAbortMsg, stCol, errCol, gIdCol, seqCol, rowsToDelete);
        }
      }
    }

    // 4. Penghapusan Aman dari Arah Bawah ke Atas (Reverse Loop)
    if (rowsToDelete.length > 0) {
      var uniqueRows = [];
      for (var u = 0; u < rowsToDelete.length; u++) {
        if (uniqueRows.indexOf(rowsToDelete[u]) === -1) uniqueRows.push(rowsToDelete[u]);
      }
      uniqueRows.sort(function(a, b) { return b - a; });
      for (var d = 0; d < uniqueRows.length; d++) {
        qSheet.deleteRow(uniqueRows[d]);
      }
    }

    SpreadsheetApp.flush();

  } catch (outerErr) {
    Logger.log('Exception in processPublishingQueue: ' + outerErr.message);
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      Logger.log('Lock release warning (processPublishingQueue): ' + e.message);
    }
  }
}

/**
 * Helper: Menggagalkan seluruh utas lanjutan di Queue_Posts jika utas induk gagal dan memindahkannya ke Published_Posts
 */
function archiveSubsequentChainsToPublished(qSheet, pSheet, qData, groupId, failedSeq, errorMsg, stCol, errCol, gIdCol, seqCol, rowsToDelete) {
  for (var i = 1; i < qData.length; i++) {
    var rowGId = String(qData[i][gIdCol] || qData[i][0]);
    var curSeq = Number(qData[i][seqCol] || 1);
    if (rowGId === groupId && curSeq > failedSeq) {
      var currentStatus = String(qData[i][stCol] || '');
      if (currentStatus === 'Scheduled' || currentStatus === 'Waiting Views' || currentStatus === 'Processing') {
        qData[i][stCol] = 'Failed - Parent Error';
        qData[i][errCol] = errorMsg;
        pSheet.appendRow(qData[i]);
        if (rowsToDelete.indexOf(i + 1) === -1) {
          rowsToDelete.push(i + 1);
        }
      }
    }
  }
}

/**
 * Alias untuk kompatibilitas trigger penamaan lama
 */
function processPostQueue() {
  return processPublishingQueue();
}

function publishContainer(metaUserId, creationId, token) {
  return publishMetaContainer(metaUserId, creationId, token);
}

function createMediaSnapshot(rawMediaList, mediaRows) {
  if (!Array.isArray(rawMediaList)) return [];
  var snapshots = [];
  var idCol = 0, typeCol = 4, urlCol = 5;
  if (mediaRows && mediaRows.length > 0) {
    var h = mediaRows[0];
    if (h.indexOf('media_id') !== -1) idCol = h.indexOf('media_id');
    if (h.indexOf('media_type') !== -1) typeCol = h.indexOf('media_type');
    if (h.indexOf('cloudinary_url') !== -1) urlCol = h.indexOf('cloudinary_url');
  }

  for (var i = 0; i < rawMediaList.length; i++) {
    var item = rawMediaList[i];
    if (!item) continue;
    var extId = '';
    var extUrl = '';
    var extType = 'IMAGE';

    if (typeof item === 'object' && item !== null) {
      extId = item.media_id ? String(item.media_id).trim() : '';
      var candUrl = String(item.url || item.cloudinary_url || '').trim();
      if (/^https?:\/\//i.test(candUrl)) extUrl = candUrl;
      extType = (item.type || item.media_type || 'IMAGE').toUpperCase();
    } else if (typeof item === 'string') {
      var s = item.trim();
      if (/^https?:\/\//i.test(s)) extUrl = s;
      else extId = s;
    }

    if (extId && mediaRows && mediaRows.length > 1) {
      for (var m = 1; m < mediaRows.length; m++) {
        if (String(mediaRows[m][idCol]).trim() === extId) {
          extUrl = String(mediaRows[m][urlCol] || '').trim();
          extType = String(mediaRows[m][typeCol] || 'IMAGE').toUpperCase();
          break;
        }
      }
    }

    if (extUrl) {
      var isVid = extType === 'VIDEO' || /\.(mp4|mov|webm)$/i.test(extUrl);
      snapshots.push({
        media_id: extId,
        url: extUrl,
        type: isVid ? 'VIDEO' : 'IMAGE'
      });
    }
  }
  return snapshots;
}

function findMediaById(mediaId, mediaRows) {
  if (!mediaId) return null;
  var idStr = typeof mediaId === 'object' && mediaId !== null ? (mediaId.media_id || mediaId.url || mediaId.cloudinary_url || '') : String(mediaId).trim();
  if (!idStr) return null;

  if (/^https?:\/\//i.test(idStr)) {
    var isVideo = /\.(mp4|mov|webm)$/i.test(idStr) || (idStr.indexOf('/video/') !== -1);
    return {
      media_type: isVideo ? 'VIDEO' : 'IMAGE',
      cloudinary_url: idStr
    };
  }

  if (mediaRows && mediaRows.length > 1) {
    var headers = mediaRows[0];
    var idCol = headers.indexOf('media_id');
    var typeCol = headers.indexOf('media_type');
    var urlCol = headers.indexOf('cloudinary_url');
    if (idCol === -1) idCol = 0;
    if (typeCol === -1) typeCol = 4;
    if (urlCol === -1) urlCol = 5;

    for (var i = 1; i < mediaRows.length; i++) {
      if (String(mediaRows[i][idCol]).trim() === idStr) {
        var rawUrl = String(mediaRows[i][urlCol] || '').trim();
        var rawType = String(mediaRows[i][typeCol] || '').toUpperCase();
        var isVid = rawType === 'VIDEO' || /\.(mp4|mov|webm)$/i.test(rawUrl);
        return {
          media_type: isVid ? 'VIDEO' : 'IMAGE',
          cloudinary_url: rawUrl
        };
      }
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

    var picCol = headers.indexOf('profile_picture_url');
    var nameCol = headers.indexOf('name');
    var userCol = headers.indexOf('username');

    if (rowIdx > 0) {
      if (activeCol >= 0) sheet.getRange(rowIdx, activeCol + 1).setValue('TRUE');
      if (picCol >= 0 && data.threads_profile_picture_url) {
        sheet.getRange(rowIdx, picCol + 1).setValue(data.threads_profile_picture_url);
      }
      if (nameCol >= 0 && data.name) {
        sheet.getRange(rowIdx, nameCol + 1).setValue(data.name);
      }
      if (userCol >= 0 && data.username) {
        sheet.getRange(rowIdx, userCol + 1).setValue(data.username);
      }
      SpreadsheetApp.flush();
    }

    return {
      ok: true,
      data: {
        account_id: accountId,
        is_active: 'TRUE',
        username: data.username || targetAcc.username,
        name: data.name || targetAcc.name,
        profile_picture_url: data.threads_profile_picture_url || targetAcc.profile_picture_url || '',
        meta_user_id: data.id
      },
      message: 'Akun @' + (data.username || targetAcc.username) + ' aktif dan data profil berhasil disinkronkan!'
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
 * Post Insights & Comments Synchronizer (Smart Window 48 Jam)
 * Menyatukan penarikan metrik (views, likes, reposts) dan komentar baru
 * Khusus untuk postingan yang aktif/diterbitkan dalam 48 jam terakhir.
 */
function syncPostInsights() {
  return syncPostAnalytics();
}

function syncPostAnalytics() {
  var lock = LockService.getScriptLock();
  var acquired = false;
  try {
    acquired = lock.tryLock(10000);
  } catch (e) {
    acquired = false;
  }
  if (!acquired) {
    Logger.log('Sinkronisasi analitik & komentar dibatalkan: Sedang dikunci proses lain.');
    return;
  }

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var analyticsSheet = ss.getSheetByName('Post_Analytics');
    if (!analyticsSheet) {
      analyticsSheet = ss.insertSheet('Post_Analytics');
      analyticsSheet.appendRow(['analytics_id', 'threads_post_id', 'views_count', 'likes_count', 'reposts_count', 'replies_count', 'last_synced_at']);
    }

    var publishedRows = readSheetAsJson('Published_Posts');
    var accounts = readSheetAsJson('Accounts');
    var anaRows = analyticsSheet.getDataRange().getValues();
    var anaHeaders = anaRows[0] || [];
    var thColIdx = anaHeaders.indexOf('threads_post_id');
    if (thColIdx === -1) thColIdx = 1;

    var nowTime = new Date().getTime();
    var windowMs = 48 * 60 * 60 * 1000;
    var syncedCommentGroupIds = {};

    for (var i = publishedRows.length - 1; i >= 0; i--) {
      var pubThread = publishedRows[i];
      var status = String(pubThread.status || '').trim();
      var threadsPostId = String(pubThread.threads_post_id || '').trim();

      // Guard Clause: Hanya proses pos berstatus Published murni dengan Meta threads_post_id valid
      if (status !== 'Published' || !threadsPostId) {
        continue;
      }

      var pubDateVal = pubThread.published_at || pubThread.scheduled_at || pubThread.created_at;
      if (pubDateVal) {
        var pubTime = new Date(pubDateVal).getTime();
        if (!isNaN(pubTime) && (nowTime - pubTime > windowMs)) {
          continue;
        }
      }

      var token = '';
      var threadAccId = String(pubThread.account_id || '').trim();
      if (!threadAccId && pubThread.chain_group_id) {
        for (var pr = 0; pr < publishedRows.length; pr++) {
          if (String(publishedRows[pr].chain_group_id || publishedRows[pr].post_id) === String(pubThread.chain_group_id) && publishedRows[pr].account_id) {
            threadAccId = String(publishedRows[pr].account_id).trim();
            break;
          }
        }
      }
      for (var a = 0; a < accounts.length; a++) {
        if (String(accounts[a].account_id) === threadAccId) {
          token = accounts[a].access_token;
          break;
        }
      }
      if (!token) continue;

      try {
        var insightUrl = 'https://graph.threads.net/v1.0/' + encodeURIComponent(threadsPostId) + '/insights?metric=views,likes,replies,reposts&access_token=' + encodeURIComponent(token);
        var resp = UrlFetchApp.fetch(insightUrl, { muteHttpExceptions: true });
        var data = JSON.parse(resp.getContentText());

        if (!data.error && data.data && Array.isArray(data.data)) {
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

          var anaRowIdx = -1;
          for (var r = 1; r < anaRows.length; r++) {
            if (String(anaRows[r][thColIdx]) === threadsPostId) {
              anaRowIdx = r + 1;
              break;
            }
          }

          var nowIso = new Date().toISOString();
          if (anaRowIdx > 0) {
            analyticsSheet.getRange(anaRowIdx, 3, 1, 5).setValues([[views, likes, reposts, replies, nowIso]]);
          } else {
            var anaId = 'ana_' + Utilities.getUuid().slice(0, 8);
            analyticsSheet.appendRow([anaId, threadsPostId, views, likes, reposts, replies, nowIso]);
            anaRows.push([anaId, threadsPostId, views, likes, reposts, replies, nowIso]);
          }
        }
      } catch (err) {
        Logger.log('Gagal sinkronisasi analitik [' + threadsPostId + ']: ' + err.message);
      }

      var grpId = String(pubThread.chain_group_id || pubThread.post_id);
      if (!syncedCommentGroupIds[grpId]) {
        syncedCommentGroupIds[grpId] = true;
        try {
          syncPostComments(grpId);
        } catch (cErr) {
          Logger.log('Gagal auto-sync komentar [' + grpId + ']: ' + cErr.message);
        }
      }
    }
  } catch (outerErr) {
    Logger.log('Exception in syncPostAnalytics: ' + outerErr.message);
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      Logger.log('Lock release warning (syncPostAnalytics): ' + e.message);
    }
  }
}
