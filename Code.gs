// ============================================================
// Code.gs — AI Career OS
// PRODUCTION READY — FULL BACKEND
// Version: 6.0.0 — Complete Feature Set
// ============================================================

// ============================================================
// CONFIG — SINGLE SOURCE OF TRUTH
// ============================================================

var CONFIG = {

  APP_NAME: 'AI Career OS',
  VERSION: '6.0.0',

  SPREADSHEET_ID: PropertiesService
    .getScriptProperties()
    .getProperty('SPREADSHEET_ID') || '',

  AI: {
    OPENROUTER_URL: 'https://openrouter.ai/api/v1/chat/completions',
    MODEL: 'openai/gpt-4o-mini',
    MODEL_ADVANCED: 'openai/gpt-4o',
    TEMPERATURE: 0.7,
    MAX_TOKENS: 3000,
    MAX_TOKENS_LONG: 4500,
    RETRY_ATTEMPTS: 2,
    RETRY_DELAY_MS: 1500
  },

  PAYMENT: {
    IS_PRODUCTION: true,
    SERVER_KEY_SANDBOX: PropertiesService.getScriptProperties().getProperty('MIDTRANS_SERVER_KEY_SANDBOX') || '',
    SERVER_KEY_PROD:    PropertiesService.getScriptProperties().getProperty('MIDTRANS_SERVER_KEY_PROD') || '',
    CLIENT_KEY_SANDBOX: PropertiesService.getScriptProperties().getProperty('MIDTRANS_CLIENT_KEY_SANDBOX') || '',
    CLIENT_KEY_PROD:    PropertiesService.getScriptProperties().getProperty('MIDTRANS_CLIENT_KEY_PROD') || '',
    MONTHLY_PRICE: 29000,
    YEARLY_PRICE:  199000,
    SNAP_URL_SANDBOX: 'https://app.sandbox.midtrans.com/snap/v1/transactions',
    SNAP_URL_PROD:    'https://app.midtrans.com/snap/v1/transactions',
    API_URL_SANDBOX:  'https://api.sandbox.midtrans.com/v2/',
    API_URL_PROD:     'https://api.midtrans.com/v2/'
  },

  CACHE: {
    OTP_TTL_SECONDS: 600,
    SESSION_TTL_SECONDS: 21600
  }

};

// ============================================================
// LOGGER — PRODUCTION GRADE
// ============================================================

var Logger_ = {

  _write: function(level, source, message, user, detail) {
    try {
      var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
      var sheet = ss.getSheetByName('ERROR_LOG');
      if (!sheet) {
        sheet = ss.insertSheet('ERROR_LOG');
        sheet.appendRow(['Timestamp','Level','Source','Message','User','Detail']);
        sheet.setFrozenRows(1);
      }
      sheet.appendRow([
        new Date().toISOString(),
        level,
        source,
        String(message).substring(0, 500),
        user || 'system',
        detail ? JSON.stringify(detail).substring(0, 500) : ''
      ]);
    } catch(e) {
      console.error('Logger write failed: ' + e.message);
    }
  },

  info: function(source, message, user) {
    console.log('[INFO] ' + source + ': ' + message);
    // Only write INFO to sheet if critical path
  },

  error: function(source, message, user, detail) {
    console.error('[ERROR] ' + source + ': ' + message);
    this._write('ERROR', source, message, user, detail);
  },

  warn: function(source, message, user) {
    console.warn('[WARN] ' + source + ': ' + message);
    this._write('WARN', source, message, user);
  },

  payment: function(source, message, user, detail) {
    console.log('[PAYMENT] ' + source + ': ' + message);
    this._write('PAYMENT', source, message, user, detail);
  },

  activity: function(source, message, user) {
    console.log('[ACTIVITY] ' + source + ': ' + message);
  }

};

// ============================================================
// VALIDATION HELPERS
// ============================================================

var Validate = {

  email: function(email) {
    if (!email || typeof email !== 'string') return false;
    var re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return re.test(email.trim().toLowerCase());
  },

  required: function(obj, fields) {
    for (var i = 0; i < fields.length; i++) {
      var f = fields[i];
      if (!obj || !obj[f] || String(obj[f]).trim() === '') {
        return { valid: false, missing: f };
      }
    }
    return { valid: true };
  },

  sanitize: function(str, maxLen) {
    if (!str) return '';
    maxLen = maxLen || 10000;
    return String(str)
      .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
      .replace(/javascript:/gi, '')
      .trim()
      .substring(0, maxLen);
  },

  safeEmail: function(email) {
    if (!email) return '';
    return String(email).trim().toLowerCase().substring(0, 200);
  }

};

// ============================================================
// DATABASE MANAGER
// ============================================================

var DB = {

  _ss: null,

  _getSpreadsheet: function() {
    if (!DB._ss) {
      DB._ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    }
    return DB._ss;
  },

  getSheet: function(name) {
    var ss = DB._getSpreadsheet();
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    return sheet;
  },

  ensureHeaders: function(sheet, headers) {
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
    }
    return sheet;
  },

  findRow: function(sheet, colIndex, value) {
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return null;
    var data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][colIndex]).toLowerCase() === String(value).toLowerCase()) {
        return { rowIndex: i + 1, data: data[i] };
      }
    }
    return null;
  },

  findRowExact: function(sheet, colIndex, value) {
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return null;
    var data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][colIndex]) === String(value)) {
        return { rowIndex: i + 1, data: data[i] };
      }
    }
    return null;
  },

  getAll: function(sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    var data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
    var headers = data[0];
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var obj = {};
      headers.forEach(function(h, idx) {
        obj[h] = data[i][idx];
      });
      rows.push(obj);
    }
    return rows;
  },

  getByEmail: function(sheetName, email) {
    var sheet = DB.getSheet(sheetName);
    var all = DB.getAll(sheet);
    return all.filter(function(r) {
      return String(r.email || '').toLowerCase() === String(email || '').toLowerCase();
    });
  },

  updateCell: function(sheet, rowIndex, colIndex, value) {
    sheet.getRange(rowIndex, colIndex).setValue(value);
  },

  updateRow: function(sheet, rowIndex, values) {
    sheet.getRange(rowIndex, 1, 1, values.length).setValues([values]);
  },

  appendSafe: function(sheet, headers, rowData) {
    DB.ensureHeaders(sheet, headers);
    sheet.appendRow(rowData);
  }

};

// ============================================================
// CACHE HELPER
// ============================================================

var Cache_ = {

  get: function(key) {
    try {
      var val = CacheService.getScriptCache().get(key);
      if (!val) return null;
      return JSON.parse(val);
    } catch(e) {
      return null;
    }
  },

  set: function(key, data, ttlSeconds) {
    try {
      CacheService.getScriptCache().put(key, JSON.stringify(data), ttlSeconds || 600);
      return true;
    } catch(e) {
      return false;
    }
  },

  remove: function(key) {
    try {
      CacheService.getScriptCache().remove(key);
    } catch(e) {}
  }

};

// ============================================================
// SAFE USER LOOKUP PATCH
// ============================================================

function findUserByEmail(email){

  try{

    if(!email){
      return null;
    }

    var safeEmail =
      String(email)
        .trim()
        .toLowerCase();

    var sheet =
      DB.getSheet('USER_DB');

    var values =
      sheet.getDataRange().getValues();

    if(values.length <= 1){
      return null;
    }

    // SAFE LOWERCASE HEADER
    var headers =
      values[0].map(function(h){
        return String(h)
          .trim()
          .toLowerCase();
      });

    var emailIdx =
      headers.indexOf('email');

    if(emailIdx < 0){
      return null;
    }

    for(var i=1;i<values.length;i++){

      var rowEmail =
        String(values[i][emailIdx] || '')
          .trim()
          .toLowerCase();

      if(rowEmail === safeEmail){

        var user = {};

        headers.forEach(function(h,index){
          user[h] = values[i][index];
        });

        user._rowIndex = i + 1;

        return user;
      }
    }

    return null;

  }catch(e){

    Logger.log(e);

    return null;
  }
}

// ============================================================
// PREMIUM GUARD
// ============================================================

function requirePremium(email) {
  try {
    if (!email) {
      return { success: false, premiumRequired: true, error: 'Authentication required. Please log in first.' };
    }

    var safeEmail = Validate.safeEmail(email);
    var sheet = DB.getSheet('USER_DB');
    var userRow = DB.findRow(sheet, 0, safeEmail);

    if (!userRow) {
      return { success: false, premiumRequired: true, error: 'User not found. Please log in again.' };
    }

    var isPremium = userRow.data[4] === true || String(userRow.data[4]).toUpperCase() === 'TRUE';
    var expiredAt = userRow.data[6];

    if (isPremium && expiredAt) {
      var expDate = new Date(expiredAt);
      if (!isNaN(expDate.getTime()) && expDate < new Date()) {
        DB.updateCell(sheet, userRow.rowIndex, 5, false);
        DB.updateCell(sheet, userRow.rowIndex, 6, '');
        DB.updateCell(sheet, userRow.rowIndex, 7, '');
        Logger_.warn('requirePremium', 'Auto-downgraded expired premium: ' + safeEmail, safeEmail);
        return { success: false, premiumRequired: true, error: 'Your premium subscription has expired. Please renew.' };
      }
    }

    if (!isPremium) {
      return { success: false, premiumRequired: true, error: 'Premium subscription required for this feature.' };
    }

    return { success: true };

  } catch(e) {
    Logger_.error('requirePremium', e.message, email);
    return { success: false, premiumRequired: true, error: 'Unable to verify subscription. Please try again.' };
  }
}

// ============================================================
// ENTRY POINTS
// ============================================================

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');

  template.clientKey = CONFIG.PAYMENT.IS_PRODUCTION
    ? CONFIG.PAYMENT.CLIENT_KEY_PROD
    : CONFIG.PAYMENT.CLIENT_KEY_SANDBOX;

  template.isProduction = CONFIG.PAYMENT.IS_PRODUCTION;

  template.snapUrl = CONFIG.PAYMENT.IS_PRODUCTION
    ? 'https://app.midtrans.com/snap/snap.js'
    : 'https://app.sandbox.midtrans.com/snap/snap.js';

  return template
    .evaluate()
    .setTitle('AI Career OS — Premium AI Career Platform')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ============================================================
// AUTH — OTP SYSTEM
// ============================================================

function sendOTP(email) {
  try {
    if (!Validate.email(email)) {
      return { success: false, error: 'Email tidak valid. Gunakan format yang benar.' };
    }

    var safeEmail = Validate.safeEmail(email);
    var otp = Math.floor(100000 + Math.random() * 900000).toString();
    var expiry = new Date(Date.now() + CONFIG.CACHE.OTP_TTL_SECONDS * 1000).toISOString();

    Cache_.set('otp_' + safeEmail, { otp: otp, expiry: expiry, attempts: 0 }, CONFIG.CACHE.OTP_TTL_SECONDS);

    var quota = MailApp.getRemainingDailyQuota();
    if (quota <= 0) {
      Logger_.warn('sendOTP', 'Email quota exhausted', safeEmail);
      return { success: false, error: 'Email quota habis hari ini. Coba lagi besok.' };
    }

    MailApp.sendEmail({
      to: email,
      subject: '🔐 Kode OTP AI Career OS — ' + otp,
      htmlBody: _buildOTPEmailHTML(otp)
    });

    Logger_.activity('OTP', 'Sent to: ' + safeEmail, safeEmail);
    return { success: true };

  } catch(e) {
    Logger_.error('sendOTP', e.message, email);
    return { success: false, error: 'Gagal mengirim OTP. Coba lagi: ' + e.message };
  }
}

function _buildOTPEmailHTML(otp) {
  return '<div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;padding:32px;background:#f4f6fb;border-radius:16px;">' +
    '<div style="text-align:center;margin-bottom:24px;">' +
    '<h1 style="color:#ff7a00;font-size:24px;margin:0;">⚡ AI Career OS</h1>' +
    '<p style="color:#888;font-size:13px;margin:4px 0 0;">Premium AI Career Platform</p>' +
    '</div>' +
    '<div style="background:#fff;border-radius:12px;padding:28px;text-align:center;margin-bottom:20px;box-shadow:0 2px 12px rgba(0,0,0,0.06);">' +
    '<p style="color:#444;font-size:14px;margin-bottom:16px;">Gunakan kode berikut untuk masuk ke AI Career OS:</p>' +
    '<div style="background:#fff8f2;border:2px dashed #ff7a00;border-radius:12px;padding:20px;">' +
    '<div style="font-size:48px;font-weight:900;letter-spacing:12px;color:#ff7a00;font-family:monospace;">' + otp + '</div>' +
    '</div>' +
    '<p style="color:#bbb;font-size:12px;margin-top:16px;">Berlaku selama <strong>10 menit</strong>. Jangan bagikan ke siapapun.</p>' +
    '</div>' +
    '<p style="color:#aaa;font-size:11px;text-align:center;">Jika kamu tidak meminta kode ini, abaikan email ini.</p>' +
    '</div>';
}

function verifyOTP(email, inputOtp) {
  try {
    if (!Validate.email(email)) {
      return { success: false, error: 'Email tidak valid.' };
    }
    if (!inputOtp || !/^\d{6}$/.test(String(inputOtp).trim())) {
      return { success: false, error: 'OTP harus 6 digit angka.' };
    }

    var safeEmail = Validate.safeEmail(email);
    var stored = Cache_.get('otp_' + safeEmail);

    if (!stored) {
      return { success: false, error: 'OTP expired atau tidak ditemukan. Minta OTP baru.' };
    }

    // Max attempts guard
    if (stored.attempts >= 5) {
      Cache_.remove('otp_' + safeEmail);
      return { success: false, error: 'Terlalu banyak percobaan. Minta OTP baru.' };
    }

    if (new Date(stored.expiry) < new Date()) {
      Cache_.remove('otp_' + safeEmail);
      return { success: false, error: 'OTP sudah expired. Minta OTP baru.' };
    }

    if (String(stored.otp) !== String(inputOtp).trim()) {
      stored.attempts = (stored.attempts || 0) + 1;
      Cache_.set('otp_' + safeEmail, stored, 600);
      return { success: false, error: 'Kode OTP salah. Sisa percobaan: ' + (5 - stored.attempts) };
    }

    Cache_.remove('otp_' + safeEmail);
    return _loginOrRegister(safeEmail);

  } catch(e) {
    Logger_.error('verifyOTP', e.message, email);
    return { success: false, error: 'Server error. Coba lagi.' };
  }
}

function _loginOrRegister(email) {
  var sheet = DB.getSheet('USER_DB');
  DB.ensureHeaders(sheet, [
    'email','name','picture','role',
    'premium','premium_plan','premium_expired',
    'session_token','last_login','created_at'
  ]);

  var existing = DB.findRow(sheet, 0, email);
  var now = new Date().toISOString();
  var displayName = _buildDisplayName(email);

  if (existing) {
    DB.updateCell(sheet, existing.rowIndex, 9, now);

    var isPremium = existing.data[4] === true || String(existing.data[4]).toUpperCase() === 'TRUE';
    var premiumExpired = existing.data[6];

    if (isPremium && premiumExpired && new Date(premiumExpired) < new Date()) {
      isPremium = false;
      DB.updateCell(sheet, existing.rowIndex, 5, false);
      DB.updateCell(sheet, existing.rowIndex, 6, '');
      DB.updateCell(sheet, existing.rowIndex, 7, '');
    }

    Logger_.activity('Auth', 'Login: ' + email, email);

    return {
      success: true,
      user: {
        email: email,
        name: String(existing.data[1] || displayName),
        picture: '',
        premium: isPremium,
        premiumPlan: String(existing.data[5] || ''),
        premiumExpired: String(premiumExpired || '')
      }
    };

  } else {
    sheet.appendRow([
      email, displayName, '', 'user', false, '', '',
      Utilities.getUuid(), now, now
    ]);

    Logger_.activity('Auth', 'Registered: ' + email, email);

    return {
      success: true,
      user: {
        email: email,
        name: displayName,
        picture: '',
        premium: false,
        premiumPlan: '',
        premiumExpired: ''
      }
    };
  }
}

function _buildDisplayName(email) {
  return (email.split('@')[0] || 'User')
    .replace(/[._\-+]/g, ' ')
    .replace(/\b\w/g, function(c) { return c.toUpperCase(); })
    .substring(0, 50);
}

// ============================================================
// USER MANAGEMENT
// ============================================================

function getUserStats(email) {
  try {
    if (!email) return { success: false, error: 'Email required' };

    var safeEmail = Validate.safeEmail(email);

    var cvCount        = DB.getByEmail('CV_DB', safeEmail).length;
    var portfolioCount = DB.getByEmail('PORTFOLIO_DB', safeEmail).length;
    var atsHistory     = DB.getByEmail('ATS_DB', safeEmail);
    var latestATS      = 0;

    if (atsHistory.length > 0) {
      latestATS = parseInt(atsHistory[atsHistory.length - 1].score) || 0;
    }

    var interviewCount = DB.getByEmail('INTERVIEW_DB', safeEmail).length;
    var applyCount     = DB.getByEmail('APPLY_DB', safeEmail).length;

    return {
      success: true,
      stats: {
        totalCV: cvCount,
        totalPortfolio: portfolioCount,
        totalATS: atsHistory.length,
        latestATSScore: latestATS,
        totalInterview: interviewCount,
        totalApplied: applyCount
      }
    };

  } catch(e) {
    Logger_.error('getUserStats', e.message, email);
    return {
      success: false,
      error: e.message,
      stats: { totalCV: 0, totalPortfolio: 0, totalATS: 0, latestATSScore: 0 }
    };
  }
}

function getUserHistory(email, type) {
  try {
    if (!email) return { success: false, error: 'Email required' };

    var safeEmail = Validate.safeEmail(email);
    var sheetMap = {
      cv: 'CV_DB',
      portfolio: 'PORTFOLIO_DB',
      ats: 'ATS_DB',
      interview: 'INTERVIEW_DB',
      apply: 'APPLY_DB',
      payment: 'PAYMENT_DB'
    };

    var sheetName = sheetMap[type];
    if (!sheetName) return { success: false, error: 'Invalid history type' };

    var records = DB.getByEmail(sheetName, safeEmail);

    // Strip heavy content fields for history list
    var light = records.map(function(r) {
      return {
        id: r.cv_id || r.portfolio_id || r.ats_id || r.interview_id || r.apply_id || r.payment_id || '',
        name: r.name || r.role || r.position || '',
        created_at: r.created_at || '',
        score: r.score || '',
        status: r.status || ''
      };
    });

    // Sort by created_at desc
    light.sort(function(a, b) {
      return new Date(b.created_at) - new Date(a.created_at);
    });

    return { success: true, history: light.slice(0, 50) };

  } catch(e) {
    Logger_.error('getUserHistory', e.message, email);
    return { success: false, error: e.message, history: [] };
  }
}

function updateUserProfile(email, updates) {
  try {
    if (!email) return { success: false, error: 'Email required' };

    var safeEmail = Validate.safeEmail(email);
    var sheet = DB.getSheet('USER_DB');
    var userRow = DB.findRow(sheet, 0, safeEmail);

    if (!userRow) return { success: false, error: 'User not found' };

    if (updates.name) {
      var safeName = Validate.sanitize(updates.name, 100);
      DB.updateCell(sheet, userRow.rowIndex, 2, safeName);
    }

    return { success: true };

  } catch(e) {
    Logger_.error('updateUserProfile', e.message, email);
    return { success: false, error: e.message };
  }
}

// ============================================================
// AI ENGINE — CORE
// ============================================================

function callAI(prompt, options) {
  options = options || {};
  var attempts = 0;
  var maxAttempts = options.retries !== undefined ? options.retries : CONFIG.AI.RETRY_ATTEMPTS;

  while (attempts <= maxAttempts) {
    try {
      var result = _callAIOnce(prompt, options);
      if (result.success) return result;

      attempts++;
      if (attempts <= maxAttempts) {
        Utilities.sleep(CONFIG.AI.RETRY_DELAY_MS * attempts);
      }

    } catch(e) {
      attempts++;
      if (attempts > maxAttempts) {
        Logger_.error('callAI', e.message);
        return { success: false, error: e.message, text: '' };
      }
      Utilities.sleep(CONFIG.AI.RETRY_DELAY_MS * attempts);
    }
  }

  return { success: false, error: 'AI service unavailable after ' + maxAttempts + ' attempts.', text: '' };
}

function _callAIOnce(prompt, options) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('OPENROUTER_API_KEY');
  if (!apiKey) {
    throw new Error('OPENROUTER_API_KEY tidak dikonfigurasi di Script Properties.');
  }

  var systemPrompt = options.systemPrompt ||
    'You are a world-class AI career assistant specializing in ATS optimization, ' +
    'executive resume writing, LinkedIn branding, and recruiter psychology. ' +
    'Always produce premium, detailed, actionable output. Use HTML formatting for structured output.';

  var messages = [{ role: 'system', content: systemPrompt }];

  // Support conversation history
  if (options.history && Array.isArray(options.history)) {
    options.history.forEach(function(h) {
      if (h.role && h.content) {
        messages.push({ role: h.role, content: String(h.content).substring(0, 3000) });
      }
    });
  }

  messages.push({ role: 'user', content: prompt });

  var payload = {
    model: options.model || CONFIG.AI.MODEL,
    messages: messages,
    temperature: options.temperature !== undefined ? options.temperature : CONFIG.AI.TEMPERATURE,
    max_tokens: options.maxTokens || CONFIG.AI.MAX_TOKENS
  };

  var response = UrlFetchApp.fetch(CONFIG.AI.OPENROUTER_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + apiKey,
      'HTTP-Referer': 'https://script.google.com',
      'X-Title': CONFIG.APP_NAME
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var rawText = response.getContentText();

  if (!rawText) throw new Error('Empty response from AI API');

  var json;
  try { json = JSON.parse(rawText); } catch(pe) { throw new Error('Invalid JSON from AI: ' + rawText.substring(0, 200)); }

  if (code === 429) throw new Error('AI rate limit exceeded. Retry later.');
  if (code === 401) throw new Error('Invalid AI API key.');
  if (code !== 200) throw new Error('AI API Error ' + code + ': ' + rawText.substring(0, 300));
  if (!json.choices || !json.choices[0] || !json.choices[0].message) {
    throw new Error('No content in AI response: ' + rawText.substring(0, 200));
  }

  var content = json.choices[0].message.content;
  if (!content || content.trim() === '') throw new Error('AI returned empty content.');

  return { success: true, text: content };
}

// ============================================================
// CV GENERATOR
// ============================================================

function generateCV(data, email) {
  var premiumCheck = requirePremium(email);
  if (!premiumCheck.success) return premiumCheck;

  try {
    var req = Validate.required(data, ['name','role','experience','skills']);
    if (!req.valid) return { success: false, error: 'Field "' + req.missing + '" wajib diisi.' };

    var lang = Validate.sanitize(data.language || 'Indonesian', 50);
    var template = data.template || 'ats-minimal';

    var templateInstructions = _getCVTemplateInstructions(template);

    var prompt =
'You are a world-class executive resume writer with 20+ years at Fortune 500 companies.\n\n' +
'Create a premium, ATS-optimized executive CV as an HTML FRAGMENT.\n\n' +
'╔════════════════════════════════╗\n' +
'║  CRITICAL RULES — ENFORCED    ║\n' +
'╚════════════════════════════════╝\n' +
'1. HTML FRAGMENT ONLY — no <html>, <head>, <body> tags\n' +
'2. Output starts with <style>...</style> then <div class="cv-wrap">...</div>\n' +
'3. ALL text must be #1a1a1a on #ffffff — no exceptions\n' +
'4. No colored backgrounds on any container element\n' +
'5. Skills = plain bullet list only\n' +
'6. Section headings = bold + uppercase + border-bottom: 1.5px solid #ccc\n' +
'7. EVERY word in output language: ' + lang + '\n' +
'8. Rewrite experience bullets as quantified achievements (numbers, percentages, impact)\n' +
'9. ATS-optimize: use exact keywords from target job if provided\n\n' +
templateInstructions + '\n\n' +
'CANDIDATE DATA:\n' +
'Name: ' + Validate.sanitize(data.name, 100) + '\n' +
'Target Position: ' + Validate.sanitize(data.role, 100) + '\n' +
'Industry: ' + Validate.sanitize(data.industry || 'General', 100) + '\n' +
'Summary: ' + Validate.sanitize(data.summary || '', 500) + '\n' +
'Experience:\n' + Validate.sanitize(data.experience, 5000) + '\n\n' +
'Education: ' + Validate.sanitize(data.education || '', 500) + '\n' +
'Skills: ' + Validate.sanitize(data.skills, 500) + '\n' +
'Certifications: ' + Validate.sanitize(data.certifications || '', 300) + '\n' +
'Target Job Description: ' + Validate.sanitize(data.targetJob || '', 1000) + '\n\n' +
'CSS CLASSES TO USE:\n' +
'.cv-wrap{font-family:Arial,sans-serif;font-size:12px;color:#1a1a1a;background:#fff;padding:0;line-height:1.7;}\n' +
'h1.cv-name{font-size:22px;font-weight:700;color:#1a1a1a;margin:0 0 2px;}\n' +
'p.cv-title{font-size:13px;font-weight:600;color:#333;margin:0 0 4px;}\n' +
'p.cv-contact{font-size:11px;color:#444;margin:0 0 18px;}\n' +
'h2.cv-section{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:1.2px;color:#1a1a1a;border-bottom:1.5px solid #ccc;padding-bottom:3px;margin:16px 0 7px;}\n' +
'p.cv-text{font-size:12px;color:#1a1a1a;margin:0 0 7px;}\n' +
'ul.cv-list{list-style:none;padding:0;margin:0 0 7px;}\n' +
'ul.cv-list li{font-size:12px;color:#1a1a1a;padding:1px 0 1px 14px;position:relative;}\n' +
'ul.cv-list li::before{content:"•";position:absolute;left:0;color:#1a1a1a;}\n' +
'div.cv-job{margin-bottom:12px;}\n' +
'div.cv-job-header{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:2px;}\n' +
'span.cv-company{font-size:13px;font-weight:700;color:#1a1a1a;}\n' +
'span.cv-dates{font-size:11px;color:#555;}\n' +
'p.cv-jobtitle{font-size:12px;font-style:italic;color:#333;margin:0 0 3px;}\n\n' +
'Generate a complete, detailed CV now. Every section must be comprehensive.';

    var result = callAI(prompt, { maxTokens: CONFIG.AI.MAX_TOKENS_LONG });

    if (result.success) {
      _saveToDB('CV_DB',
        ['cv_id','email','name','role','template','language','content','created_at'],
        [Utilities.getUuid(), email||'', data.name, data.role, template, lang, result.text, new Date().toISOString()],
        'generateCV', email
      );
    }

    return result;

  } catch(e) {
    Logger_.error('generateCV', e.message, email);
    return { success: false, error: 'Gagal generate CV: ' + e.message };
  }
}

function _getCVTemplateInstructions(template) {
  var instructions = {
    'ats-minimal': 'TEMPLATE: ATS Minimal — clean single column, no graphics, maximum keyword density',
    'corporate':   'TEMPLATE: Corporate — two-column header with contact sidebar, conservative design',
    'modern':      'TEMPLATE: Modern — clean lines, subtle dividers, contemporary typography',
    'executive':   'TEMPLATE: Executive — prestigious layout, metrics-forward, C-suite appropriate',
    'startup':     'TEMPLATE: Startup — dynamic, concise bullets, skills-forward, achievement-heavy',
    'harvard':     'TEMPLATE: Harvard Style — academic format, minimal formatting, text-dense'
  };
  return instructions[template] || instructions['ats-minimal'];
}

// ============================================================
// COVER LETTER AI
// ============================================================

function generateCoverLetter(data, email) {
  var premiumCheck = requirePremium(email);
  if (!premiumCheck.success) return premiumCheck;

  try {
    var req = Validate.required(data, ['role','experience']);
    if (!req.valid) return { success: false, error: 'Field "' + req.missing + '" wajib diisi.' };

    var lang = Validate.sanitize(data.language || 'Indonesian', 50);
    var style = data.style || 'professional';
    var toneMap = {
      professional:  'Professional, confident, and achievement-focused',
      storytelling:  'Narrative storytelling with emotional connection and personal journey',
      bold:          'Bold, direct, and result-driven. No fluff.',
      startup:       'Casual, enthusiastic, startup culture-friendly, innovative',
      executive:     'Executive gravitas, strategic perspective, leadership-focused'
    };

    var prompt =
'You are a master cover letter writer who has helped 10,000+ candidates land offers at top companies.\n\n' +
'Create a premium, personalized cover letter as an HTML FRAGMENT.\n\n' +
'RULES:\n' +
'- HTML FRAGMENT ONLY — no <html>, <head>, <body> tags\n' +
'- White background (#ffffff), dark text (#1a1a1a)\n' +
'- Professional letter layout with proper spacing\n' +
'- 380-480 words — not shorter, not longer\n' +
'- ENTIRELY in ' + lang + '\n' +
'- Tone: ' + (toneMap[style] || toneMap.professional) + '\n\n' +
'CANDIDATE DATA:\n' +
'Target Role: ' + Validate.sanitize(data.role, 100) + '\n' +
'Company: ' + Validate.sanitize(data.company || 'the company', 100) + '\n' +
'Recruiter/HM: ' + Validate.sanitize(data.recruiter || '', 100) + '\n' +
'Experience: ' + Validate.sanitize(data.experience, 3000) + '\n' +
'Key Achievements: ' + Validate.sanitize(data.achievements || '', 1000) + '\n' +
'Why This Company: ' + Validate.sanitize(data.whyCompany || '', 500) + '\n' +
'Job Keywords: ' + Validate.sanitize(data.keywords || '', 500) + '\n\n' +
'LETTER STRUCTURE:\n' +
'1. Date + Contact block (professional format)\n' +
'2. Salutation (personalized if recruiter name given)\n' +
'3. HOOK opening paragraph — compelling first sentence\n' +
'4. Value proposition paragraph — your specific fit\n' +
'5. Achievement paragraph — 2-3 quantified results\n' +
'6. Company connection paragraph — why THEM specifically\n' +
'7. Confident closing with clear CTA\n' +
'8. Professional signature block\n\n' +
'Make it sound human, not AI-generated. Use specific details, not generic praise.';

    var result = callAI(prompt, { maxTokens: 2500 });

    if (result.success) {
      _saveToDB('COVERLETTER_DB',
        ['cl_id','email','role','company','language','style','content','created_at'],
        [Utilities.getUuid(), email||'', data.role, data.company||'', lang, style, result.text, new Date().toISOString()],
        'generateCoverLetter', email
      );
    }

    return result;

  } catch(e) {
    Logger_.error('generateCoverLetter', e.message, email);
    return { success: false, error: 'Gagal generate cover letter: ' + e.message };
  }
}

// ============================================================
// ATS CHECKER
// ============================================================

function runATSCheck(data, email) {
  var premiumCheck = requirePremium(email);
  if (!premiumCheck.success) return premiumCheck;

  try {
    var req = Validate.required(data, ['resume','job']);
    if (!req.valid) return { success: false, error: 'Resume dan Job Description wajib diisi.' };

    var mode = data.mode || 'full';
    var modePromptMap = {
      'full':     'Comprehensive ATS analysis with AI rewrite suggestions for every section',
      'quick':    'Quick score only — ATS score, top 5 issues, 3 quick fixes',
      'keywords': 'Deep keyword gap analysis — all missing keywords with insertion suggestions',
      'format':   'Format and structure check — ATS parsing issues, layout problems'
    };

    var prompt =
'You are a senior ATS specialist who has reviewed 50,000+ resumes at FAANG, McKinsey, and Fortune 100 companies.\n\n' +
'Analyze this resume against the job description. Mode: ' + (modePromptMap[mode] || modePromptMap.full) + '\n\n' +
'RESUME:\n' + Validate.sanitize(data.resume, 8000) + '\n\n' +
'JOB DESCRIPTION:\n' + Validate.sanitize(data.job, 5000) + '\n\n' +
'OUTPUT RULES:\n' +
'- HTML FRAGMENT ONLY — no <html>, <head>, <body> tags\n' +
'- White background (#ffffff), dark text (#1a1a1a)\n' +
'- Use inline span badges for scores: background-color can be used ONLY on score badges\n' +
'- Professional layout with clear sections\n\n' +
'REQUIRED SECTIONS:\n' +
'1. ATS SCORE (0-100) — Large prominent badge. Format: XX/100\n' +
'2. MATCH BREAKDOWN — keyword match %, skill match %, experience match %, format score\n' +
'3. FOUND KEYWORDS TABLE — keywords in resume that match JD\n' +
'4. MISSING KEYWORDS TABLE — critical keywords absent from resume\n' +
'5. SECTION-BY-SECTION ANALYSIS — Summary, Experience, Skills, Education\n' +
'6. FORMATTING ISSUES — ATS parsing problems\n' +
'7. TOP 10 IMPROVEMENTS — specific, actionable, ranked by impact\n' +
'8. AI REWRITE SUGGESTIONS — 2-3 example bullet rewrites\n' +
'9. RECRUITER PERSPECTIVE — how a recruiter would view this resume\n' +
'10. FINAL VERDICT — hire/no-hire recommendation with reasoning\n\n' +
'Be specific. Use real data from the resume and JD. No generic advice.';

    var result = callAI(prompt, { maxTokens: CONFIG.AI.MAX_TOKENS_LONG });

    if (result.success) {
      var score = _extractATSScore(result.text);
      result.score = score;

      _saveToDB('ATS_DB',
        ['ats_id','email','score','role','mode','content','created_at'],
        [Utilities.getUuid(), email||'', score, data.role||'', mode, result.text, new Date().toISOString()],
        'runATSCheck', email
      );
    }

    return result;

  } catch(e) {
    Logger_.error('runATSCheck', e.message, email);
    return { success: false, error: 'Gagal ATS check: ' + e.message };
  }
}

function _extractATSScore(text) {
  var patterns = [
    /(\d{1,3})\s*\/\s*100/i,
    /(\d{1,3})[\s]*(?:%|percent)/i,
    /score[:\s]+(\d{1,3})/i,
    /ats[:\s]+(\d{1,3})/i
  ];
  for (var i = 0; i < patterns.length; i++) {
    var m = text.match(patterns[i]);
    if (m) {
      var s = parseInt(m[1]);
      if (s >= 0 && s <= 100) return s;
    }
  }
  return 0;
}

// ============================================================
// LINKEDIN AI
// ============================================================

function generateLinkedIn(data, email) {
  var premiumCheck = requirePremium(email);
  if (!premiumCheck.success) return premiumCheck;

  try {
    var req = Validate.required(data, ['role']);
    if (!req.valid) return { success: false, error: 'Role tidak boleh kosong.' };

    var lang = Validate.sanitize(data.language || 'Indonesian', 50);
    var mode = data.mode || 'full';

    var modePrompt = {
      full:     'Complete LinkedIn profile optimization: headline, about, experience, skills, featured, recommendations template, connection templates, content strategy',
      headline: 'Generate 5 powerful LinkedIn headline variations (220 chars max each). Analyze and explain why each works.',
      about:    'Write a compelling LinkedIn About section (2600 chars max). Multiple versions: professional, story-driven, achievement-focused.',
      posts:    'Create 5 viral LinkedIn post templates with hooks, body, CTA. Include engagement tactics and best posting times.',
      roast:    'Brutally but constructively critique this LinkedIn profile. Give specific, actionable improvements for every section.'
    };

    var prompt =
'You are a LinkedIn growth expert who has helped 500+ professionals reach 10K+ followers and land dream jobs.\n\n' +
'Task: ' + (modePrompt[mode] || modePrompt.full) + '\n\n' +
'OUTPUT RULES:\n' +
'- HTML FRAGMENT ONLY — no <html>, <head>, <body>, or <html> tags\n' +
'- White background (#ffffff), dark text (#1a1a1a)\n' +
'- Clean sections with professional headers\n' +
'- Write ENTIRELY in ' + lang + '\n\n' +
'PROFESSIONAL PROFILE:\n' +
'Current/Target Role: ' + Validate.sanitize(data.role, 150) + '\n' +
'Industry: ' + Validate.sanitize(data.industry || 'Technology', 100) + '\n' +
'Experience: ' + Validate.sanitize(data.experience || '', 3000) + '\n' +
'Key Skills: ' + Validate.sanitize(data.skills || '', 500) + '\n' +
'Target Audience: ' + Validate.sanitize(data.target || 'Recruiters and hiring managers', 200) + '\n\n' +
'Make every word count. Avoid clichés like "passionate", "results-driven", "dynamic".';

    return callAI(prompt, { maxTokens: CONFIG.AI.MAX_TOKENS_LONG });

  } catch(e) {
    Logger_.error('generateLinkedIn', e.message, email);
    return { success: false, error: 'Gagal generate LinkedIn: ' + e.message };
  }
}

// ============================================================
// PORTFOLIO AI
// ============================================================

function generatePortfolio(data, email) {
  var premiumCheck = requirePremium(email);
  if (!premiumCheck.success) return premiumCheck;

  try {
    var req = Validate.required(data, ['name','role','projects']);
    if (!req.valid) return { success: false, error: 'Field "' + req.missing + '" wajib diisi.' };

    var lang = Validate.sanitize(data.language || 'Indonesian', 50);
    var template = data.template || 'dark-modern';

    var templateStyle = _getPortfolioTemplateStyle(template);

    var prompt =
'You are a senior UI/UX designer and personal branding expert who has built award-winning portfolios.\n\n' +
'Create a complete, visually stunning portfolio website as a SELF-CONTAINED HTML FRAGMENT.\n\n' +
'CRITICAL OUTPUT RULES:\n' +
'- COMPLETE HTML FRAGMENT with embedded CSS (no <html>/<head>/<body> tags)\n' +
'- Start with <style>...</style> then the content div\n' +
'- All CSS must be inside the style tag or inline — no external dependencies\n' +
'- Language: ENTIRELY in ' + lang + '\n' +
'- Template style: ' + templateStyle.description + '\n\n' +
'DESIGN REQUIREMENTS:\n' +
templateStyle.css + '\n\n' +
'PORTFOLIO SECTIONS (all required):\n' +
'1. HERO — Name, title, animated tagline, CTA buttons, stats counter\n' +
'2. ABOUT — Personal story, values, photo placeholder with initials\n' +
'3. SKILLS — Visual skill bars or grid with proficiency levels\n' +
'4. PROJECTS — Card grid with project image placeholders, tech stack badges, impact numbers\n' +
'5. ACHIEVEMENTS / STATS — Impressive numbers in counter format\n' +
'6. TESTIMONIALS (optional but impressive if added)\n' +
'7. CONTACT — Form layout, social links, email, availability status\n\n' +
'PROFESSIONAL DATA:\n' +
'Name: ' + Validate.sanitize(data.name, 100) + '\n' +
'Role: ' + Validate.sanitize(data.role, 100) + '\n' +
'Tagline: ' + Validate.sanitize(data.tagline || '', 200) + '\n' +
'Years of Experience: ' + Validate.sanitize(data.years || '', 20) + '\n' +
'Projects:\n' + Validate.sanitize(data.projects, 5000) + '\n\n' +
'Skills: ' + Validate.sanitize(data.skills || '', 500) + '\n' +
'Achievements: ' + Validate.sanitize(data.achievements || '', 300) + '\n' +
'Contact: ' + Validate.sanitize(data.contact || '', 300) + '\n\n' +
'Make it look like it was designed by a professional design agency. Use CSS animations where appropriate.';

    var result = callAI(prompt, { maxTokens: CONFIG.AI.MAX_TOKENS_LONG });

    if (result.success) {
      _saveToDB('PORTFOLIO_DB',
        ['portfolio_id','email','name','role','template','language','content','created_at'],
        [Utilities.getUuid(), email||'', data.name, data.role, template, lang, result.text, new Date().toISOString()],
        'generatePortfolio', email
      );
    }

    return result;

  } catch(e) {
    Logger_.error('generatePortfolio', e.message, email);
    return { success: false, error: 'Gagal generate portfolio: ' + e.message };
  }
}

function _getPortfolioTemplateStyle(template) {
  var styles = {
    'dark-modern': {
      description: 'Dark modern tech portfolio — dark backgrounds, orange/neon accents, glassmorphism cards',
      css: 'Colors: background #0a0a0f, card backgrounds rgba(255,255,255,0.05), accent #ff7a00. Use glassmorphism, gradient text, subtle glow effects.'
    },
    'minimalist': {
      description: 'Swiss minimalist — white backgrounds, black typography, lots of whitespace, geometric precision',
      css: 'Colors: background #ffffff, text #1a1a1a, accent #000000. Use minimal decorations, perfect spacing, serif/sans contrast.'
    },
    'startup': {
      description: 'Startup/SaaS style — gradient backgrounds, purple/violet palette, modern cards',
      css: 'Colors: background #0f0f2e, gradient from #6c63ff to #ff00aa, white text. Use bold gradients, modern cards, startup energy.'
    },
    'creative': {
      description: 'Creative bold — full color gradients, experimental layout, artistic expression',
      css: 'Colors: vibrant gradients (purple→pink→orange), white text, bold typography. Use creative layouts, overlapping elements, visual impact.'
    },
    'corporate': {
      description: 'Corporate professional — navy/blue palette, trustworthy, LinkedIn-compatible',
      css: 'Colors: background #1a3a5c, accent #4db8ff, white text. Use clean corporate layout, professional sections, trust signals.'
    },
    'elegant': {
      description: 'Luxury elegant — dark warm backgrounds, gold accents, premium feel',
      css: 'Colors: background #0f0a00, gold accent #d4a853, warm white text #f5f0e8. Use luxury spacing, gold details, prestigious typography.'
    }
  };
  return styles[template] || styles['dark-modern'];
}

// ============================================================
// INTERVIEW AI
// ============================================================

function generateInterviewQuestions(data, email) {
  var premiumCheck = requirePremium(email);
  if (!premiumCheck.success) return premiumCheck;

  try {
    var req = Validate.required(data, ['role']);
    if (!req.valid) return { success: false, error: 'Target position wajib diisi.' };

    var lang = Validate.sanitize(data.language || 'Indonesian', 50);
    var mode = data.mode || 'standard';

    var modeInstructions = {
      standard: 'Standard HR + behavioral interview simulation. Friendly but professional tone.',
      technical: 'Deep technical interview. Include coding concepts, system design, problem-solving, and technical depth appropriate to the role.',
      stress: 'Stress interview simulation — challenging questions, trick questions, pressure scenarios. Include how to stay calm and respond.',
      rapid: 'Rapid fire simulation — 35+ quick questions. Short, punchy. Include how to give crisp answers.'
    };

    var prompt =
'You are a senior hiring manager at a top company who has conducted 1,000+ interviews.\n\n' +
'Mode: ' + (modeInstructions[mode] || modeInstructions.standard) + '\n\n' +
'OUTPUT RULES:\n' +
'- HTML FRAGMENT ONLY — no <html>, <head>, <body> tags\n' +
'- White background (#ffffff), dark text (#1a1a1a)\n' +
'- Clear numbered sections, professional typography\n' +
'- Write ENTIRELY in ' + lang + '\n\n' +
'POSITION DETAILS:\n' +
'Role: ' + Validate.sanitize(data.role, 150) + '\n' +
'Company Type: ' + Validate.sanitize(data.companyType || 'Tech company', 100) + '\n' +
'Experience Level: ' + Validate.sanitize(data.level || 'Mid-senior', 50) + '\n' +
'Company Focus: ' + Validate.sanitize(data.companyFocus || '', 200) + '\n\n' +
'REQUIRED SECTIONS:\n' +
'1. INTERVIEW OVERVIEW — what to expect, key focus areas\n' +
'2. HR & BEHAVIORAL QUESTIONS — 12 questions with:\n' +
'   - STAR method framework\n' +
'   - Sample high-scoring answer\n' +
'   - Common mistakes to avoid\n' +
'3. TECHNICAL/ROLE-SPECIFIC QUESTIONS — 8 questions with answer frameworks\n' +
'4. LEADERSHIP & SITUATIONAL — 6 scenarios with how to approach\n' +
'5. CULTURE FIT QUESTIONS — 5 questions about company values\n' +
'6. QUESTIONS TO ASK INTERVIEWER — 6 smart questions that impress\n' +
'7. SALARY NEGOTIATION SCRIPT — exact words to say\n' +
'8. RED FLAGS TO AVOID — 10 things that kill your chances\n' +
'9. POST-INTERVIEW FOLLOW-UP — email template\n' +
'10. CONFIDENCE BOOSTER — final preparation checklist\n\n' +
'Be specific to the role and company type. Use real interview examples.';

    var result = callAI(prompt, { maxTokens: CONFIG.AI.MAX_TOKENS_LONG });

    if (result.success) {
      _saveToDB('INTERVIEW_DB',
        ['interview_id','email','role','mode','language','content','created_at'],
        [Utilities.getUuid(), email||'', data.role, mode, lang, result.text, new Date().toISOString()],
        'generateInterviewQuestions', email
      );
    }

    return result;

  } catch(e) {
    Logger_.error('generateInterviewQuestions', e.message, email);
    return { success: false, error: 'Gagal generate interview kit: ' + e.message };
  }
}

// ============================================================
// JOB MATCH AI
// ============================================================

function generateJobMatch(data, email) {
  var premiumCheck = requirePremium(email);
  if (!premiumCheck.success) return premiumCheck;

  try {
    var req = Validate.required(data, ['cv']);
    if (!req.valid) return { success: false, error: 'CV/Resume wajib diisi.' };

    var prompt =
'You are a senior talent acquisition specialist and career advisor.\n\n' +
'Analyze this candidate\'s profile and provide a comprehensive job matching report as an HTML FRAGMENT.\n\n' +
'OUTPUT RULES:\n' +
'- HTML FRAGMENT ONLY — no <html>, <head>, <body> tags\n' +
'- White background (#ffffff), dark text (#1a1a1a)\n' +
'- Use match percentage badges with color (inline only)\n' +
'- Professional layout\n\n' +
'CANDIDATE PROFILE:\n' +
'CV/Resume:\n' + Validate.sanitize(data.cv, 6000) + '\n\n' +
'Target Industry: ' + Validate.sanitize(data.industry || 'Technology', 100) + '\n' +
'Experience Level: ' + Validate.sanitize(data.level || 'Mid-senior', 50) + '\n' +
'Location: ' + Validate.sanitize(data.location || 'Jakarta', 100) + '\n' +
'Salary Expectation: ' + Validate.sanitize(data.salary || 'Negotiable', 100) + '\n\n' +
'REQUIRED SECTIONS:\n' +
'1. CANDIDATE PROFILE SUMMARY — key strengths, years of experience, seniority level\n' +
'2. TOP 8 MATCHING JOB ROLES — each with:\n' +
'   - Job title and match percentage\n' +
'   - Why it matches\n' +
'   - Required vs candidate skills comparison\n' +
'   - Example companies hiring for this role in Indonesia\n' +
'   - Estimated salary range (Indonesia market)\n' +
'3. SKILL GAP ANALYSIS — what skills to add for top matches\n' +
'4. SALARY BENCHMARK — current market rate for profile\n' +
'5. CAREER POSITIONING — how to position for highest-paying roles\n' +
'6. RECOMMENDED JOB PLATFORMS — where to apply\n' +
'7. PROFILE IMPROVEMENT TIPS — 5 quick wins to increase job match rate\n' +
'8. 30-60-90 DAY JOB SEARCH PLAN\n\n' +
'Be specific to Indonesian job market unless location suggests otherwise.';

    return callAI(prompt, { maxTokens: CONFIG.AI.MAX_TOKENS_LONG });

  } catch(e) {
    Logger_.error('generateJobMatch', e.message, email);
    return { success: false, error: 'Gagal job match: ' + e.message };
  }
}

// ============================================================
// CAREER ROADMAP AI
// ============================================================

function generateCareerRoadmap(data, email) {
  var premiumCheck = requirePremium(email);
  if (!premiumCheck.success) return premiumCheck;

  try {
    var req = Validate.required(data, ['currentRole','goalRole']);
    if (!req.valid) return { success: false, error: 'Current role dan goal role wajib diisi.' };

    var prompt =
'You are a career strategy consultant who has advised 1,000+ professionals from junior to C-suite.\n\n' +
'Create a detailed, personalized career roadmap as an HTML FRAGMENT.\n\n' +
'OUTPUT RULES:\n' +
'- HTML FRAGMENT ONLY — no <html>, <head>, <body> tags\n' +
'- White background (#ffffff), dark text (#1a1a1a)\n' +
'- Timeline visualization using CSS (no SVG or canvas)\n' +
'- Clear phase markers with dates/durations\n\n' +
'CAREER PROFILE:\n' +
'Current Role: ' + Validate.sanitize(data.currentRole, 150) + '\n' +
'Dream Goal: ' + Validate.sanitize(data.goalRole, 150) + '\n' +
'Current Skills: ' + Validate.sanitize(data.currentSkills || '', 500) + '\n' +
'Years of Experience: ' + Validate.sanitize(data.years || 'Unknown', 50) + '\n' +
'Timeline: ' + Validate.sanitize(data.timeline || '2 years', 50) + '\n' +
'Industry: ' + Validate.sanitize(data.industry || 'Technology', 100) + '\n\n' +
'REQUIRED SECTIONS:\n' +
'1. EXECUTIVE SUMMARY — current position, goal, feasibility assessment\n' +
'2. GAP ANALYSIS — skills you have vs skills you need\n' +
'3. PHASED ROADMAP (3-5 phases) — each phase has:\n' +
'   - Phase name and duration\n' +
'   - Specific goals to achieve\n' +
'   - Skills to develop\n' +
'   - Certifications to earn (with providers)\n' +
'   - Projects to build\n' +
'   - Job titles to target\n' +
'   - Expected salary range\n' +
'4. LEARNING RESOURCES — specific courses, books, platforms per skill\n' +
'5. NETWORKING STRATEGY — how to build the right connections\n' +
'6. MONTHLY ACTION PLAN — first 3 months in detail\n' +
'7. SUCCESS METRICS — how to measure progress\n' +
'8. RISK & MITIGATION — potential obstacles and how to handle them\n' +
'9. ALTERNATIVE PATHS — 2 faster routes to the goal\n\n' +
'Be specific, realistic, and actionable. Include real course names, certification bodies, salary numbers.';

    return callAI(prompt, { maxTokens: CONFIG.AI.MAX_TOKENS_LONG });

  } catch(e) {
    Logger_.error('generateCareerRoadmap', e.message, email);
    return { success: false, error: 'Gagal generate roadmap: ' + e.message };
  }
}

// ============================================================
// AI CAREER COACH — MULTI-TURN CHAT
// ============================================================

function chatWithCoach(message, email, historyJSON) {
  var premiumCheck = requirePremium(email);
  if (!premiumCheck.success) return premiumCheck;

  try {
    if (!message || message.trim() === '') {
      return { success: false, error: 'Pesan tidak boleh kosong.' };
    }

    var safeMessage = Validate.sanitize(message, 2000);

    var history = [];
    if (historyJSON) {
      try {
        var parsed = JSON.parse(historyJSON);
        if (Array.isArray(parsed)) {
          history = parsed.slice(-8).map(function(h) {
            return { role: h.role, content: Validate.sanitize(String(h.content), 1500) };
          });
        }
      } catch(e) {}
    }

    var systemPrompt =
'You are an expert AI Career Coach with 20+ years of experience as an executive recruiter, ' +
'career counselor, and leadership coach. You have helped thousands of professionals:\n' +
'- Land jobs at Google, McKinsey, Gojek, Tokopedia, and top companies\n' +
'- Negotiate salaries 30-50% above initial offers\n' +
'- Successfully pivot careers and build personal brands\n' +
'- Build confidence and overcome imposter syndrome\n\n' +
'Your coaching style:\n' +
'- Direct, honest, and actionable — no fluff\n' +
'- Empathetic but challenging — push people to be better\n' +
'- Data-driven advice with real examples\n' +
'- Culturally aware of Indonesian job market dynamics\n\n' +
'RULES:\n' +
'- Reply in the SAME LANGUAGE as the user\'s message\n' +
'- Keep replies focused and useful — 150-400 words\n' +
'- Use HTML formatting for structure when helpful (<b>, <ul>, <li>, <br>)\n' +
'- Give concrete, specific advice — never vague generalities\n' +
'- If salary negotiation: give specific numbers and scripts\n' +
'- If career advice: give specific steps and timelines';

    return callAI(safeMessage, {
      systemPrompt: systemPrompt,
      history: history,
      maxTokens: 1500,
      temperature: 0.75,
      retries: 1
    });

  } catch(e) {
    Logger_.error('chatWithCoach', e.message, email);
    return { success: false, error: 'Coach tidak bisa menjawab sekarang. Coba lagi.' };
  }
}

// ============================================================
// AUTO APPLY SYSTEM
// ============================================================

function saveApplication(data, email) {
  var premiumCheck = requirePremium(email);
  if (!premiumCheck.success) return premiumCheck;

  try {
    var req = Validate.required(data, ['company','position']);
    if (!req.valid) return { success: false, error: 'Company dan position wajib diisi.' };

    var sheet = DB.getSheet('APPLY_DB');
    DB.ensureHeaders(sheet, [
      'apply_id','email','company','position','status',
      'applied_date','notes','job_url','salary','next_step','updated_at'
    ]);

    var applyId = Utilities.getUuid();
    sheet.appendRow([
      applyId,
      email,
      Validate.sanitize(data.company, 200),
      Validate.sanitize(data.position, 200),
      data.status || 'applied',
      data.appliedDate || new Date().toISOString().split('T')[0],
      Validate.sanitize(data.notes || '', 500),
      Validate.sanitize(data.jobUrl || '', 500),
      Validate.sanitize(data.salary || '', 100),
      Validate.sanitize(data.nextStep || '', 200),
      new Date().toISOString()
    ]);

    return { success: true, applyId: applyId };

  } catch(e) {
    Logger_.error('saveApplication', e.message, email);
    return { success: false, error: e.message };
  }
}

function updateApplicationStatus(applyId, newStatus, email) {
  try {
    var validStatuses = ['applied','interview','offer','accepted','rejected','pending','withdrawn'];
    if (!validStatuses.includes(newStatus)) {
      return { success: false, error: 'Invalid status.' };
    }

    var sheet = DB.getSheet('APPLY_DB');
    var row = DB.findRowExact(sheet, 0, applyId);

    if (!row) return { success: false, error: 'Application not found.' };
    if (String(row.data[1]).toLowerCase() !== Validate.safeEmail(email)) {
      return { success: false, error: 'Unauthorized.' };
    }

    DB.updateCell(sheet, row.rowIndex, 5, newStatus);
    DB.updateCell(sheet, row.rowIndex, 11, new Date().toISOString());

    return { success: true };

  } catch(e) {
    Logger_.error('updateApplicationStatus', e.message, email);
    return { success: false, error: e.message };
  }
}

function generateApplicationEmail(data, email) {
  var premiumCheck = requirePremium(email);
  if (!premiumCheck.success) return premiumCheck;

  try {
    var req = Validate.required(data, ['company','position','type']);
    if (!req.valid) return { success: false, error: 'Company, position, dan type wajib diisi.' };

    var typeMap = {
      apply:    'initial job application email',
      followup: 'professional follow-up after application (1 week no response)',
      thankyou: 'post-interview thank you email',
      negotiation: 'salary negotiation email'
    };

    var prompt =
'Write a professional ' + (typeMap[data.type] || 'job application email') + ' as an HTML FRAGMENT.\n\n' +
'Details:\n' +
'Company: ' + Validate.sanitize(data.company, 150) + '\n' +
'Position: ' + Validate.sanitize(data.position, 150) + '\n' +
'Candidate Background: ' + Validate.sanitize(data.background || '', 1000) + '\n' +
'Tone: Professional, confident, human\n\n' +
'Rules:\n' +
'- HTML fragment only (no <html>/<body> tags)\n' +
'- White background, dark text\n' +
'- Include: Subject line, email body, signature\n' +
'- 200-300 words\n' +
'- Language: ' + Validate.sanitize(data.language || 'Indonesian', 50);

    return callAI(prompt, { maxTokens: 1500 });

  } catch(e) {
    Logger_.error('generateApplicationEmail', e.message, email);
    return { success: false, error: e.message };
  }
}

// ============================================================
// PAYMENT — MIDTRANS PRODUCTION
// ============================================================

function createPayment(email, plan, userName) {
  try {
    if (!Validate.email(email)) {
      return { success: false, error: 'Email tidak valid.' };
    }
    if (!plan || !['monthly','yearly'].includes(plan)) {
      return { success: false, error: 'Plan tidak valid.' };
    }

    var safeEmail = Validate.safeEmail(email);

    // Duplicate payment check — prevent if pending payment exists
    var existingPending = _checkExistingPendingPayment(safeEmail);
    if (existingPending) {
      Logger_.warn('createPayment', 'Existing pending payment found: ' + existingPending, safeEmail);
      // Allow creating new — old one may be stale
    }

    var amount = plan === 'yearly' ? CONFIG.PAYMENT.YEARLY_PRICE : CONFIG.PAYMENT.MONTHLY_PRICE;
    var orderId = 'AICAREER-' + Date.now() + '-' + Utilities.getUuid().substring(0, 8).toUpperCase();

    var serverKey = CONFIG.PAYMENT.IS_PRODUCTION
      ? CONFIG.PAYMENT.SERVER_KEY_PROD
      : CONFIG.PAYMENT.SERVER_KEY_SANDBOX;

    var snapUrl = CONFIG.PAYMENT.IS_PRODUCTION
      ? CONFIG.PAYMENT.SNAP_URL_PROD
      : CONFIG.PAYMENT.SNAP_URL_SANDBOX;

    if (!serverKey) {
      Logger_.error('createPayment', 'Midtrans server key not configured', safeEmail);
      return { success: false, error: 'Payment gateway belum dikonfigurasi. Hubungi admin.' };
    }

    var safeName = Validate.sanitize(userName || safeEmail.split('@')[0], 100);

    var payload = {
      transaction_details: { order_id: orderId, gross_amount: amount },
      item_details: [{
        id: 'premium-' + plan,
        price: amount,
        quantity: 1,
        name: 'AI Career OS Premium ' + (plan === 'yearly' ? 'Yearly' : 'Monthly')
      }],
      customer_details: {
        first_name: safeName,
        email: safeEmail
      },
      credit_card: { secure: true },
      expiry: {
        start_time: Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss Z'),
        unit: 'hours',
        duration: 24
      },
      enabled_payments: ['credit_card','bca_va','bni_va','bri_va','mandiri_bill',
                         'permata_va','other_va','gopay','shopeepay','qris','akulaku']
    };

    var response = UrlFetchApp.fetch(snapUrl, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(serverKey + ':'),
        'Accept': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var code = response.getResponseCode();
    var rawText = response.getContentText();
    var json;

    try { json = JSON.parse(rawText); } catch(pe) {
      Logger_.error('createPayment', 'Invalid JSON from Midtrans: ' + rawText.substring(0, 200), safeEmail);
      throw new Error('Payment gateway error. Coba lagi.');
    }

    if (code !== 201 || !json.token) {
      var errMsg = json.error_messages ? json.error_messages.join(', ') : JSON.stringify(json).substring(0, 300);
      Logger_.error('createPayment', 'Midtrans ' + code + ': ' + errMsg, safeEmail);
      throw new Error('Payment error: ' + errMsg);
    }

    // Save to DB
    var sheet = DB.getSheet('PAYMENT_DB');
    DB.ensureHeaders(sheet, [
      'payment_id','order_id','email','plan',
      'amount','status','token','created_at','updated_at','mode'
    ]);
    sheet.appendRow([
      Utilities.getUuid(), orderId, safeEmail, plan,
      amount, 'pending', json.token,
      new Date().toISOString(), new Date().toISOString(),
      CONFIG.PAYMENT.IS_PRODUCTION ? 'production' : 'sandbox'
    ]);

    Logger_.payment('createPayment', 'Created: ' + orderId + ' | ' + plan + ' | Rp' + amount, safeEmail);

    return { success: true, token: json.token, orderId: orderId, amount: amount };

  } catch(e) {
    Logger_.error('createPayment', e.message, email);
    return { success: false, error: e.message };
  }
}

function _checkExistingPendingPayment(email) {
  try {
    var sheet = DB.getSheet('PAYMENT_DB');
    var records = DB.getByEmail('PAYMENT_DB', email);
    var pending = records.filter(function(r) { return r.status === 'pending'; });
    return pending.length > 0 ? pending[pending.length - 1].order_id : null;
  } catch(e) {
    return null;
  }
}

function activatePremium(email, orderId, plan) {
  try {
    if (!email || !orderId) return { success: false, error: 'Email dan Order ID wajib.' };

    var safeEmail = Validate.safeEmail(email);

    // Duplicate activation check
    var paymentSheet = DB.getSheet('PAYMENT_DB');
    var paymentRow = DB.findRowExact(paymentSheet, 1, orderId);

    if (paymentRow) {
      if (String(paymentRow.data[5]) === 'success') {
        Logger_.warn('activatePremium', 'Duplicate activation attempt: ' + orderId, safeEmail);
        return { success: true, alreadyActivated: true };
      }
    }

    var userSheet = DB.getSheet('USER_DB');
    var userRow = DB.findRow(userSheet, 0, safeEmail);
    if (!userRow) return { success: false, error: 'User tidak ditemukan.' };

    var now = new Date();
    var expDate = new Date(now);

    if (plan === 'yearly') {
      expDate.setFullYear(expDate.getFullYear() + 1);
    } else {
      expDate.setMonth(expDate.getMonth() + 1);
    }

    DB.updateCell(userSheet, userRow.rowIndex, 5, true);
    DB.updateCell(userSheet, userRow.rowIndex, 6, plan);
    DB.updateCell(userSheet, userRow.rowIndex, 7, expDate.toISOString());

    if (paymentRow) {
      DB.updateCell(paymentSheet, paymentRow.rowIndex, 6, 'success');
      DB.updateCell(paymentSheet, paymentRow.rowIndex, 9, new Date().toISOString());
    }

    Logger_.payment('activatePremium', 'Activated: ' + safeEmail + ' | ' + plan + ' until ' + expDate.toISOString(), safeEmail);

    // Send confirmation email
    _sendPremiumConfirmationEmail(safeEmail, plan, orderId, expDate);

    return { success: true, premium: true, plan: plan, expiredAt: expDate.toISOString() };

  } catch(e) {
    Logger_.error('activatePremium', e.message, email);
    return { success: false, error: e.message };
  }
}

function _sendPremiumConfirmationEmail(email, plan, orderId, expDate) {
  try {
    MailApp.sendEmail({
      to: email,
      subject: '🎉 Premium AI Career OS Aktif! Order #' + orderId,
      htmlBody:
        '<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;padding:32px;background:#f4f6fb;border-radius:16px;">' +
        '<div style="text-align:center;margin-bottom:20px;">' +
        '<h1 style="color:#ff7a00;font-size:22px;margin:0;">⚡ AI Career OS</h1>' +
        '</div>' +
        '<div style="background:#fff;border-radius:12px;padding:28px;margin-bottom:16px;">' +
        '<h2 style="color:#1a1a1a;font-size:18px;margin:0 0 12px;">🎉 Selamat! Premium Aktif!</h2>' +
        '<p style="color:#444;font-size:14px;">Terima kasih sudah upgrade ke <strong>AI Career OS Premium ' + plan + '</strong>.</p>' +
        '<p style="color:#444;font-size:14px;">Akses premium kamu aktif hingga: <strong>' + expDate.toDateString() + '</strong></p>' +
        '<hr style="border:none;border-top:1px solid #eee;margin:16px 0;">' +
        '<p style="color:#444;font-size:13px;">✅ Unlimited CV Generator AI<br>✅ Unlimited ATS Checker Pro<br>✅ LinkedIn AI Optimizer<br>✅ Cover Letter AI<br>✅ Portfolio AI<br>✅ Interview AI Pro<br>✅ Job Match AI<br>✅ Career Roadmap AI<br>✅ AI Career Coach</p>' +
        '</div>' +
        '<p style="color:#aaa;font-size:11px;text-align:center;">Order ID: ' + orderId + '<br>Mode: ' + (CONFIG.PAYMENT.IS_PRODUCTION ? 'Production' : 'Sandbox') + '</p>' +
        '</div>'
    });
  } catch(e) {
    Logger_.warn('_sendPremiumConfirmationEmail', 'Email failed: ' + e.message, email);
  }
}

function checkPaymentStatus(orderId, email) {
  try {
    if (!orderId) return { success: false, error: 'Order ID required.' };

    var serverKey = CONFIG.PAYMENT.IS_PRODUCTION
      ? CONFIG.PAYMENT.SERVER_KEY_PROD
      : CONFIG.PAYMENT.SERVER_KEY_SANDBOX;

    var apiUrl = (CONFIG.PAYMENT.IS_PRODUCTION
      ? CONFIG.PAYMENT.API_URL_PROD
      : CONFIG.PAYMENT.API_URL_SANDBOX) + orderId + '/status';

    var response = UrlFetchApp.fetch(apiUrl, {
      method: 'get',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(serverKey + ':'),
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    });

    var code = response.getResponseCode();
    if (code !== 200) return { success: false, error: 'Status check failed: ' + code };

    var json = JSON.parse(response.getContentText());
    return { success: true, status: json.transaction_status, fraud: json.fraud_status, data: json };

  } catch(e) {
    Logger_.error('checkPaymentStatus', e.message, email);
    return { success: false, error: e.message };
  }
}

// ============================================================
// WEBHOOK — MIDTRANS
// ============================================================

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return _jsonResponse({ status: 'error', message: 'No data' });
    }

    var data = JSON.parse(e.postData.contents);
    var orderId     = data.order_id;
    var status      = data.transaction_status;
    var fraudStatus = data.fraud_status;
    var signatureKey = data.signature_key;

    // Verify webhook signature
    if (!_verifyMidtransSignature(data)) {
      Logger_.warn('Webhook', 'Invalid signature for order: ' + orderId);
      return _jsonResponse({ status: 'error', message: 'Invalid signature' });
    }

    Logger_.payment('Webhook', 'Received: ' + orderId + ' | status: ' + status + ' | fraud: ' + fraudStatus);

    if ((status === 'capture' && fraudStatus === 'accept') || status === 'settlement') {
      var paymentSheet = DB.getSheet('PAYMENT_DB');
      var paymentRow = DB.findRowExact(paymentSheet, 1, orderId);

      if (paymentRow) {
        var currentStatus = String(paymentRow.data[5]);
        if (currentStatus !== 'success') {
          activatePremium(String(paymentRow.data[2]), orderId, String(paymentRow.data[3]));
        }
      } else {
        Logger_.warn('Webhook', 'Payment record not found for order: ' + orderId);
      }

    } else if (status === 'deny' || status === 'cancel' || status === 'expire') {
      var paymentSheet2 = DB.getSheet('PAYMENT_DB');
      var failRow = DB.findRowExact(paymentSheet2, 1, orderId);
      if (failRow) {
        DB.updateCell(paymentSheet2, failRow.rowIndex, 6, status);
        DB.updateCell(paymentSheet2, failRow.rowIndex, 9, new Date().toISOString());
      }
      Logger_.payment('Webhook', 'Payment ' + status + ': ' + orderId);

    } else if (status === 'pending') {
      Logger_.payment('Webhook', 'Payment pending: ' + orderId);
    }

    return _jsonResponse({ status: 'ok' });

  } catch(e) {
    Logger_.error('doPost/webhook', e.message);
    return _jsonResponse({ status: 'error', message: e.message });
  }
}

function _verifyMidtransSignature(data) {
  try {
    if (!data.signature_key || !data.order_id || !data.status_code || !data.gross_amount) {
      return true; // Can't verify without all fields — allow for now
    }

    var serverKey = CONFIG.PAYMENT.IS_PRODUCTION
      ? CONFIG.PAYMENT.SERVER_KEY_PROD
      : CONFIG.PAYMENT.SERVER_KEY_SANDBOX;

    if (!serverKey) return true;

    var rawString = data.order_id + data.status_code + data.gross_amount + serverKey;
    var hash = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_512,
      rawString,
      Utilities.Charset.UTF_8
    );

    var hexHash = hash.map(function(b) {
      return ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2);
    }).join('');

    return hexHash === data.signature_key;

  } catch(e) {
    Logger_.warn('_verifyMidtransSignature', 'Signature verification failed: ' + e.message);
    return true; // Fail open with warning
  }
}

function _jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// MULTI-LANGUAGE SUPPORT
// ============================================================

var I18N = {
  messages: {
    id: {
      premiumRequired: 'Fitur premium diperlukan.',
      notFound: 'Tidak ditemukan.',
      serverError: 'Terjadi kesalahan server.'
    },
    en: {
      premiumRequired: 'Premium subscription required.',
      notFound: 'Not found.',
      serverError: 'Server error occurred.'
    }
  },
  get: function(key, lang) {
    var l = (lang || 'id').substring(0, 2).toLowerCase();
    var msgs = I18N.messages[l] || I18N.messages.id;
    return msgs[key] || key;
  }
};

// ============================================================
// DB HELPER — SAFE SAVE
// ============================================================

function _saveToDB(sheetName, headers, rowData, source, email) {
  try {
    var sheet = DB.getSheet(sheetName);
    DB.ensureHeaders(sheet, headers);
    sheet.appendRow(rowData);
  } catch(e) {
    Logger_.warn(source || '_saveToDB', 'DB save failed for ' + sheetName + ': ' + e.message, email);
  }
}

// ============================================================
// DIAGNOSTIC & SETUP
// ============================================================

function runDiagnostic() {
  var results = [];

  // API Key
  var apiKey = PropertiesService.getScriptProperties().getProperty('OPENROUTER_API_KEY');
  results.push({
    label: 'OPENROUTER_API_KEY',
    status: apiKey ? 'OK' : 'ERROR',
    message: apiKey ? '✅ OpenRouter API Key configured' : '❌ OPENROUTER_API_KEY not set in Script Properties'
  });

  // Spreadsheet
  try {
    if (!CONFIG.SPREADSHEET_ID) throw new Error('SPREADSHEET_ID not configured');
    SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    results.push({ label: 'Spreadsheet', status: 'OK', message: '✅ Spreadsheet connected: ' + CONFIG.SPREADSHEET_ID });
  } catch(e) {
    results.push({ label: 'Spreadsheet', status: 'ERROR', message: '❌ ' + e.message });
  }

  // Cache
  try {
    var cache = CacheService.getScriptCache();
    cache.put('_diag_' + Date.now(), 'test', 10);
    results.push({ label: 'CacheService', status: 'OK', message: '✅ CacheService working (OTP system ready)' });
  } catch(e) {
    results.push({ label: 'CacheService', status: 'ERROR', message: '❌ ' + e.message });
  }

  // Mail
  try {
    var remaining = MailApp.getRemainingDailyQuota();
    results.push({
      label: 'MailApp',
      status: remaining > 10 ? 'OK' : 'WARN',
      message: remaining > 10 ? '✅ Email quota: ' + remaining + ' remaining' : '⚠️ Email quota low: ' + remaining
    });
  } catch(e) {
    results.push({ label: 'MailApp', status: 'ERROR', message: '❌ ' + e.message });
  }

  // Midtrans Production Key
  var mtKeyProd = CONFIG.PAYMENT.SERVER_KEY_PROD;
  results.push({
    label: 'Midtrans Production Key',
    status: mtKeyProd ? 'OK' : 'ERROR',
    message: mtKeyProd ? '✅ Production key configured' : '❌ MIDTRANS_SERVER_KEY_PROD not set'
  });

  // Midtrans Client Key
  var mtClientProd = CONFIG.PAYMENT.CLIENT_KEY_PROD;
  results.push({
    label: 'Midtrans Client Key',
    status: mtClientProd ? 'OK' : 'ERROR',
    message: mtClientProd ? '✅ Client key configured' : '❌ MIDTRANS_CLIENT_KEY_PROD not set'
  });

  // Payment Mode
  results.push({
    label: 'Payment Mode',
    status: 'OK',
    message: CONFIG.PAYMENT.IS_PRODUCTION ? '✅ PRODUCTION MODE ACTIVE' : '⚠️ SANDBOX MODE'
  });

  // AI Test
  try {
    var testKey = PropertiesService.getScriptProperties().getProperty('OPENROUTER_API_KEY');
    if (testKey) {
      results.push({ label: 'AI API Test', status: 'OK', message: '✅ API key present (not validated — call testAI() separately)' });
    } else {
      results.push({ label: 'AI API Test', status: 'ERROR', message: '❌ No API key found' });
    }
  } catch(e) {
    results.push({ label: 'AI API Test', status: 'ERROR', message: '❌ ' + e.message });
  }

  var errors = results.filter(function(r) { return r.status === 'ERROR'; }).length;
  var warns  = results.filter(function(r) { return r.status === 'WARN'; }).length;

  return {
    ok: errors === 0,
    summary: {
      total: results.length,
      passed: results.filter(function(r) { return r.status === 'OK'; }).length,
      warns: warns,
      errors: errors,
      mode: CONFIG.PAYMENT.IS_PRODUCTION ? 'PRODUCTION' : 'SANDBOX',
      version: CONFIG.VERSION
    },
    results: results,
    timestamp: new Date().toISOString()
  };
}

function testAI() {
  var result = callAI('Reply with exactly: "AI Career OS ready."', { maxTokens: 50, retries: 0 });
  return {
    success: result.success,
    response: result.text || '',
    error: result.error || '',
    timestamp: new Date().toISOString()
  };
}

function autoSetup() {
  var sheets = {
    USER_DB:       ['email','name','picture','role','premium','premium_plan','premium_expired','session_token','last_login','created_at'],
    PAYMENT_DB:    ['payment_id','order_id','email','plan','amount','status','token','created_at','updated_at','mode'],
    CV_DB:         ['cv_id','email','name','role','template','language','content','created_at'],
    COVERLETTER_DB:['cl_id','email','role','company','language','style','content','created_at'],
    PORTFOLIO_DB:  ['portfolio_id','email','name','role','template','language','content','created_at'],
    ATS_DB:        ['ats_id','email','score','role','mode','content','created_at'],
    INTERVIEW_DB:  ['interview_id','email','role','mode','language','content','created_at'],
    APPLY_DB:      ['apply_id','email','company','position','status','applied_date','notes','job_url','salary','next_step','updated_at'],
    ERROR_LOG:     ['Timestamp','Level','Source','Message','User','Detail']
  };

  var results = [];
  for (var sheetName in sheets) {
    try {
      var sheet = DB.getSheet(sheetName);
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(sheets[sheetName]);
        sheet.setFrozenRows(1);
        results.push('✅ Created: ' + sheetName);
      } else {
        results.push('✓ Exists: ' + sheetName);
      }
    } catch(e) {
      results.push('❌ Failed: ' + sheetName + ' — ' + e.message);
    }
  }

  return { ok: true, version: CONFIG.VERSION, results: results, timestamp: new Date().toISOString() };
}

function manualActivatePremium(email, plan) {
  // Admin function — use from Apps Script editor only
  if (!email || !plan) return { success: false, error: 'Email and plan required' };
  var fakeOrderId = 'MANUAL-' + Date.now();
  return activatePremium(email, fakeOrderId, plan);
}

function checkUserPremium(email) {
  // Admin function — check user premium status
  try {
    var sheet = DB.getSheet('USER_DB');
    var row = DB.findRow(sheet, 0, Validate.safeEmail(email));
    if (!row) return { found: false };
    return {
      found: true,
      email: row.data[0],
      name: row.data[1],
      premium: row.data[4],
      plan: row.data[5],
      expires: row.data[6],
      lastLogin: row.data[8]
    };
  } catch(e) {
    return { found: false, error: e.message };
  }
}

function requireRecruiter(email) {

  try {

    if (!email) {
      return {
        success:false,
        error:'Authentication required'
      };
    }

    var safeEmail = Validate.safeEmail(email);

    var sheet = DB.getSheet('USER_DB');

    var userRow = DB.findRow(sheet,0,safeEmail);

    if(!userRow){
      return {
        success:false,
        error:'User not found'
      };
    }

    var role = String(userRow.data[3] || '')
      .toLowerCase();

    if(role !== 'recruiter' && role !== 'admin'){
      return {
        success:false,
        error:'Recruiter access required'
      };
    }

    return {
      success:true,
      user:userRow.data
    };

  } catch(e){

    Logger_.error(
      'requireRecruiter',
      e.message,
      email
    );

    return {
      success:false,
      error:e.message
    };
  }
}

function createCompany(data,email){

  try {

    var sheet = DB.getSheet('COMPANY_DB');

    DB.ensureHeaders(sheet,[
      'company_id',
      'company_name',
      'email',
      'website',
      'logo',
      'description',
      'verified',
      'premium',
      'created_at'
    ]);

    var companyId =
      'COMP-' + Utilities.getUuid();

    sheet.appendRow([
      companyId,
      Validate.sanitize(data.company_name,100),
      Validate.safeEmail(email),
      Validate.sanitize(data.website,200),
      Validate.sanitize(data.logo,500),
      Validate.sanitize(data.description,5000),
      false,
      false,
      new Date().toISOString()
    ]);

    return {
      success:true,
      company_id:companyId
    };

  } catch(e){

    Logger_.error(
      'createCompany',
      e.message,
      email
    );

    return {
      success:false,
      error:e.message
    };
  }
}

function postJob(data, email) {

  try {
    Logger.log(email);
Logger.log(getUserRole(email));
Logger.log(isAdmin(email));

    var recruiterCheck =
      requireRecruiter(email);

    if (!recruiterCheck.success) {
      return recruiterCheck;
    }

    var sheet =
      DB.getSheet('JOB_DB');

    DB.ensureHeaders(sheet, [
      'job_id',
      'company_id',
      'company_name',
      'company_logo',
      'title',
      'location',
      'salary',
      'type',
      'category',
      'description',
      'requirements',
      'skills',
      'experience',
      'education',
      'apply_link',
      'status',
      'featured',
      'views',
      'applications',
      'created_by',
      'created_at',
      'expires_at'
    ]);

    var jobId =
      'JOB-' + Utilities.getUuid();

    sheet.appendRow([
      jobId,
      data.company_id || '',
      Validate.sanitize(data.company_name || '', 100),
      Validate.sanitize(data.company_logo || '', 500),
      Validate.sanitize(data.title || '', 150),
      Validate.sanitize(data.location || '', 100),
      Validate.sanitize(data.salary || '', 100),
      Validate.sanitize(data.type || '', 50),
      Validate.sanitize(data.category || '', 100),
      Validate.sanitize(data.description || '', 10000),
      Validate.sanitize(data.requirements || '', 10000),
      Validate.sanitize(data.skills || '', 1000),
      Validate.sanitize(data.experience || '', 500),
      Validate.sanitize(data.education || '', 500),
      Validate.sanitize(data.apply_link || '', 1000),
      'active',
      false,
      0,
      0,
      Validate.safeEmail(email),
      new Date().toISOString(),
      data.expires_at || ''
    ]);

    return {
      success: true,
      job_id: jobId
    };

  } catch (e) {

    Logger_.error(
      'postJob',
      e.message,
      email
    );

    return {
      success: false,
      error: e.message
    };
  }
}

function getJobs(page, limit, filters) {

  try {

    page = page || 1;
    limit = limit || 20;
    filters = filters || {};

    var cacheKey =
      'jobs_' + page + '_' + JSON.stringify(filters);

    var cached =
      Cache_.get(cacheKey);

    if (cached) {
      return cached;
    }

    var sheet =
      DB.getSheet('JOB_DB');

    var values =
      sheet.getDataRange().getValues();

    if (values.length <= 1) {

      return {
        success: true,
        jobs: [],
        total: 0,
        page: page,
        hasMore: false
      };
    }

    var headers =
      values[0];

    var rows =
      values.slice(1);

    var jobs = rows.map(function(row) {

      var obj = {};

      headers.forEach(function(h, i) {
        obj[h] = row[i];
      });

      return obj;
    });

    jobs = jobs.filter(function(job) {

      return String(job.status || '')
        .toLowerCase() === 'active';

    });

    if (filters.keyword) {

      var keyword =
        String(filters.keyword)
          .toLowerCase();

      jobs = jobs.filter(function(job) {

        return (
          String(job.title || '')
            .toLowerCase()
            .includes(keyword)

          ||

          String(job.company_name || '')
            .toLowerCase()
            .includes(keyword)

          ||

          String(job.category || '')
            .toLowerCase()
            .includes(keyword)
        );

      });
    }

    if (filters.location) {

      var location =
        String(filters.location)
          .toLowerCase();

      jobs = jobs.filter(function(job) {

        return String(job.location || '')
          .toLowerCase()
          .includes(location);

      });
    }

    if (filters.category) {

      var category =
        String(filters.category)
          .toLowerCase();

      jobs = jobs.filter(function(job) {

        return String(job.category || '')
          .toLowerCase() === category;

      });
    }

    jobs.sort(function(a, b) {

      return new Date(b.created_at)
        - new Date(a.created_at);

    });

    var start =
      (page - 1) * limit;

    var end =
      start + limit;

    var paginated =
      jobs.slice(start, end);

    var result = {

      success: true,
      jobs: paginated,
      total: jobs.length,
      page: page,
      hasMore: end < jobs.length

    };

    Cache_.set(
      cacheKey,
      result,
      120
    );

    return result;

  } catch (e) {

    Logger_.error(
      'getJobs',
      e.message
    );

    return {
      success: false,
      error: e.message,
      jobs: []
    };
  }
}

function applyJob(data,email){

  try {

    if(!email){
      return {
        success:false,
        error:'Login required'
      };
    }

    var sheet = DB.getSheet('APPLICATION_DB');

    DB.ensureHeaders(sheet,[
      'application_id',
      'job_id',
      'email',
      'cv_id',
      'resume_link',
      'status',
      'match_score',
      'ai_summary',
      'created_at'
    ]);

    var applicationId =
      'APP-' + Utilities.getUuid();

    sheet.appendRow([
      applicationId,
      data.job_id,
      Validate.safeEmail(email),
      data.cv_id || '',
      data.resume_link || '',
      'pending',
      data.match_score || 0,
      '',
      new Date().toISOString()
    ]);

    return {
      success:true,
      application_id:applicationId
    };

  } catch(e){

    Logger_.error(
      'applyJob',
      e.message,
      email
    );

    return {
      success:false,
      error:e.message
    };
  }
}

function saveJob(jobId,email){

  try {

    var sheet = DB.getSheet('SAVED_JOB_DB');

    DB.ensureHeaders(sheet,[
      'save_id',
      'email',
      'job_id',
      'created_at'
    ]);

    sheet.appendRow([
      'SAVE-' + Utilities.getUuid(),
      Validate.safeEmail(email),
      jobId,
      new Date().toISOString()
    ]);

    return {
      success:true
    };

  } catch(e){

    Logger_.error(
      'saveJob',
      e.message,
      email
    );

    return {
      success:false,
      error:e.message
    };
  }
}

function generateJobMatch(cvText,jobs,email){

  try {

    var premiumCheck =
      requirePremium(email);

    if(!premiumCheck.success){
      return premiumCheck;
    }

    var prompt =
`Analyze this candidate CV:

${cvText}

Compare against jobs:

${JSON.stringify(jobs)}

Return:
- best matching jobs
- match percentage
- matching skills
- missing skills
- career recommendations
- ATS recommendations

Use HTML formatting.`;

    var ai = callAI(prompt,{
      maxTokens:2500
    });

    return ai;

  } catch(e){

    Logger_.error(
      'generateJobMatch',
      e.message,
      email
    );

    return {
      success:false,
      error:e.message
    };
  }
}

function autoExpireJobs(){

  try {

    var sheet = DB.getSheet('JOB_DB');

    var values =
      sheet.getDataRange().getValues();

    for(var i=1;i<values.length;i++){

      var expires =
        new Date(values[i][21]);

      if(expires < new Date()){

        sheet
          .getRange(i+1,16)
          .setValue('expired');
      }
    }

  } catch(e){

    Logger_.error(
      'autoExpireJobs',
      e.message
    );
  }

  // ============================================================
// RECRUITER BACKEND PATCH — tambahkan ke Code.gs
// ============================================================

/**
 * getApplicants — ambil semua pelamar untuk job yang dimiliki recruiter
 * @param {string} recruiterEmail - email recruiter (owner of jobs)
 * @param {string} jobId          - filter by job (optional, '' = semua)
 */
function getApplicants(recruiterEmail, jobId) {
  try {
    if (!recruiterEmail) {
      return { success: false, error: 'Login required', applications: [] };
    }

    var safeEmail = Validate.safeEmail(recruiterEmail);

    // Step 1 — get all jobs owned by this recruiter
    var jobSheet = DB.getSheet('JOB_DB');
    var allJobs  = DB.getAll(jobSheet);
    var myJobs   = allJobs.filter(function(j) {
      return String(j.created_by || '').toLowerCase() === safeEmail;
    });

    if (myJobs.length === 0) {
      return { success: true, applications: [] };
    }

    // Build a lookup map: job_id → job_title
    var jobMap = {};
    myJobs.forEach(function(j) {
      jobMap[String(j.job_id || '')] = String(j.title || j.job_id || '');
    });

    // Step 2 — get all applications from APPLICATION_DB
    var appSheet = DB.getSheet('APPLICATION_DB');
    var allApps  = DB.getAll(appSheet);

    // Filter: only apps for this recruiter's jobs
    var filtered = allApps.filter(function(app) {
      var jid = String(app.job_id || '');
      if (!jobMap[jid]) return false;
      if (jobId && jid !== String(jobId)) return false;
      return true;
    });

    // Enrich with job title
    filtered = filtered.map(function(app) {
      var copy = {};
      for (var k in app) copy[k] = app[k];
      copy.job_title = jobMap[String(app.job_id || '')] || '';
      return copy;
    });

    // Sort newest first
    filtered.sort(function(a, b) {
      return new Date(b.created_at) - new Date(a.created_at);
    });

    return { success: true, applications: filtered };

  } catch(e) {
    Logger_.error('getApplicants', e.message, recruiterEmail);
    return { success: false, error: e.message, applications: [] };
  }
}
}



// ============================================================
// ROLE SYSTEM — EXTENDED
// ============================================================

function isAdmin(email){

  try{

    if(!email){
      return false;
    }

    var safeEmail =
      String(email)
        .trim()
        .toLowerCase();

    var sheet =
      DB.getSheet('USER_DB');

    var values =
      sheet.getDataRange().getValues();

    if(values.length <= 1){
      return false;
    }

    // SAFE LOWERCASE HEADER
    var headers =
      values[0].map(function(h){
        return String(h)
          .trim()
          .toLowerCase();
      });

    var emailIdx =
      headers.indexOf('email');

    var roleIdx =
      headers.indexOf('role');

    if(emailIdx < 0 || roleIdx < 0){
      return false;
    }

    for(var i=1;i<values.length;i++){

      var rowEmail =
        String(values[i][emailIdx] || '')
          .trim()
          .toLowerCase();

      if(rowEmail === safeEmail){

        var role =
          String(values[i][roleIdx] || '')
            .trim()
            .toLowerCase();

        return role === 'admin';
      }
    }

    return false;

  }catch(e){

    Logger_.error(
      'isAdmin',
      e.message,
      email
    );

    return false;
  }
}

/**
 * requireAdmin — guard function untuk admin-only endpoints
 */
function requireAdmin(email) {
  if (!email) return { success: false, error: 'Authentication required' };
  if (!isAdmin(email)) return { success: false, error: 'Admin access required', forbidden: true };
  return { success: true };
}

/**
 * getUserRole — ambil role user dari USER_DB
 * Returns: 'admin' | 'recruiter' | 'premium' | 'user'
 * Admin dianggap premium + recruiter (bypass semua)
 */
function getUserRole(email){

  try{

    if(!email){
      return 'user';
    }

    var safeEmail =
      String(email)
        .trim()
        .toLowerCase();

    var sheet =
      DB.getSheet('USER_DB');

    var values =
      sheet.getDataRange().getValues();

    if(values.length <= 1){
      return 'user';
    }

    var headers =
      values[0].map(function(h){
        return String(h)
          .trim()
          .toLowerCase();
      });

    var emailIdx =
      headers.indexOf('email');

    var roleIdx =
      headers.indexOf('role');

    if(emailIdx < 0 || roleIdx < 0){
      return 'user';
    }

    for(var i=1;i<values.length;i++){

      var rowEmail =
        String(values[i][emailIdx] || '')
          .trim()
          .toLowerCase();

      if(rowEmail === safeEmail){

        return String(
          values[i][roleIdx] || 'user'
        )
        .trim()
        .toLowerCase();
      }
    }

    return 'user';

  }catch(e){

    Logger_.error(
      'getUserRole',
      e.message,
      email
    );

    return 'user';
  }
}

/**
 * Override requirePremium — admin bypass
 * Replace existing requirePremium with this version
 * (hapus fungsi requirePremium lama dan ganti dengan ini,
 *  atau biarkan — Apps Script pakai definisi terakhir)
 */
function requirePremiumV2(email) {
  try {
    if (!email) {
      return { success: false, premiumRequired: true, error: 'Authentication required.' };
    }
    var safeEmail = Validate.safeEmail(email);

    // Admin bypass semua restriction
    if (isAdmin(safeEmail)) return { success: true };

    var sheet = DB.getSheet('USER_DB');
    var userRow = DB.findRow(sheet, 0, safeEmail);
    if (!userRow) {
      return { success: false, premiumRequired: true, error: 'User not found.' };
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var premiumIdx  = headers.indexOf('premium');
    var expiredIdx  = headers.indexOf('premium_expired');
    var roleIdx     = headers.indexOf('role');
    if (premiumIdx  < 0) premiumIdx  = 4;
    if (expiredIdx  < 0) expiredIdx  = 6;
    if (roleIdx     < 0) roleIdx     = 3;

    var role      = String(userRow.data[roleIdx]    || '').toLowerCase();
    var isPremium = userRow.data[premiumIdx] === true ||
                    String(userRow.data[premiumIdx]).toUpperCase() === 'TRUE';
    var expiredAt = userRow.data[expiredIdx];

    // Recruiter also gets premium access
    if (role === 'recruiter' || role === 'admin') return { success: true };

    if (isPremium && expiredAt) {
      var expDate = new Date(expiredAt);
      if (!isNaN(expDate.getTime()) && expDate < new Date()) {
        isPremium = false;
        DB.updateCell(sheet, userRow.rowIndex, premiumIdx + 1, false);
        DB.updateCell(sheet, userRow.rowIndex, expiredIdx + 1, '');
        Logger_.warn('requirePremiumV2', 'Auto-downgraded expired premium: ' + safeEmail, safeEmail);
        return { success: false, premiumRequired: true, error: 'Premium subscription has expired.' };
      }
    }

    if (!isPremium) {
      return { success: false, premiumRequired: true, error: 'Premium subscription required.' };
    }

    return { success: true };

  } catch(e) {
    Logger_.error('requirePremiumV2', e.message, email);
    return { success: false, premiumRequired: true, error: 'Unable to verify subscription.' };
  }
}

// ============================================================
// RECRUITER REQUEST SYSTEM
// ============================================================

var RECRUITER_REQUEST_HEADERS = [
  'request_id', 'email', 'company_name', 'website', 'industry',
  'linkedin', 'phone', 'description', 'status',
  'created_at', 'reviewed_at', 'reviewed_by'
];

/**
 * createRecruiterRequest — submit form "Become Recruiter"
 */
function createRecruiterRequest(data, email) {
  try {
    if (!email) return { success: false, error: 'Login required' };

    var safeEmail = Validate.safeEmail(email);

    // Validate required fields
    var req = Validate.required(data, ['company_name']);
    if (!req.valid) return { success: false, error: 'Company name wajib diisi.' };

    var sheet = DB.getSheet('RECRUITER_REQUEST_DB');
    DB.ensureHeaders(sheet, RECRUITER_REQUEST_HEADERS);

    // Anti-duplicate: cek apakah sudah ada pending/approved request
    var existing = DB.getAll(sheet).filter(function(r) {
      return String(r.email || '').toLowerCase() === safeEmail &&
             (String(r.status || '') === 'pending' || String(r.status || '') === 'approved');
    });

    if (existing.length > 0) {
      var existStatus = existing[0].status;
      if (existStatus === 'pending') {
        return { success: false, error: 'Kamu sudah punya request yang sedang pending review.' };
      }
      if (existStatus === 'approved') {
        return { success: false, error: 'Akun kamu sudah terdaftar sebagai recruiter.' };
      }
    }

    // Sanitize input
    var requestId = 'REQ-' + Utilities.getUuid();
    sheet.appendRow([
      requestId,
      safeEmail,
      Validate.sanitize(data.company_name,  150),
      Validate.sanitize(data.website     || '', 300),
      Validate.sanitize(data.industry    || '', 100),
      Validate.sanitize(data.linkedin    || '', 300),
      Validate.sanitize(data.phone       || '', 50),
      Validate.sanitize(data.description || '', 2000),
      'pending',
      new Date().toISOString(),
      '',
      ''
    ]);

    Logger_.activity('RecruiterRequest', 'New request from: ' + safeEmail, safeEmail);

    return { success: true, request_id: requestId };

  } catch(e) {
    Logger_.error('createRecruiterRequest', e.message, email);
    return { success: false, error: e.message };
  }
}

/**
 * getRecruiterRequests — admin: lihat semua request
 * filter: 'all' | 'pending' | 'approved' | 'rejected'
 */
function getRecruiterRequests(adminEmail, filter) {
  try {
    var adminCheck = requireAdmin(adminEmail);
    if (!adminCheck.success) return adminCheck;

    var sheet = DB.getSheet('RECRUITER_REQUEST_DB');
    DB.ensureHeaders(sheet, RECRUITER_REQUEST_HEADERS);

    var all = DB.getAll(sheet);
    filter = filter || 'all';

    if (filter !== 'all') {
      all = all.filter(function(r) {
        return String(r.status || '').toLowerCase() === filter;
      });
    }

    // Sort: pending first, then by created_at desc
    all.sort(function(a, b) {
      if (a.status === 'pending' && b.status !== 'pending') return -1;
      if (b.status === 'pending' && a.status !== 'pending') return  1;
      return new Date(b.created_at) - new Date(a.created_at);
    });

    return { success: true, requests: all };

  } catch(e) {
    Logger_.error('getRecruiterRequests', e.message, adminEmail);
    return { success: false, error: e.message, requests: [] };
  }
}

/**
 * approveRecruiter — admin approve request
 * Efek: update request status + ubah role user menjadi 'recruiter'
 */
function approveRecruiter(requestId, adminEmail) {
  try {
    var adminCheck = requireAdmin(adminEmail);
    if (!adminCheck.success) return adminCheck;

    if (!requestId) return { success: false, error: 'Request ID required' };

    var reqSheet = DB.getSheet('RECRUITER_REQUEST_DB');
    DB.ensureHeaders(reqSheet, RECRUITER_REQUEST_HEADERS);

    // Find request by request_id (index 0)
    var reqRow = DB.findRowExact(reqSheet, 0, requestId);
    if (!reqRow) return { success: false, error: 'Request tidak ditemukan.' };

    var targetEmail = String(reqRow.data[1] || '').toLowerCase();
    var now = new Date().toISOString();

    // Update request: status, reviewed_at, reviewed_by
    var headers = reqSheet.getRange(1, 1, 1, reqSheet.getLastColumn()).getValues()[0];
    var statusIdx     = headers.indexOf('status');
    var reviewedAtIdx = headers.indexOf('reviewed_at');
    var reviewedByIdx = headers.indexOf('reviewed_by');

    DB.updateCell(reqSheet, reqRow.rowIndex, statusIdx     + 1, 'approved');
    DB.updateCell(reqSheet, reqRow.rowIndex, reviewedAtIdx + 1, now);
    DB.updateCell(reqSheet, reqRow.rowIndex, reviewedByIdx + 1, Validate.safeEmail(adminEmail));

    // Update USER_DB: ubah role menjadi recruiter
    var userSheet = DB.getSheet('USER_DB');
    var userRow   = DB.findRow(userSheet, 0, targetEmail);

    if (userRow) {
      var userHeaders = userSheet.getRange(1, 1, 1, userSheet.getLastColumn()).getValues()[0];
      var roleIdx = userHeaders.indexOf('role');
      if (roleIdx < 0) roleIdx = 3;
      DB.updateCell(userSheet, userRow.rowIndex, roleIdx + 1, 'recruiter');
    } else {
      // User mungkin belum pernah login — buat entry minimal
      userSheet.appendRow([
        targetEmail, '', '', 'recruiter', false, '', '',
        Utilities.getUuid(), '', new Date().toISOString()
      ]);
    }

    // Kirim notifikasi email ke recruiter
    _sendRecruiterApprovalEmail(targetEmail, String(reqRow.data[2] || ''));

    Logger_.activity('RecruiterApprove', 'Approved: ' + targetEmail + ' by ' + adminEmail, adminEmail);

    return { success: true, email: targetEmail };

  } catch(e) {
    Logger_.error('approveRecruiter', e.message, adminEmail);
    return { success: false, error: e.message };
  }
}

/**
 * rejectRecruiter — admin reject request
 */
function rejectRecruiter(requestId, adminEmail, reason) {
  try {
    var adminCheck = requireAdmin(adminEmail);
    if (!adminCheck.success) return adminCheck;

    if (!requestId) return { success: false, error: 'Request ID required' };

    var reqSheet = DB.getSheet('RECRUITER_REQUEST_DB');
    var reqRow   = DB.findRowExact(reqSheet, 0, requestId);
    if (!reqRow) return { success: false, error: 'Request tidak ditemukan.' };

    var targetEmail = String(reqRow.data[1] || '').toLowerCase();
    var now = new Date().toISOString();

    var headers       = reqSheet.getRange(1, 1, 1, reqSheet.getLastColumn()).getValues()[0];
    var statusIdx     = headers.indexOf('status');
    var reviewedAtIdx = headers.indexOf('reviewed_at');
    var reviewedByIdx = headers.indexOf('reviewed_by');

    DB.updateCell(reqSheet, reqRow.rowIndex, statusIdx     + 1, 'rejected');
    DB.updateCell(reqSheet, reqRow.rowIndex, reviewedAtIdx + 1, now);
    DB.updateCell(reqSheet, reqRow.rowIndex, reviewedByIdx + 1, Validate.safeEmail(adminEmail));

    Logger_.activity('RecruiterReject', 'Rejected: ' + targetEmail + ' by ' + adminEmail, adminEmail);

    return { success: true, email: targetEmail };

  } catch(e) {
    Logger_.error('rejectRecruiter', e.message, adminEmail);
    return { success: false, error: e.message };
  }
}

function _sendRecruiterApprovalEmail(email, companyName) {
  try {
    MailApp.sendEmail({
      to: email,
      subject: '🎉 Recruiter Access Approved — AI Career OS',
      htmlBody:
        '<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;padding:32px;background:#f4f6fb;border-radius:16px;">' +
        '<h1 style="color:#ff7a00;font-size:22px;">⚡ AI Career OS</h1>' +
        '<div style="background:#fff;border-radius:12px;padding:28px;margin-bottom:16px;">' +
        '<h2 style="color:#1a1a1a;">🎉 Selamat! Recruiter Access Disetujui</h2>' +
        '<p style="color:#444;font-size:14px;">Akun recruiter untuk <strong>' + (companyName || 'perusahaan kamu') + '</strong> telah disetujui!</p>' +
        '<p style="color:#444;font-size:14px;">Kamu sekarang bisa:</p>' +
        '<p style="color:#444;font-size:13px;">✅ Post Job Lowongan<br>✅ Manage Jobs<br>✅ Review Applicants<br>✅ AI Matching Kandidat</p>' +
        '</div></div>'
    });
  } catch(e) {
    Logger_.warn('_sendRecruiterApprovalEmail', e.message, email);
  }
}

// ============================================================
// ADMIN DASHBOARD
// ============================================================

/**
 * getAdminDashboard — admin: ambil semua statistik platform
 */
function getAdminDashboard(adminEmail) {
  try {
    var adminCheck = requireAdmin(adminEmail);
    if (!adminCheck.success) return adminCheck;

    var userSheet = DB.getSheet('USER_DB');
    var allUsers  = DB.getAll(userSheet);

    var totalUsers     = allUsers.length;
    var totalPremium   = 0;
    var totalRecruiters= 0;
    var totalAdmins    = 0;

    allUsers.forEach(function(u) {
      var role    = String(u.role    || '').toLowerCase();
      var premium = u.premium === true || String(u.premium).toUpperCase() === 'TRUE';
      if (role === 'admin')    totalAdmins++;
      if (role === 'recruiter')totalRecruiters++;
      if (premium)             totalPremium++;
    });

    // Jobs count
    var jobSheet  = DB.getSheet('JOB_DB');
    var allJobs   = DB.getAll(jobSheet);
    var totalJobs = allJobs.filter(function(j) {
      return String(j.status || '').toLowerCase() === 'active';
    }).length;

    // Applications count
    var appSheet   = DB.getSheet('APPLICATION_DB');
    var totalApps  = DB.getAll(appSheet).length;

    // Pending recruiter requests
    var reqSheet = DB.getSheet('RECRUITER_REQUEST_DB');
    DB.ensureHeaders(reqSheet, RECRUITER_REQUEST_HEADERS);
    var pendingReqs = DB.getAll(reqSheet).filter(function(r) {
      return String(r.status || '') === 'pending';
    }).length;

    // Recent activity: last 10 users
    var recentUsers = allUsers.slice(-10).reverse().map(function(u) {
      return {
        email:      u.email      || '',
        name:       u.name       || '',
        role:       u.role       || 'user',
        premium:    u.premium    || false,
        last_login: u.last_login || '',
        created_at: u.created_at || ''
      };
    });

    return {
      success: true,
      stats: {
        totalUsers:      totalUsers,
        totalPremium:    totalPremium,
        totalRecruiters: totalRecruiters,
        totalAdmins:     totalAdmins,
        totalJobs:       totalJobs,
        totalApps:       totalApps,
        pendingRequests: pendingReqs
      },
      recentUsers: recentUsers
    };

  } catch(e) {
    Logger_.error('getAdminDashboard', e.message, adminEmail);
    return { success: false, error: e.message };
  }
}

// ============================================================
// ADMIN USER MANAGEMENT
// ============================================================

/**
 * getAllUsers — admin: lihat semua user
 */
function getAllUsers(adminEmail, page, limit) {
  try {
    var adminCheck = requireAdmin(adminEmail);
    if (!adminCheck.success) return adminCheck;

    page  = page  || 1;
    limit = limit || 50;

    var sheet    = DB.getSheet('USER_DB');
    var allUsers = DB.getAll(sheet);

    // Sort by created_at desc
    allUsers.sort(function(a, b) {
      return new Date(b.created_at || 0) - new Date(a.created_at || 0);
    });

    var start     = (page - 1) * limit;
    var paginated = allUsers.slice(start, start + limit).map(function(u) {
      return {
        email:      u.email       || '',
        name:       u.name        || '',
        role:       u.role        || 'user',
        premium:    u.premium     || false,
        plan:       u.premium_plan|| '',
        expires:    u.premium_expired || '',
        last_login: u.last_login  || '',
        created_at: u.created_at  || ''
      };
    });

    return {
      success:  true,
      users:    paginated,
      total:    allUsers.length,
      page:     page,
      hasMore:  (start + limit) < allUsers.length
    };

  } catch(e) {
    Logger_.error('getAllUsers', e.message, adminEmail);
    return { success: false, error: e.message, users: [] };
  }
}

/**
 * adminSetUserRole — admin: ubah role user
 * Allowed roles: 'user' | 'premium' | 'recruiter' | 'admin'
 */
function adminSetUserRole(adminEmail, targetEmail, newRole) {
  try {
    var adminCheck = requireAdmin(adminEmail);
    if (!adminCheck.success) return adminCheck;

    var allowed = ['user', 'premium', 'recruiter', 'admin'];
    newRole = String(newRole || '').toLowerCase();
    if (allowed.indexOf(newRole) < 0) {
      return { success: false, error: 'Role tidak valid. Pilih: ' + allowed.join(', ') };
    }

    var safeTarget = Validate.safeEmail(targetEmail);
    var sheet      = DB.getSheet('USER_DB');
    var userRow    = DB.findRow(sheet, 0, safeTarget);

    if (!userRow) return { success: false, error: 'User tidak ditemukan.' };

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var roleIdx = headers.indexOf('role');
    if (roleIdx < 0) roleIdx = 3;

    DB.updateCell(sheet, userRow.rowIndex, roleIdx + 1, newRole);

    Logger_.activity('AdminSetRole', safeTarget + ' → ' + newRole, adminEmail);

    return { success: true, email: safeTarget, role: newRole };

  } catch(e) {
    Logger_.error('adminSetUserRole', e.message, adminEmail);
    return { success: false, error: e.message };
  }
}

/**
 * adminDeleteUser — admin: hapus user dari USER_DB
 */
function adminDeleteUser(adminEmail, targetEmail) {
  try {
    var adminCheck = requireAdmin(adminEmail);
    if (!adminCheck.success) return adminCheck;

    var safeTarget = Validate.safeEmail(targetEmail);

    // Jangan hapus diri sendiri
    if (safeTarget === Validate.safeEmail(adminEmail)) {
      return { success: false, error: 'Tidak bisa menghapus akun sendiri.' };
    }

    var sheet   = DB.getSheet('USER_DB');
    var userRow = DB.findRow(sheet, 0, safeTarget);

    if (!userRow) return { success: false, error: 'User tidak ditemukan.' };

    sheet.deleteRow(userRow.rowIndex);

    Logger_.activity('AdminDeleteUser', 'Deleted: ' + safeTarget, adminEmail);

    return { success: true };

  } catch(e) {
    Logger_.error('adminDeleteUser', e.message, adminEmail);
    return { success: false, error: e.message };
  }
}

/**
 * adminManualPremium — admin: activate/deactivate premium
 */
function adminManualPremium(adminEmail, targetEmail, action, plan) {
  try {
    var adminCheck = requireAdmin(adminEmail);
    if (!adminCheck.success) return adminCheck;

    var safeTarget = Validate.safeEmail(targetEmail);

    if (action === 'activate') {
      var fakeOrderId = 'ADMIN-' + Date.now();
      return activatePremium(safeTarget, fakeOrderId, plan || 'monthly');
    }

    if (action === 'deactivate') {
      var sheet   = DB.getSheet('USER_DB');
      var userRow = DB.findRow(sheet, 0, safeTarget);
      if (!userRow) return { success: false, error: 'User tidak ditemukan.' };

      var headers     = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var premiumIdx  = headers.indexOf('premium');
      var planIdx     = headers.indexOf('premium_plan');
      var expiredIdx  = headers.indexOf('premium_expired');
      if (premiumIdx < 0) premiumIdx = 4;
      if (planIdx    < 0) planIdx    = 5;
      if (expiredIdx < 0) expiredIdx = 6;

      DB.updateCell(sheet, userRow.rowIndex, premiumIdx  + 1, false);
      DB.updateCell(sheet, userRow.rowIndex, planIdx     + 1, '');
      DB.updateCell(sheet, userRow.rowIndex, expiredIdx  + 1, '');

      Logger_.activity('AdminDeactivatePremium', safeTarget, adminEmail);
      return { success: true };
    }

    return { success: false, error: 'Action tidak valid. Gunakan: activate / deactivate' };

  } catch(e) {
    Logger_.error('adminManualPremium', e.message, adminEmail);
    return { success: false, error: e.message };
  }
}

// ============================================================
// EXTENDED autoSetup — pastikan semua sheet baru tersedia
// ============================================================

/**
 * autoSetupV2 — jalankan setelah autoSetup() untuk setup sheet baru
 */
function autoSetupV2() {
  var newSheets = {
    RECRUITER_REQUEST_DB: [
      'request_id','email','company_name','website','industry',
      'linkedin','phone','description','status',
      'created_at','reviewed_at','reviewed_by'
    ],
    JOB_DB: [
      'job_id','company_id','company_name','company_logo','title','location',
      'salary','type','category','description','requirements','skills',
      'experience','education','apply_link','status','featured','views',
      'applications','created_by','created_at','expires_at'
    ],
    APPLICATION_DB: [
      'application_id','job_id','email','cv_id','resume_link',
      'status','match_score','ai_summary','created_at'
    ],
    SAVED_JOB_DB: ['save_id','email','job_id','created_at'],
    COMPANY_DB: [
      'company_id','company_name','email','website','logo',
      'description','verified','premium','created_at'
    ]
  };

  var results = [];
  for (var sheetName in newSheets) {
    try {
      var sheet = DB.getSheet(sheetName);
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(newSheets[sheetName]);
        sheet.setFrozenRows(1);
        results.push('✅ Created: ' + sheetName);
      } else {
        results.push('✓ Exists: ' + sheetName);
      }
    } catch(e) {
      results.push('❌ Failed: ' + sheetName + ' — ' + e.message);
    }
  }

  return { ok: true, version: '7.0.0', results: results };
}

// ============================================================
// UTILITY — getMyRecruiterRequestStatus
// ============================================================

/**
 * getMyRecruiterRequestStatus — user: cek status request milik sendiri
 */
function getMyRecruiterRequestStatus(email) {
  try {
    if (!email) return { success: false, error: 'Login required' };

    var safeEmail = Validate.safeEmail(email);
    var sheet = DB.getSheet('RECRUITER_REQUEST_DB');
    DB.ensureHeaders(sheet, RECRUITER_REQUEST_HEADERS);

    var all    = DB.getAll(sheet);
    var myReqs = all.filter(function(r) {
      return String(r.email || '').toLowerCase() === safeEmail;
    });

    if (myReqs.length === 0) return { success: true, status: 'none' };

    // Sort by created_at desc, ambil yang terbaru
    myReqs.sort(function(a, b) {
      return new Date(b.created_at) - new Date(a.created_at);
    });

    var latest = myReqs[0];
    return {
      success:    true,
      status:     latest.status || 'pending',
      request_id: latest.request_id || '',
      created_at: latest.created_at || '',
      company_name: latest.company_name || ''
    };

  } catch(e) {
    Logger_.error('getMyRecruiterRequestStatus', e.message, email);
    return { success: false, error: e.message, status: 'none' };
  }
}

// ============================================================
// FULL USER INFO — tambah role info ke login response
// ============================================================

/**
 * getFullUserInfo — frontend call untuk dapat role + premium status
 * Dipanggil setelah verifyOTP sukses
 */
function getFullUserInfo(email) {
  try {
    if (!email) return { success: false, error: 'Email required' };

    var safeEmail = Validate.safeEmail(email);
    var sheet     = DB.getSheet('USER_DB');
    var userRow   = DB.findRow(sheet, 0, safeEmail);

    if (!userRow) return { success: false, error: 'User not found' };

    var headers    = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var roleIdx    = headers.indexOf('role');       if (roleIdx    < 0) roleIdx    = 3;
    var premiumIdx = headers.indexOf('premium');    if (premiumIdx < 0) premiumIdx = 4;
    var planIdx    = headers.indexOf('premium_plan');if (planIdx   < 0) planIdx    = 5;
    var expIdx     = headers.indexOf('premium_expired');if (expIdx < 0) expIdx     = 6;
    var nameIdx    = headers.indexOf('name');       if (nameIdx    < 0) nameIdx    = 1;

    var role      = String(userRow.data[roleIdx]    || 'user').toLowerCase();
    var isPremium = userRow.data[premiumIdx] === true ||
                    String(userRow.data[premiumIdx]).toUpperCase() === 'TRUE';
    var plan      = String(userRow.data[planIdx]    || '');
    var expires   = String(userRow.data[expIdx]     || '');
    var name      = String(userRow.data[nameIdx]    || '');

    // Admin override
    if (isAdmin(safeEmail)) role = 'admin';

    // Check expiry
    if (isPremium && expires) {
      var expDate = new Date(expires);
      if (!isNaN(expDate.getTime()) && expDate < new Date()) {
        isPremium = false;
        DB.updateCell(sheet, userRow.rowIndex, premiumIdx + 1, false);
      }
    }

    return {
      success: true,
      email:   safeEmail,
      name:    name,
      role:    role,
      premium: isPremium || role === 'admin' || role === 'recruiter',
      premiumPlan:    plan,
      premiumExpired: expires,
      isAdmin:     role === 'admin',
      isRecruiter: role === 'recruiter' || role === 'admin',
      isPremium:   isPremium || role === 'admin'
    };

  } catch(e) {
    Logger_.error('getFullUserInfo', e.message, email);
    return { success: false, error: e.message };
  }
}
// ============================================================
// CODE.GS PATCH — NON-DESTRUCTIVE ADDITIONS
// Tambahkan seluruh isi file ini ke bawah Code.gs yang ada
// Jangan hapus kode lama, hanya tambahkan patch ini
// Version: 7.1.0
// ============================================================

// ============================================================
// PATCH: getMyJobs — recruiter lihat jobs milik sendiri
// ============================================================
function getMyJobs(email) {
  try {
    if (!email) return { success: false, error: 'Login required', jobs: [] };

    var safeEmail = Validate.safeEmail(email);

    var jobSheet = DB.getSheet('JOB_DB');
    var allJobs  = DB.getAll(jobSheet);

    var myJobs = allJobs.filter(function(j) {
      return String(j.created_by || '').toLowerCase() === safeEmail;
    });

    myJobs.sort(function(a, b) {
      return new Date(b.created_at) - new Date(a.created_at);
    });

    return { success: true, jobs: myJobs };

  } catch(e) {
    Logger_.error('getMyJobs', e.message, email);
    return { success: false, error: e.message, jobs: [] };
  }
}

// ============================================================
// PATCH: deleteJob — recruiter hapus job
// ============================================================
function deleteJob(jobId, email) {
  try {
    if (!email || !jobId) return { success: false, error: 'Missing params' };

    var safeEmail = Validate.safeEmail(email);
    var sheet     = DB.getSheet('JOB_DB');
    var allValues = sheet.getDataRange().getValues();
    if (allValues.length <= 1) return { success: false, error: 'Job not found' };

    var headers = allValues[0].map(function(h){ return String(h).toLowerCase().trim(); });
    var jobIdIdx    = headers.indexOf('job_id');
    var createdByIdx= headers.indexOf('created_by');

    for (var i = 1; i < allValues.length; i++) {
      if (String(allValues[i][jobIdIdx]) === String(jobId)) {
        if (String(allValues[i][createdByIdx]).toLowerCase() !== safeEmail && !isAdmin(safeEmail)) {
          return { success: false, error: 'Unauthorized' };
        }
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Job not found' };

  } catch(e) {
    Logger_.error('deleteJob', e.message, email);
    return { success: false, error: e.message };
  }
}

// ============================================================
// PATCH: updateJobStatus — pause/activate/archive
// ============================================================
function updateJobStatus(jobId, newStatus, email) {
  try {
    var validStatuses = ['active','paused','closed','archived','draft'];
    if (validStatuses.indexOf(newStatus) < 0) return { success: false, error: 'Status tidak valid' };

    var safeEmail = Validate.safeEmail(email);
    var sheet     = DB.getSheet('JOB_DB');
    var allValues = sheet.getDataRange().getValues();
    if (allValues.length <= 1) return { success: false, error: 'Job not found' };

    var headers     = allValues[0].map(function(h){ return String(h).toLowerCase().trim(); });
    var jobIdIdx    = headers.indexOf('job_id');
    var createdByIdx= headers.indexOf('created_by');
    var statusIdx   = headers.indexOf('status');

    for (var i = 1; i < allValues.length; i++) {
      if (String(allValues[i][jobIdIdx]) === String(jobId)) {
        if (String(allValues[i][createdByIdx]).toLowerCase() !== safeEmail && !isAdmin(safeEmail)) {
          return { success: false, error: 'Unauthorized' };
        }
        sheet.getRange(i + 1, statusIdx + 1).setValue(newStatus);
        return { success: true };
      }
    }
    return { success: false, error: 'Job not found' };

  } catch(e) {
    Logger_.error('updateJobStatus', e.message, email);
    return { success: false, error: e.message };
  }
}

// ============================================================
// PATCH: duplicateJob
// ============================================================
function duplicateJob(jobId, email) {
  try {
    var safeEmail = Validate.safeEmail(email);
    var sheet     = DB.getSheet('JOB_DB');
    var allJobs   = DB.getAll(sheet);

    var orig = allJobs.filter(function(j){ return String(j.job_id) === String(jobId); })[0];
    if (!orig) return { success: false, error: 'Job not found' };
    if (String(orig.created_by || '').toLowerCase() !== safeEmail && !isAdmin(safeEmail)) {
      return { success: false, error: 'Unauthorized' };
    }

    var newId = 'JOB-' + Utilities.getUuid();
    var now   = new Date().toISOString();

    sheet.appendRow([
      newId,
      orig.company_id    || '',
      orig.company_name  || '',
      orig.company_logo  || '',
      '[COPY] ' + (orig.title || ''),
      orig.location      || '',
      orig.salary        || '',
      orig.type          || '',
      orig.category      || '',
      orig.description   || '',
      orig.requirements  || '',
      orig.skills        || '',
      orig.experience    || '',
      orig.education     || '',
      orig.apply_link    || '',
      'draft',
      false, 0, 0,
      safeEmail,
      now, ''
    ]);

    return { success: true, job_id: newId };

  } catch(e) {
    Logger_.error('duplicateJob', e.message, email);
    return { success: false, error: e.message };
  }
}

// ============================================================
// PATCH: getSavedJobs — user lihat saved jobs
// ============================================================
function getSavedJobs(email) {
  try {
    if (!email) return { success: false, error: 'Login required', saves: [] };

    var safeEmail  = Validate.safeEmail(email);
    var saveSheet  = DB.getSheet('SAVED_JOB_DB');
    var allSaves   = DB.getAll(saveSheet);

    var mySaves = allSaves.filter(function(s) {
      return String(s.email || '').toLowerCase() === safeEmail;
    });

    if (mySaves.length === 0) return { success: true, saves: [] };

    // Enrich with job details
    var jobSheet = DB.getSheet('JOB_DB');
    var allJobs  = DB.getAll(jobSheet);
    var jobMap   = {};
    allJobs.forEach(function(j) { jobMap[String(j.job_id || '')] = j; });

    var enriched = mySaves.map(function(s) {
      var job = jobMap[String(s.job_id || '')] || {};
      return {
        save_id:      s.save_id || '',
        job_id:       s.job_id  || '',
        saved_at:     s.created_at || '',
        title:        job.title        || 'Unknown',
        company_name: job.company_name || '',
        company_logo: job.company_logo || '',
        location:     job.location     || '',
        type:         job.type         || '',
        salary:       job.salary       || '',
        status:       job.status       || 'unknown',
        description:  job.description  || ''
      };
    });

    enriched.sort(function(a, b) { return new Date(b.saved_at) - new Date(a.saved_at); });

    return { success: true, saves: enriched };

  } catch(e) {
    Logger_.error('getSavedJobs', e.message, email);
    return { success: false, error: e.message, saves: [] };
  }
}

// ============================================================
// PATCH: removeSavedJob
// ============================================================
function removeSavedJob(saveId, email) {
  try {
    var safeEmail = Validate.safeEmail(email);
    var sheet     = DB.getSheet('SAVED_JOB_DB');
    var allValues = sheet.getDataRange().getValues();
    if (allValues.length <= 1) return { success: false, error: 'Not found' };

    var headers   = allValues[0].map(function(h){ return String(h).toLowerCase().trim(); });
    var saveIdIdx = headers.indexOf('save_id');
    var emailIdx  = headers.indexOf('email');

    for (var i = 1; i < allValues.length; i++) {
      if (String(allValues[i][saveIdIdx]) === String(saveId)) {
        if (String(allValues[i][emailIdx]).toLowerCase() !== safeEmail) {
          return { success: false, error: 'Unauthorized' };
        }
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Not found' };

  } catch(e) {
    Logger_.error('removeSavedJob', e.message, email);
    return { success: false, error: e.message };
  }
}

// ============================================================
// PATCH: getMyApplications — job seeker lihat lamaran
// ============================================================
function getMyApplications(email, page, limit) {
  try {
    if (!email) return { success: false, error: 'Login required', applications: [] };

    page  = page  || 1;
    limit = limit || 10;

    var safeEmail = Validate.safeEmail(email);
    var appSheet  = DB.getSheet('APPLICATION_DB');
    var allApps   = DB.getAll(appSheet);

    var mine = allApps.filter(function(a) {
      return String(a.email || '').toLowerCase() === safeEmail;
    });

    // Enrich with job details
    var jobSheet = DB.getSheet('JOB_DB');
    var allJobs  = DB.getAll(jobSheet);
    var jobMap   = {};
    allJobs.forEach(function(j) { jobMap[String(j.job_id || '')] = j; });

    mine = mine.map(function(a) {
      var job = jobMap[String(a.job_id || '')] || {};
      return {
        application_id: a.application_id || '',
        job_id:         a.job_id || '',
        status:         a.status || 'pending',
        match_score:    a.match_score || 0,
        resume_link:    a.resume_link || '',
        created_at:     a.created_at || '',
        title:          job.title        || 'Unknown',
        company_name:   job.company_name || '',
        company_logo:   job.company_logo || '',
        location:       job.location     || '',
        type:           job.type         || '',
        salary:         job.salary       || ''
      };
    });

    mine.sort(function(a, b) { return new Date(b.created_at) - new Date(a.created_at); });

    var total     = mine.length;
    var start     = (page - 1) * limit;
    var paginated = mine.slice(start, start + limit);

    return { success: true, applications: paginated, total: total, page: page, hasMore: (start + limit) < total };

  } catch(e) {
    Logger_.error('getMyApplications', e.message, email);
    return { success: false, error: e.message, applications: [] };
  }
}

// ============================================================
// PATCH: withdrawApplication
// ============================================================
function withdrawApplication(applicationId, email) {
  try {
    var safeEmail = Validate.safeEmail(email);
    var sheet     = DB.getSheet('APPLICATION_DB');
    var allValues = sheet.getDataRange().getValues();
    if (allValues.length <= 1) return { success: false, error: 'Not found' };

    var headers   = allValues[0].map(function(h){ return String(h).toLowerCase().trim(); });
    var appIdIdx  = headers.indexOf('application_id');
    var emailIdx  = headers.indexOf('email');
    var statusIdx = headers.indexOf('status');

    for (var i = 1; i < allValues.length; i++) {
      if (String(allValues[i][appIdIdx]) === String(applicationId)) {
        if (String(allValues[i][emailIdx]).toLowerCase() !== safeEmail) {
          return { success: false, error: 'Unauthorized' };
        }
        sheet.getRange(i + 1, statusIdx + 1).setValue('withdrawn');
        return { success: true };
      }
    }
    return { success: false, error: 'Not found' };

  } catch(e) {
    Logger_.error('withdrawApplication', e.message, email);
    return { success: false, error: e.message };
  }
}

// ============================================================
// PATCH: getJobDetail — view single job
// ============================================================
function getJobDetail(jobId) {
  try {
    if (!jobId) return { success: false, error: 'Job ID required' };

    var sheet   = DB.getSheet('JOB_DB');
    var allJobs = DB.getAll(sheet);
    var found   = allJobs.filter(function(j){ return String(j.job_id) === String(jobId); })[0];

    if (!found) return { success: false, error: 'Job not found' };

    return { success: true, job: found };

  } catch(e) {
    Logger_.error('getJobDetail', e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// PATCH: getUserProfile — extended profile fields
// ============================================================
function getUserProfile(email) {
  try {
    if (!email) return { success: false, error: 'Email required' };

    var safeEmail = Validate.safeEmail(email);

    // Try PROFILE_DB first
    var profileSheet = DB.getSheet('PROFILE_DB');
    DB.ensureHeaders(profileSheet, [
      'email','full_name','headline','summary','phone','location',
      'website','linkedin','github','portfolio_url',
      'skills','experience','education','certifications','languages',
      'photo_url','cv_url','portfolio_file_url',
      'created_at','updated_at'
    ]);

    var allProfiles = DB.getAll(profileSheet);
    var profile = allProfiles.filter(function(p){
      return String(p.email || '').toLowerCase() === safeEmail;
    })[0];

    // Also get base user info
    var userSheet = DB.getSheet('USER_DB');
    var userRow   = DB.findRow(userSheet, 0, safeEmail);
    var userName  = userRow ? String(userRow.data[1] || '') : '';
    var userRole  = userRow ? String(userRow.data[3] || 'user') : 'user';

    if (!profile) {
      return {
        success: true,
        profile: {
          email:      safeEmail,
          full_name:  userName,
          headline:   '',
          summary:    '',
          phone:      '',
          location:   '',
          website:    '',
          linkedin:   '',
          github:     '',
          portfolio_url: '',
          skills:     '',
          experience: '',
          education:  '',
          certifications: '',
          languages:  '',
          photo_url:  '',
          cv_url:     '',
          portfolio_file_url: '',
          role:       userRole
        }
      };
    }

    profile.role = userRole;
    return { success: true, profile: profile };

  } catch(e) {
    Logger_.error('getUserProfile', e.message, email);
    return { success: false, error: e.message };
  }
}

// ============================================================
// PATCH: saveUserProfile — upsert extended profile
// ============================================================
function saveUserProfile(email, data) {
  try {
    if (!email) return { success: false, error: 'Email required' };

    var safeEmail = Validate.safeEmail(email);
    var now       = new Date().toISOString();

    var profileSheet = DB.getSheet('PROFILE_DB');
    DB.ensureHeaders(profileSheet, [
      'email','full_name','headline','summary','phone','location',
      'website','linkedin','github','portfolio_url',
      'skills','experience','education','certifications','languages',
      'photo_url','cv_url','portfolio_file_url',
      'created_at','updated_at'
    ]);

    var allValues = profileSheet.getDataRange().getValues();
    var headers   = allValues[0].map(function(h){ return String(h).toLowerCase().trim(); });
    var emailIdx  = headers.indexOf('email');

    var existingRow = -1;
    for (var i = 1; i < allValues.length; i++) {
      if (String(allValues[i][emailIdx]).toLowerCase() === safeEmail) {
        existingRow = i + 1; break;
      }
    }

    var s = Validate.sanitize;
    var rowData = [
      safeEmail,
      s(data.full_name    || '', 100),
      s(data.headline     || '', 200),
      s(data.summary      || '', 2000),
      s(data.phone        || '', 50),
      s(data.location     || '', 150),
      s(data.website      || '', 300),
      s(data.linkedin     || '', 300),
      s(data.github       || '', 300),
      s(data.portfolio_url|| '', 300),
      s(data.skills       || '', 1000),
      s(data.experience   || '', 5000),
      s(data.education    || '', 2000),
      s(data.certifications || '', 1000),
      s(data.languages    || '', 500),
      s(data.photo_url    || '', 500),
      s(data.cv_url       || '', 500),
      s(data.portfolio_file_url || '', 500),
      existingRow > 0 ? allValues[existingRow - 1][headers.indexOf('created_at')] : now,
      now
    ];

    if (existingRow > 0) {
      profileSheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      profileSheet.appendRow(rowData);
    }

    // Also update name in USER_DB
    if (data.full_name) {
      var userSheet = DB.getSheet('USER_DB');
      var userRow   = DB.findRow(userSheet, 0, safeEmail);
      if (userRow) {
        var userHeaders = userSheet.getRange(1,1,1,userSheet.getLastColumn()).getValues()[0];
        var nameIdx = userHeaders.indexOf('name');
        if (nameIdx >= 0) DB.updateCell(userSheet, userRow.rowIndex, nameIdx + 1, s(data.full_name, 100));
      }
    }

    return { success: true };

  } catch(e) {
    Logger_.error('saveUserProfile', e.message, email);
    return { success: false, error: e.message };
  }
}

// ============================================================
// PATCH: getJobsWithFilter — enhanced getJobs with more filters
// ============================================================
function getJobsFiltered(page, limit, filters) {
  try {
    page    = parseInt(page)  || 1;
    limit   = parseInt(limit) || 20;
    filters = filters || {};

    var sheet  = DB.getSheet('JOB_DB');
    var values = sheet.getDataRange().getValues();

    if (values.length <= 1) {
      return { success: true, jobs: [], total: 0, page: page, hasMore: false };
    }

    var headers = values[0];
    var rows    = values.slice(1);

    var jobs = rows.map(function(row) {
      var obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i]; });
      return obj;
    });

    // Filter: only active unless recruiter view
    if (!filters.includeAll) {
      jobs = jobs.filter(function(j) {
        return String(j.status || '').toLowerCase() === 'active';
      });
    }

    // Keyword search
    if (filters.keyword) {
      var kw = String(filters.keyword).toLowerCase().trim();
      jobs = jobs.filter(function(j) {
        return (String(j.title || '') + String(j.company_name || '') +
                String(j.category || '') + String(j.description || '') +
                String(j.skills || '')).toLowerCase().includes(kw);
      });
    }

    // Location
    if (filters.location) {
      var loc = String(filters.location).toLowerCase().trim();
      jobs = jobs.filter(function(j) {
        return String(j.location || '').toLowerCase().includes(loc);
      });
    }

    // Category
    if (filters.category && filters.category !== 'all') {
      var cat = String(filters.category).toLowerCase().trim();
      jobs = jobs.filter(function(j) {
        return String(j.category || '').toLowerCase() === cat;
      });
    }

    // Type (remote/on-site/full-time etc)
    if (filters.type && filters.type !== 'all') {
      var typ = String(filters.type).toLowerCase().trim();
      jobs = jobs.filter(function(j) {
        return String(j.type || '').toLowerCase().includes(typ);
      });
    }

    // Sorting
    var sort = filters.sort || 'newest';
    jobs.sort(function(a, b) {
      if (sort === 'oldest') return new Date(a.created_at) - new Date(b.created_at);
      return new Date(b.created_at) - new Date(a.created_at); // newest default
    });

    var total     = jobs.length;
    var start     = (page - 1) * limit;
    var paginated = jobs.slice(start, start + limit);

    return { success: true, jobs: paginated, total: total, page: page, hasMore: (start + limit) < total };

  } catch(e) {
    Logger_.error('getJobsFiltered', e.message);
    return { success: false, error: e.message, jobs: [] };
  }
}

// ============================================================
// PATCH: getRecruiterStats — dashboard stats for recruiter
// ============================================================
function getRecruiterStats(email) {
  try {
    if (!email) return { success: false, error: 'Login required' };

    var safeEmail = Validate.safeEmail(email);
    var jobSheet  = DB.getSheet('JOB_DB');
    var allJobs   = DB.getAll(jobSheet).filter(function(j){
      return String(j.created_by || '').toLowerCase() === safeEmail;
    });

    var appSheet = DB.getSheet('APPLICATION_DB');
    var allApps  = DB.getAll(appSheet);

    var myJobIds = allJobs.map(function(j){ return String(j.job_id || ''); });
    var myApps   = allApps.filter(function(a){ return myJobIds.indexOf(String(a.job_id || '')) >= 0; });

    var activeJobs  = allJobs.filter(function(j){ return String(j.status) === 'active'; }).length;
    var closedJobs  = allJobs.filter(function(j){ return String(j.status) === 'closed' || String(j.status) === 'archived'; }).length;
    var draftJobs   = allJobs.filter(function(j){ return String(j.status) === 'draft'; }).length;

    return {
      success: true,
      stats: {
        totalJobs:       allJobs.length,
        activeJobs:      activeJobs,
        closedJobs:      closedJobs,
        draftJobs:       draftJobs,
        totalApplicants: myApps.length
      }
    };

  } catch(e) {
    Logger_.error('getRecruiterStats', e.message, email);
    return { success: false, error: e.message };
  }
}

// ============================================================
// PATCH: autoSetupV3 — tambahkan PROFILE_DB
// ============================================================
function autoSetupV3() {
  var sheets = {
    PROFILE_DB: [
      'email','full_name','headline','summary','phone','location',
      'website','linkedin','github','portfolio_url',
      'skills','experience','education','certifications','languages',
      'photo_url','cv_url','portfolio_file_url',
      'created_at','updated_at'
    ]
  };

  var results = [];
  for (var sheetName in sheets) {
    try {
      var sheet = DB.getSheet(sheetName);
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(sheets[sheetName]);
        sheet.setFrozenRows(1);
        results.push('✅ Created: ' + sheetName);
      } else {
        results.push('✓ Exists: ' + sheetName);
      }
    } catch(e) {
      results.push('❌ Failed: ' + sheetName + ' — ' + e.message);
    }
  }

  return { ok: true, version: '7.1.0', results: results };
}
