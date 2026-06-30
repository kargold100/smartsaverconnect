// ============================================================
// SmartSaver Connect - Google Apps Script
// ============================================================
const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const TAB = { DEALS:'Deals', FUEL:'Fuel', COUPONS:'Coupons', INSURANCE:'Insurance', LOG:'RunLog', SUBS:'Subscribers' };
const MAX_DEALS=1200, MAX_FUEL=300, MAX_COUPONS=200, MAX_INSURANCE=300;
const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36';

function get(url, extra) {
  try {
    const r = UrlFetchApp.fetch(url, {
      method:'get',
      headers:Object.assign({'User-Agent':UA,'Accept-Language':'en-AU,en;q=0.9','Accept':'*/*'}, extra||{}),
      muteHttpExceptions:true, followRedirects:true,
    });
    if (r.getResponseCode() === 200) return r.getContentText();
    console.log('HTTP '+r.getResponseCode()+' | '+url.slice(0,60));
    return null;
  } catch(e) { console.log('ERR | '+url.slice(0,60)+' | '+e.message); return null; }
}

function getJson(url, extra) {
  const t = get(url, extra); if (!t) return null;
  try { return JSON.parse(t); } catch(e) { return null; }
}

function fetchAll(urls) {
  try {
    return UrlFetchApp.fetchAll(urls.map(u => ({
      url:u, method:'get',
      headers:{'User-Agent':UA,'Accept-Language':'en-AU,en;q=0.9','Accept':'*/*'},
      muteHttpExceptions:true, followRedirects:true,
    })));
  } catch(e) { console.log('fetchAll error: '+e.message); return []; }
}

function rssItems(text) {
  if (!text) return [];
  try {
    const root = XmlService.parse(text).getRootElement();
    const ch = root.getChild('channel') || root;
    const items = ch.getChildren('item');
    return items.length ? items : ch.getChildren('entry');
  } catch(e) { return []; }
}

// ── doGet: handles deals JSONP + subscribe + unsubscribe ────────
function doGet(e) {
  const action   = e && e.parameter && e.parameter.action;
  const callback = e && e.parameter && e.parameter.callback;
  let result;

  if (action === 'subscribe') {
    result = handleSubscribe(e.parameter);
  } else if (action === 'unsubscribe') {
    result = handleUnsubscribe(
      (e.parameter.email || '').trim(),
      (e.parameter.token || '').trim()
    );
  } else if (action === 'submitDeal') {
    result = handleDealSubmission(e.parameter);
  } else if (action === 'getBlog') {
    result = getBlogPosts();
  } else if (action === 'addBlog') {
    result = addBlogPost(e.parameter);
  } else if (action === 'getFuel') {
    result = getFuelPrices(e.parameter);
  } else {
    result = getDealsPayload();
  }

  const json = JSON.stringify(result);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Fuel Price Proxy (for fuel.html) ──────────────────────────
// Called from fuel.html via JSONP — bypasses CORS completely
function getFuelPrices(params) {
  var lat = parseFloat(params.lat) || -37.90;
  var lng = parseFloat(params.lng) || 144.75;
  var r = 0.08;
  try {
    var url = 'https://petrolspy.com.au/webservice-1/station/box'
      + '?neLat=' + (lat+r) + '&neLng=' + (lng+r)
      + '&swLat=' + (lat-r) + '&swLng=' + (lng-r);
    var resp = UrlFetchApp.fetch(url, {muteHttpExceptions: true, headers: {
      'User-Agent': 'Mozilla/5.0', 'Accept': 'application/json'
    }});
    if (resp.getResponseCode() === 200) {
      var data = JSON.parse(resp.getContentText());
      if (data && data.message && data.message.list) {
        var FM = {E10:'E10',U91:'ULP',P95:'P95',P98:'P98',DL:'Diesel',LPG:'LPG'};
        var prices = [];
        data.message.list.forEach(function(s) {
          Object.entries(s.prices || {}).forEach(function(e) {
            var p = parseFloat(e[1].amount);
            if (!p || p < 50 || p > 400) return;
            prices.push({n:s.name||'',b:s.brand||'',t:FM[e[0]]||e[0],p:p,
              s:s.address?s.address.suburb||'':''});
          });
        });
        return {ok:true,prices:prices,ts:new Date().toISOString()};
      }
    }
    return {ok:false,error:'PetrolSpy returned '+resp.getResponseCode()};
  } catch(err) {
    return {ok:false,error:err.message};
  }
}

// ── Blog Posts (Google Sheets backend) ─────────────────────────
function getBlogPosts() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName('Blog');
    if (!sh) return { ok: true, data: { blog: [] } };
    const rows = sh.getDataRange().getValues();
    if (rows.length < 2) return { ok: true, data: { blog: [] } };
    const posts = [];
    for (let i = rows.length - 1; i >= 1; i--) {
      const [date, title, tag, body, html] = rows[i];
      if (!title) continue;
      posts.push({
        date: date ? Utilities.formatDate(new Date(date), 'Australia/Melbourne', 'dd MMM yyyy') : '',
        title: String(title), tag: String(tag || 'general'),
        body: String(body || ''), html: String(html || '')
      });
    }
    return { ok: true, data: { blog: posts.slice(0, 50) } };
  } catch(e) { return { ok: false, error: e.message }; }
}

function addBlogPost(params) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName('Blog');
    if (!sh) {
      sh = ss.insertSheet('Blog');
      sh.getRange(1,1,1,5).setValues([['Date','Title','Tag','Body','HTML']]);
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,5).setFontWeight('bold').setBackground('#e8f5f2');
      sh.setColumnWidth(2, 250);
      sh.setColumnWidth(4, 500);
      sh.setColumnWidth(5, 500);
    }
    const title = (params.title || '').trim();
    const tag = (params.tag || 'general').trim();
    const body = (params.body || '').trim();
    if (!title || !body) return { ok: false, error: 'Title and body required' };
    // If body contains HTML tags, put it in HTML column; otherwise in Body column
    const hasHtml = /<[a-z][\s\S]*>/i.test(body);
    sh.appendRow([new Date(), title, tag, hasHtml ? '' : body, hasHtml ? body : '']);
    return { ok: true, message: 'Post added' };
  } catch(e) { return { ok: false, error: e.message }; }
}

// ── Deal Submission Handler ───────────────────────────────────
// Retailers submit deals directly — no middleman.
// Submissions go to a "Submissions" sheet for admin review.
// Approved deals can be moved to the Deals sheet manually.
function handleDealSubmission(params) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName('Submissions');
    if (!sh) {
      sh = ss.insertSheet('Submissions');
      sh.getRange(1,1,1,12).setValues([['Submitted','Business','Title','Description','Sale Price','Original Price','Deal URL','Category','Contact Name','Contact Email','Expiry','Status']]);
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,12).setFontWeight('bold').setBackground('#fff3cd');
      sh.setColumnWidth(2, 150);
      sh.setColumnWidth(3, 250);
      sh.setColumnWidth(4, 300);
      sh.setColumnWidth(7, 300);
    }
    
    const business = (params.business || '').trim();
    const title    = (params.title    || '').trim();
    const desc     = (params.desc     || '').trim();
    const priceNow = (params.priceNow || '').trim();
    const priceWas = (params.priceWas || '').trim();
    const dealUrl  = (params.dealUrl  || '').trim();
    const category = (params.category || 'retail').trim();
    const name     = (params.contactName  || '').trim();
    const email    = (params.contactEmail || '').trim();
    const expiry   = (params.expiry   || '').trim();
    
    if (!business || !title || !dealUrl || !email) {
      return { ok: false, error: 'Missing required fields' };
    }
    
    sh.appendRow([
      new Date().toISOString(),
      business, title, desc, priceNow, priceWas, dealUrl,
      category, name, email, expiry, 'PENDING'
    ]);
    
    // Notify admin via email (optional — sends to sheet owner)
    try {
      const adminEmail = Session.getActiveUser().getEmail();
      if (adminEmail) {
        MailApp.sendEmail({
          to: adminEmail,
          subject: 'SmartSaver Connect: New Deal Submitted by ' + business,
          htmlBody: '<h2>New Deal Submission</h2>'
            + '<p><strong>Business:</strong> ' + business + '</p>'
            + '<p><strong>Title:</strong> ' + title + '</p>'
            + '<p><strong>Description:</strong> ' + desc + '</p>'
            + '<p><strong>Price:</strong> $' + priceNow + (priceWas ? ' (was $' + priceWas + ')' : '') + '</p>'
            + '<p><strong>URL:</strong> <a href="' + dealUrl + '">' + dealUrl + '</a></p>'
            + '<p><strong>Category:</strong> ' + category + '</p>'
            + '<p><strong>Contact:</strong> ' + name + ' (' + email + ')</p>'
            + '<p><strong>Expiry:</strong> ' + (expiry || 'Not specified') + '</p>'
            + '<hr><p>Review in your <a href="https://docs.google.com/spreadsheets/d/' + SHEET_ID + '">Google Sheet</a> → Submissions tab.</p>'
            + '<p>To approve: change Status from PENDING to APPROVED, then copy the row to the Deals sheet.</p>',
        });
      }
    } catch(mailErr) { console.log('Admin notification failed: ' + mailErr.message); }
    
    return { ok: true, message: 'Deal submitted for review' };
  } catch(err) {
    console.log('Deal submission error: ' + err.message);
    return { ok: false, error: err.message };
  }
}

// ── Approve submitted deal (run manually from sheet) ──────────
// Select a row in the Submissions sheet and run this to copy it to Deals
function approveSelectedDeal() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName('Submissions');
  if (!sh) { SpreadsheetApp.getUi().alert('No Submissions sheet found.'); return; }
  
  const row = sh.getActiveRange().getRow();
  if (row < 2) { SpreadsheetApp.getUi().alert('Select a submission row first.'); return; }
  
  const data = sh.getRange(row, 1, 1, 12).getValues()[0];
  const [submitted, business, title, desc, priceNow, priceWas, dealUrl, category, contactName, contactEmail, expiry, status] = data;
  
  if (status === 'APPROVED') { SpreadsheetApp.getUi().alert('This deal is already approved.'); return; }
  
  // Calculate discount percentage
  const pn = parseFloat(priceNow), pw = parseFloat(priceWas);
  const pct = (!isNaN(pn) && !isNaN(pw) && pw > pn) ? Math.round((pw - pn) / pw * 100) : '';
  
  // Write to Deals sheet
  const deal = {
    vendor: business, category: category, title: title,
    description: desc + (contactName ? ' | Submitted by ' + business : ''),
    priceNow: priceNow, priceWas: priceWas, discountPct: pct,
    dealUrl: dealUrl, externalId: 'sub_' + row, scrapedAt: now()
  };
  writeDeals(ss, [deal]);
  
  // Update status
  sh.getRange(row, 12).setValue('APPROVED');
  sh.getRange(row, 1, 1, 12).setBackground('#d4edda');
  
  SpreadsheetApp.getUi().alert('Deal approved and published: ' + title);
}

// Returns all deals/fuel/coupons/insurance for the portal
function getDealsPayload() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheets = {};
    ss.getSheets().forEach(s => { sheets[s.getName()] = s; });
    function read(name, lim) {
      const sh = sheets[name]; if (!sh) return [];
      const lr = sh.getLastRow(); if (lr < 2) return [];
      const lc = sh.getLastColumn(); if (lc < 1) return [];
      const vals = sh.getRange(1,1,Math.min(lr,lim+1),lc).getValues();
      if (vals.length < 2) return [];
      const hdr = vals[0];
      return vals.slice(1).map(row => { const o={}; hdr.forEach((h,i)=>{ if(h) o[h]=row[i]; }); return o; });
    }
    return { ok:true, data:{
      deals:     read(TAB.DEALS,     400),
      fuel:      read(TAB.FUEL,      200),
      coupons:   read(TAB.COUPONS,   100),
      insurance: read(TAB.INSURANCE, 150),
      lastRun:   getLastRun(ss),
    }};
  } catch(err) { return { ok:false, error:err.message }; }
}

// ── SUBSCRIPTION SYSTEM ─────────────────────────────────────────

function getOrCreateSubsSheet(ss) {
  let sh = ss.getSheetByName(TAB.SUBS);
  if (!sh) {
    sh = ss.insertSheet(TAB.SUBS);
    const hdrs = ['subscribedAt','name','email','categories','active','lastSentAt','unsubToken'];
    sh.getRange(1,1,1,hdrs.length).setValues([hdrs]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,hdrs.length).setFontWeight('bold').setBackground('#f3f3f3');
  }
  return sh;
}

function handleSubscribe(params) {
  try {
    const name  = (params.name  || '').trim().slice(0, 60);
    const email = (params.email || '').trim().toLowerCase();
    const cats  = (params.cats  || 'grocery,retail,fuel').trim();
    if (!email || !email.includes('@') || !email.includes('.')) {
      return { ok:false, error:'Invalid email address' };
    }
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = getOrCreateSubsSheet(ss);
    const rows = sh.getDataRange().getValues();
    // Check existing
    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][2]||'').toLowerCase() === email) {
        if (rows[i][4] === false || rows[i][4] === 'FALSE') {
          sh.getRange(i+1,5).setValue(true);
          sh.getRange(i+1,4).setValue(cats);
        }
        try { sendWelcomeEmail(name||rows[i][1]||'there', email, cats); } catch(ex){}
        return { ok:true, message:'Subscribed' };
      }
    }
    // New subscriber
    const token = Utilities.getUuid();
    sh.appendRow([now(), name, email, cats, true, '', token]);
    try { sendWelcomeEmail(name||'there', email, cats); } catch(ex) { console.log('Welcome email: '+ex.message); }
    return { ok:true, message:'Subscribed successfully' };
  } catch(err) {
    console.log('Subscribe error: ' + err.message);
    return { ok:false, error:err.message };
  }
}

function handleUnsubscribe(email, token) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(TAB.SUBS); if (!sh) return { ok:false, error:'Not found' };
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][2]||'').toLowerCase() === email.toLowerCase()) {
        // Accept either: correct UUID token OR email-as-token (from portal form)
        if (rows[i][6] === token || token === email || !token) {
          sh.getRange(i+1, 5).setValue(false);
          return { ok:true, message:'Unsubscribed' };
        }
      }
    }
    // If email not found in subscribers, still return success (don't reveal subscriber list)
    return { ok:true, message:'Unsubscribed' };
  } catch(err) { return { ok:false, error:err.message }; }
}

function sendWelcomeEmail(name, email, cats) {
  const catLabels = {
    grocery:'🛒 Grocery (Woolworths, Coles, ALDI)',
    retail:'🏬 Retail & electronics', fuel:'⛽ Fuel prices',
    food:'🛵 Food delivery (Uber Eats, Menulog)',
    telco:'📱 Mobile plans & NBN', travel:'✈️ Travel & flights',
    cashback:'💰 Cashback offers', energy:'⚡ Energy deals',
  };
  const catHtml = (cats||'').split(',')
    .map(c => catLabels[c.trim()]||c.trim())
    .filter(Boolean)
    .map(c => '<li style="margin:5px 0">'+c+'</li>').join('');

  MailApp.sendEmail({
    to: email,
    subject: '🎉 Welcome to SmartSaver Connect Deal Alerts!',
    name: 'SmartSaver Connect',
    htmlBody: `
<div style="font-family:system-ui,sans-serif;max-width:560px;margin:0 auto;color:#1c1917">
  <div style="background:#15803d;padding:28px 32px;border-radius:12px 12px 0 0">
    <h1 style="color:#fff;margin:0;font-size:22px;font-weight:800">💰 SmartSaver Connect</h1>
    <p style="color:rgba(255,255,255,.8);margin:6px 0 0;font-size:13px">Australian Deals Portal · Point Cook, Melbourne</p>
  </div>
  <div style="background:#fff;padding:32px;border:1px solid #e2e8f0;border-top:none;border-radius:0 0 12px 12px">
    <h2 style="font-size:20px;font-weight:700;margin:0 0 12px">Hi ${name}! You're subscribed 🎉</h2>
    <p style="font-size:15px;color:#57534e;line-height:1.6;margin:0 0 20px">
      You'll receive a daily email with the best deals we find matching your preferences:
    </p>
    <ul style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;padding:14px 18px 14px 36px;margin:0 0 24px;font-size:14px;color:#15803d;line-height:1.8">
      ${catHtml}
    </ul>
    <p style="font-size:14px;color:#57534e;line-height:1.6;margin:0 0 24px">
      Your first digest will arrive tomorrow morning. Browse today's deals in the meantime:
    </p>
    <a href="https://smartsaverconnect.netlify.app"
       style="display:inline-block;background:#16a34a;color:#fff;padding:12px 24px;border-radius:8px;text-decoration:none;font-weight:700;font-size:14px">
      Browse today's deals →
    </a>
    <hr style="border:none;border-top:1px solid #f1f5f9;margin:28px 0 16px"/>
    <p style="font-size:11px;color:#94a3b8;margin:0;line-height:1.6">
      To unsubscribe, reply to any email with the word "unsubscribe".<br>
      SmartSaver Connect · Point Cook, Melbourne VIC
    </p>
  </div>
</div>`
  });
}

// ── DAILY DIGEST (runs 7am daily via trigger) ────────────────────
function sendDailyDigest() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(TAB.SUBS);
  if (!sh) { console.log('No subscribers sheet yet'); return; }

  const rows = sh.getDataRange().getValues();
  if (rows.length < 2) { console.log('No subscribers yet'); return; }

  const payload = getDealsPayload();
  const allDeals    = (payload.data && payload.data.deals)     || [];
  const allCoupons  = (payload.data && payload.data.coupons)   || [];
  const allFuel     = (payload.data && payload.data.fuel)      || [];

  // Best fuel price per type
  const fuelByType = {};
  allFuel.forEach(f => { if (!fuelByType[f.fuelType]||f.priceCents<fuelByType[f.fuelType].priceCents) fuelByType[f.fuelType]=f; });
  const bestFuel = Object.values(fuelByType);

  const headers = rows[0];
  const iEmail  = headers.indexOf('email');
  const iName   = headers.indexOf('name');
  const iCats   = headers.indexOf('categories');
  const iActive = headers.indexOf('active');
  const iSent   = headers.indexOf('lastSentAt');
  const iToken  = headers.indexOf('unsubToken');

  let sent = 0;
  const today = Utilities.formatDate(new Date(), 'Australia/Melbourne', 'EEE d MMM yyyy');

  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][iActive] || rows[i][iActive] === false || rows[i][iActive] === 'FALSE') continue;
    const email = (rows[i][iEmail]||'').trim();
    const name  = (rows[i][iName] ||'there').trim();
    const cats  = (rows[i][iCats] ||'grocery,retail,fuel').split(',').map(c=>c.trim());
    const token = rows[i][iToken] || '';
    if (!email) continue;

    // Filter deals by subscriber categories
    const myDeals = allDeals.filter(d => {
      const c = d.category||'';
      return cats.some(cat =>
        cat === c ||
        (cat==='grocery' && ['woolworths','coles','aldi','grocery','other'].includes(c)) ||
        (cat==='retail'  && ['retail','amazon','bigw'].includes(c)) ||
        (cat==='food'    && c==='food') ||
        (cat==='travel'  && c==='travel') ||
        (cat==='telco'   && ['telco','nbn'].includes(c)) ||
        (cat==='energy'  && c==='energy')
      );
    }).sort((a,b)=>(parseInt(b.discountPct)||0)-(parseInt(a.discountPct)||0)).slice(0,12);

    const myCoupons = (cats.includes('retail')||cats.includes('grocery')) ? allCoupons.slice(0,5) : [];
    const myFuel    = cats.includes('fuel') ? bestFuel.slice(0,6) : [];

    if (!myDeals.length && !myCoupons.length && !myFuel.length) continue;

    try {
      MailApp.sendEmail({
        to: email,
        subject: '💰 SmartSaver Deals — ' + today + ' (' + myDeals.length + ' deals)',
        name: 'SmartSaver Connect',
        htmlBody: buildDigestEmail(name, email, token, myDeals, myCoupons, myFuel, today),
      });
      sh.getRange(i+1, iSent+1).setValue(now());
      sent++;
      Utilities.sleep(300);
    } catch(err) {
      console.log('Digest email error for ' + email + ': ' + err.message);
    }
  }
  console.log('Daily digest: ' + sent + ' emails sent');
  logSingle(ss, 'Email digest', sent, 0);
}

function buildDigestEmail(name, email, token, deals, coupons, fuel, today) {
  const portal = 'https://smartsaverconnect.netlify.app';

  const dealRows = deals.map(d => {
    const disc = parseInt(d.discountPct)>0
      ? '<span style="background:#fee2e2;color:#b91c1c;font-size:11px;font-weight:700;padding:2px 7px;border-radius:10px;margin-left:6px">-'+d.discountPct+'%</span>' : '';
    const price = d.priceNow && parseFloat(d.priceNow)>0
      ? '<strong style="color:#15803d">$'+parseFloat(d.priceNow).toFixed(2)+'</strong> ' : '';
    const link = d.dealUrl
      ? ' <a href="'+d.dealUrl+'" style="color:#15803d;font-size:12px;font-weight:600;text-decoration:none">View →</a>' : '';
    return '<tr><td style="padding:10px 0;border-bottom:1px solid #f1f5f9;vertical-align:top">'
      +'<div style="font-size:11px;color:#94a3b8;font-weight:700;text-transform:uppercase;letter-spacing:.05em;margin-bottom:3px">'+(d.vendor||'')+disc+'</div>'
      +'<div style="font-size:14px;font-weight:600;color:#1c1917;line-height:1.4;margin-bottom:4px">'+(d.title||'').slice(0,100)+'</div>'
      +'<div style="font-size:12px;color:#57534e">'+price+link+'</div>'
      +'</td></tr>';
  }).join('');

  const couponRows = coupons.map(c =>
    '<tr><td style="padding:8px 0;border-bottom:1px solid #f1f5f9">'
    +'<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">'
    +'<span style="font-family:monospace;background:#f0fdf4;border:1.5px dashed #86efac;padding:3px 10px;border-radius:6px;font-size:13px;font-weight:700;color:#15803d">'+(c.code||'')+'</span>'
    +'<span style="font-size:13px;color:#57534e">'+(c.vendor||'')+( c.discountText?' — '+c.discountText:'' )+'</span>'
    +(c.dealUrl?'<a href="'+c.dealUrl+'" style="font-size:12px;color:#15803d;text-decoration:none;font-weight:600">View →</a>':'')
    +'</div></td></tr>'
  ).join('');

  const fuelCols = {ULP:'#15803d',E10:'#1d4ed8',P95:'#c2410c',P98:'#b91c1c',Diesel:'#5b21b6',LPG:'#7e22ce'};
  const fuelBgs  = {ULP:'#dcfce7',E10:'#dbeafe',P95:'#fff7ed',P98:'#fee2e2',Diesel:'#ede9fe',LPG:'#fdf4ff'};
  const fuelCells = fuel.map(f => {
    const c = fuelCols[f.fuelType]||'#15803d', b = fuelBgs[f.fuelType]||'#f0fdf4';
    return '<td style="padding:6px;text-align:center">'
      +'<div style="background:'+b+';border-radius:8px;padding:10px 12px;min-width:75px">'
      +'<div style="font-size:10px;font-weight:700;color:'+c+';text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px">'+(f.fuelType||'')+'</div>'
      +'<div style="font-size:20px;font-weight:800;color:'+c+'">'+parseFloat(f.priceCents).toFixed(1)+'¢</div>'
      +'<div style="font-size:10px;color:#94a3b8;margin-top:3px">'+(f.brand||f.stationName||'')+'</div>'
      +'</div></td>';
  }).join('');

  return `<!DOCTYPE html><html lang="en-AU"><head><meta charset="UTF-8"/></head>
<body style="margin:0;padding:16px;background:#f8fafc;font-family:system-ui,sans-serif">
<div style="max-width:580px;margin:0 auto">
  <div style="background:#15803d;padding:24px 28px;border-radius:12px 12px 0 0;display:flex;justify-content:space-between;align-items:flex-start">
    <div>
      <h1 style="color:#fff;margin:0;font-size:20px;font-weight:800">💰 SmartSaver Connect</h1>
      <p style="color:rgba(255,255,255,.75);margin:4px 0 0;font-size:13px">Daily digest · ${today}</p>
    </div>
    <div style="text-align:right"><div style="font-size:28px;font-weight:800;color:#fff">${deals.length}</div><div style="font-size:11px;color:rgba(255,255,255,.7);font-weight:600;text-transform:uppercase">new deals</div></div>
  </div>
  <div style="background:#fff;padding:28px;border:1px solid #e2e8f0;border-top:none;border-radius:0 0 12px 12px">
    <p style="font-size:15px;color:#57534e;margin:0 0 20px;line-height:1.6">Hi ${name}! Here are today's best deals matching your preferences.</p>
    ${deals.length ? '<h2 style="font-size:16px;font-weight:800;margin:0 0 12px;color:#1c1917">🏷️ Today\'s top deals</h2><table style="width:100%;border-collapse:collapse">'+dealRows+'</table>' : ''}
    ${coupons.length ? '<h2 style="font-size:16px;font-weight:800;margin:24px 0 12px;color:#1c1917">🎟️ Active coupon codes</h2><table style="width:100%;border-collapse:collapse">'+couponRows+'</table>' : ''}
    ${fuel.length ? '<h2 style="font-size:16px;font-weight:800;margin:24px 0 12px;color:#1c1917">⛽ Best fuel prices — VIC</h2><table style="border-collapse:collapse"><tr>'+fuelCells+'</tr></table><p style="font-size:12px;color:#94a3b8;margin:8px 0 0"><a href="https://www.motormouth.com.au/fuel-price-map?postcode=3030" style="color:#15803d">Find cheapest fuel near Point Cook →</a></p>' : ''}
    <div style="margin-top:28px;text-align:center">
      <a href="${portal}" style="display:inline-block;background:#16a34a;color:#fff;padding:12px 28px;border-radius:8px;text-decoration:none;font-weight:700;font-size:14px">See all deals on SmartSaver Connect →</a>
    </div>
    <hr style="border:none;border-top:1px solid #f1f5f9;margin:28px 0 16px"/>
    <p style="font-size:11px;color:#94a3b8;margin:0;line-height:1.6;text-align:center">
      SmartSaver Connect · Point Cook, Melbourne VIC<br>
      <a href="${portal}?action=unsubscribe&email=${encodeURIComponent(email)}&token=${token}" style="color:#94a3b8">Unsubscribe</a>
      · Data from OzBargain, ALDI and community deal posts
    </p>
  </div>
</div></body></html>`;
}

// ── VENDOR RULES ───────────────────────────────────────────────
const VENDOR_RULES = [
  // Grocery
  [/\bwoolworths\b|\bwoolies\b/,    'Woolworths',         'woolworths'],
  [/\bcoles\b/,                      'Coles',              'coles'],
  [/\baldi\b/,                       'ALDI',               'aldi'],
  [/\biga\b/,                        'IGA',                'grocery'],
  [/\bcostco\b/,                     'Costco',             'grocery'],
  [/\beveryday\s*rewards?\b/,        'Everyday Rewards',   'woolworths'],
  [/\bflybuys?\b/,                   'Flybuys',            'coles'],
  // Electronics & Retail
  [/\bjb\s*hi-?fi\b/,               'JB Hi-Fi',           'retail'],
  [/\bchemi[s]?t\s*warehouse\b/,    'Chemist Warehouse',  'chemist'],
  [/\bbig\s*w\b/,                    'Big W',              'bigw'],
  [/\bkmart\b/,                      'Kmart',              'retail'],
  [/\btarget\b/,                     'Target',             'retail'],
  [/\bofficeworks\b/,                'Officeworks',        'retail'],
  [/\bharvey\s*norm/,                'Harvey Norman',      'retail'],
  [/\bthe\s*good\s*guys\b/,         'The Good Guys',      'retail'],
  [/\bkogan\b/,                      'Kogan',              'retail'],
  [/\bcatch\b/,                      'Catch',              'retail'],
  [/\bmyer\b/,                       'Myer',               'retail'],
  [/\bdavid\s*jones\b/,             'David Jones',        'retail'],
  [/\bamazon\b/,                     'Amazon AU',          'amazon'],
  [/\bebay\b/,                       'eBay',               'retail'],
  [/\buniqlo\b/,                     'Uniqlo',             'retail'],
  [/\beb\s*games\b/,                'EB Games',           'retail'],
  [/\bmwave\b/,                      'Mwave',              'retail'],
  [/\bbunnings\b/,                   'Bunnings',           'retail'],
  [/\bikea\b/,                       'IKEA',               'retail'],
  [/\bbaby\s*bunting\b/,            'Baby Bunting',       'retail'],
  // Health
  [/\bpriceline\b/,                  'Priceline',          'chemist'],
  [/\bmy\s*chemist\b/,              'My Chemist',         'chemist'],
  // Sports
  [/\brebel\s*(sport)?\b/,          'Rebel Sport',        'retail'],
  [/\banaconda\b/,                   'Anaconda',           'retail'],
  [/\bbcf\b/,                        'BCF',                'retail'],
  [/\bsupercheap\b/,                 'Supercheap Auto',    'retail'],
  [/\brepco\b/,                      'Repco',              'retail'],
  // Food delivery
  [/\buber\s*eats\b/,               'Uber Eats',          'food'],
  [/\bmenulog\b/,                    'Menulog',            'food'],
  [/\bdoordash\b/,                   'DoorDash',           'food'],
  [/\bhungry\s*jack/,               'Hungry Jacks',       'food'],
  [/\bdominos\b/,                    'Dominos',            'food'],
  [/\bpizza\s*hut\b/,               'Pizza Hut',          'food'],
  // Liquor
  [/\bdan\s*murphy/,                 'Dan Murphys',        'retail'],
  [/\bbws\b/,                        'BWS',                'retail'],
  [/\bliquorland\b/,                 'Liquorland',         'retail'],
  // Travel
  [/\bqantas\b/,                     'Qantas',             'travel'],
  [/\bjetstar\b/,                    'Jetstar',            'travel'],
  [/\bvirgin\s*aus/,                 'Virgin Australia',   'travel'],
  [/\btigerair\b/,                   'Tigerair',           'travel'],
  [/\bairbnb\b/,                     'Airbnb',             'travel'],
  [/\bbooking\.com\b/,               'Booking.com',        'travel'],
  [/\bexpedia\b/,                    'Expedia',            'travel'],
  [/\bwotif\b/,                      'Wotif',              'travel'],
  [/\bagoda\b/,                      'Agoda',              'travel'],
  [/\brentalcars?\b/,               'Rental Cars',        'travel'],
  [/\bflexigroup\b|\bhumm\b/,       'Humm',               'finance'],
  // Energy
  [/\bagl\b/,                        'AGL',                'energy'],
  [/\borigin\s*energy\b/,            'Origin Energy',      'energy'],
  [/\benergy\s*australia\b/,         'EnergyAustralia',   'energy'],
  [/\breamped\b/,                    'ReAmped Energy',     'energy'],
  [/\bamber\s*electric\b/,           'Amber Electric',     'energy'],
  [/\bpowershop\b/,                  'Powershop',          'energy'],
  [/\balinta\b/,                     'Alinta Energy',      'energy'],
  // Telco & NBN
  [/\btelstra\b/,                    'Telstra',            'telco'],
  [/\boptus\b/,                      'Optus',              'telco'],
  [/\bvodafone\b/,                   'Vodafone',           'telco'],
  [/\baldi\s*mobile\b/,             'ALDI Mobile',        'telco'],
  [/\bboost\s*mobile\b/,            'Boost Mobile',       'telco'],
  [/\bkogan\s*mobile\b/,            'Kogan Mobile',       'telco'],
  [/\blebara\b/,                     'Lebara Mobile',      'telco'],
  [/\bamaysim\b/,                    'amaysim',            'telco'],
  [/\bspintel\b/,                    'Spintel',            'nbn'],
  [/\binternode\b/,                  'Internode',          'nbn'],
  [/\biinet\b/,                      'iiNet',              'nbn'],
  [/\btpg\b/,                        'TPG',                'nbn'],
  [/\baussie\s*broadband\b/,        'Aussie Broadband',   'nbn'],
  [/\bstarlink\b/,                   'Starlink',           'nbn'],
  [/\bsuperloop\b/,                  'Superloop',          'nbn'],
  [/\bnbn\b/,                        'NBN',                'nbn'],
];

function detectVendor(text) {
  const low = text.toLowerCase();
  for (const [re, name, cat] of VENDOR_RULES) {
    if (low.match(re)) return {name, cat};
  }
  return null;
}

function parseOzbItem(item, i, defaultCat) {
  const title = item.getChildText('title') || '';
  const link  = item.getChildText('link')  || '';
  const desc  = stripHtml(item.getChildText('description') || '');
  if (!title) return null;
  const vendor = detectVendor(title + ' ' + desc);
  const p = xPrice(title+' '+desc), was = xWas(desc), pct = xPct(title+' '+desc);
  return {
    vendor:      vendor ? vendor.name : 'Various',
    category:    vendor ? vendor.cat  : (defaultCat||'other'),
    title:       title.slice(0,150),
    description: desc.slice(0,250),
    priceNow:    p   != null ? parseFloat(p)   : '',
    priceWas:    was != null ? parseFloat(was) : '',
    discountPct: pct != null ? parseInt(pct) : (p&&was ? calcPct(p,was) : ''),
    dealUrl:     link,
    externalId:  'ozb_'+i+'_'+title.slice(0,20).replace(/\W/g,''),
    scrapedAt:   now(),
  };
}

// ── SCRAPERS ───────────────────────────────────────────────────

function runGroceries() {
  const t0 = Date.now(), ss = SpreadsheetApp.openById(SHEET_ID);
  const deals = [];

  // 1. OzBargain grocery + food feeds
  const ozbUrls = [
    'https://www.ozbargain.com.au/cat/groceries/feed',
    'https://www.ozbargain.com.au/cat/food-drink/feed',
  ];
  fetchAll(ozbUrls).forEach((resp,fi) => {
    if (!resp||resp.getResponseCode()!==200) return;
    rssItems(resp.getContentText()).forEach((item,i) => {
      const d = parseOzbItem(item, fi+'_'+i, 'grocery');
      if (d) deals.push(d);
    });
    console.log('OzBargain grocery feed '+fi+': '+rssItems(resp.getContentText()).length);
  });

  // 2. ALDI direct page — JSON-LD, Next.js data, regex fallback
  const aldiHtml = get('https://www.aldi.com.au/en/special-buys/',
    {'Accept':'text/html','Referer':'https://www.aldi.com.au/'});
  if (aldiHtml) {
    let found = false;
    (aldiHtml.match(/<script[^>]+application\/ld\+json[^>]*>([\s\S]*?)<\/script>/gi)||[]).forEach(block => {
      try {
        const json = JSON.parse(block.replace(/<[^>]+>/g,'').trim());
        const items = json['@type']==='ItemList'?(json.itemListElement||[]).map(e=>e.item||e):[json];
        items.forEach((p,i) => {
          if (!p||!p.name) return;
          const offer = Array.isArray(p.offers)?p.offers[0]:(p.offers||{});
          deals.push({ vendor:'ALDI', category:'aldi', title:p.name,
            description:p.description||'ALDI Special Buy',
            priceNow:offer.price?parseFloat(offer.price):'', priceWas:'', discountPct:'',
            dealUrl:p.url||'https://www.aldi.com.au/en/special-buys/',
            externalId:'aldi_ld_'+(p.sku||i), scrapedAt:now() });
          found = true;
        });
      } catch(e){}
    });
    if (!found) {
      const ndM = aldiHtml.match(/<script id="__NEXT_DATA__"[^>]*>([\s\S]*?)<\/script>/);
      if (ndM) try {
        const pp = (JSON.parse(ndM[1]).props||{}).pageProps||{};
        (pp.products||pp.items||pp.specialBuys||[]).slice(0,40).forEach((p,i) => {
          if (!p.name&&!p.title) return;
          deals.push({ vendor:'ALDI', category:'aldi', title:p.name||p.title,
            description:'ALDI Special Buy',
            priceNow:p.price?parseFloat(p.price):'', priceWas:'', discountPct:'',
            dealUrl:p.url||'https://www.aldi.com.au/en/special-buys/',
            externalId:'aldi_nd_'+i, scrapedAt:now() });
          found = true;
        });
      } catch(e){}
    }
    if (!found) {
      const h3s=[], prs=[];
      let m;
      const h3Re=/<h3[^>]*>([^<]{5,80})<\/h3>/gi, prRe=/\$(\d[\d.]*)/g;
      while((m=h3Re.exec(aldiHtml))!==null) h3s.push(m[1].trim());
      while((m=prRe.exec(aldiHtml))!==null) prs.push(m[1]);
      h3s.slice(0,25).forEach((t,i) => deals.push({ vendor:'ALDI', category:'aldi',
        title:t, description:'ALDI Special Buy',
        priceNow:prs[i]?parseFloat(prs[i]):'', priceWas:'', discountPct:'',
        dealUrl:'https://www.aldi.com.au/en/special-buys/',
        externalId:'aldi_re_'+i, scrapedAt:now() }));
    }
    console.log('ALDI: '+deals.filter(d=>d.vendor==='ALDI').length);
  }

  // 3. Woolworths/Coles via ScraperAPI (if key set)
  const sk = PropertiesService.getScriptProperties().getProperty('SCRAPER_API_KEY')||'';
  if (sk) {
    const wwData = getJson('https://api.scraperapi.com/?api_key='+sk+'&url='+
      encodeURIComponent('https://www.woolworths.com.au/apis/ui/browse/category?categoryId=1_E5BEE36E&pageNumber=1&pageSize=36&sortType=TraderRelevance&url=%2Fshop%2Fspecials%2Fall&includeChildren=true&inStockProductsOnly=false&filters=%5B%5D'));
    if (wwData) {
      const items=[];
      (wwData.Bundles||wwData.Products||[]).forEach(b=>b.Products?items.push(...b.Products):items.push(b));
      items.slice(0,36).forEach((p,i) => {
        const info=p.Product||p; if (!info.Name) return;
        deals.push({vendor:'Woolworths',category:'woolworths',title:info.Name,
          description:info.CupString||'',
          priceNow:info.Price?parseFloat(info.Price):'',
          priceWas:info.WasPrice?parseFloat(info.WasPrice):'',
          discountPct:info.PercentageSaved?parseInt(info.PercentageSaved):'',
          dealUrl:'https://www.woolworths.com.au/shop/productdetails/'+(info.Stockcode||''),
          externalId:'ww_'+(info.Stockcode||i), scrapedAt:now()});
      });
      console.log('Woolworths (ScraperAPI): '+deals.filter(d=>d.category==='woolworths').length);
    }
    const colesData = getJson('https://api.scraperapi.com/?api_key='+sk+'&url='+
      encodeURIComponent('https://www.coles.com.au/api/2.0/page/search?slug=on-special&page=1&pageSize=36'));
    if (colesData) {
      const results=colesData.results||(colesData.data&&colesData.data.results)||[];
      results.slice(0,36).forEach((item,i) => {
        const p=item._source||item, pr=p.pricing||{}; if (!p.name) return;
        deals.push({vendor:'Coles',category:'coles',title:p.name,description:p.description||'',
          priceNow:pr.now?parseFloat(pr.now):'',priceWas:pr.was?parseFloat(pr.was):'',
          discountPct:pr.savePercent?parseInt(pr.savePercent):'',
          dealUrl:'https://www.coles.com.au/product/'+(p.slug||p.id||''),
          externalId:'coles_'+(p.id||i), scrapedAt:now()});
      });
      console.log('Coles (ScraperAPI): '+deals.filter(d=>d.category==='coles').length);
    }
  }

  if (deals.length) writeDeals(ss, dedupeDeals(deals,'externalId'));
  logSingle(ss,'Groceries',deals.length,Date.now()-t0);
}

function runRetail() {
  const t0 = Date.now(), ss = SpreadsheetApp.openById(SHEET_ID);
  const deals = [];
  const seen  = new Set();

  // Split into TWO batches of ~8 feeds each to avoid Apps Script timeout.
  // 17 feeds in one fetchAll reliably times out; 8-9 is safe.
  const batch1 = [
    { url:'https://www.ozbargain.com.au/cat/computing/feed',             cat:'retail',  label:'Computing' },
    { url:'https://www.ozbargain.com.au/cat/gaming/feed',                cat:'retail',  label:'Gaming' },
    { url:'https://www.ozbargain.com.au/cat/electrical-appliances/feed', cat:'retail',  label:'Electrical' },
    { url:'https://www.ozbargain.com.au/cat/home-garden/feed',           cat:'retail',  label:'Home' },
    { url:'https://www.ozbargain.com.au/cat/clothing/feed',              cat:'retail',  label:'Clothing' },
    { url:'https://www.ozbargain.com.au/cat/health-beauty/feed',         cat:'chemist', label:'Health' },
    { url:'https://www.ozbargain.com.au/cat/baby-kids/feed',             cat:'retail',  label:'Baby' },
    { url:'https://www.ozbargain.com.au/cat/sports-outdoors/feed',       cat:'retail',  label:'Sports' },
  ];
  const batch2 = [
    { url:'https://www.ozbargain.com.au/cat/entertainment/feed',         cat:'entertainment', label:'Entertainment' },
    { url:'https://www.ozbargain.com.au/cat/automotive/feed',            cat:'retail',  label:'Auto' },
    { url:'https://www.ozbargain.com.au/cat/food-drink/feed',            cat:'food',    label:'Food' },
    { url:'https://www.ozbargain.com.au/cat/finance/feed',               cat:'finance', label:'Finance' },
    { url:'https://www.ozbargain.com.au/deals/feed',                     cat:'retail',  label:'All' },
  ];

  function processBatch(feeds) {
    let responses = [];
    try {
      responses = UrlFetchApp.fetchAll(feeds.map(f => ({
        url: f.url, method: 'get',
        headers: { 'User-Agent': UA, 'Accept-Language': 'en-AU,en;q=0.9' },
        muteHttpExceptions: true, followRedirects: true,
      })));
    } catch(e) { console.log('Retail fetchAll error: ' + e.message); return; }

    responses.forEach((resp, fi) => {
      if (!resp || resp.getResponseCode() !== 200) {
        console.log(feeds[fi].label + ': HTTP ' + (resp ? resp.getResponseCode() : 'err'));
        return;
      }
      const items = rssItems(resp.getContentText());
      console.log(feeds[fi].label + ': ' + items.length + ' items');
      items.forEach((item, i) => {
        const d = parseOzbItem(item, feeds[fi].label + '_' + i, feeds[fi].cat);
        if (!d || seen.has(d.externalId)) return;
        seen.add(d.externalId);
        deals.push(d);
      });
    });
  }

  processBatch(batch1);
  processBatch(batch2);

  console.log('Retail total: ' + deals.length);
  if (deals.length) writeDeals(ss, deals);
  logSingle(ss, 'Retail', deals.length, Date.now()-t0);
}

function runCoupons() {
  const t0 = Date.now(), ss = SpreadsheetApp.openById(SHEET_ID);
  const coupons = [];
  rssItems(get('https://www.ozbargain.com.au/deals/feed')).slice(0,100).forEach((item,i) => {
    const title=item.getChildText('title')||'', link=item.getChildText('link')||'';
    const desc=stripHtml(item.getChildText('description')||'').slice(0,200);
    const combined=(title+' '+desc).toLowerCase();
    const vendor = detectVendor(combined);
    const codeM=(title+' '+desc).match(/\b([A-Z][A-Z0-9]{3,15})\b/);
    const discM=(title+' '+desc).match(/(\d+%\s*off|save\s*\$[\d.]+|\$[\d.]+\s*off)/i);
    const code=codeM?codeM[1]:'SEE_LINK';
    if (code==='SEE_LINK'&&!discM) return;
    coupons.push({
      vendor:vendor?vendor.name:'Various', code,
      description:title.slice(0,150), discountText:discM?discM[1]:'',
      minSpend:'', dealUrl:link, source:'ozbargain',
      externalId:'ozb_cpn_'+i, scrapedAt:now()
    });
  });
  if (coupons.length) writeCoupons(ss,coupons);
  logSingle(ss,'Coupons',coupons.length,Date.now()-t0);
}

function runInsurance() {
  const t0 = Date.now(), ss = SpreadsheetApp.openById(SHEET_ID);
  const offers = [];
  ['https://www.ozbargain.com.au/cat/finance/feed','https://www.ozbargain.com.au/deals/feed'].forEach(url => {
    rssItems(get(url)).slice(0,100).forEach((item,i) => {
      const title=item.getChildText('title')||'', link=item.getChildText('link')||'';
      const desc=stripHtml(item.getChildText('description')||'').slice(0,220);
      const combined=(title+' '+desc).toLowerCase();
      const insType=detectInsType(combined); if (!insType) return;
      offers.push(makeIns(detectInsurer(combined),insType,title,desc,
        xPct(title+' '+desc),xDollar(title+' '+desc),link,'ozb_ins_'+url.length+'_'+i));
    });
  });
  offers.push(...getHardcodedInsurance());
  const seen=new Set(), final=[];
  offers.forEach(o => {
    const k=(o.title||'').toLowerCase().slice(0,50);
    if (!k||seen.has(k)) return; seen.add(k); final.push(o);
  });
  if (final.length) writeInsurance(ss,final);
  logSingle(ss,'Insurance',final.length,Date.now()-t0);
}

function runTelco() {
  const t0 = Date.now(), ss = SpreadsheetApp.openById(SHEET_ID);
  const deals = [];
  const seen  = new Set();

  // Dedicated telco/NBN/energy feeds — runs separately from runRetail
  // so these never get pushed off the sheet by retail overwrites
  const feeds = [
    { url:'https://www.ozbargain.com.au/cat/mobile/feed',   cat:'telco',   label:'Mobile' },
    { url:'https://www.ozbargain.com.au/cat/internet/feed', cat:'nbn',     label:'Internet' },
    { url:'https://www.ozbargain.com.au/cat/finance/feed',  cat:'finance', label:'Finance' },
  ];

  // Telco/NBN/energy keywords for filtering the finance feed
  const TELCO_KW = [
    'telstra','optus','vodafone','aldi mobile','boost mobile','kogan mobile',
    'lebara','amaysim','tpg','iinet','internode','aussie broadband','superloop',
    'spintel','skymesh','starlink','nbn','sim plan','mobile plan','prepaid',
    'postpaid','broadband','internet plan','unlimited data','5g plan','5g home',
    'agl','origin energy','energy australia','electricity plan','gas plan',
    'power bill','reamped','powershop','alinta','amber electric','solar rebate',
    'energy offer','energy deal',
  ];

  let responses = [];
  try {
    responses = UrlFetchApp.fetchAll(feeds.map(f => ({
      url: f.url, method: 'get',
      headers: { 'User-Agent': UA, 'Accept-Language': 'en-AU,en;q=0.9' },
      muteHttpExceptions: true, followRedirects: true,
    })));
  } catch(e) { console.log('Telco fetchAll error: ' + e.message); }

  responses.forEach((resp, fi) => {
    if (!resp || resp.getResponseCode() !== 200) {
      console.log(feeds[fi].label + ': HTTP ' + (resp ? resp.getResponseCode() : 'err'));
      return;
    }
    const items = rssItems(resp.getContentText());
    console.log(feeds[fi].label + ': ' + items.length + ' items');
    items.forEach((item, i) => {
      const title = item.getChildText('title') || '';
      const link  = item.getChildText('link')  || '';
      const desc  = stripHtml(item.getChildText('description') || '');
      const low   = (title + ' ' + desc).toLowerCase();
      if (!title) return;

      // Finance feed: only keep telco/energy items
      if (fi === 2 && !TELCO_KW.some(k => low.includes(k))) return;

      const key = title.slice(0, 40).toLowerCase();
      if (seen.has(key)) return; seen.add(key);

      // Detect vendor — if not matched, use feed category as fallback
      const vendor = detectVendor(low) || { name: feeds[fi].cat === 'nbn' ? 'NBN Deal' : 'Mobile Plan', cat: feeds[fi].cat };
      const p   = xPrice(title + ' ' + desc);
      const pct = xPct(title + ' ' + desc);

      deals.push({
        vendor:      vendor.name,
        category:    vendor.cat,
        title:       title.slice(0, 150),
        description: desc.slice(0, 250),
        priceNow:    p   ? parseFloat(p)   : '',
        priceWas:    '',
        discountPct: pct ? parseInt(pct)   : '',
        dealUrl:     link,
        externalId:  'telco_' + feeds[fi].label + '_' + i,
        scrapedAt:   now(),
      });
    });
  });

  // Always add hardcoded current plans — these cover official carrier pages
  // that OzBargain never captures (long-term prepaid, signup deals etc.)
  getHardcodedTelcoPlans().forEach(p => {
    if (!seen.has(p.externalId)) {
      seen.add(p.externalId);
      deals.push(p);
    }
  });

  console.log('Telco/NBN/Energy total (incl. hardcoded): ' + deals.length);
  if (deals.length) writeDeals(ss, dedupeDeals(deals, 'externalId'));
  logSingle(ss, 'Telco/NBN', deals.length, Date.now()-t0);
}

// ── Hardcoded telco/NBN/internet plans ──────────────────────────
// IMPORTANT: No prices hardcoded here — carrier pricing changes constantly
// (e.g. Vodafone $350 plan is currently discounted to $219 but that changes).
// We show plan names and direct links only. User sees live price on carrier site.
// Update titles/descriptions/URLs here when plans are discontinued or renamed.
// Last reviewed: May 2026
function getHardcodedTelcoPlans() {
  const t = now();
  // price = null means "check current price on carrier website"
  function mp(vendor, cat, title, desc, url, id) {
    return { vendor, category: cat, title, description: desc,
      priceNow: '', priceWas: '', discountPct: '',
      dealUrl: url, externalId: 'hc_telco_' + id, scrapedAt: t };
  }

  return [
    // ── VODAFONE PREPAID ───────────────────────────────────────
    mp('Vodafone','telco',
      'Vodafone Prepaid Plus — 365-Day Expiry, Large Data Inclusion',
      'Vodafone long-expiry prepaid: large data valid for 365 days. Currently discounted — check site for live price. Great for light users wanting annual coverage.',
      'https://www.vodafone.com.au/prepaid/plans/350-plus',
      'voda_350plus'),

    mp('Vodafone','telco',
      'Vodafone Prepaid — All Long Expiry Plans (180 & 365 Day)',
      'All Vodafone long-expiry prepaid plans. Prices subject to promotional discounts — always check the live page before purchasing.',
      'https://www.vodafone.com.au/prepaid/plans',
      'voda_prepaid_all'),

    mp('Vodafone','telco',
      'Vodafone SIM-Only — Monthly No Lock-In Plans',
      'Vodafone monthly SIM-only with international calls/texts included. Check current pricing and any active promotions on site.',
      'https://www.vodafone.com.au/mobile/sim-only-plans',
      'voda_simonly'),

    // ── TELSTRA PREPAID ────────────────────────────────────────
    mp('Telstra','telco',
      'Telstra Prepaid — Long Life Plans (180 & 365 Day)',
      'Telstra long-life prepaid on Australia\'s largest network. Best regional VIC coverage. Check live prices — often promotional discounts available.',
      'https://www.telstra.com.au/mobile-phones/prepaid-mobiles/plans',
      'telstra_prepaid_longlife'),

    mp('Telstra','telco',
      'Telstra SIM-Only Monthly Plans — No Lock-In',
      'Telstra SIM-only plans. Best network for regional areas. Check current pricing — Telstra runs frequent new customer promotions.',
      'https://www.telstra.com.au/mobile-phones/mobile-plans',
      'telstra_simonly'),

    // ── OPTUS PREPAID ──────────────────────────────────────────
    mp('Optus','telco',
      'Optus Prepaid — 365-Day Expiry Plans',
      'Optus long-expiry prepaid with 365-day validity. Check live prices — promotional discounts apply regularly.',
      'https://www.optus.com.au/mobile/prepaid/prepaid-plans',
      'optus_prepaid_365'),

    mp('Optus','telco',
      'Optus SIM-Only Plans — Monthly No Lock-In',
      'Optus SIM-only plans with Optus Sport streaming. Bundle savings with Optus NBN. Check current pricing on site.',
      'https://www.optus.com.au/mobile/mobile-plans',
      'optus_simonly'),

    // ── ALDI MOBILE ────────────────────────────────────────────
    mp('ALDI Mobile','telco',
      'ALDI Mobile Prepaid — Telstra Network at Budget Prices',
      'ALDI Mobile on Telstra network with rollover data. No lock-in. Check current plan prices — often very competitive vs Telstra direct.',
      'https://www.aldimobile.com.au/pages/plans',
      'aldi_mobile_plans'),

    mp('ALDI Mobile','telco',
      'ALDI Mobile Long Expiry — 365-Day Recharge Options',
      'ALDI Mobile annual recharge on Telstra network. Check live prices — great value for light users wanting 12 months coverage.',
      'https://www.aldimobile.com.au/pages/long-expiry-plans',
      'aldi_mobile_365'),

    // ── BOOST MOBILE ──────────────────────────────────────────
    mp('Boost Mobile','telco',
      'Boost Mobile — Annual Prepaid Plans on Telstra Network',
      'Boost Mobile annual prepaid on Telstra network. Usually cheaper than Telstra direct for same network. Check current prices.',
      'https://www.boost.com.au/pages/plans',
      'boost_annual'),

    // ── KOGAN MOBILE ──────────────────────────────────────────
    mp('Kogan Mobile','telco',
      'Kogan Mobile — Large Data Prepaid Plans (Vodafone Network)',
      'Kogan Mobile on Vodafone network. Frequent sale pricing. Check current deals — prices fluctuate regularly with promotions.',
      'https://www.koganmobile.com.au/plans/',
      'kogan_plans'),

    // ── LEBARA ────────────────────────────────────────────────
    mp('Lebara Mobile','telco',
      'Lebara Mobile — International Calls Included (Optus Network)',
      'Lebara includes international calls to 40+ countries. Great for calling overseas from Melbourne. Check current plan pricing.',
      'https://www.lebara.com.au/prepaid-plans',
      'lebara_plans'),

    // ── AMAYSIM ───────────────────────────────────────────────
    mp('amaysim','telco',
      'amaysim — Unlimited Data Plans (Optus Network)',
      'amaysim unlimited data on Optus network with data rollover and no lock-in. Check current pricing on site.',
      'https://www.amaysim.com.au/plans/mobile-plans',
      'amaysim_plans'),

    // ── NBN ───────────────────────────────────────────────────
    mp('Aussie Broadband','nbn',
      'Aussie Broadband — ACCC #1 Rated for Customer Satisfaction',
      'Rated #1 by ACCC 2023-24. No throttling, Melbourne-based support. Check current signup deals — often includes modem or bonus months.',
      'https://www.aussiebb.com.au/nbn',
      'abb_nbn'),

    mp('Superloop','nbn',
      'Superloop NBN — Frequently Offers Months Free on Signup',
      'Superloop often includes months free or modem on new connections. Check current promotional pricing before signing up.',
      'https://www.superloop.com/consumer/nbn.html',
      'superloop_nbn'),

    mp('TPG','nbn',
      'TPG NBN — Budget No Lock-In Plans',
      'TPG budget NBN plans. No lock-in contracts. Check current pricing — competitive on 50Mbps and 100Mbps tiers.',
      'https://www.tpg.com.au/nbn',
      'tpg_nbn'),

    mp('iiNet','nbn',
      'iiNet NBN — Reliable with Local Support',
      'iiNet NBN with Toolbox home network tools. Australian-based support. Check current plan pricing and any new customer offers.',
      'https://www.iinet.net.au/internet/nbn/',
      'iinet_nbn'),

    mp('Telstra','nbn',
      'Telstra NBN — Bundle Discount with Telstra Mobile',
      'Telstra NBN with bundle discount when combined with a Telstra mobile plan. Check for current signup offers including months free.',
      'https://www.telstra.com.au/internet/nbn',
      'telstra_nbn'),

    mp('Optus','nbn',
      'Optus NBN & 5G Home Internet',
      'Optus NBN and 5G Home Internet (no nbn line needed in some areas). Multi-service discount with Optus mobile. Check current pricing.',
      'https://www.optus.com.au/internet',
      'optus_nbn'),

    mp('Starlink','nbn',
      'Starlink Satellite Internet — Best for Regional & Rural Areas',
      'SpaceX Starlink. No data cap. Best option for areas without reliable NBN. Check current hardware and monthly pricing.',
      'https://www.starlink.com/au',
      'starlink_au'),

    // ── ENERGY ────────────────────────────────────────────────
    mp('ReAmped Energy','energy',
      'ReAmped Energy — Consistently Cheapest Electricity in AU',
      'Repeatedly ranks among cheapest on Energy Made Easy. Flat-rate, no lock-in, no exit fees. Check current rates for VIC on their site.',
      'https://www.reamped.com.au',
      'reamped_elec'),

    mp('Amber Electric','energy',
      'Amber Electric — Wholesale Electricity Pricing',
      'Amber passes through wholesale electricity prices in real time. Best for solar + battery homes. Pricing varies with the wholesale market.',
      'https://www.amber.com.au',
      'amber_elec'),

    mp('AGL','energy',
      'AGL — Electricity & Gas with Solar Buyback',
      'AGL electricity, gas and solar buyback plans. Check current new customer discounts — rates and offers change frequently.',
      'https://www.agl.com.au/electricity-gas/plans',
      'agl_plans'),

    mp('Origin Energy','energy',
      'Origin Energy — Spike Rewards for Off-Peak Usage',
      'Origin electricity and gas with Spike rewards for off-peak shifting. Check current plan rates — always compare via Energy Made Easy first.',
      'https://www.origin.com.au/plans/electricity',
      'origin_plans'),
  ];
}

function runFuel() {
  const t0 = Date.now(), ss = SpreadsheetApp.openById(SHEET_ID);
  let prices = [];

  // Source 1: VIC Fuel Price Direct API (needs key from https://fuelprices.nsw.gov.au)
  const apiKey = PropertiesService.getScriptProperties().getProperty('VIC_FUEL_API_KEY')||'';
  if (apiKey) {
    const data = getJson('https://fppdirectapi-prod.azurewebsites.net/Subscriber/GetFullSiteDetails?countryId=21&geoRegionLevel=3&geoRegionId=4',
      {'Authorization':'FPDAPI SubscriberToken='+apiKey});
    if (data) {
      (data.S||[]).forEach(site=>(site.Prices||[]).forEach(p=>{
        const c=parseFloat(p.Price)/10; if (c<80||c>350) return;
        prices.push({stationName:site.N||'',brand:site.Brand||guessBrand(site.N||''),
          fuelType:normFuel(p.FuelId),priceCents:c,suburb:site.Suburb||'',
          postcode:site.Postcode||'',state:'VIC',scrapedAt:now()});
      }));
      console.log('VIC Fuel API: '+prices.length);
    }
  }

  // Source 2: NSW FuelCheck API (free, no key needed for basic access)
  if (prices.length === 0) {
    try {
      const nswData = getJson('https://api.onegov.nsw.gov.au/FuelPriceCheck/v1/fuel/prices/bylocation?latitude=-37.90&longitude=144.75&radius=15&fueltype=P95&sortby=price&sortascending=true',
        {'apikey': 'empty', 'transactionid': 'ssc_' + Date.now(), 'requesttimestamp': new Date().toISOString()});
      if (nswData && nswData.stations) {
        nswData.stations.forEach((s, i) => {
          if (i > 30) return;
          prices.push({
            stationName: s.name || s.stationname || '', brand: s.brand || guessBrand(s.name || ''),
            fuelType: normFuel(s.fueltype || 'ULP'), priceCents: parseFloat(s.price) || 0,
            suburb: s.suburb || '', postcode: s.postcode || '', state: 'VIC', scrapedAt: now()
          });
        });
        console.log('NSW API fallback: ' + prices.length);
      }
    } catch(e) { console.log('NSW Fuel API: ' + e.message); }
  }

  // Source 3: Scrape Motormouth RSS (free, public)
  if (prices.length === 0) {
    try {
      const mmText = get('https://www.motormouth.com.au/rss/fuel-prices-3030.xml');
      if (mmText) {
        const items = rssItems(mmText);
        items.forEach((item, i) => {
          const title = item.getChildText('title') || '';
          const desc = item.getChildText('description') || '';
          const priceMatch = desc.match(/([\d.]+)\s*c/i);
          const typeMatch = title.match(/(ULP|E10|P95|P98|Diesel|LPG)/i);
          if (priceMatch && typeMatch) {
            prices.push({
              stationName: title.replace(/\s*[-–]\s*\d.*/, '').trim(),
              brand: guessBrand(title), fuelType: normFuel(typeMatch[1]),
              priceCents: parseFloat(priceMatch[1]),
              suburb: 'Point Cook', postcode: '3030', state: 'VIC', scrapedAt: now()
            });
          }
        });
        console.log('Motormouth RSS: ' + prices.length);
      }
    } catch(e) { console.log('Motormouth: ' + e.message); }
  }

  // Source 4: PetrolSpy API (free, public, covers VIC)
  if (prices.length === 0) {
    try {
      const psData = getJson('https://petrolspy.com.au/webservice-1/station/box?neLat=-37.80&neLng=144.82&swLat=-37.95&swLng=144.68');
      if (psData && psData.message && psData.message.list) {
        psData.message.list.forEach(s => {
          Object.entries(s.prices || {}).forEach(([fuelId, priceObj]) => {
            const p = parseFloat(priceObj.amount);
            if (!p || p < 80 || p > 350) return;
            prices.push({
              stationName: s.name||'', brand: s.brand||guessBrand(s.name||''),
              fuelType: normFuel(fuelId), priceCents: p,
              suburb: s.suburb||'Point Cook area', postcode: s.postcode||'3030',
              state:'VIC', scrapedAt: now()
            });
          });
        });
        console.log('PetrolSpy: ' + prices.length);
      }
    } catch(e) { console.log('PetrolSpy: ' + e.message); }
  }

  // Source 5: Fallback hardcoded (last resort — clearly marked as estimates)
  if (prices.length === 0) {
    console.log('All fuel APIs failed. Using hardcoded fallback. For live prices, set VIC_FUEL_API_KEY in Script Properties.');
    console.log('Free key signup: https://www.fuelcheck.nsw.gov.au/app/fuel-api-subscriber');
    getFuelFallback().forEach(f => prices.push(f));
  }

  const seen={},out=[];
  prices.forEach(p=>{
    if (!p.priceCents||!p.fuelType||p.priceCents<80||p.priceCents>350) return;
    const k=(p.stationName+'|'+p.fuelType).toLowerCase();
    if (!seen[k]){seen[k]=true;out.push(p);}
  });
  if (out.length) writeFuel(ss,out);
  logSingle(ss,'Fuel',out.length,Date.now()-t0);
}

// ── CheapShark — FREE gaming deals API (no key needed) ────────
function runCheapShark() {
  const t0 = Date.now(), ss = SpreadsheetApp.openById(SHEET_ID);
  const deals = [];
  // CheapShark returns PC game deals with real prices from official stores
  const stores = [
    {id:'1', name:'Steam'}, {id:'7', name:'GOG'}, {id:'11', name:'Humble Store'},
    {id:'13', name:'Uplay'}, {id:'25', name:'Epic Games Store'}, {id:'3', name:'Green Man Gaming'},
  ];
  stores.forEach(store => {
    const data = getJson('https://www.cheapshark.com/api/1.0/deals?storeID=' + store.id 
      + '&upperPrice=50&pageSize=15&sortBy=Deal+Rating&onSale=1');
    if (!data || !data.length) return;
    console.log('CheapShark ' + store.name + ': ' + data.length + ' deals');
    data.forEach((d, i) => {
      if (!d.title) return;
      const pNow = parseFloat(d.salePrice) || null;
      const pWas = parseFloat(d.normalPrice) || null;
      const pct  = (pNow && pWas && pWas > pNow) ? Math.round((pWas - pNow) / pWas * 100) : '';
      deals.push({
        vendor: store.name, category: 'gaming', title: d.title,
        description: store.name + ' deal. Steam rating: ' + (d.steamRatingText || 'N/A') + '. Metacritic: ' + (d.metacriticScore || 'N/A'),
        priceNow: pNow ? ('$' + pNow.toFixed(2) + ' USD') : '',
        priceWas: pWas ? ('$' + pWas.toFixed(2) + ' USD') : '',
        discountPct: pct, dealUrl: 'https://www.cheapshark.com/redirect?dealID=' + d.dealID,
        externalId: 'cs_' + d.dealID, scrapedAt: now()
      });
    });
  });
  console.log('CheapShark total: ' + deals.length);
  if (deals.length) writeDeals(ss, deals);
  logSingle(ss, 'CheapShark', deals.length, Date.now() - t0);
}

// ── Amazon Product Advertising API (PAAPI 5.0) ───────────────
// Requires: Script Properties → AMAZON_ACCESS_KEY, AMAZON_SECRET_KEY, AMAZON_PARTNER_TAG
// Sign up free at: https://affiliate-program.amazon.com.au
// It takes 3 days to approve. Once approved, add your keys to Script Properties.
function runAmazonPAAPI() {
  const t0 = Date.now(), ss = SpreadsheetApp.openById(SHEET_ID);
  const props = PropertiesService.getScriptProperties();
  const accessKey = props.getProperty('AMAZON_ACCESS_KEY');
  const secretKey = props.getProperty('AMAZON_SECRET_KEY');
  const partnerTag = props.getProperty('AMAZON_PARTNER_TAG');
  
  if (!accessKey || !secretKey || !partnerTag) {
    console.log('Amazon PAAPI: No API keys configured. Set AMAZON_ACCESS_KEY, AMAZON_SECRET_KEY, AMAZON_PARTNER_TAG in Script Properties.');
    console.log('Sign up free at: https://affiliate-program.amazon.com.au');
    // Fall back to OzBargain Amazon tag
    runAmazonFallback(ss, t0);
    return;
  }
  
  const categories = [
    'Electronics', 'Computers', 'HomeAndGarden', 'Kitchen',
    'Toys', 'SportsAndOutdoors', 'Beauty', 'Books',
  ];
  const deals = [];

  categories.forEach(cat => {
    try {
      const payload = JSON.stringify({
        SearchIndex: cat,
        BrowseNodeId: '',
        ItemCount: 10,
        PartnerTag: partnerTag,
        PartnerType: 'Associates',
        Resources: [
          'ItemInfo.Title', 'ItemInfo.Features',
          'Offers.Listings.Price', 'Offers.Listings.SavingBasis',
          'Offers.Listings.MerchantInfo', 'Images.Primary.Medium'
        ]
      });
      
      // PAAPI 5.0 requires AWS Signature V4
      const host = 'webservices.amazon.com.au';
      const path = '/paapi5/searchitems';
      const dt = new Date();
      const amzDate = Utilities.formatDate(dt, 'UTC', "yyyyMMdd'T'HHmmss'Z'");
      const dateStamp = Utilities.formatDate(dt, 'UTC', 'yyyyMMdd');
      const region = 'us-west-2';
      const service = 'ProductAdvertisingAPI';
      
      // Simplified signing — Apps Script doesn't have native AWS signing
      // Use the direct JSON endpoint with basic auth
      const resp = UrlFetchApp.fetch('https://' + host + path, {
        method: 'post',
        contentType: 'application/json',
        headers: {
          'Host': host,
          'X-Amz-Date': amzDate,
          'X-Amz-Target': 'com.amazon.paapi5.v1.ProductAdvertisingAPIv1.SearchItems',
          'Content-Encoding': 'amz-1.0',
          'Authorization': 'AWS4-HMAC-SHA256 Credential=' + accessKey + '/' + dateStamp + '/' + region + '/' + service + '/aws4_request',
        },
        payload: payload,
        muteHttpExceptions: true,
      });
      
      if (resp.getResponseCode() !== 200) {
        console.log('PAAPI ' + cat + ': HTTP ' + resp.getResponseCode());
        return;
      }
      
      const data = JSON.parse(resp.getContentText());
      if (!data.SearchResult || !data.SearchResult.Items) return;
      
      data.SearchResult.Items.forEach((item, i) => {
        const title = item.ItemInfo && item.ItemInfo.Title && item.ItemInfo.Title.DisplayValue;
        if (!title) return;
        const listing = item.Offers && item.Offers.Listings && item.Offers.Listings[0];
        const pNow = listing && listing.Price && listing.Price.Amount;
        const pWas = listing && listing.SavingBasis && listing.SavingBasis.Amount;
        const pct = (pNow && pWas && pWas > pNow) ? Math.round((pWas - pNow) / pWas * 100) : '';
        
        deals.push({
          vendor: 'Amazon AU', category: 'amazon',
          title: title.slice(0, 150),
          description: (item.ItemInfo && item.ItemInfo.Features && item.ItemInfo.Features.DisplayValues)
            ? item.ItemInfo.Features.DisplayValues.slice(0, 3).join('. ').slice(0, 200)
            : 'Amazon AU ' + cat + ' deal',
          priceNow: pNow || '', priceWas: pWas || '',
          discountPct: pct, dealUrl: item.DetailPageURL || '',
          externalId: 'amz_' + (item.ASIN || i), scrapedAt: now()
        });
      });
    } catch(e) {
      console.log('PAAPI ' + cat + ' error: ' + e.message);
    }
  });

  console.log('Amazon PAAPI total: ' + deals.length);
  if (deals.length) writeDeals(ss, deals);
  logSingle(ss, 'AmazonPAAPI', deals.length, Date.now() - t0);
}

function runAmazonFallback(ss, t0) {
  // When no PAAPI key: try OzBargain Amazon tag feed
  const text = get('https://www.ozbargain.com.au/tag/amazon/feed');
  const items = rssItems(text);
  const deals = [];
  items.forEach((item, i) => {
    const d = parseOzbItem(item, 'amz_ozb_' + i, 'amazon');
    if (d) { d.vendor = 'Amazon AU (via OzBargain)'; deals.push(d); }
  });
  console.log('Amazon fallback (OzBargain): ' + deals.length);
  if (deals.length) writeDeals(ss, deals);
  logSingle(ss, 'AmazonFB', deals.length, Date.now() - t0);
}

// ── Cashback portals — curated best offers ────────────────────
function runCashback() {
  const t0 = Date.now(), ss = SpreadsheetApp.openById(SHEET_ID);
  const deals = [];
  
  // Try ShopBack public deals page
  const sbText = get('https://www.shopback.com.au/cashback-deals');
  if (sbText) {
    // Parse deal snippets from the page
    const titleMatches = sbText.match(/<h[23][^>]*>(.*?)<\/h[23]>/g) || [];
    titleMatches.slice(0, 20).forEach((m, i) => {
      const title = stripHtml(m);
      if (title.length > 10 && title.length < 150) {
        deals.push({
          vendor: 'ShopBack', category: 'retail',
          title: 'ShopBack: ' + title,
          description: 'Earn cashback on this deal via ShopBack. Click through ShopBack before purchasing.',
          priceNow: '', priceWas: '', discountPct: '',
          dealUrl: 'https://www.shopback.com.au/cashback-deals',
          externalId: 'sb_' + i, scrapedAt: now()
        });
      }
    });
  }
  
  // Try Cashrewards deals page
  const crText = get('https://www.cashrewards.com.au/featured-stores');
  if (crText) {
    const titleMatches = crText.match(/<h[23][^>]*>(.*?)<\/h[23]>/g) || [];
    titleMatches.slice(0, 15).forEach((m, i) => {
      const title = stripHtml(m);
      if (title.length > 5 && title.length < 150) {
        deals.push({
          vendor: 'Cashrewards', category: 'retail',
          title: 'Cashrewards: ' + title,
          description: 'Earn cashback via Cashrewards. Always click through Cashrewards before purchasing to track your cashback.',
          priceNow: '', priceWas: '', discountPct: '',
          dealUrl: 'https://www.cashrewards.com.au/featured-stores',
          externalId: 'cr_' + i, scrapedAt: now()
        });
      }
    });
  }
  
  // Try TopCashback deals
  const tcText = get('https://www.topcashback.com.au/trending-offers/');
  if (tcText) {
    const titleMatches = tcText.match(/<h[23][^>]*>(.*?)<\/h[23]>/g) || [];
    titleMatches.slice(0, 15).forEach((m, i) => {
      const title = stripHtml(m);
      if (title.length > 5 && title.length < 150) {
        deals.push({
          vendor: 'TopCashback', category: 'retail',
          title: 'TopCashback: ' + title,
          description: 'Earn cashback via TopCashback AU. Sign up free and click through before buying.',
          priceNow: '', priceWas: '', discountPct: '',
          dealUrl: 'https://www.topcashback.com.au/trending-offers/',
          externalId: 'tc_' + i, scrapedAt: now()
        });
      }
    });
  }
  
  console.log('Cashback total: ' + deals.length);
  if (deals.length) writeDeals(ss, deals);
  logSingle(ss, 'Cashback', deals.length, Date.now() - t0);
}

// ── Travel deals via RSS ──────────────────────────────────────
function runTravel() {
  const t0 = Date.now(), ss = SpreadsheetApp.openById(SHEET_ID);
  const deals = [];
  const feeds = [
    { url: 'https://www.ozbargain.com.au/cat/travel/feed', label: 'OzB Travel' },
  ];
  feeds.forEach(f => {
    const text = get(f.url);
    const items = rssItems(text);
    console.log(f.label + ': ' + items.length);
    items.forEach((item, i) => {
      const d = parseOzbItem(item, 'trv_' + i, 'travel');
      if (d) deals.push(d);
    });
  });
  console.log('Travel total: ' + deals.length);
  if (deals.length) writeDeals(ss, deals);
  logSingle(ss, 'Travel', deals.length, Date.now() - t0);
}

function scrapeAll() {
  runGroceries(); runRetail(); runTravel();
  runTelco(); runCoupons(); runInsurance(); runFuel();
  runCheapShark(); runAmazonPAAPI(); runCashback();
}

// ── UTILITIES ──────────────────────────────────────────────────
function makeIns(insurer,type,title,description,discountPct,saveDollars,dealUrl,externalId){
  return{insurer:insurer||'Various',type:type||'General',
    title:(title||'').slice(0,150).trim(),description:(description||'').slice(0,250).trim(),
    discountPct:discountPct!=null?discountPct:'',saveDollars:saveDollars!=null?saveDollars:'',
    dealUrl:dealUrl||'',externalId:externalId||'',scrapedAt:now()};}
function stripHtml(s){return(s||'').replace(/<[^>]+>/g,' ').replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&nbsp;/g,' ').replace(/&quot;/g,'"').replace(/\s+/g,' ').trim();}
function xPrice(t){const m=(t||'').match(/\$\s*([\d,]+\.?\d*)/);return m?parseFloat(m[1].replace(/,/g,'')):null;}
function xWas(t){const m=(t||'').match(/(?:was|rrp|from)\s*\$\s*([\d,]+\.?\d*)/i);return m?parseFloat(m[1].replace(/,/g,'')):null;}
function xPct(t){const m=(t||'').match(/(\d+)%\s*off/i);return m?parseInt(m[1]):null;}
function xDollar(t){const m=(t||'').match(/save\s*\$\s*([\d,]+)/i);return m?parseFloat(m[1].replace(/,/g,'')):null;}
function dedupeDeals(deals,key){const seen=new Set();return deals.filter(d=>{const k=(d[key]||'').toLowerCase().slice(0,80);if(!k||seen.has(k))return false;seen.add(k);return true;});}
function calcPct(now,was){const n=parseFloat(now),w=parseFloat(was);if(!isNaN(n)&&!isNaN(w)&&w>0&&n<w)return Math.round((w-n)/w*100);return '';}
function now(){return new Date().toISOString();}
function normFuel(raw){const t=String(raw||'').toUpperCase().trim();if(t.includes('98')||t==='6')return 'P98';if(t.includes('95')||t==='5')return 'P95';if(t.includes('E10')||t==='2')return 'E10';if(t.includes('DIESEL')||t==='4'||t==='D')return 'Diesel';if(t.includes('LPG')||t==='3')return 'LPG';return 'ULP';}
function guessBrand(n){n=(n||'').toUpperCase();if(n.includes('BP'))return 'BP';if(n.includes('AMPOL')||n.includes('CALTEX'))return 'Ampol';if(n.includes('SHELL'))return 'Shell';if(n.includes('7-ELEVEN'))return '7-Eleven';if(n.includes('UNITED'))return 'United';if(n.includes('PUMA'))return 'Puma';return '';}
function detectInsType(t){if(t.match(/\bpet\s*insur/))return'Pet';if(t.match(/\bhealth\s*insur|\bprivate\s*health/))return'Health';if(t.match(/\bcar\s*insur|\bcomprehensive\s*car/))return'Car';if(t.match(/\bhome\s*insur|\bcontents\s*insur/))return'Home';if(t.match(/\btravel\s*insur/))return'Travel';if(t.match(/\blife\s*insur|\bterm\s*life/))return'Life';if(t.match(/\bincome\s*protection/))return'Income Protection';if(t.match(/\binsur/))return'General';return null;}
function detectInsurer(t){const map=[['budget direct','Budget Direct'],['nrma','NRMA'],['racv','RACV'],['bupa','Bupa'],['medibank','Medibank'],['hbf','HBF'],['ahm','ahm'],['allianz','Allianz'],['aami','AAMI'],['youi','Youi'],['bow wow meow','Bow Wow Meow'],['petsure','PetSure'],['tal','TAL']];for(const [k,v] of map){if(t.includes(k))return v;}return 'Various';}
function getFuelFallback(){
  // No hardcoded prices — they go stale immediately.
  // Live fuel data comes from: VIC Fuel API, NSW FuelCheck, Motormouth, PetrolSpy.
  // Browser-side PetrolSpy fetch provides real-time prices when users view Fuel page.
  console.log('Fuel fallback: no stale data. Set VIC_FUEL_API_KEY for server-side live prices.');
  console.log('Free signup: https://www.fuelcheck.nsw.gov.au/app/fuel-api-subscriber');
  return [];
}

// ── Hardcoded insurance offers ─────────────────────────────────
// No hardcoded discount percentages — these change constantly with promotions.
// Descriptions mention "check site for current offer" so users know to verify.
// Last reviewed: May 2026
function getHardcodedInsurance(){
  const mk = makeIns;
  return [
    // Car insurance
    mk('Budget Direct','Car',
      'Budget Direct Car Insurance — Online Discount Available',
      'Budget Direct frequently offers online signup discounts. Check their site for the current promotion before purchasing.',
      null,null,'https://www.budgetdirect.com.au/car-insurance.html','hc_car_bd'),
    mk('AAMI','Car',
      'AAMI Car Insurance — Bundle Car & Home for Savings',
      'Combine car and home insurance with AAMI for a multi-policy discount. Check current bundle pricing on their site.',
      null,null,'https://www.aami.com.au/car-insurance.html','hc_car_aami'),
    mk('RACV','Car',
      'RACV Comprehensive Car Insurance — Member Discounts',
      'RACV members receive exclusive discounts on car insurance. Roadside assist included. Check current member pricing.',
      null,null,'https://www.racv.com.au/insurance','hc_car_racv'),
    mk('Youi','Car',
      'Youi Car Insurance — Pay for What You Use',
      'Youi usage-based pricing — you only pay for when you drive. Check their site for a personalised quote.',
      null,null,'https://www.youi.com.au/car-insurance','hc_car_youi'),
    mk('Allianz','Car',
      'Allianz Car Insurance — Online Purchase Discount',
      'Allianz offers a discount when purchasing car insurance online. Check their site for the current online discount percentage.',
      null,null,'https://www.allianz.com.au/car-insurance.html','hc_car_allianz'),
    // Home insurance
    mk('Budget Direct','Home',
      'Budget Direct Home & Contents — Online Purchase Discount',
      'Budget Direct frequently discounts home and contents insurance purchased online. Check current promotion before buying.',
      null,null,'https://www.budgetdirect.com.au/home-insurance.html','hc_home_bd'),
    mk('RACV','Home',
      'RACV Home Insurance — Comprehensive VIC Cover',
      'Covers flood, fire, storm and more. RACV member discounts available. Check current pricing for Point Cook area.',
      null,null,'https://www.racv.com.au/insurance/home-insurance.html','hc_home_racv'),
    mk('Allianz','Home',
      'Allianz Home & Contents — Online Purchase Discount',
      'Allianz offers a discount when buying home and contents insurance online. Check their site for current percentage.',
      null,null,'https://www.allianz.com.au/home-insurance.html','hc_home_allianz'),
    // Health insurance
    mk('Medibank','Health',
      'Medibank Health Insurance — New Member Offer',
      'Medibank frequently offers weeks-free promotions for new members. Check their site for the current new member deal.',
      null,null,'https://www.medibank.com.au/health-insurance/','hc_health_med'),
    mk('Bupa','Health',
      'Bupa Health Insurance — New Member Offer',
      'Bupa runs new member promotions including weeks free on combined cover. Check current offer on their site.',
      null,null,'https://www.bupa.com.au/health-insurance','hc_health_bupa'),
    mk('ahm','Health',
      'ahm Health Insurance — Join Offer',
      'ahm periodically offers gift cards or other incentives for new members. Check their site for the current joining offer.',
      null,null,'https://www.ahm.com.au','hc_health_ahm'),
    mk('RACV','Health',
      'RACV Health Insurance — Member Discounts',
      'RACV members receive discounted health insurance premiums. Check current member pricing and available extras covers.',
      null,null,'https://www.racv.com.au/insurance/health-insurance.html','hc_health_racv'),
    // Pet insurance
    mk('Bow Wow Meow','Pet',
      'Bow Wow Meow Pet Insurance — New Policy Offer',
      'Bow Wow Meow runs new policy promotions. Check their site for current first-month or discount offers.',
      null,null,'https://www.bowwowmeow.com.au','hc_pet_bwm'),
    mk('PetSure','Pet',
      'PetSure — Up to 80% of Vet Bills Covered',
      'PetSure covers accident, illness and more across multiple plan tiers. Check their site for current pricing.',
      null,null,'https://www.petsure.com.au','hc_pet_petsure'),
    // Travel insurance
    mk('Allianz','Travel',
      'Allianz Travel Insurance — Online Purchase Discount',
      'Allianz offers a discount on travel insurance purchased online. Check current percentage and coverage options.',
      null,null,'https://www.allianz.com.au/travel-insurance.html','hc_travel_allianz'),
    mk('Budget Direct','Travel',
      'Budget Direct Travel Insurance — Competitive Daily Rates',
      'Budget Direct travel insurance with medical, cancellation and baggage cover. Check current daily rate pricing.',
      null,null,'https://www.budgetdirect.com.au/travel-insurance.html','hc_travel_bd'),
    // Life and income protection
    mk('Budget Direct','Life',
      'Budget Direct Life Insurance — Online Purchase Discount',
      'Term life insurance with competitive premiums. Online purchase discount available — check current pricing.',
      null,null,'https://www.budgetdirect.com.au/life-insurance.html','hc_life_bd'),
    mk('TAL','Life',
      'TAL Life Insurance — Flexible Cover Options',
      'TAL covers life, TPD, trauma and income protection. Check current pricing and any active promotional offers.',
      null,null,'https://www.tal.com.au','hc_life_tal'),
    mk('TAL','Income Protection',
      'TAL Income Protection — Up to 70% of Income Covered',
      'TAL income protection pays monthly benefits if you cannot work due to illness or injury. Check current pricing.',
      null,null,'https://www.tal.com.au/products/income-protection','hc_ip_tal'),
  ];
}
// ── SHEET HELPERS ──────────────────────────────────────────────
const HEADERS={
  [TAB.DEALS]:    ['scrapedAt','vendor','category','title','description','priceNow','priceWas','discountPct','dealUrl','externalId'],
  [TAB.FUEL]:     ['scrapedAt','stationName','brand','fuelType','priceCents','suburb','postcode','state'],
  [TAB.COUPONS]:  ['scrapedAt','vendor','code','description','discountText','minSpend','dealUrl','source','externalId'],
  [TAB.INSURANCE]:['scrapedAt','insurer','type','title','description','discountPct','saveDollars','dealUrl','externalId'],
  [TAB.LOG]:      ['startedAt','finishedAt','source','count','errors','ms','notes','spare'],
};
function writeDeals(ss,rows){prependRows(ss.getSheetByName(TAB.DEALS),rows.map(d=>HEADERS[TAB.DEALS].map(h=>d[h]??'')),MAX_DEALS);}
function writeFuel(ss,rows){prependRows(ss.getSheetByName(TAB.FUEL),rows.map(f=>HEADERS[TAB.FUEL].map(h=>f[h]??'')),MAX_FUEL);}
function writeCoupons(ss,rows){prependRows(ss.getSheetByName(TAB.COUPONS),rows.map(c=>HEADERS[TAB.COUPONS].map(h=>c[h]??'')),MAX_COUPONS);}
function writeInsurance(ss,rows){prependRows(ss.getSheetByName(TAB.INSURANCE),rows.map(i=>HEADERS[TAB.INSURANCE].map(h=>i[h]??'')),MAX_INSURANCE);}
function prependRows(sheet,rows,maxRows){if(!rows||!rows.length)return;sheet.insertRowsAfter(1,rows.length);sheet.getRange(2,1,rows.length,rows[0].length).setValues(rows);const last=sheet.getLastRow();if(last>maxRows+1)sheet.deleteRows(maxRows+2,last-maxRows-1);}
function getLastRun(ss){try{const sh=ss.getSheetByName(TAB.LOG);if(!sh)return null;const d=sh.getDataRange().getValues();if(d.length<2)return null;const o={};d[0].forEach((k,i)=>{o[k]=d[1][i];});return o;}catch(e){return null;}}
function logSingle(ss,source,count,ms){const sh=ss.getSheetByName(TAB.LOG);if(!sh)return;sh.insertRowAfter(1);sh.getRange(2,1,1,8).setValues([[now(),now(),source,count,0,ms,'OK','']]);const last=sh.getLastRow();if(last>101)sh.deleteRows(102,last-101);}

// ── SETUP ──────────────────────────────────────────────────────
function setupSheets(){const ss=SpreadsheetApp.openById(SHEET_ID);Object.entries(HEADERS).forEach(([name,headers])=>{let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);if(sh.getRange(1,1).getValue()!==headers[0]){sh.getRange(1,1,1,headers.length).setValues([headers]);sh.setFrozenRows(1);sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#f3f3f3');}});const def=ss.getSheetByName('Sheet1');if(def&&ss.getSheets().length>1)ss.deleteSheet(def);SpreadsheetApp.getUi().alert('Sheets ready.');}
function setupTriggers(){
  const fns=['scrapeAll','runGroceries','runRetail','runTelco','runCoupons','runInsurance','runFuel','runCheapShark','runAmazonPAAPI','runCashback','runTravel','sendDailyDigest'];
  ScriptApp.getProjectTriggers().filter(t=>fns.includes(t.getHandlerFunction())).forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendDailyDigest').timeBased().atHour(7).everyDays(1).create();
  ScriptApp.newTrigger('runGroceries').timeBased().everyHours(3).create();
  ScriptApp.newTrigger('runRetail').timeBased().everyHours(3).create();
  ScriptApp.newTrigger('runTelco').timeBased().everyHours(6).create();
  ScriptApp.newTrigger('runCoupons').timeBased().everyHours(4).create();
  ScriptApp.newTrigger('runInsurance').timeBased().everyHours(8).create();
  ScriptApp.newTrigger('runFuel').timeBased().everyHours(6).create(); // runs 4x/day for fresh fuel data
  ScriptApp.newTrigger('runCheapShark').timeBased().everyHours(4).create();
  ScriptApp.newTrigger('runAmazonPAAPI').timeBased().everyHours(3).create();
  ScriptApp.newTrigger('runCashback').timeBased().everyHours(6).create();
  ScriptApp.newTrigger('runTravel').timeBased().everyHours(6).create();
  SpreadsheetApp.getUi().alert('11 triggers created. Daily digest at 7am. CheapShark/Amazon/Cashback included.');
}
function onOpen(){
  SpreadsheetApp.getUi().createMenu('SmartSaver Connect')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Test scrapers')
      .addItem('Groceries (OzBargain + ALDI)','runGroceries')
      .addItem('Retail (two batches)','runRetail')
      .addItem('Telco / NBN / Energy','runTelco')
      .addItem('Coupons','runCoupons')
      .addItem('Insurance','runInsurance')
      .addItem('Fuel','runFuel')
      .addItem('Travel','runTravel')
      .addItem('CheapShark (gaming deals)','runCheapShark')
      .addItem('Amazon PAAPI','runAmazonPAAPI')
      .addItem('Cashback portals','runCashback'))
    .addSeparator()
    .addItem('Run ALL scrapers','scrapeAll')
    .addItem('Setup sheets','setupSheets')
    .addItem('Setup triggers','setupTriggers')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Deal submissions')
      .addItem('View submissions','openSubmissionsSheet')
      .addItem('Approve selected deal','approveSelectedDeal'))
    .addItem('View run log','openRunLog')
    .addItem('View subscribers','openSubsSheet')
    .addItem('Send daily digest now (test)','sendDailyDigest')
    .addToUi();
}
function openSubmissionsSheet(){const ss=SpreadsheetApp.openById(SHEET_ID);let sh=ss.getSheetByName('Submissions');if(!sh){sh=ss.insertSheet('Submissions');sh.getRange(1,1,1,12).setValues([['Submitted','Business','Title','Description','Sale Price','Original Price','Deal URL','Category','Contact Name','Contact Email','Expiry','Status']]);sh.setFrozenRows(1);}ss.setActiveSheet(sh);}
function openRunLog(){const ss=SpreadsheetApp.openById(SHEET_ID);ss.setActiveSheet(ss.getSheetByName(TAB.LOG));}
function openSubsSheet(){const ss=SpreadsheetApp.openById(SHEET_ID);const sh=getOrCreateSubsSheet(ss);ss.setActiveSheet(sh);}
