// ============================================================
// KOL AGENCY OPERATING SYSTEM — Google Apps Script
// Master backend: routing, creator, campaign, assignment, scoring
// ============================================================

// ─── SHEET NAMES ────────────────────────────────────────────
const SHEET = {
  CREATORS:     'CREATORS',
  CAMPAIGNS:    'CAMPAIGNS',
  ASSIGNMENTS:  'ASSIGNMENTS',
  PERFORMANCE:  'PERFORMANCE',
  CONFIG:       'CONFIG',
};

// ─── CORS HEADERS ───────────────────────────────────────────
function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST',
    'Access-Control-Allow-Headers': 'Content-Type',
  };
}

function makeResponse(data, status) {
  const output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ─── ROUTER: GET ─────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action || '';
  let result;

  try {
    switch (action) {
      case 'checkDuplicate':
        result = checkDuplicate(e.parameter.ig);
        break;
      case 'getCreator':
        result = getCreatorForApply(e.parameter.ig, e.parameter.campaignId);
        break;
      case 'getCampaign':
        result = getCampaignPublic(e.parameter.campaignId);
        break;
      case 'getOpenCampaigns':
        result = getOpenCampaigns();
        break;
      // Admin endpoints
      case 'getAllCreators':
        result = getAllCreators();
        break;
      case 'getWaitingList':
        result = getWaitingList(e.parameter.campaignId);
        break;
      case 'getAllCampaigns':
        result = getAllCampaigns();
        break;
      case 'getDashboardStats':
        result = getDashboardStats();
        break;
      default:
        result = { success: false, message: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { success: false, message: err.toString() };
    Logger.log('doGet error: ' + err.toString());
  }

  return makeResponse(result);
}

// ─── ROUTER: POST ────────────────────────────────────────────
function doPost(e) {
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return makeResponse({ success: false, message: 'Invalid JSON payload' });
  }

  const action = payload.action || '';
  let result;

  try {
    switch (action) {
      case 'registerCreator':
        result = registerCreator(payload);
        break;
      case 'applyToCampaign':
        result = applyToCampaign(payload);
        break;
      case 'updateAssignmentStatus':
        result = updateAssignmentStatus(payload);
        break;
      case 'createCampaign':
        result = createCampaign(payload);
        break;
      case 'updateCreatorStatus':
        result = updateCreatorStatus(payload);
        break;
      case 'recalculateScore':
        result = recalculateAllScores();
        break;
      default:
        result = { success: false, message: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { success: false, message: err.toString() };
    Logger.log('doPost error: ' + err.toString());
  }

  return makeResponse(result);
}

// ============================================================
// CREATOR FUNCTIONS
// ============================================================

function checkDuplicate(igUsername) {
  if (!igUsername) return { isDuplicate: false };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.CREATORS);
  const data = sheet.getDataRange().getValues();

  // Col index: 0=creator_id, 1=created_at, 2=updated_at, 3=full_name,
  // 4=ig_username, 5=tiktok_username
  const ig = igUsername.toLowerCase().replace('@', '');
  for (let i = 1; i < data.length; i++) {
    if ((data[i][4] || '').toLowerCase() === ig) {
      return {
        isDuplicate: true,
        message: 'Username IG ini sudah terdaftar.',
      };
    }
  }
  return { isDuplicate: false };
}

function registerCreator(payload) {
  // 1. Dedup check
  const dupCheck = checkDuplicate(payload.igUsername);
  if (dupCheck.isDuplicate) {
    return { success: false, message: 'Username IG sudah terdaftar.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.CREATORS);
  const now = new Date();

  // Creator ID = IG username (lowercase, no @)
  const creatorId = (payload.igUsername || '').toLowerCase().replace('@', '');

  const row = [
    creatorId,                          // creator_id
    now,                                // created_at
    now,                                // updated_at
    payload.fullName || '',             // full_name
    creatorId,                          // ig_username
    (payload.tiktokUsername || '').toLowerCase().replace('@', ''), // tiktok_username
    payload.primaryPlatform || '',      // platform_primary
    parseInt(payload.followersIG) || 0, // followers_ig
    parseInt(payload.followersTT) || 0, // followers_tiktok
    payload.niche || '',                // niche
    payload.domisili || '',             // domisili
    payload.audienceAge || '',          // audience_demographics
    payload.styleCategory || '',        // style_category
    payload.postingFrequency || '',     // posting_frequency
    payload.endorseExperience || '',    // endorse_experience
    payload.availabilityDays || '',     // availability_days
    payload.contactWA || '',            // contact_wa
    payload.contactEmail || '',         // contact_email
    payload.portfolioUrl || '',         // portfolio_url
    'inactive',                         // status (admin must approve)
    '',                                 // blacklist_reason
    0,                                  // score_total (calculated later)
    '',                                 // notes_internal
    'self-apply',                       // source
  ];

  sheet.appendRow(row);
  Logger.log('New creator registered: ' + creatorId);

  return {
    success: true,
    creatorId: creatorId,
    message: 'Pendaftaran berhasil. Tim kami akan review dalam 1-2 hari kerja.',
  };
}

function getCreatorById(creatorId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.CREATORS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toLowerCase() === creatorId.toLowerCase()) {
      return rowToObject(headers, data[i]);
    }
  }
  return null;
}

function getAllCreators() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.CREATORS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const creators = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) creators.push(rowToObject(headers, data[i]));
  }
  return { success: true, data: creators };
}

function updateCreatorStatus(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.CREATORS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === payload.creatorId) {
      // Col 19 = status, Col 20 = blacklist_reason, Col 2 = updated_at
      sheet.getRange(i + 1, 20).setValue(payload.status);
      sheet.getRange(i + 1, 21).setValue(payload.blacklistReason || '');
      sheet.getRange(i + 1, 3).setValue(new Date());
      return { success: true };
    }
  }
  return { success: false, message: 'Creator tidak ditemukan.' };
}

// ============================================================
// CAMPAIGN FUNCTIONS
// ============================================================

function getCampaignPublic(campaignId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.CAMPAIGNS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === campaignId) {
      const c = rowToObject(headers, data[i]);
      // Count filled slots
      const filled = countFilledSlots(campaignId);
      c.slotsFilled = filled;
      c.slotsRemaining = (c.slots_total || 0) - filled;
      c.isOpen = c.status === 'open' && c.slotsRemaining > 0;
      return { success: true, data: c };
    }
  }
  return { success: false, message: 'Campaign tidak ditemukan.' };
}

function getOpenCampaigns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.CAMPAIGNS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const open = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][9] === 'open') { // col 9 = status
      const c = rowToObject(headers, data[i]);
      const filled = countFilledSlots(c.campaign_id);
      c.slotsFilled = filled;
      c.slotsRemaining = (c.slots_total || 0) - filled;
      if (c.slotsRemaining > 0) open.push(c);
    }
  }
  return { success: true, data: open };
}

function getAllCampaigns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.CAMPAIGNS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const campaigns = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      const c = rowToObject(headers, data[i]);
      c.slotsFilled = countFilledSlots(c.campaign_id);
      c.slotsRemaining = (c.slots_total || 0) - c.slotsFilled;
      c.waitingCount = countWaiting(c.campaign_id);
      campaigns.push(c);
    }
  }
  return { success: true, data: campaigns };
}

function createCampaign(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.CAMPAIGNS);
  const now = new Date();

  // Auto-generate campaign ID
  const lastRow = sheet.getLastRow();
  const seq = String(lastRow).padStart(3, '0');
  const year = now.getFullYear().toString().slice(2);
  const campaignId = `CAM-${year}-${seq}`;

  const row = [
    campaignId,
    payload.brandName || '',
    payload.campaignName || '',
    payload.objective || '',
    payload.platform || '',
    payload.deliverables || '',
    parseFloat(payload.budgetTotal) || 0,
    parseInt(payload.creatorSlots) || 1,
    payload.deadlineContent || '',
    payload.deadlinePosting || '',
    'open',               // status
    payload.picAdmin || '',
    payload.briefUrl || '',
    payload.notes || '',
    now,                  // created_at
    now,                  // updated_at
    payload.applyDeadline || '',
    payload.minFollowers || 0,
    payload.targetNiche || '',
  ];

  sheet.appendRow(row);
  return { success: true, campaignId: campaignId };
}

function countFilledSlots(campaignId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.ASSIGNMENTS);
  const data = sheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === campaignId &&
        ['approved', 'assigned', 'delivered', 'posted'].includes(data[i][5])) {
      count++;
    }
  }
  return count;
}

function countWaiting(campaignId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.ASSIGNMENTS);
  const data = sheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === campaignId && data[i][5] === 'waiting') count++;
  }
  return count;
}

// ============================================================
// ASSIGNMENT / APPLY FUNCTIONS
// ============================================================

function getCreatorForApply(igUsername, campaignId) {
  if (!igUsername) return { success: false, message: 'Username diperlukan.' };

  const ig = igUsername.toLowerCase().replace('@', '');
  const creator = getCreatorById(ig);

  if (!creator) return { success: false, found: false, message: 'Creator tidak ditemukan.' };

  // Eligibility checks
  if (creator.status === 'blacklist') {
    return { success: true, found: true, eligible: false, reason: 'blacklist', creator: sanitizeCreator(creator) };
  }
  if (creator.status !== 'active') {
    return { success: true, found: true, eligible: false, reason: 'inactive', creator: sanitizeCreator(creator) };
  }

  // Check active campaign
  const activeCampaign = getActiveCampaignForCreator(ig);
  if (activeCampaign) {
    return {
      success: true, found: true, eligible: false, reason: 'active_campaign',
      activeCampaignName: activeCampaign,
      creator: sanitizeCreator(creator),
    };
  }

  // Check already applied
  if (campaignId && hasApplied(ig, campaignId)) {
    return { success: true, found: true, eligible: false, reason: 'already_applied', creator: sanitizeCreator(creator) };
  }

  // Check slots
  if (campaignId) {
    const camp = getCampaignPublic(campaignId);
    if (camp.success && camp.data.slotsRemaining <= 0) {
      return { success: true, found: true, eligible: false, reason: 'slot_full', creator: sanitizeCreator(creator) };
    }
  }

  return { success: true, found: true, eligible: true, creator: sanitizeCreator(creator) };
}

function sanitizeCreator(c) {
  // Only return non-sensitive fields to public
  return {
    name: c.full_name,
    ig: '@' + c.ig_username,
    followers: formatFollowers(c.followers_ig),
    niche: c.niche,
    score: c.score_total || '—',
    domisili: c.domisili,
    status: c.status,
  };
}

function getActiveCampaignForCreator(creatorId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.ASSIGNMENTS);
  const data = sheet.getDataRange().getValues();
  const activeStatuses = ['approved', 'assigned', 'delivered'];

  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === creatorId && activeStatuses.includes(data[i][5])) {
      // Return campaign name
      return data[i][14] || data[i][1]; // col 14 = campaign_name_cached
    }
  }
  return null;
}

function hasApplied(creatorId, campaignId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.ASSIGNMENTS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === campaignId && data[i][2] === creatorId) return true;
  }
  return false;
}

function applyToCampaign(payload) {
  const ig = (payload.creatorIG || '').toLowerCase().replace('@', '');
  const campaignId = payload.campaignId;

  // Re-check eligibility server-side
  const check = getCreatorForApply(ig, campaignId);
  if (!check.eligible) {
    return { success: false, message: 'Tidak eligible: ' + check.reason };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.ASSIGNMENTS);
  const now = new Date();

  // Get campaign info for denormalization
  const campRes = getCampaignPublic(campaignId);
  const campName = campRes.success ? campRes.data.campaign_name : campaignId;

  // Count waiting position
  const waitPos = countWaiting(campaignId) + 1;

  // Auto-generate assignment ID
  const lastRow = sheet.getLastRow();
  const assignmentId = `ASG-${String(lastRow).padStart(4, '0')}`;

  const row = [
    assignmentId,               // 0: assignment_id
    campaignId,                 // 1: campaign_id
    ig,                         // 2: creator_id
    check.creator.name,         // 3: creator_name (denorm)
    check.creator.ig,           // 4: ig_username (denorm)
    'waiting',                  // 5: status
    0,                          // 6: rate_agreed
    '',                         // 7: deliverable_detail
    '',                         // 8: deadline_draft
    '',                         // 9: deadline_post
    '',                         // 10: draft_submitted_at
    '',                         // 11: posted_at
    '',                         // 12: post_url
    'pending',                  // 13: payment_status
    campName,                   // 14: campaign_name_cached
    now,                        // 15: applied_at
    waitPos,                    // 16: waitlist_position
    payload.reason || '',       // 17: apply_reason
    payload.referenceUrl || '', // 18: reference_url
    payload.availableFrom || '',// 19: available_from
    payload.question || '',     // 20: creator_question
    '',                         // 21: payment_date
    '',                         // 22: notes_admin
  ];

  sheet.appendRow(row);

  return {
    success: true,
    assignmentId: assignmentId,
    waitlistPosition: waitPos,
    message: 'Apply berhasil. Kamu ada di posisi #' + waitPos + ' waiting list.',
  };
}

function getWaitingList(campaignId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.ASSIGNMENTS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const list = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === campaignId) {
      const row = rowToObject(headers, data[i]);
      // Enrich with creator score
      const creator = getCreatorById(data[i][2]);
      if (creator) {
        row.score_total = creator.score_total;
        row.followers_ig = creator.followers_ig;
        row.niche = creator.niche;
        row.domisili = creator.domisili;
        row.endorse_experience = creator.endorse_experience;
      }
      list.push(row);
    }
  }

  // Sort: approved first, then waiting by applied_at
  list.sort((a, b) => {
    if (a.status === b.status) {
      return new Date(a.applied_at) - new Date(b.applied_at);
    }
    const order = ['approved', 'assigned', 'waiting', 'rejected'];
    return order.indexOf(a.status) - order.indexOf(b.status);
  });

  return { success: true, data: list };
}

function updateAssignmentStatus(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.ASSIGNMENTS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === payload.assignmentId) {
      sheet.getRange(i + 1, 6).setValue(payload.status); // col 6 = status (1-indexed = 6)
      if (payload.rateAgreed) sheet.getRange(i + 1, 7).setValue(payload.rateAgreed);
      if (payload.notesAdmin) sheet.getRange(i + 1, 23).setValue(payload.notesAdmin);
      return { success: true };
    }
  }
  return { success: false, message: 'Assignment tidak ditemukan.' };
}

// ============================================================
// SCORING SYSTEM
// ============================================================

function calculateScore(creatorId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SHEET.CONFIG);
  const configData = configSheet.getDataRange().getValues();
  const config = {};
  configData.forEach(row => { if (row[0]) config[row[0]] = row[1]; });

  const weights = {
    engagement:   parseFloat(config['weight_engagement'])   || 0.30,
    reliability:  parseFloat(config['weight_reliability'])  || 0.25,
    consistency:  parseFloat(config['weight_consistency'])  || 0.20,
    cost:         parseFloat(config['weight_cost'])         || 0.15,
    content:      parseFloat(config['weight_content'])      || 0.10,
  };

  const creator = getCreatorById(creatorId);
  if (!creator) return null;

  // Get all completed assignments for this creator
  const assignSheet = ss.getSheetByName(SHEET.ASSIGNMENTS);
  const assignData = assignSheet.getDataRange().getValues();
  const headers = assignData[0];
  const assignments = [];
  for (let i = 1; i < assignData.length; i++) {
    if (assignData[i][2] === creatorId) {
      assignments.push(rowToObject(headers, assignData[i]));
    }
  }

  const completed = assignments.filter(a => ['posted', 'paid'].includes(a.status));
  const total = assignments.filter(a => a.status !== 'waiting').length;
  const onTime = completed.filter(a => a.posted_at && a.deadline_post &&
    new Date(a.posted_at) <= new Date(a.deadline_post)).length;

  // 1. Reliability score
  const reliabilityScore = total > 0 ? (onTime / total) * 100 : 60;

  // 2. Engagement from PERFORMANCE sheet
  const perfSheet = ss.getSheetByName(SHEET.PERFORMANCE);
  const perfData = perfSheet.getDataRange().getValues();
  const perfHeaders = perfData[0];
  const perfs = [];
  for (let i = 1; i < perfData.length; i++) {
    if (perfData[i][3] === creatorId) perfs.push(rowToObject(perfHeaders, perfData[i]));
  }

  const avgER = perfs.length > 0
    ? perfs.reduce((s, p) => s + (parseFloat(p.engagement_rate) || 0), 0) / perfs.length
    : 0;
  const erBenchmark = parseFloat(config['er_benchmark_fb']) || 2.5;
  const engagementScore = Math.min((avgER / erBenchmark) * 100, 100);

  // 3. Consistency score (std dev of ER)
  let consistencyScore = 60; // default for new creators
  if (perfs.length >= 2) {
    const mean = avgER;
    const variance = perfs.reduce((s, p) => s + Math.pow((parseFloat(p.engagement_rate) || 0) - mean, 2), 0) / perfs.length;
    const stdDev = Math.sqrt(variance);
    consistencyScore = mean > 0 ? Math.max(0, 100 - (stdDev / mean * 100)) : 60;
  }

  // 4. Content quality (manual admin rating, avg from assignments)
  const rated = completed.filter(a => a.content_rating > 0);
  const contentScore = rated.length > 0
    ? rated.reduce((s, a) => s + (parseFloat(a.content_rating) || 0), 0) / rated.length * 20
    : 60;

  // 5. Cost efficiency (placeholder — compare against follower tier avg)
  const costScore = 70; // Will be refined once rate benchmarks are established

  const totalScore = (
    engagementScore  * weights.engagement +
    reliabilityScore * weights.reliability +
    consistencyScore * weights.consistency +
    contentScore     * weights.content +
    costScore        * weights.cost
  );

  return {
    total: Math.round(totalScore * 10) / 10,
    breakdown: {
      engagement:  Math.round(engagementScore),
      reliability: Math.round(reliabilityScore),
      consistency: Math.round(consistencyScore),
      content:     Math.round(contentScore),
      cost:        Math.round(costScore),
    },
  };
}

function recalculateAllScores() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.CREATORS);
  const data = sheet.getDataRange().getValues();
  let updated = 0;

  for (let i = 1; i < data.length; i++) {
    const creatorId = data[i][0];
    if (!creatorId) continue;
    const score = calculateScore(creatorId);
    if (score) {
      sheet.getRange(i + 1, 22).setValue(score.total); // col 22 = score_total (1-indexed)
      sheet.getRange(i + 1, 3).setValue(new Date());   // updated_at
      updated++;
    }
  }

  return { success: true, updated: updated };
}

// ============================================================
// DASHBOARD STATS
// ============================================================

function getDashboardStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const creatorSheet = ss.getSheetByName(SHEET.CREATORS);
  const creatorData = creatorSheet.getDataRange().getValues();
  let totalCreators = 0, activeCreators = 0, inactiveCreators = 0;
  for (let i = 1; i < creatorData.length; i++) {
    if (!creatorData[i][0]) continue;
    totalCreators++;
    if (creatorData[i][19] === 'active') activeCreators++;
    if (creatorData[i][19] === 'inactive') inactiveCreators++;
  }

  const campSheet = ss.getSheetByName(SHEET.CAMPAIGNS);
  const campData = campSheet.getDataRange().getValues();
  let openCampaigns = 0, ongoingCampaigns = 0;
  for (let i = 1; i < campData.length; i++) {
    if (!campData[i][0]) continue;
    if (campData[i][10] === 'open') openCampaigns++;
    if (campData[i][10] === 'ongoing') ongoingCampaigns++;
  }

  const assignSheet = ss.getSheetByName(SHEET.ASSIGNMENTS);
  const assignData = assignSheet.getDataRange().getValues();
  let pendingApprovals = 0;
  for (let i = 1; i < assignData.length; i++) {
    if (assignData[i][5] === 'waiting') pendingApprovals++;
  }

  return {
    success: true,
    data: {
      totalCreators, activeCreators, inactiveCreators,
      openCampaigns, ongoingCampaigns, pendingApprovals,
    }
  };
}

// ============================================================
// UTILITIES
// ============================================================

function rowToObject(headers, row) {
  const obj = {};
  headers.forEach((h, i) => { obj[h] = row[i]; });
  return obj;
}

function formatFollowers(n) {
  if (!n) return '0';
  n = parseInt(n);
  if (n >= 1000000) return (n / 1000000).toFixed(1) + 'M';
  if (n >= 1000) return (n / 1000).toFixed(1) + 'K';
  return n.toString();
}

// ─── TIME-BASED TRIGGERS ─────────────────────────────────────
// Run this function once to set up triggers:
// function setupTriggers() {
//   ScriptApp.newTrigger('recalculateAllScores')
//     .timeBased().everyDays(1).atHour(2).create();
// }
