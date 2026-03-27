const CONFIG = {
  spreadsheetId: '1wfcDAuHKCh4gV1LOR3X7U_YlAiQiCI4IZs0NCQJYrlo',
  masterSheetName: 'Master Applicants',
  sourceProvider: 'DAC',
  searchWindowDays: 180,
  maxThreadsPerRun: 100,
  enableInformationalLaneLabels: true,
  labels: {
    root: 'Applicants',
    trigger: 'Applicants/DAC',
    processed: 'Applicants/Processed',
    error: 'Applicants/Error',
    manualReview: 'Applicants/ManualReview',
    laneGiggle: 'Applicants/DAC/Giggle',
    laneBankBreezy: 'Applicants/DAC/BankBreezy',
  },
  leadStatusMap: {
    submissionStarted: {
      lead_status: 'New',
      priority: 'High',
      follow_up_status: 'Queued',
      bank_connection_status: 'Pending',
      documents_status: ''
    },
    completedApplication: {
      lead_status: 'In Progress',
      priority: 'Medium',
      follow_up_status: 'Monitor',
      bank_connection_status: 'Connected',
      documents_status: ''
    },
    initialUnderwriting: {
      lead_status: 'In Progress',
      priority: 'Medium',
      follow_up_status: 'Monitor',
      bank_connection_status: 'Connected',
      documents_status: ''
    },
    waitingDocs: {
      lead_status: 'Needs Applicant Action',
      priority: 'High',
      follow_up_status: 'Queued',
      bank_connection_status: 'Connected',
      documents_status: 'Requested'
    },
    applicationIncomplete: {
      lead_status: 'Needs Applicant Action',
      priority: 'High',
      follow_up_status: 'Queued',
      bank_connection_status: 'Pending',
      documents_status: 'Requested'
    },
    fileClosedLost: {
      lead_status: 'Lost',
      priority: 'Low',
      follow_up_status: 'Closed',
      bank_connection_status: '',
      documents_status: ''
    },
  }
};

function processApplicantEmails() {
  setupApplicantsLabels();

  const triggerLabel = GmailApp.getUserLabelByName(CONFIG.labels.trigger);
  if (!triggerLabel) throw new Error('Missing Gmail label: ' + CONFIG.labels.trigger);

  const processedLabel = GmailApp.getUserLabelByName(CONFIG.labels.processed);
  const errorLabel = GmailApp.getUserLabelByName(CONFIG.labels.error);
  const manualReviewLabel = GmailApp.getUserLabelByName(CONFIG.labels.manualReview);
  const laneGiggleLabel = CONFIG.enableInformationalLaneLabels ? GmailApp.getUserLabelByName(CONFIG.labels.laneGiggle) : null;
  const laneBankBreezyLabel = CONFIG.enableInformationalLaneLabels ? GmailApp.getUserLabelByName(CONFIG.labels.laneBankBreezy) : null;

  const query = 'label:"' + CONFIG.labels.trigger + '" -label:"' + CONFIG.labels.processed + '" newer_than:' + CONFIG.searchWindowDays + 'd';
  const threads = GmailApp.search(query, 0, CONFIG.maxThreadsPerRun);
  if (!threads.length) {
    Logger.log('No unprocessed applicant threads found.');
    return;
  }

  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(CONFIG.masterSheetName);
  if (!sheet) throw new Error('Missing sheet: ' + CONFIG.masterSheetName);

  const headerMap = getHeaderMap_(sheet);

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    try {
      const record = buildRecordFromThread_(thread);
      if (!record || (!record.email && !record.applicant_full_name && !record.business_name)) {
        thread.addLabel(errorLabel);
        thread.removeLabel(triggerLabel);
        continue;
      }

      upsertApplicantRow_(sheet, headerMap, record);

      thread.addLabel(processedLabel);
      thread.removeLabel(triggerLabel);
      thread.removeLabel(errorLabel);

      if (record.needs_manual_review) thread.addLabel(manualReviewLabel);
      else thread.removeLabel(manualReviewLabel);

      if (CONFIG.enableInformationalLaneLabels) {
        if (record.funding_lane === 'Giggle') {
          if (laneGiggleLabel) thread.addLabel(laneGiggleLabel);
          if (laneBankBreezyLabel) thread.removeLabel(laneBankBreezyLabel);
        } else if (record.funding_lane === 'BankBreezy') {
          if (laneBankBreezyLabel) thread.addLabel(laneBankBreezyLabel);
          if (laneGiggleLabel) thread.removeLabel(laneGiggleLabel);
        } else {
          if (laneGiggleLabel) thread.removeLabel(laneGiggleLabel);
          if (laneBankBreezyLabel) thread.removeLabel(laneBankBreezyLabel);
        }
      }
    } catch (err) {
      Logger.log('Error processing thread ' + thread.getId() + ': ' + err.message);
      thread.addLabel(errorLabel);
      thread.removeLabel(triggerLabel);
    }
  }
}

function backfillDacApplicants() {
  processApplicantEmails();
}

function setupApplicantsLabels() {
  [
    CONFIG.labels.root,
    CONFIG.labels.trigger,
    CONFIG.labels.processed,
    CONFIG.labels.error,
    CONFIG.labels.manualReview,
    CONFIG.labels.laneGiggle,
    CONFIG.labels.laneBankBreezy,
  ].forEach(getOrCreateLabel_);
}

function buildRecordFromThread_(thread) {
  const messages = thread.getMessages().slice().sort(function(a, b) {
    return a.getDate().getTime() - b.getDate().getTime();
  });

  const record = blankRecord_(thread);
  let bestLaneConfidence = 0;
  let latestStageTimestamp = 0;

  record.created_at = messages.length ? messages[0].getDate() : new Date();

  for (let i = 0; i < messages.length; i++) {
    const message = messages[i];
    const parsed = parseMessage_(message);

    record.gmail_thread_id = thread.getId();
    record.gmail_message_id = message.getId();

    fillIfBlank_(record, 'applicant_full_name', parsed.applicant_full_name);
    fillIfBlank_(record, 'business_name', parsed.business_name);
    fillIfBlank_(record, 'email', normalizeEmail_(parsed.email));
    fillIfBlank_(record, 'phone', normalizePhone_(parsed.phone));
    fillIfBlank_(record, 'account_type_raw', parsed.account_type_raw);
    fillIfBlank_(record, 'time_in_business_raw', parsed.time_in_business_raw);
    fillIfBlank_(record, 'monthly_revenue_band', parsed.monthly_revenue_band);
    fillIfBlank_(record, 'requested_amount', parsed.requested_amount);
    fillIfBlank_(record, 'missing_documents', parsed.missing_documents);

    if (parsed.close_reason) fillIfBlank_(record, 'close_reason', parsed.close_reason);

    if (parsed.monthly_revenue_min !== '') record.monthly_revenue_min = parsed.monthly_revenue_min;
    if (parsed.monthly_revenue_max !== '') record.monthly_revenue_max = parsed.monthly_revenue_max;

    if (parsed.application_started_at && !record.application_started_at) record.application_started_at = parsed.application_started_at;
    if (parsed.application_submitted_at) record.application_submitted_at = parsed.application_submitted_at;

    if (parsed.funding_lane && parsed.lane_confidence >= bestLaneConfidence) {
      record.funding_lane = parsed.funding_lane;
      bestLaneConfidence = parsed.lane_confidence;
    }

    if (parsed.provider_stage_raw) {
      const eventMs = parsed.event_time ? parsed.event_time.getTime() : message.getDate().getTime();
      if (eventMs >= latestStageTimestamp) {
        latestStageTimestamp = eventMs;
        record.provider_stage_raw = parsed.provider_stage_raw;
        record.lead_status = parsed.lead_status;
        record.priority = parsed.priority;
        record.follow_up_status = parsed.follow_up_status;
        record.bank_connection_status = parsed.bank_connection_status;
        record.documents_status = parsed.documents_status;
        if (parsed.close_reason) record.close_reason = parsed.close_reason;
      }
    }
  }

  if (!record.funding_lane) {
    record.funding_lane = inferFundingLane_(
      record.account_type_raw,
      record.time_in_business_raw,
      record.monthly_revenue_min,
      record.monthly_revenue_max,
      record.monthly_revenue_band
    );
  }

  record.updated_at = new Date();
  record.source_provider = CONFIG.sourceProvider;
  record.source_channel = 'Gmail';
  record.gmail_label = thread.getLabels().map(function(l) { return l.getName(); }).join(', ');
  record.first_name = firstWord_(record.applicant_full_name);
  record.last_name = restWords_(record.applicant_full_name);
  record.dedupe_key = buildDedupeKey_(record.email, record.applicant_full_name, record.business_name);
  record.needs_manual_review = needsManualReview_(record);
  record.notes = buildNotes_(record, messages);

  return record;
}

function parseMessage_(message) {
  const subject = safeTrim_(message.getSubject());
  const body = getBestMessageBody_(message);
  const trusted = getTrustedContentZone_(subject, body);
  const text = (subject + '\n' + trusted).toLowerCase();
  const date = message.getDate();

  const applicant_full_name = firstNonEmpty_(
    matchFirst_(trusted, /\bName:\s*(.+)/i),
    matchFirst_(trusted, /\bOwner Name:\s*(.+)/i),
    matchFirst_(subject, /for\s+(.+)$/i),
    matchFirst_(subject, /-\s*(.+)$/i)
  );

  const business_name = extractBusinessName_(subject, trusted);

  const email = firstNonEmpty_(
    matchFirst_(trusted, /\bOwner Email:\s*([^\s<]+@[^\s>]+)/i),
    matchFirst_(trusted, /\bEmail Address:\s*([^\s<]+@[^\s>]+)/i),
    matchFirst_(trusted, /\bEmail:\s*([^\s<]+@[^\s>]+)/i),
    extractFirstEmail_(trusted),
    extractFirstEmail_(subject)
  );

  const phone = firstNonEmpty_(
    matchFirst_(trusted, /\bOwner Phone:\s*([+\d\-\(\)\s]+)/i),
    matchFirst_(trusted, /\bPhone(?: Number)?:\s*([+\d\-\(\)\s]+)/i)
  );

  const routeText = firstNonEmpty_(
    matchFirst_(trusted, /Based on their answers,?\s*we(?: have)?\s*advanced their application(?: for completion)? with\s+(.+?)(?:\.|\n|$)/i),
    matchFirst_(trusted, /Based on their answers,?\s*we(?: have)?\s*advanced their application to\s+(.+?)(?:\.|\n|$)/i),
    matchFirst_(trusted, /advanced their application(?: for completion)? with\s+(.+?)(?:\.|\n|$)/i),
    matchFirst_(trusted, /routed\s+them\s+to\s+(.+?)(?:\.|\n|$)/i)
  );
  const explicitLane = classifyLaneFromRouteText_(routeText);

  const detailsBlock = firstNonEmpty_(
    matchFirst_(trusted, /(Applicant Details:[\s\S]+?)(?:\n\s*Current Status:|\n\s*If they were declined|\n\s*If sent to|$)/i),
    trusted
  );

  const monthlyRevenueRaw = firstNonEmpty_(
    matchFirst_(detailsBlock, /\bMonthly Revenue(?: Band)?:\s*(.+?)(?:\n|$)/i),
    matchFirst_(detailsBlock, /lowest monthly revenue of\s+(.+?)(?:\s+and a bank account type|\.\s|\n|$)/i),
    matchFirst_(detailsBlock, /monthly revenue of\s+(.+?)(?:\s+and a bank account type|\.\s|\n|$)/i),
    matchFirst_(detailsBlock, /revenue\s*(?:is|of|:)\s*(.+?)(?:\n|$)/i)
  );
  const revenueParsed = parseRevenueBand_(monthlyRevenueRaw);

  const accountTypeRaw = normalizeAccountType_(firstNonEmpty_(
    matchFirst_(detailsBlock, /\bAccount Type:\s*(.+?)(?:\n|$)/i),
    matchFirst_(detailsBlock, /bank account type of\s+(personal|business)/i),
    matchFirst_(detailsBlock, /using a\s+(personal|business)\s+bank account/i),
    matchFirst_(detailsBlock, /\bBank Account Type:\s*(.+?)(?:\n|$)/i)
  ));

  const timeInBusinessRaw = firstNonEmpty_(
    matchFirst_(detailsBlock, /\bTime in Business:\s*(.+?)(?:\n|$)/i),
    matchFirst_(detailsBlock, /time in business(?:[^:\n]*?)\s*(?:is|of|:)\s*(.+?)(?:\.|\n|$)/i),
    matchFirst_(detailsBlock, /business age(?:[^:\n]*?)\s*(?:is|of|:)\s*(.+?)(?:\.|\n|$)/i)
  );

  const missingDocuments = extractMissingDocuments_(trusted);
  const closeReason = extractCloseReason_(trusted);

  const parsed = {
    event_time: date,
    applicant_full_name: applicant_full_name,
    business_name: business_name,
    email: email,
    phone: phone,
    account_type_raw: accountTypeRaw,
    time_in_business_raw: timeInBusinessRaw,
    monthly_revenue_band: revenueParsed.band,
    monthly_revenue_min: revenueParsed.min,
    monthly_revenue_max: revenueParsed.max,
    funding_lane: explicitLane.lane,
    lane_confidence: explicitLane.confidence,
    provider_stage_raw: '',
    lead_status: '',
    priority: '',
    bank_connection_status: '',
    documents_status: '',
    missing_documents: missingDocuments,
    close_reason: closeReason,
    follow_up_status: '',
    application_started_at: '',
    application_submitted_at: '',
    requested_amount: matchFirst_(detailsBlock, /Requested Amount:\s*\$?([\d,]+)/i),
  };

  const hardDeclineSignals = [
    /\bcurrent status\s*:\s*(?:file closed|deal lost|declined|ineligible|not qualified)\b/i,
    /\bapplication (?:was|is) declined\b/i,
    /\bdo not qualify for funding at this time\b/i,
    /\bnot qualified\b/i,
    /\bineligible\b/i,
    /\bdeclined due to\b/i,
    /\breason for close\b/i,
    /\bclose reason\b/i
  ];

  if (containsAny_(trusted, hardDeclineSignals) && !isHypotheticalDeclineOnly_(trusted)) {
    applyStatusMap_(parsed, 'File Closed - Deal Lost');
    if (!parsed.close_reason) {
      parsed.close_reason = /revenue was too low|low revenue/i.test(trusted)
        ? 'Revenue was too low'
        : 'Declined / Not eligible';
    }
    return parsed;
  }

  const currentStatus = firstNonEmpty_(
    matchFirst_(trusted, /Current Status:\s*"?(.+?)"?(?:\n|$)/i),
    matchFirst_(trusted, /Current Stage:\s*"?(.+?)"?(?:\n|$)/i)
  );

  if (currentStatus) {
    const normalizedCurrent = normalizeProviderStage_(currentStatus);
    if (/waiting on required documents/i.test(normalizedCurrent)) {
      applyStatusMap_(parsed, 'Waiting on Required Documents');
      return parsed;
    }
    if (/initial underwriting/i.test(normalizedCurrent)) {
      applyStatusMap_(parsed, 'Initial Underwriting');
      return parsed;
    }
    if (/application incomplete/i.test(normalizedCurrent)) {
      applyStatusMap_(parsed, 'Application Incomplete');
      return parsed;
    }
    if (/file closed|deal lost/i.test(normalizedCurrent)) {
      applyStatusMap_(parsed, 'File Closed - Deal Lost');
      if (!parsed.close_reason) parsed.close_reason = closeReason || 'File closed';
      return parsed;
    }
    if (/submission started/i.test(normalizedCurrent)) {
      applyStatusMap_(parsed, 'Submission Started');
      parsed.application_started_at = date;
      return parsed;
    }
    if (/completed application|completed/i.test(normalizedCurrent)) {
      applyStatusMap_(parsed, 'Completed Application');
      parsed.application_submitted_at = date;
      return parsed;
    }

    parsed.provider_stage_raw = normalizedCurrent;
    parsed.lead_status = 'In Progress';
    parsed.priority = 'Medium';
    parsed.follow_up_status = 'Monitor';
    parsed.bank_connection_status = 'Connected';
    return parsed;
  }

  if (/application is incomplete/i.test(subject) || /\bincomplete\b/i.test(text)) {
    applyStatusMap_(parsed, 'Application Incomplete');
    return parsed;
  }

  if (/completed .*funding application/i.test(subject) || /completed and signed/i.test(text)) {
    applyStatusMap_(parsed, 'Completed Application');
    parsed.application_submitted_at = date;
    return parsed;
  }

  if (/submission started/i.test(subject) || /\bsubmission started\b/i.test(text)) {
    applyStatusMap_(parsed, 'Submission Started');
    parsed.application_started_at = date;
    return parsed;
  }

  if (routeText || applicant_full_name || email) {
    applyStatusMap_(parsed, 'Submission Started');
    parsed.application_started_at = date;
    return parsed;
  }

  return parsed;
}

function applyStatusMap_(parsed, providerStageRaw) {
  parsed.provider_stage_raw = providerStageRaw;

  const stageToConfigKey = {
    'Submission Started': 'submissionStarted',
    'Completed Application': 'completedApplication',
    'Initial Underwriting': 'initialUnderwriting',
    'Waiting on Required Documents': 'waitingDocs',
    'Application Incomplete': 'applicationIncomplete',
    'File Closed - Deal Lost': 'fileClosedLost'
  };

  const mapKey = stageToConfigKey[providerStageRaw];
  if (mapKey && CONFIG.leadStatusMap[mapKey]) Object.assign(parsed, CONFIG.leadStatusMap[mapKey]);
}

function upsertApplicantRow_(sheet, headerMap, record) {
  const targetRow = findTargetRow_(sheet, headerMap, record);
  const isExisting = safeTrim_(getCellByHeader_(sheet, targetRow, headerMap, 'record_id')) !== '';

  record.record_id = isExisting
    ? (getCellByHeader_(sheet, targetRow, headerMap, 'record_id') || buildRecordId_(record.source_provider, targetRow, record.created_at))
    : buildRecordId_(record.source_provider, targetRow, record.created_at);

  const writeFields = [
    'record_id', 'created_at', 'updated_at', 'source_provider', 'source_channel', 'gmail_message_id',
    'gmail_thread_id', 'gmail_label', 'applicant_full_name', 'first_name', 'last_name', 'business_name',
    'email', 'phone', 'city', 'state', 'website', 'entity_name', 'entity_type', 'industry',
    'funding_product', 'requested_amount', 'account_type_raw', 'time_in_business_raw', 'monthly_revenue_band',
    'monthly_revenue_min', 'monthly_revenue_max', 'funding_lane', 'needs_manual_review', 'provider_stage_raw',
    'lead_status', 'priority', 'application_started_at', 'application_submitted_at', 'bank_connection_status',
    'bank_connection_last_checked_at', 'documents_status', 'missing_documents', 'close_reason', 'follow_up_status',
    'follow_up_owner', 'last_follow_up_at', 'next_follow_up_at', 'notes', 'partner_portal_link', 'resume_link',
    'document_folder_link', 'notion_page_url', 'dedupe_key'
  ];

  for (let i = 0; i < writeFields.length; i++) {
    const field = writeFields[i];
    if (!(field in record)) continue;
    if (field === 'created_at' && isExisting) continue;

    if (field === 'updated_at') {
      setCellIfHeaderExists_(sheet, targetRow, headerMap, field, record[field]);
      continue;
    }

    const incoming = record[field];
    if (isExisting && isBlank_(incoming)) continue;

    setCellIfHeaderExists_(sheet, targetRow, headerMap, field, incoming);
  }
}

function findTargetRow_(sheet, headerMap, record) {
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const lastCol = sheet.getLastColumn();
  const rowCount = Math.max(lastRow - 1, 1);
  const data = sheet.getRange(2, 1, rowCount, lastCol).getValues();

  const emailCol = headerMap.email;
  const threadCol = headerMap.gmail_thread_id;
  const dedupeCol = headerMap.dedupe_key;
  const recordIdCol = headerMap.record_id;

  const normalizedEmail = normalizeEmail_(record.email);
  const threadId = safeTrim_(record.gmail_thread_id);
  const dedupeKey = safeTrim_(record.dedupe_key);

  for (let i = 0; i < data.length; i++) {
    const rowNum = i + 2;
    const existingEmail = emailCol ? normalizeEmail_(data[i][emailCol - 1]) : '';
    const existingThreadId = threadCol ? safeTrim_(data[i][threadCol - 1]) : '';
    const existingDedupe = dedupeCol ? safeTrim_(data[i][dedupeCol - 1]) : '';

    if (normalizedEmail && existingEmail && normalizedEmail === existingEmail) return rowNum;
    if (threadId && existingThreadId && threadId === existingThreadId) return rowNum;
    if (dedupeKey && existingDedupe && dedupeKey === existingDedupe) return rowNum;
  }

  for (let i = 0; i < data.length; i++) {
    const rowNum = i + 2;
    const existingRecordId = recordIdCol ? safeTrim_(data[i][recordIdCol - 1]) : '';
    const existingEmail = emailCol ? safeTrim_(data[i][emailCol - 1]) : '';
    const existingThreadId = threadCol ? safeTrim_(data[i][threadCol - 1]) : '';
    if (!existingRecordId && !existingEmail && !existingThreadId) return rowNum;
  }

  return lastRow + 1;
}

function classifyLaneFromRouteText_(routeText) {
  const text = String(routeText || '').toLowerCase();
  if (!text) return { lane: '', confidence: 0 };
  if (/bankbreezy|bigger funding|higher tier|larger funding/.test(text)) return { lane: 'BankBreezy', confidence: 100 };
  if (/giggle|same-day funding|starter lane|starter funding/.test(text)) return { lane: 'Giggle', confidence: 100 };
  return { lane: '', confidence: 0 };
}

function inferFundingLane_(accountTypeRaw, timeInBusinessRaw, revenueMin, revenueMax, monthlyRevenueBand) {
  const accountType = normalizeAccountType_(accountTypeRaw);
  const hasTimeInfo = safeTrim_(timeInBusinessRaw) !== '';
  const isTimeEligible = timeInBusinessEligible_(timeInBusinessRaw);
  if (!accountType || !hasTimeInfo || !isTimeEligible) return '';

  const parsedBand = parseRevenueBand_(monthlyRevenueBand);
  const min = toNumber_(revenueMin) !== '' ? toNumber_(revenueMin) : parsedBand.min;
  const max = toNumber_(revenueMax) !== '' ? toNumber_(revenueMax) : parsedBand.max;

  const hasAtLeast3k = min !== '' ? min > 3000 : (max !== '' ? max > 3000 : false);
  if (!hasAtLeast3k) return '';

  if (accountType === 'personal') {
    return 'Giggle';
  }

  if (accountType === 'business') {
    if (min !== '' && min >= 15000) return 'BankBreezy';
    if (max !== '' && max < 15000) return 'Giggle';
    if (min !== '' && min > 3000 && min < 15000) return 'Giggle';
    if (String(monthlyRevenueBand || '').match(/\$?\s*15,?000\s*\+/i) || /greater than\s*\$?\s*15,?000/i.test(String(monthlyRevenueBand || ''))) {
      return 'BankBreezy';
    }
  }

  return '';
}

function parseRevenueBand_(raw) {
  const text = String(raw || '').replace(/[–—]/g, '-').replace(/\s+/g, ' ').trim();
  if (!text) return { band: '', min: '', max: '' };

  let m;

  m = text.match(/less than\s*\$?\s*([\d,]+)\s*but\s*greater than\s*\$?\s*([\d,]+)/i);
  if (m) {
    const upper = toNumber_(m[1]);
    const lower = toNumber_(m[2]);
    if (upper !== '' && lower !== '') {
      const min = Math.min(lower, upper);
      const max = Math.max(lower, upper);
      return { band: '>$' + commaNumber_(min) + ' and <$' + commaNumber_(max), min: min, max: max };
    }
  }

  m = text.match(/\$?\s*([\d,]+)\s*(?:-|to)\s*\$?\s*([\d,]+)/i);
  if (m) {
    const left = toNumber_(m[1]);
    const right = toNumber_(m[2]);
    if (left !== '' && right !== '') {
      const min = Math.min(left, right);
      const max = Math.max(left, right);
      return { band: '$' + commaNumber_(min) + '-$' + commaNumber_(max), min: min, max: max };
    }
  }

  m = text.match(/greater than\s*\$?\s*([\d,]+)/i);
  if (m) {
    const n = toNumber_(m[1]);
    if (n !== '') return { band: '>$' + commaNumber_(n), min: n, max: '' };
  }

  m = text.match(/\$?\s*([\d,]+)\s*\+/i);
  if (m) {
    const n = toNumber_(m[1]);
    if (n !== '') return { band: '$' + commaNumber_(n) + '+', min: n, max: '' };
  }

  m = text.match(/less than\s*\$?\s*([\d,]+)/i);
  if (m) {
    const n = toNumber_(m[1]);
    if (n !== '') return { band: '<$' + commaNumber_(n), min: '', max: n };
  }

  m = text.match(/^\$?\s*([\d,]+)\s*$/);
  if (m) {
    const n = toNumber_(m[1]);
    if (n !== '') return { band: '$' + commaNumber_(n), min: n, max: n };
  }

  return { band: text, min: '', max: '' };
}

function extractBusinessName_(subject, body) {
  const candidates = [
    matchFirst_(body, /\bBusiness Name:\s*(.+)/i),
    matchFirst_(body, /\bCompany Name:\s*(.+)/i),
    matchFirst_(body, /\bCompany:\s*(.+)/i),
    matchFirst_(body, /\bBusiness:\s*(.+)/i),
    matchFirst_(subject, /^Time Sensitive\s*-\s*(.+?)\s+Funding Application/i),
    matchFirst_(subject, /for\s+(.+?)\s+Funding Application/i),
  ];

  for (let i = 0; i < candidates.length; i++) {
    const raw = safeTrim_(candidates[i]);
    if (!raw) continue;

    const cleaned = raw
      .replace(/[“”]/g, '"')
      .replace(/[‘’]/g, "'")
      .replace(/\s*\|[\s\S]*$/g, '')
      .replace(/\s*(?:[-:|]\s*)?(?:owner email|owner phone|requested amount|account type|email address|time in business|monthly revenue|bank account type)\b[\s\S]*$/i, '')
      .replace(/^[:\-\*\s]+/, '')
      .replace(/^['"`]+|['"`]+$/g, '')
      .replace(/\s{2,}/g, ' ')
      .trim();

    if (!cleaned) continue;
    if (looksLikeBadBusinessName_(cleaned)) continue;
    return cleaned;
  }

  return '';
}

function looksLikeBadBusinessName_(value) {
  const v = safeTrim_(value);
  if (!v) return true;

  const lower = v.toLowerCase();
  if (/(^|\b)\*?\s*details\s*:?(\s*\*?)$/i.test(v) || /^\*?details\*?:?/i.test(v)) return true;
  if (/\b(owner email|owner phone|requested amount|account type|email address|phone number|monthly revenue|time in business)\b/i.test(lower)) return true;
  if (/^(n\/a|na|none|null|unknown|tbd)$/i.test(v)) return true;
  if (/^\d+$/.test(v)) return true;
  if (/^\+?[\d\-\(\)\s]{7,}$/.test(v)) return true;
  if (/^\d{3}-\d{2}-\d{4}$/.test(v)) return true;
  if (/@/.test(v)) return true;
  if (v.length < 2 || v.length > 120) return true;
  if (!/[a-z]/i.test(v)) return true;
  if (/[^a-z0-9&.,'()\-\/\s]/i.test(v)) return true;

  return false;
}

function getBestMessageBody_(message) {
  const plain = normalizeWhitespace_(message.getPlainBody() || '');
  const html = String(message.getBody() || '');

  if (isGoodTextBody_(plain)) return plain;

  if (html) {
    const fromHtml = normalizeWhitespace_(htmlToText_(html));
    if (isGoodTextBody_(fromHtml)) return fromHtml;
    if (fromHtml && fromHtml.length > plain.length) return fromHtml;
  }

  return plain || normalizeWhitespace_(htmlToText_(html));
}

function htmlToText_(html) {
  const source = String(html || '');
  if (!source) return '';

  let text = source
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<\/li>/gi, '\n')
    .replace(/<li[^>]*>/gi, '- ')
    .replace(/<\/tr>/gi, '\n')
    .replace(/<\/td>/gi, ' ')
    .replace(/<\/th>/gi, ' ')
    .replace(/<[^>]+>/g, ' ');

  text = decodeHtmlEntities_(text);
  return text.replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
}

function decodeHtmlEntities_(text) {
  return String(text || '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&#x2F;/gi, '/')
    .replace(/&#(\d+);/g, function(_, code) { return String.fromCharCode(Number(code)); })
    .replace(/&#x([0-9a-f]+);/gi, function(_, hex) { return String.fromCharCode(parseInt(hex, 16)); });
}

function normalizeProviderStage_(value) {
  const text = safeTrim_(value);
  if (!text) return '';
  if (/waiting on required documents/i.test(text)) return 'Waiting on Required Documents';
  if (/initial underwriting/i.test(text)) return 'Initial Underwriting';
  if (/submission started/i.test(text)) return 'Submission Started';
  if (/completed/i.test(text)) return 'Completed Application';
  if (/incomplete/i.test(text)) return 'Application Incomplete';
  if (/file closed|deal lost/i.test(text)) return 'File Closed - Deal Lost';
  return text;
}

function timeInBusinessEligible_(value) {
  const text = String(value || '').toLowerCase().trim();
  if (!text) return true;
  if (/4\s*\+?\s*months|4 months or greater|four months or greater|4\+ months/.test(text)) return true;
  if (/([5-9]|[1-9]\d)\s*months/.test(text)) return true;
  if (/\b1\s*year|\b2\s*years|\b3\s*years|\b\d+\s*years/.test(text)) return true;
  if (/<\s*4|less than 4|under 4/.test(text)) return false;
  const numMatch = text.match(/(\d+(?:\.\d+)?)\s*(month|months|mo)/i);
  if (numMatch) return Number(numMatch[1]) >= 4;
  return true;
}

function needsManualReview_(record) {
  if (!record.email) return true;
  if (!record.provider_stage_raw) return true;
  if (!record.funding_lane) return true;
  return false;
}

function getTrustedContentZone_(subject, body) {
  const s = String(subject || '');
  const b = String(body || '');
  if (!b) return '';

  const lines = b.split(/\n/);
  const out = [];
  let skipHypothetical = false;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const l = line.toLowerCase().trim();

    if (/^if they were declined|^if sent to|^if they are declined|^if applicant is declined/.test(l)) {
      skipHypothetical = true;
      continue;
    }

    if (skipHypothetical && (/^applicant details\b/.test(l) || /^current status\b/.test(l) || /^reason for close\b/.test(l) || /^owner email\b/.test(l))) {
      skipHypothetical = false;
    }

    if (skipHypothetical) continue;
    if (/hypothetical|example pathway|for example only/i.test(l)) continue;

    out.push(line);
  }

  const candidate = out.join('\n');

  if (/submission started/i.test(s)) {
    const focused = firstNonEmpty_(
      matchFirst_(candidate, /(Applicant Details:[\s\S]+?)(?:\n\s*If they were declined|\n\s*If sent to|$)/i),
      candidate
    );
    return focused || candidate;
  }

  return candidate;
}

function isHypotheticalDeclineOnly_(text) {
  const t = String(text || '').toLowerCase();
  if (!t) return false;

  const hasDeclineWord = /declined|ineligible|not qualified|file closed|deal lost/.test(t);
  if (!hasDeclineWord) return false;

  const hasCurrentClosedStatus = /current status\s*:\s*(file closed|deal lost|declined|ineligible|not qualified)/.test(t);
  const hasExplicitReasonClose = /reason for close\s*:/.test(t);
  const hasHypotheticalCue = /if they were declined|if sent to|if applicant is declined/.test(t);

  return hasHypotheticalCue && !hasCurrentClosedStatus && !hasExplicitReasonClose;
}

function extractCloseReason_(text) {
  return firstNonEmpty_(
    matchFirst_(text, /Reason for close:\s*(.+?)(?:\n|$)/i),
    matchFirst_(text, /Close Reason:\s*(.+?)(?:\n|$)/i)
  );
}

function extractMissingDocuments_(body) {
  const docBlock = firstNonEmpty_(
    matchFirst_(body, /Please upload the following required documents:\s*([\s\S]+?)(?:Need help\?|This email was sent|$)/i),
    matchFirst_(body, /missing\s+(.+?bank statements.*?)(?:\.|\n|$)/i),
    matchFirst_(body, /awaiting\s+(.+?bank statements.*?)(?:\.|\n|$)/i),
    matchFirst_(body, /upload\s+(.+?bank statements.*?)(?:\.|\n|$)/i)
  );
  return docBlock ? docBlock.replace(/\n+/g, ' | ').trim() : '';
}

function extractFirstEmail_(text) {
  const m = String(text || '').match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return m ? m[0] : '';
}

function containsAny_(text, patterns) {
  const source = String(text || '');
  for (let i = 0; i < patterns.length; i++) {
    if (patterns[i].test(source)) return true;
  }
  return false;
}

function isGoodTextBody_(text) {
  const t = String(text || '').trim();
  if (!t || t.length < 40) return false;
  if (/[\uFFFD]/.test(t)) return false;
  const letters = (t.match(/[a-z]/gi) || []).length;
  return letters >= 20;
}

function blankRecord_(thread) {
  return {
    record_id: '', created_at: '', updated_at: '', source_provider: CONFIG.sourceProvider,
    source_channel: 'Gmail', gmail_message_id: '', gmail_thread_id: thread.getId(), gmail_label: '',
    applicant_full_name: '', first_name: '', last_name: '', business_name: '', email: '', phone: '', city: '',
    state: '', website: '', entity_name: '', entity_type: '', industry: '', funding_product: '', requested_amount: '',
    account_type_raw: '', time_in_business_raw: '', monthly_revenue_band: '', monthly_revenue_min: '',
    monthly_revenue_max: '', funding_lane: '', needs_manual_review: false, provider_stage_raw: '', lead_status: '',
    priority: '', application_started_at: '', application_submitted_at: '', bank_connection_status: '',
    bank_connection_last_checked_at: '', documents_status: '', missing_documents: '', close_reason: '',
    follow_up_status: '', follow_up_owner: '', last_follow_up_at: '', next_follow_up_at: '', notes: '',
    partner_portal_link: '', resume_link: '', document_folder_link: '', notion_page_url: '', dedupe_key: '',
  };
}

function buildRecordId_(sourceProvider, rowNumber, createdAt) {
  const providerSlug = String(sourceProvider || 'DAC').replace(/[^A-Za-z0-9]+/g, '_').toUpperCase().replace(/^_+|_+$/g, '');
  const dt = createdAt instanceof Date ? createdAt : new Date();
  const y = dt.getFullYear();
  const m = ('0' + (dt.getMonth() + 1)).slice(-2);
  const d = ('0' + dt.getDate()).slice(-2);
  return providerSlug + '-' + rowNumber + '-' + y + m + d;
}

function buildDedupeKey_(email, applicantFullName, businessName) {
  const normalizedEmail = normalizeEmail_(email);
  if (normalizedEmail) return normalizedEmail;
  return [safeTrim_(applicantFullName).toLowerCase(), safeTrim_(businessName).toLowerCase()].join('|');
}

function buildNotes_(record, messages) {
  const latest = messages.length ? messages[messages.length - 1] : null;
  const parts = [];

  if (latest) parts.push('Last Subject: ' + safeTrim_(latest.getSubject()));
  if (record.provider_stage_raw) parts.push('Stage: ' + record.provider_stage_raw);
  if (record.funding_lane) parts.push('Lane: ' + record.funding_lane);
  if (record.bank_connection_status) parts.push('Bank Link: ' + record.bank_connection_status);
  if (record.documents_status) parts.push('Docs: ' + record.documents_status);
  if (record.missing_documents) parts.push('Missing Docs: ' + record.missing_documents);
  if (record.close_reason) parts.push('Close Reason: ' + record.close_reason);

  return parts.join(' | ');
}

function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    if (headers[i]) map[String(headers[i]).trim()] = i + 1;
  }
  return map;
}

function setCellIfHeaderExists_(sheet, row, headerMap, headerName, value) {
  const col = headerMap[headerName];
  if (!col) return;
  sheet.getRange(row, col).setValue(value);
}

function getCellByHeader_(sheet, row, headerMap, headerName) {
  const col = headerMap[headerName];
  return col ? sheet.getRange(row, col).getValue() : '';
}

function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function fillIfBlank_(obj, key, value) {
  if (!obj[key] && value) obj[key] = value;
}

function normalizeWhitespace_(text) {
  return String(text || '').replace(/\r/g, '\n').replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n').trim();
}

function normalizeEmail_(value) {
  return safeTrim_(value).toLowerCase();
}

function normalizePhone_(value) {
  return safeTrim_(value).replace(/\s+/g, ' ');
}

function normalizeAccountType_(value) {
  const text = String(value || '').toLowerCase().replace(/\s+/g, ' ').trim();
  if (!text) return '';

  if (/\bpersonal\b/.test(text) && !/\bbusiness\b/.test(text)) return 'personal';
  if (/\bbusiness\b/.test(text) && !/\bpersonal\b/.test(text)) return 'business';
  if (/personal (checking|account)/.test(text)) return 'personal';
  if (/business (checking|account)/.test(text)) return 'business';
  if (/^(p|pers)$/.test(text)) return 'personal';
  if (/^(b|biz)$/.test(text)) return 'business';

  return '';
}

function firstWord_(value) {
  const parts = safeTrim_(value).split(/\s+/).filter(Boolean);
  return parts.length ? parts[0] : '';
}

function restWords_(value) {
  const parts = safeTrim_(value).split(/\s+/).filter(Boolean);
  return parts.length > 1 ? parts.slice(1).join(' ') : '';
}

function matchFirst_(text, regex) {
  const m = String(text || '').match(regex);
  return m && m[1] ? safeTrim_(m[1]) : '';
}

function firstNonEmpty_() {
  for (let i = 0; i < arguments.length; i++) {
    const v = arguments[i];
    if (v !== null && v !== undefined && String(v).trim() !== '') return String(v).trim();
  }
  return '';
}

function safeTrim_(value) {
  return String(value || '').trim();
}

function toNumber_(value) {
  const cleaned = String(value || '').replace(/[^\d.]/g, '');
  return cleaned ? Number(cleaned) : '';
}

function commaNumber_(value) {
  const n = toNumber_(value);
  if (n === '') return '';
  return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function isBlank_(value) {
  return value === '' || value === null || value === undefined;
}
