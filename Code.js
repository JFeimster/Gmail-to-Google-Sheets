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
    submissionStarted: { lead_status: 'New', priority: 'High', follow_up_status: 'Queued', bank_connection_status: 'Pending', documents_status: '' },
    completedApplication: { lead_status: 'In Progress', priority: 'Medium', follow_up_status: 'Monitor', bank_connection_status: 'Connected', documents_status: '' },
    initialUnderwriting: { lead_status: 'In Progress', priority: 'Medium', follow_up_status: 'Monitor', bank_connection_status: 'Connected', documents_status: '' },
    waitingDocs: { lead_status: 'Needs Applicant Action', priority: 'High', follow_up_status: 'Queued', bank_connection_status: 'Connected', documents_status: 'Requested' },
    applicationIncomplete: { lead_status: 'Needs Applicant Action', priority: 'High', follow_up_status: 'Queued', bank_connection_status: 'Pending', documents_status: 'Requested' },
    fileClosedLost: { lead_status: 'Lost', priority: 'Low', follow_up_status: 'Closed', bank_connection_status: '', documents_status: '' },
  },
};

/**
 * Main runner.
 * Watches only Applicants/DAC, upserts into Master Applicants, then relabels the thread.
 */
function processApplicantEmails() {
  setupApplicantsLabels();

  const triggerLabel = GmailApp.getUserLabelByName(CONFIG.labels.trigger);
  if (!triggerLabel) throw new Error(`Missing Gmail label: ${CONFIG.labels.trigger}`);

  const processedLabel = GmailApp.getUserLabelByName(CONFIG.labels.processed);
  const errorLabel = GmailApp.getUserLabelByName(CONFIG.labels.error);
  const manualReviewLabel = GmailApp.getUserLabelByName(CONFIG.labels.manualReview);
  const laneGiggleLabel = CONFIG.enableInformationalLaneLabels ? GmailApp.getUserLabelByName(CONFIG.labels.laneGiggle) : null;
  const laneBankBreezyLabel = CONFIG.enableInformationalLaneLabels ? GmailApp.getUserLabelByName(CONFIG.labels.laneBankBreezy) : null;

  const query = `label:"${CONFIG.labels.trigger}" -label:"${CONFIG.labels.processed}" newer_than:${CONFIG.searchWindowDays}d`;

  const threads = GmailApp.search(query, 0, CONFIG.maxThreadsPerRun);
  if (!threads.length) {
    Logger.log('No unprocessed applicant threads found.');
    return;
  }

  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(CONFIG.masterSheetName);
  if (!sheet) throw new Error(`Missing sheet: ${CONFIG.masterSheetName}`);
  
  const headerMap = getHeaderMap_(sheet);

  threads.forEach(thread => {
    try {
      const record = buildRecordFromThread_(thread);

      if (!record || (!record.email && !record.applicant_full_name && !record.business_name)) {
        thread.addLabel(errorLabel);
        thread.removeLabel(triggerLabel);
        return;
      }

      upsertApplicantRow_(sheet, headerMap, record);

      thread.addLabel(processedLabel);
      thread.removeLabel(triggerLabel);
      thread.removeLabel(errorLabel);

      if (record.needs_manual_review) {
        thread.addLabel(manualReviewLabel);
      } else {
        thread.removeLabel(manualReviewLabel);
      }

      if (CONFIG.enableInformationalLaneLabels) {
        if (record.funding_lane === 'Giggle') {
          thread.addLabel(laneGiggleLabel);
          if (laneBankBreezyLabel) thread.removeLabel(laneBankBreezyLabel);
        } else if (record.funding_lane === 'BankBreezy') {
          thread.addLabel(laneBankBreezyLabel);
          if (laneGiggleLabel) thread.removeLabel(laneGiggleLabel);
        } else {
          if (laneGiggleLabel) thread.removeLabel(laneGiggleLabel);
          if (laneBankBreezyLabel) thread.removeLabel(laneBankBreezyLabel);
        }
      }
    } catch (err) {
      Logger.log(`Error processing thread ${thread.getId()}: ${err.message}`);
      thread.addLabel(errorLabel);
      thread.removeLabel(triggerLabel);
    }
  });
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
  const messages = thread.getMessages().slice().sort((a, b) => a.getDate().getTime() - b.getDate().getTime());

  const record = blankRecord_(thread);
  let bestLaneConfidence = 0;
  let latestStageTimestamp = 0;

  record.created_at = messages.length > 0 ? messages[0].getDate() : new Date();

  messages.forEach(message => {
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
    fillIfBlank_(record, 'close_reason', parsed.close_reason);
    fillIfBlank_(record, 'missing_documents', parsed.missing_documents);

    if (parsed.monthly_revenue_min !== '') record.monthly_revenue_min = parsed.monthly_revenue_min;
    if (parsed.monthly_revenue_max !== '') record.monthly_revenue_max = parsed.monthly_revenue_max;

    if (parsed.application_started_at && !record.application_started_at) {
      record.application_started_at = parsed.application_started_at;
    }
    if (parsed.application_submitted_at) {
      record.application_submitted_at = parsed.application_submitted_at;
    }

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
        if (parsed.close_reason) {
          record.close_reason = parsed.close_reason;
        }
      }
    }
  });

  if (!record.funding_lane) {
    record.funding_lane = inferFundingLane_(
      record.account_type_raw,
      record.time_in_business_raw,
      record.monthly_revenue_min,
      record.monthly_revenue_max
    );
  }

  record.updated_at = new Date();
  record.source_provider = CONFIG.sourceProvider;
  record.source_channel = 'Gmail';
  record.gmail_label = thread.getLabels().map(l => l.getName()).join(', ');
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
  const text = `${subject}\n${body}`.toLowerCase();
  const date = message.getDate();

  const applicant_full_name = firstNonEmpty_(
    matchFirst_(body, /\bName:\s*(.+)/i),
    matchFirst_(body, /\bOwner Name:\s*(.+)/i),
    matchFirst_(subject, /for\s+(.+)$/i),
    matchFirst_(subject, /-\s*(.+)$/i)
  );

  const business_name = extractBusinessName_(subject, body);

  const email = firstNonEmpty_(
    matchFirst_(body, /\bOwner Email:\s*([^\s<]+@[^\s>]+)/i),
    matchFirst_(body, /\bEmail Address:\s*([^\s<]+@[^\s>]+)/i),
    matchFirst_(body, /\bEmail:\s*([^\s<]+@[^\s>]+)/i)
  );

  const phone = firstNonEmpty_(
    matchFirst_(body, /\bOwner Phone:\s*([+\d\-\(\)\s]+)/i),
    matchFirst_(body, /\bPhone(?: Number)?:\s*([+\d\-\(\)\s]+)/i)
  );

  const routeText = firstNonEmpty_(
    matchFirst_(body, /Based on their answers,?\s*we(?: have)?\s*advanced their application(?: for completion)? with\s+(.+?)(?:\.|\n|$)/i),
    matchFirst_(body, /Based on their answers,?\s*we(?: have)?\s*advanced their application to\s+(.+?)(?:\.|\n|$)/i),
    matchFirst_(body, /advanced their application(?: for completion)? with\s+(.+?)(?:\.|\n|$)/i)
  );
  const explicitLane = classifyLaneFromRouteText_(routeText);

  const detailsLine = firstNonEmpty_(
    matchFirst_(body, /lowest monthly revenue of\s+(.+?)\s+and a bank account type of\s+(personal|business)/i),
    matchFirst_(body, /monthly revenue of\s+(.+?)\s+and a bank account type of\s+(personal|business)/i)
  );

  const monthlyRevenueBand = firstNonEmpty_(
    matchFirst_(body, /lowest monthly revenue of\s+(.+?)(?:\s+and a bank account type|\.\s|\n|$)/i),
    matchFirst_(body, /monthly revenue of\s+(.+?)(?:\s+and a bank account type|\.\s|\n|$)/i)
  );
  const revenueParsed = parseRevenueBand_(monthlyRevenueBand);

  const accountTypeRaw = normalizeAccountType_(firstNonEmpty_(
    matchFirst_(body, /bank account type of\s+(personal|business)/i),
    matchFirst_(body, /using a\s+(personal|business)\s+bank account/i),
    detailsLine ? matchSecond_(detailsLine, /(.+?)\s+(personal|business)$/i) : ''
  ));

  const timeInBusinessRaw = firstNonEmpty_(
    matchFirst_(body, /time in business(?:[^:\n]*?)\s*(?:is|of|:)\s*(.+?)(?:\.|\n|$)/i),
    matchFirst_(body, /business age(?:[^:\n]*?)\s*(?:is|of|:)\s*(.+?)(?:\.|\n|$)/i)
  );

  let closeReason = matchFirst_(body, /Reason for close:\s*(.+)/i);

  const missingDocuments = extractMissingDocuments_(body);

  const parsed = {
    event_time: date,
    applicant_full_name,
    business_name,
    email,
    phone,
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
    requested_amount: matchFirst_(body, /Requested Amount:\s*\$?([\d,]+)/i),
  };

  const declineRegex = /they were declined|application was declined|revenue was too low|do not qualify for funding at this time|not eligible|ineligible|cannot qualify|declined|decline/i;
  if (declineRegex.test(text)) {
    applyStatusMap_(parsed, 'File Closed - Deal Lost');
    const isLowRev = /revenue was too low/i.test(text);
    if (!parsed.close_reason) {
      parsed.close_reason = isLowRev ? 'Revenue was too low' : 'Declined / Not eligible';
    }
    return parsed;
  }

  if (/client file closed/i.test(text)) {
    applyStatusMap_(parsed, 'File Closed - Deal Lost');
    return parsed;
  }

  const currentStatus = firstNonEmpty_(
    matchFirst_(body, /Current Status:\s*"?(.+?)"?(?:\n|$)/i),
    matchFirst_(body, /Current Stage:\s*"?(.+?)"?(?:\n|$)/i)
  );

  if (/application is incomplete/i.test(subject) || /\bincomplete\b/i.test(text)) {
    applyStatusMap_(parsed, 'Application Incomplete');
    return parsed;
  }

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
    if (/closed|declined|lost/i.test(normalizedCurrent)) {
      applyStatusMap_(parsed, 'File Closed - Deal Lost');
      return parsed;
    }
    if (/review|underwriting|submitted|processing/i.test(normalizedCurrent)) {
      parsed.provider_stage_raw = normalizedCurrent;
      parsed.lead_status = 'In Progress';
      parsed.priority = 'Medium';
      parsed.follow_up_status = 'Monitor';
      parsed.bank_connection_status = 'Connected';
      return parsed;
    }

    parsed.provider_stage_raw = normalizedCurrent;
    parsed.lead_status = 'In Progress';
    parsed.priority = 'Medium';
    parsed.follow_up_status = 'Monitor';
    return parsed;
  }

  if (/completed .*funding application/i.test(subject) || /completed and signed/i.test(text)) {
    applyStatusMap_(parsed, 'Completed Application');
    parsed.application_submitted_at = date;
    return parsed;
  }

  if (/submission started/i.test(subject)) {
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
  if (mapKey && CONFIG.leadStatusMap[mapKey]) {
    Object.assign(parsed, CONFIG.leadStatusMap[mapKey]);
  }
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

  writeFields.forEach(field => {
    if (!(field in record)) return;
    if (field === 'created_at' && isExisting) return;
    if (field === 'updated_at') {
      setCellIfHeaderExists_(sheet, targetRow, headerMap, field, record[field]);
      return;
    }

    if (isExisting && (record[field] === '' || record[field] === null || record[field] === undefined)) return;

    setCellIfHeaderExists_(sheet, targetRow, headerMap, field, record[field]);
  });
}

function findTargetRow_(sheet, headerMap, record) {
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const lastCol = sheet.getLastColumn();
  const rowCount = Math.max(lastRow - 1, 1);
  const data = sheet.getRange(2, 1, rowCount, lastCol).getValues();

  const emailCol = headerMap['email'];
  const threadCol = headerMap['gmail_thread_id'];
  const dedupeCol = headerMap['dedupe_key'];
  const recordIdCol = headerMap['record_id'];

  const normalizedEmail = normalizeEmail_(record.email);
  const dedupeKey = safeTrim_(record.dedupe_key);
  const threadId = safeTrim_(record.gmail_thread_id);

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
  if (/bankbreezy|bigger funding/.test(text)) return { lane: 'BankBreezy', confidence: 100 };
  if (/giggle|same-day funding/.test(text)) return { lane: 'Giggle', confidence: 100 };
  return { lane: '', confidence: 0 };
}

function inferFundingLane_(accountTypeRaw, timeInBusinessRaw, revenueMin, revenueMax) {
  const accountType = normalizeAccountType_(accountTypeRaw);
  
  if (String(timeInBusinessRaw || '').trim() !== '' && !timeInBusinessEligible_(timeInBusinessRaw)) {
    return '';
  }

  const min = toNumber_(revenueMin);
  const max = toNumber_(revenueMax);

  if (accountType === 'personal' && min >= 3000) {
    return 'Giggle';
  }
  if (accountType === 'business') {
    if (min >= 15000) return 'BankBreezy';
    if (min >= 3000 && (max === '' || max < 15000)) return 'Giggle';
  }

  return '';
}

function parseRevenueBand_(raw) {
  const text = String(raw || '').replace(/\s+/g, ' ').trim();
  if (!text) return { band: '', min: '', max: '' };

  let m;

  m = text.match(/less than\s*\$?\s*([\d,]+)\s*but\s*greater than\s*\$?\s*([\d,]+)/i);
  if (m) {
    const upper = toNumber_(m[1]);
    const lower = toNumber_(m[2]);
    if (upper !== '' && lower !== '') {
      return { band: `>$${commaNumber_(lower)} and <$${commaNumber_(upper)}`, min: lower, max: upper };
    }
  }

  m = text.match(/\$?\s*([\d,]+)\s*(?:-|to|–)\s*\$?\s*([\d,]+)/i);
  if (m) {
    const left = toNumber_(m[1]);
    const right = toNumber_(m[2]);
    if (left !== '' && right !== '') {
      const min = Math.min(left, right);
      const max = Math.max(left, right);
      return { band: `$${commaNumber_(min)}-$${commaNumber_(max)}`, min: min, max: max };
    }
  }

  m = text.match(/\$?\s*([\d,]+)\s*\+/i);
  if (m) {
    const n = toNumber_(m[1]);
    if (n !== '') return { band: `$${commaNumber_(n)}+`, min: n, max: '' };
  }

  m = text.match(/greater than\s*\$?\s*([\d,]+)/i);
  if (m) {
    const n = toNumber_(m[1]);
    if (n !== '') return { band: `>$${commaNumber_(n)}`, min: n, max: '' };
  }

  m = text.match(/less than\s*\$?\s*([\d,]+)/i);
  if (m) {
    const n = toNumber_(m[1]);
    if (n !== '') return { band: `<$${commaNumber_(n)}`, min: '', max: n };
  }

  m = text.match(/^\$?\s*([\d,]+)\s*$/);
  if (m) {
    const n = toNumber_(m[1]);
    if (n !== '') return { band: `$${commaNumber_(n)}`, min: n, max: n };
  }

  return { band: text, min: '', max: '' };
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

function needsManualReview_(record) {
  return !record.email || !record.funding_lane || !record.provider_stage_raw;
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
  return `${providerSlug}-${rowNumber}-${y}${m}${d}`;
}

function buildDedupeKey_(email, applicantFullName, businessName) {
  const normalizedEmail = normalizeEmail_(email);
  if (normalizedEmail) return normalizedEmail;
  return [safeTrim_(applicantFullName).toLowerCase(), safeTrim_(businessName).toLowerCase()].join('|');
}

function buildNotes_(record, messages) {
  const latest = messages.length ? messages[messages.length - 1] : null;
  const parts = [];

  if (latest) parts.push(`Last Subject: ${safeTrim_(latest.getSubject())}`);
  if (record.provider_stage_raw) parts.push(`Stage: ${record.provider_stage_raw}`);
  if (record.funding_lane) parts.push(`Lane: ${record.funding_lane}`);
  if (record.bank_connection_status) parts.push(`Bank Link: ${record.bank_connection_status}`);
  if (record.documents_status) parts.push(`Docs: ${record.documents_status}`);
  if (record.missing_documents) parts.push(`Missing Docs: ${record.missing_documents}`);
  if (record.close_reason) parts.push(`Close Reason: ${record.close_reason}`);

  return parts.join(' | ');
}

function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const map = {};
  headers.forEach((header, idx) => {
    if (header) map[String(header).trim()] = idx + 1;
  });
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

function extractBusinessName_(subject, body) {
  const candidates = [
    matchFirst_(body, /\bCompany:\s*(.+)/i),
    matchFirst_(body, /\bBusiness Name:\s*(.+)/i),
    matchFirst_(body, /\bCompany Name:\s*(.+)/i),
    matchFirst_(body, /\bBusiness:\s*(.+)/i),
    matchFirst_(subject, /^Time Sensitive\s*-\s*(.+?)\s+Funding Application/i),
    matchFirst_(subject, /for\s+(.+?)\s+Funding Application/i)
  ];

  for (let i = 0; i < candidates.length; i++) {
    const raw = safeTrim_(candidates[i]);
    if (!raw) continue;
    const cleaned = raw
      .replace(/[“”]/g, '"')
      .replace(/[‘’]/g, "'")
      .replace(/\s*\|[\s\S]*$/g, '')
      .replace(/\s*(?:[-:|]\s*)?(?:owner phone|owner email|email address|requested amount|account type|time in business|business age|monthly revenue|bank account type)\b[\s\S]*$/i, '')
      .replace(/^["'`]+|["'`]+$/g, '')
      .replace(/\s{2,}/g, ' ')
      .trim();
    if (cleaned && !looksLikeBadBusinessName_(cleaned)) return cleaned;
  }

  return '';
}

function looksLikeBadBusinessName_(value) {
  const v = safeTrim_(value);
  if (!v) return true;

  const lower = v.toLowerCase();
  
  if (/(^|\b)\*?\s*details\s*:?\s*\*?/i.test(lower)) return true;
  if (/\b(owner phone|owner email|email address|requested amount|account type|monthly revenue|time in business)\b/i.test(lower)) return true;
  if (/^(n\/a|na|none|null|unknown|tbd)$/i.test(v)) return true;
  if (/^\d+$/.test(v)) return true;
  if (/^\+?[\d\-\(\)\s]{7,}$/.test(v)) return true; 
  if (/^\d{3}-\d{2}-\d{4}$/.test(v)) return true; 
  if (/\bphone number\b/i.test(lower)) return true;
  if (/@/.test(v)) return true;
  if (v.length < 2 || v.length > 120) return true;

  if (!/[a-z]/i.test(v)) return true;
  if (/[^a-z0-9&.,'()\-\/\s]/i.test(v)) return true;

  return false;
}

function getBestMessageBody_(message) {
  const plain = normalizeWhitespace_(message.getPlainBody() || '');
  const html = String(message.getBody() || '');

  if (plain && !/[]/.test(plain) && plain.length >= 60) return plain;

  if (html) {
    const fromHtml = normalizeWhitespace_(htmlToText_(html));
    if (fromHtml && fromHtml.length > plain.length * 0.5) return fromHtml;
  }

  return plain;
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
  if (/closed|deal lost|declined|lost/i.test(text)) return 'File Closed - Deal Lost';
  return text;
}

function timeInBusinessEligible_(value) {
  const text = String(value || '').toLowerCase().trim();
  if (!text) return true; 
  if (/4\s*\+?\s*months|4 months or greater|four months or greater|4\+ months/.test(text)) return true;
  if (/([5-9]|[1-9]\d)\s*months/.test(text)) return true;
  if (/\b1\s*year|\b2\s*years|\b3\s*years|\b\d+\s*years/.test(text)) return true;
  if (/<\s*4|less than 4|under 4/.test(text)) return false;
  return true;
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

function matchSecond_(text, regex) {
  const m = String(text || '').match(regex);
  return m && m[2] ? safeTrim_(m[2]) : '';
}

function firstNonEmpty_(...args) {
  for (const v of args) {
    if (v !== null && v !== undefined && String(v).trim() !== '') {
      return String(v).trim();
    }
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