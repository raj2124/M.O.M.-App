const fs = require('fs');
const path = require('path');

function ensureDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

function toIsoString(value) {
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) {
    return new Date().toISOString();
  }
  return date.toISOString();
}

function parseDateMs(value) {
  const ms = Date.parse(String(value || ''));
  return Number.isNaN(ms) ? 0 : ms;
}

function normalizeInt(value, fallback, min) {
  const parsed = Number.parseInt(value, 10);
  if (Number.isNaN(parsed)) {
    return fallback;
  }
  return Math.max(min, parsed);
}

function createRecordsStore(config) {
  const filePath = config?.records?.filePath;
  const dataDir = config?.records?.dataDir || path.dirname(filePath);
  const retentionDays = normalizeInt(config?.records?.retentionDays, 30, 1);
  const maxCount = normalizeInt(config?.records?.maxCount, 60, 50);
  const retentionMs = retentionDays * 24 * 60 * 60 * 1000;

  ensureDir(dataDir);

  function readAll() {
    try {
      if (!fs.existsSync(filePath)) {
        return [];
      }
      const raw = fs.readFileSync(filePath, 'utf8');
      if (!raw.trim()) {
        return [];
      }
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : [];
    } catch (_error) {
      return [];
    }
  }

  function writeAll(records) {
    fs.writeFileSync(filePath, JSON.stringify(records, null, 2), 'utf8');
  }

  function sortByRecent(records) {
    return [...records].sort((a, b) => parseDateMs(b.createdAt) - parseDateMs(a.createdAt));
  }

  function cleanup(recordsInput) {
    const now = Date.now();
    const recentSorted = sortByRecent(recordsInput);
    const keptByAge = recentSorted.filter((record) => {
      const createdAtMs = parseDateMs(record.createdAt);
      if (!createdAtMs) {
        return false;
      }
      return now - createdAtMs <= retentionMs;
    });

    const trimmed = keptByAge.slice(0, maxCount);
    const trimmedIds = new Set(trimmed.map((record) => record.id));
    const removed = recentSorted.filter((record) => !trimmedIds.has(record.id));

    return {
      records: trimmed,
      removed,
      changed:
        trimmed.length !== recordsInput.length || removed.length !== Math.max(0, recordsInput.length - trimmed.length)
    };
  }

  function persistCleanup() {
    const all = readAll();
    const cleaned = cleanup(all);
    if (cleaned.changed) {
      writeAll(cleaned.records);
    }
    return cleaned;
  }

  function buildRecordId(now = new Date()) {
    const stamp = now.toISOString().replace(/[-:.TZ]/g, '').slice(0, 14);
    const suffix = Math.floor(1000 + Math.random() * 9000);
    return `MOMREC-${stamp}-${suffix}`;
  }

  function addRecord(input) {
    const cleaned = persistCleanup();
    const now = new Date();
    const createdAt = now.toISOString();
    const expiresAt = toIsoString(new Date(now.getTime() + retentionMs));

    const record = {
      id: buildRecordId(now),
      createdAt,
      expiresAt,
      documentId: String(input.documentId || '').trim(),
      meetingTitle: String(input.meetingTitle || '').trim(),
      projectName: String(input.projectName || '').trim(),
      projectNoWorkOrderNo: String(input.projectNoWorkOrderNo || '').trim(),
      clientName: String(input.clientName || '').trim(),
      meetingDate: String(input.meetingDate || '').trim(),
      meetingTime: String(input.meetingTime || '').trim(),
      meetingLocation: String(input.meetingLocation || '').trim(),
      projectSource: String(input.projectSource || '').trim(),
      zohoProjectId: String(input.zohoProjectId || '').trim(),
      taskRows: Array.isArray(input.taskRows) ? input.taskRows : [],
      zohoTaskCommentSync: Array.isArray(input.zohoTaskCommentSync) ? input.zohoTaskCommentSync : [],
      output: {
        generatePdf: Boolean(input.output?.generatePdf),
        printPdf: Boolean(input.output?.printPdf),
        sendEmail: Boolean(input.output?.sendEmail)
      },
      pdfUrl: String(input.pdfUrl || '').trim(),
      pdfFileName: String(input.pdfFileName || '').trim(),
      pdfAbsoluteUrl: String(input.pdfAbsoluteUrl || '').trim()
    };

    const merged = sortByRecent([record, ...cleaned.records]);
    const finalCleaned = cleanup(merged);
    writeAll(finalCleaned.records);

    return {
      record,
      removed: [...cleaned.removed, ...finalCleaned.removed]
    };
  }

  function listRecords(query = '') {
    const cleaned = persistCleanup();
    const text = String(query || '').trim().toLowerCase();
    if (!text) {
      return {
        records: cleaned.records,
        removed: cleaned.removed
      };
    }

    const records = cleaned.records.filter((record) => {
      return [
        record.documentId,
        record.projectName,
        record.projectNoWorkOrderNo,
        record.meetingTitle,
        record.clientName,
        record.meetingDate
      ]
        .join(' ')
        .toLowerCase()
        .includes(text);
    });

    return {
      records,
      removed: cleaned.removed
    };
  }

  function deleteRecord(recordId) {
    const cleaned = persistCleanup();
    const targetId = String(recordId || '').trim();
    if (!targetId) {
      return {
        deleted: false,
        record: null,
        removed: cleaned.removed
      };
    }

    const match = cleaned.records.find((record) => String(record.id || '') === targetId) || null;
    const remaining = cleaned.records.filter((record) => String(record.id || '') !== targetId);
    if (!match) {
      return {
        deleted: false,
        record: null,
        removed: cleaned.removed
      };
    }

    writeAll(remaining);
    return {
      deleted: true,
      record: match,
      removed: cleaned.removed
    };
  }

  function updateRecord(recordId, updates = {}) {
    const cleaned = persistCleanup();
    const targetId = String(recordId || '').trim();
    if (!targetId) {
      return {
        updated: false,
        record: null,
        removed: cleaned.removed
      };
    }

    let updatedRecord = null;
    const nextRecords = cleaned.records.map((record) => {
      if (String(record.id || '') !== targetId) {
        return record;
      }

      updatedRecord = {
        ...record,
        ...updates
      };
      return updatedRecord;
    });

    if (!updatedRecord) {
      return {
        updated: false,
        record: null,
        removed: cleaned.removed
      };
    }

    writeAll(nextRecords);
    return {
      updated: true,
      record: updatedRecord,
      removed: cleaned.removed
    };
  }

  return {
    listRecords,
    addRecord,
    deleteRecord,
    updateRecord,
    settings: {
      retentionDays,
      maxCount
    }
  };
}

module.exports = {
  createRecordsStore
};
