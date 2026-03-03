function toArray(value) {
  if (Array.isArray(value)) {
    return value;
  }
  return [];
}

function sanitizeString(value) {
  if (value === undefined || value === null) {
    return '';
  }
  return String(value).trim();
}

function sanitizeMomPayload(payload) {
  const meetingType = toArray(payload.meetingType).map(sanitizeString).filter(Boolean);

  return {
    meetingTitle: sanitizeString(payload.meetingTitle),
    projectName: sanitizeString(payload.projectName),
    projectNoWorkOrderNo: sanitizeString(payload.projectNoWorkOrderNo),
    clientName: sanitizeString(payload.clientName),
    meetingDate: sanitizeString(payload.meetingDate),
    meetingTime: sanitizeString(payload.meetingTime),
    entryTime: sanitizeString(payload.entryTime),
    exitTime: sanitizeString(payload.exitTime),
    meetingLocation: sanitizeString(payload.meetingLocation),
    meetingCalledBy: sanitizeString(payload.meetingCalledBy),
    meetingType,
    meetingTypeOther: sanitizeString(payload.meetingTypeOther),
    facilitatorRepresentative: sanitizeString(payload.facilitatorRepresentative),
    elegrowRepresentative: sanitizeString(payload.elegrowRepresentative),
    clientRepresentative: sanitizeString(payload.clientRepresentative),
    agendaRows: toArray(payload.agendaRows).map((row, index) => ({
      srNo: sanitizeString(row?.srNo) || String(index + 1),
      agenda: sanitizeString(row?.agenda),
      actionPlan: sanitizeString(row?.actionPlan),
      responsibility: sanitizeString(row?.responsibility)
    })),
    taskRows: toArray(payload.taskRows).map((row, index) => ({
      srNo: sanitizeString(row?.srNo) || String(index + 1),
      taskId: sanitizeString(row?.taskId),
      taskName: sanitizeString(row?.taskName),
      quantityDescription: sanitizeString(row?.quantityDescription),
      remarks: sanitizeString(row?.remarks),
      status: sanitizeString(row?.status)
    })),
    attendeeRows: toArray(payload.attendeeRows).map((row, index) => ({
      srNo: sanitizeString(row?.srNo) || String(index + 1),
      elegrowName: sanitizeString(row?.elegrowName),
      clientName: sanitizeString(row?.clientName)
    })),
    organizationAddress:
      sanitizeString(payload.organizationAddress) ||
      '302, Sangini Aspire, Beside Sanskruti Township Near Pal RTO, Pal-Hajira Road, Pal Gam, Surat, Gujarat - 395009'
  };
}

function validateMomPayload(mom) {
  const errors = [];

  if (!mom.meetingTitle) {
    errors.push('Meeting Title is required.');
  }
  if (!mom.projectName) {
    errors.push('Project Name is required (fetch from Zoho).');
  }
  if (!mom.projectNoWorkOrderNo) {
    errors.push('Project No / Work Order No is required (fetch from Zoho).');
  }
  if (!mom.meetingDate) {
    errors.push('Meeting Date is required.');
  }
  if (!mom.meetingLocation) {
    errors.push('Meeting Location is required.');
  }

  return errors;
}

module.exports = {
  sanitizeMomPayload,
  validateMomPayload
};
