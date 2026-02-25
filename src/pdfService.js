const fs = require('fs');
const path = require('path');
const PDFDocument = require('pdfkit');

function safeText(value) {
  if (value === undefined || value === null || value === '') {
    return '-';
  }
  return String(value);
}

function ensureDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

function sectionTitle(doc, text) {
  doc.moveDown(0.6);
  doc.font('Helvetica-Bold').fontSize(12).fillColor('#0f172a').text(text, { underline: false });
  doc.moveDown(0.2);
  doc.moveTo(40, doc.y).lineTo(555, doc.y).strokeColor('#94a3b8').lineWidth(0.8).stroke();
  doc.moveDown(0.5);
}

function drawFieldRow(doc, leftLabel, leftValue, rightLabel, rightValue) {
  const y = doc.y;
  doc.font('Helvetica-Bold').fontSize(10).fillColor('#111827').text(`${leftLabel}:`, 40, y, {
    width: 120
  });
  doc.font('Helvetica').text(safeText(leftValue), 165, y, { width: 155 });

  doc.font('Helvetica-Bold').text(`${rightLabel}:`, 320, y, { width: 120 });
  doc.font('Helvetica').text(safeText(rightValue), 445, y, { width: 110 });

  doc.moveDown(1);
}

function drawAgendaTable(doc, rows) {
  const tableX = 40;
  let y = doc.y;
  const rowHeight = 28;
  const widths = [45, 185, 165, 120];
  const headers = ['Sr. No', 'Agenda', 'Action Plan', 'Responsibility'];

  function drawHeaderOrRow(values, isHeader = false) {
    let x = tableX;
    for (let i = 0; i < widths.length; i += 1) {
      doc
        .rect(x, y, widths[i], rowHeight)
        .lineWidth(0.8)
        .strokeColor('#475569')
        .stroke();
      doc
        .font(isHeader ? 'Helvetica-Bold' : 'Helvetica')
        .fontSize(9)
        .fillColor('#0f172a')
        .text(safeText(values[i]), x + 4, y + 8, {
          width: widths[i] - 8,
          height: rowHeight - 6,
          ellipsis: true
        });
      x += widths[i];
    }
    y += rowHeight;
  }

  drawHeaderOrRow(headers, true);
  const effectiveRows = rows && rows.length ? rows : [{ srNo: '1', agenda: '', actionPlan: '', responsibility: '' }];
  for (const row of effectiveRows) {
    if (y > 740) {
      doc.addPage();
      y = 60;
      drawHeaderOrRow(headers, true);
    }
    drawHeaderOrRow([row.srNo, row.agenda, row.actionPlan, row.responsibility], false);
  }

  doc.y = y + 6;
}

function drawAttendeeTable(doc, rows) {
  const tableX = 40;
  let y = doc.y;
  const rowHeight = 28;
  const widths = [55, 230, 230];
  const headers = ['Sr.', 'Elegrow Technology (Name)', 'Client (Name)'];

  function draw(values, isHeader = false) {
    let x = tableX;
    for (let i = 0; i < widths.length; i += 1) {
      doc
        .rect(x, y, widths[i], rowHeight)
        .lineWidth(0.8)
        .strokeColor('#475569')
        .stroke();
      doc
        .font(isHeader ? 'Helvetica-Bold' : 'Helvetica')
        .fontSize(9)
        .fillColor('#0f172a')
        .text(safeText(values[i]), x + 4, y + 7, {
          width: widths[i] - 8,
          height: rowHeight - 6,
          ellipsis: true
        });
      x += widths[i];
    }
    y += rowHeight;
  }

  draw(headers, true);
  const effectiveRows = rows && rows.length ? rows : [{ srNo: '1', elegrowName: '', clientName: '' }];
  for (const row of effectiveRows) {
    if (y > 740) {
      doc.addPage();
      y = 60;
      draw(headers, true);
    }
    draw([row.srNo, row.elegrowName, row.clientName], false);
  }

  doc.y = y;
}

function generateFileName() {
  const stamp = new Date().toISOString().replace(/[.:]/g, '-');
  return `MOM-${stamp}.pdf`;
}

async function generateMomPdf(mom, outputDir) {
  ensureDir(outputDir);

  const fileName = generateFileName();
  const filePath = path.join(outputDir, fileName);

  await new Promise((resolve, reject) => {
    const doc = new PDFDocument({
      size: 'A4',
      margin: 40
    });

    const stream = fs.createWriteStream(filePath);
    doc.pipe(stream);

    doc.font('Helvetica-Bold').fontSize(18).fillColor('#0f172a').text('Minutes of Meeting', {
      align: 'center'
    });

    doc.moveDown(0.3);
    doc
      .font('Helvetica')
      .fontSize(9)
      .fillColor('#334155')
      .text(mom.organizationAddress, {
        align: 'center'
      });

    sectionTitle(doc, '1. GENERAL INFORMATION');
    drawFieldRow(doc, 'Meeting Title', mom.meetingTitle, 'Project Name', mom.projectName);
    drawFieldRow(
      doc,
      'Project No / Work Order No',
      mom.projectNoWorkOrderNo,
      'Client Name',
      mom.clientName
    );
    drawFieldRow(doc, 'Meeting Date', mom.meetingDate, 'Meeting Time', mom.meetingTime);
    drawFieldRow(doc, 'Entry Time', mom.entryTime, 'Exit Time', mom.exitTime);

    sectionTitle(doc, '2. DETAILS OF MEETING');
    drawFieldRow(doc, 'Meeting Location', mom.meetingLocation, 'Meeting Called by', mom.meetingCalledBy);
    drawFieldRow(
      doc,
      'Type of Meeting',
      [...mom.meetingType, mom.meetingTypeOther].filter(Boolean).join(', '),
      'Facilitator Representative',
      mom.facilitatorRepresentative
    );
    drawFieldRow(
      doc,
      'Elegrow Representative',
      mom.elegrowRepresentative,
      'Client Representative',
      mom.clientRepresentative
    );

    sectionTitle(doc, '3. MINUTES');
    doc
      .font('Helvetica')
      .fontSize(10)
      .fillColor('#0f172a')
      .text(safeText(mom.minutes), {
        width: 515,
        align: 'left'
      });

    sectionTitle(doc, '4. AGENDA');
    drawAgendaTable(doc, mom.agendaRows);

    sectionTitle(doc, 'Attendees');
    drawAttendeeTable(doc, mom.attendeeRows);

    doc.end();

    stream.on('finish', resolve);
    stream.on('error', reject);
  });

  return { fileName, filePath };
}

module.exports = {
  generateMomPdf
};
