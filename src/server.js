const path = require('path');
const fs = require('fs');
const express = require('express');

const config = require('./config');
const { getProjects, getProjectUsers } = require('./zohoClient');
const { sanitizeMomPayload, validateMomPayload } = require('./momTemplate');
const { generateMomPdf } = require('./pdfService');
const { sendMomEmail } = require('./emailService');

const app = express();

if (!fs.existsSync(config.app.generatedDir)) {
  fs.mkdirSync(config.app.generatedDir, { recursive: true });
}

app.use(express.json({ limit: '5mb' }));
app.use(express.urlencoded({ extended: true }));

app.use('/generated-pdfs', express.static(config.app.generatedDir));
app.use(express.static(path.join(config.app.rootDir, 'public')));

app.get('/api/health', (_req, res) => {
  const zohoAutoRefreshConfigured = Boolean(
    config.zoho.refreshToken && config.zoho.clientId && config.zoho.clientSecret
  );

  res.json({
    success: true,
    timestamp: new Date().toISOString(),
    zohoMode: config.zoho.useMock ? 'mock' : 'live',
    zohoAutoRefreshConfigured,
    emailEnabled: config.email.enabled
  });
});

app.get('/api/zoho/projects', async (req, res) => {
  try {
    const query = String(req.query.query || '').trim();
    const projects = await getProjects(query);

    res.json({
      success: true,
      projects
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      message: error.message || 'Failed to fetch projects from Zoho Projects'
    });
  }
});

app.get('/api/zoho/projects/:projectId/users', async (req, res) => {
  try {
    const projectId = String(req.params.projectId || '').trim();
    const users = await getProjectUsers(projectId);

    res.json({
      success: true,
      users
    });
  } catch (error) {
    const projectId = String(req.params.projectId || '').trim();
    res.status(500).json({
      success: false,
      message:
        error.message || `Failed to fetch Zoho project users for project ID: ${projectId || '(empty)'}`
    });
  }
});

app.post('/api/mom/submit', async (req, res) => {
  try {
    const mom = sanitizeMomPayload(req.body.mom || {});
    const options = req.body.options || {};

    const validationErrors = validateMomPayload(mom);
    if (validationErrors.length) {
      return res.status(400).json({
        success: false,
        message: 'Validation failed',
        errors: validationErrors
      });
    }

    const shouldGeneratePdf = Boolean(options.generatePdf || options.printPdf || options.sendEmail);

    if (!shouldGeneratePdf) {
      return res.status(400).json({
        success: false,
        message: 'Select at least one output option (Email, PDF, Print).'
      });
    }

    const { fileName, filePath } = await generateMomPdf(mom, config.app.generatedDir);
    const pdfUrl = `/generated-pdfs/${fileName}`;

    let emailSent = false;
    if (options.sendEmail) {
      await sendMomEmail({
        to: options.emailTo,
        cc: options.emailCc,
        subject: options.emailSubject || `Minutes of Meeting - ${mom.projectName}`,
        body:
          options.emailBody ||
          `<p>Dear Team,</p><p>Please find attached the Minutes of Meeting for <strong>${mom.projectName}</strong>.</p><p>Regards,<br/>M.O.M App</p>`,
        attachmentPath: filePath,
        attachmentName: fileName
      });
      emailSent = true;
    }

    return res.json({
      success: true,
      message: 'M.O.M processed successfully.',
      result: {
        pdfUrl,
        pdfFileName: fileName,
        emailSent,
        printRequested: Boolean(options.printPdf)
      }
    });
  } catch (error) {
    return res.status(500).json({
      success: false,
      message: error.message || 'Failed to process M.O.M submission'
    });
  }
});

app.get('*', (_req, res) => {
  res.sendFile(path.join(config.app.rootDir, 'public', 'index.html'));
});

const server = app.listen(config.app.port, config.app.host, () => {
  // eslint-disable-next-line no-console
  console.log(`M.O.M app running at ${config.app.baseUrl} on port ${config.app.port}`);
});

server.on('error', (error) => {
  // eslint-disable-next-line no-console
  console.error(`Server startup failed: ${error.code || error.message}`);
  process.exit(1);
});
