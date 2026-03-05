const path = require('path');
const dotenv = require('dotenv');

dotenv.config();

const rootDir = path.resolve(__dirname, '..');
const generatedDir = path.join(rootDir, 'generated-pdfs');
const dataDir = path.join(rootDir, 'data');
const appHost = process.env.APP_HOST || '127.0.0.1';
const appPort = Number.parseInt(process.env.PORT || '3000', 10);

function toBool(value, defaultValue = false) {
  if (value === undefined || value === null || value === '') {
    return defaultValue;
  }
  return String(value).toLowerCase() === 'true';
}

const config = {
  app: {
    host: appHost,
    port: appPort,
    rootDir,
    generatedDir,
    baseUrl: process.env.APP_BASE_URL || `http://${appHost}:${appPort}`,
    generatedBy: process.env.MOM_GENERATED_BY || 'ETPL AI_M.O.M System'
  },
  records: {
    dataDir,
    filePath: process.env.MOM_RECORDS_FILE || path.join(dataDir, 'mom-records.json'),
    retentionDays: Number.parseInt(process.env.MOM_RECORD_RETENTION_DAYS || '30', 10),
    maxCount: Number.parseInt(process.env.MOM_RECORD_MAX_COUNT || '60', 10)
  },
  zoho: {
    useMock: toBool(process.env.ZOHO_USE_MOCK, true),
    enrichProjectStage: toBool(process.env.ZOHO_ENRICH_PROJECT_STAGE, true),
    baseUrl: process.env.ZOHO_BASE_URL || 'https://projectsapi.zoho.com/restapi',
    accountsBaseUrl: process.env.ZOHO_ACCOUNTS_BASE_URL || 'https://accounts.zoho.com',
    organizationUserEmailDomain: process.env.ORG_USER_EMAIL_DOMAIN || 'elegrow.com',
    portalId: process.env.ZOHO_PORTAL_ID || '',
    accessToken: process.env.ZOHO_ACCESS_TOKEN || '',
    refreshToken: process.env.ZOHO_REFRESH_TOKEN || '',
    clientId: process.env.ZOHO_CLIENT_ID || '',
    clientSecret: process.env.ZOHO_CLIENT_SECRET || '',
    projectsEndpoint: process.env.ZOHO_PROJECTS_ENDPOINT || ''
  },
  email: {
    enabled: toBool(process.env.SMTP_ENABLED, false),
    host: process.env.SMTP_HOST || '',
    port: Number.parseInt(process.env.SMTP_PORT || '587', 10),
    secure: toBool(process.env.SMTP_SECURE, false),
    user: process.env.SMTP_USER || '',
    pass: process.env.SMTP_PASS || '',
    from: process.env.SMTP_FROM || 'mom-app@example.com'
  }
};

module.exports = config;
