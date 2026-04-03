const path = require('path');
const dotenv = require('dotenv');

dotenv.config();

const rootDir = path.resolve(__dirname, '..');
const generatedDir = path.join(rootDir, 'generated-pdfs');
const dataDir = path.join(rootDir, 'data');
const appHost = process.env.APP_HOST || '127.0.0.1';
const appPort = Number.parseInt(process.env.PORT || '3000', 10);
const DEFAULT_GENERATED_BY = 'ETPL_AI M.O.M System';

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
    generatedBy: process.env.MOM_GENERATED_BY || DEFAULT_GENERATED_BY
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
  },
  graph: {
    enabled: toBool(process.env.MS_GRAPH_ENABLED, false),
    authorityHost: process.env.MS_GRAPH_AUTHORITY_HOST || 'https://login.microsoftonline.com',
    tenantId: process.env.MS_GRAPH_TENANT_ID || '',
    clientId: process.env.MS_GRAPH_CLIENT_ID || '',
    clientSecret: process.env.MS_GRAPH_CLIENT_SECRET || '',
    mailboxUser: process.env.MS_GRAPH_MAILBOX_USER || '',
    baseUrl: process.env.MS_GRAPH_BASE_URL || 'https://graph.microsoft.com/v1.0',
    scope: process.env.MS_GRAPH_SCOPE || 'https://graph.microsoft.com/.default',
    openGraphWebLink: toBool(process.env.MS_GRAPH_OPEN_WEBLINK, false)
  },
  gemini: {
    enabled: toBool(process.env.GEMINI_ENABLED, true),
    apiKey: process.env.GEMINI_API_KEY || '',
    model: process.env.GEMINI_MODEL || 'gemini-2.5-pro',
    maxUploadBytes: Number.parseInt(process.env.GEMINI_MAX_UPLOAD_BYTES || '15728640', 10)
  }
};

module.exports = config;
