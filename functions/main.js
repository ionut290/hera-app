"use strict";

const existingFunctions = require("./index");
const userNotificationFunctions = require("./user-notifications");
const centralNotificationFunctions = require("./central-notifications");
const sharedCalendarViewFunctions = require("./shared-calendar-view");
const sharedOperationalViewFunctions = require("./shared-operational-views");
const adminPasswordPrivateRequestFunctions = require("./admin-password-private-requests");
const passwordRecoveryCodeFunctions = require("./password-recovery-code");
const impiantoChangeIndexFunctions = require("./impianto-change-index");
const errorReportingFunctions = require("./error-reporting");
const errorCenterFunctions = require("./error-center");
const errorCenterResetFunctions = require("./error-center-reset");
const errorEmailUsageFunctions = require("./error-email-usage");
const cleanupWhazzupPdfFunctions = require("./cleanup-whazzup-pdfs");
const whazzupPdfDriveFunctions = require("./whazzup-pdf-drive");
const userAccessApprovalFunctions = require("./user-access-approval");
const vargaGestionaleSyncFunctions = require("./varga-gestionale-sync");

Object.assign(
  exports,
  existingFunctions,
  userNotificationFunctions,
  centralNotificationFunctions,
  sharedCalendarViewFunctions,
  sharedOperationalViewFunctions,
  adminPasswordPrivateRequestFunctions,
  passwordRecoveryCodeFunctions,
  impiantoChangeIndexFunctions,
  errorReportingFunctions,
  errorCenterFunctions,
  errorCenterResetFunctions,
  errorEmailUsageFunctions,
  cleanupWhazzupPdfFunctions,
  whazzupPdfDriveFunctions,
  userAccessApprovalFunctions,
  vargaGestionaleSyncFunctions
);
