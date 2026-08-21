"use strict";

const existingFunctions = require("./index");
const userNotificationFunctions = require("./user-notifications");
const centralNotificationFunctions = require("./central-notifications");
const sharedCalendarViewFunctions = require("./shared-calendar-view");
const sharedOperationalViewFunctions = require("./shared-operational-views");
const adminPasswordPrivateRequestFunctions = require("./admin-password-private-requests");
const impiantoChangeIndexFunctions = require("./impianto-change-index");
const errorReportingFunctions = require("./error-reporting");
const errorEmailUsageFunctions = require("./error-email-usage");
const cleanupWhazzupPdfFunctions = require("./cleanup-whazzup-pdfs");

Object.assign(
  exports,
  existingFunctions,
  userNotificationFunctions,
  centralNotificationFunctions,
  sharedCalendarViewFunctions,
  sharedOperationalViewFunctions,
  adminPasswordPrivateRequestFunctions,
  impiantoChangeIndexFunctions,
  errorReportingFunctions,
  errorEmailUsageFunctions,
  cleanupWhazzupPdfFunctions
);