"use strict";

const existingFunctions = require("./index");
const userNotificationFunctions = require("./user-notifications");
const centralNotificationFunctions = require("./central-notifications");
const sharedCalendarViewFunctions = require("./shared-calendar-view");
const sharedOperationalViewFunctions = require("./shared-operational-views");
const operatorUsernameLoginFunctions = require("./operator-username-login");

Object.assign(
  exports,
  existingFunctions,
  userNotificationFunctions,
  centralNotificationFunctions,
  sharedCalendarViewFunctions,
  sharedOperationalViewFunctions,
  operatorUsernameLoginFunctions
);
