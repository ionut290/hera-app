"use strict";

const existingFunctions = require("./index");
const userNotificationFunctions = require("./user-notifications");
const centralNotificationFunctions = require("./central-notifications");

Object.assign(exports, existingFunctions, userNotificationFunctions, centralNotificationFunctions);
