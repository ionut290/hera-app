"use strict";

const existingFunctions = require("./index");
const userNotificationFunctions = require("./user-notifications");

Object.assign(exports, existingFunctions, userNotificationFunctions);
