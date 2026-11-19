"use strict";

// Browser-free behavioral regression for the actual chat listener function.
// Fake documents only: never reads or writes production Firebase data.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const path = require("node:path");

const app = fs.readFileSync(path.resolve(__dirname, "../app.js"), "utf8");
const start = app.indexOf("function subscribeChat() {");
const end = app.indexOf("async function notifyForNewChatMessages(", start);
assert.ok(start >= 0 && end > start, "Cannot isolate subscribeChat for regression test.");
const source = app.slice(start, end);
const listeners = new Map();
const stopped = [];
const notified = [];
const rendered = [];
const ui = { chatFeedback: { textContent: "" } };

function key(field, value) {
  return field + ":" + value;
}
const db = {
  collection(name) {
    assert.equal(name, "chatMessages");
    return {
      where(field, op, value) {
        assert.equal(op, "==");
        return {
          orderBy(orderField, direction) {
            assert.equal(orderField, "createdAt");
            assert.equal(direction, "desc");
            return {
              limit(n) {
                assert.equal(n, 500);
                return {
                  onSnapshot(onData, onError) {
                    const id = key(field, value);
                    assert.ok(!listeners.has(id), "Duplicate scoped subscription: " + id);
                    listeners.set(id, { onData, onError });
                    return () => stopped.push(id);
                  }
                };
              }
            };
          }
        };
      }
    };
  }
};
function doc(id, createdAt, data = {}) {
  return { id, data: () => ({ createdAt, ...data }) };
}
function emit(field, value, docs, changes = []) {
  const l = listeners.get(key(field, value));
  assert.ok(l, "Subscription missing for " + field + ":" + value);
  l.onData({ docs, docChanges: () => changes });
}

const context = {
  db,
  currentUser: null,
  chatMessages: [],
  chatNotificationsInitialized: false,
  unsubscribeChat: null,
  renderChat(messages) { rendered.push(messages.map(message => message.id)); },
  notifyForNewChatMessages(changes) { notified.push(changes); },
  getChatMessageCreatedAtMs(message) { return message.createdAt; },
  ui,
  console: { error() {} }
};
vm.runInNewContext(source + "\nsubscribeChat();", context, { filename: "app.js:subscribeChat" });
assert.equal(listeners.size, 0, "Anonymous client should not start chat listeners.");

context.currentUser = { uid: "synthetic-user-a" };
vm.runInNewContext("subscribeChat();", context, { filename: "app.js:subscribeChat" });
assert.deepEqual(Array.from(listeners.keys()).sort(), [
  "recipientId:", "recipientId:synthetic-user-a", "senderId:synthetic-user-a"
].sort(), "Chat queries must be scoped to public, sender, recipient.");
const a = doc("message-a", 10, { senderId: "synthetic-user-a", recipientId: "synthetic-user-b" });
const b = doc("message-b", 20, { senderId: "synthetic-user-b", recipientId: "synthetic-user-a" });
const c = doc("public", 5, { senderId: "synthetic-user-b", recipientId: "" });
emit("recipientId", "synthetic-user-a", [b]);
emit("recipientId", "", [c]);
emit("senderId", "synthetic-user-a", [a, b]); // duplicate ID across streams
assert.deepEqual(Array.from(context.chatMessages, m => m.id), ["public", "message-a", "message-b"]);
assert.equal(notified.length, 0, "Initial snapshots may not create notifications.");
assert.equal(context.chatNotificationsInitialized, true);

const d = doc("message-d", 30, { senderId: "synthetic-user-b", recipientId: "synthetic-user-a" });
emit("recipientId", "synthetic-user-a", [b, d], [{ type: "added", doc: d }]);
assert.deepEqual(Array.from(context.chatMessages, m => m.id), ["public", "message-a", "message-b", "message-d"]);
assert.equal(notified.length, 1, "Subsequent changes should create one notification batch.");
assert.equal(notified[0].length, 1);

const oldRendered = rendered.length;
context.currentUser = { uid: "synthetic-other" };
emit("recipientId", "synthetic-user-a", [d]);
assert.equal(rendered.length, oldRendered, "Old session listener must ignore changed account.");

context.currentUser = { uid: "synthetic-user-a" };
listeners.get("senderId:synthetic-user-a").onError(new Error("mock permission denial"));
assert.match(ui.chatFeedback.textContent, /Errore caricamento chat/);
assert.equal(typeof context.unsubscribeChat, "function");
context.unsubscribeChat();
assert.equal(stopped.length, 3, "Sign-out must close all scoped listeners.");
const afterStop = rendered.length;
emit("recipientId", "", [c, doc("late", 99)]);
assert.equal(rendered.length, afterStop, "No rendering after listener shutdown.");
ui.chatFeedback.textContent = "";
listeners.get("recipientId:").onError(new Error("ignore after disposal"));
assert.equal(ui.chatFeedback.textContent, "", "No stale error after logout.");

console.log("PASS: scoped chat streams, deduplication, initialization, session isolation and teardown.");
