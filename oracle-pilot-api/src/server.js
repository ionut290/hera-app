import { createApp } from "./app.js";
import { loadConfig } from "./config.js";
import { createPool, createRepository } from "./repository.js";

const config = loadConfig();
const repository = createRepository(createPool(config));
const app = createApp({ config, repository });
const server = app.listen(config.port, config.host, () => {
  console.log(JSON.stringify({ level: "info", message: "pilot API listening", host: config.host, port: config.port }));
});

async function shutdown(signal) {
  console.log(JSON.stringify({ level: "info", message: "shutdown", signal }));
  server.close(async () => { await repository.close(); process.exit(0); });
  setTimeout(() => process.exit(1), 10_000).unref();
}
process.on("SIGINT", () => shutdown("SIGINT"));
process.on("SIGTERM", () => shutdown("SIGTERM"));
