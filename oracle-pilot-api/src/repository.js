import pg from "pg";

const { Pool } = pg;
const COLLECTIONS = new Set(["commesse", "impianti"]);

export function createPool(config) {
  return new Pool({
    connectionString: config.databaseUrl,
    ssl: config.dbSsl ? { rejectUnauthorized: true } : false,
    max: 10,
    idleTimeoutMillis: 10_000,
    connectionTimeoutMillis: 3_000,
    statement_timeout: 5_000,
    query_timeout: 6_000,
    application_name: "varga-oracle-pilot-api"
  });
}

export function parseListQuery(query) {
  const collection = String(query.collection || "commesse").toLowerCase();
  if (!COLLECTIONS.has(collection)) throw Object.assign(new Error("invalid collection"), { status: 400 });
  const limit = Math.min(Math.max(Number.parseInt(query.limit || "100", 10) || 100, 1), 500);
  const offset = Math.max(Number.parseInt(query.offset || "0", 10) || 0, 0);
  if (offset > 100_000) throw Object.assign(new Error("offset too large"), { status: 400 });
  return { collection, limit, offset };
}

export function createRepository(pool) {
  return {
    async ready() { await pool.query("SELECT 1"); },
    async catalog({ collection, limit, offset }) {
      const result = await pool.query(
        `SELECT collection_name, document_id, denominazione, comune, indirizzo, stato, eliminato
           FROM public.varga_catalogo_api_v1
          WHERE collection_name = $1
          ORDER BY document_id
          LIMIT $2 OFFSET $3`,
        [collection, limit, offset]
      );
      return result.rows;
    },
    async map({ collection, limit, offset }) {
      const result = await pool.query(
        `SELECT collection_name, document_id, latitudine, longitudine
           FROM public.varga_mappa_api_v1
          WHERE collection_name = $1
          ORDER BY document_id
          LIMIT $2 OFFSET $3`,
        [collection, limit, offset]
      );
      return result.rows;
    },
    async close() { await pool.end(); }
  };
}
