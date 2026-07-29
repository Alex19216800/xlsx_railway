import pg from "pg";
import { env } from "./env.js";

const { Pool } = pg;

export const pool = new Pool({
  connectionString: env.DATABASE_URL,
  max: env.DATABASE_POOL_MAX,
  idleTimeoutMillis: 30_000,
  connectionTimeoutMillis: 10_000,
  ssl: env.DATABASE_SSL
    ? {
        rejectUnauthorized: env.DATABASE_SSL_REJECT_UNAUTHORIZED,
      }
    : false,
});

pool.on("error", (error) => {
  console.error("Unexpected PostgreSQL pool error", {
    name: error.name,
    message: error.message,
  });
});

export async function closeDatabase() {
  await pool.end();
}
