import { buildApp } from "./app.js";
import { closeDatabase, pool } from "./lib/db.js";
import { env } from "./lib/env.js";

const app = buildApp();

async function shutdown(signal: string) {
  app.log.info({ signal }, "Shutting down");
  await app.close();
  await closeDatabase();
  process.exit(0);
}

process.on("SIGTERM", () => void shutdown("SIGTERM"));
process.on("SIGINT", () => void shutdown("SIGINT"));

try {
  await app.listen({
    host: env.HOST,
    port: env.PORT,
  });

  try {
    const result = await pool.query<{
      database: string;
      database_user: string;
    }>(
      `SELECT
         current_database() AS database,
         current_user AS database_user`,
    );
    app.log.info(
      {
        database: result.rows[0]?.database,
        databaseUser: result.rows[0]?.database_user,
      },
      "Database connection established",
    );
  } catch (error) {
    const normalizedError =
      error instanceof Error ? error : new Error("Unknown database error");
    const errorCode =
      typeof error === "object" &&
      error !== null &&
      "code" in error &&
      typeof error.code === "string"
        ? error.code
        : undefined;

    app.log.error(
      {
        err: {
          name: normalizedError.name,
          message: normalizedError.message,
          code: errorCode,
        },
      },
      "Database connection failed; API is running in not-ready state",
    );
  }
} catch (error) {
  app.log.fatal(error, "Failed to start API");
  await closeDatabase();
  process.exit(1);
}
