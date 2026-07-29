import type { FastifyInstance } from "fastify";
import { pool } from "../lib/db.js";

export async function healthRoutes(app: FastifyInstance) {
  // Railway uses this endpoint only to confirm that the HTTP server is alive.
  // Database availability is checked separately by /ready.
  app.get("/health", async () => {
    return {
      status: "ok",
      timestamp: new Date().toISOString(),
    };
  });

  app.get("/ready", async (request, reply) => {
    try {
      const result = await pool.query<{
        database: string;
        database_user: string;
      }>(
        `SELECT
           current_database() AS database,
           current_user AS database_user`,
      );
      const connection = result.rows[0];

      return {
        status: "ready",
        database: "connected",
        databaseName: connection?.database ?? null,
        databaseUser: connection?.database_user ?? null,
        timestamp: new Date().toISOString(),
      };
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

      request.log.error(
        {
          err: {
            name: normalizedError.name,
            message: normalizedError.message,
            code: errorCode,
          },
        },
        "Database readiness check failed",
      );

      return reply.code(503).send({
        status: "not_ready",
        database: "unavailable",
        timestamp: new Date().toISOString(),
      });
    }
  });
}
