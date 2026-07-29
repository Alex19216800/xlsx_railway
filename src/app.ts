import cors from "@fastify/cors";
import helmet from "@fastify/helmet";
import rateLimit from "@fastify/rate-limit";
import Fastify from "fastify";
import { requireApiKey } from "./lib/auth.js";
import { env } from "./lib/env.js";
import { departmentRoutes } from "./routes/departments.js";
import { employeeRoutes } from "./routes/employees.js";
import { healthRoutes } from "./routes/health.js";

export function buildApp() {
  const app = Fastify({
    logger: {
      level: env.LOG_LEVEL,
      redact: {
        paths: [
          "req.headers.authorization",
          "req.headers.x-api-key",
          "headers.authorization",
          "headers.x-api-key",
        ],
        censor: "[REDACTED]",
      },
    },
    trustProxy: true,
    bodyLimit: 1_048_576,
  });

  app.register(helmet, {
    contentSecurityPolicy: false,
  });

  app.register(rateLimit, {
    max: 120,
    timeWindow: "1 minute",
  });

  app.register(cors, {
    credentials: false,
    methods: ["GET", "OPTIONS"],
    allowedHeaders: ["Content-Type", "X-API-Key"],
    origin(origin, callback) {
      if (!origin || env.ALLOWED_ORIGINS.includes(origin)) {
        callback(null, true);
        return;
      }
      callback(null, false);
    },
  });

  app.addHook("onRequest", async (request, reply) => {
    if (request.url.startsWith("/api/")) {
      return requireApiKey(request, reply);
    }
  });

  app.addHook("onSend", async (request, reply, payload) => {
    if (request.url.startsWith("/api/")) {
      reply.header("Cache-Control", "no-store");
    }
    return payload;
  });

  app.register(healthRoutes);
  app.register(departmentRoutes);
  app.register(employeeRoutes);

  app.setNotFoundHandler((_request, reply) => {
    return reply.code(404).send({
      error: "not_found",
    });
  });

  app.setErrorHandler((error, request, reply) => {
    const normalizedError =
      error instanceof Error ? error : new Error("Unknown error");

    request.log.error(
      {
        err: {
          name: normalizedError.name,
          message: normalizedError.message,
          stack:
            env.NODE_ENV === "development"
              ? normalizedError.stack
              : undefined,
        },
      },
      "Request failed",
    );

    return reply.code(500).send({
      error: "internal_server_error",
      message: "The request could not be completed.",
    });
  });

  return app;
}
