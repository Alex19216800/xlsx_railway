import { timingSafeEqual } from "node:crypto";
import type { FastifyReply, FastifyRequest } from "fastify";
import { env } from "./env.js";

function tokenMatches(received: string | undefined) {
  if (!received) return false;

  const expectedBuffer = Buffer.from(env.API_ACCESS_TOKEN);
  const receivedBuffer = Buffer.from(received);

  if (expectedBuffer.length !== receivedBuffer.length) return false;
  return timingSafeEqual(expectedBuffer, receivedBuffer);
}

export async function requireApiKey(
  request: FastifyRequest,
  reply: FastifyReply,
) {
  const received = request.headers["x-api-key"];
  const token = Array.isArray(received) ? received[0] : received;

  if (!tokenMatches(token)) {
    return reply.code(401).send({
      error: "unauthorized",
      message: "A valid X-API-Key header is required.",
    });
  }
}
