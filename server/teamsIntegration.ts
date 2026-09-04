import type { Request, Response } from "express";
import { processMeetingTranscript } from "./meetingPipeline";

function timingSafeEqual(left: string, right: string) {
  if (!left || !right || left.length !== right.length) return false;
  let result = 0;
  for (let i = 0; i < left.length; i += 1) result |= left.charCodeAt(i) ^ right.charCodeAt(i);
  return result === 0;
}

export function teamsGraphNotificationHandler(req: Request, res: Response) {
  const validationToken = typeof req.query.validationToken === "string" ? req.query.validationToken : undefined;
  if (validationToken) return res.type("text/plain").status(200).send(validationToken);

  const expectedState = process.env.TEAMS_GRAPH_CLIENT_STATE;
  const notifications = Array.isArray(req.body?.value) ? req.body.value : [];
  if (expectedState && notifications.some((item: { clientState?: string }) => !timingSafeEqual(item.clientState || "", expectedState))) return res.status(401).json({ error: "invalid_client_state" });
  return res.sendStatus(202);
}

export async function powerAutomateTranscriptHandler(req: Request, res: Response) {
  const expectedToken = process.env.TEAMS_AGENT_INGEST_TOKEN;
  const receivedToken = req.header("x-teams-agent-token") || "";
  if (!expectedToken || !timingSafeEqual(receivedToken, expectedToken)) return res.status(401).json({ error: "invalid_ingest_token" });

  const { graphMeetingId, title, organizerName, organizerEmail, scheduledAt, attendeesCount, transcript } = req.body || {};
  if (typeof graphMeetingId !== "string" || typeof transcript !== "string" || transcript.length < 20) return res.status(400).json({ error: "graphMeetingId_and_transcript_are_required" });

  try {
    const result = await processMeetingTranscript({ graphMeetingId, title, organizerName, organizerEmail, scheduledAt, attendeesCount, transcript });
    return res.status(200).json({ ok: true, ...result });
  } catch (error) {
    console.error("[Teams] transcript processing failed", error);
    return res.status(500).json({ error: error instanceof Error ? error.message : "processing_failed" });
  }
}
