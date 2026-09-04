import { TRPCError } from "@trpc/server";
import { buildMeetingFiles } from "./documentGenerator";
import { getMeetingById, getMeetingByGraphId, markMeetingProcessing, upsertMeetingFromTeams } from "./db";
import { generateMeetingDocuments } from "./meetingAgent";

export async function processMeetingTranscript(input: { meetingId?: number; graphMeetingId?: string; title?: string; organizerName?: string; organizerEmail?: string; scheduledAt?: string; attendeesCount?: number; transcript: string }) {
  let meeting = input.meetingId ? await getMeetingById(input.meetingId) : undefined;
  if (!meeting && input.graphMeetingId) meeting = await getMeetingByGraphId(input.graphMeetingId);
  if (!meeting && input.graphMeetingId) {
    meeting = await upsertMeetingFromTeams({
      graphMeetingId: input.graphMeetingId,
      title: input.title || "Reunión de Teams",
      organizerName: input.organizerName || "Sin identificar",
      organizerEmail: input.organizerEmail,
      scheduledAt: input.scheduledAt ? new Date(input.scheduledAt) : new Date(),
      attendeesCount: input.attendeesCount || 0,
    });
  }
  if (!meeting) throw new TRPCError({ code: "NOT_FOUND", message: "No se encontró la reunión para procesar." });
  await markMeetingProcessing(meeting.id);
  const output = await generateMeetingDocuments({ meetingTitle: meeting.title, transcript: input.transcript });
  const files = await buildMeetingFiles({ meetingId: meeting.id, meetingTitle: meeting.title, output });
  return { meetingId: meeting.id, output, files };
}
