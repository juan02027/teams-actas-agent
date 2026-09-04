import { and, desc, eq } from "drizzle-orm";
import { drizzle } from "drizzle-orm/mysql2";
import {
  AgentSettings,
  InsertUser,
  agentSettings,
  commitments,
  meetingDocuments,
  meetings,
  users,
} from "../drizzle/schema";
import { ENV } from "./_core/env";

let _db: ReturnType<typeof drizzle> | null = null;

export async function getDb() {
  if (!_db && process.env.DATABASE_URL) {
    try {
      _db = drizzle(process.env.DATABASE_URL);
    } catch (error) {
      console.warn("[Database] Failed to connect:", error);
      _db = null;
    }
  }
  return _db;
}

export async function upsertUser(user: InsertUser): Promise<void> {
  if (!user.openId) throw new Error("User openId is required for upsert");
  const db = await getDb();
  if (!db) return;

  const values: InsertUser = { openId: user.openId };
  const updateSet: Record<string, unknown> = {};
  const textFields = ["name", "email", "loginMethod"] as const;
  for (const field of textFields) {
    if (user[field] !== undefined) {
      values[field] = user[field] ?? null;
      updateSet[field] = user[field] ?? null;
    }
  }
  if (user.lastSignedIn !== undefined) {
    values.lastSignedIn = user.lastSignedIn;
    updateSet.lastSignedIn = user.lastSignedIn;
  }
  if (user.role !== undefined) {
    values.role = user.role;
    updateSet.role = user.role;
  } else if (user.openId === ENV.ownerOpenId) {
    values.role = "admin";
    updateSet.role = "admin";
  }
  if (!values.lastSignedIn) values.lastSignedIn = new Date();
  if (Object.keys(updateSet).length === 0) updateSet.lastSignedIn = new Date();
  await db.insert(users).values(values).onDuplicateKeyUpdate({ set: updateSet });
}

export async function getUserByOpenId(openId: string) {
  const db = await getDb();
  if (!db) return undefined;
  const result = await db.select().from(users).where(eq(users.openId, openId)).limit(1);
  return result[0];
}

export async function listMeetings() {
  const db = await getDb();
  if (!db) return [];
  return db.select().from(meetings).orderBy(desc(meetings.scheduledAt)).limit(100);
}

export async function listCommitments() {
  const db = await getDb();
  if (!db) return [];
  return db.select().from(commitments).orderBy(desc(commitments.createdAt)).limit(200);
}

export async function getMeetingById(id: number) {
  const db = await getDb();
  if (!db) return undefined;
  const result = await db.select().from(meetings).where(eq(meetings.id, id)).limit(1);
  return result[0];
}

export async function getMeetingByGraphId(graphMeetingId: string) {
  const db = await getDb();
  if (!db) return undefined;
  const result = await db.select().from(meetings).where(eq(meetings.graphMeetingId, graphMeetingId)).limit(1);
  return result[0];
}

export async function upsertMeetingFromTeams(input: { graphMeetingId: string; title: string; organizerName: string; organizerEmail?: string; scheduledAt: Date; attendeesCount: number }) {
  const db = await getDb();
  if (!db) throw new Error("Database is not available");
  const existing = await getMeetingByGraphId(input.graphMeetingId);
  if (existing) return existing;
  await db.insert(meetings).values({
    graphMeetingId: input.graphMeetingId,
    title: input.title,
    organizerName: input.organizerName,
    organizerEmail: input.organizerEmail || null,
    scheduledAt: input.scheduledAt,
    attendeesCount: input.attendeesCount,
    recordingEnabled: 1,
    processingEnabled: 1,
    status: "processing",
  });
  return getMeetingByGraphId(input.graphMeetingId);
}

export async function updateMeetingControls(id: number, recordingEnabled: boolean, processingEnabled: boolean) {
  const db = await getDb();
  if (!db) return { id, recordingEnabled: Number(recordingEnabled), processingEnabled: Number(processingEnabled) };
  await db.update(meetings).set({ recordingEnabled: Number(recordingEnabled), processingEnabled: Number(processingEnabled), updatedAt: new Date() }).where(eq(meetings.id, id));
  const result = await db.select().from(meetings).where(eq(meetings.id, id)).limit(1);
  return result[0];
}

export async function markMeetingProcessing(id: number) {
  const db = await getDb();
  if (!db) return;
  await db.update(meetings).set({ status: "processing", updatedAt: new Date() }).where(eq(meetings.id, id));
}

export async function saveMeetingOutput(input: { meetingId: number; documents: Array<{ kind: "minutes" | "commitments"; format: "docx" | "pdf"; fileName: string; storageKey: string; storageUrl: string }>; commitmentsToSave: Array<{ personName: string; personEmail?: string; action: string; dueDate?: string }> }) {
  const db = await getDb();
  if (!db) throw new Error("Database is not available");
  await db.insert(meetingDocuments).values(input.documents.map((document) => ({ meetingId: input.meetingId, ...document })));
  if (input.commitmentsToSave.length > 0) {
    await db.insert(commitments).values(input.commitmentsToSave.map((commitment) => ({ meetingId: input.meetingId, personName: commitment.personName, personEmail: commitment.personEmail || null, action: commitment.action, dueDate: commitment.dueDate || null })));
  }
  await db.update(meetings).set({ status: "review", updatedAt: new Date() }).where(eq(meetings.id, input.meetingId));
}

export async function getAgentSettings(): Promise<AgentSettings | undefined> {
  const db = await getDb();
  if (!db) return undefined;
  const result = await db.select().from(agentSettings).limit(1);
  return result[0];
}

export async function saveAgentSettings(input: Partial<AgentSettings>) {
  const db = await getDb();
  if (!db) return input;
  const current = await getAgentSettings();
  if (current) {
    await db.update(agentSettings).set({ ...input, updatedAt: new Date() }).where(eq(agentSettings.id, current.id));
    const next = await db.select().from(agentSettings).where(eq(agentSettings.id, current.id)).limit(1);
    return next[0];
  }
  await db.insert(agentSettings).values({ companyName: input.companyName || "Toda la empresa", recordByDefault: input.recordByDefault ?? 0, processByDefault: input.processByDefault ?? 1, requireReview: input.requireReview ?? 1, destination: input.destination || "SharePoint / Actas de reuniones", aiModel: input.aiModel || "gpt-5-mini" });
  return getAgentSettings();
}
