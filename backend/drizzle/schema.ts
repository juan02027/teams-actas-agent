import { index, int, mysqlEnum, mysqlTable, text, timestamp, varchar } from "drizzle-orm/mysql-core";

/** Core user table backing the Manus OAuth flow. */
export const users = mysqlTable("users", {
  id: int("id").autoincrement().primaryKey(),
  openId: varchar("openId", { length: 64 }).notNull().unique(),
  name: text("name"),
  email: varchar("email", { length: 320 }),
  loginMethod: varchar("loginMethod", { length: 64 }),
  role: mysqlEnum("role", ["user", "admin"]).default("user").notNull(),
  createdAt: timestamp("createdAt").defaultNow().notNull(),
  updatedAt: timestamp("updatedAt").defaultNow().onUpdateNow().notNull(),
  lastSignedIn: timestamp("lastSignedIn").defaultNow().notNull(),
});

export type User = typeof users.$inferSelect;
export type InsertUser = typeof users.$inferInsert;

export const meetings = mysqlTable(
  "meetings",
  {
    id: int("id").autoincrement().primaryKey(),
    graphMeetingId: varchar("graphMeetingId", { length: 180 }).notNull().unique(),
    title: varchar("title", { length: 255 }).notNull(),
    organizerName: varchar("organizerName", { length: 180 }).notNull(),
    organizerEmail: varchar("organizerEmail", { length: 320 }),
    scheduledAt: timestamp("scheduledAt").notNull(),
    durationMinutes: int("durationMinutes").default(0).notNull(),
    attendeesCount: int("attendeesCount").default(0).notNull(),
    recordingEnabled: int("recordingEnabled").default(0).notNull(),
    processingEnabled: int("processingEnabled").default(1).notNull(),
    status: mysqlEnum("status", ["scheduled", "recording", "processing", "ready", "review", "error"]).default("scheduled").notNull(),
    transcriptUrl: text("transcriptUrl"),
    recordingUrl: text("recordingUrl"),
    createdAt: timestamp("createdAt").defaultNow().notNull(),
    updatedAt: timestamp("updatedAt").defaultNow().onUpdateNow().notNull(),
  },
  (table) => ({ graphMeetingIdx: index("meetings_graph_meeting_idx").on(table.graphMeetingId) }),
);

export type Meeting = typeof meetings.$inferSelect;
export type InsertMeeting = typeof meetings.$inferInsert;

export const meetingDocuments = mysqlTable("meetingDocuments", {
  id: int("id").autoincrement().primaryKey(),
  meetingId: int("meetingId").notNull(),
  kind: mysqlEnum("kind", ["minutes", "commitments"]).notNull(),
  format: mysqlEnum("format", ["docx", "pdf"]).notNull(),
  fileName: varchar("fileName", { length: 255 }).notNull(),
  storageKey: text("storageKey"),
  storageUrl: text("storageUrl"),
  generatedAt: timestamp("generatedAt").defaultNow().notNull(),
}, (table) => ({ meetingIdx: index("meeting_documents_meeting_idx").on(table.meetingId) }));

export type MeetingDocument = typeof meetingDocuments.$inferSelect;
export type InsertMeetingDocument = typeof meetingDocuments.$inferInsert;

export const commitments = mysqlTable("commitments", {
  id: int("id").autoincrement().primaryKey(),
  meetingId: int("meetingId").notNull(),
  personName: varchar("personName", { length: 180 }).notNull(),
  personEmail: varchar("personEmail", { length: 320 }),
  action: text("action").notNull(),
  dueDate: varchar("dueDate", { length: 64 }),
  status: mysqlEnum("status", ["open", "in_progress", "done", "blocked"]).default("open").notNull(),
  createdAt: timestamp("createdAt").defaultNow().notNull(),
}, (table) => ({ meetingIdx: index("commitments_meeting_idx").on(table.meetingId) }));

export type Commitment = typeof commitments.$inferSelect;
export type InsertCommitment = typeof commitments.$inferInsert;

export const agentSettings = mysqlTable("agentSettings", {
  id: int("id").autoincrement().primaryKey(),
  companyName: varchar("companyName", { length: 180 }).default("Toda la empresa").notNull(),
  recordByDefault: int("recordByDefault").default(0).notNull(),
  processByDefault: int("processByDefault").default(1).notNull(),
  requireReview: int("requireReview").default(1).notNull(),
  destination: varchar("destination", { length: 255 }).default("SharePoint / Actas de reuniones").notNull(),
  aiModel: varchar("aiModel", { length: 120 }).default("gpt-5-mini").notNull(),
  updatedAt: timestamp("updatedAt").defaultNow().onUpdateNow().notNull(),
});

export type AgentSettings = typeof agentSettings.$inferSelect;
export type InsertAgentSettings = typeof agentSettings.$inferInsert;
