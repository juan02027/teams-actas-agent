import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";

export type LocalMeeting = {
  id: number;
  graphMeetingId: string;
  title: string;
  organizerName: string;
  organizerEmail: string | null;
  scheduledAt: string;
  durationMinutes: number;
  attendeesCount: number;
  recordingEnabled: number;
  processingEnabled: number;
  status: "scheduled" | "recording" | "processing" | "ready" | "review" | "error";
  transcriptUrl: string | null;
  recordingUrl: string | null;
  recordingFileName: string | null;
  recordingMimeType: string | null;
  recordingSizeBytes: number | null;
  errorMessage: string | null;
  createdAt: string;
  updatedAt: string;
};

export type LocalRecording = {
  id: number;
  meetingId: number;
  filePath: string;
  fileName: string;
  mimeType: string;
  sizeBytes: number;
  durationSeconds: number | null;
  processingFilePath?: string;
  processingMimeType?: string;
  createdAt: string;
};

export type LocalDocument = {
  id: number;
  meetingId: number;
  kind: "minutes" | "commitments";
  format: "docx" | "pdf";
  fileName: string;
  filePath: string;
  storageKey: string | null;
  storageUrl: string;
  generatedAt: string;
};

export type LocalCommitment = {
  id: number;
  meetingId: number;
  personName: string;
  personEmail: string | null;
  action: string;
  dueDate: string;
  status: "open" | "in_progress" | "done" | "blocked";
  evidence?: string;
  confidence?: "high" | "medium" | "low";
  createdAt: string;
};

type LocalSettings = {
  companyName: string;
  recordByDefault: number;
  processByDefault: number;
  requireReview: number;
  destination: string;
  aiModel: string;
};

type LocalState = {
  version: 2;
  nextId: number;
  meetings: LocalMeeting[];
  recordings: LocalRecording[];
  documents: LocalDocument[];
  commitments: LocalCommitment[];
  settings?: LocalSettings;
};

const DATA_DIR = path.resolve(process.env.LOCAL_DATA_DIR || path.join(process.cwd(), "data"));
const INDEX_PATH = path.join(DATA_DIR, "recordings-index.json");
const RECORDINGS_DIR = path.join(DATA_DIR, "recordings");
const DOCUMENTS_DIR = path.join(DATA_DIR, "documents");

function now() { return new Date().toISOString(); }
function ensureDirs() {
  fs.mkdirSync(RECORDINGS_DIR, { recursive: true });
  fs.mkdirSync(DOCUMENTS_DIR, { recursive: true });
}
function emptyState(): LocalState {
  return { version: 2, nextId: 1, meetings: [], recordings: [], documents: [], commitments: [] };
}
function readState(): LocalState {
  ensureDirs();
  try {
    const parsed = JSON.parse(fs.readFileSync(INDEX_PATH, "utf8")) as Partial<LocalState>;
    return {
      ...emptyState(),
      ...parsed,
      meetings: Array.isArray(parsed.meetings) ? parsed.meetings : [],
      recordings: Array.isArray(parsed.recordings) ? parsed.recordings : [],
      documents: Array.isArray(parsed.documents) ? parsed.documents : [],
      commitments: Array.isArray(parsed.commitments) ? parsed.commitments : [],
      nextId: typeof parsed.nextId === "number" ? parsed.nextId : 1,
    };
  } catch {
    return emptyState();
  }
}
function writeState(state: LocalState) {
  ensureDirs();
  const tmp = `${INDEX_PATH}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(state, null, 2), "utf8");
  fs.renameSync(tmp, INDEX_PATH);
}
function nextId(state: LocalState) { const id = state.nextId; state.nextId += 1; return id; }
function safeName(value: string) {
  return value.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9-_ ]/g, "").trim().replace(/\s+/g, "-").slice(0, 70) || "reunion";
}
function extForMime(mimeType: string) {
  return mimeType.startsWith("video/") ? "webm" : mimeType.includes("mp4") ? "m4a" : "webm";
}

export function localDataDirectory() { ensureDirs(); return DATA_DIR; }

export function listLocalMeetings() {
  return readState().meetings.sort((a, b) => b.scheduledAt.localeCompare(a.scheduledAt));
}
export function listLocalCommitments() {
  return readState().commitments.sort((a, b) => b.createdAt.localeCompare(a.createdAt)).map((item) => {
    const meeting = readState().meetings.find((candidate) => candidate.id === item.meetingId);
    return { ...item, meetingTitle: meeting?.title || "Reunión local" };
  });
}
export function getLocalMeetingById(id: number) { return readState().meetings.find((meeting) => meeting.id === id); }
export function getLocalMeetingByGraphId(graphMeetingId: string) { return readState().meetings.find((meeting) => meeting.graphMeetingId === graphMeetingId); }

export function createLocalMeeting(input: { title: string; graphMeetingId?: string; organizerName?: string; organizerEmail?: string; scheduledAt?: Date | string; durationMinutes?: number; attendeesCount?: number }) {
  const state = readState();
  const graphId = input.graphMeetingId || `local-${crypto.randomUUID()}`;
  const existing = state.meetings.find((meeting) => meeting.graphMeetingId === graphId);
  if (existing) return existing;
  const timestamp = now();
  const meeting: LocalMeeting = {
    id: nextId(state), graphMeetingId: graphId, title: input.title || "Reunión grabada localmente",
    organizerName: input.organizerName || "Operador local", organizerEmail: input.organizerEmail || null,
    scheduledAt: new Date(input.scheduledAt || timestamp).toISOString(), durationMinutes: input.durationMinutes || 0,
    attendeesCount: input.attendeesCount || 0, recordingEnabled: 1, processingEnabled: 1, status: "scheduled",
    transcriptUrl: null, recordingUrl: null, recordingFileName: null, recordingMimeType: null,
    recordingSizeBytes: null, errorMessage: null, createdAt: timestamp, updatedAt: timestamp,
  };
  state.meetings.push(meeting); writeState(state); return meeting;
}

export function updateLocalMeeting(id: number, patch: Partial<LocalMeeting>) {
  const state = readState();
  const meeting = state.meetings.find((candidate) => candidate.id === id);
  if (!meeting) return undefined;
  Object.assign(meeting, patch, { updatedAt: now() });
  writeState(state); return meeting;
}

export function updateLocalMeetingControls(id: number, recordingEnabled: boolean, processingEnabled: boolean) {
  return updateLocalMeeting(id, { recordingEnabled: Number(recordingEnabled), processingEnabled: Number(processingEnabled) });
}

export function saveLocalRecording(input: { meetingId?: number; graphMeetingId?: string; title: string; mimeType: string; buffer: Buffer; processingBuffer?: Buffer; processingMimeType?: string; durationSeconds?: number | null }) {
  const state = readState();
  let meeting = input.meetingId ? state.meetings.find((candidate) => candidate.id === input.meetingId) : undefined;
  const timestamp = now();
  if (!meeting) {
    const graphMeetingId = input.graphMeetingId || `local-${crypto.randomUUID()}`;
    meeting = {
      id: nextId(state), graphMeetingId, title: input.title || "Reunión grabada localmente", organizerName: "Operador local", organizerEmail: null,
      scheduledAt: timestamp, durationMinutes: input.durationSeconds ? Math.round(input.durationSeconds / 60) : 0, attendeesCount: 0,
      recordingEnabled: 1, processingEnabled: 1, status: "processing", transcriptUrl: null, recordingUrl: null, recordingFileName: null,
      recordingMimeType: null, recordingSizeBytes: null, errorMessage: null, createdAt: timestamp, updatedAt: timestamp,
    };
    state.meetings.push(meeting);
  }
  const recordingId = nextId(state);
  const extension = extForMime(input.mimeType);
  const fileName = `${safeName(meeting.title)}-${new Date(timestamp).toISOString().replace(/[:.]/g, "-")}.${extension}`;
  const filePath = path.join(RECORDINGS_DIR, `${recordingId}-${fileName}`);
  fs.writeFileSync(filePath, input.buffer);
  const processingFilePath = input.processingBuffer ? path.join(RECORDINGS_DIR, `${recordingId}-processing.webm`) : undefined;
  if (processingFilePath && input.processingBuffer) fs.writeFileSync(processingFilePath, input.processingBuffer);
  const recording: LocalRecording = { id: recordingId, meetingId: meeting.id, filePath, fileName, mimeType: input.mimeType, sizeBytes: input.buffer.length, durationSeconds: input.durationSeconds ?? null, processingFilePath, processingMimeType: input.processingMimeType, createdAt: timestamp };
  state.recordings = state.recordings.filter((candidate) => candidate.meetingId !== meeting!.id);
  state.recordings.push(recording);
  Object.assign(meeting, { status: "processing", recordingUrl: `/api/local-recordings/${meeting.id}/download`, recordingFileName: fileName, recordingMimeType: input.mimeType, recordingSizeBytes: input.buffer.length, errorMessage: null, updatedAt: timestamp });
  writeState(state);
  return { meeting, recording };
}

export function getLocalRecording(meetingId: number) {
  const state = readState();
  return state.recordings.find((recording) => recording.meetingId === meetingId);
}
export function listLocalRecordings() {
  const state = readState();
  return state.recordings.slice().sort((a, b) => b.createdAt.localeCompare(a.createdAt)).map((recording) => ({
    meeting: state.meetings.find((meeting) => meeting.id === recording.meetingId),
    recording,
    documents: state.documents.filter((document) => document.meetingId === recording.meetingId),
  })).filter((item): item is { meeting: LocalMeeting; recording: LocalRecording; documents: LocalDocument[] } => Boolean(item.meeting));
}

export function clearLocalRecordingReference(meetingId: number) {
  const state = readState();
  state.recordings.filter((recording) => recording.meetingId === meetingId && recording.processingFilePath).forEach((recording) => { try { fs.unlinkSync(recording.processingFilePath!); } catch { /* already removed */ } });
  state.recordings = state.recordings.filter((recording) => recording.meetingId !== meetingId);
  const meeting = state.meetings.find((candidate) => candidate.id === meetingId);
  if (meeting) Object.assign(meeting, { recordingUrl: null, recordingFileName: null, recordingMimeType: null, recordingSizeBytes: null, updatedAt: now() });
  writeState(state); return { meetingId };
}

export function saveLocalDocuments(input: { meetingId: number; documents: Array<{ kind: "minutes" | "commitments"; format: "docx" | "pdf"; fileName: string; bytes: Buffer; storageKey?: string | null }>; commitments: Array<{ personName: string; personEmail?: string; action: string; dueDate: string; evidence?: string; confidence?: "high" | "medium" | "low" }> }) {
  const state = readState();
  state.documents = state.documents.filter((document) => document.meetingId !== input.meetingId);
  state.commitments = state.commitments.filter((commitment) => commitment.meetingId !== input.meetingId);
  const timestamp = now();
  const documents = input.documents.map((item) => {
    const id = nextId(state); const filePath = path.join(DOCUMENTS_DIR, `${id}-${item.fileName}`); fs.writeFileSync(filePath, item.bytes);
    const document: LocalDocument = { id, meetingId: input.meetingId, kind: item.kind, format: item.format, fileName: item.fileName, filePath, storageKey: item.storageKey || null, storageUrl: `/api/local-documents/${id}/download`, generatedAt: timestamp };
    state.documents.push(document); return document;
  });
  input.commitments.forEach((item) => state.commitments.push({ id: nextId(state), meetingId: input.meetingId, personName: item.personName, personEmail: item.personEmail || null, action: item.action, dueDate: item.dueDate, status: "open", evidence: item.evidence, confidence: item.confidence, createdAt: timestamp }));
  const meeting = state.meetings.find((candidate) => candidate.id === input.meetingId);
  if (meeting) Object.assign(meeting, { status: "review", errorMessage: null, updatedAt: timestamp });
  writeState(state); return documents;
}

export function listLocalDocuments() {
  const state = readState();
  return state.documents.slice().sort((a, b) => b.generatedAt.localeCompare(a.generatedAt)).map((document) => ({ ...document, meetingTitle: state.meetings.find((meeting) => meeting.id === document.meetingId)?.title || "Reunión local" }));
}
export function getLocalDocument(id: number) { return readState().documents.find((document) => document.id === id); }
export function deleteLocalDocument(id: number) {
  const state = readState(); state.documents = state.documents.filter((document) => document.id !== id); writeState(state); return { id };
}
export function updateLocalCommitmentStatus(id: number, status: LocalCommitment["status"]) {
  const state = readState(); const item = state.commitments.find((commitment) => commitment.id === id); if (!item) return undefined; item.status = status; writeState(state); return item;
}
export function updateLocalCommitment(id: number, patch: { personName?: string; action?: string; dueDate?: string }) {
  const state = readState(); const item = state.commitments.find((commitment) => commitment.id === id); if (!item) return undefined;
  if (patch.personName?.trim()) item.personName = patch.personName.trim();
  if (patch.action?.trim()) item.action = patch.action.trim();
  if (patch.dueDate?.trim()) item.dueDate = patch.dueDate.trim();
  writeState(state); return item;
}
export function deleteLocalCommitment(id: number) {
  const state = readState(); state.commitments = state.commitments.filter((commitment) => commitment.id !== id); writeState(state); return { id };
}
export function getLocalSettings() { return readState().settings; }
export function saveLocalSettings(input: Partial<LocalSettings>) { const state = readState(); state.settings = { companyName: "Toda la empresa", recordByDefault: 0, processByDefault: 1, requireReview: 1, destination: "SharePoint / Actas de reuniones", aiModel: "groq/compound-mini", ...state.settings, ...input }; writeState(state); return state.settings; }
export function getLocalFile(pathname: string) { return fs.existsSync(pathname) ? pathname : undefined; }
