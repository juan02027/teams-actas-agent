import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { makeLocalDocuments } from "./localDocuments";

const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "teams-actas-agent-"));
process.env.LOCAL_DATA_DIR = tempDir;
const store = await import("./localStore");

describe("local recording index", () => {
  it("stores WebM metadata and removes only the history reference", () => {
    const meeting = store.createLocalMeeting({ title: "Prueba de descarga" });
    const bytes = Buffer.from("webm-test-data");
    const saved = store.saveLocalRecording({ meetingId: meeting.id, title: meeting.title, mimeType: "video/webm", buffer: bytes, durationSeconds: 3600 });
    expect(saved.recording.fileName).toMatch(/\.webm$/);
    expect(fs.readFileSync(saved.recording.filePath)).toEqual(bytes);
    expect(store.listLocalRecordings()).toHaveLength(1);
    store.clearLocalRecordingReference(meeting.id);
    expect(store.listLocalRecordings()).toHaveLength(0);
    expect(fs.existsSync(saved.recording.filePath)).toBe(true);
  });

  it("keeps only explicit commitments and tracks status", () => {
    const meeting = store.createLocalMeeting({ title: "Prueba de compromisos" });
    const docs = store.saveLocalDocuments({ meetingId: meeting.id, documents: [{ kind: "minutes", format: "pdf", fileName: "acta.pdf", bytes: Buffer.from("pdf") }], commitments: [{ personName: "Ana", action: "Entregar informe", dueDate: "2026-09-10", evidence: "Ana se compromete a entregarlo el 10 de septiembre", confidence: "high" }] });
    expect(docs[0].storageUrl).toContain("/api/local-documents/");
    const item = store.listLocalCommitments()[0];
    expect(item.personName).toBe("Ana");
    store.updateLocalCommitmentStatus(item.id, "done");
    expect(store.listLocalCommitments()[0].status).toBe("done");
    store.deleteLocalCommitment(item.id);
    expect(store.listLocalCommitments()).toHaveLength(0);
  });

  it("generates non-empty DOCX and PDF files", async () => {
    const generated = await makeLocalDocuments({ meetingTitle: "Prueba de documentos", output: { executiveSummary: "Resumen breve", objective: "Validar salida", decisions: ["Continuar"], openTopics: [], risks: [], commitments: [] } });
    expect(generated.documents.every((document) => document.bytes.length > 100)).toBe(true);
    expect(generated.documents.map((document) => document.format)).toEqual(["docx", "docx", "pdf"]);
  });
});
