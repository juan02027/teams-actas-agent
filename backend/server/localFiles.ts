import type { Express, Response } from "express";
import { getLocalDocument, getLocalRecording } from "./localStore";

function sendFile(res: Response, filePath: string, fileName: string, mimeType: string) {
  res.type(mimeType);
  res.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`);
  res.setHeader("X-Content-Type-Options", "nosniff");
  res.download(filePath, fileName, (error: unknown) => {
    if (error && !res.headersSent) res.status(404).json({ error: "file_not_found" });
  });
}

export function registerLocalFileRoutes(app: Express) {
  app.get("/api/local-recordings/:meetingId/download", (req, res) => {
    const meetingId = Number(req.params.meetingId);
    const recording = Number.isInteger(meetingId) ? getLocalRecording(meetingId) : undefined;
    if (!recording) return res.status(404).json({ error: "recording_not_found" });
    return sendFile(res, recording.filePath, recording.fileName, recording.mimeType || "video/webm");
  });

  app.get("/api/local-documents/:id/download", (req, res) => {
    const id = Number(req.params.id);
    const document = Number.isInteger(id) ? getLocalDocument(id) : undefined;
    if (!document) return res.status(404).json({ error: "document_not_found" });
    const mimeType = document.format === "pdf" ? "application/pdf" : "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    return sendFile(res, document.filePath, document.fileName, mimeType);
  });
}
