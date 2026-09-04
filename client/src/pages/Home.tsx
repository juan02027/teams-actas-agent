import DashboardLayout from "@/components/DashboardLayout";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Separator } from "@/components/ui/separator";
import { Switch } from "@/components/ui/switch";
import { Textarea } from "@/components/ui/textarea";
import { cn } from "@/lib/utils";
import { PublicClientApplication } from "@azure/msal-browser";
import type { AccountInfo } from "@azure/msal-browser";
import { CalendarClock, Check, CircleHelp, Cloud, Download, FileCheck2, FileText, Link2, ListChecks, LockKeyhole, Mic2, RefreshCw, Search, Settings2, ShieldCheck, Sparkles, Trash2, Upload, UsersRound, Video, WandSparkles } from "lucide-react";
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { useLocation } from "wouter";
import { trpc } from "@/lib/trpc";

const microsoftClientId = import.meta.env.VITE_MICROSOFT_CLIENT_ID as string | undefined;
const microsoftTenantId = (import.meta.env.VITE_MICROSOFT_TENANT_ID as string | undefined) || "common";
const publicAppUrl = ((import.meta.env.VITE_PUBLIC_APP_URL as string | undefined) || window.location.origin).replace(/\/$/, "");
const microsoftScopes = ["openid", "profile", "User.Read", "Calendars.Read", "Sites.ReadWrite.All"];
const microsoftApp = microsoftClientId ? new PublicClientApplication({ auth: { clientId: microsoftClientId, authority: `https://login.microsoftonline.com/${microsoftTenantId}`, redirectUri: `${publicAppUrl}/`, postLogoutRedirectUri: `${publicAppUrl}/` } }) : null;

type Attendee = { name: string; email: string; role?: string };
interface Meeting { id: string | number; title: string; organizerName: string; scheduledAt: string; durationMinutes: number; attendeesCount: number; attendees?: Attendee[]; recordingEnabled: boolean; processingEnabled: boolean; status: string; source?: string; }
const demoMeetings: Meeting[] = [
  { id: 101, title: "Comité de operaciones · Q3", organizerName: "Mariana López", scheduledAt: "2026-08-28T15:00:00.000Z", durationMinutes: 60, attendeesCount: 9, recordingEnabled: true, processingEnabled: true, status: "Por revisar" },
  { id: 102, title: "Seguimiento de lanzamiento CRM", organizerName: "Diego Ramírez", scheduledAt: "2026-08-28T17:30:00.000Z", durationMinutes: 45, attendeesCount: 6, recordingEnabled: true, processingEnabled: true, status: "Programada" },
];
const demoCommitments = [{ personName: "Mariana López", action: "Compartir el tablero de capacidad", dueDate: "30 ago", status: "Abierto" }, { personName: "Diego Ramírez", action: "Validar el flujo de aprobación con Legal", dueDate: "02 sep", status: "En curso" }];

function dateLabel(value: string) { return new Date(value).toLocaleDateString("es-MX", { weekday: "short", day: "numeric", month: "short" }); }
function timeLabel(value: string) { return new Date(value).toLocaleTimeString("es-MX", { hour: "2-digit", minute: "2-digit" }); }
function initials(value: string) { return value.split(" ").map((part) => part[0]).slice(0, 2).join("").toUpperCase(); }
function sharePointNote(output: { executiveSummary: string; objective: string; decisions: string[]; openTopics?: string[]; risks?: string[]; commitments: Array<{ personName: string; action: string; dueDate: string }> }, meetingTitle: string, attendees: Attendee[] = []) { const lines = ["Acta de Reunión", "Código: SG-FO-07 - Versión: 10 - Fecha: 10/11/2017", meetingTitle.toUpperCase(), `Fecha: ${new Date().toLocaleDateString("es-CO")} || [hora inicio] - [hora fin]`, `Asistentes (${attendees.length || 0})`, "Nombre completo | Correo | Cargo", ...(attendees.length ? attendees.map((person) => `${person.name} | ${person.email} | ${person.role || "Participante"}`) : ["[No se encontraron asistentes en el calendario]"]), "", "Información de planeación de la reunión", `[Objetivo y contexto] ${output.objective}`, "", "Notas de la Reunión", output.executiveSummary || "[Sin notas relevantes]", "", "Temas abiertos", ...(output.openTopics?.length ? output.openTopics.map((item) => `- ${item}`) : ["[No se identificaron temas abiertos]" ]), "", "Riesgos o bloqueos", ...(output.risks?.length ? output.risks.map((item) => `- ${item}`) : ["[No se identificaron riesgos]" ]), "", "Tareas de la Reunión", "Tarea | Nota | Responsable | Estado | Fecha"]; if (output.commitments.length) lines.push(...output.commitments.map((item) => `${item.action} | Compromiso identificado en la reunión | ${item.personName} | Pendiente | ${item.dueDate}`)); else lines.push("[Sin tareas con responsable y fecha de entrega]"); if (output.decisions.length) lines.push("", "Decisiones", ...output.decisions.map((item) => `- ${item}`)); return lines.join("\n").slice(0, 63999); }
async function resolveSharePointSite(accessToken: string, siteUrl: URL) {
  const headers = { Authorization: `Bearer ${accessToken}` };
  // El usuario puede pegar por error la URL de la vista de la lista
  // (.../Lists/Reuniones Efectivas/AllItems.aspx). Graph necesita únicamente
  // la ruta del sitio (.../transformaciondigital).
  const rawPath = decodeURIComponent(siteUrl.pathname).replace(/\/$/, "") || "/";
  const pathName = rawPath.split(/\/lists\//i)[0].replace(/\/$/, "") || "/";
  const direct = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteUrl.hostname}:${pathName}`, { headers });
  if (direct.ok) return await direct.json() as { id?: string };
  const siteName = decodeURIComponent(pathName.split("/").filter(Boolean).pop() || "Transformación Digital");
  const search = await fetch(`https://graph.microsoft.com/v1.0/sites?search=${encodeURIComponent(siteName)}`, { headers });
  if (search.ok) {
    const payload = await search.json() as { value?: Array<{ id?: string; webUrl?: string; displayName?: string }> };
    const hostnameMatch = (payload.value || []).find((site) => site.webUrl && new URL(site.webUrl).hostname.toLowerCase() === siteUrl.hostname.toLowerCase());
    if (hostnameMatch?.id) return hostnameMatch;
  }
  const detail = await direct.text().catch(() => "");
  throw new Error(`No se pudo localizar el sitio de SharePoint (${direct.status}). Revisa VITE_SHAREPOINT_SITE_URL; Graph recibió ${siteUrl.origin}${pathName}. ${detail}`);
}
async function downloadFile(url: string, fileName: string) { const response = await fetch(url, { credentials: "same-origin" }); if (!response.ok) throw new Error(`No se pudo descargar el archivo (${response.status})`); const blob = await response.blob(); const objectUrl = URL.createObjectURL(blob); const anchor = document.createElement("a"); anchor.href = objectUrl; anchor.download = fileName; anchor.style.display = "none"; document.body.appendChild(anchor); anchor.click(); anchor.remove(); window.setTimeout(() => URL.revokeObjectURL(objectUrl), 1000); }

async function uploadBlobToSharePoint(blob: Blob, fileName: string, folder = "Grabaciones") {
  if (!microsoftApp) throw new Error("Microsoft 365 no está vinculado");
  await microsoftApp.initialize();
  const account = microsoftApp.getAllAccounts()[0];
  if (!account) throw new Error("No hay una sesión de Microsoft 365 activa");
  const token = await microsoftApp.acquireTokenSilent({ account, scopes: ["Sites.ReadWrite.All"] });
  const siteUrl = new URL(import.meta.env.VITE_SHAREPOINT_SITE_URL || "https://abcstorage.sharepoint.com/transformaciondigital");
  const site = await resolveSharePointSite(token.accessToken, siteUrl);
  if (!site.id) throw new Error("Graph no devolvió el sitio de SharePoint");
  const libraryName = import.meta.env.VITE_SHAREPOINT_DOCUMENT_LIBRARY || "Documentos";
  const driveResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(site.id)}/drives`, { headers: { Authorization: `Bearer ${token.accessToken}` } });
  if (!driveResponse.ok) throw new Error(`No se pudieron consultar las bibliotecas de SharePoint (${driveResponse.status})`);
  const drives = await driveResponse.json() as { value?: Array<{ id: string; name?: string; driveType?: string }> };
  const drive = drives.value?.find((item) => item.name?.toLowerCase() === libraryName.toLowerCase()) || drives.value?.find((item) => item.driveType === "documentLibrary");
  if (!drive?.id) throw new Error(`No se encontró la biblioteca de documentos «${libraryName}»`);
  const clean = (value: string) => value.replace(/[\\/:*?"<>|#%]+/g, "-").trim() || "archivo";
  // Se usa la raíz de la biblioteca para que la primera carga no dependa de
  // que existan previamente carpetas en SharePoint. La carpeta recibida se
  // conserva en la función para poder organizarla en una futura migración.
  void folder;
  const path = encodeURIComponent(clean(fileName));
  const sessionResponse = await fetch(`https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(drive.id)}/root:/${path}:/createUploadSession`, { method: "POST", headers: { Authorization: `Bearer ${token.accessToken}`, "Content-Type": "application/json" }, body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "replace" } }) });
  if (!sessionResponse.ok) throw new Error(`No se pudo preparar la carga en SharePoint (${sessionResponse.status})`);
  const session = await sessionResponse.json() as { uploadUrl?: string };
  if (!session.uploadUrl) throw new Error("SharePoint no devolvió la URL de carga");
  const chunkSize = 10 * 320 * 1024;
  for (let start = 0; start < blob.size; start += chunkSize) {
    const end = Math.min(start + chunkSize, blob.size);
    const chunk = blob.slice(start, end);
    const response = await fetch(session.uploadUrl, { method: "PUT", headers: { "Content-Length": String(end - start), "Content-Range": `bytes ${start}-${end - 1}/${blob.size}` }, body: chunk });
    if (!response.ok && response.status !== 202) throw new Error(`SharePoint rechazó un bloque de carga (${response.status})`);
  }
  return `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(drive.id)}/root:/${path}`;
}

async function uploadGeneratedDocuments(documents: Array<{ fileName: string; storageUrl?: string; kind?: string }>) {
  for (const document of documents) {
    if (!document.storageUrl) continue;
    const response = await fetch(document.storageUrl, { credentials: "same-origin" });
    if (!response.ok) throw new Error(`No se pudo leer ${document.fileName} para subirlo`);
    await uploadBlobToSharePoint(await response.blob(), document.fileName, "Documentos");
  }
}
type RemoteSharePointFile = { id: string; name: string; size: number; modified: string; webUrl?: string };
async function getSharePointDrive() {
  if (!microsoftApp) throw new Error("Microsoft 365 no está vinculado");
  await microsoftApp.initialize();
  const account = microsoftApp.getAllAccounts()[0]; if (!account) throw new Error("No hay una sesión de Microsoft 365 activa");
  const token = await microsoftApp.acquireTokenSilent({ account, scopes: ["Sites.ReadWrite.All"] });
  const siteUrl = new URL(import.meta.env.VITE_SHAREPOINT_SITE_URL || "https://abcstorage.sharepoint.com/transformaciondigital");
  const site = await resolveSharePointSite(token.accessToken, siteUrl); const libraryName = import.meta.env.VITE_SHAREPOINT_DOCUMENT_LIBRARY || "Documentos";
  const response = await fetch(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(site.id || "")}/drives`, { headers: { Authorization: `Bearer ${token.accessToken}` } });
  if (!response.ok) throw new Error(`No se pudieron consultar las bibliotecas de SharePoint (${response.status})`);
  const drives = await response.json() as { value?: Array<{ id: string; name?: string; driveType?: string }> };
  const drive = drives.value?.find((item) => item.name?.toLowerCase() === libraryName.toLowerCase()) || drives.value?.find((item) => item.driveType === "documentLibrary");
  if (!drive?.id) throw new Error(`No se encontró la biblioteca «${libraryName}»`);
  return { driveId: drive.id, accessToken: token.accessToken };
}
async function listSharePointFiles() { const { driveId, accessToken } = await getSharePointDrive(); const response = await fetch(`https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root/children?$top=200&$orderby=lastModifiedDateTime%20desc`, { headers: { Authorization: `Bearer ${accessToken}` } }); if (!response.ok) throw new Error(`No se pudieron cargar las grabaciones de SharePoint (${response.status})`); const payload = await response.json() as { value?: Array<{ id: string; name: string; size?: number; lastModifiedDateTime?: string; webUrl?: string; file?: { mimeType?: string } }> }; return (payload.value || []).filter((item) => !!item.file).map((item) => ({ id: item.id, name: item.name, size: item.size || 0, modified: item.lastModifiedDateTime || "", webUrl: item.webUrl })); }
async function deleteSharePointFile(id: string) { const { driveId, accessToken } = await getSharePointDrive(); const response = await fetch(`https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(id)}`, { method: "DELETE", headers: { Authorization: `Bearer ${accessToken}` } }); if (!response.ok && response.status !== 404) throw new Error(`No se pudo eliminar el archivo de SharePoint (${response.status})`); }
async function downloadSharePointFile(id: string, fileName: string) { const { driveId, accessToken } = await getSharePointDrive(); const response = await fetch(`https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(id)}/content`, { headers: { Authorization: `Bearer ${accessToken}` } }); if (!response.ok) throw new Error(`No se pudo descargar el archivo (${response.status})`); const blob = await response.blob(); const url = URL.createObjectURL(blob); const anchor = document.createElement("a"); anchor.href = url; anchor.download = fileName; document.body.appendChild(anchor); anchor.click(); anchor.remove(); window.setTimeout(() => URL.revokeObjectURL(url), 1000); }

export default function Home() {
  const [location, setLocation] = useLocation();
  const [meetings, setMeetings] = useState<Meeting[]>([]);
  const [connected, setConnected] = useState(() => localStorage.getItem("m365-connected") === "true");
  const [connectOpen, setConnectOpen] = useState(false);
  const [connecting, setConnecting] = useState(false);
  const [authReady, setAuthReady] = useState(!microsoftApp);
  const [search, setSearch] = useState("");
  const [transcriptOpen, setTranscriptOpen] = useState(false);
  const [transcript, setTranscript] = useState("Mariana: Acordamos entregar el tablero de capacidad el viernes.\nDiego: Validaré el flujo con Legal antes del lunes.");
  const [testMeetingId, setTestMeetingId] = useState<string>("");
  const [personalRecording, setPersonalRecording] = useState(true);
  const [processAtEnd, setProcessAtEnd] = useState(true);
  const [reviewBeforeSend, setReviewBeforeSend] = useState(true);
  const [localRecording, setLocalRecording] = useState<MediaRecorder | null>(null);
  const [recordingMeetingId, setRecordingMeetingId] = useState<Meeting["id"] | null>(null);
  const [longRecording, setLongRecording] = useState(false);
  const [recordingStartedAt, setRecordingStartedAt] = useState<number | null>(null);
  const [sharePointItems, setSharePointItems] = useState<Array<{ id: string; note: string; modified: string }>>([]);
  const [sharePointLoading, setSharePointLoading] = useState(false);
  const [remoteFiles, setRemoteFiles] = useState<RemoteSharePointFile[]>([]);
  const [actaPreview, setActaPreview] = useState("");
  const [actaAttendees, setActaAttendees] = useState<Attendee[]>([]);
  const [pendingSharePoint, setPendingSharePoint] = useState<{ meetingTitle: string; graphMeetingId: string } | null>(null);
  const [actaStatus, setActaStatus] = useState<{ type: "idle" | "processing" | "generated" | "sharepoint" | "error"; text: string }>({ type: "idle", text: "Aún no se ha generado un acta en esta sesión." });
  useEffect(() => {
    const handleSignedOut = () => {
      setConnected(false);
      setMeetings([]);
      setActaAttendees([]);
      setPendingSharePoint(null);
      toast.success("Sesión de Microsoft 365 cerrada");
    };
    window.addEventListener("m365-signed-out", handleSignedOut);
    return () => window.removeEventListener("m365-signed-out", handleSignedOut);
  }, []);
  const loadRemoteFiles = async () => { setSharePointLoading(true); try { setRemoteFiles(await listSharePointFiles()); } catch (error) { toast.error(error instanceof Error ? error.message : "No se pudieron cargar los archivos de SharePoint"); } finally { setSharePointLoading(false); } };
  useEffect(() => { if (connected && microsoftApp) void loadRemoteFiles(); }, [connected]);
  const processLocalRecording = trpc.agent.processLocalRecording.useMutation({ onMutate: () => setActaStatus({ type: "processing", text: "Procesando grabación y generando acta…" }), onSuccess: (result) => { setTranscript(result.transcript); setActaAttendees(result.attendees?.length ? result.attendees : meetings.find((meeting) => String(meeting.id) === String(result.graphMeetingId))?.attendees || []); setActaPreview(sharePointNote(result.output, result.meetingTitle, meetings.find((meeting) => String(meeting.id) === String(result.graphMeetingId))?.attendees || [])); setPendingSharePoint(result.graphMeetingId ? { meetingTitle: result.meetingTitle, graphMeetingId: result.graphMeetingId } : null); setActaStatus({ type: "generated", text: "Acta generada. Revisa o edita el texto antes de enviarlo a SharePoint." }); void uploadGeneratedDocuments(result.documents).then(() => toast.success("Acta y compromisos respaldados en SharePoint")).catch((error) => toast.warning(`Acta generada, pero no se pudo respaldar en SharePoint: ${error instanceof Error ? error.message : "error de carga"}`)); toast.success("Grabación transcrita y acta generada"); }, onError: (error) => { setActaStatus({ type: "error", text: `Error al generar el acta: ${error.message}` }); toast.error(error.message); } });
  const processTextTranscript = trpc.agent.processTextTranscript.useMutation({ onSuccess: (result) => { void trpcUtils.agent.documents.invalidate(); void trpcUtils.agent.commitments.invalidate(); setTranscriptOpen(false); const selected = meetings.find((meeting) => String(meeting.id) === testMeetingId); if (selected && typeof selected.id === "string") { setActaAttendees(selected.attendees || []); setActaPreview(sharePointNote(result.output, selected.title, selected.attendees || [])); setPendingSharePoint({ meetingTitle: selected.title, graphMeetingId: selected.id }); } setActaStatus({ type: "generated", text: "Acta generada. Revisa o edita el texto antes de enviarlo a SharePoint." }); toast.success("Acta generada desde la reunión seleccionada"); }, onError: (error) => toast.error(error.message) });
  const recordingsQuery = trpc.agent.recordings.useQuery(undefined, { enabled: true });
  const documentsQuery = trpc.agent.documents.useQuery(undefined, { enabled: true });
  const commitmentsQuery = trpc.agent.commitments.useQuery(undefined, { enabled: true });
  const trpcUtils = trpc.useUtils();
  const clearRecording = trpc.agent.clearRecordingReference.useMutation({ onSuccess: () => { void trpcUtils.agent.recordings.invalidate(); toast.success("Referencia eliminada del historial"); }, onError: (error) => toast.error(error.message) });
  const deleteDocument = trpc.agent.deleteDocument.useMutation({ onSuccess: () => { void trpcUtils.agent.documents.invalidate(); toast.success("Acta eliminada"); }, onError: (error) => toast.error(error.message) });
  const reprocessRecording = trpc.agent.reprocessRecording.useMutation({ onMutate: () => setActaStatus({ type: "processing", text: "Reprocesando grabación y generando acta…" }), onSuccess: (result) => { setActaStatus({ type: "generated", text: "Acta generada nuevamente. Revisa o edita el texto antes de enviarlo a SharePoint." }); setActaAttendees(result.attendees?.length ? result.attendees : meetings.find((meeting) => String(meeting.id) === String(result.graphMeetingId))?.attendees || []); setActaPreview(sharePointNote(result.output, result.meetingTitle, meetings.find((meeting) => String(meeting.id) === String(result.graphMeetingId))?.attendees || [])); setPendingSharePoint(result.graphMeetingId ? { meetingTitle: result.meetingTitle, graphMeetingId: result.graphMeetingId } : null); void trpcUtils.agent.recordings.invalidate(); void trpcUtils.agent.documents.invalidate(); void trpcUtils.agent.commitments.invalidate(); toast.success("Grabación reprocesada y acta generada"); }, onError: (error) => toast.error(error.message) });

  const filtered = useMemo(() => meetings.filter((item) => `${item.title} ${item.organizerName}`.toLowerCase().includes(search.toLowerCase())), [meetings, search]);
  const importCalendar = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const text = String(reader.result || "");
      const blocks = text.split("BEGIN:VEVENT").slice(1);
      const imported: Meeting[] = blocks.map((block, index) => {
        const value = (name: string) => block.match(new RegExp(`(?:^|\\n)${name}[^:]*:(.*)`, "i"))?.[1]?.trim() || "";
        const rawDate = value("DTSTART").replace(/[^0-9TZ]/g, "");
        const iso = rawDate.length >= 8 ? `${rawDate.slice(0, 4)}-${rawDate.slice(4, 6)}-${rawDate.slice(6, 8)}T${rawDate.slice(9, 11) || "09"}:${rawDate.slice(11, 13) || "00"}:00Z` : new Date().toISOString();
        return { id: `ics-${index}`, title: value("SUMMARY") || "Reunión de calendario", organizerName: "Mi calendario Microsoft 365", scheduledAt: iso, durationMinutes: 60, attendeesCount: 0, attendees: [], recordingEnabled: false, processingEnabled: true, status: "Importada", source: "ICS" };
      });
      if (!imported.length) { toast.error("No encontré reuniones en el archivo .ics"); return; }
      setMeetings(imported);
      toast.success(`${imported.length} reuniones importadas`);
    };
    reader.readAsText(file);
  };
  const syncMicrosoftCalendar = async (account: AccountInfo) => {
    if (!microsoftApp) return;
    const token = await microsoftApp.acquireTokenSilent({ account, scopes: ["Calendars.Read"] });
      const start = new Date();
      const end = new Date(Date.now() + 14 * 24 * 60 * 60 * 1000);
      const params = new URLSearchParams({ startDateTime: start.toISOString(), endDateTime: end.toISOString(), "$top": "50", "$orderby": "start/dateTime" });
      const response = await fetch(`https://graph.microsoft.com/v1.0/me/calendarView?${params}`, { headers: { Authorization: `Bearer ${token.accessToken}`, Prefer: 'IdType="ImmutableId"' } });
      if (!response.ok) throw new Error(`Graph respondió ${response.status}`);
      const payload = await response.json() as { value: Array<{ id: string; subject?: string; start?: { dateTime?: string }; end?: { dateTime?: string }; organizer?: { emailAddress?: { name?: string } }; attendees?: Array<{ emailAddress?: { name?: string; address?: string } }>; isOnlineMeeting?: boolean }> };
      const imported: Meeting[] = payload.value.map((event) => ({ id: event.id, title: event.subject || "Reunión de Outlook", organizerName: event.organizer?.emailAddress?.name || "Mi calendario Microsoft 365", scheduledAt: event.start?.dateTime ? new Date(`${event.start.dateTime}Z`).toISOString() : new Date().toISOString(), durationMinutes: event.start?.dateTime && event.end?.dateTime ? Math.max(15, Math.round((new Date(`${event.end.dateTime}Z`).getTime() - new Date(`${event.start.dateTime}Z`).getTime()) / 60000)) : 60, attendeesCount: event.attendees?.length || 0, attendees: (event.attendees || []).map((attendee) => { const item = attendee as { emailAddress?: { name?: string; address?: string } }; return { name: item.emailAddress?.name || item.emailAddress?.address || "Participante", email: item.emailAddress?.address || "", role: "Participante" }; }), recordingEnabled: false, processingEnabled: true, status: event.isOnlineMeeting ? "Teams" : "Calendario", source: "Microsoft Graph" }));
      setMeetings(imported);
      setConnected(true);
      const profile = { name: account.name || account.username.split("@")[0] || "Usuario Microsoft 365", email: account.username };
      localStorage.setItem("m365-profile", JSON.stringify(profile));
      window.dispatchEvent(new CustomEvent("m365-profile-updated"));
      localStorage.setItem("m365-connected", "true");
      setConnectOpen(false);
      toast.success(`${imported.length} reuniones sincronizadas desde Microsoft 365`);
  };
  const loadSharePointItems = async (meetingId?: string) => {
    if (!microsoftApp) { toast.info("Primero vincula Microsoft 365 para consultar SharePoint."); return; }
    const account = microsoftApp.getAllAccounts()[0];
    if (!account) { toast.info("Primero vincula Microsoft 365 para consultar SharePoint."); return; }
    setSharePointLoading(true);
    try {
      const token = await microsoftApp.acquireTokenSilent({ account, scopes: ["Sites.ReadWrite.All"] });
      const siteUrl = new URL(import.meta.env.VITE_SHAREPOINT_SITE_URL || "https://abcstorage.sharepoint.com/transformaciondigital");
      const relativePath = siteUrl.pathname.replace(/\/$/, "") || "/";
      const headers = { Authorization: `Bearer ${token.accessToken}` };
      const site = await resolveSharePointSite(token.accessToken, siteUrl);
      const listName = import.meta.env.VITE_SHAREPOINT_LIST_NAME || "Reuniones Efectivas";
      const listResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(site.id || "")}/lists/${encodeURIComponent(listName)}?$select=id,displayName`, { headers });
      if (!listResponse.ok) throw new Error(`No se encontró la lista «${listName}» (${listResponse.status})`);
      const list = await listResponse.json() as { id?: string };
      const itemsResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(site.id || "")}/lists/${encodeURIComponent(list.id || "")}/items?$expand=fields&$top=999`, { headers });
      if (!itemsResponse.ok) throw new Error(`No se pudieron leer los registros (${itemsResponse.status})`);
      const payload = await itemsResponse.json() as { value?: Array<{ id: string; fields?: Record<string, unknown> }> };
      const normalize = (value: unknown) => String(value ?? "").replace(/[\u200B-\u200D\uFEFF]/g, "").replace(/[\r\n\t ]+/g, "").trim().toLowerCase();
      const selectedId = normalize(meetingId);
      const visible = selectedId ? (payload.value || []).filter((item) => Object.values(item.fields || {}).some((value) => normalize(value) === selectedId)) : (payload.value || []).slice(0, 25);
      setSharePointItems(visible.map((item) => ({ id: item.id, note: String(item.fields?.Nota || "(sin Nota)"), modified: String(item.fields?.Modified || "") })));
      toast.success(selectedId ? `${visible.length} registro(s) encontrados para el ID seleccionado` : "Registros de SharePoint cargados");
    } catch (error) { toast.error(error instanceof Error ? error.message : "No se pudo cargar SharePoint"); } finally { setSharePointLoading(false); }
  };
  const syncNoteToSharePoint = async (output: { executiveSummary: string; objective: string; decisions: string[]; commitments: Array<{ personName: string; action: string; dueDate: string }> }, meetingTitle: string, graphMeetingId?: string, editedNote?: string) => {
    if (!microsoftApp) return;
    const account = microsoftApp.getAllAccounts()[0];
    if (!account) return;
    const realMeetingId = graphMeetingId || meetings.find((item) => item.title.trim().toLowerCase() === meetingTitle.trim().toLowerCase())?.id;
    if (!realMeetingId || typeof realMeetingId !== "string") { toast.warning("Acta local lista, pero no está vinculada a una reunión sincronizada de Teams. No se modificó SharePoint."); return; }
    try {
      const token = await microsoftApp.acquireTokenSilent({ account, scopes: ["Sites.ReadWrite.All"] });
      const headers = { Authorization: `Bearer ${token.accessToken}` };
      const siteUrl = new URL(import.meta.env.VITE_SHAREPOINT_SITE_URL || "https://abcstorage.sharepoint.com/transformaciondigital");
      const relativePath = siteUrl.pathname.replace(/\/$/, "") || "/";
      const site = await resolveSharePointSite(token.accessToken, siteUrl);
      if (!site.id) throw new Error("Graph no devolvió el sitio de SharePoint");
      const listName = import.meta.env.VITE_SHAREPOINT_LIST_NAME || "Reuniones Efectivas";
      const listResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(listName)}?$select=id,displayName`, { headers });
      if (!listResponse.ok) throw new Error(`No se encontró la lista «${listName}» (${listResponse.status})`);
      const list = await listResponse.json() as { id?: string };
      if (!list.id) throw new Error("Graph no devolvió la lista de SharePoint");
      const columnsResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(list.id)}/columns?$select=name,displayName`, { headers });
      if (!columnsResponse.ok) throw new Error(`No se pudieron leer las columnas de SharePoint (${columnsResponse.status})`);
      const columns = await columnsResponse.json() as { value?: Array<{ name: string; displayName?: string }> };
      const configuredIdColumn = import.meta.env.VITE_SHAREPOINT_MEETING_ID_COLUMN;
      const detectedIdColumn = columns.value?.find((column) => /id\s*reuni[oó]n/i.test(column.displayName || ""))?.name;
      const idColumns = [configuredIdColumn, detectedIdColumn, "Id_x0020_reuni_x00f3_n", "Id reunión", "Id reunion"].filter((value, index, list): value is string => Boolean(value) && list.indexOf(value) === index);
      const items: Array<{ id: string; fields?: Record<string, unknown> }> = [];
      let nextUrl: string | null = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(list.id)}/items?$expand=fields&$top=999`;
      while (nextUrl && items.length < 10000) {
        const itemResponse = await fetch(nextUrl, { headers });
        if (!itemResponse.ok) throw new Error(`No se pudieron buscar las reuniones (${itemResponse.status})`);
        const payload = await itemResponse.json() as { value?: Array<{ id: string; fields?: Record<string, unknown> }>; [key: string]: unknown };
        items.push(...(payload.value || []));
        nextUrl = typeof payload["@odata.nextLink"] === "string" ? payload["@odata.nextLink"] : null;
      }
      const normalizeSharePointValue = (value: unknown) => String(value ?? "").replace(/[\u200B-\u200D\uFEFF]/g, "").replace(/[\r\n\t ]+/g, "").trim().toLowerCase();
      const normalizedMeetingId = normalizeSharePointValue(realMeetingId);
      const normalizedTitle = meetingTitle.trim().replace(/[\u200B-\u200D\uFEFF]/g, "").replace(/[\r\n\t ]+/g, " ").toLowerCase();
      const idMatches = items.filter((item) => idColumns.some((column) => normalizeSharePointValue(item.fields?.[column]) === normalizedMeetingId) || Object.values(item.fields || {}).some((value) => normalizeSharePointValue(value) === normalizedMeetingId));
      const titleMatches = items.filter((item) => String(item.fields?.Title ?? item.fields?.title ?? "").replace(/[\u200B-\u200D\uFEFF]/g, "").replace(/[\r\n\t ]+/g, " ").trim().toLowerCase() === normalizedTitle);
      const matches = idMatches.length ? idMatches : titleMatches;
      const matchMode = idMatches.length ? `Id reunión (${realMeetingId})` : `título exacto (${meetingTitle})`;
      if (matches.length === 0) throw new Error(`No encontré una fila por Id reunión ni por título exacto. Id Teams = ${realMeetingId}; título = ${meetingTitle}; filas leídas = ${items.length}; sitio = ${site.id}; lista = ${list.id}. Verifica que VITE_SHAREPOINT_SITE_URL apunte al mismo sitio que estás viendo. No se creó ninguna fila nueva.`);
      if (matches.length > 1) throw new Error(`Encontré ${matches.length} filas con ${matchMode}. No se modificó SharePoint para evitar actualizar la reunión equivocada.`);
      const note = editedNote?.trim() || sharePointNote(output, meetingTitle);
      const patchResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(list.id)}/items/${encodeURIComponent(matches[0].id)}/fields`, { method: "PATCH", headers: { ...headers, "Content-Type": "application/json" }, body: JSON.stringify({ [import.meta.env.VITE_SHAREPOINT_COLUMN_NAME || "Nota"]: note }) });
      if (!patchResponse.ok) throw new Error(`No se pudo actualizar Nota (${patchResponse.status})`);
      setTestMeetingId(realMeetingId);
      setSharePointItems([{ id: matches[0].id, note, modified: new Date().toISOString() }]);
      setPendingSharePoint(null);
      setActaStatus({ type: "sharepoint", text: `Acta guardada en SharePoint · fila ${matches[0].id} · coincidencia por ${matchMode}` });
      toast.success(`Acta guardada en la reunión correcta de SharePoint (fila ${matches[0].id})`);
    } catch (error) {
      toast.warning(`Acta local lista; SharePoint no se actualizó: ${error instanceof Error ? error.message : "revisa Sites.ReadWrite.All"}`);
    }
  };
  useEffect(() => {
    if (!microsoftApp) return;
    void (async () => {
      try {
        await microsoftApp.initialize();
        const result = await microsoftApp.handleRedirectPromise();
        const account = result?.account || microsoftApp.getAllAccounts()[0];
        if (account) await syncMicrosoftCalendar(account);
      } catch (error) {
        const code = error && typeof error === "object" && "errorCode" in error ? String((error as { errorCode?: string }).errorCode) : "";
        if (code === "interaction_in_progress") { sessionStorage.removeItem("msal.interaction.status"); toast.error("Microsoft conservaba un inicio anterior. Ya limpié el estado local; recarga el panel y vuelve a intentarlo una sola vez."); }
        else toast.error(error instanceof Error ? error.message : "No se pudo completar la conexión Microsoft 365");
      } finally { setAuthReady(true); }
    })();
  }, []);
  const connectMicrosoft = async () => {
    if (!microsoftApp) { toast.info("Falta VITE_MICROSOFT_CLIENT_ID. Para probar ahora usa Importar .ics."); return; }
    if (connecting || !authReady) { toast.info("Espera a que termine de cargar Microsoft 365."); return; }
    setConnecting(true);
    try {
      await microsoftApp.initialize();
      const existingAccount = microsoftApp.getAllAccounts()[0];
      if (existingAccount) { await syncMicrosoftCalendar(existingAccount); setConnecting(false); return; }
      if (sessionStorage.getItem("msal.interaction.status")) sessionStorage.removeItem("msal.interaction.status");
      await microsoftApp.loginRedirect({ scopes: microsoftScopes });
    } catch (error) {
      const code = error && typeof error === "object" && "errorCode" in error ? String((error as { errorCode?: string }).errorCode) : "";
      if (code === "interaction_in_progress") { sessionStorage.removeItem("msal.interaction.status"); toast.error("Se limpió una sesión anterior de Microsoft. Recarga la página e inténtalo una sola vez."); }
      else toast.error(error instanceof Error ? error.message : "No se pudo iniciar la conexión Microsoft 365");
      setConnecting(false);
    }
  };
  const toggle = (id: Meeting["id"], field: "recordingEnabled" | "processingEnabled") => setMeetings((items) => items.map((item) => item.id === id ? { ...item, [field]: !item[field] } : item));
  const recordLocalAudio = async (meetingTitle = "Reunión grabada localmente", meetingId: Meeting["id"] | null = null) => {
    if (localRecording) { localRecording.stop(); return; }
    try {
      const screen = await navigator.mediaDevices.getDisplayMedia({ video: true, audio: true });
      const microphone = await navigator.mediaDevices.getUserMedia({ audio: true });
      const audioContext = new AudioContext(); const destination = audioContext.createMediaStreamDestination();
      if (screen?.getAudioTracks().length) audioContext.createMediaStreamSource(new MediaStream(screen.getAudioTracks())).connect(destination);
      audioContext.createMediaStreamSource(microphone).connect(destination);
      const recordingStream = longRecording ? destination.stream : new MediaStream([...screen!.getVideoTracks(), ...destination.stream.getAudioTracks()]);
      const mimeType = longRecording ? "audio/webm" : "video/webm";
      const audioMimeType = "audio/webm";
      const startedAt = Date.now(); const recorder = new MediaRecorder(recordingStream, { mimeType, audioBitsPerSecond: longRecording ? 16000 : 64000, videoBitsPerSecond: 120000 }); const chunks: Blob[] = [];
      const audioChunks: Blob[] = []; const audioRecorder = longRecording ? recorder : new MediaRecorder(destination.stream, { mimeType: audioMimeType, audioBitsPerSecond: 16000 });
      recorder.ondataavailable = (event) => { if (event.data.size) chunks.push(event.data); };
      if (!longRecording) audioRecorder.ondataavailable = (event) => { if (event.data.size) audioChunks.push(event.data); };
      const audioFinished = longRecording ? Promise.resolve() : new Promise<void>((resolve) => { audioRecorder.addEventListener("stop", () => resolve(), { once: true }); });
      recorder.onstop = async () => { if (!longRecording && audioRecorder.state !== "inactive") audioRecorder.stop(); await audioFinished; screen?.getTracks().forEach((track) => track.stop()); microphone.getTracks().forEach((track) => track.stop()); await audioContext.close(); const blob = new Blob(chunks, { type: mimeType }); const audioBlob = longRecording ? blob : new Blob(audioChunks, { type: audioMimeType }); const durationSeconds = Math.max(1, Math.round((Date.now() - startedAt) / 1000)); if (blob.size < 100 || audioBlob.size < 100) { toast.error("No se capturó audio. Mantén la reunión activa unos segundos y verifica que Teams tenga el audio compartido."); setLocalRecording(null); return; } if (blob.size > 180 * 1024 * 1024) { toast.error("La grabación supera 180 MB. Reduce la resolución de pantalla o divide la reunión en dos partes."); setLocalRecording(null); return; } const toDataUrl = (value: Blob) => new Promise<string>((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(String(reader.result)); reader.onerror = () => reject(reader.error || new Error("No se pudo leer la grabación")); reader.readAsDataURL(value); }); try { const [audioBase64, recordingBase64] = await Promise.all([toDataUrl(audioBlob), longRecording ? Promise.resolve(undefined) : toDataUrl(blob)]); try { const extension = longRecording ? "webm" : "webm"; await uploadBlobToSharePoint(blob, `${meetingTitle}-${new Date().toISOString().replace(/[:.]/g, "-")}.${extension}`, "Grabaciones"); toast.success("Grabación respaldada en SharePoint"); } catch (error) { toast.warning(`La grabación se procesará, pero no se pudo respaldar en SharePoint: ${error instanceof Error ? error.message : "error de carga"}`); } processLocalRecording.mutate({ audioBase64, recordingBase64, mimeType: audioMimeType, recordingMimeType: longRecording ? audioMimeType : mimeType, title: meetingTitle, meetingId: typeof meetingId === "number" ? meetingId : undefined, graphMeetingId: typeof meetingId === "string" ? meetingId : undefined, durationSeconds, attendees: (meetings.find((meeting) => String(meeting.id) === String(meetingId))?.attendees || []).map((person) => ({ name: person.name, email: person.email, role: person.role })) }); } catch (error) { toast.error(error instanceof Error ? error.message : "No se pudo preparar la grabación"); } setLocalRecording(null); setRecordingStartedAt(null); setRecordingMeetingId(null); };
      recorder.start(1000); if (!longRecording) audioRecorder.start(1000); setLocalRecording(recorder); setRecordingStartedAt(Date.now()); setRecordingMeetingId(meetingId); toast.success(`${longRecording ? "Audio largo de bajo tamaño (más de 3 horas)" : "Pantalla y audio"} activo para «${meetingTitle}». Informa y obtén consentimiento.`);
    } catch (error) { toast.error(error instanceof Error ? error.message : "No se pudo iniciar la grabación"); }
  };

  const processUploadedRecording = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    if (file.size > 180 * 1024 * 1024) { toast.error("El archivo supera 180 MB. Divide la grabación antes de cargarla."); return; }
    try {
      const dataUrl = await new Promise<string>((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(String(reader.result)); reader.onerror = () => reject(reader.error || new Error("No se pudo leer el archivo")); reader.readAsDataURL(file); });
      const mimeType = file.type || (file.name.toLowerCase().endsWith(".mp3") ? "audio/mpeg" : "audio/webm");
      processLocalRecording.mutate({ audioBase64: dataUrl, recordingBase64: dataUrl, mimeType, recordingMimeType: mimeType, title: file.name.replace(/\.[^.]+$/, "") || "Grabación cargada", durationSeconds: 0, attendees: [] });
    } catch (error) { toast.error(error instanceof Error ? error.message : "No se pudo cargar la grabación"); }
  };
  const processedCount = documentsQuery.data?.filter((item) => item.kind === "minutes").length ?? 0;
  const pendingReviewCount = recordingsQuery.data?.filter((item) => item.meeting.status === "review").length ?? 0;
  const openCommitmentsCount = commitmentsQuery.data?.filter((item) => item.status !== "done").length ?? 0;
  const pageTitle = location === "/reuniones" ? "Mis reuniones" : location === "/actas" ? "Actas generadas" : location === "/grabaciones" ? "Grabaciones" : location === "/compromisos" ? "Mis compromisos" : location === "/configuracion" ? "Configuración personal" : "Reuniones bajo control.";
  return <DashboardLayout>
    <div className="mx-auto max-w-[1450px] px-5 py-7 lg:px-9 lg:py-9">
      <header className="flex flex-col justify-between gap-5 xl:flex-row xl:items-end"><div><div className="mb-4 flex items-center gap-2 text-xs font-semibold uppercase tracking-[0.2em] text-[#0d776c]"><span className="h-2 w-2 rounded-full bg-[#0d9a85]" />Agente personal · Microsoft 365</div><h1 className="text-4xl font-semibold tracking-[-0.04em] text-[#102c36] sm:text-5xl">{pageTitle}</h1><p className="mt-3 max-w-2xl text-base leading-7 text-[#66818a]">El agente trabaja solo con tus reuniones, genera el acta y deja los compromisos listos para revisión.</p></div><div className="flex flex-wrap gap-3"><Badge className={cn("rounded-full border px-3 py-1.5", connected ? "border-[#b9e5da] bg-[#e8f7f1] text-[#24765f]" : "border-[#f0d9a4] bg-[#fff6df] text-[#9a731e]")}><Cloud className="mr-1.5 h-3.5 w-3.5" />{connected ? "Microsoft 365 vinculado" : "Microsoft 365 no vinculado"}</Badge><Button onClick={() => setConnectOpen(true)} variant="outline" className="rounded-xl border-[#d7e4e5] bg-white"><Link2 className="mr-2 h-4 w-4 text-[#0d776c]" />Vincular Microsoft 365</Button></div></header>

      {location === "/configuracion" ? <Configuration connected={connected} connect={connectMicrosoft} personalRecording={personalRecording} setPersonalRecording={setPersonalRecording} processAtEnd={processAtEnd} setProcessAtEnd={setProcessAtEnd} reviewBeforeSend={reviewBeforeSend} setReviewBeforeSend={setReviewBeforeSend} /> : location === "/actas" ? <Documents data={documentsQuery.data} recordings={recordingsQuery.data} sharePointItems={sharePointItems} sharePointLoading={sharePointLoading} preview={actaPreview} actaStatus={actaStatus} attendees={actaAttendees} onLoadSharePoint={() => void loadSharePointItems(testMeetingId || undefined)} onEditPreview={setActaPreview} onSendSharePoint={() => { if (pendingSharePoint) void syncNoteToSharePoint({ executiveSummary: "", objective: "", decisions: [], commitments: [] }, pendingSharePoint.meetingTitle, pendingSharePoint.graphMeetingId, actaPreview); else toast.info("Primero genera un acta vinculada a Microsoft 365."); }} onGenerateFromRecording={(id, url) => { const meeting = meetings.find((item) => Number(item.id) === id); reprocessRecording.mutate({ meetingId: id, recordingUrl: url, attendees: meeting?.attendees || [] }); }} onTest={() => { const first = meetings.find((meeting) => typeof meeting.id === "string"); setTestMeetingId(first ? String(first.id) : ""); setTranscriptOpen(true); }} onDelete={(id) => deleteDocument.mutate({ id })} /> : location === "/grabaciones" ? <Recordings data={recordingsQuery.data} onClear={(meetingId) => clearRecording.mutate({ meetingId })} onReprocess={(meetingId, recordingUrl) => { const meeting = meetings.find((item) => Number(item.id) === meetingId); reprocessRecording.mutate({ meetingId, recordingUrl, attendees: meeting?.attendees || [] }); }} onUpload={processUploadedRecording} remoteFiles={remoteFiles} onRefreshRemote={loadRemoteFiles} onDeleteRemote={async (id) => { if (!window.confirm("¿Eliminar este archivo de SharePoint?")) return; try { await deleteSharePointFile(id); await loadRemoteFiles(); toast.success("Archivo eliminado de SharePoint"); } catch (error) { toast.error(error instanceof Error ? error.message : "No se pudo eliminar el archivo"); } }} /> : location === "/compromisos" ? <Commitments data={commitmentsQuery.data} onUpdate={() => void trpcUtils.agent.commitments.invalidate()} /> : <>{location === "/" && <ControlCenter meetings={meetings} onOpenMeetings={() => setLocation("/reuniones")} />}</>}
        {location === "/reuniones" && <><section className="mt-8 grid gap-4 sm:grid-cols-2 xl:grid-cols-4">{([{ label: "Mis reuniones esta semana", value: meetings.length, Icon: CalendarClock, color: "#e6f6f1" }, { label: "Procesadas", value: processedCount, Icon: Sparkles, color: "#f0edff" }, { label: "Pendientes de revisión", value: pendingReviewCount, Icon: FileCheck2, color: "#fff4d8" }, { label: "Compromisos abiertos", value: openCommitmentsCount, Icon: ListChecks, color: "#fff0ed" }] as const).map(({ label, value, Icon, color }) => <Card key={label} className="rounded-2xl border-0 bg-white shadow-sm"><CardContent className="flex items-start justify-between p-5"><div><p className="text-sm text-[#6b858c]">{label}</p><p className="mt-3 text-3xl font-semibold text-[#102c36]">{value}</p></div><div style={{ backgroundColor: color }} className="flex h-10 w-10 items-center justify-center rounded-xl"><Icon className="h-5 w-5 text-[#0d776c]" /></div></CardContent></Card>)}</section>
        <Card className="mt-7 overflow-hidden rounded-2xl border-0 bg-white shadow-sm"><CardHeader className="border-b px-5 py-5 sm:px-6"><div className="flex flex-col justify-between gap-4 sm:flex-row sm:items-center"><div><CardTitle className="text-lg text-[#102c36]">{connected ? "Reuniones de mi calendario" : "Reuniones disponibles para probar"}</CardTitle><p className="mt-1 text-sm text-[#789098]">Activa procesamiento y graba audio localmente con consentimiento.</p></div><div className="flex flex-wrap gap-2"><Button disabled={processLocalRecording.isPending} onClick={() => { void recordLocalAudio(); }} className={cn("rounded-xl text-xs text-white", localRecording ? "bg-[#b24e45] hover:bg-[#963d36]" : "bg-[#0d776c] hover:bg-[#095f57]")}>{localRecording ? <><Mic2 className="mr-2 h-3.5 w-3.5" />Detener grabación</> : <><Mic2 className="mr-2 h-3.5 w-3.5" />Grabar localmente</>}</Button><Button type="button" variant="outline" onClick={() => setLongRecording((value) => !value)} disabled={!!localRecording} className={cn("rounded-xl text-xs", longRecording ? "border-[#0d776c] bg-[#e8f7f1] text-[#0d776c]" : "border-[#dce8e8] text-[#52717a]")}>{longRecording ? "Modo largo · solo audio" : "Modo pantalla + audio"}</Button><label className="inline-flex cursor-pointer items-center gap-2 rounded-xl border border-[#dce8e8] px-3 py-2 text-xs font-semibold text-[#52717a] hover:bg-[#f1f8f7]"><Upload className="h-3.5 w-3.5" />Importar .ics<input type="file" accept=".ics,text/calendar" className="hidden" onChange={importCalendar} /></label><div className="relative"><Search className="absolute left-3 top-2.5 h-4 w-4 text-[#9ab0b4]" /><Input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Buscar" className="h-9 w-[160px] rounded-xl pl-9 text-xs" /></div></div></div></CardHeader><CardContent className="p-0">{filtered.map((meeting) => <div key={meeting.id} className="flex flex-col gap-4 border-b px-5 py-5 sm:flex-row sm:items-center sm:px-6"><div className="flex min-w-0 flex-1 items-start gap-4"><div className="flex h-11 w-11 shrink-0 flex-col items-center justify-center rounded-xl bg-[#edf7f5] text-[#0d776c]"><span className="text-[10px] font-semibold uppercase">{dateLabel(meeting.scheduledAt).split(" ")[0]}</span><span className="text-sm font-semibold">{new Date(meeting.scheduledAt).getDate()}</span></div><div className="min-w-0"><div className="flex flex-wrap items-center gap-2"><h3 className="truncate text-sm font-semibold text-[#183841]">{meeting.title}</h3><Badge className="rounded-full bg-[#eff5f5] px-2 py-0.5 text-[10px] text-[#52717a]">{meeting.status}</Badge></div><p className="mt-1.5 flex flex-wrap gap-3 text-xs text-[#789098]"><span>{timeLabel(meeting.scheduledAt)} · {meeting.durationMinutes} min</span><span className="flex items-center gap-1"><UsersRound className="h-3.5 w-3.5" />{meeting.attendeesCount || "—"} asistentes</span><span>{meeting.organizerName}</span></p></div></div><div className="flex items-center gap-5"><div className="flex items-center gap-2"><Mic2 className={cn("h-4 w-4", meeting.recordingEnabled ? "text-[#c05a4d]" : "text-[#b6c7c9]")} /><div><p className="text-[10px] font-semibold uppercase tracking-[0.12em] text-[#8aa0a5]">Grabar</p><Switch checked={meeting.recordingEnabled} onCheckedChange={(enabled) => { toggle(meeting.id, "recordingEnabled"); if (enabled && !localRecording) void recordLocalAudio(meeting.title, meeting.id); if (!enabled && localRecording && recordingMeetingId === meeting.id) localRecording.stop(); }} className="mt-1 data-[state=checked]:bg-[#c05a4d]" /></div></div><Separator orientation="vertical" className="h-9" /><div className="flex items-center gap-2"><Sparkles className="h-4 w-4 text-[#6f5ab5]" /><div><p className="text-[10px] font-semibold uppercase tracking-[0.12em] text-[#8aa0a5]">Procesar</p><Switch checked={meeting.processingEnabled} onCheckedChange={() => { toggle(meeting.id, "processingEnabled"); toast.success("Procesamiento actualizado"); }} className="mt-1 data-[state=checked]:bg-[#6f5ab5]" /></div></div></div></div>)}<div className="flex flex-wrap items-center justify-between gap-3 px-6 py-4"><p className="text-xs text-[#83999e]">{connected ? "Sincronización conectada a tu cuenta Microsoft 365." : "Para datos reales, vincula Microsoft 365 o importa un archivo .ics."}</p><Button variant="ghost" onClick={() => setConnectOpen(true)} className="text-xs font-semibold text-[#0d776c]">{connected ? "Actualizar calendario" : "Vincular calendario"} <RefreshCw className="ml-1.5 h-3.5 w-3.5" /></Button></div></CardContent></Card></>}
      <div className="mt-8 flex items-center justify-between border-t border-[#dce8e8] pt-5 text-xs text-[#80979c]"><span><ShieldCheck className="mr-1 inline h-3.5 w-3.5" />Solo tú tienes acceso</span><span>Teams · Outlook · Power Apps</span></div>
    </div>
    <Dialog open={connectOpen} onOpenChange={setConnectOpen}><DialogContent className="max-w-lg rounded-2xl"><DialogHeader><DialogTitle>Vincular tu Microsoft 365</DialogTitle><DialogDescription>El vínculo usará permisos delegados para leer tu calendario y escribir el acta en tu lista de SharePoint. No compartas tu contraseña con el agente.</DialogDescription></DialogHeader><div className="space-y-4"><div className="rounded-xl bg-[#eff8f6] p-4 text-sm leading-6 text-[#52717a]"><ShieldCheck className="mr-2 inline h-4 w-4 text-[#0d776c]" />Se solicitarán `Calendars.Read` y `Sites.ReadWrite.All`. Al terminar cada procesamiento, el agente escribirá solo el campo `Nota` de `Reuniones Efectivas`.</div><Button disabled={connecting || !authReady} onClick={() => { void connectMicrosoft(); }} className="w-full rounded-xl bg-[#0d776c] text-white hover:bg-[#095f57]">{connecting || !authReady ? <RefreshCw className="mr-2 h-4 w-4 animate-spin" /> : <Link2 className="mr-2 h-4 w-4" />}{!authReady ? "Preparando Microsoft…" : connecting ? "Redirigiendo en esta pestaña…" : "Iniciar conexión Microsoft"}</Button><div className="text-center text-xs text-[#83999e]">No se abrirá un popup. Microsoft regresará automáticamente a la dirección desde la que abriste la aplicación.</div></div><DialogFooter><Button variant="outline" onClick={() => setConnectOpen(false)} className="rounded-xl">Cerrar</Button></DialogFooter></DialogContent></Dialog>
    <Dialog open={transcriptOpen} onOpenChange={setTranscriptOpen}><DialogContent className="max-w-2xl rounded-2xl"><DialogHeader><DialogTitle>Generar acta de reunión programada</DialogTitle><DialogDescription>Selecciona la reunión sincronizada. El texto pegado se procesará con su ID real de Teams y se actualizará la fila correspondiente en SharePoint.</DialogDescription></DialogHeader><select value={testMeetingId} onChange={(event) => setTestMeetingId(event.target.value)} className="h-10 rounded-xl border border-[#dce8e8] bg-white px-3 text-sm text-[#183841]"><option value="">Selecciona una reunión de Microsoft 365</option>{meetings.filter((meeting) => typeof meeting.id === "string").map((meeting) => <option key={meeting.id} value={String(meeting.id)}>{meeting.title} · {dateLabel(meeting.scheduledAt)}</option>)}</select><Textarea value={transcript} onChange={(event) => setTranscript(event.target.value)} rows={8} /><DialogFooter><Button variant="outline" onClick={() => setTranscriptOpen(false)}>Cancelar</Button><Button disabled={processTextTranscript.isPending || !testMeetingId} onClick={() => { const selected = meetings.find((meeting) => String(meeting.id) === testMeetingId); if (selected) processTextTranscript.mutate({ title: selected.title, transcript }); }} className="bg-[#0d776c] text-white">{processTextTranscript.isPending ? "Generando…" : "Generar y actualizar SharePoint"}</Button></DialogFooter></DialogContent></Dialog>
  </DashboardLayout>;
}

function ControlCenter({ meetings, onOpenMeetings }: { meetings: Meeting[]; onOpenMeetings: () => void }) {
  const nextMeeting = [...meetings].sort((a, b) => new Date(a.scheduledAt).getTime() - new Date(b.scheduledAt).getTime())[0];
  const activeCount = meetings.filter((meeting) => meeting.recordingEnabled || meeting.processingEnabled).length;
  return <div className="mt-8 space-y-6">
    <section className="relative overflow-hidden rounded-3xl bg-[#102c36] p-7 text-white shadow-[0_20px_55px_rgba(16,44,54,0.18)] sm:p-9">
      <div className="relative z-10 max-w-2xl"><Badge className="border-0 bg-[#1b5960] text-[#b4e8dd]">Resumen operativo</Badge><h2 className="mt-4 text-3xl font-semibold tracking-tight sm:text-4xl">Todo listo para tu próxima reunión.</h2><p className="mt-3 max-w-xl text-sm leading-6 text-[#b8d0d0]">Desde aquí puedes ver el estado general del agente. Entra a Reuniones para seleccionar una reunión específica, grabar y configurar su procesamiento.</p><Button onClick={onOpenMeetings} className="mt-6 rounded-xl bg-[#8de0cd] text-[#10333b] hover:bg-[#b1f0e2]">Ver mis reuniones <CalendarClock className="ml-2 h-4 w-4" /></Button></div><div className="absolute -right-16 -top-24 h-72 w-72 rounded-full bg-[#0d9a85]/25 blur-2xl" /><div className="absolute -bottom-24 right-24 h-48 w-48 rounded-full bg-[#8975d6]/20 blur-2xl" />
    </section>
    <section className="grid gap-4 sm:grid-cols-3">
      {[{ label: "Reuniones cargadas", value: meetings.length, note: "en tu agenda" }, { label: "Con automatización", value: activeCount, note: "listas para actuar" }, { label: "Próxima reunión", value: nextMeeting ? timeLabel(nextMeeting.scheduledAt) : "—", note: nextMeeting?.title || "Sin reuniones" }].map((item) => <Card key={item.label} className="rounded-2xl border-0 bg-white shadow-[0_12px_32px_rgba(30,74,74,0.07)]"><CardContent className="p-5"><p className="text-xs font-semibold uppercase tracking-[0.12em] text-[#8aa0a5]">{item.label}</p><p className="mt-3 truncate text-2xl font-semibold text-[#102c36]">{item.value}</p><p className="mt-1 truncate text-xs text-[#789098]">{item.note}</p></CardContent></Card>)}
    </section>
    <Card className="rounded-2xl border-0 bg-white shadow-sm"><CardHeader><CardTitle className="text-lg text-[#183841]">Flujo recomendado</CardTitle><p className="text-sm text-[#789098]">Una vista rápida del proceso de trabajo.</p></CardHeader><CardContent className="grid gap-3 pb-6 md:grid-cols-3"><div className="rounded-xl bg-[#e8f7f1] p-4"><CalendarClock className="h-5 w-5 text-[#0d776c]" /><p className="mt-3 text-sm font-semibold text-[#183841]">1. Elige la reunión</p><p className="mt-1 text-xs leading-5 text-[#66818a]">Revisa fecha, asistentes y estado desde Reuniones.</p></div><div className="rounded-xl bg-[#f0edff] p-4"><Mic2 className="h-5 w-5 text-[#6f5ab5]" /><p className="mt-3 text-sm font-semibold text-[#183841]">2. Graba con consentimiento</p><p className="mt-1 text-xs leading-5 text-[#66818a]">Usa pantalla + audio o modo largo solo audio.</p></div><div className="rounded-xl bg-[#fff4d8] p-4"><FileCheck2 className="h-5 w-5 text-[#a27720]" /><p className="mt-3 text-sm font-semibold text-[#183841]">3. Revisa el resultado</p><p className="mt-1 text-xs leading-5 text-[#66818a]">Consulta actas y compromisos antes de distribuirlos.</p></div></CardContent></Card>
  </div>;
}

function Configuration({ connected, connect, personalRecording, setPersonalRecording, processAtEnd, setProcessAtEnd, reviewBeforeSend, setReviewBeforeSend }: { connected: boolean; connect: () => void; personalRecording: boolean; setPersonalRecording: (v: boolean) => void; processAtEnd: boolean; setProcessAtEnd: (v: boolean) => void; reviewBeforeSend: boolean; setReviewBeforeSend: (v: boolean) => void }) {
  const controls: Array<{ label: string; checked: boolean; setter: (value: boolean) => void; Icon: typeof Mic2 }> = [{ label: "Grabar por defecto", checked: personalRecording, setter: setPersonalRecording, Icon: Mic2 }, { label: "Procesar al finalizar", checked: processAtEnd, setter: setProcessAtEnd, Icon: WandSparkles }, { label: "Revisar antes de enviar", checked: reviewBeforeSend, setter: setReviewBeforeSend, Icon: ShieldCheck }];
  return <div className="mt-8 grid max-w-4xl gap-6 md:grid-cols-2"><Card className="rounded-2xl border-0 bg-[#102c36] text-white shadow-sm"><CardHeader><Badge className="w-fit bg-[#1b5960] text-[#b4e8dd]">Configuración personal</Badge><CardTitle className="mt-3 text-xl">Cómo trabaja tu agente</CardTitle></CardHeader><CardContent className="space-y-5">{controls.map(({ label, checked, setter, Icon }) => <div key={label} className="flex items-center justify-between"><span className="flex items-center gap-3 text-sm"><Icon className="h-4 w-4 text-[#9edbd1]" />{label}</span><Switch checked={checked} onCheckedChange={setter} /></div>)}</CardContent></Card><Card className="rounded-2xl border-0 bg-white shadow-sm"><CardHeader><CardTitle className="text-lg text-[#183841]">Conexión</CardTitle></CardHeader><CardContent><div className="rounded-xl bg-[#f2f8f7] p-4 text-sm text-[#52717a]">{connected ? "Tu cuenta aparece vinculada en esta sesión." : "Todavía no hay una cuenta Microsoft 365 vinculada."}</div><Button onClick={connect} className="mt-4 w-full rounded-xl bg-[#0d776c] text-white"><Link2 className="mr-2 h-4 w-4" />{connected ? "Actualizar vínculo" : "Vincular Microsoft 365"}</Button><p className="mt-4 text-xs leading-5 text-[#83999e]">El acceso real requiere una app registrada en Entra ID. La importación .ics permite probar el flujo sin credenciales.</p></CardContent></Card></div>;
}
function downloadEditableActaPdf(text: string) {
  const esc = (value: string) => value.replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)").replace(/[^\x20-\x7E]/g, "?");
  const raw = text.split("\n").map((line) => line.trim());
  const wrap = (value: string, max = 82) => { const result: string[] = []; if (!value) return [""]; for (let i = 0; i < value.length; i += max) result.push(value.slice(i, i + max)); return result; };
  const pages: string[] = []; let stream = ""; let y = 750; let row = 0;
  const textAt = (value: string, x: number, yy: number, size = 9, color = "0.15 0.23 0.28") => { stream += `BT ${color} rg /F1 ${size} Tf ${x} ${yy} Td (${esc(value)}) Tj ET\n`; };
  const rect = (x: number, yy: number, w: number, h: number, color: string) => { stream += `q ${color} rg ${x} ${yy} ${w} ${h} re f Q\n`; };
  const footer = () => { textAt("Teams Actas Agent  |  Documento generado para revisión", 174, 24, 7, "0.38 0.47 0.50"); pages.push(stream); stream = ""; };
  const startPage = () => { stream = ""; rect(0, 762, 612, 30, "0.07 0.23 0.36"); rect(48, 770, 5, 16, "0.83 0.61 0.20"); textAt("ACTA DE REUNIÓN", 66, 775, 18, "1 1 1"); textAt("SG-FO-07  |  Versión 10  |  10/11/2017", 66, 766, 7, "0.85 0.91 0.93"); y = 738; row = 0; };
  const section = (label: string) => { if (y < 70) { footer(); startPage(); } rect(48, y - 5, 516, 19, "0.05 0.47 0.42"); textAt(label.toUpperCase(), 59, y + 1, 9, "1 1 1"); y -= 25; };
  startPage();
  raw.forEach((line) => {
    const heading = /^(Acta de Reunión|Asistentes|Información de planeación de la reunión|Notas de la Reunión|Tareas de la Reunión|Decisiones)/i.test(line);
    if (heading) { section(line); return; }
    if (/^Código:/.test(line)) { textAt(line, 48, y, 8, "0.38 0.47 0.50"); y -= 15; return; }
    if (row === 0 && line && !/^Fecha:/.test(line)) { textAt(line, 48, y, 15, "0.07 0.23 0.36"); y -= 20; row = 1; return; }
    if (/^Fecha:/.test(line)) { rect(48, y - 5, 516, 22, "0.93 0.96 0.96"); textAt(line, 58, y + 2, 9, "0.15 0.23 0.28"); y -= 32; return; }
    if (/^Tarea \| Nota \| Responsable \| Estado \| Fecha$/i.test(line)) { rect(48, y - 5, 516, 21, "0.07 0.23 0.36"); textAt("TAREA", 53, y + 2, 7, "1 1 1"); textAt("NOTA", 184, y + 2, 7, "1 1 1"); textAt("RESPONSABLE", 315, y + 2, 7, "1 1 1"); textAt("ESTADO", 420, y + 2, 7, "1 1 1"); textAt("FECHA", 487, y + 2, 7, "1 1 1"); y -= 28; return; }
    if (line.includes(" | ") && y < 680) { const parts = line.split(" | "); rect(48, y - 5, 516, 26, row % 2 ? "0.97 0.98 0.98" : "0.91 0.96 0.95"); textAt(parts[0] || "-", 53, y + 2, 7); textAt(parts[1] || "-", 184, y + 2, 7); textAt(parts[2] || "-", 315, y + 2, 7); textAt(parts[3] || "-", 420, y + 2, 7); textAt(parts[4] || "-", 487, y + 2, 7); y -= 28; row += 1; return; }
    for (const part of wrap(line)) { if (y < 55) { footer(); startPage(); } textAt(part, 52, y, 9); y -= 14; }
  });
  footer();
  const objects: string[] = ["<< /Type /Catalog /Pages 2 0 R >>", "", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"];
  const pageIds: number[] = []; const contentIds: number[] = [];
  pages.forEach((page, index) => { const pageId = 4 + index * 2; const contentId = pageId + 1; pageIds.push(pageId); contentIds.push(contentId); objects.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R >> >> /Contents ${contentId} 0 R >>`); objects.push(`<< /Length ${page.length} >>\nstream\n${page}endstream`); });
  objects[1] = `<< /Type /Pages /Kids [${pageIds.map((id) => `${id} 0 R`).join(" ")}] /Count ${pages.length} >>`;
  const body = objects.map((object, index) => `${index + 1} 0 obj\n${object}\nendobj\n`).join("");
  const blob = new Blob([`%PDF-1.4\n${body}trailer\n<< /Root 1 0 R /Size ${objects.length + 1} >>\n%%EOF`], { type: "application/pdf" });
  const url = URL.createObjectURL(blob); const anchor = document.createElement("a"); anchor.href = url; anchor.download = "acta-plantilla-SG-FO-07.pdf"; document.body.appendChild(anchor); anchor.click(); anchor.remove(); URL.revokeObjectURL(url);
}

function Documents({ data, recordings, sharePointItems, sharePointLoading, preview, actaStatus, onLoadSharePoint, onEditPreview, onSendSharePoint, onGenerateFromRecording, onTest, onDelete, attendees }: { data?: Array<{ id: number; kind: string; format: string; fileName: string; storageUrl: string | null; meetingTitle?: string }>; recordings?: RecordingItem[]; sharePointItems: Array<{ id: string; note: string; modified: string }>; sharePointLoading: boolean; preview: string; actaStatus: { type: string; text: string }; onLoadSharePoint: () => void; onEditPreview: (value: string) => void; onSendSharePoint: () => void; onGenerateFromRecording: (id: number, url: string) => void; onTest: () => void; onDelete: (id: number) => void; attendees: Attendee[] }) {
  const [emailOpen, setEmailOpen] = useState(false);
  const [selectedEmails, setSelectedEmails] = useState<string[]>([]);
  const sendEmail = trpc.agent.sendMinutesEmail.useMutation({ onSuccess: () => { toast.success("Acta enviada por correo"); setEmailOpen(false); }, onError: (error) => toast.error(`No se pudo enviar el correo: ${error.message}`) });
  const sendActaByEmail = () => { const recipients = selectedEmails.filter(Boolean); if (!recipients.length) { toast.error("Selecciona al menos un asistente con correo"); return; } if (!preview.trim()) { toast.error("Primero genera o edita el acta"); return; } sendEmail.mutate({ recipients, subject: `Acta de reunión - ${attendees.length ? "reunión seleccionada" : "Teams Actas Agent"}`, text: `Hola,\n\nComparto el acta de la reunión para su revisión.\n\n${preview}\n\nSaludos.` }); };
  const statusClass = actaStatus.type === "error" ? "bg-[#fff0ed] text-[#a44238]" : actaStatus.type === "processing" ? "bg-[#fff6df] text-[#8b671b]" : actaStatus.type === "idle" ? "bg-[#f2f7f7] text-[#52717a]" : "bg-[#e8f7f1] text-[#16614f]";
  return <div className="mt-8 max-w-6xl space-y-6"><Card className="rounded-2xl border-0 bg-white shadow-sm"><CardHeader><div className="flex flex-wrap items-center justify-between gap-3"><div><CardTitle className="text-lg text-[#183841]">Actas y SharePoint</CardTitle><p className="mt-1 text-sm text-[#789098]">Revisa qué genera el agente y qué se guarda en Reuniones Efectivas.</p><div className={cn("mt-3 rounded-lg px-3 py-2 text-sm font-medium", statusClass)}>{actaStatus.type === "generated" || actaStatus.type === "sharepoint" ? "✓ " : ""}{actaStatus.text}</div></div><div className="flex gap-2"><Button onClick={onLoadSharePoint} disabled={sharePointLoading} variant="outline" className="rounded-xl border-[#b9e5da] text-[#164f4a]"><RefreshCw className={cn("mr-2 h-4 w-4", sharePointLoading && "animate-spin")} />Cargar SharePoint</Button><Button onClick={onTest} className="rounded-xl bg-[#0d776c] text-white"><WandSparkles className="mr-2 h-4 w-4" />Probar reunión Teams</Button><Button variant="outline" onClick={() => { if (!preview) { toast.info("Primero genera un acta para habilitar el envío por correo."); return; } setEmailOpen((value) => !value); }} className="rounded-xl border-[#f47c20] text-[#c85d12]"><UsersRound className="mr-2 h-4 w-4" />Enviar por correo</Button></div></div></CardHeader><CardContent><div className="rounded-xl border border-[#dce8e8] bg-[#f7fbfa] p-4"><p className="text-xs font-bold uppercase tracking-[0.12em] text-[#52717a]">Qué llena el agente</p><p className="mt-2 text-sm leading-6 text-[#36545d]">Puedes editar el texto antes de actualizar la columna <strong>Nota</strong>. No se escribe en SharePoint hasta pulsar el botón.</p>{preview && <div className="mt-4 rounded-lg border border-[#b9e5da] bg-white p-4"><p className="text-xs font-bold uppercase tracking-[0.12em] text-[#0d776c]">Texto editable para SharePoint</p><Textarea value={preview} onChange={(event) => onEditPreview(event.target.value)} rows={10} className="mt-2 bg-white text-sm leading-6 text-[#36545d]" /><div className="mt-4 overflow-hidden rounded-xl border border-[#ccdedd] bg-white shadow-sm"><div className="flex items-center justify-between bg-[#0b2f55] px-5 py-3 text-white"><span className="text-xs font-bold tracking-[0.14em]">VISTA DEL ACTA</span><span className="text-[10px] text-[#d8e7ed]">Así se verá en el PDF</span></div><div className="p-5"><div className="border-b-2 border-[#f47c20] pb-3"><p className="text-lg font-bold text-[#123b5d]">ACTA DE REUNIÓN</p><p className="text-[10px] text-[#607780]">SG-FO-07 | Versión 10 | 10/11/2017</p></div><div className="mt-4 space-y-2 text-xs leading-5 text-[#263b46]">{preview.split("\n").map((line, index) => { const isSection = /^(Asistentes|Información de planeación|Notas de la Reunión|Tareas de la Reunión|Decisiones)/i.test(line.trim()); const isTable = /^Tarea \| Nota \| Responsable \| Estado \| Fecha$/i.test(line.trim()); return isSection ? <div key={index} className="mt-4 rounded-md bg-[#0f6b78] px-3 py-2 font-bold tracking-[0.08em] text-white">{line}</div> : isTable ? <div key={index} className="rounded-md bg-[#f47c20] px-3 py-2 font-bold text-white">{line}</div> : <p key={index} className="whitespace-pre-wrap">{line || " "}</p>; })}</div><p className="mt-5 border-t pt-3 text-center text-[10px] text-[#607780]">Teams Actas Agent · Documento para revisión</p></div></div><div className="mt-3 flex flex-wrap gap-2"><Button variant="outline" onClick={() => downloadEditableActaPdf(preview)} className="rounded-xl text-xs text-[#52717a]"><Download className="mr-2 h-4 w-4" />Descargar acta de la plantilla (PDF)</Button><Button onClick={onSendSharePoint} className="rounded-xl bg-[#0d776c] text-white"><Cloud className="mr-2 h-4 w-4" />Enviar texto editado a SharePoint</Button><Button variant="outline" onClick={() => { setEmailOpen((value) => !value); if (!selectedEmails.length) setSelectedEmails(attendees.filter((item) => item.email).map((item) => item.email)); }} className="rounded-xl border-[#f47c20] text-[#c85d12]"><UsersRound className="mr-2 h-4 w-4" />Enviar por correo</Button></div>{emailOpen && <div className="mt-3 rounded-xl border border-[#f5c39f] bg-[#fffaf6] p-4"><p className="text-xs font-bold uppercase tracking-[0.12em] text-[#c85d12]">Selecciona los destinatarios</p>{!attendees.length ? <p className="mt-2 text-sm text-[#789098]">No hay asistentes cargados. Vincula Microsoft 365 y actualiza el calendario.</p> : <div className="mt-3 grid gap-2 sm:grid-cols-2">{attendees.map((person) => <label key={person.email || person.name} className="flex items-center gap-2 rounded-lg bg-white p-2 text-xs text-[#36545d]"><input type="checkbox" checked={selectedEmails.includes(person.email)} disabled={!person.email} onChange={(event) => setSelectedEmails((current) => event.target.checked ? [...current, person.email] : current.filter((email) => email !== person.email))} /><span><b>{person.name}</b><br /><span className="text-[#789098]">{person.email || "Sin correo disponible"}</span></span></label>)}</div>}<Button onClick={sendActaByEmail} disabled={!selectedEmails.length || sendEmail.isPending} className="mt-3 rounded-xl bg-[#f47c20] text-white"><UsersRound className="mr-2 h-4 w-4" />{sendEmail.isPending ? "Enviando…" : "Enviar acta por SMTP"}</Button><p className="mt-2 text-[11px] text-[#8d6b56]">El mensaje se envía desde el buzón fijo configurado en el servidor.</p></div>}</div>}</div></CardContent></Card><Card className="rounded-2xl border-0 bg-white shadow-sm"><CardHeader><CardTitle className="text-lg text-[#183841]">Grabaciones listas para generar acta</CardTitle><p className="text-sm text-[#789098]">El botón muestra el resultado arriba cuando termina.</p></CardHeader><CardContent>{!recordings?.length ? <p className="text-sm text-[#789098]">Todavía no hay grabaciones.</p> : <div className="space-y-3">{recordings.map((item) => <div key={item.meeting.id} className="flex flex-wrap items-center justify-between gap-3 rounded-xl border border-[#e1ecea] p-4"><div><p className="text-sm font-semibold text-[#183841]">{item.meeting.title}</p><p className="mt-1 text-xs text-[#789098]">Estado: {item.meeting.status}</p></div>{item.meeting.recordingUrl && <Button onClick={() => onGenerateFromRecording(item.meeting.id, item.meeting.recordingUrl!)} className="rounded-lg bg-[#164f4a] text-xs text-white">Generar desde grabación</Button>}</div>)}</div>}</CardContent></Card><Card className="rounded-2xl border-0 bg-white shadow-sm"><CardHeader><CardTitle className="text-lg text-[#183841]">Registros actuales de SharePoint</CardTitle><p className="text-sm text-[#789098]">Últimos 25 elementos de la lista configurada.</p></CardHeader><CardContent>{!sharePointItems.length ? <p className="rounded-xl border border-dashed p-6 text-sm text-[#789098]">Pulsa “Cargar SharePoint” para consultar la lista.</p> : <div className="space-y-3">{sharePointItems.map((item) => <div key={item.id} className="rounded-xl border border-[#e1ecea] p-4"><div className="flex justify-between gap-3"><span className="text-xs font-semibold text-[#52717a]">Registro #{item.id}</span><span className="text-xs text-[#789098]">{item.modified ? new Date(item.modified).toLocaleString() : ""}</span></div><p className="mt-2 whitespace-pre-wrap text-sm leading-6 text-[#36545d]">{item.note}</p></div>)}</div>}</CardContent></Card>{data?.filter((item) => item.kind === "minutes").map((item) => <Card key={item.id} className="rounded-2xl border-0 bg-white shadow-sm"><CardContent className="flex flex-wrap items-center justify-between gap-3 p-4"><div><p className="text-sm font-semibold text-[#36545d]">{item.meetingTitle || "Reunión"}</p><p className="text-xs text-[#789098]">{item.fileName}</p></div><>{item.storageUrl && item.format === "pdf" && <Button variant="outline" onClick={() => void downloadFile(item.storageUrl!, item.fileName).catch((error) => toast.error(error instanceof Error ? error.message : "No se pudo descargar el PDF"))} className="text-xs text-[#0d776c]"><Download className="mr-1 h-3 w-3" />Descargar PDF diseñado</Button>}<Button variant="outline" onClick={() => { if (window.confirm("¿Eliminar esta acta?")) onDelete(item.id); }} className="text-xs text-[#b24e45]"><Trash2 className="mr-1 h-3 w-3" />Eliminar</Button></></CardContent></Card>)}</div>;
}
function Commitments({ data, onUpdate }: { data?: Array<{ id: number; personName: string; personEmail?: string | null; action: string; dueDate?: string | null; status?: string; meetingTitle?: string }>; onUpdate: () => void }) { const update = trpc.agent.updateCommitmentStatus.useMutation({ onSuccess: onUpdate, onError: (error) => toast.error(error.message) }); const remove = trpc.agent.deleteCommitment.useMutation({ onSuccess: onUpdate, onError: (error) => toast.error(error.message) }); const edit = trpc.agent.updateCommitment.useMutation({ onSuccess: () => { onUpdate(); toast.success("Compromiso actualizado"); }, onError: (error) => toast.error(error.message) }); const labels: Record<string, string> = { open: "Pendiente", in_progress: "En progreso", done: "Completado", blocked: "Bloqueado" }; return <Card className="mt-8 max-w-5xl rounded-2xl border-0 bg-white shadow-sm"><CardHeader><CardTitle className="text-lg text-[#183841]">Compromisos por persona</CardTitle><p className="text-sm text-[#789098]">Las tareas y compromisos explícitos aparecen con responsable, acción y fecha; puedes editarlos antes de enviarlos.</p></CardHeader><CardContent className="space-y-3">{!data?.length ? <p className="rounded-xl border border-dashed p-6 text-sm text-[#789098]">Todavía no hay compromisos válidos.</p> : data.map((item) => <div key={item.id} className="flex flex-col gap-3 rounded-xl border p-4 sm:flex-row sm:items-center"><div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-full bg-[#dff3ee] text-xs font-semibold text-[#0d776c]">{initials(item.personName)}</div><div className="flex-1"><p className="text-sm font-semibold text-[#36545d]">{item.personName || "Responsable por definir"}</p><p className="text-[11px] text-[#0d776c]">{item.personEmail || "Sin correo asociado"}</p><p className="text-xs text-[#789098]">{item.action} · entrega {item.dueDate || "Por definir"}{item.meetingTitle ? ` · ${item.meetingTitle}` : ""}</p></div><div className="flex flex-wrap gap-1">{(["open", "in_progress", "done", "blocked"] as const).map((status) => <Button key={status} variant="outline" onClick={() => update.mutate({ id: item.id, status })} className={cn("h-8 rounded-lg px-2 text-[10px]", item.status === status && "border-[#0d776c] bg-[#e8f7f1] text-[#0d776c]")}>{labels[status]}</Button>)}<Button variant="outline" onClick={() => { const personName = window.prompt("Responsable", item.personName); if (personName === null) return; const action = window.prompt("Tarea o compromiso", item.action); if (action === null) return; const dueDate = window.prompt("Fecha o plazo", item.dueDate || "Por definir"); if (dueDate === null) return; edit.mutate({ id: item.id, personName, action, dueDate }); }} className="h-8 rounded-lg border-[#b9e5da] px-2 text-[10px] text-[#0d776c]">Editar</Button><Button variant="outline" onClick={() => { if (window.confirm("¿Eliminar este compromiso? Esta acción no se puede deshacer.")) remove.mutate({ id: item.id }); }} className="h-8 rounded-lg border-[#f0d5d1] px-2 text-[10px] text-[#b24e45]"><Trash2 className="mr-1 h-3 w-3" />Eliminar</Button></div></div>)}</CardContent></Card>; }


type RecordingItem = { meeting: { id: number; title: string; scheduledAt: Date | string; status: string; recordingUrl: string | null }; documents: Array<{ id: number; kind: string; format: string; fileName: string; storageUrl: string | null }> };
function Recordings({ data, onClear, onReprocess, onUpload, remoteFiles, onRefreshRemote, onDeleteRemote }: { data?: RecordingItem[]; onClear: (meetingId: number) => void; onReprocess: (meetingId: number, recordingUrl: string) => void; onUpload: (event: React.ChangeEvent<HTMLInputElement>) => void; remoteFiles: RemoteSharePointFile[]; onRefreshRemote: () => Promise<void>; onDeleteRemote: (id: string) => Promise<void> }) {
  return <div className="mt-8 max-w-5xl"><Card className="mb-6 rounded-2xl border-0 bg-white shadow-sm"><CardHeader><div className="flex flex-wrap items-center justify-between gap-3"><div><CardTitle className="text-lg text-[#183841]">Archivos permanentes en SharePoint</CardTitle><p className="mt-1 text-sm text-[#789098]">Estos archivos permanecen aunque Render se reinicie.</p></div><Button variant="outline" onClick={() => void onRefreshRemote()} className="rounded-xl border-[#b9e5da] text-xs text-[#0d776c]"><RefreshCw className="mr-2 h-3.5 w-3.5" />Actualizar SharePoint</Button></div></CardHeader><CardContent className="space-y-2">{!remoteFiles.length ? <p className="rounded-xl border border-dashed p-5 text-sm text-[#789098]">No hay archivos cargados o Microsoft 365 no está vinculado.</p> : remoteFiles.map((file) => <div key={file.id} className="flex flex-wrap items-center justify-between gap-3 rounded-xl border border-[#e1eceb] p-3"><div><p className="text-sm font-semibold text-[#183841]">{file.name}</p><p className="text-xs text-[#789098]">{Math.ceil(file.size / 1024)} KB · {file.modified ? new Date(file.modified).toLocaleString("es-CO") : ""}</p></div><div className="flex gap-2"><Button variant="outline" onClick={() => void downloadSharePointFile(file.id, file.name).catch((error) => toast.error(error instanceof Error ? error.message : "No se pudo descargar"))} className="h-8 rounded-lg px-2 text-xs text-[#0d776c]"><Download className="mr-1 h-3 w-3" />Descargar</Button><Button variant="outline" onClick={() => void onDeleteRemote(file.id)} className="h-8 rounded-lg px-2 text-xs text-[#b24e45]"><Trash2 className="mr-1 h-3 w-3" />Eliminar</Button></div></div>)}</CardContent></Card><Card className="rounded-2xl border-0 bg-white shadow-sm"><CardHeader><div className="flex items-center gap-3"><div className="flex h-10 w-10 items-center justify-center rounded-xl bg-[#e8f7f1]"><Video className="h-5 w-5 text-[#0d776c]" /></div><div><CardTitle className="text-lg text-[#183841]">Historial de grabaciones</CardTitle><p className="mt-1 text-sm text-[#789098]">Archivos capturados localmente y documentos producidos.</p></div></div><label className="inline-flex cursor-pointer items-center gap-2 rounded-xl bg-[#0d776c] px-3 py-2 text-xs font-semibold text-white hover:bg-[#095f57]"><Upload className="h-3.5 w-3.5" />Cargar grabación para reprocesar<input type="file" accept="audio/*,video/webm,video/mp4,.webm,.m4a,.mp3,.wav" className="hidden" onChange={onUpload} /></label></CardHeader><CardContent className="space-y-4">{!data?.length ? <div className="rounded-xl border border-dashed border-[#cfe0df] px-5 py-10 text-center text-sm text-[#789098]">Todavía no hay grabaciones procesadas. Activa <b>Grabar</b> en una reunión para comenzar.</div> : data.map(({ meeting, documents }) => <div key={meeting.id} className="rounded-2xl border border-[#e1eceb] p-5"><div className="flex flex-col justify-between gap-4 lg:flex-row lg:items-start"><div><div className="flex flex-wrap items-center gap-2"><h3 className="font-semibold text-[#183841]">{meeting.title}</h3><Badge className="bg-[#e8f7f1] text-[#24765f]">{meeting.status === "review" ? "Lista para revisión" : meeting.status}</Badge></div><p className="mt-2 text-xs text-[#789098]">{new Date(meeting.scheduledAt).toLocaleString("es-MX")} · Archivo WebM almacenado de forma privada</p></div><div className="flex flex-wrap gap-2">{meeting.recordingUrl && <Button variant="outline" onClick={() => void downloadFile(meeting.recordingUrl!, `${meeting.title.replace(/[^a-zA-Z0-9-_ ]/g, "").trim() || "grabacion"}.webm`).catch((error) => toast.error(error instanceof Error ? error.message : "No se pudo descargar la grabación"))} className="rounded-xl border-[#dce8e8] text-xs text-[#52717a]"><Download className="mr-1.5 h-3.5 w-3.5" />Descargar grabación</Button>}{meeting.recordingUrl && <Button variant="outline" onClick={() => onReprocess(meeting.id, meeting.recordingUrl!)} className="rounded-xl border-[#b9e5da] text-xs text-[#0d776c]"><RefreshCw className="mr-1.5 h-3.5 w-3.5" />Procesar nuevamente</Button>}<Button variant="outline" onClick={() => onClear(meeting.id)} className="rounded-xl border-[#f0d5d1] text-xs text-[#b24e45] hover:bg-[#fff3ef]"><Trash2 className="mr-1.5 h-3.5 w-3.5" />Eliminar referencia</Button></div></div><div className="mt-4 flex flex-wrap gap-2">{documents.map((document) => document.storageUrl && <Button key={document.id} variant="outline" onClick={() => void downloadFile(document.storageUrl!, document.fileName).catch((error) => toast.error(error instanceof Error ? error.message : "No se pudo descargar el documento"))} className="inline-flex items-center rounded-xl bg-[#f2f6f6] px-3 py-2 text-xs font-semibold text-[#52717a] hover:bg-[#e5efee]"><FileText className="mr-1.5 h-3.5 w-3.5 text-[#0d776c]" />{document.kind === "commitments" ? "Compromisos" : "Acta"} · {document.format.toUpperCase()}</Button>)}</div></div>)}</CardContent></Card></div>;
}
