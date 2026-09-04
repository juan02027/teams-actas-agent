import { ENV } from "./_core/env";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
function graphUrl(pathname: string) { return `${GRAPH_BASE}${pathname}`; }
async function graphRequest(accessToken: string, pathname: string, init: RequestInit = {}) {
  const response = await fetch(graphUrl(pathname), { ...init, headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json", ...(init.headers || {}) } });
  if (!response.ok) {
    const detail = await response.text().catch(() => "");
    throw new Error(`Microsoft Graph respondió ${response.status}: ${detail || response.statusText}`);
  }
  return response.status === 204 ? null : response.json();
}
function siteParts(siteUrl: string) {
  const parsed = new URL(siteUrl);
  const relativePath = parsed.pathname.replace(/\/$/, "") || "/";
  return { hostname: parsed.hostname, relativePath };
}

export function buildSharePointNote(input: { meetingTitle: string; executiveSummary: string; objective: string; decisions: string[]; commitments: Array<{ personName: string; action: string; dueDate: string }> }) {
  const lines = [`Reunión: ${input.meetingTitle}`, `Resumen: ${input.executiveSummary}`, `Objetivo: ${input.objective}`];
  if (input.decisions.length) lines.push(`Decisiones: ${input.decisions.join("; ")}`);
  if (input.commitments.length) lines.push(`Compromisos: ${input.commitments.map((item) => `${item.personName}: ${item.action} (entrega ${item.dueDate})`).join("; ")}`);
  return lines.join("\n").slice(0, 63999);
}

export async function createSharePointMeetingNote(input: { accessToken: string; note: string }) {
  if (!input.accessToken) throw new Error("Falta el token delegado de Microsoft 365.");
  const { hostname, relativePath } = siteParts(ENV.sharePointSiteUrl);
  const site = await graphRequest(input.accessToken, `/sites/${encodeURIComponent(hostname)}:${relativePath === "/" ? "/" : `/${relativePath.replace(/^\//, "")}`}`) as { id?: string };
  if (!site?.id) throw new Error("Graph no devolvió el identificador del sitio de SharePoint.");
  const listName = encodeURIComponent(ENV.sharePointListName);
  const list = await graphRequest(input.accessToken, `/sites/${encodeURIComponent(site.id)}/lists/${listName}?$select=id,displayName`) as { id?: string };
  if (!list?.id) throw new Error(`No se encontró la lista «${ENV.sharePointListName}».`);
  const item = await graphRequest(input.accessToken, `/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(list.id)}/items`, { method: "POST", body: JSON.stringify({ fields: { [ENV.sharePointColumnName]: input.note } }) }) as { id?: string };
  return { siteId: site.id, listId: list.id, itemId: item?.id || null, column: ENV.sharePointColumnName };
}
