import { describe, expect, it } from "vitest";
import { powerAutomateTranscriptHandler, teamsGraphNotificationHandler } from "./teamsIntegration";

function responseMock() {
  const state: { statusCode: number; body?: unknown; type?: string } = { statusCode: 200 };
  const res = {
    type(value: string) { state.type = value; return res; },
    status(value: number) { state.statusCode = value; return res; },
    send(value: unknown) { state.body = value; return res; },
    json(value: unknown) { state.body = value; return res; },
    sendStatus(value: number) { state.statusCode = value; return res; },
  };
  return { res, state };
}

describe("Teams integration endpoints", () => {
  it("answers Graph validation tokens as plain text", () => {
    const { res, state } = responseMock();
    teamsGraphNotificationHandler({ query: { validationToken: "graph-token" }, body: {} } as never, res as never);
    expect(state.statusCode).toBe(200);
    expect(state.type).toBe("text/plain");
    expect(state.body).toBe("graph-token");
  });

  it("rejects transcript ingestion without the configured token", async () => {
    const { res, state } = responseMock();
    await powerAutomateTranscriptHandler({ headers: {}, header: () => undefined, body: {} } as never, res as never);
    expect(state.statusCode).toBe(401);
  });
});
