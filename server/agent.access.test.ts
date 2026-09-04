import { describe, expect, it } from "vitest";
import { appRouter } from "./routers";
import type { TrpcContext } from "./_core/context";

type AuthenticatedUser = NonNullable<TrpcContext["user"]>;

function createContext(role: "user" | "admin"): TrpcContext {
  const user: AuthenticatedUser = {
    id: role === "admin" ? 99 : 98,
    openId: `${role}-sample`,
    email: `${role}@example.com`,
    name: role === "admin" ? "Operador" : "Colaborador",
    loginMethod: "manus",
    role,
    createdAt: new Date(),
    updatedAt: new Date(),
    lastSignedIn: new Date(),
  };
  return {
    user,
    req: { protocol: "https", headers: {} } as TrpcContext["req"],
    res: {} as TrpcContext["res"],
  };
}

describe("agent operator access", () => {
  it("blocks a regular user from reading company meetings", async () => {
    const caller = appRouter.createCaller(createContext("user"));
    await expect(caller.agent.meetings()).rejects.toMatchObject({ code: "FORBIDDEN" });
  });

  it("blocks a regular user from changing meeting controls", async () => {
    const caller = appRouter.createCaller(createContext("user"));
    await expect(caller.agent.setMeetingControls({
      id: 1,
      recordingEnabled: true,
      processingEnabled: true,
    })).rejects.toMatchObject({ code: "FORBIDDEN" });
  });
});
