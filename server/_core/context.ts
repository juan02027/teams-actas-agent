import type { CreateExpressContextOptions } from "@trpc/server/adapters/express";
import type { User } from "../../drizzle/schema";
import { sdk } from "./sdk";

export type TrpcContext = {
  req: CreateExpressContextOptions["req"];
  res: CreateExpressContextOptions["res"];
  user: User | null;
};

export async function createContext(
  opts: CreateExpressContextOptions
): Promise<TrpcContext> {
  let user: User | null = null;

  if (process.env.LOCAL_OPERATOR_MODE === "true") {
    user = {
      id: 0,
      openId: "local-operator",
      email: "local@teams-actas-agent",
      name: process.env.LOCAL_OPERATOR_NAME || "Operador local",
      loginMethod: "local",
      role: "admin",
      createdAt: new Date(0),
      updatedAt: new Date(),
      lastSignedIn: new Date(),
    };
  }

  try {
    if (user) return { req: opts.req, res: opts.res, user };
    user = await sdk.authenticateRequest(opts.req);
  } catch (error) {
    // Authentication is optional for public procedures.
    user = null;
  }

  return {
    req: opts.req,
    res: opts.res,
    user,
  };
}
