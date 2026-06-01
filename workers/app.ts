// Copyright (c) 2026 Cloudflare, Inc.
// Licensed under the Apache 2.0 license found in the LICENSE file or at:
//     https://opensource.org/licenses/Apache-2.0

import { Hono } from "hono";
import { createRequestHandler } from "react-router";
import { app as apiApp } from "./index";
import type { Env } from "./types";
import { handleScheduled } from "./agenticEmail";

declare module "react-router" {
	export interface AppLoadContext {
		cloudflare: {
			env: Env;
			ctx: ExecutionContext;
		};
	}
}

const requestHandler = createRequestHandler(
	() => import("virtual:react-router/server-build"),
	import.meta.env.MODE,
);

// Main app that wraps the API and adds React Router fallback
const app = new Hono<{ Bindings: Env }>();

// TODO: Add Microsoft Entra ID authentication middleware here

// Mount the API routes
app.route("/", apiApp);

// Test route to manually trigger the background cron job in local dev!
app.get("/test-cron", async (c) => {
	try {
		await handleScheduled({} as any, c.env, c.executionCtx as ExecutionContext);
		return c.text("Agentic Cron Job executed successfully! Check your terminal logs for details.");
	} catch (e: any) {
		return c.text("Cron Job failed: " + e.message, 500);
	}
});

// React Router catch-all: serves the SPA for all non-API routes
app.all("*", (c) => {
	return requestHandler(c.req.raw, {
		cloudflare: { env: c.env, ctx: c.executionCtx as ExecutionContext },
	});
});

// Export the Hono app as the default export
export default {
	fetch: app.fetch,
	scheduled: handleScheduled,
};

export { ChatSession } from "./ChatSession";
