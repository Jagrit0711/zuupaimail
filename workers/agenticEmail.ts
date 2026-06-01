import type { Env } from "./types";

export async function handleScheduled(controller: ScheduledController, env: Env, ctx: ExecutionContext) {
	console.log(`Cron triggered at ${new Date().toISOString()}`);
	console.log("Email polling via cron is deprecated in favor of real-time Cloudflare Email Routing. Exiting.");
}

