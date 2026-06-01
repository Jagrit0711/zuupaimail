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

// React Router catch-all: serves the SPA for all non-API routes
app.all("*", (c) => {
	return requestHandler(c.req.raw, {
		cloudflare: { env: c.env, ctx: c.executionCtx as ExecutionContext },
	});
});

import PostalMime from "postal-mime";

// Export the Hono app as the default export
export default {
	fetch: app.fetch,
	scheduled: handleScheduled,
	async email(message: any, env: Env, ctx: ExecutionContext) {
		try {
			console.log(`Received email for ${message.to} from ${message.from}`);
			
			const id = env.CHAT_SESSION.idFromName("default");
			const stub = env.CHAT_SESSION.get(id);
			const doSettingsRes = await stub.fetch(new Request("http://do/settings"));
			if (!doSettingsRes.ok) {
				console.log("Could not fetch settings. Exiting.");
				return;
			}
			
			const doSettings = await doSettingsRes.json() as any;
			const autoReplyEnabled = doSettings.agentAutoReplyEnabled === true;
			
			if (!autoReplyEnabled) {
				console.log("Auto-reply is disabled globally.");
				return;
			}
			
			// Find token for this mailbox
			const token = doSettings[`token_${message.to}`];
			if (!token) {
				console.error(`No Graph token stored for mailbox ${message.to}. Cannot use Copilot.`);
				return;
			}

			// Parse email
			const rawEmail = await new Response(message.raw).arrayBuffer();
			const parsedEmail = await PostalMime.parse(rawEmail);
			const emailSubject = parsedEmail.subject || "(No Subject)";
			const emailText = parsedEmail.text || parsedEmail.html || "No body";

			const systemPrompt = `You are the autonomous email agent for the owner of ${message.to}.
First, analyze the incoming email. 
1. If it's a newsletter, marketing, spam, calendar invite, or a purely automated notification, set <action>ignore</action>.
2. If it requires human attention, complex negotiation, extremely sensitive info, or you cannot confidently answer, set <action>human_required</action>.
3. If it is a standard customer interaction, greeting, simple query, or something you can easily resolve (like providing a slack link), set <action>reply</action>.

Output MUST be wrapped in XML like this:
<response>
  <action>ignore or human_required or reply</action>
  <reply_text>Your drafted text here perfectly mimicking the user's style, or empty if ignored/forwarded.</reply_text>
</response>`;

			const userPrompt = `Incoming email from ${message.from}:
Subject: ${emailSubject}
Body: ${emailText}`;

			const fullPrompt = `${systemPrompt}\n\n${userPrompt}`;

			// 1. Create Copilot Conversation
			const convRes = await fetch("https://graph.microsoft.com/beta/copilot/conversations", {
				method: "POST",
				headers: { "Authorization": token, "Content-Type": "application/json" },
				body: JSON.stringify({})
			});
			
			if (!convRes.ok) {
				console.error("Failed to create Copilot conversation:", await convRes.text());
				return;
			}
			
			const convData = await convRes.json() as { id: string };
			const convId = convData.id;

			// 2. Send Chat Prompt
			const chatRes = await fetch(`https://graph.microsoft.com/beta/copilot/conversations/${convId}/chat`, {
				method: "POST",
				headers: { "Authorization": token, "Content-Type": "application/json" },
				body: JSON.stringify({
					message: {
						text: fullPrompt
					}
				})
			});

			if (!chatRes.ok) {
				console.error("Failed to execute Copilot chat:", await chatRes.text());
				return;
			}

			const chatData = await chatRes.json() as any;
			const msgs = chatData.messages || [];
			const copilotMsg = msgs.length > 0 ? msgs[msgs.length - 1] : null;
			let rawOutput = "";
			if (copilotMsg && copilotMsg.text) {
				rawOutput = copilotMsg.text;
			}

			let action = "human_required";
			let replyText = "";
			
			const actionMatch = rawOutput.match(/<action>\s*([\s\S]*?)\s*<\/action>/i);
			if (actionMatch) action = actionMatch[1].trim();
			
			const textMatch = rawOutput.match(/<reply_text>\s*([\s\S]*?)\s*<\/reply_text>/i);
			if (textMatch) replyText = textMatch[1].trim();

			console.log(`AI Decision for email from ${message.from}: ${action}`);

			if (action === "reply" && replyText) {
				const signature = `\n\n---\nNote: This email has been sent autonomously by ZuupMail AI Agent.`;
				const finalReplyText = replyText + signature;

				await env.SEB.send({
					from: message.to,
					to: message.from,
					subject: emailSubject.startsWith("Re:") ? emailSubject : `Re: ${emailSubject}`,
					text: finalReplyText,
				});
				console.log("Auto-reply sent successfully via SEB!");
			}
		} catch (err) {
			console.error("Error processing incoming email:", err);
		}
	}
};

export { ChatSession } from "./ChatSession";
