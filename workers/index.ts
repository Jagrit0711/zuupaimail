// Copyright (c) 2026 Cloudflare, Inc.
// Licensed under the Apache 2.0 license found in the LICENSE file or at:
//     https://opensource.org/licenses/Apache-2.0

import { Hono } from "hono";
import { cors } from "hono/cors";
import type { Env } from "./types";

const app = new Hono<{ Bindings: Env }>();

app.use("/api/*", cors({
	origin: (origin) => {
		if (!origin) return origin;
		try {
			const url = new URL(origin);
			if (url.hostname === "localhost" || url.hostname === "127.0.0.1") return origin;
		} catch { /* invalid origin */ }
		return undefined;
	},
}));

// Helper to get Graph API client
const graphFetch = async (url: string, c: any, options: RequestInit = {}) => {
	const token = c.req.header("Authorization");
	if (!token) throw new Error("Missing Authorization header");

	const res = await fetch(`https://graph.microsoft.com/v1.0${url}`, {
		...options,
		headers: {
			"Authorization": token,
			"Content-Type": "application/json",
			...(options.headers || {})
		}
	});
	if (!res.ok) throw new Error(`Graph API Error: ${res.statusText}`);
	return res.json();
};

// -- Config ---------------------------------------------------------
app.get("/api/v1/config", (c) => {
	const domainsRaw = c.env.DOMAINS || "";
	const domains = domainsRaw.split(",").map((d) => d.trim()).filter(Boolean);
	return c.json({ domains, emailAddresses: [] });
});

// -- Mailboxes ------------------------------------------------------
app.get("/api/v1/mailboxes", async (c) => {
	try {
		const user = await graphFetch("/me", c);
		return c.json([{
			id: user.mail || user.userPrincipalName,
			email: user.mail || user.userPrincipalName,
			name: user.displayName,
			settings: { fromName: user.displayName }
		}]);
	} catch (e) {
		return c.json([{ id: "default", email: "auth_required@example.com", name: "Please Login" }]);
	}
});

app.get("/api/v1/mailboxes/:mailboxId", async (c) => {
	const user = await graphFetch("/me", c);
	return c.json({
		id: user.mail || user.userPrincipalName,
		email: user.mail || user.userPrincipalName,
		name: user.displayName,
		settings: { fromName: user.displayName }
	});
});

// -- Helpers for Mapping Graph to Agentic Inbox --
const mapGraphMessageToEmail = (msg: any) => ({
	id: msg.id,
	thread_id: msg.conversationId,
	subject: msg.subject || "(No Subject)",
	sender: msg.from?.emailAddress?.address || "unknown",
	recipient: msg.toRecipients?.map((r: any) => r.emailAddress?.address).join(", ") || "",
	cc: msg.ccRecipients?.map((r: any) => r.emailAddress?.address).join(", ") || "",
	bcc: msg.bccRecipients?.map((r: any) => r.emailAddress?.address).join(", ") || "",
	date: msg.receivedDateTime,
	read: msg.isRead,
	starred: false, // E5 doesn't map directly to a simple starred flag easily without flag status
	body: msg.body?.content || "",
	snippet: msg.bodyPreview || "",
	has_attachment: msg.hasAttachments
});

// -- Emails ---------------------------------------------------------
app.get("/api/v1/mailboxes/:mailboxId/emails", async (c) => {
	const folder = c.req.query("folder");
	let url = `/me/messages?$top=50&$orderby=receivedDateTime DESC`;
	
	if (folder === "sent") url = `/me/mailFolders/SentItems/messages?$top=50&$orderby=receivedDateTime DESC`;
	else if (folder === "draft") url = `/me/mailFolders/Drafts/messages?$top=50&$orderby=receivedDateTime DESC`;
	else if (folder === "trash") url = `/me/mailFolders/DeletedItems/messages?$top=50&$orderby=receivedDateTime DESC`;

	try {
		const response = await graphFetch(url, c);
		const emails = response.value.map(mapGraphMessageToEmail);
		return c.json({ emails, totalCount: response.value.length });
	} catch (e) {
		return c.json({ error: (e as Error).message }, 500);
	}
});

app.post("/api/v1/mailboxes/:mailboxId/emails", async (c) => {
	const body = await c.req.json();
	const token = c.req.header("Authorization");
	
	const message = {
		message: {
			subject: body.subject,
			body: { contentType: "HTML", content: body.html || body.text },
			toRecipients: body.to.split(",").map((t: string) => ({ emailAddress: { address: t.trim() } }))
		}
	};

	const res = await fetch(`https://graph.microsoft.com/v1.0/me/sendMail`, {
		method: "POST",
		headers: { "Authorization": token!, "Content-Type": "application/json" },
		body: JSON.stringify(message)
	});
	
	if (!res.ok) return c.json({ error: "Failed to send email" }, 500);
	return c.json({ id: "sent", status: "sent" }, 202);
});

app.get("/api/v1/mailboxes/:mailboxId/emails/:id", async (c) => {
	const msg = await graphFetch(`/me/messages/${c.req.param("id")}`, c);
	return c.json(mapGraphMessageToEmail(msg));
});

app.post("/api/v1/mailboxes/:mailboxId/emails/:id/reply", async (c) => {
	const id = c.req.param("id");
	const body = await c.req.json();
	const token = c.req.header("Authorization");
	
	const replyPayload = {
		message: {
			toRecipients: body.to ? body.to.split(",").map((t: string) => ({ emailAddress: { address: t.trim() } })) : []
		},
		comment: body.html || body.text || ""
	};

	const res = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${id}/reply`, {
		method: "POST",
		headers: { "Authorization": token!, "Content-Type": "application/json" },
		body: JSON.stringify(replyPayload)
	});
	
	if (!res.ok) return c.json({ error: "Failed to reply to email" }, 500);
	return c.json({ status: "replied" }, 202);
});

app.post("/api/v1/mailboxes/:mailboxId/emails/:id/forward", async (c) => {
	const id = c.req.param("id");
	const body = await c.req.json();
	const token = c.req.header("Authorization");
	
	const forwardPayload = {
		toRecipients: body.to ? body.to.split(",").map((t: string) => ({ emailAddress: { address: t.trim() } })) : [],
		comment: body.html || body.text || ""
	};

	const res = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${id}/forward`, {
		method: "POST",
		headers: { "Authorization": token!, "Content-Type": "application/json" },
		body: JSON.stringify(forwardPayload)
	});
	
	if (!res.ok) return c.json({ error: "Failed to forward email" }, 500);
	return c.json({ status: "forwarded" }, 202);
});

// -- Folders --------------------------------------------------------
app.get("/api/v1/mailboxes/:mailboxId/folders", async (c) => {
	return c.json([
		{ id: "inbox", name: "Inbox", unreadCount: 0 },
		{ id: "sent", name: "Sent", unreadCount: 0 },
		{ id: "draft", name: "Drafts", unreadCount: 0 },
		{ id: "trash", name: "Trash", unreadCount: 0 }
	]);
});

// -- Search ---------------------------------------------------------
app.get("/api/v1/mailboxes/:mailboxId/search", async (c) => {
	const query = c.req.query("query") || "";
	const url = `/me/messages?$search="${query}"`;
	const response = await graphFetch(url, c);
	const emails = response.value.map(mapGraphMessageToEmail);
	return c.json({ emails, totalCount: response.value.length });
});

// -- AI Stateless Agent ---------------------------------------------
app.post("/api/v1/ai/draft", async (c) => {
	const body = await c.req.json();
	const emailText = body.emailText || "";
	const userPrompt = body.userPrompt || "";
	
	const systemPrompt = `You are a highly capable, professional executive assistant for Jagrit Sachdev (Founder & CEO of Zuup).
Your job is to draft concise, highly professional, and natural-sounding email replies.
Avoid corporate buzzwords, generic templates, or robotic phrasing like "Dear [Name]" or "I will take care of the matter".
Write exactly as a busy CEO would: direct, polite, and to the point.
NEVER include placeholders like [Your Name] or [Insert Date]. ALWAYS use context clues to fill in names, or omit them if unknown.
Respond ONLY with the final drafted email body. Do not include any meta-commentary, subject lines, or explanations.`;

	const finalUserPrompt = `I need you to draft an email reply based on my instructions.
My instructions: "${userPrompt}"

Here is the email context you are replying to (if any):
${emailText}

Draft the reply now.`;

	try {
		const response = await c.env.AI.run("@cf/meta/llama-3.1-8b-instruct", {
			messages: [
				{ role: "system", content: systemPrompt },
				{ role: "user", content: finalUserPrompt }
			]
		});
		
		// @ts-ignore
		return c.json({ draft: response.response });
	} catch (e) {
		return c.json({ error: "Failed to generate AI draft" }, 500);
	}
});

export { app };
