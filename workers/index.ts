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
	const id = c.env.CHAT_SESSION.idFromName("default");
	const stub = c.env.CHAT_SESSION.get(id);

	try {
		// Attempt to register the current logged-in user if a token is provided
		const token = c.req.header("Authorization");
		if (token && token !== "null" && token !== "undefined") {
			const user = await graphFetch("/me", c);
			const email = user.mail || user.userPrincipalName;
			
			// Fetch existing profile so we can merge settings
			const existingSettingsRes = await stub.fetch(new Request("http://do/settings"));
			const existingSettings = existingSettingsRes.ok ? await existingSettingsRes.json() as any : {};
			const existingProfile = existingSettings[`profile_${email}`] || {};

			// Merge: preserve any existing per-mailbox settings the user has saved
			await stub.fetch(new Request("http://do/settings", {
				method: "POST",
				headers: { "Content-Type": "application/json" },
				body: JSON.stringify({ 
					[`profile_${email}`]: {
						id: email,
						email: email,
						name: user.displayName,
						settings: { 
							fromName: user.displayName,
							...(existingProfile.settings || {})
						}
					},
					[`token_${email}`]: token
				})
			}));
		}
	} catch (e) {
		console.warn("Failed to fetch /me from graph (token may be expired or missing)", e);
	}

	// Fetch all stored profiles from the DO and return them
	const doSettingsRes = await stub.fetch(new Request("http://do/settings"));
	if (doSettingsRes.ok) {
		const doSettings = await doSettingsRes.json() as any;
		const profiles = Object.keys(doSettings)
			.filter(key => key.startsWith("profile_"))
			.map(key => {
				const profile = doSettings[key];
				// Merge per-mailbox settings stored under mailbox_settings_<email>
				const perMailboxSettings = doSettings[`mailbox_settings_${profile.email}`] || {};
				return {
					...profile,
					settings: { ...(profile.settings || {}), ...perMailboxSettings }
				};
			});
		
		if (profiles.length > 0) {
			return c.json(profiles);
		}
	}

	return c.json([{ id: "default", email: "auth_required@example.com", name: "Please Login" }]);
});

// POST /api/v1/mailboxes - manually register a mailbox (for non-OAuth accounts like other Zuup members)
app.post("/api/v1/mailboxes", async (c) => {
	try {
		const body = await c.req.json();
		const email = body.email;
		const name = body.name || email?.split("@")[0] || email;
		
		if (!email) return c.json({ error: "email is required" }, 400);

		const id = c.env.CHAT_SESSION.idFromName("default");
		const stub = c.env.CHAT_SESSION.get(id);

		// Fetch existing profile (don't overwrite settings if profile exists)
		const existingSettingsRes = await stub.fetch(new Request("http://do/settings"));
		const existingSettings = existingSettingsRes.ok ? await existingSettingsRes.json() as any : {};
		const existingProfile = existingSettings[`profile_${email}`];

		const profile = existingProfile || {
			id: email,
			email,
			name,
			settings: { fromName: name, ...(body.settings || {}) }
		};

		if (!existingProfile) {
			await stub.fetch(new Request("http://do/settings", {
				method: "POST",
				headers: { "Content-Type": "application/json" },
				body: JSON.stringify({ [`profile_${email}`]: profile })
			}));
		}

		return c.json(profile, 201);
	} catch (e: any) {
		return c.json({ error: e.message }, 500);
	}
});

app.get("/api/v1/mailboxes/:mailboxId", async (c) => {
	const mailboxId = c.req.param("mailboxId");
	
	// Fetch from DO first (stored profiles are the source of truth)
	const id = c.env.CHAT_SESSION.idFromName("default");
	const stub = c.env.CHAT_SESSION.get(id);
	const doSettingsRes = await stub.fetch(new Request("http://do/settings"));
	const doSettings = doSettingsRes.ok ? await doSettingsRes.json() as any : {};

	// Try to find the profile by email key
	const profileKey = `profile_${mailboxId}`;
	const storedProfile = doSettings[profileKey];
	// Merge per-mailbox knowledge settings
	const perMailboxSettings = doSettings[`mailbox_settings_${mailboxId}`] || {};

	if (storedProfile) {
		return c.json({
			...storedProfile,
			settings: { ...(storedProfile.settings || {}), ...perMailboxSettings }
		});
	}

	// Fallback: try to get from Graph if a valid token is available
	try {
		const token = c.req.header("Authorization");
		if (token && token !== "null") {
			const user = await graphFetch("/me", c);
			const email = user.mail || user.userPrincipalName;
			return c.json({
				id: email,
				email,
				name: user.displayName,
				settings: { fromName: user.displayName, ...perMailboxSettings }
			});
		}
	} catch (e) {
		console.warn("Fallback /me fetch failed for mailboxId:", mailboxId, e);
	}

	return c.json({ id: mailboxId, email: mailboxId, name: mailboxId, settings: perMailboxSettings });
});

app.delete("/api/v1/mailboxes/:mailboxId", async (c) => {
	const mailboxId = c.req.param("mailboxId");
	const id = c.env.CHAT_SESSION.idFromName("default");
	const stub = c.env.CHAT_SESSION.get(id);

	// Delete the profile and associated token/settings from DO
	// We store null to mark it as deleted (DO doesn't have a delete key API, so we overwrite with empty)
	await stub.fetch(new Request("http://do/settings", {
		method: "POST",
		headers: { "Content-Type": "application/json" },
		body: JSON.stringify({ 
			[`profile_${mailboxId}`]: null,
			[`token_${mailboxId}`]: null,
			[`mailbox_settings_${mailboxId}`]: null
		})
	}));

	return c.json({ status: "deleted" });
});

app.put("/api/v1/mailboxes/:mailboxId", async (c) => {
	const mailboxId = c.req.param("mailboxId");
	const body = await c.req.json();
	const settings = body.settings || {};
	
	// Forward settings to DO
	const id = c.env.CHAT_SESSION.idFromName("default");
	const stub = c.env.CHAT_SESSION.get(id);

	// Update both the profile settings and any top-level settings keys
	// First, read existing profile to merge fromName/agentSystemPrompt etc.
	const doSettingsRes = await stub.fetch(new Request("http://do/settings"));
	const doSettings = doSettingsRes.ok ? await doSettingsRes.json() as any : {};
	const existingProfile = doSettings[`profile_${mailboxId}`];

	const updatePayload: Record<string, any> = {};

	// Extract per-profile fields vs top-level settings
	const { agentKnowledgeBase, agentMailboxPurpose, agentAutoReplyEnabled, fromName, agentSystemPrompt, ...restSettings } = settings;

	// Update the profile's nested settings
	if (existingProfile) {
		updatePayload[`profile_${mailboxId}`] = {
			...existingProfile,
			settings: {
				...(existingProfile.settings || {}),
				...(fromName !== undefined ? { fromName } : {}),
				...(agentSystemPrompt !== undefined ? { agentSystemPrompt } : {}),
				...(agentKnowledgeBase !== undefined ? { agentKnowledgeBase } : {}),
				...(agentMailboxPurpose !== undefined ? { agentMailboxPurpose } : {}),
				...(agentAutoReplyEnabled !== undefined ? { agentAutoReplyEnabled } : {}),
			}
		};
	}

	// Handle any nested mailbox_settings_<email> keys passed in (from the settings page double-save)
	for (const [key, value] of Object.entries(restSettings)) {
		if (key.startsWith("mailbox_settings_")) {
			updatePayload[key] = value;
		}
	}

	// Also save agentAutoReplyEnabled at top level for backward compat
	if (agentAutoReplyEnabled !== undefined) {
		updatePayload.agentAutoReplyEnabled = agentAutoReplyEnabled;
	}

	if (Object.keys(updatePayload).length > 0) {
		await stub.fetch(new Request("http://do/settings", {
			method: "POST",
			headers: { "Content-Type": "application/json" },
			body: JSON.stringify(updatePayload)
		}));
	}

	return c.json({ status: "success" });
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
		
		// Group by conversationId to get thread_counts
		const threadCounts: Record<string, number> = {};
		response.value.forEach((msg: any) => {
			if (msg.conversationId) {
				threadCounts[msg.conversationId] = (threadCounts[msg.conversationId] || 0) + 1;
			}
		});

		// Deduplicate the list to only show the newest message per thread
		const uniqueThreads = new Map<string, any>();
		const nonThreads: any[] = [];
		
		response.value.forEach((msg: any) => {
			if (msg.conversationId) {
				// Since graphFetch orderby is receivedDateTime DESC, the first one we see is the newest
				if (!uniqueThreads.has(msg.conversationId)) {
					uniqueThreads.set(msg.conversationId, msg);
				}
			} else {
				nonThreads.push(msg);
			}
		});

		const merged = [...Array.from(uniqueThreads.values()), ...nonThreads]
			.sort((a, b) => new Date(b.receivedDateTime).getTime() - new Date(a.receivedDateTime).getTime());

		const emails = merged.map((msg: any) => ({
			...mapGraphMessageToEmail(msg),
			thread_count: threadCounts[msg.conversationId] || 1
		}));
		
		return c.json({ emails, totalCount: emails.length });
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

// -- AI Stateful Agent (Durable Object Proxy) ----------------------
app.post("/api/v1/ai/chat", async (c) => {
	const id = c.env.CHAT_SESSION.idFromName("default");
	const stub = c.env.CHAT_SESSION.get(id);
	
	const req = new Request("http://do/chat", {
		method: "POST",
		headers: c.req.raw.headers,
		body: await c.req.raw.clone().arrayBuffer()
	});
	return await stub.fetch(req);
});

app.get("/api/v1/ai/history", async (c) => {
	const id = c.env.CHAT_SESSION.idFromName("default");
	const stub = c.env.CHAT_SESSION.get(id);
	
	const req = new Request("http://do/history", {
		method: "GET"
	});
	return await stub.fetch(req);
});

export { app };
export { ChatSession } from "./ChatSession";
