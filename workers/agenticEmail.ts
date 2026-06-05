import type { Env } from "./types";

export async function handleScheduled(controller: ScheduledController, env: Env, ctx: ExecutionContext) {
	try {
		console.log(`Cron triggered at ${new Date().toISOString()}`);

		// 1. Check Global Settings & fetch all registered mailbox profiles
		let autoReplyEnabled = true;
		let doSettings: Record<string, any> = {};

		try {
			const id = env.CHAT_SESSION.idFromName("default");
			const stub = env.CHAT_SESSION.get(id);
			const doSettingsRes = await stub.fetch(new Request("http://do/settings"));
			if (doSettingsRes.ok) {
				doSettings = await doSettingsRes.json() as any;
				autoReplyEnabled = doSettings.agentAutoReplyEnabled !== false;
			}
		} catch (e) {
			console.error("Failed to fetch DO settings, defaulting to enabled", e);
		}

		if (!autoReplyEnabled) {
			console.log("Agentic Auto-Reply is disabled. Exiting cron.");
			return;
		}

		// 2. Collect all registered mailbox profiles from DO settings
		const mailboxProfiles: Array<{
			email: string;
			name: string;
			customKnowledge: string;
			mailboxPurpose: string;
		}> = [];

		for (const [key, value] of Object.entries(doSettings)) {
			if (key.startsWith("profile_") && value && typeof value === "object") {
				const profile = value as any;
				const email = profile.email || profile.id;
				if (!email) continue;

				// Look for per-mailbox settings stored under settings_<email>
				const mailboxSettings = doSettings[`mailbox_settings_${email}`] || profile.settings || {};
				
				mailboxProfiles.push({
					email,
					name: profile.name || email,
					customKnowledge: mailboxSettings.agentKnowledgeBase || "",
					mailboxPurpose: mailboxSettings.agentMailboxPurpose || "",
				});
			}
		}

		// If no profiles registered, fall back to HUMAN_FALLBACK_EMAIL so there's always something to process
		if (mailboxProfiles.length === 0) {
			console.log("No mailbox profiles found in DO. Falling back to HUMAN_FALLBACK_EMAIL.");
			mailboxProfiles.push({
				email: env.HUMAN_FALLBACK_EMAIL,
				name: "Zuup AI",
				customKnowledge: "",
				mailboxPurpose: "",
			});
		}

		console.log(`Processing ${mailboxProfiles.length} mailbox(es): ${mailboxProfiles.map(m => m.email).join(", ")}`);

		// 3. Authenticate with Microsoft Graph (once, shared across all mailboxes)
		const params = new URLSearchParams();
		params.append("client_id", env.AZURE_CLIENT_ID);
		params.append("scope", "https://graph.microsoft.com/.default");
		params.append("client_secret", env.AZURE_CLIENT_SECRET || "");
		params.append("grant_type", "client_credentials");

		const tokenRes = await fetch(`https://login.microsoftonline.com/${env.AZURE_TENANT_ID}/oauth2/v2.0/token`, {
			method: "POST",
			headers: { "Content-Type": "application/x-www-form-urlencoded" },
			body: params
		});

		if (!tokenRes.ok) {
			const errorText = await tokenRes.text();
			console.error("Failed to authenticate with Graph:", errorText);
			throw new Error("Graph API Auth Failed! Details: " + errorText);
		}

		const tokenData = await tokenRes.json() as { access_token: string };
		const accessToken = tokenData.access_token;

		// 4. Process each registered mailbox independently
		for (const mailbox of mailboxProfiles) {
			console.log(`\n--- Processing mailbox: ${mailbox.email} ---`);
			await processMailbox(mailbox, accessToken, env);
		}

	} catch (error) {
		console.error("Error in cron handler:", error);
	}
}

async function processMailbox(
	mailbox: { email: string; name: string; customKnowledge: string; mailboxPurpose: string },
	accessToken: string,
	env: Env
) {
	// Build knowledge base from past sent emails of this specific mailbox
	let pastEmailsContext = "No past context available.";
	try {
		const kbRes = await fetch(
			`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox.email)}/mailFolders/SentItems/messages?$top=30&$select=subject,bodyPreview&$orderby=sentDateTime DESC`,
			{ headers: { "Authorization": `Bearer ${accessToken}` } }
		);

		if (kbRes.ok) {
			const kbData = await kbRes.json() as any;
			if (kbData.value && kbData.value.length > 0) {
				pastEmailsContext = kbData.value
					.map((msg: any) => `Subject: ${msg.subject}\nContent: ${msg.bodyPreview}`)
					.join("\n\n---\n\n");
			}
		} else {
			console.warn(`Could not fetch sent emails for ${mailbox.email}:`, await kbRes.text());
		}
	} catch (e) {
		console.error(`Error fetching knowledge base for ${mailbox.email}:`, e);
	}

	// Fetch unread emails for this mailbox
	let unreadEmails: any[] = [];
	try {
		const unreadRes = await fetch(
			`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox.email)}/mailFolders/Inbox/messages?$filter=isRead eq false&$top=10&$select=id,subject,from,toRecipients,bodyPreview,body,conversationId`,
			{
				headers: {
					"Authorization": `Bearer ${accessToken}`,
					"Prefer": 'outlook.body-content-type="text"'
				}
			}
		);

		if (!unreadRes.ok) {
			console.error(`Failed to fetch unread emails for ${mailbox.email}:`, await unreadRes.text());
			return;
		}

		const unreadData = await unreadRes.json() as any;
		unreadEmails = unreadData.value || [];
		console.log(`Found ${unreadEmails.length} unread email(s) for ${mailbox.email}`);
	} catch (e) {
		console.error(`Error fetching unread emails for ${mailbox.email}:`, e);
		return;
	}

	// Process each unread email
	for (const email of unreadEmails) {
		const sender = email.from?.emailAddress?.address;
		// Skip emails sent by this mailbox itself or by the fallback
		if (!sender || sender.toLowerCase() === mailbox.email.toLowerCase()) continue;

		console.log(`Processing email from ${sender}: "${email.subject}"`);

		// Get thread context
		let threadContext = "No specific thread history found.";
		if (email.conversationId) {
			try {
				const threadRes = await fetch(
					`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox.email)}/messages?$filter=conversationId eq '${email.conversationId}'&$select=id,from,bodyPreview,receivedDateTime&$orderby=receivedDateTime desc&$top=6`,
					{ headers: { "Authorization": `Bearer ${accessToken}` } }
				);
				if (threadRes.ok) {
					const threadData = await threadRes.json() as any;
					const threadMsgs = (threadData.value || []).filter((m: any) => m.id !== email.id);
					if (threadMsgs.length > 0) {
						threadContext = threadMsgs
							.reverse()
							.map((msg: any) => `From: ${msg.from?.emailAddress?.address || "Unknown"}\nDate: ${msg.receivedDateTime}\nContent: ${msg.bodyPreview}`)
							.join("\n\n---\n\n");
					}
				}
			} catch (e) {
				console.error("Error fetching thread context:", e);
			}
		}

		// Build the system prompt with custom knowledge and purpose
		const purposeSection = mailbox.mailboxPurpose
			? `\n\nMailbox Purpose / Role:\n${mailbox.mailboxPurpose}`
			: "";

		const customKnowledgeSection = mailbox.customKnowledge
			? `\n\nCustom Knowledge Base for this mailbox:\n<custom_knowledge>\n${mailbox.customKnowledge}\n</custom_knowledge>`
			: "";

		const systemPrompt = `You are the autonomous Zuup AI agent managing the inbox for: ${mailbox.email} (${mailbox.name}).${purposeSection}

First, analyze the incoming email:
1. If it is purely an automated notification, newsletter, or spam, set <action>ignore</action>.
2. For ANY user question, greeting, or request (even if you don't know the exact answer), you MUST set <action>reply</action>. If you don't know the answer, draft a polite reply saying you will follow up. STILL set <action>reply</action>. Only set <action>human_required</action> for severe emergencies.

Use the following knowledge base to answer questions:
<knowledge_base>
${pastEmailsContext}
</knowledge_base>${customKnowledgeSection}

You MUST sign off as "Zuup AI". Do NOT pretend to be the mailbox owner personally.

Output MUST be wrapped in XML:
<response>
  <action>ignore or human_required or reply</action>
  <reply_text>Your drafted reply text, or empty if ignored.</reply_text>
</response>`;

		const userPrompt = `Incoming email from ${sender}:
Subject: ${email.subject}

[Previous Thread History]:
${threadContext}

[Newest Unread Message Body]:
${email.body?.content || email.bodyPreview || "No text body found."}`;

		let rawOutput = "";
		try {
			const response = await env.AI.run("@cf/meta/llama-3.3-70b-instruct-fp8-fast", {
				messages: [
					{ role: "system", content: systemPrompt },
					{ role: "user", content: userPrompt }
				]
			}) as { response: string };
			rawOutput = response.response;
		} catch (err) {
			console.error("Error communicating with Cloudflare AI API:", err);
			continue;
		}

		let action = "human_required";
		let replyText = "";

		const actionMatch = rawOutput.match(/<action>\s*([\s\S]*?)\s*<\/action>/i);
		if (actionMatch) action = actionMatch[1].trim();

		const textMatch = rawOutput.match(/<reply_text>\s*([\s\S]*?)\s*<\/reply_text>/i);
		if (textMatch) replyText = textMatch[1].trim();

		console.log(`AI Decision for ${email.id}: ${action}`);

		if (action === "ignore") {
			await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox.email)}/messages/${email.id}`, {
				method: "PATCH",
				headers: { "Authorization": `Bearer ${accessToken}`, "Content-Type": "application/json" },
				body: JSON.stringify({ isRead: true })
			});
			continue;
		}

		if (action === "human_required") {
			await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox.email)}/messages/${email.id}`, {
				method: "PATCH",
				headers: { "Authorization": `Bearer ${accessToken}`, "Content-Type": "application/json" },
				body: JSON.stringify({ isRead: true, flag: { flagStatus: "flagged" } })
			});
			continue;
		}

		if (action === "reply" && replyText) {
			const signature = `\n\n---\nHello, I am Zuup AI. If you need further assistance, just reply to this email.`;
			const finalReplyText = replyText + signature;

			const replyPayload = {
				message: {
					subject: email.subject.startsWith("Re:") ? email.subject : `Re: ${email.subject}`,
					body: { contentType: "text", content: finalReplyText },
					toRecipients: [{ emailAddress: { address: sender } }],
					from: { emailAddress: { address: mailbox.email } }
				}
			};

			try {
				const sendRes = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox.email)}/sendMail`, {
					method: "POST",
					headers: {
						"Authorization": `Bearer ${accessToken}`,
						"Content-Type": "application/json"
					},
					body: JSON.stringify(replyPayload)
				});

				if (sendRes.ok) {
					await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox.email)}/messages/${email.id}`, {
						method: "PATCH",
						headers: { "Authorization": `Bearer ${accessToken}`, "Content-Type": "application/json" },
						body: JSON.stringify({ isRead: true })
					});
					console.log(`Autonomous reply sent from ${mailbox.email} to ${sender}`);
				} else {
					console.error(`Failed to send reply from ${mailbox.email}:`, await sendRes.text());
				}
			} catch (e) {
				console.error(`Error sending reply from ${mailbox.email}:`, e);
			}
		}
	}
}
