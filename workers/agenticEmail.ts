import type { Env } from "./types";

export async function handleScheduled(controller: ScheduledController, env: Env, ctx: ExecutionContext) {
	try {
		console.log(`Cron triggered at ${new Date().toISOString()}`);

		// 1. Check Global Settings
		let autoReplyEnabled = false;
		try {
			const id = env.CHAT_SESSION.idFromName("default");
			const stub = env.CHAT_SESSION.get(id);
			const doSettingsRes = await stub.fetch(new Request("http://do/settings"));
			if (doSettingsRes.ok) {
				const doSettings = await doSettingsRes.json() as any;
				autoReplyEnabled = doSettings.agentAutoReplyEnabled === true;
			}
		} catch (e) {
			console.error("Failed to fetch DO settings for auto-reply check", e);
		}

		if (!autoReplyEnabled) {
			console.log("Agentic Auto-Reply is disabled. Exiting cron.");
			return;
		}

		// 2. Authenticate with Microsoft Graph using Client Credentials
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
		
		let pastEmailsContext = "No past context available.";
		let accessToken = "";
		
		if (tokenRes.ok) {
			const tokenData = await tokenRes.json() as { access_token: string };
			accessToken = tokenData.access_token;

			// Fetch the last 50 sent emails to learn writing style
			const sentRes = await fetch(`https://graph.microsoft.com/v1.0/users/${env.HUMAN_FALLBACK_EMAIL}/mailFolders/SentItems/messages?$top=50&$select=subject,bodyPreview&$orderby=receivedDateTime DESC`, {
				headers: { "Authorization": `Bearer ${accessToken}` }
			});
			
			if (sentRes.ok) {
				const sentData = await sentRes.json() as any;
				pastEmailsContext = sentData.value.map((msg: any) => `Subject: ${msg.subject}\nContent: ${msg.bodyPreview}`).join("\n\n---\n\n");
			}
		} else {
			console.error("Failed to authenticate with Graph:", await tokenRes.text());
			return;
		}

		// 3. Fetch Unread Emails from Inbox
		const unreadRes = await fetch(`https://graph.microsoft.com/v1.0/users/${env.HUMAN_FALLBACK_EMAIL}/mailFolders/Inbox/messages?$filter=isRead eq false&$top=10&$select=id,subject,from,bodyPreview,body`, {
			headers: { 
				"Authorization": `Bearer ${accessToken}`,
				"Prefer": 'outlook.body-content-type="text"'
			}
		});

		if (!unreadRes.ok) {
			console.error("Failed to fetch unread emails:", await unreadRes.text());
			return;
		}

		const unreadData = await unreadRes.json() as any;
		const unreadEmails = unreadData.value || [];
		console.log(`Found ${unreadEmails.length} unread emails to process.`);

		// 4. Process each email
		for (const email of unreadEmails) {
			const sender = email.from?.emailAddress?.address;
			if (!sender || sender === env.HUMAN_FALLBACK_EMAIL) continue;

			console.log(`Processing email from ${sender}: ${email.subject}`);

			const systemPrompt = `You are the autonomous email agent for Jagrit Sachdev (Founder & CEO of Zuup).
First, analyze the incoming email. 
1. If it's a newsletter, marketing, spam, calendar invite, or a purely automated notification, set <action>ignore</action>.
2. If it requires human attention, complex negotiation, extremely sensitive info, or you cannot confidently answer, set <action>human_required</action>.
3. If it is a standard customer interaction, greeting, simple query, or something you can easily resolve (like providing a slack link), set <action>reply</action>.

When replying, YOU MUST mimic Jagrit's EXACT writing style, tone, signature style, and length based on these past sent emails:
${pastEmailsContext}

Output MUST be wrapped in XML like this:
<response>
  <action>ignore or human_required or reply</action>
  <reply_text>Your drafted text here perfectly mimicking Jagrit, or empty if ignored/forwarded.</reply_text>
</response>`;

			const userPrompt = `Incoming email from ${sender}:
Subject: ${email.subject}
Body: ${email.body?.content || email.bodyPreview || "No text body found."}`;

			let rawOutput = "";
			try {
				// 1. Create Copilot Conversation
				const convRes = await fetch("https://graph.microsoft.com/beta/copilot/conversations", {
					method: "POST",
					headers: { "Authorization": `Bearer ${accessToken}`, "Content-Type": "application/json" },
					body: JSON.stringify({})
				});
				
				if (!convRes.ok) {
					console.error("Failed to create Copilot conversation:", await convRes.text());
					continue;
				}
				
				const convData = await convRes.json() as { id: string };
				const convId = convData.id;

				// 2. Send Chat Prompt
				const fullPrompt = `${systemPrompt}\n\n${userPrompt}`;
				const chatRes = await fetch(`https://graph.microsoft.com/beta/copilot/conversations/${convId}/chat`, {
					method: "POST",
					headers: { "Authorization": `Bearer ${accessToken}`, "Content-Type": "application/json" },
					body: JSON.stringify({
						message: {
							text: fullPrompt
						}
					})
				});

				if (!chatRes.ok) {
					console.error("Failed to execute Copilot chat:", await chatRes.text());
					continue;
				}

				const chatData = await chatRes.json() as any;
				// Extract response text from copilotConversation response structure
				const msgs = chatData.messages || [];
				// The API returns the conversation history. The last message is usually Copilot's response.
				const copilotMsg = msgs.length > 0 ? msgs[msgs.length - 1] : null;
				if (copilotMsg && copilotMsg.text) {
					rawOutput = copilotMsg.text;
				}
			} catch (err) {
				console.error("Error communicating with Copilot API:", err);
				continue;
			}
			
			let action = "human_required";
			let replyText = "";
			
			const actionMatch = rawOutput.match(/<action>\s*([\s\S]*?)\s*<\/action>/i);
			if (actionMatch) action = actionMatch[1].trim();
			
			const textMatch = rawOutput.match(/<reply_text>\s*([\s\S]*?)\s*<\/reply_text>/i);
			if (textMatch) replyText = textMatch[1].trim();

			console.log(`AI Decision for ${email.id}: ${action}`);

			// Execute Action
			if (action === "ignore") {
				// Mark as read so we don't process it again
				await fetch(`https://graph.microsoft.com/v1.0/users/${env.HUMAN_FALLBACK_EMAIL}/messages/${email.id}`, {
					method: "PATCH",
					headers: { "Authorization": `Bearer ${accessToken}`, "Content-Type": "application/json" },
					body: JSON.stringify({ isRead: true })
				});
				continue;
			}

			if (action === "human_required") {
				// We don't mark it as read so the human sees it as unread.
				// But to prevent infinite loops, we can tag it or move it, or just mark it read and flag it.
				// For now, let's just mark it read and flag it so it stays in the inbox but doesn't get processed again.
				await fetch(`https://graph.microsoft.com/v1.0/users/${env.HUMAN_FALLBACK_EMAIL}/messages/${email.id}`, {
					method: "PATCH",
					headers: { "Authorization": `Bearer ${accessToken}`, "Content-Type": "application/json" },
					body: JSON.stringify({ isRead: true, flag: { flagStatus: "flagged" } })
				});
			}

			if (action === "reply" && replyText) {
				const signature = `\n\n---\nNote: This email has been sent by ZuupMail AI Agent. If you have more questions, please reply back.`;
				const finalReplyText = replyText + signature;

				const replyPayload = {
					message: {
						subject: email.subject.startsWith("Re:") ? email.subject : `Re: ${email.subject}`,
						body: { contentType: "text", content: finalReplyText },
						toRecipients: [{ emailAddress: { address: sender } }]
					}
				};

				try {
					// 1. Send the autonomous reply via Microsoft Graph
					const sendRes = await fetch(`https://graph.microsoft.com/v1.0/users/${env.HUMAN_FALLBACK_EMAIL}/sendMail`, {
						method: "POST",
						headers: { 
							"Authorization": `Bearer ${accessToken}`,
							"Content-Type": "application/json"
						},
						body: JSON.stringify(replyPayload)
					});
					
					if (sendRes.ok) {
						// 2. Mark the original as read
						await fetch(`https://graph.microsoft.com/v1.0/users/${env.HUMAN_FALLBACK_EMAIL}/messages/${email.id}`, {
							method: "PATCH",
							headers: { "Authorization": `Bearer ${accessToken}`, "Content-Type": "application/json" },
							body: JSON.stringify({ isRead: true })
						});
						console.log("Autonomous reply sent successfully!");
					}
				} catch(e) {
					console.error("Failed to send auto-reply:", e);
				}
			}
		}
	} catch (error) {
		console.error("Error in cron handler:", error);
	}
}
