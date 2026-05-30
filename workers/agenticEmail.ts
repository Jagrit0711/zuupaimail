import PostalMime from "postal-mime";
import { createMimeMessage } from "mimetext";
import { EmailMessage } from "cloudflare:email";
import type { Env } from "./types";

export async function handleIncomingEmail(message: ForwardableEmailMessage, env: Env, ctx: ExecutionContext) {
	try {
		console.log(`Received email from ${message.from} to ${message.to}`);

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
			console.log("Agentic Auto-Reply is disabled. Routing normally.");
			await message.forward(env.HUMAN_FALLBACK_EMAIL);
			return;
		}

		// 2. Parse the incoming email
		const parsedEmail = await PostalMime.parse(message.raw);
		const sender = message.from;

		// 3. Authenticate with Microsoft Graph using Client Credentials
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
			} else {
				console.error("Failed to fetch sent items:", await sentRes.text());
			}
		} else {
			console.error("Failed to authenticate with Graph:", await tokenRes.text());
		}

		// 4. Triage & Draft via Llama 3.1
		const systemPrompt = `You are the autonomous email agent for Jagrit Sachdev (Founder & CEO of Zuup).
First, analyze the incoming email. 
1. If it's a newsletter, marketing, spam, calendar invite, or a purely automated notification, set "action": "ignore".
2. If it requires human attention, complex negotiation, extremely sensitive info, or you cannot confidently answer, set "action": "human_required".
3. If it is a standard customer interaction, greeting, simple query, or something you can easily resolve, set "action": "reply".

When replying, YOU MUST mimic Jagrit's EXACT writing style, tone, signature style, and length based on these past sent emails:
${pastEmailsContext}

Output MUST be a valid JSON object ONLY:
{
  "action": "ignore" | "human_required" | "reply",
  "reply_text": "Your drafted text here perfectly mimicking Jagrit, or empty if ignored/forwarded."
}`;

		const userPrompt = `Incoming email from ${sender} to ${message.to}:
Subject: ${parsedEmail.subject}
Body: ${parsedEmail.text || "No text body found."}`;

		const aiResponse = await env.AI.run("@cf/meta/llama-3.1-8b-instruct", {
			messages: [
				{ role: "system", content: systemPrompt },
				{ role: "user", content: userPrompt }
			]
		});
		
		// @ts-ignore
		const rawOutput = aiResponse.response || "";
		// Extract JSON from potential markdown wrapping
		const jsonMatch = rawOutput.match(/\{[\s\S]*\}/);
		let decision = { action: "human_required", reply_text: "" };
		
		if (jsonMatch) {
			try { decision = JSON.parse(jsonMatch[0]); } catch(e) {}
		}

		console.log(`AI Decision: ${decision.action}`);

		// 5. Execute Action
		if (decision.action === "ignore") {
			console.log("Ignored as spam/newsletter.");
			return;
		}

		if (decision.action === "human_required") {
			const ticketId = `#ZP-${Math.floor(Math.random() * 100000)}`;
			
			// Auto-reply with ticket ID using MS Graph
			const autoReplyPayload = {
				message: {
					subject: `Re: ${parsedEmail.subject} (Ticket ${ticketId})`,
					body: { contentType: "text", content: `Hi there,\n\nWe received your message and a human will look into this shortly.\n\nYour reference token is ${ticketId}.\n\nBest,\nZuup AI Agent` },
					toRecipients: [{ emailAddress: { address: sender } }]
				}
			};

			try {
				if (accessToken) {
					await fetch(`https://graph.microsoft.com/v1.0/users/${env.HUMAN_FALLBACK_EMAIL}/sendMail`, {
						method: "POST",
						headers: { "Authorization": `Bearer ${accessToken}`, "Content-Type": "application/json" },
						body: JSON.stringify(autoReplyPayload)
					});
				}
				
				// Forward the original email to the human (jagrit@zuup.dev)
				await message.forward("jagrit@zuup.dev");
			} catch (e) {
				console.error("Failed to process human fallback routing:", e);
			}
		}

		if (decision.action === "reply" && decision.reply_text) {
			const signature = `\n\n---\nNote: This email has been sent by ZuupMail AI Agent. If you have more questions, please reply back.`;
			const finalReplyText = decision.reply_text + signature;

			const replyPayload = {
				message: {
					subject: `Re: ${parsedEmail.subject || "Your Message"}`,
					body: { contentType: "text", content: finalReplyText },
					toRecipients: [{ emailAddress: { address: sender } }]
				}
			};

			try {
				if (!accessToken) throw new Error("No Graph API token available to send email.");
				
				// 1. Send the autonomous reply via Microsoft Graph
				const sendRes = await fetch(`https://graph.microsoft.com/v1.0/users/${env.HUMAN_FALLBACK_EMAIL}/sendMail`, {
					method: "POST",
					headers: { 
						"Authorization": `Bearer ${accessToken}`,
						"Content-Type": "application/json"
					},
					body: JSON.stringify(replyPayload)
				});
				
				if (!sendRes.ok) {
					console.error("Failed to send auto-reply in Graph:", await sendRes.text());
				} else {
					console.log("Autonomous reply sent via MS Graph successfully!");
				}

				// 2. Forward the original email to jagrit@zuup.dev so they have it
				await message.forward("jagrit@zuup.dev");
			} catch(e) {
				console.error("Failed to send auto-reply:", e);
				// Fallback: just forward it
				await message.forward("jagrit@zuup.dev");
			}
		}

	} catch (error) {
		console.error("Error in email handler:", error);
		// Always forward to human if something completely crashes
		await message.forward("jagrit@zuup.dev");
	}
}
