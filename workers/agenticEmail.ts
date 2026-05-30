import PostalMime from "postal-mime";
import { createMimeMessage } from "mimetext";
import { EmailMessage } from "cloudflare:email";
import type { Env } from "./types";

export async function handleIncomingEmail(message: ForwardableEmailMessage, env: Env, ctx: ExecutionContext) {
	try {
		console.log(`Received email from ${message.from} to ${message.to}`);

		// 1. Parse the incoming email
		const parsedEmail = await PostalMime.parse(message.raw);
		const sender = message.from;

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
			} else {
				console.error("Failed to fetch sent items:", await sentRes.text());
			}
		} else {
			console.error("Failed to authenticate with Graph:", await tokenRes.text());
		}

		// 3. Triage & Draft via Llama 3.1
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

		// 4. Execute Action
		if (decision.action === "ignore") {
			console.log("Ignored as spam/newsletter.");
			return;
		}

		if (decision.action === "human_required") {
			const ticketId = `#ZP-${Math.floor(Math.random() * 100000)}`;
			
			// Auto-reply with ticket ID
			const autoReply = createMimeMessage();
			autoReply.setSender({ name: "Zuup Support", addr: message.to });
			autoReply.setRecipient(message.from);
			autoReply.setSubject(`Re: ${parsedEmail.subject} (Ticket ${ticketId})`);
			autoReply.addMessage({
				contentType: "text/plain",
				data: `Hi there,\n\nWe received your message and a human will look into this shortly.\n\nYour reference token is ${ticketId}.\n\nBest,\nZuup AI Agent`
			});

			try {
				const autoReplyMessage = new EmailMessage(message.to, message.from, autoReply.asRaw());
				await env.SEB.send(autoReplyMessage);
				
				// Forward the original email to the human
				await message.forward(env.HUMAN_FALLBACK_EMAIL);
			} catch (e) {
				console.error("Failed to process human fallback routing:", e);
			}
		}

		if (decision.action === "reply" && decision.reply_text) {
			const draftPayload = {
				subject: `Re: ${parsedEmail.subject || "Your Message"}`,
				body: { contentType: "text", content: decision.reply_text },
				toRecipients: [{ emailAddress: { address: sender } }]
			};

			try {
				if (!accessToken) throw new Error("No Graph API token available to create draft.");
				
				// 1. Create the draft in Microsoft Graph
				const draftRes = await fetch(`https://graph.microsoft.com/v1.0/users/${env.HUMAN_FALLBACK_EMAIL}/messages`, {
					method: "POST",
					headers: { 
						"Authorization": `Bearer ${accessToken}`,
						"Content-Type": "application/json"
					},
					body: JSON.stringify(draftPayload)
				});
				
				if (!draftRes.ok) {
					console.error("Failed to create draft in Graph:", await draftRes.text());
				} else {
					console.log("Autonomous draft created in MS Graph Drafts folder successfully!");
				}

				// 2. Forward the original email to the human inbox so they can see the context of the draft
				await message.forward(env.HUMAN_FALLBACK_EMAIL);
			} catch(e) {
				console.error("Failed to create auto-draft:", e);
				// Fallback: just forward it
				await message.forward(env.HUMAN_FALLBACK_EMAIL);
			}
		}

	} catch (error) {
		console.error("Error in email handler:", error);
		// Always forward to human if something completely crashes
		await message.forward(env.HUMAN_FALLBACK_EMAIL);
	}
}
