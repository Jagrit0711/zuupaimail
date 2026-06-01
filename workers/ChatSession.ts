import { DurableObject } from "cloudflare:workers";
import type { Env } from "./types";

export class ChatSession extends DurableObject<Env> {
	sql: SqlStorage;

	constructor(ctx: DurableObjectState, env: Env) {
		super(ctx, env);
		// @ts-ignore - TS types for new SQLite DOs might not be fully up to date in the environment
		this.sql = ctx.storage.sql;
		
		this.sql.exec(`
			CREATE TABLE IF NOT EXISTS messages (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				role TEXT NOT NULL,
				content TEXT NOT NULL,
				type TEXT NOT NULL,
				timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
			);
			CREATE TABLE IF NOT EXISTS settings (
				key TEXT PRIMARY KEY,
				value TEXT NOT NULL
			);
		`);
	}

	async fetch(request: Request): Promise<Response> {
		const url = new URL(request.url);

		if (url.pathname.endsWith("/settings")) {
			if (request.method === "GET") {
				const cursor = this.sql.exec("SELECT key, value FROM settings");
				const settings = [...cursor].reduce((acc: any, row: any) => {
					acc[row.key] = JSON.parse(row.value);
					return acc;
				}, {});
				return Response.json(settings);
			}
			
			if (request.method === "POST") {
				const body = await request.json() as any;
				for (const [key, value] of Object.entries(body)) {
					this.sql.exec("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", key, JSON.stringify(value));
				}
				return Response.json({ success: true });
			}
		}

		if (request.method === "GET" && url.pathname.endsWith("/history")) {
			const cursor = this.sql.exec("SELECT role, content, type FROM messages ORDER BY timestamp ASC");
			return Response.json([...cursor]);
		}

		if (request.method === "POST" && url.pathname.endsWith("/chat")) {
			const body = await request.json() as any;
			const { userPrompt, emailText } = body;
			const graphToken = request.headers.get("Authorization");

			// Store user message
			this.sql.exec("INSERT INTO messages (role, content, type) VALUES (?, ?, ?)", "user", userPrompt, "chat");

			// Fetch history for AI context
			const historyCursor = this.sql.exec("SELECT role, content FROM messages ORDER BY timestamp ASC LIMIT 10");
			const history = [...historyCursor].map((r: any) => ({ 
				role: r.role === "ai" ? "assistant" : r.role, 
				content: r.content 
			}));

			const systemPrompt = `You are the Zuup Agent, an autonomous UI assistant.
You have access to the user's Microsoft Graph Inbox natively. You can search their inbox and send emails on their behalf.
Available Tools:
1. <tool><name>search_emails</name><query>your search query</query></tool>
2. <tool><name>send_email</name><to>recipient email</to><subject>email subject</subject><body>html or text body</body></tool>

If you need to use a tool, output ONLY the tool XML block. Wait for the user to provide the <tool_result> before continuing.
If you are finished and ready to respond to the user, you MUST output your final response wrapped in XML tags in this exact format:
<response>
  <type>chat or draft</type>
  <text>Your conversational response OR the drafted email body</text>
</response>
Do not use JSON. Use the exact XML format above.`;

			let finalResponseText = "";
			let finalResponseType = "chat";

			try {
				let messages: any[] = [
					{ role: "system", content: systemPrompt },
					...history,
					{ role: "user", content: `Current email context (if any):\n${emailText || "None"}\n\nUser Instruction: ${userPrompt}` }
				];

				let iterations = 0;
				let isDone = false;
				let rawText = "";

				while (iterations < 4 && !isDone) {
					iterations++;
					
					const aiRes = await this.env.AI.run("@cf/meta/llama-3.3-70b-instruct-fp8-fast", {
						messages
					}) as { response: string };

					rawText = aiRes.response || "No response generated.";
					messages.push({ role: "assistant", content: rawText });
					
					console.log(`[AI Iteration ${iterations}]`, rawText);

					// Check for Tool Calls
					const toolMatch = rawText.match(/<tool>([\s\S]*?)<\/tool>/i);
					if (toolMatch && graphToken) {
						const toolBody = toolMatch[1];
						const nameMatch = toolBody.match(/<name>(.*?)<\/name>/i);
						const toolName = nameMatch ? nameMatch[1].trim() : "";
						
						if (toolName === "search_emails") {
							const queryMatch = toolBody.match(/<query>(.*?)<\/query>/i);
							const query = queryMatch ? queryMatch[1].trim() : "";
							
							const searchRes = await fetch(`https://graph.microsoft.com/v1.0/me/messages?$search="${query}"&$select=sender,toRecipients,subject,bodyPreview&$top=10`, {
								headers: { "Authorization": graphToken }
							});
							const searchData = await searchRes.json() as any;
							
							const results = (searchData.value || []).map((m: any) => 
								`Subject: ${m.subject}\nSender: ${m.sender?.emailAddress?.address}\nTo: ${m.toRecipients?.map((r:any)=>r.emailAddress?.address).join(", ")}\nPreview: ${m.bodyPreview}`
							).join("\n\n---\n\n");
							
							messages.push({ role: "user", content: `<tool_result>${results || "No emails found."}</tool_result>` });
							continue;
						}
						
						if (toolName === "send_email") {
							const toMatch = toolBody.match(/<to>(.*?)<\/to>/i);
							const subjMatch = toolBody.match(/<subject>(.*?)<\/subject>/i);
							const bodyMatch = toolBody.match(/<body>([\s\S]*?)<\/body>/i);
							
							const to = toMatch ? toMatch[1].trim() : "";
							const subject = subjMatch ? subjMatch[1].trim() : "";
							const bodyContent = bodyMatch ? bodyMatch[1].trim() : "";
							
							const sendPayload = {
								message: {
									subject,
									body: { contentType: "HTML", content: bodyContent },
									toRecipients: to.split(",").map(t => ({ emailAddress: { address: t.trim() } }))
								}
							};

							const sendRes = await fetch(`https://graph.microsoft.com/v1.0/me/sendMail`, {
								method: "POST",
								headers: { "Authorization": graphToken, "Content-Type": "application/json" },
								body: JSON.stringify(sendPayload)
							});
							
							if (sendRes.ok) {
								messages.push({ role: "user", content: `<tool_result>Email successfully sent!</tool_result>` });
							} else {
								messages.push({ role: "user", content: `<tool_result>Failed to send email: ${await sendRes.text()}</tool_result>` });
							}
							continue;
						}
					}
					
					isDone = true;
				}
				
				let parsed = { type: "chat", text: rawText };
				
				// Try parsing XML first
				const typeMatch = rawText.match(/<type>\s*([\s\S]*?)\s*<\/type>/i);
				if (typeMatch) parsed.type = typeMatch[1].trim();
				
				const textMatch = rawText.match(/<text>\s*([\s\S]*?)\s*<\/text>/i);
				if (textMatch) {
					parsed.text = textMatch[1].trim();
				} else {
					const jsonMatch = rawText.match(/\{[\s\S]*\}/);
					if (jsonMatch) {
						try {
							const temp = JSON.parse(jsonMatch[0]);
							if (temp.type) parsed.type = temp.type;
							if (temp.text || temp.content) parsed.text = temp.text || temp.content;
						} catch (e) {}
					}
				}

				finalResponseText = parsed.text || rawText || "No response generated.";
				finalResponseType = parsed.type === "draft" ? "draft" : "chat";

			} catch (e) {
				console.error("AI execution failed:", e);
				finalResponseText = `An error occurred: ${e}`;
			}

			// Store AI message
			this.sql.exec("INSERT INTO messages (role, content, type) VALUES (?, ?, ?)", "ai", finalResponseText, finalResponseType);

			return Response.json({ type: finalResponseType, text: finalResponseText });
		}

		return new Response("Not found", { status: 404 });
	}
}
