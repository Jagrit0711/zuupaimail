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

			const systemPrompt = `You are the Zuup Agent, a stateful UI assistant.
You have access to the user's Microsoft Graph Inbox via tools. If they ask you to search, summarize, or read their emails, YOU MUST USE THE TOOLS PROVIDED. Do not hallucinate tools.

When you are ready to respond to the user (after using tools, or if no tools are needed), you MUST output your final response wrapped in XML tags in this exact format:
<response>
  <type>chat or draft</type>
  <text>Your conversational response OR the drafted email body</text>
</response>
Do not use JSON. Use the exact XML format above.`;

			let messages: any[] = [
				{ role: "system", content: systemPrompt },
				...history,
				{ role: "user", content: `Current email context (if any):\n${emailText || "None"}\n\nUser Instruction: ${userPrompt}` }
			];

			const tools = [
				{
					name: "get_recent_emails",
					description: "Fetch the user's most recent emails from their inbox.",
					parameters: { type: "object", properties: {} }
				},
				{
					name: "get_unread_emails",
					description: "Fetch the user's unread emails from their inbox.",
					parameters: { type: "object", properties: {} }
				},
				{
					name: "search_emails",
					description: "Search the user's inbox for a specific query.",
					parameters: {
						type: "object",
						properties: {
							query: { type: "string", description: "The search query (e.g. 'from:john' or 'meeting')" }
						},
						required: ["query"]
					}
				}
			];

			let finalResponseText = "";
			let finalResponseType = "chat";

			try {
				let aiResponse = await this.env.AI.run("@cf/meta/llama-3.1-8b-instruct", {
					messages,
					// @ts-ignore
					tools
				});

				// Handle tool calls
				// @ts-ignore
				let toolCall = null;
				// @ts-ignore
				if (aiResponse.tool_calls && aiResponse.tool_calls.length > 0) {
					// @ts-ignore
					toolCall = aiResponse.tool_calls[0];
				} else {
					// Fallback for when Llama 3.1 hallucinates tool calls as raw text
					// @ts-ignore
					const rawRes = aiResponse.response || "";
					const match = rawRes.match(/\{"name":\s*"([^"]+)",\s*"arguments":\s*(\{.*?\})\}/);
					if (match) {
						try {
							toolCall = { name: match[1], arguments: JSON.parse(match[2]) };
						} catch (e) {}
					}
				}

				if (toolCall) {
					let toolResult = "No data found.";
					
					if (graphToken) {
						try {
							if (toolCall.name === "get_recent_emails") {
								const res = await fetch(`https://graph.microsoft.com/v1.0/me/messages?$top=20&$select=subject,from,body,receivedDateTime&$orderby=receivedDateTime DESC`, {
									headers: { 
										"Authorization": graphToken,
										"Prefer": 'outlook.body-content-type="text"'
									}
								});
								if (res.ok) {
									const data = await res.json() as any;
									toolResult = JSON.stringify(data.value.map((m: any) => ({ 
										from: m.from?.emailAddress?.address, 
										subject: m.subject, 
										date: m.receivedDateTime,
										content: m.body?.content ? m.body.content.substring(0, 800) : ""
									})));
								}
							} else if (toolCall.name === "get_unread_emails") {
								const res = await fetch(`https://graph.microsoft.com/v1.0/me/messages?$filter=isRead eq false&$top=15&$select=subject,from,body,receivedDateTime`, {
									headers: { 
										"Authorization": graphToken,
										"Prefer": 'outlook.body-content-type="text"'
									}
								});
								if (res.ok) {
									const data = await res.json() as any;
									toolResult = JSON.stringify(data.value.map((m: any) => ({ 
										from: m.from?.emailAddress?.address, 
										subject: m.subject, 
										date: m.receivedDateTime,
										content: m.body?.content ? m.body.content.substring(0, 800) : ""
									})));
								}
							} else if (toolCall.name === "search_emails") {
								const res = await fetch(`https://graph.microsoft.com/v1.0/me/messages?$search="${toolCall.arguments.query}"&$top=15&$select=subject,from,body,receivedDateTime`, {
									headers: { 
										"Authorization": graphToken,
										"Prefer": 'outlook.body-content-type="text"'
									}
								});
								if (res.ok) {
									const data = await res.json() as any;
									toolResult = JSON.stringify(data.value.map((m: any) => ({ 
										from: m.from?.emailAddress?.address, 
										subject: m.subject, 
										date: m.receivedDateTime,
										content: m.body?.content ? m.body.content.substring(0, 800) : ""
									})));
								}
							}
						} catch (e) {
							toolResult = `Tool execution failed: ${e}`;
						}
					} else {
						toolResult = "Missing Microsoft Graph token, cannot execute tool.";
					}

					// Append tool result and run again using standard roles to avoid Cloudflare AI schema strictness
					messages.push({ role: "assistant", content: `I need to use the ${toolCall.name} tool.` });
					messages.push({ role: "user", content: `The tool '${toolCall.name}' returned the following data:\n${toolResult}\n\nNow, fulfill the user's request. If the user asked for a summary, YOU MUST provide a highly detailed, comprehensive summary. Extract specific names, dates, amounts, and actionable items from the data. Do not just give a high-level overview. YOU MUST output your final response wrapped in XML in this format:\n<response>\n  <type>chat</type>\n  <text>your detailed response here</text>\n</response>` });

					aiResponse = await this.env.AI.run("@cf/meta/llama-3.1-8b-instruct", { messages });
				}

				const rawText = aiResponse.response || "";
				let parsed = { type: "chat", text: rawText };
				
				// Try parsing XML first
				const typeMatch = rawText.match(/<type>\s*([\s\S]*?)\s*<\/type>/i);
				if (typeMatch) parsed.type = typeMatch[1].trim();
				
				const textMatch = rawText.match(/<text>\s*([\s\S]*?)\s*<\/text>/i);
				if (textMatch) {
					parsed.text = textMatch[1].trim();
				} else {
					// Fallback for JSON if it ignored the XML instruction
					const jsonMatch = rawText.match(/\{[\s\S]*\}/);
					if (jsonMatch) {
						try {
							// Try native parse first
							const temp = JSON.parse(jsonMatch[0]);
							if (temp.type) parsed.type = temp.type;
							if (temp.text || temp.content) parsed.text = temp.text || temp.content;
						} catch (e) {
							// Try regex extraction as last resort
							const fallbackText = rawText.match(/"(?:text|content)"\s*:\s*"([\s\S]*?)"\s*\}/);
							if (fallbackText) parsed.text = fallbackText[1].replace(/\\n/g, "\n").replace(/\\"/g, '"');
						}
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
