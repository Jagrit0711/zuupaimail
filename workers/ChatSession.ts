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
You have access to the user's Microsoft Graph Inbox natively. Please search their inbox if needed.
When you are ready to respond to the user, you MUST output your final response wrapped in XML tags in this exact format:
<response>
  <type>chat or draft</type>
  <text>Your conversational response OR the drafted email body</text>
</response>
Do not use JSON. Use the exact XML format above.`;

			let finalResponseText = "";
			let finalResponseType = "chat";

			if (!graphToken) {
				finalResponseText = "Missing Microsoft Graph token, cannot contact Copilot.";
			} else {
				try {
					// 1. Create Copilot Conversation
					const convRes = await fetch("https://graph.microsoft.com/beta/copilot/conversations", {
						method: "POST",
						headers: { "Authorization": graphToken, "Content-Type": "application/json" },
						body: JSON.stringify({})
					});
					
					if (!convRes.ok) {
						throw new Error(`Failed to create Copilot conversation: ${await convRes.text()}`);
					}
					
					const convData = await convRes.json() as { id: string };
					const convId = convData.id;

					// Format history into the prompt
					const historyText = history.map((h: any) => `${h.role.toUpperCase()}: ${h.content}`).join("\n\n");
					
					const fullPrompt = `${systemPrompt}\n\nChat History:\n${historyText}\n\nCurrent email context (if any):\n${emailText || "None"}\n\nUser Instruction: ${userPrompt}`;

					// 2. Send Chat Prompt
					const chatRes = await fetch(`https://graph.microsoft.com/beta/copilot/conversations/${convId}/chat`, {
						method: "POST",
						headers: { "Authorization": graphToken, "Content-Type": "application/json" },
						body: JSON.stringify({
							message: {
								text: fullPrompt
							}
						})
					});

					if (!chatRes.ok) {
						throw new Error(`Failed to execute Copilot chat: ${await chatRes.text()}`);
					}

					const chatData = await chatRes.json() as any;
					const msgs = chatData.messages || [];
					const copilotMsg = msgs.length > 0 ? msgs[msgs.length - 1] : null;
					
					const rawText = (copilotMsg && copilotMsg.text) ? copilotMsg.text : "No response from Copilot.";
					
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
								const temp = JSON.parse(jsonMatch[0]);
								if (temp.type) parsed.type = temp.type;
								if (temp.text || temp.content) parsed.text = temp.text || temp.content;
							} catch (e) {
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
			}

			// Store AI message
			this.sql.exec("INSERT INTO messages (role, content, type) VALUES (?, ?, ?)", "ai", finalResponseText, finalResponseType);

			return Response.json({ type: finalResponseType, text: finalResponseText });
		}

		return new Response("Not found", { status: 404 });
	}
}
