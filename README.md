<div align="center">
  <img src="https://raw.githubusercontent.com/Jagrit0711/zuup-main/bc25cc6dafa9026827ffffa84f5d6740d86950ab/public/lovable-uploads/b44b8051-6117-4b37-999d-014c4c33dd13.png" alt="Zuup Logo" width="120" height="120">
  <br>
  <h1>Zuup Agentic Inbox</h1>
  <p><strong>Your Inbox, Automated.</strong></p>
  <br>
</div>

Welcome to **Zuup Agentic Inbox**! This isn't just an email client; it's a fully autonomous, stateful AI agent that lives inside your inbox. Built entirely on Cloudflare Workers, Durable Objects, and Llama 3.1, Zuup Mail learns how you write, triages incoming emails, and auto-replies on your behalf.

> Built for founders, by founders who are tired of writing the same "Here's the Slack link!" email for the 400th time.

---

## ✨ Features

- **🧠 Autonomous Auto-Reply**: The AI Agent dynamically scrapes your past Sent Items to learn your exact tone, writing style, and context. If someone asks a question you've answered before, the agent just handles it.
- **🛡️ Human Fallback Escalation**: If an email is too complex, involves a negotiation, or requires sensitive info, the AI generates a Support Ticket ID and forwards the thread to your fallback email address.
- **💬 Conversational UI Panel**: Chat with your inbox! Ask "Summarize my unread emails" or "Find emails from John" and watch the agent securely query the Microsoft Graph API using native tool-calling.
- **💾 100% Persistent State**: No expensive vector databases here. Your chat history and global agent settings are stored on the edge using Cloudflare Durable Objects + Embedded SQLite.
- **🚀 Edge-Native Architecture**: No cold starts, no massive Node.js servers. Everything runs on Cloudflare Workers and Cloudflare Email Routing bindings.

## 🛠️ Architecture Deep Dive

Zuup Mail is powered by three core pillars:
1. **The Edge Worker** (`index.ts`): The React Router SSR server and REST API bridge.
2. **The Stateful Brain** (`ChatSession.ts`): A Cloudflare Durable Object containing an embedded SQLite database that persists all chat history and global user settings. It handles all tool-dispatching for the AI.
3. **The Email Router** (`agenticEmail.ts`): A background worker triggered directly by Cloudflare Email Routing. When an email hits your domain, it wakes up, fetches your MS Graph context, runs a Llama 3.1 inference, and either sends an auto-reply or forwards it to you.

## 🚀 Getting Started

Want this for yourself? Fork it, deploy it, and never write a repetitive email again.

### Prerequisites
- A Cloudflare account (with Workers AI and Email Routing enabled)
- A Microsoft Azure App Registration (for Microsoft Graph API access)
- Node.js (v18+)

### Installation

1. **Clone the repo**
   ```bash
   git clone https://github.com/Jagrit0711/zuupaimail.git
   cd zuupaimail
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Configure Environment Variables**
   Create a `.dev.vars` file in the root directory:
   ```env
   AZURE_CLIENT_ID=your_client_id
   AZURE_CLIENT_SECRET=your_client_secret
   AZURE_TENANT_ID=your_tenant_id
   HUMAN_FALLBACK_EMAIL=you@yourdomain.com
   ```

4. **Local Development**
   ```bash
   npm run dev
   ```

5. **Deploy to Edge**
   ```bash
   npx wrangler deploy
   ```

> **Note**: Cloudflare Email Routing only triggers workers that are deployed to the edge. To test the autonomous auto-reply feature, you must deploy your worker to production!

## 🧑‍💻 Contributing

Pull requests are welcome! If you want to add support for Google Workspace or anthropic models, feel free to open an issue or PR.

## 📄 License

Apache 2.0. Go build cool things. 🚀
