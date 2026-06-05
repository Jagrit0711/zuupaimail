// Copyright (c) 2026 Cloudflare, Inc.
// Licensed under the Apache 2.0 license found in the LICENSE file or at:
//     https://opensource.org/licenses/Apache-2.0

import { Badge, Button, Input, Loader, useKumoToastManager } from "@cloudflare/kumo";
import { BrainIcon, RobotIcon, ArrowCounterClockwiseIcon, NotepadIcon } from "@phosphor-icons/react";
import { useEffect, useState } from "react";
import { useParams } from "react-router";
import { useMailbox, useUpdateMailbox } from "~/queries/mailboxes";

// Placeholder shown in the textarea when no custom prompt is set.
// The authoritative default prompt lives in workers/agent/index.ts (DEFAULT_SYSTEM_PROMPT).
const PROMPT_PLACEHOLDER = `You are an email assistant that helps manage this inbox. You read emails, draft replies, and help organize conversations.\n\nWrite like a real person. Short, direct, flowing prose. Plain text only.\n\n(Leave empty to use the full built-in default prompt)`;

export default function SettingsRoute() {
	const { mailboxId } = useParams<{ mailboxId: string }>();
	const toastManager = useKumoToastManager();
	const { data: mailbox } = useMailbox(mailboxId);
	const updateMailboxMutation = useUpdateMailbox();

	const [displayName, setDisplayName] = useState("");
	const [agentPrompt, setAgentPrompt] = useState("");
	const [agentKnowledgeBase, setAgentKnowledgeBase] = useState("");
	const [agentMailboxPurpose, setAgentMailboxPurpose] = useState("");
	const [agentAutoReplyEnabled, setAgentAutoReplyEnabled] = useState(false);
	const [isSaving, setIsSaving] = useState(false);

	useEffect(() => {
		if (mailbox) {
			setDisplayName(mailbox.settings?.fromName || mailbox.name || "");
			setAgentPrompt(mailbox.settings?.agentSystemPrompt || "");
			setAgentKnowledgeBase(mailbox.settings?.agentKnowledgeBase || "");
			setAgentMailboxPurpose(mailbox.settings?.agentMailboxPurpose || "");
			setAgentAutoReplyEnabled(mailbox.settings?.agentAutoReplyEnabled || false);
		}
	}, [mailbox]);

	const handleSave = async () => {
		if (!mailbox || !mailboxId) return;
		setIsSaving(true);
		const settings = {
			...mailbox.settings,
			fromName: displayName,
			agentSystemPrompt: agentPrompt.trim() || undefined,
			agentKnowledgeBase: agentKnowledgeBase.trim() || undefined,
			agentMailboxPurpose: agentMailboxPurpose.trim() || undefined,
			agentAutoReplyEnabled,
		};
		try {
			await updateMailboxMutation.mutateAsync({ mailboxId, settings });
			toastManager.add({ title: "Settings saved!" });
		} catch {
			toastManager.add({
				title: "Failed to save settings",
				variant: "error",
			});
		} finally {
			setIsSaving(false);
		}
	};

	const handleResetPrompt = () => {
		setAgentPrompt("");
	};

	if (!mailbox) {
		return (
			<div className="flex justify-center py-20">
				<Loader size="lg" />
			</div>
		);
	}

	const isCustomPrompt = agentPrompt.trim().length > 0;

	return (
		<div className="max-w-2xl px-4 py-4 md:px-8 md:py-6 h-full overflow-y-auto">
			<h1 className="text-lg font-semibold text-kumo-default mb-6">Settings</h1>

			<div className="space-y-6">
				{/* Account */}
				<div className="rounded-lg border border-kumo-line bg-kumo-base p-5">
					<div className="text-sm font-medium text-kumo-default mb-4">
						Account
					</div>
					<div className="space-y-3">
						<Input
							label="Display Name"
							value={displayName}
							onChange={(e) => setDisplayName(e.target.value)}
						/>
						<Input label="Email" type="email" value={mailbox.email} disabled />
					</div>
				</div>

				{/* Agent System Prompt */}
				<div className="rounded-lg border border-kumo-line bg-kumo-base p-5">
					<div className="flex items-center justify-between mb-4">
						<div className="flex items-center gap-2">
							<RobotIcon size={16} weight="duotone" className="text-kumo-subtle" />
							<span className="text-sm font-medium text-kumo-default">
								AI Agent Prompt
							</span>
							{isCustomPrompt ? (
								<Badge variant="primary">Custom</Badge>
							) : (
								<Badge variant="secondary">Default</Badge>
							)}
						</div>
						{isCustomPrompt && (
							<Button
								variant="ghost"
								size="xs"
								icon={<ArrowCounterClockwiseIcon size={14} />}
								onClick={handleResetPrompt}
							>
								Reset to default
							</Button>
						)}
					</div>
					<p className="text-xs text-kumo-subtle mb-3">
						Customize how the AI agent behaves for this mailbox.
						Leave empty to use the built-in default prompt.
					</p>
					<textarea
						value={agentPrompt}
						onChange={(e) => setAgentPrompt(e.target.value)}
						placeholder={PROMPT_PLACEHOLDER}
						rows={12}
						className="w-full resize-y rounded-lg border border-kumo-line bg-kumo-recessed px-3 py-2 text-xs text-kumo-default placeholder:text-kumo-subtle focus:outline-none focus:ring-1 focus:ring-kumo-ring font-mono leading-relaxed"
					/>
					<p className="text-xs text-kumo-subtle mt-2">
						The prompt is sent as the system message to the AI model.
						It controls the agent's personality, writing style, and behavior rules.
					</p>
				</div>

				{/* AI Knowledge Base */}
				<div className="rounded-lg border border-kumo-line bg-kumo-base p-5">
					<div className="flex items-center gap-2 mb-4">
						<NotepadIcon size={16} weight="duotone" className="text-kumo-subtle" />
						<span className="text-sm font-medium text-kumo-default">
							AI Knowledge Base
						</span>
					</div>
					<p className="text-xs text-kumo-subtle mb-3">
						Add custom facts, FAQs, company info, or any text the cron AI agent should know when auto-replying for this specific mailbox. This is used by the background agent that processes emails autonomously.
					</p>
					<textarea
						value={agentKnowledgeBase}
						onChange={(e) => setAgentKnowledgeBase(e.target.value)}
						placeholder={`Example:\n- hello@zuup.dev handles general inquiries about Zuup products\n- FAR AWAY hackathon registration deadline is July 1st 2026\n- Refunds take 5–7 business days\n- For urgent issues, escalate to jagrit@zuup.dev`}
						rows={10}
						className="w-full resize-y rounded-lg border border-kumo-line bg-kumo-recessed px-3 py-2 text-xs text-kumo-default placeholder:text-kumo-subtle focus:outline-none focus:ring-1 focus:ring-kumo-ring font-mono leading-relaxed"
					/>
				</div>

				{/* Mailbox Purpose */}
				<div className="rounded-lg border border-kumo-line bg-kumo-base p-5">
					<div className="flex items-center gap-2 mb-4">
						<BrainIcon size={16} weight="duotone" className="text-kumo-subtle" />
						<span className="text-sm font-medium text-kumo-default">
							Mailbox Purpose / Role
						</span>
					</div>
					<p className="text-xs text-kumo-subtle mb-3">
						Describe what this mailbox is for and who typically reads it. This tells the AI agent the context and tone for auto-replies. Example: <em>"Handle sponsorship inquiries for FAR AWAY hackathon. Respond formally."</em>
					</p>
					<textarea
						value={agentMailboxPurpose}
						onChange={(e) => setAgentMailboxPurpose(e.target.value)}
						placeholder="e.g. This mailbox handles all general inquiries from Zuup community members. Replies should be warm, concise, and helpful. Jagrit reviews flagged emails."
						rows={4}
						className="w-full resize-y rounded-lg border border-kumo-line bg-kumo-recessed px-3 py-2 text-xs text-kumo-default placeholder:text-kumo-subtle focus:outline-none focus:ring-1 focus:ring-kumo-ring font-mono leading-relaxed"
					/>
				</div>

				{/* Agentic Auto-Reply */}
				<div className="rounded-lg border border-kumo-line bg-kumo-base p-5">
					<div className="flex items-center justify-between mb-2">
						<div className="flex items-center gap-2">
							<RobotIcon size={16} weight="duotone" className="text-kumo-primary" />
							<span className="text-sm font-medium text-kumo-default">
								Agentic Auto-Reply
							</span>
						</div>
						<div className="flex items-center">
							<input 
								type="checkbox" 
								className="toggle toggle-primary" 
								checked={agentAutoReplyEnabled}
								onChange={(e) => setAgentAutoReplyEnabled(e.target.checked)}
							/>
						</div>
					</div>
					<p className="text-xs text-kumo-subtle mb-3">
						When enabled, the AI will automatically read your incoming emails, learn from your past sent emails, and reply automatically. 
						Complex questions will be forwarded to your fallback email address for manual review.
					</p>
				</div>

				{/* Save */}
				<div className="flex justify-end">
					<Button variant="primary" onClick={handleSave} loading={isSaving}>
						Save Changes
					</Button>
				</div>
			</div>
		</div>
	);
}
