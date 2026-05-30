// Copyright (c) 2026 Cloudflare, Inc.
// Licensed under the Apache 2.0 license found in the LICENSE file or at:
//     https://opensource.org/licenses/Apache-2.0

import { Badge, Button, Loader, Tooltip, useKumoToastManager } from "@cloudflare/kumo";
import { RobotIcon, ArrowUpIcon, StopIcon, TrashIcon, PaperPlaneTiltIcon, PencilSimpleIcon, UserIcon, CopyIcon } from "@phosphor-icons/react";
import { useState, useRef, useEffect } from "react";
import { useParams } from "react-router";
import Markdown from "react-markdown";
import remarkGfm from "remark-gfm";
import { useUIStore } from "~/hooks/useUIStore";
import { useEmail, useReplyToEmail } from "~/queries/emails";
import api from "~/services/api";

export default function AgentPanel() {
	const { mailboxId } = useParams<{ mailboxId: string }>();
	const { selectedEmailId, startCompose } = useUIStore();
	// @ts-ignore
	const { data: currentEmail } = useEmail(mailboxId, selectedEmailId);
	const replyMut = useReplyToEmail();
	const toastManager = useKumoToastManager();

	const [messages, setMessages] = useState<{ role: "user" | "ai", text: string, isDraft?: boolean }[]>([]);
	const [inputValue, setInputValue] = useState("");
	const [isLoading, setIsLoading] = useState(false);
	const scrollRef = useRef<HTMLDivElement>(null);
	const inputRef = useRef<HTMLTextAreaElement>(null);

	// Fetch history on mount
	useEffect(() => {
		api.aiHistory().then(history => {
			if (history && history.length > 0) {
				const formattedHistory = history.map((msg: any) => ({
					role: msg.role === "user" ? "user" : "ai",
					text: msg.content,
					isDraft: msg.type === "draft"
				}));
				setMessages(formattedHistory as any);
			}
		}).catch(e => console.error("Failed to load history:", e));
	}, []);

	useEffect(() => {
		const el = scrollRef.current;
		if (el) el.scrollTop = el.scrollHeight;
	}, [messages, isLoading]);

	const handleSend = async () => {
		const text = inputValue.trim();
		if (!text || isLoading) return;

		setInputValue("");
		setMessages(prev => [...prev, { role: "user", text }]);
		setIsLoading(true);

		// Prepare context if an email is open
		let emailText = "";
		if (currentEmail) {
			const cleanBody = currentEmail.body ? currentEmail.body.replace(/<[^>]*>?/gm, '') : "";
			emailText = `From: ${currentEmail.sender}\nSubject: ${currentEmail.subject}\nBody: ${cleanBody}`;
		}

		try {
			const data = await api.aiChat({ emailText, userPrompt: text });
			const replyType = data.type || "chat";
			const replyText = data.text || "No response generated.";
			const isDraft = replyType === "draft";
			setMessages(prev => [...prev, { role: "ai", text: replyText, isDraft }]);
		} catch (e) {
			setMessages(prev => [...prev, { role: "ai", text: "Network error." }]);
		} finally {
			setIsLoading(false);
		}
	};

	const handleQuickSend = async (text: string) => {
		if (!currentEmail || !mailboxId) {
			toastManager.add({ title: "No email selected to reply to.", variant: "error" });
			return;
		}
		toastManager.add({ title: "Sending reply..." });
		try {
			await replyMut.mutateAsync({
				mailboxId,
				emailId: currentEmail.id,
				email: {
					html: `<p>${text.replace(/\n/g, "<br>")}</p>`,
					text: text,
					to: currentEmail.sender
				}
			});
			toastManager.add({ title: "Reply sent successfully!" });
		} catch (e) {
			toastManager.add({ title: "Failed to send reply.", variant: "error" });
		}
	};

	const handleEditDraft = (text: string) => {
		if (!currentEmail) {
			toastManager.add({ title: "No email selected to reply to.", variant: "error" });
			return;
		}
		startCompose({
			mode: "reply",
			originalEmail: currentEmail,
			draftEmail: { 
				recipient: currentEmail.sender,
				subject: currentEmail.subject.startsWith("Re:") ? currentEmail.subject : `Re: ${currentEmail.subject}`,
				body: `<p>${text.replace(/\n/g, "<br>")}</p>`,
			} as any
		});
	};

	return (
		<div className="flex flex-col h-full">
			{/* Header */}
			<div className="flex items-center justify-between px-3 py-1.5 border-b border-kumo-line shrink-0">
				<div className="flex items-center gap-2">
					<Badge variant="beta">AI</Badge>
					<span className="text-xs text-kumo-subtle">Zuup Agent</span>
				</div>
				{messages.length > 0 && (
					<Tooltip content="Clear chat" asChild>
						<Button
							variant="ghost"
							shape="square"
							size="sm"
							icon={<TrashIcon size={14} />}
							onClick={() => setMessages([])}
						/>
					</Tooltip>
				)}
			</div>

			{/* Messages */}
			<div ref={scrollRef} className="flex-1 overflow-y-auto px-3 py-4 flex flex-col gap-3">
				{messages.length === 0 ? (
					<div className="flex flex-col items-center justify-center h-full gap-4 opacity-70">
						<RobotIcon size={32} weight="duotone" className="text-kumo-brand" />
						<p className="text-xs text-center px-4">
							{currentEmail ? "I am reading the currently opened email. Tell me how to reply!" : "I am your stateless AI assistant. Open an email to get started!"}
						</p>
					</div>
				) : (
					messages.map((msg, i) => (
						<div key={i} className={`flex flex-col gap-1.5 ${msg.role === "user" ? "items-end" : "items-start"}`}>
							<div className={`flex gap-2 ${msg.role === "user" ? "flex-row-reverse" : ""}`}>
								<div className={`flex h-6 w-6 shrink-0 items-center justify-center rounded-full ${msg.role === "user" ? "bg-kumo-brand text-white" : "bg-kumo-surface border border-kumo-border"}`}>
									{msg.role === "user" ? <UserIcon size={12} /> : <RobotIcon size={12} />}
								</div>
								<div className={`rounded-lg px-3 py-2 text-[13px] leading-relaxed max-w-[85%] ${msg.role === "user" ? "bg-kumo-brand text-white rounded-br-sm" : "bg-kumo-surface border border-kumo-border rounded-bl-sm"}`}>
									<Markdown remarkPlugins={[remarkGfm]}>{msg.text}</Markdown>
								</div>
							</div>
							
							{/* Action Buttons for AI Drafts */}
							{msg.role === "ai" && (
								<div className="flex items-center gap-1.5 pl-8 mt-1">
									{msg.isDraft && (
										<>
											<Button 
												variant="primary" 
												size="sm" 
												icon={<PaperPlaneTiltIcon size={12} />}
												onClick={() => handleQuickSend(msg.text)}
												disabled={replyMut.isPending}
											>
												Send
											</Button>
											<Button 
												variant="secondary" 
												size="sm" 
												icon={<PencilSimpleIcon size={12} />}
												onClick={() => handleEditDraft(msg.text)}
												disabled={replyMut.isPending}
											>
												Edit
											</Button>
										</>
									)}
									<Button 
										variant="ghost" 
										size="sm" 
										icon={<CopyIcon size={12} />}
										onClick={() => {
											navigator.clipboard.writeText(msg.text);
											toastManager.add({ title: "Copied to clipboard!" });
										}}
									>
										Copy
									</Button>
								</div>
							)}
						</div>
					))
				)}
				{isLoading && (
					<div className="flex gap-2">
						<div className="flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-kumo-surface border border-kumo-border">
							<RobotIcon size={12} />
						</div>
						<div className="flex items-center gap-1.5 px-3 py-2 rounded-lg bg-kumo-surface border border-kumo-border rounded-bl-sm">
							<Loader size="sm" />
							<span className="text-xs">Working on it...</span>
						</div>
					</div>
				)}
			</div>

			{/* Input */}
			<div className="shrink-0 border-t border-kumo-line px-3 py-2">
				<div className="flex items-end gap-1.5">
					<textarea
						ref={inputRef}
						value={inputValue}
						onChange={(e) => setInputValue(e.target.value)}
						onKeyDown={(e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleSend(); } }}
						placeholder={currentEmail ? "E.g. Tell them I will do it..." : "Open an email first..."}
						rows={1}
						className="flex-1 resize-none rounded-lg border border-kumo-line bg-kumo-control px-3 py-2 text-xs focus:outline-none focus:ring-1 focus:ring-kumo-brand min-h-[36px] max-h-[100px] text-kumo-default"
					/>
					<Button
						variant="primary"
						shape="square"
						size="sm"
						disabled={!inputValue.trim() || isLoading}
						icon={<ArrowUpIcon size={14} weight="bold" />}
						onClick={handleSend}
					/>
				</div>
			</div>
		</div>
	);
}
