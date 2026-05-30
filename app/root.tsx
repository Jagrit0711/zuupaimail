// Copyright (c) 2026 Cloudflare, Inc.
// Licensed under the Apache 2.0 license found in the LICENSE file or at:
//     https://opensource.org/licenses/Apache-2.0

import {
	Button,
	Empty,
	LinkProvider,
	Loader,
	Toasty,
	TooltipProvider,
} from "@cloudflare/kumo";
import { WarningIcon } from "@phosphor-icons/react";
import { MutationCache, QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { forwardRef, useState, useEffect, useState as reactUseState } from "react";
import { MsalProvider, AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from "@azure/msal-react";
import { msalInstance, loginRequest } from "~/lib/authConfig";
import {
	isRouteErrorResponse,
	Links,
	Meta,
	Outlet,
	Link as RouterLink,
	Scripts,
	ScrollRestoration,
} from "react-router";
import { ApiError } from "~/services/api";
import "./index.css";

function makeQueryClient() {
	return new QueryClient({
		defaultOptions: {
			queries: {
				staleTime: 30_000,
				refetchOnWindowFocus: false,
				retry: (failureCount, error) => {
					// Don't retry 4xx errors (not found, unauthorized, etc.)
					if (error instanceof ApiError && error.status >= 400 && error.status < 500) {
						return false;
					}
					return failureCount < 2;
				},
			},
		},
		mutationCache: new MutationCache({
			onError: (error) => {
				// Global fallback for mutations that don't handle errors themselves.
				// Consumers using mutateAsync + try/catch handle their own errors.
				console.error("Mutation failed:", error);
			},
		}),
	});
}

// Lazy singleton for the browser — avoids module-scope instantiation that
// leaks cache across SSR requests.
let browserQueryClient: QueryClient | undefined;
function getQueryClient() {
	if (typeof window === "undefined") {
		// SSR: always create a fresh client per request to prevent cross-user cache leaks
		return makeQueryClient();
	}
	// Browser: reuse the same client across navigations
	if (!browserQueryClient) browserQueryClient = makeQueryClient();
	return browserQueryClient;
}

const KumoLink = forwardRef<
	HTMLAnchorElement,
	React.AnchorHTMLAttributes<HTMLAnchorElement> & { href?: string }
>(function KumoLink({ href, ...props }, ref) {
	if (href && !href.startsWith("http")) {
		return (
			<RouterLink to={href} ref={ref} {...(props as Record<string, unknown>)} />
		);
	}
	return <a href={href} ref={ref} {...props} />;
});

export function Layout({ children }: { children: React.ReactNode }) {
	return (
		<html lang="en" className="dark">
			<head>
				<meta charSet="UTF-8" />
				<link rel="icon" type="image/png" href="https://raw.githubusercontent.com/Jagrit0711/zuup-main/bc25cc6dafa9026827ffffa84f5d6740d86950ab/public/lovable-uploads/b44b8051-6117-4b37-999d-014c4c33dd13.png" />
				<meta name="viewport" content="width=device-width, initial-scale=1.0" />
				<title>Zuup Mail</title>
				<Meta />
				<Links />
			</head>
			<body className="bg-kumo-recessed text-kumo-default antialiased">
				{children}
				<ScrollRestoration />
				<Scripts />
			</body>
		</html>
	);
}

export function HydrateFallback() {
	return (
		<div className="flex items-center justify-center h-screen">
			<Loader size="lg" />
		</div>
	);
}

function LoginScreen() {
	const { instance, inProgress } = useMsal();
	return (
		<div className="flex flex-col min-h-screen bg-[#09090b] text-kumo-default relative overflow-x-hidden font-sans">
			{/* Aesthetic Background Glows */}
			<div className="absolute top-[-10%] left-[-10%] w-[40%] h-[40%] bg-rose-500/10 blur-[120px] rounded-full pointer-events-none" />
			<div className="absolute top-[40%] right-[-10%] w-[40%] h-[40%] bg-blue-500/10 blur-[120px] rounded-full pointer-events-none" />
			
			{/* Top Navigation */}
			<header className="sticky top-0 z-50 flex items-center justify-between px-6 py-4 border-b border-zinc-800/50 bg-[#09090b]/80 backdrop-blur-md">
				<div className="flex items-center gap-8">
					<div className="flex items-center gap-2 cursor-pointer" onClick={() => window.scrollTo(0,0)}>
						<img 
							src="https://raw.githubusercontent.com/Jagrit0711/zuup-main/bc25cc6dafa9026827ffffa84f5d6740d86950ab/public/lovable-uploads/b44b8051-6117-4b37-999d-014c4c33dd13.png" 
							alt="Zuup Logo" 
							className="w-6 h-6 rounded-md" 
						/>
						<span className="font-bold text-lg text-white tracking-wide">Zuup</span>
						<span className="text-zinc-400 font-medium">Mail</span>
						<span className="ml-2 text-[10px] font-semibold tracking-wider uppercase bg-rose-500/10 text-rose-400 px-2 py-0.5 rounded-full border border-rose-500/20">Agentic</span>
					</div>
					<nav className="hidden md:flex items-center gap-6 text-sm font-medium text-zinc-400">
						<a href="#features" className="hover:text-white transition-colors">Features</a>
						<a href="#docs" className="hover:text-white transition-colors">Docs</a>
						<a href="#apps" className="hover:text-white transition-colors">Apps</a>
					</nav>
				</div>
				<div className="flex items-center gap-4">
					<button 
						onClick={() => instance.loginRedirect(loginRequest).catch(console.error)}
						disabled={inProgress !== "none"}
						className={`text-sm font-medium text-zinc-400 hover:text-white transition-colors hidden sm:block ${inProgress !== "none" ? "opacity-50 cursor-not-allowed" : ""}`}
					>
						Sign In
					</button>
					<button 
						onClick={() => instance.loginRedirect(loginRequest).catch(console.error)}
						disabled={inProgress !== "none"}
						className={`text-sm font-medium bg-gradient-to-r from-rose-500 to-pink-500 text-white px-4 py-2 rounded-md hover:opacity-90 transition-opacity shadow-[0_0_15px_rgba(244,63,94,0.3)] ${inProgress !== "none" ? "opacity-50 cursor-not-allowed" : ""}`}
					>
						Get Started →
					</button>
				</div>
			</header>

			{/* Hero Section */}
			<section className="relative z-10 flex flex-col items-center justify-center text-center px-4 pt-32 pb-24 border-b border-zinc-800/50">
				<div className="inline-flex items-center gap-2 px-3 py-1 rounded-full border border-zinc-800 bg-zinc-900/50 text-xs font-medium text-zinc-400 mb-8">
					<svg className="w-3.5 h-3.5 text-rose-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
						<path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z" />
					</svg>
					100% Edge-Native on Cloudflare
				</div>
				
				<h1 className="text-5xl md:text-7xl font-extrabold tracking-tight text-white mb-2 leading-tight">
					Mail for the <br />
					<span className="text-transparent bg-clip-text bg-gradient-to-r from-rose-400 to-pink-500">Zuup ecosystem.</span>
				</h1>
				<h1 className="text-5xl md:text-7xl font-extrabold tracking-tight text-zinc-500 mb-8 leading-tight">
					Done right.
				</h1>
				
				<p className="text-lg md:text-xl text-zinc-400 max-w-2xl mx-auto mb-10 leading-relaxed">
					One inbox across every Zuup service. Agentic AI, autonomous auto-replies, smart tool calling, and 100% persistent state — built for founders who value their time.
				</p>
				
				<div className="flex flex-col sm:flex-row items-center gap-4 mb-10">
					<button 
						onClick={() => instance.loginRedirect(loginRequest).catch(console.error)}
						disabled={inProgress !== "none"}
						className={`w-full sm:w-auto px-8 py-3.5 bg-gradient-to-r from-rose-500 to-pink-500 text-white font-medium rounded-lg hover:opacity-90 transition-opacity shadow-[0_0_20px_rgba(244,63,94,0.4)] flex items-center justify-center gap-2 ${inProgress !== "none" ? "opacity-50 cursor-not-allowed" : ""}`}
					>
						Create Account →
					</button>
					<a 
						href="#docs"
						className="w-full sm:w-auto px-8 py-3.5 bg-zinc-900 border border-zinc-800 text-zinc-300 font-medium rounded-lg hover:bg-zinc-800 transition-colors flex items-center justify-center gap-2"
					>
						View Docs
					</a>
				</div>
				
				<div className="flex flex-wrap justify-center items-center gap-4 text-xs font-medium text-zinc-500 uppercase tracking-widest">
					<span>Agentic AI</span>
					<span className="w-1 h-1 bg-zinc-700 rounded-full"></span>
					<span>Auto-Replies</span>
					<span className="w-1 h-1 bg-zinc-700 rounded-full"></span>
					<span>SQLite-Backed</span>
					<span className="w-1 h-1 bg-zinc-700 rounded-full"></span>
					<span>Cloudflare Workers</span>
				</div>
			</section>

			{/* Features Section */}
			<section id="features" className="relative z-10 px-6 py-24 border-b border-zinc-800/50">
				<div className="max-w-4xl mx-auto">
					<h2 className="text-sm font-bold tracking-widest text-rose-500 uppercase mb-2">Features</h2>
					<h3 className="text-3xl md:text-4xl font-bold text-white mb-6">Everything you need. Nothing you don't.</h3>
					<p className="text-zinc-400 mb-12 text-lg">Built 100% on Cloudflare's Edge Network for zero latency and infinite scale. No legacy databases, no slow servers.</p>
					
					<div className="grid md:grid-cols-2 gap-8">
						<div className="bg-zinc-900/50 border border-zinc-800 p-6 rounded-xl">
							<h4 className="text-xl font-bold text-white mb-2">Cloudflare Workers</h4>
							<p className="text-zinc-400">Deployed globally on Cloudflare's edge. Our background agent processes your incoming email routing instantly at edge nodes closest to you.</p>
						</div>
						<div className="bg-zinc-900/50 border border-zinc-800 p-6 rounded-xl">
							<h4 className="text-xl font-bold text-white mb-2">Durable Objects + SQLite</h4>
							<p className="text-zinc-400">Stateful, persistent AI memory. Every chat and setting is securely saved in a high-performance SQLite database embedded inside a Cloudflare Durable Object.</p>
						</div>
						<div className="bg-zinc-900/50 border border-zinc-800 p-6 rounded-xl">
							<h4 className="text-xl font-bold text-white mb-2">Autonomous Llama 3.1</h4>
							<p className="text-zinc-400">Powered by Cloudflare Workers AI. Our Llama 3.1 8B agent scrapes your MS Graph context, triages emails, and auto-replies directly.</p>
						</div>
						<div className="bg-zinc-900/50 border border-zinc-800 p-6 rounded-xl">
							<h4 className="text-xl font-bold text-white mb-2">Human Escapement</h4>
							<p className="text-zinc-400">If the bot encounters a complex negotiation or a question it can't confidently answer, it generates a Ticket ID and forwards it to your human inbox.</p>
						</div>
					</div>
				</div>
			</section>

			{/* Docs Section */}
			<section id="docs" className="relative z-10 px-6 py-24 bg-[#050505] border-b border-zinc-800/50">
				<div className="max-w-4xl mx-auto">
					<h2 className="text-sm font-bold tracking-widest text-rose-500 uppercase mb-2">Docs</h2>
					<h3 className="text-3xl md:text-4xl font-bold text-white mb-6">Four steps to Auto-Pilot.</h3>
					<p className="text-zinc-400 mb-12 text-lg">Since the AI agent runs securely on Cloudflare's backend, it needs an Azure App Registration (Client Credentials flow) to read your emails securely without human intervention.</p>
					
					<div className="space-y-12">
						<div className="flex gap-6">
							<div className="shrink-0 w-8 h-8 rounded-full bg-rose-500/20 text-rose-400 border border-rose-500/30 flex items-center justify-center font-bold">1</div>
							<div>
								<h4 className="text-xl font-bold text-white mb-2">Create an Azure App Registration</h4>
								<p className="text-zinc-400">Go to the Azure Portal &gt; Entra ID &gt; App Registrations. Create a new app. Grab your <code className="text-pink-400 bg-pink-400/10 px-1.5 py-0.5 rounded">Tenant ID</code> and <code className="text-pink-400 bg-pink-400/10 px-1.5 py-0.5 rounded">Client ID</code>.</p>
							</div>
						</div>
						<div className="flex gap-6">
							<div className="shrink-0 w-8 h-8 rounded-full bg-rose-500/20 text-rose-400 border border-rose-500/30 flex items-center justify-center font-bold">2</div>
							<div>
								<h4 className="text-xl font-bold text-white mb-2">Generate a Client Secret</h4>
								<p className="text-zinc-400">In your App Registration, navigate to "Certificates & secrets" and create a new Client Secret. Save the secret value.</p>
							</div>
						</div>
						<div className="flex gap-6">
							<div className="shrink-0 w-8 h-8 rounded-full bg-rose-500/20 text-rose-400 border border-rose-500/30 flex items-center justify-center font-bold">3</div>
							<div>
								<h4 className="text-xl font-bold text-white mb-2">Grant API Permissions</h4>
								<p className="text-zinc-400">Under "API Permissions", add Microsoft Graph &gt; <strong>Application Permissions</strong> (not Delegated). Select <code className="text-pink-400 bg-pink-400/10 px-1.5 py-0.5 rounded">Mail.ReadWrite</code> and <code className="text-pink-400 bg-pink-400/10 px-1.5 py-0.5 rounded">Mail.Send</code>. Click "Grant admin consent for your organization".</p>
							</div>
						</div>
						<div className="flex gap-6">
							<div className="shrink-0 w-8 h-8 rounded-full bg-rose-500/20 text-rose-400 border border-rose-500/30 flex items-center justify-center font-bold">4</div>
							<div>
								<h4 className="text-xl font-bold text-white mb-2">Set Cloudflare Secrets</h4>
								<p className="text-zinc-400 mb-4">Run the following commands to securely store your credentials in Cloudflare:</p>
								<pre className="bg-[#09090b] border border-zinc-800 p-4 rounded-lg overflow-x-auto text-sm text-zinc-300 font-mono">
npx wrangler secret put AZURE_TENANT_ID
npx wrangler secret put AZURE_CLIENT_ID
npx wrangler secret put AZURE_CLIENT_SECRET
npx wrangler secret put HUMAN_FALLBACK_EMAIL
								</pre>
								<p className="text-zinc-500 mt-3 text-sm">Note: <code className="text-zinc-400">HUMAN_FALLBACK_EMAIL</code> is the inbox the AI will forward complex questions to (e.g., jagrit@zuup.dev).</p>
							</div>
						</div>
					</div>
				</div>
			</section>

			{/* Apps / Ecosystem Section */}
			<section id="apps" className="relative z-10 px-6 py-24">
				<div className="max-w-4xl mx-auto text-center">
					<h2 className="text-sm font-bold tracking-widest text-rose-500 uppercase mb-2">Ecosystem</h2>
					<h3 className="text-3xl md:text-4xl font-bold text-white mb-12">Works across all Zuup apps.</h3>
					
					<div className="grid md:grid-cols-3 gap-6 text-left">
						<div className="bg-zinc-900/30 border border-zinc-800 p-6 rounded-xl hover:bg-zinc-800/50 transition-colors cursor-pointer">
							<h4 className="text-lg font-bold text-white mb-1 flex items-center justify-between">ZuupCode <span className="text-[10px] bg-rose-500/10 text-rose-400 px-2 py-0.5 rounded-full border border-rose-500/20">First-party</span></h4>
							<p className="text-zinc-400 text-sm">Browser-based IDE with 30+ languages.</p>
						</div>
						<div className="bg-zinc-900/30 border border-zinc-800 p-6 rounded-xl hover:bg-zinc-800/50 transition-colors cursor-pointer">
							<h4 className="text-lg font-bold text-white mb-1 flex items-center justify-between">ZuupTime <span className="text-[10px] bg-rose-500/10 text-rose-400 px-2 py-0.5 rounded-full border border-rose-500/20">First-party</span></h4>
							<p className="text-zinc-400 text-sm">Time tracking for developers.</p>
						</div>
						<div className="bg-zinc-900/30 border border-zinc-800 p-6 rounded-xl hover:bg-zinc-800/50 transition-colors cursor-pointer">
							<h4 className="text-lg font-bold text-white mb-1 flex items-center justify-between">Zuup Auth <span className="text-[10px] bg-rose-500/10 text-rose-400 px-2 py-0.5 rounded-full border border-rose-500/20">First-party</span></h4>
							<p className="text-zinc-400 text-sm">OAuth 2.1 identity provider for the ecosystem.</p>
						</div>
					</div>
					
					<div className="mt-20 pt-8 border-t border-zinc-800/50 text-sm text-zinc-500 flex flex-col items-center gap-2">
						<p>© 2026 Zuup · Made by Jagrit Sachdev</p>
						<a 
							href="https://github.com/Jagrit0711/zuupaimail" 
							target="_blank" 
							rel="noreferrer"
							className="hover:text-white transition-colors flex items-center gap-1 mt-2"
						>
							View Source on GitHub
						</a>
					</div>
				</div>
			</section>
		</div>
	);
}

export default function App() {
	const [queryClient] = useState(getQueryClient);
	const [isMsalInitialized, setIsMsalInitialized] = reactUseState(false);

	useEffect(() => {
		msalInstance.initialize().then(() => {
			setIsMsalInitialized(true);
		});
	}, []);

	if (!isMsalInitialized) {
		return <HydrateFallback />;
	}

	return (
		<MsalProvider instance={msalInstance}>
			<QueryClientProvider client={queryClient}>
				<LinkProvider component={KumoLink}>
					<TooltipProvider>
						<Toasty>
							<AuthenticatedTemplate>
								<Outlet />
							</AuthenticatedTemplate>
							<UnauthenticatedTemplate>
								<LoginScreen />
							</UnauthenticatedTemplate>
						</Toasty>
					</TooltipProvider>
				</LinkProvider>
			</QueryClientProvider>
		</MsalProvider>
	);
}

export function ErrorBoundary({ error }: { error: unknown }) {
	let title = "Something went wrong";
	let description = "An unexpected error occurred. Please try again.";
	let status: number | null = null;

	if (isRouteErrorResponse(error)) {
		status = error.status;
		if (error.status === 404) {
			title = "Page not found";
			description =
				"The page you're looking for doesn't exist or has been moved.";
		} else {
			title = `Error ${error.status}`;
			description = error.statusText || description;
		}
	} else if (error instanceof Error && import.meta.env.DEV) {
		description = error.message;
	}

	return (
		<div className="flex items-center justify-center min-h-screen p-8">
			<Empty
				icon={<WarningIcon size={48} className="text-kumo-inactive" />}
				title={status === 404 ? "404 — Page not found" : title}
				description={description}
				contents={
					<Button
						variant="primary"
						onClick={() => {
							window.location.href = "/";
						}}
					>
						Go Home
					</Button>
				}
			/>
		</div>
	);
}
