import { type Configuration, PublicClientApplication } from "@azure/msal-browser";

export const msalConfig: Configuration = {
  auth: {
    clientId: import.meta.env.VITE_AZURE_CLIENT_ID || "9a7b8329-1c9e-4dc5-a47c-1d4eef806c5d",
    authority: `https://login.microsoftonline.com/${import.meta.env.VITE_AZURE_TENANT_ID || "6ac6a7c6-5c84-42e2-b558-236d04db78f1"}`,
    redirectUri: typeof window !== "undefined" ? window.location.origin : "http://localhost:5173",
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
};

export const msalInstance = new PublicClientApplication(msalConfig);

export const loginRequest = {
  scopes: ["User.Read", "Mail.ReadWrite", "Mail.Send"]
};
