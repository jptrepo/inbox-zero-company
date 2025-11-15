# Office365 Integration - Detailed Implementation Plan

## Overview

This document provides a detailed, step-by-step implementation plan for adding Office365/Outlook support to Inbox Zero using Microsoft Graph API with NextAuth.js.

**Approach**: Microsoft Graph API with NextAuth.js Provider (Approach 1 from evaluation)
**Timeline**: 4-5 weeks
**Estimated Effort**: 80-100 development hours

---

## Prerequisites

### Azure Portal Setup

1. **Create Azure AD Application**
   - Go to [Azure Portal](https://portal.azure.com)
   - Navigate to "Azure Active Directory" → "App registrations"
   - Click "New registration"
   - Name: "Inbox Zero Email Client"
   - Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"
   - Redirect URI: Web - `https://yourdomain.com/api/auth/callback/azure-ad`
   - Save Application (client) ID and Directory (tenant) ID

2. **Configure API Permissions**
   - Go to "API permissions" → "Add a permission"
   - Select "Microsoft Graph" → "Delegated permissions"
   - Add these permissions:
     - `User.Read` - Sign in and read user profile
     - `Mail.ReadWrite` - Read and write mail
     - `Mail.Send` - Send mail
     - `MailboxSettings.Read` - Read mailbox settings
     - `offline_access` - Maintain access to data
     - `openid` - Sign in
     - `profile` - View users' basic profile
     - `email` - View users' email address
   - Click "Grant admin consent" (if available)

3. **Create Client Secret**
   - Go to "Certificates & secrets" → "Client secrets"
   - Click "New client secret"
   - Description: "Inbox Zero Production"
   - Expires: 24 months
   - Save the secret value (only shown once!)

4. **Configure Token Configuration** (optional)
   - Go to "Token configuration"
   - Add optional claims if needed
   - Configure token lifetime settings

---

## Phase 1: Authentication Setup (Week 1)

### Task 1.1: Environment Configuration

**File**: `apps/web/.env.example`

```bash
# Add these variables
AZURE_AD_CLIENT_ID=your_client_id_here
AZURE_AD_CLIENT_SECRET=your_client_secret_here
AZURE_AD_TENANT_ID=common
```

**File**: `apps/web/env.ts`

```typescript
// Add to server config
AZURE_AD_CLIENT_ID: z.string().optional(), // Make optional for now
AZURE_AD_CLIENT_SECRET: z.string().optional(),
AZURE_AD_TENANT_ID: z.string().default("common"),
MICROSOFT_GRAPH_WEBHOOK_VALIDATION_TOKEN: z.string().optional(),
```

**Testing**:
- [ ] Verify environment variables load correctly
- [ ] Run `pnpm run dev` without errors

---

### Task 1.2: Add Microsoft Provider to NextAuth

**File**: `apps/web/utils/auth.ts`

**Changes**:

1. Import Azure AD provider:
```typescript
import AzureADProvider from "next-auth/providers/azure-ad";
```

2. Define Microsoft scopes:
```typescript
export const MICROSOFT_SCOPES = [
  "openid",
  "profile",
  "email",
  "offline_access",
  "User.Read",
  "Mail.ReadWrite",
  "Mail.Send",
  "MailboxSettings.Read",
];
```

3. Add provider to configuration (conditional based on env):
```typescript
providers: [
  GoogleProvider({
    clientId: env.GOOGLE_CLIENT_ID,
    clientSecret: env.GOOGLE_CLIENT_SECRET,
    authorization: {
      url: "https://accounts.google.com/o/oauth2/v2/auth",
      params: {
        scope: SCOPES.join(" "),
        access_type: "offline",
        response_type: "code",
        ...(options?.consent ? { prompt: "consent" } : {}),
      },
    },
  }),
  ...(env.AZURE_AD_CLIENT_ID && env.AZURE_AD_CLIENT_SECRET
    ? [
        AzureADProvider({
          clientId: env.AZURE_AD_CLIENT_ID,
          clientSecret: env.AZURE_AD_CLIENT_SECRET,
          tenantId: env.AZURE_AD_TENANT_ID,
          authorization: {
            params: {
              scope: MICROSOFT_SCOPES.join(" "),
              ...(options?.consent ? { prompt: "consent" } : {}),
            },
          },
        }),
      ]
    : []),
],
```

**Testing**:
- [ ] Application starts without errors
- [ ] Visit `/api/auth/signin` and see both Google and Microsoft options
- [ ] Test Microsoft OAuth flow (should succeed if credentials are correct)

---

### Task 1.3: Database Schema Updates

**File**: `apps/web/prisma/schema.prisma`

**Changes**:

```prisma
model User {
  // ... existing fields
  
  // Add email provider tracking
  emailProvider String? @default("gmail") // "gmail" or "outlook"
}

// Account model already supports multiple providers via the provider field
// No changes needed here
```

**Migration**:
```bash
cd apps/web
pnpm prisma migrate dev --name add_email_provider_to_user
```

**Testing**:
- [ ] Migration runs successfully
- [ ] Database schema updated
- [ ] Prisma client regenerated
- [ ] No errors when starting application

---

### Task 1.4: Provider Selection Logic

**File**: `apps/web/utils/actions/user.ts` (new file or add to existing)

```typescript
import prisma from "@/utils/prisma";
import { auth } from "@/app/api/auth/[...nextauth]/auth";

export async function updateUserEmailProvider(provider: "gmail" | "outlook") {
  const session = await auth();
  if (!session?.user?.email) throw new Error("Not authenticated");

  await prisma.user.update({
    where: { email: session.user.email },
    data: { emailProvider: provider },
  });

  return { success: true };
}

export async function getUserEmailProvider() {
  const session = await auth();
  if (!session?.user?.email) return null;

  const user = await prisma.user.findUnique({
    where: { email: session.user.email },
    select: { emailProvider: true },
  });

  return user?.emailProvider || "gmail";
}
```

**Testing**:
- [ ] Function compiles without errors
- [ ] Can update user's email provider
- [ ] Can retrieve user's email provider

---

### Task 1.5: Login Page Updates

**File**: `apps/web/app/(landing)/login/page.tsx`

**Changes**: Update login page to show both provider options clearly.

Example button additions:
```typescript
<Button
  variant="outline"
  onClick={() => signIn("azure-ad", { callbackUrl: "/mail" })}
>
  <MicrosoftIcon className="mr-2" />
  Continue with Microsoft
</Button>
```

**File**: `apps/web/components/icons/MicrosoftIcon.tsx` (new)

```typescript
export function MicrosoftIcon({ className }: { className?: string }) {
  return (
    <svg className={className} viewBox="0 0 21 21" fill="none">
      <rect x="1" y="1" width="9" height="9" fill="#F25022"/>
      <rect x="1" y="11" width="9" height="9" fill="#00A4EF"/>
      <rect x="11" y="1" width="9" height="9" fill="#7FBA00"/>
      <rect x="11" y="11" width="9" height="9" fill="#FFB900"/>
    </svg>
  );
}
```

**Testing**:
- [ ] Login page displays both Google and Microsoft buttons
- [ ] Clicking Microsoft button initiates OAuth flow
- [ ] Successful authentication redirects to application
- [ ] User account is created in database

---

## Phase 2: Email Provider Abstraction (Week 2)

### Task 2.1: Define Email Provider Interface

**File**: `apps/web/utils/email/types.ts` (new)

```typescript
export type EmailProvider = "gmail" | "outlook";

export interface Message {
  id: string;
  threadId: string;
  from: EmailAddress;
  to: EmailAddress[];
  cc?: EmailAddress[];
  bcc?: EmailAddress[];
  subject: string;
  snippet: string;
  textPlain?: string;
  textHtml?: string;
  date: Date;
  labels?: string[]; // Gmail: label IDs, Outlook: folder IDs
  unread: boolean;
  attachments?: Attachment[];
  headers?: Record<string, string>;
}

export interface EmailAddress {
  address: string;
  name?: string;
}

export interface Attachment {
  id: string;
  filename: string;
  mimeType: string;
  size: number;
  contentId?: string;
}

export interface Folder {
  id: string;
  name: string;
  displayName: string;
  unreadCount?: number;
  totalCount?: number;
}

export interface Thread {
  id: string;
  messages: Message[];
}

export interface ListMessagesOptions {
  maxResults?: number;
  pageToken?: string;
  labelIds?: string[];
  q?: string;
  includeSpamTrash?: boolean;
}

export interface SendMessageOptions {
  to: EmailAddress[];
  cc?: EmailAddress[];
  bcc?: EmailAddress[];
  subject: string;
  textPlain?: string;
  textHtml?: string;
  replyTo?: string;
  inReplyTo?: string;
  references?: string[];
  threadId?: string;
}
```

---

### Task 2.2: Create Email Provider Interface

**File**: `apps/web/utils/email/provider.ts` (new)

```typescript
import type {
  Message,
  Folder,
  Thread,
  ListMessagesOptions,
  SendMessageOptions,
  Attachment,
} from "./types";

export interface IEmailProvider {
  // Message operations
  listMessages(options: ListMessagesOptions): Promise<{
    messages: Message[];
    nextPageToken?: string;
  }>;
  
  getMessage(id: string): Promise<Message>;
  
  sendMessage(options: SendMessageOptions): Promise<Message>;
  
  deleteMessage(id: string): Promise<void>;
  
  markAsRead(id: string): Promise<void>;
  
  markAsUnread(id: string): Promise<void>;
  
  archiveMessage(id: string): Promise<void>;
  
  // Folder/Label operations
  listFolders(): Promise<Folder[]>;
  
  createFolder(name: string, color?: string): Promise<Folder>;
  
  updateFolder(id: string, updates: Partial<Folder>): Promise<Folder>;
  
  deleteFolder(id: string): Promise<void>;
  
  addToFolder(messageId: string, folderId: string): Promise<void>;
  
  removeFromFolder(messageId: string, folderId: string): Promise<void>;
  
  // Thread operations
  getThread(id: string): Promise<Thread>;
  
  // Search
  searchMessages(query: string): Promise<Message[]>;
  
  // Attachments
  getAttachment(messageId: string, attachmentId: string): Promise<Attachment & { data: Buffer }>;
  
  // Webhook/Watch
  setupWebhook?(callbackUrl: string): Promise<{ expirationDate: Date }>;
  
  refreshWebhook?(): Promise<void>;
  
  // Utility
  getProviderType(): "gmail" | "outlook";
}
```

---

### Task 2.3: Create Microsoft Graph Client

**File**: `apps/web/utils/microsoft/client.ts` (new)

```typescript
import { Client } from "@microsoft/microsoft-graph-client";
import { saveRefreshToken } from "@/utils/auth";
import { env } from "@/env";
import { createScopedLogger } from "@/utils/logger";

const logger = createScopedLogger("microsoft/client");

type ClientOptions = {
  accessToken?: string;
  refreshToken?: string;
};

export const getMicrosoftGraphClient = (session: ClientOptions) => {
  if (!session.accessToken) {
    throw new Error("No access token available");
  }

  const client = Client.init({
    authProvider: (done) => {
      done(null, session.accessToken);
    },
  });

  return client;
};

export const getMicrosoftGraphClientWithRefresh = async (
  session: ClientOptions & { refreshToken: string; expiryDate?: number | null },
  providerAccountId: string,
) => {
  // If token is still valid, return client
  if (session.expiryDate && session.expiryDate > Date.now()) {
    return getMicrosoftGraphClient(session);
  }

  // Refresh the token
  try {
    const response = await fetch(
      `https://login.microsoftonline.com/${env.AZURE_AD_TENANT_ID}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: env.AZURE_AD_CLIENT_ID,
          client_secret: env.AZURE_AD_CLIENT_SECRET,
          grant_type: "refresh_token",
          refresh_token: session.refreshToken,
          scope: "offline_access User.Read Mail.ReadWrite Mail.Send MailboxSettings.Read",
        }),
      }
    );

    if (!response.ok) {
      const error = await response.text();
      logger.error("Error refreshing Microsoft access token", { error });
      throw new Error("Failed to refresh token");
    }

    const tokens = await response.json();

    // Save new tokens
    await saveRefreshToken(
      {
        access_token: tokens.access_token,
        refresh_token: tokens.refresh_token || session.refreshToken,
        expires_at: Math.floor(Date.now() / 1000 + tokens.expires_in),
      },
      {
        refresh_token: session.refreshToken,
        providerAccountId,
      }
    );

    return getMicrosoftGraphClient({ accessToken: tokens.access_token });
  } catch (error) {
    logger.error("Error refreshing Microsoft token", { error });
    throw error;
  }
};
```

**Install dependency**:
```bash
cd apps/web
pnpm add @microsoft/microsoft-graph-client @microsoft/microsoft-graph-types isomorphic-fetch
```

**Testing**:
- [ ] Can create Graph client with valid token
- [ ] Token refresh works correctly
- [ ] Errors are logged appropriately

---

### Task 2.4: Implement Microsoft Graph Provider

**File**: `apps/web/utils/microsoft/provider.ts` (new)

```typescript
import { Client } from "@microsoft/microsoft-graph-client";
import type { Message as GraphMessage, MailFolder } from "@microsoft/microsoft-graph-types";
import { getMicrosoftGraphClient } from "./client";
import type { IEmailProvider } from "../email/provider";
import type {
  Message,
  Folder,
  Thread,
  ListMessagesOptions,
  SendMessageOptions,
  Attachment,
  EmailAddress,
} from "../email/types";

export class MicrosoftGraphProvider implements IEmailProvider {
  private client: Client;

  constructor(accessToken: string) {
    this.client = getMicrosoftGraphClient({ accessToken });
  }

  getProviderType(): "outlook" {
    return "outlook";
  }

  async listMessages(options: ListMessagesOptions = {}) {
    const { maxResults = 50, pageToken, labelIds, q } = options;

    let query = this.client.api("/me/messages").top(maxResults);

    // Apply folder filter (labelIds in Outlook are folder IDs)
    if (labelIds && labelIds.length > 0) {
      // For Outlook, we need to query specific folders
      const folderId = labelIds[0]; // Start with first folder
      query = this.client.api(`/me/mailFolders/${folderId}/messages`).top(maxResults);
    }

    // Apply search filter
    if (q) {
      query = query.search(q);
    }

    // Apply pagination
    if (pageToken) {
      query = query.skipToken(pageToken);
    }

    const response = await query.get();

    return {
      messages: response.value.map((msg: GraphMessage) => this.convertToMessage(msg)),
      nextPageToken: response["@odata.nextLink"]
        ? new URL(response["@odata.nextLink"]).searchParams.get("$skiptoken") || undefined
        : undefined,
    };
  }

  async getMessage(id: string): Promise<Message> {
    const message = await this.client
      .api(`/me/messages/${id}`)
      .select("*")
      .expand("attachments")
      .get();

    return this.convertToMessage(message);
  }

  async sendMessage(options: SendMessageOptions): Promise<Message> {
    const message = {
      subject: options.subject,
      body: {
        contentType: options.textHtml ? "HTML" : "Text",
        content: options.textHtml || options.textPlain || "",
      },
      toRecipients: options.to.map(addr => ({
        emailAddress: { address: addr.address, name: addr.name },
      })),
      ccRecipients: options.cc?.map(addr => ({
        emailAddress: { address: addr.address, name: addr.name },
      })),
      bccRecipients: options.bcc?.map(addr => ({
        emailAddress: { address: addr.address, name: addr.name },
      })),
    };

    const sent = await this.client.api("/me/sendMail").post({ message });
    
    // Graph API doesn't return the sent message, so we fetch it
    // This is a limitation - we might need to return a partial message
    return this.convertToMessage(sent);
  }

  async deleteMessage(id: string): Promise<void> {
    await this.client.api(`/me/messages/${id}`).delete();
  }

  async markAsRead(id: string): Promise<void> {
    await this.client.api(`/me/messages/${id}`).patch({ isRead: true });
  }

  async markAsUnread(id: string): Promise<void> {
    await this.client.api(`/me/messages/${id}`).patch({ isRead: false });
  }

  async archiveMessage(id: string): Promise<void> {
    // In Outlook, archiving means moving to the Archive folder
    const folders = await this.listFolders();
    const archiveFolder = folders.find(f => f.name === "Archive");
    
    if (archiveFolder) {
      await this.client.api(`/me/messages/${id}/move`).post({
        destinationId: archiveFolder.id,
      });
    }
  }

  async listFolders(): Promise<Folder[]> {
    const response = await this.client.api("/me/mailFolders").get();

    return response.value.map((folder: MailFolder) => ({
      id: folder.id!,
      name: folder.displayName!,
      displayName: folder.displayName!,
      unreadCount: folder.unreadItemCount,
      totalCount: folder.totalItemCount,
    }));
  }

  async createFolder(name: string): Promise<Folder> {
    const folder = await this.client.api("/me/mailFolders").post({
      displayName: name,
    });

    return {
      id: folder.id,
      name: folder.displayName,
      displayName: folder.displayName,
      unreadCount: 0,
      totalCount: 0,
    };
  }

  async updateFolder(id: string, updates: Partial<Folder>): Promise<Folder> {
    const folder = await this.client.api(`/me/mailFolders/${id}`).patch({
      displayName: updates.displayName || updates.name,
    });

    return {
      id: folder.id,
      name: folder.displayName,
      displayName: folder.displayName,
    };
  }

  async deleteFolder(id: string): Promise<void> {
    await this.client.api(`/me/mailFolders/${id}`).delete();
  }

  async addToFolder(messageId: string, folderId: string): Promise<void> {
    await this.client.api(`/me/messages/${messageId}/move`).post({
      destinationId: folderId,
    });
  }

  async removeFromFolder(messageId: string, folderId: string): Promise<void> {
    // Outlook doesn't have a "remove from folder" concept
    // Messages exist in one folder at a time
    // This would need to move to inbox or another default folder
    const folders = await this.listFolders();
    const inboxFolder = folders.find(f => f.name === "Inbox");
    
    if (inboxFolder) {
      await this.addToFolder(messageId, inboxFolder.id);
    }
  }

  async getThread(id: string): Promise<Thread> {
    // Outlook uses conversationId for threads
    const message = await this.getMessage(id);
    
    // Get all messages in the conversation
    const response = await this.client
      .api(`/me/messages`)
      .filter(`conversationId eq '${message.threadId}'`)
      .orderby("receivedDateTime asc")
      .get();

    return {
      id: message.threadId,
      messages: response.value.map((msg: GraphMessage) => this.convertToMessage(msg)),
    };
  }

  async searchMessages(query: string): Promise<Message[]> {
    const response = await this.client.api("/me/messages").search(query).top(50).get();

    return response.value.map((msg: GraphMessage) => this.convertToMessage(msg));
  }

  async getAttachment(messageId: string, attachmentId: string): Promise<Attachment & { data: Buffer }> {
    const attachment = await this.client
      .api(`/me/messages/${messageId}/attachments/${attachmentId}`)
      .get();

    return {
      id: attachment.id,
      filename: attachment.name,
      mimeType: attachment.contentType,
      size: attachment.size,
      data: Buffer.from(attachment.contentBytes, "base64"),
    };
  }

  async setupWebhook(callbackUrl: string): Promise<{ expirationDate: Date }> {
    // Microsoft Graph webhooks expire after max 4230 minutes (about 3 days)
    const expirationDateTime = new Date();
    expirationDateTime.setMinutes(expirationDateTime.getMinutes() + 4230);

    const subscription = await this.client.api("/subscriptions").post({
      changeType: "created,updated",
      notificationUrl: callbackUrl,
      resource: "/me/messages",
      expirationDateTime: expirationDateTime.toISOString(),
      clientState: "secretClientValue", // Should be from env
    });

    return {
      expirationDate: new Date(subscription.expirationDateTime),
    };
  }

  async refreshWebhook(): Promise<void> {
    // Need to store subscription ID to refresh
    // This is a simplified version
    throw new Error("Webhook refresh not implemented");
  }

  private convertToMessage(msg: GraphMessage): Message {
    return {
      id: msg.id!,
      threadId: msg.conversationId!,
      from: {
        address: msg.from?.emailAddress?.address || "",
        name: msg.from?.emailAddress?.name,
      },
      to: (msg.toRecipients || []).map(r => ({
        address: r.emailAddress?.address || "",
        name: r.emailAddress?.name,
      })),
      cc: (msg.ccRecipients || []).map(r => ({
        address: r.emailAddress?.address || "",
        name: r.emailAddress?.name,
      })),
      bcc: (msg.bccRecipients || []).map(r => ({
        address: r.emailAddress?.address || "",
        name: r.emailAddress?.name,
      })),
      subject: msg.subject || "",
      snippet: msg.bodyPreview || "",
      textPlain: msg.body?.contentType === "text" ? msg.body?.content : undefined,
      textHtml: msg.body?.contentType === "html" ? msg.body?.content : undefined,
      date: new Date(msg.receivedDateTime || msg.createdDateTime || Date.now()),
      unread: !msg.isRead,
      attachments: (msg.attachments || []).map(att => ({
        id: att.id || "",
        filename: att.name || "",
        mimeType: att.contentType || "",
        size: att.size || 0,
      })),
      headers: {}, // Graph API doesn't expose all headers easily
    };
  }
}
```

**Testing**:
- [ ] Can list messages
- [ ] Can get single message
- [ ] Can send message
- [ ] Can delete message
- [ ] Can mark as read/unread
- [ ] Folder operations work

---

### Task 2.5: Refactor Gmail Provider to Use Interface

**File**: `apps/web/utils/gmail/provider.ts` (new)

```typescript
import type { gmail_v1 } from "@googleapis/gmail";
import { getGmailClient } from "./client";
import type { IEmailProvider } from "../email/provider";
import type {
  Message,
  Folder,
  Thread,
  ListMessagesOptions,
  SendMessageOptions,
  Attachment,
} from "../email/types";
// Import existing Gmail utilities
import { parseMessage } from "./message";
import { getThread as gmailGetThread } from "./thread";
// ... other imports

export class GmailProvider implements IEmailProvider {
  private gmail: gmail_v1.Gmail;

  constructor(accessToken: string, refreshToken?: string) {
    this.gmail = getGmailClient({ accessToken, refreshToken });
  }

  getProviderType(): "gmail" {
    return "gmail";
  }

  async listMessages(options: ListMessagesOptions = {}) {
    const response = await this.gmail.users.messages.list({
      userId: "me",
      maxResults: options.maxResults || 50,
      pageToken: options.pageToken,
      labelIds: options.labelIds,
      q: options.q,
      includeSpamTrash: options.includeSpamTrash,
    });

    const messages = await Promise.all(
      (response.data.messages || []).map(async (msg) => {
        const full = await this.getMessage(msg.id!);
        return full;
      })
    );

    return {
      messages,
      nextPageToken: response.data.nextPageToken || undefined,
    };
  }

  async getMessage(id: string): Promise<Message> {
    const response = await this.gmail.users.messages.get({
      userId: "me",
      id,
      format: "full",
    });

    return parseMessage(response.data);
  }

  // Implement all other IEmailProvider methods using existing Gmail utils
  // ...
}
```

**Note**: This task involves refactoring existing Gmail code to implement the interface. Much of the logic already exists in `apps/web/utils/gmail/*.ts` files.

---

### Task 2.6: Create Provider Factory

**File**: `apps/web/utils/email/factory.ts` (new)

```typescript
import type { IEmailProvider } from "./provider";
import { GmailProvider } from "../gmail/provider";
import { MicrosoftGraphProvider } from "../microsoft/provider";
import { auth } from "@/app/api/auth/[...nextauth]/auth";
import prisma from "@/utils/prisma";

export async function getEmailProvider(): Promise<IEmailProvider> {
  const session = await auth();
  
  if (!session?.user?.email || !session.accessToken) {
    throw new Error("Not authenticated");
  }

  const user = await prisma.user.findUnique({
    where: { email: session.user.email },
    select: { emailProvider: true },
    include: {
      accounts: {
        select: {
          provider: true,
          access_token: true,
          refresh_token: true,
        },
      },
    },
  });

  if (!user) {
    throw new Error("User not found");
  }

  // Determine provider from user settings or account
  const provider = user.emailProvider || user.accounts[0]?.provider || "google";
  const account = user.accounts.find(acc => 
    provider === "gmail" ? acc.provider === "google" : acc.provider === "azure-ad"
  );

  if (!account) {
    throw new Error("No account found for provider");
  }

  if (provider === "gmail" || provider === "google") {
    return new GmailProvider(
      account.access_token!,
      account.refresh_token!
    );
  } else if (provider === "outlook" || provider === "azure-ad") {
    return new MicrosoftGraphProvider(account.access_token!);
  }

  throw new Error(`Unsupported provider: ${provider}`);
}

export async function getEmailProviderForUser(userId: string): Promise<IEmailProvider> {
  const user = await prisma.user.findUnique({
    where: { id: userId },
    select: { emailProvider: true },
    include: {
      accounts: {
        select: {
          provider: true,
          access_token: true,
          refresh_token: true,
        },
      },
    },
  });

  if (!user || !user.accounts.length) {
    throw new Error("User or account not found");
  }

  const provider = user.emailProvider || user.accounts[0].provider;
  const account = user.accounts[0];

  if (provider === "gmail") {
    return new GmailProvider(account.access_token!, account.refresh_token!);
  } else {
    return new MicrosoftGraphProvider(account.access_token!);
  }
}
```

---

## Phase 3: Feature Implementation (Week 3)

### Task 3.1: Update API Routes to Use Provider Factory

Update all email-related API routes to use the provider factory instead of directly calling Gmail functions.

**Example - File**: `apps/web/app/api/google/messages/[id]/route.ts`

**Before**:
```typescript
import { getGmailClient } from "@/utils/gmail/client";

export async function GET(request: Request, { params }: { params: { id: string } }) {
  const gmail = getGmailClient({ accessToken: session.accessToken });
  const message = await gmail.users.messages.get({ userId: "me", id: params.id });
  // ...
}
```

**After**:
```typescript
import { getEmailProvider } from "@/utils/email/factory";

export async function GET(request: Request, { params }: { params: { id: string } }) {
  const provider = await getEmailProvider();
  const message = await provider.getMessage(params.id);
  // ...
}
```

**Files to update**:
- [ ] `apps/web/app/api/google/messages/[id]/route.ts`
- [ ] `apps/web/app/api/google/threads/[id]/route.ts`
- [ ] `apps/web/app/api/user/*/route.ts` files
- [ ] All other API routes that interact with email

---

### Task 3.2: Implement Webhook Handler for Microsoft

**File**: `apps/web/app/api/microsoft/webhook/route.ts` (new)

```typescript
import { NextResponse } from "next/server";
import { headers } from "next/headers";
import { env } from "@/env";
import { createScopedLogger } from "@/utils/logger";

const logger = createScopedLogger("Microsoft Webhook");

export async function POST(request: Request) {
  const headersList = headers();
  const validationToken = new URL(request.url).searchParams.get("validationToken");

  // Microsoft sends validation request when setting up webhook
  if (validationToken) {
    logger.info("Webhook validation request received");
    return new Response(validationToken, {
      status: 200,
      headers: { "Content-Type": "text/plain" },
    });
  }

  // Verify client state
  const body = await request.json();
  const notifications = body.value || [];

  for (const notification of notifications) {
    if (notification.clientState !== env.MICROSOFT_GRAPH_WEBHOOK_VALIDATION_TOKEN) {
      logger.error("Invalid client state in webhook notification");
      continue;
    }

    // Process the notification
    logger.info("Processing Microsoft webhook notification", {
      resource: notification.resource,
      changeType: notification.changeType,
    });

    // TODO: Fetch the changed message and process it
    // Similar to how Google PubSub webhook works
  }

  return NextResponse.json({ success: true });
}

export async function GET(request: Request) {
  // Handle validation during setup
  const url = new URL(request.url);
  const validationToken = url.searchParams.get("validationToken");

  if (validationToken) {
    return new Response(validationToken, {
      status: 200,
      headers: { "Content-Type": "text/plain" },
    });
  }

  return NextResponse.json({ error: "Missing validation token" }, { status: 400 });
}
```

---

### Task 3.3: Update Settings Page

**File**: `apps/web/app/(app)/settings/page.tsx`

Add UI to show current email provider and allow switching (for users with multiple accounts):

```typescript
"use client";

import { useState } from "react";
import { Button } from "@/components/ui/button";
import { updateUserEmailProvider, getUserEmailProvider } from "@/utils/actions/user";

export function EmailProviderSettings() {
  const [provider, setProvider] = useState<"gmail" | "outlook">("gmail");

  return (
    <div className="space-y-4">
      <h3 className="text-lg font-medium">Email Provider</h3>
      <div className="flex gap-2">
        <Button
          variant={provider === "gmail" ? "default" : "outline"}
          onClick={() => {
            setProvider("gmail");
            updateUserEmailProvider("gmail");
          }}
        >
          Gmail
        </Button>
        <Button
          variant={provider === "outlook" ? "default" : "outline"}
          onClick={() => {
            setProvider("outlook");
            updateUserEmailProvider("outlook");
          }}
        >
          Outlook
        </Button>
      </div>
    </div>
  );
}
```

---

## Phase 4: Testing (Week 4)

### Task 4.1: Unit Tests

Create unit tests for both providers:

**File**: `apps/web/utils/microsoft/provider.test.ts` (new)

```typescript
import { describe, it, expect, vi } from "vitest";
import { MicrosoftGraphProvider } from "./provider";

describe("MicrosoftGraphProvider", () => {
  it("should list messages", async () => {
    // Mock Graph client
    // Test message listing
  });

  it("should send message", async () => {
    // Test message sending
  });

  // More tests...
});
```

---

### Task 4.2: Integration Tests

Test with real APIs using test accounts:

**File**: `apps/web/utils/email/provider.integration.test.ts` (new)

```typescript
import { describe, it, expect } from "vitest";
import { getEmailProvider } from "./factory";

describe("Email Provider Integration", () => {
  it("should work with Gmail", async () => {
    // Use test Gmail account
  });

  it("should work with Outlook", async () => {
    // Use test Outlook account
  });
});
```

---

### Task 4.3: Manual Testing Checklist

- [ ] Test Gmail OAuth flow
- [ ] Test Microsoft OAuth flow
- [ ] Test switching between providers
- [ ] Test sending email (Gmail)
- [ ] Test sending email (Outlook)
- [ ] Test receiving email (Gmail)
- [ ] Test receiving email (Outlook)
- [ ] Test folder operations (Gmail labels)
- [ ] Test folder operations (Outlook folders)
- [ ] Test search (Gmail)
- [ ] Test search (Outlook)
- [ ] Test AI features with Gmail
- [ ] Test AI features with Outlook
- [ ] Test bulk unsubscribe with Gmail
- [ ] Test bulk unsubscribe with Outlook
- [ ] Test webhook notifications (Gmail)
- [ ] Test webhook notifications (Outlook)
- [ ] Test token refresh (Gmail)
- [ ] Test token refresh (Outlook)
- [ ] Test with personal Microsoft account
- [ ] Test with work/school Microsoft account
- [ ] Test error handling for both providers

---

## Phase 5: Documentation and Deployment

### Task 5.1: Update Documentation

**Files to update**:
- [ ] `README.md` - Add Office365 setup instructions
- [ ] `apps/web/.env.example` - Add Microsoft environment variables
- [ ] Create `docs/OFFICE365_SETUP.md` - Detailed setup guide
- [ ] Update FAQ - Add Office365 support info

---

### Task 5.2: Update Landing Page

**File**: `apps/web/app/(landing)/home/FAQs.tsx`

Update FAQ entry:
```typescript
{
  question: "What email providers do you support?",
  answer: "We support Gmail, Google Workspace, and Office365/Outlook accounts. You can connect any of these email providers to use Inbox Zero."
}
```

---

### Task 5.3: Deployment Checklist

- [ ] Set up Azure AD app in production
- [ ] Add environment variables to production
- [ ] Run database migration in production
- [ ] Deploy to staging first
- [ ] Test on staging with real accounts
- [ ] Monitor error logs
- [ ] Deploy to production
- [ ] Monitor production metrics
- [ ] Announce feature launch

---

## Rollback Plan

If issues arise:

1. **Immediate**: Feature flag to disable Microsoft login
   ```typescript
   const ENABLE_MICROSOFT_AUTH = env.ENABLE_MICROSOFT_AUTH || false;
   ```

2. **Database**: Migration can be safely rolled back (only added nullable field)

3. **Code**: Git revert to previous version

4. **Users**: Existing Gmail users are unaffected

---

## Success Metrics

Track these metrics after launch:

1. **Adoption**:
   - % of new signups using Office365
   - Total Office365 users

2. **Technical**:
   - API error rate (Gmail vs Outlook)
   - Token refresh success rate
   - Webhook delivery rate

3. **User Experience**:
   - Time to first email load
   - Support tickets related to Office365
   - User satisfaction (NPS)

4. **Business**:
   - Conversion rate (trial → paid) for Office365 users
   - Retention rate for Office365 users

---

## Post-Launch Tasks

After successful launch:

- [ ] Monitor error logs for Office365-specific issues
- [ ] Collect user feedback via surveys
- [ ] Optimize performance based on metrics
- [ ] Add Office365 to marketing materials
- [ ] Write blog post announcing feature
- [ ] Update product tour/onboarding
- [ ] Consider adding more Microsoft features (Calendar, Contacts)

---

## Maintenance

Ongoing maintenance tasks:

- [ ] Monthly: Review Microsoft Graph API changelog
- [ ] Quarterly: Review Azure AD deprecations
- [ ] Update dependencies regularly
- [ ] Monitor for breaking changes in Graph API
- [ ] Keep NextAuth.js updated

---

## Future Enhancements

After Office365 support is stable:

1. **Multiple account support**: Let users connect both Gmail and Outlook
2. **Calendar integration**: Add calendar features using Graph API
3. **Teams integration**: Integrate with Microsoft Teams
4. **OneDrive**: Save attachments to OneDrive
5. **More providers**: Yahoo Mail, ProtonMail, etc.

---

## Conclusion

This implementation plan provides a detailed roadmap for adding Office365 support to Inbox Zero. The approach leverages existing architecture, minimizes risk, and provides a solid foundation for future multi-provider support.

**Total Estimated Timeline**: 4-5 weeks
**Total Estimated Effort**: 80-100 hours

The plan is designed to be executed incrementally with frequent testing and validation at each phase.
