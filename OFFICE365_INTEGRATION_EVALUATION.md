# Office365/Outlook Mailbox Integration - Technical Evaluation

## Executive Summary

This document evaluates different approaches for adding Office365/Outlook mailbox connection capabilities to Inbox Zero. After analyzing three distinct approaches, we recommend implementing a multi-provider architecture using Microsoft Graph API with NextAuth.js provider support.

## Current Architecture

Inbox Zero currently uses:
- **Authentication**: NextAuth.js with Google OAuth provider
- **Email API**: Google Gmail API (`@googleapis/gmail`)
- **Database**: PostgreSQL with Prisma ORM
- **Token Management**: Encrypted storage of OAuth tokens with refresh token rotation
- **Real-time Updates**: Google PubSub webhooks for push notifications

### Key Files:
- `apps/web/utils/auth.ts` - NextAuth.js configuration with Google OAuth
- `apps/web/utils/gmail/client.ts` - Gmail API client initialization
- `apps/web/utils/gmail/*.ts` - Gmail-specific operations (messages, threads, labels, etc.)
- `apps/web/prisma/schema.prisma` - Database schema with Account model

## Approach 1: Microsoft Graph API with NextAuth.js Provider

### Description
Extend the existing NextAuth.js setup by adding a Microsoft/Azure AD provider alongside the existing Google provider. Use Microsoft Graph API for email operations.

### Technical Details

**Authentication Flow:**
```typescript
// apps/web/utils/auth.ts
import AzureADProvider from "next-auth/providers/azure-ad";

const MICROSOFT_SCOPES = [
  "openid",
  "profile",
  "email",
  "offline_access",
  "Mail.ReadWrite",
  "Mail.Send",
  "MailboxSettings.Read",
  "User.Read"
];

providers: [
  GoogleProvider({ ... }),
  AzureADProvider({
    clientId: env.AZURE_AD_CLIENT_ID,
    clientSecret: env.AZURE_AD_CLIENT_SECRET,
    tenantId: "common", // or "organizations" for work/school accounts
    authorization: {
      params: {
        scope: MICROSOFT_SCOPES.join(" "),
      }
    }
  })
]
```

**API Client:**
- Use `@microsoft/microsoft-graph-client` for Graph API operations
- Create abstraction layer similar to `utils/gmail/` structure
- Implement provider-agnostic interfaces

**Database Changes:**
```prisma
model Account {
  // ... existing fields
  provider String  // "google" or "azure-ad"
}

model User {
  // ... existing fields
  emailProvider String? // "gmail" or "outlook"
}
```

### Pros
✅ **Minimal architectural changes** - Leverages existing NextAuth.js infrastructure
✅ **Native NextAuth.js support** - Azure AD provider is well-maintained
✅ **Consistent user experience** - Similar OAuth flow for both providers
✅ **Token management handled** - NextAuth.js manages refresh tokens automatically
✅ **Well-documented** - Extensive Microsoft Graph API documentation
✅ **Feature parity possible** - Graph API supports most Gmail API features
✅ **Real-time updates** - Supports webhooks via Microsoft Graph subscriptions
✅ **Single codebase** - Can share business logic between providers

### Cons
❌ **API differences** - Graph API has different request/response formats than Gmail API
❌ **Abstraction complexity** - Need to create provider-agnostic interfaces
❌ **Testing overhead** - Must test both providers
❌ **Migration path** - Existing users remain on Gmail
❌ **Feature gaps** - Some Gmail-specific features may not have Graph API equivalents
❌ **Azure AD setup** - Requires Azure AD app registration (additional setup)

### Implementation Effort
**Estimated: 3-4 weeks**
- 1 week: Auth setup and provider abstraction
- 1 week: Graph API client and core operations
- 1 week: Feature implementation (labels, filters, etc.)
- 1 week: Testing and bug fixes

### Key Dependencies
```json
{
  "@microsoft/microsoft-graph-client": "^3.0.7",
  "@microsoft/microsoft-graph-types": "^2.40.0",
  "next-auth": "^4.x (already installed)"
}
```

---

## Approach 2: IMAP/SMTP Protocol Implementation

### Description
Implement direct IMAP/SMTP support to work with any email provider (Office365, Gmail, self-hosted, etc.) using standard email protocols.

### Technical Details

**Authentication:**
- Store IMAP/SMTP credentials separately from OAuth providers
- Support both password and OAuth2 for IMAP/SMTP
- Microsoft supports OAuth2 for IMAP/SMTP

**Protocol Libraries:**
```typescript
import Imap from "imap";
import nodemailer from "nodemailer";

// IMAP connection
const imap = new Imap({
  user: "user@outlook.com",
  password: "password", // or OAuth2 token
  host: "outlook.office365.com",
  port: 993,
  tls: true,
  xoauth2: accessToken, // if using OAuth2
});

// SMTP sending
const transporter = nodemailer.createTransport({
  host: "smtp.office365.com",
  port: 587,
  secure: false,
  auth: {
    type: "OAuth2",
    user: "user@outlook.com",
    accessToken: accessToken,
  },
});
```

**Database Changes:**
```prisma
model EmailConnection {
  id            String   @id @default(cuid())
  userId        String
  user          User     @relation(fields: [userId], references: [id])
  provider      String   // "imap"
  host          String
  port          Int
  email         String
  credentials   String   @db.Text // encrypted
  createdAt     DateTime @default(now())
  updatedAt     DateTime @updatedAt
}
```

### Pros
✅ **Universal support** - Works with any email provider (Office365, Gmail, custom domains)
✅ **No API restrictions** - Not limited by API quotas or rate limits
✅ **Self-hosted friendly** - Works with any IMAP/SMTP server
✅ **Simple authentication** - Can use basic auth or OAuth2
✅ **Direct control** - Full control over email operations
✅ **Offline capabilities** - Can cache emails locally

### Cons
❌ **Protocol limitations** - IMAP/SMTP lack features like labels, filters, advanced search
❌ **Performance issues** - Slower than REST APIs, requires persistent connections
❌ **Complex state management** - Need to manage IMAP connection pooling
❌ **No real-time updates** - Must poll for new messages (IDLE extension available but limited)
❌ **Feature degradation** - Loss of provider-specific features (Gmail labels → Outlook folders)
❌ **Security concerns** - More complex to secure credential storage
❌ **Limited metadata** - Less rich email metadata compared to APIs
❌ **Compatibility issues** - Different IMAP server implementations vary
❌ **No webhook support** - Cannot receive push notifications
❌ **Development effort** - Must implement low-level protocol handling

### Implementation Effort
**Estimated: 5-6 weeks**
- 2 weeks: IMAP/SMTP client implementation and connection management
- 1 week: Email parsing and synchronization logic
- 1 week: Feature mapping (labels → folders, etc.)
- 1-2 weeks: Testing across different providers and edge cases

### Key Dependencies
```json
{
  "imap": "^0.8.19",
  "nodemailer": "^6.9.7",
  "mailparser": "^3.6.5",
  "imap-simple": "^5.1.0"
}
```

---

## Approach 3: Unified Email Abstraction Layer (Nylas/SendGrid/Similar)

### Description
Use a third-party email infrastructure service that provides unified APIs for multiple email providers. These services handle provider-specific complexities.

### Technical Details

**Example: Nylas Email API**
```typescript
import Nylas from "nylas";

const nylas = new Nylas({
  apiKey: env.NYLAS_API_KEY,
  apiUri: "https://api.us.nylas.com",
});

// User connects account
const { url, successUrl } = await nylas.auth.urlForAuthentication({
  clientId: env.NYLAS_CLIENT_ID,
  provider: "microsoft", // or "google"
  redirectUri: "https://app.com/oauth/callback",
  scopes: ["mail.read", "mail.send"],
});

// Unified API for all providers
const messages = await nylas.messages.list({
  identifier: grantId,
  queryParams: { limit: 10 },
});
```

**Alternative Services:**
- **Nylas**: Email, calendar, contacts API ($99-$399/month)
- **Microsoft Graph Data Connect**: Bulk data access (enterprise only)
- **CloudMaiL**: Email parsing service
- **Context.io**: Email sync and webhook service (deprecated)

### Pros
✅ **Rapid implementation** - Pre-built provider integrations
✅ **Unified interface** - Single API for Gmail, Office365, and others
✅ **Production-ready** - Battle-tested, scalable infrastructure
✅ **Webhook support** - Built-in real-time notifications
✅ **Feature abstraction** - Handles provider differences automatically
✅ **Reduced maintenance** - Provider updates handled by service
✅ **Additional features** - Calendar, contacts, scheduling APIs included
✅ **Security compliance** - SOC2, GDPR compliant

### Cons
❌ **Cost** - Significant monthly/annual fees ($99-$399+/month per tier)
❌ **Vendor lock-in** - Difficult to migrate away from service
❌ **Privacy concerns** - Third-party processes user emails
❌ **API limitations** - Constrained by service's API design
❌ **External dependency** - Service downtime affects your app
❌ **Data residency** - Email data flows through third-party servers
❌ **Feature delays** - Dependent on vendor for new features
❌ **Pricing scaling** - Costs increase with user base
❌ **Limited customization** - Cannot extend or modify service APIs
❌ **Trust factor** - Users may be uncomfortable with third-party email access

### Implementation Effort
**Estimated: 2-3 weeks**
- 1 week: Service integration and auth flow
- 1 week: Feature implementation using service APIs
- 1 week: Testing and migration planning

### Key Dependencies
```json
{
  "nylas": "^7.0.0"
  // or alternative service SDK
}
```

### Cost Analysis
**Nylas Pricing (example):**
- Starter: $99/month - 5 accounts
- Growth: $299/month - 50 accounts
- Scale: $399/month - 100 accounts
- Enterprise: Custom pricing

**Annual cost for 1000 active users:** $20,000 - $50,000+

---

## Comparative Analysis

| Criteria | Graph API (Approach 1) | IMAP/SMTP (Approach 2) | Third-Party (Approach 3) |
|----------|------------------------|------------------------|--------------------------|
| **Implementation Time** | 3-4 weeks | 5-6 weeks | 2-3 weeks |
| **Development Complexity** | Medium | High | Low |
| **Ongoing Maintenance** | Medium | High | Low |
| **Cost** | Free (Azure AD free tier) | Free | $1,200-$4,800+/year |
| **Feature Completeness** | 95% | 60% | 90% |
| **Performance** | Excellent | Good | Excellent |
| **Real-time Updates** | Yes (webhooks) | Limited (IDLE) | Yes (webhooks) |
| **Security** | OAuth2 | OAuth2 or basic auth | OAuth2 via provider |
| **Scalability** | Excellent | Good | Excellent |
| **Provider Control** | Direct | Direct | Indirect |
| **User Privacy** | Direct connection | Direct connection | Via third-party |
| **Future Flexibility** | High | Medium | Low |

---

## Recommendation: Approach 1 (Microsoft Graph API)

### Rationale

After evaluating all approaches, **Approach 1 (Microsoft Graph API with NextAuth.js)** is the recommended path forward for the following reasons:

1. **Architectural Fit**: Seamlessly extends existing NextAuth.js infrastructure without major refactoring
2. **Cost-Effective**: Zero additional costs (Azure AD free tier supports OAuth)
3. **Feature Parity**: Microsoft Graph API provides comparable features to Gmail API
4. **User Control**: Direct authentication, no third-party data processing
5. **Scalability**: Can handle growth without significant cost increases
6. **Maintainability**: Leverages well-documented, stable APIs
7. **Future-Proof**: Provides foundation for adding more providers (Yahoo, ProtonMail, etc.)

### When to Consider Alternatives

- **IMAP/SMTP (Approach 2)**: If you need to support custom/self-hosted email servers
- **Third-Party Service (Approach 3)**: If rapid time-to-market is critical and budget permits

---

## Implementation Plan for Approach 1

### Phase 1: Foundation (Week 1)
**Goal**: Set up authentication and basic infrastructure

- [ ] Register Azure AD application in Microsoft Azure Portal
- [ ] Configure OAuth consent screen and scopes
- [ ] Add environment variables for Azure AD credentials
- [ ] Extend NextAuth.js configuration to add Azure AD provider
- [ ] Update database schema to track email provider per user
- [ ] Create provider selection UI in onboarding/settings
- [ ] Test OAuth flow end-to-end

**Deliverables**:
- Users can authenticate with Microsoft accounts
- Database tracks which provider each user uses
- Basic provider selection interface

### Phase 2: Email Abstraction Layer (Week 2)
**Goal**: Create provider-agnostic interfaces

- [ ] Design interface for email operations (IEmailProvider)
- [ ] Implement Microsoft Graph API client (`utils/microsoft/client.ts`)
- [ ] Create Graph API wrapper for basic operations:
  - [ ] List messages (`utils/microsoft/message.ts`)
  - [ ] Get message details
  - [ ] Send message
  - [ ] Delete message
- [ ] Refactor existing Gmail code to implement IEmailProvider interface
- [ ] Create provider factory to return appropriate client
- [ ] Update API routes to use provider-agnostic interfaces

**Deliverables**:
- Abstract interfaces for email operations
- Both Gmail and Graph clients implement same interface
- Dynamic provider selection based on user's account

### Phase 3: Feature Implementation (Week 3)
**Goal**: Implement Office365-specific features

- [ ] Implement folder operations (equivalent to Gmail labels)
  - [ ] List folders
  - [ ] Create folders
  - [ ] Move messages between folders
- [ ] Implement filtering/rules (Inbox rules)
- [ ] Implement search functionality
- [ ] Implement thread operations
- [ ] Implement draft operations
- [ ] Implement contact operations
- [ ] Set up Microsoft Graph webhooks for real-time updates
- [ ] Implement webhook handler for Office365 notifications

**Deliverables**:
- Full feature parity between Gmail and Office365
- Real-time email updates for Office365 users
- Folder management UI

### Phase 4: Testing and Refinement (Week 4)
**Goal**: Ensure reliability and fix bugs

- [ ] Write unit tests for Graph API client
- [ ] Write integration tests for email operations
- [ ] Test with multiple Office365 account types:
  - [ ] Personal Microsoft accounts (outlook.com)
  - [ ] Work/School accounts (Azure AD)
  - [ ] Office365 business accounts
- [ ] Test token refresh and error handling
- [ ] Performance testing and optimization
- [ ] Update documentation
- [ ] Create migration guide for testing with Office365 accounts

**Deliverables**:
- Comprehensive test suite
- Bug-free experience for both providers
- Updated documentation

### Phase 5: AI Feature Compatibility (Week 5, if needed)
**Goal**: Ensure AI features work with Office365

- [ ] Test AI assistant with Office365 accounts
- [ ] Test bulk unsubscriber with Office365
- [ ] Test cold email blocker with Office365
- [ ] Test email analytics with Office365
- [ ] Adjust prompts/logic for provider differences if needed

**Deliverables**:
- All AI features work with Office365
- Provider-specific optimizations if needed

---

## Technical Architecture

### Provider Abstraction Interface

```typescript
// apps/web/utils/email/provider.ts
export interface IEmailProvider {
  // Message operations
  listMessages(options: ListMessagesOptions): Promise<Message[]>;
  getMessage(id: string): Promise<Message>;
  sendMessage(message: OutgoingMessage): Promise<void>;
  deleteMessage(id: string): Promise<void>;
  
  // Folder/Label operations
  listFolders(): Promise<Folder[]>;
  createFolder(name: string): Promise<Folder>;
  moveToFolder(messageId: string, folderId: string): Promise<void>;
  
  // Thread operations
  getThread(id: string): Promise<Thread>;
  
  // Search
  search(query: string): Promise<Message[]>;
  
  // Real-time updates
  setupWebhook(callbackUrl: string): Promise<void>;
  refreshWebhook(): Promise<void>;
}

// Provider factory
export function getEmailProvider(user: User): IEmailProvider {
  switch (user.emailProvider) {
    case "gmail":
      return new GmailProvider(user);
    case "outlook":
      return new MicrosoftGraphProvider(user);
    default:
      throw new Error(`Unsupported provider: ${user.emailProvider}`);
  }
}
```

### Directory Structure

```
apps/web/utils/
├── email/
│   ├── provider.ts          # Interface definitions
│   ├── factory.ts           # Provider factory
│   └── types.ts            # Shared types
├── gmail/                   # Gmail implementation (existing)
│   ├── client.ts
│   ├── message.ts
│   └── ...
└── microsoft/              # Microsoft Graph implementation (new)
    ├── client.ts
    ├── message.ts
    ├── folder.ts
    ├── thread.ts
    └── webhook.ts
```

---

## Environment Variables

Add to `apps/web/.env`:

```bash
# Microsoft/Azure AD OAuth
AZURE_AD_CLIENT_ID=
AZURE_AD_CLIENT_SECRET=
AZURE_AD_TENANT_ID=common  # or specific tenant ID

# Microsoft Graph API
MICROSOFT_GRAPH_WEBHOOK_VALIDATION_TOKEN=
```

Add to `apps/web/env.ts`:

```typescript
export const env = createEnv({
  server: {
    // ... existing vars
    AZURE_AD_CLIENT_ID: z.string().min(1),
    AZURE_AD_CLIENT_SECRET: z.string().min(1),
    AZURE_AD_TENANT_ID: z.string().default("common"),
    MICROSOFT_GRAPH_WEBHOOK_VALIDATION_TOKEN: z.string().optional(),
  },
  // ...
});
```

---

## Database Migrations

```prisma
// apps/web/prisma/schema.prisma

model User {
  // ... existing fields
  emailProvider String? @default("gmail") // "gmail" or "outlook"
}

// No changes needed to Account model - provider field already exists
```

Migration:
```bash
pnpm prisma migrate dev --name add_email_provider_to_user
```

---

## Dependencies to Add

```bash
pnpm add @microsoft/microsoft-graph-client @microsoft/microsoft-graph-types isomorphic-fetch
```

Update `apps/web/package.json`:
```json
{
  "dependencies": {
    "@microsoft/microsoft-graph-client": "^3.0.7",
    "@microsoft/microsoft-graph-types": "^2.40.0",
    "isomorphic-fetch": "^3.0.0"
  }
}
```

---

## Security Considerations

1. **Token Storage**: Continue using encrypted storage for OAuth tokens
2. **Scopes**: Request minimal necessary scopes
3. **Webhook Validation**: Validate Microsoft Graph webhook signatures
4. **Error Handling**: Properly handle token expiration and refresh
5. **Rate Limiting**: Implement rate limiting for API calls
6. **Audit Logging**: Log provider-specific operations for debugging

---

## Testing Strategy

1. **Unit Tests**: Test provider implementations independently
2. **Integration Tests**: Test with real APIs (test accounts)
3. **E2E Tests**: Test full user flows with both providers
4. **Provider Comparison Tests**: Ensure feature parity
5. **Performance Tests**: Benchmark API response times
6. **Error Scenario Tests**: Test token refresh, rate limits, etc.

---

## Rollout Strategy

### Beta Testing (Week 5-6)
- [ ] Deploy to staging environment
- [ ] Invite 10-20 beta testers with Office365 accounts
- [ ] Collect feedback and fix critical issues
- [ ] Monitor error rates and performance

### Gradual Rollout (Week 7-8)
- [ ] Add Office365 option to production
- [ ] Monitor new user signups
- [ ] Keep Gmail as default for existing users
- [ ] Add banner promoting Office365 support

### Full Launch (Week 9+)
- [ ] Announce Office365 support in marketing materials
- [ ] Update FAQ and documentation
- [ ] Blog post about the new feature
- [ ] Monitor support tickets and feedback

---

## Success Metrics

- **Adoption Rate**: % of new users choosing Office365
- **Feature Parity**: % of features working on both providers
- **Error Rate**: API error rate for Office365 vs Gmail
- **Performance**: Response time comparison
- **User Satisfaction**: NPS score from Office365 users

---

## Future Considerations

After successfully implementing Office365 support, the abstraction layer makes it easier to add:

1. **Yahoo Mail**: Using Yahoo Mail API
2. **ProtonMail**: Using ProtonMail Bridge or API
3. **FastMail**: Using JMAP protocol
4. **Self-hosted**: Using IMAP/SMTP fallback
5. **Multiple Accounts**: Users can connect multiple email accounts

---

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| API differences cause bugs | High | Comprehensive testing, staged rollout |
| Token refresh issues | Medium | Robust error handling, monitoring |
| Feature gaps in Graph API | Medium | Document limitations, find workarounds |
| Increased support burden | Medium | Comprehensive docs, FAQ updates |
| Development timeline slips | Low | Phased approach, MVP first |

---

## Conclusion

Implementing Office365 support via Microsoft Graph API (Approach 1) provides the best balance of:
- **Technical feasibility**: Extends existing architecture naturally
- **Cost effectiveness**: No additional service fees
- **Feature completeness**: Achieves ~95% parity with Gmail
- **User privacy**: Direct authentication without third-party data access
- **Future flexibility**: Foundation for multi-provider support

The estimated 4-5 week implementation timeline is reasonable and provides significant value to users who prefer or require Office365/Outlook for work.

**Next Steps**: Review and approve this plan, then proceed with Phase 1 implementation.

---

## References

- [Microsoft Graph API Documentation](https://learn.microsoft.com/en-us/graph/overview)
- [NextAuth.js Azure AD Provider](https://next-auth.js.org/providers/azure-ad)
- [Microsoft Graph Email API](https://learn.microsoft.com/en-us/graph/api/resources/message)
- [Microsoft Graph Webhooks](https://learn.microsoft.com/en-us/graph/webhooks)
- [Gmail API Reference](https://developers.google.com/gmail/api/reference/rest)
