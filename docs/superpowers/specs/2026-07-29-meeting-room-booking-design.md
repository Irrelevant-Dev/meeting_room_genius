# Design Specification: M365 Meeting Room Booking Application (Meeting Room Genius)

## 1. Overview & Objectives
Meeting Room Genius is a high-performance, modern web application that interfaces directly with Microsoft 365 / Exchange Online via Microsoft Graph API. It empowers corporate employees to discover office meeting rooms, view real-time free/busy availability timelines, and book room reservations on-the-fly outside of the standard Outlook desktop/web client.

The application is built on Next.js 14/15 (App Router, Server Components & API Routes), TypeScript, Tailwind CSS, Lucide React, and NextAuth.js v5 (Auth.js) with Microsoft Entra ID. It is designed for seamless containerized deployment on Railway.

---

## 2. Authentication & Authorization Flow

### 2.1 Microsoft Entra ID Integration
* **Protocol:** OAuth 2.0 OpenID Connect with Authorization Code Flow + PKCE.
* **Framework Provider:** NextAuth.js v5 (Auth.js) configured with Entra ID (Azure AD) provider.
* **Delegated Scopes:**
  * `openid`
  * `profile`
  * `email`
  * `offline_access` (enables refresh token acquisition)
  * `User.Read` (user profile context)
  * `Calendars.ReadWrite` (delegated event creation and calendar scheduling)
  * `Places.Read.All` (room metadata discovery, with fallback to `findRooms`)

### 2.2 Token Management & Session Refresh
* **Encrypted JWT Session:** Access token, refresh token, and token expiration epoch (`expiresAt`) are stored inside encrypted HTTP-only session cookies.
* **Auto-Refresh Mechanism:**
  * Prior to forwarding requests to Microsoft Graph API in server routes, session validity is verified against `expiresAt`.
  * If the access token is near expiration (<= 5 minutes remaining), `lib/auth.ts` issues a background OAuth refresh request to `https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token` using `grant_type=refresh_token`.
  * The newly issued access and refresh tokens update the user session transparently.

---

## 3. Core Architecture & Component Structure

```
meeting_room_genius/
├── app/
│   ├── api/
│   │   ├── auth/[...nextauth]/route.ts  # Auth.js Entra ID handler
│   │   ├── rooms/
│   │   │   ├── route.ts                 # GET /api/rooms (Room listing)
│   │   │   └── schedule/route.ts        # POST /api/rooms/schedule (Free/busy blocks)
│   │   └── bookings/route.ts            # POST /api/bookings (Event creation)
│   ├── dashboard/
│   │   ├── page.tsx                     # Main room grid & timeline dashboard
│   │   └── loading.tsx                  # Skeleton loader
│   ├── layout.tsx                       # Root layout with SessionProvider and Toast notifications
│   ├── page.tsx                         # Landing page with corporate SSO login CTA
│   └── globals.css                      # Modern CSS custom tokens, glassmorphism, animations
├── components/
│   ├── auth/
│   │   ├── UserNav.tsx                  # Logged-in user badge & logout control
│   │   └── LoginButton.tsx              # Microsoft SSO trigger button
│   ├── rooms/
│   │   ├── RoomCard.tsx                 # Room details, capacity pill, features & status indicator
│   │   ├── RoomGrid.tsx                 # Grid container with search, building & capacity filters
│   │   ├── RoomTimeline.tsx             # Interactive 30-min free/busy time-bar visualizer
│   │   ├── DatePickerHeader.tsx         # Date quick-nav (Today, Tomorrow, custom picker)
│   │   └── BookingModal.tsx             # Room reservation dialog with subject, time, description
│   └── ui/                              # Primitives: Modal, Toast, Skeleton, Button, Input
├── lib/
│   ├── auth.ts                          # NextAuth configuration and token refresh callbacks
│   ├── env.ts                           # Zod environment variable validation
│   ├── graph/
│   │   ├── client.ts                    # Microsoft Graph Client SDK initializer with Retry handler
│   │   ├── rooms.ts                     # Graph query helper for Places / findRooms
│   │   ├── availability.ts             # Graph query helper for getSchedule
│   │   └── booking.ts                  # Graph query helper for creating calendar events
│   └── types/
│       ├── room.ts                      # Interfaces for Room, Equipment, ScheduleItem
│       └── booking.ts                   # Interfaces for Event payloads and FreeBusy schedules
└── public/                              # Branding assets
```

---

## 4. Microsoft Graph API Endpoint Mappings

| Feature | Graph API Endpoint | Method | Key Parameters & Body |
| :--- | :--- | :--- | :--- |
| **Room Listing** | `/v1.0/places/microsoft.graph.room` | `GET` | Select: `id, displayName, emailAddress, capacity, building, floorNumber, phone` (Fallback: `/v1.0/me/findRooms`) |
| **Free/Busy Schedule** | `/v1.0/users/{user-id\|me}/calendar/getSchedule` | `POST` | `schedules: [email1, email2]`, `startTime`, `endTime`, `availabilityViewInterval: 30` |
| **Create Booking** | `/v1.0/me/events` | `POST` | `subject`, `start`, `end`, `location`, `attendees: [{emailAddress: roomEmail, type: "resource"}]` |

---

## 5. UI/UX Aesthetics & Design System

* **Theme:** Sleek, high-contrast dark mode with glassmorphic cards, subtle borders (`border-white/10`), and vibrant status indicators.
* **Palette:**
  * Background: Deep Slate `#090D16` / Obsidian Card Surfaces `#111827`
  * Accents: Cyan Glow `#06B6D4`, Emerald Available `#10B981`, Amber Pending `#F59E0B`, Rose Occupied `#EF4444`
* **Typography:** `Inter` & `Outfit` (via Google Fonts).
* **Interactive Experience:**
  * Real-time free/busy timeline visualization with clickable 30-minute availability blocks.
  * Instant room search by name, building, floor, minimum capacity, or feature tags (AV Screen, Whiteboard, Video Conf).
  * Smooth modal dialog for reservation entry with live validation and toast status feedback.

---

## 6. Resiliency & Error Handling Strategy

1. **Graph API Rate Limits (HTTP 429):**
   * Handled by custom Graph client middleware reading `Retry-After` headers and retrying with exponential backoff up to 3 attempts.
2. **Exchange Booking Conflicts:**
   * When an event is created, Exchange Online evaluates room availability. If auto-decline occurs or a double-booking conflict occurs, the booking engine parses the response status and notifies the user with actionable feedback.
3. **Environment Validation:**
   * Strict Zod schema (`lib/env.ts`) evaluates mandatory variables (`AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`, `AZURE_TENANT_ID`, `NEXTAUTH_SECRET`, `NEXTAUTH_URL`) at startup, preventing runtime crashes.

---

## 7. Deployment Configuration

* **Platform:** Railway
* **Runtime:** Node.js 20 LTS
* **Build Target:** Next.js Standalone Output (`output: 'standalone'` in `next.config.mjs`)
* **Environment Variables:** Production Entra ID App Registration credentials and HTTPS redirect URI.
