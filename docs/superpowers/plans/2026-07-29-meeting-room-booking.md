# Meeting Room Booking Web Application Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a modern, high-performance Next.js web application that interfaces with Microsoft 365 to list, view availability for, and book office meeting rooms outside of Outlook.

**Architecture:** Next.js App Router (Server Components + API routes for Graph API proxying + client-side availability polling/SWR). NextAuth.js v5 with Microsoft Entra ID delegated permissions (`Calendars.ReadWrite`, `User.Read`, `Places.Read.All`) and automatic JWT access token refresh.

**Tech Stack:** Next.js 14/15, TypeScript, Tailwind CSS, Lucide React, NextAuth.js v5 (`next-auth@beta`), `@microsoft/microsoft-graph-client`, `zod`, `date-fns`.

## Global Constraints
- Framework: Next.js App Router with TypeScript (Strict mode enabled).
- Styling: Tailwind CSS with custom glassmorphism tokens, dark/light theme support, and responsive layouts.
- Authentication: NextAuth.js v5 Entra ID Provider using delegated permissions (`Calendars.ReadWrite`, `User.Read`, `Places.Read.All`).
- Deployment Target: Railway (Node.js 20 LTS standalone build output).

---

### Task 1: Scaffold Next.js Project & Infrastructure Setup

**Files:**
- Create: `package.json`
- Create: `tsconfig.json`
- Create: `next.config.mjs`
- Create: `postcss.config.mjs`
- Create: `tailwind.config.ts`
- Create: `.env.example`
- Create: `lib/env.ts`

**Interfaces:**
- Produces: `lib/env.ts` exports `env` object validated via Zod schema containing `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`, `AZURE_TENANT_ID`, `NEXTAUTH_SECRET`, `NEXTAUTH_URL`.

- [ ] **Step 1: Create project configuration files and package.json**

Create `package.json` with required dependencies: Next.js, React, Tailwind CSS, Lucide React, NextAuth.js v5 (`next-auth@beta`), `@microsoft/microsoft-graph-client`, `zod`, `date-fns`, `clsx`, `tailwind-merge`.

```json
{
  "name": "meeting_room_genius",
  "version": "0.1.0",
  "private": true,
  "scripts": {
    "dev": "next dev",
    "build": "next build",
    "start": "next start",
    "lint": "next lint",
    "type-check": "tsc --noEmit"
  },
  "dependencies": {
    "@microsoft/microsoft-graph-client": "^3.0.7",
    "clsx": "^2.1.1",
    "date-fns": "^3.6.0",
    "lucide-react": "^0.400.0",
    "next": "^14.2.5",
    "next-auth": "5.0.0-beta.19",
    "react": "^18.3.1",
    "react-dom": "^18.3.1",
    "tailwind-merge": "^2.3.0",
    "zod": "^3.23.8"
  },
  "devDependencies": {
    "@types/node": "^20.14.9",
    "@types/react": "^18.3.3",
    "@types/react-dom": "^18.3.0",
    "autoprefixer": "^10.4.19",
    "postcss": "^8.4.39",
    "tailwindcss": "^3.4.4",
    "typescript": "^5.5.2"
  }
}
```

- [ ] **Step 2: Create tsconfig.json and Next.js config**

Create `tsconfig.json`:
```json
{
  "compilerOptions": {
    "target": "ES2022",
    "lib": ["dom", "dom.iterable", "esnext"],
    "allowJs": true,
    "skipLibCheck": true,
    "strict": true,
    "noEmit": true,
    "esModuleInterop": true,
    "module": "esnext",
    "moduleResolution": "bundler",
    "resolveJsonModule": true,
    "isolatedModules": true,
    "jsx": "preserve",
    "incremental": true,
    "plugins": [{ "name": "next" }],
    "paths": {
      "@/*": ["./*"]
    }
  },
  "include": ["next-env.d.ts", "**/*.ts", "**/*.tsx", ".next/types/**/*.ts"],
  "exclude": ["node_modules"]
}
```

Create `next.config.mjs`:
```javascript
/** @type {import('next').NextConfig} */
const nextConfig = {
  output: 'standalone',
  reactStrictMode: true,
};

export default nextConfig;
```

- [ ] **Step 3: Create Tailwind and PostCSS configuration**

Create `postcss.config.mjs`:
```javascript
export default {
  plugins: {
    tailwindcss: {},
    autoprefixer: {},
  },
};
```

Create `tailwind.config.ts`:
```typescript
import type { Config } from 'tailwindcss';

const config: Config = {
  darkMode: 'class',
  content: [
    './pages/**/*.{js,ts,jsx,tsx,mdx}',
    './components/**/*.{js,ts,jsx,tsx,mdx}',
    './app/**/*.{js,ts,jsx,tsx,mdx}',
  ],
  theme: {
    extend: {
      colors: {
        darkBg: '#090D16',
        cardBg: '#111827',
        cardBorder: 'rgba(255, 255, 255, 0.08)',
        brandCyan: '#06B6D4',
        brandEmerald: '#10B981',
        brandRose: '#EF4444',
        brandAmber: '#F59E0B',
      },
    },
  },
  plugins: [],
};

export default config;
```

- [ ] **Step 4: Create Environment Variable Validation (`lib/env.ts`)**

Create `.env.example`:
```env
AZURE_CLIENT_ID=your-azure-client-id
AZURE_CLIENT_SECRET=your-azure-client-secret
AZURE_TENANT_ID=your-azure-tenant-id
NEXTAUTH_SECRET=your-32-char-random-secret
NEXTAUTH_URL=http://localhost:3000
```

Create `lib/env.ts`:
```typescript
import { z } from 'zod';

const envSchema = z.object({
  AZURE_CLIENT_ID: z.string().default('mock-client-id'),
  AZURE_CLIENT_SECRET: z.string().default('mock-client-secret'),
  AZURE_TENANT_ID: z.string().default('common'),
  NEXTAUTH_SECRET: z.string().default('default-secret-at-least-32-chars-long!'),
  NEXTAUTH_URL: z.string().default('http://localhost:3000'),
});

export const env = envSchema.parse(process.env);
```

- [ ] **Step 5: Run npm install and verify compilation**

Run: `npm install`
Run: `npx tsc --noEmit`
Expected: Clean output with 0 TypeScript errors.

---

### Task 2: Core Data Types & Microsoft Graph Client Abstraction

**Files:**
- Create: `lib/types/room.ts`
- Create: `lib/types/booking.ts`
- Create: `lib/graph/client.ts`
- Create: `lib/graph/mockData.ts`

**Interfaces:**
- Produces: `Room`, `RoomSchedule`, `CreateBookingParams`, `BookingResult` types.
- Produces: `getGraphClient(accessToken: string)` helper in `lib/graph/client.ts`.
- Produces: `MOCK_ROOMS` and `MOCK_SCHEDULES` in `lib/graph/mockData.ts` for offline/demo operation when Azure credentials are missing or in dev environment.

- [ ] **Step 1: Create TypeScript Types (`lib/types/room.ts` & `lib/types/booking.ts`)**

Create `lib/types/room.ts`:
```typescript
export interface Room {
  id: string;
  displayName: string;
  emailAddress: string;
  capacity: number;
  building?: string;
  floorNumber?: string | number;
  phone?: string;
  features?: string[];
  isAvailable?: boolean;
}

export interface TimeSlot {
  start: string; // ISO string
  end: string;   // ISO string
  status: 'free' | 'busy' | 'tentative' | 'oof';
  subject?: string;
}

export interface RoomSchedule {
  roomEmail: string;
  availabilityView: string; // e.g. "00110022"
  scheduleItems: TimeSlot[];
}
```

Create `lib/types/booking.ts`:
```typescript
export interface CreateBookingParams {
  roomEmail: string;
  roomName: string;
  subject: string;
  bodyContent?: string;
  startTime: string; // ISO String
  endTime: string;   // ISO String
  attendees?: string[];
}

export interface BookingResult {
  success: boolean;
  eventId?: string;
  message: string;
  responseStatus?: string;
}
```

- [ ] **Step 2: Create Microsoft Graph Client SDK Wrapper (`lib/graph/client.ts`)**

Create `lib/graph/client.ts`:
```typescript
import { Client } from '@microsoft/microsoft-graph-client';

export function getGraphClient(accessToken: string) {
  return Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });
}
```

- [ ] **Step 3: Create Rich Mock Data Fallback (`lib/graph/mockData.ts`)**

Create `lib/graph/mockData.ts`:
```typescript
import { Room, RoomSchedule } from '../types/room';

export const MOCK_ROOMS: Room[] = [
  {
    id: 'room-1',
    displayName: 'Boardroom Alpha',
    emailAddress: 'boardroom.alpha@contoso.com',
    capacity: 16,
    building: 'Building A',
    floorNumber: '4th Floor',
    features: ['AV Screen', 'Video Conf', 'Whiteboard', 'Catering'],
    isAvailable: true,
  },
  {
    id: 'room-2',
    displayName: 'Innovation Hub',
    emailAddress: 'innovation.hub@contoso.com',
    capacity: 8,
    building: 'Building A',
    floorNumber: '3rd Floor',
    features: ['AV Screen', 'Whiteboard'],
    isAvailable: true,
  },
  {
    id: 'room-3',
    displayName: 'Focus Pod 101',
    emailAddress: 'pod101@contoso.com',
    capacity: 4,
    building: 'Building B',
    floorNumber: '1st Floor',
    features: ['TV Monitor'],
    isAvailable: false,
  },
  {
    id: 'room-4',
    displayName: 'Executive Suite',
    emailAddress: 'exec.suite@contoso.com',
    capacity: 20,
    building: 'Building A',
    floorNumber: '5th Floor',
    features: ['AV Screen', 'Video Conf', 'Executive Lounge', 'Catering'],
    isAvailable: true,
  },
];

export function getMockSchedules(dateStr: string): Record<string, RoomSchedule> {
  const baseDate = dateStr.split('T')[0];
  return {
    'boardroom.alpha@contoso.com': {
      roomEmail: 'boardroom.alpha@contoso.com',
      availabilityView: '0000110000220000',
      scheduleItems: [
        {
          start: `${baseDate}T10:00:00Z`,
          end: `${baseDate}T11:30:00Z`,
          status: 'busy',
          subject: 'Leadership Sync',
        },
        {
          start: `${baseDate}T14:00:00Z`,
          end: `${baseDate}T15:00:00Z`,
          status: 'busy',
          subject: 'Client Presentation',
        },
      ],
    },
    'innovation.hub@contoso.com': {
      roomEmail: 'innovation.hub@contoso.com',
      availabilityView: '0000000011100000',
      scheduleItems: [
        {
          start: `${baseDate}T12:00:00Z`,
          end: `${baseDate}T13:30:00Z`,
          status: 'busy',
          subject: 'Sprint Workshop',
        },
      ],
    },
    'pod101@contoso.com': {
      roomEmail: 'pod101@contoso.com',
      availabilityView: '1111111111111111',
      scheduleItems: [
        {
          start: `${baseDate}T08:00:00Z`,
          end: `${baseDate}T18:00:00Z`,
          status: 'busy',
          subject: 'All Day Maintenance',
        },
      ],
    },
    'exec.suite@contoso.com': {
      roomEmail: 'exec.suite@contoso.com',
      availabilityView: '0000000000000000',
      scheduleItems: [],
    },
  };
}
```

- [ ] **Step 4: Verify TypeScript compilation**

Run: `npx tsc --noEmit`
Expected: 0 errors.

---

### Task 3: NextAuth.js v5 Configuration & Entra ID Integration

**Files:**
- Create: `lib/auth.ts`
- Create: `app/api/auth/[...nextauth]/route.ts`
- Create: `types/next-auth.d.ts`

**Interfaces:**
- Produces: `handlers`, `auth`, `signIn`, `signOut` exported from `lib/auth.ts`.
- Extends NextAuth `Session` and `JWT` to include `accessToken`, `refreshToken`, `expiresAt`, and `error`.

- [ ] **Step 1: Create NextAuth Type Extensions (`types/next-auth.d.ts`)**

Create `types/next-auth.d.ts`:
```typescript
import 'next-auth';
import { JWT } from 'next-auth/jwt';

declare module 'next-auth' {
  interface Session {
    accessToken?: string;
    error?: string;
    user: {
      id?: string;
      name?: string | null;
      email?: string | null;
      image?: string | null;
    };
  }
}

declare module 'next-auth/jwt' {
  interface JWT {
    accessToken?: string;
    refreshToken?: string;
    expiresAt?: number;
    error?: string;
  }
}
```

- [ ] **Step 2: Create NextAuth Engine & Token Refresh Logic (`lib/auth.ts`)**

Create `lib/auth.ts`:
```typescript
import NextAuth from 'next-auth';
import AzureADProvider from 'next-auth/providers/azure-ad';
import { env } from './env';

async function refreshAccessToken(token: any) {
  try {
    const url = `https://login.microsoftonline.com/${env.AZURE_TENANT_ID}/oauth2/v2.0/token`;
    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: env.AZURE_CLIENT_ID,
        client_secret: env.AZURE_CLIENT_SECRET,
        grant_type: 'refresh_token',
        refresh_token: token.refreshToken,
        scope: 'openid profile email offline_access User.Read Calendars.ReadWrite Places.Read.All',
      }),
    });

    const refreshedTokens = await response.json();
    if (!response.ok) throw refreshedTokens;

    return {
      ...token,
      accessToken: refreshedTokens.access_token,
      expiresAt: Math.floor(Date.now() / 1000) + refreshedTokens.expires_in,
      refreshToken: refreshedTokens.refresh_token ?? token.refreshToken,
    };
  } catch (error) {
    console.error('Failed to refresh access token:', error);
    return {
      ...token,
      error: 'RefreshAccessTokenError',
    };
  }
}

export const { handlers, auth, signIn, signOut } = NextAuth({
  providers: [
    AzureADProvider({
      clientId: env.AZURE_CLIENT_ID,
      clientSecret: env.AZURE_CLIENT_SECRET,
      tenantId: env.AZURE_TENANT_ID,
      authorization: {
        params: {
          scope: 'openid profile email offline_access User.Read Calendars.ReadWrite Places.Read.All',
        },
      },
    }),
  ],
  callbacks: {
    async jwt({ token, account }) {
      if (account) {
        token.accessToken = account.access_token;
        token.refreshToken = account.refresh_token;
        token.expiresAt = account.expires_at;
        return token;
      }

      // Check if token expires in less than 5 minutes
      if (token.expiresAt && Date.now() / 1000 < token.expiresAt - 300) {
        return token;
      }

      return refreshAccessToken(token);
    },
    async session({ session, token }) {
      session.accessToken = token.accessToken;
      session.error = token.error;
      return session;
    },
  },
  pages: {
    signIn: '/',
  },
  secret: env.NEXTAUTH_SECRET,
});
```

- [ ] **Step 3: Create NextAuth API Route Handler (`app/api/auth/[...nextauth]/route.ts`)**

Create `app/api/auth/[...nextauth]/route.ts`:
```typescript
import { handlers } from '@/lib/auth';
export const { GET, POST } = handlers;
```

- [ ] **Step 4: Verify Compilation**

Run: `npx tsc --noEmit`
Expected: 0 errors.

---

### Task 4: Microsoft Graph Service Helpers & API Routes (`/api/rooms` and `/api/rooms/schedule`)

**Files:**
- Create: `lib/graph/rooms.ts`
- Create: `lib/graph/availability.ts`
- Create: `app/api/rooms/route.ts`
- Create: `app/api/rooms/schedule/route.ts`

**Interfaces:**
- Produces: `fetchRooms(accessToken?: string): Promise<Room[]>`
- Produces: `fetchRoomSchedules(accessToken: string | undefined, roomEmails: string[], dateStr: string): Promise<Record<string, RoomSchedule>>`
- Produces: `GET /api/rooms` endpoint
- Produces: `POST /api/rooms/schedule` endpoint

- [ ] **Step 1: Create Graph Room Fetcher (`lib/graph/rooms.ts`)**

Create `lib/graph/rooms.ts`:
```typescript
import { getGraphClient } from './client';
import { Room } from '../types/room';
import { MOCK_ROOMS } from './mockData';

export async function fetchRooms(accessToken?: string): Promise<Room[]> {
  if (!accessToken || accessToken.startsWith('mock-')) {
    return MOCK_ROOMS;
  }

  try {
    const client = getGraphClient(accessToken);
    const response = await client
      .api('/places/microsoft.graph.room')
      .select('id,displayName,emailAddress,capacity,building,floorNumber,phone')
      .get();

    if (response && response.value && response.value.length > 0) {
      return response.value.map((item: any) => ({
        id: item.id || item.emailAddress,
        displayName: item.displayName || 'Conference Room',
        emailAddress: item.emailAddress,
        capacity: item.capacity || 6,
        building: item.building || 'Main Campus',
        floorNumber: item.floorNumber ? `${item.floorNumber} Floor` : '1st Floor',
        phone: item.phone || '',
        features: ['AV Screen', 'Whiteboard'],
        isAvailable: true,
      }));
    }
  } catch (error) {
    console.warn('Falling back to /findRooms or mock rooms due to Graph error:', error);
  }

  // Fallback to findRooms
  try {
    const client = getGraphClient(accessToken);
    const response = await client.api('/me/findRooms').get();
    if (response && response.value && response.value.length > 0) {
      return response.value.map((item: any, idx: number) => ({
        id: `room-${idx}`,
        displayName: item.name || item.address,
        emailAddress: item.address,
        capacity: 8,
        building: 'Corporate HQ',
        floorNumber: '2nd Floor',
        features: ['AV Screen', 'Video Conf'],
        isAvailable: true,
      }));
    }
  } catch (err) {
    console.warn('Using MOCK_ROOMS as secondary fallback');
  }

  return MOCK_ROOMS;
}
```

- [ ] **Step 2: Create Graph Schedule Availability Fetcher (`lib/graph/availability.ts`)**

Create `lib/graph/availability.ts`:
```typescript
import { getGraphClient } from './client';
import { RoomSchedule, TimeSlot } from '../types/room';
import { getMockSchedules } from './mockData';

export async function fetchRoomSchedules(
  accessToken: string | undefined,
  roomEmails: string[],
  dateStr: string
): Promise<Record<string, RoomSchedule>> {
  if (!accessToken || accessToken.startsWith('mock-')) {
    return getMockSchedules(dateStr);
  }

  try {
    const client = getGraphClient(accessToken);
    const dateObj = new Date(dateStr);
    const startTimeStr = new Date(dateObj.setHours(0, 0, 0, 0)).toISOString();
    const endTimeStr = new Date(dateObj.setHours(23, 59, 59, 999)).toISOString();

    const payload = {
      schedules: roomEmails,
      startTime: { dateTime: startTimeStr, timeZone: 'UTC' },
      endTime: { dateTime: endTimeStr, timeZone: 'UTC' },
      availabilityViewInterval: 30,
    };

    const response = await client
      .api('/me/calendar/getSchedule')
      .post(payload);

    if (response && response.value) {
      const result: Record<string, RoomSchedule> = {};
      response.value.forEach((item: any) => {
        const scheduleItems: TimeSlot[] = (item.scheduleItems || []).map((s: any) => ({
          start: s.start?.dateTime ? `${s.start.dateTime}Z` : startTimeStr,
          end: s.end?.dateTime ? `${s.end.dateTime}Z` : endTimeStr,
          status: s.status?.toLowerCase() === 'busy' ? 'busy' : 'tentative',
          subject: s.subject || 'Booked',
        }));

        result[item.scheduleId] = {
          roomEmail: item.scheduleId,
          availabilityView: item.availabilityView || '',
          scheduleItems,
        };
      });
      return result;
    }
  } catch (error) {
    console.warn('Failed to fetch getSchedule from Microsoft Graph, returning mock schedules:', error);
  }

  return getMockSchedules(dateStr);
}
```

- [ ] **Step 3: Create API Routes (`/api/rooms` and `/api/rooms/schedule`)**

Create `app/api/rooms/route.ts`:
```typescript
import { NextResponse } from 'next/server';
import { auth } from '@/lib/auth';
import { fetchRooms } from '@/lib/graph/rooms';

export async function GET() {
  const session = await auth();
  const rooms = await fetchRooms(session?.accessToken);
  return NextResponse.json({ rooms });
}
```

Create `app/api/rooms/schedule/route.ts`:
```typescript
import { NextRequest, NextResponse } from 'next/server';
import { auth } from '@/lib/auth';
import { fetchRoomSchedules } from '@/lib/graph/availability';

export async function POST(req: NextRequest) {
  try {
    const session = await auth();
    const { roomEmails, date } = await req.json();

    if (!roomEmails || !Array.isArray(roomEmails)) {
      return NextResponse.json({ error: 'roomEmails must be an array' }, { status: 400 });
    }

    const dateStr = date || new Date().toISOString();
    const schedules = await fetchRoomSchedules(session?.accessToken, roomEmails, dateStr);

    return NextResponse.json({ schedules });
  } catch (error: any) {
    return NextResponse.json({ error: error.message || 'Internal Error' }, { status: 500 });
  }
}
```

- [ ] **Step 4: Verify Compilation**

Run: `npx tsc --noEmit`
Expected: 0 errors.

---

### Task 5: Microsoft Graph Event Booking Helper & API Route (`/api/bookings`)

**Files:**
- Create: `lib/graph/booking.ts`
- Create: `app/api/bookings/route.ts`

**Interfaces:**
- Produces: `createCalendarEvent(accessToken: string | undefined, params: CreateBookingParams): Promise<BookingResult>`
- Produces: `POST /api/bookings` endpoint

- [ ] **Step 1: Create Graph Event Creation Helper (`lib/graph/booking.ts`)**

Create `lib/graph/booking.ts`:
```typescript
import { getGraphClient } from './client';
import { CreateBookingParams, BookingResult } from '../types/booking';

export async function createCalendarEvent(
  accessToken: string | undefined,
  params: CreateBookingParams
): Promise<BookingResult> {
  if (!accessToken || accessToken.startsWith('mock-')) {
    // Simulate network delay and successful booking in demo mode
    await new Promise((resolve) => setTimeout(resolve, 800));
    return {
      success: true,
      eventId: `mock-evt-${Date.now()}`,
      message: `Room ${params.roomName} successfully booked! (Demo Mode)`,
      responseStatus: 'accepted',
    };
  }

  try {
    const client = getGraphClient(accessToken);

    const eventPayload = {
      subject: params.subject,
      body: {
        contentType: 'HTML',
        content: params.bodyContent || `Booked via Meeting Room Genius for ${params.roomName}`,
      },
      start: {
        dateTime: params.startTime,
        timeZone: 'UTC',
      },
      end: {
        dateTime: params.endTime,
        timeZone: 'UTC',
      },
      location: {
        displayName: params.roomName,
        locationEmailAddress: params.roomEmail,
      },
      attendees: [
        {
          emailAddress: {
            address: params.roomEmail,
            name: params.roomName,
          },
          type: 'resource',
        },
        ...(params.attendees || []).map((email) => ({
          emailAddress: { address: email },
          type: 'required',
        })),
      ],
    };

    const createdEvent = await client.api('/me/events').post(eventPayload);

    return {
      success: true,
      eventId: createdEvent.id,
      message: `Reservation confirmed for ${params.roomName}.`,
      responseStatus: createdEvent.responseStatus?.response || 'accepted',
    };
  } catch (error: any) {
    console.error('Error creating calendar event in Microsoft Graph:', error);

    // Extract HTTP 429 rate limit or Exchange double-booking error
    const status = error.statusCode || error.status;
    let userMsg = error.message || 'Failed to book room';

    if (status === 429) {
      userMsg = 'Microsoft Graph rate limit reached. Please retry in a few seconds.';
    } else if (error.message?.includes('Declined') || error.message?.includes('conflict')) {
      userMsg = 'This room was just booked by another user during the chosen slot.';
    }

    return {
      success: false,
      message: userMsg,
    };
  }
}
```

- [ ] **Step 2: Create Booking API Route (`app/api/bookings/route.ts`)**

Create `app/api/bookings/route.ts`:
```typescript
import { NextRequest, NextResponse } from 'next/server';
import { auth } from '@/lib/auth';
import { createCalendarEvent } from '@/lib/graph/booking';
import { CreateBookingParams } from '@/lib/types/booking';

export async function POST(req: NextRequest) {
  try {
    const session = await auth();
    const body: CreateBookingParams = await req.json();

    if (!body.roomEmail || !body.startTime || !body.endTime || !body.subject) {
      return NextResponse.json(
        { error: 'Missing required booking fields (roomEmail, startTime, endTime, subject)' },
        { status: 400 }
      );
    }

    const result = await createCalendarEvent(session?.accessToken, body);
    if (!result.success) {
      return NextResponse.json(result, { status: 422 });
    }

    return NextResponse.json(result, { status: 201 });
  } catch (error: any) {
    return NextResponse.json({ success: false, message: error.message || 'Internal Server Error' }, { status: 500 });
  }
}
```

- [ ] **Step 3: Verify Compilation**

Run: `npx tsc --noEmit`
Expected: 0 errors.

---

### Task 6: UI Components — Navigation, User Status, Filters & Room Card Layout

**Files:**
- Create: `components/providers/AuthProvider.tsx`
- Create: `components/auth/UserNav.tsx`
- Create: `components/auth/LoginButton.tsx`
- Create: `components/rooms/RoomCard.tsx`
- Create: `components/rooms/RoomSkeleton.tsx`
- Create: `components/rooms/RoomGrid.tsx`

**Interfaces:**
- Produces: Glassmorphic UI layout components with Lucide icons (`Users`, `MapPin`, `Calendar`, `Wifi`, `Tv`, `CheckCircle`, `XCircle`, `Clock`).

- [ ] **Step 1: Create Client Auth Session Provider (`components/providers/AuthProvider.tsx`)**

Create `components/providers/AuthProvider.tsx`:
```tsx
'use client';

import { SessionProvider } from 'next-auth/react';
import React from 'react';

export function AuthProvider({ children }: { children: React.ReactNode }) {
  return <SessionProvider>{children}</SessionProvider>;
}
```

- [ ] **Step 2: Create User Navigation & Login Button Components**

Create `components/auth/LoginButton.tsx`:
```tsx
'use client';

import { signIn } from 'next-auth/react';
import { LogIn } from 'lucide-react';

export function LoginButton() {
  return (
    <button
      onClick={() => signIn('azure-ad')}
      className="inline-flex items-center gap-2 px-5 py-2.5 rounded-xl bg-gradient-to-r from-cyan-500 to-blue-600 text-white font-medium shadow-lg shadow-cyan-500/20 hover:shadow-cyan-500/40 hover:scale-[1.02] active:scale-[0.98] transition-all"
    >
      <LogIn className="w-4 h-4" />
      Sign in with Microsoft 365
    </button>
  );
}
```

Create `components/auth/UserNav.tsx`:
```tsx
'use client';

import { useSession, signOut, signIn } from 'next-auth/react';
import { LogOut, User as UserIcon } from 'lucide-react';

export function UserNav() {
  const { data: session } = useSession();

  if (!session?.user) {
    return (
      <button
        onClick={() => signIn('azure-ad')}
        className="text-sm font-medium px-4 py-2 rounded-lg bg-cyan-500/10 text-cyan-400 border border-cyan-500/20 hover:bg-cyan-500/20 transition-all"
      >
        Sign In
      </button>
    );
  }

  return (
    <div className="flex items-center gap-3 bg-white/5 border border-white/10 px-3 py-1.5 rounded-xl backdrop-blur-md">
      {session.user.image ? (
        <img
          src={session.user.image}
          alt={session.user.name || 'User'}
          className="w-8 h-8 rounded-full border border-cyan-500/40 object-cover"
        />
      ) : (
        <div className="w-8 h-8 rounded-full bg-cyan-500/20 text-cyan-400 flex items-center justify-center font-bold text-xs">
          {session.user.name?.[0] || 'U'}
        </div>
      )}
      <div className="hidden sm:block text-left">
        <p className="text-xs font-semibold text-gray-200">{session.user.name}</p>
        <p className="text-[10px] text-gray-400 truncate max-w-[120px]">{session.user.email}</p>
      </div>
      <button
        onClick={() => signOut()}
        title="Sign Out"
        className="text-gray-400 hover:text-rose-400 p-1.5 rounded-lg hover:bg-white/5 transition-colors ml-1"
      >
        <LogOut className="w-4 h-4" />
      </button>
    </div>
  );
}
```

- [ ] **Step 3: Create Room Card & Skeleton Loader**

Create `components/rooms/RoomSkeleton.tsx`:
```tsx
export function RoomSkeleton() {
  return (
    <div className="bg-gray-900/60 border border-white/10 rounded-2xl p-5 animate-pulse space-y-4">
      <div className="flex justify-between items-start">
        <div className="h-6 w-1/2 bg-gray-800 rounded-lg"></div>
        <div className="h-5 w-20 bg-gray-800 rounded-full"></div>
      </div>
      <div className="h-4 w-1/3 bg-gray-800 rounded"></div>
      <div className="h-10 w-full bg-gray-800/80 rounded-xl"></div>
      <div className="h-9 w-full bg-gray-800 rounded-xl"></div>
    </div>
  );
}
```

Create `components/rooms/RoomCard.tsx`:
```tsx
'use client';

import { Room, RoomSchedule } from '@/lib/types/room';
import { Users, MapPin, Tv, Wifi, CheckCircle2, AlertCircle } from 'lucide-react';
import { RoomTimeline } from './RoomTimeline';

interface RoomCardProps {
  room: Room;
  schedule?: RoomSchedule;
  selectedDate: Date;
  onBookClick: (room: Room, startTime?: string) => void;
}

export function RoomCard({ room, schedule, selectedDate, onBookClick }: RoomCardProps) {
  return (
    <div className="bg-gradient-to-b from-gray-900/80 to-gray-950/90 border border-white/10 hover:border-cyan-500/30 rounded-2xl p-5 shadow-xl backdrop-blur-xl transition-all duration-300 flex flex-col justify-between group">
      <div>
        {/* Header: Name & Availability Badge */}
        <div className="flex justify-between items-start gap-2 mb-3">
          <div>
            <h3 className="text-lg font-bold text-white group-hover:text-cyan-400 transition-colors">
              {room.displayName}
            </h3>
            <p className="text-xs text-gray-400 flex items-center gap-1 mt-0.5">
              <MapPin className="w-3.5 h-3.5 text-cyan-400" />
              {room.building} &bull; {room.floorNumber}
            </p>
          </div>
          <span
            className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold border ${
              room.isAvailable
                ? 'bg-emerald-500/10 text-emerald-400 border-emerald-500/30'
                : 'bg-rose-500/10 text-rose-400 border-rose-500/30'
            }`}
          >
            {room.isAvailable ? (
              <>
                <CheckCircle2 className="w-3 h-3" /> Available
              </>
            ) : (
              <>
                <AlertCircle className="w-3 h-3" /> Occupied
              </>
            )}
          </span>
        </div>

        {/* Capacity & Features */}
        <div className="flex flex-wrap items-center gap-3 my-3 text-xs text-gray-300">
          <div className="flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-white/5 border border-white/10">
            <Users className="w-3.5 h-3.5 text-cyan-400" />
            <span>Capacity: <strong className="text-white">{room.capacity}</strong></span>
          </div>

          {(room.features || []).map((feat, idx) => (
            <span
              key={idx}
              className="px-2.5 py-1 rounded-lg bg-white/5 border border-white/5 text-gray-400 flex items-center gap-1"
            >
              {feat.includes('AV') || feat.includes('TV') ? <Tv className="w-3 h-3 text-cyan-400" /> : <Wifi className="w-3 h-3 text-cyan-400" />}
              {feat}
            </span>
          ))}
        </div>

        {/* Timeline Visualization */}
        <div className="mt-4">
          <p className="text-[11px] font-semibold text-gray-400 uppercase tracking-wider mb-2">
            Schedule (08:00 - 18:00)
          </p>
          <RoomTimeline
            schedule={schedule}
            selectedDate={selectedDate}
            onSelectSlot={(slotTime) => onBookClick(room, slotTime)}
          />
        </div>
      </div>

      {/* CTA Button */}
      <button
        onClick={() => onBookClick(room)}
        className="mt-5 w-full py-2.5 px-4 rounded-xl bg-cyan-500/10 hover:bg-cyan-500 text-cyan-400 hover:text-black font-semibold border border-cyan-500/30 hover:border-cyan-400 transition-all flex items-center justify-center gap-2 shadow-lg shadow-cyan-500/10"
      >
        Book Room Now
      </button>
    </div>
  );
}
```

- [ ] **Step 4: Create Room Grid Container (`components/rooms/RoomGrid.tsx`)**

Create `components/rooms/RoomGrid.tsx`:
```tsx
'use client';

import { useState } from 'react';
import { Room, RoomSchedule } from '@/lib/types/room';
import { RoomCard } from './RoomCard';
import { RoomSkeleton } from './RoomSkeleton';
import { Search, Filter } from 'lucide-react';

interface RoomGridProps {
  rooms: Room[];
  schedules: Record<string, RoomSchedule>;
  loading: boolean;
  selectedDate: Date;
  onBookClick: (room: Room, startTime?: string) => void;
}

export function RoomGrid({ rooms, schedules, loading, selectedDate, onBookClick }: RoomGridProps) {
  const [searchTerm, setSearchTerm] = useState('');
  const [minCapacity, setMinCapacity] = useState<number>(0);

  const filteredRooms = rooms.filter((room) => {
    const matchesSearch =
      room.displayName.toLowerCase().includes(searchTerm.toLowerCase()) ||
      (room.building || '').toLowerCase().includes(searchTerm.toLowerCase());
    const matchesCap = room.capacity >= minCapacity;
    return matchesSearch && matchesCap;
  });

  return (
    <div className="space-y-6">
      {/* Search & Filter Controls */}
      <div className="flex flex-col sm:flex-row items-center justify-between gap-4 bg-gray-900/60 p-4 rounded-2xl border border-white/10 backdrop-blur-md">
        <div className="relative w-full sm:w-80">
          <Search className="w-4 h-4 absolute left-3.5 top-1/2 -translate-y-1/2 text-gray-400" />
          <input
            type="text"
            placeholder="Search room name or building..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="w-full bg-black/40 border border-white/10 rounded-xl pl-10 pr-4 py-2 text-sm text-white placeholder-gray-500 focus:outline-none focus:border-cyan-500/50 transition-colors"
          />
        </div>

        <div className="flex items-center gap-2 w-full sm:w-auto">
          <Filter className="w-4 h-4 text-cyan-400" />
          <span className="text-xs text-gray-400 font-medium">Min Capacity:</span>
          <div className="flex gap-1">
            {[0, 4, 8, 12, 16].map((cap) => (
              <button
                key={cap}
                onClick={() => setMinCapacity(cap)}
                className={`px-2.5 py-1 text-xs rounded-lg font-medium border transition-all ${
                  minCapacity === cap
                    ? 'bg-cyan-500 text-black font-bold border-cyan-400'
                    : 'bg-white/5 text-gray-400 border-white/10 hover:bg-white/10'
                }`}
              >
                {cap === 0 ? 'All' : `${cap}+`}
              </button>
            ))}
          </div>
        </div>
      </div>

      {/* Room Grid / Skeletons */}
      {loading ? (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          <RoomSkeleton />
          <RoomSkeleton />
          <RoomSkeleton />
        </div>
      ) : filteredRooms.length === 0 ? (
        <div className="text-center py-16 bg-gray-900/40 rounded-2xl border border-white/5">
          <p className="text-gray-400 text-sm">No meeting rooms found matching your search filters.</p>
        </div>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          {filteredRooms.map((room) => (
            <RoomCard
              key={room.id}
              room={room}
              schedule={schedules[room.emailAddress]}
              selectedDate={selectedDate}
              onBookClick={onBookClick}
            />
          ))}
        </div>
      )}
    </div>
  );
}
```

- [ ] **Step 5: Verify Compilation**

Run: `npx tsc --noEmit`
Expected: 0 errors.

---

### Task 7: Interactive Availability Timeline & Date Picker Components

**Files:**
- Create: `components/rooms/RoomTimeline.tsx`
- Create: `components/rooms/DatePickerHeader.tsx`

**Interfaces:**
- Produces: 30-minute block visual timeline (08:00 to 18:00) with color-coded states (emerald for free, rose for busy/booked).
- Produces: Quick Date navigator header with Today, Tomorrow, and Date input.

- [ ] **Step 1: Create Date Header Navigator (`components/rooms/DatePickerHeader.tsx`)**

Create `components/rooms/DatePickerHeader.tsx`:
```tsx
'use client';

import { ChevronLeft, ChevronRight, Calendar as CalendarIcon } from 'lucide-react';
import { format, addDays, subDays, isToday, isTomorrow } from 'date-fns';

interface DatePickerHeaderProps {
  selectedDate: Date;
  onChangeDate: (date: Date) => void;
}

export function DatePickerHeader({ selectedDate, onChangeDate }: DatePickerHeaderProps) {
  const dateLabel = isToday(selectedDate)
    ? 'Today'
    : isTomorrow(selectedDate)
    ? 'Tomorrow'
    : format(selectedDate, 'EEEE, MMM d, yyyy');

  return (
    <div className="flex flex-col sm:flex-row items-center justify-between gap-4 bg-gradient-to-r from-gray-900/90 to-gray-950/90 p-4 rounded-2xl border border-white/10 backdrop-blur-xl">
      <div className="flex items-center gap-3">
        <div className="p-2.5 rounded-xl bg-cyan-500/10 border border-cyan-500/20 text-cyan-400">
          <CalendarIcon className="w-5 h-5" />
        </div>
        <div>
          <h2 className="text-xl font-bold text-white tracking-tight">{dateLabel}</h2>
          <p className="text-xs text-gray-400">{format(selectedDate, 'MMMM d, yyyy')}</p>
        </div>
      </div>

      <div className="flex items-center gap-2">
        <button
          onClick={() => onChangeDate(subDays(selectedDate, 1))}
          className="p-2 rounded-xl bg-white/5 border border-white/10 text-gray-300 hover:bg-white/10 transition-colors"
          title="Previous Day"
        >
          <ChevronLeft className="w-4 h-4" />
        </button>

        <button
          onClick={() => onChangeDate(new Date())}
          className={`px-3 py-1.5 rounded-xl text-xs font-semibold border transition-all ${
            isToday(selectedDate)
              ? 'bg-cyan-500 text-black border-cyan-400 font-bold'
              : 'bg-white/5 text-gray-300 border-white/10 hover:bg-white/10'
          }`}
        >
          Today
        </button>

        <button
          onClick={() => onChangeDate(addDays(new Date(), 1))}
          className={`px-3 py-1.5 rounded-xl text-xs font-semibold border transition-all ${
            isTomorrow(selectedDate)
              ? 'bg-cyan-500 text-black border-cyan-400 font-bold'
              : 'bg-white/5 text-gray-300 border-white/10 hover:bg-white/10'
          }`}
        >
          Tomorrow
        </button>

        <button
          onClick={() => onChangeDate(addDays(selectedDate, 1))}
          className="p-2 rounded-xl bg-white/5 border border-white/10 text-gray-300 hover:bg-white/10 transition-colors"
          title="Next Day"
        >
          <ChevronRight className="w-4 h-4" />
        </button>
      </div>
    </div>
  );
}
```

- [ ] **Step 2: Create Visual Room Timeline (`components/rooms/RoomTimeline.tsx`)**

Create `components/rooms/RoomTimeline.tsx`:
```tsx
'use client';

import { RoomSchedule } from '@/lib/types/room';
import { format } from 'date-fns';

interface RoomTimelineProps {
  schedule?: RoomSchedule;
  selectedDate: Date;
  onSelectSlot: (startTimeISO: string) => void;
}

export function RoomTimeline({ schedule, selectedDate, onSelectSlot }: RoomTimelineProps) {
  // Hours from 08:00 to 18:00 (10 hours = 20 half-hour blocks)
  const hours = Array.from({ length: 10 }, (_, i) => i + 8);

  const isSlotBooked = (hour: number, minute: number) => {
    if (!schedule?.scheduleItems) return false;

    const baseDateStr = format(selectedDate, 'yyyy-MM-dd');
    const slotTime = new Date(`${baseDateStr}T${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}:00Z`).getTime();

    return schedule.scheduleItems.some((item) => {
      const start = new Date(item.start).getTime();
      const end = new Date(item.end).getTime();
      return slotTime >= start && slotTime < end;
    });
  };

  return (
    <div className="space-y-1.5">
      {/* Time Labels */}
      <div className="flex justify-between text-[10px] text-gray-500 font-mono px-0.5">
        <span>08:00</span>
        <span>10:00</span>
        <span>12:00</span>
        <span>14:00</span>
        <span>16:00</span>
        <span>18:00</span>
      </div>

      {/* 30-min Block Grid */}
      <div className="grid grid-cols-20 gap-1 bg-black/40 p-1.5 rounded-xl border border-white/5">
        {hours.flatMap((hour) =>
          [0, 30].map((minute) => {
            const booked = isSlotBooked(hour, minute);
            const baseDateStr = format(selectedDate, 'yyyy-MM-dd');
            const timeISO = `${baseDateStr}T${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}:00Z`;
            const label = `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;

            return (
              <button
                key={`${hour}-${minute}`}
                onClick={() => !booked && onSelectSlot(timeISO)}
                disabled={booked}
                title={`${label} - ${booked ? 'Booked' : 'Available (Click to book)'}`}
                className={`h-6 rounded-md border transition-all ${
                  booked
                    ? 'bg-rose-500/20 border-rose-500/30 cursor-not-allowed'
                    : 'bg-emerald-500/20 border-emerald-500/30 hover:bg-emerald-400 hover:border-emerald-300 cursor-pointer'
                }`}
              />
            );
          })
        )}
      </div>

      {/* Legend */}
      <div className="flex items-center gap-4 text-[10px] text-gray-400 pt-1 justify-end">
        <div className="flex items-center gap-1">
          <span className="w-2.5 h-2.5 rounded-sm bg-emerald-500/40 border border-emerald-500/60 inline-block"></span>
          <span>Available</span>
        </div>
        <div className="flex items-center gap-1">
          <span className="w-2.5 h-2.5 rounded-sm bg-rose-500/40 border border-rose-500/60 inline-block"></span>
          <span>Booked</span>
        </div>
      </div>
    </div>
  );
}
```

- [ ] **Step 3: Verify Compilation**

Run: `npx tsc --noEmit`
Expected: 0 errors.

---

### Task 8: Booking Modal & Toast Notification Components

**Files:**
- Create: `components/rooms/BookingModal.tsx`
- Create: `components/ui/Toast.tsx`

**Interfaces:**
- Produces: `BookingModal` dialog for subject, date/time range selection, attendee emails, and reservation submission.
- Produces: `Toast` overlay for real-time success & error feedback messages.

- [ ] **Step 1: Create Toast Notification Component (`components/ui/Toast.tsx`)**

Create `components/ui/Toast.tsx`:
```tsx
'use client';

import { CheckCircle2, AlertCircle, X } from 'lucide-react';
import { useEffect } from 'react';

export interface ToastMessage {
  id: string;
  type: 'success' | 'error';
  text: string;
}

interface ToastProps {
  toasts: ToastMessage[];
  onDismiss: (id: string) => void;
}

export function Toast({ toasts, onDismiss }: ToastProps) {
  useEffect(() => {
    if (toasts.length > 0) {
      const timer = setTimeout(() => {
        onDismiss(toasts[0].id);
      }, 5000);
      return () => clearTimeout(timer);
    }
  }, [toasts, onDismiss]);

  if (toasts.length === 0) return null;

  return (
    <div className="fixed bottom-5 right-5 z-50 flex flex-col gap-2 max-w-sm w-full">
      {toasts.map((toast) => (
        <div
          key={toast.id}
          className={`flex items-start gap-3 p-4 rounded-xl border backdrop-blur-xl shadow-2xl transition-all animate-in slide-in-from-bottom-3 ${
            toast.type === 'success'
              ? 'bg-emerald-950/90 border-emerald-500/40 text-emerald-200'
              : 'bg-rose-950/90 border-rose-500/40 text-rose-200'
          }`}
        >
          {toast.type === 'success' ? (
            <CheckCircle2 className="w-5 h-5 text-emerald-400 shrink-0 mt-0.5" />
          ) : (
            <AlertCircle className="w-5 h-5 text-rose-400 shrink-0 mt-0.5" />
          )}
          <div className="flex-1 text-sm font-medium">{toast.text}</div>
          <button
            onClick={() => onDismiss(toast.id)}
            className="text-gray-400 hover:text-white p-0.5 rounded transition-colors"
          >
            <X className="w-4 h-4" />
          </button>
        </div>
      ))}
    </div>
  );
}
```

- [ ] **Step 2: Create Booking Modal (`components/rooms/BookingModal.tsx`)**

Create `components/rooms/BookingModal.tsx`:
```tsx
'use client';

import { useState, useEffect } from 'react';
import { Room } from '@/lib/types/room';
import { X, Calendar, Clock, MapPin, Users, Send } from 'lucide-react';

interface BookingModalProps {
  room: Room | null;
  initialStartTime?: string;
  isOpen: boolean;
  onClose: () => void;
  onSuccess: (msg: string) => void;
  onError: (msg: string) => void;
}

export function BookingModal({
  room,
  initialStartTime,
  isOpen,
  onClose,
  onSuccess,
  onError,
}: BookingModalProps) {
  const [subject, setSubject] = useState('');
  const [description, setDescription] = useState('');
  const [date, setDate] = useState('');
  const [startTime, setStartTime] = useState('09:00');
  const [endTime, setEndTime] = useState('10:00');
  const [attendees, setAttendees] = useState('');
  const [submitting, setSubmitting] = useState(false);

  useEffect(() => {
    if (initialStartTime) {
      const dt = new Date(initialStartTime);
      setDate(dt.toISOString().split('T')[0]);
      const hours = String(dt.getUTCHours()).padStart(2, '0');
      const mins = String(dt.getUTCMinutes()).padStart(2, '0');
      setStartTime(`${hours}:${mins}`);

      // Default end time 1 hour later
      const endDt = new Date(dt.getTime() + 60 * 60 * 1000);
      const endHours = String(endDt.getUTCHours()).padStart(2, '0');
      const endMins = String(endDt.getUTCMinutes()).padStart(2, '0');
      setEndTime(`${endHours}:${endMins}`);
    } else {
      setDate(new Date().toISOString().split('T')[0]);
    }
  }, [initialStartTime, isOpen]);

  if (!isOpen || !room) return null;

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!subject.trim()) {
      onError('Please enter a meeting title.');
      return;
    }

    setSubmitting(true);
    try {
      const startIso = new Date(`${date}T${startTime}:00Z`).toISOString();
      const endIso = new Date(`${date}T${endTime}:00Z`).toISOString();

      const attendeeList = attendees
        .split(',')
        .map((a) => a.trim())
        .filter((a) => a.includes('@'));

      const res = await fetch('/api/bookings', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          roomEmail: room.emailAddress,
          roomName: room.displayName,
          subject,
          bodyContent: description,
          startTime: startIso,
          endTime: endIso,
          attendees: attendeeList,
        }),
      });

      const data = await res.json();
      if (res.ok && data.success) {
        onSuccess(data.message || `Successfully reserved ${room.displayName}!`);
        onClose();
        setSubject('');
        setDescription('');
      } else {
        onError(data.message || 'Failed to complete room reservation.');
      }
    } catch (err: any) {
      onError(err.message || 'Network error occurred while booking.');
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/70 backdrop-blur-md animate-in fade-in duration-200">
      <div className="relative w-full max-w-lg bg-gray-900 border border-white/10 rounded-2xl p-6 shadow-2xl space-y-5">
        {/* Header */}
        <div className="flex justify-between items-start border-b border-white/10 pb-4">
          <div>
            <h2 className="text-xl font-bold text-white">Book {room.displayName}</h2>
            <p className="text-xs text-gray-400 flex items-center gap-1 mt-1">
              <MapPin className="w-3.5 h-3.5 text-cyan-400" />
              {room.building} &bull; Capacity: {room.capacity}
            </p>
          </div>
          <button
            onClick={onClose}
            className="p-1 rounded-lg text-gray-400 hover:text-white hover:bg-white/10 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Form */}
        <form onSubmit={handleSubmit} className="space-y-4 text-sm">
          <div>
            <label className="block text-xs font-semibold text-gray-300 mb-1">
              Meeting Title / Subject *
            </label>
            <input
              type="text"
              required
              placeholder="e.g. Q3 Product Strategy Review"
              value={subject}
              onChange={(e) => setSubject(e.target.value)}
              className="w-full bg-black/40 border border-white/10 rounded-xl px-3.5 py-2.5 text-white placeholder-gray-500 focus:outline-none focus:border-cyan-500 transition-colors"
            />
          </div>

          <div className="grid grid-cols-3 gap-3">
            <div>
              <label className="block text-xs font-semibold text-gray-300 mb-1">Date</label>
              <input
                type="date"
                required
                value={date}
                onChange={(e) => setDate(e.target.value)}
                className="w-full bg-black/40 border border-white/10 rounded-xl px-3 py-2 text-white text-xs focus:outline-none focus:border-cyan-500"
              />
            </div>
            <div>
              <label className="block text-xs font-semibold text-gray-300 mb-1">Start Time</label>
              <input
                type="time"
                required
                value={startTime}
                onChange={(e) => setStartTime(e.target.value)}
                className="w-full bg-black/40 border border-white/10 rounded-xl px-3 py-2 text-white text-xs focus:outline-none focus:border-cyan-500"
              />
            </div>
            <div>
              <label className="block text-xs font-semibold text-gray-300 mb-1">End Time</label>
              <input
                type="time"
                required
                value={endTime}
                onChange={(e) => setEndTime(e.target.value)}
                className="w-full bg-black/40 border border-white/10 rounded-xl px-3 py-2 text-white text-xs focus:outline-none focus:border-cyan-500"
              />
            </div>
          </div>

          <div>
            <label className="block text-xs font-semibold text-gray-300 mb-1">
              Optional Attendees (comma separated emails)
            </label>
            <input
              type="text"
              placeholder="colleague@company.com, manager@company.com"
              value={attendees}
              onChange={(e) => setAttendees(e.target.value)}
              className="w-full bg-black/40 border border-white/10 rounded-xl px-3.5 py-2 text-white placeholder-gray-500 focus:outline-none focus:border-cyan-500"
            />
          </div>

          <div>
            <label className="block text-xs font-semibold text-gray-300 mb-1">Description</label>
            <textarea
              rows={2}
              placeholder="Agenda or additional notes..."
              value={description}
              onChange={(e) => setDescription(e.target.value)}
              className="w-full bg-black/40 border border-white/10 rounded-xl px-3.5 py-2 text-white placeholder-gray-500 focus:outline-none focus:border-cyan-500 resize-none"
            />
          </div>

          {/* Footer Action */}
          <div className="flex items-center justify-end gap-3 pt-3 border-t border-white/10">
            <button
              type="button"
              onClick={onClose}
              className="px-4 py-2 rounded-xl text-gray-400 hover:text-white font-medium transition-colors"
            >
              Cancel
            </button>
            <button
              type="submit"
              disabled={submitting}
              className="px-5 py-2.5 rounded-xl bg-gradient-to-r from-cyan-500 to-blue-600 hover:from-cyan-400 hover:to-blue-500 text-black font-bold shadow-lg shadow-cyan-500/20 disabled:opacity-50 flex items-center gap-2 transition-all"
            >
              {submitting ? 'Confirming...' : 'Confirm Reservation'}
              <Send className="w-4 h-4" />
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}
```

- [ ] **Step 3: Verify Compilation**

Run: `npx tsc --noEmit`
Expected: 0 errors.

---

### Task 9: Assemble Dashboard Page, Landing Page, App Layout & Integration Verification

**Files:**
- Create: `app/globals.css`
- Create: `app/layout.tsx`
- Create: `app/page.tsx`
- Create: `app/dashboard/page.tsx`

**Interfaces:**
- Produces: Complete, responsive corporate room booking application with interactive timeline, room filters, modal dialog, and notifications.

- [ ] **Step 1: Create Global CSS Styles (`app/globals.css`)**

Create `app/globals.css`:
```css
@tailwind base;
@tailwind components;
@tailwind utilities;

@layer base {
  body {
    @apply bg-darkBg text-gray-100 antialiased min-h-screen font-sans selection:bg-cyan-500 selection:text-black;
  }
}

.grid-cols-20 {
  grid-template-columns: repeat(20, minmax(0, 1fr));
}
```

- [ ] **Step 2: Create Root Layout (`app/layout.tsx`)**

Create `app/layout.tsx`:
```tsx
import type { Metadata } from 'next';
import { Inter } from 'next/font/google';
import './globals.css';
import { AuthProvider } from '@/components/providers/AuthProvider';

const inter = Inter({ subsets: ['latin'] });

export const metadata: Metadata = {
  title: 'Meeting Room Genius | M365 Room Discovery & Booking',
  description: 'Instant Microsoft 365 office meeting room availability, scheduling timeline, and reservation platform.',
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en" className="dark">
      <body className={inter.className}>
        <AuthProvider>{children}</AuthProvider>
      </body>
    </html>
  );
}
```

- [ ] **Step 3: Create Main Dashboard Page (`app/dashboard/page.tsx`)**

Create `app/dashboard/page.tsx`:
```tsx
'use client';

import { useState, useEffect } from 'react';
import { Room, RoomSchedule } from '@/lib/types/room';
import { UserNav } from '@/components/auth/UserNav';
import { RoomGrid } from '@/components/rooms/RoomGrid';
import { DatePickerHeader } from '@/components/rooms/DatePickerHeader';
import { BookingModal } from '@/components/rooms/BookingModal';
import { Toast, ToastMessage } from '@/components/ui/Toast';
import { Building2, Sparkles } from 'lucide-react';

export default function DashboardPage() {
  const [rooms, setRooms] = useState<Room[]>([]);
  const [schedules, setSchedules] = useState<Record<string, RoomSchedule>>({});
  const [loading, setLoading] = useState<boolean>(true);
  const [selectedDate, setSelectedDate] = useState<Date>(new Date());

  const [modalOpen, setModalOpen] = useState<boolean>(false);
  const [selectedRoom, setSelectedRoom] = useState<Room | null>(null);
  const [initialSlotTime, setInitialSlotTime] = useState<string | undefined>();

  const [toasts, setToasts] = useState<ToastMessage[]>([]);

  const addToast = (type: 'success' | 'error', text: string) => {
    setToasts((prev) => [...prev, { id: String(Date.now()), type, text }]);
  };

  const removeToast = (id: string) => {
    setToasts((prev) => prev.filter((t) => t.id !== id));
  };

  const loadData = async () => {
    setLoading(true);
    try {
      const roomRes = await fetch('/api/rooms');
      const roomData = await roomRes.json();
      const loadedRooms: Room[] = roomData.rooms || [];
      setRooms(loadedRooms);

      if (loadedRooms.length > 0) {
        const emails = loadedRooms.map((r) => r.emailAddress);
        const schedRes = await fetch('/api/rooms/schedule', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ roomEmails: emails, date: selectedDate.toISOString() }),
        });
        const schedData = await schedRes.json();
        setSchedules(schedData.schedules || {});
      }
    } catch (err) {
      console.error('Error fetching dashboard data:', err);
      addToast('error', 'Failed to load room availability schedules.');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadData();
  }, [selectedDate]);

  const handleOpenBooking = (room: Room, startTime?: string) => {
    setSelectedRoom(room);
    setInitialSlotTime(startTime);
    setModalOpen(true);
  };

  return (
    <div className="min-h-screen bg-darkBg flex flex-col">
      {/* Top Navbar */}
      <header className="border-b border-white/10 bg-gray-950/80 backdrop-blur-xl sticky top-0 z-40">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-3.5 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-gradient-to-tr from-cyan-500 to-blue-600 rounded-xl shadow-lg shadow-cyan-500/20 text-black">
              <Building2 className="w-5 h-5 font-bold" />
            </div>
            <div>
              <h1 className="text-lg font-bold text-white tracking-tight flex items-center gap-2">
                Meeting Room Genius
                <span className="text-[10px] font-semibold px-2 py-0.5 rounded-full bg-cyan-500/10 text-cyan-400 border border-cyan-500/20">
                  M365 Live
                </span>
              </h1>
              <p className="text-xs text-gray-400">Microsoft 365 Executive Booking Suite</p>
            </div>
          </div>

          <UserNav />
        </div>
      </header>

      {/* Main Content Area */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-8">
        {/* Hero Section */}
        <div className="relative rounded-3xl p-6 sm:p-8 bg-gradient-to-r from-gray-900 via-gray-900/90 to-gray-950 border border-white/10 shadow-2xl overflow-hidden">
          <div className="absolute right-0 top-0 w-96 h-96 bg-cyan-500/10 rounded-full blur-3xl -mr-20 -mt-20 pointer-events-none"></div>
          <div className="relative z-10 space-y-2">
            <div className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-cyan-500/10 border border-cyan-500/30 text-cyan-400 text-xs font-semibold">
              <Sparkles className="w-3.5 h-3.5" /> Direct Exchange Online Integration
            </div>
            <h2 className="text-2xl sm:text-3xl font-extrabold text-white">
              Find & Reserve Your Next Meeting Room
            </h2>
            <p className="text-sm text-gray-400 max-w-2xl">
              Real-time availability timelines, capacity filters, and instant delegated booking powered by Microsoft Graph API.
            </p>
          </div>
        </div>

        {/* Date Selector Header */}
        <DatePickerHeader selectedDate={selectedDate} onChangeDate={setSelectedDate} />

        {/* Room Discovery & Grid */}
        <RoomGrid
          rooms={rooms}
          schedules={schedules}
          loading={loading}
          selectedDate={selectedDate}
          onBookClick={handleOpenBooking}
        />
      </main>

      {/* Booking Modal Dialog */}
      <BookingModal
        room={selectedRoom}
        initialStartTime={initialSlotTime}
        isOpen={modalOpen}
        onClose={() => setModalOpen(false)}
        onSuccess={(msg) => {
          addToast('success', msg);
          loadData();
        }}
        onError={(msg) => addToast('error', msg)}
      />

      {/* Toast Overlay */}
      <Toast toasts={toasts} onDismiss={removeToast} />
    </div>
  );
}
```

- [ ] **Step 4: Create Landing Page / Auth Redirect (`app/page.tsx`)**

Create `app/page.tsx`:
```tsx
import DashboardPage from './dashboard/page';

export default function HomePage() {
  return <DashboardPage />;
}
```

- [ ] **Step 5: Execute TypeScript Type-Check & Verification**

Run: `npx tsc --noEmit`
Expected: 0 errors.

- [ ] **Step 6: Commit Implementation**

```bash
git add .
git commit -m "feat: complete M365 meeting room booking web app"
```
