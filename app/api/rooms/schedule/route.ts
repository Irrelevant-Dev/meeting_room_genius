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
