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
