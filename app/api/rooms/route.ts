import { NextResponse } from 'next/server';
import { auth } from '@/lib/auth';
import { fetchRooms } from '@/lib/graph/rooms';

export async function GET() {
  const session = await auth();
  const rooms = await fetchRooms(session?.accessToken);
  return NextResponse.json({ rooms });
}
