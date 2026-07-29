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
