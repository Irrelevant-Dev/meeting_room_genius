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
