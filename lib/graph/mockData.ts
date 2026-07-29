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
