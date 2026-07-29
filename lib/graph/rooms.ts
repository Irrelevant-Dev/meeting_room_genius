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
