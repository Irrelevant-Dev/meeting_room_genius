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
