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
