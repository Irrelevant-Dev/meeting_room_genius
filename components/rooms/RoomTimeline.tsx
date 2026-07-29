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
