'use client';

import { Room, RoomSchedule } from '@/lib/types/room';
import { Users, MapPin, Tv, Wifi, CheckCircle2, AlertCircle } from 'lucide-react';
import { RoomTimeline } from './RoomTimeline';

interface RoomCardProps {
  room: Room;
  schedule?: RoomSchedule;
  selectedDate: Date;
  onBookClick: (room: Room, startTime?: string) => void;
}

export function RoomCard({ room, schedule, selectedDate, onBookClick }: RoomCardProps) {
  return (
    <div className="bg-gradient-to-b from-gray-900/80 to-gray-950/90 border border-white/10 hover:border-cyan-500/30 rounded-2xl p-5 shadow-xl backdrop-blur-xl transition-all duration-300 flex flex-col justify-between group">
      <div>
        {/* Header: Name & Availability Badge */}
        <div className="flex justify-between items-start gap-2 mb-3">
          <div>
            <h3 className="text-lg font-bold text-white group-hover:text-cyan-400 transition-colors">
              {room.displayName}
            </h3>
            <p className="text-xs text-gray-400 flex items-center gap-1 mt-0.5">
              <MapPin className="w-3.5 h-3.5 text-cyan-400" />
              {room.building} &bull; {room.floorNumber}
            </p>
          </div>
          <span
            className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold border ${
              room.isAvailable
                ? 'bg-emerald-500/10 text-emerald-400 border-emerald-500/30'
                : 'bg-rose-500/10 text-rose-400 border-rose-500/30'
            }`}
          >
            {room.isAvailable ? (
              <>
                <CheckCircle2 className="w-3 h-3" /> Available
              </>
            ) : (
              <>
                <AlertCircle className="w-3 h-3" /> Occupied
              </>
            )}
          </span>
        </div>

        {/* Capacity & Features */}
        <div className="flex flex-wrap items-center gap-3 my-3 text-xs text-gray-300">
          <div className="flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-white/5 border border-white/10">
            <Users className="w-3.5 h-3.5 text-cyan-400" />
            <span>Capacity: <strong className="text-white">{room.capacity}</strong></span>
          </div>

          {(room.features || []).map((feat, idx) => (
            <span
              key={idx}
              className="px-2.5 py-1 rounded-lg bg-white/5 border border-white/5 text-gray-400 flex items-center gap-1"
            >
              {feat.includes('AV') || feat.includes('TV') ? <Tv className="w-3 h-3 text-cyan-400" /> : <Wifi className="w-3 h-3 text-cyan-400" />}
              {feat}
            </span>
          ))}
        </div>

        {/* Timeline Visualization */}
        <div className="mt-4">
          <p className="text-[11px] font-semibold text-gray-400 uppercase tracking-wider mb-2">
            Schedule (08:00 - 18:00)
          </p>
          <RoomTimeline
            schedule={schedule}
            selectedDate={selectedDate}
            onSelectSlot={(slotTime) => onBookClick(room, slotTime)}
          />
        </div>
      </div>

      {/* CTA Button */}
      <button
        onClick={() => onBookClick(room)}
        className="mt-5 w-full py-2.5 px-4 rounded-xl bg-cyan-500/10 hover:bg-cyan-500 text-cyan-400 hover:text-black font-semibold border border-cyan-500/30 hover:border-cyan-400 transition-all flex items-center justify-center gap-2 shadow-lg shadow-cyan-500/10"
      >
        Book Room Now
      </button>
    </div>
  );
}
