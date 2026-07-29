'use client';

import { useState } from 'react';
import { Room, RoomSchedule } from '@/lib/types/room';
import { RoomCard } from './RoomCard';
import { RoomSkeleton } from './RoomSkeleton';
import { Search, Filter } from 'lucide-react';

interface RoomGridProps {
  rooms: Room[];
  schedules: Record<string, RoomSchedule>;
  loading: boolean;
  selectedDate: Date;
  onBookClick: (room: Room, startTime?: string) => void;
}

export function RoomGrid({ rooms, schedules, loading, selectedDate, onBookClick }: RoomGridProps) {
  const [searchTerm, setSearchTerm] = useState('');
  const [minCapacity, setMinCapacity] = useState<number>(0);

  const filteredRooms = rooms.filter((room) => {
    const matchesSearch =
      room.displayName.toLowerCase().includes(searchTerm.toLowerCase()) ||
      (room.building || '').toLowerCase().includes(searchTerm.toLowerCase());
    const matchesCap = room.capacity >= minCapacity;
    return matchesSearch && matchesCap;
  });

  return (
    <div className="space-y-6">
      {/* Search & Filter Controls */}
      <div className="flex flex-col sm:flex-row items-center justify-between gap-4 bg-gray-900/60 p-4 rounded-2xl border border-white/10 backdrop-blur-md">
        <div className="relative w-full sm:w-80">
          <Search className="w-4 h-4 absolute left-3.5 top-1/2 -translate-y-1/2 text-gray-400" />
          <input
            type="text"
            placeholder="Search room name or building..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="w-full bg-black/40 border border-white/10 rounded-xl pl-10 pr-4 py-2 text-sm text-white placeholder-gray-500 focus:outline-none focus:border-cyan-500/50 transition-colors"
          />
        </div>

        <div className="flex items-center gap-2 w-full sm:w-auto">
          <Filter className="w-4 h-4 text-cyan-400" />
          <span className="text-xs text-gray-400 font-medium">Min Capacity:</span>
          <div className="flex gap-1">
            {[0, 4, 8, 12, 16].map((cap) => (
              <button
                key={cap}
                onClick={() => setMinCapacity(cap)}
                className={`px-2.5 py-1 text-xs rounded-lg font-medium border transition-all ${
                  minCapacity === cap
                    ? 'bg-cyan-500 text-black font-bold border-cyan-400'
                    : 'bg-white/5 text-gray-400 border-white/10 hover:bg-white/10'
                }`}
              >
                {cap === 0 ? 'All' : `${cap}+`}
              </button>
            ))}
          </div>
        </div>
      </div>

      {/* Room Grid / Skeletons */}
      {loading ? (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          <RoomSkeleton />
          <RoomSkeleton />
          <RoomSkeleton />
        </div>
      ) : filteredRooms.length === 0 ? (
        <div className="text-center py-16 bg-gray-900/40 rounded-2xl border border-white/5">
          <p className="text-gray-400 text-sm">No meeting rooms found matching your search filters.</p>
        </div>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          {filteredRooms.map((room) => (
            <RoomCard
              key={room.id}
              room={room}
              schedule={schedules[room.emailAddress]}
              selectedDate={selectedDate}
              onBookClick={onBookClick}
            />
          ))}
        </div>
      )}
    </div>
  );
}
