'use client';

import { useState, useEffect } from 'react';
import { Room, RoomSchedule } from '@/lib/types/room';
import { UserNav } from '@/components/auth/UserNav';
import { RoomGrid } from '@/components/rooms/RoomGrid';
import { DatePickerHeader } from '@/components/rooms/DatePickerHeader';
import { BookingModal } from '@/components/rooms/BookingModal';
import { Toast, ToastMessage } from '@/components/ui/Toast';
import { Building2, Sparkles } from 'lucide-react';

export default function DashboardPage() {
  const [rooms, setRooms] = useState<Room[]>([]);
  const [schedules, setSchedules] = useState<Record<string, RoomSchedule>>({});
  const [loading, setLoading] = useState<boolean>(true);
  const [selectedDate, setSelectedDate] = useState<Date>(new Date());

  const [modalOpen, setModalOpen] = useState<boolean>(false);
  const [selectedRoom, setSelectedRoom] = useState<Room | null>(null);
  const [initialSlotTime, setInitialSlotTime] = useState<string | undefined>();

  const [toasts, setToasts] = useState<ToastMessage[]>([]);

  const addToast = (type: 'success' | 'error', text: string) => {
    setToasts((prev) => [...prev, { id: String(Date.now()), type, text }]);
  };

  const removeToast = (id: string) => {
    setToasts((prev) => prev.filter((t) => t.id !== id));
  };

  const loadData = async () => {
    setLoading(true);
    try {
      const roomRes = await fetch('/api/rooms');
      const roomData = await roomRes.json();
      const loadedRooms: Room[] = roomData.rooms || [];
      setRooms(loadedRooms);

      if (loadedRooms.length > 0) {
        const emails = loadedRooms.map((r) => r.emailAddress);
        const schedRes = await fetch('/api/rooms/schedule', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ roomEmails: emails, date: selectedDate.toISOString() }),
        });
        const schedData = await schedRes.json();
        setSchedules(schedData.schedules || {});
      }
    } catch (err) {
      console.error('Error fetching dashboard data:', err);
      addToast('error', 'Failed to load room availability schedules.');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadData();
  }, [selectedDate]);

  const handleOpenBooking = (room: Room, startTime?: string) => {
    setSelectedRoom(room);
    setInitialSlotTime(startTime);
    setModalOpen(true);
  };

  return (
    <div className="min-h-screen bg-darkBg flex flex-col">
      {/* Top Navbar */}
      <header className="border-b border-white/10 bg-gray-950/80 backdrop-blur-xl sticky top-0 z-40">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-3.5 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-gradient-to-tr from-cyan-500 to-blue-600 rounded-xl shadow-lg shadow-cyan-500/20 text-black">
              <Building2 className="w-5 h-5 font-bold" />
            </div>
            <div>
              <h1 className="text-lg font-bold text-white tracking-tight flex items-center gap-2">
                Meeting Room Genius
                <span className="text-[10px] font-semibold px-2 py-0.5 rounded-full bg-cyan-500/10 text-cyan-400 border border-cyan-500/20">
                  M365 Live
                </span>
              </h1>
              <p className="text-xs text-gray-400">Microsoft 365 Executive Booking Suite</p>
            </div>
          </div>

          <UserNav />
        </div>
      </header>

      {/* Main Content Area */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-8">
        {/* Hero Section */}
        <div className="relative rounded-3xl p-6 sm:p-8 bg-gradient-to-r from-gray-900 via-gray-900/90 to-gray-950 border border-white/10 shadow-2xl overflow-hidden">
          <div className="absolute right-0 top-0 w-96 h-96 bg-cyan-500/10 rounded-full blur-3xl -mr-20 -mt-20 pointer-events-none"></div>
          <div className="relative z-10 space-y-2">
            <div className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-cyan-500/10 border border-cyan-500/30 text-cyan-400 text-xs font-semibold">
              <Sparkles className="w-3.5 h-3.5" /> Direct Exchange Online Integration
            </div>
            <h2 className="text-2xl sm:text-3xl font-extrabold text-white">
              Find & Reserve Your Next Meeting Room
            </h2>
            <p className="text-sm text-gray-400 max-w-2xl">
              Real-time availability timelines, capacity filters, and instant delegated booking powered by Microsoft Graph API.
            </p>
          </div>
        </div>

        {/* Date Selector Header */}
        <DatePickerHeader selectedDate={selectedDate} onChangeDate={setSelectedDate} />

        {/* Room Discovery & Grid */}
        <RoomGrid
          rooms={rooms}
          schedules={schedules}
          loading={loading}
          selectedDate={selectedDate}
          onBookClick={handleOpenBooking}
        />
      </main>

      {/* Booking Modal Dialog */}
      <BookingModal
        room={selectedRoom}
        initialStartTime={initialSlotTime}
        isOpen={modalOpen}
        onClose={() => setModalOpen(false)}
        onSuccess={(msg) => {
          addToast('success', msg);
          loadData();
        }}
        onError={(msg) => addToast('error', msg)}
      />

      {/* Toast Overlay */}
      <Toast toasts={toasts} onDismiss={removeToast} />
    </div>
  );
}
