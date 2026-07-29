'use client';

import { useState, useEffect } from 'react';
import { Room } from '@/lib/types/room';
import { X, Calendar, Clock, MapPin, Users, Send } from 'lucide-react';

interface BookingModalProps {
  room: Room | null;
  initialStartTime?: string;
  isOpen: boolean;
  onClose: () => void;
  onSuccess: (msg: string) => void;
  onError: (msg: string) => void;
}

export function BookingModal({
  room,
  initialStartTime,
  isOpen,
  onClose,
  onSuccess,
  onError,
}: BookingModalProps) {
  const [subject, setSubject] = useState('');
  const [description, setDescription] = useState('');
  const [date, setDate] = useState('');
  const [startTime, setStartTime] = useState('09:00');
  const [endTime, setEndTime] = useState('10:00');
  const [attendees, setAttendees] = useState('');
  const [submitting, setSubmitting] = useState(false);

  useEffect(() => {
    if (initialStartTime) {
      const dt = new Date(initialStartTime);
      setDate(dt.toISOString().split('T')[0]);
      const hours = String(dt.getUTCHours()).padStart(2, '0');
      const mins = String(dt.getUTCMinutes()).padStart(2, '0');
      setStartTime(`${hours}:${mins}`);

      // Default end time 1 hour later
      const endDt = new Date(dt.getTime() + 60 * 60 * 1000);
      const endHours = String(endDt.getUTCHours()).padStart(2, '0');
      const endMins = String(endDt.getUTCMinutes()).padStart(2, '0');
      setEndTime(`${endHours}:${endMins}`);
    } else {
      setDate(new Date().toISOString().split('T')[0]);
    }
  }, [initialStartTime, isOpen]);

  if (!isOpen || !room) return null;

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!subject.trim()) {
      onError('Please enter a meeting title.');
      return;
    }

    setSubmitting(true);
    try {
      const startIso = new Date(`${date}T${startTime}:00Z`).toISOString();
      const endIso = new Date(`${date}T${endTime}:00Z`).toISOString();

      const attendeeList = attendees
        .split(',')
        .map((a) => a.trim())
        .filter((a) => a.includes('@'));

      const res = await fetch('/api/bookings', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          roomEmail: room.emailAddress,
          roomName: room.displayName,
          subject,
          bodyContent: description,
          startTime: startIso,
          endTime: endIso,
          attendees: attendeeList,
        }),
      });

      const data = await res.json();
      if (res.ok && data.success) {
        onSuccess(data.message || `Successfully reserved ${room.displayName}!`);
        onClose();
        setSubject('');
        setDescription('');
      } else {
        onError(data.message || 'Failed to complete room reservation.');
      }
    } catch (err: any) {
      onError(err.message || 'Network error occurred while booking.');
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/70 backdrop-blur-md animate-in fade-in duration-200">
      <div className="relative w-full max-w-lg bg-gray-900 border border-white/10 rounded-2xl p-6 shadow-2xl space-y-5">
        {/* Header */}
        <div className="flex justify-between items-start border-b border-white/10 pb-4">
          <div>
            <h2 className="text-xl font-bold text-white">Book {room.displayName}</h2>
            <p className="text-xs text-gray-400 flex items-center gap-1 mt-1">
              <MapPin className="w-3.5 h-3.5 text-cyan-400" />
              {room.building} &bull; Capacity: {room.capacity}
            </p>
          </div>
          <button
            onClick={onClose}
            className="p-1 rounded-lg text-gray-400 hover:text-white hover:bg-white/10 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Form */}
        <form onSubmit={handleSubmit} className="space-y-4 text-sm">
          <div>
            <label className="block text-xs font-semibold text-gray-300 mb-1">
              Meeting Title / Subject *
            </label>
            <input
              type="text"
              required
              placeholder="e.g. Q3 Product Strategy Review"
              value={subject}
              onChange={(e) => setSubject(e.target.value)}
              className="w-full bg-black/40 border border-white/10 rounded-xl px-3.5 py-2.5 text-white placeholder-gray-500 focus:outline-none focus:border-cyan-500 transition-colors"
            />
          </div>

          <div className="grid grid-cols-3 gap-3">
            <div>
              <label className="block text-xs font-semibold text-gray-300 mb-1">Date</label>
              <input
                type="date"
                required
                value={date}
                onChange={(e) => setDate(e.target.value)}
                className="w-full bg-black/40 border border-white/10 rounded-xl px-3 py-2 text-white text-xs focus:outline-none focus:border-cyan-500"
              />
            </div>
            <div>
              <label className="block text-xs font-semibold text-gray-300 mb-1">Start Time</label>
              <input
                type="time"
                required
                value={startTime}
                onChange={(e) => setStartTime(e.target.value)}
                className="w-full bg-black/40 border border-white/10 rounded-xl px-3 py-2 text-white text-xs focus:outline-none focus:border-cyan-500"
              />
            </div>
            <div>
              <label className="block text-xs font-semibold text-gray-300 mb-1">End Time</label>
              <input
                type="time"
                required
                value={endTime}
                onChange={(e) => setEndTime(e.target.value)}
                className="w-full bg-black/40 border border-white/10 rounded-xl px-3 py-2 text-white text-xs focus:outline-none focus:border-cyan-500"
              />
            </div>
          </div>

          <div>
            <label className="block text-xs font-semibold text-gray-300 mb-1">
              Optional Attendees (comma separated emails)
            </label>
            <input
              type="text"
              placeholder="colleague@company.com, manager@company.com"
              value={attendees}
              onChange={(e) => setAttendees(e.target.value)}
              className="w-full bg-black/40 border border-white/10 rounded-xl px-3.5 py-2 text-white placeholder-gray-500 focus:outline-none focus:border-cyan-500"
            />
          </div>

          <div>
            <label className="block text-xs font-semibold text-gray-300 mb-1">Description</label>
            <textarea
              rows={2}
              placeholder="Agenda or additional notes..."
              value={description}
              onChange={(e) => setDescription(e.target.value)}
              className="w-full bg-black/40 border border-white/10 rounded-xl px-3.5 py-2 text-white placeholder-gray-500 focus:outline-none focus:border-cyan-500 resize-none"
            />
          </div>

          {/* Footer Action */}
          <div className="flex items-center justify-end gap-3 pt-3 border-t border-white/10">
            <button
              type="button"
              onClick={onClose}
              className="px-4 py-2 rounded-xl text-gray-400 hover:text-white font-medium transition-colors"
            >
              Cancel
            </button>
            <button
              type="submit"
              disabled={submitting}
              className="px-5 py-2.5 rounded-xl bg-gradient-to-r from-cyan-500 to-blue-600 hover:from-cyan-400 hover:to-blue-500 text-black font-bold shadow-lg shadow-cyan-500/20 disabled:opacity-50 flex items-center gap-2 transition-all"
            >
              {submitting ? 'Confirming...' : 'Confirm Reservation'}
              <Send className="w-4 h-4" />
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}
