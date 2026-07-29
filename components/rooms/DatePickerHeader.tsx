'use client';

import { ChevronLeft, ChevronRight, Calendar as CalendarIcon } from 'lucide-react';
import { format, addDays, subDays, isToday, isTomorrow } from 'date-fns';

interface DatePickerHeaderProps {
  selectedDate: Date;
  onChangeDate: (date: Date) => void;
}

export function DatePickerHeader({ selectedDate, onChangeDate }: DatePickerHeaderProps) {
  const dateLabel = isToday(selectedDate)
    ? 'Today'
    : isTomorrow(selectedDate)
    ? 'Tomorrow'
    : format(selectedDate, 'EEEE, MMM d, yyyy');

  return (
    <div className="flex flex-col sm:flex-row items-center justify-between gap-4 bg-gradient-to-r from-gray-900/90 to-gray-950/90 p-4 rounded-2xl border border-white/10 backdrop-blur-xl">
      <div className="flex items-center gap-3">
        <div className="p-2.5 rounded-xl bg-cyan-500/10 border border-cyan-500/20 text-cyan-400">
          <CalendarIcon className="w-5 h-5" />
        </div>
        <div>
          <h2 className="text-xl font-bold text-white tracking-tight">{dateLabel}</h2>
          <p className="text-xs text-gray-400">{format(selectedDate, 'MMMM d, yyyy')}</p>
        </div>
      </div>

      <div className="flex items-center gap-2">
        <button
          onClick={() => onChangeDate(subDays(selectedDate, 1))}
          className="p-2 rounded-xl bg-white/5 border border-white/10 text-gray-300 hover:bg-white/10 transition-colors"
          title="Previous Day"
        >
          <ChevronLeft className="w-4 h-4" />
        </button>

        <button
          onClick={() => onChangeDate(new Date())}
          className={`px-3 py-1.5 rounded-xl text-xs font-semibold border transition-all ${
            isToday(selectedDate)
              ? 'bg-cyan-500 text-black border-cyan-400 font-bold'
              : 'bg-white/5 text-gray-300 border-white/10 hover:bg-white/10'
          }`}
        >
          Today
        </button>

        <button
          onClick={() => onChangeDate(addDays(new Date(), 1))}
          className={`px-3 py-1.5 rounded-xl text-xs font-semibold border transition-all ${
            isTomorrow(selectedDate)
              ? 'bg-cyan-500 text-black border-cyan-400 font-bold'
              : 'bg-white/5 text-gray-300 border-white/10 hover:bg-white/10'
          }`}
        >
          Tomorrow
        </button>

        <button
          onClick={() => onChangeDate(addDays(selectedDate, 1))}
          className="p-2 rounded-xl bg-white/5 border border-white/10 text-gray-300 hover:bg-white/10 transition-colors"
          title="Next Day"
        >
          <ChevronRight className="w-4 h-4" />
        </button>
      </div>
    </div>
  );
}
