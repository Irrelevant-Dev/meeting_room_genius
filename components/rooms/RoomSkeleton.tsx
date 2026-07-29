export function RoomSkeleton() {
  return (
    <div className="bg-gray-900/60 border border-white/10 rounded-2xl p-5 animate-pulse space-y-4">
      <div className="flex justify-between items-start">
        <div className="h-6 w-1/2 bg-gray-800 rounded-lg"></div>
        <div className="h-5 w-20 bg-gray-800 rounded-full"></div>
      </div>
      <div className="h-4 w-1/3 bg-gray-800 rounded"></div>
      <div className="h-10 w-full bg-gray-800/80 rounded-xl"></div>
      <div className="h-9 w-full bg-gray-800 rounded-xl"></div>
    </div>
  );
}
