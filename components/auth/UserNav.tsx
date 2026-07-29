'use client';

import { useSession, signOut, signIn } from 'next-auth/react';
import { LogOut, User as UserIcon } from 'lucide-react';

export function UserNav() {
  const { data: session } = useSession();

  if (!session?.user) {
    return (
      <button
        onClick={() => signIn('azure-ad')}
        className="text-sm font-medium px-4 py-2 rounded-lg bg-cyan-500/10 text-cyan-400 border border-cyan-500/20 hover:bg-cyan-500/20 transition-all"
      >
        Sign In
      </button>
    );
  }

  return (
    <div className="flex items-center gap-3 bg-white/5 border border-white/10 px-3 py-1.5 rounded-xl backdrop-blur-md">
      {session.user.image ? (
        <img
          src={session.user.image}
          alt={session.user.name || 'User'}
          className="w-8 h-8 rounded-full border border-cyan-500/40 object-cover"
        />
      ) : (
        <div className="w-8 h-8 rounded-full bg-cyan-500/20 text-cyan-400 flex items-center justify-center font-bold text-xs">
          {session.user.name?.[0] || 'U'}
        </div>
      )}
      <div className="hidden sm:block text-left">
        <p className="text-xs font-semibold text-gray-200">{session.user.name}</p>
        <p className="text-[10px] text-gray-400 truncate max-w-[120px]">{session.user.email}</p>
      </div>
      <button
        onClick={() => signOut()}
        title="Sign Out"
        className="text-gray-400 hover:text-rose-400 p-1.5 rounded-lg hover:bg-white/5 transition-colors ml-1"
      >
        <LogOut className="w-4 h-4" />
      </button>
    </div>
  );
}
