'use client';

import { signIn } from 'next-auth/react';
import { LogIn } from 'lucide-react';

export function LoginButton() {
  return (
    <button
      onClick={() => signIn('azure-ad')}
      className="inline-flex items-center gap-2 px-5 py-2.5 rounded-xl bg-gradient-to-r from-cyan-500 to-blue-600 text-white font-medium shadow-lg shadow-cyan-500/20 hover:shadow-cyan-500/40 hover:scale-[1.02] active:scale-[0.98] transition-all"
    >
      <LogIn className="w-4 h-4" />
      Sign in with Microsoft 365
    </button>
  );
}
