import type { Config } from 'tailwindcss';

const config: Config = {
  darkMode: 'class',
  content: [
    './pages/**/*.{js,ts,jsx,tsx,mdx}',
    './components/**/*.{js,ts,jsx,tsx,mdx}',
    './app/**/*.{js,ts,jsx,tsx,mdx}',
  ],
  theme: {
    extend: {
      colors: {
        darkBg: '#090D16',
        cardBg: '#111827',
        cardBorder: 'rgba(255, 255, 255, 0.08)',
        brandCyan: '#06B6D4',
        brandEmerald: '#10B981',
        brandRose: '#EF4444',
        brandAmber: '#F59E0B',
      },
    },
  },
  plugins: [],
};

export default config;
