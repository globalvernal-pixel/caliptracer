/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        navy: {
          50: '#F0F4F8',
          100: '#D9E2EC',
          200: '#BCCCDC',
          500: '#2B4C7E',
          700: '#1A365D',
          800: '#102A43',
          900: '#0F172A',
        },
        primaryNavy: '#1A365D',
        primaryHover: '#2A4365',
        slateText: '#333333',
        appBg: '#F8F9FA',
        cardBg: '#FFFFFF',
        panelBg: '#FFFFFF',
        panelBorder: '#E2E8F0',
        brand: {
          DEFAULT: "#1A365D",
          hover: "#2A4365",
          glow: "rgba(26, 54, 93, 0.12)",
        },
        starColor: {
          DEFAULT: "#D97706",
          glow: "rgba(217, 119, 6, 0.12)",
        },
        tallyColor: {
          DEFAULT: "#0284C7",
          glow: "rgba(2, 132, 199, 0.12)",
        }
      },
      fontFamily: {
        sans: ['Inter', 'Poppins', 'Plus Jakarta Sans', 'system-ui', 'sans-serif'],
      },
      borderRadius: {
        'xl': '16px',
        'lg': '12px',
        'md': '8px',
        'sm': '6px',
      },
      boxShadow: {
        'card': '0 2px 8px -2px rgba(26, 54, 93, 0.06), 0 1px 4px -1px rgba(0, 0, 0, 0.04)',
        'card-hover': '0 8px 20px -4px rgba(26, 54, 93, 0.12), 0 4px 6px -2px rgba(0, 0, 0, 0.04)',
        'header': '0 4px 14px 0 rgba(26, 54, 93, 0.15)',
        'bottom-nav': '0 -4px 16px rgba(0, 0, 0, 0.06)',
      }
    },
  },
  plugins: [],
}
