import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'HCT-COHS EHS KPI Dashboard',
  description: 'Health, Safety & Environment KPI Dashboard',
  icons: {
    icon: '/icon.svg',
  },
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <head>
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css" crossOrigin="anonymous" referrerPolicy="no-referrer" />
      </head>
      <body>{children}</body>
    </html>
  );
}
