import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'HCT-COHS EHS KPI Dashboard',
  description: 'Health, Safety & Environment KPI Dashboard',
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
