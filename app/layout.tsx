import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'مولّد خطة الدرس - DOCX',
  description: 'إنشاء خطة درس وفق قالب الوزارة بصيغة DOCX'
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="ar" dir="rtl">
      <body>{children}</body>
    </html>
  );
}
