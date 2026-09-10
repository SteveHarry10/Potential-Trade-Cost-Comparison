import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'Trade Cost Dashboard',
  description: 'Compare potential trade bids with the latest plan costs from your Floorplans workbook.',
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body className="antialiased">{children}</body>
    </html>
  );
}
