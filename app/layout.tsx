import type { Metadata } from 'next'
import './globals.css'

export const metadata: Metadata = {
  title: 'Excel Data Extractor',
  description: 'Extract data from Excel files',
}

export default function RootLayout({
  children,
}: {
  children: React.ReactNode
}) {
  return (
    <html lang="hi">
      <body>{children}</body>
    </html>
  )
}
