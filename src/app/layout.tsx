
import type {Metadata} from 'next';
import { Inter } from 'next/font/google'; // Changed from Geist to Inter
import './globals.css';
import { Toaster } from "@/components/ui/toaster" // Import Toaster

const inter = Inter({ subsets: ['latin'], variable: '--font-sans' }); // Use Inter

export const metadata: Metadata = {
  title: 'SCA - Sistema para conversão de arquivos v1.2.0', // Updated App Name and version
  description: 'Converta arquivos Excel (XLS, XLSX, ODS) para layouts TXT ou CSV personalizados. Seus dados não são armazenados, garantindo conformidade com a LGPD.', // Updated description with LGPD note
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    // suppressHydrationWarning is kept as it was previously added to handle intermittent whitespace issues.
    // The main fix here is removing the comment below, which might be the source of the whitespace.
    <html lang="pt-BR" suppressHydrationWarning> {/* Default language to Portuguese & suppress warning */}
      <body className={`${inter.variable} font-sans antialiased`}>
        {children}
        <Toaster /> {/* Add Toaster component here */}
      </body>
    </html>
  );
}

