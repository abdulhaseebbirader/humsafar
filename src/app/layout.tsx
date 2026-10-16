import type { Metadata, Viewport } from "next";
import ServiceWorkerRegistrar from "@/components/ServiceWorkerRegistrar";
import "./globals.css";

export const metadata: Metadata = {
  title: "HumsafarAI — Settlement reconciliation",
  description: "Turn a payment-advice PDF and a detailed Excel export into one defensible settlement answer.",
  appleWebApp: {
    capable: true,
    title: "HumsafarAI",
    statusBarStyle: "black-translucent",
  },
  icons: {
    icon: [{ url: "/icons/icon-192.png", sizes: "192x192", type: "image/png" }],
    apple: [{ url: "/icons/apple-touch-icon.png", sizes: "180x180", type: "image/png" }],
  },
};

export const viewport: Viewport = {
  themeColor: "#102e2d",
  width: "device-width",
  initialScale: 1,
  viewportFit: "cover",
};

export default function RootLayout({ children }: LayoutProps<"/">) {
  return (
    <html lang="en">
      <body>
        {children}
        <ServiceWorkerRegistrar />
      </body>
    </html>
  );
}
