import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "HumsafarAI — Settlement reconciliation",
  description: "Turn a payment-advice PDF and a detailed Excel export into one defensible settlement answer.",
};

export default function RootLayout({ children }: LayoutProps<"/">) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
