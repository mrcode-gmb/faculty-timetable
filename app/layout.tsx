import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "ExamGrid | Faculty Examination Timetable",
  description: "Prepare faculty examination timetables from Excel courses and rooms while checking student and room clashes.",
  icons: {
    icon: "/favicon.svg",
    shortcut: "/favicon.svg",
  },
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
