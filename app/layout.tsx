export const metadata = {
  title: "Excel → Apple Takvim (.ics)",
  description: "Excel yükle, isim seç, .ics indir (tarayıcıda).",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="tr">
      <body>{children}</body>
    </html>
  );
}
