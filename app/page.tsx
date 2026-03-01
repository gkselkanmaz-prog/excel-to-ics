"use client";

import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";

type RowObj = Record<string, any>;
type Task = { date: Date; unit: string };

function normalizeName(s: string) {
  return (s ?? "")
    .toString()
    .trim()
    .toUpperCase()
    .replaceAll("İ", "I")
    .replaceAll("İ", "I")
    .replaceAll("Ş", "S")
    .replaceAll("Ğ", "G")
    .replaceAll("Ü", "U")
    .replaceAll("Ö", "O")
    .replaceAll("Ç", "C");
}

function pad(n: number) {
  return String(n).padStart(2, "0");
}

function toICSDateUTC(d: Date) {
  return (
    d.getUTCFullYear() +
    pad(d.getUTCMonth() + 1) +
    pad(d.getUTCDate()) +
    "T" +
    pad(d.getUTCHours()) +
    pad(d.getUTCMinutes()) +
    pad(d.getUTCSeconds()) +
    "Z"
  );
}

function toDateValue(d: Date) {
  return d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate());
}

function escapeICS(text: string) {
  return (text ?? "")
    .replaceAll("\\", "\\\\")
    .replaceAll("\n", "\\n")
    .replaceAll(",", "\\,")
    .replaceAll(";", "\\;");
}

function buildICS(events: { date: Date; title: string; description?: string }[]) {
  const dtstamp = toICSDateUTC(new Date());

  const lines: string[] = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//ExcelToICS//TR",
    "CALSCALE:GREGORIAN",
    "METHOD:PUBLISH",
  ];

  for (const e of events) {
    const uid =
      typeof crypto !== "undefined" && "randomUUID" in crypto
        ? crypto.randomUUID()
        : String(Math.random()).slice(2) + "@exceltoics";

    // All-day event: start date -> next day as DTEND
    const startDate = new Date(e.date.getFullYear(), e.date.getMonth(), e.date.getDate());
    const endDate = new Date(startDate);
    endDate.setDate(endDate.getDate() + 1);

    lines.push("BEGIN:VEVENT");
    lines.push(`UID:${uid}`);
    lines.push(`DTSTAMP:${dtstamp}`);
    lines.push(`DTSTART;VALUE=DATE:${toDateValue(startDate)}`);
    lines.push(`DTEND;VALUE=DATE:${toDateValue(endDate)}`);
    lines.push(`SUMMARY:${escapeICS(e.title)}`);
    if (e.description) lines.push(`DESCRIPTION:${escapeICS(e.description)}`);
    lines.push("END:VEVENT");
  }

  lines.push("END:VCALENDAR");
  return lines.join("\r\n");
}

function downloadTextFile(filename: string, content: string, mime = "text/calendar") {
  const blob = new Blob([content], { type: `${mime};charset=utf-8` });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function asDate(value: any): Date | null {
  if (!value) return null;

  if (value instanceof Date && !isNaN(value.getTime())) return value;

  if (typeof value === "number") {
    const d = XLSX.SSF.parse_date_code(value);
    if (d && d.y && d.m && d.d) return new Date(d.y, d.m - 1, d.d);
  }

  const str = value.toString().trim();

  const m1 = str.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
  if (m1) return new Date(Number(m1[3]), Number(m1[2]) - 1, Number(m1[1]));

  const m2 = str.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m2) return new Date(Number(m2[3]), Number(m2[2]) - 1, Number(m2[1]));

  const m3 = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m3) return new Date(Number(m3[1]), Number(m3[2]) - 1, Number(m3[3]));

  return null;
}

function sheetToRowsWithSmartHeader(ws: XLSX.WorkSheet): RowObj[] {
  const grid: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  if (!grid.length) return [];

  const headerKeyRowIdx = grid.findIndex((r) =>
    (r ?? []).some((c) => {
      const t = normalizeName(String(c ?? ""));
      return t.includes("TARİH") || t.includes("TARIH") || t.includes("DATE");
    })
  );

  const baseIdx = headerKeyRowIdx >= 0 ? headerKeyRowIdx : 0;
  const headerRow = grid[baseIdx + 1] ?? grid[baseIdx] ?? [];
  const dataStart = baseIdx + 2;

  let dateCol = headerRow.findIndex((c) => {
    const t = normalizeName(String(c ?? ""));
    return t === "TARİH" || t === "TARIH" || t.includes("TARİH") || t.includes("TARIH") || t.includes("DATE");
  });
  if (dateCol < 0) dateCol = 0;

  const labels = headerRow.map((c, idx) => {
    const raw = (c ?? "").toString().trim();
    if (idx === dateCol) return "TARİH";
    if (!raw) return `COL_${idx}`;
    return raw;
  });

  const rows: RowObj[] = [];
  for (let r = dataStart; r < grid.length; r++) {
    const row = grid[r];
    if (!row || row.every((v) => v === null || v === "")) continue;

    const obj: RowObj = {};
    for (let c = 0; c < labels.length; c++) obj[labels[c]] = row[c] ?? null;

    const dt = asDate(obj["TARİH"]);
    if (!dt) continue;
    obj["TARİH"] = dt;

    rows.push(obj);
  }

  return rows;
}

export default function Page() {
  const [rows, setRows] = useState<RowObj[]>([]);
  const [fileName, setFileName] = useState<string>("");
  const [personInput, setPersonInput] = useState<string>("");
  const [personPicked, setPersonPicked] = useState<string>("");

  const people = useMemo(() => {
    const set = new Set<string>();
    for (const r of rows) {
      for (const [k, v] of Object.entries(r)) {
        if (k === "TARİH") continue;
        if (typeof v !== "string") continue;
        const s = v.trim();
        if (!s) continue;
        if (s.length >= 8 && s.includes(" ")) set.add(s);
      }
    }
    return Array.from(set).sort((a, b) => a.localeCompare(b, "tr"));
  }, [rows]);

  const tasks = useMemo(() => {
    const target = normalizeName(personPicked || personInput);
    if (!target || !rows.length) return [];

    const out: Task[] = [];
    for (const r of rows) {
      const dt: Date = r["TARİH"];
      for (const [unit, val] of Object.entries(r)) {
        if (unit === "TARİH") continue;
        if (!val) continue;
        if (normalizeName(String(val)) === target) out.push({ date: dt, unit });
      }
    }

    out.sort((a, b) => a.date.getTime() - b.date.getTime() || a.unit.localeCompare(b.unit, "tr"));
    return out;
  }, [rows, personPicked, personInput]);

  async function onFile(file: File) {
    setFileName(file.name);
    setPersonPicked("");
    setPersonInput("");

    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const parsedRows = sheetToRowsWithSmartHeader(ws);
    setRows(parsedRows);
  }

  function makeICS() {
    const who = (personPicked || personInput).trim();
    if (!who) return alert("Lütfen isim yaz veya listeden seç.");
    if (!tasks.length) return alert("Bu isimle eşleşen görev bulunamadı. Yazımı kontrol et.");

    const ics = buildICS(
      tasks.map((t) => ({
        date: t.date,
        title: `${who} - ${t.unit}`,
        description: `${fileName || "Excel"} üzerinden üretildi (tarayıcıda).`,
      }))
    );

    const safe = who
      .replace(/[^A-Za-z0-9 _-ğüşöçıİĞÜŞÖÇ]/g, "")
      .trim()
      .replace(/\s+/g, "_");

    downloadTextFile(`${safe || "takvim"}.ics`, ics);
  }

  return (
    <main style={{ maxWidth: 980, margin: "40px auto", padding: 16, fontFamily: "system-ui, -apple-system, Segoe UI, Roboto" }}>
      <h1 style={{ fontSize: 28, marginBottom: 8 }}>Excel → Apple Takvim (.ics)</h1>
      <p style={{ marginTop: 0, color: "#444" }}>
        Excel dosyanı yükle, ismi seç/yaz, takvime eklenebilir <b>.ics</b> dosyasını indir. (Dosya sunucuya gitmez.)
      </p>

      <section style={{ padding: 16, border: "1px solid #ddd", borderRadius: 12, marginBottom: 16 }}>
        <h2 style={{ fontSize: 18, marginTop: 0 }}>1) Excel yükle</h2>
        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={(e) => {
            const f = e.target.files?.[0];
            if (f) onFile(f);
          }}
        />
        {rows.length > 0 && (
          <div style={{ marginTop: 10, color: "#333" }}>
            Okunan satır sayısı: <b>{rows.length}</b>
          </div>
        )}
      </section>

      <section style={{ padding: 16, border: "1px solid #ddd", borderRadius: 12, marginBottom: 16 }}>
        <h2 style={{ fontSize: 18, marginTop: 0 }}>2) İsim seç / yaz</h2>

        <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
          <div style={{ flex: "1 1 320px" }}>
            <label style={{ display: "block", fontSize: 13, color: "#444", marginBottom: 6 }}>Listeden seç</label>
            <select
              disabled={!people.length}
              value={personPicked}
              onChange={(e) => setPersonPicked(e.target.value)}
              style={{ width: "100%", padding: 10, borderRadius: 10, border: "1px solid #ccc" }}
            >
              <option value="">{people.length ? "Seç..." : "Önce Excel yükle"}</option>
              {people.map((p) => (
                <option key={p} value={p}>
                  {p}
                </option>
              ))}
            </select>
          </div>

          <div style={{ flex: "1 1 320px" }}>
            <label style={{ display: "block", fontSize: 13, color: "#444", marginBottom: 6 }}>Veya isim yaz (tam eşleşme)</label>
            <input
              value={personInput}
              onChange={(e) => setPersonInput(e.target.value)}
              placeholder="Örn: AHKAM GÖKSEL KANMAZ"
              style={{ width: "100%", padding: 10, borderRadius: 10, border: "1px solid #ccc" }}
            />
            <div style={{ fontSize: 12, color: "#666", marginTop: 6 }}>
              Büyük/küçük harf ve Türkçe karakter toleranslı (tam eşleşme).
            </div>
          </div>
        </div>
      </section>

      <section style={{ padding: 16, border: "1px solid #ddd", borderRadius: 12 }}>
        <h2 style={{ fontSize: 18, marginTop: 0 }}>3) Önizleme & İndir</h2>

        <div style={{ marginBottom: 10 }}>
          Bulunan görev sayısı: <b>{tasks.length}</b>
        </div>

        {tasks.length > 0 && (
          <div style={{ maxHeight: 240, overflow: "auto", border: "1px solid #eee", borderRadius: 10, padding: 10, marginBottom: 12 }}>
            <ul style={{ margin: 0, paddingLeft: 18 }}>
              {tasks.slice(0, 200).map((t, i) => (
                <li key={i}>
                  {t.date.toLocaleDateString("tr-TR")} — {t.unit}
                </li>
              ))}
            </ul>
            {tasks.length > 200 && <div style={{ fontSize: 12, color: "#666" }}>İlk 200 satır gösteriliyor…</div>}
          </div>
        )}

        <button
          onClick={makeICS}
          style={{
            padding: "12px 14px",
            borderRadius: 12,
            border: "1px solid #333",
            background: "#111",
            color: "white",
            cursor: "pointer",
            fontWeight: 600,
          }}
        >
          .ics indir
        </button>
      </section>

      <footer style={{ marginTop: 18, fontSize: 12, color: "#666" }}>
        Bu MVP tamamen tarayıcıda çalışır; Excel dosyası sunucuya gönderilmez.
      </footer>
    </main>
  );
}
