import React, { useState, useCallback, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell,
} from "recharts";

// ---------------------------------------------------------------------------
// Design tokens
// ---------------------------------------------------------------------------
const COLOR = {
  bg: "#F4F5F1",
  surface: "#FFFFFF",
  ink: "#1B2321",
  inkSoft: "#5B655F",
  line: "#DCE1DB",
  teal: "#0E6E6E",
  tealDark: "#0A4F4F",
  tealSoft: "#DCEEEC",
  coral: "#C85A3E",
  amber: "#D89A3C",
  grey: "#9AA5A0",
  greySoft: "#EDEFEC",
};

const SANS =
  "'Inter', ui-sans-serif, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif";
const MONO =
  "ui-monospace, 'SF Mono', 'Cascadia Mono', Menlo, Consolas, monospace";

// ---------------------------------------------------------------------------
// Domain logic
// ---------------------------------------------------------------------------
const SPM_GROUPS = [
  "Ibu Hamil", "Populasi Umum", "Penyakit TB", "Pasangan Risti", "WPS", "LSL", "Penyakit IMS",
  "Pelanggan PS", "Waria", "Pasangan ODHIV", "Penyakit Hepatitis",
  "Anak dari Ibu ODHIV", "WBP",
];
const WANTED_HEADERS = [
  "No", "Nama UPK", "Kelompok Populasi", "Jenis Layanan", "Status ODHIV",
];

function isFacilityCell(v) {
  if (v === null || v === undefined) return false;
  const s = String(v).trim();
  if (s === "") return false;
  return isNaN(Number(s));
}

function isSPM(kp) {
  if (!kp) return false;
  const s = String(kp);
  return SPM_GROUPS.some((g) => s.includes(g));
}

function findMeta(aoa) {
  // Scans the first ~10 rows for label/value pairs like "Provinsi" | "Jawa Barat"
  const wanted = ["Provinsi", "Kabupaten/Kota", "Periode"];
  const meta = {};
  for (let i = 0; i < Math.min(aoa.length, 12); i++) {
    const row = aoa[i] || [];
    for (let c = 0; c < row.length; c++) {
      const label = row[c] == null ? "" : String(row[c]).trim();
      if (wanted.includes(label)) {
        for (let c2 = c + 1; c2 < row.length; c2++) {
          if (row[c2] != null && String(row[c2]).trim() !== "") {
            meta[label] = String(row[c2]).trim();
            break;
          }
        }
      }
    }
  }
  return meta;
}

function parseWorkbook(aoa) {
  let headerRow = -1;
  let colMap = {};
  for (let i = 0; i < Math.min(aoa.length, 40); i++) {
    const row = aoa[i] || [];
    const found = WANTED_HEADERS.filter((w) =>
      row.some((c) => c != null && String(c).trim() === w)
    );
    if (found.length >= 4) {
      headerRow = i;
      row.forEach((c, idx) => {
        if (c != null) colMap[String(c).trim()] = idx;
      });
      break;
    }
  }
  if (headerRow === -1) {
    throw new Error(
      "Format tidak dikenali: kolom 'No', 'Nama UPK', 'Kelompok Populasi', dst tidak ditemukan. Pastikan file adalah Register Tes HIV dengan format standar."
    );
  }
  const need = ["No", "Nama UPK", "Kelompok Populasi", "Jenis Layanan", "Status ODHIV"];
  const missing = need.filter((k) => !(k in colMap));
  if (missing.length) {
    throw new Error(`Kolom berikut tidak ditemukan di header: ${missing.join(", ")}`);
  }

  let dataStart = headerRow + 1;
  while (dataStart < aoa.length && !isFacilityCell((aoa[dataStart] || [])[colMap["Nama UPK"]])) {
    dataStart++;
  }

  const rows = [];
  for (let i = dataStart; i < aoa.length; i++) {
    const row = aoa[i] || [];
    const upk = row[colMap["Nama UPK"]];
    if (!isFacilityCell(upk)) continue;
    rows.push({
      upk: String(upk).trim(),
      kelompokPopulasi: row[colMap["Kelompok Populasi"]] != null ? String(row[colMap["Kelompok Populasi"]]) : "",
      jenisLayanan: row[colMap["Jenis Layanan"]] != null ? String(row[colMap["Jenis Layanan"]]) : "",
      statusOdhiv: row[colMap["Status ODHIV"]] != null ? String(row[colMap["Status ODHIV"]]).trim() : "",
    });
  }
  if (rows.length === 0) {
    throw new Error("Header ditemukan tapi tidak ada baris data di bawahnya.");
  }
  return { headerRow, colMap, headerRowRaw: aoa[headerRow], rows, meta: findMeta(aoa) };
}

function summarize(rows) {
  const byUpk = new Map();
  let totalOdhiv = 0, totalBelumTahu = 0, totalSpm = 0;

  for (const r of rows) {
    if (!byUpk.has(r.upk)) {
      byUpk.set(r.upk, { upk: r.upk, total: 0, odhiv: 0, belumTahu: 0, bukanOdhiv: 0, spm: 0 });
    }
    const b = byUpk.get(r.upk);
    b.total += 1;
    if (r.statusOdhiv === "ODHIV") { b.odhiv += 1; totalOdhiv += 1; }
    else if (r.statusOdhiv === "Belum Tahu") { b.belumTahu += 1; totalBelumTahu += 1; }
    else if (r.statusOdhiv === "Bukan ODHIV") { b.bukanOdhiv += 1; }
    if (isSPM(r.kelompokPopulasi)) { b.spm += 1; totalSpm += 1; }
  }

  const perLayanan = Array.from(byUpk.values()).sort((a, b) => b.total - a.total);
  const total = rows.length;
  return {
    total, totalOdhiv, totalBelumTahu, totalSpm,
    totalNonSpm: total - totalSpm,
    layananCount: perLayanan.length,
    perLayanan,
  };
}

function buildWorkbook(parsed, summary) {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Data (raw + helper column)
  const dataAoa = [[...parsed.headerRowRaw.map((h) => (h == null ? "" : h)), "Kategori Populasi (SPM)"]];
  for (const r of parsed.rows) {
    const base = [];
    base[parsed.colMap["Nama UPK"]] = r.upk;
    base[parsed.colMap["Kelompok Populasi"]] = r.kelompokPopulasi;
    base[parsed.colMap["Jenis Layanan"]] = r.jenisLayanan;
    base[parsed.colMap["Status ODHIV"]] = r.statusOdhiv;
    const row = [];
    for (let i = 0; i < parsed.headerRowRaw.length; i++) row[i] = base[i] !== undefined ? base[i] : "";
    row.push(isSPM(r.kelompokPopulasi) ? "SPM" : "Non-SPM (Calon Pengantin)");
    dataAoa.push(row);
  }
  const wsData = XLSX.utils.aoa_to_sheet(dataAoa);
  XLSX.utils.book_append_sheet(wb, wsData, "Data");

  // Sheet 2: Rekap per Layanan
  const rekapAoa = [["No", "Nama UPK / Layanan", "Total Kunjungan/Tes", "ODHIV", "Belum Tahu Status", "Bukan ODHIV", "Capaian SPM", "% SPM"]];
  summary.perLayanan.forEach((r, i) => {
    rekapAoa.push([i + 1, r.upk, r.total, r.odhiv, r.belumTahu, r.bukanOdhiv, r.spm, r.total ? r.spm / r.total : 0]);
  });
  rekapAoa.push([null, "TOTAL", summary.total, summary.totalOdhiv, summary.totalBelumTahu, null, summary.totalSpm, summary.total ? summary.totalSpm / summary.total : 0]);
  const wsRekap = XLSX.utils.aoa_to_sheet(rekapAoa);
  XLSX.utils.book_append_sheet(wb, wsRekap, "Rekap per Layanan");

  // Sheet 3: Rekap Kelompok Populasi
  const popAoa = [
    ["No", "Indikator", "Jumlah Tes HIV", "% dari Total Tes"],
    [1, "Capaian SPM (di luar Calon Pengantin)", summary.totalSpm, summary.total ? summary.totalSpm / summary.total : 0],
    [2, "Non-SPM (Calon Pengantin)", summary.totalNonSpm, summary.total ? summary.totalNonSpm / summary.total : 0],
    [3, "Capaian Tes Seluruh Kelompok Populasi (Total)", summary.total, 1],
  ];
  const wsPop = XLSX.utils.aoa_to_sheet(popAoa);
  XLSX.utils.book_append_sheet(wb, wsPop, "Rekap Kelompok Populasi");

  return wb;
}

// ---------------------------------------------------------------------------
// UI bits
// ---------------------------------------------------------------------------
function CoverageRail({ spm, total, height = 8, width = 90 }) {
  const pct = total ? spm / total : 0;
  return (
    <div style={{ width, height, borderRadius: height / 2, overflow: "hidden", background: COLOR.greySoft, display: "flex" }}>
      <div style={{ width: `${pct * 100}%`, background: COLOR.teal }} />
      <div style={{ flex: 1, background: COLOR.grey, opacity: 0.35 }} />
    </div>
  );
}

function StatCard({ eyebrow, value, sub, accent }) {
  return (
    <div style={{
      background: COLOR.surface, border: `1px solid ${COLOR.line}`, borderRadius: 10,
      padding: "16px 18px", flex: "1 1 160px", minWidth: 150,
    }}>
      <div style={{ fontSize: 10.5, letterSpacing: "0.09em", textTransform: "uppercase", color: COLOR.inkSoft, fontFamily: SANS, marginBottom: 8 }}>
        {eyebrow}
      </div>
      <div style={{ fontFamily: MONO, fontSize: 26, fontWeight: 700, color: accent || COLOR.ink, fontVariantNumeric: "tabular-nums" }}>
        {value}
      </div>
      {sub && <div style={{ fontSize: 12, color: COLOR.inkSoft, marginTop: 4, fontFamily: SANS }}>{sub}</div>}
    </div>
  );
}

const fmt = (n) => new Intl.NumberFormat("id-ID").format(n);
const pct = (n, d) => (d ? `${((n / d) * 100).toFixed(1)}%` : "0%");

export default function App() {
  const [parsed, setParsed] = useState(null);
  const [summary, setSummary] = useState(null);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [sortKey, setSortKey] = useState("total");
  const [sortDir, setSortDir] = useState("desc");
  const [query, setQuery] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const fileInputRef = useRef(null);

  const handleFile = useCallback((file) => {
    if (!file) return;
    setError("");
    setLoading(true);
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array", cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: true });
        const p = parseWorkbook(aoa);
        const s = summarize(p.rows);
        setParsed(p);
        setSummary(s);
      } catch (err) {
        setError(err.message || "Gagal membaca file.");
        setParsed(null);
        setSummary(null);
      } finally {
        setLoading(false);
      }
    };
    reader.onerror = () => {
      setError("Gagal membaca file.");
      setLoading(false);
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const onDrop = useCallback((e) => {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files && e.dataTransfer.files[0];
    handleFile(file);
  }, [handleFile]);

  const sortedRows = useMemo(() => {
    if (!summary) return [];
    let rows = summary.perLayanan;
    if (query.trim()) {
      const q = query.trim().toLowerCase();
      rows = rows.filter((r) => r.upk.toLowerCase().includes(q));
    }
    const dir = sortDir === "asc" ? 1 : -1;
    return [...rows].sort((a, b) => {
      if (sortKey === "upk") return a.upk.localeCompare(b.upk) * dir;
      return (a[sortKey] - b[sortKey]) * dir;
    });
  }, [summary, sortKey, sortDir, query]);

  const topChartData = useMemo(() => {
    if (!summary) return [];
    return summary.perLayanan.slice(0, 12).map((r) => ({
      name: r.upk.length > 22 ? r.upk.slice(0, 20) + "…" : r.upk,
      Total: r.total,
    })).reverse();
  }, [summary]);

  const pieData = useMemo(() => {
    if (!summary) return [];
    return [
      { name: "Capaian SPM", value: summary.totalSpm },
      { name: "Non-SPM (Catin/Pop. Umum)", value: summary.totalNonSpm },
    ];
  }, [summary]);

  const handleDownload = () => {
    if (!parsed || !summary) return;
    const wb = buildWorkbook(parsed, summary);
    XLSX.writeFile(wb, `hasil_rekap_hiv_${(fileName || "data").replace(/\.[^/.]+$/, "")}.xlsx`);
  };

  const reset = () => {
    setParsed(null);
    setSummary(null);
    setFileName("");
    setError("");
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  const toggleSort = (key) => {
    if (sortKey === key) setSortDir(sortDir === "asc" ? "desc" : "asc");
    else { setSortKey(key); setSortDir("desc"); }
  };

  const arrow = (key) => (sortKey === key ? (sortDir === "asc" ? " ↑" : " ↓") : "");

  return (
    <div style={{ background: COLOR.bg, minHeight: "100%", fontFamily: SANS, color: COLOR.ink, padding: "28px 24px 60px" }}>
      <style>{`
        * { box-sizing: border-box; }
        input:focus, button:focus { outline: 2px solid ${COLOR.teal}; outline-offset: 2px; }
        th { cursor: pointer; user-select: none; }
        th:hover { color: ${COLOR.teal}; }
        tr.data-row:hover { background: ${COLOR.tealSoft}; }
        ::-webkit-scrollbar { height: 8px; width: 8px; }
        ::-webkit-scrollbar-thumb { background: ${COLOR.line}; border-radius: 4px; }
      `}</style>

      <div style={{ maxWidth: 1080, margin: "0 auto" }}>
        {/* Header */}
        <div style={{ marginBottom: 22 }}>
          <div style={{ fontSize: 11, letterSpacing: "0.14em", textTransform: "uppercase", color: COLOR.teal, fontWeight: 700, marginBottom: 6 }}>
            Sistem Rekap · Register Tes HIV
          </div>
          <h1 style={{ margin: 0, fontSize: 27, fontWeight: 700, letterSpacing: "-0.01em" }}>
            Olah Data Tes HIV per Layanan
          </h1>
          <p style={{ margin: "6px 0 0", color: COLOR.inkSoft, fontSize: 14, maxWidth: 620 }}>
            Unggah file Register Tes HIV (format sama seperti biasa) untuk langsung mendapatkan rekap
            jumlah tes per layanan dan capaian SPM per kelompok populasi — tanpa olah manual.
          </p>
        </div>

        {/* Upload zone */}
        <div
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={onDrop}
          style={{
            border: `1.5px dashed ${dragOver ? COLOR.teal : COLOR.line}`,
            background: dragOver ? COLOR.tealSoft : COLOR.surface,
            borderRadius: 12, padding: "22px 20px", display: "flex",
            alignItems: "center", justifyContent: "space-between", gap: 16, flexWrap: "wrap",
            transition: "all 0.15s ease",
          }}
        >
          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            <div style={{
              width: 40, height: 40, borderRadius: 9, background: COLOR.tealSoft,
              display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
            }}>
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke={COLOR.tealDark} strokeWidth="2">
                <path d="M12 3v12m0-12l-4 4m4-4l4 4" strokeLinecap="round" strokeLinejoin="round" />
                <path d="M4 17v2a2 2 0 002 2h12a2 2 0 002-2v-2" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </div>
            <div>
              <div style={{ fontWeight: 600, fontSize: 14 }}>
                {fileName ? fileName : "Seret file .xlsx ke sini, atau pilih file"}
              </div>
              <div style={{ fontSize: 12.5, color: COLOR.inkSoft, marginTop: 2 }}>
                {loading ? "Membaca data…" : fileName ? "Klik \"Ganti file\" untuk memakai data lain" : "Format: Register Tes HIV (.xlsx)"}
              </div>
            </div>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls"
              style={{ display: "none" }}
              onChange={(e) => handleFile(e.target.files[0])}
            />
            <button
              onClick={() => fileInputRef.current && fileInputRef.current.click()}
              style={{
                background: COLOR.teal, color: "#fff", border: "none", borderRadius: 8,
                padding: "10px 16px", fontSize: 13.5, fontWeight: 600, cursor: "pointer",
              }}
            >
              {fileName ? "Ganti file" : "Pilih file"}
            </button>
            {fileName && (
              <button
                onClick={reset}
                style={{
                  background: "transparent", color: COLOR.inkSoft, border: `1px solid ${COLOR.line}`,
                  borderRadius: 8, padding: "10px 14px", fontSize: 13.5, fontWeight: 600, cursor: "pointer",
                }}
              >
                Reset
              </button>
            )}
          </div>
        </div>

        {error && (
          <div style={{
            marginTop: 14, background: "#FBEAE6", border: `1px solid ${COLOR.coral}`, color: "#8A3624",
            borderRadius: 8, padding: "12px 14px", fontSize: 13.5,
          }}>
            {error}
          </div>
        )}

        {summary && parsed && (
          <>
            {/* Meta line */}
            {(parsed.meta["Provinsi"] || parsed.meta["Kabupaten/Kota"] || parsed.meta["Periode"]) && (
              <div style={{ marginTop: 20, fontSize: 12.5, color: COLOR.inkSoft }}>
                {[parsed.meta["Kabupaten/Kota"], parsed.meta["Provinsi"]].filter(Boolean).join(", ")}
                {parsed.meta["Periode"] ? ` · Periode ${parsed.meta["Periode"]}` : ""}
                {" · "}{fmt(summary.layananCount)} layanan
              </div>
            )}

            {/* Stat cards */}
            <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginTop: 14 }}>
              <StatCard eyebrow="Total Tes HIV" value={fmt(summary.total)} sub={`${fmt(summary.layananCount)} layanan`} />
              <StatCard eyebrow="Capaian SPM" value={fmt(summary.totalSpm)} sub={pct(summary.totalSpm, summary.total) + " dari total"} accent={COLOR.teal} />
              <StatCard eyebrow="Non-SPM" value={fmt(summary.totalNonSpm)} sub="Calon Pengantin" accent={COLOR.grey} />
              <StatCard eyebrow="ODHIV Terkonfirmasi" value={fmt(summary.totalOdhiv)} sub={pct(summary.totalOdhiv, summary.total) + " dari total"} accent={COLOR.coral} />
            </div>

            {/* Charts */}
            <div style={{ display: "flex", gap: 14, marginTop: 18, flexWrap: "wrap" }}>
              <div style={{ background: COLOR.surface, border: `1px solid ${COLOR.line}`, borderRadius: 10, padding: "16px 18px", flex: "2 1 420px", minWidth: 320 }}>
                <div style={{ fontSize: 12.5, fontWeight: 600, marginBottom: 10 }}>
                  Top 12 Layanan · Jumlah Tes
                </div>
                <ResponsiveContainer width="100%" height={280}>
                  <BarChart data={topChartData} layout="vertical" margin={{ left: 8, right: 16 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke={COLOR.line} horizontal={false} />
                    <XAxis type="number" tick={{ fontSize: 11, fill: COLOR.inkSoft }} />
                    <YAxis type="category" dataKey="name" width={150} tick={{ fontSize: 11, fill: COLOR.ink }} />
                    <Tooltip contentStyle={{ fontSize: 12, borderRadius: 6, border: `1px solid ${COLOR.line}` }} />
                    <Bar dataKey="Total" fill={COLOR.teal} radius={[0, 3, 3, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>

              <div style={{ background: COLOR.surface, border: `1px solid ${COLOR.line}`, borderRadius: 10, padding: "16px 18px", flex: "1 1 260px", minWidth: 260 }}>
                <div style={{ fontSize: 12.5, fontWeight: 600, marginBottom: 10 }}>
                  Capaian SPM vs Non-SPM
                </div>
                <ResponsiveContainer width="100%" height={220}>
                  <PieChart>
                    <Pie data={pieData} dataKey="value" nameKey="name" innerRadius={52} outerRadius={80} paddingAngle={2}>
                      <Cell fill={COLOR.teal} />
                      <Cell fill={COLOR.grey} />
                    </Pie>
                    <Tooltip contentStyle={{ fontSize: 12, borderRadius: 6, border: `1px solid ${COLOR.line}` }} />
                  </PieChart>
                </ResponsiveContainer>
                <div style={{ display: "flex", justifyContent: "center", gap: 16, fontSize: 12, marginTop: 4 }}>
                  <span><span style={{ display: "inline-block", width: 8, height: 8, borderRadius: 4, background: COLOR.teal, marginRight: 5 }} />SPM</span>
                  <span><span style={{ display: "inline-block", width: 8, height: 8, borderRadius: 4, background: COLOR.grey, marginRight: 5 }} />Non-SPM</span>
                </div>
              </div>
            </div>

            {/* Table */}
            <div style={{ background: COLOR.surface, border: `1px solid ${COLOR.line}`, borderRadius: 10, marginTop: 18, overflow: "hidden" }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 18px", borderBottom: `1px solid ${COLOR.line}`, gap: 12, flexWrap: "wrap" }}>
                <div style={{ fontSize: 13.5, fontWeight: 600 }}>Rekap per Layanan</div>
                <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
                  <input
                    value={query}
                    onChange={(e) => setQuery(e.target.value)}
                    placeholder="Cari nama layanan…"
                    style={{
                      border: `1px solid ${COLOR.line}`, borderRadius: 7, padding: "7px 10px",
                      fontSize: 12.5, width: 190, fontFamily: SANS,
                    }}
                  />
                  <button
                    onClick={handleDownload}
                    style={{
                      background: COLOR.tealDark, color: "#fff", border: "none", borderRadius: 7,
                      padding: "8px 14px", fontSize: 12.5, fontWeight: 600, cursor: "pointer",
                      display: "flex", alignItems: "center", gap: 6,
                    }}
                  >
                    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.2">
                      <path d="M12 3v12m0 0l-4-4m4 4l4-4M4 19h16" strokeLinecap="round" strokeLinejoin="round" />
                    </svg>
                    Unduh Rekap (.xlsx)
                  </button>
                </div>
              </div>

              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.8 }}>
                  <thead>
                    <tr style={{ background: COLOR.greySoft, textAlign: "left" }}>
                      <th style={{ padding: "9px 12px" }} onClick={() => toggleSort("upk")}>Layanan{arrow("upk")}</th>
                      <th style={{ padding: "9px 12px", textAlign: "right" }} onClick={() => toggleSort("total")}>Total Tes{arrow("total")}</th>
                      <th style={{ padding: "9px 12px", textAlign: "right" }} onClick={() => toggleSort("odhiv")}>ODHIV{arrow("odhiv")}</th>
                      <th style={{ padding: "9px 12px", textAlign: "right" }} onClick={() => toggleSort("belumTahu")}>Belum Tahu{arrow("belumTahu")}</th>
                      <th style={{ padding: "9px 12px", textAlign: "right" }} onClick={() => toggleSort("spm")}>Capaian SPM{arrow("spm")}</th>
                      <th style={{ padding: "9px 12px" }}>Cakupan SPM</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sortedRows.map((r) => (
                      <tr key={r.upk} className="data-row" style={{ borderTop: `1px solid ${COLOR.line}` }}>
                        <td style={{ padding: "8px 12px", maxWidth: 260 }}>{r.upk}</td>
                        <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: MONO }}>{fmt(r.total)}</td>
                        <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: MONO, color: r.odhiv ? COLOR.coral : COLOR.inkSoft }}>{fmt(r.odhiv)}</td>
                        <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: MONO, color: COLOR.inkSoft }}>{fmt(r.belumTahu)}</td>
                        <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: MONO }}>{fmt(r.spm)} <span style={{ color: COLOR.inkSoft }}>({pct(r.spm, r.total)})</span></td>
                        <td style={{ padding: "8px 12px" }}><CoverageRail spm={r.spm} total={r.total} /></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <div style={{ fontSize: 11.5, color: COLOR.inkSoft, marginTop: 12, lineHeight: 1.5 }}>
              Capaian SPM dihitung dari seluruh kelompok populasi di luar Calon Pengantin.
              File yang diunduh berisi 3 sheet: Data (mentah + kategori SPM), Rekap per Layanan, dan Rekap Kelompok Populasi.
            </div>
          </>
        )}

        {!summary && !error && !loading && (
          <div style={{ marginTop: 40, textAlign: "center", color: COLOR.inkSoft, fontSize: 13 }}>
            Belum ada data. Unggah file Register Tes HIV untuk mulai.
          </div>
        )}
      </div>
    </div>
  );
}
