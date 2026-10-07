import React, { useState, useCallback, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, Legend,
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

// ---------------------------------------------------------------------------
// Target matching (Indikator Target per Layanan / Dinkes)
// ---------------------------------------------------------------------------
const ORG_STOPWORDS = new Set([
  "PUSKESMAS", "RSUD", "RSIA", "RSK", "RS", "KLINIK", "PRATAMA", "KABUPATEN",
]);

function normalizeFacility(name) {
  return String(name || "")
    .toUpperCase()
    .replace(/\./g, "")
    .split(/[^A-Z0-9]+/)
    .filter((w) => w && !ORG_STOPWORDS.has(w))
    .join("");
}

function levenshtein(a, b) {
  const m = a.length, n = b.length;
  const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      dp[i][j] = a[i - 1] === b[j - 1] ? dp[i - 1][j - 1] : 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
    }
  }
  return dp[m][n];
}

function parseTargetSheet(aoa) {
  let headerRow = -1;
  for (let i = 0; i < Math.min(aoa.length, 40); i++) {
    const row = aoa[i] || [];
    if (row.some((c) => c != null && String(c).trim() === "PUSKESMAS")) {
      headerRow = i;
      break;
    }
  }
  if (headerRow === -1) return null;
  const mainHeader = aoa[headerRow] || [];
  const subHeader = aoa[headerRow + 1] || [];
  const colMap = {};
  mainHeader.forEach((c, idx) => {
    if (c != null && String(c).trim() !== "") colMap[String(c).trim().toUpperCase()] = idx;
  });
  const subLabels = ["BUMIL", "TB", "WARIA", "PENASUN", "WPS", "LSL", "WB"];
  subHeader.forEach((c, idx) => {
    if (c != null) {
      const label = String(c).trim().toUpperCase();
      if (subLabels.includes(label) && !(label in colMap)) colMap[label] = idx;
    }
  });
  if (!("PUSKESMAS" in colMap)) return null;

  const rows = [];
  for (let i = headerRow + 2; i < aoa.length; i++) {
    const row = aoa[i] || [];
    const nameRaw = row[colMap["PUSKESMAS"]];
    if (nameRaw == null) continue;
    const name = String(nameRaw).trim();
    if (!name || /^jumlah$/i.test(name)) continue;
    const num = (key) => {
      const v = colMap[key] != null ? row[colMap[key]] : null;
      const n = Number(v);
      return isNaN(n) ? 0 : n;
    };
    const bumil = num("BUMIL"), tb = num("TB"), waria = num("WARIA"),
      penasun = num("PENASUN"), wps = num("WPS"), lsl = num("LSL"), wb = num("WB");
    let total = num("TOTAL TARGET");
    if (!total) total = bumil + tb + waria + penasun + wps + lsl + wb;
    if (!total) continue;
    rows.push({ name, bumil, tb, waria, penasun, wps, lsl, wb, total });
  }
  return rows;
}

function parseTargetWorkbook(wb) {
  let best = null;
  let bestYear = -1;
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: true });
    const rows = parseTargetSheet(aoa);
    if (!rows || rows.length === 0) continue;
    let year = 0;
    for (let i = 0; i < Math.min(aoa.length, 6); i++) {
      const text = (aoa[i] || []).map((c) => (c == null ? "" : String(c))).join(" ");
      const m = text.match(/20\d{2}/);
      if (m) year = Math.max(year, Number(m[0]));
    }
    if (year >= bestYear) { bestYear = year; best = rows; }
  }
  if (!best) {
    throw new Error(
      "Format target tidak dikenali: tidak ditemukan kolom 'PUSKESMAS' pada file. Pastikan file target sesuai format standar (ada kolom PUSKESMAS dan TOTAL TARGET)."
    );
  }
  return best;
}

function facilityType(name) {
  const s = String(name || "").trim().toUpperCase();
  if (/^(RSUD|RSIA|RSK|RS)\b/.test(s)) return "RS";
  if (/^PUSKESMAS\b/.test(s)) return "PUSKESMAS";
  return "OTHER";
}

function matchTargetsToFacilities(targetRows, facilityNames) {
  const regKeys = facilityNames.map((n) => ({ raw: n, key: normalizeFacility(n), type: facilityType(n) }));
  const byFacility = new Map();
  const unmatched = [];
  for (const t of targetRows) {
    const tKey = normalizeFacility(t.name);
    if (!tKey) continue;
    const tType = facilityType(t.name);
    // Batasi kandidat ke tipe yang sama (RS hanya dicocokkan ke RS; selain itu ke Puskesmas/lainnya)
    // supaya nama pendek seperti "KARAWANG" tidak salah nyantol ke "RSUD Karawang".
    const candidates = tType === "RS" ? regKeys.filter((r) => r.type === "RS") : regKeys.filter((r) => r.type !== "RS");
    let match = candidates.find((r) => r.key === tKey);
    if (!match) match = candidates.find((r) => r.key === tKey + "KARAWANG");
    if (!match) {
      const contains = candidates.filter((r) => tKey.length >= 4 && r.key.includes(tKey));
      if (contains.length === 1) match = contains[0];
    }
    if (!match) {
      let bestR = null, bestDist = Infinity, secondDist = Infinity;
      for (const r of candidates) {
        const d = levenshtein(tKey, r.key);
        if (d < bestDist) { secondDist = bestDist; bestDist = d; bestR = r; }
        else if (d < secondDist) secondDist = d;
      }
      if (bestR && bestDist <= 4 && bestDist < secondDist) match = bestR;
    }
    if (match) {
      const cur = byFacility.get(match.raw) || { bumil: 0, tb: 0, waria: 0, penasun: 0, wps: 0, lsl: 0, wb: 0, total: 0 };
      cur.bumil += t.bumil; cur.tb += t.tb; cur.waria += t.waria; cur.penasun += t.penasun;
      cur.wps += t.wps; cur.lsl += t.lsl; cur.wb += t.wb; cur.total += t.total;
      byFacility.set(match.raw, cur);
    } else {
      unmatched.push(t.name);
    }
  }
  return { byFacility, unmatched };
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

// ---------------------------------------------------------------------------
// Laporan Capaian Bulanan (berdasarkan kolom "Tanggal Kunjungan")
// ---------------------------------------------------------------------------
const BULAN_ID = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];

function parseTanggalKunjungan(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === "number" && isFinite(v)) {
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    return isNaN(d.getTime()) ? null : d;
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (m) {
    const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
    return isNaN(d.getTime()) ? null : d;
  }
  m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) {
    const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    return isNaN(d.getTime()) ? null : d;
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function monthKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function monthLabel(key) {
  const [y, m] = key.split("-");
  return `${BULAN_ID[Number(m) - 1]} ${y}`;
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

  const tglIdx = colMap["Tanggal Kunjungan"]; // opsional — boleh tidak ada di beberapa format file
  const artIdx = colMap["Tanggal Mulai ART"]; // opsional — terisi berarti ODHIV sudah mulai ART

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
      tanggalKunjungan: tglIdx != null ? row[tglIdx] : null,
      tanggalMulaiArt: artIdx != null ? row[artIdx] : null,
    });
  }
  if (rows.length === 0) {
    throw new Error("Header ditemukan tapi tidak ada baris data di bawahnya.");
  }
  return { headerRow, colMap, headerRowRaw: aoa[headerRow], rows, meta: findMeta(aoa) };
}

function isMulaiArt(v) {
  return v != null && String(v).trim() !== "";
}

function summarize(rows) {
  const byUpk = new Map();
  const byMonth = new Map();
  const byMonthUpk = new Map(); // key: `${monthKey}|${upk}`
  let totalOdhiv = 0, totalBelumTahu = 0, totalSpm = 0, totalMulaiArt = 0;
  let tglTerbaca = 0;

  for (const r of rows) {
    if (!byUpk.has(r.upk)) {
      byUpk.set(r.upk, { upk: r.upk, total: 0, odhiv: 0, belumTahu: 0, bukanOdhiv: 0, spm: 0, mulaiArt: 0 });
    }
    const b = byUpk.get(r.upk);
    b.total += 1;
    if (r.statusOdhiv === "ODHIV") { b.odhiv += 1; totalOdhiv += 1; }
    else if (r.statusOdhiv === "Belum Tahu") { b.belumTahu += 1; totalBelumTahu += 1; }
    else if (r.statusOdhiv === "Bukan ODHIV") { b.bukanOdhiv += 1; }
    const spm = isSPM(r.kelompokPopulasi);
    if (spm) { b.spm += 1; totalSpm += 1; }
    const mulaiArt = isMulaiArt(r.tanggalMulaiArt);
    if (mulaiArt) { b.mulaiArt += 1; totalMulaiArt += 1; }

    const tgl = parseTanggalKunjungan(r.tanggalKunjungan);
    if (tgl) {
      tglTerbaca += 1;
      const key = monthKey(tgl);

      if (!byMonth.has(key)) byMonth.set(key, { key, label: monthLabel(key), total: 0, spm: 0, nonSpm: 0, odhiv: 0, mulaiArt: 0 });
      const m = byMonth.get(key);
      m.total += 1;
      if (spm) m.spm += 1; else m.nonSpm += 1;
      if (r.statusOdhiv === "ODHIV") m.odhiv += 1;
      if (mulaiArt) m.mulaiArt += 1;

      const muKey = `${key}|${r.upk}`;
      if (!byMonthUpk.has(muKey)) {
        byMonthUpk.set(muKey, { month: key, monthLabel: monthLabel(key), upk: r.upk, total: 0, spm: 0, nonSpm: 0, odhiv: 0, mulaiArt: 0 });
      }
      const mu = byMonthUpk.get(muKey);
      mu.total += 1;
      if (spm) mu.spm += 1; else mu.nonSpm += 1;
      if (r.statusOdhiv === "ODHIV") mu.odhiv += 1;
      if (mulaiArt) mu.mulaiArt += 1;
    }
  }

  const perLayanan = Array.from(byUpk.values()).sort((a, b) => b.total - a.total);
  const perBulan = Array.from(byMonth.values()).sort((a, b) => a.key.localeCompare(b.key));
  const perBulanLayanan = Array.from(byMonthUpk.values());
  const total = rows.length;
  return {
    total, totalOdhiv, totalBelumTahu, totalSpm, totalMulaiArt,
    totalNonSpm: total - totalSpm,
    layananCount: perLayanan.length,
    perLayanan,
    perBulan,
    perBulanLayanan,
    adaTanggalKunjungan: tglTerbaca > 0,
  };
}

function getAtomicCategoryCounts(rows) {
  // Rincian per kelompok populasi (bukan hanya SPM/Non-SPM), khusus untuk file unduhan.
  const set = new Set();
  rows.forEach((r) => {
    if (r.kelompokPopulasi) {
      r.kelompokPopulasi.split(",").forEach((p) => {
        const t = p.trim();
        if (t) set.add(t);
      });
    }
  });
  const cats = Array.from(set).sort((a, b) => {
    const countA = rows.filter((r) => r.kelompokPopulasi.includes(a)).length;
    const countB = rows.filter((r) => r.kelompokPopulasi.includes(b)).length;
    return countB - countA;
  });
  return cats.map((cat) => ({
    cat,
    count: rows.filter((r) => r.kelompokPopulasi.includes(cat)).length,
  }));
}

function buildWorkbook(parsed, summary, targetInfo) {
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

  // Sheet 4: Detail per Kelompok Populasi (rincian, khusus file unduhan)
  const atomicCounts = getAtomicCategoryCounts(parsed.rows);
  const detailAoa = [["No", "Kelompok Populasi", "Jumlah Tes", "% dari Total Baris Data"]];
  atomicCounts.forEach((c, i) => {
    detailAoa.push([i + 1, c.cat, c.count, summary.total ? c.count / summary.total : 0]);
  });
  detailAoa.push([null, "TOTAL BARIS DATA", summary.total, null]);
  const wsDetail = XLSX.utils.aoa_to_sheet(detailAoa);
  XLSX.utils.book_append_sheet(wb, wsDetail, "Detail Kelompok Populasi");

  // Sheet 5: Capaian SPM vs Target (hanya jika file target diunggah)
  if (targetInfo && targetInfo.targetRows && targetInfo.targetRows.length) {
    const { byFacility, totalTargetAll } = targetInfo;
    const targetAoa = [["No", "Nama UPK / Layanan", "Capaian SPM", "Target SPM", "% Capaian SPM"]];
    summary.perLayanan.forEach((r, i) => {
      const t = byFacility.get(r.upk);
      targetAoa.push([i + 1, r.upk, r.spm, t ? t.total : null, t && t.total ? r.spm / t.total : null]);
    });
    targetAoa.push([null, "TOTAL DINKES (KABUPATEN)", summary.totalSpm, totalTargetAll, totalTargetAll ? summary.totalSpm / totalTargetAll : null]);
    const wsTarget = XLSX.utils.aoa_to_sheet(targetAoa);
    XLSX.utils.book_append_sheet(wb, wsTarget, "Capaian SPM vs Target");
  }

  // Sheet 6: Capaian Bulanan — ringkasan Dinkes (hanya jika kolom Tanggal Kunjungan terbaca)
  if (summary.adaTanggalKunjungan && summary.perBulan && summary.perBulan.length) {
    const targetBulananDinkes = targetInfo ? targetInfo.totalTargetAll / 12 : null;
    const bulanHeader = ["Bulan", "Total Tes", "Capaian SPM", "Non-SPM", "ODHIV", "Mulai ART", "% SPM"];
    if (targetBulananDinkes) bulanHeader.push("Target Bulanan", "% Capaian Bulanan");
    const bulanAoa = [bulanHeader];
    summary.perBulan.forEach((m) => {
      const row = [m.label, m.total, m.spm, m.nonSpm, m.odhiv, m.mulaiArt, m.total ? m.spm / m.total : 0];
      if (targetBulananDinkes) row.push(targetBulananDinkes, m.spm / targetBulananDinkes);
      bulanAoa.push(row);
    });
    const wsBulan = XLSX.utils.aoa_to_sheet(bulanAoa);
    XLSX.utils.book_append_sheet(wb, wsBulan, "Capaian Bulanan");

    // Sheet 7: Capaian Bulanan per Layanan (rincian tiap layanan, tiap bulan)
    const bulanLayananHeader = ["Bulan", "Nama UPK / Layanan", "Total Tes", "Capaian SPM", "Non-SPM", "ODHIV", "Mulai ART"];
    if (targetInfo) bulanLayananHeader.push("Target Bulanan", "% Capaian");
    const bulanLayananAoa = [bulanLayananHeader];
    const sortedMonthly = [...summary.perBulanLayanan].sort((a, b) => a.month.localeCompare(b.month) || b.total - a.total);
    sortedMonthly.forEach((r) => {
      const t = targetInfo ? targetInfo.byFacility.get(r.upk) : null;
      const targetBulananLayanan = t ? t.total / 12 : null;
      const row = [r.monthLabel, r.upk, r.total, r.spm, r.nonSpm, r.odhiv, r.mulaiArt];
      if (targetInfo) row.push(targetBulananLayanan, targetBulananLayanan ? r.spm / targetBulananLayanan : null);
      bulanLayananAoa.push(row);
    });
    const wsBulanLayanan = XLSX.utils.aoa_to_sheet(bulanLayananAoa);
    XLSX.utils.book_append_sheet(wb, wsBulanLayanan, "Capaian Bulanan per Layanan");
  }

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

function CapaianBar({ capaian, height = 8, width = 90 }) {
  if (capaian == null) {
    return <div style={{ width, height, borderRadius: height / 2, background: COLOR.greySoft }} />;
  }
  const color = capaian >= 1 ? COLOR.teal : capaian >= 0.75 ? COLOR.amber : COLOR.coral;
  const fillPct = Math.min(capaian, 1) * 100;
  return (
    <div style={{ width, height, borderRadius: height / 2, overflow: "hidden", background: COLOR.greySoft, position: "relative" }}>
      <div style={{ width: `${fillPct}%`, height: "100%", background: color }} />
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

  const [selectedMonth, setSelectedMonth] = useState(null);

  const [targetRows, setTargetRows] = useState(null);
  const [targetFileName, setTargetFileName] = useState("");
  const [targetError, setTargetError] = useState("");
  const [targetLoading, setTargetLoading] = useState(false);
  const targetInputRef = useRef(null);

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

  const handleTargetFile = useCallback((file) => {
    if (!file) return;
    setTargetError("");
    setTargetLoading(true);
    setTargetFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array", cellDates: true });
        const rows = parseTargetWorkbook(wb);
        setTargetRows(rows);
      } catch (err) {
        setTargetError(err.message || "Gagal membaca file target.");
        setTargetRows(null);
      } finally {
        setTargetLoading(false);
      }
    };
    reader.onerror = () => {
      setTargetError("Gagal membaca file target.");
      setTargetLoading(false);
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const resetTarget = () => {
    setTargetRows(null);
    setTargetFileName("");
    setTargetError("");
    if (targetInputRef.current) targetInputRef.current.value = "";
  };

  const targetMatch = useMemo(() => {
    if (!targetRows || !summary) return null;
    const { byFacility, unmatched } = matchTargetsToFacilities(targetRows, summary.perLayanan.map((r) => r.upk));
    const totalTargetAll = targetRows.reduce((s, t) => s + t.total, 0);
    return { byFacility, unmatched, totalTargetAll };
  }, [targetRows, summary]);

  const enrichedRows = useMemo(() => {
    if (!summary) return [];
    return summary.perLayanan.map((r) => {
      const t = targetMatch ? targetMatch.byFacility.get(r.upk) : null;
      const target = t ? t.total : null;
      const capaian = target ? r.spm / target : null;
      return { ...r, target, capaian };
    });
  }, [summary, targetMatch]);

  const sortedRows = useMemo(() => {
    if (!summary) return [];
    let rows = enrichedRows;
    if (query.trim()) {
      const q = query.trim().toLowerCase();
      rows = rows.filter((r) => r.upk.toLowerCase().includes(q));
    }
    const dir = sortDir === "asc" ? 1 : -1;
    return [...rows].sort((a, b) => {
      if (sortKey === "upk") return a.upk.localeCompare(b.upk) * dir;
      const av = a[sortKey], bv = b[sortKey];
      const an = av == null ? -1 : av, bn = bv == null ? -1 : bv;
      return (an - bn) * dir;
    });
  }, [summary, enrichedRows, sortKey, sortDir, query]);

  const topChartData = useMemo(() => {
    if (!summary) return [];
    return summary.perLayanan.slice(0, 12).map((r) => ({
      name: r.upk.length > 22 ? r.upk.slice(0, 20) + "…" : r.upk,
      Total: r.total,
    })).reverse();
  }, [summary]);

  const monthlyChartData = useMemo(() => {
    if (!summary || !summary.perBulan) return [];
    return summary.perBulan.map((m) => ({ label: m.label, SPM: m.spm, "Non-SPM": m.nonSpm }));
  }, [summary]);

  const selectedMonthEffective = useMemo(() => {
    if (!summary || !summary.perBulan.length) return null;
    if (selectedMonth && summary.perBulan.some((m) => m.key === selectedMonth)) return selectedMonth;
    return summary.perBulan[summary.perBulan.length - 1].key;
  }, [summary, selectedMonth]);

  const selectedMonthSummary = useMemo(() => {
    if (!summary || !selectedMonthEffective) return null;
    return summary.perBulan.find((m) => m.key === selectedMonthEffective) || null;
  }, [summary, selectedMonthEffective]);

  const monthlyLayananRows = useMemo(() => {
    if (!summary || !selectedMonthEffective) return [];
    return summary.perBulanLayanan
      .filter((r) => r.month === selectedMonthEffective)
      .sort((a, b) => b.total - a.total);
  }, [summary, selectedMonthEffective]);

  const pieData = useMemo(() => {
    if (!summary) return [];
    return [
      { name: "Capaian SPM", value: summary.totalSpm },
      { name: "Non-SPM (Catin/Pop. Umum)", value: summary.totalNonSpm },
    ];
  }, [summary]);

  const handleDownload = () => {
    if (!parsed || !summary) return;
    const wb = buildWorkbook(parsed, summary, targetMatch ? { ...targetMatch, targetRows } : null);
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

            {/* Target upload (opsional) */}
            <div style={{
              marginTop: 14, border: `1px solid ${COLOR.line}`, borderRadius: 10,
              background: COLOR.surface, padding: "12px 16px", display: "flex",
              alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap",
            }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <div style={{
                  width: 30, height: 30, borderRadius: 7, background: COLOR.greySoft,
                  display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
                }}>
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={COLOR.inkSoft} strokeWidth="2">
                    <circle cx="12" cy="12" r="9" />
                    <circle cx="12" cy="12" r="4.5" />
                    <circle cx="12" cy="12" r="0.7" fill={COLOR.inkSoft} />
                  </svg>
                </div>
                <div>
                  <div style={{ fontWeight: 600, fontSize: 13 }}>
                    {targetFileName ? targetFileName : "Indikator Target (opsional)"}
                  </div>
                  <div style={{ fontSize: 12, color: COLOR.inkSoft, marginTop: 1 }}>
                    {targetLoading
                      ? "Membaca target…"
                      : targetMatch
                      ? `${fmt(targetMatch.byFacility.size)} layanan cocok${targetMatch.unmatched.length ? `, ${fmt(targetMatch.unmatched.length)} tidak cocok` : ""}`
                      : "Unggah file target SPM Dinkes untuk melihat capaian per layanan"}
                  </div>
                </div>
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                <input
                  ref={targetInputRef}
                  type="file"
                  accept=".xlsx,.xls"
                  style={{ display: "none" }}
                  onChange={(e) => handleTargetFile(e.target.files[0])}
                />
                <button
                  onClick={() => targetInputRef.current && targetInputRef.current.click()}
                  style={{
                    background: targetFileName ? "transparent" : COLOR.teal,
                    color: targetFileName ? COLOR.inkSoft : "#fff",
                    border: targetFileName ? `1px solid ${COLOR.line}` : "none",
                    borderRadius: 7, padding: "7px 12px", fontSize: 12.5, fontWeight: 600, cursor: "pointer",
                  }}
                >
                  {targetFileName ? "Ganti" : "Unggah Target"}
                </button>
                {targetFileName && (
                  <button
                    onClick={resetTarget}
                    style={{
                      background: "transparent", color: COLOR.inkSoft, border: `1px solid ${COLOR.line}`,
                      borderRadius: 7, padding: "7px 12px", fontSize: 12.5, fontWeight: 600, cursor: "pointer",
                    }}
                  >
                    Hapus
                  </button>
                )}
              </div>
            </div>

            {targetError && (
              <div style={{
                marginTop: 10, background: "#FBEAE6", border: `1px solid ${COLOR.coral}`, color: "#8A3624",
                borderRadius: 8, padding: "10px 14px", fontSize: 13,
              }}>
                {targetError}
              </div>
            )}

            {targetMatch && targetMatch.unmatched.length > 0 && (
              <div style={{
                marginTop: 10, background: "#FBF3E4", border: `1px solid ${COLOR.amber}`, color: "#7A5A17",
                borderRadius: 8, padding: "10px 14px", fontSize: 12.5, lineHeight: 1.5,
              }}>
                {fmt(targetMatch.unmatched.length)} baris target tidak dapat dicocokkan otomatis ke nama layanan di data: {targetMatch.unmatched.slice(0, 8).join(", ")}{targetMatch.unmatched.length > 8 ? `, +${targetMatch.unmatched.length - 8} lainnya` : ""}.
              </div>
            )}

            {/* Stat cards */}
            <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginTop: 14 }}>
              <StatCard eyebrow="Total Tes HIV" value={fmt(summary.total)} sub={`${fmt(summary.layananCount)} layanan`} />
              <StatCard eyebrow="Capaian SPM" value={fmt(summary.totalSpm)} sub={pct(summary.totalSpm, summary.total) + " dari total"} accent={COLOR.teal} />
              <StatCard eyebrow="Non-SPM" value={fmt(summary.totalNonSpm)} sub="Calon Pengantin" accent={COLOR.grey} />
              <StatCard eyebrow="ODHIV Terkonfirmasi" value={fmt(summary.totalOdhiv)} sub={pct(summary.totalOdhiv, summary.total) + " dari total"} accent={COLOR.coral} />
              {targetMatch && (
                <StatCard
                  eyebrow="Capaian Dinkes (Kabupaten) — SPM"
                  value={pct(summary.totalSpm, targetMatch.totalTargetAll)}
                  sub={`${fmt(summary.totalSpm)} / ${fmt(Math.round(targetMatch.totalTargetAll))} target SPM`}
                  accent={summary.totalSpm / targetMatch.totalTargetAll >= 1 ? COLOR.teal : summary.totalSpm / targetMatch.totalTargetAll >= 0.75 ? COLOR.amber : COLOR.coral}
                />
              )}
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
                      {targetMatch && (
                        <>
                          <th style={{ padding: "9px 12px", textAlign: "right" }} onClick={() => toggleSort("target")}>Target SPM{arrow("target")}</th>
                          <th style={{ padding: "9px 12px", textAlign: "right" }} onClick={() => toggleSort("capaian")}>% Capaian SPM{arrow("capaian")}</th>
                          <th style={{ padding: "9px 12px" }}>Capaian SPM vs Target</th>
                        </>
                      )}
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
                        {targetMatch && (
                          <>
                            <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: MONO, color: r.target != null ? COLOR.ink : COLOR.inkSoft }}>
                              {r.target != null ? fmt(Math.round(r.target)) : "–"}
                            </td>
                            <td style={{
                              padding: "8px 12px", textAlign: "right", fontFamily: MONO, fontWeight: 600,
                              color: r.capaian == null ? COLOR.inkSoft : r.capaian >= 1 ? COLOR.teal : r.capaian >= 0.75 ? COLOR.amber : COLOR.coral,
                            }}>
                              {r.capaian != null ? pct(r.spm, r.target) : "–"}
                            </td>
                            <td style={{ padding: "8px 12px" }}><CapaianBar capaian={r.capaian} /></td>
                          </>
                        )}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Laporan Capaian Bulanan per Layanan (dari kolom Tanggal Kunjungan) */}
            {summary.adaTanggalKunjungan && summary.perBulan.length > 0 && (
              <div style={{ background: COLOR.surface, border: `1px solid ${COLOR.line}`, borderRadius: 10, marginTop: 18, padding: "16px 18px" }}>
                <div style={{ fontSize: 13.5, fontWeight: 600, marginBottom: 2 }}>Laporan Capaian Bulanan per Layanan</div>
                <div style={{ fontSize: 12, color: COLOR.inkSoft, marginBottom: 12 }}>
                  Berdasarkan kolom Tanggal Kunjungan. "Mulai ART" dihitung dari kolom Tanggal Mulai ART — terisi berarti ODHIV tsb sudah mulai ART.
                  {targetMatch ? " Target bulanan per layanan = target SPM tahunan layanan ÷ 12." : ""}
                </div>

                <ResponsiveContainer width="100%" height={220}>
                  <BarChart data={monthlyChartData} margin={{ left: 0, right: 8, top: 4 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke={COLOR.line} vertical={false} />
                    <XAxis dataKey="label" tick={{ fontSize: 11, fill: COLOR.inkSoft }} />
                    <YAxis tick={{ fontSize: 11, fill: COLOR.inkSoft }} />
                    <Tooltip contentStyle={{ fontSize: 12, borderRadius: 6, border: `1px solid ${COLOR.line}` }} />
                    <Legend wrapperStyle={{ fontSize: 12 }} />
                    <Bar dataKey="SPM" stackId="a" fill={COLOR.teal} radius={[0, 0, 0, 0]} />
                    <Bar dataKey="Non-SPM" stackId="a" fill={COLOR.grey} radius={[3, 3, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>

                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginTop: 16, marginBottom: 10 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <span style={{ fontSize: 12.5, fontWeight: 600, color: COLOR.inkSoft }}>Pilih bulan:</span>
                    <select
                      value={selectedMonthEffective || ""}
                      onChange={(e) => setSelectedMonth(e.target.value)}
                      style={{ border: `1px solid ${COLOR.line}`, borderRadius: 7, padding: "7px 10px", fontSize: 12.5, fontFamily: SANS, background: COLOR.surface }}
                    >
                      {summary.perBulan.map((m) => (
                        <option key={m.key} value={m.key}>{m.label}</option>
                      ))}
                    </select>
                  </div>
                  {selectedMonthSummary && (
                    <div style={{ fontSize: 12, color: COLOR.inkSoft }}>
                      {selectedMonthSummary.label}: {fmt(selectedMonthSummary.total)} tes · {fmt(selectedMonthSummary.spm)} SPM · {fmt(selectedMonthSummary.odhiv)} ODHIV · {fmt(selectedMonthSummary.mulaiArt)} mulai ART
                    </div>
                  )}
                </div>

                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.8 }}>
                    <thead>
                      <tr style={{ background: COLOR.greySoft, textAlign: "left" }}>
                        <th style={{ padding: "9px 12px" }}>Layanan</th>
                        <th style={{ padding: "9px 12px", textAlign: "right" }}>Total Tes</th>
                        <th style={{ padding: "9px 12px", textAlign: "right" }}>Capaian SPM</th>
                        <th style={{ padding: "9px 12px", textAlign: "right" }}>Non-SPM</th>
                        <th style={{ padding: "9px 12px", textAlign: "right" }}>ODHIV</th>
                        <th style={{ padding: "9px 12px", textAlign: "right" }}>Mulai ART</th>
                        {targetMatch && (
                          <>
                            <th style={{ padding: "9px 12px", textAlign: "right" }}>Target Bulanan</th>
                            <th style={{ padding: "9px 12px", textAlign: "right" }}>% Capaian</th>
                          </>
                        )}
                      </tr>
                    </thead>
                    <tbody>
                      {monthlyLayananRows.map((r) => {
                        const t = targetMatch ? targetMatch.byFacility.get(r.upk) : null;
                        const targetBulanan = t ? t.total / 12 : null;
                        const capaianBulanan = targetBulanan ? r.spm / targetBulanan : null;
                        return (
                          <tr key={r.upk} className="data-row" style={{ borderTop: `1px solid ${COLOR.line}` }}>
                            <td style={{ padding: "8px 12px", maxWidth: 240 }}>{r.upk}</td>
                            <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: MONO }}>{fmt(r.total)}</td>
                            <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: MONO }}>{fmt(r.spm)}</td>
                            <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: MONO, color: COLOR.inkSoft }}>{fmt(r.nonSpm)}</td>
                            <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: MONO, color: r.odhiv ? COLOR.coral : COLOR.inkSoft }}>{fmt(r.odhiv)}</td>
                            <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: MONO, color: r.mulaiArt ? COLOR.teal : COLOR.inkSoft }}>{fmt(r.mulaiArt)}</td>
                            {targetMatch && (
                              <>
                                <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: MONO, color: COLOR.inkSoft }}>
                                  {targetBulanan != null ? fmt(Math.round(targetBulanan)) : "–"}
                                </td>
                                <td style={{
                                  padding: "8px 12px", textAlign: "right", fontFamily: MONO, fontWeight: 600,
                                  color: capaianBulanan == null ? COLOR.inkSoft : capaianBulanan >= 1 ? COLOR.teal : capaianBulanan >= 0.75 ? COLOR.amber : COLOR.coral,
                                }}>
                                  {capaianBulanan != null ? pct(r.spm, targetBulanan) : "–"}
                                </td>
                              </>
                            )}
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            <div style={{ fontSize: 11.5, color: COLOR.inkSoft, marginTop: 12, lineHeight: 1.5 }}>
              Capaian SPM dihitung dari seluruh kelompok populasi di luar Calon Pengantin.
              {" "}File yang diunduh berisi sheet: Data (mentah + kategori SPM), Rekap per Layanan, Rekap Kelompok Populasi, Detail Kelompok Populasi{targetMatch ? ", Capaian SPM vs Target" : ""}{summary.adaTanggalKunjungan ? ", Capaian Bulanan, dan Capaian Bulanan per Layanan" : ""}.
              {!targetMatch ? " Unggah file target di atas untuk menambahkan sheet Capaian SPM vs Target." : ""}
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
