import { useCallback, useEffect, useMemo, useState } from "react";

/* =====================================================================
   PENGATURAN
   ===================================================================== */
// Tempel URL Web App dari Google Apps Script (berakhiran /exec).
// Dikosongkan = MODE LOKAL (memakai data awal dari Jubah.xlsx, disimpan di browser).
const API_URL =
  "https://script.google.com/macros/s/AKfycbz_iBpTgH8J_z2WD_njQwu8f0LCt8yo3NJgp-TEuYAUVS13kQzwLArBexOr7tvXqMfRXA/exec";
// Harus sama dengan CONFIG.TOKEN di Apps Script (kosongkan jika tidak dipakai).
const API_TOKEN = "123";
const STORAGE_KEY = "stok-barang-v1";

/* =====================================================================
   TIPE
   ===================================================================== */
type SheetKey = "jubah" | "campur";
type Status = "habis" | "menipis" | "tersedia";
type ColKey =
  | "merek"
  | "kode"
  | "model"
  | "warna"
  | "ukuran"
  | "stokLama"
  | "masuk"
  | "terjual"
  | "stokAkhir"
  | "hargaBeli"
  | "hargaJual"
  | "aset"
  | "untung";
type FormKey =
  | "merek"
  | "kode"
  | "model"
  | "warna"
  | "ukuran"
  | "stokLama"
  | "masuk"
  | "terjual"
  | "hargaBeli"
  | "hargaJual";

interface Item {
  row: number; // nomor baris di Google Sheets
  merek: string;
  kode: string;
  model: string;
  warna: string;
  ukuran: string;
  stokLama: number;
  masuk: number;
  terjual: number;
  hargaBeli: number;
  hargaJual: number;
}
type Data = Record<SheetKey, Item[]>;
type SeedRow = [
  string,
  string,
  string,
  string,
  string,
  number,
  number,
  number,
  number,
  number
];
interface Guide {
  merek: string;
  head: string[];
  rows: string[][];
}

const SHEETS: Record<SheetKey, { label: string; sheetName: string }> = {
  jubah: { label: "Stok Jubah", sheetName: "Stock Jubah" },
  campur: { label: "Stok Barang Campur", sheetName: "Stock Barang Campur" },
};

/* =====================================================================
   DATA AWAL: sengaja dikosongkan. Data dimuat langsung dari Google Sheets
   (sheet "Stock Jubah" dan "Stock Barang Campur") lewat API_URL di atas.
   Urutan kolom jika ingin mengisi manual: merek, kode, model, warna, ukuran,
   stokLama, masuk, terjual, hargaBeli, hargaJual
   ===================================================================== */
const SEED_JUBAH: SeedRow[] = [];

const SEED_CAMPUR: SeedRow[] = [];

// Panduan ukuran jubah (opsional). Kosong = bagian panduan disembunyikan.
const GUIDES: Guide[] = [];

const toItems = (rows: SeedRow[]): Item[] =>
  rows.map((r, i) => ({
    row: i + 2,
    merek: r[0],
    kode: r[1],
    model: r[2],
    warna: r[3],
    ukuran: r[4],
    stokLama: r[5],
    masuk: r[6],
    terjual: r[7],
    hargaBeli: r[8],
    hargaJual: r[9],
  }));
const seedData = (): Data => ({
  jubah: toItems(SEED_JUBAH),
  campur: toItems(SEED_CAMPUR),
});

/* =====================================================================
   HELPER
   ===================================================================== */
const stokAkhir = (i: Item) => i.stokLama + i.masuk - i.terjual;
const rp = (n: number) => "Rp " + Math.round(n).toLocaleString("id-ID");
const num = (n: number) => n.toLocaleString("id-ID");
const errMsg = (e: unknown) => (e instanceof Error ? e.message : String(e));
const uniq = (a: string[]) =>
  Array.from(new Set(a.filter(Boolean))).sort((x, y) =>
    x.localeCompare(y, "id", { numeric: true })
  );
const getStatus = (i: Item, batas: number): Status => {
  const s = stokAkhir(i);
  return s <= 0 ? "habis" : s <= batas ? "menipis" : "tersedia";
};
const keyOf = (i: Item) => ({
  merek: i.merek,
  kode: i.kode,
  model: i.model,
  warna: i.warna,
  ukuran: i.ukuran,
});
const payload = (i: Item) => ({
  merek: i.merek,
  kode: i.kode,
  model: i.model,
  warna: i.warna,
  ukuran: i.ukuran,
  stokLama: i.stokLama,
  masuk: i.masuk,
  terjual: i.terjual,
  hargaBeli: i.hargaBeli,
  hargaJual: i.hargaJual,
});

const COLS: {
  key: ColKey;
  label: string;
  num: boolean;
  money?: boolean;
  get: (i: Item) => string | number;
}[] = [
  { key: "merek", label: "Merek", num: false, get: (i) => i.merek },
  { key: "kode", label: "Kode", num: false, get: (i) => i.kode },
  { key: "model", label: "Model", num: false, get: (i) => i.model },
  { key: "warna", label: "Warna", num: false, get: (i) => i.warna },
  { key: "ukuran", label: "Ukuran", num: false, get: (i) => i.ukuran },
  { key: "stokLama", label: "Stok Lama", num: true, get: (i) => i.stokLama },
  { key: "masuk", label: "Masuk", num: true, get: (i) => i.masuk },
  { key: "terjual", label: "Terjual", num: true, get: (i) => i.terjual },
  { key: "stokAkhir", label: "Stok Akhir", num: true, get: stokAkhir },
  {
    key: "hargaBeli",
    label: "Harga Beli",
    num: true,
    money: true,
    get: (i) => i.hargaBeli,
  },
  {
    key: "hargaJual",
    label: "Harga Jual",
    num: true,
    money: true,
    get: (i) => i.hargaJual,
  },
  {
    key: "aset",
    label: "Total Aset",
    num: true,
    money: true,
    get: (i) => Math.max(stokAkhir(i), 0) * i.hargaBeli,
  },
  {
    key: "untung",
    label: "Untung/Pcs",
    num: true,
    money: true,
    get: (i) => i.hargaJual - i.hargaBeli,
  },
];

const FORM_FIELDS: {
  key: FormKey;
  label: string;
  num: boolean;
  list?: string;
}[] = [
  { key: "merek", label: "Merek *", num: false, list: "dl-merek" },
  { key: "kode", label: "Kode", num: false },
  { key: "model", label: "Model", num: false, list: "dl-model" },
  { key: "warna", label: "Warna", num: false, list: "dl-warna" },
  { key: "ukuran", label: "Ukuran", num: false, list: "dl-ukuran" },
  { key: "stokLama", label: "Stok Lama", num: true },
  { key: "masuk", label: "Barang Masuk", num: true },
  { key: "terjual", label: "Terjual", num: true },
  { key: "hargaBeli", label: "Harga Beli (Rp)", num: true },
  { key: "hargaJual", label: "Harga Jual (Rp)", num: true },
];

/* =====================================================================
   API APPS SCRIPT
   ===================================================================== */
async function apiList(): Promise<Data> {
  const url = `${API_URL}?action=list${
    API_TOKEN ? `&token=${encodeURIComponent(API_TOKEN)}` : ""
  }`;
  const res = await fetch(url);
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || "Gagal memuat data");
  return { jubah: json.jubah as Item[], campur: json.campur as Item[] };
}

async function apiPost(
  body: Record<string, unknown>
): Promise<{ item?: Item }> {
  // text/plain agar browser tidak mengirim preflight CORS (Apps Script tidak mendukung OPTIONS)
  const res = await fetch(API_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ ...body, token: API_TOKEN }),
  });
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || "Gagal menyimpan");
  return json;
}

/* =====================================================================
   KOMPONEN
   ===================================================================== */
function Card(props: {
  label: string;
  value: string;
  note?: string;
  warn?: boolean;
}) {
  return (
    <div className={"card" + (props.warn ? " warn" : "")}>
      <div className="cl">{props.label}</div>
      <div className="cv">{props.value}</div>
      {props.note && <div className="cn">{props.note}</div>}
    </div>
  );
}

export default function App() {
  const [data, setData] = useState<Data>(() => {
    if (API_URL) return { jubah: [], campur: [] };
    try {
      const s = localStorage.getItem(STORAGE_KEY);
      if (s) return JSON.parse(s) as Data;
    } catch {
      /* abaikan, pakai data awal */
    }
    return seedData();
  });
  const [tab, setTab] = useState<SheetKey>("jubah");
  const [loading, setLoading] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [toast, setToast] = useState<{ text: string; bad: boolean } | null>(
    null
  );

  const [q, setQ] = useState("");
  const [fMerek, setFMerek] = useState("");
  const [fUkuran, setFUkuran] = useState("");
  const [fStatus, setFStatus] = useState<"" | Status>("");
  const [batas, setBatas] = useState(1);
  const [sort, setSort] = useState<{ key: ColKey; dir: 1 | -1 } | null>(null);
  const [form, setForm] = useState<{
    mode: "add" | "edit";
    row: number;
    v: Record<FormKey, string>;
  } | null>(null);

  const flash = (text: string, bad = false) => {
    setToast({ text, bad });
    window.setTimeout(() => setToast(null), 3200);
  };

  const reload = useCallback(async () => {
    if (!API_URL) return;
    setLoading(true);
    setError("");
    try {
      setData(await apiList());
    } catch (e) {
      setError("Tidak bisa memuat dari Google Sheets: " + errMsg(e));
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    reload();
  }, [reload]);

  useEffect(() => {
    if (API_URL) return;
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    } catch {
      /* abaikan */
    }
  }, [data]);

  useEffect(() => {
    setQ("");
    setFMerek("");
    setFUkuran("");
    setFStatus("");
    setSort(null);
  }, [tab]);

  const list = data[tab];
  const mereks = useMemo(() => uniq(list.map((i) => i.merek)), [list]);
  const ukurans = useMemo(() => uniq(list.map((i) => i.ukuran)), [list]);

  const rows = useMemo(() => {
    const k = q.trim().toLowerCase();
    let r = list.filter(
      (i) =>
        (!k ||
          [i.merek, i.kode, i.model, i.warna, i.ukuran]
            .join(" ")
            .toLowerCase()
            .includes(k)) &&
        (!fMerek || i.merek === fMerek) &&
        (!fUkuran || i.ukuran === fUkuran) &&
        (!fStatus || getStatus(i, batas) === fStatus)
    );
    if (sort) {
      const col = COLS.find((c) => c.key === sort.key)!;
      r = [...r].sort((a, b) => {
        const x = col.get(a),
          y = col.get(b);
        const d =
          typeof x === "number" && typeof y === "number"
            ? x - y
            : String(x).localeCompare(String(y), "id", { numeric: true });
        return d * sort.dir;
      });
    }
    return r;
  }, [list, q, fMerek, fUkuran, fStatus, batas, sort]);

  const sum = useMemo(
    () =>
      rows.reduce(
        (a, i) => {
          const s = Math.max(stokAkhir(i), 0);
          a.pcs += s;
          a.aset += s * i.hargaBeli;
          a.untung += s * (i.hargaJual - i.hargaBeli);
          const st = getStatus(i, batas);
          if (st === "habis") a.habis++;
          else if (st === "menipis") a.menipis++;
          return a;
        },
        { pcs: 0, aset: 0, untung: 0, habis: 0, menipis: 0 }
      ),
    [rows, batas]
  );

  /* ---------- operasi data ---------- */
  const upsert = (sheet: SheetKey, it: Item) =>
    setData((d) => {
      const arr = d[sheet];
      const next = arr.some((x) => x.row === it.row)
        ? arr.map((x) => (x.row === it.row ? it : x))
        : [...arr, it].sort((a, b) => a.row - b.row);
      return { ...d, [sheet]: next };
    });

  const save = async (
    sheet: SheetKey,
    item: Item,
    isNew: boolean,
    orig?: Item
  ): Promise<boolean> => {
    setBusy(true);
    try {
      if (API_URL) {
        const res = await apiPost({
          action: isNew ? "add" : "update",
          sheet,
          row: item.row,
          data: payload(item),
          check: orig ? keyOf(orig) : undefined,
        });
        if (res.item) upsert(sheet, res.item);
      } else {
        const nextRow = data[sheet].reduce((m, x) => Math.max(m, x.row), 1) + 1;
        upsert(sheet, isNew ? { ...item, row: nextRow } : item);
      }
      return true;
    } catch (e) {
      flash(errMsg(e), true);
      return false;
    } finally {
      setBusy(false);
    }
  };

  const remove = async (i: Item) => {
    const nama = [i.merek, i.kode, i.warna, i.ukuran].filter(Boolean).join(" ");
    const info = API_URL ? "\nIsi barisnya di sheet akan dikosongkan." : "";
    if (!window.confirm(`Hapus "${nama}"?${info}`)) return;
    setBusy(true);
    try {
      if (API_URL)
        await apiPost({
          action: "delete",
          sheet: tab,
          row: i.row,
          check: keyOf(i),
        });
      setData((d) => ({ ...d, [tab]: d[tab].filter((x) => x.row !== i.row) }));
      flash("Barang dihapus");
    } catch (e) {
      flash(errMsg(e), true);
    } finally {
      setBusy(false);
    }
  };

  const quick = async (i: Item, field: "masuk" | "terjual") => {
    if (field === "terjual" && stokAkhir(i) <= 0)
      return flash("Stok sudah habis", true);
    const ok = await save(tab, { ...i, [field]: i[field] + 1 }, false, i);
    if (ok) flash(field === "masuk" ? "Barang masuk +1" : "Terjual +1");
  };

  /* ---------- form ---------- */
  const openAdd = () =>
    setForm({
      mode: "add",
      row: 0,
      v: {
        merek: "",
        kode: "",
        model: "",
        warna: "",
        ukuran: "",
        stokLama: "0",
        masuk: "0",
        terjual: "0",
        hargaBeli: "0",
        hargaJual: "0",
      },
    });

  const openEdit = (i: Item) =>
    setForm({
      mode: "edit",
      row: i.row,
      v: {
        merek: i.merek,
        kode: i.kode,
        model: i.model,
        warna: i.warna,
        ukuran: i.ukuran,
        stokLama: String(i.stokLama),
        masuk: String(i.masuk),
        terjual: String(i.terjual),
        hargaBeli: String(i.hargaBeli),
        hargaJual: String(i.hargaJual),
      },
    });

  const setField = (k: FormKey, val: string) =>
    setForm((f) => (f ? { ...f, v: { ...f.v, [k]: val } } : f));

  const submit = async () => {
    if (!form) return;
    const v = form.v;
    if (!v.merek.trim()) return flash("Merek wajib diisi", true);
    if (tab === "campur" && !v.kode.trim())
      return flash("Kode wajib diisi untuk barang campur", true);
    const n = (s: string) => (s.trim() === "" ? 0 : Number(s));
    const nums = [v.stokLama, v.masuk, v.terjual, v.hargaBeli, v.hargaJual].map(
      n
    );
    if (nums.some((x) => !Number.isFinite(x) || x < 0))
      return flash("Angka tidak boleh kosong/negatif", true);
    const item: Item = {
      row: form.row,
      merek: v.merek.trim(),
      kode: v.kode.trim(),
      model: v.model.trim(),
      warna: v.warna.trim(),
      ukuran: v.ukuran.trim(),
      stokLama: nums[0],
      masuk: nums[1],
      terjual: nums[2],
      hargaBeli: nums[3],
      hargaJual: nums[4],
    };
    const orig =
      form.mode === "edit" ? list.find((x) => x.row === form.row) : undefined;
    const ok = await save(tab, item, form.mode === "add", orig);
    if (ok) {
      flash(form.mode === "add" ? "Barang ditambahkan" : "Perubahan disimpan");
      setForm(null);
    }
  };

  const toggleSort = (key: ColKey) =>
    setSort((s) =>
      !s || s.key !== key
        ? { key, dir: 1 }
        : s.dir === 1
        ? { key, dir: -1 }
        : null
    );

  const exportCsv = () => {
    const esc = (v: string | number) => `"${String(v).replace(/"/g, '""')}"`;
    const head = COLS.map((c) => esc(c.label)).join(";");
    const body = rows
      .map((i) => COLS.map((c) => esc(c.get(i))).join(";"))
      .join("\n");
    const blob = new Blob(["\ufeff" + head + "\n" + body], {
      type: "text/csv;charset=utf-8",
    });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download =
      SHEETS[tab].sheetName.replace(/\s+/g, "-").toLowerCase() + ".csv";
    a.click();
    URL.revokeObjectURL(a.href);
  };

  const resetLocal = () => {
    if (window.confirm("Hapus semua data lokal di browser ini?")) {
      setData(seedData());
      flash("Data lokal dihapus");
    }
  };

  const preview = form
    ? (Number(form.v.stokLama) || 0) +
      (Number(form.v.masuk) || 0) -
      (Number(form.v.terjual) || 0)
    : 0;

  /* ---------- tampilan ---------- */
  return (
    <div className="app">
      <style>{CSS}</style>

      <header className="top">
        <div>
          <h1>Stok Barang</h1>
          <p className="sub">
            {API_URL
              ? "Tersambung ke Google Sheets"
              : "Mode lokal: belum tersambung ke Google Sheets (isi API_URL di App.tsx)"}
          </p>
        </div>
        <div className="actions">
          {API_URL ? (
            <button className="btn" onClick={reload} disabled={loading || busy}>
              {loading ? "Memuat…" : "Muat ulang"}
            </button>
          ) : (
            <button className="btn" onClick={resetLocal}>
              Hapus data lokal
            </button>
          )}
          <button className="btn" onClick={exportCsv}>
            Ekspor CSV
          </button>
          <button className="btn primary" onClick={openAdd}>
            + Tambah barang
          </button>
        </div>
      </header>

      {error && <div className="alert">{error}</div>}

      <nav className="tabs">
        {(Object.keys(SHEETS) as SheetKey[]).map((k) => (
          <button
            key={k}
            className={"tab" + (tab === k ? " on" : "")}
            onClick={() => setTab(k)}
          >
            {SHEETS[k].label} <span className="pill">{data[k].length}</span>
          </button>
        ))}
      </nav>

      <section className="cards">
        <Card
          label="Jenis barang"
          value={num(rows.length)}
          note={
            rows.length !== list.length
              ? `dari ${num(list.length)} baris`
              : "baris data"
          }
        />
        <Card label="Total stok" value={`${num(sum.pcs)} pcs`} />
        <Card label="Nilai aset (modal)" value={rp(sum.aset)} />
        <Card
          label="Potensi untung"
          value={rp(sum.untung)}
          note="jika semua stok terjual"
        />
        <Card
          label="Perlu perhatian"
          value={`${sum.habis} habis · ${sum.menipis} menipis`}
          warn={sum.habis + sum.menipis > 0}
        />
      </section>

      <section className="bar">
        <input
          className="in search"
          placeholder="Cari merek, kode, model, warna, ukuran…"
          value={q}
          onChange={(e) => setQ(e.target.value)}
        />
        <select
          className="in"
          value={fMerek}
          onChange={(e) => setFMerek(e.target.value)}
        >
          <option value="">Semua merek</option>
          {mereks.map((m) => (
            <option key={m} value={m}>
              {m}
            </option>
          ))}
        </select>
        <select
          className="in"
          value={fUkuran}
          onChange={(e) => setFUkuran(e.target.value)}
        >
          <option value="">Semua ukuran</option>
          {ukurans.map((u) => (
            <option key={u} value={u}>
              {u}
            </option>
          ))}
        </select>
        <select
          className="in"
          value={fStatus}
          onChange={(e) => setFStatus(e.target.value as "" | Status)}
        >
          <option value="">Semua status</option>
          <option value="tersedia">Tersedia</option>
          <option value="menipis">Menipis</option>
          <option value="habis">Habis</option>
        </select>
        <label className="lim">
          Menipis jika stok ≤
          <input
            className="in sm"
            type="number"
            min={0}
            value={batas}
            onChange={(e) => setBatas(Math.max(0, Number(e.target.value) || 0))}
          />
        </label>
      </section>

      <div className="wrap">
        {loading && (
          <div className="loading">Memuat data dari Google Sheets…</div>
        )}
        <table className="tbl">
          <thead>
            <tr>
              {COLS.map((c) => (
                <th
                  key={c.key}
                  className={c.num ? "r" : ""}
                  onClick={() => toggleSort(c.key)}
                >
                  {c.label}
                  <span className="arr">
                    {sort?.key === c.key ? (sort.dir === 1 ? " ▲" : " ▼") : ""}
                  </span>
                </th>
              ))}
              <th>Aksi</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((i) => {
              const st = getStatus(i, batas);
              return (
                <tr key={i.row}>
                  {COLS.map((c) => {
                    const v = c.get(i);
                    if (c.key === "stokAkhir")
                      return (
                        <td key={c.key} className="r">
                          <span className={"badge " + st}>
                            {num(Number(v))}
                          </span>
                        </td>
                      );
                    return (
                      <td key={c.key} className={c.num ? "r" : ""}>
                        {c.money
                          ? rp(Number(v))
                          : c.num
                          ? num(Number(v))
                          : String(v)}
                      </td>
                    );
                  })}
                  <td className="act">
                    <button
                      className="mini"
                      disabled={busy}
                      onClick={() => quick(i, "masuk")}
                      title="Tambah barang masuk 1"
                    >
                      + Masuk
                    </button>
                    <button
                      className="mini"
                      disabled={busy || stokAkhir(i) <= 0}
                      onClick={() => quick(i, "terjual")}
                      title="Catat terjual 1"
                    >
                      Jual
                    </button>
                    <button
                      className="mini"
                      disabled={busy}
                      onClick={() => openEdit(i)}
                    >
                      Edit
                    </button>
                    <button
                      className="mini danger"
                      disabled={busy}
                      onClick={() => remove(i)}
                    >
                      Hapus
                    </button>
                  </td>
                </tr>
              );
            })}
            {!rows.length && !loading && (
              <tr>
                <td colSpan={COLS.length + 1} className="empty">
                  {list.length === 0
                    ? API_URL
                      ? "Sheet ini belum berisi data."
                      : "Belum ada data. Isi API_URL di App.tsx agar data dimuat dari Google Sheets."
                    : "Tidak ada data yang cocok."}
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {tab === "jubah" && GUIDES.length > 0 && (
        <details className="guide">
          <summary>Panduan ukuran jubah</summary>
          <div className="guides">
            {GUIDES.map((g) => (
              <div key={g.merek}>
                <h3>{g.merek}</h3>
                <table className="gt">
                  <thead>
                    <tr>
                      {g.head.map((h) => (
                        <th key={h}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {g.rows.map((r, ri) => (
                      <tr key={ri}>
                        {r.map((c, ci) => (
                          <td key={ci}>{c}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ))}
          </div>
        </details>
      )}

      <datalist id="dl-merek">
        {mereks.map((v) => (
          <option key={v} value={v} />
        ))}
      </datalist>
      <datalist id="dl-model">
        {uniq(list.map((i) => i.model)).map((v) => (
          <option key={v} value={v} />
        ))}
      </datalist>
      <datalist id="dl-warna">
        {uniq(list.map((i) => i.warna)).map((v) => (
          <option key={v} value={v} />
        ))}
      </datalist>
      <datalist id="dl-ukuran">
        {ukurans.map((v) => (
          <option key={v} value={v} />
        ))}
      </datalist>

      {form && (
        <div
          className="ov"
          onMouseDown={(e) => {
            if (e.target === e.currentTarget) setForm(null);
          }}
        >
          <div className="modal" role="dialog" aria-modal="true">
            <h2>
              {form.mode === "add" ? "Tambah barang" : "Edit barang"} ·{" "}
              {SHEETS[tab].label}
            </h2>
            <div className="grid">
              {FORM_FIELDS.map((f) => (
                <label key={f.key}>
                  {f.label}
                  <input
                    className="in"
                    type={f.num ? "number" : "text"}
                    min={f.num ? 0 : undefined}
                    list={f.list}
                    value={form.v[f.key]}
                    onChange={(e) => setField(f.key, e.target.value)}
                  />
                </label>
              ))}
            </div>
            <p className="calc">
              Stok akhir (otomatis): <b>{num(preview)}</b> · Total aset:{" "}
              <b>
                {rp(Math.max(preview, 0) * (Number(form.v.hargaBeli) || 0))}
              </b>
            </p>
            {API_URL && (
              <p className="hint">
                Jika kolom Barang Masuk/Terjual di sheet berisi rumus (dari
                sheet Transaksi), perubahan disimpan sebagai penyesuaian +/− di
                akhir rumus, jadi rumusnya tidak hilang.
              </p>
            )}
            <div className="mact">
              <button className="btn" onClick={() => setForm(null)}>
                Batal
              </button>
              <button className="btn primary" disabled={busy} onClick={submit}>
                {busy ? "Menyimpan…" : "Simpan"}
              </button>
            </div>
          </div>
        </div>
      )}

      {toast && (
        <div className={"toast" + (toast.bad ? " bad" : "")}>{toast.text}</div>
      )}
    </div>
  );
}

/* =====================================================================
   GAYA
   ===================================================================== */
const CSS = `
:root{--bg:#f4f5f7;--card:#fff;--ink:#1c2230;--mut:#6b7385;--line:#e3e6ec;--brand:#0f766e;--brandd:#0b5d57;
--red:#b91c1c;--redbg:#fee2e2;--amb:#92400e;--ambbg:#fef3c7;--grn:#166534;--grnbg:#dcfce7}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.45 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif}
.app{max-width:1500px;margin:0 auto;padding:16px}
.top{display:flex;flex-wrap:wrap;gap:12px;justify-content:space-between;align-items:center;margin-bottom:12px}
h1{margin:0;font-size:22px}
.sub{margin:2px 0 0;color:var(--mut);font-size:13px}
.actions{display:flex;flex-wrap:wrap;gap:8px}
.btn{border:1px solid var(--line);background:#fff;color:var(--ink);padding:8px 14px;border-radius:8px;font:inherit;cursor:pointer}
.btn:hover:not(:disabled){background:#f0f2f5}
.btn.primary{background:var(--brand);border-color:var(--brand);color:#fff}
.btn.primary:hover:not(:disabled){background:var(--brandd)}
.btn:disabled,.mini:disabled{opacity:.5;cursor:not-allowed}
.alert{background:var(--redbg);color:var(--red);padding:10px 14px;border-radius:8px;margin-bottom:12px}
.tabs{display:flex;gap:6px;margin-bottom:12px;flex-wrap:wrap}
.tab{border:1px solid var(--line);background:#fff;padding:9px 16px;border-radius:999px;font:inherit;cursor:pointer;color:var(--mut)}
.tab.on{background:var(--ink);border-color:var(--ink);color:#fff}
.pill{display:inline-block;margin-left:6px;padding:0 7px;border-radius:999px;background:rgba(127,127,127,.2);font-size:12px}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px;margin-bottom:12px}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px 14px}
.card.warn{border-color:#f5d28a;background:#fffbeb}
.cl{color:var(--mut);font-size:12px}
.cv{font-size:18px;font-weight:650;margin-top:2px}
.cn{color:var(--mut);font-size:12px;margin-top:2px}
.bar{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:10px;align-items:center}
.in{border:1px solid var(--line);background:#fff;border-radius:8px;padding:8px 10px;font:inherit;color:var(--ink);min-width:0}
.in:focus{outline:2px solid var(--brand);outline-offset:-1px}
.search{flex:1 1 240px}
.in.sm{width:64px;margin-left:6px;padding:6px 8px}
.lim{color:var(--mut);font-size:13px;display:flex;align-items:center}
.wrap{position:relative;background:#fff;border:1px solid var(--line);border-radius:10px;overflow:auto;max-height:68vh}
.loading{position:absolute;inset:0;background:rgba(255,255,255,.75);display:flex;align-items:center;justify-content:center;z-index:3;color:var(--mut)}
.tbl{border-collapse:collapse;width:100%;min-width:1180px}
.tbl th{position:sticky;top:0;background:#f8f9fb;text-align:left;padding:9px 10px;border-bottom:1px solid var(--line);font-size:12px;color:var(--mut);white-space:nowrap;cursor:pointer;user-select:none;z-index:2}
.tbl td{padding:8px 10px;border-bottom:1px solid #f0f1f4;white-space:nowrap}
.tbl tr:hover td{background:#fafbfc}
.r{text-align:right!important}
.arr{color:var(--brand)}
.badge{display:inline-block;min-width:28px;text-align:center;padding:1px 8px;border-radius:999px;font-weight:650}
.badge.tersedia{background:var(--grnbg);color:var(--grn)}
.badge.menipis{background:var(--ambbg);color:var(--amb)}
.badge.habis{background:var(--redbg);color:var(--red)}
.act{display:flex;gap:4px}
.mini{border:1px solid var(--line);background:#fff;border-radius:6px;padding:3px 8px;font:inherit;font-size:12px;cursor:pointer}
.mini:hover:not(:disabled){background:#f0f2f5}
.mini.danger{color:var(--red)}
.empty{text-align:center;color:var(--mut);padding:28px!important}
.guide{margin-top:14px;background:#fff;border:1px solid var(--line);border-radius:10px;padding:10px 14px}
.guide summary{cursor:pointer;font-weight:600}
.guides{display:flex;flex-wrap:wrap;gap:24px;margin-top:10px}
.guides h3{margin:0 0 6px;font-size:14px}
.gt{border-collapse:collapse;font-size:13px}
.gt th,.gt td{border:1px solid var(--line);padding:5px 10px;text-align:left}
.gt th{background:#f8f9fb}
.ov{position:fixed;inset:0;background:rgba(20,24,34,.45);display:flex;align-items:center;justify-content:center;padding:16px;z-index:20}
.modal{background:#fff;border-radius:12px;padding:18px;width:100%;max-width:640px;max-height:92vh;overflow:auto}
.modal h2{margin:0 0 12px;font-size:17px}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:10px}
.grid label{display:flex;flex-direction:column;gap:4px;font-size:12px;color:var(--mut)}
.grid .in{color:var(--ink);font-size:14px}
.calc{margin:12px 0 4px}
.hint{margin:4px 0;color:var(--mut);font-size:12px}
.mact{display:flex;justify-content:flex-end;gap:8px;margin-top:12px}
.toast{position:fixed;left:50%;bottom:24px;transform:translateX(-50%);background:var(--ink);color:#fff;padding:10px 18px;border-radius:999px;z-index:30;max-width:90vw}
.toast.bad{background:var(--red)}
@media(max-width:600px){.app{padding:10px}h1{font-size:19px}}
`;
