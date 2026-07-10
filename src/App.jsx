import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { initializeApp } from "firebase/app";
import { getFirestore, doc, setDoc, onSnapshot } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyDR_5nDCLZBQO7xJbcjHSLlqppthYiYpsg",
  authDomain: "bugu-kanri.firebaseapp.com",
  projectId: "bugu-kanri",
  storageBucket: "bugu-kanri.firebasestorage.app",
  messagingSenderId: "420973865056",
  appId: "1:420973865056:web:302b12e5d9df985ba9fb05"
};
const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

const DISCORD_WEBHOOK_URL = "https://discordapp.com/api/webhooks/1519560284430274581/sP-JE_JuI_z0qhwSDmDx0d5e78wrue0djfxoc70aK-M5-FoT30821ZobItAemXkt3q0X";
const sendDiscordNotification = async (type, data) => {
  try {
    let title, fields;
    if (type === "loan") {
      title = "📋 新しい貸出申請が届きました";
      fields = [
        { name: "申請者", value: data.userName, inline: true },
        { name: "借りるもの", value: data.itemName, inline: true },
        { name: "使用目的", value: data.purpose, inline: false },
        { name: "返却予定日", value: data.dueDate, inline: true },
        { name: "申請日", value: data.requestedAt, inline: true },
      ];
    } else {
      title = "🔄 返却申請が届きました";
      fields = [
        { name: "申請者", value: data.userName, inline: true },
        { name: "返却する物品", value: data.itemName, inline: true },
        { name: "メッセージ", value: data.returnMessage || "（なし）", inline: false },
        { name: "申請日", value: data.returnRequestedAt, inline: true },
      ];
    }
    await fetch(DISCORD_WEBHOOK_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ embeds: [{ title, color: 0x2c3f7a, fields, footer: { text: "武道具管理システム" }, timestamp: new Date().toISOString() }] }),
    });
  } catch (e) { console.error("Discord通知エラー:", e); }
};

const INITIAL_ITEMS = [
  { id: "K1", name: "小太刀 K1", category: "小太刀", status: "利用可能", note: "", hidden: false },
  { id: "K2", name: "小太刀 K2", category: "小太刀", status: "利用可能", note: "", hidden: false },
  { id: "K3", name: "小太刀 K3", category: "小太刀", status: "利用可能", note: "", hidden: false },
  { id: "K4", name: "小太刀 K4", category: "小太刀", status: "利用可能", note: "", hidden: false },
  { id: "K5", name: "小太刀 K5", category: "小太刀", status: "利用可能", note: "", hidden: false },
  { id: "C1", name: "長剣 C1", category: "長剣", status: "利用可能", note: "", hidden: false },
  { id: "C2", name: "長剣 C2", category: "長剣", status: "利用可能", note: "", hidden: false },
];

const ADMIN_PASSWORD = "Kumistral";
const today = () => new Date().toISOString().slice(0, 10);

const STATUS_COLORS = {
  "利用可能": { bg: "#e8f5e9", text: "#2e7d32", border: "#a5d6a7" },
  "申請中":   { bg: "#e8eaf6", text: "#283593", border: "#9fa8da" },
  "返却申請中": { bg: "#f3e5f5", text: "#6a1b9a", border: "#ce93d8" },
  "貸出中":   { bg: "#fff3e0", text: "#e65100", border: "#ffb74d" },
  "故障":     { bg: "#fce4ec", text: "#b71c1c", border: "#ef9a9a" },
};

const C = {
  navy: "#1a2744", indigo: "#2c3f7a", gold: "#c9a84c", cream: "#f8f6f0",
  white: "#ffffff", gray50: "#f9fafb", gray100: "#f3f4f6", gray300: "#d1d5db",
  gray500: "#6b7280", gray700: "#374151", red: "#dc2626", green: "#16a34a", purple: "#7c3aed",
};

const s = {
  app: { minHeight: "100vh", background: C.cream, fontFamily: "'Helvetica Neue', Arial, 'Hiragino Kaku Gothic ProN', sans-serif" },
  header: { background: `linear-gradient(135deg, ${C.navy} 0%, ${C.indigo} 100%)`, padding: "0 24px", display: "flex", alignItems: "center", justifyContent: "space-between", height: 64, boxShadow: "0 2px 12px rgba(0,0,0,0.3)" },
  headerTitle: { color: C.gold, fontSize: 20, fontWeight: 700, letterSpacing: "0.08em", margin: 0 },
  headerSub: { color: "rgba(255,255,255,0.6)", fontSize: 12, marginTop: 2 },
  badge: (role) => ({ background: role === "管理者" ? C.gold : "rgba(255,255,255,0.2)", color: role === "管理者" ? C.navy : "#fff", padding: "4px 12px", borderRadius: 20, fontSize: 12, fontWeight: 700 }),
  nav: { background: C.navy, display: "flex", gap: 0, padding: "0 16px" },
  navBtn: (active) => ({ background: active ? C.gold : "transparent", color: active ? C.navy : "rgba(255,255,255,0.7)", border: "none", padding: "10px 20px", cursor: "pointer", fontSize: 13, fontWeight: active ? 700 : 400, transition: "all 0.2s" }),
  main: { maxWidth: 1100, margin: "0 auto", padding: "24px 16px" },
  card: { background: C.white, borderRadius: 10, boxShadow: "0 1px 4px rgba(0,0,0,0.08)", marginBottom: 16, overflow: "hidden" },
  cardHeader: { background: C.navy, color: C.gold, padding: "12px 20px", fontSize: 14, fontWeight: 700, letterSpacing: "0.06em", display: "flex", alignItems: "center", justifyContent: "space-between" },
  table: { width: "100%", borderCollapse: "collapse", fontSize: 13 },
  th: { background: C.gray100, color: C.gray700, padding: "10px 12px", textAlign: "left", fontWeight: 600, borderBottom: `2px solid ${C.gray300}`, whiteSpace: "nowrap" },
  td: { padding: "10px 12px", borderBottom: `1px solid ${C.gray100}`, verticalAlign: "middle" },
  statusBadge: (st) => ({ display: "inline-block", padding: "3px 10px", borderRadius: 12, fontSize: 12, fontWeight: 600, background: STATUS_COLORS[st]?.bg || "#eee", color: STATUS_COLORS[st]?.text || "#333", border: `1px solid ${STATUS_COLORS[st]?.border || "#ccc"}` }),
  btn: (variant = "primary", size = "md") => ({
    background: variant === "primary" ? C.indigo : variant === "gold" ? C.gold : variant === "danger" ? C.red : variant === "success" ? C.green : variant === "purple" ? C.purple : "transparent",
    color: variant === "gold" ? C.navy : variant === "outline" ? C.indigo : "#fff",
    border: variant === "outline" ? `1.5px solid ${C.indigo}` : "none",
    padding: size === "sm" ? "5px 12px" : "8px 18px",
    borderRadius: 6, cursor: "pointer", fontSize: size === "sm" ? 12 : 13, fontWeight: 600, transition: "opacity 0.15s",
  }),
  input: { width: "100%", padding: "9px 12px", border: `1.5px solid ${C.gray300}`, borderRadius: 6, fontSize: 13, outline: "none", boxSizing: "border-box", background: C.white },
  select: { width: "100%", padding: "9px 12px", border: `1.5px solid ${C.gray300}`, borderRadius: 6, fontSize: 13, outline: "none", background: C.white, cursor: "pointer", boxSizing: "border-box" },
  label: { display: "block", fontSize: 12, fontWeight: 600, color: C.gray700, marginBottom: 5 },
  grid3: { display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14 },
  row: { display: "flex", gap: 10, alignItems: "center" },
  tag: (color) => ({ background: color + "22", color: color, padding: "2px 8px", borderRadius: 4, fontSize: 11, fontWeight: 700 }),
  modal: { position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000 },
  modalBox: { background: C.white, borderRadius: 12, maxWidth: 520, width: "90%", maxHeight: "90vh", overflow: "auto", boxShadow: "0 8px 32px rgba(0,0,0,0.3)" },
  modalHeader: { background: C.navy, color: C.gold, padding: "16px 20px", fontSize: 15, fontWeight: 700, borderRadius: "12px 12px 0 0" },
  modalBody: { padding: 24 },
  empty: { textAlign: "center", padding: "40px 20px", color: C.gray500, fontSize: 13 },
  alert: (type) => ({ background: type === "success" ? "#e8f5e9" : "#fce4ec", color: type === "success" ? "#2e7d32" : "#b71c1c", border: `1px solid ${type === "success" ? "#a5d6a7" : "#ef9a9a"}`, borderRadius: 6, padding: "10px 14px", marginBottom: 14, fontSize: 13 }),
};

const saveToFirebase = async (items, loans, notice) => {
  await setDoc(doc(db, "appData", "main"), { items, loans, notice: notice ?? "" });
};

// ── マーキー（自動スクロールテキスト）──────────────────────
function Marquee({ text }) {
  const ref = useRef();
  const [overflow, setOverflow] = useState(false);
  useEffect(() => {
    if (ref.current) setOverflow(ref.current.scrollWidth > ref.current.clientWidth);
  }, [text]);
  return (
    <div style={{ overflow: "hidden", whiteSpace: "nowrap", width: "100%" }} ref={ref}>
      {overflow ? (
        <span style={{ display: "inline-block", animation: "marquee 12s linear infinite", paddingLeft: "100%" }}>
          {text}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{text}
          <style>{`@keyframes marquee { 0% { transform: translateX(0); } 100% { transform: translateX(-50%); } }`}</style>
        </span>
      ) : <span>{text}</span>}
    </div>
  );
}

// ── ログイン画面 ─────────────────────────────────────────────
function LoginScreen({ onLogin, notice }) {
  const [role, setRole] = useState("一般ユーザー");
  const [name, setName] = useState("");
  const [pw, setPw] = useState("");
  const [err, setErr] = useState("");

  const handle = () => {
    if (!name.trim()) { setErr("名前を入力してください"); return; }
    if (role === "管理者" && pw !== ADMIN_PASSWORD) { setErr("管理者パスワードが違います"); return; }
    onLogin({ role, name: name.trim() });
  };

  return (
    <div style={{ minHeight: "100vh", background: `linear-gradient(160deg, ${C.navy} 0%, ${C.indigo} 60%, #3d5296 100%)`, display: "flex", alignItems: "center", justifyContent: "center", flexDirection: "column", gap: 16 }}>
      {notice && (
        <div style={{ width: "min(400px, 90%)", background: "rgba(201,168,76,0.15)", border: `1px solid ${C.gold}`, borderRadius: 8, padding: "10px 16px", color: C.gold, fontSize: 13, fontWeight: 600, overflow: "hidden" }}>
          <Marquee text={`📢 ${notice}`} />
        </div>
      )}
      <div style={{ background: C.white, borderRadius: 16, padding: "40px 36px", width: "min(360px, 90%)", boxShadow: "0 12px 48px rgba(0,0,0,0.4)" }}>
        <div style={{ textAlign: "center", marginBottom: 28 }}>
          <div style={{ fontSize: 36, marginBottom: 8 }}>⚔️</div>
          <h1 style={{ color: C.navy, fontSize: 20, fontWeight: 800, margin: 0, letterSpacing: "0.06em" }}>武道具管理システム</h1>
          <p style={{ color: C.gray500, fontSize: 12, marginTop: 6 }}>Bugu Management System</p>
        </div>
        {err && <div style={s.alert("error")}>{err}</div>}
        <div style={{ marginBottom: 14 }}>
          <label style={s.label}>ログインの種類</label>
          <select style={s.select} value={role} onChange={e => { setRole(e.target.value); setErr(""); }}>
            <option>一般ユーザー</option>
            <option>管理者</option>
          </select>
        </div>
        <div style={{ marginBottom: 14 }}>
          <label style={s.label}>氏名</label>
          <input style={s.input} placeholder="山田 太郎" value={name} onChange={e => setName(e.target.value)} onKeyDown={e => e.key === "Enter" && handle()} />
        </div>
        {role === "管理者" && (
          <div style={{ marginBottom: 14 }}>
            <label style={s.label}>管理者パスワード</label>
            <input type="password" style={s.input} placeholder="••••••••" value={pw} onChange={e => setPw(e.target.value)} onKeyDown={e => e.key === "Enter" && handle()} />
          </div>
        )}
        <button style={{ ...s.btn("primary"), width: "100%", padding: "11px", fontSize: 14, marginTop: 8, background: C.navy }} onClick={handle}>ログイン →</button>
      </div>
    </div>
  );
}

// ── 物品一覧タブ ─────────────────────────────────────────────
function ItemsTab({ items, loans, user, onRequestLoan, onChangeStatus, onChangeNote, onChangeName, onDeleteItem, onToggleHidden, onAddItem, onImportXlsx, onRequestReturn }) {
  const [statusFilter, setStatusFilter] = useState("すべて");
  const [showAdd, setShowAdd] = useState(false);
  const [newItem, setNewItem] = useState({ id: "", name: "", category: "" });
  const [editStatus, setEditStatus] = useState(null);
  const [addErr, setAddErr] = useState("");
  const [sortKey, setSortKey] = useState("id");
  const [sortDir, setSortDir] = useState("asc");
  const [confirmDelete, setConfirmDelete] = useState(null);
  const fileRef = useRef();

  const handleSort = (key) => {
    if (sortKey === key) setSortDir(d => d === "asc" ? "desc" : "asc");
    else { setSortKey(key); setSortDir("asc"); }
  };

  const STATUS_ORDER = { "利用可能": 0, "申請中": 1, "返却申請中": 2, "貸出中": 3, "故障": 4 };
  const getDueDate = (itemId) => loans.find(l => l.itemId === itemId && l.status === "承認済" && !l.returned)?.dueDate || "";

  const visibleItems = items.filter(i => user.role === "管理者" || !i.hidden);
  const filtered = visibleItems
    .filter(i => statusFilter === "すべて" || i.status === statusFilter)
    .sort((a, b) => {
      let valA, valB;
      if (sortKey === "id") { valA = a.id; valB = b.id; }
      else if (sortKey === "category") { valA = a.category; valB = b.category; }
      else if (sortKey === "status") { valA = STATUS_ORDER[a.status] ?? 9; valB = STATUS_ORDER[b.status] ?? 9; return sortDir === "asc" ? valA - valB : valB - valA; }
      else if (sortKey === "dueDate") { valA = getDueDate(a.id) || "9999"; valB = getDueDate(b.id) || "9999"; }
      if (valA < valB) return sortDir === "asc" ? -1 : 1;
      if (valA > valB) return sortDir === "asc" ? 1 : -1;
      return 0;
    });

  const pendingLoan = (itemId) => loans.find(l => l.itemId === itemId && l.status === "審査中");
  const activeLoan = (itemId) => loans.find(l => l.itemId === itemId && l.status === "承認済" && !l.returned);
  const returnPending = (itemId) => loans.find(l => l.itemId === itemId && l.returnStatus === "返却申請中");

  const adminStatusOptions = (currentStatus) => {
    if (currentStatus === "申請中") return ["申請中", "利用可能", "故障"];
    if (currentStatus === "貸出中" || currentStatus === "返却申請中") return [currentStatus, "利用可能", "故障"];
    return ["利用可能", "故障"];
  };

  const handleAdd = () => {
    if (!newItem.id.trim() || !newItem.name.trim() || !newItem.category.trim()) { setAddErr("すべて入力してください"); return; }
    if (items.find(i => i.id === newItem.id.trim())) { setAddErr("IDが重複しています"); return; }
    onAddItem({ ...newItem, id: newItem.id.trim(), name: newItem.name.trim(), category: newItem.category.trim(), status: "利用可能", note: "", hidden: false });
    setNewItem({ id: "", name: "", category: "" }); setShowAdd(false); setAddErr("");
  };

  const handleFile = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const wb = XLSX.read(ev.target.result, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      onImportXlsx(XLSX.utils.sheet_to_json(ws));
    };
    reader.readAsArrayBuffer(file); e.target.value = "";
  };

  const SortTh = ({ label, sortK }) => (
    <th style={{ ...s.th, cursor: sortK ? "pointer" : "default", userSelect: "none" }} onClick={() => sortK && handleSort(sortK)}>
      {label}
      {sortK && sortKey === sortK && <span style={{ marginLeft: 4 }}>{sortDir === "asc" ? "▲" : "▼"}</span>}
      {sortK && sortKey !== sortK && <span style={{ marginLeft: 4, color: C.gray300 }}>▲▼</span>}
    </th>
  );

  return (
    <div>
      <div style={{ ...s.row, justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
        <div style={{ ...s.row, flexWrap: "wrap", gap: 6 }}>
          {["すべて", "利用可能", "申請中", "返却申請中", "貸出中", "故障"].map(f => (
            <button key={f} style={{ ...s.btn(statusFilter === f ? "primary" : "outline", "sm"), borderColor: statusFilter === f ? C.indigo : C.gray300 }} onClick={() => setStatusFilter(f)}>{f}</button>
          ))}
        </div>
        {user.role === "管理者" && (
          <div style={s.row}>
            <button style={s.btn("outline", "sm")} onClick={() => fileRef.current.click()}>📂 Excel読込</button>
            <button style={s.btn("gold", "sm")} onClick={() => setShowAdd(true)}>＋ 物品追加</button>
            <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleFile} />
          </div>
        )}
      </div>

      <div style={s.card}>
        <div style={s.cardHeader}>
          <span>物品一覧</span>
          <span style={{ fontSize: 12, fontWeight: 400, color: "rgba(255,255,255,0.6)" }}>{filtered.length} 件</span>
        </div>
        <div style={{ overflowX: "auto" }}>
          <table style={s.table}>
            <thead>
              <tr>
                <SortTh label="ID" sortK="id" />
                <th style={s.th}>名称</th>
                <SortTh label="種別" sortK="category" />
                <SortTh label="状態" sortK="status" />
                <SortTh label="関連情報（返却予定日）" sortK="dueDate" />
                <th style={s.th}>備考</th>
                <th style={s.th}>操作</th>
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr><td colSpan={7} style={s.empty}>物品がありません</td></tr>
              ) : filtered.map(item => {
                const pending = pendingLoan(item.id);
                const active = activeLoan(item.id);
                const retPending = returnPending(item.id);
                const relatedLoan = pending || active;
                const isHidden = item.hidden;
                return (
                  <tr key={item.id} style={{ background: isHidden ? "#f0f0f0" : item.status === "故障" ? "#fff5f5" : item.status === "申請中" ? "#f5f6ff" : item.status === "返却申請中" ? "#fdf4ff" : "white", opacity: isHidden ? 0.6 : 1 }}>
                    <td style={s.td}><span style={{ fontFamily: "monospace", fontWeight: 700, color: C.indigo }}>{item.id}</span>{isHidden && <span style={{ ...s.tag("#888"), marginLeft: 6 }}>非表示</span>}</td>

                    {/* 名称（管理者は編集可） */}
                    <td style={s.td}>
                      {user.role === "管理者" && editStatus?.id === item.id + "_name" ? (
                        <div style={s.row}>
                          <input style={{ ...s.input, padding: "4px 8px", fontSize: 12 }} value={editStatus.val} onChange={e => setEditStatus({ ...editStatus, val: e.target.value })}
                            onKeyDown={e => { if (e.key === "Enter") { onChangeName(item.id, editStatus.val); setEditStatus(null); } if (e.key === "Escape") setEditStatus(null); }} autoFocus />
                          <button style={s.btn("success", "sm")} onClick={() => { onChangeName(item.id, editStatus.val); setEditStatus(null); }}>✓</button>
                          <button style={s.btn("outline", "sm")} onClick={() => setEditStatus(null)}>✕</button>
                        </div>
                      ) : (
                        <div style={s.row}>
                          <span>{item.name}</span>
                          {user.role === "管理者" && <button style={{ ...s.btn("outline", "sm"), fontSize: 11 }} onClick={() => setEditStatus({ id: item.id + "_name", val: item.name })}>✏️</button>}
                        </div>
                      )}
                    </td>

                    <td style={s.td}><span style={s.tag(C.indigo)}>{item.category}</span></td>

                    {/* 状態 */}
                    <td style={s.td}>
                      {user.role === "管理者" && editStatus?.id === item.id + "_status" ? (
                        <div style={s.row}>
                          <select style={{ ...s.select, width: "auto" }} value={editStatus.val} onChange={e => setEditStatus({ ...editStatus, val: e.target.value })}>
                            {adminStatusOptions(item.status).map(st => <option key={st}>{st}</option>)}
                          </select>
                          <button style={s.btn("success", "sm")} onClick={() => { onChangeStatus(item.id, editStatus.val); setEditStatus(null); }}>✓</button>
                          <button style={s.btn("outline", "sm")} onClick={() => setEditStatus(null)}>✕</button>
                        </div>
                      ) : (
                        <div style={s.row}>
                          <span style={s.statusBadge(item.status)}>{item.status}</span>
                          {user.role === "管理者" && <button style={{ ...s.btn("outline", "sm"), fontSize: 11 }} onClick={() => setEditStatus({ id: item.id + "_status", val: item.status })}>変更</button>}
                        </div>
                      )}
                    </td>

                    {/* 関連情報 */}
                    <td style={s.td}>
                      {retPending ? (
                        <div style={{ fontSize: 12 }}>
                          <div style={{ fontWeight: 600, color: C.purple }}>{retPending.userName}</div>
                          <div style={{ color: C.purple }}>返却申請中</div>
                          {retPending.returnMessage && <div style={{ color: C.gray500 }}>「{retPending.returnMessage}」</div>}
                        </div>
                      ) : relatedLoan ? (
                        <div style={{ fontSize: 12 }}>
                          <div style={{ fontWeight: 600 }}>{relatedLoan.userName}</div>
                          {relatedLoan.status === "審査中"
                            ? <div style={{ color: "#283593" }}>申請中（{relatedLoan.requestedAt}）</div>
                            : <div style={{ color: C.gray500 }}>返却予定: <b style={{ color: relatedLoan.dueDate < today() ? C.red : "inherit" }}>{relatedLoan.dueDate}</b>{relatedLoan.dueDate < today() && <span style={{ color: C.red }}> ⚠</span>}</div>}
                        </div>
                      ) : <span style={{ color: C.gray300 }}>—</span>}
                    </td>

                    {/* 備考 */}
                    <td style={{ ...s.td, fontSize: 12 }}>
                      {user.role === "管理者" && editStatus?.id === item.id + "_note" ? (
                        <div style={s.row}>
                          <input style={{ ...s.input, padding: "4px 8px", fontSize: 12 }} value={editStatus.val} onChange={e => setEditStatus({ ...editStatus, val: e.target.value })}
                            onKeyDown={e => { if (e.key === "Enter") { onChangeNote(item.id, editStatus.val); setEditStatus(null); } if (e.key === "Escape") setEditStatus(null); }} autoFocus />
                          <button style={s.btn("success", "sm")} onClick={() => { onChangeNote(item.id, editStatus.val); setEditStatus(null); }}>✓</button>
                          <button style={s.btn("outline", "sm")} onClick={() => setEditStatus(null)}>✕</button>
                        </div>
                      ) : (
                        <div style={s.row}>
                          <span style={{ color: item.note ? C.gray700 : C.gray300 }}>{item.note || "—"}</span>
                          {user.role === "管理者" && <button style={{ ...s.btn("outline", "sm"), fontSize: 11 }} onClick={() => setEditStatus({ id: item.id + "_note", val: item.note || "" })}>編集</button>}
                        </div>
                      )}
                    </td>

                    {/* 操作 */}
                    <td style={s.td}>
                      <div style={{ ...s.row, flexWrap: "wrap", gap: 4 }}>
                        {user.role !== "管理者" && item.status === "利用可能" && !isHidden && (
                          <button style={s.btn("primary", "sm")} onClick={() => onRequestLoan(item)}>申請</button>
                        )}
                        {user.role !== "管理者" && item.status === "申請中" && pending?.userName === user.name && (
                          <span style={{ fontSize: 12, color: "#283593", fontWeight: 600 }}>審査待ち</span>
                        )}
                        {user.role !== "管理者" && item.status === "貸出中" && active?.userName === user.name && (
                          <button style={s.btn("purple", "sm")} onClick={() => onRequestReturn(item, active)}>返却申請</button>
                        )}
                        {user.role !== "管理者" && item.status === "返却申請中" && retPending?.userName === user.name && (
                          <span style={{ fontSize: 12, color: C.purple, fontWeight: 600 }}>返却審査待ち</span>
                        )}
                        {user.role === "管理者" && (
                          <>
                            <button style={{ ...s.btn("outline", "sm"), fontSize: 11, color: isHidden ? C.green : C.gray500, borderColor: isHidden ? C.green : C.gray300 }} onClick={() => onToggleHidden(item.id)}>
                              {isHidden ? "表示" : "非表示"}
                            </button>
                            <button style={{ ...s.btn("danger", "sm"), fontSize: 11 }} onClick={() => setConfirmDelete(item)}>削除</button>
                          </>
                        )}
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      {/* 削除確認モーダル */}
      {confirmDelete && (
        <div style={s.modal}>
          <div style={s.modalBox}>
            <div style={s.modalHeader}>物品の削除</div>
            <div style={s.modalBody}>
              <p style={{ marginBottom: 20 }}>「<b>{confirmDelete.name}</b>」を削除します。この操作は取り消せません。</p>
              <div style={s.row}>
                <button style={s.btn("danger")} onClick={() => { onDeleteItem(confirmDelete.id); setConfirmDelete(null); }}>削除する</button>
                <button style={s.btn("outline")} onClick={() => setConfirmDelete(null)}>キャンセル</button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* 物品追加モーダル */}
      {showAdd && (
        <div style={s.modal} onClick={e => e.target === e.currentTarget && setShowAdd(false)}>
          <div style={s.modalBox}>
            <div style={s.modalHeader}>物品を追加</div>
            <div style={s.modalBody}>
              {addErr && <div style={s.alert("error")}>{addErr}</div>}
              <div style={{ ...s.grid3, marginBottom: 14 }}>
                <div><label style={s.label}>ID</label><input style={s.input} placeholder="K6" value={newItem.id} onChange={e => setNewItem({ ...newItem, id: e.target.value })} /></div>
                <div><label style={s.label}>名称</label><input style={s.input} placeholder="小太刀 K6" value={newItem.name} onChange={e => setNewItem({ ...newItem, name: e.target.value })} /></div>
                <div><label style={s.label}>種別</label><input style={s.input} placeholder="小太刀" value={newItem.category} onChange={e => setNewItem({ ...newItem, category: e.target.value })} /></div>
              </div>
              <div style={s.row}>
                <button style={s.btn("primary")} onClick={handleAdd}>追加</button>
                <button style={s.btn("outline")} onClick={() => { setShowAdd(false); setAddErr(""); }}>キャンセル</button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ── 貸出申請モーダル ──────────────────────────────────────────
function LoanRequestModal({ item, user, onSubmit, onClose }) {
  const [purpose, setPurpose] = useState("");
  const [dueDate, setDueDate] = useState("");
  const [err, setErr] = useState("");
  const handle = () => {
    if (!purpose.trim()) { setErr("使用目的を入力してください"); return; }
    if (!dueDate) { setErr("返却予定日を入力してください"); return; }
    if (dueDate <= today()) { setErr("返却予定日は明日以降を指定してください"); return; }
    onSubmit({ itemId: item.id, itemName: item.name, userName: user.name, purpose: purpose.trim(), dueDate, requestedAt: today(), status: "審査中", returned: false, id: Date.now().toString() });
    onClose();
  };
  return (
    <div style={s.modal}>
      <div style={s.modalBox}>
        <div style={s.modalHeader}>貸出申請 — {item.name}</div>
        <div style={s.modalBody}>
          {err && <div style={s.alert("error")}>{err}</div>}
          <div style={{ background: C.gray50, border: `1px solid ${C.gray100}`, borderRadius: 8, padding: "12px 16px", marginBottom: 18 }}>
            <div style={{ fontSize: 12, color: C.gray500 }}>申請者</div>
            <div style={{ fontSize: 15, fontWeight: 700, color: C.navy }}>{user.name}</div>
          </div>
          <div style={{ marginBottom: 14 }}><label style={s.label}>借りるもの</label><input style={{ ...s.input, background: C.gray50 }} value={item.name} readOnly /></div>
          <div style={{ marginBottom: 14 }}><label style={s.label}>使用目的 <span style={{ color: C.red }}>*</span></label><textarea style={{ ...s.input, minHeight: 80, resize: "vertical" }} placeholder="例: 演武練習のため" value={purpose} onChange={e => setPurpose(e.target.value)} /></div>
          <div style={{ marginBottom: 22 }}><label style={s.label}>返却予定日 <span style={{ color: C.red }}>*</span></label><input type="date" style={s.input} min={today()} value={dueDate} onChange={e => setDueDate(e.target.value)} /></div>
          <div style={s.row}><button style={s.btn("primary")} onClick={handle}>申請を送信</button><button style={s.btn("outline")} onClick={onClose}>キャンセル</button></div>
        </div>
      </div>
    </div>
  );
}

// ── 返却申請モーダル ──────────────────────────────────────────
function ReturnRequestModal({ item, loan, user, onSubmit, onClose }) {
  const [message, setMessage] = useState("");
  const handle = () => {
    onSubmit({ loanId: loan.id, itemId: item.id, itemName: item.name, userName: user.name, returnMessage: message.trim(), returnRequestedAt: today() });
    onClose();
  };
  return (
    <div style={s.modal}>
      <div style={s.modalBox}>
        <div style={{ ...s.modalHeader, background: C.purple }}>返却申請 — {item.name}</div>
        <div style={s.modalBody}>
          <div style={{ background: C.gray50, border: `1px solid ${C.gray100}`, borderRadius: 8, padding: "12px 16px", marginBottom: 18 }}>
            <div style={{ fontSize: 12, color: C.gray500 }}>返却する物品</div>
            <div style={{ fontSize: 15, fontWeight: 700, color: C.navy }}>{item.name}</div>
          </div>
          <div style={{ marginBottom: 22 }}>
            <label style={s.label}>管理者へのメッセージ（任意）</label>
            <textarea style={{ ...s.input, minHeight: 80, resize: "vertical" }} placeholder="例: 返却しました。ロッカーに置いています。" value={message} onChange={e => setMessage(e.target.value)} />
          </div>
          <div style={s.row}><button style={{ ...s.btn("purple") }} onClick={handle}>返却申請を送信</button><button style={s.btn("outline")} onClick={onClose}>キャンセル</button></div>
        </div>
      </div>
    </div>
  );
}

// ── 申請管理タブ ──────────────────────────────────────────────
function RequestsTab({ loans, onApprove, onReject, onApproveReturn, onRejectReturn, user }) {
  const [filter, setFilter] = useState("審査中");
  const allRequests = loans.flatMap(l => {
    const result = [{ ...l, _type: "loan" }];
    if (l.returnStatus === "返却申請中") result.push({ ...l, _type: "return", _id: l.id + "_return" });
    return result;
  });
  const filtered = allRequests.filter(l => {
    if (filter === "すべて") return true;
    if (filter === "審査中") return (l._type === "loan" && l.status === "審査中") || (l._type === "return" && l.returnStatus === "返却申請中");
    if (filter === "承認済") return l._type === "loan" && l.status === "承認済";
    if (filter === "却下") return l._type === "loan" && l.status === "却下";
    return true;
  });
  const pendingCount = loans.filter(l => l.status === "審査中").length + loans.filter(l => l.returnStatus === "返却申請中").length;
  const statusColor = { "審査中": C.gold, "承認済": C.green, "却下": C.red };

  return (
    <div>
      <div style={{ ...s.row, marginBottom: 16, flexWrap: "wrap", gap: 8 }}>
        {["すべて", "審査中", "承認済", "却下"].map(f => (
          <button key={f} style={{ ...s.btn(filter === f ? "primary" : "outline", "sm") }} onClick={() => setFilter(f)}>
            {f}{f === "審査中" && pendingCount > 0 && <span style={{ marginLeft: 6, background: C.red + "33", color: C.red, padding: "1px 7px", borderRadius: 8, fontSize: 11, fontWeight: 700 }}>{pendingCount}</span>}
          </button>
        ))}
      </div>
      {filtered.length === 0 ? (
        <div style={s.card}><div style={s.empty}>該当する申請はありません</div></div>
      ) : filtered.map(loan => {
        const isReturn = loan._type === "return";
        return (
          <div key={loan._id || loan.id} style={{ ...s.card, borderLeft: `4px solid ${isReturn ? C.purple : (statusColor[loan.status] || C.gray300)}` }}>
            <div style={{ padding: "14px 20px", display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12 }}>
              <div style={{ flex: 1 }}>
                <div style={{ ...s.row, gap: 8, marginBottom: 6 }}>
                  <span style={{ fontSize: 15, fontWeight: 700, color: C.navy }}>{loan.itemName}</span>
                  {isReturn
                    ? <span style={{ background: C.purple + "22", color: C.purple, padding: "2px 10px", borderRadius: 10, fontSize: 12, fontWeight: 700 }}>返却申請</span>
                    : <span style={{ background: (statusColor[loan.status] || C.gray500) + "22", color: statusColor[loan.status] || C.gray500, padding: "2px 10px", borderRadius: 10, fontSize: 12, fontWeight: 700 }}>{loan.status}</span>
                  }
                  {!isReturn && loan.returned && <span style={s.tag(C.green)}>返却済</span>}
                </div>
                {isReturn ? (
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: "4px 16px", fontSize: 13, color: C.gray700 }}>
                    <div><span style={{ color: C.gray500 }}>申請者: </span>{loan.userName}</div>
                    <div><span style={{ color: C.gray500 }}>返却申請日: </span>{loan.returnRequestedAt}</div>
                    {loan.returnMessage && <div><span style={{ color: C.gray500 }}>メッセージ: </span>{loan.returnMessage}</div>}
                  </div>
                ) : (
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: "4px 16px", fontSize: 13, color: C.gray700 }}>
                    <div><span style={{ color: C.gray500 }}>申請者: </span>{loan.userName}</div>
                    <div><span style={{ color: C.gray500 }}>申請日: </span>{loan.requestedAt}</div>
                    <div><span style={{ color: C.gray500 }}>返却予定: </span><b style={{ color: loan.dueDate < today() && !loan.returned ? C.red : "inherit" }}>{loan.dueDate}</b>{loan.dueDate < today() && !loan.returned && loan.status === "承認済" && <span style={{ color: C.red }}> ⚠ 期限超過</span>}</div>
                    <div><span style={{ color: C.gray500 }}>使用目的: </span>{loan.purpose}</div>
                  </div>
                )}
                {loan.rejectedReason && <div style={{ marginTop: 6, fontSize: 12, color: C.red }}>却下理由: {loan.rejectedReason}</div>}
              </div>
              <div style={{ ...s.row, gap: 8, flexShrink: 0 }}>
                {isReturn ? (
                  <>
                    <button style={s.btn("success", "sm")} onClick={() => onApproveReturn(loan.id)}>✓ 返却承認</button>
                    <button style={s.btn("danger", "sm")} onClick={() => onRejectReturn(loan.id)}>✕ 却下</button>
                  </>
                ) : (
                  <>
                    {loan.status === "審査中" && <><button style={s.btn("success", "sm")} onClick={() => onApprove(loan.id)}>✓ 承認</button><button style={s.btn("danger", "sm")} onClick={() => onReject(loan.id)}>✕ 却下</button></>}
                  </>
                )}
              </div>
            </div>
          </div>
        );
      })}
    </div>
  );
}

// ── 却下理由モーダル ──────────────────────────────────────────
function RejectModal({ onConfirm, onClose }) {
  const [reason, setReason] = useState("");
  return (
    <div style={s.modal}>
      <div style={s.modalBox}>
        <div style={s.modalHeader}>却下理由</div>
        <div style={s.modalBody}>
          <label style={s.label}>却下理由（任意）</label>
          <textarea style={{ ...s.input, minHeight: 80, marginBottom: 16 }} placeholder="例: 既に他の利用者が予約済みのため" value={reason} onChange={e => setReason(e.target.value)} />
          <div style={s.row}><button style={s.btn("danger")} onClick={() => onConfirm(reason)}>却下する</button><button style={s.btn("outline")} onClick={onClose}>キャンセル</button></div>
        </div>
      </div>
    </div>
  );
}

// ── マイ申請タブ ──────────────────────────────────────────────
function MyRequestsTab({ loans, user }) {
  const myLoans = loans.filter(l => l.userName === user.name);
  const statusColor = { "審査中": C.gold, "承認済": C.green, "却下": C.red };
  return (
    <div>
      {myLoans.length === 0 ? (
        <div style={s.card}><div style={s.empty}>申請履歴がありません<br /><span style={{ fontSize: 12 }}>物品一覧から貸出申請を行えます</span></div></div>
      ) : myLoans.map(loan => (
        <div key={loan.id} style={{ ...s.card, borderLeft: `4px solid ${statusColor[loan.status] || C.gray300}` }}>
          <div style={{ padding: "14px 20px" }}>
            <div style={{ ...s.row, marginBottom: 6, flexWrap: "wrap", gap: 6 }}>
              <span style={{ fontSize: 15, fontWeight: 700, color: C.navy }}>{loan.itemName}</span>
              <span style={{ background: (statusColor[loan.status] || C.gray500) + "22", color: statusColor[loan.status] || C.gray500, padding: "2px 10px", borderRadius: 10, fontSize: 12, fontWeight: 700 }}>{loan.status}</span>
              {loan.returned && <span style={s.tag(C.green)}>返却済</span>}
              {loan.returnStatus === "返却申請中" && <span style={{ background: C.purple + "22", color: C.purple, padding: "2px 10px", borderRadius: 10, fontSize: 12, fontWeight: 700 }}>返却申請中</span>}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: "4px 16px", fontSize: 13, color: C.gray700 }}>
              <div><span style={{ color: C.gray500 }}>申請日: </span>{loan.requestedAt}</div>
              <div><span style={{ color: C.gray500 }}>返却予定: </span><b style={{ color: loan.dueDate < today() && !loan.returned && loan.status === "承認済" ? C.red : "inherit" }}>{loan.dueDate}</b></div>
              <div><span style={{ color: C.gray500 }}>使用目的: </span>{loan.purpose}</div>
            </div>
            {loan.rejectedReason && <div style={{ marginTop: 8, fontSize: 12, color: C.red, background: "#fce4ec", padding: "6px 10px", borderRadius: 4 }}>却下理由: {loan.rejectedReason}</div>}
          </div>
        </div>
      ))}
    </div>
  );
}

// ── ダッシュボード ────────────────────────────────────────────
function Dashboard({ items, loans }) {
  const counts = {
    total: items.filter(i => !i.hidden).length,
    available: items.filter(i => i.status === "利用可能" && !i.hidden).length,
    applying: items.filter(i => i.status === "申請中").length,
    returning: items.filter(i => i.status === "返却申請中").length,
    lent: items.filter(i => i.status === "貸出中").length,
    broken: items.filter(i => i.status === "故障").length,
    overdue: loans.filter(l => l.status === "承認済" && !l.returned && l.dueDate < today()).length,
  };
  const stats = [
    { label: "総物品数", value: counts.total, color: C.navy, icon: "⚔️" },
    { label: "利用可能", value: counts.available, color: C.green, icon: "✅" },
    { label: "貸出申請中", value: counts.applying, color: "#283593", icon: "📝" },
    { label: "返却申請中", value: counts.returning, color: C.purple, icon: "🔄" },
    { label: "貸出中", value: counts.lent, color: "#e65100", icon: "📤" },
    { label: "故障", value: counts.broken, color: C.red, icon: "⚠️" },
    { label: "期限超過", value: counts.overdue, color: C.red, icon: "🔴" },
  ];
  return (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 12 }}>
      {stats.map(st => (
        <div key={st.label} style={{ ...s.card, marginBottom: 0 }}>
          <div style={{ padding: "16px", textAlign: "center" }}>
            <div style={{ fontSize: 28, marginBottom: 4 }}>{st.icon}</div>
            <div style={{ fontSize: 28, fontWeight: 800, color: st.color }}>{st.value}</div>
            <div style={{ fontSize: 12, color: C.gray500, marginTop: 2 }}>{st.label}</div>
          </div>
        </div>
      ))}
    </div>
  );
}

// ── お知らせ編集モーダル ──────────────────────────────────────
function NoticeModal({ current, onSave, onClose }) {
  const [text, setText] = useState(current || "");
  return (
    <div style={s.modal}>
      <div style={s.modalBox}>
        <div style={s.modalHeader}>ログイン画面のお知らせ文</div>
        <div style={s.modalBody}>
          <label style={s.label}>お知らせ（1行・長い場合は自動スクロール）</label>
          <input style={{ ...s.input, marginBottom: 16 }} placeholder="例: 7月の活動は休止します" value={text} onChange={e => setText(e.target.value)} maxLength={200} />
          <div style={{ fontSize: 12, color: C.gray500, marginBottom: 16 }}>空にすると非表示になります</div>
          <div style={s.row}><button style={s.btn("primary")} onClick={() => onSave(text)}>保存</button><button style={s.btn("outline")} onClick={onClose}>キャンセル</button></div>
        </div>
      </div>
    </div>
  );
}

// ── メインアプリ ─────────────────────────────────────────────
export default function App() {
  const [user, setUser] = useState(null);
  const [items, setItems] = useState([]);
  const [loans, setLoans] = useState([]);
  const [notice, setNotice] = useState("");
  const [loading, setLoading] = useState(true);
  const [tab, setTab] = useState("物品一覧");
  const [loanTarget, setLoanTarget] = useState(null);
  const [returnTarget, setReturnTarget] = useState(null);
  const [rejectTarget, setRejectTarget] = useState(null);
  const [rejectReturnTarget, setRejectReturnTarget] = useState(null);
  const [showNoticeModal, setShowNoticeModal] = useState(false);
  const [toast, setToast] = useState(null);

  useEffect(() => {
    const unsub = onSnapshot(doc(db, "appData", "main"), (snap) => {
      if (snap.exists()) {
        const data = snap.data();
        setItems(data.items || []);
        setLoans(data.loans || []);
        setNotice(data.notice || "");
      } else {
        saveToFirebase(INITIAL_ITEMS, [], "");
        setItems(INITIAL_ITEMS); setLoans([]);
      }
      setLoading(false);
    });
    return () => unsub();
  }, []);

  const showToast = (msg, type = "success") => { setToast({ msg, type }); setTimeout(() => setToast(null), 3500); };

  const save = async (newItems, newLoans, newNotice) => {
    await setDoc(doc(db, "appData", "main"), { items: newItems ?? items, loans: newLoans ?? loans, notice: newNotice ?? notice });
  };

  if (loading) return <div style={{ minHeight: "100vh", background: C.navy, display: "flex", alignItems: "center", justifyContent: "center" }}><div style={{ color: C.gold, fontSize: 18, fontWeight: 700 }}>⚔️ 読み込み中...</div></div>;
  if (!user) return <LoginScreen onLogin={setUser} notice={notice} />;

  const tabs = user.role === "管理者" ? ["物品一覧", "申請管理", "ダッシュボード"] : ["物品一覧", "マイ申請"];
  const pendingCount = loans.filter(l => l.status === "審査中").length + loans.filter(l => l.returnStatus === "返却申請中").length;

  // 貸出申請送信
  const handleSubmitLoan = async (loan) => {
    const newLoans = [loan, ...loans];
    const newItems = items.map(i => i.id === loan.itemId ? { ...i, status: "申請中" } : i);
    await save(newItems, newLoans);
    sendDiscordNotification("loan", loan);
    showToast("申請を送信しました。管理者の承認をお待ちください。");
    setTab("マイ申請");
  };

  // 貸出承認
  const handleApprove = async (loanId) => {
    const loan = loans.find(l => l.id === loanId);
    const newLoans = loans.map(l => l.id === loanId ? { ...l, status: "承認済" } : l);
    const newItems = items.map(i => i.id === loan?.itemId ? { ...i, status: "貸出中" } : i);
    await save(newItems, newLoans);
    showToast("申請を承認しました。");
  };

  // 貸出却下
  const handleRejectConfirm = async (reason) => {
    const loan = loans.find(l => l.id === rejectTarget);
    const newLoans = loans.map(l => l.id === rejectTarget ? { ...l, status: "却下", rejectedReason: reason } : l);
    const newItems = items.map(i => i.id === loan?.itemId ? { ...i, status: "利用可能" } : i);
    await save(newItems, newLoans);
    setRejectTarget(null);
    showToast("申請を却下しました。", "error");
  };

  // 返却申請送信（一般ユーザー）
  const handleSubmitReturn = async ({ loanId, itemId, itemName, userName, returnMessage, returnRequestedAt }) => {
    const newLoans = loans.map(l => l.id === loanId ? { ...l, returnStatus: "返却申請中", returnMessage, returnRequestedAt, userName } : l);
    const newItems = items.map(i => i.id === itemId ? { ...i, status: "返却申請中" } : i);
    await save(newItems, newLoans);
    sendDiscordNotification("return", { itemName, userName, returnMessage, returnRequestedAt });
    showToast("返却申請を送信しました。管理者の承認をお待ちください。");
  };

  // 返却承認（管理者）
  const handleApproveReturn = async (loanId) => {
    const loan = loans.find(l => l.id === loanId);
    const newLoans = loans.map(l => l.id === loanId ? { ...l, returned: true, returnStatus: "返却承認済" } : l);
    const newItems = items.map(i => i.id === loan?.itemId ? { ...i, status: "利用可能" } : i);
    await save(newItems, newLoans);
    showToast("返却を承認しました。");
  };

  // 返却申請却下（管理者）
  const handleRejectReturn = async (loanId) => {
    const loan = loans.find(l => l.id === loanId);
    const newLoans = loans.map(l => l.id === loanId ? { ...l, returnStatus: null } : l);
    const newItems = items.map(i => i.id === loan?.itemId ? { ...i, status: "貸出中" } : i);
    await save(newItems, newLoans);
    setRejectReturnTarget(null);
    showToast("返却申請を却下しました。", "error");
  };

  // 状態変更（管理者）
  const handleChangeStatus = async (itemId, newStatus) => {
    const prevItem = items.find(i => i.id === itemId);
    const newItems = items.map(i => i.id === itemId ? { ...i, status: newStatus } : i);
    let newLoans = [...loans];
    if (newStatus === "利用可能" && (prevItem?.status === "貸出中" || prevItem?.status === "申請中" || prevItem?.status === "返却申請中")) {
      newLoans = loans.map(l => {
        if (l.itemId === itemId && !l.returned) {
          if (l.status === "承認済") return { ...l, returned: true, returnStatus: null };
          if (l.status === "審査中") return { ...l, returned: true, status: "却下", rejectedReason: "管理者が状態を変更したため" };
        }
        return l;
      });
      showToast(`${itemId} の状態を「利用可能」に変更しました（返却完了）。`);
    } else {
      showToast(`${itemId} の状態を「${newStatus}」に変更しました。`);
    }
    await save(newItems, newLoans);
  };

  const handleChangeNote = async (itemId, note) => {
    const newItems = items.map(i => i.id === itemId ? { ...i, note } : i);
    await save(newItems, loans);
    showToast("備考を更新しました。");
  };

  const handleChangeName = async (itemId, name) => {
    const newItems = items.map(i => i.id === itemId ? { ...i, name } : i);
    await save(newItems, loans);
    showToast("名称を更新しました。");
  };

  const handleDeleteItem = async (itemId) => {
    const newItems = items.filter(i => i.id !== itemId);
    await save(newItems, loans);
    showToast("物品を削除しました。", "error");
  };

  const handleToggleHidden = async (itemId) => {
    const newItems = items.map(i => i.id === itemId ? { ...i, hidden: !i.hidden } : i);
    await save(newItems, loans);
    const item = items.find(i => i.id === itemId);
    showToast(`${itemId} を${item?.hidden ? "表示" : "非表示"}にしました。`);
  };

  const handleAddItem = async (item) => {
    const newItems = [...items, item];
    await save(newItems, loans);
    showToast(`${item.name} を追加しました。`);
  };

  const handleImportXlsx = async (rows) => {
    const newItems = rows.map(r => ({ id: String(r["ID"] || r["id"] || "").trim(), name: String(r["名称"] || r["name"] || "").trim(), category: String(r["種別"] || r["category"] || "").trim(), note: String(r["備考"] || r["note"] || "").trim(), status: "利用可能", hidden: false })).filter(i => i.id && i.name);
    const deduped = [...items, ...newItems.filter(n => !items.find(ex => ex.id === n.id))];
    await save(deduped, loans);
    showToast(`${newItems.length} 件の物品をインポートしました。`);
  };

  const handleSaveNotice = async (text) => {
    await save(items, loans, text);
    setNotice(text);
    setShowNoticeModal(false);
    showToast("お知らせを更新しました。");
  };

  return (
    <div style={s.app}>
      <header style={s.header}>
        <div>
          <h1 style={s.headerTitle}>⚔️ 武道具管理システム</h1>
          <div style={s.headerSub}>Bugu Management System</div>
        </div>
        <div style={s.row}>
          {user.role === "管理者" && <button style={{ ...s.btn("gold", "sm") }} onClick={() => setShowNoticeModal(true)}>📢 お知らせ編集</button>}
          <span style={{ color: "rgba(255,255,255,0.8)", fontSize: 13 }}>{user.name}</span>
          <span style={s.badge(user.role)}>{user.role}</span>
          <button style={{ ...s.btn("outline", "sm"), color: "rgba(255,255,255,0.7)", borderColor: "rgba(255,255,255,0.3)" }} onClick={() => { setUser(null); setTab("物品一覧"); }}>ログアウト</button>
        </div>
      </header>
      <nav style={s.nav}>
        {tabs.map(t => (
          <button key={t} style={{ ...s.navBtn(tab === t), position: "relative" }} onClick={() => setTab(t)}>
            {t}
            {t === "申請管理" && pendingCount > 0 && <span style={{ position: "absolute", top: 6, right: 4, background: C.red, color: "#fff", borderRadius: 8, fontSize: 10, fontWeight: 700, padding: "1px 5px", lineHeight: 1.4 }}>{pendingCount}</span>}
          </button>
        ))}
      </nav>
      <main style={s.main}>
        {toast && <div style={{ ...s.alert(toast.type), position: "fixed", top: 80, right: 20, zIndex: 2000, minWidth: 280, boxShadow: "0 4px 16px rgba(0,0,0,0.15)" }}>{toast.msg}</div>}
        {tab === "ダッシュボード" && <Dashboard items={items} loans={loans} />}
        {tab === "物品一覧" && <ItemsTab items={items} loans={loans} user={user} onRequestLoan={setLoanTarget} onChangeStatus={handleChangeStatus} onChangeNote={handleChangeNote} onChangeName={handleChangeName} onDeleteItem={handleDeleteItem} onToggleHidden={handleToggleHidden} onAddItem={handleAddItem} onImportXlsx={handleImportXlsx} onRequestReturn={(item, loan) => setReturnTarget({ item, loan })} />}
        {tab === "申請管理" && user.role === "管理者" && <RequestsTab loans={loans} onApprove={handleApprove} onReject={setRejectTarget} onApproveReturn={handleApproveReturn} onRejectReturn={(id) => { setRejectReturnTarget(id); handleRejectReturn(id); }} user={user} />}
        {tab === "マイ申請" && user.role !== "管理者" && <MyRequestsTab loans={loans} user={user} />}
      </main>
      {loanTarget && <LoanRequestModal item={loanTarget} user={user} onSubmit={handleSubmitLoan} onClose={() => setLoanTarget(null)} />}
      {returnTarget && <ReturnRequestModal item={returnTarget.item} loan={returnTarget.loan} user={user} onSubmit={handleSubmitReturn} onClose={() => setReturnTarget(null)} />}
      {rejectTarget && <RejectModal onConfirm={handleRejectConfirm} onClose={() => setRejectTarget(null)} />}
      {showNoticeModal && <NoticeModal current={notice} onSave={handleSaveNotice} onClose={() => setShowNoticeModal(false)} />}
    </div>
  );
}
