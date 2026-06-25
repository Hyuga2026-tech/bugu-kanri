import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { initializeApp } from "firebase/app";
import { getFirestore, doc, setDoc, getDoc, onSnapshot } from "firebase/firestore";

// ── Firebase設定 ────────────────────────────────────────────
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

// ── Discord通知 ──────────────────────────────────────────────
const DISCORD_WEBHOOK_URL = "https://discordapp.com/api/webhooks/1519560284430274581/sP-JE_JuI_z0qhwSDmDxOd5e78wrue0djfxoc70aK-M5-FoT3O82lZobItAemXkt3q0X";

const sendDiscordNotification = async (loan) => {
  try {
    await fetch(DISCORD_WEBHOOK_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        embeds: [{
          title: "📋 新しい貸出申請が届きました",
          color: 0x2c3f7a,
          fields: [
            { name: "申請者", value: loan.userName, inline: true },
            { name: "借りるもの", value: loan.itemName, inline: true },
            { name: "使用目的", value: loan.purpose, inline: false },
            { name: "返却予定日", value: loan.dueDate, inline: true },
            { name: "申請日", value: loan.requestedAt, inline: true },
          ],
          footer: { text: "武道具管理システム" },
          timestamp: new Date().toISOString(),
        }]
      }),
    });
  } catch (e) {
    console.error("Discord通知エラー:", e);
  }
};

// ── 初期データ ──────────────────────────────────────────────
const INITIAL_ITEMS = [
  { id: "K1", name: "小太刀 K1", category: "小太刀", status: "利用可能", note: "" },
  { id: "K2", name: "小太刀 K2", category: "小太刀", status: "利用可能", note: "" },
  { id: "K3", name: "小太刀 K3", category: "小太刀", status: "利用可能", note: "" },
  { id: "K4", name: "小太刀 K4", category: "小太刀", status: "利用可能", note: "" },
  { id: "K5", name: "小太刀 K5", category: "小太刀", status: "利用可能", note: "" },
  { id: "C1", name: "長剣 C1", category: "長剣", status: "利用可能", note: "" },
  { id: "C2", name: "長剣 C2", category: "長剣", status: "利用可能", note: "" },
];

const ADMIN_PASSWORD = "Kumistral";
const today = () => new Date().toISOString().slice(0, 10);

const STATUS_COLORS = {
  "利用可能": { bg: "#e8f5e9", text: "#2e7d32", border: "#a5d6a7" },
  "申請中":   { bg: "#e8eaf6", text: "#283593", border: "#9fa8da" },
  "貸出中":   { bg: "#fff3e0", text: "#e65100", border: "#ffb74d" },
  "故障":     { bg: "#fce4ec", text: "#b71c1c", border: "#ef9a9a" },
};

const C = {
  navy: "#1a2744", indigo: "#2c3f7a", gold: "#c9a84c", cream: "#f8f6f0",
  white: "#ffffff", gray50: "#f9fafb", gray100: "#f3f4f6", gray300: "#d1d5db",
  gray500: "#6b7280", gray700: "#374151", gray900: "#111827", red: "#dc2626", green: "#16a34a",
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
    background: variant === "primary" ? C.indigo : variant === "gold" ? C.gold : variant === "danger" ? C.red : variant === "success" ? C.green : "transparent",
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

// ── Firebase読み書き ─────────────────────────────────────────
const saveToFirebase = async (items, loans) => {
  await setDoc(doc(db, "appData", "main"), { items, loans });
};

// ── ログイン画面 ─────────────────────────────────────────────
function LoginScreen({ onLogin }) {
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
    <div style={{ minHeight: "100vh", background: `linear-gradient(160deg, ${C.navy} 0%, ${C.indigo} 60%, #3d5296 100%)`, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ background: C.white, borderRadius: 16, padding: "40px 36px", width: 360, boxShadow: "0 12px 48px rgba(0,0,0,0.4)" }}>
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
        <button style={{ ...s.btn("primary"), width: "100%", padding: "11px", fontSize: 14, marginTop: 8, background: C.navy }} onClick={handle}>
          ログイン →
        </button>
      </div>
    </div>
  );
}

// ── 物品一覧タブ ─────────────────────────────────────────────
function ItemsTab({ items, loans, user, onRequestLoan, onChangeStatus, onAddItem, onImportXlsx }) {
  const [statusFilter, setStatusFilter] = useState("すべて");
  const [showAdd, setShowAdd] = useState(false);
  const [newItem, setNewItem] = useState({ id: "", name: "", category: "" });
  const [editStatus, setEditStatus] = useState(null);
  const [addErr, setAddErr] = useState("");
  const [sortKey, setSortKey] = useState("id");
  const [sortDir, setSortDir] = useState("asc");
  const fileRef = useRef();

  const handleSort = (key) => {
    if (sortKey === key) {
      setSortDir(d => d === "asc" ? "desc" : "asc");
    } else {
      setSortKey(key);
      setSortDir("asc");
    }
  };

  const STATUS_ORDER = { "利用可能": 0, "申請中": 1, "貸出中": 2, "故障": 3 };

  const getDueDate = (itemId) => {
    const loan = loans.find(l => l.itemId === itemId && l.status === "承認済" && !l.returned);
    return loan?.dueDate || "";
  };

  const filtered = items
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

  const handleAdd = () => {
    if (!newItem.id.trim() || !newItem.name.trim() || !newItem.category.trim()) { setAddErr("すべて入力してください"); return; }
    if (items.find(i => i.id === newItem.id.trim())) { setAddErr("IDが重複しています"); return; }
    onAddItem({ ...newItem, id: newItem.id.trim(), name: newItem.name.trim(), category: newItem.category.trim(), status: "利用可能", note: "" });
    setNewItem({ id: "", name: "", category: "" }); setShowAdd(false); setAddErr("");
  };

  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const wb = XLSX.read(ev.target.result, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws);
      onImportXlsx(rows);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const adminStatusOptions = (currentStatus) => {
    if (currentStatus === "申請中") return ["申請中", "利用可能", "故障"];
    if (currentStatus === "貸出中") return ["貸出中", "利用可能", "故障"];
    return ["利用可能", "故障"];
  };

  return (
    <div>
      <div style={{ ...s.row, justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
        <div style={s.row}>
          {["すべて", "利用可能", "申請中", "貸出中", "故障"].map(f => (
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
                {[
                  { label: "ID", key: "id" },
                  { label: "名称", key: null },
                  { label: "種別", key: "category" },
                  { label: "状態", key: "status" },
                  { label: "関連情報（返却予定日）", key: "dueDate" },
                  { label: "備考", key: null },
                  { label: "操作", key: null },
                ].map(h => (
                  <th key={h.label} style={{ ...s.th, cursor: h.key ? "pointer" : "default", userSelect: "none" }}
                    onClick={() => h.key && handleSort(h.key)}>
                    {h.label}
                    {h.key && sortKey === h.key && (
                      <span style={{ marginLeft: 4 }}>{sortDir === "asc" ? "▲" : "▼"}</span>
                    )}
                    {h.key && sortKey !== h.key && (
                      <span style={{ marginLeft: 4, color: C.gray300 }}>▲▼</span>
                    )}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr><td colSpan={7} style={s.empty}>物品がありません</td></tr>
              ) : filtered.map(item => {
                const pending = pendingLoan(item.id);
                const active = activeLoan(item.id);
                const relatedLoan = pending || active;
                return (
                  <tr key={item.id} style={{ background: item.status === "故障" ? "#fff5f5" : item.status === "申請中" ? "#f5f6ff" : "white" }}>
                    <td style={s.td}><span style={{ fontFamily: "monospace", fontWeight: 700, color: C.indigo }}>{item.id}</span></td>
                    <td style={s.td}>{item.name}</td>
                    <td style={s.td}><span style={s.tag(C.indigo)}>{item.category}</span></td>
                    <td style={s.td}>
                      {user.role === "管理者" && editStatus?.id === item.id ? (
                        <div style={s.row}>
                          <select style={{ ...s.select, width: "auto" }} value={editStatus.status} onChange={e => setEditStatus({ ...editStatus, status: e.target.value })}>
                            {adminStatusOptions(item.status).map(st => <option key={st}>{st}</option>)}
                          </select>
                          <button style={s.btn("success", "sm")} onClick={() => { onChangeStatus(item.id, editStatus.status); setEditStatus(null); }}>✓</button>
                          <button style={s.btn("outline", "sm")} onClick={() => setEditStatus(null)}>✕</button>
                        </div>
                      ) : (
                        <div style={s.row}>
                          <span style={s.statusBadge(item.status)}>{item.status}</span>
                          {user.role === "管理者" && <button style={{ ...s.btn("outline", "sm"), fontSize: 11 }} onClick={() => setEditStatus({ id: item.id, status: item.status })}>変更</button>}
                        </div>
                      )}
                    </td>
                    <td style={s.td}>
                      {relatedLoan ? (
                        <div style={{ fontSize: 12 }}>
                          <div style={{ fontWeight: 600 }}>{relatedLoan.userName}</div>
                          {relatedLoan.status === "審査中"
                            ? <div style={{ color: "#283593" }}>申請中（{relatedLoan.requestedAt}）</div>
                            : <div style={{ color: C.gray500 }}>返却予定: {relatedLoan.dueDate}</div>}
                        </div>
                      ) : <span style={{ color: C.gray300 }}>—</span>}
                    </td>
                    <td style={{ ...s.td, color: C.gray500, fontSize: 12 }}>{item.note || "—"}</td>
                    <td style={s.td}>
                      {user.role !== "管理者" && item.status === "利用可能" && (
                        <button style={s.btn("primary", "sm")} onClick={() => onRequestLoan(item)}>申請</button>
                      )}
                      {user.role !== "管理者" && item.status === "申請中" && pending?.userName === user.name && (
                        <span style={{ fontSize: 12, color: "#283593", fontWeight: 600 }}>審査待ち</span>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

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
          <div style={{ marginBottom: 14 }}>
            <label style={s.label}>借りるもの</label>
            <input style={{ ...s.input, background: C.gray50 }} value={item.name} readOnly />
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={s.label}>使用目的 <span style={{ color: C.red }}>*</span></label>
            <textarea style={{ ...s.input, minHeight: 80, resize: "vertical" }} placeholder="例: 演武練習のため" value={purpose} onChange={e => setPurpose(e.target.value)} />
          </div>
          <div style={{ marginBottom: 22 }}>
            <label style={s.label}>返却予定日 <span style={{ color: C.red }}>*</span></label>
            <input type="date" style={s.input} min={today()} value={dueDate} onChange={e => setDueDate(e.target.value)} />
          </div>
          <div style={s.row}>
            <button style={s.btn("primary")} onClick={handle}>申請を送信</button>
            <button style={s.btn("outline")} onClick={onClose}>キャンセル</button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ── 申請管理タブ ──────────────────────────────────────────────
function RequestsTab({ loans, onApprove, onReject, onReturn, user }) {
  const [filter, setFilter] = useState("審査中");
  const filtered = loans.filter(l => filter === "すべて" || l.status === filter);
  const statusColor = { "審査中": C.gold, "承認済": C.green, "却下": C.red };

  return (
    <div>
      <div style={{ ...s.row, marginBottom: 16, flexWrap: "wrap", gap: 8 }}>
        {["すべて", "審査中", "承認済", "却下"].map(f => (
          <button key={f} style={{ ...s.btn(filter === f ? "primary" : "outline", "sm") }} onClick={() => setFilter(f)}>
            {f}{f !== "すべて" && <span style={{ marginLeft: 6, background: (statusColor[f] || C.gray500) + "33", color: statusColor[f] || C.gray500, padding: "1px 7px", borderRadius: 8, fontSize: 11, fontWeight: 700 }}>{loans.filter(l => l.status === f).length}</span>}
          </button>
        ))}
      </div>
      {filtered.length === 0 ? (
        <div style={s.card}><div style={s.empty}>該当する申請はありません</div></div>
      ) : filtered.map(loan => (
        <div key={loan.id} style={{ ...s.card, borderLeft: `4px solid ${statusColor[loan.status] || C.gray300}` }}>
          <div style={{ padding: "14px 20px", display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12 }}>
            <div style={{ flex: 1 }}>
              <div style={{ ...s.row, gap: 8, marginBottom: 6 }}>
                <span style={{ fontSize: 15, fontWeight: 700, color: C.navy }}>{loan.itemName}</span>
                <span style={{ background: (statusColor[loan.status] || C.gray500) + "22", color: statusColor[loan.status] || C.gray500, padding: "2px 10px", borderRadius: 10, fontSize: 12, fontWeight: 700 }}>{loan.status}</span>
                {loan.returned && <span style={s.tag(C.green)}>返却済</span>}
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: "4px 16px", fontSize: 13, color: C.gray700 }}>
                <div><span style={{ color: C.gray500 }}>申請者: </span>{loan.userName}</div>
                <div><span style={{ color: C.gray500 }}>申請日: </span>{loan.requestedAt}</div>
                <div><span style={{ color: C.gray500 }}>返却予定: </span><b style={{ color: loan.dueDate < today() && !loan.returned ? C.red : "inherit" }}>{loan.dueDate}</b>{loan.dueDate < today() && !loan.returned && loan.status === "承認済" && <span style={{ color: C.red }}> ⚠ 期限超過</span>}</div>
              </div>
              <div style={{ marginTop: 6, fontSize: 13 }}><span style={{ color: C.gray500 }}>使用目的: </span>{loan.purpose}</div>
              {loan.rejectedReason && <div style={{ marginTop: 6, fontSize: 12, color: C.red }}>却下理由: {loan.rejectedReason}</div>}
            </div>
            <div style={{ ...s.row, gap: 8, flexShrink: 0 }}>
              {user.role === "管理者" && loan.status === "審査中" && (
                <><button style={s.btn("success", "sm")} onClick={() => onApprove(loan.id)}>✓ 承認</button><button style={s.btn("danger", "sm")} onClick={() => onReject(loan.id)}>✕ 却下</button></>
              )}
              {user.role === "管理者" && loan.status === "承認済" && !loan.returned && (
                <button style={s.btn("gold", "sm")} onClick={() => onReturn(loan.id)}>返却確認</button>
              )}
            </div>
          </div>
        </div>
      ))}
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
          <div style={s.row}>
            <button style={s.btn("danger")} onClick={() => onConfirm(reason)}>却下する</button>
            <button style={s.btn("outline")} onClick={onClose}>キャンセル</button>
          </div>
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
            <div style={{ ...s.row, marginBottom: 6 }}>
              <span style={{ fontSize: 15, fontWeight: 700, color: C.navy }}>{loan.itemName}</span>
              <span style={{ background: (statusColor[loan.status] || C.gray500) + "22", color: statusColor[loan.status] || C.gray500, padding: "2px 10px", borderRadius: 10, fontSize: 12, fontWeight: 700 }}>{loan.status}</span>
              {loan.returned && <span style={s.tag(C.green)}>返却済</span>}
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
    total: items.length,
    available: items.filter(i => i.status === "利用可能").length,
    applying: items.filter(i => i.status === "申請中").length,
    lent: items.filter(i => i.status === "貸出中").length,
    broken: items.filter(i => i.status === "故障").length,
    overdue: loans.filter(l => l.status === "承認済" && !l.returned && l.dueDate < today()).length,
  };
  const stats = [
    { label: "総物品数", value: counts.total, color: C.navy, icon: "⚔️" },
    { label: "利用可能", value: counts.available, color: C.green, icon: "✅" },
    { label: "申請中", value: counts.applying, color: "#283593", icon: "📝" },
    { label: "貸出中", value: counts.lent, color: "#e65100", icon: "📤" },
    { label: "故障", value: counts.broken, color: C.red, icon: "⚠️" },
    { label: "期限超過", value: counts.overdue, color: C.red, icon: "🔴" },
  ];
  return (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12 }}>
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

// ── メインアプリ ─────────────────────────────────────────────
export default function App() {
  const [user, setUser] = useState(null);
  const [items, setItems] = useState([]);
  const [loans, setLoans] = useState([]);
  const [loading, setLoading] = useState(true);
  const [tab, setTab] = useState("物品一覧");
  const [loanTarget, setLoanTarget] = useState(null);
  const [rejectTarget, setRejectTarget] = useState(null);
  const [toast, setToast] = useState(null);

  // Firebaseからリアルタイムでデータ取得
  useEffect(() => {
    const unsub = onSnapshot(doc(db, "appData", "main"), (snap) => {
      if (snap.exists()) {
        const data = snap.data();
        setItems(data.items || []);
        setLoans(data.loans || []);
      } else {
        // 初回のみ初期データを書き込む
        saveToFirebase(INITIAL_ITEMS, []);
        setItems(INITIAL_ITEMS);
        setLoans([]);
      }
      setLoading(false);
    });
    return () => unsub();
  }, []);

  const showToast = (msg, type = "success") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3500);
  };

  if (loading) return (
    <div style={{ minHeight: "100vh", background: C.navy, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ color: C.gold, fontSize: 18, fontWeight: 700 }}>⚔️ 読み込み中...</div>
    </div>
  );

  if (!user) return <LoginScreen onLogin={setUser} />;

  const tabs = user.role === "管理者"
    ? ["物品一覧", "申請管理", "ダッシュボード"]
    : ["物品一覧", "マイ申請"];

  const handleSubmitLoan = async (loan) => {
    const newLoans = [loan, ...loans];
    const newItems = items.map(i => i.id === loan.itemId ? { ...i, status: "申請中" } : i);
    await saveToFirebase(newItems, newLoans);
    sendDiscordNotification(loan);
    showToast("申請を送信しました。管理者の承認をお待ちください。");
    setTab("マイ申請");
  };

  const handleApprove = async (loanId) => {
    const loan = loans.find(l => l.id === loanId);
    const newLoans = loans.map(l => l.id === loanId ? { ...l, status: "承認済" } : l);
    const newItems = items.map(i => i.id === loan?.itemId ? { ...i, status: "貸出中" } : i);
    await saveToFirebase(newItems, newLoans);
    showToast("申請を承認しました。");
  };

  const handleRejectConfirm = async (reason) => {
    const loan = loans.find(l => l.id === rejectTarget);
    const newLoans = loans.map(l => l.id === rejectTarget ? { ...l, status: "却下", rejectedReason: reason } : l);
    const newItems = items.map(i => i.id === loan?.itemId ? { ...i, status: "利用可能" } : i);
    await saveToFirebase(newItems, newLoans);
    setRejectTarget(null);
    showToast("申請を却下しました。", "error");
  };

  const handleReturn = async (loanId) => {
    const loan = loans.find(l => l.id === loanId);
    if (!loan) { showToast("申請データが見つかりません。", "error"); return; }
    const newLoans = loans.map(l => l.id === loanId ? { ...l, returned: true } : l);
    const newItems = items.map(i => i.id === loan.itemId ? { ...i, status: "利用可能" } : i);
    try {
      await setDoc(doc(db, "appData", "main"), { items: newItems, loans: newLoans });
      showToast("返却を確認しました。");
    } catch (e) {
      console.error(e);
      showToast("保存に失敗しました。もう一度お試しください。", "error");
    }
  };

  const handleChangeStatus = async (itemId, newStatus) => {
    const prevItem = items.find(i => i.id === itemId);
    const newItems = items.map(i => i.id === itemId ? { ...i, status: newStatus } : i);
    let newLoans = [...loans];

    if (newStatus === "利用可能" && (prevItem?.status === "貸出中" || prevItem?.status === "申請中")) {
      newLoans = loans.map(l => {
        if (l.itemId === itemId && !l.returned) {
          if (l.status === "承認済") {
            return { ...l, returned: true };
          }
          if (l.status === "審査中") {
            return { ...l, returned: true, status: "却下", rejectedReason: "管理者が状態を変更したため" };
          }
        }
        return l;
      });
      try {
        await setDoc(doc(db, "appData", "main"), { items: newItems, loans: newLoans });
        showToast(`${itemId} の状態を「利用可能」に変更しました（返却完了）。`);
      } catch (e) {
        console.error(e);
        showToast("保存に失敗しました。", "error");
      }
    } else {
      try {
        await setDoc(doc(db, "appData", "main"), { items: newItems, loans: newLoans });
        showToast(`${itemId} の状態を「${newStatus}」に変更しました。`);
      } catch (e) {
        console.error(e);
        showToast("保存に失敗しました。", "error");
      }
    }
  };

  const handleAddItem = async (item) => {
    const newItems = [...items, item];
    await saveToFirebase(newItems, loans);
    showToast(`${item.name} を追加しました。`);
  };

  const handleImportXlsx = async (rows) => {
    const newItems = rows.map(r => ({
      id: String(r["ID"] || r["id"] || "").trim(),
      name: String(r["名称"] || r["name"] || "").trim(),
      category: String(r["種別"] || r["category"] || "").trim(),
      note: String(r["備考"] || r["note"] || "").trim(),
      status: "利用可能",
    })).filter(i => i.id && i.name);
    const deduped = [...items, ...newItems.filter(n => !items.find(ex => ex.id === n.id))];
    await saveToFirebase(deduped, loans);
    showToast(`${newItems.length} 件の物品をインポートしました。`);
  };

  const pendingCount = loans.filter(l => l.status === "審査中").length;

  return (
    <div style={s.app}>
      <header style={s.header}>
        <div>
          <h1 style={s.headerTitle}>⚔️ 武道具管理システム</h1>
          <div style={s.headerSub}>Bugu Management System</div>
        </div>
        <div style={s.row}>
          <span style={{ color: "rgba(255,255,255,0.8)", fontSize: 13 }}>{user.name}</span>
          <span style={s.badge(user.role)}>{user.role}</span>
          <button style={{ ...s.btn("outline", "sm"), color: "rgba(255,255,255,0.7)", borderColor: "rgba(255,255,255,0.3)" }} onClick={() => { setUser(null); setTab("物品一覧"); }}>ログアウト</button>
        </div>
      </header>

      <nav style={s.nav}>
        {tabs.map(t => (
          <button key={t} style={{ ...s.navBtn(tab === t), position: "relative" }} onClick={() => setTab(t)}>
            {t}
            {t === "申請管理" && pendingCount > 0 && (
              <span style={{ position: "absolute", top: 6, right: 4, background: C.red, color: "#fff", borderRadius: 8, fontSize: 10, fontWeight: 700, padding: "1px 5px", lineHeight: 1.4 }}>{pendingCount}</span>
            )}
          </button>
        ))}
      </nav>

      <main style={s.main}>
        {toast && (
          <div style={{ ...s.alert(toast.type), position: "fixed", top: 80, right: 20, zIndex: 2000, minWidth: 280, boxShadow: "0 4px 16px rgba(0,0,0,0.15)" }}>
            {toast.msg}
          </div>
        )}
        {tab === "ダッシュボード" && <Dashboard items={items} loans={loans} />}
        {tab === "物品一覧" && <ItemsTab items={items} loans={loans} user={user} onRequestLoan={setLoanTarget} onChangeStatus={handleChangeStatus} onAddItem={handleAddItem} onImportXlsx={handleImportXlsx} />}
        {tab === "申請管理" && user.role === "管理者" && <RequestsTab loans={loans} onApprove={handleApprove} onReject={setRejectTarget} onReturn={handleReturn} user={user} />}
        {tab === "マイ申請" && user.role !== "管理者" && <MyRequestsTab loans={loans} user={user} />}
      </main>

      {loanTarget && <LoanRequestModal item={loanTarget} user={user} onSubmit={handleSubmitLoan} onClose={() => setLoanTarget(null)} />}
      {rejectTarget && <RejectModal onConfirm={handleRejectConfirm} onClose={() => setRejectTarget(null)} />}
    </div>
  );
}
