const { useEffect, useMemo, useState } = React;

const STORAGE_KEY = 'umkm_ledger_pro_state_v2';
const LEGACY_TX_KEY = 'umkm_transactions';
const LEGACY_INV_KEY = 'umkm_inventory';
const BEP_TARGET = 5000000;

function pad(n) {
  return String(n).padStart(2, '0');
}

function todayKey(date = new Date()) {
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

function nowTime() {
  const d = new Date();
  return `${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function formatCurrency(value) {
  const num = Number(value || 0);
  return new Intl.NumberFormat('id-ID', {
    style: 'currency',
    currency: 'IDR',
    maximumFractionDigits: 0,
  }).format(num).replace('Rp', 'Rp ');
}

function formatDate(dateStr) {
  if (!dateStr) return '-';
  const d = new Date(`${dateStr}T00:00:00`);
  return new Intl.DateTimeFormat('id-ID', {
    day: '2-digit',
    month: 'short',
    year: 'numeric',
  }).format(d);
}

function formatDateTime(dateStr, timeStr) {
  return `${formatDate(dateStr)} • ${timeStr || '--:--'}`;
}

function safeParse(raw, fallback) {
  try {
    if (!raw) return fallback;
    return JSON.parse(raw);
  } catch {
    return fallback;
  }
}

function uid(prefix = 'id') {
  if (window.crypto && crypto.randomUUID) return `${prefix}_${crypto.randomUUID()}`;
  return `${prefix}_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

const defaultInventory = [
  { id: uid('item'), name: 'Nasi', cogs: 3000, price: 5000, qty: 50, minStock: 10, archived: false, createdAt: new Date().toISOString() },
  { id: uid('item'), name: 'Tahu', cogs: 1000, price: 2000, qty: 100, minStock: 20, archived: false, createdAt: new Date().toISOString() },
  { id: uid('item'), name: 'Tempe', cogs: 1500, price: 3000, qty: 80, minStock: 15, archived: false, createdAt: new Date().toISOString() },
];

function migrateLegacyState() {
  const oldTx = safeParse(localStorage.getItem(LEGACY_TX_KEY), []);
  const oldInv = safeParse(localStorage.getItem(LEGACY_INV_KEY), {});

  const inventory = Object.keys(oldInv || {}).length
    ? Object.entries(oldInv).map(([name, details]) => ({
        id: uid('item'),
        name,
        cogs: Number(details?.cogs ?? 0),
        price: Number(details?.price ?? 0),
        qty: Number(details?.qty ?? 0),
        minStock: Number(details?.minStock ?? 0),
        archived: false,
        createdAt: new Date().toISOString(),
      }))
    : clone(defaultInventory);

  const byName = new Map(inventory.map((item) => [item.name, item.id]));
  const transactions = Array.isArray(oldTx)
    ? oldTx.map((tx) => ({
        id: tx.id || uid('tx'),
        type: tx.type === 'expense' ? 'purchase' : 'sale',
        itemId: byName.get(tx.item) || uid('item'),
        itemName: tx.item || 'Unknown Item',
        qty: Number(tx.qty ?? 0),
        unitPrice: Number(tx.qty ? (tx.amount / tx.qty) : (tx.amount || 0)),
        total: Number(tx.amount ?? 0),
        channel: tx.channel || 'N/A',
        notes: tx.notes || '',
        date: todayKey(),
        time: tx.timestamp || nowTime(),
        createdAt: new Date().toISOString(),
      }))
    : [];

  // Ensure any legacy transaction itemIds exist in inventory
  const itemIds = new Set(inventory.map((item) => item.id));
  transactions.forEach((tx) => {
    if (!itemIds.has(tx.itemId)) {
      const newItem = { id: tx.itemId, name: tx.itemName, cogs: 0, price: 0, qty: 0, minStock: 0, archived: true, createdAt: new Date().toISOString() };
      inventory.push(newItem);
      itemIds.add(tx.itemId);
    }
  });

  return { inventory, transactions };
}

function loadInitialState() {
  const raw = safeParse(localStorage.getItem(STORAGE_KEY), null);
  if (raw && Array.isArray(raw.inventory) && Array.isArray(raw.transactions)) {
    return {
      inventory: raw.inventory,
      transactions: raw.transactions,
    };
  }

  const legacyExists = localStorage.getItem(LEGACY_TX_KEY) || localStorage.getItem(LEGACY_INV_KEY);
  if (legacyExists) return migrateLegacyState();

  return {
    inventory: clone(defaultInventory),
    transactions: [],
  };
}

function sum(list, getter) {
  return list.reduce((acc, item) => acc + Number(getter(item) || 0), 0);
}

function normalizeTransactionDraft(draft, inventory) {
  const item = inventory.find((it) => it.id === draft.itemId);
  if (!item) throw new Error('Item tidak ditemukan.');

  const qty = Number(draft.qty);
  const unitPrice = Number(draft.unitPrice);
  if (!Number.isFinite(qty) || qty <= 0) throw new Error('Qty harus lebih dari 0.');
  if (!Number.isFinite(unitPrice) || unitPrice < 0) throw new Error('Harga unit tidak valid.');

  return {
    id: draft.id || uid('tx'),
    type: draft.type,
    itemId: item.id,
    itemName: item.name,
    qty,
    unitPrice,
    total: qty * unitPrice,
    channel: draft.channel || 'Cash',
    notes: (draft.notes || '').trim(),
    date: draft.date || todayKey(),
    time: draft.time || nowTime(),
    createdAt: draft.createdAt || new Date().toISOString(),
  };
}

function applyTransactionEffect(inventory, tx, direction = 1) {
  const next = clone(inventory);
  const item = next.find((it) => it.id === tx.itemId);
  if (!item) throw new Error(`Item ${tx.itemName} tidak ditemukan di inventory.`);

  const delta = tx.type === 'sale' ? -tx.qty : tx.qty;
  const effective = delta * direction;
  const newQty = item.qty + effective;

  if (newQty < 0) {
    throw new Error(`Stok ${item.name} tidak cukup.`);
  }

  item.qty = newQty;
  return next;
}

function exportWorkbook(filename, sheets) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(({ name, data, type = 'json' }) => {
    const ws = type === 'aoa' ? XLSX.utils.aoa_to_sheet(data) : XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, name);
  });
  XLSX.writeFile(wb, filename);
}

function App() {
  const [appState, setAppState] = useState(loadInitialState);
  const [activePage, setActivePage] = useState('dashboard');
  const [toast, setToast] = useState('');
  const [quick, setQuick] = useState({
    type: 'sale',
    itemId: '',
    qty: 1,
    unitPrice: 0,
    channel: 'Cash',
    date: todayKey(),
    notes: '',
  });
  const [txFilters, setTxFilters] = useState({ q: '', type: 'all', itemId: 'all', channel: 'all', from: '', to: '', sort: 'newest' });
  const [inventoryFilters, setInventoryFilters] = useState({ q: '', status: 'all', sort: 'name' });
  const [txModal, setTxModal] = useState(null);
  const [itemModal, setItemModal] = useState(null);

  const inventory = appState.inventory;
  const transactions = appState.transactions;
  const activeItems = useMemo(() => inventory.filter((item) => !item.archived), [inventory]);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(appState));
  }, [appState]);

  useEffect(() => {
    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.register('./sw.js').catch(() => {});
    }
  }, []);

  useEffect(() => {
    const firstActive = activeItems[0];
    if (!quick.itemId && firstActive) {
      setQuick((prev) => ({
        ...prev,
        itemId: firstActive.id,
        unitPrice: firstActive.price || firstActive.cogs || 0,
      }));
    }
  }, [activeItems, quick.itemId]);

  useEffect(() => {
    const item = activeItems.find((it) => it.id === quick.itemId);
    if (!item) return;
    const defaultPrice = quick.type === 'sale' ? item.price : item.cogs;
    setQuick((prev) => ({ ...prev, unitPrice: defaultPrice || prev.unitPrice }));
  }, [quick.type, quick.itemId, activeItems]);

  function notify(message) {
    setToast(message);
    window.clearTimeout(notify.timer);
    notify.timer = window.setTimeout(() => setToast(''), 2200);
  }

  function updateInventory(nextInventory) {
    setAppState((prev) => ({ ...prev, inventory: nextInventory }));
  }

  function updateTransactions(nextTransactions) {
    setAppState((prev) => ({ ...prev, transactions: nextTransactions }));
  }

  function addQuickTransaction(event) {
    event.preventDefault();
    try {
      const draft = normalizeTransactionDraft(quick, inventory);
      const item = activeItems.find((it) => it.id === draft.itemId);
      if (!item) throw new Error('Item tidak tersedia.');
      if (draft.type === 'sale' && item.qty < draft.qty) {
        throw new Error(`Stok ${item.name} hanya ${item.qty}.`);
      }

      const nextInventory = applyTransactionEffect(inventory, draft, 1);
      const nextTransactions = [draft, ...transactions];
      setAppState({ inventory: nextInventory, transactions: nextTransactions });
      setQuick((prev) => ({ ...prev, qty: 1, notes: '', date: todayKey(), unitPrice: draft.type === 'sale' ? (item.price || 0) : (item.cogs || 0) }));
      notify('Transaksi tersimpan.');
    } catch (err) {
      alert(err.message || 'Gagal menyimpan transaksi.');
    }
  }

  function openNewItem() {
    setItemModal({ mode: 'create', item: { name: '', cogs: 0, price: 0, qty: 0, minStock: 0, archived: false } });
  }

  function saveItemDraft(draft) {
    try {
      const name = draft.name.trim();
      if (!name) throw new Error('Nama item wajib diisi.');
      const qty = Number(draft.qty);
      const cogs = Number(draft.cogs);
      const price = Number(draft.price);
      const minStock = Number(draft.minStock);
      if (![qty, cogs, price, minStock].every(Number.isFinite)) throw new Error('Angka item tidak valid.');
      if (qty < 0 || cogs < 0 || price < 0 || minStock < 0) throw new Error('Nilai item tidak boleh negatif.');

      if (itemModal.mode === 'create') {
        const newItem = {
          id: uid('item'),
          name,
          cogs,
          price,
          qty,
          minStock,
          archived: false,
          createdAt: new Date().toISOString(),
        };
        updateInventory([newItem, ...inventory]);
        notify('Item baru ditambahkan.');
      } else {
        updateInventory(
          inventory.map((item) =>
            item.id === itemModal.item.id
              ? { ...item, name, cogs, price, qty, minStock, archived: !!draft.archived }
              : item
          )
        );
        notify('Item berhasil diperbarui.');
      }
      setItemModal(null);
    } catch (err) {
      alert(err.message || 'Gagal menyimpan item.');
    }
  }

  function toggleArchiveItem(itemId) {
    const target = inventory.find((item) => item.id === itemId);
    if (!target) return;
    const nextArchived = !target.archived;
    const confirmed = window.confirm(nextArchived ? `Arsipkan item ${target.name}?` : `Pulihkan item ${target.name}?`);
    if (!confirmed) return;
    updateInventory(inventory.map((item) => item.id === itemId ? { ...item, archived: nextArchived } : item));
    notify(nextArchived ? 'Item diarsipkan.' : 'Item dipulihkan.');
  }


  function editItem(item) {
    setItemModal({ mode: 'edit', item: clone(item) });
  }

  function openTransactionModal(tx) {
    setTxModal({ mode: 'edit', tx: clone(tx) });
  }

  function saveTransactionDraft(draft) {
    try {
      const currentTx = appState.transactions.find((tx) => tx.id === txModal.tx.id);
      if (!currentTx) throw new Error('Transaksi tidak ditemukan.');

      const invWithoutCurrent = applyTransactionEffect(inventory, currentTx, -1);
      const normalized = normalizeTransactionDraft(draft, invWithoutCurrent);
      const item = invWithoutCurrent.find((it) => it.id === normalized.itemId);
      if (normalized.type === 'sale' && item.qty < normalized.qty) {
        throw new Error(`Stok ${item.name} hanya ${item.qty}.`);
      }
      const nextInventory = applyTransactionEffect(invWithoutCurrent, normalized, 1);
      const nextTransactions = appState.transactions.map((tx) => tx.id === currentTx.id ? normalized : tx);
      setAppState({ inventory: nextInventory, transactions: nextTransactions });
      setTxModal(null);
      notify('Transaksi diperbarui.');
    } catch (err) {
      alert(err.message || 'Gagal memperbarui transaksi.');
    }
  }

  function deleteTransaction(txId) {
    const tx = transactions.find((t) => t.id === txId);
    if (!tx) return;
    if (!window.confirm('Hapus transaksi ini? Stok inventory akan disesuaikan kembali.')) return;
    try {
      const nextInventory = applyTransactionEffect(inventory, tx, -1);
      const nextTransactions = transactions.filter((t) => t.id !== txId);
      setAppState({ inventory: nextInventory, transactions: nextTransactions });
      notify('Transaksi dihapus.');
    } catch (err) {
      alert(err.message || 'Gagal menghapus transaksi.');
    }
  }

  function exportTransactions() {
    const rows = filteredTransactions.map((tx, index) => ({
      No: index + 1,
      Tanggal: tx.date,
      Waktu: tx.time,
      Tipe: tx.type === 'sale' ? 'Sale' : 'Purchase',
      Item: tx.itemName,
      Qty: tx.qty,
      'Unit Price': tx.unitPrice,
      Total: tx.total,
      Channel: tx.channel,
      Notes: tx.notes,
    }));
    exportWorkbook(`UMKM_Transactions_${todayKey()}.xlsx`, [
      { name: 'Transactions', data: rows },
      { name: 'Summary', type: 'aoa', data: [
        ['UMKM Ledger Pro', '', ''],
        ['Exported At', `${formatDate(todayKey())} ${nowTime()}`, ''],
        [],
        ['Filter Search', txFilters.q || '(none)', ''],
        ['Date From', txFilters.from || '(all)', ''],
        ['Date To', txFilters.to || '(all)', ''],
        ['Type', txFilters.type, ''],
        ['Item', txFilters.itemId, ''],
        ['Rows', String(rows.length), ''],
      ]},
    ]);
    notify('Excel transaksi diekspor.');
  }

  function exportInventory() {
    const rows = inventory.map((item, index) => ({
      No: index + 1,
      Name: item.name,
      Qty: item.qty,
      COGS: item.cogs,
      Price: item.price,
      MinStock: item.minStock,
      Status: item.archived ? 'Archived' : item.qty <= item.minStock ? 'Low Stock' : 'Active',
      CreatedAt: item.createdAt,
    }));
    exportWorkbook(`UMKM_Inventory_${todayKey()}.xlsx`, [
      { name: 'Inventory', data: rows },
      { name: 'Active Items', data: rows.filter((r) => r.Status !== 'Archived') },
      { name: 'Archived', data: rows.filter((r) => r.Status === 'Archived') },
    ]);
    notify('Excel inventory diekspor.');
  }

  const today = todayKey();
  const todayTransactions = useMemo(() => transactions.filter((tx) => tx.date === today), [transactions, today]);
  const dailyRevenue = useMemo(() => sum(todayTransactions.filter((tx) => tx.type === 'sale'), (tx) => tx.total), [todayTransactions]);
  const dailyPurchaseValue = useMemo(() => sum(todayTransactions.filter((tx) => tx.type === 'purchase'), (tx) => tx.total), [todayTransactions]);
  const activeStockValue = useMemo(() => sum(activeItems, (item) => item.qty * item.cogs), [activeItems]);
  const lowStockCount = useMemo(() => activeItems.filter((item) => item.qty <= item.minStock).length, [activeItems]);
  const bepProgress = Math.min(100, (dailyRevenue / BEP_TARGET) * 100);

  const filteredTransactions = useMemo(() => {
    const q = txFilters.q.trim().toLowerCase();
    let list = [...transactions];
    if (txFilters.type !== 'all') list = list.filter((tx) => tx.type === txFilters.type);
    if (txFilters.itemId !== 'all') list = list.filter((tx) => tx.itemId === txFilters.itemId);
    if (txFilters.channel !== 'all') list = list.filter((tx) => tx.channel === txFilters.channel);
    if (txFilters.from) list = list.filter((tx) => tx.date >= txFilters.from);
    if (txFilters.to) list = list.filter((tx) => tx.date <= txFilters.to);
    if (q) {
      list = list.filter((tx) => {
        const haystack = [tx.itemName, tx.notes, tx.channel, tx.date, tx.time, tx.type, String(tx.qty), String(tx.total)].join(' ').toLowerCase();
        return haystack.includes(q);
      });
    }
    list.sort((a, b) => {
      const left = `${a.date} ${a.time}`;
      const right = `${b.date} ${b.time}`;
      return txFilters.sort === 'oldest' ? left.localeCompare(right) : right.localeCompare(left);
    });
    return list;
  }, [transactions, txFilters]);

  const filteredInventory = useMemo(() => {
    const q = inventoryFilters.q.trim().toLowerCase();
    let list = [...inventory];
    if (inventoryFilters.status === 'active') list = list.filter((item) => !item.archived);
    if (inventoryFilters.status === 'archived') list = list.filter((item) => item.archived);
    if (inventoryFilters.status === 'low') list = list.filter((item) => !item.archived && item.qty <= item.minStock);
    if (q) {
      list = list.filter((item) => {
        const haystack = [item.name, item.qty, item.cogs, item.price, item.minStock].join(' ').toLowerCase();
        return haystack.includes(q);
      });
    }
    list.sort((a, b) => {
      if (inventoryFilters.sort === 'qty') return a.qty - b.qty;
      return a.name.localeCompare(b.name, 'id');
    });
    return list;
  }, [inventory, inventoryFilters]);

  const channelOptions = useMemo(() => {
    const set = new Set(transactions.map((tx) => tx.channel).filter(Boolean));
    return ['all', ...set];
  }, [transactions]);

  const currentQuickItem = activeItems.find((item) => item.id === quick.itemId) || activeItems[0];

  return (
    <div className="app-shell">
      <div className="bg-blur bg-blur-a" />
      <div className="bg-blur bg-blur-b" />

      <header className="topbar card glass">
        <div className="brand-block">
          <div className="brand-mark">UL</div>
          <div>
            <div className="eyebrow">UMKM Financial Control</div>
            <h1>UMKM Ledger Pro</h1>
            <p>Dashboard harian, history transaksi, dan inventory manager dalam satu workspace landscape.</p>
          </div>
        </div>
        <div className="topbar-meta">
          <div className="pill success">Ready</div>
          <div className="date-stack">
            <span>Zona waktu: Asia/Jakarta</span>
            <strong>{formatDate(today)}</strong>
          </div>
        </div>
      </header>

      <nav className="tabbar card glass">
        <TabButton active={activePage === 'dashboard'} onClick={() => setActivePage('dashboard')}>Page 1 · Dashboard</TabButton>
        <TabButton active={activePage === 'transactions'} onClick={() => setActivePage('transactions')}>Page 2 · History</TabButton>
        <TabButton active={activePage === 'inventory'} onClick={() => setActivePage('inventory')}>Page 3 · Inventory</TabButton>
      </nav>

      {toast ? <div className="toast">{toast}</div> : null}

      {activePage === 'dashboard' && (
        <section className="page-stack">
          <div className="stats-grid">
            <StatCard title="Revenue Hari Ini" value={formatCurrency(dailyRevenue)} hint={`${todayTransactions.filter((tx) => tx.type === 'sale').length} transaksi penjualan hari ini`} />
            <StatCard title="Nilai Stok Aktif" value={formatCurrency(activeStockValue)} hint="Nilai modal yang tertahan di inventory aktif" />
            <StatCard title="Transaksi Hari Ini" value={String(todayTransactions.length)} hint={`${dailyPurchaseValue > 0 ? formatCurrency(dailyPurchaseValue) + ' pembelian' : 'Tanpa pembelian hari ini'}`} />
            <StatCard title="Low Stock Alert" value={String(lowStockCount)} hint="Item yang sudah menyentuh minimum stock" />
          </div>

          <div className="dashboard-grid">
            <Panel className="page-panel entry-panel" title="Quick Entry" subtitle="Tambah transaksi hari ini">
              <div className="section-head-row">
                <span className="chip">Auto update stok</span>
                <span className="chip soft">Landscape layout</span>
              </div>
              <form onSubmit={addQuickTransaction} className="form-grid">
                <Field label="Transaction Type">
                  <select value={quick.type} onChange={(e) => setQuick((prev) => ({ ...prev, type: e.target.value }))}>
                    <option value="sale">Sale / Income (stok berkurang)</option>
                    <option value="purchase">Purchase / Expense (stok bertambah)</option>
                  </select>
                </Field>
                <Field label="Item">
                  <select value={quick.itemId} onChange={(e) => setQuick((prev) => ({ ...prev, itemId: e.target.value }))}>
                    {activeItems.map((item) => <option key={item.id} value={item.id}>{item.name}</option>)}
                  </select>
                </Field>
                <div className="two-col">
                  <Field label="Qty">
                    <input type="number" min="1" value={quick.qty} onChange={(e) => setQuick((prev) => ({ ...prev, qty: e.target.value }))} />
                  </Field>
                  <Field label="Amount (IDR)">
                    <input type="number" min="0" value={quick.unitPrice} onChange={(e) => setQuick((prev) => ({ ...prev, unitPrice: e.target.value }))} />
                  </Field>
                </div>
                <Field label="Payment Channel">
                  <select value={quick.channel} onChange={(e) => setQuick((prev) => ({ ...prev, channel: e.target.value }))}>
                    <option>Cash</option>
                    <option>QRIS / Digital Pay</option>
                    <option>Transfer</option>
                    <option>Other</option>
                  </select>
                </Field>
                <Field label="Transaction Date">
                  <input type="date" value={quick.date} onChange={(e) => setQuick((prev) => ({ ...prev, date: e.target.value }))} />
                </Field>
                <Field label="Notes">
                  <textarea rows="4" placeholder="Optional notes..." value={quick.notes} onChange={(e) => setQuick((prev) => ({ ...prev, notes: e.target.value }))} />
                </Field>
                <div className="inline-summary">
                  <div>
                    <span className="muted">Selected item</span>
                    <strong>{currentQuickItem ? currentQuickItem.name : '-'}</strong>
                  </div>
                  <div>
                    <span className="muted">Stock after action</span>
                    <strong>{currentQuickItem ? (quick.type === 'sale' ? currentQuickItem.qty - Number(quick.qty || 0) : currentQuickItem.qty + Number(quick.qty || 0)) : '-'}</strong>
                  </div>
                  <div>
                    <span className="muted">Unit default</span>
                    <strong>{currentQuickItem ? formatCurrency(quick.type === 'sale' ? currentQuickItem.price : currentQuickItem.cogs) : '-'}</strong>
                  </div>
                </div>
                <button className="primary-btn" type="submit">Save Transaction</button>
              </form>
            </Panel>

            <Panel className="page-panel ledger-panel" title="Ledger history" subtitle="Hari ini saja">
              <div className="panel-toolbar">
                <span className="pill">{todayTransactions.length} records</span>
                <span className="pill ghost">{formatCurrency(dailyRevenue)} sales</span>
              </div>
              <div className="ledger-list">
                {todayTransactions.length === 0 ? (
                  <div className="empty-state">Belum ada transaksi untuk hari ini.</div>
                ) : todayTransactions.map((tx) => <LedgerCard key={tx.id} tx={tx} />)}
              </div>
            </Panel>
          </div>
        </section>
      )}

      {activePage === 'transactions' && (
        <section className="page-stack">
          <Panel title="Transaction History" subtitle="Full history dengan filter detail">
            <div className="toolbar-grid">
              <Field label="Search">
                <input placeholder="Cari item, catatan, channel, tanggal..." value={txFilters.q} onChange={(e) => setTxFilters((prev) => ({ ...prev, q: e.target.value }))} />
              </Field>
              <Field label="Type">
                <select value={txFilters.type} onChange={(e) => setTxFilters((prev) => ({ ...prev, type: e.target.value }))}>
                  <option value="all">All</option>
                  <option value="sale">Sale</option>
                  <option value="purchase">Purchase</option>
                </select>
              </Field>
              <Field label="Item">
                <select value={txFilters.itemId} onChange={(e) => setTxFilters((prev) => ({ ...prev, itemId: e.target.value }))}>
                  <option value="all">All items</option>
                  {inventory.map((item) => <option key={item.id} value={item.id}>{item.name}</option>)}
                </select>
              </Field>
              <Field label="Channel">
                <select value={txFilters.channel} onChange={(e) => setTxFilters((prev) => ({ ...prev, channel: e.target.value }))}>
                  {channelOptions.map((ch) => <option key={ch} value={ch}>{ch === 'all' ? 'All channels' : ch}</option>)}
                </select>
              </Field>
              <Field label="From">
                <input type="date" value={txFilters.from} onChange={(e) => setTxFilters((prev) => ({ ...prev, from: e.target.value }))} />
              </Field>
              <Field label="To">
                <input type="date" value={txFilters.to} onChange={(e) => setTxFilters((prev) => ({ ...prev, to: e.target.value }))} />
              </Field>
              <Field label="Sort">
                <select value={txFilters.sort} onChange={(e) => setTxFilters((prev) => ({ ...prev, sort: e.target.value }))}>
                  <option value="newest">Newest</option>
                  <option value="oldest">Oldest</option>
                </select>
              </Field>
              <div className="export-stack">
                <button className="secondary-btn" type="button" onClick={exportTransactions}>Export Excel</button>
                <button className="ghost-btn" type="button" onClick={() => setTxFilters({ q: '', type: 'all', itemId: 'all', channel: 'all', from: '', to: '', sort: 'newest' })}>Reset Filter</button>
              </div>
            </div>
          </Panel>

          <Panel title={`Transactions (${filteredTransactions.length})`} subtitle="Edit / delete tersedia di setiap row">
            <div className="table-wrap">
              <table className="data-table">
                <thead>
                  <tr>
                    <th>Date</th>
                    <th>Type</th>
                    <th>Item</th>
                    <th>Qty</th>
                    <th>Unit</th>
                    <th>Total</th>
                    <th>Channel</th>
                    <th>Notes</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredTransactions.length === 0 ? (
                    <tr><td colSpan="9"><div className="empty-state">Tidak ada transaksi yang cocok dengan filter.</div></td></tr>
                  ) : filteredTransactions.map((tx) => (
                    <tr key={tx.id}>
                      <td>{formatDateTime(tx.date, tx.time)}</td>
                      <td><span className={`pill ${tx.type === 'sale' ? 'success' : 'warning'}`}>{tx.type === 'sale' ? 'Sale' : 'Purchase'}</span></td>
                      <td>
                        <div className="stacked-cell">
                          <strong>{tx.itemName}</strong>
                          <span className="muted">ID: {tx.itemId.slice(0, 8)}</span>
                        </div>
                      </td>
                      <td>{tx.qty}</td>
                      <td>{formatCurrency(tx.unitPrice)}</td>
                      <td>{formatCurrency(tx.total)}</td>
                      <td>{tx.channel}</td>
                      <td className="ellipsis">{tx.notes || '-'}</td>
                      <td>
                        <div className="row-actions">
                          <button className="tiny-btn" onClick={() => openTransactionModal(tx)}>Edit</button>
                          <button className="tiny-btn danger" onClick={() => deleteTransaction(tx.id)}>Delete</button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Panel>
        </section>
      )}

      {activePage === 'inventory' && (
        <section className="page-stack">
          <div className="dashboard-grid">
            <Panel className="page-panel entry-panel" title="Inventory Editor" subtitle="Add / edit / delete item">
              <div className="panel-toolbar">
                <span className="chip">Data storage export ready</span>
                <button type="button" className="secondary-btn compact" onClick={openNewItem}>Add Item</button>
              </div>
              <form className="form-grid" onSubmit={(e) => { e.preventDefault(); if (itemModal?.mode === 'edit') saveItemDraft(itemModal.item); else setItemModal({ mode: 'create', item: { name: '', cogs: 0, price: 0, qty: 0, minStock: 0, archived: false } }); }}>
                <Field label="Selected Item">
                  <select value={itemModal?.item?.id || ''} onChange={(e) => {
                    const found = inventory.find((item) => item.id === e.target.value);
                    if (found) setItemModal({ mode: 'edit', item: clone(found) });
                  }}>
                    <option value="">Choose existing item</option>
                    {inventory.map((item) => <option key={item.id} value={item.id}>{item.name}{item.archived ? ' (archived)' : ''}</option>)}
                  </select>
                </Field>
                <Field label="Name">
                  <input value={itemModal?.item?.name ?? ''} onChange={(e) => setItemModal((prev) => prev ? { ...prev, item: { ...prev.item, name: e.target.value } } : prev)} placeholder="Item name" />
                </Field>
                <div className="two-col">
                  <Field label="Qty">
                    <input type="number" min="0" value={itemModal?.item?.qty ?? 0} onChange={(e) => setItemModal((prev) => prev ? { ...prev, item: { ...prev.item, qty: e.target.value } } : prev)} />
                  </Field>
                  <Field label="Min Stock">
                    <input type="number" min="0" value={itemModal?.item?.minStock ?? 0} onChange={(e) => setItemModal((prev) => prev ? { ...prev, item: { ...prev.item, minStock: e.target.value } } : prev)} />
                  </Field>
                </div>
                <div className="two-col">
                  <Field label="COGS">
                    <input type="number" min="0" value={itemModal?.item?.cogs ?? 0} onChange={(e) => setItemModal((prev) => prev ? { ...prev, item: { ...prev.item, cogs: e.target.value } } : prev)} />
                  </Field>
                  <Field label="Price">
                    <input type="number" min="0" value={itemModal?.item?.price ?? 0} onChange={(e) => setItemModal((prev) => prev ? { ...prev, item: { ...prev.item, price: e.target.value } } : prev)} />
                  </Field>
                </div>
                <label className="checkbox-row">
                  <input type="checkbox" checked={!!itemModal?.item?.archived} onChange={(e) => setItemModal((prev) => prev ? { ...prev, item: { ...prev.item, archived: e.target.checked } } : prev)} />
                  Archive item
                </label>
                <button className="primary-btn" type="button" onClick={() => itemModal ? saveItemDraft(itemModal.item) : null}>Save Item</button>
              </form>
            </Panel>

            <Panel className="page-panel ledger-panel" title="Inventory list" subtitle="Search / filter / export">
              <div className="toolbar-grid compact-grid">
                <Field label="Search">
                  <input placeholder="Cari nama / angka" value={inventoryFilters.q} onChange={(e) => setInventoryFilters((prev) => ({ ...prev, q: e.target.value }))} />
                </Field>
                <Field label="Status">
                  <select value={inventoryFilters.status} onChange={(e) => setInventoryFilters((prev) => ({ ...prev, status: e.target.value }))}>
                    <option value="all">All</option>
                    <option value="active">Active</option>
                    <option value="low">Low stock</option>
                    <option value="archived">Archived</option>
                  </select>
                </Field>
                <Field label="Sort">
                  <select value={inventoryFilters.sort} onChange={(e) => setInventoryFilters((prev) => ({ ...prev, sort: e.target.value }))}>
                    <option value="name">Name</option>
                    <option value="qty">Qty</option>
                  </select>
                </Field>
                <div className="export-stack">
                  <button className="secondary-btn" type="button" onClick={exportInventory}>Export Excel</button>
                  <button className="ghost-btn" type="button" onClick={() => setInventoryFilters({ q: '', status: 'all', sort: 'name' })}>Reset</button>
                </div>
              </div>

              <div className="inventory-list">
                {filteredInventory.length === 0 ? (
                  <div className="empty-state">Tidak ada item yang cocok.</div>
                ) : filteredInventory.map((item) => {
                  const status = item.archived ? 'Archived' : item.qty <= item.minStock ? 'Low Stock' : 'Active';
                  return (
                    <div key={item.id} className={`inventory-row ${item.archived ? 'archived' : ''}`}>
                      <div className="row-main">
                        <div className="stacked-cell">
                          <strong>{item.name}</strong>
                          <span className="muted">Qty {item.qty} · COGS {formatCurrency(item.cogs)} · Price {formatCurrency(item.price)}</span>
                        </div>
                        <div className="stacked-cell align-right">
                          <span className={`pill ${status === 'Active' ? 'success' : status === 'Low Stock' ? 'warning' : ''}`}>{status}</span>
                          <span className="muted">Min stock {item.minStock}</span>
                        </div>
                      </div>
                      <div className="row-actions">
                        <button className="tiny-btn" onClick={() => editItem(item)}>Edit</button>
                        <button className={`tiny-btn ${item.archived ? '' : 'danger'}`} onClick={() => toggleArchiveItem(item.id)}>{item.archived ? 'Restore' : 'Delete'}</button>
                      </div>
                    </div>
                  );
                })}
              </div>
            </Panel>
          </div>
        </section>
      )}

      {txModal ? (
        <Modal title="Edit transaction" onClose={() => setTxModal(null)}>
          <TransactionForm
            initial={txModal.tx}
            items={inventory}
            onCancel={() => setTxModal(null)}
            onSave={saveTransactionDraft}
          />
        </Modal>
      ) : null}

      {itemModal ? (
        <Modal title={itemModal.mode === 'create' ? 'Add item' : 'Edit item'} onClose={() => setItemModal(null)}>
          <ItemForm
            initial={itemModal.item}
            onCancel={() => setItemModal(null)}
            onSave={(next) => saveItemDraft(next)}
          />
        </Modal>
      ) : null}
    </div>
  );

}

function TabButton({ active, onClick, children }) {
  return <button className={`tab-button ${active ? 'active' : ''}`} onClick={onClick}>{children}</button>;
}

function StatCard({ title, value, hint }) {
  return (
    <div className="card stat-card glass">
      <span className="eyebrow">{title}</span>
      <strong className="stat-value">{value}</strong>
      <span className="muted">{hint}</span>
    </div>
  );
}

function Panel({ title, subtitle, children, className = '' }) {
  return (
    <section className={`card panel glass ${className}`}>
      <div className="panel-head">
        <div>
          <span className="eyebrow">{subtitle}</span>
          <h2>{title}</h2>
        </div>
      </div>
      {children}
    </section>
  );
}

function Field({ label, children }) {
  return (
    <label className="field">
      <span>{label}</span>
      {children}
    </label>
  );
}

function LedgerCard({ tx }) {
  return (
    <div className="ledger-card">
      <div className="ledger-left">
        <div className={`pill ${tx.type === 'sale' ? 'success' : 'warning'}`}>{tx.type === 'sale' ? 'Sale' : 'Purchase'}</div>
        <div className="stacked-cell">
          <strong>{tx.itemName} · {tx.qty} pcs</strong>
          <span className="muted">{formatDateTime(tx.date, tx.time)} · {tx.channel}</span>
          <span className="muted ellipsis">{tx.notes || 'No notes'}</span>
        </div>
      </div>
      <strong className={tx.type === 'sale' ? 'positive' : 'negative'}>{tx.type === 'sale' ? '+' : '-'}{formatCurrency(tx.total)}</strong>
    </div>
  );
}

function Modal({ title, onClose, children }) {
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal card glass" onClick={(e) => e.stopPropagation()}>
        <div className="modal-head">
          <div>
            <span className="eyebrow">Form</span>
            <h2>{title}</h2>
          </div>
          <button className="ghost-btn icon-btn" onClick={onClose}>×</button>
        </div>
        {children}
      </div>
    </div>
  );
}

function TransactionForm({ initial, items, onSave, onCancel }) {
  const firstItem = items.find((item) => !item.archived) || items[0];
  const [draft, setDraft] = useState(() => ({
    ...initial,
    itemId: initial.itemId || (firstItem?.id || ''),
  }));

  useEffect(() => {
    const item = items.find((it) => it.id === draft.itemId);
    if (!item) return;
    if (!draft.unitPrice || draft.unitPrice === 0) {
      setDraft((prev) => ({ ...prev, unitPrice: prev.type === 'sale' ? item.price : item.cogs }));
    }
  }, [draft.itemId, draft.type, items]);

  return (
    <form className="form-grid modal-form" onSubmit={(e) => { e.preventDefault(); onSave(draft); }}>
      <div className="two-col">
        <Field label="Type">
          <select value={draft.type} onChange={(e) => setDraft((prev) => ({ ...prev, type: e.target.value }))}>
            <option value="sale">Sale</option>
            <option value="purchase">Purchase</option>
          </select>
        </Field>
        <Field label="Date">
          <input type="date" value={draft.date} onChange={(e) => setDraft((prev) => ({ ...prev, date: e.target.value }))} />
        </Field>
      </div>
      <Field label="Item">
        <select value={draft.itemId} onChange={(e) => setDraft((prev) => ({ ...prev, itemId: e.target.value }))}>
          {items.map((item) => <option key={item.id} value={item.id}>{item.name}{item.archived ? ' (archived)' : ''}</option>)}
        </select>
      </Field>
      <div className="two-col">
        <Field label="Qty">
          <input type="number" min="1" value={draft.qty} onChange={(e) => setDraft((prev) => ({ ...prev, qty: e.target.value }))} />
        </Field>
        <Field label="Unit Price">
          <input type="number" min="0" value={draft.unitPrice} onChange={(e) => setDraft((prev) => ({ ...prev, unitPrice: e.target.value }))} />
        </Field>
      </div>
      <Field label="Payment Channel">
        <select value={draft.channel} onChange={(e) => setDraft((prev) => ({ ...prev, channel: e.target.value }))}>
          <option>Cash</option>
          <option>QRIS / Digital Pay</option>
          <option>Transfer</option>
          <option>Other</option>
        </select>
      </Field>
      <Field label="Notes">
        <textarea rows="4" value={draft.notes} onChange={(e) => setDraft((prev) => ({ ...prev, notes: e.target.value }))} />
      </Field>
      <div className="modal-actions">
        <button type="button" className="ghost-btn" onClick={onCancel}>Cancel</button>
        <button type="submit" className="primary-btn">Save Transaction</button>
      </div>
    </form>
  );
}

function ItemForm({ initial, onSave, onCancel }) {
  const [draft, setDraft] = useState(() => initial || { name: '', qty: 0, cogs: 0, price: 0, minStock: 0, archived: false });
  return (
    <form className="form-grid modal-form" onSubmit={(e) => { e.preventDefault(); onSave(draft); }}>
      <Field label="Name">
        <input value={draft.name} onChange={(e) => setDraft((prev) => ({ ...prev, name: e.target.value }))} placeholder="Item name" />
      </Field>
      <div className="two-col">
        <Field label="Qty">
          <input type="number" min="0" value={draft.qty} onChange={(e) => setDraft((prev) => ({ ...prev, qty: e.target.value }))} />
        </Field>
        <Field label="Min Stock">
          <input type="number" min="0" value={draft.minStock} onChange={(e) => setDraft((prev) => ({ ...prev, minStock: e.target.value }))} />
        </Field>
      </div>
      <div className="two-col">
        <Field label="COGS">
          <input type="number" min="0" value={draft.cogs} onChange={(e) => setDraft((prev) => ({ ...prev, cogs: e.target.value }))} />
        </Field>
        <Field label="Price">
          <input type="number" min="0" value={draft.price} onChange={(e) => setDraft((prev) => ({ ...prev, price: e.target.value }))} />
        </Field>
      </div>
      <label className="checkbox-row">
        <input type="checkbox" checked={!!draft.archived} onChange={(e) => setDraft((prev) => ({ ...prev, archived: e.target.checked }))} />
        Archived
      </label>
      <div className="modal-actions">
        <button type="button" className="ghost-btn" onClick={onCancel}>Cancel</button>
        <button type="submit" className="primary-btn">Save Item</button>
      </div>
    </form>
  );
}

ReactDOM.createRoot(document.getElementById('root')).render(<App />);
