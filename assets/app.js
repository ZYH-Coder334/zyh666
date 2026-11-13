/* 巷子里·小馆 — 共享数据层（菜单 / 订单库 / 跨页同步 / 示例数据） */

const MENU = [
  { id: 'n1', cat: '招牌', name: '红烧牛肉面', price: 38, icon: '面', desc: '牛腩慢炖三小时，汤浓肉烂', tag: '热销' },
  { id: 'n2', cat: '招牌', name: '金汤酸菜鱼', price: 58, icon: '鱼', desc: '现杀黑鱼片，酸辣开胃', tag: '招牌' },
  { id: 'n3', cat: '招牌', name: '葱油拌面', price: 22, icon: '葱', desc: '手熬葱油，一人食刚刚好' },
  { id: 's1', cat: '小菜', name: '口水鸡', price: 28, icon: '鸡', desc: '川味红油，麻香回甘' },
  { id: 's2', cat: '小菜', name: '凉拌木耳', price: 16, icon: '木', desc: '爽脆弹牙，蒜香清爽', tag: '素食' },
  { id: 's3', cat: '小菜', name: '手工春卷', price: 18, icon: '卷', desc: '现包现炸，六只装' },
  { id: 'z1', cat: '主食', name: '扬州炒饭', price: 26, icon: '饭', desc: '粒粒分明，虾仁火腿' },
  { id: 'z2', cat: '主食', name: '鲜虾馄饨', price: 24, icon: '馄', desc: '整只虾仁，紫菜汤底' },
  { id: 'd1', cat: '饮品', name: '酸梅汤', price: 12, icon: '梅', desc: '乌梅山楂现熬，冰镇' },
  { id: 'd2', cat: '饮品', name: '茉莉冷萃', price: 16, icon: '莉', desc: '冷萃八小时，花香清冽' },
  { id: 'd3', cat: '饮品', name: '手作豆乳', price: 14, icon: '豆', desc: '当日现磨，微甜' },
  { id: 't1', cat: '甜品', name: '桂花酒酿圆子', price: 16, icon: '圆', desc: '桂花蜜配手工小圆子' }
];

const CATS = ['全部', ...Array.from(new Set(MENU.map(m => m.cat)))];

const STATUS = {
  pending: { label: '待接单', cls: 'st-pending' },
  cooking: { label: '制作中', cls: 'st-cooking' },
  served: { label: '已出餐', cls: 'st-served' },
  paid: { label: '已结账', cls: 'st-paid' }
};

const FLOW = ['pending', 'cooking', 'served', 'paid'];

const Store = (() => {
  const KEY = 'xiangzi_store_v1';
  const bc = typeof BroadcastChannel !== 'undefined' ? new BroadcastChannel('xiangzi_store') : null;
  const listeners = new Set();

  const read = () => {
    try { return JSON.parse(localStorage.getItem(KEY) || 'null'); } catch (e) { return null; }
  };
  const emit = () => {
    listeners.forEach(fn => { try { fn(); } catch (e) { /* 忽略单个订阅错误 */ } });
    if (bc) bc.postMessage('sync');
  };
  const write = state => {
    localStorage.setItem(KEY, JSON.stringify(state));
    emit();
  };
  const ensure = () => {
    let s = read();
    if (!s) { s = { orders: [], seq: 1 }; localStorage.setItem(KEY, JSON.stringify(s)); }
    return s;
  };

  if (bc) bc.onmessage = () => listeners.forEach(fn => { try { fn(); } catch (e) { /* noop */ } });
  window.addEventListener('storage', e => {
    if (e.key === KEY) listeners.forEach(fn => { try { fn(); } catch (err) { /* noop */ } });
  });

  return {
    on(fn) { listeners.add(fn); return () => listeners.delete(fn); },
    orders() { return ensure().orders.slice().sort((a, b) => b.createdAt - a.createdAt); },
    get(id) { return ensure().orders.find(o => o.id === id); },
    add({ table, items, note }) {
      const s = ensure();
      const id = 'A' + String(s.seq).padStart(3, '0');
      s.seq += 1;
      const t = Date.now();
      const order = {
        id,
        table,
        items,
        note: note || '',
        total: items.reduce((sum, i) => sum + i.price * i.qty, 0),
        status: 'pending',
        createdAt: t,
        timeline: { pending: t }
      };
      s.orders.push(order);
      write(s);
      return order;
    },
    setStatus(id, status) {
      const s = ensure();
      const o = s.orders.find(x => x.id === id);
      if (!o) return;
      o.status = status;
      o.timeline = o.timeline || {};
      o.timeline[status] = Date.now();
      write(s);
    },
    reset() { write({ orders: [], seq: 1 }); },
    seed() {
      const base = new Date();
      const at = (h, m) => { const d = new Date(base); d.setHours(h, m, 0, 0); return d.getTime(); };
      const mk = (id, qty) => {
        const m = MENU.find(x => x.id === id);
        return { id, name: m.name, price: m.price, icon: m.icon, qty };
      };
      const plan = [
        [11, 42, '03', 'paid', [['n1', 1], ['s2', 1], ['d1', 1]]],
        [12, 8, '01', 'paid', [['n2', 1], ['s1', 1], ['d2', 2]]],
        [12, 36, '05', 'paid', [['n1', 2], ['s3', 1], ['d1', 2]]],
        [13, 4, '02', 'paid', [['n3', 1], ['s1', 1], ['d3', 1]]],
        [17, 55, '11', 'paid', [['z1', 1], ['z2', 2], ['d1', 2]]],
        [18, 26, '07', 'paid', [['n2', 1], ['s2', 1], ['d1', 2]]],
        [19, 3, '04', 'paid', [['n1', 1], ['s1', 2], ['d2', 1]]],
        [19, 41, '06', 'paid', [['n1', 1], ['t1', 2], ['d3', 1]]],
        [20, 5, '06', 'served', [['n2', 2], ['s3', 1], ['d3', 1]]],
        [20, 16, '08', 'cooking', [['n1', 1], ['s1', 1]]],
        [20, 24, '09', 'pending', [['n3', 1], ['s2', 1], ['d1', 1]]]
      ];
      const orders = plan.map((p, i) => {
        const [h, m, table, status, raw] = p;
        const items = raw.map(([id, qty]) => mk(id, qty));
        const createdAt = at(h, m);
        return {
          id: 'A' + String(i + 1).padStart(3, '0'),
          table,
          items,
          note: '',
          total: items.reduce((sum, it) => sum + it.price * it.qty, 0),
          status,
          createdAt,
          timeline: { pending: createdAt }
        };
      });
      write({ orders, seq: orders.length + 1 });
    }
  };
})();

/* 首次打开自动载入一份示例数据，方便演示；店长看板可一键清空 */
if (!localStorage.getItem('xiangzi_store_v1') && !new URLSearchParams(location.search).has('empty')) {
  Store.seed();
}

const money = n => '¥' + n;
const hhmm = ts => {
  const d = new Date(ts);
  return String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0');
};
const isToday = ts => new Date(ts).toDateString() === new Date().toDateString();
const elapsed = ts => {
  const min = Math.max(0, Math.floor((Date.now() - ts) / 60000));
  return min < 60 ? min + ' 分钟' : Math.floor(min / 60) + ' 小时 ' + (min % 60) + ' 分';
};
const tableNo = () => (new URLSearchParams(location.search).get('table') || '08').padStart(2, '0');
