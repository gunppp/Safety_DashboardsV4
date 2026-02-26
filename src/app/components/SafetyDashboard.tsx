import { useEffect, useMemo, useRef, useState, useCallback } from 'react';
import {
  Activity,
  AlertTriangle,
  Calendar as CalendarIcon,
  CheckCircle2,
  Edit,
  Flame,
  Image as ImageIcon,
  Lock,
  Megaphone,
  Plus,
  BarChart3,
  Save,
  Shield,
  Target,
  Trash2,
  Unlock,
  Upload,
  X,
} from 'lucide-react';
import { DndProvider, useDrag, useDrop } from 'react-dnd';
import { HTML5Backend } from 'react-dnd-html5-backend';
import nhkLogo from '@/assets/nhk-logo.png';

import {
  Dialog,
  DialogContent,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from '@/app/components/ui/dialog';
import { ScrollArea } from '@/app/components/ui/scroll-area';

import {
  Bar,
  CartesianGrid,
  ComposedChart,
  XAxis,
  YAxis,
} from 'recharts';

import {
  ChartContainer,
  ChartLegend,
  ChartLegendContent,
  ChartTooltip,
  ChartTooltipContent,
} from '@/app/components/ui/chart';

type DayStatus = 'safe' | 'near_miss' | 'accident' | null;
interface DailyStatistic { day: number; status: DayStatus }
interface MonthlyData { month: number; year: number; days: DailyStatistic[] }
interface Announcement { id: string; text: string }
interface SafetyMetric { id: string; label: string; value: string; unit?: string }
interface SafetyTrendRow {
  year: number;
  firstAid: number;
  nonAbsent: number;
  absent: number;
  fire: number;
  ifr?: number;
  isr?: number;
}

type PanelKey =
  | 'slogan'
  | 'safetyData'
  | 'announcements'
  | 'calendar'
  | 'streak'
  | 'policy'
  | 'poster';

type SlotKey =
  | 'leftTop'
  | 'leftBottom'
  | 'centerTop'
  | 'centerMid'
  | 'centerBottom'
  | 'rightTop'
  | 'rightBottom';

interface LayoutState {
  cols: [number, number, number];
  leftRows: [number, number];
  centerRows: [number, number, number];
  rightRows: [number, number];
}

const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const DAY_HEADERS = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
const DEFAULT_ANNOUNCEMENTS: Announcement[] = [
  { id: '1', text: 'PPE Audit ประจำสัปดาห์ทุกวันพฤหัสบดี เวลา 09:00 น.' },
  { id: '2', text: 'Emergency Drill ไตรมาสนี้กำหนดวันที่ 28 มีนาคม 2026' },
];
const DEFAULT_POLICY_LINES = [
  'ปฏิบัติตามกฎความปลอดภัยและสวม PPE ก่อนเข้าพื้นที่ผลิต',
  'แจ้ง Near Miss / Unsafe Condition ทันทีเมื่อพบความเสี่ยง',
  'หยุดงานทันทีเมื่อพบสภาพไม่ปลอดภัย (Stop Work Authority)',
  'ทุกคนมีส่วนร่วมรักษา Zero Accident Workplace',
];
const DEFAULT_METRICS: SafetyMetric[] = [
  { id: 'm1', label: 'First Aid', value: '0', unit: 'case' },
  { id: 'm2', label: 'Non-Absent', value: '0', unit: 'case' },
  { id: 'm3', label: 'Absent', value: '0', unit: 'case' },
  { id: 'm4', label: 'Fire', value: '0', unit: 'case' },
  { id: 'm5', label: 'IFR', value: '0', unit: '' },
  { id: 'm6', label: 'ISR', value: '1.2', unit: '' },
];

const DEFAULT_POSTER_URL = '/company-policy-poster.png';

function defaultTrendRows(year: number): SafetyTrendRow[] {
  // 5-year window (4 years back + current year)
  return Array.from({ length: 5 }, (_, i) => {
    const y = year - 4 + i;
    return { year: y, firstAid: 0, nonAbsent: 0, absent: 0, fire: 0, ifr: 0, isr: 0 };
  });
}

const BASE_VIEWPORT = { width: 1920, height: 1080 };
const DEFAULT_LAYOUT: LayoutState = {
  cols: [25, 45, 30],
  leftRows: [28, 72],
  centerRows: [60, 18, 22],
  rightRows: [34, 66],
};
const DEFAULT_SLOTS: Record<SlotKey, PanelKey> = {
  leftTop: 'slogan',
  leftBottom: 'poster',
  centerTop: 'safetyData',
  centerMid: 'policy',
  centerBottom: 'announcements',
  rightTop: 'streak',
  rightBottom: 'calendar',
};

function clamp(n: number, min: number, max: number) { return Math.min(max, Math.max(min, n)); }
function sum(arr: number[]) { return arr.reduce((a,b)=>a+b,0); }
function normalized<T extends number[]>(arr: T): T {
  const s = sum(arr as number[]);
  return arr.map((v)=> (v/s)*100) as T;
}
function rootFontSize(w:number,h:number){
  const scale = Math.min(w/BASE_VIEWPORT.width, h/BASE_VIEWPORT.height);
  return clamp(16 * Math.pow(Math.max(scale,0.35), 0.45), 14, 24);
}
function panelScaleFromSize(w:number,h:number){
  const ratio = Math.min(w/540, h/320);
  return clamp(Math.pow(Math.max(ratio, 0.35), 0.42), 0.68, 1.18);
}
function scaledPx(base:number, panelScale:number, min?:number, max?:number){
  const px = clamp(base * panelScale, min ?? base*0.8, max ?? base*1.35);
  return `${px/16}rem`;
}
function nextDayStatus(status: DayStatus): DayStatus {
  if (status === null) return 'safe';
  if (status === 'safe') return 'near_miss';
  if (status === 'near_miss') return 'accident';
  return null;
}
function createYearData(year:number): MonthlyData[] {
  return Array.from({length:12}, (_,m)=> ({
    month:m, year,
    days: Array.from({length:new Date(year,m+1,0).getDate()},(_,i)=>({day:i+1,status:null}))
  }));
}

const AUTO_SAFE_HOUR = 16;

function applyAutoSafe(prev: MonthlyData[], now: Date, year: number): MonthlyData[] {
  if (now.getFullYear() !== year) return prev;
  const todayStart = new Date(year, now.getMonth(), now.getDate());
  const afterCutoff = (now.getHours() > AUTO_SAFE_HOUR) || (now.getHours() === AUTO_SAFE_HOUR && now.getMinutes() >= 0);

  let next: MonthlyData[] | null = null;
  const ensureNext = () => {
    if (!next) next = prev.map((mm) => ({ ...mm, days: mm.days.map((dd) => ({ ...dd })) }));
    return next;
  };

  let changed = false;

  for (let m = 0; m < 12; m++) {
    const month = prev[m];
    if (!month) continue;
    for (let i = 0; i < month.days.length; i++) {
      const dd = month.days[i];
      if (dd.status !== null) continue;
      const dt = new Date(year, m, dd.day);
      if (dt < todayStart) {
        const tgt = ensureNext()[m].days[i];
        tgt.status = 'safe';
        changed = true;
        continue;
      }
      if (afterCutoff && m === now.getMonth() && dd.day === now.getDate()) {
        const tgt = ensureNext()[m].days[i];
        tgt.status = 'safe';
        changed = true;
      }
    }
  }

  return changed && next ? next : prev;
}

function uid(prefix='id'){ return `${prefix}-${Math.random().toString(16).slice(2)}-${Date.now()}`; }
function isValidMonthlyData(data: unknown, year:number): data is MonthlyData[] {
  return Array.isArray(data) && data.length===12 && data.every((m:any, idx)=>
    m && m.month===idx && m.year===year && Array.isArray(m.days) && m.days.length===new Date(year, idx+1, 0).getDate()
  );
}

function isValidTrendRows(data: unknown): data is SafetyTrendRow[] {
  if (!Array.isArray(data) || data.length < 1 || data.length > 12) return false;
  return data.every((r: any) =>
    r && Number.isFinite(Number(r.year)) &&
    ['firstAid','nonAbsent','absent','fire'].every((k) => Number.isFinite(Number(r[k])))
  );
}
function isValidLayout(data: any): data is LayoutState {
  if (!data) return false;
  const keys: (keyof LayoutState)[] = ['cols','leftRows','centerRows','rightRows'];
  return keys.every((k)=> Array.isArray(data[k]) && data[k].every((v:number)=> typeof v === 'number'));
}
function isValidSlots(data:any): data is Record<SlotKey, PanelKey> {
  const slots: SlotKey[] = ['leftTop','leftBottom','centerTop','centerMid','centerBottom','rightTop','rightBottom'];
  const panels: PanelKey[] = ['slogan','safetyData','announcements','calendar','streak','policy','poster'];
  return data && typeof data === 'object' && slots.every((s)=> panels.includes((data as any)[s]));
}

function useResizeGroup(values: number[], setValues: (next:number[])=>void, minEach = 12){
  return (index:number)=> (e: React.MouseEvent) => {
    e.preventDefault();
    const startX = e.clientX;
    const startY = e.clientY;
    const target = e.currentTarget as HTMLElement;
    const orientation = target.dataset.orientation as 'horizontal'|'vertical';
    const container = target.parentElement as HTMLElement;
    if (!container) return;
    const rect = container.getBoundingClientRect();
    const totalPx = orientation === 'vertical' ? rect.width : rect.height;
    const start = [...values];
    const move = (ev: MouseEvent) => {
      const deltaPx = orientation === 'vertical' ? (ev.clientX - startX) : (ev.clientY - startY);
      const deltaPct = (deltaPx / Math.max(1, totalPx)) * 100;
      let a = start[index] + deltaPct;
      let b = start[index+1] - deltaPct;
      const rest = sum(start) - start[index] - start[index+1];
      const maxA = 100 - rest - minEach;
      a = clamp(a, minEach, maxA);
      b = 100 - rest - a;
      if (b < minEach) {
        b = minEach;
        a = 100 - rest - b;
      }
      const next = [...start];
      next[index] = a;
      next[index+1] = b;
      setValues(normalized(next));
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      document.body.style.userSelect = '';
      document.body.style.cursor = '';
    };
    document.body.style.userSelect = 'none';
    document.body.style.cursor = orientation === 'vertical' ? 'col-resize' : 'row-resize';
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };
}

function Splitter({ orientation, onMouseDown }: { orientation:'vertical'|'horizontal'; onMouseDown:(e:React.MouseEvent)=>void }) {
  return (
    <div
      data-orientation={orientation}
      onMouseDown={onMouseDown}
      className={orientation === 'vertical'
        ? 'relative w-2 -mx-1 cursor-col-resize group'
        : 'relative h-2 -my-1 cursor-row-resize group'}
      title="ลากเพื่อปรับขนาด"
    >
      <div className={orientation === 'vertical'
        ? 'absolute left-1/2 top-0 h-full w-[3px] -translate-x-1/2 rounded-full bg-sky-300/60 group-hover:bg-sky-500'
        : 'absolute top-1/2 left-0 w-full h-[3px] -translate-y-1/2 rounded-full bg-sky-300/60 group-hover:bg-sky-500'}
      />
    </div>
  );
}

function HeaderClock({ centered = false }: { centered?: boolean }) {
  const [now, setNow] = useState(() => new Date());
  useEffect(() => {
    const t = window.setInterval(() => setNow(new Date()), 1000);
    return () => window.clearInterval(t);
  }, []);
  const dateLabel = now.toLocaleDateString([], { weekday: 'short', day: '2-digit', month: 'short', year: 'numeric' });
  const timeLabel = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
  return (
    <div
      className={`${centered ? 'flex' : 'hidden md:flex'} flex-col ${centered ? 'items-center' : 'items-end'} tabular-nums leading-none`}
      aria-label="Current date and time"
    >
      <div className={`font-extrabold text-slate-800 ${centered ? 'header-date' : 'text-sm'}`}>{dateLabel}</div>
      <div className={`font-black text-slate-900 ${centered ? 'header-time' : 'text-xs'} mt-1`}>{timeLabel}</div>
    </div>
  );
}

function Card({
  title,
  icon,
  actions,
  children,
  className='',
  tone='sky',
  panelScale=1,
}:{
  title:string;
  icon:React.ReactNode;
  actions?: React.ReactNode;
  children:React.ReactNode;
  className?:string;
  tone?: 'sky'|'amber'|'green'|'blue'|'teal';
  panelScale?: number;
}) {
  const toneMap = {
    sky: {
      outer: 'border-sky-200 bg-gradient-to-b from-sky-50/70 to-white',
      header: 'from-sky-100 via-white to-sky-50 border-sky-200',
      body: 'bg-gradient-to-b from-white to-sky-50/25',
    },
    amber: {
      outer: 'border-amber-200 bg-gradient-to-b from-amber-50/80 to-white',
      header: 'from-amber-100 via-white to-yellow-50 border-amber-200',
      body: 'bg-gradient-to-b from-white to-amber-50/20',
    },
    green: {
      outer: 'border-emerald-200 bg-gradient-to-b from-emerald-50/70 to-white',
      header: 'from-emerald-100 via-white to-lime-50 border-emerald-200',
      body: 'bg-gradient-to-b from-white to-emerald-50/20',
    },
    blue: {
      outer: 'border-blue-200 bg-gradient-to-b from-blue-50/70 to-white',
      header: 'from-blue-100 via-white to-cyan-50 border-blue-200',
      body: 'bg-gradient-to-b from-white to-blue-50/20',
    },
    teal: {
      outer: 'border-cyan-200 bg-gradient-to-b from-cyan-50/70 to-white',
      header: 'from-cyan-100 via-white to-teal-50 border-cyan-200',
      body: 'bg-gradient-to-b from-white to-cyan-50/20',
    },
  } as const;
  const toneCls = toneMap[tone];
  return (
    <section className={`rounded-2xl border shadow-sm min-h-0 flex flex-col overflow-hidden ${toneCls.outer} ${className}`}>
      <div className={`relative px-4 py-3 border-b bg-gradient-to-r ${toneCls.header} flex items-center gap-2 text-slate-800 font-semibold`}>
        <div className="h-7 w-7 rounded-lg bg-white/80 border border-white shadow-sm flex items-center justify-center shrink-0">
          {icon}
        </div>
        <h2 className="truncate font-extrabold tracking-tight" style={{ fontSize: scaledPx(17, panelScale, 14, 22) }}>{title}</h2>
        {actions ? (
          <div className="ml-auto flex items-center gap-1">
            {actions}
          </div>
        ) : null}
      </div>
      <div className={`p-3 min-h-0 flex-1 overflow-hidden flex flex-col ${toneCls.body}`} style={{ fontSize: scaledPx(14, 0.95 + (panelScale-1)*0.35, 11, 16) }}>{children}</div>
    </section>
  );
}

const DND_TYPE = 'DASH_PANEL';

function DroppableSlot({
  slot,
  locked,
  onSwap,
  children,
}: {
  slot: SlotKey;
  locked: boolean;
  onSwap: (from: SlotKey, to: SlotKey) => void;
  children: React.ReactNode;
}) {
  const [{ isOver }, drop] = useDrop(() => ({
    accept: DND_TYPE,
    canDrop: () => !locked,
    drop: (item: any) => {
      if (!item?.fromSlot) return;
      if (item.fromSlot === slot) return;
      onSwap(item.fromSlot as SlotKey, slot);
    },
    collect: (monitor) => ({ isOver: monitor.isOver({ shallow: true }) }),
  }), [slot, locked, onSwap]);

  return (
    <div ref={drop as any} className={isOver && !locked ? 'ring-2 ring-sky-400 rounded-2xl' : ''}>
      {children}
    </div>
  );
}



type RenderPanelFn = (panel: PanelKey, panelScale: number) => React.ReactNode;

function DashboardSlot({
  slot,
  panel,
  layoutLocked,
  onSwap,
  renderPanel,
}: {
  slot: SlotKey;
  panel: PanelKey;
  layoutLocked: boolean;
  onSwap: (from: SlotKey, to: SlotKey) => void;
  renderPanel: RenderPanelFn;
}) {
  const slotContainerRef = useRef<HTMLDivElement>(null);
  const [panelScale, setPanelScale] = useState(1);

  useEffect(() => {
    const el = slotContainerRef.current;
    if (!el || typeof ResizeObserver === 'undefined') return;
    let raf = 0;
    const ro = new ResizeObserver((entries) => {
      const rect = entries[0]?.contentRect;
      if (!rect) return;
      cancelAnimationFrame(raf);
      raf = window.requestAnimationFrame(() => {
        const next = panelScaleFromSize(rect.width, rect.height);
        setPanelScale((prev) => (Math.abs(prev - next) > 0.02 ? next : prev));
      });
    });
    ro.observe(el);
    return () => {
      cancelAnimationFrame(raf);
      ro.disconnect();
    };
  }, []);

  const [{ isDragging }, drag] = useDrag(() => ({
    type: DND_TYPE,
    item: { fromSlot: slot },
    canDrag: !layoutLocked,
    collect: (monitor) => ({ isDragging: monitor.isDragging() }),
  }), [slot, layoutLocked]);

  return (
    <DroppableSlot slot={slot} locked={layoutLocked} onSwap={onSwap}>
      <div ref={slotContainerRef} className={`h-full ${isDragging ? 'opacity-60' : ''}`}>
        <div className="relative h-full">
          {!layoutLocked && (
            <div
              ref={drag as any}
              className="absolute left-3 right-28 top-3 z-20 h-10 rounded-xl border border-sky-200/80 bg-white/70 backdrop-blur cursor-grab active:cursor-grabbing flex items-center px-3 gap-2 shadow-sm"
              title="คลิกค้างแล้วลากเพื่อย้ายช่อง"
            >
              <Unlock className="h-4 w-4 text-sky-700" />
              <span className="text-xs font-extrabold tracking-wide text-sky-800">ลากย้ายช่อง</span>
            </div>
          )}
          {renderPanel(panel, panelScale)}
        </div>
      </div>
    </DroppableSlot>
  );
}

export function SafetyDashboard() {
  const now = new Date();
  const [displayMonth, setDisplayMonth] = useState(now.getMonth());
  const [currentYear, setCurrentYear] = useState(now.getFullYear());
  const [monthlyData, setMonthlyData] = useState<MonthlyData[]>(() => createYearData(now.getFullYear()));
  const [announcements, setAnnouncements] = useState<Announcement[]>(DEFAULT_ANNOUNCEMENTS);
  const [policyPoster, setPolicyPoster] = useState<string | null>(DEFAULT_POSTER_URL);
  const [posterZoom, setPosterZoom] = useState(1);
  const [policyTitle, setPolicyTitle] = useState('Safety Policy');
  const [policyLines, setPolicyLines] = useState<string[]>(DEFAULT_POLICY_LINES);
  const [sloganTh, setSloganTh] = useState('ความปลอดภัย เริ่มที่ตัวเรา');
  const [sloganEn, setSloganEn] = useState('Safety Starts With Me');
  const [metrics, setMetrics] = useState<SafetyMetric[]>(DEFAULT_METRICS);
  const [trendRows, setTrendRows] = useState<SafetyTrendRow[]>(() => defaultTrendRows(now.getFullYear()));
  const [layout, setLayout] = useState<LayoutState>(DEFAULT_LAYOUT);
  const [slots, setSlots] = useState<Record<SlotKey, PanelKey>>(DEFAULT_SLOTS);
  const [layoutLocked, setLayoutLocked] = useState(true);

  const [editingAnnId, setEditingAnnId] = useState<string | null>(null);
  const [annDraft, setAnnDraft] = useState('');
  const [editSlogan, setEditSlogan] = useState(false);
  const [sloganThDraft, setSloganThDraft] = useState('');
  const [sloganEnDraft, setSloganEnDraft] = useState('');
  const [editPolicy, setEditPolicy] = useState(false);
  const [policyTitleDraft, setPolicyTitleDraft] = useState('');
  const [policyLinesDraft, setPolicyLinesDraft] = useState('');
  const [editMetrics, setEditMetrics] = useState(false);
  const [metricsDraft, setMetricsDraft] = useState<SafetyMetric[]>([]);
  const [editTrend, setEditTrend] = useState(false);
  const [trendDraft, setTrendDraft] = useState<SafetyTrendRow[]>([]);
  const [uiScale, setUiScale] = useState<number>(() => {
    try {
      const raw = localStorage.getItem('safety-dashboard-ui-scale');
      const v = raw ? Number(raw) : 1;
      return Number.isFinite(v) ? clamp(v, 0.8, 1.4) : 1;
    } catch {
      return 1;
    }
  });

  const fileInputRef = useRef<HTMLInputElement>(null);

  const storageKey = `safety-dashboard-${currentYear}`;
  // v7: layout + defaults updated
  const layoutKey = 'safety-dashboard-layout-v7';
  const slotKey = 'safety-dashboard-slots-v4';

  // Auto mark SAFE at 16:00 (today) and backfill past days (before today) as SAFE when still NOT SET.
// - Backfill is applied on load when reading from storage (see storage effect below)
// - This scheduler triggers at exactly 16:00:00 each day (local time)
useEffect(() => {
  const now = new Date();
  if (now.getFullYear() !== currentYear) return;

  let timer: number | undefined;

  const scheduleNext = () => {
    const current = new Date();
    const next = new Date(current);
    next.setHours(AUTO_SAFE_HOUR, 0, 0, 0);
    if (current.getTime() >= next.getTime()) next.setDate(next.getDate() + 1);

    const ms = Math.max(250, next.getTime() - current.getTime());
    timer = window.setTimeout(() => {
      const fireNow = new Date();
      setMonthlyData((prev) => applyAutoSafe(prev, fireNow, currentYear));
      scheduleNext();
    }, ms);
  };

  scheduleNext();
  return () => {
    if (timer) window.clearTimeout(timer);
  };
}, [currentYear]);

  useEffect(() => {
    const onResize = () => {
      const root = document.documentElement;
      const base = rootFontSize(window.innerWidth, window.innerHeight);
      root.style.setProperty('--font-size', `${base * uiScale}px`);
    };
    onResize();
    window.addEventListener('resize', onResize);
    return () => window.removeEventListener('resize', onResize);
  }, [uiScale]);

  useEffect(() => {
    try { localStorage.setItem('safety-dashboard-ui-scale', String(uiScale)); } catch {}
  }, [uiScale]);

  useEffect(() => {
    try {
      const raw = localStorage.getItem(storageKey);
      if (raw) {
        const parsed = JSON.parse(raw);
        const loadedMonthly = isValidMonthlyData(parsed.monthlyData, currentYear) ? parsed.monthlyData : createYearData(currentYear);
        setMonthlyData(applyAutoSafe(loadedMonthly, new Date(), currentYear));
        setAnnouncements(Array.isArray(parsed.announcements) && parsed.announcements.length ? parsed.announcements : DEFAULT_ANNOUNCEMENTS);
        const poster = Object.prototype.hasOwnProperty.call(parsed, 'policyPoster')
          ? (typeof parsed.policyPoster === 'string' ? parsed.policyPoster : null)
          : DEFAULT_POSTER_URL;
        setPolicyPoster(poster);
        setPosterZoom(typeof parsed.posterZoom === 'number' ? clamp(parsed.posterZoom, 0.5, 2.5) : 1);
        setPolicyTitle(typeof parsed.policyTitle === 'string' ? parsed.policyTitle : 'Safety Policy');
        setPolicyLines(Array.isArray(parsed.policyLines) && parsed.policyLines.length ? parsed.policyLines : DEFAULT_POLICY_LINES);
        setSloganTh(typeof parsed.sloganTh === 'string' ? parsed.sloganTh : 'ความปลอดภัย เริ่มที่ตัวเรา');
        setSloganEn(typeof parsed.sloganEn === 'string' ? parsed.sloganEn : 'Safety Starts With Me');
        setMetrics(Array.isArray(parsed.metrics) && parsed.metrics.length ? parsed.metrics : DEFAULT_METRICS);
        setTrendRows(isValidTrendRows(parsed.trendRows) ? parsed.trendRows : defaultTrendRows(currentYear));
      } else {
        setMonthlyData(applyAutoSafe(createYearData(currentYear), new Date(), currentYear));
        setPolicyPoster(DEFAULT_POSTER_URL);
        setTrendRows(defaultTrendRows(currentYear));
      }
    } catch {
      setMonthlyData(applyAutoSafe(createYearData(currentYear), new Date(), currentYear));
      setPolicyPoster(DEFAULT_POSTER_URL);
      setTrendRows(defaultTrendRows(currentYear));
    }
  }, [storageKey, currentYear]);

  useEffect(() => {
    const payload = {
      monthlyData,
      announcements,
      policyPoster,
      posterZoom,
      policyTitle,
      policyLines,
      sloganTh,
      sloganEn,
      metrics,
      trendRows,
    };
    const t = window.setTimeout(() => {
      try { localStorage.setItem(storageKey, JSON.stringify(payload)); } catch {}
    }, 450);
    return () => window.clearTimeout(t);
  }, [storageKey, monthlyData, announcements, policyPoster, posterZoom, policyTitle, policyLines, sloganTh, sloganEn, metrics, trendRows]);

  useEffect(() => {
    try {
      const raw = localStorage.getItem(layoutKey);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (isValidLayout(parsed)) setLayout({
          cols: normalized(parsed.cols),
          leftRows: normalized(parsed.leftRows),
          centerRows: normalized(parsed.centerRows),
          rightRows: normalized(parsed.rightRows),
        });
      }
    } catch {}
    try {
      const rawS = localStorage.getItem(slotKey);
      if (rawS) {
        const parsed = JSON.parse(rawS);
        if (isValidSlots(parsed)) setSlots(parsed);
      }
    } catch {}
  }, []);

  useEffect(() => {
    const t = window.setTimeout(() => { try { localStorage.setItem(layoutKey, JSON.stringify(layout)); } catch {} }, 250);
    return () => window.clearTimeout(t);
  }, [layout]);
  useEffect(() => {
    const t = window.setTimeout(() => { try { localStorage.setItem(slotKey, JSON.stringify(slots)); } catch {} }, 250);
    return () => window.clearTimeout(t);
  }, [slots]);


  const onResizeCols = useResizeGroup(layout.cols, (next) => setLayout((p) => ({ ...p, cols: next as any })), 18);
  const onResizeLeft = useResizeGroup(layout.leftRows, (next) => setLayout((p) => ({ ...p, leftRows: next as any })), 14);
  const onResizeCenter = useResizeGroup(layout.centerRows, (next) => setLayout((p) => ({ ...p, centerRows: next as any })), 18);
  const onResizeRight = useResizeGroup(layout.rightRows, (next) => setLayout((p) => ({ ...p, rightRows: next as any })), 16);

  const displayMonthData = monthlyData[displayMonth];

  const safetyStreak = useMemo(() => {
    const today = new Date();
    const y = today.getFullYear();
    if (y !== currentYear) return 0;
    const end = new Date(y, today.getMonth(), today.getDate());
    let streak = 0;
    for (let dt = new Date(end); ; ) {
      const m = dt.getMonth();
      const d = dt.getDate();
      const st = monthlyData[m]?.days?.[d - 1]?.status ?? null;
      // "Safety streak" in plants typically means "days without accident".
      // Count SAFE and NEAR MISS as streak days; break on ACCIDENT or NOT SET.
      if (st === 'safe' || st === 'near_miss') streak += 1;
      else break;
      dt.setDate(dt.getDate() - 1);
      if (dt.getFullYear() !== y) break;
    }
    return streak;
  }, [monthlyData, currentYear]);

  const monthSummary = useMemo(() => {
    if (!displayMonthData) return { safe: 0, near: 0, accident: 0 };
    let safe = 0, near = 0, accident = 0;
    for (const d of displayMonthData.days) {
      if (d.status === 'safe') safe += 1;
      if (d.status === 'near_miss') near += 1;
      if (d.status === 'accident') accident += 1;
    }
    return { safe, near, accident };
  }, [displayMonthData]);

  const firstDayOffset = useMemo(() => new Date(currentYear, displayMonth, 1).getDay(), [currentYear, displayMonth]);
  const daysInMonth = useMemo(() => new Date(currentYear, displayMonth + 1, 0).getDate(), [currentYear, displayMonth]);
  const gridCells = useMemo(() => {
    const cells: Array<{ day: number | null }> = [];
    for (let i = 0; i < firstDayOffset; i++) cells.push({ day: null });
    for (let d = 1; d <= daysInMonth; d++) cells.push({ day: d });
    while (cells.length % 7 !== 0) cells.push({ day: null });
    return cells;
  }, [firstDayOffset, daysInMonth]);

  const setDayStatus = (day: number, status: DayStatus) => {
    setMonthlyData((prev) => {
      const next = prev.map((mm) => ({ ...mm, days: mm.days.map((dd) => ({ ...dd })) }));
      const month = next[displayMonth];
      if (!month) return prev;
      const target = month.days[day - 1];
      if (!target) return prev;
      target.status = status;
      return next;
    });
  };

  const cycleDayStatus = (day: number) => {
    const current = displayMonthData?.days?.[day - 1]?.status ?? null;
    setDayStatus(day, nextDayStatus(current));
  };

  const swapSlots = (from: SlotKey, to: SlotKey) => {
    setSlots((prev) => {
      const next = { ...prev };
      const a = next[from];
      const b = next[to];
      next[from] = b;
      next[to] = a;
      return next;
    });
  };

  const resetLayout = () => {
    setLayout(DEFAULT_LAYOUT);
    setSlots(DEFAULT_SLOTS);
    setPosterZoom(1);
    try { localStorage.removeItem(layoutKey); } catch {}
    try { localStorage.removeItem(slotKey); } catch {}
  };

  const startEditSlogan = () => {
    setSloganThDraft(sloganTh);
    setSloganEnDraft(sloganEn);
    setEditSlogan(true);
  };
  const saveSlogan = () => {
    setSloganTh(sloganThDraft.trim() || sloganTh);
    setSloganEn(sloganEnDraft.trim() || sloganEn);
    setEditSlogan(false);
  };

  const startEditPolicy = () => {
    setPolicyTitleDraft(policyTitle);
    setPolicyLinesDraft(policyLines.join('\n'));
    setEditPolicy(true);
  };
  const savePolicy = () => {
    const lines = policyLinesDraft
      .split('\n')
      .map((s) => s.trim())
      .filter(Boolean);
    setPolicyTitle(policyTitleDraft.trim() || policyTitle);
    setPolicyLines(lines.length ? lines : policyLines);
    setEditPolicy(false);
  };

  const addAnnouncement = () => {
    const id = uid('ann');
    setAnnouncements((p) => [{ id, text: 'New announcement...' }, ...p]);
    setEditingAnnId(id);
    setAnnDraft('New announcement...');
  };
  const startEditAnn = (ann: Announcement) => {
    setEditingAnnId(ann.id);
    setAnnDraft(ann.text);
  };
  const saveAnn = (id: string) => {
    setAnnouncements((p) => p.map((a) => (a.id === id ? { ...a, text: annDraft.trim() || a.text } : a)));
    setEditingAnnId(null);
    setAnnDraft('');
  };
  const deleteAnn = (id: string) => {
    setAnnouncements((p) => p.filter((a) => a.id !== id));
    if (editingAnnId === id) { setEditingAnnId(null); setAnnDraft(''); }
  };

  const openMetricsEditor = useCallback(() => {
    setMetricsDraft(metrics.map((m) => ({ ...m })));
    setEditMetrics(true);
  }, [metrics]);

  const closeMetricsEditor = useCallback(() => setEditMetrics(false), []);

  const saveMetricsEditor = useCallback(() => {
    setMetrics(metricsDraft.map((m) => ({ ...m })));
    setEditMetrics(false);
  }, [metricsDraft]);

  // Performance: keep handlers stable so inputs don't lose focus and large lists stay responsive.
  const addMetricDraft = useCallback(() => {
    setMetricsDraft((p) => [{ id: uid('m'), label: 'New Metric', value: '0', unit: '' }, ...p]);
  }, []);

  const updateMetricDraft = useCallback((id: string, patch: Partial<SafetyMetric>) => {
    setMetricsDraft((p) => p.map((m) => (m.id === id ? { ...m, ...patch } : m)));
  }, []);

  const deleteMetricDraft = useCallback((id: string) => {
    setMetricsDraft((p) => p.filter((m) => m.id !== id));
  }, []);

  const openTrendEditor = useCallback(() => {
    setTrendDraft(trendRows.map((r) => ({ ...r })));
    setEditTrend(true);
  }, [trendRows]);

  const closeTrendEditor = useCallback(() => setEditTrend(false), []);

  const saveTrendEditor = useCallback(() => {
    // sort by year and de-dupe
    const next = [...trendDraft]
      .map((r) => ({
        ...r,
        year: Number(r.year),
        firstAid: Number(r.firstAid) || 0,
        nonAbsent: Number(r.nonAbsent) || 0,
        absent: Number(r.absent) || 0,
        fire: Number(r.fire) || 0,
      }))
      .sort((a, b) => a.year - b.year);
    const dedup: SafetyTrendRow[] = [];
    for (const r of next) {
      if (!dedup.length || dedup[dedup.length - 1].year !== r.year) dedup.push(r);
    }
    setTrendRows(dedup.length ? dedup : defaultTrendRows(currentYear));
    setEditTrend(false);
  }, [trendDraft, currentYear]);

  const addTrendYear = useCallback(() => {
    setTrendDraft((p) => {
      const years = p.map((x) => x.year);
      const maxY = years.length ? Math.max(...years) : currentYear;
      return [...p, { year: maxY + 1, firstAid: 0, nonAbsent: 0, absent: 0, fire: 0 }];
    });
  }, [currentYear]);

  const deleteTrendRow = useCallback((idx: number) => {
    setTrendDraft((p) => p.filter((_, i) => i !== idx));
  }, []);

  const updateTrendRow = useCallback((idx: number, patch: Partial<SafetyTrendRow>) => {
    setTrendDraft((p) => p.map((r, i) => (i === idx ? { ...r, ...patch } : r)));
  }, []);

const onPosterSelected = (file?: File | null) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      setPolicyPoster(String(reader.result));
      setPosterZoom(1);
    };
    reader.readAsDataURL(file);
  };

  const renderPanel = useCallback((panel: PanelKey, panelScale = 1) => {
    if (panel === 'slogan') {
      return (
        <Card
          title="Safety Slogan"
          icon={<Target className="h-5 w-5 text-sky-600" />}
          tone="blue"
          panelScale={panelScale}
          actions={
            !editSlogan ? (
              <button onClick={startEditSlogan} className="p-2 rounded-lg hover:bg-white/70" title="Edit">
                <Edit className="h-4 w-4 text-slate-600" />
              </button>
            ) : (
              <>
                <button onClick={saveSlogan} className="p-2 rounded-lg hover:bg-emerald-50" title="Save">
                  <Save className="h-4 w-4 text-emerald-700" />
                </button>
                <button onClick={() => setEditSlogan(false)} className="p-2 rounded-lg hover:bg-white/70" title="Cancel">
                  <X className="h-4 w-4 text-slate-600" />
                </button>
              </>
            )
          }
        >
          {!editSlogan ? (
            <div className="h-full flex items-center justify-center">
              <div className="w-full rounded-2xl border border-sky-100 bg-gradient-to-br from-sky-50 via-white to-cyan-50 p-4 md:p-5 h-full flex flex-col justify-center">
                <div className="font-extrabold text-slate-900 leading-tight text-center break-words" style={{ fontSize: scaledPx(28, panelScale, 16, 40) }}>{sloganTh}</div>
                <div className="mt-2 font-semibold text-slate-600 text-center break-words" style={{ fontSize: scaledPx(18, panelScale, 12, 24) }}>{sloganEn}</div>
              </div>
            </div>
          ) : (
            <div className="space-y-3">
              <label className="block text-sm font-semibold text-slate-700">Thai</label>
              <textarea value={sloganThDraft} onChange={(e)=>setSloganThDraft(e.target.value)} className="w-full rounded-xl border border-slate-200 p-3 bg-white" rows={2} />
              <label className="block text-sm font-semibold text-slate-700">English</label>
              <textarea value={sloganEnDraft} onChange={(e)=>setSloganEnDraft(e.target.value)} className="w-full rounded-xl border border-slate-200 p-3 bg-white" rows={2} />
            </div>
          )}
        </Card>
      );
    }

    if (panel === 'safetyData') {
      const metricCols = panelScale < 0.82 ? 1 : 2;
      const metricRows = Math.max(1, Math.ceil(metrics.length / metricCols));
      const densityScale = clamp((6 / Math.max(metrics.length, 1)) ** 0.28, 0.78, 1);
      const cardScale = clamp(panelScale * densityScale * (metricRows >= 4 ? 0.92 : 1), 0.68, 1.08);

      const metricsFr = panelScale < 0.78 ? 1.45 : 1.65;
      const chartFr = 1;
      const chartMinPx = Math.round(clamp(110 * panelScale, 90, 160));

      const chartConfig = {
        firstAid: { label: 'First Aid', color: '#16a34a' },
        nonAbsent: { label: 'Non-Absent', color: '#0284c7' },
        absent: { label: 'Absent', color: '#f59e0b' },
        fire: { label: 'Fire', color: '#ef4444' },
      } as const;

      return (
        <Card
          title="Safety Data"
          icon={<Activity className="h-5 w-5 text-sky-600" />}
          tone="teal"
          panelScale={panelScale}
          actions={
            <>
              <button type="button" onClick={openMetricsEditor} className="p-2 rounded-lg hover:bg-white/70" title="Edit Safety Data" aria-label="Edit Safety Data">
                <Edit className="h-4 w-4 text-slate-600" />
              </button>
              <button type="button" onClick={openTrendEditor} className="p-2 rounded-lg hover:bg-white/70" title="Edit Trend / Graph" aria-label="Edit Trend / Graph">
                <BarChart3 className="h-4 w-4 text-slate-600" />
              </button>
            </>
          }
        >

          <div className="h-full min-h-0 grid gap-2 overflow-hidden" style={{ gridTemplateRows: `minmax(0, ${metricsFr}fr) minmax(0, ${chartFr}fr)` }}>
            <div className="min-h-0 overflow-hidden">
              {/* Metrics grid */}
              <div
                className="grid gap-2 h-full min-h-0"
                style={{
                  gridTemplateColumns: `repeat(${metricCols}, minmax(0, 1fr))`,
                  gridTemplateRows: `repeat(${metricRows}, minmax(0, 1fr))`,
                }}
              >
                {metrics.map((m) => (
                  <div
                    key={m.id}
                    className="rounded-2xl border border-cyan-100 bg-gradient-to-br from-white to-cyan-50 p-3 flex flex-col justify-between shadow-[0_1px_0_rgba(2,132,199,0.05)] min-h-0 overflow-hidden"
                  >
                    <div className="font-semibold text-slate-700 leading-tight break-words line-clamp-2" style={{ fontSize: scaledPx(14, cardScale, 10, 16) }}>
                      {m.label}
                    </div>
                    <div className="mt-1 flex items-end gap-1 min-w-0">
                      <div className="font-extrabold text-slate-900 leading-none truncate" style={{ fontSize: scaledPx(38, cardScale, 18, 48) }}>
                        {m.value}
                      </div>
                      <div className="font-semibold text-slate-500 pb-[0.12rem] truncate" style={{ fontSize: scaledPx(16, cardScale, 9, 18) }}>
                        {m.unit}
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </div>

            {/* Trend chart: render into the remaining free space under the metric cards */}
            <div
              className="rounded-2xl bg-white/40 p-2 min-h-0 overflow-hidden flex flex-col"
              style={{ minHeight: chartMinPx }}
            >
              <div className="flex items-center justify-between px-1 pb-1 shrink-0">
                <div className="text-xs font-extrabold text-slate-700">Case Trend (5Y)</div>
                <div className="text-[11px] font-semibold text-slate-500 truncate">ย้อนหลัง 4–5 ปี</div>
              </div>
              <div className="flex-1 min-h-0">
                <ChartContainer
                  id="safety-trend"
                  className="h-full w-full aspect-auto"
                  config={chartConfig}
                >
                  <ComposedChart data={trendRows} margin={{ top: 6, right: 10, bottom: 0, left: -4 }}>
                    <CartesianGrid vertical={false} strokeDasharray="3 3" />
                    <XAxis dataKey="year" tickLine={false} axisLine={false} />
                    <YAxis tickLine={false} axisLine={false} allowDecimals={false} width={28} />
                    <ChartTooltip content={<ChartTooltipContent indicator="dot" />} />
                    <ChartLegend content={<ChartLegendContent />} />
                    <Bar dataKey="firstAid" fill="var(--color-firstAid)" radius={[6, 6, 0, 0]} />
                    <Bar dataKey="nonAbsent" fill="var(--color-nonAbsent)" radius={[6, 6, 0, 0]} />
                    <Bar dataKey="absent" fill="var(--color-absent)" radius={[6, 6, 0, 0]} />
                    <Bar dataKey="fire" fill="var(--color-fire)" radius={[6, 6, 0, 0]} />
                  </ComposedChart>
                </ChartContainer>
              </div>
            </div>
          </div>
        </Card>
      );
    }

if (panel === 'announcements') {
      return (
        <Card
          title="Announcements"
          icon={<Megaphone className="h-5 w-5 text-amber-600" />}
          tone="amber"
          panelScale={panelScale}
          actions={
            <button type="button" onClick={addAnnouncement} className="p-2 rounded-lg hover:bg-amber-50" title="Add">
              <Plus className="h-4 w-4 text-amber-700" />
            </button>
          }
        >
          <div className="h-full overflow-hidden pr-1 space-y-2">
            {announcements.map((a) => (
              <div key={a.id} className="rounded-2xl border border-slate-100 bg-white p-3">
                {editingAnnId !== a.id ? (
                  <div className="flex items-start gap-3">
                    <div className="mt-1 h-2 w-2 rounded-full bg-amber-400" />
                    <div className="flex-1 text-slate-800 font-medium leading-snug line-clamp-2" style={{ fontSize: scaledPx(14, panelScale, 11, 16) }}>{a.text}</div>
                    <button onClick={()=>startEditAnn(a)} className="p-2 rounded-xl hover:bg-slate-50" title="Edit">
                      <Edit className="h-4 w-4 text-slate-600" />
                    </button>
                    <button onClick={()=>deleteAnn(a.id)} className="p-2 rounded-xl hover:bg-rose-50" title="Delete">
                      <Trash2 className="h-4 w-4 text-rose-600" />
                    </button>
                  </div>
                ) : (
                  <div className="space-y-2">
                    <textarea value={annDraft} onChange={(e)=>setAnnDraft(e.target.value)} className="w-full rounded-xl border border-slate-200 p-3" rows={2} />
                    <div className="flex justify-end gap-2">
                      <button onClick={()=>saveAnn(a.id)} className="px-3 py-2 rounded-xl bg-emerald-600 text-white font-semibold hover:bg-emerald-700">Save</button>
                      <button onClick={()=>{ setEditingAnnId(null); setAnnDraft(''); }} className="px-3 py-2 rounded-xl border border-slate-200 hover:bg-slate-50">Cancel</button>
                    </div>
                  </div>
                )}
              </div>
            ))}
          </div>
        </Card>
      );
    }

    if (panel === 'calendar') {
      const weekRows = Math.max(4, Math.ceil(gridCells.length / 7));
      const compactScale = clamp(panelScale * (weekRows >= 6 ? 0.9 : 1), 0.62, 1.08);
      return (
        <Card title="Safety Calendar" icon={<CalendarIcon className="h-5 w-5 text-sky-600" />} tone="sky" panelScale={panelScale}>
          <div className="h-full flex flex-col min-h-0">
            <div className="flex items-center justify-between gap-2 mb-2 shrink-0">
              <div className="flex items-center gap-2 min-w-0">
                <button type="button" className="px-2 py-1.5 rounded-xl border border-slate-200 hover:bg-slate-50" onClick={() => setDisplayMonth((m) => (m + 11) % 12)} title="Prev">‹</button>
                <div className="text-base md:text-lg font-extrabold text-slate-900 truncate">{MONTHS[displayMonth]} {currentYear}</div>
                <button type="button" className="px-2 py-1.5 rounded-xl border border-slate-200 hover:bg-slate-50" onClick={() => setDisplayMonth((m) => (m + 1) % 12)} title="Next">›</button>
              </div>
              <div className="grid grid-cols-3 gap-1 text-[10px] md:text-xs font-bold shrink-0">
                <div className="px-2 py-1 rounded-lg bg-emerald-50 border border-emerald-100 text-emerald-700">{monthSummary.safe} SAFE</div>
                <div className="px-2 py-1 rounded-lg bg-amber-50 border border-amber-100 text-amber-800">{monthSummary.near} NM</div>
                <div className="px-2 py-1 rounded-lg bg-rose-50 border border-rose-100 text-rose-700">{monthSummary.accident} ACC</div>
              </div>
            </div>

            <div className="grid grid-cols-7 gap-1 mb-1 shrink-0">
              {DAY_HEADERS.map((d) => (
                <div key={d} className="text-[10px] md:text-xs font-bold text-slate-500 text-center">{d}</div>
              ))}
            </div>
            <div className="mb-1 rounded-lg border border-sky-100 bg-sky-50/80 px-2 py-1 text-[10px] md:text-xs font-semibold text-sky-800 shrink-0 leading-tight">
              คลิกวันเพื่อเปลี่ยนสถานะ: ว่าง → SAFE → NEAR MISS → ACCIDENT → ว่าง
            </div>

            <div className="grid grid-cols-7 gap-1 h-full min-h-0" style={{ gridTemplateRows: `repeat(${weekRows}, minmax(0, 1fr))` }}>
              {gridCells.map((c, idx) => {
                if (!c.day) return <div key={idx} className="rounded-xl bg-transparent min-h-0" />;
                const st = displayMonthData?.days?.[c.day - 1]?.status ?? null;
                const cls = st === 'safe' ? 'border-emerald-200 bg-emerald-50'
                  : st === 'near_miss' ? 'border-amber-200 bg-amber-50'
                  : st === 'accident' ? 'border-rose-200 bg-rose-50'
                  : 'border-slate-200 bg-white';
                const statusText = st === 'safe' ? 'SAFE' : st === 'near_miss' ? 'NEAR MISS' : st === 'accident' ? 'ACCIDENT' : 'NOT SET';
                const statusTone = st === 'safe'
                  ? 'text-emerald-700 bg-emerald-100/80 border-emerald-200'
                  : st === 'near_miss'
                  ? 'text-amber-800 bg-amber-100/80 border-amber-200'
                  : st === 'accident'
                  ? 'text-rose-700 bg-rose-100/80 border-rose-200'
                  : 'text-slate-500 bg-slate-50 border-slate-200';
                return (
                  <button
                    key={idx}
                    type="button"
                    onClick={() => cycleDayStatus(c.day!)}
                    className={`rounded-xl border p-1 md:p-1.5 flex flex-col gap-1 text-left cursor-pointer hover:shadow-sm transition-shadow min-h-0 overflow-hidden ${cls}`}
                    title="คลิกเพื่อเปลี่ยนสถานะ"
                  >
                    <div className="flex items-center justify-between gap-1 shrink-0">
                      <div className="font-extrabold text-slate-900" style={{ fontSize: scaledPx(13, compactScale, 10, 16) }}>{c.day}</div>
                      {st === 'safe' && <CheckCircle2 className="h-3.5 w-3.5 text-emerald-700 shrink-0" />}
                      {st === 'near_miss' && <AlertTriangle className="h-3.5 w-3.5 text-amber-700 shrink-0" />}
                      {st === 'accident' && <AlertTriangle className="h-3.5 w-3.5 text-rose-700 shrink-0" />}
                    </div>
                    <div className={`mt-auto rounded-md border px-1 py-0.5 font-bold text-center leading-tight ${statusTone}`} style={{ fontSize: scaledPx(10, compactScale, 8, 12) }}>
                      {statusText}
                    </div>
                  </button>
                );
              })}
            </div>
          </div>
        </Card>
      );
    }

    if (panel === 'streak') {
      return (
        <Card title="Safety Streak" icon={<Flame className="h-5 w-5 text-emerald-700" />} tone="green" panelScale={panelScale}>
          <div className="h-full flex flex-col items-center justify-center text-center">
            <div className="font-bold text-slate-600" style={{ fontSize: scaledPx(14, panelScale, 12, 18) }}>Zero Accident Days</div>
            <div className="mt-2 font-extrabold text-emerald-700 leading-none" style={{ fontSize: scaledPx(84, panelScale, 46, 120) }}>{safetyStreak}</div>
            <div className="mt-2 font-semibold text-slate-700" style={{ fontSize: scaledPx(16, panelScale, 12, 22) }}>days</div>
            <div className="mt-4 inline-flex items-center gap-2 rounded-2xl bg-sky-50 border border-sky-100 px-4 py-2">
              <Shield className="h-4 w-4 text-sky-700" />
              <span className="text-sm font-bold text-slate-700">Zero Accident Workplace</span>
            </div>
          </div>
        </Card>
      );
    }

    if (panel === 'policy') {
      return (
        <Card
          title={policyTitle}
          icon={<Shield className="h-5 w-5 text-sky-700" />}
          tone="sky"
          panelScale={panelScale}
          actions={
            !editPolicy ? (
              <button onClick={startEditPolicy} className="p-2 rounded-lg hover:bg-white/70" title="Edit">
                <Edit className="h-4 w-4 text-slate-600" />
              </button>
            ) : (
              <>
                <button onClick={savePolicy} className="p-2 rounded-lg hover:bg-emerald-50" title="Save">
                  <Save className="h-4 w-4 text-emerald-700" />
                </button>
                <button onClick={() => setEditPolicy(false)} className="p-2 rounded-lg hover:bg-white/70" title="Cancel">
                  <X className="h-4 w-4 text-slate-600" />
                </button>
              </>
            )
          }
        >
          {!editPolicy ? (
            <ul className="space-y-2 h-full overflow-hidden">
              {policyLines.map((l, i) => (
                <li key={i} className="flex items-start gap-3">
                  <div className="mt-1 h-2 w-2 rounded-full bg-sky-500" />
                  <div className="text-slate-800 font-medium leading-snug line-clamp-2" style={{ fontSize: scaledPx(14, panelScale, 11, 16) }}>{l}</div>
                </li>
              ))}
            </ul>
          ) : (
            <div className="space-y-3">
              <label className="block text-sm font-semibold text-slate-700">Title</label>
              <input value={policyTitleDraft} onChange={(e)=>setPolicyTitleDraft(e.target.value)} className="w-full rounded-xl border border-slate-200 px-3 py-2" />
              <label className="block text-sm font-semibold text-slate-700">Lines (one per line)</label>
              <textarea value={policyLinesDraft} onChange={(e)=>setPolicyLinesDraft(e.target.value)} className="w-full rounded-xl border border-slate-200 p-3" rows={6} />
            </div>
          )}
        </Card>
      );
    }

    return (
      <Card
        title="Company Policy Poster"
        icon={<ImageIcon className="h-5 w-5 text-amber-700" />}
        tone="amber"
        panelScale={panelScale}
        actions={
          <>
            <input ref={fileInputRef} type="file" accept="image/*" className="hidden" onChange={(e) => onPosterSelected(e.target.files?.[0])} />
            <button onClick={() => fileInputRef.current?.click()} className="p-2 rounded-lg hover:bg-amber-50" title="Upload">
              <Upload className="h-4 w-4 text-amber-700" />
            </button>
            {policyPoster ? (
              <>
                <button onClick={() => setPosterZoom((z) => clamp(z - 0.1, 0.5, 2.5))} className="p-2 rounded-lg hover:bg-white border border-transparent hover:border-amber-200" title="Zoom out">
                  <span className="text-amber-800 font-bold">−</span>
                </button>
                <button onClick={() => setPosterZoom(1)} className="px-2 py-1.5 rounded-lg text-xs font-bold border border-amber-200 bg-white hover:bg-amber-50" title="Reset size">
                  {Math.round(posterZoom * 100)}%
                </button>
                <button onClick={() => setPosterZoom((z) => clamp(z + 0.1, 0.5, 2.5))} className="p-2 rounded-lg hover:bg-white border border-transparent hover:border-amber-200" title="Zoom in">
                  <span className="text-amber-800 font-bold">+</span>
                </button>
                <button onClick={() => setPolicyPoster(null)} className="p-2 rounded-lg hover:bg-rose-50" title="Remove">
                  <Trash2 className="h-4 w-4 text-rose-600" />
                </button>
              </>
            ) : null}
          </>
        }
      >
        {!policyPoster ? (
          <div className="h-full flex flex-col items-center justify-center text-center gap-3 rounded-2xl border-2 border-dashed border-slate-200 bg-slate-50 p-6">
            <Upload className="h-8 w-8 text-slate-500" />
            <div className="font-bold text-slate-700">Upload Poster (Vertical)</div>
            <div className="text-sm text-slate-500">แนะนำอัตราส่วน 3:4 หรือ A4 แนวตั้ง</div>
          </div>
        ) : (
          <div className="h-full flex flex-col min-h-0">
            <div className="w-full flex-1 min-h-0 rounded-2xl border border-slate-200 bg-white overflow-hidden flex items-center justify-center relative">
              <div className="absolute left-2 top-2 z-10 rounded-lg border border-amber-200 bg-amber-50/90 px-2 py-1 text-[10px] font-extrabold text-amber-800">
                Zoom: {Math.round(posterZoom * 100)}%
              </div>
              <img
                src={policyPoster}
                alt="Policy Poster"
                className="max-h-full max-w-full object-contain select-none"
                style={{ transform: `scale(${posterZoom})`, transformOrigin: 'center center' }}
              />
            </div>
          </div>
        )}
      </Card>
    );
  }, [
    editSlogan, sloganTh, sloganEn, sloganThDraft, sloganEnDraft,
    editMetrics, metrics, openMetricsEditor, openTrendEditor, trendRows, editPolicy, policyTitle, policyLines, policyTitleDraft, policyLinesDraft,
    policyPoster, posterZoom, announcements, editingAnnId, annDraft, monthSummary, displayMonth, currentYear,
    displayMonthData, gridCells, safetyStreak
  ]);

  return (
    <DndProvider backend={HTML5Backend}>
      <div className="h-screen max-h-screen max-w-[100vw] overflow-hidden flex flex-col bg-[radial-gradient(circle_at_top_left,_#dbeafe_0%,_#f0f9ff_30%,_#ffffff_55%,_#fefce8_80%,_#ecfdf5_100%)] text-slate-900" style={{ fontSize: 'var(--font-size)' }}>
        <header className="px-6 py-4 flex items-center gap-4 rounded-b-3xl border-b border-white/70 bg-white/70 backdrop-blur supports-[backdrop-filter]:bg-white/60 shadow-sm">
          <div className="flex items-center gap-3 shrink-0">
            <div className="h-12 flex items-center rounded-2xl bg-white/80 border border-slate-200 shadow-sm px-3 overflow-hidden">
              <img src={nhkLogo} alt="NHK SPRING (THAILAND)" className="h-9 w-auto" />
            </div>
            <div className="text-2xl font-extrabold">Safety Dashboard</div>
          </div>

          {/* Centered date/time (with seconds) */}
          <div className="flex-1 flex justify-center">
            <HeaderClock centered />
          </div>

          <div className="flex items-center gap-3 shrink-0">
            <div className="flex items-center gap-1 rounded-2xl border border-slate-200 bg-white px-3 py-2 shadow-sm">
              <button
                type="button"
                onClick={() => setUiScale((s) => clamp(Number((s - 0.05).toFixed(2)), 0.8, 1.4))}
                className="h-8 w-10 rounded-xl border border-slate-200 bg-white font-extrabold hover:bg-slate-50"
                aria-label="Decrease font size"
                title="ลดขนาดตัวอักษร"
              >
                A-
              </button>
              <div className="w-14 text-center text-xs font-extrabold text-slate-600">{Math.round(uiScale * 100)}%</div>
              <button
                type="button"
                onClick={() => setUiScale((s) => clamp(Number((s + 0.05).toFixed(2)), 0.8, 1.4))}
                className="h-8 w-10 rounded-xl border border-slate-200 bg-white font-extrabold hover:bg-slate-50"
                aria-label="Increase font size"
                title="เพิ่มขนาดตัวอักษร"
              >
                A+
              </button>
              <button
                type="button"
                onClick={() => setUiScale(1)}
                className="h-8 px-3 rounded-xl border border-slate-200 bg-white text-xs font-extrabold hover:bg-slate-50"
                aria-label="Reset font size"
                title="รีเซ็ตขนาดตัวอักษร"
              >
                Reset
              </button>
            </div>
            <button
              type="button"
              onClick={() => setLayoutLocked((v) => !v)}
              className={`px-4 py-2 rounded-2xl border font-extrabold flex items-center gap-2 ${layoutLocked ? 'border-slate-200 bg-white hover:bg-slate-50' : 'border-sky-200 bg-sky-50 hover:bg-sky-100'}`}
              title={layoutLocked ? 'Unlock layout to move panels' : 'Lock layout'}
            >
              {layoutLocked ? <Lock className="h-4 w-4" /> : <Unlock className="h-4 w-4" />}
              {layoutLocked ? 'LOCKED' : 'UNLOCKED'}
            </button>
            <button type="button" onClick={resetLayout} className="px-4 py-2 rounded-2xl border border-slate-200 bg-white font-extrabold hover:bg-slate-50" title="Reset layout">
              Reset Layout
            </button>
          </div>
        </header>

        <main className="px-6 pt-2 pb-4 flex-1 min-h-0 overflow-hidden">
          {/* Use flex-grow weights instead of fixed % widths/heights so panels never overflow beyond screen bounds */}
          <div className="flex w-full h-full max-w-full max-h-full gap-2 overflow-hidden">
            <div className="flex flex-col gap-2 min-w-0" style={{ flex: `${layout.cols[0]} 1 0%` }}>
              <div className="min-h-0 min-w-0" style={{ flex: `${layout.leftRows[0]} 1 0%` }}><DashboardSlot slot="leftTop" panel={slots.leftTop} layoutLocked={layoutLocked} onSwap={swapSlots} renderPanel={renderPanel} /></div>
              <Splitter orientation="horizontal" onMouseDown={onResizeLeft(0)} />
              <div className="min-h-0 min-w-0" style={{ flex: `${layout.leftRows[1]} 1 0%` }}><DashboardSlot slot="leftBottom" panel={slots.leftBottom} layoutLocked={layoutLocked} onSwap={swapSlots} renderPanel={renderPanel} /></div>
            </div>

            <Splitter orientation="vertical" onMouseDown={onResizeCols(0)} />

            <div className="flex flex-col gap-2 min-w-0" style={{ flex: `${layout.cols[1]} 1 0%` }}>
              <div className="min-h-0 min-w-0" style={{ flex: `${layout.centerRows[0]} 1 0%` }}><DashboardSlot slot="centerTop" panel={slots.centerTop} layoutLocked={layoutLocked} onSwap={swapSlots} renderPanel={renderPanel} /></div>
              <Splitter orientation="horizontal" onMouseDown={onResizeCenter(0)} />
              <div className="min-h-0 min-w-0" style={{ flex: `${layout.centerRows[1]} 1 0%` }}><DashboardSlot slot="centerMid" panel={slots.centerMid} layoutLocked={layoutLocked} onSwap={swapSlots} renderPanel={renderPanel} /></div>
              <Splitter orientation="horizontal" onMouseDown={onResizeCenter(1)} />
              <div className="min-h-0 min-w-0" style={{ flex: `${layout.centerRows[2]} 1 0%` }}><DashboardSlot slot="centerBottom" panel={slots.centerBottom} layoutLocked={layoutLocked} onSwap={swapSlots} renderPanel={renderPanel} /></div>
            </div>

            <Splitter orientation="vertical" onMouseDown={onResizeCols(1)} />

            <div className="flex flex-col gap-2 min-w-0" style={{ flex: `${layout.cols[2]} 1 0%` }}>
              <div className="min-h-0 min-w-0" style={{ flex: `${layout.rightRows[0]} 1 0%` }}><DashboardSlot slot="rightTop" panel={slots.rightTop} layoutLocked={layoutLocked} onSwap={swapSlots} renderPanel={renderPanel} /></div>
              <Splitter orientation="horizontal" onMouseDown={onResizeRight(0)} />
              <div className="min-h-0 min-w-0" style={{ flex: `${layout.rightRows[1]} 1 0%` }}><DashboardSlot slot="rightBottom" panel={slots.rightBottom} layoutLocked={layoutLocked} onSwap={swapSlots} renderPanel={renderPanel} /></div>
            </div>
          </div>
        </main>
        <Dialog open={editMetrics} onOpenChange={(open) => { if (!open) closeMetricsEditor(); }}>
          <DialogContent className="w-[min(980px,96vw)] max-w-none">
            <DialogHeader>
              <DialogTitle className="text-slate-900 font-extrabold">Safety Data Settings</DialogTitle>
              <div className="text-sm text-slate-600 font-semibold">
                เพิ่ม/ลบ/แก้ไขหัวข้อได้สะดวก • ใช้เมาส์/คีย์บอร์ดได้ • กด Save เพื่อบันทึก
              </div>
            </DialogHeader>

            <div className="flex items-center justify-between gap-2">
              <button
                type="button"
                onClick={addMetricDraft}
                className="px-4 py-2 rounded-xl bg-cyan-600 text-white font-extrabold hover:bg-cyan-700 inline-flex items-center gap-2"
              >
                <Plus className="h-4 w-4" /> เพิ่มหัวข้อ
              </button>
              <div className="text-xs font-bold text-slate-500">รายการทั้งหมด: {metricsDraft.length}</div>
            </div>

            <ScrollArea className="h-[60vh] rounded-2xl border border-slate-200 bg-slate-50/40 p-3 tv-scroll">
              <div className="space-y-2">
                {metricsDraft.map((m, idx) => (
                  <div key={m.id} className="rounded-2xl border border-slate-200 bg-white p-3 shadow-sm">
                    <div className="grid gap-2 items-end" style={{ gridTemplateColumns: 'minmax(220px,1.6fr) minmax(110px,0.5fr) minmax(90px,0.45fr) auto' }}>
                      <div className="min-w-0">
                        <label className="block text-xs font-extrabold text-slate-500 mb-1">หัวข้อ #{idx + 1}</label>
                        <input
                          value={m.label}
                          onChange={(e) => updateMetricDraft(m.id, { label: e.target.value })}
                          className="w-full rounded-xl border border-slate-200 px-3 py-2 bg-white focus:outline-none focus:ring-2 focus:ring-cyan-200"
                          placeholder="เช่น PPE Compliance"
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-extrabold text-slate-500 mb-1">ค่า</label>
                        <input
                          value={m.value}
                          onChange={(e) => updateMetricDraft(m.id, { value: e.target.value })}
                          className="w-full rounded-xl border border-slate-200 px-3 py-2 text-right bg-white focus:outline-none focus:ring-2 focus:ring-cyan-200"
                          placeholder="0"
                          inputMode="decimal"
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-extrabold text-slate-500 mb-1">หน่วย</label>
                        <input
                          value={m.unit || ''}
                          onChange={(e) => updateMetricDraft(m.id, { unit: e.target.value })}
                          className="w-full rounded-xl border border-slate-200 px-3 py-2 bg-white focus:outline-none focus:ring-2 focus:ring-cyan-200"
                          placeholder="%, case"
                        />
                      </div>
                      <div className="flex items-end justify-end">
                        <button
                          type="button"
                          onClick={() => deleteMetricDraft(m.id)}
                          className="h-10 px-3 rounded-xl border border-rose-200 bg-rose-50 hover:bg-rose-100 text-rose-700 font-extrabold inline-flex items-center gap-1"
                          aria-label="Delete metric"
                          title="ลบหัวข้อ"
                        >
                          <Trash2 className="h-4 w-4" /> ลบ
                        </button>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </ScrollArea>

            <DialogFooter className="gap-2">
              <button type="button" onClick={closeMetricsEditor} className="px-4 py-2 rounded-xl border border-slate-200 bg-white font-extrabold hover:bg-slate-50">
                Cancel
              </button>
              <button type="button" onClick={saveMetricsEditor} className="px-4 py-2 rounded-xl bg-emerald-600 text-white font-extrabold hover:bg-emerald-700 inline-flex items-center gap-2">
                <Save className="h-4 w-4" /> Save
              </button>
            </DialogFooter>
          </DialogContent>
        </Dialog>

        <Dialog open={editTrend} onOpenChange={(open) => { if (!open) closeTrendEditor(); }}>
          <DialogContent className="w-[min(980px,96vw)] max-w-none">
            <DialogHeader>
              <DialogTitle className="text-slate-900 font-extrabold">Safety Trend (Graph) Settings</DialogTitle>
              <div className="text-sm text-slate-600 font-semibold">
                ใส่ข้อมูล Case ย้อนหลัง 4–5 ปีเพื่อสร้างกราฟด้านล่าง • กด Save เพื่อบันทึก
              </div>
            </DialogHeader>

            <div className="flex flex-wrap items-center justify-between gap-2">
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={addTrendYear}
                  className="px-4 py-2 rounded-xl bg-sky-600 text-white font-extrabold hover:bg-sky-700 inline-flex items-center gap-2"
                >
                  <Plus className="h-4 w-4" /> เพิ่มปี
                </button>
                <button
                  type="button"
                  onClick={() => setTrendDraft(defaultTrendRows(currentYear))}
                  className="px-4 py-2 rounded-xl border border-slate-200 bg-white font-extrabold hover:bg-slate-50"
                >
                  Reset 5Y
                </button>
              </div>
              <div className="text-xs font-bold text-slate-500">ปีทั้งหมด: {trendDraft.length}</div>
            </div>

            <ScrollArea className="h-[60vh] rounded-2xl border border-slate-200 bg-slate-50/40 p-3 tv-scroll">
              <div className="space-y-2">
                <div className="grid gap-2 px-2 text-[11px] font-extrabold text-slate-500" style={{ gridTemplateColumns: '90px repeat(4, minmax(90px,1fr)) auto' }}>
                  <div>Year</div>
                  <div>First Aid</div>
                  <div>Non-Absent</div>
                  <div>Absent</div>
                  <div>Fire</div>
                  <div className="text-right">Action</div>
                </div>
                {trendDraft.map((r, idx) => (
                  <div key={`${r.year}-${idx}`} className="rounded-2xl border border-slate-200 bg-white p-3 shadow-sm">
                    <div className="grid gap-2 items-end" style={{ gridTemplateColumns: '90px repeat(4, minmax(90px,1fr)) auto' }}>
                      <div>
                        <label className="block text-[10px] font-extrabold text-slate-500 mb-1">Year</label>
                        <input
                          value={String(r.year)}
                          onChange={(e) => updateTrendRow(idx, { year: Number(e.target.value) })}
                          className="w-full rounded-xl border border-slate-200 px-3 py-2 bg-white focus:outline-none focus:ring-2 focus:ring-sky-200"
                          inputMode="numeric"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-extrabold text-slate-500 mb-1">case</label>
                        <input
                          value={String(r.firstAid)}
                          onChange={(e) => updateTrendRow(idx, { firstAid: Number(e.target.value) })}
                          className="w-full rounded-xl border border-slate-200 px-3 py-2 text-right bg-white focus:outline-none focus:ring-2 focus:ring-sky-200"
                          inputMode="numeric"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-extrabold text-slate-500 mb-1">case</label>
                        <input
                          value={String(r.nonAbsent)}
                          onChange={(e) => updateTrendRow(idx, { nonAbsent: Number(e.target.value) })}
                          className="w-full rounded-xl border border-slate-200 px-3 py-2 text-right bg-white focus:outline-none focus:ring-2 focus:ring-sky-200"
                          inputMode="numeric"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-extrabold text-slate-500 mb-1">case</label>
                        <input
                          value={String(r.absent)}
                          onChange={(e) => updateTrendRow(idx, { absent: Number(e.target.value) })}
                          className="w-full rounded-xl border border-slate-200 px-3 py-2 text-right bg-white focus:outline-none focus:ring-2 focus:ring-sky-200"
                          inputMode="numeric"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-extrabold text-slate-500 mb-1">case</label>
                        <input
                          value={String(r.fire)}
                          onChange={(e) => updateTrendRow(idx, { fire: Number(e.target.value) })}
                          className="w-full rounded-xl border border-slate-200 px-3 py-2 text-right bg-white focus:outline-none focus:ring-2 focus:ring-sky-200"
                          inputMode="numeric"
                        />
                      </div>
                      <div className="flex items-end justify-end">
                        <button
                          type="button"
                          onClick={() => deleteTrendRow(idx)}
                          className="h-10 px-3 rounded-xl border border-rose-200 bg-rose-50 hover:bg-rose-100 text-rose-700 font-extrabold inline-flex items-center gap-1"
                          aria-label="Delete year"
                          title="ลบปีนี้"
                        >
                          <Trash2 className="h-4 w-4" /> ลบ
                        </button>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </ScrollArea>

            <DialogFooter className="gap-2">
              <button type="button" onClick={closeTrendEditor} className="px-4 py-2 rounded-xl border border-slate-200 bg-white font-extrabold hover:bg-slate-50">
                Cancel
              </button>
              <button type="button" onClick={saveTrendEditor} className="px-4 py-2 rounded-xl bg-emerald-600 text-white font-extrabold hover:bg-emerald-700 inline-flex items-center gap-2">
                <Save className="h-4 w-4" /> Save
              </button>
            </DialogFooter>
          </DialogContent>
        </Dialog>

      </div>
    </DndProvider>
  );
}
