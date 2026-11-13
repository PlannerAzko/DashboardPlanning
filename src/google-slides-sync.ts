import { initializeApp, getApps, getApp } from 'firebase/app';
import {
  getAuth,
  signInWithPopup,
  GoogleAuthProvider,
  onAuthStateChanged,
  User,
} from 'firebase/auth';
import PptxGenJS from 'pptxgenjs';
import firebaseConfig from '../firebase-applet-config.json';

const SCOPES = ['https://www.googleapis.com/auth/presentations'];
const STORAGE_KEY_PRES_ID = 'ahi_synced_presentation_id';
const STORAGE_KEY_PRES_TITLE = 'ahi_synced_presentation_title';
const STORAGE_KEY_WATCH_SYNC = 'ahi_watch_spreadsheet_sync';

const app = getApps().length ? getApp() : initializeApp(firebaseConfig);
const auth = getAuth(app);

const provider = new GoogleAuthProvider();
SCOPES.forEach((scope) => provider.addScope(scope));
provider.setCustomParameters({ prompt: 'select_account' });

// Strictly in-memory OAuth token cache (never stored in localStorage/sessionStorage)
let cachedAccessToken: string | null = null;
let currentUser: User | null = null;
let isSigningIn = false;
let pendingConfirmAction: (() => Promise<void>) | null = null;
let lastSyncedDataSignature: string | null = null;

interface ArmadaRow {
  type: string;
  plans: string;
  totalCbm: string;
  avgCbm: string;
  pct: string;
}

interface StoreAgingRow {
  store: string;
  shipTo: string;
  sudahPlan: string;
  belumPlan: string;
}

interface CityVolumeRow {
  kota: string;
  cbm: number;
  plans: number;
}

interface AhiSnapshot {
  generatedAt: string;
  lastUpdate: string;
  filtersText: string;
  planning: {
    caseIdPlan: string;
    caseIdShipped: string;
    caseIdRatio: string;
    cbmPlan: string;
    cbmShipped: string;
    cbmRatio: string;
    shift1Cbm: string;
    shift2Cbm: string;
    slaCustomer: string;
    slaGrw: string;
    batches: Array<{ batch: string; caseId: string; cbm: string }>;
    armadas: ArmadaRow[];
  };
  outstanding: {
    totalBacklog: string;
    sudahPlan: string;
    belumPlan: string;
    aging7Up: string;
    allocUnsuccessful: string;
    topBelumPlanStores: StoreAgingRow[];
  };
  vendorAndHub: {
    vendorSupporting: string;
    completedLc: string;
    onTimeLc: string;
    terlambatLc: string;
    delayedLc: string;
    hubSla: string;
    hubSudahLoading: string;
    hubBelumLoading: string;
    astorSla: string;
    astorHit: string;
    astorMiss: string;
    astorOs: string;
    astorOsCbm: string;
    astorOccupancy: string;
  };
  summary: {
    avgSlaCustomer: string;
    avgSlaGrw: string;
    cbmRatio: string;
    totalCbmShipped: string;
    topCities: CityVolumeRow[];
    bottomCities: CityVolumeRow[];
    insights: string[];
  };
  signature: string;
}

function getElText(id: string, fallback = '-'): string {
  const el = document.getElementById(id);
  const txt = el?.textContent?.trim();
  return txt && txt.length > 0 ? txt : fallback;
}

function parseIndoNum(val: string): number {
  if (!val) return 0;
  const cleaned = val.toString().replace(/\./g, '').replace(/,/g, '.').replace(/[^0-9.-]/g, '');
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

function fmtId(num: number, decimals = 2): string {
  return num.toLocaleString('id-ID', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
}

function extractPresentationId(input: string): string {
  const raw = (input || '').trim();
  if (!raw) return '';
  const match = raw.match(/\/presentation\/d\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) return match[1];
  return raw;
}

function safeStr(val: string | number | undefined | null, fallback = '-'): string {
  const s = val === undefined || val === null ? '' : String(val).trim();
  return s.length > 0 ? s : fallback;
}

function collectAhiDashboardSnapshot(): AhiSnapshot {
  const win = window as any;
  const nowStr = new Date().toLocaleString('id-ID', {
    day: '2-digit',
    month: 'short',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });

  const lastUpdate = getElText('last-update-text', nowStr)
    .replace(/^.*LAST UPDATE\s*[:|-]?\s*/i, '')
    .trim() || nowStr;

  const bulanDisp = getElText('bulan-display', 'Semua Bulan');
  const dateDisp = getElText('stuffingDate-display', 'Semua Tanggal');
  const regDisp = getElText('regional-display', 'Semua Regional');
  const areaDisp = getElText('area-display', 'Semua Area');
  const kotaDisp = getElText('kota-display', 'Semua Kota');

  const filtersText = `Bulan: ${bulanDisp} | Tanggal: ${dateDisp} | Regional: ${regDisp} | Area: ${areaDisp} | Kota: ${kotaDisp}`;

  // Top & bottom cities from filteredData if available
  const cityMap: Record<string, { cbm: number; plans: number }> = {};
  const filteredRows: any[] = Array.isArray(win.filteredData) ? win.filteredData : [];
  filteredRows.forEach((row) => {
    const kota = (row?.AR || '').toString().trim();
    if (!kota || kota === '-') return;
    const rawZ = (row?.Z || '0').toString().replace(/,/g, '.').replace(/[^0-9.-]/g, '');
    const cbm = parseFloat(rawZ) || 0;
    if (!cityMap[kota]) cityMap[kota] = { cbm: 0, plans: 0 };
    cityMap[kota].cbm += cbm;
    cityMap[kota].plans += 1;
  });

  const cityList: CityVolumeRow[] = Object.entries(cityMap).map(([kota, stat]) => ({
    kota,
    cbm: stat.cbm,
    plans: stat.plans,
  }));
  const topCities = [...cityList].sort((a, b) => b.cbm - a.cbm).slice(0, 6);
  const bottomCities = [...cityList].filter((c) => c.cbm > 0).sort((a, b) => a.cbm - b.cbm).slice(0, 6);

  // Top stores Belum Plan from pivot-table-body DOM
  const topBelumPlanStores: StoreAgingRow[] = [];
  const pivotRows = document.querySelectorAll('#pivot-table-body tr');
  pivotRows.forEach((tr) => {
    const cells = tr.querySelectorAll('td');
    if (cells.length >= 12 && !tr.classList.contains('total-row')) {
      const store = (cells[0]?.textContent || '').trim();
      const shipTo = (cells[1]?.textContent || '').trim();
      const sudahPlan = (cells[6]?.textContent || '0').trim();
      const belumPlan = (cells[11]?.textContent || '0').trim();
      if (store && !store.toUpperCase().includes('TOTAL') && parseIndoNum(belumPlan) > 0) {
        topBelumPlanStores.push({ store, shipTo, sudahPlan, belumPlan });
      }
    }
  });
  topBelumPlanStores.sort((a, b) => parseIndoNum(b.belumPlan) - parseIndoNum(a.belumPlan));

  const planning = {
    caseIdPlan: getElText('caseIdPlan', '0'),
    caseIdShipped: getElText('caseIdShipped', '0'),
    caseIdRatio: getElText('caseIdRatioShipped', '0%'),
    cbmPlan: getElText('cbmPlan', '0'),
    cbmShipped: getElText('cbmShipped', '0'),
    cbmRatio: getElText('cbmRatioShipped', '0%'),
    shift1Cbm: getElText('cbmPlanShift1', '0'),
    shift2Cbm: getElText('cbmPlanShift2', '0'),
    slaCustomer: getElText('slaCustomerKPI', '0%'),
    slaGrw: getElText('slaGrwKPI', '0%'),
    batches: [
      { batch: 'Batch 1 (08:00-10:59)', caseId: getElText('caseIdPlanB1', '0'), cbm: getElText('cbmPlanB1', '0') },
      { batch: 'Batch 2 (13:00-13:59)', caseId: getElText('caseIdPlanB2', '0'), cbm: getElText('cbmPlanB2', '0') },
      { batch: 'Batch 3 (16:00-18:59)', caseId: getElText('caseIdPlanB3', '0'), cbm: getElText('cbmPlanB3', '0') },
      { batch: 'Batch 4 (19:00-21:59)', caseId: getElText('caseIdPlanB4', '0'), cbm: getElText('cbmPlanB4', '0') },
    ],
    armadas: [
      {
        type: 'TOTAL PLAN',
        plans: getElText('armada-total-plans', '0'),
        totalCbm: getElText('armada-total-cbm', '0'),
        avgCbm: getElText('armada-average-cbm', '0'),
        pct: getElText('armada-total-pct', '100%'),
      },
      {
        type: 'Cont 20',
        plans: getElText('kpi-cont20-plans', '0'),
        totalCbm: getElText('kpi-cont20-totalcbm', '0'),
        avgCbm: getElText('kpi-cont20-avgcbm', '0'),
        pct: getElText('kpi-cont20-pct', '0%'),
      },
      {
        type: 'Cont 40',
        plans: getElText('kpi-cont40-plans', '0'),
        totalCbm: getElText('kpi-cont40-totalcbm', '0'),
        avgCbm: getElText('kpi-cont40-avgcbm', '0'),
        pct: getElText('kpi-cont40-pct', '0%'),
      },
      {
        type: 'Wingbox',
        plans: getElText('kpi-wingbox-plans', '0'),
        totalCbm: getElText('kpi-wingbox-totalcbm', '0'),
        avgCbm: getElText('kpi-wingbox-avgcbm', '0'),
        pct: getElText('kpi-wingbox-pct', '0%'),
      },
      {
        type: 'Fuso',
        plans: getElText('kpi-fuso-plans', '0'),
        totalCbm: getElText('kpi-fuso-totalcbm', '0'),
        avgCbm: getElText('kpi-fuso-avgcbm', '0'),
        pct: getElText('kpi-fuso-pct', '0%'),
      },
      {
        type: 'CDD Long',
        plans: getElText('kpi-cdd-plans', '0'),
        totalCbm: getElText('kpi-cdd-totalcbm', '0'),
        avgCbm: getElText('kpi-cdd-avgcbm', '0'),
        pct: getElText('kpi-cdd-pct', '0%'),
      },
      {
        type: 'CDE Long',
        plans: getElText('kpi-cde-plans', '0'),
        totalCbm: getElText('kpi-cde-totalcbm', '0'),
        avgCbm: getElText('kpi-cde-avgcbm', '0'),
        pct: getElText('kpi-cde-pct', '0%'),
      },
    ],
  };

  const outstanding = {
    totalBacklog: getElText('kpi-backlog-outstanding', '0'),
    sudahPlan: getElText('kpi-backlog-sudah', '0'),
    belumPlan: getElText('kpi-backlog-belum', '0'),
    aging7Up: getElText('kpi-backlog-aging7', '0'),
    allocUnsuccessful: getElText('kpi-alloc-unsuccessful', '0'),
    topBelumPlanStores: topBelumPlanStores.slice(0, 7),
  };

  const vendorAndHub = {
    vendorSupporting: getElText('kpi-supporting-percent', '0%'),
    completedLc: getElText('kpi-total-lc', '0'),
    onTimeLc: getElText('kpi-ontime-lc', '0'),
    terlambatLc: getElText('kpi-terlambat-lc', '0'),
    delayedLc: getElText('kpi-delayed-lc', '0'),
    hubSla:
      typeof win.hubSlaRateLast === 'number'
        ? `${win.hubSlaRateLast.toFixed(1)}%`
        : getElText('hub-sla-pct', '0%'),
    hubSudahLoading: getElText('hub-sudah-loading-cbm', '0 CBM'),
    hubBelumLoading: getElText('hub-belum-loading-cbm', '0 CBM'),
    astorSla:
      typeof win.astorSlaRateLast === 'number'
        ? `${win.astorSlaRateLast.toFixed(1)}%`
        : getElText('astor-pct', '0%'),
    astorHit: getElText('astor-hit', '0'),
    astorMiss: getElText('astor-miss', '0'),
    astorOs: getElText('astor-outstanding', '0'),
    astorOsCbm: getElText('astor-outstanding-cbm', '0'),
    astorOccupancy: getElText('astor-occupancy-pct', '0%'),
  };

  const avgSlaCustomer =
    getElText('summary-kpi-sla-customer', '--') !== '--' &&
    getElText('summary-kpi-sla-customer', '--') !== '100.0%'
      ? getElText('summary-kpi-sla-customer')
      : planning.slaCustomer;
  const avgSlaGrw =
    getElText('summary-kpi-sla-grw', '-- %').includes('--')
      ? planning.slaGrw
      : getElText('summary-kpi-sla-grw');
  const cbmRatio =
    getElText('summary-kpi-cbm-ratio', '-- %').includes('--')
      ? planning.cbmRatio
      : getElText('summary-kpi-cbm-ratio');
  const totalCbmShipped =
    getElText('summary-kpi-cbm-shipped', '-- CBM').includes('--')
      ? `${planning.cbmShipped} CBM`
      : getElText('summary-kpi-cbm-shipped');

  const belumNum = parseIndoNum(outstanding.belumPlan);
  const aging7Num = parseIndoNum(outstanding.aging7Up);
  const terlambatNum = parseIndoNum(vendorAndHub.terlambatLc);

  const insights = [
    `Realisasi Order Planning NDC AHI Sidoarjo mencapai ${planning.cbmShipped} CBM dari target ${planning.cbmPlan} CBM (${planning.cbmRatio} fulfillment) dengan ${planning.armadas[0].plans} total rencana armada.`,
    `Performa ketepatan waktu tercatat pada SLA Customer ${planning.slaCustomer} dan SLA GRW ${planning.slaGrw}, didukung Vendor Supporting ${vendorAndHub.vendorSupporting} (${vendorAndHub.onTimeLc}/${vendorAndHub.completedLc} LC On-Time).`,
    belumNum > 0
      ? `Prioritas Outstanding: Terdapat ${outstanding.belumPlan} CBM berstatus Belum Plan (Aging >7 Hari: ${outstanding.aging7Up} CBM) yang perlu segera dialokasikan.`
      : `Outstanding Belum Plan dalam kondisi terkendali (${outstanding.belumPlan} CBM).`,
    `Monitoring HUB GRW SLA di angka ${vendorAndHub.hubSla} dan SLA ASTOR ALL BU di angka ${vendorAndHub.astorSla} (HIT: ${vendorAndHub.astorHit} | MISS: ${vendorAndHub.astorMiss} | Line Occupancy: ${vendorAndHub.astorOccupancy}).`,
    terlambatNum > 0 || aging7Num > 0
      ? `Tindak lanjut armada terlambat (${vendorAndHub.terlambatLc} LC) dan percepatan alokasi kuota store aging >7 hari sangat disarankan.`
      : `Operasional armada dan clearance backlog berjalan optimal sesuai jadwal.`,
  ];

  const signature = [
    planning.caseIdPlan,
    planning.cbmPlan,
    planning.cbmShipped,
    outstanding.totalBacklog,
    outstanding.belumPlan,
    vendorAndHub.astorHit,
    vendorAndHub.astorMiss,
    lastUpdate,
  ].join('|');

  return {
    generatedAt: nowStr,
    lastUpdate,
    filtersText,
    planning,
    outstanding,
    vendorAndHub,
    summary: {
      avgSlaCustomer,
      avgSlaGrw,
      cbmRatio,
      totalCbmShipped,
      topCities,
      bottomCities,
      insights,
    },
    signature,
  };
}

// =========================================================
// GOOGLE SLIDES API REQUEST BUILDER (Widescreen 720 x 405 PT)
// =========================================================
function hexToRgb(hex: string): { red: number; green: number; blue: number } {
  const clean = hex.replace('#', '');
  const r = parseInt(clean.substring(0, 2), 16) / 255;
  const g = parseInt(clean.substring(2, 4), 16) / 255;
  const b = parseInt(clean.substring(4, 6), 16) / 255;
  return { red: r, green: g, blue: b };
}

function buildGoogleSlidesRequests(snap: AhiSnapshot, oldSlideIds: string[]): any[] {
  const reqs: any[] = [];
  const uid = `ahi_${Date.now().toString(36)}`;
  let seq = 0;
  const nextId = (prefix: string) => `${uid}_${prefix}_${++seq}`;

  const addSlide = (bgHex = '#0B1120') => {
    const slideId = nextId('slide');
    reqs.push({
      createSlide: {
        objectId: slideId,
        slideLayoutReference: { predefinedLayout: 'BLANK' },
      },
    });
    reqs.push({
      updatePageProperties: {
        objectId: slideId,
        pageProperties: {
          pageBackgroundFill: {
            solidFill: { color: { rgbColor: hexToRgb(bgHex) } },
          },
        },
        fields: 'pageBackgroundFill.solidFill.color',
      },
    });
    return slideId;
  };

  const addRect = (
    pageId: string,
    x: number,
    y: number,
    w: number,
    h: number,
    fillHex: string,
    borderHex?: string
  ) => {
    const id = nextId('rect');
    reqs.push({
      createShape: {
        objectId: id,
        shapeType: 'ROUND_RECTANGLE',
        elementProperties: {
          pageObjectId: pageId,
          size: { width: { magnitude: w, unit: 'PT' }, height: { magnitude: h, unit: 'PT' } },
          transform: { scaleX: 1, scaleY: 1, translateX: x, translateY: y, unit: 'PT' },
        },
      },
    });
    reqs.push({
      updateShapeProperties: {
        objectId: id,
        shapeProperties: {
          shapeBackgroundFill: { solidFill: { color: { rgbColor: hexToRgb(fillHex) } } },
          outline: borderHex
            ? {
                outlineFill: { solidFill: { color: { rgbColor: hexToRgb(borderHex) } } },
                weight: { magnitude: 1, unit: 'PT' },
              }
            : { propertyState: 'NOT_RENDERED' },
        },
        fields: 'shapeBackgroundFill.solidFill.color,outline',
      },
    });
    return id;
  };

  const addTextBox = (
    pageId: string,
    text: string,
    x: number,
    y: number,
    w: number,
    h: number,
    opts: {
      fontSize?: number;
      bold?: boolean;
      color?: string;
      align?: 'START' | 'CENTER' | 'END';
      fillHex?: string;
      borderHex?: string;
    } = {}
  ) => {
    const id = nextId('txt');
    const content = safeStr(text, '-');
    reqs.push({
      createShape: {
        objectId: id,
        shapeType: 'TEXT_BOX',
        elementProperties: {
          pageObjectId: pageId,
          size: { width: { magnitude: w, unit: 'PT' }, height: { magnitude: h, unit: 'PT' } },
          transform: { scaleX: 1, scaleY: 1, translateX: x, translateY: y, unit: 'PT' },
        },
      },
    });
    if (opts.fillHex || opts.borderHex) {
      reqs.push({
        updateShapeProperties: {
          objectId: id,
          shapeProperties: {
            ...(opts.fillHex
              ? { shapeBackgroundFill: { solidFill: { color: { rgbColor: hexToRgb(opts.fillHex) } } } }
              : {}),
            outline: opts.borderHex
              ? {
                  outlineFill: { solidFill: { color: { rgbColor: hexToRgb(opts.borderHex) } } },
                  weight: { magnitude: 1, unit: 'PT' },
                }
              : { propertyState: 'NOT_RENDERED' },
          },
          fields: opts.fillHex ? 'shapeBackgroundFill.solidFill.color,outline' : 'outline',
        },
      });
    }
    reqs.push({
      insertText: { objectId: id, text: content, insertionIndex: 0 },
    });
    reqs.push({
      updateTextStyle: {
        objectId: id,
        textRange: { type: 'ALL' },
        style: {
          fontFamily: 'Inter',
          fontSize: { magnitude: opts.fontSize || 11, unit: 'PT' },
          bold: Boolean(opts.bold),
          foregroundColor: { opaqueColor: { rgbColor: hexToRgb(opts.color || '#F8FAFC') } },
        },
        fields: 'fontFamily,fontSize,bold,foregroundColor',
      },
    });
    if (opts.align) {
      reqs.push({
        updateParagraphStyle: {
          objectId: id,
          textRange: { type: 'ALL' },
          style: { alignment: opts.align },
          fields: 'alignment',
        },
      });
    }
    return id;
  };

  const addKpiCard = (
    pageId: string,
    x: number,
    y: number,
    w: number,
    h: number,
    label: string,
    value: string,
    accentHex: string,
    sub = ''
  ) => {
    addRect(pageId, x, y, w, h, '#111827', accentHex);
    addRect(pageId, x, y, 4, h, accentHex);
    addTextBox(pageId, label, x + 10, y + 6, w - 16, 16, {
      fontSize: 8.5,
      bold: true,
      color: '#94A3B8',
    });
    addTextBox(pageId, value, x + 10, y + 22, w - 16, 26, {
      fontSize: 17,
      bold: true,
      color: accentHex,
    });
    if (sub) {
      addTextBox(pageId, sub, x + 10, y + h - 18, w - 16, 14, {
        fontSize: 7.5,
        color: '#CBD5E1',
      });
    }
  };

  const addTable = (
    pageId: string,
    x: number,
    y: number,
    w: number,
    h: number,
    headers: string[],
    rows: string[][],
    headerBgHex = '#991B1B'
  ) => {
    const id = nextId('tbl');
    const totalRows = rows.length + 1;
    const cols = headers.length;
    reqs.push({
      createTable: {
        objectId: id,
        rows: totalRows,
        columns: cols,
        elementProperties: {
          pageObjectId: pageId,
          size: { width: { magnitude: w, unit: 'PT' }, height: { magnitude: h, unit: 'PT' } },
          transform: { scaleX: 1, scaleY: 1, translateX: x, translateY: y, unit: 'PT' },
        },
      },
    });

    const writeCell = (r: number, c: number, text: string, isHeader: boolean) => {
      const val = safeStr(text, '-');
      reqs.push({
        insertText: {
          objectId: id,
          cellLocation: { rowIndex: r, columnIndex: c },
          text: val,
          insertionIndex: 0,
        },
      });
      reqs.push({
        updateTableCellProperties: {
          objectId: id,
          tableRange: {
            location: { rowIndex: r, columnIndex: c },
            rowSpan: 1,
            columnSpan: 1,
          },
          tableCellProperties: {
            tableCellBackgroundFill: {
              solidFill: {
                color: {
                  rgbColor: hexToRgb(isHeader ? headerBgHex : r % 2 === 1 ? '#111827' : '#0F172A'),
                },
              },
            },
            contentAlignment: 'MIDDLE',
          },
          fields: 'tableCellBackgroundFill.solidFill.color,contentAlignment',
        },
      });
      reqs.push({
        updateTextStyle: {
          objectId: id,
          cellLocation: { rowIndex: r, columnIndex: c },
          textRange: { type: 'ALL' },
          style: {
            fontFamily: 'Inter',
            fontSize: { magnitude: isHeader ? 8.5 : 8, unit: 'PT' },
            bold: isHeader || c === 0,
            foregroundColor: {
              opaqueColor: { rgbColor: hexToRgb(isHeader ? '#FACC15' : '#F1F5F9') },
            },
          },
          fields: 'fontFamily,fontSize,bold,foregroundColor',
        },
      });
    };

    headers.forEach((hText, cIdx) => writeCell(0, cIdx, hText, true));
    rows.forEach((rowArr, rIdx) => {
      for (let cIdx = 0; cIdx < cols; cIdx++) {
        writeCell(rIdx + 1, cIdx, rowArr[cIdx] ?? '-', false);
      }
    });
    return id;
  };

  const addSlideChrome = (pageId: string, title: string, subtitle: string, pageNum: number) => {
    addRect(pageId, 0, 0, 720, 48, '#991B1B');
    addRect(pageId, 0, 48, 720, 2.5, '#FACC15');
    addTextBox(pageId, title, 20, 7, 500, 22, {
      fontSize: 14,
      bold: true,
      color: '#FFFFFF',
    });
    addTextBox(pageId, subtitle, 20, 28, 520, 15, {
      fontSize: 8.5,
      color: '#FDE68A',
    });
    addTextBox(pageId, `NDC AHI SIDOARJO  |  Slide ${pageNum}/5`, 540, 15, 160, 18, {
      fontSize: 8.5,
      bold: true,
      color: '#FACC15',
      align: 'END',
    });
    addTextBox(
      pageId,
      `Synced from Spreadsheet • Update: ${snap.lastUpdate} • ${snap.filtersText}`,
      20,
      386,
      680,
      14,
      { fontSize: 7, color: '#64748B' }
    );
  };

  // SLIDE 1: COVER & EXECUTIVE SNAPSHOT
  const s1 = addSlide('#090D16');
  addRect(s1, 0, 0, 14, 405, '#DC2626');
  addRect(s1, 14, 0, 4, 405, '#FACC15');
  addTextBox(s1, 'LOGISTICS OPERATIONS & ORDER PLANNING', 42, 42, 620, 20, {
    fontSize: 11,
    bold: true,
    color: '#FACC15',
  });
  addTextBox(s1, 'NDC AHI SIDOARJO\nPERFORMANCE REPORT', 42, 68, 640, 82, {
    fontSize: 28,
    bold: true,
    color: '#FFFFFF',
  });
  addTextBox(s1, `Sumber Data: Live Dashboard Spreadsheet  |  Update: ${snap.lastUpdate}`, 42, 158, 640, 18, {
    fontSize: 10.5,
    color: '#CBD5E1',
  });
  addTextBox(s1, snap.filtersText, 42, 180, 640, 18, {
    fontSize: 9.5,
    color: '#94A3B8',
  });
  addKpiCard(s1, 42, 224, 152, 76, 'CBM PLAN / SHIPPED', snap.planning.cbmPlan, '#EF4444', `Shipped: ${snap.planning.cbmShipped} (${snap.planning.cbmRatio})`);
  addKpiCard(s1, 206, 224, 152, 76, 'SLA CUSTOMER / GRW', snap.planning.slaCustomer, '#10B981', `SLA GRW: ${snap.planning.slaGrw}`);
  addKpiCard(s1, 370, 224, 152, 76, 'OUTSTANDING BELUM PLAN', snap.outstanding.belumPlan, '#F59E0B', `Total Backlog: ${snap.outstanding.totalBacklog} CBM`);
  addKpiCard(s1, 534, 224, 152, 76, 'SLA ASTOR & HUB GRW', snap.vendorAndHub.astorSla, '#38BDF8', `HUB GRW: ${snap.vendorAndHub.hubSla}`);
  addTextBox(s1, `Generated & Synced: ${snap.generatedAt}`, 42, 368, 640, 18, {
    fontSize: 8.5,
    color: '#64748B',
  });

  // SLIDE 2: ORDER PLANNING & ARMADA BREAKDOWN
  const s2 = addSlide('#0B1120');
  addSlideChrome(s2, 'ORDER PLANNING & ARMADA BREAKDOWN', 'Ringkasan Realisasi Case ID, CBM, Batch Loading & Kapasitas Armada', 2);
  addKpiCard(s2, 20, 62, 164, 62, 'CASE ID PLAN', snap.planning.caseIdPlan, '#FACC15', `Shipped: ${snap.planning.caseIdShipped} (${snap.planning.caseIdRatio})`);
  addKpiCard(s2, 192, 62, 164, 62, 'CBM PLAN', snap.planning.cbmPlan, '#EF4444', `Shipped: ${snap.planning.cbmShipped} (${snap.planning.cbmRatio})`);
  addKpiCard(s2, 364, 62, 164, 62, 'SHIFT 1 vs SHIFT 2 (CBM)', snap.planning.shift1Cbm, '#38BDF8', `Shift 2: ${snap.planning.shift2Cbm}`);
  addKpiCard(s2, 536, 62, 164, 62, 'SLA PERFORMANCE', snap.planning.slaCustomer, '#10B981', `SLA GRW: ${snap.planning.slaGrw}`);

  addTextBox(s2, 'BREAKDOWN BATCH LOADING (B1 - B4)', 20, 134, 250, 16, { fontSize: 9.5, bold: true, color: '#FACC15' });
  addTable(
    s2,
    20,
    152,
    255,
    150,
    ['BATCH', 'CASE ID', 'CBM'],
    snap.planning.batches.map((b) => [b.batch, b.caseId, b.cbm]),
    '#1E3A8A'
  );

  addTextBox(s2, 'SUMMARY PLAN BY ARMADA TYPE', 288, 134, 412, 16, { fontSize: 9.5, bold: true, color: '#FACC15' });
  addTable(
    s2,
    288,
    152,
    412,
    218,
    ['TIPE ARMADA', 'PLANS', 'TOTAL CBM', 'AVG CBM', 'SHARE'],
    snap.planning.armadas.map((a) => [a.type, a.plans, a.totalCbm, a.avgCbm, a.pct]),
    '#991B1B'
  );

  // SLIDE 3: OUTSTANDING & ORDER AGING
  const s3 = addSlide('#0B1120');
  addSlideChrome(s3, 'MONITORING OUTSTANDING & ORDER AGING', 'Status Backlog CBM Sudah Plan vs Belum Plan & Prioritas Store', 3);
  addKpiCard(s3, 20, 62, 130, 62, 'TOTAL BACKLOG', snap.outstanding.totalBacklog, '#A855F7', 'CBM Total');
  addKpiCard(s3, 157, 62, 130, 62, 'SUDAH PLAN', snap.outstanding.sudahPlan, '#10B981', 'CBM');
  addKpiCard(s3, 294, 62, 130, 62, 'BELUM PLAN', snap.outstanding.belumPlan, '#EF4444', 'Urgent Action');
  addKpiCard(s3, 431, 62, 130, 62, 'AGING > 7 HARI', snap.outstanding.aging7Up, '#F97316', 'CBM');
  addKpiCard(s3, 568, 62, 132, 62, 'ALLOC UNSUCCESSFUL', snap.outstanding.allocUnsuccessful, '#FACC15', 'CBM');

  addTextBox(s3, 'TOP STORE PRIORITAS — BACKLOG BELUM PLAN TERTINGGI', 20, 134, 680, 16, { fontSize: 9.5, bold: true, color: '#FACC15' });
  const storeRows =
    snap.outstanding.topBelumPlanStores.length > 0
      ? snap.outstanding.topBelumPlanStores.map((s) => [s.store, s.shipTo, s.sudahPlan, s.belumPlan])
      : [['-', 'Tidak ada store Belum Plan pada filter aktif', '0', '0']];
  addTable(s3, 20, 152, 680, 220, ['STOREBOOKING', 'SHIP TO', 'SUDAH PLAN (CBM)', 'BELUM PLAN (CBM)'], storeRows, '#991B1B');

  // SLIDE 4: SUPPLY PARTNER, HUB GRW & ASTOR SLA
  const s4 = addSlide('#0B1120');
  addSlideChrome(s4, 'SUPPLY PARTNER, HUB GRW & ASTOR SLA', 'Evaluasi Kinerja Vendor Armada, HUB GRW & SLA All BU (AZKO, KWI, TGI, FBI, OTHER, CUSTOMER)', 4);
  addKpiCard(s4, 20, 62, 164, 64, 'VENDOR SUPPORTING', snap.vendorAndHub.vendorSupporting, '#FACC15', `On-Time: ${snap.vendorAndHub.onTimeLc} / ${snap.vendorAndHub.completedLc} LC`);
  addKpiCard(s4, 192, 62, 164, 64, 'TERLAMBAT / DELAYED', `${snap.vendorAndHub.terlambatLc} LC`, '#EF4444', `Delayed Support: ${snap.vendorAndHub.delayedLc} LC`);
  addKpiCard(s4, 364, 62, 164, 64, 'SLA HUB GRW', snap.vendorAndHub.hubSla, '#2DD4BF', `Sudah: ${snap.vendorAndHub.hubSudahLoading}`);
  addKpiCard(s4, 536, 62, 164, 64, 'SLA ASTOR (ALL BU)', snap.vendorAndHub.astorSla, '#38BDF8', `HIT: ${snap.vendorAndHub.astorHit} | MISS: ${snap.vendorAndHub.astorMiss}`);

  addTextBox(s4, 'RINGKASAN MONITORING HUB GRW & LINE OCCUPANCY ASTOR', 20, 138, 680, 16, { fontSize: 9.5, bold: true, color: '#FACC15' });
  addTable(
    s4,
    20,
    158,
    680,
    206,
    ['INDIKATOR OPERASIONAL', 'NILAI SAAT INI', 'KETERANGAN'],
    [
      ['Vendor On-Time Achievement', snap.vendorAndHub.vendorSupporting, `${snap.vendorAndHub.onTimeLc} On-Time dari ${snap.vendorAndHub.completedLc} Completed LC`],
      ['Vendor Terlambat / Delayed', `${snap.vendorAndHub.terlambatLc} / ${snap.vendorAndHub.delayedLc}`, 'Total LC Terlambat & Delayed Admin Support'],
      ['HUB GRW SLA Performance', snap.vendorAndHub.hubSla, `Sudah Loading: ${snap.vendorAndHub.hubSudahLoading} | Belum: ${snap.vendorAndHub.hubBelumLoading}`],
      ['ASTOR ALL BU SLA (HIT vs MISS)', snap.vendorAndHub.astorSla, `HIT: ${snap.vendorAndHub.astorHit} Order | MISS: ${snap.vendorAndHub.astorMiss} Order`],
      ['ASTOR Outstanding & Line Occupancy', `${snap.vendorAndHub.astorOs} Order (${snap.vendorAndHub.astorOsCbm} CBM)`, `Occupancy Yard: ${snap.vendorAndHub.astorOccupancy}`],
    ],
    '#0F766E'
  );

  // SLIDE 5: TOP & BOTTOM CITIES + INSIGHTS
  const s5 = addSlide('#0B1120');
  addSlideChrome(s5, 'DISTRIBUSI VOLUME KOTA & ACTION SUMMARY', 'Top/Bottom Kota Tujuan Berdasarkan CBM Plan & Rekomendasi Operasional', 5);
  addTextBox(s5, 'TOP KOTA (CBM TERTINGGI)', 20, 60, 330, 16, { fontSize: 9, bold: true, color: '#FACC15' });
  addTable(
    s5,
    20,
    78,
    332,
    172,
    ['KOTA', 'PLANS', 'TOTAL CBM'],
    (snap.summary.topCities.length > 0 ? snap.summary.topCities : [{ kota: '-', plans: 0, cbm: 0 }]).map((c) => [
      c.kota,
      String(c.plans),
      fmtId(c.cbm, 2),
    ]),
    '#15803D'
  );

  addTextBox(s5, 'BOTTOM KOTA (CBM TERENDAH)', 368, 60, 332, 16, { fontSize: 9, bold: true, color: '#F87171' });
  addTable(
    s5,
    368,
    78,
    332,
    172,
    ['KOTA', 'PLANS', 'TOTAL CBM'],
    (snap.summary.bottomCities.length > 0 ? snap.summary.bottomCities : [{ kota: '-', plans: 0, cbm: 0 }]).map((c) => [
      c.kota,
      String(c.plans),
      fmtId(c.cbm, 2),
    ]),
    '#991B1B'
  );

  addRect(s5, 20, 260, 680, 118, '#111827', '#FACC15');
  addTextBox(s5, 'EXECUTIVE TAKEAWAYS & REKOMENDASI', 32, 266, 656, 16, { fontSize: 9.5, bold: true, color: '#FACC15' });
  addTextBox(
    s5,
    snap.summary.insights.map((line, i) => `${i + 1}. ${line}`).join('\n'),
    32,
    284,
    656,
    88,
    { fontSize: 8.5, color: '#E2E8F0' }
  );

  // Remove old slides AFTER creating new slides so the deck is never empty
  oldSlideIds.forEach((oldId) => {
    reqs.push({ deleteObject: { objectId: oldId } });
  });

  return reqs;
}

// =========================================================
// LOCAL PPTX EXPORT (PptxGenJS)
// =========================================================
function generateLocalPptx(snap: AhiSnapshot) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';
  pptx.title = 'NDC AHI Sidoarjo - Order Planning Report';
  pptx.author = 'Logistics Dashboard NDC AHI Sidoarjo';

  const applyChrome = (slide: any, title: string, subtitle: string, pageNum: number) => {
    slide.background = { color: '0B1120' };
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.68, fill: { color: '991B1B' } });
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0.68, w: 10, h: 0.04, fill: { color: 'FACC15' } });
    slide.addText(title, { x: 0.3, y: 0.08, w: 7.2, h: 0.32, fontSize: 14, bold: true, color: 'FFFFFF', fontFace: 'Arial' });
    slide.addText(subtitle, { x: 0.3, y: 0.38, w: 7.4, h: 0.22, fontSize: 8.5, color: 'FDE68A', fontFace: 'Arial' });
    slide.addText(`NDC AHI SIDOARJO  |  Slide ${pageNum}/5`, {
      x: 7.5,
      y: 0.2,
      w: 2.2,
      h: 0.25,
      fontSize: 8.5,
      bold: true,
      color: 'FACC15',
      align: 'right',
      fontFace: 'Arial',
    });
    slide.addText(`Update: ${snap.lastUpdate}  •  ${snap.filtersText}`, {
      x: 0.3,
      y: 5.36,
      w: 9.4,
      h: 0.2,
      fontSize: 7,
      color: '64748B',
      fontFace: 'Arial',
    });
  };

  const addCard = (slide: any, x: number, y: number, w: number, h: number, label: string, value: string, accent: string, sub = '') => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y,
      w,
      h,
      fill: { color: '111827' },
      line: { color: accent, width: 1 },
    });
    slide.addShape(pptx.ShapeType.rect, { x, y, w: 0.06, h, fill: { color: accent } });
    slide.addText(label, { x: x + 0.14, y: y + 0.08, w: w - 0.2, h: 0.2, fontSize: 8.5, bold: true, color: '94A3B8' });
    slide.addText(value, { x: x + 0.14, y: y + 0.3, w: w - 0.2, h: 0.38, fontSize: 17, bold: true, color: accent });
    if (sub) {
      slide.addText(sub, { x: x + 0.14, y: y + h - 0.25, w: w - 0.2, h: 0.2, fontSize: 7.5, color: 'CBD5E1' });
    }
  };

  // Slide 1
  const s1 = pptx.addSlide();
  s1.background = { color: '090D16' };
  s1.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.2, h: 5.625, fill: { color: 'DC2626' } });
  s1.addShape(pptx.ShapeType.rect, { x: 0.2, y: 0, w: 0.06, h: 5.625, fill: { color: 'FACC15' } });
  s1.addText('LOGISTICS OPERATIONS & ORDER PLANNING', { x: 0.6, y: 0.55, w: 8.8, h: 0.3, fontSize: 11, bold: true, color: 'FACC15' });
  s1.addText('NDC AHI SIDOARJO\nPERFORMANCE REPORT', { x: 0.6, y: 0.95, w: 8.8, h: 1.2, fontSize: 28, bold: true, color: 'FFFFFF' });
  s1.addText(`Sumber Data: Live Dashboard Spreadsheet  |  Update: ${snap.lastUpdate}`, { x: 0.6, y: 2.25, w: 8.8, h: 0.25, fontSize: 10.5, color: 'CBD5E1' });
  s1.addText(snap.filtersText, { x: 0.6, y: 2.55, w: 8.8, h: 0.25, fontSize: 9.5, color: '94A3B8' });
  addCard(s1, 0.6, 3.1, 2.1, 1.05, 'CBM PLAN / SHIPPED', snap.planning.cbmPlan, 'EF4444', `Shipped: ${snap.planning.cbmShipped} (${snap.planning.cbmRatio})`);
  addCard(s1, 2.86, 3.1, 2.1, 1.05, 'SLA CUSTOMER / GRW', snap.planning.slaCustomer, '10B981', `SLA GRW: ${snap.planning.slaGrw}`);
  addCard(s1, 5.12, 3.1, 2.1, 1.05, 'OUTSTANDING BELUM PLAN', snap.outstanding.belumPlan, 'F59E0B', `Total Backlog: ${snap.outstanding.totalBacklog} CBM`);
  addCard(s1, 7.38, 3.1, 2.1, 1.05, 'SLA ASTOR & HUB GRW', snap.vendorAndHub.astorSla, '38BDF8', `HUB GRW: ${snap.vendorAndHub.hubSla}`);

  // Slide 2
  const s2 = pptx.addSlide();
  applyChrome(s2, 'ORDER PLANNING & ARMADA BREAKDOWN', 'Ringkasan Realisasi Case ID, CBM, Batch Loading & Kapasitas Armada', 2);
  addCard(s2, 0.3, 0.88, 2.25, 0.88, 'CASE ID PLAN', snap.planning.caseIdPlan, 'FACC15', `Shipped: ${snap.planning.caseIdShipped} (${snap.planning.caseIdRatio})`);
  addCard(s2, 2.68, 0.88, 2.25, 0.88, 'CBM PLAN', snap.planning.cbmPlan, 'EF4444', `Shipped: ${snap.planning.cbmShipped} (${snap.planning.cbmRatio})`);
  addCard(s2, 5.06, 0.88, 2.25, 0.88, 'SHIFT 1 vs SHIFT 2', snap.planning.shift1Cbm, '38BDF8', `Shift 2: ${snap.planning.shift2Cbm}`);
  addCard(s2, 7.44, 0.88, 2.25, 0.88, 'SLA PERFORMANCE', snap.planning.slaCustomer, '10B981', `SLA GRW: ${snap.planning.slaGrw}`);

  s2.addText('BREAKDOWN BATCH LOADING', { x: 0.3, y: 1.9, w: 3.4, h: 0.22, fontSize: 9.5, bold: true, color: 'FACC15' });
  s2.addTable(
    [
      [
        { text: 'BATCH', options: { bold: true, fill: { color: '1E3A8A' }, color: 'FACC15' } },
        { text: 'CASE ID', options: { bold: true, fill: { color: '1E3A8A' }, color: 'FACC15' } },
        { text: 'CBM', options: { bold: true, fill: { color: '1E3A8A' }, color: 'FACC15' } },
      ],
      ...snap.planning.batches.map((b, i) => [
        { text: b.batch, options: { fill: { color: i % 2 ? '0F172A' : '111827' }, color: 'F1F5F9', bold: true } },
        { text: b.caseId, options: { fill: { color: i % 2 ? '0F172A' : '111827' }, color: 'F1F5F9' } },
        { text: b.cbm, options: { fill: { color: i % 2 ? '0F172A' : '111827' }, color: 'F1F5F9' } },
      ]),
    ],
    { x: 0.3, y: 2.16, w: 3.5, colW: [1.9, 0.8, 0.8], fontSize: 8.5, border: { color: '334155', pt: 0.5 } }
  );

  s2.addText('SUMMARY PLAN BY ARMADA TYPE', { x: 4.0, y: 1.9, w: 5.7, h: 0.22, fontSize: 9.5, bold: true, color: 'FACC15' });
  s2.addTable(
    [
      [
        { text: 'TIPE ARMADA', options: { bold: true, fill: { color: '991B1B' }, color: 'FACC15' } },
        { text: 'PLANS', options: { bold: true, fill: { color: '991B1B' }, color: 'FACC15' } },
        { text: 'TOTAL CBM', options: { bold: true, fill: { color: '991B1B' }, color: 'FACC15' } },
        { text: 'AVG CBM', options: { bold: true, fill: { color: '991B1B' }, color: 'FACC15' } },
        { text: 'SHARE', options: { bold: true, fill: { color: '991B1B' }, color: 'FACC15' } },
      ],
      ...snap.planning.armadas.map((a, i) => [
        { text: a.type, options: { fill: { color: i % 2 ? '0F172A' : '111827' }, color: 'F1F5F9', bold: true } },
        { text: a.plans, options: { fill: { color: i % 2 ? '0F172A' : '111827' }, color: 'F1F5F9' } },
        { text: a.totalCbm, options: { fill: { color: i % 2 ? '0F172A' : '111827' }, color: 'F1F5F9' } },
        { text: a.avgCbm, options: { fill: { color: i % 2 ? '0F172A' : '111827' }, color: 'F1F5F9' } },
        { text: a.pct, options: { fill: { color: i % 2 ? '0F172A' : '111827' }, color: 'F1F5F9' } },
      ]),
    ],
    { x: 4.0, y: 2.16, w: 5.7, colW: [1.4, 0.9, 1.3, 1.1, 1.0], fontSize: 8.5, border: { color: '334155', pt: 0.5 } }
  );

  // Slide 3
  const s3 = pptx.addSlide();
  applyChrome(s3, 'MONITORING OUTSTANDING & ORDER AGING', 'Status Backlog CBM Sudah Plan vs Belum Plan & Prioritas Store', 3);
  addCard(s3, 0.3, 0.88, 1.78, 0.88, 'TOTAL BACKLOG', snap.outstanding.totalBacklog, 'A855F7', 'CBM');
  addCard(s3, 2.2, 0.88, 1.78, 0.88, 'SUDAH PLAN', snap.outstanding.sudahPlan, '10B981', 'CBM');
  addCard(s3, 4.1, 0.88, 1.78, 0.88, 'BELUM PLAN', snap.outstanding.belumPlan, 'EF4444', 'Urgent');
  addCard(s3, 6.0, 0.88, 1.78, 0.88, 'AGING > 7 HARI', snap.outstanding.aging7Up, 'F97316', 'CBM');
  addCard(s3, 7.9, 0.88, 1.8, 0.88, 'UNSUCCESSFUL', snap.outstanding.allocUnsuccessful, 'FACC15', 'CBM');

  s3.addText('TOP STORE PRIORITAS — BACKLOG BELUM PLAN TERTINGGI', { x: 0.3, y: 1.9, w: 9.4, h: 0.22, fontSize: 9.5, bold: true, color: 'FACC15' });
  const storeRows =
    snap.outstanding.topBelumPlanStores.length > 0
      ? snap.outstanding.topBelumPlanStores
      : [{ store: '-', shipTo: 'Tidak ada store Belum Plan pada filter aktif', sudahPlan: '0', belumPlan: '0' }];
  s3.addTable(
    [
      [
        { text: 'STOREBOOKING', options: { bold: true, fill: { color: '991B1B' }, color: 'FACC15' } },
        { text: 'SHIP TO', options: { bold: true, fill: { color: '991B1B' }, color: 'FACC15' } },
        { text: 'SUDAH PLAN (CBM)', options: { bold: true, fill: { color: '991B1B' }, color: 'FACC15' } },
        { text: 'BELUM PLAN (CBM)', options: { bold: true, fill: { color: '991B1B' }, color: 'FACC15' } },
      ],
      ...storeRows.map((r, i) => [
        { text: r.store, options: { fill: { color: i % 2 ? '0F172A' : '111827' }, color: 'F1F5F9', bold: true } },
        { text: r.shipTo, options: { fill: { color: i % 2 ? '0F172A' : '111827' }, color: 'F1F5F9' } },
        { text: r.sudahPlan, options: { fill: { color: i % 2 ? '0F172A' : '111827' }, color: '10B981' } },
        { text: r.belumPlan, options: { fill: { color: i % 2 ? '0F172A' : '111827' }, color: 'F87171', bold: true } },
      ]),
    ],
    { x: 0.3, y: 2.16, w: 9.4, colW: [1.5, 4.9, 1.5, 1.5], fontSize: 8.5, border: { color: '334155', pt: 0.5 } }
  );

  // Slide 4
  const s4 = pptx.addSlide();
  applyChrome(s4, 'SUPPLY PARTNER, HUB GRW & ASTOR SLA', 'Evaluasi Kinerja Vendor Armada, HUB GRW & SLA All BU', 4);
  addCard(s4, 0.3, 0.88, 2.25, 0.9, 'VENDOR SUPPORTING', snap.vendorAndHub.vendorSupporting, 'FACC15', `On-Time: ${snap.vendorAndHub.onTimeLc}/${snap.vendorAndHub.completedLc} LC`);
  addCard(s4, 2.68, 0.88, 2.25, 0.9, 'TERLAMBAT LC', `${snap.vendorAndHub.terlambatLc} LC`, 'EF4444', `Delayed Support: ${snap.vendorAndHub.delayedLc} LC`);
  addCard(s4, 5.06, 0.88, 2.25, 0.9, 'SLA HUB GRW', snap.vendorAndHub.hubSla, '2DD4BF', `Sudah Loading: ${snap.vendorAndHub.hubSudahLoading}`);
  addCard(s4, 7.44, 0.88, 2.25, 0.9, 'SLA ASTOR (ALL BU)', snap.vendorAndHub.astorSla, '38BDF8', `HIT: ${snap.vendorAndHub.astorHit} | MISS: ${snap.vendorAndHub.astorMiss}`);

  s4.addText('RINGKASAN MONITORING HUB GRW & LINE OCCUPANCY ASTOR', { x: 0.3, y: 1.94, w: 9.4, h: 0.22, fontSize: 9.5, bold: true, color: 'FACC15' });
  const hubAstorTableRows = [
    ['Vendor On-Time Achievement', snap.vendorAndHub.vendorSupporting, `${snap.vendorAndHub.onTimeLc} On-Time dari ${snap.vendorAndHub.completedLc} Completed LC`],
    ['Vendor Terlambat / Delayed', `${snap.vendorAndHub.terlambatLc} / ${snap.vendorAndHub.delayedLc}`, 'Total LC Terlambat & Delayed Admin Support'],
    ['HUB GRW SLA Performance', snap.vendorAndHub.hubSla, `Sudah Loading: ${snap.vendorAndHub.hubSudahLoading} | Belum: ${snap.vendorAndHub.hubBelumLoading}`],
    ['ASTOR ALL BU SLA (HIT vs MISS)', snap.vendorAndHub.astorSla, `HIT: ${snap.vendorAndHub.astorHit} Order | MISS: ${snap.vendorAndHub.astorMiss} Order`],
    ['ASTOR Outstanding & Line Occupancy', `${snap.vendorAndHub.astorOs} Order (${snap.vendorAndHub.astorOsCbm} CBM)`, `Occupancy Yard: ${snap.vendorAndHub.astorOccupancy}`],
  ];
  s4.addTable(
    [
      [
        { text: 'INDIKATOR OPERASIONAL', options: { bold: true, fill: { color: '0F766E' }, color: 'FACC15' } },
        { text: 'NILAI SAAT INI', options: { bold: true, fill: { color: '0F766E' }, color: 'FACC15' } },
        { text: 'KETERANGAN', options: { bold: true, fill: { color: '0F766E' }, color: 'FACC15' } },
      ],
      ...hubAstorTableRows.map((r, i) => [
        { text: r[0], options: { fill: { color: i % 2 ? '0F172A' : '111827' }, color: 'F1F5F9', bold: true } },
        { text: r[1], options: { fill: { color: i % 2 ? '0F172A' : '111827' }, color: 'FACC15', bold: true } },
        { text: r[2], options: { fill: { color: i % 2 ? '0F172A' : '111827' }, color: 'CBD5E1' } },
      ]),
    ],
    { x: 0.3, y: 2.22, w: 9.4, colW: [3.1, 2.2, 4.1], fontSize: 9, border: { color: '334155', pt: 0.5 } }
  );

  // Slide 5
  const s5 = pptx.addSlide();
  applyChrome(s5, 'DISTRIBUSI VOLUME KOTA & ACTION SUMMARY', 'Top/Bottom Kota Tujuan Berdasarkan CBM Plan & Rekomendasi Operasional', 5);
  s5.addText('TOP KOTA (CBM TERTINGGI)', { x: 0.3, y: 0.85, w: 4.5, h: 0.22, fontSize: 9, bold: true, color: 'FACC15' });
  s5.addTable(
    [
      [
        { text: 'KOTA', options: { bold: true, fill: { color: '15803D' }, color: 'FACC15' } },
        { text: 'PLANS', options: { bold: true, fill: { color: '15803D' }, color: 'FACC15' } },
        { text: 'TOTAL CBM', options: { bold: true, fill: { color: '15803D' }, color: 'FACC15' } },
      ],
      ...(snap.summary.topCities.length ? snap.summary.topCities : [{ kota: '-', plans: 0, cbm: 0 }]).map((c, i) => [
        { text: c.kota, options: { fill: { color: i % 2 ? '0F172A' : '111827' }, color: 'F1F5F9', bold: true } },
        { text: String(c.plans), options: { fill: { color: i % 2 ? '0F172A' : '111827' }, color: 'F1F5F9' } },
        { text: fmtId(c.cbm, 2), options: { fill: { color: i % 2 ? '0F172A' : '111827' }, color: '4ADE80' } },
      ]),
    ],
    { x: 0.3, y: 1.1, w: 4.55, colW: [2.35, 0.9, 1.3], fontSize: 8.5, border: { color: '334155', pt: 0.5 } }
  );

  s5.addText('BOTTOM KOTA (CBM TERENDAH)', { x: 5.15, y: 0.85, w: 4.55, h: 0.22, fontSize: 9, bold: true, color: 'F87171' });
  s5.addTable(
    [
      [
        { text: 'KOTA', options: { bold: true, fill: { color: '991B1B' }, color: 'FACC15' } },
        { text: 'PLANS', options: { bold: true, fill: { color: '991B1B' }, color: 'FACC15' } },
        { text: 'TOTAL CBM', options: { bold: true, fill: { color: '991B1B' }, color: 'FACC15' } },
      ],
      ...(snap.summary.bottomCities.length ? snap.summary.bottomCities : [{ kota: '-', plans: 0, cbm: 0 }]).map((c, i) => [
        { text: c.kota, options: { fill: { color: i % 2 ? '0F172A' : '111827' }, color: 'F1F5F9', bold: true } },
        { text: String(c.plans), options: { fill: { color: i % 2 ? '0F172A' : '111827' }, color: 'F1F5F9' } },
        { text: fmtId(c.cbm, 2), options: { fill: { color: i % 2 ? '0F172A' : '111827' }, color: 'F87171' } },
      ]),
    ],
    { x: 5.15, y: 1.1, w: 4.55, colW: [2.35, 0.9, 1.3], fontSize: 8.5, border: { color: '334155', pt: 0.5 } }
  );

  s5.addShape(pptx.ShapeType.roundRect, { x: 0.3, y: 3.55, w: 9.4, h: 1.65, fill: { color: '111827' }, line: { color: 'FACC15', width: 1 } });
  s5.addText('EXECUTIVE TAKEAWAYS & REKOMENDASI', { x: 0.48, y: 3.63, w: 9.0, h: 0.22, fontSize: 9.5, bold: true, color: 'FACC15' });
  s5.addText(
    snap.summary.insights.map((line, i) => `${i + 1}. ${line}`).join('\n'),
    { x: 0.48, y: 3.88, w: 9.0, h: 1.25, fontSize: 8.5, color: 'E2E8F0', valign: 'top' }
  );

  const fileStamp = new Date().toISOString().slice(0, 16).replace(/[-:T]/g, '');
  pptx.writeFile({ fileName: `Order_Planning_NDC_AHI_Sidoarjo_${fileStamp}.pptx` });
}

// =========================================================
// GOOGLE SLIDES API CLIENT HELPERS
// =========================================================
async function slidesFetch(url: string, options: RequestInit = {}): Promise<any> {
  if (!cachedAccessToken) {
    throw new Error('AUTH_REQUIRED');
  }
  const res = await fetch(url, {
    ...options,
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${cachedAccessToken}`,
      ...(options.headers || {}),
    },
  });
  if (res.status === 401) {
    cachedAccessToken = null;
    updateAuthUIState();
    throw new Error('AUTH_REQUIRED');
  }
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const msg = data?.error?.message || `Google Slides API Error (${res.status})`;
    throw new Error(msg);
  }
  return data;
}

async function createNewGooglePresentation(snap: AhiSnapshot) {
  const titleDate = new Date().toLocaleDateString('id-ID', {
    day: '2-digit',
    month: 'short',
    year: 'numeric',
  });
  const title = `NDC AHI Sidoarjo - Order Planning Report (${titleDate})`;

  const created = await slidesFetch('https://slides.googleapis.com/v1/presentations', {
    method: 'POST',
    body: JSON.stringify({ title }),
  });

  const presentationId: string = created.presentationId;
  const oldSlideIds: string[] = (created.slides || []).map((s: any) => s.objectId).filter(Boolean);
  const requests = buildGoogleSlidesRequests(snap, oldSlideIds);

  await slidesFetch(`https://slides.googleapis.com/v1/presentations/${presentationId}:batchUpdate`, {
    method: 'POST',
    body: JSON.stringify({ requests }),
  });

  localStorage.setItem(STORAGE_KEY_PRES_ID, presentationId);
  localStorage.setItem(STORAGE_KEY_PRES_TITLE, title);
  lastSyncedDataSignature = snap.signature;
  return { presentationId, title };
}

async function inspectGooglePresentation(presentationId: string) {
  const data = await slidesFetch(
    `https://slides.googleapis.com/v1/presentations/${encodeURIComponent(presentationId)}`
  );
  return {
    presentationId: data.presentationId as string,
    title: (data.title as string) || 'Presentasi Google Slides',
    slideIds: ((data.slides as any[]) || []).map((s) => s.objectId).filter(Boolean) as string[],
  };
}

async function syncToExistingGooglePresentation(
  presentationId: string,
  title: string,
  oldSlideIds: string[],
  snap: AhiSnapshot
) {
  const requests = buildGoogleSlidesRequests(snap, oldSlideIds);
  await slidesFetch(`https://slides.googleapis.com/v1/presentations/${encodeURIComponent(presentationId)}:batchUpdate`, {
    method: 'POST',
    body: JSON.stringify({ requests }),
  });
  localStorage.setItem(STORAGE_KEY_PRES_ID, presentationId);
  localStorage.setItem(STORAGE_KEY_PRES_TITLE, title);
  lastSyncedDataSignature = snap.signature;
}

// =========================================================
// UI CONTROLLERS
// =========================================================
function setSlidesStatus(msg: string, tone: 'info' | 'success' | 'error' = 'info') {
  const el = document.getElementById('ahi-slides-status-msg');
  if (!el) return;
  const colors = {
    info: { bg: 'rgba(56, 189, 248, 0.12)', border: '#0284c7', text: '#bae6fd' },
    success: { bg: 'rgba(16, 185, 129, 0.14)', border: '#10b981', text: '#a7f3d0' },
    error: { bg: 'rgba(239, 68, 68, 0.14)', border: '#ef4444', text: '#fecaca' },
  }[tone];
  el.style.display = msg ? 'block' : 'none';
  el.style.background = colors.bg;
  el.style.borderColor = colors.border;
  el.style.color = colors.text;
  el.innerHTML = msg;
}

function updateAuthUIState() {
  const signedOutBox = document.getElementById('ahi-slides-auth-signedout');
  const signedInBox = document.getElementById('ahi-slides-auth-signedin');
  const userNameEl = document.getElementById('ahi-slides-user-name');
  const syncControls = document.getElementById('ahi-slides-sync-controls');

  const isAuthed = Boolean(currentUser && cachedAccessToken);
  if (signedOutBox) signedOutBox.style.display = isAuthed ? 'none' : 'flex';
  if (signedInBox) signedInBox.style.display = isAuthed ? 'flex' : 'none';
  if (userNameEl && currentUser) {
    userNameEl.textContent = currentUser.displayName || currentUser.email || 'Akun Google Terhubung';
  }
  if (syncControls) {
    syncControls.style.opacity = isAuthed ? '1' : '0.6';
  }
}

function renderLinkedPresentationCard() {
  const presId = localStorage.getItem(STORAGE_KEY_PRES_ID) || '';
  const presTitle = localStorage.getItem(STORAGE_KEY_PRES_TITLE) || 'Presentasi Google Slides';
  const card = document.getElementById('ahi-slides-linked-card');
  const titleEl = document.getElementById('ahi-slides-linked-title');
  const openBtn = document.getElementById('ahi-slides-open-link') as HTMLAnchorElement | null;
  const pptxLink = document.getElementById('ahi-slides-export-pptx-link') as HTMLAnchorElement | null;
  const inputEl = document.getElementById('ahi-slides-target-input') as HTMLInputElement | null;
  const headerBadge = document.getElementById('header-slides-sync-badge');

  if (inputEl && presId && !inputEl.value.trim()) {
    inputEl.value = `https://docs.google.com/presentation/d/${presId}/edit`;
  }

  if (card) card.style.display = presId ? 'flex' : 'none';
  if (titleEl) titleEl.textContent = presTitle;
  if (openBtn) openBtn.href = `https://docs.google.com/presentation/d/${presId}/edit`;
  if (pptxLink) pptxLink.href = `https://docs.google.com/presentation/d/${presId}/export/pptx`;
  if (headerBadge) {
    headerBadge.textContent = presId ? 'SYNCED' : 'PPT';
  }
}

function renderSnapshotPreview(snap: AhiSnapshot) {
  const previewEl = document.getElementById('ahi-slides-snapshot-preview');
  if (!previewEl) return;
  previewEl.innerHTML = `
    <div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap:8px;">
      <div style="background:#0f172a; border:1px solid #334155; border-radius:6px; padding:8px;">
        <div style="font-size:0.7em; color:#94a3b8;">SLIDE 1-2: PLANNING</div>
        <div style="font-size:0.95em; font-weight:800; color:#ffd700; margin-top:2px;">${snap.planning.cbmPlan} CBM</div>
        <div style="font-size:0.72em; color:#cbd5e1;">Shipped: ${snap.planning.cbmShipped} (${snap.planning.cbmRatio})</div>
      </div>
      <div style="background:#0f172a; border:1px solid #334155; border-radius:6px; padding:8px;">
        <div style="font-size:0.7em; color:#94a3b8;">SLIDE 3: OUTSTANDING</div>
        <div style="font-size:0.95em; font-weight:800; color:#f87171; margin-top:2px;">${snap.outstanding.belumPlan} CBM</div>
        <div style="font-size:0.72em; color:#cbd5e1;">Belum Plan (Total: ${snap.outstanding.totalBacklog})</div>
      </div>
      <div style="background:#0f172a; border:1px solid #334155; border-radius:6px; padding:8px;">
        <div style="font-size:0.7em; color:#94a3b8;">SLIDE 4: HUB &amp; ASTOR</div>
        <div style="font-size:0.95em; font-weight:800; color:#2dd4bf; margin-top:2px;">ASTOR ${snap.vendorAndHub.astorSla}</div>
        <div style="font-size:0.72em; color:#cbd5e1;">HUB GRW: ${snap.vendorAndHub.hubSla} • Vendor: ${snap.vendorAndHub.vendorSupporting}</div>
      </div>
      <div style="background:#0f172a; border:1px solid #334155; border-radius:6px; padding:8px;">
        <div style="font-size:0.7em; color:#94a3b8;">SLIDE 5: KOTA &amp; SUMMARY</div>
        <div style="font-size:0.95em; font-weight:800; color:#38bdf8; margin-top:2px;">SLA ${snap.planning.slaCustomer}</div>
        <div style="font-size:0.72em; color:#cbd5e1;">Top Kota: ${snap.summary.topCities[0]?.kota || '-'}</div>
      </div>
    </div>
    <div style="margin-top:8px; font-size:0.74em; color:#94a3b8;">
      <i class="fas fa-clock" style="color:#ffd700; margin-right:4px;"></i>Update Spreadsheet: <strong style="color:#e2e8f0;">${snap.lastUpdate}</strong> • ${snap.filtersText}
    </div>
  `;
}

function hideConfirmBox() {
  const box = document.getElementById('ahi-slides-confirm-box');
  if (box) box.style.display = 'none';
  pendingConfirmAction = null;
}

function showConfirmBox(title: string, detailsHtml: string, onConfirm: () => Promise<void>) {
  const box = document.getElementById('ahi-slides-confirm-box');
  const titleEl = document.getElementById('ahi-slides-confirm-title');
  const descEl = document.getElementById('ahi-slides-confirm-desc');
  if (!box || !titleEl || !descEl) return;
  titleEl.textContent = title;
  descEl.innerHTML = detailsHtml;
  pendingConfirmAction = onConfirm;
  box.style.display = 'block';
}

onAuthStateChanged(auth, (user) => {
  currentUser = user;
  if (!user) {
    cachedAccessToken = null;
  } else if (!cachedAccessToken && !isSigningIn) {
    cachedAccessToken = null;
  }
  updateAuthUIState();
});

const win = window as any;

win.openAhiSlidesModal = function () {
  hideConfirmBox();
  setSlidesStatus('', 'info');
  const snap = collectAhiDashboardSnapshot();
  renderSnapshotPreview(snap);
  renderLinkedPresentationCard();
  updateAuthUIState();

  const watchCheckbox = document.getElementById('ahi-slides-watch-toggle') as HTMLInputElement | null;
  if (watchCheckbox) {
    watchCheckbox.checked = localStorage.getItem(STORAGE_KEY_WATCH_SYNC) === '1';
  }

  const modal = document.getElementById('modal-ahi-slides-sync');
  if (modal) {
    modal.classList.remove('hidden');
    document.body.style.overflow = 'hidden';
  }

  const banner = document.getElementById('ahi-slides-update-banner');
  if (banner) banner.style.display = 'none';
};

win.closeAhiSlidesModal = function () {
  hideConfirmBox();
  const modal = document.getElementById('modal-ahi-slides-sync');
  if (modal) modal.classList.add('hidden');
  document.body.style.overflow = '';
};

win.signInGoogleSlides = async function () {
  try {
    isSigningIn = true;
    setSlidesStatus('<i class="fas fa-spinner fa-spin"></i> Menghubungkan ke akun Google Slides...', 'info');
    const result = await signInWithPopup(auth, provider);
    const credential = GoogleAuthProvider.credentialFromResult(result);
    if (!credential?.accessToken) {
      throw new Error('Gagal mendapatkan izin akses Google Slides. Silakan coba login ulang.');
    }
    cachedAccessToken = credential.accessToken;
    currentUser = result.user;
    updateAuthUIState();
    setSlidesStatus('<i class="fas fa-check-circle"></i> Akun Google terhubung! Anda siap membuat atau menyinkronkan Google Slides.', 'success');
  } catch (err: any) {
    setSlidesStatus(`<i class="fas fa-exclamation-triangle"></i> Login gagal: ${err?.message || 'Dibatalkan'}`, 'error');
  } finally {
    isSigningIn = false;
  }
};

win.signOutGoogleSlides = async function () {
  await auth.signOut();
  cachedAccessToken = null;
  currentUser = null;
  hideConfirmBox();
  updateAuthUIState();
  setSlidesStatus('Sesi Google Slides telah diputuskan.', 'info');
};

win.downloadAhiPptxLocal = function () {
  try {
    const snap = collectAhiDashboardSnapshot();
    generateLocalPptx(snap);
    setSlidesStatus('<i class="fas fa-check-circle"></i> File PowerPoint (.pptx) berhasil dibuat dan diunduh sesuai data dashboard saat ini!', 'success');
  } catch (err: any) {
    setSlidesStatus(`<i class="fas fa-exclamation-triangle"></i> Gagal membuat PPTX: ${err?.message || 'Error'}`, 'error');
  }
};

win.handleCreateNewGoogleSlides = async function () {
  if (!cachedAccessToken) {
    setSlidesStatus('<i class="fas fa-lock"></i> Silakan klik <strong>Sign in with Google</strong> terlebih dahulu.', 'error');
    return;
  }
  const snap = collectAhiDashboardSnapshot();
  showConfirmBox(
    'Buat Presentasi Google Slides Baru?',
    `Aplikasi akan membuat presentasi <strong>Google Slides</strong> baru berisi <strong>5 slide laporan NDC AHI Sidoarjo</strong> sesuai data dashboard terakhir (${snap.lastUpdate}).`,
    async () => {
      try {
        setSlidesStatus('<i class="fas fa-spinner fa-spin"></i> Membuat presentasi baru di Google Slides...', 'info');
        const created = await createNewGooglePresentation(snap);
        const inputEl = document.getElementById('ahi-slides-target-input') as HTMLInputElement | null;
        if (inputEl) inputEl.value = `https://docs.google.com/presentation/d/${created.presentationId}/edit`;
        renderLinkedPresentationCard();
        setSlidesStatus(
          `<i class="fas fa-check-circle"></i> Presentasi <strong>${created.title}</strong> berhasil dibuat &amp; disinkronkan!`,
          'success'
        );
      } catch (err: any) {
        const msg = err?.message === 'AUTH_REQUIRED'
          ? 'Sesi Google berakhir. Silakan Sign in with Google kembali.'
          : err?.message || 'Gagal membuat Google Slides';
        setSlidesStatus(`<i class="fas fa-exclamation-triangle"></i> ${msg}`, 'error');
      }
    }
  );
};

win.handleSyncExistingGoogleSlides = async function () {
  if (!cachedAccessToken) {
    setSlidesStatus('<i class="fas fa-lock"></i> Silakan klik <strong>Sign in with Google</strong> terlebih dahulu.', 'error');
    return;
  }
  const inputEl = document.getElementById('ahi-slides-target-input') as HTMLInputElement | null;
  const targetId = extractPresentationId(inputEl?.value || localStorage.getItem(STORAGE_KEY_PRES_ID) || '');
  if (!targetId) {
    setSlidesStatus('<i class="fas fa-info-circle"></i> Tempel link atau ID Google Slides tujuan, atau klik <strong>Buat Google Slides Baru</strong>.', 'error');
    return;
  }

  try {
    setSlidesStatus('<i class="fas fa-spinner fa-spin"></i> Memeriksa presentasi Google Slides tujuan...', 'info');
    const info = await inspectGooglePresentation(targetId);
    const snap = collectAhiDashboardSnapshot();
    setSlidesStatus('', 'info');

    showConfirmBox(
      `Konfirmasi Sinkronisasi: "${info.title}"`,
      `Tindakan ini akan memperbarui isi presentasi <strong>${info.title}</strong> (ID: <code>${info.presentationId}</code>) dengan mengganti <strong>${info.slideIds.length} slide lama</strong> menjadi <strong>5 slide terbaru</strong> dari dashboard NDC AHI Sidoarjo (Update: ${snap.lastUpdate}).`,
      async () => {
        try {
          setSlidesStatus('<i class="fas fa-spinner fa-spin"></i> Menyinkronkan 5 slide ke Google Slides...', 'info');
          await syncToExistingGooglePresentation(info.presentationId, info.title, info.slideIds, snap);
          renderLinkedPresentationCard();
          setSlidesStatus(
            `<i class="fas fa-check-circle"></i> Presentasi <strong>${info.title}</strong> berhasil disinkronkan dengan data spreadsheet terbaru!`,
            'success'
          );
        } catch (err: any) {
          const msg = err?.message === 'AUTH_REQUIRED'
            ? 'Sesi Google berakhir. Silakan Sign in with Google kembali.'
            : err?.message || 'Gagal sinkronisasi ke Google Slides';
          setSlidesStatus(`<i class="fas fa-exclamation-triangle"></i> ${msg}`, 'error');
        }
      }
    );
  } catch (err: any) {
    const msg = err?.message === 'AUTH_REQUIRED'
      ? 'Sesi Google berakhir. Silakan Sign in with Google kembali.'
      : `Tidak dapat mengakses presentasi tersebut: ${err?.message || 'Periksa URL/ID'}`;
    setSlidesStatus(`<i class="fas fa-exclamation-triangle"></i> ${msg}`, 'error');
  }
};

win.confirmAhiSlidesAction = async function () {
  const fn = pendingConfirmAction;
  hideConfirmBox();
  if (fn) await fn();
};

win.cancelAhiSlidesAction = function () {
  hideConfirmBox();
};

win.toggleAhiSlidesWatchSync = function (checked: boolean) {
  localStorage.setItem(STORAGE_KEY_WATCH_SYNC, checked ? '1' : '0');
};

win.clearLinkedGooglePresentation = function () {
  localStorage.removeItem(STORAGE_KEY_PRES_ID);
  localStorage.removeItem(STORAGE_KEY_PRES_TITLE);
  const inputEl = document.getElementById('ahi-slides-target-input') as HTMLInputElement | null;
  if (inputEl) inputEl.value = '';
  renderLinkedPresentationCard();
  setSlidesStatus('Link presentasi Google Slides dilepaskan.', 'info');
};

win.notifyAhiSpreadsheetSynced = function () {
  if (document.body.classList.contains('theme-purple') || document.body.classList.contains('theme-cyan')) {
    return;
  }
  const watchEnabled = localStorage.getItem(STORAGE_KEY_WATCH_SYNC) === '1';
  const linkedId = localStorage.getItem(STORAGE_KEY_PRES_ID);
  if (!watchEnabled || !linkedId) return;

  const snap = collectAhiDashboardSnapshot();
  if (lastSyncedDataSignature === snap.signature) return;

  const banner = document.getElementById('ahi-slides-update-banner');
  if (banner) {
    banner.style.display = 'flex';
  }
};

document.addEventListener('DOMContentLoaded', () => {
  renderLinkedPresentationCard();
  updateAuthUIState();
});
