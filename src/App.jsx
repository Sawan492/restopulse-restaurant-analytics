import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  BarChart, Bar, LineChart, Line, AreaChart, Area, PieChart, Pie, Cell,
  ScatterChart, Scatter, ZAxis, XAxis, YAxis, CartesianGrid, Tooltip,
  Legend, ResponsiveContainer, ComposedChart, ReferenceLine
} from 'recharts';
import {
  LayoutDashboard, TrendingUp, Utensils, Users, Sparkles, UploadCloud,
  Download, FileSpreadsheet, CheckCircle2, AlertTriangle, Info, ArrowUpRight,
  ArrowDownRight, Search, Filter, RefreshCw, Send, Bot, User, HelpCircle,
  ChevronRight, ChevronLeft, Layers, DollarSign, ShoppingBag, Repeat,
  Award, AlertCircle, Trash2, Eye, Sliders, FileText, Check, X, Menu,
  Calendar, PieChart as PieChartIcon, BarChart3, Zap, Target, ShieldAlert,
  ArrowRight, Database
} from 'lucide-react';

// ============================================================================
// INDIAN REGIONAL SETTINGS, CURRENCY & DATE FORMATTERS (en-IN / IST)
// ============================================================================

const CURRENCY_OPTIONS = {
  INR: { code: 'INR', symbol: '₹', locale: 'en-IN', label: 'INR (₹) — Indian Rupee' },
  USD: { code: 'USD', symbol: '$', locale: 'en-US', label: 'USD ($) — US Dollar' },
  EUR: { code: 'EUR', symbol: '€', locale: 'en-IE', label: 'EUR (€) — Euro' },
  GBP: { code: 'GBP', symbol: '£', locale: 'en-GB', label: 'GBP (£) — British Pound' },
  AED: { code: 'AED', symbol: 'AED ', locale: 'en-AE', label: 'AED (د.إ) — UAE Dirham' },
};

const GST_OPTIONS = [
  { id: 'none', label: 'GST Not Specified (Use Raw Values)', rate: 0, mode: 'none' },
  { id: 'incl_5', label: 'Prices Inclusive of 5% GST (Restaurant Std)', rate: 5, mode: 'inclusive' },
  { id: 'excl_5', label: 'Prices Exclusive of 5% GST (+5% GST)', rate: 5, mode: 'exclusive' },
  { id: 'incl_18', label: 'Prices Inclusive of 18% GST (AC/Fine Dining)', rate: 18, mode: 'inclusive' },
  { id: 'excl_18', label: 'Prices Exclusive of 18% GST (+18% GST)', rate: 18, mode: 'exclusive' },
];

// Active currency configuration (defaults to Indian Rupees ₹ and en-IN numbering: ₹1,25,000)
let activeCurrency = CURRENCY_OPTIONS.INR;

const setGlobalCurrency = (code) => {
  if (CURRENCY_OPTIONS[code]) {
    activeCurrency = CURRENCY_OPTIONS[code];
  }
};

const formatCurrency = (val, decimals = 0) => {
  if (val === null || val === undefined || isNaN(val)) return '—';
  return new Intl.NumberFormat(activeCurrency.locale, {
    style: 'currency',
    currency: activeCurrency.code,
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  }).format(val);
};

const formatCompactCurrency = (val) => {
  if (val === null || val === undefined || isNaN(val)) return '—';
  const abs = Math.abs(val);
  const sym = activeCurrency.symbol;
  if (activeCurrency.code === 'INR') {
    if (abs >= 10000000) return `${sym}${(val / 10000000).toFixed(2)}Cr`;
    if (abs >= 100000) return `${sym}${(val / 100000).toFixed(2)}L`;
    if (abs >= 1000) return `${sym}${(val / 1000).toFixed(1)}k`;
    return `${sym}${val.toFixed(0)}`;
  }
  if (abs >= 1000000) return `${sym}${(val / 1000000).toFixed(2)}M`;
  if (abs >= 1000) return `${sym}${(val / 1000).toFixed(1)}k`;
  return `${sym}${val.toFixed(0)}`;
};

const formatNumber = (val, decimals = 0) => {
  if (val === null || val === undefined || isNaN(val)) return '—';
  return new Intl.NumberFormat(activeCurrency.locale, {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  }).format(val);
};

const formatPercent = (val, decimals = 1) => {
  if (val === null || val === undefined || isNaN(val)) return '—';
  return `${Number(val).toFixed(decimals)}%`;
};

// Format ISO YYYY-MM-DD or Date into Indian DD-MM-YYYY format
const formatDateDDMMYYYY = (dateInput) => {
  if (!dateInput) return '—';
  if (typeof dateInput === 'string') {
    const trimmed = dateInput.trim();
    // Already DD-MM-YYYY
    if (/^\d{2}-\d{2}-\d{4}$/.test(trimmed)) return trimmed;
    // ISO YYYY-MM-DD
    const isoMatch = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (isoMatch) return `${isoMatch[3]}-${isoMatch[2]}-${isoMatch[1]}`;
  }
  const dt = dateInput instanceof Date ? dateInput : new Date(dateInput);
  if (isNaN(dt.getTime())) return String(dateInput);
  const dd = String(dt.getUTCDate()).padStart(2, '0');
  const mm = String(dt.getUTCMonth() + 1).padStart(2, '0');
  const yyyy = dt.getUTCFullYear();
  return `${dd}-${mm}-${yyyy}`;
};

// Parse both DD-MM-YYYY / DD/MM/YYYY and YYYY-MM-DD dates accurately
const parseFlexibleDate = (rawStr) => {
  if (!rawStr) return null;
  if (rawStr instanceof Date && !isNaN(rawStr.getTime())) return rawStr;
  const s = String(rawStr).trim();
  // Match DD-MM-YYYY or DD/MM/YYYY
  const dmy = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (dmy) {
    const day = parseInt(dmy[1], 10);
    const month = parseInt(dmy[2], 10) - 1;
    const year = parseInt(dmy[3], 10);
    const dt = new Date(Date.UTC(year, month, day, 12, 0, 0));
    if (!isNaN(dt.getTime())) return dt;
  }
  // Match YYYY-MM-DD
  const ymd = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (ymd) {
    const year = parseInt(ymd[1], 10);
    const month = parseInt(ymd[2], 10) - 1;
    const day = parseInt(ymd[3], 10);
    const dt = new Date(Date.UTC(year, month, day, 12, 0, 0));
    if (!isNaN(dt.getTime())) return dt;
  }
  const fallback = new Date(s);
  return isNaN(fallback.getTime()) ? null : fallback;
};

const getISTTimeString = () =>
  new Date().toLocaleTimeString('en-IN', {
    timeZone: 'Asia/Kolkata',
    hour: '2-digit',
    minute: '2-digit',
    hour12: true,
  });

const computeGSTBreakdown = (rawRevenue, gstOptionId) => {
  const opt = GST_OPTIONS.find((o) => o.id === gstOptionId) || GST_OPTIONS[0];
  if (opt.mode === 'none' || opt.rate === 0) {
    return {
      isActive: false,
      label: opt.label,
      rate: 0,
      mode: 'none',
      rawRevenue,
      netRevenueExGst: rawRevenue,
      gstAmount: 0,
      grossRevenueInclGst: rawRevenue,
    };
  }
  if (opt.mode === 'inclusive') {
    const net = rawRevenue / (1 + opt.rate / 100);
    const gst = rawRevenue - net;
    return {
      isActive: true,
      label: opt.label,
      rate: opt.rate,
      mode: 'inclusive',
      rawRevenue,
      netRevenueExGst: Number(net.toFixed(2)),
      gstAmount: Number(gst.toFixed(2)),
      grossRevenueInclGst: Number(rawRevenue.toFixed(2)),
    };
  }
  // Exclusive
  const gst = rawRevenue * (opt.rate / 100);
  const gross = rawRevenue + gst;
  return {
    isActive: true,
    label: opt.label,
    rate: opt.rate,
    mode: 'exclusive',
    rawRevenue,
    netRevenueExGst: Number(rawRevenue.toFixed(2)),
    gstAmount: Number(gst.toFixed(2)),
    grossRevenueInclGst: Number(gross.toFixed(2)),
  };
};

const parseNumeric = (val) => {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  const cleaned = String(val).replace(/[^0-9.-]+/g, '');
  const parsed = parseFloat(cleaned);
  return isNaN(parsed) ? 0 : parsed;
};

const normalizeHeader = (str) =>
  String(str || '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');

// ============================================================================
// REALISTIC INDIAN RESTAURANT SAMPLE DEMO DATASETS (INR ₹ & DD-MM-YYYY)
// ============================================================================

const MENU_CATALOG = [
  { name: 'Chicken Biryani', price: 380, category: 'Biryani & Rice', role: 'Star' },
  { name: 'Paneer Butter Masala', price: 320, category: 'North Indian Mains', role: 'Star' },
  { name: 'Veg Biryani', price: 290, category: 'Biryani & Rice', role: 'Star' },
  { name: 'Dal Makhani', price: 280, category: 'North Indian Mains', role: 'Star' },
  { name: 'Veg Thali', price: 260, category: 'Thali & Combos', role: 'Star' },
  { name: 'Masala Dosa', price: 140, category: 'South Indian Tiffin', role: 'Plowhorse' },
  { name: 'Butter Naan', price: 65, category: 'Indian Breads', role: 'Plowhorse' },
  { name: 'Masala Chai', price: 45, category: 'Beverages', role: 'Plowhorse' },
  { name: 'Tandoori Roti', price: 40, category: 'Indian Breads', role: 'Plowhorse' },
  { name: 'Idli Sambar', price: 90, category: 'South Indian Tiffin', role: 'Plowhorse' },
  { name: 'Paneer Tikka', price: 310, category: 'Tandoor Starters', role: 'Puzzle' },
  { name: 'Chole Bhature', price: 180, category: 'North Indian Classics', role: 'Puzzle' },
  { name: 'Pav Bhaji', price: 160, category: 'Mumbai Street Classics', role: 'Dog' },
  { name: 'Lassi', price: 110, category: 'Beverages', role: 'Dog' },
  { name: 'Gulab Jamun', price: 95, category: 'Desserts', role: 'Dog' },
];

// Authentic Indian meal co-ordering affinities for realistic customer profiles
const INDIAN_MEAL_COMBOS = [
  [MENU_CATALOG[1], MENU_CATALOG[6], MENU_CATALOG[3]], // Paneer Butter Masala + Butter Naan + Dal Makhani
  [MENU_CATALOG[0], MENU_CATALOG[13], MENU_CATALOG[14]], // Chicken Biryani + Lassi + Gulab Jamun
  [MENU_CATALOG[5], MENU_CATALOG[7], MENU_CATALOG[9]], // Masala Dosa + Masala Chai + Idli Sambar
  [MENU_CATALOG[3], MENU_CATALOG[8], MENU_CATALOG[10]], // Dal Makhani + Tandoori Roti + Paneer Tikka
  [MENU_CATALOG[2], MENU_CATALOG[6], MENU_CATALOG[7]], // Veg Biryani + Butter Naan + Masala Chai
  [MENU_CATALOG[4], MENU_CATALOG[7], MENU_CATALOG[14]], // Veg Thali + Masala Chai + Gulab Jamun
  [MENU_CATALOG[11], MENU_CATALOG[13], MENU_CATALOG[7]], // Chole Bhature + Lassi + Masala Chai
];

// Deterministic pseudo-random generator so sample data is stable and consistent
const seededRandom = (seed) => {
  let x = Math.sin(seed) * 10000;
  return x - Math.floor(x);
};

const generateSampleTransactionDataset = () => {
  const rows = [];
  const startDate = new Date('2026-05-01T12:00:00Z');
  const endDate = new Date('2026-10-31T22:00:00Z');
  const totalDays = Math.floor((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;

  // 135 distinct Indian restaurant customers with varying loyalty profiles
  const customers = Array.from({ length: 135 }, (_, i) => {
    const idNum = 1001 + i;
    const tierSeed = seededRandom(i + 42);
    let profile = 'regular';
    if (tierSeed > 0.82) profile = 'vip';
    else if (tierSeed < 0.25) profile = 'single_or_declining';
    else if (tierSeed < 0.45) profile = 'occasional_big';

    const combo = INDIAN_MEAL_COMBOS[i % INDIAN_MEAL_COMBOS.length];

    return {
      customerId: `CUST-IN-${idNum}`,
      email: `patron${idNum}@spicegarden.co.in`,
      profile,
      fav1: combo[0],
      fav2: combo[1],
      fav3: combo[2],
    };
  });

  // Weighted item picker reflecting realistic Indian restaurant demand across the 15 items
  const itemWeights = [
    15, 14, 12, 13, 12, // Stars (Chicken Biryani, Paneer Butter Masala, Veg Biryani, Dal Makhani, Veg Thali)
    18, 22, 24, 19, 16, // Plowhorses (Masala Dosa, Butter Naan, Masala Chai, Tandoori Roti, Idli Sambar)
    6.5, 8.5,           // Puzzles (Paneer Tikka, Chole Bhature)
    4.0, 4.5, 4.2       // Dogs (Pav Bhaji, Lassi, Gulab Jamun)
  ];
  const totalWeight = itemWeights.reduce((a, b) => a + b, 0);

  const pickWeightedItem = (seedVal) => {
    let r = seededRandom(seedVal) * totalWeight;
    for (let i = 0; i < MENU_CATALOG.length; i++) {
      r -= itemWeights[i];
      if (r <= 0) return MENU_CATALOG[i];
    }
    return MENU_CATALOG[0];
  };

  let txCounter = 1;
  for (let d = 0; d < totalDays; d++) {
    const curDate = new Date(startDate.getTime() + d * 86400000);
    const ddmmyyyyDate = formatDateDDMMYYYY(curDate);
    const monthIdx = curDate.getUTCMonth(); // 4=May ... 9=Oct
    const dayOfWeek = curDate.getUTCDay(); // 0=Sun, 5=Fri, 6=Sat

    // Gradual 6-month growth trend + weekend & festive dining boost
    const monthGrowthFactor = 1 + (monthIdx - 4) * 0.07;
    const isWeekend = dayOfWeek === 0 || dayOfWeek === 5 || dayOfWeek === 6;
    const baseDailyOrders = isWeekend ? 8 : 5;
    const dailyCount = Math.max(3, Math.round(baseDailyOrders * monthGrowthFactor + (seededRandom(d * 19) * 3 - 1)));

    for (let o = 0; o < dailyCount; o++) {
      const seed = d * 100 + o * 17 + 7;
      let custIdx = Math.floor(seededRandom(seed) * customers.length);
      let cust = customers[custIdx];
      if (cust.profile === 'single_or_declining' && monthIdx > 6) {
        custIdx = Math.floor(seededRandom(seed + 99) * 60);
        cust = customers[custIdx];
      }

      const useFav = seededRandom(seed + 3) < 0.65;
      let chosenItem;
      if (useFav) {
        const favRoll = seededRandom(seed + 5);
        chosenItem = favRoll < 0.48 ? cust.fav1 : favRoll < 0.8 ? cust.fav2 : cust.fav3;
      } else {
        chosenItem = pickWeightedItem(seed + 13);
      }

      // Breads & Chai often ordered in higher quantities per table
      const isBreadOrChai = chosenItem.category === 'Indian Breads' || chosenItem.name === 'Masala Chai';
      const qtyRoll = seededRandom(seed + 29);
      const quantity = isBreadOrChai
        ? (qtyRoll > 0.7 ? 4 : qtyRoll > 0.35 ? 2 : 1)
        : (qtyRoll > 0.85 ? 3 : qtyRoll > 0.60 ? 2 : 1);
      const unitPrice = chosenItem.price;
      const totalAmount = Number((quantity * unitPrice).toFixed(2));

      rows.push({
        'Transaction ID': `INV-IN-${String(txCounter++).padStart(5, '0')}`,
        'Transaction Date': ddmmyyyyDate,
        'Customer ID': cust.customerId,
        'Customer Email ID': cust.email,
        'Item Name': chosenItem.name,
        'Category': chosenItem.category,
        'Quantity': quantity,
        'Unit Price': unitPrice,
        'Total Amount': totalAmount,
      });
    }
  }

  return rows;
};

const generateSampleCustomerSummaryDataset = () => {
  // Build directly from the exact Customer-Summary schema specified in the prompt with Indian Menu Items
  const rows = [];
  const totalCustomers = 115;

  const itemRevenueMap = {};
  MENU_CATALOG.forEach((it) => { itemRevenueMap[it.name] = 0; });

  const tempRows = [];

  for (let i = 0; i < totalCustomers; i++) {
    const idNum = 2001 + i;
    const s1 = i * 13 + 7;
    const isVip = seededRandom(s1) > 0.78;
    const isRepeat = seededRandom(s1 + 1) > 0.22;

    const useCombo = seededRandom(s1 + 9) > 0.35;
    let item1, item2, item3;

    if (useCombo) {
      const combo = INDIAN_MEAL_COMBOS[i % INDIAN_MEAL_COMBOS.length];
      item1 = combo[0];
      item2 = combo[1];
      item3 = combo[2];
    } else {
      const idx1 = Math.floor(seededRandom(s1 + 2) * 10);
      let idx2 = Math.floor(seededRandom(s1 + 3) * 13);
      if (idx2 === idx1) idx2 = (idx2 + 1) % MENU_CATALOG.length;
      let idx3 = Math.floor(seededRandom(s1 + 4) * MENU_CATALOG.length);
      while (idx3 === idx1 || idx3 === idx2) {
        idx3 = (idx3 + 1) % MENU_CATALOG.length;
      }
      item1 = MENU_CATALOG[idx1];
      item2 = MENU_CATALOG[idx2];
      item3 = MENU_CATALOG[idx3];
    }

    const freq1 = isVip
      ? Math.floor(seededRandom(s1 + 5) * 14) + 12
      : isRepeat
      ? Math.floor(seededRandom(s1 + 5) * 7) + 4
      : 1;

    const freq2 = isRepeat
      ? Math.max(1, Math.floor(freq1 * (0.55 + seededRandom(s1 + 6) * 0.35)))
      : 0;

    const freq3 = isRepeat && freq2 > 1
      ? Math.max(1, Math.floor(freq2 * (0.45 + seededRandom(s1 + 7) * 0.35)))
      : 0;

    const rev1 = freq1 * item1.price;
    const rev2 = freq2 * item2.price;
    const rev3 = freq3 * item3.price;

    // Additional long-tail orders by this customer beyond top 3 items
    const extraFactor = isVip ? 1.18 : isRepeat ? 1.08 : 1.0;
    const totalCustomerRev = Number(Math.round((rev1 + rev2 + rev3) * extraFactor));

    itemRevenueMap[item1.name] = (itemRevenueMap[item1.name] || 0) + rev1;
    if (freq2 > 0) itemRevenueMap[item2.name] = (itemRevenueMap[item2.name] || 0) + rev2;
    if (freq3 > 0) itemRevenueMap[item3.name] = (itemRevenueMap[item3.name] || 0) + rev3;

    tempRows.push({
      'Customer ID': `CUST-IN-${idNum}`,
      'Customer Email ID': `diner${idNum}@indianbistro.in`,
      'Most Ordered Item': item1.name,
      'Frequency of Most Ordered Item': freq1,
      'Price of Most Ordered Item': item1.price,
      'Second Most Ordered Item': freq2 > 0 ? item2.name : '',
      'Frequency of Second Most Ordered Item': freq2,
      'Price of Second Most Ordered Item': freq2 > 0 ? item2.price : 0,
      'Third Most Ordered Item': freq3 > 0 ? item3.name : '',
      'Frequency of Third Most Ordered Item': freq3,
      'Price of Third Most Ordered Item': freq3 > 0 ? item3.price : 0,
      'Total Revenue Generated by Each Customer': totalCustomerRev,
    });
  }

  return tempRows.map((r) => ({
    ...r,
    'Revenue Generated by Each Menu Item': Number((itemRevenueMap[r['Most Ordered Item']] || 0).toFixed(2)),
  }));
};

// ============================================================================
// CSV / EXCEL PARSER & SCHEMA AUTO-DETECTION ENGINE
// ============================================================================

const parseCSVText = (text) => {
  const lines = [];
  let curLine = [];
  let curField = '';
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    const nextCh = text[i + 1];

    if (inQuotes) {
      if (ch === '"' && nextCh === '"') {
        curField += '"';
        i++;
      } else if (ch === '"') {
        inQuotes = false;
      } else {
        curField += ch;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
      } else if (ch === ',') {
        curLine.push(curField.trim());
        curField = '';
      } else if (ch === '\n' || (ch === '\r' && nextCh === '\n')) {
        curLine.push(curField.trim());
        if (curLine.some((cell) => cell !== '')) {
          lines.push(curLine);
        }
        curLine = [];
        curField = '';
        if (ch === '\r') i++;
      } else {
        curField += ch;
      }
    }
  }
  if (curField !== '' || curLine.length > 0) {
    curLine.push(curField.trim());
    if (curLine.some((cell) => cell !== '')) {
      lines.push(curLine);
    }
  }

  if (lines.length < 2) return { headers: [], rows: [] };
  const headers = lines[0].map((h, idx) => h || `Column_${idx + 1}`);
  const rows = lines.slice(1).map((line) => {
    const obj = {};
    headers.forEach((h, idx) => {
      obj[h] = line[idx] !== undefined ? line[idx] : '';
    });
    return obj;
  });
  return { headers, rows };
};

// Schema Field Definitions for Auto-Mapping
const SUMMARY_SCHEMA_FIELDS = [
  { key: 'customerId', label: 'Customer ID', required: true, aliases: ['customerid', 'custid', 'id', 'clientid', 'memberid'] },
  { key: 'customerEmail', label: 'Customer Email ID', required: false, aliases: ['customeremailid', 'customeremail', 'email', 'emailid', 'mail'] },
  { key: 'item1Name', label: 'Most Ordered Item', required: true, aliases: ['mostordereditem', 'firstmostordereditem', 'topitem', 'item1', 'primaryitem', 'favoriteitem'] },
  { key: 'item1Freq', label: 'Frequency of Most Ordered Item', required: true, aliases: ['frequencyofmostordereditem', 'mostordereditemfrequency', 'freq1', 'frequency1', 'topitemfrequency', 'firstitemfrequency'] },
  { key: 'item1Price', label: 'Price of Most Ordered Item', required: true, aliases: ['priceofmostordereditem', 'mostordereditemprice', 'price1', 'topitemprice', 'firstitemprice'] },
  { key: 'item2Name', label: 'Second Most Ordered Item', required: false, aliases: ['secondmostordereditem', 'item2', 'seconditem', '2ndmostordereditem'] },
  { key: 'item2Freq', label: 'Frequency of Second Most Ordered Item', required: false, aliases: ['frequencyofsecondmostordereditem', 'secondmostordereditemfrequency', 'freq2', 'frequency2', '2nditemfrequency'] },
  { key: 'item2Price', label: 'Price of Second Most Ordered Item', required: false, aliases: ['priceofsecondmostordereditem', 'secondmostordereditemprice', 'price2', '2nditemprice'] },
  { key: 'item3Name', label: 'Third Most Ordered Item', required: false, aliases: ['thirdmostordereditem', 'item3', 'thirditem', '3rdmostordereditem'] },
  { key: 'item3Freq', label: 'Frequency of Third Most Ordered Item', required: false, aliases: ['frequencyofthirdmostordereditem', 'thirdmostordereditemfrequency', 'freq3', 'frequency3', '3rditemfrequency'] },
  { key: 'item3Price', label: 'Price of Third Most Ordered Item', required: false, aliases: ['priceofthirdmostordereditem', 'thirdmostordereditemprice', 'price3', '3rditemprice'] },
  { key: 'totalCustomerRevenue', label: 'Total Revenue Generated by Each Customer', required: false, aliases: ['totalrevenuegeneratedbyeachcustomer', 'totalcustomerrevenue', 'customerspend', 'totalrevenue', 'totalspent', 'lifetimevalue'] },
  { key: 'menuItemRevenue', label: 'Revenue Generated by Each Menu Item', required: false, aliases: ['revenuegeneratedbyeachmenuitem', 'menuitemrevenue', 'itemrevenue'] },
];

const TRANSACTION_SCHEMA_FIELDS = [
  { key: 'transactionDate', label: 'Transaction Date', required: false, aliases: ['transactiondate', 'date', 'orderdate', 'timestamp', 'datetime', 'day'] },
  { key: 'customerId', label: 'Customer ID', required: false, aliases: ['customerid', 'custid', 'clientid', 'guestid', 'memberid'] },
  { key: 'customerEmail', label: 'Customer Email ID', required: false, aliases: ['customeremailid', 'customeremail', 'email', 'emailid'] },
  { key: 'itemName', label: 'Item Name', required: true, aliases: ['itemname', 'menuitem', 'item', 'product', 'productname', 'dish', 'dishname'] },
  { key: 'category', label: 'Category (Optional)', required: false, aliases: ['category', 'itemcategory', 'menucategory', 'department', 'group'] },
  { key: 'quantity', label: 'Quantity', required: true, aliases: ['quantity', 'qty', 'units', 'count', 'orders', 'itemsold'] },
  { key: 'unitPrice', label: 'Unit Price', required: false, aliases: ['unitprice', 'price', 'itemprice', 'rate', 'cost'] },
  { key: 'totalAmount', label: 'Total Amount / Revenue', required: false, aliases: ['totalamount', 'total', 'revenue', 'amount', 'sales', 'lineamount', 'netamount'] },
];

const autoDetectSchemaAndMapping = (headers) => {
  const normHeaders = headers.map((h) => ({ raw: h, norm: normalizeHeader(h) }));

  // Check if headers strongly match Customer-Summary schema
  const hasMostOrdered = normHeaders.some((h) =>
    h.norm.includes('mostordered') || h.norm.includes('secondmost') || h.norm.includes('thirdmost')
  );

  const schemaType = hasMostOrdered ? 'customer_summary' : 'transaction';
  const targetFields = schemaType === 'customer_summary' ? SUMMARY_SCHEMA_FIELDS : TRANSACTION_SCHEMA_FIELDS;

  const mapping = {};
  targetFields.forEach((field) => {
    // 1. Exact alias match
    let match = normHeaders.find((h) => field.aliases.includes(h.norm));
    // 2. Substring match if no exact match
    if (!match) {
      match = normHeaders.find((h) => field.aliases.some((alias) => h.norm.includes(alias) || alias.includes(h.norm)));
    }
    mapping[field.key] = match ? match.raw : '';
  });

  return { schemaType, mapping };
};

// ============================================================================
// CORE ANALYTICS & BUSINESS INTELLIGENCE ENGINE
// ============================================================================

const processRestaurantData = (rawRows, schemaType, mapping, datasetMeta = {}) => {
  const hygiene = {
    rawRowCount: rawRows.length,
    validRowCount: 0,
    duplicatesRemoved: 0,
    missingValuesHandled: 0,
    unattributedRevenueTotal: 0,
    hasDateData: false,
    dateRange: null,
  };

  const seenRowSignatures = new Set();
  const cleanRows = [];

  // 1. Deduplicate and sanitize rows
  rawRows.forEach((row) => {
    const sig = JSON.stringify(row);
    if (seenRowSignatures.has(sig)) {
      hygiene.duplicatesRemoved++;
      return;
    }
    seenRowSignatures.add(sig);
    cleanRows.push(row);
  });

  // Maps for aggregation
  const itemMap = {}; // itemName -> { name, category, unitsSold, revenue, orderCount, customerSet, prices: [], explicitItemRevenue: null }
  const customerMap = {}; // customerId -> { id, email, totalRevenue, accountedItemRevenue, totalUnits, orderCount, itemsOrdered: { itemName: qty }, firstDate, lastDate, dates: [] }
  const coOrderPairs = {}; // "Item A + Item B" -> count
  const dailyMap = {}; // YYYY-MM-DD -> { date, revenue, units, orders, customers: Set }
  const monthlyMap = {}; // YYYY-MM -> { month, label, revenue, units, orders, customers: Set }
  const weeklyMap = {}; // YYYY-Wxx -> { week, revenue, units, orders }

  const ensureItem = (name, category = 'Menu Item') => {
    const cleanName = String(name || '').trim();
    if (!cleanName) return null;
    if (!itemMap[cleanName]) {
      itemMap[cleanName] = {
        name: cleanName,
        category: category || 'Menu Item',
        unitsSold: 0,
        revenue: 0,
        orderFrequency: 0,
        customerIds: new Set(),
        prices: [],
        explicitCatalogRevenue: null,
      };
    }
    return itemMap[cleanName];
  };

  const ensureCustomer = (id, email = '') => {
    const cleanId = String(id || '').trim() || `GUEST-${hygiene.validRowCount + 1}`;
    if (!customerMap[cleanId]) {
      customerMap[cleanId] = {
        id: cleanId,
        email: String(email || '').trim() || '—',
        totalRevenue: 0,
        accountedItemRevenue: 0,
        totalUnits: 0,
        orderFrequency: 0,
        itemsOrdered: {},
        explicitTopItems: [],
        dates: [],
      };
    } else if (email && customerMap[cleanId].email === '—') {
      customerMap[cleanId].email = String(email).trim();
    }
    return customerMap[cleanId];
  };

  const recordCoOrder = (itemsList) => {
    const unique = Array.from(new Set(itemsList.filter(Boolean)));
    for (let i = 0; i < unique.length; i++) {
      for (let j = i + 1; j < unique.length; j++) {
        const pair = [unique[i], unique[j]].sort().join(' + ');
        coOrderPairs[pair] = (coOrderPairs[pair] || 0) + 1;
      }
    }
  };

  // ==========================================================================
  // BRANCH A: CUSTOMER-SUMMARY SCHEMA PROCESSING
  // ==========================================================================
  if (schemaType === 'customer_summary') {
    cleanRows.forEach((row) => {
      const custId = row[mapping.customerId];
      const item1Name = row[mapping.item1Name];

      if (!custId && !item1Name) {
        hygiene.missingValuesHandled++;
        return;
      }
      hygiene.validRowCount++;

      const cust = ensureCustomer(custId, row[mapping.customerEmail]);
      const customerRowItems = [];

      // Process 1st, 2nd, and 3rd Most Ordered Items
      const slots = [
        { nameKey: 'item1Name', freqKey: 'item1Freq', priceKey: 'item1Price', rank: 1 },
        { nameKey: 'item2Name', freqKey: 'item2Freq', priceKey: 'item2Price', rank: 2 },
        { nameKey: 'item3Name', freqKey: 'item3Freq', priceKey: 'item3Price', rank: 3 },
      ];

      let rowAccountedRev = 0;
      let rowUnits = 0;

      slots.forEach((slot) => {
        const rawName = mapping[slot.nameKey] ? row[mapping[slot.nameKey]] : '';
        const freq = mapping[slot.freqKey] ? parseNumeric(row[mapping[slot.freqKey]]) : 0;
        const price = mapping[slot.priceKey] ? parseNumeric(row[mapping[slot.priceKey]]) : 0;

        if (rawName && String(rawName).trim() !== '' && freq > 0) {
          const itemObj = ensureItem(rawName);
          if (itemObj) {
            const slotRev = freq * price;
            itemObj.unitsSold += freq;
            itemObj.revenue += slotRev;
            itemObj.orderFrequency += 1; // 1 customer has this as a frequent item
            itemObj.customerIds.add(cust.id);
            if (price > 0) itemObj.prices.push(price);

            // If row contains "Revenue Generated by Each Menu Item" on slot 1, record it
            if (slot.rank === 1 && mapping.menuItemRevenue && row[mapping.menuItemRevenue] !== undefined && row[mapping.menuItemRevenue] !== '') {
              const expRev = parseNumeric(row[mapping.menuItemRevenue]);
              if (expRev > 0) itemObj.explicitCatalogRevenue = Math.max(itemObj.explicitCatalogRevenue || 0, expRev);
            }

            rowAccountedRev += slotRev;
            rowUnits += freq;
            cust.itemsOrdered[itemObj.name] = (cust.itemsOrdered[itemObj.name] || 0) + freq;
            cust.explicitTopItems.push({ rank: slot.rank, name: itemObj.name, freq, price, revenue: slotRev });
            customerRowItems.push(itemObj.name);
          }
        } else if (rawName && freq <= 0) {
          hygiene.missingValuesHandled++;
        }
      });

      const explicitTotalCustRev = mapping.totalCustomerRevenue
        ? parseNumeric(row[mapping.totalCustomerRevenue])
        : 0;

      const finalCustRev = explicitTotalCustRev > 0 ? explicitTotalCustRev : rowAccountedRev;
      if (explicitTotalCustRev > rowAccountedRev) {
        hygiene.unattributedRevenueTotal += (explicitTotalCustRev - rowAccountedRev);
      }

      cust.totalRevenue += finalCustRev;
      cust.accountedItemRevenue += rowAccountedRev;
      cust.totalUnits += rowUnits;
      cust.orderFrequency += rowUnits; // Total item order occasions

      if (customerRowItems.length >= 2) {
        recordCoOrder(customerRowItems);
      }
    });
  }
  // ==========================================================================
  // BRANCH B: TRANSACTION-LEVEL SCHEMA PROCESSING
  // ==========================================================================
  else {
    cleanRows.forEach((row) => {
      const itemName = mapping.itemName ? row[mapping.itemName] : '';
      if (!itemName || String(itemName).trim() === '') {
        hygiene.missingValuesHandled++;
        return;
      }

      const qtyRaw = mapping.quantity ? parseNumeric(row[mapping.quantity]) : 1;
      const qty = qtyRaw > 0 ? qtyRaw : 1;
      const unitPriceRaw = mapping.unitPrice ? parseNumeric(row[mapping.unitPrice]) : 0;
      const totalAmountRaw = mapping.totalAmount ? parseNumeric(row[mapping.totalAmount]) : 0;

      let lineRevenue = totalAmountRaw;
      let unitPrice = unitPriceRaw;

      if (lineRevenue <= 0 && unitPrice > 0) {
        lineRevenue = qty * unitPrice;
        hygiene.missingValuesHandled++;
      } else if (unitPrice <= 0 && lineRevenue > 0) {
        unitPrice = lineRevenue / qty;
        hygiene.missingValuesHandled++;
      }

      hygiene.validRowCount++;

      const category = mapping.category && row[mapping.category] ? String(row[mapping.category]).trim() : 'Menu Item';
      const itemObj = ensureItem(itemName, category);
      const cust = ensureCustomer(
        mapping.customerId ? row[mapping.customerId] : `ANON-${hygiene.validRowCount}`,
        mapping.customerEmail ? row[mapping.customerEmail] : ''
      );

      itemObj.unitsSold += qty;
      itemObj.revenue += lineRevenue;
      itemObj.orderFrequency += 1;
      itemObj.customerIds.add(cust.id);
      if (unitPrice > 0) itemObj.prices.push(unitPrice);

      cust.totalRevenue += lineRevenue;
      cust.accountedItemRevenue += lineRevenue;
      cust.totalUnits += qty;
      cust.orderFrequency += 1;
      cust.itemsOrdered[itemObj.name] = (cust.itemsOrdered[itemObj.name] || 0) + qty;

      // Date parsing if available
      if (mapping.transactionDate && row[mapping.transactionDate]) {
        const rawDate = row[mapping.transactionDate];
        const parsedDate = parseFlexibleDate(rawDate);
        if (parsedDate && !isNaN(parsedDate.getTime())) {
          hygiene.hasDateData = true;
          const isoDate = parsedDate.toISOString().split('T')[0]; // YYYY-MM-DD for chronological sorting
          const displayDate = formatDateDDMMYYYY(parsedDate); // DD-MM-YYYY for Indian display
          const monthKey = isoDate.slice(0, 7); // YYYY-MM
          const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
          const monthLabel = `${monthNames[parsedDate.getUTCMonth()]} ${parsedDate.getUTCFullYear()}`;

          // ISO week approximation
          const startOfYear = new Date(Date.UTC(parsedDate.getUTCFullYear(), 0, 1));
          const weekNum = Math.ceil((((parsedDate - startOfYear) / 86400000) + startOfYear.getUTCDay() + 1) / 7);
          const weekKey = `${parsedDate.getUTCFullYear()}-W${String(weekNum).padStart(2, '0')}`;

          cust.dates.push(isoDate);

          if (!dailyMap[isoDate]) {
            dailyMap[isoDate] = { date: isoDate, displayDate, revenue: 0, units: 0, orders: 0, customers: new Set() };
          }
          dailyMap[isoDate].revenue += lineRevenue;
          dailyMap[isoDate].units += qty;
          dailyMap[isoDate].orders += 1;
          dailyMap[isoDate].customers.add(cust.id);

          if (!monthlyMap[monthKey]) {
            monthlyMap[monthKey] = { month: monthKey, label: monthLabel, revenue: 0, units: 0, orders: 0, customers: new Set() };
          }
          monthlyMap[monthKey].revenue += lineRevenue;
          monthlyMap[monthKey].units += qty;
          monthlyMap[monthKey].orders += 1;
          monthlyMap[monthKey].customers.add(cust.id);

          if (!weeklyMap[weekKey]) {
            weeklyMap[weekKey] = { week: weekKey, revenue: 0, units: 0, orders: 0 };
          }
          weeklyMap[weekKey].revenue += lineRevenue;
          weeklyMap[weekKey].units += qty;
          weeklyMap[weekKey].orders += 1;
        }
      }
    });
  }

  // ==========================================================================
  // POST-PROCESSING & METRIC SYNTHESIS
  // ==========================================================================

  // 1. Finalize Menu Items List
  const itemsList = Object.values(itemMap).map((it) => {
    const avgPrice = it.prices.length > 0
      ? it.prices.reduce((a, b) => a + b, 0) / it.prices.length
      : it.unitsSold > 0 ? it.revenue / it.unitsSold : 0;
    const effectiveRevenue = it.revenue;
    return {
      name: it.name,
      category: it.category,
      unitsSold: it.unitsSold,
      revenue: Number(effectiveRevenue.toFixed(2)),
      explicitCatalogRevenue: it.explicitCatalogRevenue,
      avgPrice: Number(avgPrice.toFixed(2)),
      orderFrequency: it.orderFrequency,
      uniqueCustomers: it.customerIds.size,
    };
  });

  const totalItemRevenue = itemsList.reduce((sum, it) => sum + it.revenue, 0);
  const totalUnitsSold = itemsList.reduce((sum, it) => sum + it.unitsSold, 0);

  // Sort items by revenue & calculate Pareto cumulative %
  const itemsByRevenue = [...itemsList].sort((a, b) => b.revenue - a.revenue);
  let cumRev = 0;
  itemsByRevenue.forEach((it, idx) => {
    cumRev += it.revenue;
    it.revenueShare = totalItemRevenue > 0 ? (it.revenue / totalItemRevenue) * 100 : 0;
    it.unitShare = totalUnitsSold > 0 ? (it.unitsSold / totalUnitsSold) * 100 : 0;
    it.cumulativeRevenueShare = totalItemRevenue > 0 ? (cumRev / totalItemRevenue) * 100 : 0;
    it.revenueRank = idx + 1;
  });

  // Menu Engineering Matrix Benchmarks (Average Volume & Average Revenue)
  const avgItemUnits = itemsList.length > 0 ? totalUnitsSold / itemsList.length : 0;
  const avgItemRev = itemsList.length > 0 ? totalItemRevenue / itemsList.length : 0;

  itemsList.forEach((it) => {
    const highVol = it.unitsSold >= avgItemUnits;
    const highRev = it.revenue >= avgItemRev;
    if (highVol && highRev) {
      it.quadrant = 'Star';
      it.quadrantLabel = 'Prioritize (Star)';
      it.recommendation = 'Flagship Indian menu leader. Maintain consistent spice/quality standards and prime menu visibility.';
    } else if (highVol && !highRev) {
      it.quadrant = 'Plowhorse';
      it.quadrantLabel = 'Reprice / Combo (Plowhorse)';
      it.recommendation = 'High order volume (e.g. breads, tiffin, chai) but lower revenue yield. Test a modest ₹10–₹25 price increase or pair in a Thali/combo.';
    } else if (!highVol && highRev) {
      it.quadrant = 'Puzzle';
      it.quadrantLabel = 'Promote (Puzzle)';
      it.recommendation = 'Strong revenue per plate with lower order volume. Train staff to recommend as a table starter or main upgrade.';
    } else {
      it.quadrant = 'Dog';
      it.quadrantLabel = 'Reconsider / Remove (Dog)';
      it.recommendation = 'Low volume and low revenue contribution. Consider menu rationalisation or offering only as a weekend/festive special.';
    }
  });

  const itemsByUnits = [...itemsList].sort((a, b) => b.unitsSold - a.unitsSold);
  const itemsByFrequency = [...itemsList].sort((a, b) => b.orderFrequency - a.orderFrequency);
  const lowestPerformingItems = [...itemsList].sort((a, b) => a.revenue - b.revenue);

  // 2. Finalize Customers List & Segmentation
  const dailyDatesSorted = Object.keys(dailyMap).sort();
  const latestDatasetDate = dailyDatesSorted.length > 0
    ? new Date(dailyDatesSorted[dailyDatesSorted.length - 1])
    : null;

  if (dailyDatesSorted.length > 0) {
    hygiene.dateRange = {
      start: formatDateDDMMYYYY(dailyDatesSorted[0]),
      end: formatDateDDMMYYYY(dailyDatesSorted[dailyDatesSorted.length - 1]),
      daysCount: dailyDatesSorted.length,
    };
  }

  const customersList = Object.values(customerMap).map((c) => {
    // Derive 1st, 2nd, 3rd preferred items if not explicitly populated
    let topItems = c.explicitTopItems;
    if (!topItems || topItems.length === 0) {
      const sortedEntries = Object.entries(c.itemsOrdered)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 3);
      topItems = sortedEntries.map(([name, freq], i) => {
        const itObj = itemMap[name];
        const price = itObj && itObj.prices.length > 0 ? itObj.prices[0] : 0;
        return { rank: i + 1, name, freq, price, revenue: freq * price };
      });
    }

    // Recency calculation if dates exist
    let daysSinceLastOrder = null;
    let firstOrderDate = null;
    let lastOrderDate = null;
    if (c.dates.length > 0 && latestDatasetDate) {
      const sortedCustDates = [...c.dates].sort();
      firstOrderDate = formatDateDDMMYYYY(sortedCustDates[0]);
      lastOrderDate = formatDateDDMMYYYY(sortedCustDates[sortedCustDates.length - 1]);
      const lastDt = new Date(sortedCustDates[sortedCustDates.length - 1]);
      daysSinceLastOrder = Math.max(0, Math.round((latestDatasetDate - lastDt) / 86400000));
    }

    // Compute co-orders for transaction schema customers
    if (schemaType === 'transaction' && topItems.length >= 2) {
      recordCoOrder(topItems.map((t) => t.name));
    }

    const isRepeat = c.totalUnits > 1 || c.orderFrequency > 1;
    const avgSpendPerUnit = c.totalUnits > 0 ? c.totalRevenue / c.totalUnits : c.totalRevenue;

    return {
      id: c.id,
      email: c.email,
      totalRevenue: Number(c.totalRevenue.toFixed(2)),
      accountedItemRevenue: Number(c.accountedItemRevenue.toFixed(2)),
      unattributedRevenue: Number(Math.max(0, c.totalRevenue - c.accountedItemRevenue).toFixed(2)),
      totalUnits: c.totalUnits,
      orderFrequency: c.orderFrequency,
      avgSpendPerUnit: Number(avgSpendPerUnit.toFixed(2)),
      isRepeat,
      top1Item: topItems[0]?.name || '—',
      top1Freq: topItems[0]?.freq || 0,
      top2Item: topItems[1]?.name || '—',
      top2Freq: topItems[1]?.freq || 0,
      top3Item: topItems[2]?.name || '—',
      top3Freq: topItems[2]?.freq || 0,
      firstOrderDate,
      lastOrderDate,
      daysSinceLastOrder,
    };
  });

  const totalCustomerRevenue = customersList.reduce((sum, c) => sum + c.totalRevenue, 0);
  const overallRevenue = Math.max(totalCustomerRevenue, totalItemRevenue);
  const avgCustomerSpend = customersList.length > 0 ? overallRevenue / customersList.length : 0;
  const avgCustomerUnits = customersList.length > 0 ? totalUnitsSold / customersList.length : 0;

  // Segment Customers
  customersList.forEach((c) => {
    const highSpend = c.totalRevenue >= avgCustomerSpend * 1.25;
    const highFreq = c.totalUnits >= avgCustomerUnits * 1.2;
    const isInactive = c.daysSinceLastOrder !== null && c.daysSinceLastOrder > 60;

    if (isInactive && c.totalUnits >= 2) {
      c.segment = 'At-Risk / Declining';
    } else if (highSpend && highFreq) {
      c.segment = 'VIP Champions';
    } else if (highFreq && !highSpend) {
      c.segment = 'High-Frequency Loyalists';
    } else if (highSpend && !highFreq) {
      c.segment = 'Big-Ticket Occasional';
    } else if (!c.isRepeat) {
      c.segment = 'Single-Visit / New';
    } else {
      c.segment = 'Core Regulars';
    }
  });

  const customersBySpend = [...customersList].sort((a, b) => b.totalRevenue - a.totalRevenue);
  const customersByFrequency = [...customersList].sort((a, b) => b.totalUnits - a.totalUnits);

  let cumCustRev = 0;
  customersBySpend.forEach((c, idx) => {
    cumCustRev += c.totalRevenue;
    c.revenueShare = overallRevenue > 0 ? (c.totalRevenue / overallRevenue) * 100 : 0;
    c.cumulativeRevenueShare = overallRevenue > 0 ? (cumCustRev / overallRevenue) * 100 : 0;
    c.spendRank = idx + 1;
  });

  // 3. Time-Series Trends & Period Comparisons
  const dailySeries = Object.values(dailyMap)
    .sort((a, b) => a.date.localeCompare(b.date))
    .map((d) => ({
      date: d.displayDate || formatDateDDMMYYYY(d.date),
      isoDate: d.date,
      revenue: Number(d.revenue.toFixed(2)),
      units: d.units,
      orders: d.orders,
      activeCustomers: d.customers.size,
    }));

  const weeklySeries = Object.values(weeklyMap)
    .sort((a, b) => a.week.localeCompare(b.week))
    .map((w) => ({
      period: w.week,
      revenue: Number(w.revenue.toFixed(2)),
      units: w.units,
      orders: w.orders,
    }));

  const monthlySeries = Object.values(monthlyMap)
    .sort((a, b) => a.month.localeCompare(b.month))
    .map((m, idx, arr) => {
      const prev = idx > 0 ? arr[idx - 1] : null;
      const momGrowth = prev && prev.revenue > 0 ? ((m.revenue - prev.revenue) / prev.revenue) * 100 : null;
      return {
        month: m.month,
        label: m.label,
        revenue: Number(m.revenue.toFixed(2)),
        units: m.units,
        orders: m.orders,
        activeCustomers: m.customers.size,
        momGrowth: momGrowth !== null ? Number(momGrowth.toFixed(1)) : null,
      };
    });

  // Period Comparison Calculations
  let periodComparison = null;
  if (monthlySeries.length >= 2) {
    const lastM = monthlySeries[monthlySeries.length - 1];
    const prevM = monthlySeries[monthlySeries.length - 2];
    const lastMonthDelta = prevM.revenue > 0 ? ((lastM.revenue - prevM.revenue) / prevM.revenue) * 100 : 0;

    const mid = Math.floor(monthlySeries.length / 2);
    const firstHalf = monthlySeries.slice(0, mid);
    const secondHalf = monthlySeries.slice(mid);
    const firstHalfRev = firstHalf.reduce((s, m) => s + m.revenue, 0);
    const secondHalfRev = secondHalf.reduce((s, m) => s + m.revenue, 0);
    const halfDelta = firstHalfRev > 0 ? ((secondHalfRev - firstHalfRev) / firstHalfRev) * 100 : 0;

    periodComparison = {
      latestMonthLabel: lastM.label,
      latestMonthRevenue: lastM.revenue,
      previousMonthLabel: prevM.label,
      previousMonthRevenue: prevM.revenue,
      latestVsPrevPct: Number(lastMonthDelta.toFixed(1)),
      firstHalfLabel: `${firstHalf[0]?.label} – ${firstHalf[firstHalf.length - 1]?.label}`,
      firstHalfRevenue: Number(firstHalfRev.toFixed(2)),
      secondHalfLabel: `${secondHalf[0]?.label} – ${secondHalf[secondHalf.length - 1]?.label}`,
      secondHalfRevenue: Number(secondHalfRev.toFixed(2)),
      halfOverHalfPct: Number(halfDelta.toFixed(1)),
    };
  }

  // 4. Repeat Customer Percentage & Top Co-Ordered Pairs
  const repeatCustomersCount = customersList.filter((c) => c.isRepeat).length;
  const repeatCustomerPct = customersList.length > 0
    ? (repeatCustomersCount / customersList.length) * 100
    : 0;

  const topCoOrders = Object.entries(coOrderPairs)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8)
    .map(([pair, count]) => ({ pair, count }));

  // 5. Customer Segment Distribution
  const segmentMap = {};
  customersList.forEach((c) => {
    if (!segmentMap[c.segment]) {
      segmentMap[c.segment] = { segment: c.segment, count: 0, revenue: 0, units: 0 };
    }
    segmentMap[c.segment].count += 1;
    segmentMap[c.segment].revenue += c.totalRevenue;
    segmentMap[c.segment].units += c.totalUnits;
  });
  const customerSegments = Object.values(segmentMap)
    .map((s) => ({
      ...s,
      revenue: Number(s.revenue.toFixed(2)),
      shareOfCustomers: customersList.length > 0 ? Number(((s.count / customersList.length) * 100).toFixed(1)) : 0,
      shareOfRevenue: overallRevenue > 0 ? Number(((s.revenue / overallRevenue) * 100).toFixed(1)) : 0,
    }))
    .sort((a, b) => b.revenue - a.revenue);

  // 6. Top 3 Items Revenue Concentration
  const top3ItemsRev = itemsByRevenue.slice(0, 3).reduce((s, it) => s + it.revenue, 0);
  const top3ItemsRevPct = totalItemRevenue > 0 ? (top3ItemsRev / totalItemRevenue) * 100 : 0;

  // Top 10% Customers Revenue Share
  const top10PctCustCount = Math.max(1, Math.ceil(customersList.length * 0.1));
  const top10PctCustRev = customersBySpend.slice(0, top10PctCustCount).reduce((s, c) => s + c.totalRevenue, 0);
  const top10PctCustShare = overallRevenue > 0 ? (top10PctCustRev / overallRevenue) * 100 : 0;

  // Quadrants lists
  const stars = itemsByRevenue.filter((i) => i.quadrant === 'Star');
  const plowhorses = itemsByUnits.filter((i) => i.quadrant === 'Plowhorse');
  const puzzles = itemsByRevenue.filter((i) => i.quadrant === 'Puzzle');
  const dogs = lowestPerformingItems.filter((i) => i.quadrant === 'Dog');

  return {
    meta: {
      datasetName: datasetMeta.datasetName || 'Uploaded Sales Dataset',
      isDemo: Boolean(datasetMeta.isDemo),
      schemaType,
      processedAt: `${getISTTimeString()} IST`,
    },
    hygiene,
    kpis: {
      totalRevenue: Number(overallRevenue.toFixed(2)),
      totalItemRevenue: Number(totalItemRevenue.toFixed(2)),
      totalUnitsSold,
      totalCustomers: customersList.length,
      avgRevenuePerCustomer: Number(avgCustomerSpend.toFixed(2)),
      avgItemPrice: totalUnitsSold > 0 ? Number((totalItemRevenue / totalUnitsSold).toFixed(2)) : 0,
      repeatCustomerPct: Number(repeatCustomerPct.toFixed(1)),
      repeatCustomersCount,
      mostPopularItemByUnits: itemsByUnits[0] || null,
      highestRevenueItem: itemsByRevenue[0] || null,
      top3ItemsRevPct: Number(top3ItemsRevPct.toFixed(1)),
      top10PctCustShare: Number(top10PctCustShare.toFixed(1)),
      avgItemUnits: Number(avgItemUnits.toFixed(1)),
      avgItemRev: Number(avgItemRev.toFixed(2)),
    },
    items: {
      all: itemsByRevenue,
      byRevenue: itemsByRevenue,
      byUnits: itemsByUnits,
      byFrequency: itemsByFrequency,
      lowestPerforming: lowestPerformingItems,
      stars,
      plowhorses,
      puzzles,
      dogs,
    },
    customers: {
      all: customersBySpend,
      bySpend: customersBySpend,
      byFrequency: customersByFrequency,
      segments: customerSegments,
      topCoOrders,
      decliningOrAtRisk: customersBySpend.filter((c) => c.segment === 'At-Risk / Declining'),
    },
    timeSeries: {
      hasDateData: hygiene.hasDateData,
      daily: dailySeries,
      weekly: weeklySeries,
      monthly: monthlySeries,
      periodComparison,
    },
    rawPreview: cleanRows.slice(0, 100),
  };
};

// ============================================================================
// STRUCTURED AI BUSINESS INTELLIGENCE GENERATOR & DETERMINISTIC QA FALLBACK
// ============================================================================

const generateStructuredBIReport = (analytics) => {
  if (!analytics) return [];
  const { kpis, items, customers, timeSeries, hygiene } = analytics;
  const topRevItem = items.byRevenue[0];
  const secondRevItem = items.byRevenue[1];
  const topUnitItem = items.byUnits[0];
  const topPlowhorse = items.plowhorses[0];
  const topPuzzle = items.puzzles[0];
  const bottomDog = items.dogs[0] || items.lowestPerforming[0];
  const topCustomer = customers.bySpend[0];
  const vipSegment = customers.segments.find((s) => s.segment === 'VIP Champions');

  return [
    {
      id: 'performing-well',
      question: 'What is performing well?',
      badge: 'Core Momentum',
      color: 'emerald',
      headline: `${formatCurrency(kpis.totalRevenue)} total revenue with a ${formatPercent(kpis.repeatCustomerPct)} repeat diner rate`,
      analysis: timeSeries.periodComparison
        ? `Sales momentum is strong: revenue in the second half of the period (${timeSeries.periodComparison.secondHalfLabel}) reached ${formatCurrency(timeSeries.periodComparison.secondHalfRevenue)}, representing a ${timeSeries.periodComparison.halfOverHalfPct >= 0 ? '+' : ''}${timeSeries.periodComparison.halfOverHalfPct}% shift compared to ${timeSeries.periodComparison.firstHalfLabel}. Repeat diners account for ${kpis.repeatCustomersCount} of ${kpis.totalCustomers} patrons.`
        : `Across ${formatNumber(kpis.totalCustomers)} active diners and ${formatNumber(kpis.totalUnitsSold)} plates/items sold, your average revenue per customer stands at ${formatCurrency(kpis.avgRevenuePerCustomer)}. ${kpis.repeatCustomersCount} patrons (${formatPercent(kpis.repeatCustomerPct)}) have ordered multiple times, anchoring steady recurring demand.`,
      metrics: [
        { label: 'Total Revenue', value: formatCurrency(kpis.totalRevenue) },
        { label: 'Repeat Diner Rate', value: formatPercent(kpis.repeatCustomerPct) },
        { label: 'Avg Spend / Customer', value: formatCurrency(kpis.avgRevenuePerCustomer) },
      ],
    },
    {
      id: 'top-revenue-items',
      question: 'Which menu items are generating the most revenue?',
      badge: 'Revenue Leaders',
      color: 'indigo',
      headline: topRevItem
        ? `${topRevItem.name} leads with ${formatCurrency(topRevItem.revenue)} (${formatPercent(topRevItem.revenueShare)} of item sales)`
        : 'Insufficient item data',
      analysis: topRevItem
        ? `Your top 3 revenue-generating dishes (${items.byRevenue.slice(0, 3).map((i) => i.name).join(', ')}) collectively drive ${formatPercent(kpis.top3ItemsRevPct)} of total attributed menu revenue. ${topRevItem.name} alone sold ${formatNumber(topRevItem.unitsSold)} units at an average price of ${formatCurrency(topRevItem.avgPrice)}${secondRevItem ? `, followed by ${secondRevItem.name} at ${formatCurrency(secondRevItem.revenue)}.` : '.'}`
        : 'Upload item-level pricing and quantities to calculate revenue leaders.',
      metrics: items.byRevenue.slice(0, 3).map((it) => ({
        label: it.name,
        value: `${formatCurrency(it.revenue)} (${formatPercent(it.revenueShare)})`,
      })),
    },
    {
      id: 'popular-low-revenue',
      question: 'Which items are popular but generating relatively little revenue?',
      badge: 'Plowhorse Analysis',
      color: 'amber',
      headline: topPlowhorse
        ? `${topPlowhorse.name} has high volume (${formatNumber(topPlowhorse.unitsSold)} units) but lower revenue (${formatCurrency(topPlowhorse.revenue)})`
        : `All high-volume items currently exceed the ${formatCurrency(kpis.avgItemRev)} average revenue benchmark`,
      analysis: topPlowhorse
        ? `High-frequency Indian staples in the Plowhorse quadrant—led by ${items.plowhorses.slice(0, 3).map((i) => `${i.name} (${i.unitsSold} units at ${formatCurrency(i.avgPrice)})`).join(', ')}—exceed the menu average of ${formatNumber(kpis.avgItemUnits, 1)} units sold, but generate less than the ${formatCurrency(kpis.avgItemRev)} item revenue benchmark due to lower per-unit pricing.`
        : `Your highest-selling items by volume (${topUnitItem?.name || 'N/A'}) also command strong price points, meaning popular demand is translating directly into top-line revenue.`,
      metrics: items.plowhorses.slice(0, 3).map((it) => ({
        label: it.name,
        value: `${formatNumber(it.unitsSold)} units · ${formatCurrency(it.revenue)}`,
      })),
    },
    {
      id: 'prioritize-items',
      question: 'Which items should the restaurant prioritize?',
      badge: 'Star Items',
      color: 'emerald',
      headline: `${items.stars.length} "Star" menu items drive both high order volume and high rupee revenue`,
      analysis: items.stars.length > 0
        ? `Prioritize kitchen prep, spice/gravy consistency, and prominent menu placement for ${items.stars.slice(0, 4).map((i) => i.name).join(', ')}. These flagship dishes outperform both the volume benchmark (${formatNumber(kpis.avgItemUnits, 0)} units) and revenue benchmark (${formatCurrency(kpis.avgItemRev)}).`
        : `Focus on your top revenue contributors (${items.byRevenue.slice(0, 3).map((i) => i.name).join(', ')}) to protect core cash flow.`,
      metrics: items.stars.slice(0, 3).map((it) => ({
        label: it.name,
        value: `${formatCurrency(it.revenue)} · ${it.unitsSold} sold`,
      })),
    },
    {
      id: 'promote-or-reconsider',
      question: 'Which items should be promoted or reconsidered?',
      badge: 'Menu Engineering',
      color: 'purple',
      headline: `${items.puzzles.length} high-margin Puzzle items to promote; ${items.dogs.length} low-performing Dog items to reconsider`,
      analysis: `${topPuzzle ? `Promote "${topPuzzle.name}" (${formatCurrency(topPuzzle.avgPrice)} unit price, ${formatCurrency(topPuzzle.revenue)} revenue from ${topPuzzle.unitsSold} units)—suggestive upselling as a table starter or family add-on will significantly lift average bill value.` : ''} ${bottomDog ? `Conversely, reconsider "${bottomDog.name}" (${formatNumber(bottomDog.unitsSold)} units, ${formatCurrency(bottomDog.revenue)} revenue, ${formatPercent(bottomDog.revenueShare)} share), which ranks at the bottom of the menu.` : ''}`,
      metrics: [
        ...(topPuzzle ? [{ label: `Promote: ${topPuzzle.name}`, value: `${formatCurrency(topPuzzle.avgPrice)} / unit` }] : []),
        ...(bottomDog ? [{ label: `Reconsider: ${bottomDog.name}`, value: `${formatCurrency(bottomDog.revenue)} total` }] : []),
      ],
    },
    {
      id: 'customer-patterns',
      question: 'What purchasing patterns can be observed among customers?',
      badge: 'Customer Behavior',
      color: 'sky',
      headline: `Top 10% of customers contribute ${formatPercent(kpis.top10PctCustShare)} of total restaurant revenue`,
      analysis: `${topCustomer ? `Your highest-spending patron is ${topCustomer.id} (${formatCurrency(topCustomer.totalRevenue)} across ${topCustomer.totalUnits} items, favoring ${topCustomer.top1Item}). ` : ''}${vipSegment ? `The "VIP Champions" segment (${vipSegment.count} diners) generates ${formatCurrency(vipSegment.revenue)} (${formatPercent(vipSegment.shareOfRevenue)} of revenue). ` : ''}${customers.topCoOrders.length > 0 ? `Strongest Indian meal pairing affinity is "${customers.topCoOrders[0].pair}" (${customers.topCoOrders[0].count} co-occurrences).` : ''}${hygiene.unattributedRevenueTotal > 0 ? ` Note: ${formatCurrency(hygiene.unattributedRevenueTotal)} of total customer spend comes from additional items beyond their top 3 favorites.` : ''}`,
      metrics: [
        { label: 'Top 10% Customer Share', value: formatPercent(kpis.top10PctCustShare) },
        { label: 'Top Customer Spend', value: topCustomer ? `${topCustomer.id} (${formatCurrency(topCustomer.totalRevenue)})` : '—' },
        { label: 'Top Meal Pairing', value: customers.topCoOrders[0]?.pair || 'N/A' },
      ],
    },
    {
      id: 'practical-steps',
      question: 'What practical steps can the restaurant take to improve revenue?',
      badge: 'Action Plan',
      color: 'emerald',
      headline: '4 data-backed operational levers to increase restaurant revenue over the next 30–90 days',
      analysis: `1) Reprice or Bundle High-Volume Plowhorses: ${topPlowhorse ? `Apply a modest ${formatCurrency(15)} price adjustment on ${topPlowhorse.name} (${topPlowhorse.unitsSold} units) or bundle it into a Meal Combo.` : 'Bundle high-frequency breads and beverages with signature curries.'} 2) Upsell Tandoor & Specialty Puzzles: ${topPuzzle ? `Target +25% volume on ${topPuzzle.name} (${formatCurrency(topPuzzle.avgPrice)}) via table recommendations.` : 'Incentivize staff to recommend high-value starters.'} 3) Curated Meal Pairings: ${customers.topCoOrders[0] ? `Feature "${customers.topCoOrders[0].pair}" as a fixed-price Family/Meal Combo to increase average bill size.` : 'Introduce a curated 2-course meal combo.'} 4) Loyalty & Win-Back Campaign: ${customers.decliningOrAtRisk.length > 0 ? `Send a complimentary dessert or starter voucher to the ${customers.decliningOrAtRisk.length} inactive repeat diners who haven't ordered in 60+ days.` : `Reward your ${kpis.repeatCustomersCount} repeat patrons with a loyalty benefit.`}`,
      metrics: [
        { label: 'Est. Plowhorse Reprice Lift', value: topPlowhorse ? `+${formatCurrency(topPlowhorse.unitsSold * 15)}` : '+4–6% Rev' },
        { label: 'Puzzle +25% Volume Lift', value: topPuzzle ? `+${formatCurrency(topPuzzle.revenue * 0.25)}` : '+5–8% Rev' },
        { label: 'At-Risk Diners to Re-engage', value: hygiene.hasDateData ? `${customers.decliningOrAtRisk.length} diners` : 'Requires Date Logs' },
      ],
    },
  ];
};

// Deterministic Local Answer Engine (Guarantees exact INR figures from active dataset)
const buildDeterministicAnswer = (query, analytics, gstBreakdown) => {
  if (!analytics) return 'Please upload a dataset or load a sample Indian restaurant dataset first so I can analyze your performance.';
  const q = query.toLowerCase();
  const { kpis, items, customers, timeSeries, hygiene } = analytics;

  if (q.includes('gst') || q.includes('tax')) {
    if (!gstBreakdown || !gstBreakdown.isActive) {
      return `**GST Configuration Notice:** Currently, GST treatment is set to **"GST Not Specified (Raw Uploaded Values)"**, so your reported total revenue of **${formatCurrency(kpis.totalRevenue)}** reflects the exact raw amounts in your dataset without assuming whether prices are inclusive or exclusive of GST.\n\nYou can select **5% GST (Restaurant Standard)** or **18% GST** (Inclusive or Exclusive) using the **GST selector** in the top header or Upload page to automatically calculate Net Revenue (Ex-GST) and Estimated GST Liability.`;
    }
    return `Based on your selected tax setting (**${gstBreakdown.label}**):\n\n• **Raw Dataset Revenue:** ${formatCurrency(gstBreakdown.rawRevenue)}\n• **Net Revenue (Ex-GST):** **${formatCurrency(gstBreakdown.netRevenueExGst)}**\n• **Estimated GST Component (${gstBreakdown.rate}%):** **${formatCurrency(gstBreakdown.gstAmount)}**\n• **Gross Revenue (Incl. GST):** **${formatCurrency(gstBreakdown.grossRevenueInclGst)}**`;
  }

  if (q.includes('highest revenue') || (q.includes('which item') && q.includes('revenue'))) {
    const top = items.byRevenue[0];
    const second = items.byRevenue[1];
    const third = items.byRevenue[2];
    if (!top) return 'No menu items were found in the active dataset.';
    return `**${top.name}** generated the highest revenue at **${formatCurrency(top.revenue)}**, accounting for **${formatPercent(top.revenueShare)}** of total item revenue across **${formatNumber(top.unitsSold)} units sold** (average unit price: ${formatCurrency(top.avgPrice)}).\n\n**Top 3 Revenue Generators:**\n1. **${top.name}** — ${formatCurrency(top.revenue)} (${formatNumber(top.unitsSold)} units · ${formatPercent(top.revenueShare)})\n${second ? `2. **${second.name}** — ${formatCurrency(second.revenue)} (${formatNumber(second.unitsSold)} units · ${formatPercent(second.revenueShare)})\n` : ''}${third ? `3. **${third.name}** — ${formatCurrency(third.revenue)} (${formatNumber(third.unitsSold)} units · ${formatPercent(third.revenueShare)})` : ''}`;
  }

  if (q.includes('top five') || q.includes('top 5') || q.includes('best-selling') || q.includes('best selling')) {
    const top5Units = items.byUnits.slice(0, 5);
    const top5Rev = items.byRevenue.slice(0, 5);
    return `Here are your **Top 5 Best-Selling Menu Items by Volume (Units Sold)** alongside their rupee revenue contribution:\n\n${top5Units.map((it, idx) => `${idx + 1}. **${it.name}** — **${formatNumber(it.unitsSold)} units sold** (${formatPercent(it.unitShare)} of volume) · Revenue: **${formatCurrency(it.revenue)}** (${formatCurrency(it.avgPrice)}/unit)`).join('\n')}\n\n*By comparison, the top 5 items by total revenue are: ${top5Rev.map((i) => `${i.name} (${formatCurrency(i.revenue)})`).join(', ')}.*`;
  }

  if (q.includes('customer') && (q.includes('spend') || q.includes('most') || q.includes('valuable') || q.includes('top'))) {
    const top5Cust = customers.bySpend.slice(0, 5);
    return `Your top-spending customer is **${top5Cust[0]?.id}** (${top5Cust[0]?.email !== '—' ? top5Cust[0]?.email : 'No email on file'}) with **${formatCurrency(top5Cust[0]?.totalRevenue)}** in total spending across **${formatNumber(top5Cust[0]?.totalUnits)} items**.\n\n**Top 5 Highest-Spending Customers:**\n${top5Cust.map((c, i) => `${i + 1}. **${c.id}** — **${formatCurrency(c.totalRevenue)}** (${c.totalUnits} items · Top choice: *${c.top1Item}* [${c.top1Freq}x])`).join('\n')}\n\nOverall, your **top 10% of customers** account for **${formatPercent(kpis.top10PctCustShare)}** of total restaurant revenue, and your average revenue per customer is **${formatCurrency(kpis.avgRevenuePerCustomer)}**.`;
  }

  if (q.includes('percentage') && q.includes('top three') || q.includes('top 3')) {
    const top3 = items.byRevenue.slice(0, 3);
    const top3RevSum = top3.reduce((s, i) => s + i.revenue, 0);
    return `Your **top three menu items generate ${formatPercent(kpis.top3ItemsRevPct)}** (**${formatCurrency(top3RevSum)}**) of your total attributed menu item revenue (${formatCurrency(kpis.totalItemRevenue)}).\n\n**Breakdown of the Top 3 Items:**\n${top3.map((it, idx) => `${idx + 1}. **${it.name}**: ${formatCurrency(it.revenue)} (**${formatPercent(it.revenueShare)}** share · ${it.unitsSold} units)`).join('\n')}`;
  }

  if (q.includes('focus on next month') || q.includes('prioritize') || q.includes('promote')) {
    const stars = items.stars.slice(0, 3);
    const puzzles = items.puzzles.slice(0, 3);
    const plowhorses = items.plowhorses.slice(0, 3);
    return `Based on Menu Engineering Matrix analysis of your uploaded Indian restaurant dataset, here is what you should focus on next month:\n\n1. **Protect & Feature Your "Stars" (High Volume + High Revenue):**\n   ${stars.map((s) => `• **${s.name}** (${formatCurrency(s.revenue)} revenue, ${s.unitsSold} units)`).join('\n   ') || '• Focus on your top revenue leaders.'}\n\n2. **Actively Promote "Puzzles" (High Revenue Yield, Lower Volume):**\n   ${puzzles.length > 0 ? puzzles.map((p) => `• **${p.name}** (${formatCurrency(p.avgPrice)} price point, ${formatCurrency(p.revenue)} revenue from ${p.unitsSold} units — prime candidate for server upselling and table tent highlights)`).join('\n   ') : '• No under-ordered high-ticket puzzles detected.'}\n\n3. **Optimize Pricing or Combo Bundling on "Plowhorses" (High Volume, Lower Revenue):**\n   ${plowhorses.length > 0 ? plowhorses.map((pl) => `• **${pl.name}** (${pl.unitsSold} units sold at ${formatCurrency(pl.avgPrice)} — test a modest ${formatCurrency(15)} price lift or pair into a Curry/Biryani/Tiffin combo)`).join('\n   ') : '• Popular items are already well-priced.'}`;
  }

  if (q.includes('six months') || q.includes('change') || q.includes('trend') || q.includes('monthly')) {
    if (!timeSeries.hasDateData) {
      return `**Date-Level Data Notice:** Your currently loaded dataset uses the **Customer-Summary Schema**, which provides cumulative customer and menu item totals (${formatCurrency(kpis.totalRevenue)} across ${formatNumber(kpis.totalCustomers)} customers) but **does not contain individual transaction dates**.\n\nTo view month-over-month, weekly, or daily sales changes (` + '`DD-MM-YYYY`' + `) over six months, please upload a **Transaction-Level Dataset** containing a \`Transaction Date\` column (or click *"6M Transaction Demo"* in the top bar).`;
    }
    const months = timeSeries.monthly;
    const comp = timeSeries.periodComparison;
    return `Over the recorded 6-month timeline (**${hygiene.dateRange?.start}** to **${hygiene.dateRange?.end}**), monthly revenue progressed as follows:\n\n${months.map((m) => `• **${m.label}**: **${formatCurrency(m.revenue)}** (${formatNumber(m.units)} items sold across ${m.orders} orders${m.momGrowth !== null ? ` · **${m.momGrowth >= 0 ? '+' : ''}${m.momGrowth}%** MoM` : ''})`).join('\n')}\n\n${comp ? `**Period Comparison:** Comparing the second half (${comp.secondHalfLabel}: ${formatCurrency(comp.secondHalfRevenue)}) against the first half (${comp.firstHalfLabel}: ${formatCurrency(comp.firstHalfRevenue)}), sales changed by **${comp.halfOverHalfPct >= 0 ? '+' : ''}${comp.halfOverHalfPct}%**.` : ''}`;
  }

  if (q.includes('low sales') || q.includes('lowest') || q.includes('underperform') || q.includes('remove')) {
    const bottom5 = items.lowestPerforming.slice(0, 5);
    return `Here are the **5 Lowest-Performing Menu Items** ranked by total revenue contribution:\n\n${bottom5.map((it, idx) => `${idx + 1}. **${it.name}** — **${formatCurrency(it.revenue)}** (${formatNumber(it.unitsSold)} units sold · ${formatPercent(it.revenueShare)} of item revenue · Quadrant: *${it.quadrant}*)`).join('\n')}\n\n**Recommendation:** Items classified as **Dogs** (${items.dogs.map((d) => d.name).join(', ') || 'bottom performers'}) fall below both your menu average volume (${formatNumber(kpis.avgItemUnits, 1)} units) and average item revenue (${formatCurrency(kpis.avgItemRev)}). Consider simplifying prep or running them as weekend specials.`;
  }

  if (q.includes('increase') || q.includes('improve') || q.includes('recommendations') || q.includes('what can i do')) {
    const topPlow = items.plowhorses[0];
    const topPuz = items.puzzles[0];
    const topPair = customers.topCoOrders[0];
    return `Based on your dataset of **${formatCurrency(kpis.totalRevenue)}** across **${formatNumber(kpis.totalCustomers)} customers**, here are 4 concrete ways to increase restaurant revenue:\n\n1. **Reprice High-Volume Indian Staples (Plowhorses):** ${topPlow ? `**${topPlow.name}** sold **${topPlow.unitsSold} units** at ${formatCurrency(topPlow.avgPrice)}. Increasing its price by just ${formatCurrency(15)} would generate **+${formatCurrency(topPlow.unitsSold * 15)}** in incremental revenue assuming steady demand.` : 'Review high-frequency bread, tiffin, and beverage pricing.'}\n2. **Promote High-Margin Specialties (Puzzles):** ${topPuz ? `**${topPuz.name}** commands **${formatCurrency(topPuz.avgPrice)}/unit** and generated ${formatCurrency(topPuz.revenue)} from ${topPuz.unitsSold} units. Lifting unit volume by 25% via server upsells adds **+${formatCurrency(topPuz.revenue * 0.25)}**.` : 'Highlight high-ticket tandoor starters and mains.'}\n3. **Bundle Co-Ordered Meal Pairings:** ${topPair ? `Diners frequently order **${topPair.pair}** together (${topPair.count} patrons). Offer this as a curated Combo / Thali upgrade to boost average check size.` : 'Create a chef pairing combo.'}\n4. **Increase Repeat Dining Frequency:** Your repeat customer rate is **${formatPercent(kpis.repeatCustomerPct)}** (${kpis.repeatCustomersCount} repeat diners vs. ${kpis.totalCustomers - kpis.repeatCustomersCount} single-visit diners). Converting just 15% of single-visit guests into repeat patrons at your average spend (${formatCurrency(kpis.avgRevenuePerCustomer)}) would unlock **+${formatCurrency((kpis.totalCustomers - kpis.repeatCustomersCount) * 0.15 * kpis.avgRevenuePerCustomer)}**.`;
  }

  return `Here is a real-time summary from your active dataset (**${analytics.meta.datasetName}**):\n\n• **Total Revenue:** ${formatCurrency(kpis.totalRevenue)} (${formatNumber(kpis.totalUnitsSold)} total items sold)\n• **Customer Base:** ${formatNumber(kpis.totalCustomers)} customers (${formatPercent(kpis.repeatCustomerPct)} repeat rate · ${formatCurrency(kpis.avgRevenuePerCustomer)} avg revenue/customer)\n• **Top Revenue Item:** ${kpis.highestRevenueItem?.name || '—'} (${formatCurrency(kpis.highestRevenueItem?.revenue)} · ${formatPercent(kpis.highestRevenueItem?.revenueShare)} share)\n• **Most Popular Item by Volume:** ${kpis.mostPopularItemByUnits?.name || '—'} (${formatNumber(kpis.mostPopularItemByUnits?.unitsSold)} units)\n• **Top 3 Items Concentration:** ${formatPercent(kpis.top3ItemsRevPct)} of item revenue\n\nFeel free to ask about specific Indian menu items, customer segments, GST breakdowns, pricing opportunities, or 6-month period comparisons!`;
};

// ============================================================================
// SECURE SAME-ORIGIN GEMINI API HELPER (API KEY STAYS ON THE SERVER)
// ============================================================================

const callGeminiAnalystAPI = async (userPrompt, analytics, chatHistory = [], gstBreakdown = null) => {
  if (!analytics) return 'Please upload a dataset or load a sample Indian restaurant dataset first.';

  const dataContext = {
    datasetName: analytics.meta.datasetName,
    schemaType: analytics.meta.schemaType,
    regionalSettings: {
      locale: activeCurrency.locale,
      currencyCode: activeCurrency.code,
      currencySymbol: activeCurrency.symbol,
      timezone: 'Asia/Kolkata (IST)',
      dateFormat: 'DD-MM-YYYY',
    },
    gstConfiguration: gstBreakdown,
    hasDateData: analytics.timeSeries.hasDateData,
    dateRange: analytics.hygiene.dateRange,
    kpis: analytics.kpis,
    topItemsByRevenue: analytics.items.byRevenue.slice(0, 10).map((i) => ({
      name: i.name,
      unitsSold: i.unitsSold,
      revenue: i.revenue,
      avgPrice: i.avgPrice,
      revenueSharePct: Number(i.revenueShare.toFixed(1)),
      quadrant: i.quadrant,
    })),
    lowestItemsByRevenue: analytics.items.lowestPerforming.slice(0, 5).map((i) => ({
      name: i.name,
      unitsSold: i.unitsSold,
      revenue: i.revenue,
      avgPrice: i.avgPrice,
      quadrant: i.quadrant,
    })),
    topCustomers: analytics.customers.bySpend.slice(0, 8).map((c) => ({
      id: c.id,
      totalRevenue: c.totalRevenue,
      totalUnits: c.totalUnits,
      top1Item: c.top1Item,
      top2Item: c.top2Item,
      segment: c.segment,
    })),
    customerSegments: analytics.customers.segments,
    topCoOrders: analytics.customers.topCoOrders.slice(0, 5),
    monthlyTrends: analytics.timeSeries.monthly,
    periodComparison: analytics.timeSeries.periodComparison,
  };

  try {
    const response = await fetch('/api/analyst', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        userPrompt,
        dataContext,
        chatHistory: chatHistory.slice(-8).map(({ role, content }) => ({ role, content })),
        currencyCode: activeCurrency.code,
        currencySymbol: activeCurrency.symbol,
      }),
    });

    const result = await response.json().catch(() => ({}));
    if (!response.ok) {
      const reason = result.error || 'The Gemini service is unavailable.';
      if (response.status === 503) {
        return `${buildDeterministicAnswer(userPrompt, analytics, gstBreakdown)}\n\n---\n**Gemini connection:** ${reason} You can still use the local, data-based answers; add your API key to the server-side `.env` file to enable Gemini responses.`;
      }
      throw new Error(reason);
    }
    if (typeof result.answer === 'string' && result.answer.trim()) return result.answer;
    throw new Error('Gemini returned an empty response.');
  } catch (err) {
    const localAnswer = buildDeterministicAnswer(userPrompt, analytics, gstBreakdown);
    return `${localAnswer}\n\n---\n*Gemini could not be reached (${String(err.message || 'network error')}). Showing a local answer calculated from the active dataset.*`;
  }
};

// ============================================================================
// CUSTOM RECHARTS TOOLTIP & THEME CONSTANTS
// ============================================================================

const CHART_COLORS = ['#10B981', '#6366F1', '#F59E0B', '#EC4899', '#06B6D4', '#8B5CF6', '#14B8A6', '#F43F5E'];

const CustomChartTooltip = ({ active, payload, label, valueFormatter = formatCurrency }) => {
  if (!active || !payload || !payload.length) return null;
  return (
    <div className="bg-[#11131A]/95 backdrop-blur-md border border-white/10 rounded-lg p-3 shadow-2xl text-xs">
      {label && <div className="font-semibold text-gray-200 mb-1.5 pb-1 border-b border-white/10">{label}</div>}
      <div className="space-y-1">
        {payload.map((entry, idx) => {
          const isPct = String(entry.name).toLowerCase().includes('%') || String(entry.name).toLowerCase().includes('share');
          const isUnits = String(entry.name).toLowerCase().includes('unit') || String(entry.name).toLowerCase().includes('qty') || String(entry.name).toLowerCase().includes('order') || String(entry.name).toLowerCase().includes('freq');
          const formatted = isPct
            ? formatPercent(entry.value)
            : isUnits
            ? formatNumber(entry.value)
            : valueFormatter(entry.value);
          return (
            <div key={idx} className="flex items-center justify-between gap-4">
              <span className="flex items-center gap-1.5 text-gray-400">
                <span className="w-2 h-2 rounded-full" style={{ backgroundColor: entry.color || '#10B981' }} />
                {entry.name}:
              </span>
              <span className="font-mono font-semibold text-gray-100 tabular-nums">{formatted}</span>
            </div>
          );
        })}
      </div>
    </div>
  );
};

// Simple Markdown-like renderer for Chatbot & Insights
const FormattedText = ({ text }) => {
  if (!text) return null;
  const lines = String(text).split('\n');
  return (
    <div className="space-y-1.5 leading-relaxed">
      {lines.map((line, i) => {
        if (!line.trim()) return <div key={i} className="h-1" />;
        const parts = line.split(/(\*\*.*?\*\*|\*.*?\*|`.*?`)/g);
        return (
          <p key={i} className="text-sm text-gray-300">
            {parts.map((part, j) => {
              if (part.startsWith('**') && part.endsWith('**')) {
                return <strong key={j} className="font-semibold text-white">{part.slice(2, -2)}</strong>;
              }
              if (part.startsWith('*') && part.endsWith('*')) {
                return <em key={j} className="italic text-emerald-300">{part.slice(1, -1)}</em>;
              }
              if (part.startsWith('`') && part.endsWith('`')) {
                return <code key={j} className="px-1.5 py-0.5 bg-white/10 rounded text-xs font-mono text-emerald-300">{part.slice(1, -1)}</code>;
              }
              return part;
            })}
          </p>
        );
      })}
    </div>
  );
};

// ============================================================================
// MAIN APPLICATION COMPONENT
// ============================================================================

export default function App() {
  // Navigation & UI State
  const [activeTab, setActiveTab] = useState('overview');
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);

  // Regional Currency & GST Configuration State (Defaults to INR ₹ and Unspecified GST)
  const [currencyCode, setCurrencyCode] = useState('INR');
  const [gstMode, setGstMode] = useState('none');

  // Keep module-level formatter synchronized with selected currency
  setGlobalCurrency(currencyCode);

  // Raw File & Column Mapping State
  const [rawHeaders, setRawHeaders] = useState([]);
  const [rawRows, setRawRows] = useState([]);
  const [pendingSchemaType, setPendingSchemaType] = useState('transaction');
  const [pendingMapping, setPendingMapping] = useState({});
  const [pendingMeta, setPendingMeta] = useState({ datasetName: '', isDemo: true });
  const [showMappingModal, setShowMappingModal] = useState(false);
  const [uploadError, setUploadError] = useState(null);
  const [isParsing, setIsParsing] = useState(false);

  // Active Processed Analytics State
  const [analytics, setAnalytics] = useState(null);

  // Section-Specific Filter States
  const [timeGranularity, setTimeGranularity] = useState('monthly'); // daily | weekly | monthly
  const [menuFilterTab, setMenuFilterTab] = useState('all'); // all | stars | plowhorses | puzzles | dogs | frequency
  const [menuSearch, setMenuSearch] = useState('');
  const [menuSortField, setMenuSortField] = useState('revenue');
  const [customerSearch, setCustomerSearch] = useState('');
  const [customerSegmentFilter, setCustomerSegmentFilter] = useState('ALL');
  const [customerSortField, setCustomerSortField] = useState('totalRevenue');
  const [previewPage, setPreviewPage] = useState(1);

  // AI Analyst & Chatbot State
  const [aiEnhancedSummary, setAiEnhancedSummary] = useState(null);
  const [isGeneratingAiBrief, setIsGeneratingAiBrief] = useState(false);
  const [chatMessages, setChatMessages] = useState([
    {
      role: 'assistant',
      content: 'Namaste! I am your **Restaurant AI Analyst**. I have analyzed your Indian restaurant sales dataset in **INR (₹)**. Ask me anything about your menu item revenue, customer spending patterns, Pareto concentration, GST breakdown, or how to increase profitability.',
      timestamp: 'Just now',
    },
  ]);
  const [chatInput, setChatInput] = useState('');
  const [isChatLoading, setIsChatLoading] = useState(false);
  const chatEndRef = useRef(null);

  // SheetJS is bundled locally via the `xlsx` npm dependency (no runtime CDN required).

  // Load initial 6-Month Indian Restaurant Transaction Sample Dataset on first launch
  useEffect(() => {
    loadDemoDataset('transaction', false);
  }, []);

  useEffect(() => {
    chatEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [chatMessages, isChatLoading]);

  // Compute GST Breakdown whenever analytics or gstMode updates
  const gstBreakdown = useMemo(() => {
    if (!analytics) return null;
    return computeGSTBreakdown(analytics.kpis.totalRevenue, gstMode);
  }, [analytics, gstMode, currencyCode]);

  // Load either the 6-Month Indian Transaction Demo or the Customer-Summary Demo
  const loadDemoDataset = (type = 'transaction', navigateToOverview = true) => {
    setUploadError(null);
    setCurrencyCode('INR');
    setGlobalCurrency('INR');
    const sampleRows = type === 'customer_summary'
      ? generateSampleCustomerSummaryDataset()
      : generateSampleTransactionDataset();
    const headers = Object.keys(sampleRows[0] || {});
    const { schemaType, mapping } = autoDetectSchemaAndMapping(headers);
    const meta = {
      datasetName: type === 'customer_summary'
        ? 'Indian Restaurant Customer-Summary Demo (115 Diners · INR)'
        : 'Indian Restaurant 6-Month Transactions Demo (01-05-2026 – 31-10-2026)',
      isDemo: true,
    };

    setRawHeaders(headers);
    setRawRows(sampleRows);
    setPendingSchemaType(schemaType);
    setPendingMapping(mapping);
    setPendingMeta(meta);

    const processed = processRestaurantData(sampleRows, schemaType, mapping, meta);
    setAnalytics(processed);
    setAiEnhancedSummary(null);
    setChatMessages([
      {
        role: 'assistant',
        content: `Loaded **${meta.datasetName}** (${schemaType === 'customer_summary' ? 'Customer-Summary Schema' : 'Transaction-Level Schema'}). Total revenue is **${formatCurrency(processed.kpis.totalRevenue)}** across **${formatNumber(processed.kpis.totalCustomers)} customers** and **${formatNumber(processed.kpis.totalUnitsSold)} items sold**. What would you like to explore?`,
        timestamp: `${getISTTimeString()} IST`,
      },
    ]);
    if (navigateToOverview) setActiveTab('overview');
  };

  // Handle CSV / XLSX / XLS file upload
  const handleFileUpload = (file) => {
    if (!file) return;
    setUploadError(null);
    setIsParsing(true);

    const fileName = file.name;
    const ext = fileName.split('.').pop().toLowerCase();

    if (!['csv', 'xlsx', 'xls', 'txt'].includes(ext)) {
      setUploadError('Unsupported file format. Please upload a .CSV, .XLSX, or .XLS file.');
      setIsParsing(false);
      return;
    }

    const reader = new FileReader();

    if (ext === 'csv' || ext === 'txt') {
      reader.onload = (e) => {
        try {
          const text = e.target.result;
          const { headers, rows } = parseCSVText(text);
          if (headers.length === 0 || rows.length === 0) {
            setUploadError('The uploaded CSV file appears to be empty or improperly formatted.');
            setIsParsing(false);
            return;
          }
          const { schemaType, mapping } = autoDetectSchemaAndMapping(headers);
          setRawHeaders(headers);
          setRawRows(rows);
          setPendingSchemaType(schemaType);
          setPendingMapping(mapping);
          setPendingMeta({ datasetName: fileName, isDemo: false });
          setShowMappingModal(true);
          setIsParsing(false);
        } catch (err) {
          setUploadError(`Failed to parse CSV file: ${err.message}`);
          setIsParsing(false);
        }
      };
      reader.readAsText(file);
    } else {
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array', cellDates: true });
          const firstSheet = workbook.SheetNames[0];
          if (!firstSheet) throw new Error('No worksheet was found in the workbook.');
          const worksheet = workbook.Sheets[firstSheet];
          const jsonRows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

          if (!jsonRows || jsonRows.length === 0) {
            setUploadError('The uploaded Excel workbook contains no readable rows.');
            setIsParsing(false);
            return;
          }
          // Format any Date objects into DD-MM-YYYY strings
          const normalizedRows = jsonRows.map((r) => {
            const copy = {};
            Object.entries(r).forEach(([k, v]) => {
              if (v instanceof Date && !isNaN(v.getTime())) {
                copy[k] = formatDateDDMMYYYY(v);
              } else {
                copy[k] = v;
              }
            });
            return copy;
          });
          const headers = Object.keys(normalizedRows[0]);
          const { schemaType, mapping } = autoDetectSchemaAndMapping(headers);
          setRawHeaders(headers);
          setRawRows(normalizedRows);
          setPendingSchemaType(schemaType);
          setPendingMapping(mapping);
          setPendingMeta({ datasetName: fileName, isDemo: false });
          setShowMappingModal(true);
          setIsParsing(false);
        } catch (err) {
          setUploadError(`Failed to parse Excel file: ${err.message}`);
          setIsParsing(false);
        }
      };
      reader.readAsArrayBuffer(file);
    }
  };

  // Apply Column Mapping & Run Analytics
  const applyMappingAndAnalyze = () => {
    const processed = processRestaurantData(rawRows, pendingSchemaType, pendingMapping, pendingMeta);
    setAnalytics(processed);
    setShowMappingModal(false);
    setAiEnhancedSummary(null);
    setChatMessages([
      {
        role: 'assistant',
        content: `Successfully processed **${pendingMeta.datasetName}** (${formatNumber(processed.hygiene.validRowCount)} valid records, displayed in **${activeCurrency.code} ${activeCurrency.symbol}** without altering raw numeric values). Total revenue is **${formatCurrency(processed.kpis.totalRevenue)}** across **${formatNumber(processed.kpis.totalCustomers)} customers**. Ask me any question about your data!`,
        timestamp: `${getISTTimeString()} IST`,
      },
    ]);
    setActiveTab('overview');
  };

  // Download Sample CSV Template
  const downloadSampleCSV = (type = 'customer_summary') => {
    const rows = type === 'customer_summary'
      ? generateSampleCustomerSummaryDataset()
      : generateSampleTransactionDataset();
    if (!rows.length) return;
    const headers = Object.keys(rows[0]);
    const csvLines = [
      headers.join(','),
      ...rows.map((r) =>
        headers.map((h) => {
          const val = r[h] ?? '';
          return String(val).includes(',') ? `"${val}"` : val;
        }).join(',')
      ),
    ];
    const blob = new Blob([csvLines.join('\n')], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = type === 'customer_summary'
      ? 'indian_restaurant_customer_summary_sample.csv'
      : 'indian_restaurant_6month_transactions_sample.csv';
    a.click();
    URL.revokeObjectURL(url);
  };

  // Download Executive Summary Report (.md / text report)
  const downloadSummaryReport = () => {
    if (!analytics) return;
    const { meta, kpis, items, customers, timeSeries, hygiene } = analytics;
    const biReport = generateStructuredBIReport(analytics);

    const lines = [
      `# INDIAN RESTAURANT BUSINESS INTELLIGENCE & SALES ANALYTICS REPORT`,
      `Generated: ${formatDateDDMMYYYY(new Date())} at ${getISTTimeString()} IST (Asia/Kolkata)`,
      `Locale & Currency: ${activeCurrency.locale} · ${activeCurrency.label}`,
      `GST Treatment: ${gstBreakdown?.label || 'Not Specified'}`,
      `Dataset: ${meta.datasetName} (${meta.isDemo ? 'SAMPLE DEMO DATA' : 'LIVE UPLOADED DATA'})`,
      `Schema Type: ${meta.schemaType === 'customer_summary' ? 'Customer-Summary Schema' : 'Transaction-Level Schema'}`,
      `Valid Records Processed: ${hygiene.validRowCount} (Duplicates Removed: ${hygiene.duplicatesRemoved})`,
      hygiene.hasDateData ? `Date Coverage (DD-MM-YYYY): ${hygiene.dateRange?.start} to ${hygiene.dateRange?.end}` : `Date Coverage: Not applicable (Summary Schema)`,
      ``,
      `## 1. EXECUTIVE PERFORMANCE SUMMARY`,
      `- Total Revenue (${activeCurrency.code}): ${formatCurrency(kpis.totalRevenue)}`,
      ...(gstBreakdown?.isActive
        ? [
            `- Net Revenue (Ex-GST): ${formatCurrency(gstBreakdown.netRevenueExGst)}`,
            `- Estimated GST (${gstBreakdown.rate}%): ${formatCurrency(gstBreakdown.gstAmount)}`,
            `- Gross Revenue (Incl. GST): ${formatCurrency(gstBreakdown.grossRevenueInclGst)}`,
          ]
        : []),
      `- Total Items Sold: ${formatNumber(kpis.totalUnitsSold)}`,
      `- Total Customers: ${formatNumber(kpis.totalCustomers)}`,
      `- Average Revenue per Customer: ${formatCurrency(kpis.avgRevenuePerCustomer)}`,
      `- Average Item Price: ${formatCurrency(kpis.avgItemPrice)}`,
      `- Repeat Customer Percentage: ${formatPercent(kpis.repeatCustomerPct)} (${kpis.repeatCustomersCount} repeat customers)`,
      `- Most Popular Item (Volume): ${kpis.mostPopularItemByUnits?.name || 'N/A'} (${formatNumber(kpis.mostPopularItemByUnits?.unitsSold)} units)`,
      `- Highest Revenue Item: ${kpis.highestRevenueItem?.name || 'N/A'} (${formatCurrency(kpis.highestRevenueItem?.revenue)})`,
      `- Top 3 Items Share of Item Revenue: ${formatPercent(kpis.top3ItemsRevPct)}`,
      ``,
      `## 2. MENU ITEM PERFORMANCE & ENGINEERING MATRIX`,
      `Rank | Menu Item | Quadrant | Units Sold | Avg Price | Total Revenue | Revenue Share`,
      `--- | --- | --- | --- | --- | --- | ---`,
      ...items.byRevenue.map((it) =>
        `${it.revenueRank} | ${it.name} | ${it.quadrant} | ${it.unitsSold} | ${formatCurrency(it.avgPrice)} | ${formatCurrency(it.revenue)} | ${formatPercent(it.revenueShare)}`
      ),
      ``,
      `## 3. TOP 15 VALUABLE CUSTOMERS`,
      `Rank | Customer ID | Email | Segment | Total Spend | Units Ordered | #1 Preferred Item | #2 Preferred Item`,
      `--- | --- | --- | --- | --- | --- | --- | ---`,
      ...customers.bySpend.slice(0, 15).map((c) =>
        `${c.spendRank} | ${c.id} | ${c.email} | ${c.segment} | ${formatCurrency(c.totalRevenue)} | ${c.totalUnits} | ${c.top1Item} (${c.top1Freq}x) | ${c.top2Item}`
      ),
      ``,
      `## 4. AI BUSINESS INTELLIGENCE & STRATEGIC RECOMMENDATIONS`,
      ...biReport.flatMap((sec) => [
        `### ${sec.question}`,
        `**Headline:** ${sec.headline}`,
        `${sec.analysis}`,
        ``,
      ]),
    ];

    if (timeSeries.hasDateData && timeSeries.monthly.length > 0) {
      lines.push(
        `## 5. MONTHLY REVENUE BREAKDOWN`,
        `Month | Revenue | Units Sold | Orders | MoM Growth`,
        `--- | --- | --- | --- | ---`,
        ...timeSeries.monthly.map((m) =>
          `${m.label} | ${formatCurrency(m.revenue)} | ${m.units} | ${m.orders} | ${m.momGrowth !== null ? `${m.momGrowth}%` : '—'}`
        )
      );
    }

    const blob = new Blob([lines.join('\n')], { type: 'text/markdown;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Indian_Restaurant_BI_Report_${formatDateDDMMYYYY(new Date())}.md`;
    a.click();
    URL.revokeObjectURL(url);
  };

  // Send message to Restaurant AI Analyst Chatbot
  const handleSendChat = async (promptText) => {
    const query = (promptText ?? chatInput).trim();
    if (!query || isChatLoading) return;

    const userMsg = {
      role: 'user',
      content: query,
      timestamp: `${getISTTimeString()} IST`,
    };

    const updatedHistory = [...chatMessages, userMsg];
    setChatMessages(updatedHistory);
    if (!promptText) setChatInput('');
    setIsChatLoading(true);

    const answer = await callGeminiAnalystAPI(query, analytics, updatedHistory, gstBreakdown);
    setChatMessages((prev) => [
      ...prev,
      {
        role: 'assistant',
        content: answer,
        timestamp: `${getISTTimeString()} IST`,
      },
    ]);
    setIsChatLoading(false);
  };

  // Generate AI Executive Deep-Dive on BI Tab
  const handleGenerateAiExecutiveBrief = async () => {
    if (!analytics || isGeneratingAiBrief) return;
    setIsGeneratingAiBrief(true);
    const prompt = `Provide a concise, high-impact executive briefing for this Indian restaurant with 3 specific, quantified growth opportunities in ${activeCurrency.symbol} (${activeCurrency.code}) based on exact menu item performance and customer spending data.`;
    const res = await callGeminiAnalystAPI(prompt, analytics, [], gstBreakdown);
    setAiEnhancedSummary(res);
    setIsGeneratingAiBrief(false);
  };

  // Re-compute BI Report when analytics or currencyCode changes
  const biReportCards = useMemo(() => generateStructuredBIReport(analytics), [analytics, currencyCode]);

  // Filtered Menu Items for Menu Intelligence Table
  const filteredMenuItems = useMemo(() => {
    if (!analytics) return [];
    let list = [...analytics.items.all];
    if (menuFilterTab === 'stars') list = list.filter((i) => i.quadrant === 'Star');
    else if (menuFilterTab === 'plowhorses') list = list.filter((i) => i.quadrant === 'Plowhorse');
    else if (menuFilterTab === 'puzzles') list = list.filter((i) => i.quadrant === 'Puzzle');
    else if (menuFilterTab === 'dogs') list = list.filter((i) => i.quadrant === 'Dog');

    if (menuSearch.trim()) {
      const q = menuSearch.toLowerCase();
      list = list.filter((i) => i.name.toLowerCase().includes(q) || i.category.toLowerCase().includes(q));
    }

    list.sort((a, b) => {
      if (menuSortField === 'revenue') return b.revenue - a.revenue;
      if (menuSortField === 'revenue_asc') return a.revenue - b.revenue;
      if (menuSortField === 'unitsSold') return b.unitsSold - a.unitsSold;
      if (menuSortField === 'orderFrequency') return b.orderFrequency - a.orderFrequency;
      if (menuSortField === 'avgPrice') return b.avgPrice - a.avgPrice;
      return 0;
    });
    return list;
  }, [analytics, menuFilterTab, menuSearch, menuSortField]);

  // Filtered Customers for Customer Insights Table
  const filteredCustomers = useMemo(() => {
    if (!analytics) return [];
    let list = [...analytics.customers.all];
    if (customerSegmentFilter !== 'ALL') {
      list = list.filter((c) => c.segment === customerSegmentFilter);
    }
    if (customerSearch.trim()) {
      const q = customerSearch.toLowerCase();
      list = list.filter(
        (c) =>
          c.id.toLowerCase().includes(q) ||
          c.email.toLowerCase().includes(q) ||
          c.top1Item.toLowerCase().includes(q) ||
          c.top2Item.toLowerCase().includes(q)
      );
    }
    list.sort((a, b) => {
      if (customerSortField === 'totalRevenue') return b.totalRevenue - a.totalRevenue;
      if (customerSortField === 'totalUnits') return b.totalUnits - a.totalUnits;
      if (customerSortField === 'avgSpendPerUnit') return b.avgSpendPerUnit - a.avgSpendPerUnit;
      return 0;
    });
    return list;
  }, [analytics, customerSegmentFilter, customerSearch, customerSortField]);

  const navItems = [
    { id: 'overview', label: 'Overview', icon: LayoutDashboard },
    { id: 'sales', label: 'Sales Analytics', icon: TrendingUp },
    { id: 'menu', label: 'Menu Intelligence', icon: Utensils },
    { id: 'customers', label: 'Customer Insights', icon: Users },
    { id: 'ai', label: 'AI Analyst', icon: Sparkles, badge: 'AI' },
    { id: 'upload', label: 'Upload Data', icon: UploadCloud },
  ];

  const suggestedQuestions = [
    'Which item generated the highest revenue?',
    'What are my top five best-selling items?',
    'Which customers spend the most?',
    'What percentage of revenue comes from the top three items?',
    'Which menu items should I focus on next month?',
    'How did sales change over the last six months?',
    'Which items have low sales performance?',
    'What can I do to increase restaurant revenue?',
  ];

  return (
    <div className="min-h-screen bg-[#090A0F] text-[#F9FAFB] flex flex-col lg:flex-row font-sans selection:bg-emerald-500/30 selection:text-emerald-200">
      {/* ====================================================================
          SIDEBAR NAVIGATION (Desktop & Mobile Drawer)
         ==================================================================== */}
      <aside
        className={`fixed inset-y-0 left-0 z-40 w-64 bg-[#0D0F17] border-r border-white/[0.07] flex flex-col transition-transform duration-200 lg:static lg:translate-x-0 ${
          mobileMenuOpen ? 'translate-x-0' : '-translate-x-full'
        }`}
      >
        {/* Brand Header */}
        <div className="h-16 px-5 flex items-center justify-between border-b border-white/[0.07]">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-emerald-500/15 border border-emerald-500/30 flex items-center justify-center text-emerald-400">
              <Utensils className="w-4 h-4" />
            </div>
            <div>
              <div className="text-sm font-semibold tracking-tight text-white flex items-center gap-1.5">
                RestoPulse <span className="text-[10px] font-mono uppercase px-1.5 py-0.5 rounded bg-white/[0.06] text-emerald-400">BI</span>
              </div>
              <div className="text-[11px] text-gray-500">Revenue & Menu Intelligence</div>
            </div>
          </div>
          <button
            onClick={() => setMobileMenuOpen(false)}
            className="lg:hidden p-1.5 text-gray-400 hover:text-white rounded-lg"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Navigation Links */}
        <nav className="flex-1 p-3 space-y-1 overflow-y-auto">
          <div className="px-2.5 py-1.5 text-[10px] font-semibold uppercase tracking-wider text-gray-500">
            Analytics Workspace
          </div>
          {navItems.map((item) => {
            const Icon = item.icon;
            const isActive = activeTab === item.id;
            return (
              <button
                key={item.id}
                onClick={() => {
                  setActiveTab(item.id);
                  setMobileMenuOpen(false);
                }}
                className={`w-full flex items-center justify-between px-3 py-2.5 rounded-lg text-xs font-medium transition-all ${
                  isActive
                    ? 'bg-emerald-500/12 text-emerald-400 border border-emerald-500/25 shadow-sm'
                    : 'text-gray-400 hover:text-gray-200 hover:bg-white/[0.04] border border-transparent'
                }`}
              >
                <span className="flex items-center gap-2.5">
                  <Icon className={`w-4 h-4 ${isActive ? 'text-emerald-400' : 'text-gray-400'}`} />
                  {item.label}
                </span>
                {item.badge && (
                  <span className="px-1.5 py-0.5 text-[10px] font-mono rounded bg-indigo-500/20 text-indigo-300 border border-indigo-500/30">
                    {item.badge}
                  </span>
                )}
              </button>
            );
          })}
        </nav>

        {/* Active Dataset Status Card in Sidebar Footer */}
        {analytics && (
          <div className="p-3.5 m-3 rounded-xl bg-[#131620] border border-white/[0.07] space-y-2.5">
            <div className="flex items-center justify-between">
              <span className="text-[10px] font-mono uppercase tracking-wider text-gray-400 flex items-center gap-1.5">
                <span className={`w-2 h-2 rounded-full ${analytics.meta.isDemo ? 'bg-amber-400' : 'bg-emerald-400'}`} />
                {analytics.meta.isDemo ? 'Sample Demo Mode' : 'Live Uploaded Data'}
              </span>
              <button
                onClick={() => setActiveTab('upload')}
                className="text-[11px] text-emerald-400 hover:underline font-medium"
              >
                Change
              </button>
            </div>
            <div className="text-xs font-medium text-gray-200 truncate" title={analytics.meta.datasetName}>
              {analytics.meta.datasetName}
            </div>
            <div className="flex items-center justify-between text-[11px] text-gray-400 pt-1 border-t border-white/[0.06]">
              <span>Locale / TZ:</span>
              <span className="font-mono text-emerald-400">en-IN · IST</span>
            </div>
            <div className="flex items-center justify-between text-[11px] text-gray-400">
              <span>Currency:</span>
              <span className="font-mono text-gray-200">{activeCurrency.code} ({activeCurrency.symbol})</span>
            </div>
            <div className="flex items-center justify-between text-[11px] text-gray-400">
              <span>Records:</span>
              <span className="font-mono text-gray-300 tabular-nums">
                {formatNumber(analytics.hygiene.validRowCount)} rows
              </span>
            </div>
          </div>
        )}
      </aside>

      {/* Backdrop for mobile sidebar */}
      {mobileMenuOpen && (
        <div
          onClick={() => setMobileMenuOpen(false)}
          className="fixed inset-0 bg-black/60 z-30 lg:hidden backdrop-blur-xs"
        />
      )}

      {/* ====================================================================
          MAIN CONTENT AREA
         ==================================================================== */}
      <div className="flex-1 flex flex-col min-w-0">
        {/* Top Header Bar */}
        <header className="h-16 px-4 lg:px-8 border-b border-white/[0.07] bg-[#090A0F]/90 backdrop-blur-md sticky top-0 z-20 flex items-center justify-between gap-3">
          <div className="flex items-center gap-3 min-w-0">
            <button
              onClick={() => setMobileMenuOpen(true)}
              className="lg:hidden p-2 rounded-lg bg-[#161922] border border-white/10 text-gray-300 hover:text-white"
            >
              <Menu className="w-4 h-4" />
            </button>
            <div className="min-w-0">
              <div className="flex items-center gap-2">
                <h1 className="text-base font-semibold text-white truncate">
                  {navItems.find((n) => n.id === activeTab)?.label}
                </h1>
                {analytics?.meta.isDemo && (
                  <span className="hidden sm:inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-medium bg-amber-500/10 text-amber-300 border border-amber-500/25">
                    <Info className="w-3 h-3" /> Indian Demo Dataset
                  </span>
                )}
              </div>
              <p className="text-xs text-gray-400 truncate hidden sm:block">
                {analytics
                  ? `${analytics.meta.datasetName} · ${
                      analytics.timeSeries.hasDateData
                        ? `${analytics.hygiene.dateRange?.start} to ${analytics.hygiene.dateRange?.end}`
                        : 'Cumulative Customer & Menu Summary'
                    }`
                  : 'Upload a CSV or Excel dataset to begin'}
              </p>
            </div>
          </div>

          {/* Quick Actions: Currency Selector, GST Config, Demo Switcher, Upload & Export */}
          <div className="flex items-center gap-2">
            {/* Currency Selector (Preserves uploaded numeric values) */}
            <select
              value={currencyCode}
              onChange={(e) => setCurrencyCode(e.target.value)}
              title="Select Display Currency (Preserves original uploaded numbers)"
              className="hidden md:block px-2.5 py-1.5 rounded-lg bg-[#131620] border border-white/10 text-xs text-gray-200 font-mono focus:outline-none focus:border-emerald-500/50"
            >
              {Object.values(CURRENCY_OPTIONS).map((c) => (
                <option key={c.code} value={c.code}>
                  {c.code} ({c.symbol.trim()})
                </option>
              ))}
            </select>

            {/* Configurable GST Treatment Selector */}
            <select
              value={gstMode}
              onChange={(e) => setGstMode(e.target.value)}
              title="Configure GST Treatment (Does not assume GST unless specified)"
              className="hidden lg:block px-2.5 py-1.5 rounded-lg bg-[#131620] border border-white/10 text-xs text-gray-300 focus:outline-none focus:border-emerald-500/50 max-w-[190px] truncate"
            >
              {GST_OPTIONS.map((opt) => (
                <option key={opt.id} value={opt.id}>
                  {opt.label}
                </option>
              ))}
            </select>

            <div className="hidden xl:flex items-center bg-[#131620] p-1 rounded-lg border border-white/[0.07] text-xs">
              <button
                onClick={() => loadDemoDataset('transaction', false)}
                className={`px-2.5 py-1 rounded-md transition-all ${
                  analytics?.meta.isDemo && analytics?.meta.schemaType === 'transaction'
                    ? 'bg-emerald-500/20 text-emerald-300 font-medium border border-emerald-500/30'
                    : 'text-gray-400 hover:text-gray-200'
                }`}
                title="Load 6-Month Indian Restaurant Dated Transaction Log Demo"
              >
                6M Transaction Demo
              </button>
              <button
                onClick={() => loadDemoDataset('customer_summary', false)}
                className={`px-2.5 py-1 rounded-md transition-all ${
                  analytics?.meta.isDemo && analytics?.meta.schemaType === 'customer_summary'
                    ? 'bg-emerald-500/20 text-emerald-300 font-medium border border-emerald-500/30'
                    : 'text-gray-400 hover:text-gray-200'
                }`}
                title="Load Indian Customer-Summary Schema Demo (1st/2nd/3rd Most Ordered Items)"
              >
                Customer-Summary Demo
              </button>
            </div>

            <button
              onClick={() => setActiveTab('upload')}
              className="px-3 py-1.5 rounded-lg text-xs font-medium bg-[#161922] hover:bg-[#1E2230] text-gray-200 border border-white/10 flex items-center gap-1.5 transition-colors"
            >
              <UploadCloud className="w-3.5 h-3.5 text-emerald-400" />
              <span className="hidden sm:inline">Upload File</span>
            </button>

            {analytics && (
              <button
                onClick={downloadSummaryReport}
                className="px-3 py-1.5 rounded-lg text-xs font-medium bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-semibold flex items-center gap-1.5 transition-colors shadow-sm"
              >
                <Download className="w-3.5 h-3.5" />
                <span className="hidden md:inline">Export Report</span>
              </button>
            )}
          </div>
        </header>

        {/* Main Scrollable Viewport */}
        <main className="flex-1 p-4 lg:p-8 max-w-[1600px] w-full mx-auto space-y-6">
          {/* ================================================================
              SECTION 1: OVERVIEW DASHBOARD
             ================================================================ */}
          {activeTab === 'overview' && analytics && (
            <div className="space-y-6">
              {/* Schema, Regional Settings & Optional GST Breakdown Banner */}
              <div className="p-4 rounded-xl bg-[#131620] border border-white/[0.07] space-y-3">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                  <div className="flex items-start sm:items-center gap-3">
                    <div className="p-2.5 rounded-lg bg-emerald-500/10 border border-emerald-500/20 text-emerald-400 shrink-0">
                      <Database className="w-5 h-5" />
                    </div>
                    <div>
                      <div className="flex flex-wrap items-center gap-2">
                        <span className="text-sm font-semibold text-white">{analytics.meta.datasetName}</span>
                        <span className="px-2 py-0.5 text-[11px] font-mono rounded bg-white/[0.06] text-gray-300">
                          {analytics.meta.schemaType === 'customer_summary' ? 'Customer-Summary Schema' : 'Transaction-Level Schema'}
                        </span>
                        <span className="px-2 py-0.5 text-[11px] font-mono rounded bg-emerald-500/15 text-emerald-300 border border-emerald-500/30">
                          {activeCurrency.code} ({activeCurrency.symbol.trim()}) · DD-MM-YYYY · IST
                        </span>
                        {analytics.meta.isDemo && (
                          <span className="px-2 py-0.5 text-[11px] font-mono rounded bg-amber-500/15 text-amber-300 border border-amber-500/30">
                            Sample Data
                          </span>
                        )}
                      </div>
                      <p className="text-xs text-gray-400 mt-0.5">
                        {analytics.timeSeries.hasDateData
                          ? `Full date-level granularity enabled (${analytics.hygiene.dateRange?.start} to ${analytics.hygiene.dateRange?.end} · ${analytics.hygiene.dateRange?.daysCount} active days).`
                          : `Customer-summary mode active. Item revenue, customer segmentation, and Pareto analytics are calculated; upload transaction dates for daily/monthly timeline charts.`}
                      </p>
                    </div>
                  </div>
                  <div className="flex flex-wrap items-center gap-2 shrink-0">
                    <button
                      onClick={() => loadDemoDataset(analytics.meta.schemaType === 'transaction' ? 'customer_summary' : 'transaction', false)}
                      className="px-3 py-1.5 rounded-lg text-xs font-medium bg-white/[0.05] hover:bg-white/[0.1] text-gray-200 border border-white/10 transition-colors"
                    >
                      Switch to {analytics.meta.schemaType === 'transaction' ? 'Customer-Summary' : '6M Transaction'} Demo
                    </button>
                    <button
                      onClick={() => setActiveTab('ai')}
                      className="px-3 py-1.5 rounded-lg text-xs font-medium bg-indigo-500/20 hover:bg-indigo-500/30 text-indigo-300 border border-indigo-500/30 flex items-center gap-1.5 transition-colors"
                    >
                      <Sparkles className="w-3.5 h-3.5" /> Ask AI Analyst
                    </button>
                  </div>
                </div>

                {/* Optional GST Tax Analytics Strip when user specifies GST mode */}
                {gstBreakdown?.isActive && (
                  <div className="pt-3 border-t border-white/[0.06] flex flex-wrap items-center justify-between gap-3 text-xs">
                    <div className="flex items-center gap-2 text-amber-300">
                      <CheckCircle2 className="w-4 h-4 text-emerald-400" />
                      <span className="font-medium">Active GST Mode:</span>
                      <span className="font-mono text-gray-200">{gstBreakdown.label}</span>
                    </div>
                    <div className="flex flex-wrap items-center gap-4 font-mono">
                      <span className="text-gray-400">
                        Net Ex-GST: <strong className="text-white">{formatCurrency(gstBreakdown.netRevenueExGst)}</strong>
                      </span>
                      <span className="text-gray-400">
                        Est. GST ({gstBreakdown.rate}%): <strong className="text-amber-300">{formatCurrency(gstBreakdown.gstAmount)}</strong>
                      </span>
                      <span className="text-gray-400">
                        Gross Incl. GST: <strong className="text-emerald-400">{formatCurrency(gstBreakdown.grossRevenueInclGst)}</strong>
                      </span>
                    </div>
                  </div>
                )}
              </div>

              {/* 7 Primary Executive KPI Cards */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 xl:grid-cols-7 gap-3.5">
                <div className="p-4 rounded-xl bg-[#161922] border border-white/[0.07] flex flex-col justify-between">
                  <div className="flex items-center justify-between text-xs text-gray-400">
                    <span>Total Revenue</span>
                    <DollarSign className="w-4 h-4 text-emerald-400" />
                  </div>
                  <div className="mt-3">
                    <div className="text-xl font-bold text-white font-mono tabular-nums">
                      {formatCurrency(analytics.kpis.totalRevenue, 0)}
                    </div>
                    <div className="text-[11px] text-emerald-400 mt-1 flex items-center gap-1">
                      <CheckCircle2 className="w-3 h-3" /> Verified Total
                    </div>
                  </div>
                </div>

                <div className="p-4 rounded-xl bg-[#161922] border border-white/[0.07] flex flex-col justify-between">
                  <div className="flex items-center justify-between text-xs text-gray-400">
                    <span>Total Items Sold</span>
                    <ShoppingBag className="w-4 h-4 text-indigo-400" />
                  </div>
                  <div className="mt-3">
                    <div className="text-xl font-bold text-white font-mono tabular-nums">
                      {formatNumber(analytics.kpis.totalUnitsSold)}
                    </div>
                    <div className="text-[11px] text-gray-400 mt-1">
                      Across {analytics.items.all.length} menu items
                    </div>
                  </div>
                </div>

                <div className="p-4 rounded-xl bg-[#161922] border border-white/[0.07] flex flex-col justify-between">
                  <div className="flex items-center justify-between text-xs text-gray-400">
                    <span>Total Customers</span>
                    <Users className="w-4 h-4 text-sky-400" />
                  </div>
                  <div className="mt-3">
                    <div className="text-xl font-bold text-white font-mono tabular-nums">
                      {formatNumber(analytics.kpis.totalCustomers)}
                    </div>
                    <div className="text-[11px] text-gray-400 mt-1">
                      {analytics.kpis.repeatCustomersCount} repeat patrons
                    </div>
                  </div>
                </div>

                <div className="p-4 rounded-xl bg-[#161922] border border-white/[0.07] flex flex-col justify-between">
                  <div className="flex items-center justify-between text-xs text-gray-400">
                    <span>Avg Rev / Customer</span>
                    <TrendingUp className="w-4 h-4 text-emerald-400" />
                  </div>
                  <div className="mt-3">
                    <div className="text-xl font-bold text-white font-mono tabular-nums">
                      {formatCurrency(analytics.kpis.avgRevenuePerCustomer)}
                    </div>
                    <div className="text-[11px] text-gray-400 mt-1">
                      Customer LTV (6M)
                    </div>
                  </div>
                </div>

                <div className="p-4 rounded-xl bg-[#161922] border border-white/[0.07] flex flex-col justify-between">
                  <div className="flex items-center justify-between text-xs text-gray-400">
                    <span>Avg Item Price</span>
                    <Layers className="w-4 h-4 text-amber-400" />
                  </div>
                  <div className="mt-3">
                    <div className="text-xl font-bold text-white font-mono tabular-nums">
                      {formatCurrency(analytics.kpis.avgItemPrice)}
                    </div>
                    <div className="text-[11px] text-gray-400 mt-1">
                      Weighted menu avg
                    </div>
                  </div>
                </div>

                <div className="p-4 rounded-xl bg-[#161922] border border-white/[0.07] flex flex-col justify-between">
                  <div className="flex items-center justify-between text-xs text-gray-400">
                    <span>Repeat Customer %</span>
                    <Repeat className="w-4 h-4 text-purple-400" />
                  </div>
                  <div className="mt-3">
                    <div className="text-xl font-bold text-white font-mono tabular-nums">
                      {formatPercent(analytics.kpis.repeatCustomerPct)}
                    </div>
                    <div className="text-[11px] text-gray-400 mt-1">
                      2+ items ordered
                    </div>
                  </div>
                </div>

                <div className="p-4 rounded-xl bg-[#161922] border border-emerald-500/25 flex flex-col justify-between">
                  <div className="flex items-center justify-between text-xs text-emerald-300">
                    <span>Most Popular Item</span>
                    <Award className="w-4 h-4 text-emerald-400" />
                  </div>
                  <div className="mt-2">
                    <div className="text-sm font-bold text-white line-clamp-1" title={analytics.kpis.mostPopularItemByUnits?.name}>
                      {analytics.kpis.mostPopularItemByUnits?.name || '—'}
                    </div>
                    <div className="text-[11px] font-mono text-emerald-400 mt-1">
                      {formatNumber(analytics.kpis.mostPopularItemByUnits?.unitsSold)} units · {formatCurrency(analytics.kpis.mostPopularItemByUnits?.revenue, 0)}
                    </div>
                  </div>
                </div>
              </div>

              {/* Primary Overview Charts Row */}
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                {/* Left 2 Columns: 6-Month Revenue Trend (or Item Pareto if non-dated Summary Schema) */}
                <div className="lg:col-span-2 p-5 rounded-xl bg-[#161922] border border-white/[0.07] flex flex-col">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 mb-5">
                    <div>
                      <h2 className="text-sm font-semibold text-white">
                        {analytics.timeSeries.hasDateData
                          ? 'Six-Month Revenue & Volume Trajectory'
                          : 'Menu Item Revenue Pareto Curve (Cumulative Contribution)'}
                      </h2>
                      <p className="text-xs text-gray-400">
                        {analytics.timeSeries.hasDateData
                          ? 'Monthly revenue performance and item volume progression across the 6-month window'
                          : 'Showing item revenue concentration (upload transaction dates for chronological trends)'}
                      </p>
                    </div>
                    <button
                      onClick={() => setActiveTab('sales')}
                      className="text-xs text-emerald-400 hover:text-emerald-300 flex items-center gap-1 font-medium self-start sm:self-auto"
                    >
                      Deep Dive <ChevronRight className="w-3.5 h-3.5" />
                    </button>
                  </div>

                  <div className="h-72 w-full">
                    {analytics.timeSeries.hasDateData ? (
                      <ResponsiveContainer width="100%" height="100%">
                        <ComposedChart data={analytics.timeSeries.monthly} margin={{ top: 10, right: 15, left: 0, bottom: 5 }}>
                          <defs>
                            <linearGradient id="revGrad" x1="0" y1="0" x2="0" y2="1">
                              <stop offset="5%" stopColor="#10B981" stopOpacity={0.35} />
                              <stop offset="95%" stopColor="#10B981" stopOpacity={0.0} />
                            </linearGradient>
                          </defs>
                          <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.06)" vertical={false} />
                          <XAxis dataKey="label" stroke="#9CA3AF" fontSize={11} tickLine={false} />
                          <YAxis
                            yAxisId="left"
                            stroke="#9CA3AF"
                            fontSize={11}
                            tickLine={false}
                            tickFormatter={formatCompactCurrency}
                          />
                          <YAxis
                            yAxisId="right"
                            orientation="right"
                            stroke="#6366F1"
                            fontSize={11}
                            tickLine={false}
                          />
                          <Tooltip content={<CustomChartTooltip />} />
                          <Legend wrapperStyle={{ fontSize: '12px' }} />
                          <Area
                            yAxisId="left"
                            type="monotone"
                            dataKey="revenue"
                            name="Monthly Revenue"
                            stroke="#10B981"
                            strokeWidth={2.5}
                            fill="url(#revGrad)"
                          />
                          <Bar
                            yAxisId="right"
                            dataKey="units"
                            name="Units Sold"
                            fill="#6366F1"
                            radius={[4, 4, 0, 0]}
                            barSize={24}
                            fillOpacity={0.75}
                          />
                        </ComposedChart>
                      </ResponsiveContainer>
                    ) : (
                      <ResponsiveContainer width="100%" height="100%">
                        <ComposedChart data={analytics.items.byRevenue.slice(0, 10)} margin={{ top: 10, right: 15, left: 0, bottom: 30 }}>
                          <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.06)" vertical={false} />
                          <XAxis
                            dataKey="name"
                            stroke="#9CA3AF"
                            fontSize={10}
                            tickLine={false}
                            angle={-20}
                            textAnchor="end"
                            interval={0}
                          />
                          <YAxis yAxisId="left" stroke="#9CA3AF" fontSize={11} tickFormatter={formatCompactCurrency} />
                          <YAxis yAxisId="right" orientation="right" stroke="#6366F1" fontSize={11} domain={[0, 100]} tickFormatter={(v) => `${v}%`} />
                          <Tooltip content={<CustomChartTooltip />} />
                          <Legend wrapperStyle={{ fontSize: '12px' }} />
                          <Bar yAxisId="left" dataKey="revenue" name="Item Revenue" fill="#10B981" radius={[4, 4, 0, 0]} barSize={28} />
                          <Line yAxisId="right" type="monotone" dataKey="cumulativeRevenueShare" name="Cumulative Share %" stroke="#6366F1" strokeWidth={2.5} dot={{ r: 3 }} />
                        </ComposedChart>
                      </ResponsiveContainer>
                    )}
                  </div>
                </div>

                {/* Right Column: Revenue Share Donut Chart */}
                <div className="p-5 rounded-xl bg-[#161922] border border-white/[0.07] flex flex-col justify-between">
                  <div>
                    <h2 className="text-sm font-semibold text-white">Sales Revenue Distribution</h2>
                    <p className="text-xs text-gray-400">
                      Top menu items vs. rest of catalog ({formatPercent(analytics.kpis.top3ItemsRevPct)} from top 3)
                    </p>
                  </div>

                  <div className="h-56 w-full my-2">
                    <ResponsiveContainer width="100%" height="100%">
                      <PieChart>
                        <Pie
                          data={(() => {
                            const top5 = analytics.items.byRevenue.slice(0, 5);
                            const restRev = analytics.items.byRevenue.slice(5).reduce((s, i) => s + i.revenue, 0);
                            return [
                              ...top5.map((i) => ({ name: i.name, value: i.revenue })),
                              ...(restRev > 0 ? [{ name: 'All Other Items', value: Number(restRev.toFixed(2)) }] : []),
                            ];
                          })()}
                          cx="50%"
                          cy="50%"
                          innerRadius={52}
                          outerRadius={80}
                          paddingAngle={3}
                          dataKey="value"
                        >
                          {CHART_COLORS.map((color, idx) => (
                            <Cell key={idx} fill={color} stroke="rgba(0,0,0,0.3)" />
                          ))}
                        </Pie>
                        <Tooltip content={<CustomChartTooltip />} />
                      </PieChart>
                    </ResponsiveContainer>
                  </div>

                  <div className="space-y-1.5 pt-2 border-t border-white/[0.06]">
                    {analytics.items.byRevenue.slice(0, 4).map((item, idx) => (
                      <div key={item.name} className="flex items-center justify-between text-xs">
                        <span className="flex items-center gap-2 text-gray-300 truncate">
                          <span className="w-2.5 h-2.5 rounded-sm shrink-0" style={{ backgroundColor: CHART_COLORS[idx] }} />
                          <span className="truncate">{item.name}</span>
                        </span>
                        <span className="font-mono text-gray-200 tabular-nums shrink-0">
                          {formatCurrency(item.revenue, 0)} ({formatPercent(item.revenueShare, 1)})
                        </span>
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              {/* Bottom Overview Row: Top Menu Items & Strategic AI Highlights */}
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                {/* Top 6 Menu Items Snapshot */}
                <div className="p-5 rounded-xl bg-[#161922] border border-white/[0.07]">
                  <div className="flex items-center justify-between mb-4">
                    <div>
                      <h2 className="text-sm font-semibold text-white">Top Revenue-Generating Menu Items</h2>
                      <p className="text-xs text-gray-400">Ranked by dollar contribution & Menu Engineering classification</p>
                    </div>
                    <button
                      onClick={() => setActiveTab('menu')}
                      className="text-xs text-emerald-400 hover:underline flex items-center gap-1"
                    >
                      All {analytics.items.all.length} Items <ChevronRight className="w-3.5 h-3.5" />
                    </button>
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse">
                      <thead>
                        <tr className="border-b border-white/[0.07] text-[11px] uppercase text-gray-400">
                          <th className="pb-2.5 font-medium">Menu Item</th>
                          <th className="pb-2.5 font-medium">Matrix</th>
                          <th className="pb-2.5 font-medium text-right">Units</th>
                          <th className="pb-2.5 font-medium text-right">Revenue</th>
                          <th className="pb-2.5 font-medium text-right">Share</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-white/[0.05] text-xs">
                        {analytics.items.byRevenue.slice(0, 6).map((item) => (
                          <tr key={item.name} className="hover:bg-white/[0.02]">
                            <td className="py-2.5 font-medium text-gray-100">{item.name}</td>
                            <td className="py-2.5">
                              <span
                                className={`px-2 py-0.5 rounded text-[10px] font-medium ${
                                  item.quadrant === 'Star'
                                    ? 'bg-emerald-500/15 text-emerald-300 border border-emerald-500/30'
                                    : item.quadrant === 'Plowhorse'
                                    ? 'bg-amber-500/15 text-amber-300 border border-amber-500/30'
                                    : item.quadrant === 'Puzzle'
                                    ? 'bg-indigo-500/15 text-indigo-300 border border-indigo-500/30'
                                    : 'bg-rose-500/15 text-rose-300 border border-rose-500/30'
                                }`}
                              >
                                {item.quadrant}
                              </span>
                            </td>
                            <td className="py-2.5 text-right font-mono text-gray-300 tabular-nums">
                              {formatNumber(item.unitsSold)}
                            </td>
                            <td className="py-2.5 text-right font-mono font-semibold text-white tabular-nums">
                              {formatCurrency(item.revenue)}
                            </td>
                            <td className="py-2.5 text-right font-mono text-emerald-400 tabular-nums">
                              {formatPercent(item.revenueShare)}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>

                {/* Executive AI Takeaways Preview */}
                <div className="p-5 rounded-xl bg-[#161922] border border-white/[0.07] flex flex-col justify-between">
                  <div>
                    <div className="flex items-center justify-between mb-4">
                      <div className="flex items-center gap-2">
                        <div className="p-1.5 rounded-lg bg-indigo-500/15 text-indigo-400 border border-indigo-500/30">
                          <Sparkles className="w-4 h-4" />
                        </div>
                        <div>
                          <h2 className="text-sm font-semibold text-white">Automated Business Intelligence Highlights</h2>
                          <p className="text-xs text-gray-400">Live diagnostic findings from your active dataset</p>
                        </div>
                      </div>
                      <button
                        onClick={() => setActiveTab('ai')}
                        className="text-xs text-indigo-400 hover:underline flex items-center gap-1"
                      >
                        Full BI Report <ChevronRight className="w-3.5 h-3.5" />
                      </button>
                    </div>

                    <div className="space-y-3">
                      {biReportCards.slice(0, 3).map((card) => (
                        <div
                          key={card.id}
                          className="p-3.5 rounded-lg bg-[#11131A] border border-white/[0.06] space-y-1"
                        >
                          <div className="flex items-center justify-between">
                            <span className="text-[11px] font-medium text-emerald-400">{card.question}</span>
                            <span className="text-[10px] font-mono uppercase px-1.5 py-0.5 rounded bg-white/[0.05] text-gray-400">
                              {card.badge}
                            </span>
                          </div>
                          <div className="text-xs font-semibold text-white">{card.headline}</div>
                          <p className="text-xs text-gray-400 line-clamp-2">{card.analysis}</p>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* ================================================================
              SECTION 2: SALES ANALYTICS (REVENUE TRENDS & PARETO)
             ================================================================ */}
          {activeTab === 'sales' && analytics && (
            <div className="space-y-6">
              {/* Period-Over-Period Comparison Cards (if date data exists) or Concentration Summary */}
              {analytics.timeSeries.hasDateData && analytics.timeSeries.periodComparison ? (
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  <div className="p-4 rounded-xl bg-[#161922] border border-white/[0.07]">
                    <div className="text-xs text-gray-400">Latest Month vs. Previous Month</div>
                    <div className="mt-2 flex items-baseline justify-between">
                      <div>
                        <div className="text-lg font-bold text-white font-mono">
                          {formatCurrency(analytics.timeSeries.periodComparison.latestMonthRevenue)}
                        </div>
                        <div className="text-[11px] text-gray-400">
                          {analytics.timeSeries.periodComparison.latestMonthLabel} vs. {analytics.timeSeries.periodComparison.previousMonthLabel} ({formatCurrency(analytics.timeSeries.periodComparison.previousMonthRevenue)})
                        </div>
                      </div>
                      <span
                        className={`px-2 py-1 rounded-md text-xs font-mono font-semibold flex items-center gap-1 ${
                          analytics.timeSeries.periodComparison.latestVsPrevPct >= 0
                            ? 'bg-emerald-500/15 text-emerald-400'
                            : 'bg-rose-500/15 text-rose-400'
                        }`}
                      >
                        {analytics.timeSeries.periodComparison.latestVsPrevPct >= 0 ? (
                          <ArrowUpRight className="w-3.5 h-3.5" />
                        ) : (
                          <ArrowDownRight className="w-3.5 h-3.5" />
                        )}
                        {analytics.timeSeries.periodComparison.latestVsPrevPct >= 0 ? '+' : ''}
                        {analytics.timeSeries.periodComparison.latestVsPrevPct}%
                      </span>
                    </div>
                  </div>

                  <div className="p-4 rounded-xl bg-[#161922] border border-white/[0.07]">
                    <div className="text-xs text-gray-400">3-Month Period Comparison (2nd Half vs. 1st Half)</div>
                    <div className="mt-2 flex items-baseline justify-between">
                      <div>
                        <div className="text-lg font-bold text-white font-mono">
                          {formatCurrency(analytics.timeSeries.periodComparison.secondHalfRevenue)}
                        </div>
                        <div className="text-[11px] text-gray-400">
                          {analytics.timeSeries.periodComparison.secondHalfLabel} vs. {analytics.timeSeries.periodComparison.firstHalfLabel}
                        </div>
                      </div>
                      <span
                        className={`px-2 py-1 rounded-md text-xs font-mono font-semibold flex items-center gap-1 ${
                          analytics.timeSeries.periodComparison.halfOverHalfPct >= 0
                            ? 'bg-emerald-500/15 text-emerald-400'
                            : 'bg-rose-500/15 text-rose-400'
                        }`}
                      >
                        {analytics.timeSeries.periodComparison.halfOverHalfPct >= 0 ? (
                          <ArrowUpRight className="w-3.5 h-3.5" />
                        ) : (
                          <ArrowDownRight className="w-3.5 h-3.5" />
                        )}
                        {analytics.timeSeries.periodComparison.halfOverHalfPct >= 0 ? '+' : ''}
                        {analytics.timeSeries.periodComparison.halfOverHalfPct}%
                      </span>
                    </div>
                  </div>

                  <div className="p-4 rounded-xl bg-[#161922] border border-white/[0.07]">
                    <div className="text-xs text-gray-400">Top 3 Menu Items Revenue Concentration</div>
                    <div className="mt-2 flex items-baseline justify-between">
                      <div>
                        <div className="text-lg font-bold text-white font-mono">
                          {formatPercent(analytics.kpis.top3ItemsRevPct)}
                        </div>
                        <div className="text-[11px] text-gray-400">
                          Generated by top 3 items out of {analytics.items.all.length}
                        </div>
                      </div>
                      <span className="px-2 py-1 rounded-md text-xs font-mono bg-indigo-500/15 text-indigo-300">
                        Pareto Ratio
                      </span>
                    </div>
                  </div>
                </div>
              ) : (
                <div className="p-4 rounded-xl bg-amber-500/10 border border-amber-500/25 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3">
                  <div className="flex items-start gap-3">
                    <AlertCircle className="w-5 h-5 text-amber-400 shrink-0 mt-0.5" />
                    <div>
                      <div className="text-xs font-semibold text-amber-200">
                        Non-Dated Customer-Summary Dataset Active — Timeline Charts Require Transaction Dates
                      </div>
                      <p className="text-xs text-amber-300/80 mt-0.5">
                        Your uploaded file uses the Customer-Summary schema (cumulative customer & item totals without individual order timestamps). We never fabricate missing date trends. Below you can analyze Item Revenue Pareto and Customer Revenue Deciles, or load the 6-Month Transaction Demo to explore Daily, Weekly, and Monthly trends.
                      </p>
                    </div>
                  </div>
                  <button
                    onClick={() => loadDemoDataset('transaction', false)}
                    className="px-3 py-1.5 rounded-lg text-xs font-medium bg-amber-400 text-slate-950 font-semibold shrink-0 hover:bg-amber-300 transition-colors"
                  >
                    Load 6M Dated Demo
                  </button>
                </div>
              )}

              {/* Interactive Timeline Chart (Daily / Weekly / Monthly Toggles) */}
              {analytics.timeSeries.hasDateData && (
                <div className="p-5 rounded-xl bg-[#161922] border border-white/[0.07]">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-5">
                    <div>
                      <h2 className="text-sm font-semibold text-white">
                        Revenue & Order Volume Breakdown ({timeGranularity.toUpperCase()})
                      </h2>
                      <p className="text-xs text-gray-400">
                        Switch between Daily, Weekly, and Monthly aggregation across the 6-month history
                      </p>
                    </div>
                    <div className="flex items-center bg-[#11131A] p-1 rounded-lg border border-white/[0.07] text-xs self-start">
                      {['daily', 'weekly', 'monthly'].map((gran) => (
                        <button
                          key={gran}
                          onClick={() => setTimeGranularity(gran)}
                          className={`px-3 py-1 rounded-md capitalize transition-all ${
                            timeGranularity === gran
                              ? 'bg-emerald-500 text-slate-950 font-semibold'
                              : 'text-gray-400 hover:text-gray-200'
                          }`}
                        >
                          {gran}
                        </button>
                      ))}
                    </div>
                  </div>

                  <div className="h-80 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      {timeGranularity === 'monthly' ? (
                        <BarChart data={analytics.timeSeries.monthly} margin={{ top: 10, right: 20, left: 0, bottom: 5 }}>
                          <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.06)" vertical={false} />
                          <XAxis dataKey="label" stroke="#9CA3AF" fontSize={11} />
                          <YAxis stroke="#9CA3AF" fontSize={11} tickFormatter={formatCompactCurrency} />
                          <Tooltip content={<CustomChartTooltip />} />
                          <Legend wrapperStyle={{ fontSize: '12px' }} />
                          <Bar dataKey="revenue" name="Monthly Revenue" fill="#10B981" radius={[5, 5, 0, 0]} barSize={36} />
                        </BarChart>
                      ) : timeGranularity === 'weekly' ? (
                        <AreaChart data={analytics.timeSeries.weekly} margin={{ top: 10, right: 20, left: 0, bottom: 5 }}>
                          <defs>
                            <linearGradient id="wkGrad" x1="0" y1="0" x2="0" y2="1">
                              <stop offset="5%" stopColor="#6366F1" stopOpacity={0.35} />
                              <stop offset="95%" stopColor="#6366F1" stopOpacity={0.0} />
                            </linearGradient>
                          </defs>
                          <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.06)" vertical={false} />
                          <XAxis dataKey="period" stroke="#9CA3AF" fontSize={10} />
                          <YAxis stroke="#9CA3AF" fontSize={11} tickFormatter={formatCompactCurrency} />
                          <Tooltip content={<CustomChartTooltip />} />
                          <Legend wrapperStyle={{ fontSize: '12px' }} />
                          <Area type="monotone" dataKey="revenue" name="Weekly Revenue" stroke="#6366F1" strokeWidth={2} fill="url(#wkGrad)" />
                        </AreaChart>
                      ) : (
                        <LineChart data={analytics.timeSeries.daily} margin={{ top: 10, right: 20, left: 0, bottom: 5 }}>
                          <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.06)" vertical={false} />
                          <XAxis dataKey="date" stroke="#9CA3AF" fontSize={10} minTickGap={28} />
                          <YAxis stroke="#9CA3AF" fontSize={11} tickFormatter={formatCompactCurrency} />
                          <Tooltip content={<CustomChartTooltip />} />
                          <Legend wrapperStyle={{ fontSize: '12px' }} />
                          <Line type="monotone" dataKey="revenue" name="Daily Revenue" stroke="#10B981" strokeWidth={1.75} dot={false} />
                        </LineChart>
                      )}
                    </ResponsiveContainer>
                  </div>
                </div>
              )}

              {/* Revenue Contribution by Menu Item & Customer Concentration */}
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                {/* Item Revenue Contribution Pareto */}
                <div className="p-5 rounded-xl bg-[#161922] border border-white/[0.07]">
                  <div className="mb-4">
                    <h2 className="text-sm font-semibold text-white">Revenue Contribution by Menu Item (Pareto)</h2>
                    <p className="text-xs text-gray-400">
                      Individual dollar sales and cumulative % of total item revenue
                    </p>
                  </div>
                  <div className="h-80 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <ComposedChart
                        data={analytics.items.byRevenue.slice(0, 12)}
                        margin={{ top: 10, right: 15, left: 0, bottom: 40 }}
                      >
                        <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.06)" vertical={false} />
                        <XAxis
                          dataKey="name"
                          stroke="#9CA3AF"
                          fontSize={10}
                          angle={-28}
                          textAnchor="end"
                          interval={0}
                        />
                        <YAxis yAxisId="left" stroke="#9CA3AF" fontSize={11} tickFormatter={formatCompactCurrency} />
                        <YAxis
                          yAxisId="right"
                          orientation="right"
                          stroke="#F59E0B"
                          fontSize={11}
                          domain={[0, 100]}
                          tickFormatter={(v) => `${v}%`}
                        />
                        <Tooltip content={<CustomChartTooltip />} />
                        <Legend wrapperStyle={{ fontSize: '12px' }} />
                        <Bar yAxisId="left" dataKey="revenue" name="Item Revenue" fill="#10B981" radius={[4, 4, 0, 0]} />
                        <Line
                          yAxisId="right"
                          type="monotone"
                          dataKey="cumulativeRevenueShare"
                          name="Cumulative Share %"
                          stroke="#F59E0B"
                          strokeWidth={2}
                          dot={{ r: 3 }}
                        />
                      </ComposedChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                {/* Revenue Contribution by Customer Segment / Top Customers */}
                <div className="p-5 rounded-xl bg-[#161922] border border-white/[0.07]">
                  <div className="mb-4">
                    <h2 className="text-sm font-semibold text-white">Revenue Contribution by Customer Segment</h2>
                    <p className="text-xs text-gray-400">
                      How total spending is distributed across behavioral customer tiers
                    </p>
                  </div>
                  <div className="h-80 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart
                        data={analytics.customers.segments}
                        layout="vertical"
                        margin={{ top: 10, right: 25, left: 25, bottom: 5 }}
                      >
                        <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.06)" horizontal={false} />
                        <XAxis type="number" stroke="#9CA3AF" fontSize={11} tickFormatter={formatCompactCurrency} />
                        <YAxis type="category" dataKey="segment" stroke="#D1D5DB" fontSize={11} width={130} />
                        <Tooltip content={<CustomChartTooltip />} />
                        <Legend wrapperStyle={{ fontSize: '12px' }} />
                        <Bar dataKey="revenue" name="Segment Revenue" fill="#6366F1" radius={[0, 5, 5, 0]} barSize={24} />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* ================================================================
              SECTION 3: MENU INTELLIGENCE & ENGINEERING MATRIX
             ================================================================ */}
          {activeTab === 'menu' && analytics && (
            <div className="space-y-6">
              {/* 4 Menu Engineering Quadrant Summary Cards */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                <div
                  onClick={() => setMenuFilterTab(menuFilterTab === 'stars' ? 'all' : 'stars')}
                  className={`p-4 rounded-xl border cursor-pointer transition-all ${
                    menuFilterTab === 'stars'
                      ? 'bg-emerald-500/15 border-emerald-500/50'
                      : 'bg-[#161922] border-white/[0.07] hover:border-emerald-500/30'
                  }`}
                >
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-semibold uppercase tracking-wider text-emerald-400">
                      Stars · Prioritize
                    </span>
                    <span className="px-2 py-0.5 rounded-full text-xs font-mono bg-emerald-500/20 text-emerald-300">
                      {analytics.items.stars.length} items
                    </span>
                  </div>
                  <div className="mt-2 text-sm font-semibold text-white">High Volume + High Revenue</div>
                  <p className="text-xs text-gray-400 mt-1">
                    Flagship menu drivers. Maintain strict quality and prime visual placement.
                  </p>
                  <div className="mt-3 pt-2 border-t border-white/[0.06] text-[11px] text-gray-300 truncate">
                    Top: <strong className="text-white">{analytics.items.stars[0]?.name || '—'}</strong>
                  </div>
                </div>

                <div
                  onClick={() => setMenuFilterTab(menuFilterTab === 'plowhorses' ? 'all' : 'plowhorses')}
                  className={`p-4 rounded-xl border cursor-pointer transition-all ${
                    menuFilterTab === 'plowhorses'
                      ? 'bg-amber-500/15 border-amber-500/50'
                      : 'bg-[#161922] border-white/[0.07] hover:border-amber-500/30'
                  }`}
                >
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-semibold uppercase tracking-wider text-amber-400">
                      Plowhorses · Reprice
                    </span>
                    <span className="px-2 py-0.5 rounded-full text-xs font-mono bg-amber-500/20 text-amber-300">
                      {analytics.items.plowhorses.length} items
                    </span>
                  </div>
                  <div className="mt-2 text-sm font-semibold text-white">Strong Sales + Low Revenue</div>
                  <p className="text-xs text-gray-400 mt-1">
                    Popular staples with low revenue yield. Test modest price increases or combo bundling.
                  </p>
                  <div className="mt-3 pt-2 border-t border-white/[0.06] text-[11px] text-gray-300 truncate">
                    Top: <strong className="text-white">{analytics.items.plowhorses[0]?.name || '—'}</strong>
                  </div>
                </div>

                <div
                  onClick={() => setMenuFilterTab(menuFilterTab === 'puzzles' ? 'all' : 'puzzles')}
                  className={`p-4 rounded-xl border cursor-pointer transition-all ${
                    menuFilterTab === 'puzzles'
                      ? 'bg-indigo-500/15 border-indigo-500/50'
                      : 'bg-[#161922] border-white/[0.07] hover:border-indigo-500/30'
                  }`}
                >
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-semibold uppercase tracking-wider text-indigo-400">
                      Puzzles · Promote
                    </span>
                    <span className="px-2 py-0.5 rounded-full text-xs font-mono bg-indigo-500/20 text-indigo-300">
                      {analytics.items.puzzles.length} items
                    </span>
                  </div>
                  <div className="mt-2 text-sm font-semibold text-white">High Revenue + Low Sales</div>
                  <p className="text-xs text-gray-400 mt-1">
                    High-value items that need more attention, server upselling, or targeted promotion.
                  </p>
                  <div className="mt-3 pt-2 border-t border-white/[0.06] text-[11px] text-gray-300 truncate">
                    Top: <strong className="text-white">{analytics.items.puzzles[0]?.name || '—'}</strong>
                  </div>
                </div>

                <div
                  onClick={() => setMenuFilterTab(menuFilterTab === 'dogs' ? 'all' : 'dogs')}
                  className={`p-4 rounded-xl border cursor-pointer transition-all ${
                    menuFilterTab === 'dogs'
                      ? 'bg-rose-500/15 border-rose-500/50'
                      : 'bg-[#161922] border-white/[0.07] hover:border-rose-500/30'
                  }`}
                >
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-semibold uppercase tracking-wider text-rose-400">
                      Dogs · Reconsider
                    </span>
                    <span className="px-2 py-0.5 rounded-full text-xs font-mono bg-rose-500/20 text-rose-300">
                      {analytics.items.dogs.length} items
                    </span>
                  </div>
                  <div className="mt-2 text-sm font-semibold text-white">Low Volume + Low Revenue</div>
                  <p className="text-xs text-gray-400 mt-1">
                    Underperforming items. Consider repricing, recipe overhaul, or menu removal.
                  </p>
                  <div className="mt-3 pt-2 border-t border-white/[0.06] text-[11px] text-gray-300 truncate">
                    Lowest: <strong className="text-white">{analytics.items.dogs[0]?.name || '—'}</strong>
                  </div>
                </div>
              </div>

              {/* Menu Engineering Scatter Matrix + Top/Bottom Comparison Charts */}
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                {/* Scatter Plot Matrix */}
                <div className="p-5 rounded-xl bg-[#161922] border border-white/[0.07]">
                  <div className="mb-4">
                    <h2 className="text-sm font-semibold text-white">
                      Menu Engineering Matrix (Sales Volume vs. Total Revenue)
                    </h2>
                    <p className="text-xs text-gray-400">
                      Dashed lines mark menu averages ({formatNumber(analytics.kpis.avgItemUnits, 1)} units & {formatCurrency(analytics.kpis.avgItemRev)} revenue)
                    </p>
                  </div>
                  <div className="h-80 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <ScatterChart margin={{ top: 15, right: 20, bottom: 15, left: 5 }}>
                        <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.06)" />
                        <XAxis
                          type="number"
                          dataKey="unitsSold"
                          name="Units Sold"
                          stroke="#9CA3AF"
                          fontSize={11}
                        />
                        <YAxis
                          type="number"
                          dataKey="revenue"
                          name="Revenue"
                          stroke="#9CA3AF"
                          fontSize={11}
                          tickFormatter={formatCompactCurrency}
                        />
                        <ZAxis type="number" dataKey="avgPrice" range={[60, 260]} name="Unit Price" />
                        <Tooltip
                          cursor={{ strokeDasharray: '3 3' }}
                          content={({ active, payload }) => {
                            if (!active || !payload || !payload.length) return null;
                            const d = payload[0].payload;
                            return (
                              <div className="bg-[#11131A] border border-white/10 rounded-lg p-3 shadow-xl text-xs space-y-1">
                                <div className="font-semibold text-white">{d.name}</div>
                                <div className="text-emerald-400 font-medium">{d.quadrantLabel}</div>
                                <div className="text-gray-300 font-mono">Revenue: {formatCurrency(d.revenue)}</div>
                                <div className="text-gray-300 font-mono">Volume: {formatNumber(d.unitsSold)} units</div>
                                <div className="text-gray-300 font-mono">Avg Price: {formatCurrency(d.avgPrice)}</div>
                              </div>
                            );
                          }}
                        />
                        <ReferenceLine x={analytics.kpis.avgItemUnits} stroke="#6B7280" strokeDasharray="4 4" />
                        <ReferenceLine y={analytics.kpis.avgItemRev} stroke="#6B7280" strokeDasharray="4 4" />
                        <Scatter name="Stars" data={analytics.items.stars} fill="#10B981" />
                        <Scatter name="Plowhorses" data={analytics.items.plowhorses} fill="#F59E0B" />
                        <Scatter name="Puzzles" data={analytics.items.puzzles} fill="#6366F1" />
                        <Scatter name="Dogs" data={analytics.items.dogs} fill="#F43F5E" />
                        <Legend wrapperStyle={{ fontSize: '12px' }} />
                      </ScatterChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                {/* Top Selling by Units vs Revenue Comparison Bar Chart */}
                <div className="p-5 rounded-xl bg-[#161922] border border-white/[0.07]">
                  <div className="mb-4">
                    <h2 className="text-sm font-semibold text-white">
                      Best-Selling Items by Volume vs. Revenue Contribution
                    </h2>
                    <p className="text-xs text-gray-400">
                      Identifies high-volume items with lower revenue yield (Plowhorses) vs. high-revenue leaders
                    </p>
                  </div>
                  <div className="h-80 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <ComposedChart
                        data={analytics.items.byUnits.slice(0, 10)}
                        layout="vertical"
                        margin={{ top: 5, right: 20, left: 30, bottom: 5 }}
                      >
                        <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.06)" horizontal={false} />
                        <XAxis type="number" stroke="#9CA3AF" fontSize={11} />
                        <YAxis type="category" dataKey="name" stroke="#D1D5DB" fontSize={10} width={135} />
                        <Tooltip content={<CustomChartTooltip />} />
                        <Legend wrapperStyle={{ fontSize: '12px' }} />
                        <Bar dataKey="unitsSold" name="Units Sold" fill="#6366F1" radius={[0, 4, 4, 0]} barSize={14} />
                      </ComposedChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              </div>

              {/* Filterable & Sortable Complete Menu Intelligence Table */}
              <div className="p-5 rounded-xl bg-[#161922] border border-white/[0.07] space-y-4">
                <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-3">
                  <div className="flex flex-wrap items-center gap-1.5">
                    {[
                      { id: 'all', label: `All Items (${analytics.items.all.length})` },
                      { id: 'stars', label: `Prioritize / Stars (${analytics.items.stars.length})` },
                      { id: 'plowhorses', label: `High Sales, Low Rev (${analytics.items.plowhorses.length})` },
                      { id: 'puzzles', label: `High Rev, Low Sales (${analytics.items.puzzles.length})` },
                      { id: 'dogs', label: `Lowest Performing (${analytics.items.dogs.length})` },
                    ].map((t) => (
                      <button
                        key={t.id}
                        onClick={() => setMenuFilterTab(t.id)}
                        className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${
                          menuFilterTab === t.id
                            ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/30'
                            : 'bg-[#11131A] text-gray-400 hover:text-gray-200 border border-white/[0.06]'
                        }`}
                      >
                        {t.label}
                      </button>
                    ))}
                  </div>

                  <div className="flex flex-wrap items-center gap-2">
                    <div className="relative">
                      <Search className="w-3.5 h-3.5 text-gray-500 absolute left-3 top-1/2 -translate-y-1/2" />
                      <input
                        type="text"
                        value={menuSearch}
                        onChange={(e) => setMenuSearch(e.target.value)}
                        placeholder="Search menu item..."
                        className="pl-8 pr-3 py-1.5 bg-[#11131A] border border-white/10 rounded-lg text-xs text-white placeholder-gray-500 focus:outline-none focus:border-emerald-500/50"
                      />
                    </div>
                    <select
                      value={menuSortField}
                      onChange={(e) => setMenuSortField(e.target.value)}
                      className="px-3 py-1.5 bg-[#11131A] border border-white/10 rounded-lg text-xs text-gray-200 focus:outline-none"
                    >
                      <option value="revenue">Sort: Highest Revenue</option>
                      <option value="revenue_asc">Sort: Lowest Revenue</option>
                      <option value="unitsSold">Sort: Best-Selling (Units)</option>
                      <option value="orderFrequency">Sort: Most Frequently Ordered</option>
                      <option value="avgPrice">Sort: Highest Unit Price</option>
                    </select>
                  </div>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="border-b border-white/[0.08] text-[11px] uppercase text-gray-400">
                        <th className="py-3 px-2">Rank</th>
                        <th className="py-3 px-2">Menu Item</th>
                        <th className="py-3 px-2">Strategic Quadrant</th>
                        <th className="py-3 px-2 text-right">Unit Price</th>
                        <th className="py-3 px-2 text-right">Units Sold</th>
                        <th className="py-3 px-2 text-right">Order Freq</th>
                        <th className="py-3 px-2 text-right">Total Revenue</th>
                        <th className="py-3 px-2 text-right">Rev Share</th>
                        <th className="py-3 px-2">Recommended Action</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-white/[0.05] text-xs">
                      {filteredMenuItems.map((item) => (
                        <tr key={item.name} className="hover:bg-white/[0.02]">
                          <td className="py-3 px-2 font-mono text-gray-500">#{item.revenueRank}</td>
                          <td className="py-3 px-2 font-medium text-white">{item.name}</td>
                          <td className="py-3 px-2">
                            <span
                              className={`inline-flex items-center px-2 py-0.5 rounded text-[11px] font-medium ${
                                item.quadrant === 'Star'
                                  ? 'bg-emerald-500/15 text-emerald-300 border border-emerald-500/30'
                                  : item.quadrant === 'Plowhorse'
                                  ? 'bg-amber-500/15 text-amber-300 border border-amber-500/30'
                                  : item.quadrant === 'Puzzle'
                                  ? 'bg-indigo-500/15 text-indigo-300 border border-indigo-500/30'
                                  : 'bg-rose-500/15 text-rose-300 border border-rose-500/30'
                              }`}
                            >
                              {item.quadrantLabel}
                            </span>
                          </td>
                          <td className="py-3 px-2 text-right font-mono text-gray-300 tabular-nums">
                            {formatCurrency(item.avgPrice)}
                          </td>
                          <td className="py-3 px-2 text-right font-mono text-gray-200 tabular-nums">
                            {formatNumber(item.unitsSold)}
                          </td>
                          <td className="py-3 px-2 text-right font-mono text-gray-300 tabular-nums">
                            {formatNumber(item.orderFrequency)}
                          </td>
                          <td className="py-3 px-2 text-right font-mono font-semibold text-white tabular-nums">
                            {formatCurrency(item.revenue)}
                          </td>
                          <td className="py-3 px-2 text-right font-mono text-emerald-400 tabular-nums">
                            {formatPercent(item.revenueShare)}
                          </td>
                          <td className="py-3 px-2 text-gray-400 max-w-xs">{item.recommendation}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* ================================================================
              SECTION 4: CUSTOMER INSIGHTS & PURCHASING BEHAVIOUR
             ================================================================ */}
          {activeTab === 'customers' && analytics && (
            <div className="space-y-6">
              {/* Customer Segment Summary Cards */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3.5">
                {analytics.customers.segments.map((seg) => (
                  <div
                    key={seg.segment}
                    onClick={() =>
                      setCustomerSegmentFilter(customerSegmentFilter === seg.segment ? 'ALL' : seg.segment)
                    }
                    className={`p-4 rounded-xl border cursor-pointer transition-all ${
                      customerSegmentFilter === seg.segment
                        ? 'bg-emerald-500/15 border-emerald-500/40'
                        : 'bg-[#161922] border-white/[0.07] hover:border-white/20'
                    }`}
                  >
                    <div className="flex items-center justify-between text-xs text-gray-400">
                      <span className="font-medium text-gray-200 truncate">{seg.segment}</span>
                      <span className="font-mono text-emerald-400">{seg.count}</span>
                    </div>
                    <div className="mt-2 text-lg font-bold text-white font-mono tabular-nums">
                      {formatCurrency(seg.revenue, 0)}
                    </div>
                    <div className="mt-1 flex items-center justify-between text-[11px] text-gray-400">
                      <span>{formatPercent(seg.shareOfRevenue)} of rev</span>
                      <span>{formatPercent(seg.shareOfCustomers)} of guests</span>
                    </div>
                  </div>
                ))}
              </div>

              {/* Customer Scatter Chart & Co-Ordering Item Affinity */}
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                {/* Customer Spending vs Frequency Scatter */}
                <div className="lg:col-span-2 p-5 rounded-xl bg-[#161922] border border-white/[0.07]">
                  <div className="mb-4">
                    <h2 className="text-sm font-semibold text-white">
                      Customer Purchasing Matrix (Total Units Ordered vs. Total Revenue Spent)
                    </h2>
                    <p className="text-xs text-gray-400">
                      Each point represents a customer ({formatPercent(analytics.kpis.top10PctCustShare)} of revenue comes from top 10% of patrons)
                    </p>
                  </div>
                  <div className="h-72 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <ScatterChart margin={{ top: 10, right: 20, bottom: 10, left: 0 }}>
                        <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.06)" />
                        <XAxis
                          type="number"
                          dataKey="totalUnits"
                          name="Total Items Ordered"
                          stroke="#9CA3AF"
                          fontSize={11}
                        />
                        <YAxis
                          type="number"
                          dataKey="totalRevenue"
                          name="Total Spend"
                          stroke="#9CA3AF"
                          fontSize={11}
                          tickFormatter={formatCompactCurrency}
                        />
                        <Tooltip
                          content={({ active, payload }) => {
                            if (!active || !payload || !payload.length) return null;
                            const c = payload[0].payload;
                            return (
                              <div className="bg-[#11131A] border border-white/10 rounded-lg p-3 shadow-xl text-xs space-y-1">
                                <div className="font-semibold text-white">{c.id}</div>
                                <div className="text-emerald-400">{c.segment}</div>
                                <div className="font-mono text-gray-200">Total Spend: {formatCurrency(c.totalRevenue)}</div>
                                <div className="font-mono text-gray-300">Items Ordered: {c.totalUnits}</div>
                                <div className="text-gray-400">Favorite: {c.top1Item} ({c.top1Freq}x)</div>
                              </div>
                            );
                          }}
                        />
                        <Scatter
                          name="Customers"
                          data={analytics.customers.bySpend.slice(0, 120)}
                          fill="#10B981"
                          fillOpacity={0.75}
                        />
                      </ScatterChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                {/* Co-Ordering Affinity & Preferred Combinations */}
                <div className="p-5 rounded-xl bg-[#161922] border border-white/[0.07] flex flex-col justify-between">
                  <div>
                    <h2 className="text-sm font-semibold text-white">Frequent Item Co-Ordering Patterns</h2>
                    <p className="text-xs text-gray-400 mb-4">
                      Items most frequently ordered together by the same customer (1st, 2nd & 3rd favorites)
                    </p>

                    {analytics.customers.topCoOrders.length > 0 ? (
                      <div className="space-y-2.5">
                        {analytics.customers.topCoOrders.map((pair, idx) => (
                          <div
                            key={pair.pair}
                            className="p-2.5 rounded-lg bg-[#11131A] border border-white/[0.06] flex items-center justify-between gap-2"
                          >
                            <div className="flex items-center gap-2 min-w-0">
                              <span className="w-5 h-5 rounded bg-indigo-500/20 text-indigo-300 text-[11px] font-mono flex items-center justify-center shrink-0">
                                {idx + 1}
                              </span>
                              <span className="text-xs text-gray-200 truncate" title={pair.pair}>
                                {pair.pair}
                              </span>
                            </div>
                            <span className="text-xs font-mono text-emerald-400 shrink-0">
                              {pair.count} guests
                            </span>
                          </div>
                        ))}
                      </div>
                    ) : (
                      <div className="text-xs text-gray-500 py-8 text-center">
                        Single-item customer records detected; upload multi-item customer histories to view pairing affinities.
                      </div>
                    )}
                  </div>

                  <div className="mt-4 pt-3 border-t border-white/[0.06] text-[11px] text-gray-400">
                    <strong className="text-gray-200">Bundle Tip:</strong> Pairing top co-ordered items into a featured combo meal lifts average check value.
                  </div>
                </div>
              </div>

              {/* Customer Leaderboard Table */}
              <div className="p-5 rounded-xl bg-[#161922] border border-white/[0.07] space-y-4">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <h2 className="text-sm font-semibold text-white">
                      Customer Valuation & Purchasing Leaderboard
                    </h2>
                    <p className="text-xs text-gray-400">
                      Showing 1st, 2nd, and 3rd preferred items, order frequency, and total revenue per customer
                    </p>
                  </div>

                  <div className="flex flex-wrap items-center gap-2">
                    <div className="relative">
                      <Search className="w-3.5 h-3.5 text-gray-500 absolute left-3 top-1/2 -translate-y-1/2" />
                      <input
                        type="text"
                        value={customerSearch}
                        onChange={(e) => setCustomerSearch(e.target.value)}
                        placeholder="Search Customer ID, Email, Item..."
                        className="pl-8 pr-3 py-1.5 bg-[#11131A] border border-white/10 rounded-lg text-xs text-white placeholder-gray-500 focus:outline-none focus:border-emerald-500/50"
                      />
                    </div>
                    <select
                      value={customerSortField}
                      onChange={(e) => setCustomerSortField(e.target.value)}
                      className="px-3 py-1.5 bg-[#11131A] border border-white/10 rounded-lg text-xs text-gray-200"
                    >
                      <option value="totalRevenue">Sort: Highest Total Spending</option>
                      <option value="totalUnits">Sort: Highest Order Frequency</option>
                      <option value="avgSpendPerUnit">Sort: Highest Avg Spend / Item</option>
                    </select>
                  </div>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="border-b border-white/[0.08] text-[11px] uppercase text-gray-400">
                        <th className="py-3 px-2">Rank</th>
                        <th className="py-3 px-2">Customer ID</th>
                        <th className="py-3 px-2">Segment</th>
                        <th className="py-3 px-2">Most Ordered (1st)</th>
                        <th className="py-3 px-2">2nd Most Ordered</th>
                        <th className="py-3 px-2">3rd Most Ordered</th>
                        <th className="py-3 px-2 text-right">Items Ordered</th>
                        <th className="py-3 px-2 text-right">Total Spend</th>
                        {analytics.timeSeries.hasDateData && (
                          <th className="py-3 px-2 text-right">Recency</th>
                        )}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-white/[0.05] text-xs">
                      {filteredCustomers.slice(0, 40).map((cust) => (
                        <tr key={cust.id} className="hover:bg-white/[0.02]">
                          <td className="py-3 px-2 font-mono text-gray-500">#{cust.spendRank}</td>
                          <td className="py-3 px-2">
                            <div className="font-medium text-white font-mono">{cust.id}</div>
                            {cust.email !== '—' && (
                              <div className="text-[11px] text-gray-500 truncate max-w-[170px]">{cust.email}</div>
                            )}
                          </td>
                          <td className="py-3 px-2">
                            <span
                              className={`px-2 py-0.5 rounded text-[10px] font-medium ${
                                cust.segment === 'VIP Champions'
                                  ? 'bg-emerald-500/15 text-emerald-300 border border-emerald-500/30'
                                  : cust.segment === 'At-Risk / Declining'
                                  ? 'bg-rose-500/15 text-rose-300 border border-rose-500/30'
                                  : cust.segment === 'High-Frequency Loyalists'
                                  ? 'bg-indigo-500/15 text-indigo-300 border border-indigo-500/30'
                                  : 'bg-white/[0.06] text-gray-300'
                              }`}
                            >
                              {cust.segment}
                            </span>
                          </td>
                          <td className="py-3 px-2 text-gray-200">
                            {cust.top1Item !== '—' ? (
                              <span>
                                {cust.top1Item}{' '}
                                <span className="font-mono text-emerald-400">({cust.top1Freq}x)</span>
                              </span>
                            ) : (
                              '—'
                            )}
                          </td>
                          <td className="py-3 px-2 text-gray-300">
                            {cust.top2Item !== '—' ? (
                              <span>
                                {cust.top2Item}{' '}
                                <span className="font-mono text-gray-400">({cust.top2Freq}x)</span>
                              </span>
                            ) : (
                              '—'
                            )}
                          </td>
                          <td className="py-3 px-2 text-gray-400">
                            {cust.top3Item !== '—' ? (
                              <span>
                                {cust.top3Item}{' '}
                                <span className="font-mono text-gray-500">({cust.top3Freq}x)</span>
                              </span>
                            ) : (
                              '—'
                            )}
                          </td>
                          <td className="py-3 px-2 text-right font-mono text-gray-200 tabular-nums">
                            {formatNumber(cust.totalUnits)}
                          </td>
                          <td className="py-3 px-2 text-right font-mono font-semibold text-white tabular-nums">
                            {formatCurrency(cust.totalRevenue)}
                          </td>
                          {analytics.timeSeries.hasDateData && (
                            <td className="py-3 px-2 text-right font-mono text-gray-400 tabular-nums">
                              {cust.daysSinceLastOrder !== null ? `${cust.daysSinceLastOrder}d ago` : '—'}
                            </td>
                          )}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* ================================================================
              SECTION 5: AI ANALYST (BI ENGINE + CONVERSATIONAL CHATBOT)
             ================================================================ */}
          {activeTab === 'ai' && analytics && (
            <div className="grid grid-cols-1 xl:grid-cols-12 gap-6 items-start">
              {/* Left 7 Columns: Structured 7-Point Business Intelligence Report */}
              <div className="xl:col-span-7 space-y-4">
                <div className="p-5 rounded-xl bg-[#161922] border border-white/[0.07] flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                  <div>
                    <div className="flex items-center gap-2">
                      <Sparkles className="w-4 h-4 text-emerald-400" />
                      <h2 className="text-sm font-semibold text-white">
                        AI Business Intelligence & Strategic Recommendations
                      </h2>
                    </div>
                    <p className="text-xs text-gray-400 mt-0.5">
                      Automatically synthesized from your uploaded records ({analytics.meta.datasetName})
                    </p>
                  </div>
                  <button
                    onClick={handleGenerateAiExecutiveBrief}
                    disabled={isGeneratingAiBrief}
                    className="px-3 py-1.5 rounded-lg text-xs font-medium bg-indigo-500 hover:bg-indigo-400 disabled:opacity-50 text-white flex items-center gap-1.5 shrink-0 transition-colors"
                  >
                    <RefreshCw className={`w-3.5 h-3.5 ${isGeneratingAiBrief ? 'animate-spin' : ''}`} />
                    {isGeneratingAiBrief ? 'Synthesizing...' : 'Generate AI Deep-Dive'}
                  </button>
                </div>

                {aiEnhancedSummary && (
                  <div className="p-5 rounded-xl bg-indigo-950/30 border border-indigo-500/30 space-y-2">
                    <div className="flex items-center justify-between">
                      <span className="text-xs font-semibold text-indigo-300 flex items-center gap-1.5">
                        <Bot className="w-4 h-4" /> Gemini AI Executive Strategy Brief
                      </span>
                      <button
                        onClick={() => setAiEnhancedSummary(null)}
                        className="text-xs text-gray-400 hover:text-white"
                      >
                        Dismiss
                      </button>
                    </div>
                    <FormattedText text={aiEnhancedSummary} />
                  </div>
                )}

                <div className="space-y-3.5">
                  {biReportCards.map((card) => (
                    <div
                      key={card.id}
                      className="p-5 rounded-xl bg-[#161922] border border-white/[0.07] space-y-3"
                    >
                      <div className="flex items-center justify-between gap-2">
                        <span className="text-xs font-semibold text-emerald-400">{card.question}</span>
                        <span className="px-2 py-0.5 rounded text-[10px] font-mono uppercase bg-white/[0.06] text-gray-300">
                          {card.badge}
                        </span>
                      </div>
                      <h3 className="text-sm font-bold text-white">{card.headline}</h3>
                      <p className="text-xs text-gray-300 leading-relaxed">{card.analysis}</p>
                      {card.metrics && card.metrics.length > 0 && (
                        <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5 pt-2 border-t border-white/[0.06]">
                          {card.metrics.map((m, i) => (
                            <div key={i} className="p-2.5 rounded-lg bg-[#11131A] border border-white/[0.05]">
                              <div className="text-[10px] text-gray-400 truncate">{m.label}</div>
                              <div className="text-xs font-mono font-semibold text-white mt-0.5 truncate">
                                {m.value}
                              </div>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              </div>

              {/* Right 5 Columns: Interactive Restaurant AI Analyst Chatbot */}
              <div className="xl:col-span-5 xl:sticky xl:top-20 rounded-xl bg-[#161922] border border-white/[0.08] flex flex-col h-[680px] shadow-2xl">
                {/* Chatbot Header */}
                <div className="p-4 border-b border-white/[0.07] flex items-center justify-between">
                  <div className="flex items-center gap-2.5">
                    <div className="w-8 h-8 rounded-lg bg-emerald-500/15 border border-emerald-500/30 flex items-center justify-center text-emerald-400">
                      <Bot className="w-4 h-4" />
                    </div>
                    <div>
                      <div className="text-sm font-semibold text-white flex items-center gap-2">
                        Restaurant AI Analyst
                        <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse" />
                      </div>
                      <div className="text-[11px] text-gray-400">
                        Grounded strictly in {analytics.meta.datasetName}
                      </div>
                    </div>
                  </div>
                  <button
                    onClick={() =>
                      setChatMessages([
                        {
                          role: 'assistant',
                          content: `Chat reset. I have your dataset (**${analytics.meta.datasetName}**) ready. Ask me any question below!`,
                          timestamp: 'Just now',
                        },
                      ])
                    }
                    className="text-xs text-gray-400 hover:text-white px-2 py-1 rounded bg-white/[0.04]"
                  >
                    Clear
                  </button>
                </div>

                {/* Suggested Prompt Chips */}
                <div className="p-3 bg-[#11131A] border-b border-white/[0.06]">
                  <div className="text-[10px] font-mono uppercase tracking-wider text-gray-500 mb-2">
                    Quick Analytical Questions (Click to Ask)
                  </div>
                  <div className="flex gap-1.5 overflow-x-auto pb-1">
                    {suggestedQuestions.map((q) => (
                      <button
                        key={q}
                        onClick={() => handleSendChat(q)}
                        disabled={isChatLoading}
                        className="px-2.5 py-1 rounded-full text-[11px] bg-[#161922] hover:bg-emerald-500/15 text-gray-300 hover:text-emerald-300 border border-white/10 hover:border-emerald-500/30 whitespace-nowrap transition-colors shrink-0"
                      >
                        {q}
                      </button>
                    ))}
                  </div>
                </div>

                {/* Chat Message Stream */}
                <div className="flex-1 p-4 overflow-y-auto space-y-4">
                  {chatMessages.map((msg, idx) => (
                    <div
                      key={idx}
                      className={`flex items-start gap-2.5 ${
                        msg.role === 'user' ? 'flex-row-reverse' : ''
                      }`}
                    >
                      <div
                        className={`w-7 h-7 rounded-lg flex items-center justify-center shrink-0 ${
                          msg.role === 'user'
                            ? 'bg-indigo-500/20 text-indigo-300 border border-indigo-500/30'
                            : 'bg-emerald-500/15 text-emerald-400 border border-emerald-500/30'
                        }`}
                      >
                        {msg.role === 'user' ? <User className="w-3.5 h-3.5" /> : <Bot className="w-3.5 h-3.5" />}
                      </div>
                      <div
                        className={`max-w-[85%] rounded-xl p-3.5 ${
                          msg.role === 'user'
                            ? 'bg-indigo-600/25 border border-indigo-500/30 text-white'
                            : 'bg-[#11131A] border border-white/[0.07] text-gray-200'
                        }`}
                      >
                        <FormattedText text={msg.content} />
                        <div className="text-[10px] text-gray-500 mt-1.5 text-right">{msg.timestamp}</div>
                      </div>
                    </div>
                  ))}
                  {isChatLoading && (
                    <div className="flex items-center gap-2 text-xs text-gray-400 p-2">
                      <RefreshCw className="w-3.5 h-3.5 animate-spin text-emerald-400" />
                      Calculating exact metrics from active dataset...
                    </div>
                  )}
                  <div ref={chatEndRef} />
                </div>

                {/* Chat Input Bar */}
                <form
                  onSubmit={(e) => {
                    e.preventDefault();
                    handleSendChat();
                  }}
                  className="p-3 border-t border-white/[0.07] bg-[#11131A] flex items-center gap-2 rounded-b-xl"
                >
                  <input
                    type="text"
                    value={chatInput}
                    onChange={(e) => setChatInput(e.target.value)}
                    placeholder="Ask about menu items, top customers, revenue share..."
                    className="flex-1 bg-[#161922] border border-white/10 rounded-lg px-3.5 py-2 text-xs text-white placeholder-gray-500 focus:outline-none focus:border-emerald-500/50"
                  />
                  <button
                    type="submit"
                    disabled={!chatInput.trim() || isChatLoading}
                    className="px-3.5 py-2 rounded-lg bg-emerald-500 hover:bg-emerald-400 disabled:opacity-40 text-slate-950 font-semibold text-xs flex items-center gap-1.5 transition-colors"
                  >
                    <Send className="w-3.5 h-3.5" /> Ask
                  </button>
                </form>
              </div>
            </div>
          )}

          {/* ================================================================
              SECTION 6: UPLOAD DATA, COLUMN MAPPING & DATA PREVIEW
             ================================================================ */}
          {activeTab === 'upload' && (
            <div className="space-y-6">
              {/* Upload Zone + Demo & Template Loaders */}
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                {/* Drag & Drop File Upload Box */}
                <div className="lg:col-span-2 p-6 rounded-xl bg-[#161922] border border-white/[0.07] flex flex-col justify-between">
                  <div>
                    <h2 className="text-base font-semibold text-white">
                      Upload Restaurant Sales Dataset (CSV, XLSX, XLS)
                    </h2>
                    <p className="text-xs text-gray-400 mt-1">
                      Supports both <strong>Customer-Summary</strong> datasets (1st/2nd/3rd Most Ordered Items) and <strong>Transaction-Level</strong> logs. Columns are automatically detected and mapped.
                    </p>
                  </div>

                  <label
                    onDragOver={(e) => e.preventDefault()}
                    onDrop={(e) => {
                      e.preventDefault();
                      const file = e.dataTransfer.files?.[0];
                      if (file) handleFileUpload(file);
                    }}
                    className="my-5 border-2 border-dashed border-white/15 hover:border-emerald-500/50 rounded-xl p-8 flex flex-col items-center justify-center text-center cursor-pointer bg-[#11131A]/60 hover:bg-emerald-500/[0.03] transition-all"
                  >
                    <div className="w-12 h-12 rounded-full bg-emerald-500/15 border border-emerald-500/30 flex items-center justify-center text-emerald-400 mb-3">
                      <UploadCloud className="w-6 h-6" />
                    </div>
                    <div className="text-sm font-medium text-white">
                      Click to browse or drag & drop your sales file here
                    </div>
                    <div className="text-xs text-gray-400 mt-1">
                      Supports .CSV, .XLSX, and .XLS files covering up to 6+ months of restaurant sales
                    </div>
                    <input
                      type="file"
                      accept=".csv,.xlsx,.xls"
                      onChange={(e) => handleFileUpload(e.target.files?.[0])}
                      className="hidden"
                    />
                  </label>

                  {uploadError && (
                    <div className="p-3 rounded-lg bg-rose-500/15 border border-rose-500/30 text-xs text-rose-300 flex items-center gap-2">
                      <AlertTriangle className="w-4 h-4 shrink-0" />
                      {uploadError}
                    </div>
                  )}

                  <div className="flex flex-wrap items-center justify-between gap-2 pt-3 border-t border-white/[0.06] text-xs text-gray-400">
                    <span>Need an Indian restaurant template file to test uploading?</span>
                    <div className="flex items-center gap-3">
                      <button
                        onClick={() => downloadSampleCSV('customer_summary')}
                        className="text-emerald-400 hover:underline flex items-center gap-1"
                      >
                        <Download className="w-3.5 h-3.5" /> Indian Customer-Summary CSV
                      </button>
                      <button
                        onClick={() => downloadSampleCSV('transaction')}
                        className="text-indigo-400 hover:underline flex items-center gap-1"
                      >
                        <Download className="w-3.5 h-3.5" /> 6M Indian Transactions CSV
                      </button>
                    </div>
                  </div>
                </div>

                {/* Regional Settings, Currency, GST & Instant Sample Demo Switcher */}
                <div className="p-6 rounded-xl bg-[#161922] border border-white/[0.07] flex flex-col justify-between space-y-4">
                  <div>
                    <span className="px-2 py-0.5 text-[10px] font-mono uppercase rounded bg-emerald-500/15 text-emerald-300 border border-emerald-500/30">
                      Indian Regional & Tax Settings
                    </span>
                    <h3 className="text-sm font-semibold text-white mt-2">
                      Currency, GST & Sample Datasets
                    </h3>
                    <p className="text-xs text-gray-400 mt-1">
                      Uploaded numeric values are never blindly converted. Choose your display currency and optional GST treatment below:
                    </p>
                  </div>

                  {/* Currency & GST Controls inside Upload Page */}
                  <div className="space-y-2.5 pt-1">
                    <div>
                      <label className="block text-[11px] text-gray-400 mb-1">
                        Dataset Currency (Default: INR ₹ · en-IN)
                      </label>
                      <select
                        value={currencyCode}
                        onChange={(e) => setCurrencyCode(e.target.value)}
                        className="w-full px-3 py-1.5 rounded-lg bg-[#11131A] border border-white/10 text-xs text-white focus:outline-none focus:border-emerald-500"
                      >
                        {Object.values(CURRENCY_OPTIONS).map((c) => (
                          <option key={c.code} value={c.code}>
                            {c.label}
                          </option>
                        ))}
                      </select>
                    </div>

                    <div>
                      <label className="block text-[11px] text-gray-400 mb-1">
                        GST Treatment (Optional — Never Assumed)
                      </label>
                      <select
                        value={gstMode}
                        onChange={(e) => setGstMode(e.target.value)}
                        className="w-full px-3 py-1.5 rounded-lg bg-[#11131A] border border-white/10 text-xs text-white focus:outline-none focus:border-emerald-500"
                      >
                        {GST_OPTIONS.map((opt) => (
                          <option key={opt.id} value={opt.id}>
                            {opt.label}
                          </option>
                        ))}
                      </select>
                    </div>
                  </div>

                  <div className="space-y-2.5 pt-2 border-t border-white/[0.06]">
                    <button
                      onClick={() => loadDemoDataset('customer_summary', true)}
                      className="w-full p-3 rounded-xl bg-[#11131A] hover:bg-white/[0.04] border border-white/10 text-left transition-all group"
                    >
                      <div className="flex items-center justify-between">
                        <span className="text-xs font-semibold text-emerald-400 group-hover:underline">
                          1. Load Indian Customer-Summary Demo
                        </span>
                        <ArrowRight className="w-3.5 h-3.5 text-gray-400 group-hover:translate-x-0.5 transition-transform" />
                      </div>
                      <p className="text-[11px] text-gray-400 mt-1">
                        115 diners with 1st/2nd/3rd Most Ordered Indian Dishes (Chicken Biryani, Paneer Butter Masala, Masala Dosa, etc.) in ₹.
                      </p>
                    </button>

                    <button
                      onClick={() => loadDemoDataset('transaction', true)}
                      className="w-full p-3 rounded-xl bg-[#11131A] hover:bg-white/[0.04] border border-white/10 text-left transition-all group"
                    >
                      <div className="flex items-center justify-between">
                        <span className="text-xs font-semibold text-indigo-400 group-hover:underline">
                          2. Load 6-Month Indian Transaction Demo
                        </span>
                        <ArrowRight className="w-3.5 h-3.5 text-gray-400 group-hover:translate-x-0.5 transition-transform" />
                      </div>
                      <p className="text-[11px] text-gray-400 mt-1">
                        6 months of dated Indian restaurant orders (01-05-2026 to 31-10-2026) in DD-MM-YYYY format and INR (₹).
                      </p>
                    </button>
                  </div>
                </div>
              </div>

              {/* Column Mapping & Schema Configuration Panel (When user uploads or inspects mapping) */}
              {(showMappingModal || rawHeaders.length > 0) && (
                <div className="p-6 rounded-xl bg-[#161922] border border-emerald-500/30 space-y-5">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-4 border-b border-white/[0.07]">
                    <div>
                      <div className="flex items-center gap-2">
                        <Sliders className="w-4 h-4 text-emerald-400" />
                        <h3 className="text-sm font-semibold text-white">
                          Column Mapping & Schema Configuration ({pendingMeta.datasetName})
                        </h3>
                      </div>
                      <p className="text-xs text-gray-400 mt-0.5">
                        We automatically matched your file headers. You can customize the schema type or adjust any column mapping below.
                      </p>
                    </div>

                    <div className="flex items-center gap-2">
                      <select
                        value={pendingSchemaType}
                        onChange={(e) => {
                          const newType = e.target.value;
                          setPendingSchemaType(newType);
                          const fields = newType === 'customer_summary' ? SUMMARY_SCHEMA_FIELDS : TRANSACTION_SCHEMA_FIELDS;
                          const remapped = {};
                          fields.forEach((f) => {
                            const match = rawHeaders.find((h) => f.aliases.includes(normalizeHeader(h)));
                            remapped[f.key] = match || '';
                          });
                          setPendingMapping(remapped);
                        }}
                        className="px-3 py-1.5 rounded-lg bg-[#11131A] border border-white/15 text-xs text-white"
                      >
                        <option value="customer_summary">Customer-Summary Schema</option>
                        <option value="transaction">Transaction-Level Schema</option>
                      </select>

                      <button
                        onClick={applyMappingAndAnalyze}
                        className="px-4 py-1.5 rounded-lg bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-semibold text-xs flex items-center gap-1.5"
                      >
                        <Check className="w-3.5 h-3.5" /> Confirm & Run Analytics
                      </button>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-3.5">
                    {(pendingSchemaType === 'customer_summary' ? SUMMARY_SCHEMA_FIELDS : TRANSACTION_SCHEMA_FIELDS).map(
                      (field) => (
                        <div key={field.key} className="p-3 rounded-lg bg-[#11131A] border border-white/[0.06]">
                          <label className="block text-[11px] font-medium text-gray-300 mb-1.5">
                            {field.label}{' '}
                            {field.required && <span className="text-emerald-400">*</span>}
                          </label>
                          <select
                            value={pendingMapping[field.key] || ''}
                            onChange={(e) =>
                              setPendingMapping((prev) => ({ ...prev, [field.key]: e.target.value }))
                            }
                            className="w-full px-2.5 py-1.5 rounded bg-[#161922] border border-white/10 text-xs text-white focus:outline-none focus:border-emerald-500"
                          >
                            <option value="">-- Not Mapped / Unavailable --</option>
                            {rawHeaders.map((h) => (
                              <option key={h} value={h}>
                                {h}
                              </option>
                            ))}
                          </select>
                        </div>
                      )
                    )}
                  </div>
                </div>
              )}

              {/* Data Quality & Hygiene Audit + Data Preview Table */}
              {analytics && (
                <div className="p-6 rounded-xl bg-[#161922] border border-white/[0.07] space-y-5">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                    <div>
                      <h3 className="text-sm font-semibold text-white">
                        Data Validation, Hygiene Log & Dataset Preview
                      </h3>
                      <p className="text-xs text-gray-400">
                        Showing validated records from <strong>{analytics.meta.datasetName}</strong>
                      </p>
                    </div>
                    <div className="flex flex-wrap items-center gap-3 text-xs">
                      <span className="px-2.5 py-1 rounded-md bg-emerald-500/10 text-emerald-300 border border-emerald-500/20 font-mono">
                        Valid Rows: {formatNumber(analytics.hygiene.validRowCount)}
                      </span>
                      <span className="px-2.5 py-1 rounded-md bg-white/[0.05] text-gray-300 font-mono">
                        Duplicates Removed: {analytics.hygiene.duplicatesRemoved}
                      </span>
                      <span className="px-2.5 py-1 rounded-md bg-white/[0.05] text-gray-300 font-mono">
                        Missing Values Handled: {analytics.hygiene.missingValuesHandled}
                      </span>
                    </div>
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse">
                      <thead>
                        <tr className="border-b border-white/[0.08] text-[11px] uppercase text-gray-400">
                          {rawHeaders.slice(0, 10).map((h) => (
                            <th key={h} className="py-2.5 px-2.5 whitespace-nowrap">
                              {h}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-white/[0.05] text-xs font-mono">
                        {analytics.rawPreview
                          .slice((previewPage - 1) * 10, previewPage * 10)
                          .map((row, rIdx) => (
                            <tr key={rIdx} className="hover:bg-white/[0.02]">
                              {rawHeaders.slice(0, 10).map((h) => (
                                <td key={h} className="py-2 px-2.5 text-gray-300 whitespace-nowrap">
                                  {String(row[h] !== undefined && row[h] !== '' ? row[h] : '—')}
                                </td>
                              ))}
                            </tr>
                          ))}
                      </tbody>
                    </table>
                  </div>

                  {/* Pagination Controls */}
                  <div className="flex items-center justify-between text-xs text-gray-400 pt-2">
                    <span>
                      Showing rows {(previewPage - 1) * 10 + 1}–
                      {Math.min(previewPage * 10, analytics.rawPreview.length)} of{' '}
                      {analytics.rawPreview.length} previewed records
                    </span>
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => setPreviewPage((p) => Math.max(1, p - 1))}
                        disabled={previewPage === 1}
                        className="px-2.5 py-1 rounded bg-[#11131A] border border-white/10 disabled:opacity-40 hover:text-white"
                      >
                        Previous
                      </button>
                      <button
                        onClick={() =>
                          setPreviewPage((p) =>
                            p * 10 < analytics.rawPreview.length ? p + 1 : p
                          )
                        }
                        disabled={previewPage * 10 >= analytics.rawPreview.length}
                        className="px-2.5 py-1 rounded bg-[#11131A] border border-white/10 disabled:opacity-40 hover:text-white"
                      >
                        Next
                      </button>
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}
        </main>
      </div>
    </div>
  );
}