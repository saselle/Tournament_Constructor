import React, { useState, useMemo, useEffect } from 'react';
import { createRoot } from 'react-dom/client';
import * as XLSX from 'xlsx';
import { Settings, Calendar, Download, Trophy, AlertCircle, Check, Sparkles } from 'lucide-react';

// ============ УТИЛИТЫ ============
const timeToMin = (s) => { const [h, m] = s.split(':').map(Number); return h * 60 + (m || 0); };
const minToTime = (m) => `${String(Math.floor(m / 60) % 24).padStart(2, '0')}:${String(Math.floor(m) % 60).padStart(2, '0')}`;
const nextPow2 = (n) => { let p = 1; while (p < n) p *= 2; return p; };
const COL = (n) => XLSX.utils.encode_col(n - 1); // 1-indexed -> буква

// ============ АЛГОРИТМЫ ============
const roundRobin = (n) => {
  const arr = Array.from({ length: n }, (_, i) => i + 1);
  if (n % 2 !== 0) arr.push(0);
  const k = arr.length;
  const rounds = [];
  let pos = [...arr];
  for (let r = 0; r < k - 1; r++) {
    const round = [];
    for (let i = 0; i < k / 2; i++) {
      const a = pos[i], b = pos[k - 1 - i];
      if (a !== 0 && b !== 0) round.push([a, b]);
    }
    rounds.push(round);
    pos = [pos[0], pos[k - 1], ...pos.slice(1, k - 1)];
  }
  return rounds;
};

const playoffSeeds = (bracketSize) => {
  let seeds = [[1, 2]];
  while (seeds.length * 2 < bracketSize) {
    const nextSize = seeds.length * 4;
    const ns = [];
    for (const [a, b] of seeds) {
      ns.push([a, nextSize + 1 - a]);
      ns.push([b, nextSize + 1 - b]);
    }
    seeds = ns;
  }
  return seeds;
};

const splitIntoGroups = (totalTeams, groupSize, order) => {
  // Последовательное распределение: команды 1..N в группу 1, N+1..2N в группу 2, и т.д.
  // Если totalTeams не делится нацело — экстра-команды попадают в ПОСЛЕДНИЕ группы
  // (нормальная практика для турниров: первые группы держат заявленный размер, в последних +1).
  // `order` (опционально) — результат жеребьёвки: массив sid в том порядке, в котором их
  // раскладывают по группам (см. generateDrawOrder). По умолчанию — просто 1..N.
  const seq = (order && order.length === totalTeams) ? order : Array.from({ length: totalTeams }, (_, i) => i + 1);
  const numGroups = Math.max(1, Math.floor(totalTeams / groupSize));
  const base = Math.floor(totalTeams / numGroups);
  const extra = totalTeams - base * numGroups; // 0 <= extra < numGroups
  const groups = [];
  let idx = 0;
  for (let g = 0; g < numGroups; g++) {
    const sz = g >= numGroups - extra ? base + 1 : base;
    const teams = [];
    for (let i = 0; i < sz; i++) teams.push(seq[idx++]);
    groups.push(teams);
  }
  return groups;
};

// Перемешивание Фишера-Йетса (не для токенов/ID — обычная жеребьёвка команд, Math.random достаточно).
const shuffleArr = (arr) => {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
};

// Строит порядок распределения sid по группам для splitIntoGroups согласно режиму жеребьёвки.
const generateDrawOrder = (mode, totalTeams, groupSize, numSeeds) => {
  const seq = Array.from({ length: totalTeams }, (_, i) => i + 1);
  if (mode === 'random') return shuffleArr(seq);
  if (mode === 'seeded') {
    const numGroups = Math.max(1, Math.floor(totalTeams / groupSize));
    const base = Math.floor(totalTeams / numGroups);
    const extra = totalTeams - base * numGroups;
    const seedCount = Math.max(0, Math.min(numSeeds || 0, numGroups, totalTeams));
    const seeds = seq.slice(0, seedCount);
    const rest = shuffleArr(seq.slice(seedCount));
    const order = [];
    let restIdx = 0;
    for (let g = 0; g < numGroups; g++) {
      const sz = g >= numGroups - extra ? base + 1 : base;
      const chunk = [];
      if (g < seeds.length) chunk.push(seeds[g]);
      while (chunk.length < sz && restIdx < rest.length) chunk.push(rest[restIdx++]);
      order.push(...chunk);
    }
    return order;
  }
  return seq; // 'sequential'
};

const countGames = (sys, totalTeams, groupSize, advance) => {
  if (sys === 'playoff') return totalTeams - 1 + 1; // +бронза
  if (sys === 'group') {
    const groups = splitIntoGroups(totalTeams, groupSize);
    return groups.reduce((s, g) => s + (g.length * (g.length - 1)) / 2, 0);
  }
  if (sys === 'mixed') {
    const groups = splitIntoGroups(totalTeams, groupSize);
    const gr = groups.reduce((s, g) => s + (g.length * (g.length - 1)) / 2, 0);
    const pt = groups.length * advance;
    return gr + (pt - 1) + 1;
  }
  return 0;
};

// ============ АВТОРЕКОМЕНДАЦИЯ ============
function scoreSystem(sys, totalTeams, matchDur, groupSize, advance) {
  let s = 0;
  if (matchDur < 20) s -= 100;
  else if (matchDur < 25) s -= 30;
  else if (matchDur < 35) s += 10;
  else if (matchDur <= 50) s += 30;
  else if (matchDur <= 70) s += 15;
  else s -= 10;
  if (sys === 'mixed' && totalTeams >= 16) s += 25;
  if (sys === 'group' && totalTeams >= 32) s -= 15;
  if (sys === 'playoff' && totalTeams >= 24) s -= 10;
  if (sys === 'playoff' && totalTeams <= 8) s += 15;
  if (groupSize && totalTeams % groupSize === 0) s += 8;
  if (sys === 'mixed' && advance === 2) s += 5;
  return s;
}

const recommend = (totalTeams, fields, minutesAvailable) => {
  const cand = [];
  const poGames = totalTeams - 1 + 1;
  const poSlots = Math.ceil(poGames / fields);
  cand.push({ system: 'playoff', groupSize: 0, advance: 0, slots: poSlots, matchDur: minutesAvailable / poSlots, score: scoreSystem('playoff', totalTeams, minutesAvailable / poSlots) });
  for (const gs of [3, 4, 5, 6]) {
    if (totalTeams < gs * 2) continue;
    if (Math.floor(totalTeams / gs) < 2) continue;
    const g = countGames('group', totalTeams, gs, 0);
    const sl = Math.ceil(g / fields);
    cand.push({ system: 'group', groupSize: gs, advance: 0, slots: sl, matchDur: minutesAvailable / sl, score: scoreSystem('group', totalTeams, minutesAvailable / sl, gs) });
  }
  for (const gs of [3, 4, 5, 6]) {
    if (totalTeams < gs * 2) continue;
    const ng = Math.floor(totalTeams / gs);
    if (ng < 2) continue;
    for (const adv of [1, 2, 3]) {
      if (adv >= gs) continue;
      const pt = ng * adv;
      if (pt < 2 || pt > 128) continue;
      const g = countGames('mixed', totalTeams, gs, adv);
      const sl = Math.ceil(g / fields);
      cand.push({ system: 'mixed', groupSize: gs, advance: adv, slots: sl, matchDur: minutesAvailable / sl, score: scoreSystem('mixed', totalTeams, minutesAvailable / sl, gs, adv) });
    }
  }
  cand.sort((a, b) => b.score - a.score);
  return cand[0];
};

// ============ ПОСТРОЕНИЕ СЕТКИ ПЛЕЙ-ОФФ ============
const ROUND_NAMES = ['Финал', '1/2', '1/4', '1/8', '1/16', '1/32', '1/64'];

// "Утешительная" под-сетка мест из проигравших конкретных матчей — без BYE, т.к.
// количество лузеров каждого раунда всегда степень двойки. rankOffset — лучшее
// место, которое может занять победитель этой под-сетки (проигравший — rankOffset+1).
// Рекурсивно продолжает вглубь: полуфинальные лузеры этой же под-сетки определяют
// следующую пару мест (например, из 5-8 получаем 5-6 и 7-8).
const buildPlacementBracket = (loserMatchIds, startId, rankOffset, bracket) => {
  const matches = [];
  let id = startId;
  const bracketSize = loserMatchIds.length;
  const numRounds = Math.log2(bracketSize);
  const isSingleMatch = numRounds === 1;
  const finalLabel = `За ${rankOffset}-${rankOffset + 1} место`;

  const r1Start = id;
  for (let i = 0; i < bracketSize / 2; i++) {
    const a = loserMatchIds[i * 2], b = loserMatchIds[i * 2 + 1];
    matches.push({
      id: id++, phase: 'playoff', round: 1, bracket,
      roundName: isSingleMatch ? finalLabel : (ROUND_NAMES[numRounds - 1] || 'Раунд 1'),
      loseFrom: [a, b],
      t1: `Проигр.М${a}`, t2: `Проигр.М${b}`,
      label: isSingleMatch ? finalLabel : `${ROUND_NAMES[numRounds - 1] || 'Р1'} (места) м.${i + 1}`,
    });
  }
  const roundStarts = [r1Start];
  const roundCounts = [bracketSize / 2];
  let prevStart = r1Start, prevCount = bracketSize / 2;
  for (let r = 2; r <= numRounds; r++) {
    const cnt = prevCount / 2;
    const rStart = id;
    const isFinal = r === numRounds;
    const rName = isFinal ? finalLabel : (ROUND_NAMES[numRounds - r] || `Раунд ${r}`);
    for (let i = 0; i < cnt; i++) {
      matches.push({
        id: id++, phase: 'playoff', round: r, bracket,
        roundName: rName,
        winFrom: [prevStart + i * 2, prevStart + i * 2 + 1],
        t1: `Поб.М${prevStart + i * 2}`, t2: `Поб.М${prevStart + i * 2 + 1}`,
        label: isFinal ? rName : `${rName} (места) м.${i + 1}`,
      });
    }
    roundStarts.push(rStart); roundCounts.push(cnt);
    prevStart = rStart; prevCount = cnt;
  }

  let nextId = id;
  for (let r = 1; r <= numRounds - 1; r++) {
    const cnt2 = roundCounts[r - 1];
    if (cnt2 < 2) continue;
    const ids2 = Array.from({ length: cnt2 }, (_, i) => roundStarts[r - 1] + i);
    const subRankOffset = rankOffset + Math.pow(2, numRounds - r);
    const sub = buildPlacementBracket(ids2, nextId, subRankOffset, bracket);
    matches.push(...sub.matches);
    nextId = sub.nextId;
  }
  return { matches, nextId };
};

// Основная сетка на вылет (seed-based первый раунд, с BYE если playoffTeams — не степень
// двойки). fullPlacement=false — только бронза (как раньше); true — расписывает ВСЕ места
// через buildPlacementBracket для лузеров каждого раунда, кроме финала.
const buildBracketMatches = (startId, playoffTeams, opts) => {
  const { bracket = null, fullPlacement = false } = opts || {};
  const matches = [];
  let id = startId;
  const bracketSize = nextPow2(playoffTeams);
  const numRounds = Math.log2(bracketSize);
  const seeds = playoffSeeds(bracketSize);

  const r1Start = id;
  seeds.forEach(([a, b], i) => {
    matches.push({
      id: id++, phase: 'playoff', round: 1, bracket,
      roundName: ROUND_NAMES[numRounds - 1] || 'Раунд 1',
      seedA: a, seedB: b, playoffTeams, bracketSize,
      t1: a > playoffTeams ? 'BYE' : `СИД${a}`,
      t2: b > playoffTeams ? 'BYE' : `СИД${b}`,
      label: `${ROUND_NAMES[numRounds - 1] || 'Р1'} м.${i + 1}`,
      prevWin: null,
    });
  });

  const roundStarts = [r1Start];
  const roundCounts = [seeds.length];
  let prevStart = r1Start, prevCount = seeds.length;
  for (let r = 2; r <= numRounds; r++) {
    const cnt = prevCount / 2;
    const rStart = id;
    for (let i = 0; i < cnt; i++) {
      matches.push({
        id: id++, phase: 'playoff', round: r, bracket,
        roundName: ROUND_NAMES[numRounds - r] || `Раунд ${r}`,
        winFrom: [prevStart + i * 2, prevStart + i * 2 + 1],
        t1: `Поб.М${prevStart + i * 2}`, t2: `Поб.М${prevStart + i * 2 + 1}`,
        label: `${ROUND_NAMES[numRounds - r] || `Р${r}`} м.${i + 1}`,
      });
    }
    roundStarts.push(rStart); roundCounts.push(cnt);
    prevStart = rStart; prevCount = cnt;
  }

  let nextId = id;
  if (numRounds >= 2) {
    if (!fullPlacement) {
      // только бронза — полуфиналы это предпоследний раунд
      const sfStart = roundStarts[numRounds - 2];
      matches.push({
        id: nextId++, phase: 'playoff', round: numRounds, isBronze: true, bracket,
        roundName: 'Бронза',
        loseFrom: [sfStart, sfStart + 1],
        t1: `Проигр.М${sfStart}`, t2: `Проигр.М${sfStart + 1}`,
        label: 'За 3-е место',
      });
    } else {
      for (let r = 1; r <= numRounds - 1; r++) {
        const cnt = roundCounts[r - 1];
        const ids = Array.from({ length: cnt }, (_, i) => roundStarts[r - 1] + i);
        const rankOffset = cnt + 1;
        const sub = buildPlacementBracket(ids, nextId, rankOffset, bracket);
        matches.push(...sub.matches);
        nextId = sub.nextId;
      }
    }
  }
  return { matches, nextId };
};

// ============ ПОСТРОЕНИЕ МАТЧЕЙ ============
// Для mixed/playoff помечаем плей-офф матчи seed-метками для связи формулами
const buildMatches = (params) => {
  const { totalTeams, system, groupSize, advance, drawOrder } = params;
  const matches = [];
  let id = 1;

  if (system === 'group' || system === 'mixed' || system === 'mixed-full' || system === 'mixed-goldsilver' || system === 'mixed-goldsilver-full') {
    const groups = splitIntoGroups(totalTeams, groupSize, drawOrder);
    groups.forEach((teams, gIdx) => {
      const rounds = roundRobin(teams.length);
      rounds.forEach((round, rIdx) => {
        round.forEach(([a, b]) => {
          matches.push({
            id: id++, phase: 'group', group: gIdx + 1, round: rIdx + 1,
            t1: `G${gIdx + 1}.${a}`, t2: `G${gIdx + 1}.${b}`,
            label: `Гр.${gIdx + 1} т.${rIdx + 1}`,
          });
        });
      });
    });
  }

  const numGroups = (system === 'group' || system === 'mixed' || system === 'mixed-full' || system === 'mixed-goldsilver' || system === 'mixed-goldsilver-full')
    ? splitIntoGroups(totalTeams, groupSize, drawOrder).length : 0;

  if (system === 'playoff' || system === 'playoff-full') {
    const res = buildBracketMatches(id, totalTeams, { fullPlacement: system === 'playoff-full' });
    matches.push(...res.matches);
    id = res.nextId;
  } else if (system === 'mixed' || system === 'mixed-full') {
    const playoffTeams = numGroups * advance;
    const res = buildBracketMatches(id, playoffTeams, { fullPlacement: system === 'mixed-full' });
    matches.push(...res.matches);
    id = res.nextId;
  } else if (system === 'mixed-goldsilver' || system === 'mixed-goldsilver-full') {
    // 1-е места групп -> золотой плей-офф, 2-е места -> серебряный. Каждая сетка независима.
    const full = system === 'mixed-goldsilver-full';
    const gold = buildBracketMatches(id, numGroups, { bracket: 'gold', fullPlacement: full });
    matches.push(...gold.matches);
    id = gold.nextId;
    const silver = buildBracketMatches(id, numGroups, { bracket: 'silver', fullPlacement: full });
    matches.push(...silver.matches);
    id = silver.nextId;
  }
  return matches;
};

// ============ РАСПИСАНИЕ (со слотами, днями) ============
const scheduleMatches = (matches, fields, params) => {
  // Жадно раскидываем без конфликтов команд в одном слоте.
  // Плей-офф матчи зависят от групп — ставим строго после всех групповых.
  // minRestSlots — минимальный зазор между двумя матчами одной команды в слотах.
  // blockedSlotIdxs — набор глобальных индексов слотов, которые нельзя использовать
  // (обеденный перерыв, церемония и т.п.).
  const schedule = [];
  const slotTeams = {}; // slotIdx -> Set команд
  const slotCount = {}; // slotIdx -> кол-во матчей
  const teamLastSlot = {}; // t1/t2 -> последний слот, где команда играла
  const minRestSlots = Math.max(0, params?.minRestSlots || 0);
  const blocked = new Set(params?.blockedSlotIdxs || []);

  const isBlocked = (s) => blocked.has(s);
  const restOk = (m, s) => {
    if (minRestSlots === 0) return true;
    const t1Last = m.t1 !== 'BYE' ? teamLastSlot[m.t1] : undefined;
    const t2Last = m.t2 !== 'BYE' ? teamLastSlot[m.t2] : undefined;
    if (t1Last != null && s - t1Last <= minRestSlots) return false;
    if (t2Last != null && s - t2Last <= minRestSlots) return false;
    return true;
  };

  const place = (m, minSlot) => {
    let s = minSlot;
    while (true) {
      if (isBlocked(s)) { s++; continue; }
      const cnt = slotCount[s] || 0;
      const teams = slotTeams[s] || new Set();
      const conflict = (m.t1 !== 'BYE' && teams.has(m.t1)) || (m.t2 !== 'BYE' && teams.has(m.t2));
      if (cnt < fields && !conflict && restOk(m, s)) {
        slotCount[s] = cnt + 1;
        if (!slotTeams[s]) slotTeams[s] = new Set();
        if (m.t1 !== 'BYE') { slotTeams[s].add(m.t1); teamLastSlot[m.t1] = s; }
        if (m.t2 !== 'BYE') { slotTeams[s].add(m.t2); teamLastSlot[m.t2] = s; }
        schedule.push({ slotIdx: s, field: cnt + 1, matchId: m.id });
        return s;
      }
      s++;
    }
  };

  const groupMatches = matches.filter((m) => m.phase === 'group');
  const poMatches = matches.filter((m) => m.phase === 'playoff');

  let maxGroupSlot = -1;
  groupMatches.forEach((m) => { maxGroupSlot = Math.max(maxGroupSlot, place(m, 0)); });

  // плей-офф: каждый раунд после предыдущего; первый раунд после групп
  let poMinSlot = maxGroupSlot + 1;
  let lastRound = 0;
  poMatches.forEach((m) => {
    if (m.round !== lastRound) { poMinSlot = Math.max(poMinSlot, (schedule.length ? Math.max(...schedule.map(s => s.slotIdx)) : -1) + 1); lastRound = m.round; }
    place(m, poMinSlot);
  });

  return schedule.sort((a, b) => a.slotIdx - b.slotIdx || a.field - b.field);
};

// ============ ГЕНЕРАЦИЯ XLSX С ФОРМУЛАМИ И СТИЛЯМИ ============
// xlsx-js-style глобально подменяет window.XLSX, поэтому используем тот же API
// со свойством `s` для стилей.

// Утилиты стилей
const BD = (color) => ({
  top: { style: 'thin', color: { rgb: color } },
  bottom: { style: 'thin', color: { rgb: color } },
  left: { style: 'thin', color: { rgb: color } },
  right: { style: 'thin', color: { rgb: color } },
});
const STYLES = {
  // MSG-палитра: чёрный 0C0C0C, красный E30613, кремовый F5F2EC, серый 565656
  pageTitle: { font: { bold: true, sz: 16, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '0C0C0C' } }, alignment: { horizontal: 'left', vertical: 'center' } },
  sectionHeader: { font: { bold: true, sz: 12, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: 'E30613' } }, alignment: { horizontal: 'center', vertical: 'center' }, border: BD('B1040F') },
  groupHeader: { font: { bold: true, sz: 12, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '0C0C0C' } }, alignment: { horizontal: 'center', vertical: 'center' }, border: BD('0C0C0C') },
  tableHeader: { font: { bold: true, color: { rgb: '0C0C0C' } }, fill: { fgColor: { rgb: 'F5F2EC' } }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: BD('0C0C0C') },
  cell: { alignment: { horizontal: 'center', vertical: 'center' }, border: BD('D6D3D1') },
  cellName: { alignment: { horizontal: 'left', vertical: 'center', indent: 1 }, border: BD('D6D3D1'), font: { bold: true } },
  cellInput: { alignment: { horizontal: 'center', vertical: 'center' }, border: BD('D6D3D1'), fill: { fgColor: { rgb: 'FFF9E5' } }, numFmt: '0' },
  cellDiagonal: { alignment: { horizontal: 'center', vertical: 'center' }, border: BD('D6D3D1'), fill: { fgColor: { rgb: 'E7E5E4' } } },
  cellComputed: { alignment: { horizontal: 'center', vertical: 'center' }, border: BD('D6D3D1'), fill: { fgColor: { rgb: 'F5F2EC' } }, numFmt: '0' },
  cellPlace: { alignment: { horizontal: 'center', vertical: 'center' }, border: BD('D6D3D1'), font: { bold: true, sz: 12, color: { rgb: 'E30613' } }, fill: { fgColor: { rgb: 'F5F2EC' } }, numFmt: '0' },
  cellSeed: { alignment: { horizontal: 'center', vertical: 'center' }, border: BD('D6D3D1'), font: { bold: true }, fill: { fgColor: { rgb: 'F5F2EC' } } },
  roundFinal: { alignment: { horizontal: 'center', vertical: 'center' }, border: BD('B1040F'), font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: 'E30613' } } },
  roundSemi: { alignment: { horizontal: 'center', vertical: 'center' }, border: BD('D6D3D1'), font: { bold: true, color: { rgb: '0C0C0C' } }, fill: { fgColor: { rgb: 'F5F2EC' } } },
  roundBronze: { alignment: { horizontal: 'center', vertical: 'center' }, border: BD('B1040F'), font: { bold: true, color: { rgb: 'B1040F' } }, fill: { fgColor: { rgb: 'FCE1E3' } } },
  instruction: { alignment: { horizontal: 'left', vertical: 'top', wrapText: true }, font: { sz: 11, color: { rgb: '565656' } } },
};
const setCell = (ws, addr, cell) => { ws[addr] = cell; };
const styled = (cell, style) => ({ ...cell, s: style });

const generateXLSX = (params, structure, matches, schedule, slotDur, fieldNames = {}, varRows = []) => {
  const wb = XLSX.utils.book_new();
  const { system, totalTeams, fields, groupSize, advance, startTime, drawOrder } = params;
  const hasGroups = system === 'group' || system === 'mixed' || system === 'mixed-full' || system === 'mixed-goldsilver' || system === 'mixed-goldsilver-full';
  const isMixedAdvance = system === 'mixed' || system === 'mixed-full';
  const isGoldSilver = system === 'mixed-goldsilver' || system === 'mixed-goldsilver-full';
  const groups = hasGroups ? splitIntoGroups(totalTeams, groupSize, drawOrder) : [];
  const numGroups = groups.length;

  // ===== РАСПРЕДЕЛЕНИЕ КОМАНД =====
  // Команды циклически разнесены по группам внутри splitIntoGroups.
  // sidOf берёт sid из реального списка группы — это устойчиво к группам разного размера.
  const sidOf = (gi, pi) => groups[gi][pi - 1];

  // ===== ЛИСТ: Команды (источник истины для всех имён) =====
  // Структура: A=пусто, B=сид/номер, C=имя команды (юзер вводит), D=куда (для group/mixed)
  const teamNameRef = {}; // sid -> "Команды!$C$N" (абсолютная ссылка)
  {
    const rows = [
      [{ v: 'СПИСОК КОМАНД', s: STYLES.pageTitle }, '', '', '', ''],
      [{ v: 'Впишите названия в колонку «Команда». Они автоматически появятся во всех листах.', s: STYLES.instruction }, '', '', '', ''],
      [],
      [
        { v: '№', s: STYLES.tableHeader },
        { v: 'Сид', s: STYLES.tableHeader },
        { v: 'Команда', s: STYLES.tableHeader },
        ...(numGroups > 0 ? [{ v: 'Группа', s: STYLES.tableHeader }, { v: 'Позиция', s: STYLES.tableHeader }] : []),
      ],
    ];
    // Каждая команда — строка
    const dataStartRow = rows.length + 1; // 1-indexed
    for (let s = 1; s <= totalTeams; s++) {
      const row = [
        { v: s, s: STYLES.cell },
        { v: `Сид ${s}`, s: STYLES.cellSeed },
        { v: `Команда ${s}`, s: STYLES.cellInput },
      ];
      // Подсказка куда попадёт команда
      if (numGroups > 0) {
        // обратный поиск группы по sid
        let gi = -1, pi = -1;
        for (let g = 0; g < numGroups; g++) {
          for (let p = 1; p <= groups[g].length; p++) {
            if (sidOf(g, p) === s) { gi = g; pi = p; break; }
          }
          if (gi >= 0) break;
        }
        if (gi >= 0) {
          row.push({ v: `Группа ${gi + 1}`, s: STYLES.cellComputed });
          row.push({ v: pi, s: STYLES.cellComputed });
        } else {
          row.push({ v: '—', s: STYLES.cellComputed });
          row.push({ v: '—', s: STYLES.cellComputed });
        }
      }
      rows.push(row);
      const rowNum = rows.length; // текущий номер строки = индекс
      teamNameRef[s] = `'Команды'!$C$${rowNum}`;
    }
    const ws = aoa(rows);
    ws['!cols'] = [{ wch: 6 }, { wch: 10 }, { wch: 28 }, ...(numGroups > 0 ? [{ wch: 12 }, { wch: 10 }] : [])];
    ws['!rows'] = [{ hpt: 28 }, { hpt: 38 }];
    // merge title across all columns
    const lastCol = numGroups > 0 ? 'E' : 'C';
    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: numGroups > 0 ? 4 : 2 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: numGroups > 0 ? 4 : 2 } },
    ];
    XLSX.utils.book_append_sheet(wb, ws, 'Команды');
  }

  // ===== ЛИСТ: Группы (визуальная разбивка для group/mixed) =====
  // Показывает кто в какой группе. Имена — формулы со ссылкой на лист Команды.
  if (numGroups > 0) {
    const rows = [
      [{ v: 'СОСТАВ ГРУПП', s: STYLES.pageTitle }, ''],
      [{ v: 'Имена тянутся из листа «Команды». Состав групп меняется автоматически при изменении количества команд в конструкторе.', s: STYLES.instruction }, ''],
      [],
    ];
    groups.forEach((teams, gi) => {
      rows.push([{ v: `ГРУППА ${gi + 1}`, s: STYLES.groupHeader }, { v: '', s: STYLES.groupHeader }]);
      teams.forEach((_, ti) => {
        const sd = sidOf(gi, ti + 1);
        rows.push([
          { v: `${ti + 1}.`, s: STYLES.cell },
          { f: teamNameRef[sd], s: STYLES.cellName },
        ]);
      });
      rows.push([]); // пустая строка между группами
    });
    const ws = aoa(rows);
    ws['!cols'] = [{ wch: 5 }, { wch: 32 }];
    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 1 } },
    ];
    // merge ГРУППА X header across columns
    let cur = 4;
    groups.forEach((teams) => {
      ws['!merges'].push({ s: { r: cur - 1, c: 0 }, e: { r: cur - 1, c: 1 } });
      cur += teams.length + 2;
    });
    XLSX.utils.book_append_sheet(wb, ws, 'Группы');
  }

  // ===== ЛИСТ: Шахматки =====
  // Имена команд — формулы на Команды. Матрица голов = ввод. Очки/места = формулы.
  // Структура для каждой группы: заголовок + матрица NxN + Гз, Гп, Очки, Разн, Место.
  const placeFormula = {}; // `G${gi}:${pi}` (gi 1-indexed, pi place 1..N) → формула получения имени команды на этом месте
  if (numGroups > 0) {
    const ws = {};
    let R = 1;
    // Заголовок страницы
    setCell(ws, `A${R}`, { v: 'ШАХМАТКИ', s: STYLES.pageTitle });
    R++;
    setCell(ws, `A${R}`, { v: 'Впишите счёт каждого матча в формате X:Y (например 3:1) в жёлтых ячейках ВЕРХНЕЙ части матрицы. Серые ячейки нижней части заполнятся зеркально автоматически. Очки, разница и места считаются сами.', s: STYLES.instruction });
    R++; R++;

    let merges = [];
    let maxN = 0;
    groups.forEach((teams, gi) => {
      const N = teams.length;
      maxN = Math.max(maxN, N);
      // Заголовок группы — широкий, покрывает все колонки таблицы
      setCell(ws, `A${R}`, { v: `ГРУППА ${gi + 1}`, s: STYLES.groupHeader });
      const lastTblColIdx = 4 + N + 4 - 1; // 0-indexed: до колонки «Место» включительно
      merges.push({ s: { r: R - 1, c: 0 }, e: { r: R - 1, c: lastTblColIdx } });
      // заполним остальные ячейки той же строки чтобы заливка распространилась
      for (let c = 1; c <= lastTblColIdx; c++) setCell(ws, `${COL(c + 1)}${R}`, { v: '', s: STYLES.groupHeader });
      R++;
      // Шапка таблицы: пустой угол | имена команд (vs) | Гз | Гп | Очки | Разн | Место
      const headerRow = R;
      setCell(ws, `B${headerRow}`, { v: '#', s: STYLES.tableHeader });
      setCell(ws, `C${headerRow}`, { v: 'Команда', s: STYLES.tableHeader });
      const matC0 = 4; // D
      for (let j = 0; j < N; j++) {
        // заголовок столбца "vs" - формула с именем команды j
        const sd = sidOf(gi, j + 1);
        setCell(ws, `${COL(matC0 + j)}${headerRow}`, { f: `"vs " & ${teamNameRef[sd]}`, s: STYLES.tableHeader });
      }
      const gfCol = matC0 + N, gaCol = matC0 + N + 1, ptsCol = matC0 + N + 2, gdCol = matC0 + N + 3, plCol = matC0 + N + 4;
      setCell(ws, `${COL(gfCol)}${headerRow}`, { v: 'Гз', s: STYLES.tableHeader });
      setCell(ws, `${COL(gaCol)}${headerRow}`, { v: 'Гп', s: STYLES.tableHeader });
      setCell(ws, `${COL(ptsCol)}${headerRow}`, { v: 'Очки', s: STYLES.tableHeader });
      setCell(ws, `${COL(gdCol)}${headerRow}`, { v: 'Разн', s: STYLES.tableHeader });
      setCell(ws, `${COL(plCol)}${headerRow}`, { v: 'Место', s: STYLES.tableHeader });
      R++;
      const firstTeamRow = R;
      // Команды
      for (let i = 0; i < N; i++) {
        const row = firstTeamRow + i;
        const sd = sidOf(gi, i + 1);
        setCell(ws, `B${row}`, { v: i + 1, s: STYLES.cell });
        setCell(ws, `C${row}`, { f: teamNameRef[sd], s: STYLES.cellName });
        // МАТРИЦА СЧЁТОВ: вводится ОДНА ячейка на матч в формате "X:Y" (например "3:1").
        // Верхний треугольник (j>i) — жёлтые ячейки ввода (формат «Текст», чтобы Excel
        // не превращал "1:0" во время). Нижний (j<i) — серая зеркальная формула,
        // которая берёт значение из верхнего и переворачивает: "3:1" → "1:3".
        // Универсальный инвариант: в любой ячейке [row, col_vs_j] значение слева от ":"
        // — голы команды этой строки в матче, справа — голы соперника.
        for (let j = 0; j < N; j++) {
          const addr = `${COL(matC0 + j)}${row}`;
          if (i === j) {
            setCell(ws, addr, { v: '—', s: STYLES.cellDiagonal });
          } else if (j > i) {
            // ячейка ввода — текстовый формат
            setCell(ws, addr, { v: '', t: 's', s: { ...STYLES.cellInput, numFmt: '@' } });
          } else {
            // зеркало верхнего: [j][i] содержит счёт "X:Y" с точки зрения команды j
            // здесь нужна перевёрнутая строка "Y:X"
            const mirror = `${COL(matC0 + i)}${firstTeamRow + j}`;
            const f = `IFERROR(IF(${mirror}="","",MID(${mirror},FIND(":",${mirror})+1,10)&":"&LEFT(${mirror},FIND(":",${mirror})-1)),"")`;
            setCell(ws, addr, { f, t: 's', s: { ...STYLES.cellComputed, numFmt: '@' } });
          }
        }
        // Голы забитые = СУММ(LEFT перед ":" в каждой ячейке строки)
        // Голы пропущенные = СУММ(MID после ":" в каждой ячейке строки)
        const gfTerms = [], gaTerms = [], ptsTerms = [];
        for (let j = 0; j < N; j++) {
          if (i === j) continue;
          const a = `${COL(matC0 + j)}${row}`;
          const lf = `VALUE(LEFT(${a},FIND(":",${a})-1))`;
          const md = `VALUE(MID(${a},FIND(":",${a})+1,10))`;
          gfTerms.push(`IFERROR(${lf},0)`);
          gaTerms.push(`IFERROR(${md},0)`);
          // Очки: 3 если забил>пропустил, 1 если ничья, 0 если меньше; пустая ячейка → 0
          ptsTerms.push(`IFERROR(IF(${lf}>${md},3,IF(${lf}=${md},1,0)),0)`);
        }
        setCell(ws, `${COL(gfCol)}${row}`, { f: gfTerms.join('+'), s: STYLES.cellComputed });
        setCell(ws, `${COL(gaCol)}${row}`, { f: gaTerms.join('+'), s: STYLES.cellComputed });
        setCell(ws, `${COL(ptsCol)}${row}`, { f: ptsTerms.join('+'), s: STYLES.cellComputed });
        // Разница
        setCell(ws, `${COL(gdCol)}${row}`, { f: `${COL(gfCol)}${row}-${COL(gaCol)}${row}`, s: STYLES.cellComputed });
        // Место с тай-брейком по строке (всегда уникально)
        const ptsRange = `${COL(ptsCol)}${firstTeamRow}:${COL(ptsCol)}${firstTeamRow + N - 1}`;
        const gdRange = `${COL(gdCol)}${firstTeamRow}:${COL(gdCol)}${firstTeamRow + N - 1}`;
        const rwr = `ROW(${COL(ptsCol)}${firstTeamRow}:${COL(ptsCol)}${firstTeamRow + N - 1})`;
        const myPts = `${COL(ptsCol)}${row}`, myGd = `${COL(gdCol)}${row}`, myRow = `ROW(${COL(ptsCol)}${row})`;
        const sc = `(${ptsRange}*1000+${gdRange})`;
        const ms = `(${myPts}*1000+${myGd})`;
        setCell(ws, `${COL(plCol)}${row}`, { f: `SUMPRODUCT(--(${sc}>${ms}))+SUMPRODUCT(--(${sc}=${ms}),--(${rwr}<${myRow}))+1`, s: STYLES.cellPlace });
      }
      // Диапазоны для извлечения имени по месту
      const nameRange = `'Шахматки'!$C$${firstTeamRow}:$C$${firstTeamRow + N - 1}`;
      const placeRange = `'Шахматки'!$${COL(plCol)}$${firstTeamRow}:$${COL(plCol)}$${firstTeamRow + N - 1}`;
      for (let p = 1; p <= N; p++) {
        placeFormula[`G${gi + 1}:${p}`] = `INDEX(${nameRange},MATCH(${p},${placeRange},0))`;
      }
      R = firstTeamRow + N + 1; // пустая строка между группами
    });
    ws['!ref'] = `A1:${COL(4 + maxN + 4)}${R}`;
    // Колонки: A узкая (3), B=#(4), C=Команда(24), матрица vs Команда(14 каждая), Гз/Гп/Очки/Разн/Место по 8-9
    const lastColIdx = 4 + maxN + 4 - 1; // 0-indexed индекс последней колонки данных
    ws['!cols'] = [{ wch: 3 }, { wch: 4 }, { wch: 24 }, ...Array(maxN).fill({ wch: 14 }), { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 9 }];
    // Высота строк: заголовок и инструкция — повыше
    ws['!rows'] = [{ hpt: 28 }, { hpt: 36 }];
    // Объединения: заголовок страницы и инструкция растянуты на всю ширину
    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: lastColIdx } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: lastColIdx } },
      ...merges,
    ];
    XLSX.utils.book_append_sheet(wb, ws, 'Шахматки');
  }

  // ===== ЛИСТ: Плей-офф =====
  const poWinRef = {}; // matchId -> "Плей-офф!$I$N"
  const poLoseRef = {};
  const poT1Ref = {}; // matchId -> "Плей-офф!$D$N"
  const poT2Ref = {};
  if (system !== 'group') {
    const poMatches = matches.filter((m) => m.phase === 'playoff');
    const ws = {};
    let R = 1;
    setCell(ws, `A${R}`, { v: 'СЕТКА ПЛЕЙ-ОФФ', s: STYLES.pageTitle });
    R++;
    setCell(ws, `A${R}`, { v: 'Вписывайте голы (только в жёлтые ячейки). Победитель проходит в следующий раунд автоматически. Проигравшие полуфиналов попадают в матч за бронзу.', s: STYLES.instruction });
    R += 2;
    // Шапка
    const C = { id: 2, rnd: 3, t1: 4, g1: 5, sep: 6, g2: 7, t2: 8, win: 9, lose: 10 };
    const headerRow = R;
    setCell(ws, `B${headerRow}`, { v: '№', s: STYLES.tableHeader });
    setCell(ws, `C${headerRow}`, { v: 'Раунд', s: STYLES.tableHeader });
    setCell(ws, `D${headerRow}`, { v: 'Команда 1', s: STYLES.tableHeader });
    setCell(ws, `E${headerRow}`, { v: 'Г1', s: STYLES.tableHeader });
    setCell(ws, `F${headerRow}`, { v: '', s: STYLES.tableHeader });
    setCell(ws, `G${headerRow}`, { v: 'Г2', s: STYLES.tableHeader });
    setCell(ws, `H${headerRow}`, { v: 'Команда 2', s: STYLES.tableHeader });
    setCell(ws, `I${headerRow}`, { v: 'Победитель', s: STYLES.tableHeader });
    setCell(ws, `J${headerRow}`, { v: 'Проигравший', s: STYLES.tableHeader });
    R++;
    // Распределение матчей по строкам
    const rowOf = {};
    poMatches.forEach((m, i) => { rowOf[m.id] = R + i; });

    poMatches.forEach((m) => {
      const r = rowOf[m.id];
      // Выбираем стиль строки в зависимости от раунда
      let rowStyle = STYLES.cell;
      if (m.isBronze) rowStyle = STYLES.roundBronze;
      else if (m.roundName === 'Финал') rowStyle = STYLES.roundFinal;
      else if (m.roundName === '1/2') rowStyle = STYLES.roundSemi;

      setCell(ws, `B${r}`, { v: m.id, s: rowStyle });
      setCell(ws, `C${r}`, { v: m.roundName, s: rowStyle });
      // Формулы победитель/проигравший
      const winF = `IF(AND(${COL(C.g1)}${r}<>"",${COL(C.g2)}${r}<>""),IF(${COL(C.g1)}${r}>${COL(C.g2)}${r},${COL(C.t1)}${r},IF(${COL(C.g2)}${r}>${COL(C.g1)}${r},${COL(C.t2)}${r},"=")),"")`;
      const loseF = `IF(AND(${COL(C.g1)}${r}<>"",${COL(C.g2)}${r}<>""),IF(${COL(C.g1)}${r}>${COL(C.g2)}${r},${COL(C.t2)}${r},IF(${COL(C.g2)}${r}>${COL(C.g1)}${r},${COL(C.t1)}${r},"")),"")`;
      setCell(ws, `${COL(C.win)}${r}`, { f: winF, s: { ...rowStyle, font: { ...(rowStyle.font || {}), bold: true } } });
      setCell(ws, `${COL(C.lose)}${r}`, { f: loseF, s: { ...rowStyle, font: { ...(rowStyle.font || {}), italic: true } } });
      setCell(ws, `${COL(C.sep)}${r}`, { v: ':', s: rowStyle });
      setCell(ws, `${COL(C.g1)}${r}`, { v: '', t: 's', s: STYLES.cellInput });
      setCell(ws, `${COL(C.g2)}${r}`, { v: '', t: 's', s: STYLES.cellInput });
      poWinRef[m.id] = `'Плей-офф'!$${COL(C.win)}$${r}`;
      poLoseRef[m.id] = `'Плей-офф'!$${COL(C.lose)}$${r}`;
      poT1Ref[m.id] = `'Плей-офф'!$${COL(C.t1)}$${r}`;
      poT2Ref[m.id] = `'Плей-офф'!$${COL(C.t2)}$${r}`;

      // Команды
      const fillTeam = (col) => {
        const addr = `${COL(col)}${r}`;
        if (m.winFrom) {
          const sid = col === C.t1 ? m.winFrom[0] : m.winFrom[1];
          setCell(ws, addr, { f: poWinRef[sid], s: rowStyle });
        } else if (m.loseFrom) {
          const sid = col === C.t1 ? m.loseFrom[0] : m.loseFrom[1];
          setCell(ws, addr, { f: poLoseRef[sid], s: rowStyle });
        } else if (m.seedA !== undefined) {
          const sd = col === C.t1 ? m.seedA : m.seedB;
          if (sd > m.playoffTeams) {
            setCell(ws, addr, { v: 'BYE', s: { ...rowStyle, font: { ...(rowStyle.font || {}), italic: true, color: { rgb: '94A3B8' } } } });
          } else if (m.bracket === 'gold' || m.bracket === 'silver') {
            // Золото/серебро: сид N — это группа N, место фиксировано (1-е для золота, 2-е для серебра)
            const placeIdx = m.bracket === 'gold' ? 1 : 2;
            const grIdx = sd;
            const key = `G${grIdx}:${placeIdx}`;
            const placeOrd = placeIdx === 1 ? '1-е' : '2-е';
            const fall = `${placeOrd} место Гр.${grIdx}`;
            setCell(ws, addr, placeFormula[key]
              ? { f: `IFERROR(${placeFormula[key]},"${fall}")`, s: rowStyle }
              : { v: fall, s: rowStyle });
          } else if (isMixedAdvance) {
            const ng = numGroups;
            const placeIdx = Math.ceil(sd / ng);
            const grIdx = ((sd - 1) % ng) + 1;
            const key = `G${grIdx}:${placeIdx}`;
            const placeOrd = ['', '1-е', '2-е', '3-е', '4-е', '5-е', '6-е', '7-е', '8-е'][placeIdx] || `${placeIdx}-е`;
            const fall = `${placeOrd} место Гр.${grIdx}`;
            setCell(ws, addr, placeFormula[key]
              ? { f: `IFERROR(${placeFormula[key]},"${fall}")`, s: rowStyle }
              : { v: fall, s: rowStyle });
          } else {
            // playoff: ссылка на лист Команды
            setCell(ws, addr, { f: teamNameRef[sd], s: rowStyle });
          }
        }
      };
      fillTeam(C.t1);
      fillTeam(C.t2);
    });
    ws['!ref'] = `A1:J${R + poMatches.length}`;
    ws['!cols'] = [{ wch: 3 }, { wch: 5 }, { wch: 9 }, { wch: 24 }, { wch: 5 }, { wch: 3 }, { wch: 5 }, { wch: 24 }, { wch: 24 }, { wch: 24 }];
    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 9 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 9 } },
    ];
    XLSX.utils.book_append_sheet(wb, ws, 'Плей-офф');
  }

  // ===== ЛИСТ: Расписание =====
  // Каждая ячейка матча — формула с настоящими именами через ссылки.
  // Слоты глобальные и сквозные; день = floor(slotIdx / slotsPerDay), время = startTime + localSlot*slotDur.
  {
    const slotMap = {};
    schedule.forEach((s) => {
      if (!slotMap[s.slotIdx]) slotMap[s.slotIdx] = {};
      slotMap[s.slotIdx][s.field] = matches.find((x) => x.id === s.matchId);
    });
    const sorted = Object.keys(slotMap).map(Number).sort((a, b) => a - b);
    const slotsPerDay = structure.slotsPerDay;

    const ws = {};
    let R = 1;
    setCell(ws, `A${R}`, { v: 'РАСПИСАНИЕ', s: STYLES.pageTitle });
    R++;
    const totalDays = sorted.length === 0 ? 1 : Math.floor(sorted[sorted.length-1] / slotsPerDay) + 1;
    setCell(ws, `A${R}`, { v: `Старт ${startTime} · слот ${slotDur} мин · полей ${fields} · дней ${totalDays}. Времена и имена команд считаются автоматически.`, s: STYLES.instruction });
    R += 2;
    const headerRow = R;
    setCell(ws, `B${headerRow}`, { v: 'Слот', s: STYLES.tableHeader });
    setCell(ws, `C${headerRow}`, { v: 'Время', s: STYLES.tableHeader });
    for (let f = 0; f < fields; f++) setCell(ws, `${COL(4 + f)}${headerRow}`, { v: fieldNames[f + 1] || `Поле ${f + 1}`, s: STYLES.tableHeader });
    R++;
    const scheduleMerges = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 3 + fields - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 3 + fields - 1 } },
    ];
    let curDay = 0;
    let slotInDayCounter = 0;
    const dOffsets = structure.dayOffsets || [0];
    const sInDay = structure.slotsInDay || [structure.slotsPerDay];
    const dInfos = structure.dayInfos || [{ startTime }];
    sorted.forEach((slotIdx) => {
      // Определяем день и локальный слот
      let day = 1, localSlot = slotIdx;
      for (let d = dOffsets.length - 1; d >= 0; d--) {
        if (slotIdx >= dOffsets[d]) { day = d + 1; localSlot = slotIdx - dOffsets[d]; break; }
      }
      // Разделитель дня
      if (day !== curDay) {
        setCell(ws, `A${R}`, { v: `ДЕНЬ ${day}`, s: STYLES.groupHeader });
        for (let c = 1; c <= 3 + fields - 1; c++) setCell(ws, `${COL(c + 1)}${R}`, { v: '', s: STYLES.groupHeader });
        scheduleMerges.push({ s: { r: R - 1, c: 0 }, e: { r: R - 1, c: 3 + fields - 1 } });
        R++;
        curDay = day;
        slotInDayCounter = 0;
      }
      slotInDayCounter++;
      const dayStartTime = (dInfos[day - 1] || dInfos[0]).startTime;
      const tStart = minToTime(timeToMin(dayStartTime) + localSlot * slotDur);
      const tEnd = minToTime(timeToMin(dayStartTime) + (localSlot + 1) * slotDur);
      setCell(ws, `B${R}`, { v: slotInDayCounter, s: STYLES.cell });
      setCell(ws, `C${R}`, { v: `${tStart}–${tEnd}`, s: { ...STYLES.cell, font: { name: 'Consolas' } } });
      for (let f = 1; f <= fields; f++) {
        const addr = `${COL(3 + f)}${R}`;
        const m = slotMap[slotIdx][f];
        if (!m) { setCell(ws, addr, { v: '', s: STYLES.cell }); continue; }
        // Получаем ссылки на имена команд
        let t1Ref, t2Ref;
        if (m.phase === 'group') {
          const sd1 = sidOf(m.group - 1, parseInt(m.t1.split('.')[1]));
          const sd2 = sidOf(m.group - 1, parseInt(m.t2.split('.')[1]));
          t1Ref = teamNameRef[sd1];
          t2Ref = teamNameRef[sd2];
        } else {
          t1Ref = poT1Ref[m.id];
          t2Ref = poT2Ref[m.id];
        }
        const label = m.label;
        const isPo = m.phase === 'playoff';
        const f1 = `"${label}: " & IF(${t1Ref}="","?",${t1Ref}) & " — " & IF(${t2Ref}="","?",${t2Ref})`;
        setCell(ws, addr, { f: f1, s: isPo ? { ...STYLES.cell, fill: { fgColor: { rgb: 'FCE1E3' } }, font: { bold: true, color: { rgb: 'B1040F' } } } : STYLES.cell });
      }
      R++;
    });
    ws['!ref'] = `A1:${COL(3 + fields)}${R}`;
    ws['!cols'] = [{ wch: 3 }, { wch: 6 }, { wch: 13 }, ...Array(fields).fill({ wch: 38 })];
    ws['!merges'] = scheduleMerges;
    XLSX.utils.book_append_sheet(wb, ws, 'Расписание');
  }

  // ===== ЛИСТ: Параметры (вставляем первым) =====
  {
    // Считаем время финиша последнего матча в последнем дне
    const sortedSchedule = schedule.map((s) => s.slotIdx).sort((a, b) => a - b);
    const lastSlot = sortedSchedule.length ? sortedSchedule[sortedSchedule.length - 1] : 0;
    const lastLocalSlot = lastSlot % structure.slotsPerDay;
    const finishMin = timeToMin(startTime) + (lastLocalSlot + 1) * slotDur;
    const finishTime = minToTime(finishMin);
    const finishLabel = structure.daysNeeded > 1 ? `Финиш (день ${structure.daysNeeded})` : 'Финиш последнего матча';

    const sysName = {
      group: 'Групповая', playoff: 'Плей-офф', mixed: 'Смешанная (группы + плей-офф)',
      'playoff-full': 'Плей-офф (розыгрыш всех мест)', 'mixed-full': 'Смешанная (розыгрыш всех мест)',
      'mixed-goldsilver': 'Смешанная (золото/серебро)', 'mixed-goldsilver-full': 'Смешанная (золото/серебро, все места)',
    }[system];
    const rows = [
      [{ v: 'ПАРАМЕТРЫ ТУРНИРА', s: STYLES.pageTitle }, ''],
      [],
      [{ v: 'Команд', s: STYLES.cellName }, { v: totalTeams, s: STYLES.cell }],
      [{ v: 'Полей', s: STYLES.cellName }, { v: fields, s: STYLES.cell }],
      [{ v: 'Дней', s: STYLES.cellName }, { v: structure.days, s: STYLES.cell }],
      [{ v: 'Система', s: STYLES.cellName }, { v: sysName, s: STYLES.cell }],
      [{ v: 'Старт каждого дня', s: STYLES.cellName }, { v: startTime, s: STYLES.cell }],
      [{ v: finishLabel, s: STYLES.cellName }, { v: finishTime, s: STYLES.cell }],
      [{ v: 'Длительность слота (мин)', s: STYLES.cellName }, { v: slotDur, s: STYLES.cell }],
      [{ v: 'Слотов в день', s: STYLES.cellName }, { v: structure.slotsPerDay, s: STYLES.cell }],
    ];
    if (numGroups > 0) {
      rows.push([{ v: 'Размер группы', s: STYLES.cellName }, { v: groupSize, s: STYLES.cell }]);
      rows.push([{ v: 'Групп', s: STYLES.cellName }, { v: numGroups, s: STYLES.cell }]);
    }
    if (isMixedAdvance) rows.push([{ v: 'Из группы в плей-офф', s: STYLES.cellName }, { v: `топ-${advance}`, s: STYLES.cell }]);
    if (isGoldSilver) rows.push([{ v: 'Из группы в плей-офф', s: STYLES.cellName }, { v: '1-е — золото, 2-е — серебро', s: STYLES.cell }]);
    rows.push([]);
    rows.push([{ v: 'Всего матчей', s: STYLES.cellName }, { v: matches.length, s: STYLES.cell }]);
    rows.push([{ v: 'Всего слотов', s: STYLES.cellName }, { v: structure.slots, s: STYLES.cell }]);
    rows.push([]);
    rows.push([{ v: 'КАК ПОЛЬЗОВАТЬСЯ', s: STYLES.sectionHeader }, { v: '', s: STYLES.sectionHeader }]);
    if (system === 'playoff' || system === 'playoff-full') {
      rows.push([{ v: '1.', s: STYLES.cell }, { v: 'На листе «Команды» впишите названия (колонка «Команда»)', s: STYLES.cellName }]);
      rows.push([{ v: '2.', s: STYLES.cell }, { v: 'На листе «Плей-офф» вписывайте голы в жёлтые ячейки — победители проходят сами', s: STYLES.cellName }]);
      rows.push([{ v: '3.', s: STYLES.cell }, { v: 'Лист «Расписание» — календарь по полям', s: STYLES.cellName }]);
    } else if (system === 'group') {
      rows.push([{ v: '1.', s: STYLES.cell }, { v: 'На листе «Команды» впишите названия. Они автоматически распределятся по группам', s: STYLES.cellName }]);
      rows.push([{ v: '2.', s: STYLES.cell }, { v: 'На листе «Шахматки» вписывайте голы в жёлтые ячейки — очки и места считаются сами', s: STYLES.cellName }]);
      rows.push([{ v: '3.', s: STYLES.cell }, { v: 'Лист «Группы» показывает состав, «Расписание» — календарь', s: STYLES.cellName }]);
    } else {
      rows.push([{ v: '1.', s: STYLES.cell }, { v: 'На листе «Команды» впишите названия. Они автоматически распределятся по группам', s: STYLES.cellName }]);
      rows.push([{ v: '2.', s: STYLES.cell }, { v: 'На листе «Шахматки» вписывайте голы — очки и места считаются сами', s: STYLES.cellName }]);
      rows.push([{ v: '3.', s: STYLES.cell }, { v: isGoldSilver ? 'В золотой плей-офф выходят 1-е места групп, в серебряный — 2-е. Вписывайте голы — победители проходят сами' : ('В сетке плей-офф первые ' + advance + ' места из групп подставятся автоматически. Вписывайте голы — победители проходят сами'), s: STYLES.cellName }]);
      rows.push([{ v: '4.', s: STYLES.cell }, { v: 'Лист «Расписание» — календарь матчей по полям и слотам', s: STYLES.cellName }]);
    }
    const ws = aoa(rows);
    ws['!cols'] = [{ wch: 30 }, { wch: 60 }];
    ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
    // Вставляем первым
    wb.SheetNames.unshift('Параметры');
    wb.Sheets['Параметры'] = ws;
  }

  // ===== ЛИСТ: VAR (зафиксированные в приложении эпизоды) =====
  if (varRows.length > 0) {
    const rows = [
      [{ v: 'VAR / СОБЫТИЯ МАТЧЕЙ', s: STYLES.pageTitle }, '', '', '', ''],
      [{ v: 'Эпизоды, зафиксированные в приложении при вводе счёта.', s: STYLES.instruction }, '', '', '', ''],
      [],
      [
        { v: 'Матч', s: STYLES.tableHeader },
        { v: 'Мин.', s: STYLES.tableHeader },
        { v: 'Команда', s: STYLES.tableHeader },
        { v: 'Событие', s: STYLES.tableHeader },
        { v: 'Комментарий', s: STYLES.tableHeader },
      ],
    ];
    varRows.forEach((e) => {
      rows.push([
        { v: e.matchLabel, s: STYLES.cellName },
        { v: e.minute != null ? e.minute : '—', s: STYLES.cell },
        { v: e.team, s: STYLES.cell },
        { v: e.type, s: STYLES.cell },
        { v: e.note || '—', s: STYLES.cell },
      ]);
    });
    const ws = aoa(rows);
    ws['!cols'] = [{ wch: 20 }, { wch: 6 }, { wch: 22 }, { wch: 26 }, { wch: 40 }];
    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 4 } },
    ];
    XLSX.utils.book_append_sheet(wb, ws, 'VAR');
  }

  // КРИТИЧНО: заставляем Excel пересчитать все формулы при открытии файла.
  // xlsx-js-style не записывает calcPr, поэтому модифицируем XML вручную после генерации.
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true });

  // Распаковываем zip, добавляем calcPr fullCalcOnLoad, запаковываем
  const fflate = window.fflate;
  const zip = fflate.unzipSync(new Uint8Array(wbout));
  const workbookXml = fflate.strFromU8(zip['xl/workbook.xml']);
  let patchedXml;
  if (/<calcPr\b/.test(workbookXml)) {
    patchedXml = workbookXml.replace(/<calcPr([^/]*)\/>/, '<calcPr$1 fullCalcOnLoad="1"/>');
  } else {
    patchedXml = workbookXml.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
  }
  zip['xl/workbook.xml'] = fflate.strToU8(patchedXml);
  const newZipBytes = fflate.zipSync(zip);

  // Скачать как Blob
  const date = new Date().toISOString().slice(0, 10);
  const filename = `tournament_${totalTeams}t_${system}_${date}.xlsx`;
  const blob = new Blob([newZipBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
};

// Хелпер aoa с сохранением s свойств (XLSX.utils.aoa_to_sheet их игнорирует если объект, не строка)
function aoa(rows) {
  const ws = {};
  let maxC = 0;
  rows.forEach((row, ri) => {
    row.forEach((cell, ci) => {
      if (cell === undefined || cell === null || cell === '') return;
      maxC = Math.max(maxC, ci);
      const addr = `${COL(ci + 1)}${ri + 1}`;
      if (typeof cell === 'object' && (cell.v !== undefined || cell.f !== undefined || cell.s !== undefined)) {
        ws[addr] = { ...cell };
        // тип
        if (cell.v !== undefined && cell.t === undefined) ws[addr].t = typeof cell.v === 'number' ? 'n' : 's';
      } else {
        ws[addr] = { v: cell, t: typeof cell === 'number' ? 'n' : 's' };
      }
    });
  });
  ws['!ref'] = `A1:${COL(maxC + 1)}${rows.length}`;
  return ws;
}


// ============ СТРУКТУРА ============
// Возвращает информацию о конкретном дне с учётом переопределений в dayWindows.
const getDayWindow = (dayNum, params) => {
  const overrides = params.dayWindows || {};
  const w = overrides[dayNum];
  const startTime = (w && w.startTime) || params.startTime;
  const endTime = (w && w.endTime) || params.endTime;
  const dayMin = timeToMin(endTime) - timeToMin(startTime);
  const availMin = Math.max(60, dayMin - 30);
  return { startTime, endTime, dayMin, availMin };
};

// Принимает реальные matches и schedule, чтобы получить точное число слотов с учётом
// упаковки (конфликты команд, разделители раундов плей-офф).
const computeStructure = (params, matches, schedule, slotDurOverride) => {
  const { totalTeams, system, groupSize, advance, fields, scheduleMode, maxGamesPerDay } = params;
  let days = Math.max(1, params.days || 1);
  const games = matches.length;
  const slotsNeeded = schedule.length === 0 ? 0 : Math.max(...schedule.map((s) => s.slotIdx)) + 1;

  const dayInfosFor = (n) => { const a = []; for (let d = 1; d <= n; d++) a.push(getDayWindow(d, params)); return a; };

  // Режим «по дням»: организатор задаёт лимит игр в день, а не число дней —
  // количество дней растёт само, пока не наберётся достаточно слотов.
  const byDay = scheduleMode === 'byDay' && maxGamesPerDay > 0;
  const slotsPerDayCap = byDay ? Math.max(1, Math.ceil(maxGamesPerDay / Math.max(1, fields))) : Infinity;

  let matchDur = slotDurOverride;
  if (!matchDur) {
    if (byDay) {
      // Длительность подбираем под лимит игр в день, а не под число дней —
      // иначе оценка «сколько дней нужно» ниже была бы противоречивой.
      const firstDayInfo = getDayWindow(1, params);
      matchDur = Math.max(15, Math.min(200, Math.floor(firstDayInfo.availMin / slotsPerDayCap)));
    } else {
      const dayInfos0 = dayInfosFor(days);
      const totalAvailMin = dayInfos0.reduce((s, d) => s + d.availMin, 0);
      const totalSlotsAt = (dur) => dayInfos0.reduce((s, d) => s + Math.max(1, Math.floor(d.availMin / dur)), 0);
      matchDur = Math.max(15, Math.floor(totalAvailMin / Math.max(1, slotsNeeded)));
      // Уменьшаем пока не влезаем в дни
      while (matchDur > 15 && totalSlotsAt(matchDur) < slotsNeeded) matchDur--;
      // Или увеличиваем пока не влезаем в дни, но не более 200
      while (matchDur < 200 && totalSlotsAt(matchDur) < slotsNeeded) matchDur++;
      matchDur = Math.max(15, matchDur);
    }
  }

  if (byDay) {
    // Растим число дней, пока (лимит игр в день → слотов в день) не покроет все матчи
    while (days < 60) {
      const infos = dayInfosFor(days);
      const cap = infos.reduce((s, d) => s + Math.min(slotsPerDayCap, Math.max(1, Math.floor(d.availMin / matchDur))), 0);
      if (cap >= slotsNeeded) break;
      days++;
    }
  }

  // Массив слотов в каждом дне с учётом финального matchDur (и лимита игр/день, если задан)
  const dayInfos = dayInfosFor(days);
  const slotsInDay = dayInfos.map((d) => Math.min(slotsPerDayCap, Math.max(1, Math.floor(d.availMin / matchDur))));
  const dayOffsets = [0];
  for (let d = 0; d < days - 1; d++) dayOffsets.push(dayOffsets[d] + slotsInDay[d]);
  const totalSlots = slotsInDay.reduce((s, x) => s + x, 0);

  // Определение дня по глобальному слоту
  let daysNeeded = 1;
  if (slotsNeeded > 0) {
    let acc = 0;
    for (let d = 0; d < days; d++) {
      acc += slotsInDay[d];
      if (slotsNeeded <= acc) { daysNeeded = d + 1; break; }
      daysNeeded = d + 1;
    }
  }

  // Совместимость со старым API: slotsPerDay = средний / первый день
  const slotsPerDay = slotsInDay[0] || 1;
  const availMin = dayInfos[0].availMin;
  const dayMin = dayInfos[0].dayMin;

  let numGroups = 0, playoffTeams = 0;
  const hasGroupsSys = system === 'group' || system === 'mixed' || system === 'mixed-full' || system === 'mixed-goldsilver' || system === 'mixed-goldsilver-full';
  if (hasGroupsSys) numGroups = splitIntoGroups(totalTeams, groupSize).length;
  if (system === 'mixed' || system === 'mixed-full') playoffTeams = numGroups * advance;
  else if (system === 'mixed-goldsilver' || system === 'mixed-goldsilver-full') playoffTeams = numGroups * 2; // золото+серебро вместе, для метрик
  else if (system === 'playoff' || system === 'playoff-full') playoffTeams = totalTeams;

  return {
    games, slots: slotsNeeded, matchDur, numGroups, playoffTeams,
    availMin, dayMin, slotsPerDay, daysNeeded, days,
    slotsInDay, dayOffsets, dayInfos, totalSlots,
  };
};

// Утилита: глобальный slotIdx → { day, local } с учётом переменного размера дней
const slotToDay = (slotIdx, dayOffsets, slotsInDay) => {
  if (!dayOffsets || dayOffsets.length === 0) return { day: 1, local: slotIdx };
  for (let d = dayOffsets.length - 1; d >= 0; d--) {
    if (slotIdx >= dayOffsets[d]) return { day: d + 1, local: slotIdx - dayOffsets[d] };
  }
  return { day: 1, local: slotIdx };
};

// ============ ХРАНЕНИЕ НЕСКОЛЬКИХ ТУРНИРОВ («личный кабинет» без сервера) ============
// Список турниров — лёгкий индекс (имя, дата, кол-во команд) для отображения в
// кабинете без загрузки полных данных каждого. Данные каждого турнира — отдельным
// ключом, чтобы не читать/писать весь список турниров на каждое изменение счёта.
const TOURNAMENTS_INDEX_KEY = 'msg_tournaments_index_v1';
const LEGACY_STORAGE_KEY = 'msg_tournament_v1'; // старый формат — один турнир без ID
const tournamentDataKey = (id) => `msg_tournament_data_v1:${id}`;
const makeTournamentId = () => `t_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
const SYS_LABELS = {
  auto: 'Авто', group: 'Групповая', playoff: 'Плей-офф', mixed: 'Смешанная',
  'playoff-full': 'Плей-офф (все места)', 'mixed-full': 'Смешанная (все места)',
  'mixed-goldsilver': 'Золото/серебро', 'mixed-goldsilver-full': 'Золото/серебро (все места)',
};

// ============ ОНЛАЙН-СИНХРОНИЗАЦИЯ СУДЕЙ И ТРАНСЛЯЦИЯ ДЛЯ РОДИТЕЛЕЙ ============
// Отдельный бэкенд (Cloudflare Worker + D1 + Durable Object) — судьи на разных
// телефонах пишут счёт в одно место, организатор и родители видят это в
// реальном времени по WebSocket. Полностью опционально: без публикации всё
// работает как раньше, только в localStorage этого браузера.
const SYNC_BACKEND_URL = 'https://ivory-falcon-377.higgsfield.app';

const loadTournamentsIndex = () => {
  try {
    const raw = localStorage.getItem(TOURNAMENTS_INDEX_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch { return []; }
};
const saveTournamentsIndex = (list) => {
  try { localStorage.setItem(TOURNAMENTS_INDEX_KEY, JSON.stringify(list)); } catch (e) { console.error('index save failed', e); }
};
// Разовая миграция: старые пользователи хранили ровно один турнир под фиксированным
// ключом. При первом запуске новой версии переносим его в список как первую запись.
const migrateLegacyTournament = () => {
  let list = loadTournamentsIndex();
  if (list.length > 0) return list;
  try {
    const raw = localStorage.getItem(LEGACY_STORAGE_KEY);
    if (!raw) return list;
    const data = JSON.parse(raw);
    const id = makeTournamentId();
    localStorage.setItem(tournamentDataKey(id), JSON.stringify(data));
    list = [{ id, name: 'Турнир', savedAt: Date.now(), totalTeams: data.params?.totalTeams, system: data.params?.system }];
    saveTournamentsIndex(list);
    localStorage.removeItem(LEGACY_STORAGE_KEY);
  } catch (e) { console.error('migration failed', e); }
  return list;
};

// ============ КОМПОНЕНТ ============

// Читаем параметр URL ?judge=matchId
const getJudgeMatchId = () => {
  try {
    const url = new URL(window.location.href);
    const j = url.searchParams.get('judge');
    return j ? parseInt(j) : null;
  } catch { return null; }
};

export default function TournamentBuilder() {
  const [params, setParams] = useState({
    totalTeams: 20, fields: 2, startTime: '10:00', endTime: '20:00',
    system: 'auto', groupSize: 4, advance: 2, days: 1,
    scheduleMode: 'sequential', restMode: 'auto', maxGamesPerDay: null,
    drawMode: 'sequential', numSeeds: 4,
    refereeMode: 'manual',
  });
  const [matchDurMode, setMatchDurMode] = useState('auto');
  const [manualDur, setManualDur] = useState(40);
  const [previewSlots, setPreviewSlots] = useState(10);
  // Данные турнира — команды и счета (сохраняются между сессиями)
  const [teamNames, setTeamNames] = useState({}); // { sid: "Тигры" }
  const [teamColors, setTeamColors] = useState({}); // { sid: "#e30613" }
  const [scores, setScores] = useState({}); // { matchId: { a: 3, b: 1 } }
  const [fieldNames, setFieldNames] = useState({}); // { 1: "Лужники, поле A" }
  const [minRest, setMinRest] = useState(0); // минимум минут отдыха между матчами команды
  const [blockedSlots, setBlockedSlots] = useState([]); // [{ day, slotStart, slotEnd, label }]
  const [dayWindows, setDayWindows] = useState({}); // { 2: { startTime: "09:00", endTime: "18:00" } }
  const [drawOrder, setDrawOrder] = useState(null); // результат жеребьёвки: массив sid, null = ещё не проводили (последовательно)
  const [refereeNames, setRefereeNames] = useState({}); // список судей: { 1: "Иванов И.И." }
  const [matchReferees, setMatchReferees] = useState({}); // { matchId: refId (fromList/random) | "имя" (manual) }
  const [tab, setTab] = useState('setup');
  const [scoreModal, setScoreModal] = useState(null);
  const [importModal, setImportModal] = useState(false);
  const [qrModal, setQrModal] = useState(null); // { matchId, matchLabel }
  const [protocolModal, setProtocolModal] = useState(null); // { matchId, matchLabel, t1Label, t2Label, sid1, sid2 }

  // «Личный кабинет»: список турниров + какой сейчас открыт
  const [tournamentsIndex, setTournamentsIndex] = useState([]);
  const [tournamentId, setTournamentId] = useState(null);
  const [tournamentName, setTournamentName] = useState('Турнир');
  const [showDashboard, setShowDashboard] = useState(false);

  // Онлайн-синхронизация: если турнир опубликован, тут лежит его id на бэкенде
  const [onlineId, setOnlineId] = useState(null);
  const [publishing, setPublishing] = useState(false);

  // Подставляет в состояние данные турнира с данным id (или пустой турнир, если данных ещё нет)
  const loadTournamentData = (id, list) => {
    const entry = (list || tournamentsIndex).find((t) => t.id === id);
    setTournamentId(id);
    setTournamentName(entry ? entry.name : 'Турнир');
    let saved = {};
    try {
      const raw = localStorage.getItem(tournamentDataKey(id));
      if (raw) saved = JSON.parse(raw);
    } catch (e) { console.error('load tournament failed', e); }
    setParams({
      totalTeams: 20, fields: 2, startTime: '10:00', endTime: '20:00',
      system: 'auto', groupSize: 4, advance: 2, days: 1,
      scheduleMode: 'sequential', restMode: 'auto', maxGamesPerDay: null,
    drawMode: 'sequential', numSeeds: 4,
    refereeMode: 'manual',
      ...(saved.params || {}),
    });
    setTeamNames(saved.teamNames || {});
    setTeamColors(saved.teamColors || {});
    setScores(saved.scores || {});
    setMatchDurMode(saved.matchDurMode || 'auto');
    setManualDur(saved.manualDur || 40);
    setFieldNames(saved.fieldNames || {});
    setMinRest(saved.minRest != null ? saved.minRest : 0);
    setBlockedSlots(saved.blockedSlots || []);
    setDayWindows(saved.dayWindows || {});
    setDrawOrder(saved.drawOrder || null);
    setRefereeNames(saved.refereeNames || {});
    setMatchReferees(saved.matchReferees || {});
    setOnlineId(saved.onlineId || null);
  };

  const createTournament = () => {
    const id = makeTournamentId();
    const entry = { id, name: `Турнир ${tournamentsIndex.length + 1}`, savedAt: Date.now(), totalTeams: 20, system: 'auto' };
    const next = [...tournamentsIndex, entry];
    setTournamentsIndex(next);
    saveTournamentsIndex(next);
    loadTournamentData(id, next);
    setShowDashboard(false);
    setTab('setup');
  };
  const openTournament = (id) => {
    loadTournamentData(id, tournamentsIndex);
    setShowDashboard(false);
    setTab('setup');
  };
  const deleteTournament = (id) => {
    if (!window.confirm('Удалить этот турнир без возможности восстановления?')) return;
    const next = tournamentsIndex.filter((t) => t.id !== id);
    setTournamentsIndex(next);
    saveTournamentsIndex(next);
    try { localStorage.removeItem(tournamentDataKey(id)); } catch (e) { console.error('delete failed', e); }
    if (id === tournamentId) {
      if (next.length > 0) loadTournamentData(next[next.length - 1].id, next);
      else createTournament();
    }
  };
  const renameTournament = (id, newName) => {
    const trimmed = newName.trim();
    if (!trimmed) return;
    const next = tournamentsIndex.map((t) => (t.id === id ? { ...t, name: trimmed } : t));
    setTournamentsIndex(next);
    saveTournamentsIndex(next);
    if (id === tournamentId) setTournamentName(trimmed);
  };

  // Загрузка списка турниров при монтировании (с разовой миграцией старого формата)
  useEffect(() => {
    let list = migrateLegacyTournament();
    if (list.length === 0) {
      const id = makeTournamentId();
      list = [{ id, name: 'Турнир 1', savedAt: Date.now(), totalTeams: 20, system: 'auto' }];
      saveTournamentsIndex(list);
    }
    setTournamentsIndex(list);
    loadTournamentData(list[list.length - 1].id, list);
  }, []);
  // Сохранение при любом изменении данных ТЕКУЩЕГО турнира
  useEffect(() => {
    if (!tournamentId) return; // ждём завершения начальной загрузки
    try {
      localStorage.setItem(tournamentDataKey(tournamentId), JSON.stringify({ params, teamNames, teamColors, scores, matchDurMode, manualDur, fieldNames, minRest, blockedSlots, dayWindows, drawOrder, refereeNames, matchReferees, onlineId }));
      setTournamentsIndex((prev) => {
        const next = prev.map((t) => (t.id === tournamentId
          ? { ...t, name: tournamentName, savedAt: Date.now(), totalTeams: params.totalTeams, system: params.system }
          : t));
        saveTournamentsIndex(next);
        return next;
      });
    } catch (e) { console.error('save failed', e); }
  }, [tournamentId, params, teamNames, teamColors, scores, matchDurMode, manualDur, fieldNames, minRest, blockedSlots, dayWindows, drawOrder, refereeNames, matchReferees, tournamentName, onlineId]);

  const dayMin = timeToMin(params.endTime) - timeToMin(params.startTime);
  const availMin = Math.max(60, dayMin - 30);
  const reco = useMemo(() => recommend(params.totalTeams, params.fields, availMin), [params.totalTeams, params.fields, params.startTime, params.endTime]);

  const actualSystem = params.system === 'auto' ? reco.system : params.system;
  const actualGroupSize = params.system === 'auto' ? (reco.groupSize || params.groupSize) : params.groupSize;
  const actualAdvance = params.system === 'auto' ? (reco.advance || params.advance) : params.advance;
  // Числовые поля могут быть пустой строкой во время редактирования — фиксируем безопасные значения для расчётов.
  const safeNum = (v, min, max, fallback) => {
    const n = Number(v);
    if (!Number.isFinite(n)) return fallback;
    return Math.max(min, Math.min(max, n));
  };
  const eff = {
    ...params,
    totalTeams: safeNum(params.totalTeams, 4, 512, 4),
    fields: safeNum(params.fields, 1, 20, 1),
    days: safeNum(params.days, 1, 14, 1),
    system: actualSystem, groupSize: actualGroupSize, advance: actualAdvance,
    dayWindows, drawOrder,
    // Служебные — заполним ниже, после расчёта slotDur
    minRestSlots: 0,
    blockedSlotIdxs: [],
  };

  const matches = useMemo(() => buildMatches(eff), [eff.totalTeams, eff.system, eff.groupSize, eff.advance, drawOrder]);
  // Первичное расписание без ограничений — нужно чтобы вычислить slotDur
  const baseSchedule = useMemo(() => scheduleMatches(matches, eff.fields, eff), [matches, eff.fields]);
  const baseStruct = useMemo(() => computeStructure(eff, matches, baseSchedule), [matches, baseSchedule, eff.days, eff.startTime, eff.endTime, eff.fields, eff.scheduleMode, eff.maxGamesPerDay, JSON.stringify(dayWindows)]);
  const slotDur = matchDurMode === 'auto' ? baseStruct.matchDur : manualDur;
  // Перерыв команды применяется только в режиме расписания «с интервалом».
  // Авто — гарантированно ровно 1 слот разрыва (команда не играет два раза подряд).
  const effectiveMinRest = eff.scheduleMode === 'interval'
    ? ((eff.restMode || 'auto') === 'auto' ? slotDur : (minRest || 0))
    : 0;
  const minRestSlots = slotDur > 0 ? Math.ceil((effectiveMinRest || 0) / slotDur) : 0;
  // Заблокированные слоты — превращаем время в глобальный индекс слота
  const slotsPerDay = baseStruct.slotsPerDay;
  const blockedSlotIdxs = useMemo(() => {
    const idxs = new Set();
    (blockedSlots || []).forEach((b) => {
      const day = Math.max(1, b.day || 1);
      const dayInfo = getDayWindow(day, { ...params, dayWindows });
      const dayStartMin = timeToMin(dayInfo.startTime);
      const startLocal = Math.max(0, Math.floor((timeToMin(b.startTime) - dayStartMin) / slotDur));
      const endLocal = Math.max(startLocal + 1, Math.ceil((timeToMin(b.endTime) - dayStartMin) / slotDur));
      const dayOffset = (baseStruct.dayOffsets && baseStruct.dayOffsets[day - 1]) || 0;
      const daySlots = (baseStruct.slotsInDay && baseStruct.slotsInDay[day - 1]) || slotsPerDay;
      for (let s = startLocal; s < endLocal && s < daySlots; s++) {
        idxs.add(dayOffset + s);
      }
    });
    return Array.from(idxs);
  }, [blockedSlots, slotDur, slotsPerDay, params.startTime, JSON.stringify(dayWindows), JSON.stringify(baseStruct.dayOffsets), JSON.stringify(baseStruct.slotsInDay)]);
  // Финальное расписание с ограничениями
  const schedule = useMemo(() =>
    (minRestSlots > 0 || blockedSlotIdxs.length > 0)
      ? scheduleMatches(matches, eff.fields, { ...eff, minRestSlots, blockedSlotIdxs })
      : baseSchedule,
    [matches, eff.fields, minRestSlots, blockedSlotIdxs, baseSchedule]
  );
  const structure = { ...baseStruct, matchDur: slotDur };

  // Отображаемое имя команды
  const teamName = (sid) => teamNames[sid] || `Команда ${sid}`;

  // Группы команд (для системы group/mixed)
  const groups = useMemo(() =>
    (actualSystem === 'group' || actualSystem === 'mixed' || actualSystem === 'mixed-full' || actualSystem === 'mixed-goldsilver' || actualSystem === 'mixed-goldsilver-full')
      ? splitIntoGroups(eff.totalTeams, eff.groupSize, drawOrder)
      : [],
    [actualSystem, eff.totalTeams, eff.groupSize, drawOrder]
  );
  const sidOf = (gi, pi) => (groups[gi] && groups[gi][pi - 1]) || 0;

  // Расчёт турнирной таблицы группы
  const computeStandings = (gi) => {
    const teams = groups[gi] || [];
    const table = teams.map((sid) => ({ sid, name: teamName(sid), gz: 0, gp: 0, pts: 0, w: 0, d: 0, l: 0, played: 0 }));
    matches.filter((m) => m.phase === 'group' && m.group === gi + 1).forEach((m) => {
      const sc = scores[m.id];
      if (!sc || sc.a == null || sc.b == null) return;
      const sd1 = sidOf(gi, parseInt(m.t1.split('.')[1]));
      const sd2 = sidOf(gi, parseInt(m.t2.split('.')[1]));
      const r1 = table.find((x) => x.sid === sd1);
      const r2 = table.find((x) => x.sid === sd2);
      if (!r1 || !r2) return;
      r1.gz += sc.a; r1.gp += sc.b; r1.played++;
      r2.gz += sc.b; r2.gp += sc.a; r2.played++;
      if (sc.a > sc.b) { r1.pts += 3; r1.w++; r2.l++; }
      else if (sc.a < sc.b) { r2.pts += 3; r2.w++; r1.l++; }
      else { r1.pts++; r2.pts++; r1.d++; r2.d++; }
    });
    // Сортировка: очки → разница → забитые → номер
    table.sort((a, b) => (b.pts - a.pts) || ((b.gz - b.gp) - (a.gz - a.gp)) || (b.gz - a.gz) || (a.sid - b.sid));
    table.forEach((t, i) => { t.place = i + 1; });
    return table;
  };

  const allStandings = useMemo(() => groups.map((_, gi) => computeStandings(gi)), [groups, matches, scores, teamNames]);

  // Определение победителя матча плей-офф
  const matchWinner = (matchId) => {
    const sc = scores[matchId];
    if (!sc || sc.a == null || sc.b == null || sc.a === sc.b) return null;
    return sc.a > sc.b ? 't1' : 't2';
  };

  // Резолвит любой маркер команды в матче плей-офф в реальную команду {sid, name}
  // matchBracket — 'gold'|'silver'|null; проставляется вызывающей стороной из m.bracket
  // (для рекурсивных Поб./Проигр.-ссылок — из bracket НАЙДЕННОГО матча, не текущего).
  const resolveSlot = (marker, matchBracket) => {
    if (!marker || marker === 'BYE') return null;
    // СИД{N} — распределение семян
    const seedMatch = marker.match(/^СИД(\d+)$/);
    if (seedMatch) {
      const sd = +seedMatch[1];
      if (matchBracket === 'gold' || matchBracket === 'silver') {
        // Золото/серебро: ровно один сеяный на группу — сид N это группа N,
        // место — 1-е для золота, 2-е для серебра.
        const st = allStandings[sd - 1];
        if (!st) return null;
        const row = st.find((r) => r.place === (matchBracket === 'gold' ? 1 : 2));
        return row ? { sid: row.sid, name: row.name } : null;
      }
      if (actualSystem === 'mixed' || actualSystem === 'mixed-full') {
        const ng = groups.length;
        const placeIdx = Math.ceil(sd / ng);
        const grIdx = ((sd - 1) % ng);
        const st = allStandings[grIdx];
        if (!st) return null;
        const row = st.find((r) => r.place === placeIdx);
        return row ? { sid: row.sid, name: row.name } : null;
      }
      return { sid: sd, name: teamName(sd) };
    }
    // Поб.М{id} — победитель матча
    const winMatch = marker.match(/^Поб\.М(\d+)$/);
    if (winMatch) {
      const m = matches.find((x) => x.id === +winMatch[1]);
      if (!m) return null;
      const w = matchWinner(m.id);
      if (!w) return null;
      return resolveSlot(w === 't1' ? m.t1 : m.t2, m.bracket);
    }
    // Проигр.М{id} — проигравший
    const loseMatch = marker.match(/^Проигр\.М(\d+)$/);
    if (loseMatch) {
      const m = matches.find((x) => x.id === +loseMatch[1]);
      if (!m) return null;
      const w = matchWinner(m.id);
      if (!w) return null;
      return resolveSlot(w === 't1' ? m.t2 : m.t1, m.bracket);
    }
    return null;
  };

  // Метка команды в матче с учётом текущих результатов
  const teamLabel = (m, side) => {
    const raw = side === 't1' ? m.t1 : m.t2;
    if (m.phase === 'group') {
      const gi = m.group - 1;
      const pos = parseInt(raw.split('.')[1]);
      const sd = sidOf(gi, pos);
      return teamName(sd);
    }
    // playoff — сначала пробуем резолвить, если не выходит показываем raw-плейсхолдер
    const resolved = resolveSlot(raw, m.bracket);
    if (resolved) return resolved.name;
    // Показываем понятный плейсхолдер
    const seed = raw.match(/^СИД(\d+)$/);
    if (seed && (m.bracket === 'gold' || m.bracket === 'silver')) {
      return `${m.bracket === 'gold' ? '1-е' : '2-е'} Гр.${+seed[1]}`;
    }
    if (seed && (actualSystem === 'mixed' || actualSystem === 'mixed-full')) {
      const sd = +seed[1];
      const ng = groups.length;
      const placeIdx = Math.ceil(sd / ng);
      const grIdx = ((sd - 1) % ng) + 1;
      const ord = ['', '1-е', '2-е', '3-е', '4-е', '5-е', '6-е', '7-е', '8-е'][placeIdx] || `${placeIdx}-е`;
      return `${ord} Гр.${grIdx}`;
    }
    if (raw.startsWith('Поб.М')) return raw.replace('Поб.М', 'Поб. м.');
    if (raw.startsWith('Проигр.М')) return raw.replace('Проигр.М', 'Проигр. м.');
    return raw;
  };

  const slotMap = {};
  schedule.forEach((s) => {
    if (!slotMap[s.slotIdx]) slotMap[s.slotIdx] = {};
    slotMap[s.slotIdx][s.field] = matches.find((x) => x.id === s.matchId);
  });
  const sortedSlots = Object.keys(slotMap).map(Number).sort((a, b) => a - b);
  const totalUsedSlots = sortedSlots.length;
  const dayOffsets = structure.dayOffsets || [0];
  const slotsInDay = structure.slotsInDay || [slotsPerDay];
  const dayInfos = structure.dayInfos || [{ startTime: params.startTime }];
  // Сколько дней реально нужно (по последнему слоту)
  const lastSlotIdx = sortedSlots.length ? sortedSlots[sortedSlots.length - 1] : 0;
  const lastPos = slotToDay(lastSlotIdx, dayOffsets, slotsInDay);
  const actualDaysNeeded = sortedSlots.length ? lastPos.day : 1;
  // В режиме «по дням» число дней подбирается автоматически — там всегда «помещается».
  const declaredDays = eff.scheduleMode === 'byDay' ? structure.days : params.days;
  const fits = actualDaysNeeded <= declaredDays;
  // Время окончания последнего матча в его дне
  const lastDayStart = timeToMin((dayInfos[lastPos.day - 1] || dayInfos[0]).startTime);
  const finishMin = lastDayStart + (lastPos.local + 1) * slotDur;
  const finishTime = minToTime(finishMin);

  const handleDownload = () => {
    try {
      const varRows = [];
      matches.forEach((m) => {
        const events = scores[m.id] && scores[m.id].events;
        if (!events || !events.length) return;
        const matchLabel = m.phase === 'group' ? m.label : (m.roundName + (m.isBronze ? ' (бронза)' : ''));
        events.forEach((e) => {
          varRows.push({
            matchLabel,
            minute: e.minute != null ? e.minute : null,
            team: e.team === 'a' ? teamLabel(m, 't1') : e.team === 'b' ? teamLabel(m, 't2') : '—',
            type: eventLabel(e.type),
            note: e.note || '',
          });
        });
      });
      generateXLSX(eff, structure, matches, schedule, slotDur, fieldNames, varRows);
    }
    catch (e) { alert('Ошибка генерации: ' + e.message); console.error(e); }
  };

  // Публикует/переопубликовывает турнир на бэкенде синхронизации — после этого
  // QR судьи и ссылка для родителей ведут туда, а не в localStorage. Метки
  // команд (t1Label/t2Label) считаются здесь один раз при публикации: бэкенд
  // не знает, как резолвить победителей плей-офф — организатору нужно
  // переопубликовать турнир после завершения раунда, чтобы обновить метки.
  const publishOnline = async () => {
    setPublishing(true);
    try {
      const publishMatches = matches.map((m) => ({
        id: m.id,
        label: m.phase === 'group' ? m.label : (m.roundName + (m.isBronze ? ' (бронза)' : '')),
        t1Label: teamLabel(m, 't1'),
        t2Label: teamLabel(m, 't2'),
        phase: m.phase,
        group: m.group,
        roundName: m.roundName,
        isBronze: m.isBronze,
      }));
      const payload = { name: tournamentName, params: eff, teamNames, teamColors, matches: publishMatches, structure };
      const url = onlineId ? `${SYNC_BACKEND_URL}/api/tournaments/${onlineId}` : `${SYNC_BACKEND_URL}/api/tournaments`;
      const res = await fetch(url, {
        method: onlineId ? 'PUT' : 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(payload),
      });
      const data = await res.json();
      if (data.ok) setOnlineId(data.id || onlineId);
      else alert('Не удалось опубликовать турнир: ' + (data.error || 'ошибка сервера'));
    } catch (e) {
      alert('Не удалось опубликовать турнир — проверьте интернет-соединение.');
      console.error(e);
    } finally {
      setPublishing(false);
    }
  };

  // Пока турнир онлайн — держим WebSocket к бэкенду, чтобы счёт от судей на
  // других телефонах сразу применялся здесь (и попадал в обычный localStorage
  // автосейв — офлайн-копия остаётся источником для xlsx/протокола).
  useEffect(() => {
    if (!onlineId) return;
    const wsUrl = SYNC_BACKEND_URL.replace(/^http/, 'ws') + `/api/tournaments/${onlineId}/ws`;
    const ws = new WebSocket(wsUrl);
    ws.addEventListener('message', (event) => {
      try {
        const msg = JSON.parse(event.data);
        if (msg.type === 'score') {
          setScores((prev) => ({ ...prev, [msg.matchId]: { a: msg.a, b: msg.b, events: msg.events || [] } }));
        }
      } catch (e) { console.error('ws message parse failed', e); }
    });
    return () => ws.close();
  }, [onlineId]);

  // Организатор тоже вводит счёт локально (в этом же приложении) — если турнир
  // онлайн, отправляем и туда, чтобы судьи/родители видели изменения от
  // организатора так же, как и наоборот. Fire-and-forget: не блокируем UI.
  const syncScoreOnline = (matchId, a, b, events) => {
    if (!onlineId) return;
    fetch(`${SYNC_BACKEND_URL}/api/tournaments/${onlineId}/matches/${matchId}/score`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ a, b, events }),
    }).catch((e) => console.error('online score sync failed', e));
  };

  const sysLabel = (s) => SYS_LABELS[s] || s;
  const isMixedFamily = (s) => s === 'mixed' || s === 'mixed-full' || s === 'mixed-goldsilver' || s === 'mixed-goldsilver-full';
  const sysColor = (s) => (
    s === 'group' ? 'bg-[#f5f2ec] text-black border border-black/10'
    : s === 'playoff' || s === 'playoff-full' ? 'bg-[#e30613]/10 text-[#b1040f] border border-[#e30613]/25'
    : isMixedFamily(s) ? 'bg-black text-white'
    : 'bg-[#f5f2ec]'
  );
 const durStatus = slotDur < 25 ? 'tight' : slotDur > 70 ? 'loose' : 'ok';

  // === СУДЕЙСКИЙ РЕЖИМ ===
  // Если URL содержит ?judge=matchId — показываем только упрощённый экран для судьи
  const judgeMatchId = getJudgeMatchId();
  if (judgeMatchId != null) {
    const m = matches.find((x) => x.id === judgeMatchId);
    if (m) {
      const t1Label = teamLabel(m, 't1');
      const t2Label = teamLabel(m, 't2');
      // sid для цветовых полос
      let sid1 = null, sid2 = null;
      if (m.phase === 'group') {
        const gi = m.group - 1;
        sid1 = sidOf(gi, parseInt(m.t1.split('.')[1]));
        sid2 = sidOf(gi, parseInt(m.t2.split('.')[1]));
      } else {
        const r1 = resolveSlot(m.t1, m.bracket);
        const r2 = resolveSlot(m.t2, m.bracket);
        sid1 = r1 ? r1.sid : null;
        sid2 = r2 ? r2.sid : null;
      }
      return <JudgeView matchLabel={m.phase === 'group' ? m.label : (m.roundName + (m.isBronze ? ' (бронза)' : ''))}
        t1Label={t1Label} t2Label={t2Label} color1={sid1 ? teamColors[sid1] : null} color2={sid2 ? teamColors[sid2] : null}
        existing={scores[m.id]}
        onSave={(a, b, events) => { setScores({ ...scores, [m.id]: { a, b, events } }); syncScoreOnline(m.id, a, b, events); }} />;
    }
  }

 return (
 <div className="min-h-screen bg-white pb-24 lg:pb-6 text-[#0c0c0c]" style={{ fontFamily:"'Inter', system-ui, sans-serif" }}>
 {/* MSG-заголовок с логотипной плашкой */}
 <div className="border-b border-black/10 bg-white sticky top-0 z-10">
 <div className="max-w-6xl mx-auto px-3 sm:px-6 py-3 flex items-center gap-3">
 <div className="flex items-baseline gap-2 leading-none">
 <span className="text-lg sm:text-xl font-black tracking-tight">MITIN SPORT</span>
 <span className="text-[10px] sm:text-xs font-bold text-white bg-[#e30613] px-1.5 py-0.5 rounded-sm">GROUP</span>
 </div>
 <div className="hidden sm:block h-6 w-px bg-black/15" />
 <div className="text-xs sm:text-sm text-neutral-500 font-medium hidden sm:block">Конструктор турниров</div>
 <div className="flex-1" />
 {onlineId ? (
   <div className="flex items-center gap-1.5 flex-shrink-0">
     <span className="text-[10px] font-bold uppercase tracking-widest text-emerald-600 hidden sm:inline">🟢 Онлайн</span>
     <button onClick={() => { navigator.clipboard.writeText(`${SYNC_BACKEND_URL}/view/${onlineId}`); alert('Ссылка для родителей скопирована'); }}
       className="px-2.5 py-1.5 text-xs font-bold text-neutral-600 hover:text-[#0c0c0c] border border-black/10 rounded-sm">👀 Родителям</button>
     <button onClick={publishOnline} disabled={publishing}
       className="px-2.5 py-1.5 text-xs font-bold text-neutral-600 hover:text-[#0c0c0c] border border-black/10 rounded-sm disabled:opacity-50" title="Обновить составы/сетку на сервере после нового раунда">
       {publishing ? '…' : '↻ Обновить'}
     </button>
   </div>
 ) : (
   <button onClick={publishOnline} disabled={publishing}
     className="px-2.5 py-1.5 text-xs font-bold text-white bg-[#e30613] hover:bg-[#b1040f] rounded-sm disabled:opacity-50 flex-shrink-0">
     {publishing ? 'Публикация…' : '🌐 Опубликовать онлайн'}
   </button>
 )}
 <button onClick={() => setShowDashboard(true)} className="flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-bold text-neutral-600 hover:text-[#0c0c0c] border border-black/10 rounded-sm flex-shrink-0">
   <span>📁</span><span className="max-w-[8rem] sm:max-w-[10rem] truncate">{tournamentName}</span>
 </button>
 </div>
 </div>
 <div className="max-w-6xl mx-auto p-3 sm:p-6">
 <div className="mb-5 sm:mb-8">
 <div className="text-[10px] sm:text-xs font-bold tracking-widest text-[#e30613] mb-2">СПОРТ ПОД КОНТРОЛЕМ</div>
 <h1 className="text-2xl sm:text-4xl font-black tracking-tight text-[#0c0c0c] leading-[1.1]">
 Конструктор<br className="sm:hidden" /> турниров
 </h1>
 <p className="text-neutral-600 mt-2 text-sm sm:text-base max-w-xl">Задайте параметры, впишите команды, ведите счёт в приложении — Excel на выходе</p>
 </div>

 {/* Вкладки */}
 <div className="flex border-b border-black/10 mb-5 sm:mb-6 -mx-3 sm:mx-0 px-3 sm:px-0 overflow-x-auto">
  <button onClick={() => setTab('setup')} className={`px-3 sm:px-4 py-2 sm:py-3 text-xs sm:text-sm font-bold uppercase tracking-widest whitespace-nowrap border-b-2 transition ${tab === 'setup' ? 'border-[#e30613] text-[#0c0c0c]' : 'border-transparent text-neutral-500 hover:text-[#0c0c0c]'}`}>
   Параметры
  </button>
  <button onClick={() => setTab('tournament')} className={`px-3 sm:px-4 py-2 sm:py-3 text-xs sm:text-sm font-bold uppercase tracking-widest whitespace-nowrap border-b-2 transition ${tab === 'tournament' ? 'border-[#e30613] text-[#0c0c0c]' : 'border-transparent text-neutral-500 hover:text-[#0c0c0c]'}`}>
   Турнир
   {Object.keys(scores).length > 0 && <span className="ml-1.5 inline-block bg-[#e30613] text-white text-[10px] px-1.5 py-0.5 rounded-sm">{Object.keys(scores).length}</span>}
  </button>
  <button onClick={() => setTab('results')} className={`px-3 sm:px-4 py-2 sm:py-3 text-xs sm:text-sm font-bold uppercase tracking-widest whitespace-nowrap border-b-2 transition ${tab === 'results' ? 'border-[#e30613] text-[#0c0c0c]' : 'border-transparent text-neutral-500 hover:text-[#0c0c0c]'}`}>
   Итоги
  </button>
 </div>

 {tab === 'setup' && (
 <div className="grid grid-cols-1 lg:grid-cols-3 gap-4 lg:gap-5">
 {/* ФОРМА */}
 <div className="lg:col-span-1 space-y-4">
 <div className="bg-white rounded border border-black/10 p-4 sm:p-5">
 <h2 className="text-xs font-black text-[#0c0c0c] mb-4 uppercase tracking-widest flex items-center gap-2"><Settings className="w-5 h-5" />Параметры</h2>
 <div className="space-y-3">
 <Field label="Команд (4–512)">
 <BoundedNumber value={params.totalTeams} min={4} max={512}
 onCommit={(n) => setParams({ ...params, totalTeams: n })} />
 </Field>
 <div className="grid grid-cols-2 gap-2">
 <Field label="Полей">
 <BoundedNumber value={params.fields} min={1} max={20}
 onCommit={(n) => setParams({ ...params, fields: n })} />
 </Field>
 {params.scheduleMode === 'byDay' ? (
 <Field label="Дней (авто)">
 <div className="inp bg-[#f5f2ec] text-neutral-500 flex items-center">{declaredDays}</div>
 </Field>
 ) : (
 <Field label="Дней">
 <BoundedNumber value={params.days} min={1} max={14}
 onCommit={(n) => setParams({ ...params, days: n })} />
 </Field>
 )}
 </div>
 <div className="grid grid-cols-2 gap-2">
 <Field label="Начало"><input type="time" value={params.startTime} onChange={(e) => setParams({ ...params, startTime: e.target.value })} className="inp" /></Field>
 <Field label="Крайний срок"><input type="time" value={params.endTime} onChange={(e) => setParams({ ...params, endTime: e.target.value })} className="inp" /></Field>
 </div>
 <div className="text-xs text-neutral-500">Окно: {Math.floor(dayMin / 60)} ч {dayMin % 60} м · чистой игры ≈ {availMin} м</div>

 <Field label="Расписание">
 <select value={params.scheduleMode || 'sequential'} onChange={(e) => setParams({ ...params, scheduleMode: e.target.value })} className="inp">
 <option value="sequential">Последовательное</option>
 <option value="interval">С интервалом (не подряд)</option>
 <option value="byDay">По дням (лимит игр в день)</option>
 </select>
 </Field>
 {params.scheduleMode === 'interval' && (
 <Field label="Интервал между матчами команды">
 <div className="flex gap-2 items-center">
 <select value={params.restMode || 'auto'} onChange={(e) => setParams({ ...params, restMode: e.target.value })} className="inp flex-1">
 <option value="auto">Авто (не подряд)</option>
 <option value="manual">Вручную</option>
 </select>
 {(params.restMode || 'auto') === 'manual' && (
 <BoundedNumber value={minRest} min={0} max={120} onCommit={setMinRest} />
 )}
 </div>
 </Field>
 )}
 {params.scheduleMode === 'byDay' && (
 <Field label="Матчей в день">
 <div className="flex gap-2 items-center">
 <BoundedNumber value={params.maxGamesPerDay || Math.max(1, params.fields * 6)} min={1} max={200}
 onCommit={(n) => setParams({ ...params, maxGamesPerDay: n })} />
 <button type="button" onClick={() => setParams({ ...params, maxGamesPerDay: Math.max(1, params.fields * 6) })}
 className="text-[10px] font-bold uppercase tracking-widest text-neutral-400 hover:text-[#e30613] whitespace-nowrap">Авто</button>
 </div>
 </Field>
 )}

 <Field label="Система">
 <select value={params.system} onChange={(e) => setParams({ ...params, system: e.target.value })} className="inp">
 <option value="auto">⚡ Авто (рекомендация)</option>
 <option value="group">Групповая</option>
 <option value="playoff">Плей-офф (олимпийка)</option>
 <option value="playoff-full">Плей-офф (розыгрыш всех мест)</option>
 <option value="mixed">Смешанная</option>
 <option value="mixed-full">Смешанная (розыгрыш всех мест)</option>
 <option value="mixed-goldsilver">Смешанная (золото/серебро, на вылет)</option>
 <option value="mixed-goldsilver-full">Смешанная (золото/серебро, все места)</option>
 </select>
 </Field>

 {(actualSystem === 'group' || actualSystem === 'mixed' || actualSystem === 'mixed-full' || actualSystem === 'mixed-goldsilver' || actualSystem === 'mixed-goldsilver-full') && (
 <Field label="Команд в группе">
 <select value={actualGroupSize} disabled={params.system === 'auto'} onChange={(e) => setParams({ ...params, groupSize: +e.target.value })} className="inp disabled:bg-[#f5f2ec] disabled:text-neutral-500">
 {[3,4,5,6,7,8,9,10,11,12,13,14,15,16].map((n) => <option key={n} value={n}>{n}</option>)}
 </select>
 </Field>
 )}
 {(actualSystem === 'group' || actualSystem === 'mixed' || actualSystem === 'mixed-full' || actualSystem === 'mixed-goldsilver' || actualSystem === 'mixed-goldsilver-full') && (
 <Field label="Жеребьёвка">
 <div className="flex gap-2 items-center">
 <select value={params.drawMode || 'sequential'} onChange={(e) => setParams({ ...params, drawMode: e.target.value })} className="inp flex-1">
 <option value="sequential">Последовательная</option>
 <option value="random">Случайная</option>
 <option value="seeded">С учётом сеяных команд</option>
 </select>
 <button type="button" onClick={() => setDrawOrder(generateDrawOrder(params.drawMode || 'sequential', eff.totalTeams, actualGroupSize, params.numSeeds))}
 className="px-3 py-2 text-xs font-bold uppercase tracking-widest text-white bg-[#e30613] hover:bg-[#b1040f] whitespace-nowrap">🎲 Провести</button>
 </div>
 {(params.drawMode || 'sequential') === 'seeded' && (
 <div className="mt-2">
 <BoundedNumber value={params.numSeeds} min={1} max={Math.max(1, Math.floor(eff.totalTeams / actualGroupSize))}
 onCommit={(n) => setParams({ ...params, numSeeds: n })} />
 <div className="text-[10px] text-neutral-400 mt-1">Сколько сеяных команд — по одной на группу, остальные вразброс</div>
 </div>
 )}
 {drawOrder && <div className="text-[10px] text-neutral-400 mt-1">Жеребьёвка проведена — состав групп зафиксирован ниже</div>}
 </Field>
 )}
 {(actualSystem === 'mixed' || actualSystem === 'mixed-full') && (
 <Field label="Из группы в плей-офф">
 <select value={actualAdvance} disabled={params.system === 'auto'} onChange={(e) => setParams({ ...params, advance: +e.target.value })} className="inp disabled:bg-[#f5f2ec] disabled:text-neutral-500">
 {Array.from({length:16},(_,i)=>i+1).filter((n) => n <= actualGroupSize).map((n) => <option key={n} value={n}>топ-{n}</option>)}
 </select>
 </Field>
 )}

 {/* Время матча */}
 <Field label="Длительность матча/слота">
 <div className="flex gap-2 items-center">
 <select value={matchDurMode} onChange={(e) => setMatchDurMode(e.target.value)} className="inp flex-1">
 <option value="auto">Авто ({baseStruct.matchDur} мин)</option>
 <option value="manual">Вручную</option>
 </select>
 {matchDurMode === 'manual' && (
 <input type="number" min="10" max="120" value={manualDur}
 onChange={(e) => setManualDur(e.target.value === '' ? '' : Number(e.target.value))}
 onBlur={(e) => setManualDur(Math.max(10, Math.min(120, Number(e.target.value) || 10)))}
 className="inp w-20" />
 )}
 </div>
 </Field>
 </div>

 {params.system === 'auto' && (
 <div className="mt-4 p-3 rounded bg-[#e30613]/5 border border-[#e30613]">
 <div className="text-xs font-semibold text-[#e30613] mb-1 flex items-center gap-1"><Sparkles className="w-3 h-3" />Рекомендация</div>
 <div className="text-sm text-[#e30613]">
 <strong>{sysLabel(reco.system)}</strong>
 {reco.groupSize > 0 && `, группы по ${reco.groupSize}`}
 {reco.advance > 0 && `, топ-${reco.advance}`}
 </div>
 </div>
 )}
 </div>

 {/* Расширенные настройки */}
 <div className="bg-white border border-black/10 p-4 sm:p-5">
  <h2 className="text-xs font-black text-[#0c0c0c] mb-4 uppercase tracking-widest">Расширенные</h2>
  <div className="space-y-3">
   {eff.days > 1 && (
   <div>
    <div className="text-[10px] font-bold uppercase tracking-widest text-neutral-500 mb-1.5">Своё окно для дней (опционально)</div>
    <div className="space-y-1.5">
     {Array.from({ length: eff.days }, (_, i) => i + 1).map((d) => {
      const w = dayWindows[d] || {};
      const isOverride = !!(w.startTime || w.endTime);
      return (
       <div key={d} className="flex items-center gap-1.5 text-xs">
        <div className={`w-10 flex-shrink-0 text-[10px] font-black uppercase ${isOverride ? 'text-[#e30613]' : 'text-neutral-400'}`}>Д{d}</div>
        <input type="time" value={w.startTime || params.startTime}
         onChange={(e) => { const next = { ...dayWindows }; next[d] = { ...next[d], startTime: e.target.value }; setDayWindows(next); }}
         className="inp text-xs flex-1" />
        <span className="text-neutral-400">–</span>
        <input type="time" value={w.endTime || params.endTime}
         onChange={(e) => { const next = { ...dayWindows }; next[d] = { ...next[d], endTime: e.target.value }; setDayWindows(next); }}
         className="inp text-xs flex-1" />
        {isOverride && (
         <button onClick={() => { const next = { ...dayWindows }; delete next[d]; setDayWindows(next); }} className="w-6 h-6 flex-shrink-0 text-neutral-400 hover:text-[#e30613] text-lg">↺</button>
        )}
       </div>
      );
     })}
     <div className="text-[10px] text-neutral-400 leading-tight">Если поле не изменено — используется общее «Начало/Крайний срок». Красная Д = переопределено.</div>
    </div>
   </div>
   )}
   <div>
    <div className="text-[10px] font-bold uppercase tracking-widest text-neutral-500 mb-1.5">Названия площадок</div>
    <div className="space-y-1.5">
     {Array.from({ length: eff.fields }, (_, i) => i + 1).map((n) => (
      <input key={n} type="text" defaultValue={fieldNames[n] || ''} placeholder={`Поле ${n}`}
       onBlur={(e) => {
        const v = e.target.value.trim();
        const next = { ...fieldNames };
        if (v) next[n] = v; else delete next[n];
        setFieldNames(next);
       }}
       className="inp text-sm" />
     ))}
    </div>
   </div>
   <div>
    <div className="text-[10px] font-bold uppercase tracking-widest text-neutral-500 mb-1.5">Заблокированные слоты (обед, церемония)</div>
    <div className="space-y-2">
     {(blockedSlots || []).map((b, i) => (
      <div key={i} className="flex items-center gap-1.5 text-xs">
       <select value={b.day} onChange={(e) => { const next = [...blockedSlots]; next[i] = { ...next[i], day: +e.target.value }; setBlockedSlots(next); }} className="inp w-14 text-xs">
        {Array.from({ length: eff.days }, (_, d) => d + 1).map((d) => <option key={d} value={d}>Д{d}</option>)}
       </select>
       <input type="time" value={b.startTime} onChange={(e) => { const next = [...blockedSlots]; next[i] = { ...next[i], startTime: e.target.value }; setBlockedSlots(next); }} className="inp text-xs flex-1" />
       <input type="time" value={b.endTime} onChange={(e) => { const next = [...blockedSlots]; next[i] = { ...next[i], endTime: e.target.value }; setBlockedSlots(next); }} className="inp text-xs flex-1" />
       <button onClick={() => setBlockedSlots(blockedSlots.filter((_, k) => k !== i))} className="w-8 h-8 flex-shrink-0 text-neutral-400 hover:text-[#e30613]">×</button>
      </div>
     ))}
     <button onClick={() => setBlockedSlots([...(blockedSlots || []), { day: 1, startTime: '13:00', endTime: '14:00', label: '' }])}
      className="w-full py-2 text-[10px] font-bold uppercase tracking-widest border border-dashed border-black/20 text-neutral-500 hover:border-[#e30613] hover:text-[#e30613]">
      + Добавить перерыв
     </button>
    </div>
   </div>
   <div>
    <div className="text-[10px] font-bold uppercase tracking-widest text-neutral-500 mb-1.5">Судьи</div>
    <select value={params.refereeMode || 'manual'} onChange={(e) => setParams({ ...params, refereeMode: e.target.value })} className="inp text-sm mb-2">
     <option value="manual">Вручную (текст у матча)</option>
     <option value="fromList">Из списка</option>
     <option value="randomFromList">Рандомно из списка</option>
    </select>
    {(params.refereeMode || 'manual') !== 'manual' && (
     <div className="space-y-1.5">
      {Object.keys(refereeNames).map(Number).sort((a, b) => a - b).map((rid) => (
       <div key={rid} className="flex items-center gap-1.5">
        <input type="text" defaultValue={refereeNames[rid]} placeholder="Имя судьи"
         onBlur={(e) => {
          const v = e.target.value.trim();
          const next = { ...refereeNames };
          if (v) next[rid] = v; else delete next[rid];
          setRefereeNames(next);
         }}
         className="inp text-sm flex-1" />
        <button onClick={() => { const next = { ...refereeNames }; delete next[rid]; setRefereeNames(next); }}
         className="w-8 h-8 flex-shrink-0 text-neutral-400 hover:text-[#e30613]">×</button>
       </div>
      ))}
      <button onClick={() => {
        const ids = Object.keys(refereeNames).map(Number);
        const nextId = (ids.length ? Math.max(...ids) : 0) + 1;
        setRefereeNames({ ...refereeNames, [nextId]: '' });
       }}
       className="w-full py-2 text-[10px] font-bold uppercase tracking-widest border border-dashed border-black/20 text-neutral-500 hover:border-[#e30613] hover:text-[#e30613]">
       + Добавить судью
      </button>
      {(params.refereeMode || 'manual') === 'randomFromList' && (
       <button onClick={() => {
         const refIds = Object.keys(refereeNames).map(Number);
         if (refIds.length === 0) { alert('Сначала добавьте судей в список'); return; }
         const bySlot = {};
         schedule.forEach((s) => { (bySlot[s.slotIdx] = bySlot[s.slotIdx] || []).push(s.matchId); });
         const next = {};
         Object.keys(bySlot).forEach((slotIdx) => {
          const pool = shuffleArr(refIds);
          bySlot[slotIdx].forEach((mid, i) => { next[mid] = pool[i % pool.length]; });
         });
         setMatchReferees(next);
        }}
        className="w-full py-2 mt-1 text-xs font-bold uppercase tracking-widest text-white bg-[#e30613] hover:bg-[#b1040f]">
        🎲 Назначить судей
       </button>
      )}
     </div>
    )}
   </div>
  </div>
 </div>

 <button onClick={handleDownload} className="hidden lg:flex w-full bg-[#e30613] hover:bg-[#b1040f] text-white font-semibold py-3 px-4 rounded items-center justify-center gap-2 transition">
 <Download className="w-5 h-5" />Скачать xlsx
 </button>
 </div>

 {/* ПРАВО */}
 <div className="lg:col-span-2 space-y-4">
 <div className="bg-white rounded border border-black/10 p-5">
 <h2 className="text-xs font-black text-[#0c0c0c] mb-4 uppercase tracking-widest">Структура</h2>
 <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
 <Metric label="Система" value={sysLabel(actualSystem)} colorClass={sysColor(actualSystem)} inverted={isMixedFamily(actualSystem)} />
 <Metric label="Матчей" value={matches.length} />
 <Metric label="Слотов" value={totalUsedSlots} />
 <Metric label="Слот" value={`${slotDur} мин`} warning={durStatus !== 'ok'} />
 {structure.numGroups > 0 && <Metric label="Групп" value={structure.numGroups} />}
 {structure.playoffTeams > 0 && <Metric label="В плей-офф" value={structure.playoffTeams} />}
 <Metric label="Нужно дней" value={`${actualDaysNeeded} из ${declaredDays}`} warning={!fits} />
 <Metric label="Слотов в день" value={slotsPerDay} />
 <Metric label={actualDaysNeeded > 1 ? `Финиш (день ${actualDaysNeeded})` : 'Финиш'} value={finishTime} warning={finishMin > timeToMin(params.endTime)} />
 </div>
 {!fits && (
 <div className="mt-4 p-3 bg-[#e30613]/5 border border-[#e30613]/25 rounded text-sm text-[#b1040f] flex items-start gap-2">
 <AlertCircle className="w-4 h-4 mt-0.5 flex-shrink-0" />
 <div>Не помещается: нужно <strong>{actualDaysNeeded} дней</strong>, задано <strong>{declaredDays}</strong>. Увеличьте дней, полей или растяните дневное окно. Файл всё равно скачается — расписание просто длиннее.</div>
 </div>
 )}
 {durStatus === 'tight' && fits && (
 <div className="mt-4 p-3 bg-[#e30613]/5 border border-[#e30613]/25 rounded text-sm text-[#b1040f] flex items-start gap-2">
 <AlertCircle className="w-4 h-4 mt-0.5 flex-shrink-0" /><div>Слот менее 25 мин — плотно. Добавьте полей или времени.</div>
 </div>
 )}
 </div>

 <div className="bg-white rounded border border-black/10 p-5">
 <h2 className="text-xs font-black text-[#0c0c0c] mb-4 uppercase tracking-widest flex items-center gap-2"><Calendar className="w-5 h-5" />Превью расписания</h2>
 <div className="overflow-x-auto">
 <table className="w-full text-sm">
 <thead>
 <tr className="border-b border-black/10">
 <th className="text-left py-2 px-2 font-semibold text-neutral-700">День</th>
 <th className="text-left py-2 px-2 font-semibold text-neutral-700">№</th>
 <th className="text-left py-2 px-2 font-semibold text-neutral-700">Время</th>
 {Array.from({ length: eff.fields }, (_, i) => <th key={i} className="text-left py-2 px-2 font-semibold text-neutral-700">{fieldNames[i + 1] || `Поле ${i + 1}`}</th>)}
 </tr>
 </thead>
 <tbody>
 {sortedSlots.slice(0, previewSlots).map((sIdx, i) => {
 const pos = slotToDay(sIdx, dayOffsets, slotsInDay);
 const dayStart = timeToMin((dayInfos[pos.day - 1] || dayInfos[0]).startTime);
 return (
 <tr key={sIdx} className="border-b border-slate-100 hover:bg-[#f5f2ec]">
 <td className="py-2 px-2 font-medium text-neutral-500">Д{pos.day}</td>
 <td className="py-2 px-2 font-medium">{pos.local + 1}</td>
 <td className="py-2 px-2 font-mono text-neutral-700 text-xs sm:text-sm whitespace-nowrap">{minToTime(dayStart + pos.local * slotDur)}–{minToTime(dayStart + (pos.local + 1) * slotDur)}</td>
 {Array.from({ length: eff.fields }, (_, f) => {
 const m = slotMap[sIdx][f + 1];
 const po = m && m.phase === 'playoff';
 return (
 <td key={f} className={`py-2 px-2 ${po ? 'text-[#b1040f] bg-[#e30613]/5' : 'text-[#0c0c0c]'}`}>
 {m ? (<><div className="text-xs text-neutral-500">{m.label}</div><div>{m.t1} — {m.t2}</div></>) : <span className="text-neutral-300">—</span>}
 </td>
 );
 })}
 </tr>
 );
 })}
 </tbody>
 </table>
 </div>
 {sortedSlots.length > previewSlots && (
 <button className="mt-3 text-sm text-[#e30613] hover:text-[#e30613]" onClick={() => setPreviewSlots(previewSlots + 10)}>
 Показать ещё ({sortedSlots.length - previewSlots})
 </button>
 )}
 </div>

 <div className="bg-[#f5f2ec] rounded p-4 text-sm text-neutral-700">
 <div className="font-semibold mb-2">В xlsx (всё считается формулами):</div>
 <div className="space-y-1">
 <div>📋 <strong>Параметры</strong> + инструкция</div>
 <div>📅 <strong>Расписание</strong> по полям</div>
 {(actualSystem === 'group' || actualSystem === 'mixed' || actualSystem === 'mixed-full' || actualSystem === 'mixed-goldsilver' || actualSystem === 'mixed-goldsilver-full') && <div>👥 <strong>Группы</strong> — впишите названия команд</div>}
 {(actualSystem === 'group' || actualSystem === 'mixed' || actualSystem === 'mixed-full' || actualSystem === 'mixed-goldsilver' || actualSystem === 'mixed-goldsilver-full') && <div>📊 <strong>Шахматки</strong> — вписываете голы → очки и места сами</div>}
 {actualSystem !== 'group' && <div>🏆 <strong>Плей-офф</strong> — вписываете голы → победители проходят сами</div>}
 </div>
 </div>
 </div>
 </div>
 )}

 {tab === 'tournament' && (
 <TournamentView
  groups={groups}
  matches={matches}
  scores={scores}
  setScores={setScores}
  teamNames={teamNames}
  setTeamNames={setTeamNames}
  teamColors={teamColors}
  setTeamColors={setTeamColors}
  teamName={teamName}
  teamLabel={teamLabel}
  resolveSlot={resolveSlot}
  allStandings={allStandings}
  actualSystem={actualSystem}
  setScoreModal={setScoreModal}
  setImportModal={setImportModal}
  setQrModal={setQrModal}
  setProtocolModal={setProtocolModal}
 />
 )}

 {tab === 'results' && (
 <ResultsView
  groups={groups}
  allStandings={allStandings}
  matches={matches}
  scores={scores}
  teamColors={teamColors}
  teamName={teamName}
  teamLabel={teamLabel}
  resolveSlot={resolveSlot}
  actualSystem={actualSystem}
 />
 )}

 {/* Модалка ввода счёта */}
 {scoreModal && (
 <ScoreModal
  modal={scoreModal}
  scores={scores}
  teamColors={teamColors}
  refereeMode={params.refereeMode || 'manual'}
  refereeNames={refereeNames}
  referee={matchReferees[scoreModal.matchId]}
  onRefereeChange={(v) => setMatchReferees({ ...matchReferees, [scoreModal.matchId]: v })}
  onSave={(a, b, events) => { setScores({ ...scores, [scoreModal.matchId]: { a, b, events } }); syncScoreOnline(scoreModal.matchId, a, b, events); setScoreModal(null); }}
  onClear={() => { const s = { ...scores }; delete s[scoreModal.matchId]; setScores(s); setScoreModal(null); }}
  onClose={() => setScoreModal(null)}
 />
 )}

 {/* Модалка QR-кода судьи */}
 {qrModal && (
 <QRModal matchLabel={qrModal.matchLabel} matchId={qrModal.matchId}
  judgeUrl={onlineId ? `${SYNC_BACKEND_URL}/judge/${onlineId}/${qrModal.matchId}` : null}
  onClose={() => setQrModal(null)} />
 )}

 {/* Модалка печатного протокола */}
 {protocolModal && (
 <ProtocolModal
  modal={protocolModal}
  scores={scores}
  teamColors={teamColors}
  refereeMode={params.refereeMode || 'manual'}
  refereeNames={refereeNames}
  referee={matchReferees[protocolModal.matchId]}
  onClose={() => setProtocolModal(null)}
 />
 )}

 {/* Модалка импорта команд списком */}
 {importModal && (
 <ImportTeamsModal
  onImport={(list) => {
    const next = { ...teamNames };
    list.forEach((name, i) => {
      if (name && name.trim()) next[i + 1] = name.trim();
    });
    setTeamNames(next);
    setImportModal(false);
  }}
  onClose={() => setImportModal(false)}
 />
 )}

 {/* Личный кабинет: список турниров */}
 {showDashboard && (
 <TournamentsDashboard
  tournaments={tournamentsIndex}
  currentId={tournamentId}
  onOpen={openTournament}
  onCreate={createTournament}
  onDelete={deleteTournament}
  onRename={renameTournament}
  onClose={() => setShowDashboard(false)}
 />
 )}
 </div>
 {/* Липкая кнопка «Скачать xlsx» — только на мобильном, скрывается когда открыта любая модалка */}
 {!scoreModal && !importModal && !qrModal && !protocolModal && !showDashboard && (
 <div className="lg:hidden fixed bottom-0 left-0 right-0 p-3 bg-white/95 backdrop-blur border-t border-black/10 z-40" style={{ paddingBottom: 'max(0.75rem, env(safe-area-inset-bottom))' }}>
 <button onClick={handleDownload} className="w-full bg-[#e30613] active:bg-[#b1040f] text-white font-bold py-3.5 px-4 rounded flex items-center justify-center gap-2 text-base tracking-wide uppercase">
 <Download className="w-5 h-5" />Скачать xlsx
 </button>
 </div>
 )}
 <style>{`.inp{width:100%;padding:0.6rem 0.75rem;border:1px solid rgba(0,0,0,0.15);border-radius:6px;font-size:0.95rem;outline:none;-webkit-appearance:none;appearance:none;background:#fff;color:#0c0c0c;font-family:inherit}.inp:focus{border-color:#e30613;box-shadow:0 0 0 3px rgba(227,6,19,0.1)}.inp:disabled{background:#f5f2ec;color:#565656}@media(min-width:640px){.inp{font-size:0.9rem;padding:0.55rem 0.75rem}}
@media print {
  @page { size: A4; margin: 15mm; }
  body { background: #fff !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  body:not(.protocol-printing) > *:not(#root), body:not(.protocol-printing) #root > div > *:not(:has(#print-area)) { display: none !important; }
  header, nav, .print\\:hidden { display: none !important; }
  #print-area { padding: 0 !important; margin: 0 !important; }
  .break-inside-avoid { break-inside: avoid; page-break-inside: avoid; }
  body.protocol-printing * { visibility: hidden !important; }
  body.protocol-printing #protocol-print-root, body.protocol-printing #protocol-print-root * { visibility: visible !important; }
  body.protocol-printing #protocol-print-root { position: fixed; inset: 0; padding: 15mm !important; }
}`}</style>
 </div>
 );
}

// Ввод числа с диапазоном. Клэмп только при потере фокуса, чтобы никакие
// промежуточные значения (стирание, ввод первой цифры) не блокировались.
function BoundedNumber({ value, min, max, onCommit }) { const [raw, setRaw] = useState(String(value));
 // Синхронизируем raw с value когда value меняется извне
 useEffect(() => { setRaw(String(value)); }, [value]);
 return (
 <input type="number" min={min} max={max} value={raw}
 onChange={(e) => setRaw(e.target.value)}
 onBlur={() => {
 const parsed = raw === '' ? value : Number(raw);
 const n = Number.isFinite(parsed) ? Math.max(min, Math.min(max, parsed)) : value;
 setRaw(String(n));
 if (n !== value) onCommit(n);
 }}
 onKeyDown={(e) => { if (e.key === 'Enter') e.target.blur(); }}
 className="inp" />
 );
}

function Field({ label, children }) {
 return (<div><label className="block text-[10px] font-bold uppercase tracking-widest text-neutral-500 mb-1.5">{label}</label>{children}</div>);
}
function Metric({ label, value, colorClass, warning, inverted }) {
 return (
 <div className={`p-3 border ${warning ? 'border-[#e30613]/40 bg-[#e30613]/5' : 'border-black/10 bg-[#f5f2ec]'} ${colorClass || ''}`}>
 <div className={`text-[10px] font-bold uppercase tracking-widest ${inverted ? 'text-neutral-300' : 'text-neutral-500'}`}>{label}</div>
 <div className={`text-base font-black mt-0.5 ${inverted ? 'text-white' : 'text-[#0c0c0c]'}`}>{value}</div>
 </div>
 );
}

// Один блок плей-офф (заголовок + список матчей) — переиспользуется и для одиночной
// сетки, и для золотой/серебряной (тогда рендерится дважды с разным заголовком/списком).
function PlayoffBracketBlock({ title, poMatches, scores, teamColors, teamLabel, resolveSlot, setScoreModal, setProtocolModal, setQrModal }) {
  return (
    <div className="bg-white border border-black/10">
      <div className="px-4 py-3 bg-[#e30613] text-white flex items-baseline justify-between">
        <h3 className="font-black uppercase tracking-widest text-xs">{title}</h3>
        <div className="text-[10px] font-bold tracking-widest text-white/70 uppercase">{poMatches.filter((m) => scores[m.id]).length} / {poMatches.length}</div>
      </div>
      <div className="divide-y divide-black/5">
        {poMatches.map((m) => {
          const sc = scores[m.id];
          const played = sc && sc.a != null && sc.b != null;
          const t1 = teamLabel(m, 't1'), t2 = teamLabel(m, 't2');
          const resolved1 = resolveSlot ? resolveSlot(m.t1, m.bracket) : null;
          const resolved2 = resolveSlot ? resolveSlot(m.t2, m.bracket) : null;
          const sd1 = resolved1 ? resolved1.sid : null;
          const sd2 = resolved2 ? resolved2.sid : null;
          return (
            <div key={m.id} className="flex items-stretch hover:bg-[#f5f2ec] transition">
              <button onClick={() => setScoreModal({ matchId: m.id, sid1: sd1, sid2: sd2, t1Label: t1, t2Label: t2, matchLabel: m.roundName + (m.isBronze ? ' (бронза)' : '') })}
                className="flex-1 flex items-center gap-2 sm:gap-3 p-3 text-left min-w-0">
                <div className="text-[10px] font-bold uppercase text-neutral-400 tracking-wider w-16 sm:w-20 flex-shrink-0">{m.roundName}{m.isBronze ? ' 🥉' : m.roundName === 'Финал' ? ' 🏆' : ''}</div>
                <div className="flex-1 flex items-center justify-between gap-3 min-w-0">
                  <span className={`text-sm font-medium truncate flex items-center gap-1.5 ${played && sc.a > sc.b ? 'font-black text-[#0c0c0c]' : played && sc.a < sc.b ? 'text-neutral-500' : ''}`}>
                    {sd1 && teamColors[sd1] && <span className="inline-block w-1 h-4 flex-shrink-0" style={{ background: teamColors[sd1] }} />}
                    <span className="truncate">{t1}</span>
                  </span>
                  {played ? (
                    <span className="font-black text-base sm:text-lg tracking-tight flex-shrink-0">{sc.a}:{sc.b}</span>
                  ) : (
                    <span className="text-[10px] font-bold uppercase tracking-widest text-neutral-400 flex-shrink-0">–:–</span>
                  )}
                  <span className={`text-sm font-medium truncate text-right flex items-center gap-1.5 justify-end ${played && sc.b > sc.a ? 'font-black text-[#0c0c0c]' : played && sc.b < sc.a ? 'text-neutral-500' : ''}`}>
                    <span className="truncate">{t2}</span>
                    {sd2 && teamColors[sd2] && <span className="inline-block w-1 h-4 flex-shrink-0" style={{ background: teamColors[sd2] }} />}
                  </span>
                </div>
              </button>
              <button onClick={() => setProtocolModal({ matchId: m.id, matchLabel: m.roundName + (m.isBronze ? ' (бронза)' : ''), t1Label: t1, t2Label: t2, sid1, sid2 })}
                className="w-11 flex items-center justify-center text-neutral-400 hover:text-[#e30613] border-l border-black/5" title="Печатный протокол">
                <span className="text-lg">🖨</span>
              </button>
              <button onClick={() => setQrModal({ matchId: m.id, matchLabel: m.roundName + (m.isBronze ? ' (бронза)' : '') })}
                className="w-11 flex items-center justify-center text-neutral-400 hover:text-[#e30613] border-l border-black/5" title="QR для судьи">
                <span className="text-lg">▦</span>
              </button>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ============ ЭКРАН «ТУРНИР» ============
function TournamentView({ groups, matches, scores, setScores, teamNames, setTeamNames, teamColors, setTeamColors, teamName, teamLabel, resolveSlot, allStandings, actualSystem, setScoreModal, setImportModal, setQrModal, setProtocolModal }) {
  const [editingTeams, setEditingTeams] = useState(false);
  const totalMatches = matches.length;
  const playedMatches = matches.filter((m) => {
    const s = scores[m.id];
    return s && s.a != null && s.b != null;
  }).length;
  const progress = totalMatches ? Math.round(playedMatches / totalMatches * 100) : 0;

  const groupMatches = matches.filter((m) => m.phase === 'group');
  const poMatches = matches.filter((m) => m.phase === 'playoff');
  // Все sid из всех групп + оставшиеся команды для playoff-режима
  const allSids = groups.length ? groups.flat() : Array.from({ length: matches.length ? Math.max(...matches.flatMap((m) => {
    const nums = [];
    const m1 = m.t1.match(/^СИД(\d+)$/); if (m1) nums.push(+m1[1]);
    const m2 = m.t2.match(/^СИД(\d+)$/); if (m2) nums.push(+m2[1]);
    return nums;
  })) : 0 }, (_, i) => i + 1);

  return (
    <div className="space-y-5">
      {/* Прогресс */}
      <div className="bg-white border border-black/10 p-4 sm:p-5">
        <div className="flex items-baseline justify-between mb-3">
          <div>
            <div className="text-[10px] font-bold uppercase tracking-widest text-neutral-500">Прогресс турнира</div>
            <div className="text-2xl sm:text-3xl font-black text-[#0c0c0c] mt-0.5">{playedMatches} / {totalMatches}</div>
          </div>
          <div className="text-4xl sm:text-5xl font-black text-[#e30613] leading-none">{progress}<span className="text-xl">%</span></div>
        </div>
        <div className="h-2 bg-[#f5f2ec] overflow-hidden">
          <div className="h-full bg-[#e30613] transition-all" style={{ width: `${progress}%` }} />
        </div>
      </div>

      {/* Команды */}
      <div className="bg-white border border-black/10 p-4 sm:p-5">
        <div className="flex items-center justify-between mb-4 gap-2 flex-wrap">
          <h2 className="text-xs font-black uppercase tracking-widest">Команды</h2>
          <div className="flex items-center gap-3">
            <button onClick={() => setImportModal(true)} className="text-[10px] font-bold uppercase tracking-widest text-neutral-500 hover:text-[#e30613]">
              📥 Импорт списком
            </button>
            <button onClick={() => setEditingTeams(!editingTeams)} className="text-[10px] font-bold uppercase tracking-widest text-[#e30613] hover:text-[#b1040f]">
              {editingTeams ? '✓ Готово' : '✎ Редактировать'}
            </button>
          </div>
        </div>
        {editingTeams ? (
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
            {allSids.map((sid) => (
              <div key={sid} className="flex items-center gap-2">
                <div className="w-8 h-8 flex items-center justify-center bg-[#f5f2ec] text-xs font-black flex-shrink-0">{sid}</div>
                <input type="color" value={teamColors[sid] || '#e30613'}
                  onChange={(e) => { const next = { ...teamColors }; next[sid] = e.target.value; setTeamColors(next); }}
                  className="w-9 h-9 border border-black/10 cursor-pointer flex-shrink-0" title="Цвет команды" />
                <input type="text" defaultValue={teamNames[sid] || ''} placeholder={`Команда ${sid}`}
                  onBlur={(e) => {
                    const val = e.target.value.trim();
                    const next = { ...teamNames };
                    if (val) next[sid] = val; else delete next[sid];
                    setTeamNames(next);
                  }}
                  className="inp" />
              </div>
            ))}
          </div>
        ) : (
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-1.5 text-sm">
            {allSids.map((sid) => (
              <div key={sid} className="flex items-center gap-2 p-2 bg-[#f5f2ec] border-l-4" style={{ borderColor: teamColors[sid] || 'transparent' }}>
                <div className="text-[10px] font-black text-neutral-500 w-4">{sid}</div>
                <div className="font-medium truncate">{teamName(sid)}</div>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* Турнирные таблицы групп */}
      {groups.map((_, gi) => {
        const st = allStandings[gi];
        if (!st || !st.length) return null;
        const grMatches = groupMatches.filter((m) => m.group === gi + 1);
        return (
          <div key={gi} className="bg-white border border-black/10">
            <div className="px-4 py-3 bg-[#0c0c0c] text-white flex items-baseline justify-between">
              <h3 className="font-black uppercase tracking-widest text-xs">Группа {gi + 1}</h3>
              <div className="text-[10px] font-bold tracking-widest text-neutral-400 uppercase">{grMatches.filter((m) => scores[m.id]).length} / {grMatches.length} сыграно</div>
            </div>
            {/* Таблица */}
            <table className="w-full text-xs sm:text-sm">
              <thead>
                <tr className="border-b border-black/10 text-neutral-500">
                  <th className="text-left p-2 font-bold uppercase tracking-wider text-[10px]">Место</th>
                  <th className="text-left p-2 font-bold uppercase tracking-wider text-[10px]">Команда</th>
                  <th className="p-2 font-bold uppercase tracking-wider text-[10px]" title="Игр">И</th>
                  <th className="p-2 font-bold uppercase tracking-wider text-[10px]" title="Победы">В</th>
                  <th className="p-2 font-bold uppercase tracking-wider text-[10px]" title="Ничьи">Н</th>
                  <th className="p-2 font-bold uppercase tracking-wider text-[10px]" title="Поражения">П</th>
                  <th className="p-2 font-bold uppercase tracking-wider text-[10px]" title="Мячи">М</th>
                  <th className="p-2 font-bold uppercase tracking-wider text-[10px]">Очки</th>
                </tr>
              </thead>
              <tbody>
                {st.map((row) => (
                  <tr key={row.sid} className="border-b border-black/5 last:border-b-0">
                    <td className="p-2 font-black text-[#e30613] text-center">{row.place}</td>
                    <td className="p-2 font-medium">{row.name}</td>
                    <td className="p-2 text-center text-neutral-600">{row.played}</td>
                    <td className="p-2 text-center text-neutral-600">{row.w}</td>
                    <td className="p-2 text-center text-neutral-600">{row.d}</td>
                    <td className="p-2 text-center text-neutral-600">{row.l}</td>
                    <td className="p-2 text-center text-neutral-600 whitespace-nowrap">{row.gz}:{row.gp}</td>
                    <td className="p-2 text-center font-black">{row.pts}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            {/* Матчи группы */}
            <div className="border-t border-black/10 divide-y divide-black/5">
              {grMatches.map((m) => {
                const sc = scores[m.id];
                const played = sc && sc.a != null && sc.b != null;
                const sd1 = groups[gi][parseInt(m.t1.split('.')[1]) - 1];
                const sd2 = groups[gi][parseInt(m.t2.split('.')[1]) - 1];
                return (
                  <div key={m.id} className="flex items-stretch hover:bg-[#f5f2ec] transition">
                    <button onClick={() => setScoreModal({ matchId: m.id, sid1: sd1, sid2: sd2, t1Label: teamLabel(m, 't1'), t2Label: teamLabel(m, 't2'), matchLabel: m.label })}
                      className="flex-1 flex items-center gap-2 sm:gap-3 p-3 text-left min-w-0">
                      <div className="text-[10px] font-bold uppercase text-neutral-400 tracking-wider w-14 sm:w-16 flex-shrink-0">{m.label.replace(`Гр.${gi + 1} `, '')}</div>
                      <div className="flex-1 flex items-center justify-between gap-3 min-w-0">
                        <span className={`text-sm font-medium truncate flex items-center gap-1.5 ${played && sc.a > sc.b ? 'font-black text-[#0c0c0c]' : played && sc.a < sc.b ? 'text-neutral-500' : ''}`}>
                          {teamColors[sd1] && <span className="inline-block w-1 h-4 flex-shrink-0" style={{ background: teamColors[sd1] }} />}
                          <span className="truncate">{teamLabel(m, 't1')}</span>
                        </span>
                        {played ? (
                          <span className="font-black text-base sm:text-lg tracking-tight flex-shrink-0">{sc.a}:{sc.b}</span>
                        ) : (
                          <span className="text-[10px] font-bold uppercase tracking-widest text-neutral-400 flex-shrink-0">–:–</span>
                        )}
                        <span className={`text-sm font-medium truncate text-right flex items-center gap-1.5 justify-end ${played && sc.b > sc.a ? 'font-black text-[#0c0c0c]' : played && sc.b < sc.a ? 'text-neutral-500' : ''}`}>
                          <span className="truncate">{teamLabel(m, 't2')}</span>
                          {teamColors[sd2] && <span className="inline-block w-1 h-4 flex-shrink-0" style={{ background: teamColors[sd2] }} />}
                        </span>
                      </div>
                    </button>
                    <button onClick={() => setProtocolModal({ matchId: m.id, matchLabel: m.label, t1Label: teamLabel(m, 't1'), t2Label: teamLabel(m, 't2'), sid1: sd1, sid2: sd2 })}
                      className="w-11 flex items-center justify-center text-neutral-400 hover:text-[#e30613] border-l border-black/5" title="Печатный протокол">
                      <span className="text-lg">🖨</span>
                    </button>
                    <button onClick={() => setQrModal({ matchId: m.id, matchLabel: m.label })}
                      className="w-11 flex items-center justify-center text-neutral-400 hover:text-[#e30613] border-l border-black/5" title="QR для судьи">
                      <span className="text-lg">▦</span>
                    </button>
                  </div>
                );
              })}
            </div>
          </div>
        );
      })}

      {/* Плей-офф — если есть золото/серебро, разбиваем на две отдельные сетки */}
      {poMatches.length > 0 && (poMatches.some((m) => m.bracket === 'gold') || poMatches.some((m) => m.bracket === 'silver')) ? (
        ['gold', 'silver'].map((bracketKey) => {
          const bMatches = poMatches.filter((m) => m.bracket === bracketKey);
          if (bMatches.length === 0) return null;
          return (
            <PlayoffBracketBlock key={bracketKey} title={bracketKey === 'gold' ? '🥇 Золотой плей-офф' : '🥈 Серебряный плей-офф'}
              poMatches={bMatches} scores={scores} teamColors={teamColors} teamLabel={teamLabel} resolveSlot={resolveSlot}
              setScoreModal={setScoreModal} setProtocolModal={setProtocolModal} setQrModal={setQrModal} />
          );
        })
      ) : poMatches.length > 0 && (
        <PlayoffBracketBlock title="🏆 Плей-офф" poMatches={poMatches} scores={scores} teamColors={teamColors} teamLabel={teamLabel}
          resolveSlot={resolveSlot} setScoreModal={setScoreModal} setProtocolModal={setProtocolModal} setQrModal={setQrModal} />
      )}

      {/* Сброс */}
      <div className="pt-4 border-t border-black/10">
        <button onClick={() => {
          if (confirm('Удалить все счета и названия команд? Параметры турнира сохранятся.')) {
            setScores({});
            setTeamNames({});
          }
        }} className="text-[10px] font-bold uppercase tracking-widest text-neutral-400 hover:text-[#e30613]">
          Сбросить турнир
        </button>
      </div>
    </div>
  );
}


// ============ СУДЕЙСКИЙ ЭКРАН ============
// Упрощённый интерфейс для арбитра — большие цифры, +/− кнопки, «Сохранить».
// Открывается по прямой ссылке ?judge=matchId (обычно через QR-код).
function JudgeView({ matchLabel, t1Label, t2Label, color1, color2, existing, onSave }) {
  const [a, setA] = useState((existing && existing.a != null) ? existing.a : 0);
  const [b, setB] = useState((existing && existing.b != null) ? existing.b : 0);
  const [events, setEvents] = useState((existing && existing.events) || []);
  const [saved, setSaved] = useState(false);
  const clamp = (v) => Math.max(0, Math.min(99, v));
  const handleSave = () => {
    onSave(a, b, events);
    setSaved(true);
    setTimeout(() => setSaved(false), 2000);
  };
  return (
    <div className="min-h-screen bg-white flex flex-col text-[#0c0c0c]" style={{ fontFamily: "'Inter', system-ui, sans-serif" }}>
      {/* Шапка */}
      <div className="border-b border-black/10 px-4 py-3 flex items-center justify-between">
        <div className="flex items-baseline gap-2 leading-none">
          <span className="text-base font-black tracking-tight">MITIN SPORT</span>
          <span className="text-[10px] font-bold text-white bg-[#e30613] px-1.5 py-0.5 rounded-sm">GROUP</span>
        </div>
        <div className="text-[10px] font-bold uppercase tracking-widest text-[#e30613]">⚑ Судейский режим</div>
      </div>

      {/* Название матча */}
      <div className="px-4 py-4 border-b border-black/10 bg-[#f5f2ec]">
        <div className="text-[10px] font-bold uppercase tracking-widest text-neutral-500">Матч</div>
        <div className="text-lg font-black mt-0.5">{matchLabel}</div>
      </div>

      {/* Основа: две большие панели команд */}
      <div className="flex-1 flex flex-col sm:flex-row">
        {[
          { label: t1Label, val: a, set: setA, color: color1 },
          { label: t2Label, val: b, set: setB, color: color2 },
        ].map((side, i) => (
          <div key={i} className={`flex-1 p-4 sm:p-8 flex flex-col items-center justify-center ${i === 0 ? 'border-b sm:border-b-0 sm:border-r' : ''} border-black/10`}>
            {side.color && <div className="h-2 w-16 mb-4" style={{ background: side.color }} />}
            <div className="text-xs sm:text-sm font-bold text-center px-2 h-12 flex items-center justify-center">{side.label}</div>
            <div className="text-[120px] sm:text-[180px] font-black leading-none my-2 sm:my-4 tabular-nums select-none">{side.val}</div>
            <div className="grid grid-cols-2 gap-3 w-full max-w-xs">
              <button onClick={() => side.set(clamp(side.val - 1))} className="py-6 sm:py-8 bg-[#f5f2ec] active:bg-black/10 font-black text-3xl">−</button>
              <button onClick={() => side.set(clamp(side.val + 1))} className="py-6 sm:py-8 bg-[#e30613] active:bg-[#b1040f] text-white font-black text-3xl">+</button>
            </div>
            <button onClick={() => side.set(0)} className="mt-3 text-[10px] font-bold uppercase tracking-widest text-neutral-400 hover:text-[#e30613]">Обнулить</button>
          </div>
        ))}
      </div>

      {/* VAR / события */}
      <CollapsibleSection title={`VAR / События${events.length > 0 ? ` (${events.length})` : ''}`}>
        <EventsEditor events={events} setEvents={setEvents} t1Label={t1Label} t2Label={t2Label} compact />
      </CollapsibleSection>

      {/* Нижняя кнопка */}
      <div className="border-t border-black/10 p-4" style={{ paddingBottom: 'max(1rem, env(safe-area-inset-bottom))' }}>
        {saved && (
          <div className="mb-3 text-center text-xs font-bold uppercase tracking-widest text-[#e30613]">✓ Сохранено</div>
        )}
        <button onClick={handleSave} className="w-full bg-[#e30613] active:bg-[#b1040f] text-white font-black py-4 text-lg uppercase tracking-widest">
          Сохранить {a}:{b}
        </button>
      </div>
    </div>
  );
}

// ============ QR-код ============
// Использует window.qrcode (qrcode-generator, подключён в HTML)
function QRCode({ text, size = 128 }) {
  const [svg, setSvg] = useState('');
  useEffect(() => {
    if (typeof window === 'undefined' || !window.qrcode) return;
    try {
      const qr = window.qrcode(0, 'M');
      qr.addData(text);
      qr.make();
      const cellSize = Math.max(2, Math.floor(size / (qr.getModuleCount() + 8)));
      setSvg(qr.createSvgTag(cellSize, 2));
    } catch (e) { console.error('QR error', e); }
  }, [text, size]);
  return <div className="inline-block" style={{ width: size, height: size }} dangerouslySetInnerHTML={{ __html: svg }} />;
}

// ============ МОДАЛКА С QR-КОДОМ СУДЬИ ============
// Общая раскладка модалок: полноэкранный оверлей + белый лист (bottom-sheet на мобильном,
// центрированное окно от sm:). sheetClassName задаёт ширину/высоту конкретной модалки.
function ModalShell({ onClose, sheetClassName, children }) {
  return (
    <div className="fixed inset-0 bg-black/60 z-50 flex items-end sm:items-center justify-center p-0 sm:p-4" onClick={onClose}>
      <div className={`bg-white w-full sm:rounded ${sheetClassName || ''}`} onClick={(e) => e.stopPropagation()} style={{ paddingBottom: 'env(safe-area-inset-bottom)' }}>
        {children}
      </div>
    </div>
  );
}

function QRModal({ matchLabel, matchId, judgeUrl, onClose }) {
  const url = judgeUrl || `${window.location.origin}${window.location.pathname}?judge=${matchId}`;
  const copyLink = () => {
    try { navigator.clipboard.writeText(url); alert('Ссылка скопирована'); }
    catch { alert(url); }
  };
  return (
    <ModalShell onClose={onClose} sheetClassName="sm:max-w-sm">
      <div className="px-4 py-3 border-b border-black/10 flex items-center justify-between">
        <div>
          <div className="text-[10px] font-bold uppercase tracking-widest text-neutral-500">{matchLabel}</div>
          <div className="text-sm font-black">QR для судьи</div>
        </div>
        <button onClick={onClose} className="w-8 h-8 flex items-center justify-center text-neutral-400 hover:text-[#0c0c0c] text-xl">×</button>
      </div>
      <div className="p-6 flex flex-col items-center gap-4">
        <div className="border-4 border-[#0c0c0c] p-3">
          <QRCode text={url} size={220} />
        </div>
        <div className="text-xs text-center text-neutral-600 leading-tight">
          Судья сканирует QR камерой телефона → сразу открывается ввод счёта только этого матча
        </div>
        <button onClick={copyLink} className="w-full py-3 bg-[#0c0c0c] text-white font-bold uppercase tracking-widest text-xs">📋 Копировать ссылку</button>
      </div>
    </ModalShell>
  );
}

// ============ ПЕЧАТНЫЙ ПРОТОКОЛ МАТЧА (РФС-стиль) ============
function ProtocolModal({ modal, scores, teamColors, refereeMode, refereeNames, referee, onClose }) {
  const c1 = teamColors && modal.sid1 && teamColors[modal.sid1];
  const c2 = teamColors && modal.sid2 && teamColors[modal.sid2];
  const refereeDisplayName = referee == null || referee === ''
    ? null
    : (refereeMode === 'manual' ? referee : ((refereeNames || {})[referee] || null));
  useEffect(() => {
    document.body.classList.add('protocol-printing');
    return () => document.body.classList.remove('protocol-printing');
  }, []);
  const sc = scores[modal.matchId] || {};
  const played = sc.a != null && sc.b != null;
  const events = sc.events || [];
  const today = new Date().toLocaleDateString('ru-RU', { year: 'numeric', month: 'long', day: 'numeric' });

  return (
    <ModalShell onClose={onClose} sheetClassName="sm:max-w-2xl max-h-[92vh] overflow-y-auto">
      <div className="px-4 py-3 border-b border-black/10 flex items-center justify-between print:hidden">
        <div>
          <div className="text-[10px] font-bold uppercase tracking-widest text-neutral-500">{modal.matchLabel}</div>
          <div className="text-sm font-black">Протокол матча</div>
        </div>
        <div className="flex items-center gap-2">
          <button onClick={() => window.print()} className="px-3 py-1.5 bg-[#e30613] active:bg-[#b1040f] text-white text-xs font-bold uppercase tracking-widest">🖨 Печать</button>
          <button onClick={onClose} className="w-8 h-8 flex items-center justify-center text-neutral-400 hover:text-[#0c0c0c] text-xl">×</button>
        </div>
      </div>

      <div id="protocol-print-root" className="p-6 sm:p-8 text-[#0c0c0c]">
          <div className="text-center mb-6">
            <div className="text-[10px] font-bold tracking-widest text-[#e30613]">MITIN SPORT GROUP</div>
            <h2 className="text-xl font-black uppercase tracking-wide mt-1">Протокол матча</h2>
            <div className="text-xs text-neutral-500 mt-1">{modal.matchLabel}</div>
          </div>

          <div className="grid grid-cols-3 gap-2 text-xs mb-6 border border-black/20 p-3">
            <div><span className="text-neutral-500">Дата:</span> ______________</div>
            <div><span className="text-neutral-500">Время:</span> ______________</div>
            <div><span className="text-neutral-500">Поле:</span> ______________</div>
          </div>

          <div className="grid grid-cols-3 items-center gap-3 mb-6 border-y-2 border-black py-4">
            <div className="text-center">
              {c1 && <div className="h-1.5 mx-auto mb-1.5 w-8" style={{ background: c1 }} />}
              <div className="font-black text-sm">{modal.t1Label}</div>
            </div>
            <div className="text-center text-4xl font-black tabular-nums">{played ? `${sc.a} : ${sc.b}` : '— : —'}</div>
            <div className="text-center">
              {c2 && <div className="h-1.5 mx-auto mb-1.5 w-8" style={{ background: c2 }} />}
              <div className="font-black text-sm">{modal.t2Label}</div>
            </div>
          </div>

          <div className="mb-6">
            <div className="text-xs font-black uppercase tracking-widest mb-2 border-b border-black/20 pb-1">VAR / События матча</div>
            {events.length === 0 ? (
              <div className="text-xs text-neutral-400">Эпизодов не зафиксировано</div>
            ) : (
              <table className="w-full text-xs border-collapse">
                <thead>
                  <tr className="border-b border-black/20 text-left">
                    <th className="py-1 pr-2 font-bold">Мин.</th>
                    <th className="py-1 pr-2 font-bold">Команда</th>
                    <th className="py-1 pr-2 font-bold">Событие</th>
                    <th className="py-1 font-bold">Комментарий</th>
                  </tr>
                </thead>
                <tbody>
                  {events.map((e) => (
                    <tr key={e.id} className="border-b border-black/10">
                      <td className="py-1 pr-2 tabular-nums">{e.minute != null ? `${e.minute}'` : '—'}</td>
                      <td className="py-1 pr-2">{e.team === 'a' ? modal.t1Label : e.team === 'b' ? modal.t2Label : '—'}</td>
                      <td className="py-1 pr-2">{eventLabel(e.type)}</td>
                      <td className="py-1">{e.note || '—'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-3 gap-6 text-xs mt-10">
            {[{ role: 'Главный судья', name: refereeDisplayName }, { role: 'Представитель команды', name: modal.t1Label }, { role: 'Представитель команды', name: modal.t2Label }].map((p, i) => (
              <div key={i}>
                <div className="text-neutral-500 mb-6">{p.role}{p.name ? `: ${p.name}` : ''}</div>
                <div className="border-t border-black pt-1">ФИО, подпись</div>
              </div>
            ))}
          </div>

        <div className="text-center text-[10px] text-neutral-400 mt-8">Сформировано в приложении «Конструктор турниров» MITIN SPORT GROUP · {today}</div>
      </div>
    </ModalShell>
  );
}

function ResultsView({ groups, allStandings, matches, scores, teamColors, teamName, teamLabel, resolveSlot, actualSystem }) {
  const poMatches = matches.filter((m) => m.phase === 'playoff');
  const totalMatches = matches.length;
  const playedMatches = matches.filter((m) => {
    const s = scores[m.id];
    return s && s.a != null && s.b != null;
  }).length;
  const done = totalMatches > 0 && playedMatches === totalMatches;
  // Победитель финала. В золото/серебро смотрим только на золотую сетку — она даёт медали.
  const medalMatches = poMatches.some((m) => m.bracket === 'gold') ? poMatches.filter((m) => m.bracket === 'gold') : poMatches;
  const finalMatch = medalMatches.find((m) => m.roundName === 'Финал');
  let winner = null, silver = null, bronze = null;
  if (finalMatch) {
    const sc = scores[finalMatch.id];
    if (sc && sc.a != null && sc.b != null && sc.a !== sc.b) {
      const wSide = sc.a > sc.b ? 't1' : 't2';
      const lSide = sc.a > sc.b ? 't2' : 't1';
      winner = teamLabel(finalMatch, wSide);
      silver = teamLabel(finalMatch, lSide);
    }
    const bronzeMatch = medalMatches.find((m) => m.isBronze || m.roundName === 'За 3-4 место');
    if (bronzeMatch) {
      const bsc = scores[bronzeMatch.id];
      if (bsc && bsc.a != null && bsc.b != null && bsc.a !== bsc.b) {
        bronze = teamLabel(bronzeMatch, bsc.a > bsc.b ? 't1' : 't2');
      }
    }
  }
  // Список бомбардиров: собираем голы каждой команды суммарно (для team) — детальной статистики игроков пока нет
  return (
    <div className="space-y-5">
      {/* Действия */}
      <div className="flex items-center justify-between gap-3 print:hidden">
        <div>
          <div className="text-[10px] font-bold uppercase tracking-widest text-neutral-500">Итоги турнира</div>
          <div className="text-2xl font-black">{playedMatches} / {totalMatches} матчей</div>
        </div>
        <button onClick={() => window.print()} className="px-4 py-3 bg-[#0c0c0c] text-white font-bold uppercase tracking-widest text-xs hover:bg-[#e30613]">
          🖨 Печать / PDF
        </button>
      </div>

      {!done && playedMatches > 0 && (
        <div className="p-3 border border-[#e30613]/25 bg-[#e30613]/5 text-xs text-[#b1040f] print:hidden">
          Турнир ещё не завершён. Можно распечатать промежуточные итоги, но победитель может измениться.
        </div>
      )}

      {/* Печатный контейнер */}
      <div id="print-area" className="bg-white">
        {/* Шапка для печати */}
        <div className="border-b-4 border-[#e30613] pb-4 mb-6">
          <div className="flex items-baseline gap-2 leading-none mb-2">
            <span className="text-2xl font-black tracking-tight">MITIN SPORT</span>
            <span className="text-xs font-bold text-white bg-[#e30613] px-1.5 py-0.5">GROUP</span>
          </div>
          <div className="text-[10px] font-bold tracking-widest text-[#e30613]">СПОРТ ПОД КОНТРОЛЕМ · ИТОГОВЫЙ ПРОТОКОЛ</div>
        </div>

        {/* Пьедестал */}
        {(winner || silver || bronze) && (
          <div className="mb-8">
            <h2 className="text-xs font-black uppercase tracking-widest mb-3 pb-2 border-b border-black/10">Победители</h2>
            <div className="grid grid-cols-3 gap-3">
              {[
                { place: '1', label: 'Победитель', name: winner, color: '#e30613' },
                { place: '2', label: 'Серебро', name: silver, color: '#565656' },
                { place: '3', label: 'Бронза', name: bronze, color: '#b1040f' },
              ].map((p) => (
                <div key={p.place} className="text-center border border-black/10 p-4">
                  <div className="text-5xl font-black" style={{ color: p.color }}>{p.place}</div>
                  <div className="text-[10px] font-bold uppercase tracking-widest text-neutral-500 mt-1">{p.label}</div>
                  <div className="text-sm font-black mt-2 h-10 flex items-center justify-center">{p.name || '—'}</div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Таблицы групп */}
        {groups.map((_, gi) => {
          const st = allStandings[gi];
          if (!st || !st.length) return null;
          return (
            <div key={gi} className="mb-6 break-inside-avoid">
              <h3 className="text-xs font-black uppercase tracking-widest mb-2 pb-2 border-b border-black/10">Группа {gi + 1}</h3>
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-[10px] font-bold uppercase tracking-wider text-neutral-500 border-b border-black/10">
                    <th className="text-left p-2 w-12">Место</th>
                    <th className="text-left p-2">Команда</th>
                    <th className="p-2 w-10">И</th>
                    <th className="p-2 w-10">В</th>
                    <th className="p-2 w-10">Н</th>
                    <th className="p-2 w-10">П</th>
                    <th className="p-2 w-16">М</th>
                    <th className="p-2 w-14">Очки</th>
                  </tr>
                </thead>
                <tbody>
                  {st.map((row) => (
                    <tr key={row.sid} className="border-b border-black/5">
                      <td className="p-2 font-black text-[#e30613] text-center">{row.place}</td>
                      <td className="p-2 font-medium flex items-center gap-2">
                        {teamColors[row.sid] && <span className="inline-block w-1 h-4 flex-shrink-0" style={{ background: teamColors[row.sid] }} />}
                        {row.name}
                      </td>
                      <td className="p-2 text-center text-neutral-600">{row.played}</td>
                      <td className="p-2 text-center text-neutral-600">{row.w}</td>
                      <td className="p-2 text-center text-neutral-600">{row.d}</td>
                      <td className="p-2 text-center text-neutral-600">{row.l}</td>
                      <td className="p-2 text-center text-neutral-600 whitespace-nowrap">{row.gz}:{row.gp}</td>
                      <td className="p-2 text-center font-black">{row.pts}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          );
        })}

        {/* Результаты плей-офф */}
        {poMatches.length > 0 && (
          <div className="mb-6 break-inside-avoid">
            <h3 className="text-xs font-black uppercase tracking-widest mb-2 pb-2 border-b border-black/10">🏆 Плей-офф</h3>
            <table className="w-full text-sm">
              <tbody>
                {poMatches.map((m) => {
                  const sc = scores[m.id];
                  const played = sc && sc.a != null && sc.b != null;
                  return (
                    <tr key={m.id} className="border-b border-black/5">
                      <td className="p-2 text-[10px] font-bold uppercase tracking-wider text-neutral-500 w-24">{m.bracket === 'gold' ? '🥇 ' : m.bracket === 'silver' ? '🥈 ' : ''}{m.roundName}{m.isBronze ? ' 🥉' : m.roundName === 'Финал' ? ' 🏆' : ''}</td>
                      <td className="p-2 font-medium text-right w-2/5">{teamLabel(m, 't1')}</td>
                      <td className="p-2 text-center font-black tabular-nums w-16">{played ? `${sc.a}:${sc.b}` : '–:–'}</td>
                      <td className="p-2 font-medium w-2/5">{teamLabel(m, 't2')}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}

        <div className="mt-8 pt-4 border-t border-black/10 text-[10px] font-bold uppercase tracking-widest text-neutral-400">
          Протокол сгенерирован · {new Date().toLocaleDateString('ru-RU', { year: 'numeric', month: 'long', day: 'numeric' })}
        </div>
      </div>
    </div>
  );
}

// ============ МОДАЛКА ИМПОРТА КОМАНД ============
function ImportTeamsModal({ onImport, onClose }) {
  const [text, setText] = useState('');
  const preview = text.split('\n').map((s) => s.trim()).filter((s) => s.length > 0);
  return (
    <ModalShell onClose={onClose} sheetClassName="sm:max-w-md max-h-[90vh] flex flex-col">
      <div className="px-4 py-3 border-b border-black/10 flex items-center justify-between">
        <div>
          <div className="text-[10px] font-bold uppercase tracking-widest text-neutral-500">Список команд</div>
          <div className="text-sm font-black">Импорт из текста</div>
        </div>
        <button onClick={onClose} className="w-8 h-8 flex items-center justify-center text-neutral-400 hover:text-[#0c0c0c] text-xl">×</button>
      </div>
      <div className="p-4 flex-1 overflow-y-auto">
        <div className="text-xs text-neutral-600 mb-2">Вставьте по одной команде на строку. Порядок = номер команды.</div>
        <textarea value={text} onChange={(e) => setText(e.target.value)} rows={8} placeholder={"Тигры\nЛьвы\nОрлы\nВолки"}
          className="inp w-full font-mono text-sm" style={{ resize: 'vertical' }} />
        {preview.length > 0 && (
          <div className="mt-3 text-xs">
            <div className="font-bold text-[10px] uppercase tracking-widest text-neutral-500 mb-1">Будет добавлено: {preview.length}</div>
            <div className="max-h-40 overflow-y-auto border border-black/10 divide-y divide-black/5">
              {preview.map((name, i) => (
                <div key={i} className="flex items-center gap-2 p-2 bg-[#f5f2ec]">
                  <div className="w-6 text-center text-[10px] font-black text-neutral-500">{i + 1}</div>
                  <div className="font-medium truncate">{name}</div>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
      <div className="p-4 pt-0 flex gap-2">
        <button onClick={onClose} className="flex-1 py-3 bg-white border border-black/15 font-bold uppercase tracking-widest text-xs">Отмена</button>
        <button onClick={() => onImport(preview)} disabled={preview.length === 0}
          className="flex-[2] py-3 bg-[#e30613] active:bg-[#b1040f] disabled:bg-neutral-300 text-white font-black uppercase tracking-widest text-sm">
          Импортировать {preview.length}
        </button>
      </div>
    </ModalShell>
  );
}

// ============ ЛИЧНЫЙ КАБИНЕТ: СПИСОК ТУРНИРОВ ============
// Хранится локально в этом браузере (localStorage) — без логина и сервера.
function TournamentsDashboard({ tournaments, currentId, onOpen, onCreate, onDelete, onRename, onClose }) {
  const [editingId, setEditingId] = useState(null);
  const [editValue, setEditValue] = useState('');
  const sorted = [...tournaments].sort((a, b) => (b.savedAt || 0) - (a.savedAt || 0));

  const commitRename = () => {
    onRename(editingId, editValue);
    setEditingId(null);
  };

  return (
    <ModalShell onClose={onClose} sheetClassName="sm:max-w-lg max-h-[85vh] flex flex-col">
      <div className="px-4 py-3 border-b border-black/10 flex items-center justify-between">
        <div>
          <div className="text-[10px] font-bold uppercase tracking-widest text-neutral-500">Личный кабинет</div>
          <div className="text-sm font-black">Мои турниры</div>
        </div>
        <button onClick={onClose} className="w-8 h-8 flex items-center justify-center text-neutral-400 hover:text-[#0c0c0c] text-xl">×</button>
      </div>
      <div className="flex-1 overflow-y-auto divide-y divide-black/5">
        {sorted.length === 0 && (
          <div className="p-6 text-center text-sm text-neutral-400">Пока нет сохранённых турниров</div>
        )}
        {sorted.map((t) => (
          <div key={t.id} className={`p-3 flex items-center gap-2 ${t.id === currentId ? 'bg-[#f5f2ec]' : ''}`}>
            {editingId === t.id ? (
              <input autoFocus value={editValue} onChange={(e) => setEditValue(e.target.value)}
                onBlur={commitRename}
                onKeyDown={(e) => { if (e.key === 'Enter') commitRename(); if (e.key === 'Escape') setEditingId(null); }}
                className="inp text-sm flex-1" />
            ) : (
              <button onClick={() => onOpen(t.id)} className="flex-1 text-left min-w-0">
                <div className="font-bold text-sm truncate">{t.name}{t.id === currentId ? ' · открыт' : ''}</div>
                <div className="text-[10px] text-neutral-500 mt-0.5">
                  {t.totalTeams || '—'} команд · {SYS_LABELS[t.system] || t.system || '—'}{t.savedAt ? ` · ${new Date(t.savedAt).toLocaleDateString('ru-RU')}` : ''}
                </div>
              </button>
            )}
            <button onClick={() => { setEditingId(t.id); setEditValue(t.name); }}
              className="w-8 h-8 flex-shrink-0 flex items-center justify-center text-neutral-400 hover:text-[#0c0c0c]" title="Переименовать">✎</button>
            <button onClick={() => onDelete(t.id)}
              className="w-8 h-8 flex-shrink-0 flex items-center justify-center text-neutral-400 hover:text-[#e30613]" title="Удалить">🗑</button>
          </div>
        ))}
      </div>
      <div className="p-4 border-t border-black/10">
        <button onClick={onCreate} className="w-full py-3 bg-[#e30613] active:bg-[#b1040f] text-white font-bold uppercase tracking-widest text-xs">+ Новый турнир</button>
      </div>
    </ModalShell>
  );
}

// Раскрываемая секция — единое поведение сворачивания (используется и в JudgeView, и в ScoreModal)
function CollapsibleSection({ title, children, defaultOpen = false }) {
  const [open, setOpen] = useState(defaultOpen);
  return (
    <div className="border-t border-black/10">
      <button onClick={() => setOpen(!open)} className="w-full px-4 py-3 flex items-center justify-between text-xs font-bold uppercase tracking-widest text-neutral-500">
        <span>🟨 {title}</span>
        <span>{open ? '▲' : '▼'}</span>
      </button>
      {open && <div className="px-4 pb-4">{children}</div>}
    </div>
  );
}

// ============ VAR / СОБЫТИЯ МАТЧА ============
const EVENT_TYPES = [
  { key: 'goal_cancelled', label: 'Гол отменён (VAR)' },
  { key: 'penalty_var', label: 'Пенальти — просмотр VAR' },
  { key: 'red_card', label: 'Удаление (красная карточка)' },
  { key: 'yellow_card', label: 'Предупреждение (жёлтая карточка)' },
  { key: 'other', label: 'Другое' },
];
const eventLabel = (key) => (EVENT_TYPES.find((t) => t.key === key) || {}).label || key;

function EventsEditor({ events, setEvents, t1Label, t2Label, compact }) {
  const [minute, setMinute] = useState('');
  const [team, setTeam] = useState('a');
  const [type, setType] = useState(EVENT_TYPES[0].key);
  const [note, setNote] = useState('');

  const addEvent = () => {
    const ev = { id: Date.now() + Math.random(), minute: minute === '' ? null : Math.max(0, Math.min(150, Number(minute))), team, type, note: note.trim() };
    setEvents([...(events || []), ev].sort((x, y) => (x.minute ?? 999) - (y.minute ?? 999)));
    setMinute(''); setNote('');
  };
  const removeEvent = (id) => setEvents((events || []).filter((e) => e.id !== id));

  return (
    <div className={compact ? 'space-y-2' : 'space-y-3'}>
      {(events || []).length > 0 && (
        <div className="space-y-1.5">
          {events.map((e) => (
            <div key={e.id} className="flex items-center gap-2 text-xs bg-[#f5f2ec] border border-black/10 px-2.5 py-1.5">
              <span className="font-black w-8 flex-shrink-0 tabular-nums">{e.minute != null ? `${e.minute}'` : '—'}</span>
              <span className="flex-1 min-w-0">
                <span className="font-bold">{eventLabel(e.type)}</span>
                {e.team && <span className="text-neutral-500"> · {e.team === 'a' ? t1Label : t2Label}</span>}
                {e.note && <span className="text-neutral-500"> · {e.note}</span>}
              </span>
              <button onClick={() => removeEvent(e.id)} className="text-neutral-400 hover:text-[#e30613] flex-shrink-0 text-base leading-none">×</button>
            </div>
          ))}
        </div>
      )}
      <div className="grid grid-cols-[3.5rem_1fr] gap-1.5">
        <input type="number" min="0" max="150" placeholder="Мин." value={minute} onChange={(e) => setMinute(e.target.value)} className="inp text-xs px-1.5" />
        <select value={type} onChange={(e) => setType(e.target.value)} className="inp text-xs">
          {EVENT_TYPES.map((t) => <option key={t.key} value={t.key}>{t.label}</option>)}
        </select>
      </div>
      <div className="grid grid-cols-2 gap-1.5">
        <select value={team} onChange={(e) => setTeam(e.target.value)} className="inp text-xs">
          <option value="a">{t1Label}</option>
          <option value="b">{t2Label}</option>
          <option value="">—</option>
        </select>
        <input type="text" placeholder="Комментарий (опц.)" value={note} onChange={(e) => setNote(e.target.value)} className="inp text-xs" />
      </div>
      <button onClick={addEvent} className="w-full py-2 text-[10px] font-bold uppercase tracking-widest border border-dashed border-black/20 text-neutral-500 hover:border-[#e30613] hover:text-[#e30613]">
        + Добавить эпизод
      </button>
    </div>
  );
}

function ScoreModal({ modal, scores, teamColors, refereeMode, refereeNames, referee, onRefereeChange, onSave, onClear, onClose }) {
  const existing = scores[modal.matchId] || {};
  const [a, setA] = useState(existing.a != null ? existing.a : 0);
  const [b, setB] = useState(existing.b != null ? existing.b : 0);
  const [events, setEvents] = useState(existing.events || []);
  const clamp = (v) => Math.max(0, Math.min(99, v));
  const c1 = teamColors && modal.sid1 && teamColors[modal.sid1];
  const c2 = teamColors && modal.sid2 && teamColors[modal.sid2];

  return (
    <ModalShell onClose={onClose} sheetClassName="sm:max-w-md">
      <div className="px-4 py-3 border-b border-black/10 flex items-center justify-between">
        <div>
          <div className="text-[10px] font-bold uppercase tracking-widest text-neutral-500">{modal.matchLabel}</div>
          <div className="text-sm font-black">Ввод счёта</div>
        </div>
        <button onClick={onClose} className="w-8 h-8 flex items-center justify-center text-neutral-400 hover:text-[#0c0c0c] text-xl">×</button>
      </div>
      <div className="p-4 grid grid-cols-2 gap-4">
        {[
          { label: modal.t1Label, val: a, set: setA, color: c1 },
          { label: modal.t2Label, val: b, set: setB, color: c2 },
        ].map((side, i) => (
          <div key={i} className="text-center">
            {side.color && <div className="h-1.5 mx-auto mb-2 w-8" style={{ background: side.color }} />}
            <div className="text-xs font-bold mb-3 h-10 flex items-center justify-center leading-tight">{side.label}</div>
            <div className="text-6xl font-black text-[#0c0c0c] mb-3 tabular-nums">{side.val}</div>
            <div className="grid grid-cols-2 gap-2">
              <button onClick={() => side.set(clamp(side.val - 1))} className="py-3 bg-[#f5f2ec] active:bg-black/10 font-black text-xl">−</button>
              <button onClick={() => side.set(clamp(side.val + 1))} className="py-3 bg-[#e30613] active:bg-[#b1040f] text-white font-black text-xl">+</button>
            </div>
            <button onClick={() => side.set(0)} className="mt-2 text-[10px] font-bold uppercase tracking-widest text-neutral-400 hover:text-[#e30613]">Сбросить</button>
          </div>
        ))}
      </div>
      <CollapsibleSection title={`VAR / События${events.length > 0 ? ` (${events.length})` : ''}`} defaultOpen>
        <EventsEditor events={events} setEvents={setEvents} t1Label={modal.t1Label} t2Label={modal.t2Label} />
      </CollapsibleSection>
      {onRefereeChange && (
        <div className="px-4 pb-4">
          <div className="text-[10px] font-bold uppercase tracking-widest text-neutral-500 mb-1.5">Судья</div>
          {refereeMode === 'manual' ? (
            <input type="text" defaultValue={referee || ''} placeholder="Имя судьи"
              onBlur={(e) => onRefereeChange(e.target.value.trim() || null)} className="inp text-sm" />
          ) : (
            <select value={referee || ''} onChange={(e) => onRefereeChange(e.target.value ? +e.target.value : null)} className="inp text-sm">
              <option value="">— не назначен —</option>
              {Object.keys(refereeNames || {}).map(Number).sort((a, b) => a - b).map((rid) => (
                <option key={rid} value={rid}>{refereeNames[rid] || `Судья ${rid}`}</option>
              ))}
            </select>
          )}
        </div>
      )}
      <div className="p-4 pt-0 flex gap-2">
        {(existing.a != null) && (
          <button onClick={onClear} className="flex-1 py-3 bg-white border border-black/15 text-[#0c0c0c] font-bold uppercase tracking-widest text-xs">
            Удалить
          </button>
        )}
        <button onClick={() => onSave(a, b, events)} className="flex-[2] py-3 bg-[#e30613] active:bg-[#b1040f] text-white font-black uppercase tracking-widest text-sm">
          Сохранить {a}:{b}
        </button>
      </div>
    </ModalShell>
  );
}

createRoot(document.getElementById('root')).render(<TournamentBuilder />);

