/**
 * 時間のパース、フォーマット、マージを行う安全なヘルパー関数群
 * 完全に独立しているため、他の処理に影響を与えません。
 */

function safeParseTime(t) {
  if (t instanceof Date) return t.getHours() * 60 + t.getMinutes();
  if (typeof t === 'string') {
    const p = t.trim().split(':');
    if (p.length >= 2) return parseInt(p[0], 10) * 60 + parseInt(p[1], 10);
  }
  if (typeof t === 'number' && t >= 0 && t < 1) return Math.round(t * 24 * 60);
  return NaN;
}

function safeFormatTime(min) {
  return `${String(Math.floor(min/60)).padStart(2,'0')}:${String(min%60).padStart(2,'0')}`;
}

function safeNormalizeAndMerge(timeStrings, isSpecialDay) {
  const OPEN = isSpecialDay ? 10 * 60 : 9 * 60;   
  const CLOSE = isSpecialDay ? 19 * 60 : 21 * 60; 
  let intervals = [];
  timeStrings.forEach(str => {
    const parts = str.split(/[-~]/);
    if (parts.length !== 2) return;
    let start = safeParseTime(parts[0]);
    let end = safeParseTime(parts[1]);
    if (isNaN(start) || isNaN(end)) return;
    if (end <= OPEN || start >= CLOSE) return; 
    start = Math.max(start, OPEN);
    end = Math.min(end, CLOSE);
    if (start < end) intervals.push({ start, end });
  });
  if (intervals.length === 0) return "";
  intervals.sort((a, b) => a.start - b.start);
  let merged = [];
  let current = intervals[0];
  for (let i = 1; i < intervals.length; i++) {
    if (current.end >= intervals[i].start) {
      current.end = Math.max(current.end, intervals[i].end);
    } else {
      merged.push(current);
      current = intervals[i];
    }
  }
  merged.push(current);
  return merged.map(m => `${safeFormatTime(m.start)}-${safeFormatTime(m.end)}`).join('/');
}

function safeCalcTotalMin(mergedStr) {
  let total = 0;
  mergedStr.split('/').forEach(p => {
    const rng = p.split('-');
    if (rng.length === 2) {
      const s = safeParseTime(rng[0]);
      const e = safeParseTime(rng[1]);
      if (!isNaN(s) && !isNaN(e)) total += (e - s);
    }
  });
  return total;
}

function safeRemoveClosed(mergedTimeStr, closedTime) {
  if (!closedTime || closedTime === "全日") return "";
  let cRanges = [];
  if (closedTime === "午前") cRanges.push({s: 9*60, e: 13*60});
  else if (closedTime === "午後") cRanges.push({s: 15*60, e: 18*60});
  else if (closedTime === "夜間") cRanges.push({s: 18*60, e: 21*60});
  else if (closedTime === "午後夜間") cRanges.push({s: 15*60, e: 21*60});
  if (cRanges.length === 0) return mergedTimeStr;
  const validRanges = [];
  mergedTimeStr.split('/').forEach(range => {
    const parts = range.split('-');
    if (parts.length !== 2) return;
    let start = safeParseTime(parts[0]);
    let end = safeParseTime(parts[1]);
    for (const c of cRanges) {
      if (start >= end) break;
      if (c.s <= start && c.e >= end) start = end; 
      else if (c.s <= start && c.e > start) start = c.e; 
      else if (c.s < end && c.e >= end) end = c.s; 
      else if (c.s > start && c.e < end) {
        validRanges.push(`${safeFormatTime(start)}-${safeFormatTime(c.s)}`);
        start = c.e; 
      }
    }
    if (start < end) validRanges.push(`${safeFormatTime(start)}-${safeFormatTime(end)}`);
  });
  return validRanges.join('/');
}