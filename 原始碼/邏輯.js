/* ================= 小工具 ================= */
const $ = id => document.getElementById(id);
const XLNS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

function colName(n){ let s=""; while(n>0){ const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=(n-m-1)/26; } return s; }
function colNum(s){ let n=0; for(const ch of s) n=n*26+(ch.charCodeAt(0)-64); return n; }
function ref(c,r){ return colName(c)+r; }
function cjk(s){ return String(s==null?"":s).replace(/[^一-鿿]/g,""); }
function norm(v){ return v==null ? "" : String(v).replace(/\s+/g," ").trim(); }
function esc(s){ return String(s).replace(/[&<>"]/g, m => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[m])); }

function getMonthInfo(yearMinguo, month){
  const adYear = Number(yearMinguo) + 1911;
  const m = Number(month);
  const days = new Date(adYear, m, 0).getDate();
  const weeks = [];
  const weekNames = ["日","一","二","三","四","五","六"];
  for(let d = 1; d <= days; d++){
    const dayOfWeek = new Date(adYear, m - 1, d).getDay();
    weeks.push(weekNames[dayOfWeek]);
  }
  const isBig = days >= 31;
  const baseTri = isBig ? 18 : (days === 30 ? 17 : 16);
  const sheetName = `${yearMinguo}${String(m).padStart(2, '0')}`;
  return { year: Number(yearMinguo), month: m, days, weeks, isBig, baseTri, sheetName };
}

async function downloadTemplate(type){
  if(typeof TEMPLATE_30_B64 === "undefined" || typeof TEMPLATE_31_B64 === "undefined"){
    alert("範本資料讀取中，請稍候重試。");
    return;
  }
  const b64 = (type === 31) ? TEMPLATE_31_B64 : TEMPLATE_30_B64;
  const byteChars = atob(b64);
  const byteNums = new Array(byteChars.length);
  for (let i = 0; i < byteChars.length; i++) byteNums[i] = byteChars.charCodeAt(i);
  const byteArray = new Uint8Array(byteNums);

  try{
    const zip = await JSZip.loadAsync(byteArray);
    const mi = getSelectedMonthInfo();
    const hInp = $("inpBaseHours");
    const baseHours = (hInp && hInp.value) ? Number(hInp.value) : (type === 31 ? 168 : 160);

    // 1. 動態替換超勤公式中的基準工時（依據步驟 0 設定之公務工時）
    const sheetPath = "xl/worksheets/sheet1.xml";
    let sXml = await zip.file(sheetPath).async("string");

    sXml = sXml.replace(/(<f>MAX\(0,AM\d+)-(\d+)(\+E\d+\+F\d+\)<\/f>)/g, `$1-${baseHours}$3`);
    sXml = sXml.replace(/(<f>MAX\(0,AN\d+)-(\d+)(\+E\d+\+F\d+\)<\/f>)/g, `$1-${baseHours}$3`);
    sXml = sXml.replace(/(<f>AM\d+)-(\d+)(\+E\d+\+F\d+<\/f>)/g, `<f>MAX(0,AM$1-${baseHours}$3`);
    sXml = sXml.replace(/(<f>AN\d+)-(\d+)(\+E\d+\+F\d+<\/f>)/g, `<f>MAX(0,AN$1-${baseHours}$3`);

    // 2. 更新工作表抬頭第一列與星期列
    const P = new DOMParser();
    const Ser = new XMLSerializer();
    const doc = P.parseFromString(sXml, "application/xml");

    for(const c of doc.getElementsByTagName("c")){
      if(c.getAttribute("r") === "A1"){
        const tTag = c.getElementsByTagName("t")[0];
        if(tTag && tTag.textContent.includes("月份")){
          tTag.textContent = tTag.textContent.replace(/\d+年\d+月份/, mi.year + "年" + String(mi.month).padStart(2, "0") + "月份");
        }
      }
      const rMatch = c.getAttribute("r").match(/^([A-Z]+)3$/);
      if(rMatch && mi.weeks && mi.weeks.length){
        const colLetter = rMatch[1];
        const colIdx = colNum(colLetter);
        const dayIdx = colIdx - 9; // Day 1 is Col I (9)
        if(dayIdx >= 0 && dayIdx < type && mi.weeks[dayIdx]){
          c.setAttribute("t", "inlineStr");
          while(c.firstChild) c.removeChild(c.firstChild);
          const isEl = doc.createElementNS(XLNS, "is");
          const tEl = doc.createElementNS(XLNS, "t");
          tEl.textContent = mi.weeks[dayIdx];
          isEl.appendChild(tEl);
          c.appendChild(isEl);
        }
      }
    }
    sXml = Ser.serializeToString(doc);
    zip.file(sheetPath, sXml);

    // 3. 更新工作表名稱為當前年月，如 11510
    const wbPath = "xl/workbook.xml";
    let wbx = await zip.file(wbPath).async("string");
    wbx = wbx.replace(/(<sheet [^>]*name=")[^"]*("[^>]*\/>)/, `$1${mi.sheetName}$2`);
    zip.file(wbPath, wbx);

    await cleanCalcChain(zip);

    const blob = await zip.generateAsync({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      compression: "DEFLATE"
    });

    const name = `超勤(空白_${mi.sheetName}_${type}天_基準${baseHours}hr).xlsx`;
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 3000);
  }catch(e){
    const blob = new Blob([byteArray], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = (type === 31) ? "超勤(空白_31天).xlsx" : "超勤(空白_30天).xlsx";
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 3000);
  }
}

const FULL = 22, HALF = 10;
const FLEX = /^(\d{1,2})\s*[-–~]\s*(\d{1,2})\s*彈$/;

function parseFlexRest(sym){
  if(!sym) return null;
  const clean = String(sym).replace(/\s+/g, "");
  const m = clean.match(/^(\d{1,2})[-–~](\d{1,2})彈$/);
  if(!m) return null;
  const sH = Number(m[1]);
  const eH = Number(m[2]);
  const span = eH >= sH ? (eH - sH) : (eH + 24 - sH);
  const duty = Math.max(0, FULL - span);
  const sStr = String(sH).padStart(2, '0') + ":00";
  const eStr = String(eH).padStart(2, '0') + ":00";
  return {
    raw: sym,
    sH, eH, span, duty, sStr, eStr,
    label: "彈性休息 " + sStr + "–" + eStr + "（休 " + span + "hr／服勤 " + duty + "hr）"
  };
}

const REGULAR_SET = new Set(["", "▲", "休", "心", "超", "超○", "超●", "超●○"]);
function isRegularSymbol(sym){
  if(sym == null) return true;
  const s = String(sym).trim();
  if(REGULAR_SET.has(s)) return true;
  const clean = s.replace(/\s+/g, "");
  if(REGULAR_SET.has(clean)) return true;
  if(parseFlexRest(clean)) return true;
  return false;
}

function checkSpecialLeavesAndTri(days, weeks, symbols, baseTri, rosterTri){
  let qualifyingDays = 0;
  let triDays = (rosterTri != null && !isNaN(Number(rosterTri))) ? Number(rosterTri) : 0;
  let workDays = 0;
  const streakReasons = [];
  const specialDays = [];

  const isQualifyingLeave = sym =>
    sym.includes('公') || sym.includes('訓') ||
    sym.includes('病') || sym.includes('喪') ||
    sym.includes('產') || sym.includes('陪');

  let gridTriCount = 0;
  for(let d = 0; d < days; d++){
    const sym = symbols[d] || "";
    if(sym === '▲') gridTriCount++;
    if(!isRegularSymbol(sym)){
      let desc = "特殊假別";
      if(sym.includes("公") || sym.includes("訓")) desc = "公假/受訓";
      else if(sym.includes("喪")) desc = "喪假";
      else if(sym.includes("產") || sym.includes("陪")) desc = "產假/陪產假";
      else if(sym.includes("病")) desc = "病假";
      else if(sym.includes("事")) desc = "事假";
      else if(sym.includes("婚")) desc = "婚假";
      else if(sym.includes("家")) desc = "家庭照顧假";
      else if(sym.includes("生")) desc = "生理假";
      else if(sym.includes("傷")) desc = "公傷假";
      else if(sym.includes("局")) desc = "局本部支援";
      else if(sym.includes("◎") || sym.includes("●") || sym.includes("○")) desc = "複合特殊假別";
      specialDays.push({ day: d + 1, sym, desc });
    }
    const c = classify(sym);
    if(c.hours > 0) workDays++;
    if(isQualifyingLeave(sym)) qualifyingDays++;
  }
  if(rosterTri == null || isNaN(Number(rosterTri))){
    triDays = gridTriCount;
  }

  // 判定連續請假天數（依公務人員請休方式計算）：
  // 公假受訓、病假、喪假、產假/陪產假在放假日前與放假日後相連，
  // 中間夾帶之六日例假日及國定假日（公務人員本即放假，輪休表常以 ▲ 標記），視為連續串接！
  const isLeaveDay = new Array(days).fill(false);
  for(let d = 0; d < days; d++){
    if(isQualifyingLeave(symbols[d] || "")) isLeaveDay[d] = true;
  }

  // 串接前後請假中間的六日例假與國定假日（公務人員休假原則）
  for(let a = 0; a < days; a++){
    if(!isLeaveDay[a]) continue;
    for(let b = a + 1; b < days; b++){
      if(isLeaveDay[b]){
        const gap = b - a - 1;
        // 週末或國定假日連假（1~4 天無出勤且為六日或 ▲ 放假標示）
        if(gap > 0 && gap <= 4){
          let allRest = true;
          for(let k = a + 1; k < b; k++){
            const symK = symbols[k] || "";
            const isWe = weeks[k] === '六' || weeks[k] === '日';
            const isTriOrOff = symK === '▲' || symK === '休' || symK === '';
            const hoursK = classify(symK).hours;
            if(hoursK > 0 || (!isWe && !isTriOrOff)){
              allRest = false;
              break;
            }
          }
          if(allRest){
            for(let k = a + 1; k < b; k++){
              isLeaveDay[k] = true;
            }
          }
        }
        break; // 找到最近的請假日，繼續外層掃描
      }
    }
  }

  // 跨週隔週一請休（週一至週五公假，隔週一若請休非超休，連同六日以 7 天計）
  for(let d = 0; d < days; d++){
    if(weeks[d] === '一' && isQualifyingLeave(symbols[d] || "")){
      let monToFri = true;
      for(let offset = 1; offset <= 4; offset++){
        const nextD = d + offset;
        if(nextD >= days || !isQualifyingLeave(symbols[nextD] || "")){
          monToFri = false; break;
        }
      }
      if(monToFri && d + 7 < days){
        const nextMonSym = symbols[d + 7] || "";
        const isWork = nextMonSym === "" || (classify(nextMonSym).hours > 0);
        const isChao = nextMonSym.includes("超") && !nextMonSym.includes("休") && !isQualifyingLeave(nextMonSym);
        if(!isWork && !isChao){
          if(d + 5 < days) isLeaveDay[d + 5] = true;
          if(d + 6 < days) isLeaveDay[d + 6] = true;
        }
      }
    }
  }

  // 判定全月公假受訓：出勤 0 天且所有出勤日皆在公假受訓中
  const leaveCount = isLeaveDay.filter(Boolean).length;
  const isFullMonthTraining = (workDays === 0 && leaveCount >= days - 4);
  if(isFullMonthTraining){
    for(let d = 0; d < days; d++) isLeaveDay[d] = true;
  }

  // 計算最長連續請假天數
  let maxContinuousStreak = 0;
  let curStreak = 0;
  for(let d = 0; d < days; d++){
    if(isLeaveDay[d]){
      curStreak++;
      if(curStreak > maxContinuousStreak) maxContinuousStreak = curStreak;
    }else{
      curStreak = 0;
    }
  }

  let expectedTri = baseTri;
  let formulaStr = "";
  let rawExpected = baseTri;

  if(isFullMonthTraining){
    expectedTri = 0;
    rawExpected = 0;
    formulaStr = "全月公假受訓隨班休假（連續 " + days + " 天無出勤）：依公務人員請休方式計算，外勤應休本休為 0 天";
    streakReasons.push("全月連續公假受訓（出勤 0 天），依公務人員班表作息，外勤應休本休為 0 天；輪休表填有 " + triDays + " 天 ▲ 實為公務人員六日例假及國定假日放假");
  }else if(maxContinuousStreak >= 8){
    const deductDays = maxContinuousStreak;
    rawExpected = (baseTri / days) * (days - deductDays);
    expectedTri = Math.round(rawExpected);
    formulaStr = "連續請假達 " + maxContinuousStreak + " 天（依公務人員原則串接六日例假）：(" + baseTri + " / " + days + ") × (" + days + " - " + deductDays + ") = " + rawExpected.toFixed(2) + " → 應休本休為 " + expectedTri + " 天";
    streakReasons.push("連續請假含六日例假日達 " + maxContinuousStreak + " 天（符合連續超過 7 天原則）");
  }else{
    expectedTri = baseTri;
    rawExpected = baseTri;
    if(qualifyingDays > 0){
      formulaStr = "請假累計 " + qualifyingDays + " 天（最長連續 " + maxContinuousStreak + " 天，未達連續超過 7 天原則），依規定不予扣減，維持基準應休 " + baseTri + " 天";
    }else{
      formulaStr = "無特殊請假，維持基準應休 " + baseTri + " 天";
    }
  }

  const diff = triDays - expectedTri;
  return { specialDays, gongDays: (isFullMonthTraining ? days : maxContinuousStreak), qualifyingDays, triDays, expectedTri, diff, formulaStr, rawExpected, streakReasons, isOverThreshold: (maxContinuousStreak >= 8 || isFullMonthTraining), isFullMonthTraining, workDays, maxContinuousStreak };
}

/* ================= 讀 xlsx ================= */
async function loadBook(file){
  const buf = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(buf);
  const P = new DOMParser();
  const readXml = async p => { const f = zip.file(p); return f ? P.parseFromString(await f.async("string"), "application/xml") : null; };

  const shared = [];
  const ss = await readXml("xl/sharedStrings.xml");
  if(ss) for(const si of ss.getElementsByTagName("si")){
    let t = "";
    for(const n of si.getElementsByTagName("t")) t += n.textContent;
    shared.push(t);
  }

  const wb = await readXml("xl/workbook.xml");
  const rels = await readXml("xl/_rels/workbook.xml.rels");
  if(!wb || !rels) throw new Error("這不像是有效的 Excel 檔案。");
  const rmap = {};
  for(const r of rels.getElementsByTagName("Relationship")) rmap[r.getAttribute("Id")] = r.getAttribute("Target");

  const sheets = [];
  for(const sh of wb.getElementsByTagName("sheet")){
    const rid = sh.getAttribute("r:id") ||
      sh.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships","id");
    let t = rmap[rid] || "";
    t = t.startsWith("/") ? t.slice(1) : (t.startsWith("xl/") ? t : "xl/" + t);
    sheets.push({ name: sh.getAttribute("name"), path: t });
  }
  if(!sheets.length) throw new Error("這個檔案裡找不到任何工作表。");
  return { buf, zip, shared, sheets, readXml, fileName: file.name };
}

async function readCells(book, path){
  const doc = await book.readXml(path);
  if(!doc) throw new Error("讀不到工作表內容。");
  const cells = new Map();
  cells.formulas = new Map();
  for(const c of doc.getElementsByTagName("c")){
    const r = c.getAttribute("r"); if(!r) continue;
    const t = c.getAttribute("t");
    let v = null;
    if(t === "inlineStr"){
      let s = ""; for(const n of c.getElementsByTagName("t")) s += n.textContent; v = s;
    }else{
      const vn = c.getElementsByTagName("v")[0];
      if(vn) v = (t === "s") ? (book.shared[+vn.textContent] ?? "") : vn.textContent;
    }
    const fn = c.getElementsByTagName("f")[0];
    if(fn) cells.formulas.set(r, fn.textContent);
    if(v !== null && v !== "") cells.set(r, v);
  }
  return cells;
}

/* ================= 解析輪休表 ================= */
const STANDARD_COUNTIF_SYMBOLS = new Set([
  "", "▲", "休", "超", "○", "●", "超●", "超○", "◎", "★",
  "家", "事", "病", "婚", "喪", "傷", "公", "生", "訓", "產", "局",
  "停補外", "超●○", "休●○", "公●超○", "陪產", "心",
  "請日◎公外○", "超日◎公外○", "請日● 公○", "超日● 公○", "請日●公○", "超日●公○"
]);

function isStandardCountifSymbol(sym){
  if(STANDARD_COUNTIF_SYMBOLS.has(sym)) return true;
  const clean = String(sym || "").replace(/\s+/g, "");
  if(STANDARD_COUNTIF_SYMBOLS.has(clean)) return true;
  if(/^請日[●◎]公(?:外)?[○★]?$/.test(clean)) return true;
  if(/^超日[●◎]公(?:外)?[○★]?$/.test(clean)) return true;
  return false;
}

function parseRoster(cells){
  let nameRow = null;
  for(let r = 1; r <= 15 && !nameRow; r++){
    for(let c = 1; c <= 3; c++){
      const v = cells.get(ref(c, r));
      if(v && String(v).includes("姓名")){ nameRow = r; break; }
    }
  }
  if(!nameRow) throw new Error("輪休表裡找不到「姓名」那一列。");

  const people = [];
  let gap = 0;
  for(let c = 3; c <= 90; c++){
    const v = norm(cells.get(ref(c, nameRow)));
    if(v){ people.push({ col: c, name: v, key: cjk(v) }); gap = 0; }
    else if(++gap >= 3 && people.length) break;
  }
  if(!people.length) throw new Error("輪休表的姓名列是空的。");

  const rows = [];
  let want = 1;
  for(let r = nameRow; r <= nameRow + 70; r++){
    const raw = cells.get(ref(1, r));
    const n = Number(raw);
    if(raw != null && Number.isFinite(n) && Number.isInteger(n) && n === want){ rows.push({ row: r, day: n }); want++; }
    else if(rows.length && want > 28) break;
  }
  if(rows.length < 28) throw new Error("輪休表的日期列讀不完整（只讀到 " + rows.length + " 天）。");

  const weekCol = 2;
  const week = rows.map(d => norm(cells.get(ref(weekCol, d.row))));

  // 搜尋底部統計列：請休時數、補休時數、本休天數
  let leaveRow = null, compLeaveRow = null, triRow = null;
  for(let r = nameRow + 25; r <= nameRow + 75; r++){
    for(let c = 1; c <= 3; c++){
      const val = norm(cells.get(ref(c, r)));
      if(val.includes("請休時數")) leaveRow = r;
      if(val.includes("補休時數")) compLeaveRow = r;
      if(val.includes("本休▲") || (val.includes("本休") && val.includes("▲"))) triRow = r;
    }
  }

  const uncountedList = [];
  people.forEach(p => {
    p.leaveHours = leaveRow ? Number(cells.get(ref(p.col, leaveRow))) || 0 : 0;
    p.compLeaveHours = compLeaveRow ? Number(cells.get(ref(p.col, compLeaveRow))) || 0 : 0;
    p.triDays = (triRow && cells.get(ref(p.col, triRow)) != null && !isNaN(Number(cells.get(ref(p.col, triRow)))))
      ? Number(cells.get(ref(p.col, triRow))) : null;
    p.origLeave = p.leaveHours;
    p.origComp = p.compLeaveHours;
    p.uncounted = [];
  });

  const marks = people.map(p => rows.map((d, dIdx) => {
    const sym = norm(cells.get(ref(p.col, d.row)));
    if(!isStandardCountifSymbol(sym)){
      let desc = "未預設假別";
      const flex = parseFlexRest(sym);
      if(flex) desc = "彈性休息 " + flex.sStr + "~" + flex.eStr + "（休" + flex.span + "hr）";
      else if(/\s/.test(sym)) desc = "組合假別";
      else if(sym.includes("◎") || sym.includes("★") || sym.includes("請")) desc = "特殊請休假別";
      const item = { day: dIdx + 1, sym, desc, personName: p.name, col: p.col, flexInfo: flex };
      p.uncounted.push(item);
      uncountedList.push(item);
    }
    return sym;
  }));

  return { nameRow, people, days: rows.length, week, marks, leaveRow, compLeaveRow, triRow, uncountedList };
}

/* ================= 解析超勤清冊 ================= */
function parseDuty(cells){
  let hr = null, hc = null, nameCol = null;
  for(let r = 1; r <= 10 && hr === null; r++){
    for(let c = 1; c <= 15; c++){
      if(norm(cells.get(ref(c, r))) === "日期"){ hr = r; hc = c; break; }
    }
  }
  if(hr === null) throw new Error("超勤清冊裡找不到「日期」標題格。");
  for(let c = 1; c < hc; c++) if(norm(cells.get(ref(c, hr))) === "姓名") nameCol = c;
  if(nameCol === null) nameCol = 3;

  const dayCols = [];
  for(let c = hc + 1; c <= hc + 40; c++){
    const n = Number(cells.get(ref(c, hr)));
    if(Number.isFinite(n) && n === dayCols.length + 1) dayCols.push(c); else break;
  }
  if(dayCols.length < 28) throw new Error("超勤清冊的日期欄讀不完整（只讀到 " + dayCols.length + " 欄）。");

  const lastDayCol = dayCols[dayCols.length - 1];
  const sumCol = lastDayCol + 1;   // 30天為 AM(39), 31天為 AN(40)
  const claimCol = lastDayCol + 2; // 30天為 AN(40), 31天為 AO(41)
  const compCol = lastDayCol + 3;  // 30天為 AO(41), 31天為 AP(42)
  const wageCol = lastDayCol + 7;  // 30天為 AS(45), 31天為 AT(46)
  const amtCol = lastDayCol + 8;   // 30天為 AT(46), 31天為 AU(47)

  // 嘗試讀取月基準工時（由超勤時數公式如 AM4-160+E4+F4 獲取 160 或 168）
  let baseHours = 168;
  if(cells.formulas){
    for(let r = hr + 1; r <= hr + 30; r++){
      const f = cells.formulas.get(ref(dayCols[0], r));
      if(f){
        const mb = f.match(/-\s*(\d+)/);
        if(mb){ baseHours = Number(mb[1]); break; }
      }
    }
  }

  const people = [];
  let gap = 0;
  for(let r = hr + 1; r <= hr + 400; r++){
    const v = norm(cells.get(ref(nameCol, r)));
    if(v){
      const cur = dayCols.map(c => {
        const x = cells.get(ref(c, r));
        return x == null || x === "" ? null : Number(x);
      });
      const curLeave = Number(cells.get(ref(5, r))) || 0;
      const curComp = Number(cells.get(ref(6, r))) || 0;
      let wage = Number(cells.get(ref(wageCol, r))) || 0;
      if(wage === 0){
        const sal = Number(cells.get(ref(lastDayCol + 4, r))) || 0;
        const prof = Number(cells.get(ref(lastDayCol + 5, r))) || 0;
        const boss = Number(cells.get(ref(lastDayCol + 6, r))) || 0;
        if(sal > 0) wage = Math.round((sal + prof + boss) / 240);
      }
      people.push({
        row: r, name: v, key: cjk(v), rank: norm(cells.get(ref(nameCol - 1, r))),
        cur, curLeave, curComp, wage
      });
      gap = 0;
    }else if(++gap >= 4 && people.length) break;
  }
  if(!people.length) throw new Error("超勤清冊裡找不到任何人員資料列。");
  return { headRow: hr, nameCol, dayCols, days: dayCols.length, people, sumCol, claimCol, compCol, wageCol, amtCol, baseHours };
}

/* ================= 換算規則 ================= */
const BASE = {
  "":     [FULL, "v", "上班"],
  "▲":   [0, "v", "輪休"],
  "休":   [0, "v", "請休假"],
  "超":   [0, "v", "超勤補休"],
  "公":   [0, "v", "公假"],
  "病":   [0, "v", "病假"],
  "心":   [0, "v", "身心調適假（回補8hr）"],
  "婚":   [0, "v", "婚假"],
  "產":   [0, "v", "產假／陪產假（回補8hr）"],
  "喪":   [0, "v", "喪假（回補8hr）"],
  "陪產": [0, "v", "產假／陪產假（回補8hr）"],
  "超○": [HALF, "v", "超勤補休外宿（半天）"],
  "局":   [0, "a", "局本部支援"],
  "訓":   [0, "a", "受訓"],
  "事":   [0, "a", "事假"],
  "傷":   [0, "a", "公傷假"],
  "生":   [0, "a", "生理假"],
  "家":   [0, "a", "家庭照顧假"],
  "補":   [0, "a", "補休"],
  "退伍": [0, "a", "退伍"],
  "▲消": [0, "a", "輪休（消）"],
  "○":   [HALF, "a", "外宿（半天）"],
  "●":   [HALF, "a", "日休（半天）"],
  "超●": [HALF, "a", "超勤補休日休（半天）"],
  "◎":   [HALF, "a", "請休日休（半天）"],
  "★":   [HALF, "a", "請休外宿（半天）"],
  "請◎": [14, "a", "請休日休（依 115/07 前例）"],
  "超●○": [0, "v", "超勤補休日宿（整天）"],
  "休●○": [0, "v", "請休日宿（整天，回補4hr）"],
  "公●超○": [0, "v", "公假＋超勤外宿（整天，回補8hr）"],
  "請日◎公外○": [0, "v", "請日◎公外○（整天，回補4hr）"],
  "超日◎公外○": [0, "v", "超日◎公外○（整天，回補0hr）"],
  "請日● 公○": [0, "v", "請日◎公外○（整天，回補4hr）"],
  "超日● 公○": [0, "v", "超日◎公外○（整天，回補0hr）"]
};

function classify(sym){
  if(Object.prototype.hasOwnProperty.call(BASE, sym)){
    const [h, s, label] = BASE[sym];
    return { hours: h, status: s, label };
  }
  const clean = String(sym || "").replace(/\s+/g, "");
  if(Object.prototype.hasOwnProperty.call(BASE, clean)){
    const [h, s, label] = BASE[clean];
    return { hours: h, status: s, label };
  }
  if(/^請日[●◎]公(?:外)?[○★]?$/.test(clean)){
    return { hours: 0, status: "v", label: "請日◎公外○（整天，回補4hr）" };
  }
  if(/^超日[●◎]公(?:外)?[○★]?$/.test(clean)){
    return { hours: 0, status: "v", label: "超日◎公外○（整天，回補0hr）" };
  }
  const flex = parseFlexRest(sym);
  if(flex){
    return { hours: flex.duty, status: "v", label: flex.label, flexInfo: flex };
  }
  if(/\s/.test(sym)) return { hours: 0, status: "a", label: "組合符號（半天＋半天）" };
  return { hours: 0, status: "a", label: "未知符號" };
}

/* ================= 狀態 ================= */
const S = {
  year: 115, month: 10, days: 31, weeks: [], baseTri: 18, baseHours: 168,
  A: null, B: null, ra: null, du: null, sheet: null, rules: null, plan: null,
  over: {}, leaveOver: {}, gongAlerts: []
};

function updateMonthConfig(userTriggered){
  const yEl = $("selYear"), mEl = $("selMonth");
  if(yEl) S.year = Number(yEl.value);
  if(mEl) S.month = Number(mEl.value);
  try{
    if(yEl) localStorage.setItem("chaoqin_year", yEl.value);
    if(mEl) localStorage.setItem("chaoqin_month", mEl.value);
  }catch(e){}
  const mi = getMonthInfo(S.year, S.month);
  S.days = mi.days;
  S.weeks = mi.weeks;

  const hintEl = $("monthDaysHint");
  if(hintEl) hintEl.textContent = "共 " + mi.days + " 天（" + (mi.isBig ? "大月" : "小月") + "）";

  const mbText = $("mbText");
  if(mbText) mbText.textContent = S.year + " 年 " + S.month + " 月共 " + mi.days + " 天 · 01 日星期" + mi.weeks[0] + " · 請搭配「" + mi.days + " 天空白清冊」使用";
  const mbBadge = $("mbBadge");
  if(mbBadge) mbBadge.textContent = mi.isBig ? "大月 31 天" : (mi.days === 30 ? "小月 30 天" : mi.days + " 天");

  const triInp = $("inpBaseTri");
  if(triInp && (!userTriggered || !triInp.dataset.userEdited)){
    triInp.value = mi.baseTri;
  }
  S.baseTri = triInp ? Number(triInp.value) : mi.baseTri;
  try{ if(triInp) localStorage.setItem("chaoqin_base_tri", triInp.value); }catch(e){}

  const hInp = $("inpBaseHours");
  if(hInp && (!userTriggered || !hInp.dataset.userEdited)){
    hInp.value = 168;
  }
  S.baseHours = hInp ? Number(hInp.value) : 168;
  try{ if(hInp) localStorage.setItem("chaoqin_base_hours", hInp.value); }catch(e){}

  updateModalCalculator();

  if(S.A && S.B){
    refresh();
  }
}

function openCalcModal(){
  const m = $("calcModal");
  if(!m) return;
  const dSel = $("modalDays");
  if(dSel) dSel.value = String(S.days || 31);
  const triInp = $("modalBaseTri");
  if(triInp) triInp.value = String(S.baseTri || (S.days >= 31 ? 18 : 17));
  m.classList.remove("hide");
  updateModalCalculator();
}

function closeCalcModal(){
  const m = $("calcModal");
  if(m) m.classList.add("hide");
}

function updateModalCalculator(){
  const outEl = $("modalCalcOutput");
  if(!outEl) return;

  const dSel = $("modalDays");
  const triInp = $("modalBaseTri");
  const gInp = $("modalGongDays");
  const contBox = $("modalIsContinuous");
  const crossBox = $("modalCrossMon");

  const days = dSel ? (Number(dSel.value) || 31) : 31;
  const baseTri = triInp ? (Number(triInp.value) || (days >= 31 ? 18 : 17)) : 18;
  const g = gInp ? (Number(gInp.value) || 0) : 0;
  const isContinuous = contBox ? contBox.checked : true;
  const isCross = crossBox ? crossBox.checked : false;

  let desc = "";
  let badge = "";

  if(isContinuous && g > 7){
    const raw = (baseTri / days) * (days - g);
    const exp = Math.round(raw);
    const diff = exp - baseTri;
    badge = '<span class="badge b-a" style="font-size:12px;padding:3px 10px;">依比例折算本休：' + exp + ' 天（扣減 ' + Math.abs(diff) + ' 天）</span>';
    desc = '<div style="margin-top:8px;line-height:1.65;">' +
      '• <b>門檻判定</b>：累計請假 <strong>' + g + ' 天</strong>且<strong>符合連續超過 7 天（8 天以上）原則</strong>，依公務人員計假公式折算。<br>' +
      '• <b>折算公式</b>：<code>(' + baseTri + ' ÷ ' + days + ') × (' + days + ' - ' + g + ') = ' + raw.toFixed(2) + '</code><br>' +
      '• <b>折算結果</b>：應休本休天數（▲）四捨五入為 <strong>' + exp + ' 天</strong>（外勤基準天數 ' + baseTri + ' 天扣減 <strong>' + Math.abs(diff) + ' 天</strong>）。' +
      '</div>';
  }else{
    const reason = !isContinuous
      ? '（請假未符合連續原則，屬零星分散請假）'
      : (g <= 0 ? '（無請假天數）' : '（請假累計未超過 7 天門檻）');
    badge = '<span class="badge b-v" style="font-size:12px;padding:3px 10px;">維持基準本休：' + baseTri + ' 天（不扣減）</span>';
    desc = '<div style="margin-top:8px;line-height:1.65;">' +
      '• <b>門檻判定</b>：累計請假 <strong>' + g + ' 天 ' + reason + '</strong>，依外勤人員規定<strong>不予折算扣減本休</strong>。<br>' +
      '• <b>折算結果</b>：應休本休天數（▲）維持外勤基準天數 <strong>' + baseTri + ' 天</strong>（大月 18 天、小月 17 天）。' +
      '</div>';
  }

  outEl.innerHTML = '<div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px; border-bottom:1px dashed var(--line); padding-bottom:8px; margin-bottom:8px;">' +
    '<div><strong style="font-size:14px;color:var(--ink);">試算條件：' + days + ' 天月 · 基準本休 ' + baseTri + ' 天</strong></div>' +
    '<div>' + badge + '</div>' +
  '</div>' +
  '<div style="font-size:13px;color:var(--ink-2);">' +
    '• <b>試算假別</b>（公假/受訓、病假、喪假、產假/陪產假）：<strong>' + g + ' 天</strong>' +
    (isCross ? ' <span style="color:var(--accent-ink);font-weight:600;">（含週一至五公假且隔週一請休，六日納入合計以 7 天計）</span>' : '') +
  '</div>' +
  desc;
}

/* ================= 本地暫存（IndexedDB）防止重新整理遺失檔案 ================= */
const DB_NAME = "ChaoqinDB";
const DB_STORE = "files";

function openFileDb(){
  return new Promise((resolve) => {
    if(typeof window === "undefined" || !window.indexedDB){ resolve(null); return; }
    const req = window.indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = e => {
      const db = e.target.result;
      if(!db.objectStoreNames.contains(DB_STORE)){
        db.createObjectStore(DB_STORE);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => resolve(null);
  });
}

async function saveFileToDb(slot, file){
  try{
    const db = await openFileDb();
    if(!db) return;
    const buf = await file.arrayBuffer();
    return new Promise(resolve => {
      const tx = db.transaction(DB_STORE, "readwrite");
      const store = tx.objectStore(DB_STORE);
      store.put({ name: file.name, data: buf }, slot);
      tx.oncomplete = () => resolve(true);
      tx.onerror = () => resolve(false);
    });
  }catch(e){
    console.warn("Save DB failed:", e);
  }
}

async function getFileFromDb(slot){
  try{
    const db = await openFileDb();
    if(!db) return null;
    return new Promise(resolve => {
      const tx = db.transaction(DB_STORE, "readonly");
      const store = tx.objectStore(DB_STORE);
      const req = store.get(slot);
      req.onsuccess = () => {
        const item = req.result;
        if(item && item.data){
          const f = new File([item.data], item.name, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
          resolve(f);
        }else{
          resolve(null);
        }
      };
      req.onerror = () => resolve(null);
    });
  }catch(e){
    return null;
  }
}

async function clearDbFiles(){
  try{
    const db = await openFileDb();
    if(!db) return;
    return new Promise(resolve => {
      const tx = db.transaction(DB_STORE, "readwrite");
      const store = tx.objectStore(DB_STORE);
      store.clear();
      tx.oncomplete = () => resolve(true);
      tx.onerror = () => resolve(false);
    });
  }catch(e){}
}

async function resetToHome(){
  await clearDbFiles();
  S.A = null;
  S.B = null;
  S.plan = null;
  S.over = {};
  S.leaveOver = {};
  try{ localStorage.removeItem("chaoqin_over"); }catch(e){}
  try{ localStorage.removeItem("chaoqin_leaveOver"); }catch(e){}
  $("fileA").value = "";
  $("fileB").value = "";
  $("fnA").textContent = "";
  $("fnB").textContent = "";
  $("ddA").textContent = "休宿表 ▲／休／公／超 那份";
  $("ddB").textContent = "要被回填的那份（會自動校正日期與星期）";
  $("dropA").classList.remove("ok");
  $("dropB").classList.remove("ok");
  $("strip").classList.add("hide");
  $("controls").classList.add("hide");
  $("step2").classList.add("hide");
  $("step3").classList.add("hide");
  $("uncountedAlert").classList.add("hide");
  $("gongAlert").classList.add("hide");
  $("shieldBanner").classList.add("hide");
  $("btnGo").disabled = true;
  if($("barExportScopeWrap")) $("barExportScopeWrap").classList.add("hide");
  if($("unitOrgTitle")) $("unitOrgTitle").innerHTML = "新北市政府消防局<br>第七救災救護大隊　南勢分隊";
  $("barMsg").textContent = "先把兩份檔案丟進上面的框裡。";
}

/* ================= 檔案輸入 ================= */
function wireDrop(zone, input, slot){
  const open = () => input.click();
  zone.addEventListener("click", open);
  zone.addEventListener("keydown", e => { if(e.key === "Enter" || e.key === " "){ e.preventDefault(); open(); } });
  ["dragenter","dragover"].forEach(t => zone.addEventListener(t, e => { e.preventDefault(); zone.classList.add("over"); }));
  ["dragleave","drop"].forEach(t => zone.addEventListener(t, e => { e.preventDefault(); zone.classList.remove("over"); }));
  zone.addEventListener("drop", e => { const f = e.dataTransfer.files[0]; if(f) take(f, slot); });
  input.addEventListener("change", e => { const f = e.target.files[0]; if(f) take(f, slot); });
}

async function take(file, slot){
  if(!/\.xls[xm]$/i.test(file.name)){ alert("請選擇 .xlsx 檔案。"); return; }
  const zone = $(slot === "A" ? "dropA" : "dropB");
  const fn = $(slot === "A" ? "fnA" : "fnB");
  fn.textContent = "讀取中…";
  try{
    const book = await loadBook(file);
    S[slot] = book;
    zone.classList.add("ok");
    fn.textContent = file.name;
    $(slot === "A" ? "ddA" : "ddB").textContent = book.sheets.length + " 個工作表";
    await saveFileToDb(slot, file);
    await refresh();
  }catch(err){
    S[slot] = null; zone.classList.remove("ok");
    fn.textContent = "";
    alert("讀不開這個檔案：" + err.message);
  }
}

/* ================= 主流程 ================= */
async function refresh(){
  const both = S.A && S.B;
  if(!both){ $("barMsg").textContent = S.A || S.B ? "還差一份檔案。" : "先把兩份檔案丟進上面的框裡。"; return; }

  const mi = getMonthInfo(S.year, S.month);
  const targetSheet = mi.sheetName; // e.g. "11510"

  const nameA = new Set(S.A.sheets.map(s => s.name));
  let common = S.B.sheets.filter(s => nameA.has(s.name)).map(s => s.name);
  const sel = $("selSheet");

  if(common.includes(targetSheet)){
    S.sheet = targetSheet;
  }else if(!common.length){
    if(S.A.sheets.length === 1 && S.B.sheets.length === 1){
      common = [S.B.sheets[0].name];
      S.sheet = S.B.sheets[0].name;
    }else{
      $("controls").classList.remove("hide");
      sel.innerHTML = "";
      fatal("兩份檔案沒有名稱相同的工作表，無法判斷要處理哪個月份。輪休表有「" +
        S.A.sheets.map(s => s.name).join("、") + "」，超勤清冊有「" + S.B.sheets.map(s => s.name).join("、") + "」。");
      return;
    }
  }else if(!S.sheet || !common.includes(S.sheet)){
    S.sheet = common[0];
  }

  sel.innerHTML = common.map(n => '<option value="' + esc(n) + '">' + esc(n) + "</option>").join("");
  sel.value = S.sheet;
  updateExportScopeUI(common);
  $("controls").classList.remove("hide");

  const pathA = (S.A.sheets.find(s => s.name === S.sheet) || S.A.sheets[0]).path;
  const pathB = (S.B.sheets.find(s => s.name === S.sheet) || S.B.sheets[0]).path;

  try{
    S.ra = parseRoster(await readCells(S.A, pathA));
    S.du = parseDuty(await readCells(S.B, pathB));
  }catch(err){ fatal(err.message); return; }

  // 檢核天數與大小月
  if(S.du.days !== mi.days){
    fatal("作業月份天數不吻合：上方選定「" + S.year + " 年 " + S.month + " 月」（共 " + mi.days +
      " 天），但上傳的超勤清冊只有 " + S.du.days + " 個日期欄。請點上方下載「" + mi.days + " 天空白清冊」範本！");
    return;
  }
  if(S.ra.days !== S.du.days){
    fatal("大小月對不起來：輪休表是 " + S.ra.days + " 天，超勤清冊只有 " + S.du.days + " 個日期欄。請確認兩份是同一個月份。");
    return;
  }

  const hInp = $("inpBaseHours");
  if(hInp && hInp.value) S.du.baseHours = Number(hInp.value);

  matchPeople();
  buildRules();
  buildPlan();
  render();
  updateDetectedUnit();
}

function fatal(msg){
  $("strip").classList.add("hide");
  $("step2").classList.add("hide");
  $("step3").classList.remove("hide");
  $("grid").innerHTML = "";
  $("gridSub").textContent = "";
  $("finds").innerHTML = '<div class="find"><span class="tag t-stop">停</span><div class="body">' + esc(msg) + "</div></div>";
  $("findSub").textContent = "";
  $("btnGo").disabled = true;
  if($("barExportScopeWrap")) $("barExportScopeWrap").classList.add("hide");
  $("barMsg").textContent = "無法轉換。";
}

/* 先把兩份名冊配對起來：同名同姓依出現順序，一對一 */
function matchPeople(){
  const byKey = new Map();
  S.ra.people.forEach((p, i) => {
    const a = byKey.get(p.key) || [];
    a.push({ ...p, idx: i });
    byKey.set(p.key, a);
  });
  const used = new Map();
  S.pairs = S.du.people.map(dp => {
    const bucket = byKey.get(dp.key) || [];
    const k = used.get(dp.key) || 0;
    const match = bucket[k] || null;
    if(match) used.set(dp.key, k + 1);
    return { dp, match, dup: bucket.length > 1 };
  });
  S.usedIdx = new Set(S.pairs.filter(p => p.match).map(p => p.match.idx));
}

function buildRules(){
  const seen = new Map();
  const ignoredSet = getIgnoredSet();
  S.ra.marks.forEach((row, i) => {
    const p = S.ra.people[i];
    if(p && ignoredSet.has(p.key)) return; // 排除被屏蔽人員
    row.forEach(sym => {
      const e = seen.get(sym) || { sym, n: 0 };
      e.n++; seen.set(sym, e);
    });
  });
  const list = [...seen.values()].map(e => {
    const c = classify(e.sym);
    const hours = Object.prototype.hasOwnProperty.call(S.over, e.sym) ? S.over[e.sym] : c.hours;
    return { ...e, ...c, hours, edited: Object.prototype.hasOwnProperty.call(S.over, e.sym) };
  });
  list.sort((a, b) => (a.sym === "" ? -1 : b.sym === "" ? 1 : 0) || b.n - a.n);
  S.rules = list;
}
function hoursOf(sym){
  const r = S.rules.find(x => x.sym === sym);
  return r ? r.hours : classify(sym).hours;
}

function getIgnoredNames(){
  try{
    const saved = localStorage.getItem("chaoqin_ignored_names");
    if(saved !== null && saved !== "龔揚鈞") return String(saved);
    if(saved === "龔揚鈞"){
      localStorage.removeItem("chaoqin_ignored_names");
      return "";
    }
  }catch(e){}
  return "";
}

function setIgnoredNames(val){
  const str = String(val == null ? "" : val);
  try{
    localStorage.setItem("chaoqin_ignored_names", str);
  }catch(e){}
  renderShieldChips();
}

function getIgnoredSet(){
  const raw = String(getIgnoredNames() || "");
  const names = raw.split(/[,，、\s]+/).map(s => cjk(s).trim()).filter(Boolean);
  return new Set(names);
}

function renderShieldChips(){
  const set = getIgnoredSet();
  const c0 = $("ignoredChipContainer");
  const c3 = $("step3ShieldChips");
  const banner = $("shieldBanner");

  if(c0){
    if(set.size === 0){
      c0.innerHTML = '<span style="font-size:12px;color:var(--muted)">（目前未屏蔽任何人，全部同仁均正常計算超勤）</span>';
    }else{
      c0.innerHTML = Array.from(set).map(name =>
        '<span class="chip active" title="點擊取消屏蔽 ' + esc(name) + '">' +
        '<input type="checkbox" checked data-name="' + esc(name) + '" aria-label="取消屏蔽 ' + esc(name) + '">' +
        '<span>' + esc(name) + '</span>' +
        '<span class="chip-del" data-name="' + esc(name) + '" title="移除屏蔽">✕</span>' +
        '</span>'
      ).join("");
    }
  }

  if(c3 && banner){
    if(set.size === 0){
      banner.classList.add("hide");
    }else{
      banner.classList.remove("hide");
      c3.innerHTML = Array.from(set).map(name =>
        '<span class="chip active" title="點擊取消屏蔽 ' + esc(name) + '" style="font-size:12px;padding:2px 8px;">' +
        '<input type="checkbox" checked data-name="' + esc(name) + '" aria-label="取消屏蔽 ' + esc(name) + '">' +
        '<span>' + esc(name) + '</span>' +
        '<span class="chip-del" data-name="' + esc(name) + '" title="取消屏蔽">✕</span>' +
        '</span>'
      ).join("");
    }
  }

  const handleRemove = name => {
    const s = getIgnoredSet();
    s.delete(name);
    setIgnoredNames(Array.from(s).join("、"));
    if(S.plan){
      buildPlan(); renderGrid(); renderFinds(); renderBar();
    }
  };

  [c0, c3].forEach(c => {
    if(!c) return;
    c.querySelectorAll('input[type="checkbox"]').forEach(box => {
      box.addEventListener("change", () => handleRemove(box.dataset.name));
    });
    c.querySelectorAll('.chip-del').forEach(del => {
      del.addEventListener("click", e => { e.stopPropagation(); handleRemove(del.dataset.name); });
    });
  });
}

function buildPlan(){
  const scopeEl = document.querySelector('input[name=scope]:checked');
  const scope = scopeEl ? scopeEl.value : "blank";
  const ignoredSet = getIgnoredSet();

  const rows = S.pairs.map(({ dp, match, dup }) => {
    const isIgnored = ignoredSet.has(dp.key);
    if(isIgnored){
      return {
        dp, match, vals: null, shown: Array(S.du.days).fill(null), willFill: false, filled: false,
        changes: [], diffs: [], dup, sum: 0, leaveHours: 0, compLeaveHours: 0, overtimeHours: 0, wage: dp.wage || 0,
        claimCap: 0, claimHours: 0, claimAmount: 0, claimLess: false, hasUncounted: false,
        specialCheck: null, compOvertime: 0, compOverLimit: false, ignored: true, isRosterOnly: false
      };
    }

    const vals = match ? S.ra.marks[match.idx].map(hoursOf) : null;
    // 1. 空白表填入 (blank)：上傳空白表，系統直接將輪休表數字填入
    // 2. 檢核標示差異 (audit)：不動原檔數值，標示出與輪休表差異
    // 3. 強制覆蓋重填 (all)：不管原檔是否有填過或空白，一律以系統為主，且不標紅字
    const willFill = !!match && (scope !== "audit");
    const filled = dp.cur.some(v => v !== null && v !== 0 && v !== "");
    const shown = (scope === "audit") ? dp.cur : (vals || dp.cur);
    const sum = shown.reduce((a, b) => a + (b || 0), 0);

    const lo = S.leaveOver[dp.key] || {};
    const leaveHours = lo.leave != null ? lo.leave : ((scope === "audit" ? dp.curLeave : (match ? match.leaveHours : dp.curLeave)) || 0);
    const compLeaveHours = lo.comp != null ? lo.comp : ((scope === "audit" ? dp.curComp : (match ? match.compLeaveHours : dp.curComp)) || 0);

    // 檢核差異（僅在 audit 檢核模式下生效與記錄）
    const diffs = [];
    if(scope === "audit" && match && vals){
      for(let i = 0; i < S.du.days; i++){
        const curVal = dp.cur[i];
        const expVal = vals[i];
        if(curVal != null && expVal != null && curVal !== expVal){
          diffs.push({ type: 'day', day: i + 1, cur: curVal, exp: expVal });
        }
      }
      if(dp.curLeave !== match.leaveHours){
        diffs.push({ type: 'leave', cur: dp.curLeave, exp: match.leaveHours });
      }
      if(dp.curComp !== match.compLeaveHours){
        diffs.push({ type: 'comp', cur: dp.curComp, exp: match.compLeaveHours });
      }
    }

    const changes = [];
    if(willFill && vals){
      vals.forEach((v, i) => { if(dp.cur[i] != null && dp.cur[i] !== v) changes.push({ day: i + 1, from: dp.cur[i], to: v }); });
    }

    // 服勤與超勤時數最低為 0，絕不為負數！
    const overtimeHours = Math.max(0, sum - S.du.baseHours + leaveHours + compLeaveHours);
    const wage = dp.wage || 0;
    const claimCap = wage > 0 ? Math.ceil(19000 / wage) : 0;
    const claimHours = Math.max(0, Math.min(overtimeHours, claimCap));
    const claimAmount = Math.min(19000, claimHours * wage);
    const claimLess = claimCap > 0 && overtimeHours < claimCap;
    const compOvertime = Math.max(0, overtimeHours - claimHours);
    const compOverLimit = compOvertime > 20;
    const hasUncounted = match && match.uncounted && match.uncounted.length > 0;
    const specialCheck = match ? checkSpecialLeavesAndTri(S.du.days, S.weeks, S.ra.marks[match.idx], S.baseTri, match.triDays) : null;

    return {
      dp, match, vals, shown, willFill, filled, changes, diffs, dup,
      sum, leaveHours, compLeaveHours, overtimeHours, wage, claimCap,
      claimHours, claimAmount, claimLess, compOvertime, compOverLimit,
      hasUncounted, specialCheck, ignored: false, isRosterOnly: false
    };
  });

  // 將輪休表上存在但未列入超勤清冊的人員（如借調/局本部/未列冊同仁）一併納入
  const matchedIdxSet = new Set(rows.filter(r => r.match).map(r => r.match.idx));
  const unmatchedRoster = S.ra.people.map((p, i) => ({ ...p, idx: i })).filter(p => !matchedIdxSet.has(p.idx));

  const extraRows = unmatchedRoster.map(rp => {
    const isIgnored = ignoredSet.has(rp.key);
    const fakeDp = {
      name: rp.name,
      key: rp.key,
      rank: rp.rank || "輪休表(未列冊)",
      row: null,
      cur: Array(S.du.days).fill(null),
      wage: 0,
      curLeave: 0,
      curComp: 0
    };
    const vals = S.ra.marks[rp.idx].map(hoursOf);
    const specialCheck = checkSpecialLeavesAndTri(S.du.days, S.weeks, S.ra.marks[rp.idx], S.baseTri, rp.triDays);
    const sum = vals.reduce((a, b) => a + (b || 0), 0);

    return {
      dp: fakeDp,
      match: { ...rp, idx: rp.idx },
      vals,
      shown: vals,
      willFill: false,
      filled: false,
      changes: [],
      diffs: [],
      dup: false,
      sum,
      leaveHours: 0,
      compLeaveHours: 0,
      overtimeHours: 0,
      wage: 0,
      claimCap: 0,
      claimHours: 0,
      claimAmount: 0,
      claimLess: false,
      compOvertime: 0,
      compOverLimit: false,
      hasUncounted: false,
      specialCheck,
      ignored: isIgnored,
      isRosterOnly: true
    };
  });

  const allRows = rows.concat(extraRows);

  // 統計特殊假別與公假檢核名單（已排除屏蔽人員，凡有特殊假別均列出）
  S.specialAlerts = allRows.filter(r => !r.ignored && r.specialCheck && r.specialCheck.specialDays.length > 0).map(r => ({
    name: cjk(r.dp.name) || r.dp.name,
    rank: r.dp.rank,
    ...r.specialCheck
  }));

  const findings = [];
  const add = (tag, body, who) => findings.push({ tag, body, who });
  const nm = x => esc(cjk(x.dp.name) || x.dp.name);

  if(S.ra.days !== S.du.days){
    add("stop", "<b>大小月對不起來</b>：輪休表是 " + S.ra.days + " 天，超勤清冊只有 " +
      S.du.days + " 個日期欄。請確認兩份是同一個月份。");
  }

  // 檢查彈性休息時間符號
  const flexItems = [];
  allRows.forEach(r => {
    if(!r.match || r.ignored) return;
    const marks = S.ra.marks[r.match.idx];
    marks.forEach((sym, dIdx) => {
      const flex = parseFlexRest(sym);
      if(flex){
        flexItems.push({
          personName: cjk(r.dp.name) || r.dp.name,
          day: dIdx + 1,
          sym,
          flex
        });
      }
    });
  });

  if(flexItems.length){
    const flexDetails = flexItems.map(f =>
      "<b>" + esc(f.personName) + "</b>（" + f.day + "日「" + esc(f.sym) + "」" + f.flex.sStr + "~" + f.flex.eStr + " 休息 " + f.flex.span + "hr，換算服勤 " + f.flex.duty + "hr）"
    ).join("、");
    add("check", "<b>【⏰ 彈性休息時間醒目標示】</b>本月共有 " + flexItems.length + " 處排定彈性休息：" + flexDetails + "。" +
      "<br>💡 <b>勤務換算邏輯</b>：外勤上班以 08:00 至翌日 08:00 為基準（全日服勤 22 小時），「(開始)-(結束) 彈」依休息時段扣除（如 08-14 彈休息 6 小時，當日服勤算 16 小時；08-10 彈休息 2 小時，服勤算 20 小時）。表格中已將此類儲存格以<b>琥珀色 ⏰</b> 標示，請確認輪休表原檔 COUNTIF 未計入之請休(E)或補休(F)是否吻合。");
  }

  // 檢查輪休表未計入公式的非標準假別或彈性上班符號（自動排除已屏蔽同仁）
  const uncounted = (S.ra.uncountedList || []).filter(u => !ignoredSet.has(cjk(u.personName)));
  if(uncounted.length){
    uncounted.forEach(u => {
      const matchRow = allRows.find(r => r.match && r.match.name === u.personName);
      const curL = matchRow ? matchRow.leaveHours : "?";
      const curC = matchRow ? matchRow.compLeaveHours : "?";
      add("check", "<b>【⚠️ 輪休表公式未計入】" + esc(u.personName) + "</b> 在 " + u.day +
        " 日填「<b>" + esc(u.sym) + "</b>」（" + esc(u.desc) + "），<b>輪休表預設公式未加總此格</b>！" +
        "目前抓取的請休時數為 " + curL + "hr、補休時數為 " + curC + "hr，若有缺漏請直接在下方表格手動修改數值。");
    });
  }

  // 模式 2：檢核差異報告（條列所有不符項目）
  if(scope === "audit"){
    const diffRows = allRows.filter(r => !r.ignored && !r.isRosterOnly && r.diffs && r.diffs.length > 0);
    const totalDiffCount = diffRows.reduce((a, b) => a + b.diffs.length, 0);
    if(totalDiffCount === 0){
      add("note", "<b>【🎉 檢核比對結果：完全吻合】</b>上傳之超勤清冊每日服勤時數、請休時數(E欄)、補休時數(F欄)與輪休表比對完全一致，未發現任何差異！");
    }else{
      add("stop", "<b>【⚠️ 檢核比對結果：發現 " + totalDiffCount + " 處差異】</b>共有 " + diffRows.length +
        " 位同仁數值與輪休表換算不符。下方表格已將差異格以<b>紅框與粉紅底</b>標示（滑鼠移入可見應填時數）。若確認欲套用輪休表數值，請直接在上方切換至<b>「3. 強制覆蓋重填」</b>一鍵全數校正！");
      diffRows.forEach(r => {
        const dStr = r.diffs.map(d => {
          if(d.type === 'day') return d.day + "日（清冊 " + (d.cur ?? '空') + " ≠ 輪休 " + d.exp + "）";
          if(d.type === 'leave') return "請休E欄（清冊 " + d.cur + "hr ≠ 輪休 " + d.exp + "hr）";
          if(d.type === 'comp') return "補休F欄（清冊 " + d.cur + "hr ≠ 輪休 " + d.exp + "hr）";
          return "";
        }).join("、");
        add("check", "<b>" + nm(r) + "</b> 發現 " + r.diffs.length + " 處不符：" + dStr);
      });
    }
  }

  // 檢查特殊假別與應休本休（▲）天數（已排除屏蔽人員）
  if(S.specialAlerts && S.specialAlerts.length){
    S.specialAlerts.forEach(g => {
      const specialDetails = g.specialDays.map(s => s.day + "日「" + s.sym + "」").join("、");
      if(g.diff !== 0){
        let extraNotice = "";
        if(g.isFullMonthTraining){
          extraNotice = "<br><span style='color:var(--alert);font-weight:700;'>📌【全月公假受訓隨班休假】全月連續公假受訓（出勤 0 天，依公務人員班表作息），公務人員例假日與國定假日（如 25日、28日）本即放假，輪休表之 " + g.triDays + " 天 ▲ 實為例假及國定假日放假，依規定外勤應休本休為 0 天。</span>";
        }
        add("check", "<b>【⚠️ 本休天數不符】" + esc(g.name) + "</b>（" + specialDetails + "）" +
          "本月請假累計 " + g.gongDays + " 天，依規定應休本休為 <b>" + g.expectedTri + "</b> 天（" + g.formulaStr + "），但輪休表現有本休（▲）為 <b>" +
          g.triDays + "</b> 天（相差 <b>" + (g.diff > 0 ? "+" : "") + g.diff + "</b> 天）！請務必核對輪休表排假。" +
          (g.streakReasons.length ? "【請假判定說明】" + g.streakReasons.join("；") : "") + extraNotice);
      }else{
        add("note", "<b>【本休天數吻合】" + esc(g.name) + "</b>（" + specialDetails + "）" +
          "本月請假累計 " + g.gongDays + " 天，依規定應休本休為 " + g.expectedTri + " 天，輪休表現有本休亦為 " + g.triDays + " 天，完全符合規定。");
      }
    });
  }

  // 檢查超勤補休時數超標（> 20 小時，提醒安排彈性休息時間）
  const overLimitList = allRows.filter(r => r.willFill && !r.ignored && r.compOverLimit);
  if(overLimitList.length){
    overLimitList.forEach(r => {
      add("check", "<b>【⚠️ 超勤補休超標】" + nm(r) + "</b> 本月超勤補休時數達 <b>" + r.compOvertime +
        "</b> 小時（超過 20 小時上限），請確認是否安排<b>彈性休息時間</b>以降低補休時數。");
    });
  }

  const need = S.rules.filter(r => r.status === "a");
  need.forEach(r => {
    add("check", "符號「<b>" + esc(r.sym || "（空白）") + "</b>」" + esc(r.label) +
      "，本月出現 " + r.n + " 次，工具先當作 <b>" + r.hours + "</b> 小時。這個值沒有舊資料可以驗證，請確認。");
  });

  const dupSeen = new Set();
  allRows.filter(r => r.dup && r.match && !dupSeen.has(r.dp.key) && dupSeen.add(r.dp.key)).forEach(r => {
    const all = allRows.filter(x => x.dp.key === r.dp.key && x.match);
    add("check", "<b>" + nm(r) + "</b> 有 " + all.length + " 位同名同姓，工具依照兩份表上的<b>出現順序</b>一對一配對，請確認沒有配錯。",
      all.map(x => "輪休表 " + colName(x.match.col) + " 欄 → 超勤表第 " + x.dp.row + " 列").join("　｜　"));
  });

  allRows.filter(r => r.willFill).forEach(r => {
    const on = r.vals.filter(v => v === FULL).length;
    if(on < 6 || on > 18) add("check", "<b>" + nm(r) + "</b> 這個月上 <b>" + on +
      "</b> 天班，跟一般 8～16 天差距較大，建議回頭看一下輪休表。");
  });

  // 實報實銷提醒
  const lessList = allRows.filter(r => r.willFill && r.claimLess && r.claimCap > 0);
  if(lessList.length){
    add("note", "<b>【實報實銷提醒】</b>共有 " + lessList.length + " 位同仁超勤時數未達上限，已自動依超勤時數實報：" +
      lessList.map(r => nm(r) + "（超勤 " + r.overtimeHours + "hr／上限 " + r.claimCap + "hr → 實領 $" + r.claimAmount.toLocaleString() + "）").join("、"));
  }

  if(scope === "blank"){
    add("note", "<b>【1. 空白表填入模式】</b>已依輪休表自動換算填入 <b>" + allRows.filter(r => r.willFill).length + "</b> 位同仁服勤時數與公式。");
  }else if(scope === "all"){
    add("note", "<b>【3. 強制覆蓋重填模式】</b>一律以系統最新換算為主強制重填 <b>" + allRows.filter(r => r.willFill).length + "</b> 位同仁，畫面乾淨不標紅字。");
  }

  const hideIgnored = $("chkHideIgnored") ? $("chkHideIgnored").checked : true;
  const ignoredRows = allRows.filter(r => r.ignored);
  if(ignoredRows.length && !hideIgnored){
    add("note", "<b>【🚫 已屏蔽人員】</b>共有 " + ignoredRows.length + " 位同仁已依設定屏蔽，<b>不回填超勤資料且不計入請領總額</b>：" +
      ignoredRows.map(nm).join("、") + "。");
  }

  const noMatch = allRows.filter(r => !r.match && !r.ignored && !r.isRosterOnly);
  if(noMatch.length) add("note", "超勤清冊上這些人在輪休表裡找不到，<b>不會填</b>：" +
    noMatch.map(nm).join("、"));

  if(extraRows.length){
    const extraIgnored = extraRows.filter(r => r.ignored);
    const extraOthers = extraRows.filter(r => !r.ignored);
    if(extraIgnored.length && !hideIgnored){
      add("note", "<b>【🚫 已屏蔽人員】</b>輪休表「" + extraIgnored.map(r => esc(r.dp.name)).join("、") + "」已依設定屏蔽（長期不在分隊／支援局本部），不納入超勤計算。");
    }
    if(extraOthers.length){
      add("note", "<b>【輪休表未列入超勤清冊同仁】</b>「" + extraOthers.map(r => esc(r.dp.name)).join("、") + "」已由輪休表載入（局本部支援／未列冊）。已在預覽表末尾列出參考，不寫入原清冊。");
    }
  }

  S.plan = { rows: allRows, findings, fillCount: allRows.filter(r => r.willFill).length };
}

/* ================= 畫面 ================= */
function render(){
  $("strip").classList.remove("hide");
  $("step2").classList.remove("hide");
  $("step3").classList.remove("hide");

  const y = S.sheet.length >= 5 ? S.sheet.slice(0, 3) : "";
  const m = S.sheet.length >= 5 ? S.sheet.slice(3) : S.sheet;
  $("sMonth").innerHTML = y ? (y + " 年 " + Number(m) + " 月") : esc(S.sheet);
  $("sSize").innerHTML = S.ra.days >= 31 ? "大月 <small>31 天</small>" :
    (S.ra.days === 30 ? "小月 <small>30 天</small>" : "特殊 <small>" + S.ra.days + " 天</small>");
  $("sDaysA").innerHTML = S.ra.days + " <small>天</small>";
  $("sDaysB").innerHTML = S.du.days + " <small>欄</small>";
  $("sPeople").innerHTML = S.plan.rows.filter(r => r.match).length + " <small>／ " + S.du.people.length + " 人</small>";

  // 輪休表未計入符號提示卡片
  const uncounted = S.ra.uncountedList || [];
  const alertEl = $("uncountedAlert");
  if(alertEl){
    if(uncounted.length){
      alertEl.classList.remove("hide");
      $("uncountedCount").textContent = String(uncounted.length);
      const flexCount = uncounted.filter(u => u.flexInfo).length;
      let extraNotice = "";
      if(flexCount > 0){
        extraNotice = "<br><span style='color:var(--accent-ink);font-weight:600;'>💡【⏰ 彈性休息時間換算邏輯】</span>：外勤上班以 08:00 至翌日 08:00 為基準（全日 22hr），「(開始)-(結束) 彈」依休息時段扣除（如 08-14 彈休息 6hr，服勤算 16hr；08-10 彈休息 2hr，服勤算 20hr）。表格中已將此類儲存格以<b>琥珀色 ⏰</b> 醒目標示，請確認輪休表原檔 COUNTIF 未計入之請休(E)或補休(F)是否吻合。";
      }
      $("uncountedDesc").innerHTML =
        "偵測到 <strong>" + uncounted.length + "</strong> 處含有彈性休息或特殊假別符號，輪休表原檔 COUNTIF 無法統計，原表請休/補休時數可能少算：" +
        uncounted.map(u => "<strong>" + esc(u.personName) + "</strong>（" + u.day + "日 " + esc(u.sym) + (u.flexInfo ? " ⏰" : "") + "）").join("、") +
        "。已在下方表格中將其「請休/補休」欄位標記黃底，您可直接在格內核對或微調數字。" + extraNotice;
    }else{
      alertEl.classList.add("hide");
    }
  }

  // 特殊假別與應休本休天數檢核卡片
  const gongCard = $("gongAlert");
  const gongList = $("gongList");
  if(gongCard && gongList){
    if(S.specialAlerts && S.specialAlerts.length){
      gongCard.classList.remove("hide");
      gongList.innerHTML = S.specialAlerts.map(g => {
        const isWarn = g.diff !== 0;
        const statusBadge = isWarn
          ? '<span class="badge b-a">相差 ' + (g.diff > 0 ? '+' : '') + g.diff + ' 天</span>'
          : '<span class="badge b-v">吻合</span>';

        // 特殊假別明細
        const symTags = g.specialDays.map(s =>
          '<span class="badge b-a" style="margin-right:4px;">' + s.day + '日 ' + esc(s.sym) + '</span>'
        ).join("");

        let trainingWarning = "";
        if(g.isFullMonthTraining){
          trainingWarning = '<div style="background:var(--alert-wash); border:1px solid var(--alert-line); color:var(--alert); padding:6px 10px; border-radius:5px; margin-top:6px; font-weight:600; font-size:12px;">' +
            '⚠️ 全月連續受訓出勤 0 天本休異常，請確認是否重複排假。</div>';
        }

        return '<div style="border-left:3px solid ' + (isWarn ? 'var(--warn)' : 'var(--accent)') + '; margin-bottom:8px; padding:10px 14px; background:var(--surface); border-radius:6px; box-shadow:0 1px 3px rgba(0,0,0,0.05);">' +
          '<div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:6px;">' +
            '<div><strong style="font-size:14px;">' + esc(g.name) + '</strong>' + (g.rank ? ' <span style="font-size:11px;color:var(--muted)">' + esc(g.rank) + '</span>' : '') + '</div>' +
            '<div>' + statusBadge + '</div>' +
          '</div>' +
          '<div style="margin-top:6px; display:flex; flex-wrap:wrap; gap:4px; align-items:center;">' +
            '<span style="font-size:12px;font-weight:600;color:var(--muted);">特殊假別紀錄：</span>' + symTags +
          '</div>' +
          '<div style="font-size:12.5px; color:var(--ink-2); margin-top:6px;">' +
            '請假門檻天數：<b>' + g.gongDays + '</b> 天 ｜ 依比例應休本休：<b>' + g.expectedTri + '</b> 天 ｜ 輪休表現有本休（▲）：<b>' + g.triDays + '</b> 天' +
          '</div>' +
          '<div style="font-size:11.5px; color:var(--muted); margin-top:2px;">' +
            '折算說明：<code>' + esc(g.formulaStr) + '</code>' +
          '</div>' +
          (g.streakReasons && g.streakReasons.length ? '<div style="font-size:11.5px; color:var(--warn); margin-top:4px; font-weight:500;">📌 ' + g.streakReasons.map(esc).join('<br>📌 ') + '</div>' : '') +
          trainingWarning +
        '</div>';
      }).join("");
    }else{
      gongCard.classList.add("hide");
    }
  }

  const mapsEl = $("maps");
  mapsEl.innerHTML = S.rules.map(r => {
    const sym = r.sym === "" ? '<span class="sym blank">空白</span>' : '<span class="sym">' + esc(r.sym) + "</span>";
    const badge = r.status === "v" ? '<span class="badge b-v">已驗證</span>' : '<span class="badge b-a">請確認</span>';
    return '<div class="map' + (r.status === "a" ? " need" : "") + '">' +
      '<div class="map-left">' +
        sym +
        '<div class="lbl">' +
          '<b>' + esc(r.label) + '</b>' +
          '<span class="lbl-cnt">' + r.n + '格</span>' +
        '</div>' +
      '</div>' +
      '<div class="map-right">' +
        badge +
        '<input type="number" min="0" max="24" step="1" value="' + r.hours +
        '" data-sym="' + esc(r.sym) + '" aria-label="' + esc(r.label) + '時數">' +
        '<span style="font-size:12px;color:var(--muted);font-weight:600">hr</span>' +
      '</div>' +
    '</div>';
  }).join("");
  $("mapSub").textContent = S.rules.length + " 種符號，其中 " + S.rules.filter(r => r.status === "a").length + " 種待確認";
  mapsEl.querySelectorAll("input").forEach(inp => {
    inp.addEventListener("input", () => {
      const v = Number(inp.value);
      if(!Number.isFinite(v)) return;
      S.over[inp.dataset.sym] = Math.max(0, Math.min(24, Math.round(v)));
      try{ localStorage.setItem("chaoqin_over", JSON.stringify(S.over)); }catch(e){}
      buildRules(); buildPlan(); renderGrid(); renderFinds(); renderBar();
    });
  });

  renderShieldChips();
  renderGrid(); renderFinds(); renderBar();
}

function renderGrid(){
  const days = S.du.days, wk = S.ra.week;
  const isWe = i => /六|日/.test(wk[i] || "");
  const hideIgnored = $("chkHideIgnored") ? $("chkHideIgnored").checked : true;
  const scopeEl = document.querySelector('input[name=scope]:checked');
  const scope = scopeEl ? scopeEl.value : "blank";

  let h = "<thead><tr><th class='nm'>姓名</th>";
  for(let i = 0; i < days; i++) h += "<th class='" + (isWe(i) ? "we" : "") + "'>" + (i + 1) + "</th>";
  h += "<th class='stat-cell'>服勤</th>";
  h += "<th class='inp-cell' title='對應超勤清冊 E 欄，可直接修改'>請休(E)</th>";
  h += "<th class='inp-cell' title='對應超勤清冊 F 欄，可直接修改'>補休(F)</th>";
  h += "<th class='stat-cell' title='服勤總和 - 月基準(" + S.du.baseHours + ") + 請休 + 補休'>超勤</th>";
  h += "<th class='stat-cell' title='若超勤時數小於上限，則以超勤時數為主實報實銷'>支領時數</th>";
  h += "<th class='stat-cell' title='超勤補休時數＝超勤時數 - 支領時數。若超過 20 小時，建議安排彈性休息時間'>超勤補休(AO)</th>";
  h += "<th class='stat-cell money'>預估金額</th></tr>";

  if(wk.some(Boolean)){
    h += "<tr><th class='nm'></th>";
    for(let i = 0; i < days; i++) h += "<th class='wk " + (isWe(i) ? "we" : "") + "'>" + esc(wk[i] || "") + "</th>";
    h += "<th class='stat-cell'></th><th class='inp-cell'></th><th class='inp-cell'></th><th class='stat-cell'></th><th class='stat-cell'></th><th class='stat-cell'></th><th class='stat-cell money'></th></tr>";
  }
  h += "</thead><tbody>";

  for(const r of S.plan.rows){
    let cls = r.ignored ? "ignored" : (!r.match ? "skip" : (r.willFill ? "newly" : ""));
    if(r.ignored && hideIgnored) cls += " is-hidden";
    const badgeRoster = r.isRosterOnly ? "<span class='badge' style='background:var(--info-wash);color:var(--info);border:1px solid var(--info-line);margin-left:4px;'>清冊無此列</span>" : "";

    const nmHtml = "<div class='nm-row'>" +
      "<div class='nm-meta'>" +
        "<span class='nm-txt'>" + esc(cjk(r.dp.name) || r.dp.name) + "</span>" +
        (r.dp.rank ? "<span class='rk'>" + esc(r.dp.rank) + "</span>" : "") +
        badgeRoster +
        (r.hasUncounted ? "<span class='badge-uncounted' title='有特殊符號未自動計入'>待核對</span>" : "") +
      "</div>" +
      "<button type='button' class='btn-ig " + (r.ignored ? "is-ignored" : "") + "' data-name='" + esc(r.dp.key) +
      "' title='" + (r.ignored ? "點擊取消屏蔽，恢復填寫超勤" : "點擊屏蔽此同仁，不予填寫超勤") + "'>" +
      (r.ignored ? "已屏蔽 ✓" : "屏蔽") + "</button>" +
    "</div>";

    h += "<tr class='" + cls + "'><th class='nm'>" + nmHtml + "</th>";
    for(let i = 0; i < days; i++){
      const v = r.shown[i];
      const isDiff = (scope === "audit") && r.diffs && r.diffs.some(d => d.type === 'day' && d.day === i + 1);
      const rSym = r.match ? S.ra.marks[r.match.idx][i] : null;
      const flex = parseFlexRest(rSym);
      const k = v === FULL ? "on" : (v > 0 ? "half" : "");
      let clsName = "d " + k + (isDiff ? " audit-diff" : "") + (flex ? " flex-cell" : "");
      let titleAttr = "";
      if(flex){
        titleAttr = " title='" + esc(r.dp.name) + " " + (i + 1) + "日 填「" + esc(rSym) + "」：" + flex.sStr + "~" + flex.eStr + " 休息 " + flex.span + " 小時，服勤 " + flex.duty + " 小時'";
      }else if(isDiff){
        const dObj = r.diffs.find(d => d.type === 'day' && d.day === i + 1);
        titleAttr = " title='現有：" + (dObj.cur ?? "空") + "（輪休表換算應為：" + dObj.exp + "）'";
      }
      h += "<td class='" + clsName + "'" + titleAttr + ">" + (r.ignored ? "—" : (v == null ? "" : (v === 0 ? "·" : v))) + "</td>";
    }
    h += "<td class='stat-cell'>" + (r.ignored ? "—" : r.sum) + "</td>";

    const isLeaveDiff = (scope === "audit") && r.diffs && r.diffs.some(d => d.type === 'leave');
    const leaveDiffObj = isLeaveDiff ? r.diffs.find(d => d.type === 'leave') : null;
    const leaveTitle = leaveDiffObj ? " title='清冊現有：" + leaveDiffObj.cur + "hr（輪休表應為：" + leaveDiffObj.exp + "hr）'" : " title='對應超勤清冊 E 欄，可直接修改'";
    h += "<td class='inp-cell " + (r.hasUncounted ? "has-warn" : "") + (isLeaveDiff ? " audit-diff" : "") + "'" + leaveTitle + ">" +
      "<input type='number' min='0' max='300' step='1' value='" + r.leaveHours + "' data-key='" + esc(r.dp.key) + "' data-field='leave' aria-label='請休時數'" + ((r.ignored || r.isRosterOnly) ? " disabled" : "") + "></td>";

    const isCompDiff = (scope === "audit") && r.diffs && r.diffs.some(d => d.type === 'comp');
    const compDiffObj = isCompDiff ? r.diffs.find(d => d.type === 'comp') : null;
    const compTitle = compDiffObj ? " title='清冊現有：" + compDiffObj.cur + "hr（輪休表應為：" + compDiffObj.exp + "hr）'" : " title='對應超勤清冊 F 欄，可直接修改'";
    h += "<td class='inp-cell " + (r.hasUncounted ? "has-warn" : "") + (isCompDiff ? " audit-diff" : "") + "'" + compTitle + ">" +
      "<input type='number' min='0' max='300' step='1' value='" + r.compLeaveHours + "' data-key='" + esc(r.dp.key) + "' data-field='comp' aria-label='補休時數'" + ((r.ignored || r.isRosterOnly) ? " disabled" : "") + "></td>";

    h += "<td class='stat-cell'>" + (r.ignored ? "—" : r.overtimeHours) + "</td>";
    const tag = (!r.ignored && r.claimLess) ? " <span class='badge b-a'>實報</span>" : "";
    h += "<td class='stat-cell " + (r.claimLess ? "claim-less" : "") + "'>" + (r.ignored ? "—" : (r.claimHours + tag)) + "</td>";
    const compTag = (!r.ignored && r.compOverLimit) ? " <span class='badge b-a'>&gt;20</span>" : "";
    h += "<td class='stat-cell " + (r.compOverLimit ? "comp-limit-warn" : "") + "' title='超勤補休時數'>" +
      (r.ignored ? "—" : (r.compOvertime + compTag)) + "</td>";
    h += "<td class='stat-cell money'>" + (r.ignored ? "$0" : ("$" + r.claimAmount.toLocaleString())) + "</td>";
    h += "</tr>";
  }
  $("grid").innerHTML = h + "</tbody>";
  $("gridSub").textContent = (scope === "audit" ? "檢核比對模式（不動原檔）" : (S.plan.fillCount + " 人會被填入")) + " · 共 " + days + " 天";
  $("prevHint").textContent = (scope === "audit"
    ? "檢核模式：紅框＝與輪休表不符處（滑鼠移入看應填數值） · 點擊屏蔽可排除人員 · 若要自動校正請切換至「3. 強制覆蓋重填」"
    : "綠線＝本次寫入人員 · 實報＝超勤小於上限依實報銷 · >20＝超勤補休超標建議彈性休息 · 數值可直接在格內微調 · 點擊姓名旁「屏蔽」可排除特定同仁");

  $("grid").querySelectorAll(".btn-ig").forEach(btn => {
    btn.addEventListener("click", e => {
      e.stopPropagation();
      const name = btn.dataset.name;
      if(!name) return;
      const currentNames = getIgnoredNames().split(/[,，、\s]+/).map(s => s.trim()).filter(Boolean);
      const set = new Set(currentNames);
      if(set.has(name)){
        set.delete(name);
      }else{
        set.add(name);
      }
      const nextVal = Array.from(set).join("、");
      setIgnoredNames(nextVal);
      buildPlan(); renderGrid(); renderFinds(); renderBar();
    });
  });

  $("grid").querySelectorAll("input[data-field]").forEach(inp => {
    inp.addEventListener("change", () => {
      const v = Number(inp.value);
      if(!Number.isFinite(v)) return;
      const key = inp.dataset.key;
      const field = inp.dataset.field;
      if(!S.leaveOver[key]) S.leaveOver[key] = {};
      S.leaveOver[key][field] = Math.max(0, Math.min(500, Math.round(v)));
      try{ localStorage.setItem("chaoqin_leaveOver", JSON.stringify(S.leaveOver)); }catch(e){}
      buildPlan(); renderGrid(); renderFinds(); renderBar();
    });
  });
}

function renderFinds(){
  const f = S.plan.findings;
  const order = { stop: 0, check: 1, note: 2 };
  const label = { stop: "停", check: "請確認", note: "說明" };
  f.sort((a, b) => order[a.tag] - order[b.tag]);
  $("finds").innerHTML = f.length ? f.map(x =>
    '<div class="find"><span class="tag t-' + x.tag + '">' + label[x.tag] + "</span>" +
    '<div class="body">' + x.body + (x.who ? '<span class="who">' + esc(x.who) + "</span>" : "") + "</div></div>"
  ).join("") : '<div class="empty">沒有需要注意的地方。</div>';
  const nStop = f.filter(x => x.tag === "stop").length, nChk = f.filter(x => x.tag === "check").length;
  $("findSub").textContent = nStop ? nStop + " 項阻擋" : (nChk ? nChk + " 項待確認" : "全部通過");
}

function renderBar(){
  const scopeEl = document.querySelector('input[name=scope]:checked');
  const scope = scopeEl ? scopeEl.value : "blank";
  const stop = (scope === "audit") ? false : S.plan.findings.some(x => x.tag === "stop");
  const n = S.plan.fillCount;
  $("btnGo").disabled = (scope !== "audit" && (stop || n === 0));
  if($("barExportScopeWrap")) $("barExportScopeWrap").classList.toggle("hide", (scope !== "audit" && (stop || n === 0)));

  if(scope === "audit"){
    const diffCount = S.plan.rows.reduce((acc, r) => acc + (r.diffs ? r.diffs.length : 0), 0);
    $("barMsg").innerHTML = "檢核模式（不動原檔數值）：" + (diffCount > 0 ? "發現 <strong>" + diffCount + "</strong> 處與輪休表不符（已於畫面上紅框標示）。" : "全數與輪休表完全吻合。");
  }else if(scope === "all"){
    $("barMsg").innerHTML = stop ? "有阻擋項目，請先處理。" :
      (n === 0 ? "沒有人需要填寫。" : "強制覆蓋重填模式：準備寫入 <strong>" + n + "</strong> 人 × <strong>" + S.du.days + "</strong> 天到「" + esc(S.sheet) + "」工作表（畫面乾淨不標紅字）。");
  }else{
    $("barMsg").innerHTML = stop ? "有阻擋項目，請先處理。" :
      (n === 0 ? "沒有人需要填寫。" : "空白表填入模式：準備寫入 <strong>" + n + "</strong> 人 × <strong>" + S.du.days + "</strong> 天到「" + esc(S.sheet) + "」工作表。");
  }
}

/* ================= 匯出 ================= */
function setNum(doc, cellRef, value, styleFrom){
  const rowNum = +cellRef.match(/\d+$/)[0];
  const colIdx = colNum(cellRef.match(/^[A-Z]+/)[0]);
  let rowEl = null;
  for(const rEl of doc.getElementsByTagName("row")) if(+rEl.getAttribute("r") === rowNum){ rowEl = rEl; break; }
  if(!rowEl){
    const sd = doc.getElementsByTagName("sheetData")[0];
    rowEl = doc.createElementNS(XLNS, "row"); rowEl.setAttribute("r", String(rowNum));
    let before = null;
    for(const rEl of sd.getElementsByTagName("row")) if(+rEl.getAttribute("r") > rowNum){ before = rEl; break; }
    sd.insertBefore(rowEl, before);
  }
  let cEl = null;
  for(const c of rowEl.getElementsByTagName("c")) if(c.getAttribute("r") === cellRef){ cEl = c; break; }
  if(!cEl){
    cEl = doc.createElementNS(XLNS, "c");
    cEl.setAttribute("r", cellRef);
    if(styleFrom != null) cEl.setAttribute("s", styleFrom);
    let before = null;
    for(const c of rowEl.getElementsByTagName("c")){
      if(colNum(c.getAttribute("r").match(/^[A-Z]+/)[0]) > colIdx){ before = c; break; }
    }
    rowEl.insertBefore(cEl, before);
  }
  cEl.removeAttribute("t");
  while(cEl.firstChild) cEl.removeChild(cEl.firstChild);
  const v = doc.createElementNS(XLNS, "v");
  v.textContent = String(value);
  cEl.appendChild(v);
}

function setFormulaAndNum(doc, cellRef, formula, value, styleFrom){
  const rowNum = +cellRef.match(/\d+$/)[0];
  const colIdx = colNum(cellRef.match(/^[A-Z]+/)[0]);
  let rowEl = null;
  for(const rEl of doc.getElementsByTagName("row")) if(+rEl.getAttribute("r") === rowNum){ rowEl = rEl; break; }
  if(!rowEl){
    const sd = doc.getElementsByTagName("sheetData")[0];
    rowEl = doc.createElementNS(XLNS, "row"); rowEl.setAttribute("r", String(rowNum));
    let before = null;
    for(const rEl of sd.getElementsByTagName("row")) if(+rEl.getAttribute("r") > rowNum){ before = rEl; break; }
    sd.insertBefore(rowEl, before);
  }
  let cEl = null;
  for(const c of rowEl.getElementsByTagName("c")) if(c.getAttribute("r") === cellRef){ cEl = c; break; }
  if(!cEl){
    cEl = doc.createElementNS(XLNS, "c");
    cEl.setAttribute("r", cellRef);
    if(styleFrom != null) cEl.setAttribute("s", styleFrom);
    let before = null;
    for(const c of rowEl.getElementsByTagName("c")){
      if(colNum(c.getAttribute("r").match(/^[A-Z]+/)[0]) > colIdx){ before = c; break; }
    }
    rowEl.insertBefore(cEl, before);
  }
  cEl.removeAttribute("t");
  while(cEl.firstChild) cEl.removeChild(cEl.firstChild);
  if(formula){
    const f = doc.createElementNS(XLNS, "f");
    // Ensure raw formula operators are passed so XMLSerializer encodes them exactly once
    const cleanF = String(formula).replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&amp;/g, "&");
    f.textContent = cleanF;
    cEl.appendChild(f);
  }
  const v = doc.createElementNS(XLNS, "v");
  v.textContent = String(value);
  cEl.appendChild(v);
}

function setInlineStr(doc, cellRef, text, styleFrom){
  const rowNum = +cellRef.match(/\d+$/)[0];
  const colIdx = colNum(cellRef.match(/^[A-Z]+/)[0]);
  let rowEl = null;
  for(const rEl of doc.getElementsByTagName("row")) if(+rEl.getAttribute("r") === rowNum){ rowEl = rEl; break; }
  if(!rowEl){
    const sd = doc.getElementsByTagName("sheetData")[0];
    rowEl = doc.createElementNS(XLNS, "row"); rowEl.setAttribute("r", String(rowNum));
    let before = null;
    for(const rEl of sd.getElementsByTagName("row")) if(+rEl.getAttribute("r") > rowNum){ before = rEl; break; }
    sd.insertBefore(rowEl, before);
  }
  let cEl = null;
  for(const c of rowEl.getElementsByTagName("c")) if(c.getAttribute("r") === cellRef){ cEl = c; break; }
  if(!cEl){
    cEl = doc.createElementNS(XLNS, "c");
    cEl.setAttribute("r", cellRef);
    if(styleFrom != null) cEl.setAttribute("s", styleFrom);
    let before = null;
    for(const c of rowEl.getElementsByTagName("c")){
      if(colNum(c.getAttribute("r").match(/^[A-Z]+/)[0]) > colIdx){ before = c; break; }
    }
    rowEl.insertBefore(cEl, before);
  }
  cEl.setAttribute("t", "inlineStr");
  while(cEl.firstChild) cEl.removeChild(cEl.firstChild);
  const is = doc.createElementNS(XLNS, "is");
  const t = doc.createElementNS(XLNS, "t");
  t.textContent = String(text);
  is.appendChild(t);
  cEl.appendChild(is);
}

function toChineseCurrency(num){
  if(!num || num <= 0) return "零元整";
  const digits = ["零","壹","貳","參","肆","伍","陸","柒","捌","玖"];
  let s = String(num).padStart(8, "0");
  s = s.split("").map(d => digits[+d]).join("");
  for(const z of ["零零零零零零零","零零零零零零","零零零零零","零零零零","零零零","零零"]){
    s = s.replace(z, "");
  }
  return s + "元整";
}

function parseYearMonth(sheetName){
  let y = 115, m = 9;
  const m5 = sheetName.match(/^(\d{3})(\d{2})$/);
  if(m5){ y = +m5[1]; m = +m5[2]; }
  else{
    const ym = sheetName.match(/(\d{2,3})年?(\d{1,2})/);
    if(ym){ y = +ym[1]; m = +ym[2]; }
  }
  return { year: y, month: m, str: y + "年" + String(m).padStart(2, "0") + "月份" };
}

function extractUnitName(){
  const candidates = [];
  if(S.B && S.B.shared) candidates.push(...S.B.shared);
  if(S.A && S.A.shared) candidates.push(...S.A.shared);
  if(S.B && S.B.fileName) candidates.push(S.B.fileName);
  if(S.A && S.A.fileName) candidates.push(S.A.fileName);

  for(const s of candidates){
    if(typeof s !== "string") continue;
    const mBranch = s.match(/([^\s\d_、，,。\(\)\（\）\<\>\"\'\/]+?分隊)/);
    if(mBranch && mBranch[1].length >= 3 && mBranch[1] !== "分隊"){
      let branch = mBranch[1].trim();
      if(branch.includes("大隊")){
        const sub = branch.match(/大隊(.*分隊)/);
        if(sub) branch = sub[1];
      }
      if(branch.includes("消防局")){
        let name = branch;
        if(!name.includes("政府") && name.match(/^(\S+?[市縣])消防局/)){
          name = name.replace(/^(\S+?[市縣])消防局/, "$1政府消防局");
        }
        return name;
      }
      const mBureau = s.match(/([^\s\d_、，,。\(\)\（\）\<\>\"\'\/]*?消防局)/);
      let bureau = mBureau ? mBureau[1].trim() : "新北市政府消防局";
      if(!bureau.includes("政府") && bureau.match(/^(\S+?[市縣])消防局/)){
        bureau = bureau.replace(/^(\S+?[市縣])消防局/, "$1政府消防局");
      }
      if(!bureau.includes("消防局")) bureau = "新北市政府消防局";
      return bureau + branch;
    }
  }
  return "新北市政府消防局南勢分隊";
}

function updateDetectedUnit(){
  const unit = extractUnitName();
  const el = $("unitOrgTitle");
  if(el && unit){
    const m = unit.match(/^(.*?消防局)(.*分隊)$/);
    if(m){
      el.innerHTML = esc(m[1]) + "<br>" + esc(m[2]);
    }else{
      el.innerHTML = esc(unit);
    }
  }
}

function fillDocSheet(doc, sheetName, ra, du, plan, ymInfo, isAudit){
  if(isAudit) return 0; // 檢核模式維持原樣不動

  const bodyRows = new Set(du.people.map(p => p.row));
  const styleAt = {};
  for(const c of doc.getElementsByTagName("c")){
    const r = c.getAttribute("r"); if(!r) continue;
    if(!bodyRows.has(+r.match(/\d+$/)[0])) continue;
    const col = r.match(/^[A-Z]+/)[0];
    if(styleAt[col] == null && c.getAttribute("s") != null) styleAt[col] = c.getAttribute("s");
  }

  const firstCol = du.dayCols[0];
  const lastDayCol = du.dayCols[du.dayCols.length - 1];
  const firstColName = colName(firstCol);
  const lastDayColName = colName(lastDayCol);
  const sumColName = colName(du.sumCol);
  const claimColName = colName(du.claimCol);
  const compColName = colName(du.compCol);
  const wageColName = colName(du.wageCol);
  const amtColName = colName(du.amtCol);

  let cells = 0;

  // 1. 自動同步標題月份（A1）
  const useYear = ymInfo.year;
  const useMonth = ymInfo.month;
  const titleText = "新北市政府消防局南勢分隊" + useYear + "年" + String(useMonth).padStart(2, "0") + "月份超勤加班費申請及印領清冊";
  setInlineStr(doc, "A1", titleText, styleAt["A1"]);
  cells++;

  // 2. 自動校正日期列（第 2 列：1, 2, 3, ... N）
  du.dayCols.forEach((col, i) => {
    setNum(doc, ref(col, 2), i + 1, styleAt[colName(col)]);
    cells++;
  });

  // 3. 自動校正星期列（第 3 列：一、二、三...）
  const weekList = (ymInfo.weeks && ymInfo.weeks.length === du.days) ? ymInfo.weeks : (ra && ra.week ? ra.week : []);
  if(weekList && weekList.length){
    du.dayCols.forEach((col, i) => {
      if(weekList[i]){
        setInlineStr(doc, ref(col, 3), weekList[i], styleAt[colName(col)]);
        cells++;
      }
    });
  }

  for(const r of plan.rows){
    if(r.isRosterOnly || !r.dp || !r.dp.row) continue;
    if(r.ignored){
      // 屏蔽人員：每日工時、請休、補休設為 0，超勤公式保留但數值為 0
      du.dayCols.forEach((col) => {
        setNum(doc, ref(col, r.dp.row), 0, styleAt[colName(col)]);
        cells++;
      });
      setNum(doc, ref(5, r.dp.row), 0, styleAt["E"] || styleAt[firstColName]);
      cells++;
      setNum(doc, ref(6, r.dp.row), 0, styleAt["F"] || styleAt[firstColName]);
      cells++;
      const sumFormula = "SUM(" + firstColName + r.dp.row + ":" + lastDayColName + r.dp.row + ")";
      setFormulaAndNum(doc, ref(du.sumCol, r.dp.row), sumFormula, 0, styleAt[sumColName]);
      cells++;
      const otFormula = "MAX(0," + sumColName + r.dp.row + "-" + du.baseHours + "+E" + r.dp.row + "+F" + r.dp.row + ")";
      setFormulaAndNum(doc, ref(firstCol, r.dp.row + 1), otFormula, 0, styleAt[firstColName]);
      cells++;
      const overtimeCell = ref(firstCol, r.dp.row + 1);
      const wageCell = ref(du.wageCol, r.dp.row);
      const claimFormula = "MAX(0,MIN(" + overtimeCell + ",ROUNDUP(19000/" + wageCell + ",0)))";
      setFormulaAndNum(doc, ref(du.claimCol, r.dp.row), claimFormula, 0, styleAt[claimColName]);
      cells++;
      const claimCell = ref(du.claimCol, r.dp.row);
      const compFormula = "IF(" + overtimeCell + "-" + claimCell + "<0,0," + overtimeCell + "-" + claimCell + ")";
      setFormulaAndNum(doc, ref(du.compCol, r.dp.row), compFormula, 0, styleAt[compColName]);
      cells++;
      const amtFormula = "MIN(" + claimCell + "*" + wageCell + ",19000)";
      setFormulaAndNum(doc, ref(du.amtCol, r.dp.row), amtFormula, 0, styleAt[amtColName]);
      cells++;
      continue;
    }
    if(!r.willFill) continue;

    // 寫入每日服勤時數
    du.dayCols.forEach((col, i) => {
      setNum(doc, ref(col, r.dp.row), r.vals[i], styleAt[colName(col)]);
      cells++;
    });

    // 寫入 E 欄（請休時數）
    setNum(doc, ref(5, r.dp.row), r.leaveHours, styleAt["E"] || styleAt[firstColName]);
    cells++;

    // 寫入 F 欄（補休時數）
    setNum(doc, ref(6, r.dp.row), r.compLeaveHours, styleAt["F"] || styleAt[firstColName]);
    cells++;

    // 寫入時數總和（保留 SUM 公式並寫入真實數值）
    const sumFormula = "SUM(" + firstColName + r.dp.row + ":" + lastDayColName + r.dp.row + ")";
    setFormulaAndNum(doc, ref(du.sumCol, r.dp.row), sumFormula, r.sum, styleAt[sumColName]);
    cells++;

    // 寫入超勤時數（下層列，首欄；依當月基準工時計算公式並寫入數值，最低為 0）
    const otFormula = "MAX(0," + sumColName + r.dp.row + "-" + du.baseHours + "+E" + r.dp.row + "+F" + r.dp.row + ")";
    setFormulaAndNum(doc, ref(firstCol, r.dp.row + 1), otFormula, Math.max(0, r.overtimeHours), styleAt[firstColName]);
    cells++;

    // 更新支領時數公式與數值（超勤時數小於上限時以超勤時數為主，實報實銷）
    const overtimeCell = ref(firstCol, r.dp.row + 1);
    const wageCell = ref(du.wageCol, r.dp.row);
    const claimFormula = "MAX(0,MIN(" + overtimeCell + ",ROUNDUP(19000/" + wageCell + ",0)))";
    setFormulaAndNum(doc, ref(du.claimCol, r.dp.row), claimFormula, r.claimHours, styleAt[claimColName]);
    cells++;

    // 寫入超勤補休時數（保留 IF 公式並寫入真實數值）
    const claimCell = ref(du.claimCol, r.dp.row);
    const compFormula = "IF(" + overtimeCell + "-" + claimCell + "<0,0," + overtimeCell + "-" + claimCell + ")";
    const compVal = Math.max(0, r.overtimeHours - r.claimHours);
    setFormulaAndNum(doc, ref(du.compCol, r.dp.row), compFormula, compVal, styleAt[compColName]);
    cells++;

    // 寫入超勤金額（保留 MIN 公式並寫入真實金額）
    const amtFormula = "MIN(" + claimCell + "*" + wageCell + ",19000)";
    setFormulaAndNum(doc, ref(du.amtCol, r.dp.row), amtFormula, r.claimAmount, styleAt[amtColName]);
    cells++;
  }

  // 3. 更新表尾合計總金額與國字大寫
  const totalAmount = plan.rows.reduce((acc, row) =>
    acc + (!row.ignored && !row.isRosterOnly && row.willFill ? row.claimAmount : (!row.ignored && !row.isRosterOnly && row.dp.wage > 0 ? Math.min(19000, row.claimHours * row.dp.wage) : 0)), 0);

  for(const c of doc.getElementsByTagName("c")){
    const fn = c.getElementsByTagName("f")[0];
    if(fn && fn.textContent.includes("SUM(" + amtColName)){
      let vn = c.getElementsByTagName("v")[0];
      if(!vn){ vn = doc.createElementNS(XLNS, "v"); c.appendChild(vn); }
      vn.textContent = String(totalAmount);
      cells++;
    }
    if(fn && (fn.textContent.includes("SUBSTITUTE(") || fn.textContent.includes("TEXT(SUM("))){
      c.setAttribute("t", "str");
      let vn = c.getElementsByTagName("v")[0];
      if(!vn){ vn = doc.createElementNS(XLNS, "v"); c.appendChild(vn); }
      vn.textContent = toChineseCurrency(totalAmount);
      cells++;
    }
  }

  return cells;
}

async function cleanCalcChain(zip){
  zip.remove("xl/calcChain.xml");
  if(zip.file("xl/_rels/workbook.xml.rels")){
    let rels = await zip.file("xl/_rels/workbook.xml.rels").async("string");
    rels = rels.replace(/<Relationship[^>]*Target="[^"]*calcChain\.xml"[^>]*\/>/g, "");
    rels = rels.replace(/<Relationship[^>]*Type="[^"]*relationships\/calcChain"[^>]*\/>/g, "");
    zip.file("xl/_rels/workbook.xml.rels", rels);
  }
  if(zip.file("[Content_Types].xml")){
    let ct = await zip.file("[Content_Types].xml").async("string");
    ct = ct.replace(/<Override[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/g, "");
    zip.file("[Content_Types].xml", ct);
  }
}

async function cleanSingleSheetWorkbook(zip, keepSheetName){
  let wbXml = await zip.file("xl/workbook.xml").async("string");
  let relsXml = await zip.file("xl/_rels/workbook.xml.rels").async("string");
  let ctXml = await zip.file("[Content_Types].xml").async("string");

  const reSheet = /<sheet [^>]*name="([^"]+)"[^>]*sheetId="([^"]+)"[^>]*r:id="([^"]+)"[^>]*\/>/g;
  const sheets = [];
  let m;
  while((m = reSheet.exec(wbXml)) !== null){
    sheets.push({ raw: m[0], name: m[1], sheetId: m[2], rId: m[3] });
  }

  const toRemove = sheets.filter(s => s.name !== keepSheetName);
  for(const s of toRemove){
    wbXml = wbXml.replace(s.raw, "");
    const relMatch = relsXml.match(new RegExp(`<Relationship [^>]*Id="${s.rId}"[^>]*Target="([^"]+)"[^>]*\\/>`));
    if(relMatch){
      const targetPath = relMatch[1].startsWith("/") ? relMatch[1].slice(1) : ("xl/" + relMatch[1].replace(/^\.\.\//, ""));
      zip.remove(targetPath);
      zip.remove(targetPath.replace("worksheets/", "worksheets/_rels/") + ".rels");
      relsXml = relsXml.replace(relMatch[0], "");
      const ctPart = "/" + targetPath;
      const ctRe = new RegExp(`<Override [^>]*PartName="${ctPart.replace("/", "\\/")}"[^>]*\\/>`);
      ctXml = ctXml.replace(ctRe, "");
    }
  }

  // 重設活躍分頁為第一頁，避免 Excel 參照越界
  wbXml = wbXml.replace(/activeTab="\d+"/g, 'activeTab="0"');

  await cleanCalcChain(zip);

  zip.file("xl/workbook.xml", wbXml);
  zip.file("xl/_rels/workbook.xml.rels", relsXml);
  zip.file("[Content_Types].xml", ctXml);
}

function matchPeopleFor(ra, du){
  const byKey = new Map();
  ra.people.forEach((p, i) => {
    const a = byKey.get(p.key) || [];
    a.push({ ...p, idx: i });
    byKey.set(p.key, a);
  });
  const ptr = new Map();
  return du.people.map(p => {
    const arr = byKey.get(p.key);
    if(!arr || !arr.length) return { dp: p, match: null, dup: false };
    const cur = ptr.get(p.key) || 0;
    const match = arr[cur] || null;
    ptr.set(p.key, cur + 1);
    return { dp: p, match, dup: arr.length > 1 };
  });
}

function buildRulesForRoster(ra, ignoredSet){
  const seen = new Map();
  ra.marks.forEach((row, i) => {
    const p = ra.people[i];
    if(p && ignoredSet.has(p.key)) return;
    row.forEach(sym => {
      const e = seen.get(sym) || { sym, n: 0 };
      e.n++; seen.set(sym, e);
    });
  });
  return [...seen.values()].map(e => {
    const c = classify(e.sym);
    return { ...e, ...c, hours: c.hours };
  });
}

function computeBatchPlan(ra, du, scope, ignoredSet, monthRules){
  const pairs = matchPeopleFor(ra, du);
  const rows = pairs.map(({ dp, match, dup }) => {
    const isIgnored = ignoredSet.has(dp.key);
    if(isIgnored){
      return {
        dp, match, dup, ignored: true,
        vals: new Array(du.days).fill(0),
        sum: 0, leaveHours: 0, compLeaveHours: 0,
        overtimeHours: 0, claimHours: 0, claimAmount: 0, compOvertime: 0,
        willFill: false, changes: []
      };
    }
    if(!match){
      return {
        dp, match: null, dup, ignored: false,
        vals: dp.vals, sum: dp.sum, leaveHours: dp.leaveHours, compLeaveHours: dp.compLeaveHours,
        overtimeHours: 0, claimHours: 0, claimAmount: 0, compOvertime: 0,
        willFill: false, changes: []
      };
    }

    const marks = ra.marks[match.idx] || [];
    const vals = marks.map(sym => {
      const r = monthRules.find(x => x.sym === sym);
      return r ? r.hours : classify(sym).hours;
    });

    const sum = vals.reduce((a, b) => a + b, 0);
    const leaveHours = (match && ra.leaveHours && ra.leaveHours[match.idx] != null) ? ra.leaveHours[match.idx] : (dp.leaveHours || 0);
    const compLeaveHours = (match && ra.compLeaveHours && ra.compLeaveHours[match.idx] != null) ? ra.compLeaveHours[match.idx] : (dp.compLeaveHours || 0);
    const overtimeHours = Math.max(0, sum - du.baseHours + leaveHours + compLeaveHours);
    const claimCap = dp.wage > 0 ? Math.ceil(19000 / dp.wage) : 0;
    const claimHours = Math.max(0, Math.min(overtimeHours, claimCap));
    const claimAmount = Math.min(claimHours * dp.wage, 19000);
    const compOvertime = Math.max(0, overtimeHours - claimHours);

    const willFill = (scope !== "audit");

    return {
      dp, match, dup, ignored: false,
      vals, sum, leaveHours, compLeaveHours,
      overtimeHours, claimHours, claimAmount, compOvertime,
      willFill, changes: []
    };
  });

  return { rows };
}

function updateExportScopeUI(common){
  const curName = S.sheet || (common && common.length ? common[0] : "");
  const count = common ? common.length : 1;

  const opts = [
    { val: "current_only", text: "僅目前月份（單一工作表：" + curName + "）" },
    { val: "all_months", text: "全部月份（批次轉換回填全部 " + count + " 個月份）" },
    { val: "current_keep_all", text: "完整活頁簿（僅更新 " + curName + "，其餘保留）" }
  ];

  const html = opts.map(o => '<option value="' + o.val + '">' + esc(o.text) + "</option>").join("");
  if($("selExportScope")) $("selExportScope").innerHTML = html;
  if($("barSelExportScope")) $("barSelExportScope").innerHTML = html;

  const saved = localStorage.getItem("chaoqin_export_scope") || "current_only";
  setExportScope(saved);

  if($("barExportScopeWrap")) $("barExportScopeWrap").classList.remove("hide");
  if($("grpExportScope")) $("grpExportScope").classList.remove("hide");
}

function setExportScope(val){
  if($("selExportScope")) $("selExportScope").value = val;
  if($("barSelExportScope")) $("barSelExportScope").value = val;
  try{ localStorage.setItem("chaoqin_export_scope", val); }catch(e){}
}

async function exportFile(){
  const scopeEl = document.querySelector('input[name=scope]:checked');
  const scope = scopeEl ? scopeEl.value : "blank";
  if(scope === "audit"){
    const ok = confirm("【檢核模式提示】\n您目前選擇「2. 檢核標示差異（不動原檔）」模式。\n在此模式下匯出，檔案將維持上傳原檔數值，不會覆蓋任何服勤時數與數字。\n\n• 若要覆蓋為輪休表最新正確數值，請按「取消」並在上方改選「3. 強制覆蓋重填」後再匯出。\n• 若確定要維持原檔數值匯出，請按「確定」。");
    if(!ok) return;
  }
  const btn = $("btnGo"); const old = btn.textContent;
  btn.disabled = true; btn.textContent = "產生中…";

  const exportScope = localStorage.getItem("chaoqin_export_scope") || "current_only";

  try{
    const zip = await JSZip.loadAsync(S.B.buf);
    const P = new DOMParser(), Ser = new XMLSerializer();

    const nameA = new Map(S.A.sheets.map(s => [s.name, s]));
    const commonSheets = S.B.sheets.filter(s => nameA.has(s.name));

    let totalCells = 0;

    if(exportScope === "all_months" && commonSheets.length > 1){
      // 批次轉換回填全部共有月份
      for(let idx = 0; idx < commonSheets.length; idx++){
        const sInfo = commonSheets[idx];
        btn.textContent = "批次處理 (" + (idx + 1) + "/" + commonSheets.length + ")…";
        if(sInfo.name === S.sheet){
          const doc = P.parseFromString(await zip.file(sInfo.path).async("string"), "application/xml");
          const ym = parseYearMonth(sInfo.name);
          const ymInfo = { year: S.year || ym.year, month: S.month || ym.month, weeks: S.weeks };
          const cCount = fillDocSheet(doc, sInfo.name, S.ra, S.du, S.plan, ymInfo, scope === "audit");
          totalCells += cCount;
          let xmlStr = Ser.serializeToString(doc).replace(/^(\s*<\?xml[^>]*\?>\s*)+/i, "");
          zip.file(sInfo.path, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xmlStr);
        }else{
          const pathA = nameA.get(sInfo.name).path;
          const pathB = sInfo.path;
          const ym = parseYearMonth(sInfo.name);
          const mi = getMonthInfo(ym.year, ym.month);
          const raOther = parseRoster(await readCells(S.A, pathA));
          const duOther = parseDuty(await readCells(S.B, pathB));
          const ignoredSet = getIgnoredSet();
          const otherRules = buildRulesForRoster(raOther, ignoredSet);
          const planOther = computeBatchPlan(raOther, duOther, scope, ignoredSet, otherRules);
          const docOther = P.parseFromString(await zip.file(pathB).async("string"), "application/xml");
          const ymInfoOther = { year: ym.year, month: ym.month, weeks: mi.weeks };
          const cCount = fillDocSheet(docOther, sInfo.name, raOther, duOther, planOther, ymInfoOther, scope === "audit");
          totalCells += cCount;
          let xmlStr = Ser.serializeToString(docOther).replace(/^(\s*<\?xml[^>]*\?>\s*)+/i, "");
          zip.file(pathB, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xmlStr);
        }
      }
      await cleanCalcChain(zip);
    }else{
      // 單月份處理（目前選定月份）
      const sPath = S.B.sheets.find(s => s.name === S.sheet).path;
      const doc = P.parseFromString(await zip.file(sPath).async("string"), "application/xml");
      const ym = parseYearMonth(S.sheet);
      const ymInfo = { year: S.year || ym.year, month: S.month || ym.month, weeks: S.weeks };
      totalCells = fillDocSheet(doc, S.sheet, S.ra, S.du, S.plan, ymInfo, scope === "audit");
      let xmlStr = Ser.serializeToString(doc).replace(/^(\s*<\?xml[^>]*\?>\s*)+/i, "");
      zip.file(sPath, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xmlStr);

      if(exportScope === "current_only"){
        // 僅保留目前月份工作表，自動剔除其他月份與多餘分頁
        await cleanSingleSheetWorkbook(zip, S.sheet);
      }else{
        await cleanCalcChain(zip);
      }
    }

    // 促使 Excel 開檔時依新資料重算公式
    const wbPath = "xl/workbook.xml";
    let wbx = await zip.file(wbPath).async("string");
    if(/<calcPr\b/.test(wbx)){
      wbx = wbx.replace(/<calcPr\b([^>]*?)\/?>/, (mm, attrs) =>
        "<calcPr" + attrs.replace(/\s*fullCalcOnLoad="[^"]*"/, "") + ' fullCalcOnLoad="1"/>');
    }else{
      wbx = wbx.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>');
    }
    zip.file(wbPath, wbx);

    const blob = await zip.generateAsync({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      compression: "DEFLATE"
    });

    const unit = extractUnitName();
    const ym = parseYearMonth(S.sheet || "");
    const y = S.year || ym.year || 115;
    const m = S.month || ym.month || 9;
    const ymStr = y + "年" + String(m).padStart(2, "0") + "月份";

    let name = "";
    if(exportScope === "all_months" && commonSheets.length > 1){
      name = `${unit}${y}年度超勤加班費申請及印領清冊(全部月份).xlsx`;
    }else{
      name = `${unit}${ymStr}超勤加班費申請及印領清冊.xlsx`;
    }

    let saved = false, errMsg = "";
    try{
      const dl = (window.claude && claude.use) ? await claude.use("downloads") : null;
      if(dl){ await dl.save({ filename: name, data: blob }); saved = true; }
    }catch(e){
      if(e && e.code === "declined"){ $("barMsg").textContent = "已取消，沒有存檔。"; return; }
      errMsg = (e && e.message) ? e.message : "";
    }
    if(!saved){
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = name; document.body.appendChild(a); a.click();
      setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 4000);
    }

    let descScope = "";
    if(exportScope === "all_months" && commonSheets.length > 1){
      descScope = "（批次完成全部 " + commonSheets.length + " 個月份工作表，共寫入 " + totalCells + " 格）";
    }else if(exportScope === "current_only"){
      descScope = "（已自動排除其他未編輯月份，輸出單月共寫入 " + totalCells + " 格）";
    }else{
      descScope = "（保留其他工作表原樣，當月共寫入 " + totalCells + " 格）";
    }

    $("barMsg").innerHTML = "已產生 <strong>" + esc(name) + "</strong> " + descScope + "。時數總和、超勤金額與合計均已同步計算！" +
      (errMsg ? " <span style='color:var(--warn)'>（" + esc(errMsg) + "）</span>" : "");
  }catch(err){
    alert("產生檔案時出錯：" + err.message);
  }finally{
    btn.textContent = old; btn.disabled = false;
  }
}

/* ================= 啟動 ================= */
wireDrop($("dropA"), $("fileA"), "A");
wireDrop($("dropB"), $("fileB"), "B");

const btnHome = $("btnHome");
if(btnHome) btnHome.addEventListener("click", resetToHome);

$("selYear").addEventListener("change", () => updateMonthConfig(true));
$("selMonth").addEventListener("change", () => updateMonthConfig(true));

const triEl = $("inpBaseTri");
if(triEl){
  triEl.addEventListener("input", e => {
    triEl.dataset.userEdited = "true";
    const v = Number(e.target.value);
    if(v > 0) S.baseTri = v;
    try{ localStorage.setItem("chaoqin_base_tri", triEl.value); }catch(err){}
    updateGongCalculator();
    if(S.plan){
      buildPlan();
      renderGrid();
      renderFinds();
      renderBar();
    }
  });
}

const bInpEl = $("inpBaseHours");
if(bInpEl){
  bInpEl.addEventListener("input", e => {
    bInpEl.dataset.userEdited = "true";
    const v = Number(e.target.value);
    if(v > 0){
      S.baseHours = v;
      if(S.du) S.du.baseHours = v;
      try{ localStorage.setItem("chaoqin_base_hours", bInpEl.value); }catch(err){}
      if(S.plan){
        buildPlan();
        renderGrid();
        renderFinds();
        renderBar();
      }
    }
  });
}

const btnDl30 = $("btnDl30");
if(btnDl30) btnDl30.addEventListener("click", () => downloadTemplate(30));
const btnDl31 = $("btnDl31");
if(btnDl31) btnDl31.addEventListener("click", () => downloadTemplate(31));

// 本休試算獨立彈出視窗（Modal）事件綁定
const btnOpenCalc = $("btnOpenCalc");
if(btnOpenCalc) btnOpenCalc.addEventListener("click", openCalcModal);
const btnFabCalc = $("btnFabCalc");
if(btnFabCalc) btnFabCalc.addEventListener("click", openCalcModal);

const btnCloseCalc = $("btnCloseCalc");
if(btnCloseCalc) btnCloseCalc.addEventListener("click", closeCalcModal);
const btnDoneCalc = $("btnDoneCalc");
if(btnDoneCalc) btnDoneCalc.addEventListener("click", closeCalcModal);
const modalBdrop = $("calcModalBackdrop");
if(modalBdrop) modalBdrop.addEventListener("click", closeCalcModal);
window.addEventListener("keydown", e => { if(e.key === "Escape") closeCalcModal(); });

const modalDays = $("modalDays");
if(modalDays){
  modalDays.addEventListener("change", () => {
    const d = Number(modalDays.value) || 31;
    const triInp = $("modalBaseTri");
    if(triInp){
      triInp.value = d >= 31 ? 18 : (d === 30 ? 17 : 16);
    }
    updateModalCalculator();
  });
}
const modalBaseTri = $("modalBaseTri");
if(modalBaseTri) modalBaseTri.addEventListener("input", updateModalCalculator);
const modalGongDays = $("modalGongDays");
if(modalGongDays) modalGongDays.addEventListener("input", updateModalCalculator);
const modalIsContinuous = $("modalIsContinuous");
if(modalIsContinuous) modalIsContinuous.addEventListener("change", updateModalCalculator);
const modalCrossMon = $("modalCrossMon");
if(modalCrossMon) modalCrossMon.addEventListener("change", updateModalCalculator);

const igInp = $("inpIgnoredNames");
if(igInp){
  igInp.addEventListener("keydown", e => {
    if(e.key === "Enter"){
      e.preventDefault();
      const b = $("btnAddIgnored");
      if(b) b.click();
    }
  });
}
const btnAddIg = $("btnAddIgnored");
if(btnAddIg){
  btnAddIg.addEventListener("click", () => {
    const inp = $("inpIgnoredNames");
    if(!inp) return;
    const addNames = inp.value.split(/[,，、\s]+/).map(s => cjk(s).trim()).filter(Boolean);
    if(!addNames.length) return;
    const s = getIgnoredSet();
    addNames.forEach(n => s.add(n));
    setIgnoredNames(Array.from(s).join("、"));
    inp.value = "";
    if(S.plan){ buildPlan(); renderGrid(); renderFinds(); renderBar(); }
  });
}
const btnClearIg = $("btnClearIgnored");
if(btnClearIg){
  btnClearIg.addEventListener("click", () => {
    setIgnoredNames("");
    if(S.plan){ buildPlan(); renderGrid(); renderFinds(); renderBar(); }
  });
}

const chkHide = $("chkHideIgnored");
if(chkHide){
  try{
    const savedHide = localStorage.getItem("chaoqin_hide_ignored");
    if(savedHide !== null) chkHide.checked = (savedHide === "true");
  }catch(e){}
  chkHide.addEventListener("change", () => {
    try{ localStorage.setItem("chaoqin_hide_ignored", chkHide.checked ? "true" : "false"); }catch(e){}
    if(S.plan){ renderGrid(); }
  });
}

$("selSheet").addEventListener("change", e => { S.sheet = e.target.value; S.over = {}; S.leaveOver = {}; refresh(); });

document.querySelectorAll('input[name=scope]').forEach(el =>
  el.addEventListener("change", () => {
    try{ localStorage.setItem("chaoqin_scope", el.value); }catch(e){}
    if(S.plan){ buildPlan(); renderGrid(); renderFinds(); renderBar(); }
  }));
$("btnGo").addEventListener("click", exportFile);

const selExp = $("selExportScope");
if(selExp) selExp.addEventListener("change", e => setExportScope(e.target.value));
const barSelExp = $("barSelExportScope");
if(barSelExp) barSelExp.addEventListener("change", e => setExportScope(e.target.value));

// 頁面初始化與狀態自動恢復（支援重新整理保持檔案與輸入）
async function tryRestoreState(){
  try{
    const y = localStorage.getItem("chaoqin_year");
    if(y && $("selYear")) $("selYear").value = y;
    const m = localStorage.getItem("chaoqin_month");
    if(m && $("selMonth")) $("selMonth").value = m;
    const bh = localStorage.getItem("chaoqin_base_hours");
    if(bh && $("inpBaseHours")) { $("inpBaseHours").value = bh; $("inpBaseHours").dataset.userEdited = "true"; }
    const bt = localStorage.getItem("chaoqin_base_tri");
    if(bt && $("inpBaseTri")) { $("inpBaseTri").value = bt; $("inpBaseTri").dataset.userEdited = "true"; }
    const sc = localStorage.getItem("chaoqin_scope");
    if(sc){
      const el = document.querySelector(`input[name=scope][value="${sc}"]`);
      if(el) el.checked = true;
    }
    const ov = localStorage.getItem("chaoqin_over");
    if(ov) S.over = JSON.parse(ov);
    const lo = localStorage.getItem("chaoqin_leaveOver");
    if(lo) S.leaveOver = JSON.parse(lo);
  }catch(e){}

  renderShieldChips();
  updateMonthConfig(false);

  // 自動從 IndexedDB 恢復上次上傳之檔案（重新整理不遺失資料）
  try{
    const fA = await getFileFromDb("A");
    const fB = await getFileFromDb("B");
    if(fA && fB){
      await take(fA, "A");
      await take(fB, "B");
    }else if(fA){
      await take(fA, "A");
    }else if(fB){
      await take(fB, "B");
    }
  }catch(e){
    console.warn("Restore files from DB failed:", e);
  }
}

tryRestoreState();
