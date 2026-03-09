const pptxgen = require("pptxgenjs");
const fs = require("fs");

const DARK_BG = "0D1B2A";
const DARKER  = "071117";
const CYAN    = "00E5FF";
const ORANGE  = "FF9800";
const YELLOW  = "FFD700";
const GREEN   = "4CAF50";
const RED     = "F44336";
const WHITE   = "FFFFFF";
const GRAY    = "546E7A";
const LGRAY   = "90A4AE";
const TEAL    = "00BFA5";
const LBLUE   = "29B6F6";
const W = 10, H = 5.625;

function toB64(path) {
  const buf = fs.readFileSync(path);
  return `image/png;base64,${buf.toString('base64')}`;
}
const charts = {
  c1: toB64('/home/claude/chart1_en.png'),
  c2: toB64('/home/claude/chart2_en.png'),
  c3: toB64('/home/claude/chart3_en.png'),
  c4a: toB64('/home/claude/chart4a_en.png'),
  c4b: toB64('/home/claude/chart4b_en.png'),
  c5: toB64('/home/claude/chart5_en.png'),
  c6: toB64('/home/claude/chart6_en.png'),
  c7: toB64('/home/claude/chart7_en.png'),
};

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.title = 'Gwoniri Wind Resource Assessment Report 2025';

function addSlide() { let sl = pres.addSlide(); sl.background = { color: DARK_BG }; return sl; }
function accentBar(sl, y, h, color) {
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: y, w: 0.06, h: h, fill: { color: color||CYAN }, line: { type: 'none' } });
}
function secBadge(sl, num, x, y) {
  sl.addShape(pres.shapes.OVAL, { x, y, w: 0.4, h: 0.4, fill: { color: CYAN }, line: { type: 'none' } });
  sl.addText(String(num), { x, y, w: 0.4, h: 0.4, fontSize: 13, bold: true, color: DARK_BG, align: 'center', valign: 'middle', margin: 0 });
}
function hRule(sl, x, y, w) {
  sl.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.02, fill: { color: CYAN, transparency: 40 }, line: { type: 'none' } });
}
function statBox(sl, x, y, w, h, label, value, unit, color) {
  sl.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: DARKER }, line: { color: color||CYAN, pt: 1 } });
  sl.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.05, h, fill: { color: color||CYAN }, line: { type: 'none' } });
  sl.addText(label, { x: x+0.1, y: y+0.05, w: w-0.15, h: 0.22, fontSize: 8.5, color: LGRAY, margin: 0 });
  sl.addText(value, { x: x+0.1, y: y+0.22, w: w-0.15, h: 0.45, fontSize: 20, bold: true, color: color||CYAN, margin: 0 });
  if (unit) sl.addText(unit, { x: x+0.1, y: y+0.62, w: w-0.15, h: 0.18, fontSize: 8, color: LGRAY, margin: 0 });
}

// ── SLIDE 1: Cover ────────────────────────────────────────
{
  let sl = addSlide();
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 3.8, h: H, fill: { color: DARKER }, line: { type: 'none' } });
  sl.addShape(pres.shapes.RECTANGLE, { x: 3.8, y: 0, w: 0.04, h: H, fill: { color: CYAN }, line: { type: 'none' } });

  sl.addText('Wind Resource\nAssessment Report', {
    x: 0.3, y: 0.9, w: 3.2, h: 1.7, fontSize: 28, bold: true, color: WHITE, align: 'left', lineSpacingMultiple: 1.2
  });
  sl.addText('2025  (Jan 01 – Dec 31, 2025)', {
    x: 0.3, y: 2.68, w: 3.2, h: 0.32, fontSize: 11, color: CYAN, bold: true
  });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 3.05, w: 2.5, h: 0.03, fill: { color: CYAN, transparency: 30 }, line: { type: 'none' } });
  sl.addText('Site: Gwoniri, Munmudaewang-myeon, Gyeongju, Gyeongbuk', { x: 0.3, y: 3.12, w: 3.2, h: 0.25, fontSize: 8.5, color: LGRAY });
  sl.addText('Tower: 100m Lattice  NRG SymphoniePRO  S/N: 820616523', { x: 0.3, y: 3.34, w: 3.2, h: 0.25, fontSize: 8.5, color: LGRAY });
  sl.addText('Client: KS Power Co., Ltd.', { x: 0.3, y: 3.56, w: 3.2, h: 0.22, fontSize: 8.5, color: LGRAY });
  sl.addText('Standard: IEC 61400-1 Ed.4  /  KS C IEC 61400-12', { x: 0.3, y: 3.78, w: 3.2, h: 0.22, fontSize: 8.5, color: LGRAY });

  const stats = [
    ['Mean Wind Speed', '6.342 m/s', '@ 100m  ✓', CYAN, '✓'],
    ['Annual WPD', '281.8 W/m²', '@ 100m  ✓', ORANGE, '✓'],
    ['Weibull k/C', 'k=2.093  C=7.172', '@ 100m', LBLUE, '✓'],
    ['Prevailing Dir.', 'WNW  24.1%', 'Energy 20.4%', TEAL, '✓'],
    ['Wind Shear α', '0.222', 'IEC 0.14~0.20 ⚠', YELLOW, '⚠'],
    ['TI P90 @15m/s', '0.2300', 'IEC Class A: 0.16 ✗', RED, '✗'],
    ['Extreme V50', '42.30 m/s', 'IEC Class III: 37.5 ✗', RED, '✗'],
    ['Data Recovery', '100.0%', '2025.01~12', GREEN, '✓'],
  ];
  const boxW = 2.82, boxH = 0.88, gap = 0.06, RX = 4.0;
  stats.forEach((s, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const bx = RX + col * (boxW + gap);
    const by = 0.18 + row * (boxH + gap);
    statBox(sl, bx, by, boxW, boxH, s[0], s[1], s[2], s[3]);
    const bc = s[4]==='✓' ? GREEN : (s[4]==='✗' ? RED : YELLOW);
    sl.addShape(pres.shapes.OVAL, { x: bx+boxW-0.32, y: by+0.04, w: 0.25, h: 0.25, fill: { color: bc }, line: { type: 'none' } });
    sl.addText(s[4], { x: bx+boxW-0.32, y: by+0.04, w: 0.25, h: 0.25, fontSize: 10, bold: true, color: DARK_BG, align: 'center', valign: 'middle', margin: 0 });
  });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 4.9, w: 3.2, h: 0.4, fill: { color: CYAN, transparency: 75 }, line: { type: 'none' } });
  sl.addText('1-Year Measurement Analysis  (2025.01.01 – 2025.12.31)', { x: 0.3, y: 4.9, w: 3.2, h: 0.4, fontSize: 9.5, color: CYAN, align: 'center', valign: 'middle', bold: true });
}

// ── SLIDE 2: Measurement System ──────────────────────────
{
  let sl = addSlide();
  accentBar(sl, 0, H, CYAN);
  sl.addText('Measurement System Overview', { x: 0.15, y: 0.1, w: 7, h: 0.42, fontSize: 20, bold: true, color: WHITE, margin: 0 });
  secBadge(sl, '01', 9.45, 0.1);
  sl.addText('NRG SymphoniePRO  ·  100m Lattice Tower  ·  Channel Configuration', { x: 0.15, y: 0.52, w: 9.2, h: 0.28, fontSize: 10, color: LGRAY, italic: true });
  hRule(sl, 0.15, 0.78, 9.5);

  sl.addText('NRG SymphoniePRO logger (S/N: 820616523) on a 100m lattice tower. 2025 data recovery 100.0% — no gaps throughout the year. Air density 1.177 kg/m³ (−4.1% vs standard). Mean temperature 13.6°C, pressure 969.0 hPa. Excellent data quality for direct statistical analysis.', {
    x: 0.15, y: 0.88, w: 5.6, h: 0.85, fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });

  const infoData = [
    ['Measurement Period', '2025.01.01  ~  2025.12.31  (1-year)'],
    ['Location', '35.867740°N  129.414420°E  /  Elev. 390m ASL'],
    ['Tower Type', '100m Lattice  /  GPS-surveyed position'],
    ['Data Logger', 'NRG SymphoniePRO  (S/N: 820616523)'],
    ['Sampling Interval', '10-min averages  (144 records/day)'],
    ['Total Records', '52,559  (2025.01 ~ 2025.12)'],
    ['Data Recovery', '100.0%  (52,559 / 52,560)'],
    ['Air Density', '1.177 kg/m³  (T=13.6°C, P=969.0 hPa)'],
  ];
  infoData.forEach((row, i) => {
    const iy = 1.78 + i * 0.33;
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.15, y: iy, w: 1.7, h: 0.28, fill: { color: CYAN, transparency: 75 }, line: { type: 'none' } });
    sl.addText(row[0], { x: 0.18, y: iy, w: 1.64, h: 0.28, fontSize: 8.5, bold: true, color: CYAN, valign: 'middle', margin: 0 });
    sl.addText(row[1], { x: 1.9, y: iy, w: 4.0, h: 0.28, fontSize: 8.5, color: WHITE, valign: 'middle', margin: 0 });
  });

  // Sensor table
  sl.addShape(pres.shapes.RECTANGLE, { x: 5.8, y: 0.88, w: 3.9, h: 0.28, fill: { color: CYAN }, line: { type: 'none' } });
  const hdr = ['Ch','Sensor Type','Height','Orientation','Parameter'];
  const hw = [0.35, 1.3, 0.55, 0.72, 0.97];
  let hx = 5.83;
  hdr.forEach((h, i) => {
    sl.addText(h, { x: hx, y: 0.88, w: hw[i], h: 0.28, fontSize: 8, bold: true, color: DARK_BG, valign: 'middle', margin: 0 });
    hx += hw[i];
  });
  const tableData = [
    ['A1', 'Anemometer (Thies)', '100m', '35° (NE)', 'Speed Avg/SD/Gust'],
    ['A2', 'Anemometer (Thies)', '100m', '215° (SW)', 'Speed Avg/SD/Gust'],
    ['A3', 'Anemometer (Thies)', '80m', '35° (NE)', 'Speed Avg/SD/Gust'],
    ['A4', 'Anemometer (Thies)', '80m', '215° (SW)', 'Speed Avg/SD/Gust'],
    ['A5', 'Anemometer (Thies)', '60m', '35° (NE)', 'Speed Avg/SD/Gust'],
    ['A6', 'Anemometer (Thies)', '60m', '215° (SW)', 'Speed Avg/SD/Gust'],
    ['V1', 'Wind Vane (Thies)', '95m', '35° (NE)', 'Direction Avg'],
    ['V2', 'Wind Vane (Thies)', '75m', '35° (NE)', 'Direction Avg'],
    ['V3', 'Wind Vane (Thies)', '55m', '35° (NE)', 'Direction Avg'],
    ['T1', 'Thermometer (NRG T60)', '5m', '—', 'Temperature Avg'],
    ['B1', 'Barometer (NRG BP65)', '2m', '—', 'Pressure Avg'],
    ['H1', 'Hygrom. (NRG RH5XC)', '2m', '—', 'Humidity Avg'],
  ];
  tableData.forEach((row, ri) => {
    const ry = 1.18 + ri * 0.34;
    const bg = ri % 2 === 0 ? '0F2840' : '071117';
    sl.addShape(pres.shapes.RECTANGLE, { x: 5.8, y: ry, w: 3.9, h: 0.32, fill: { color: bg }, line: { type: 'none' } });
    let rx = 5.83;
    row.forEach((cell, ci) => {
      sl.addText(cell, { x: rx, y: ry, w: hw[ci], h: 0.32, fontSize: 7.8, color: ci === 0 ? CYAN : WHITE, valign: 'middle', margin: 0 });
      rx += hw[ci];
    });
  });
}

// ── SLIDE 3: Mean Wind Speed ──────────────────────────────
{
  let sl = addSlide();
  accentBar(sl, 0, H, LBLUE);
  sl.addText('Mean Wind Speed Analysis', { x: 0.15, y: 0.1, w: 7, h: 0.42, fontSize: 20, bold: true, color: WHITE, margin: 0 });
  secBadge(sl, '02', 9.45, 0.1);
  sl.addText('Monthly Variation  &  Vertical Profile  ·  2025 Annual  @ 100m', { x: 0.15, y: 0.52, w: 9.2, h: 0.28, fontSize: 10, color: LGRAY, italic: true });
  hRule(sl, 0.15, 0.78, 9.5);
  sl.addText('Annual mean wind speed 6.342 m/s (@ 100m) meets the IEC development threshold (6~7 m/s). Wind shear α=0.222 extrapolates to 6.604 m/s at 120m and 6.939 m/s at 150m. Peak month is February (8.832 m/s) and minimum is August (4.285 m/s), with a seasonal amplitude of 4.5 m/s requiring seasonal energy correction.', {
    x: 0.15, y: 0.88, w: 9.5, h: 0.65, fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });
  sl.addImage({ data: charts.c1, x: 0.1, y: 1.62, w: 7.0, h: 3.7 });
  const callouts = [['Annual Mean','6.342 m/s', LBLUE],['α=0.222','120m: 6.604', YELLOW]];
  callouts.forEach((c, i) => {
    const cx = 7.2, cy = 1.7 + i * 1.3;
    sl.addShape(pres.shapes.RECTANGLE, { x: cx, y: cy, w: 2.55, h: 0.95, fill: { color: DARKER }, line: { color: c[2], pt: 1 } });
    sl.addShape(pres.shapes.RECTANGLE, { x: cx, y: cy, w: 0.05, h: 0.95, fill: { color: c[2] }, line: { type: 'none' } });
    sl.addText(c[0], { x: cx+0.1, y: cy+0.06, w: 2.3, h: 0.25, fontSize: 9, color: LGRAY, margin: 0 });
    sl.addText(c[1], { x: cx+0.1, y: cy+0.3, w: 2.3, h: 0.45, fontSize: 17, bold: true, color: c[2], margin: 0 });
  });
  sl.addShape(pres.shapes.RECTANGLE, { x: 7.2, y: 4.3, w: 2.55, h: 1.0, fill: { color: DARKER }, line: { color: GRAY, pt: 1 } });
  sl.addText([
    { text: 'Peak month  Feb 8.832 m/s\n', options: { color: CYAN, fontSize: 9.5, bold: true } },
    { text: 'Min month  Aug 4.285 m/s\n', options: { color: ORANGE, fontSize: 9.5, bold: true } },
    { text: 'Seasonal amplitude  4.547 m/s', options: { color: LGRAY, fontSize: 8.5 } }
  ], { x: 7.25, y: 4.35, w: 2.45, h: 0.9, margin: 0 });
}

// ── SLIDE 4: Wind Rose ────────────────────────────────────
{
  let sl = addSlide();
  accentBar(sl, 0, H, TEAL);
  sl.addText('Wind Direction Distribution  (Wind Rose)', { x: 0.15, y: 0.1, w: 7, h: 0.42, fontSize: 20, bold: true, color: WHITE, margin: 0 });
  secBadge(sl, '03', 9.45, 0.1);
  sl.addText('Annual Frequency  &  Energy Wind Rose  @  95m  ·  2025 Annual', { x: 0.15, y: 0.52, w: 9.2, h: 0.28, fontSize: 10, color: LGRAY, italic: true });
  hRule(sl, 0.15, 0.78, 9.5);
  sl.addText('Prevailing direction is WNW (West-Northwest) with 24.1% frequency and 20.4% energy concentration. The westerly sector (WNW+NW+W) accounts for about 40.8% of total energy. WNW is recommended as the array orientation reference axis for layout optimization.', {
    x: 0.15, y: 0.88, w: 9.5, h: 0.65, fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });
  sl.addImage({ data: charts.c2, x: 0.1, y: 1.55, w: 7.2, h: 3.8 });
  const wdata = [
    ['Prevailing Dir.', 'WNW (W-Northwest)', 'Freq. 24.1%', TEAL],
    ['Energy Conc.', 'WNW 20.4%', '', ORANGE],
    ['Westerly Energy', 'WNW+NW+W', '40.8%', LBLUE],
    ['Array Opt.', 'WNW axis recommended', '', YELLOW],
  ];
  wdata.forEach((d, i) => {
    const cy = 1.65 + i * 1.0;
    sl.addShape(pres.shapes.RECTANGLE, { x: 7.3, y: cy, w: 2.5, h: 0.82, fill: { color: DARKER }, line: { color: d[3], pt: 1 } });
    sl.addShape(pres.shapes.RECTANGLE, { x: 7.3, y: cy, w: 0.05, h: 0.82, fill: { color: d[3] }, line: { type: 'none' } });
    sl.addText(d[0], { x: 7.38, y: cy+0.04, w: 2.3, h: 0.2, fontSize: 8, color: LGRAY, margin: 0 });
    sl.addText(d[1], { x: 7.38, y: cy+0.24, w: 2.3, h: 0.32, fontSize: 13, bold: true, color: d[3], margin: 0 });
    if (d[2]) sl.addText(d[2], { x: 7.38, y: cy+0.56, w: 2.3, h: 0.2, fontSize: 8.5, color: LGRAY, margin: 0 });
  });
}

// ── SLIDE 5: Weibull ─────────────────────────────────────
{
  let sl = addSlide();
  accentBar(sl, 0, H, LBLUE);
  sl.addText('Weibull Distribution', { x: 0.15, y: 0.1, w: 7, h: 0.42, fontSize: 20, bold: true, color: WHITE, margin: 0 });
  secBadge(sl, '04', 9.45, 0.1);
  sl.addText('Annual & Seasonal PDF  ·  k/C Parameters  @  100m  ·  2025 Annual', { x: 0.15, y: 0.52, w: 9.2, h: 0.28, fontSize: 10, color: LGRAY, italic: true });
  hRule(sl, 0.15, 0.78, 9.5);
  sl.addText('Shape factor k=2.093 indicates moderate dispersion close to Rayleigh distribution. Scale factor C=7.172 m/s exceeds the IEC development threshold (C≥6.0). Winter C=9.137 is highest and Summer C=5.558 is lowest — a seasonal spread of 3.6 m/s. Winter extrapolation to 120m hub height yields C≈9.6 m/s.', {
    x: 0.15, y: 0.88, w: 9.5, h: 0.65, fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });
  sl.addImage({ data: charts.c3, x: 0.1, y: 1.55, w: 6.9, h: 3.8 });
  sl.addShape(pres.shapes.RECTANGLE, { x: 7.1, y: 1.55, w: 2.7, h: 0.32, fill: { color: LBLUE }, line: { type: 'none' } });
  ['Season','k','C (m/s)','Vave'].forEach((h, i) => {
    const xs = [7.12, 7.82, 8.37, 9.07], ws = [0.7, 0.55, 0.7, 0.7];
    sl.addText(h, { x: xs[i], y: 1.55, w: ws[i], h: 0.32, fontSize: 8.5, bold: true, color: DARK_BG, margin: 0 });
  });
  const wrows = [
    ['Annual','2.093','7.172','6.353', CYAN],
    ['Spring','2.226','7.428','6.579', GREEN],
    ['Summer','2.175','5.558','4.923', ORANGE],
    ['Autumn','2.341','6.572','5.823', "FF4081"],
    ['Winter','2.361','9.137','8.097', YELLOW],
  ];
  wrows.forEach((row, ri) => {
    const ry = 1.87 + ri * 0.37;
    sl.addShape(pres.shapes.RECTANGLE, { x: 7.1, y: ry, w: 2.7, h: 0.34, fill: { color: ri%2===0?'0F2840':'071117' }, line: { type: 'none' } });
    const xs=[7.12,7.82,8.37,9.07], ws=[0.7,0.55,0.7,0.7];
    row.slice(0,4).forEach((v,ci) => {
      sl.addText(v, { x: xs[ci], y: ry, w: ws[ci], h: 0.34, fontSize: 8.5, color: ci===0?row[4]:WHITE, valign: 'middle', bold: ri===0, margin: 0 });
    });
  });
  sl.addShape(pres.shapes.RECTANGLE, { x: 7.1, y: 4.0, w: 2.7, h: 0.65, fill: { color: DARKER }, line: { color: LBLUE, pt: 1 } });
  sl.addText('Note: Winter C=9.137 → 120m extrapolation C≈9.6 m/s. Seasonal energy imbalance requires consideration.', {
    x: 7.15, y: 4.05, w: 2.6, h: 0.55, fontSize: 8.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.3
  });
}

// ── SLIDE 6: Wind Shear ───────────────────────────────────
{
  let sl = addSlide();
  accentBar(sl, 0, H, YELLOW);
  sl.addText('Wind Shear', { x: 0.15, y: 0.1, w: 6, h: 0.42, fontSize: 20, bold: true, color: WHITE, margin: 0 });
  secBadge(sl, '05', 9.45, 0.1);
  sl.addText('Power Law Exponent α  ·  Vertical Profile  &  Seasonal  ·  2025 Annual', { x: 0.15, y: 0.52, w: 9.2, h: 0.28, fontSize: 10, color: LGRAY, italic: true });
  hRule(sl, 0.15, 0.78, 9.5);
  sl.addText('Annual wind shear exponent α=0.222 slightly exceeds the IEC range (0.14~0.20). Complex terrain effects cause seasonal variation. Winter (α=0.189) falls within IEC limits while Summer (α=0.25) and Autumn (α=0.241) exceed them. Must be reflected in IEC DLC calculations. Hub height increase yields meaningful wind speed gain.', {
    x: 0.15, y: 0.88, w: 9.5, h: 0.65, fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });
  sl.addImage({ data: charts.c4a, x: 0.1, y: 1.55, w: 7.2, h: 3.8 });
  const alphaData = [
    ['Spring (MAM)', 'α = 0.225', '✗ Exceeds IEC', RED],
    ['Summer (JJA)', 'α = 0.250', '✗ Exceeds IEC', RED],
    ['Autumn (SON)', 'α = 0.241', '✗ Exceeds IEC', RED],
    ['Winter (DJF)', 'α = 0.189', '✓ Within IEC', GREEN],
    ['Annual Mean', 'α = 0.222', '⚠ Slight exceed', YELLOW],
  ];
  alphaData.forEach((d, i) => {
    const cy = 1.65 + i * 0.78;
    sl.addShape(pres.shapes.RECTANGLE, { x: 7.3, y: cy, w: 2.5, h: 0.65, fill: { color: DARKER }, line: { color: d[3], pt: 1 } });
    sl.addText(d[0], { x: 7.38, y: cy+0.04, w: 2.3, h: 0.2, fontSize: 8, color: LGRAY, margin: 0 });
    sl.addText(d[1], { x: 7.38, y: cy+0.22, w: 1.5, h: 0.28, fontSize: 14, bold: true, color: d[3], margin: 0 });
    sl.addText(d[2], { x: 7.38, y: cy+0.44, w: 2.3, h: 0.18, fontSize: 8.5, color: LGRAY, margin: 0 });
  });
}


// ── SLIDE 6B: Wind Shear Rose & by Sector ──────────────────
{
  let sl = addSlide();
  accentBar(sl, 0, H, YELLOW);
  sl.addText('Wind Shear by Direction', { x: 0.15, y: 0.1, w: 7.5, h: 0.42, fontSize: 20, bold: true, color: WHITE, margin: 0 });
  secBadge(sl, '05B', 9.1, 0.1);
  sl.addText('Wind Shear Rose  &  16-Sector Alpha Exponent  ·  2025 Annual', { x: 0.15, y: 0.52, w: 9.2, h: 0.28, fontSize: 10, color: LGRAY, italic: true });
  hRule(sl, 0.15, 0.78, 9.5);
  sl.addText('Sector-by-sector wind shear analysis. SSW (α=0.497) is the highest, indicating terrain influence. Prevailing WNW direction (α=0.184) is within IEC range. ESE/NNW sectors (α≈0.11) show very low shear. 16-sector mean α=0.217. Favourable shear characteristics in the prevailing direction for design load calculations.', {
    x: 0.15, y: 0.88, w: 9.5, h: 0.65, fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });
  sl.addImage({ data: charts.c4b, x: 0.1, y: 1.55, w: 9.75, h: 3.85 });
}

// ── SLIDE 7: TI ───────────────────────────────────────────
{
  let sl = addSlide();
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.38, fill: { color: RED, transparency: 20 }, line: { type: 'none' } });
  sl.addText('⚠  TI WARNING: Exceeds IEC Class A (0.16)  —  TI P90 @15m/s = 0.2300  (+43.8%)  →  IEC Class S High-Turbulence Certified Turbine Required', {
    x: 0.1, y: 0, w: 9.8, h: 0.38, fontSize: 9.5, bold: true, color: WHITE, valign: 'middle', margin: 0
  });
  accentBar(sl, 0.38, H-0.38, RED);
  sl.addText('Turbulence Intensity  (TI)  ·  P90', { x: 0.15, y: 0.42, w: 7, h: 0.38, fontSize: 18, bold: true, color: WHITE, margin: 0 });
  secBadge(sl, '06', 9.45, 0.42);
  sl.addText('TI P90 Curve  &  Wind Speed Bin Distribution  ·  2025 Annual  @ 100m', { x: 0.15, y: 0.8, w: 9.2, h: 0.25, fontSize: 10, color: LGRAY, italic: true });
  hRule(sl, 0.15, 1.03, 9.5);
  sl.addText('TI P90 @ 15 m/s = 0.2300, exceeding IEC Class A limit (0.16) by 43.8%. TI is highest in the low wind speed bins (3~5 m/s) and continues to exceed Class A at high wind speeds. High turbulence confirmed across all wind speed ranges. IEC Class S certified turbine is mandatory.', {
    x: 0.15, y: 1.1, w: 9.5, h: 0.6, fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });
  sl.addImage({ data: charts.c5, x: 0.1, y: 1.72, w: 9.75, h: 3.7 });
}

// ── SLIDE 8: WPD ─────────────────────────────────────────
{
  let sl = addSlide();
  accentBar(sl, 0, H, ORANGE);
  sl.addText('Wind Power Density  (WPD)', { x: 0.15, y: 0.1, w: 6, h: 0.42, fontSize: 20, bold: true, color: WHITE, margin: 0 });
  secBadge(sl, '07', 9.45, 0.1);
  sl.addText('Monthly  &  Height-wise  ·  ρ=1.177 kg/m³  ·  2025 Annual', { x: 0.15, y: 0.52, w: 9.2, h: 0.28, fontSize: 10, color: LGRAY, italic: true });
  hRule(sl, 0.15, 0.78, 9.5);
  sl.addText('Annual WPD 281.8 W/m² (@ 100m) exceeds the development threshold (200 W/m²) by 40.9%, indicating commercial development potential. Winter (Dec-Feb) peaks at ~615 W/m² while Summer (Jun-Aug) drops to ~88 W/m², a seasonal ratio of ~7x. Operational planning must address this seasonal variability.', {
    x: 0.15, y: 0.88, w: 9.5, h: 0.65, fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });
  sl.addImage({ data: charts.c6, x: 0.1, y: 1.55, w: 7.2, h: 3.8 });
  const wpdStats = [
    ['Annual WPD @ 100m', '281.8 W/m²', '≥200 W/m² ✓', ORANGE],
    ['Peak WPD (Winter)', '~615 W/m²', 'Winter maximum', CYAN],
    ['Min WPD (Summer)', '~88 W/m²', 'Summer minimum', LBLUE],
    ['Hub Height Extrap.', 'α=0.222', '+4.5%/20m increase', YELLOW],
  ];
  wpdStats.forEach((d, i) => {
    const cy = 1.65 + i * 1.0;
    sl.addShape(pres.shapes.RECTANGLE, { x: 7.3, y: cy, w: 2.5, h: 0.82, fill: { color: DARKER }, line: { color: d[3], pt: 1 } });
    sl.addShape(pres.shapes.RECTANGLE, { x: 7.3, y: cy, w: 0.05, h: 0.82, fill: { color: d[3] }, line: { type: 'none' } });
    sl.addText(d[0], { x: 7.38, y: cy+0.05, w: 2.3, h: 0.2, fontSize: 7.5, color: LGRAY, margin: 0 });
    sl.addText(d[1], { x: 7.38, y: cy+0.22, w: 2.3, h: 0.35, fontSize: 15, bold: true, color: d[3], margin: 0 });
    sl.addText(d[2], { x: 7.38, y: cy+0.57, w: 2.3, h: 0.2, fontSize: 8.5, color: LGRAY, margin: 0 });
  });
}

// ── SLIDE 9: Extreme Wind ────────────────────────────────
{
  let sl = addSlide();
  accentBar(sl, 0, H, "FF4081");
  sl.addText('Extreme Wind Speed', { x: 0.15, y: 0.1, w: 6, h: 0.42, fontSize: 20, bold: true, color: WHITE, margin: 0 });
  secBadge(sl, '08', 9.45, 0.1);
  sl.addText('Gumbel Distribution  ·  Return Period Curve  ·  2025 Annual', { x: 0.15, y: 0.52, w: 9.2, h: 0.28, fontSize: 10, color: LGRAY, italic: true });
  hRule(sl, 0.15, 0.78, 9.5);
  sl.addText('Gumbel distribution fitted to 12 monthly maximum wind speed samples. V50=42.30 m/s exceeds IEC Class III limit (37.5 m/s) by 12.8%. Statistical reliability is limited with only 12 samples from 1 year of data. Extension of measurement to 3+ years is essential to improve V50 confidence and enable proper IEC wind class determination.', {
    x: 0.15, y: 0.88, w: 9.5, h: 0.65, fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });
  sl.addImage({ data: charts.c7, x: 0.1, y: 1.55, w: 6.9, h: 3.8 });
  sl.addShape(pres.shapes.RECTANGLE, { x: 7.1, y: 1.55, w: 2.7, h: 0.3, fill: { color: "FF4081" }, line: { type: 'none' } });
  ['Return Period','Extreme WS','Class III'].forEach((h, i) => {
    const xs=[7.12,8.15,9.05], ws=[1.0,0.9,0.75];
    sl.addText(h, { x: xs[i], y: 1.55, w: ws[i], h: 0.3, fontSize: 8.5, bold: true, color: WHITE, margin: 0 });
  });
  const retData = [
    ['T = 2yr','25.48','—'],['T = 5yr','30.87','—'],['T = 10yr','34.44','—'],
    ['T = 25yr','38.96','—'],['T = 50yr (V50)','42.30','≤37.5 ✗'],['T = 100yr','45.63','—'],
  ];
  retData.forEach((row, ri) => {
    const ry = 1.85 + ri * 0.34;
    sl.addShape(pres.shapes.RECTANGLE, { x: 7.1, y: ry, w: 2.7, h: 0.32, fill: { color: ri%2===0?'0F2840':'071117' }, line: { type: 'none' } });
    const xs=[7.12,8.15,9.05], ws=[1.0,0.9,0.75];
    const isV50 = ri===4;
    row.forEach((v, ci) => {
      sl.addText(v, { x: xs[ci], y: ry, w: ws[ci], h: 0.32, fontSize: 8.5, color: isV50&&ci>0?YELLOW:WHITE, bold: isV50, valign: 'middle', margin: 0 });
    });
  });
  sl.addShape(pres.shapes.RECTANGLE, { x: 7.1, y: 4.12, w: 2.7, h: 0.65, fill: { color: DARKER }, line: { color: RED, pt: 1 } });
  sl.addText('Assessment: V50=42.30 m/s — Exceeds Class III (37.5 m/s) ✗.\n12 monthly max samples. Measurement extension required.', {
    x: 7.15, y: 4.17, w: 2.6, h: 0.55, fontSize: 8.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.3
  });
}

// ── SLIDE 10: Summary Dashboard ───────────────────────────
{
  let sl = addSlide();
  accentBar(sl, 0, H, CYAN);
  sl.addText('Comprehensive Analysis Dashboard', { x: 0.15, y: 0.1, w: 7, h: 0.42, fontSize: 20, bold: true, color: WHITE, margin: 0 });
  secBadge(sl, '09', 9.45, 0.1);
  sl.addText('7-Item Summary  ·  IEC 61400-1 Ed.4 Benchmarks  ·  2025 Annual', { x: 0.15, y: 0.52, w: 9.2, h: 0.28, fontSize: 10, color: LGRAY, italic: true });
  hRule(sl, 0.15, 0.78, 9.5);
  sl.addText('7-item assessment: WPD (281.8 W/m²), Weibull (C=7.172), prevailing direction (WNW 20.4%), and mean wind speed (6.342 m/s) meet IEC criteria (✓). Wind Shear (α=0.222) slightly exceeds IEC (⚠). TI (P90=0.230) and V50 (42.30 m/s) exceed IEC limits (✗) — these are the key design drivers.', {
    x: 0.15, y: 0.88, w: 9.5, h: 0.65, fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });
  const dashStats = [
    ['Mean Wind Speed','6.342 m/s','@ 100m  ≥6 ✓', CYAN, '✓'],
    ['Weibull C','7.172 m/s','k=2.093  ✓', LBLUE, '✓'],
    ['WPD','281.8 W/m²','≥200 W/m² ✓', ORANGE, '✓'],
    ['TI P90','0.2300','≤0.16 Class A ✗', RED, '✗'],
    ['Wind Shear α','0.222','IEC 0.14-0.20 ⚠', YELLOW, '⚠'],
    ['Extreme V50','42.30 m/s','Class III 37.5 ✗', RED, '✗'],
  ];
  const bw=3.05, bh=1.1, bgap=0.07;
  dashStats.forEach((s, i) => {
    const col=i%3, row2=Math.floor(i/2);
    const bx=0.15+col*(bw+bgap), by=1.65+(i<3?0:bh+bgap);
    sl.addShape(pres.shapes.RECTANGLE, { x: bx, y: by, w: bw, h: bh, fill: { color: DARKER }, line: { color: s[3], pt: 1 } });
    sl.addShape(pres.shapes.RECTANGLE, { x: bx, y: by, w: bw, h: 0.05, fill: { color: s[3] }, line: { type: 'none' } });
    const bc=s[4]==='✓'?GREEN:(s[4]==='✗'?RED:YELLOW);
    sl.addShape(pres.shapes.OVAL, { x: bx+bw-0.36, y: by+0.1, w: 0.3, h: 0.3, fill: { color: bc }, line: { type: 'none' } });
    sl.addText(s[4], { x: bx+bw-0.36, y: by+0.1, w: 0.3, h: 0.3, fontSize: 11, bold: true, color: DARK_BG, align: 'center', valign: 'middle', margin: 0 });
    sl.addText(s[0], { x: bx+0.12, y: by+0.12, w: bw-0.55, h: 0.22, fontSize: 9, color: LGRAY, margin: 0 });
    sl.addText(s[1], { x: bx+0.12, y: by+0.33, w: bw-0.25, h: 0.5, fontSize: 22, bold: true, color: s[3], margin: 0 });
    sl.addText(s[2], { x: bx+0.12, y: by+0.82, w: bw-0.25, h: 0.22, fontSize: 8.5, color: LGRAY, margin: 0 });
  });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.15, y: 5.15, w: 9.5, h: 0.3, fill: { color: DARKER }, line: { type: 'none' } });
  sl.addText('Data Recovery: 100.0%  |  Air Density: 1.177 kg/m³  |  Standard: IEC 61400-1 Ed.4  |  Measurement: 2025.01.01 ~ 2025.12.31', {
    x: 0.15, y: 5.15, w: 9.5, h: 0.3, fontSize: 8.5, color: LGRAY, align: 'center', valign: 'middle'
  });
}

// ── SLIDE 11: Site Suitability Assessment ─────────────────
{
  let sl = addSlide();
  accentBar(sl, 0, H, GREEN);
  sl.addText('Site Suitability Assessment', { x: 0.15, y: 0.1, w: 7, h: 0.42, fontSize: 20, bold: true, color: WHITE, margin: 0 });
  secBadge(sl, '10', 9.45, 0.1);
  sl.addText('IEC 61400-1 Ed.4  ·  2025 Annual Analysis Results', { x: 0.15, y: 0.52, w: 9.2, h: 0.28, fontSize: 10, color: LGRAY, italic: true });
  hRule(sl, 0.15, 0.78, 9.5);
  sl.addText('IEC 61400-1 Ed.4: 4 items pass (✓), 1 conditional (⚠, extended measurement), 2 exceed (✗). Statistical confidence is limited with 1-year data. Annual mean speed and WPD meet development criteria. TI and V50 exceedances require design countermeasures and extended measurement.', {
    x: 0.15, y: 0.88, w: 9.5, h: 0.65, fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });
  const hdrs = ['#', 'Parameter', 'Result', 'IEC Criterion', 'Result', 'Remarks & Recommendations'];
  const hws2 = [0.3, 1.05, 1.65, 1.3, 0.5, 4.4];
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.1, y: 1.62, w: 9.7, h: 0.3, fill: { color: GREEN }, line: { type: 'none' } });
  let hx2=0.1;
  hdrs.forEach((h, i) => {
    sl.addText(h, { x: hx2, y: 1.62, w: hws2[i], h: 0.3, fontSize: 8.5, bold: true, color: DARK_BG, valign: 'middle', margin: 0 });
    hx2 += hws2[i];
  });
  const evalRows = [
    ['01','Mean Wind Speed','6.342 m/s @ 100m','≥6~7 m/s','✓',GREEN,'Meets IEC criterion. 120m with α=0.222 → 6.604 m/s.'],
    ['02','Wind Direction','WNW 24.1% / Energy 20.4%','Single dominant','✓',GREEN,'Westerly energy 40.8%. Optimize array on WNW axis.'],
    ['03','Weibull','k=2.093  C=7.172 m/s','C≥6.0 m/s','✓',GREEN,'C meets IEC criterion. Winter peak, Summer trough.'],
    ['04','TI P90','0.2300 @ 15m/s','≤0.16 (Class A)','✗',RED,'Exceeds Class A by 43.8%. IEC Class S turbine mandatory.'],
    ['05','Wind Shear α','0.222 annual (Spring 0.225)','0.14~0.20','⚠',YELLOW,'Slightly exceeds IEC range. Must apply to DLC.'],
    ['06','WPD','281.8 W/m² @ 100m','≥200 W/m²','✓',GREEN,'Exceeds dev. threshold by 40.9%. Seasonal variability.'],
    ['07','Extreme V50','42.30 m/s (Gumbel)','≤37.5 (III)','✗',RED,'Exceeds Class III. 12 samples only — extend measurement.'],
  ];
  evalRows.forEach((row, ri) => {
    const ry = 1.94 + ri * 0.49;
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.1, y: ry, w: 9.7, h: 0.46, fill: { color: ri%2===0?'0F2840':'071117' }, line: { type: 'none' } });
    const cols = [row[0],row[1],row[2],row[3],row[4],row[6]];
    let cx=0.1;
    cols.forEach((c, ci) => {
      sl.addText(c, { x: cx, y: ry, w: hws2[ci], h: 0.46, fontSize: ci===5?7.5:8.5, color: ci===4?row[5]:(ci===0?LGRAY:WHITE), bold: ci===4, valign: 'middle', margin: 0 });
      cx += hws2[ci];
    });
  });
}

// ── SLIDE 12: Conclusions ─────────────────────────────────
{
  let sl = addSlide();
  accentBar(sl, 0, H, CYAN);
  sl.addText('Conclusions & Recommended Actions', { x: 0.15, y: 0.1, w: 7, h: 0.42, fontSize: 20, bold: true, color: WHITE, margin: 0 });
  secBadge(sl, '11', 9.45, 0.1);
  sl.addText('Based on 1-Year Analysis  ·  Measurement Extension · Turbine Selection · Hub Height Optimization', { x: 0.15, y: 0.52, w: 9.2, h: 0.28, fontSize: 10, color: LGRAY, italic: true });
  hRule(sl, 0.15, 0.78, 9.5);
  sl.addText('2025 analysis shows mean wind speed, WPD, Weibull parameters, and wind direction all meet IEC criteria. TI and V50 exceed limits and Wind Shear slightly exceeds IEC range. With only 1 year of data, statistical confidence for extreme wind analysis is limited. A phased approach from immediate actions to medium-term design enables commercial feasibility.', {
    x: 0.15, y: 0.88, w: 9.5, h: 0.65, fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });
  const actions = [
    ['🔴  Immediate Actions', RED, [
      'Root cause analysis of TI/V50 exceedances — sensor inspection',
      'Min temperature -13.5°C → review LT(-20°C) cold climate package',
      'Plan 2-3 year measurement extension to improve V50 confidence',
    ]],
    ['🟡  Short-term (MCP Analysis)', YELLOW, [
      'MCP analysis using Pohang ASOS (10yr+) + ERA5 reanalysis data',
      'Re-assess site viability after long-term mean wind speed correction',
      'CFD simulation: 3D flow analysis for complex terrain with α=0.222',
    ]],
    ['🔵  Medium-term (Design & Turbine)', LBLUE, [
      'IEC Class S high-turbulence certified turbine mandatory (TI P90=0.230)',
      'Hub height optimisation: 120m→6.604 m/s  /  150m→6.939 m/s',
      'Cold climate package: LT(-20°C) + blade heating system',
    ]],
    ['🟢  Long-term (Final Assessment)', GREEN, [
      'Achieve 3+ years of data → recalculate V50 and confirm IEC wind class',
      'Final AEP calculation: P50/P90 + wake losses + air density correction',
      'IEC certification: DLC (design load calculation) + structural verification',
    ]],
  ];
  const abw=4.7, abh=1.55, agap=0.1;
  actions.forEach((a, i) => {
    const col=i%2, row2=Math.floor(i/2);
    const ax=0.1+col*(abw+agap), ay=1.62+row2*(abh+agap);
    sl.addShape(pres.shapes.RECTANGLE, { x: ax, y: ay, w: abw, h: abh, fill: { color: DARKER }, line: { color: a[1], pt: 1 } });
    sl.addShape(pres.shapes.RECTANGLE, { x: ax, y: ay, w: abw, h: 0.3, fill: { color: a[1], transparency: 30 }, line: { type: 'none' } });
    sl.addText(a[0], { x: ax+0.1, y: ay+0.02, w: abw-0.15, h: 0.28, fontSize: 10, bold: true, color: WHITE, valign: 'middle', margin: 0 });
    a[2].forEach((bullet, bi) => {
      sl.addText([{ text: '▸  ', options: { color: a[1], bold: true } }, { text: bullet, options: { color: WHITE } }], {
        x: ax+0.1, y: ay+0.38+bi*0.37, w: abw-0.2, h: 0.35, fontSize: 8.8, wrap: true
      });
    });
  });
}

pres.writeFile({ fileName: '/home/claude/Gwoniri_WRA_2025_EN.pptx' })
  .then(() => console.log('English PPTX saved!'));
