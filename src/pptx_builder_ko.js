const pptxgen = require("pptxgenjs");
const fs = require("fs");

// Colors
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

// Slide dims: 10" x 5.625"
const W = 10, H = 5.625;

function toB64(path) {
  const buf = fs.readFileSync(path);
  const ext = path.split('.').pop().toLowerCase();
  const mime = ext === 'png' ? 'image/png' : 'image/jpeg';
  return `${mime};base64,${buf.toString('base64')}`;
}

const charts = {
  c1:  toB64('/home/claude/chart1_ko.png'),
  c2:  toB64('/home/claude/chart2_ko.png'),
  c3:  toB64('/home/claude/chart3_ko.png'),
  c4a: toB64('/home/claude/chart4a_ko.png'),
  c4b: toB64('/home/claude/chart4b_ko.png'),
  c5:  toB64('/home/claude/chart5_ko.png'),
  c6:  toB64('/home/claude/chart6_ko.png'),
  c7:  toB64('/home/claude/chart7_ko.png'),
};

function addSlide(pres) {
  let sl = pres.addSlide();
  sl.background = { color: DARK_BG };
  return sl;
}

// Left accent bar
function accentBar(slide, y, h, color) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: y, w: 0.06, h: h,
    fill: { color: color || CYAN }, line: { type: 'none' }
  });
}

// Section number badge
function secBadge(slide, num, x, y) {
  slide.addShape(pres.shapes.OVAL, {
    x: x, y: y, w: 0.4, h: 0.4,
    fill: { color: CYAN }, line: { type: 'none' }
  });
  slide.addText(String(num), {
    x: x, y: y, w: 0.4, h: 0.4,
    fontSize: 13, bold: true, color: DARK_BG, align: 'center', valign: 'middle', margin: 0
  });
}

// Horizontal rule
function hRule(slide, x, y, w) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: x, y: y, w: w, h: 0.02,
    fill: { color: CYAN, transparency: 40 }, line: { type: 'none' }
  });
}

// Stat box
function statBox(slide, x, y, w, h, label, value, unit, color, icon) {
  const c = color || CYAN;
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: DARKER }, line: { color: c, pt: 1 }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w: 0.05, h,
    fill: { color: c }, line: { type: 'none' }
  });
  slide.addText(label, {
    x: x+0.1, y: y+0.05, w: w-0.15, h: 0.22,
    fontSize: 8.5, color: LGRAY, bold: false, margin: 0
  });
  slide.addText(value, {
    x: x+0.1, y: y+0.22, w: w-0.15, h: 0.45,
    fontSize: 20, bold: true, color: c, margin: 0
  });
  if (unit) {
    slide.addText(unit, {
      x: x+0.1, y: y+0.62, w: w-0.15, h: 0.18,
      fontSize: 8, color: LGRAY, margin: 0
    });
  }
}

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.title = '권이리 풍황계측 분석보고서 2025';

// ═══════════════════════════════════════════
// SLIDE 1: 표지
// ═══════════════════════════════════════════
{
  let sl = addSlide(pres);
  // Left panel
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 3.8, h: H,
    fill: { color: DARKER }, line: { type: 'none' }
  });
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 3.8, y: 0, w: 0.04, h: H,
    fill: { color: CYAN }, line: { type: 'none' }
  });

  // Title
  sl.addText('풍황계측\n분석보고서', {
    x: 0.3, y: 1.0, w: 3.2, h: 1.6,
    fontSize: 32, bold: true, color: WHITE, align: 'left',
    lineSpacingMultiple: 1.2
  });
  sl.addText('2025년  (2025.01.01 – 2025.12.31)', {
    x: 0.3, y: 2.7, w: 3.2, h: 0.35,
    fontSize: 11, color: CYAN, bold: true, align: 'left'
  });
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 3.05, w: 2.5, h: 0.03,
    fill: { color: CYAN, transparency: 30 }, line: { type: 'none' }
  });
  sl.addText('부지: 경북 경주시 문무대왕면 권이리 산324-3', {
    x: 0.3, y: 3.15, w: 3.2, h: 0.25, fontSize: 9, color: LGRAY
  });
  sl.addText('계측탑: 100m 격자형 철탑  NRG SymphoniePRO  S/N: 820616523', {
    x: 0.3, y: 3.38, w: 3.2, h: 0.25, fontSize: 8.5, color: LGRAY
  });
  sl.addText('발주처: (주)케이에스파워', {
    x: 0.3, y: 3.60, w: 3.2, h: 0.22, fontSize: 9, color: LGRAY
  });
  sl.addText('분석기준: IEC 61400-1 Ed.4  /  KS C IEC 61400-12', {
    x: 0.3, y: 3.82, w: 3.2, h: 0.22, fontSize: 8.5, color: LGRAY
  });

  // Right panel: key stats
  const RX = 4.0;
  const stats = [
    ['연평균 풍속', '6.342 m/s', '@ 100m', CYAN, '✓'],
    ['연간 WPD', '281.8 W/m²', '@ 100m', ORANGE, '✓'],
    ['Weibull k/C', 'k=2.093  C=7.172', '@ 100m', LBLUE, '✓'],
    ['주풍향', 'WNW  24.1%', '에너지 20.4%', TEAL, '✓'],
    ['Wind Shear α', '0.222', 'IEC 0.14~0.20 ⚠', YELLOW, '⚠'],
    ['TI P90 @15m/s', '0.2300', 'IEC Class A: 0.16 ✗', RED, '✗'],
    ['극한풍속 V50', '42.30 m/s', 'IEC Class III: 37.5 ✗', RED, '✗'],
    ['데이터 회수율', '100.0%', '2025.01~12', GREEN, '✓'],
  ];
  const boxW = 2.82, boxH = 0.88;
  const gap = 0.06;
  stats.forEach((s, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const bx = RX + col * (boxW + gap);
    const by = 0.18 + row * (boxH + gap);
    statBox(sl, bx, by, boxW, boxH, s[0], s[1], s[2], s[3]);
    // badge
    const badgeColor = s[4]==='✓' ? GREEN : (s[4]==='✗' ? RED : YELLOW);
    sl.addShape(pres.shapes.OVAL, {
      x: bx+boxW-0.32, y: by+0.04, w: 0.25, h: 0.25,
      fill: { color: badgeColor }, line: { type: 'none' }
    });
    sl.addText(s[4], {
      x: bx+boxW-0.32, y: by+0.04, w: 0.25, h: 0.25,
      fontSize: 10, bold: true, color: DARK_BG, align: 'center', valign: 'middle', margin: 0
    });
  });
  // 3개년 통합 태그
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 4.9, w: 3.2, h: 0.4,
    fill: { color: CYAN, transparency: 75 }, line: { type: 'none' }
  });
  sl.addText('1년 계측 분석  (2025.01.01 – 2025.12.31)', {
    x: 0.3, y: 4.9, w: 3.2, h: 0.4,
    fontSize: 9.5, color: CYAN, align: 'center', valign: 'middle', bold: true
  });
}

// ═══════════════════════════════════════════
// SLIDE 2: 계측 시스템 개요
// ═══════════════════════════════════════════
{
  let sl = addSlide(pres);
  accentBar(sl, 0, H, CYAN);

  sl.addText('계측 시스템 개요', {
    x: 0.15, y: 0.1, w: 6, h: 0.42,
    fontSize: 20, bold: true, color: WHITE, margin: 0
  });
  secBadge(sl, '01', 9.45, 0.1);
  sl.addText('NRG SymphoniePRO  ·  100m 격자형 철탑  ·  측정 채널 구성', {
    x: 0.15, y: 0.52, w: 9.2, h: 0.28,
    fontSize: 10, color: LGRAY, italic: true
  });
  hRule(sl, 0.15, 0.78, 9.5);

  // Body text
  sl.addText('NRG SymphoniePRO 로거(S/N: 820616523) 탑재 100m 격자형 철탑에서 2025년 계측 완료. 데이터 회수율 100.0%로 전 기간 결측 없음. 공기밀도 1.177 kg/m³(표준 대비 -4.1%), 연평균 기온 13.6°C, 기압 969.0 hPa. 높은 회수율로 통계적 신뢰성 우수하며 직접 분석 활용 가능.', {
    x: 0.15, y: 0.88, w: 5.6, h: 0.85,
    fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });

  // Site info blocks
  const infoData = [
    ['계측 기간', '2025.01.01  ~  2025.12.31  (1개년)'],
    ['설치 위치', '35.867740°N  129.414420°E  /  해발 390m'],
    ['타워 제원', '100m 격자형 철탑  /  위성위치 확인'],
    ['데이터 로거', 'NRG SymphoniePRO  (S/N: 820616523)'],
    ['샘플링 간격', '10분 평균값  (144 records/day)'],
    ['총 레코드수', '52,559건  (2025.01 ~ 2025.12)'],
    ['회수율', '100.0%  (52,559 / 52,560)'],
    ['공기밀도', '1.177 kg/m³  (기온 13.6°C, 기압 969.0 hPa)'],
  ];
  infoData.forEach((row, i) => {
    const iy = 1.78 + i * 0.33;
    sl.addShape(pres.shapes.RECTANGLE, {
      x: 0.15, y: iy, w: 1.5, h: 0.28,
      fill: { color: CYAN, transparency: 75 }, line: { type: 'none' }
    });
    sl.addText(row[0], { x: 0.18, y: iy, w: 1.44, h: 0.28, fontSize: 8.5, bold: true, color: CYAN, valign: 'middle', margin: 0 });
    sl.addText(row[1], { x: 1.7, y: iy, w: 4.2, h: 0.28, fontSize: 8.5, color: WHITE, valign: 'middle', margin: 0 });
  });

  // Sensor table
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 5.8, y: 0.88, w: 3.9, h: 0.28,
    fill: { color: CYAN }, line: { type: 'none' }
  });
  const hdr = ['Ch','센서 종류','높이','방향','측정 항목'];
  const hw = [0.35, 1.3, 0.55, 0.65, 1.0];
  let hx = 5.83;
  hdr.forEach((h, i) => {
    sl.addText(h, { x: hx, y: 0.88, w: hw[i], h: 0.28, fontSize: 8, bold: true, color: DARK_BG, valign: 'middle', margin: 0 });
    hx += hw[i];
  });

  const tableData = [
    ['A1', '풍속계 (Thies FCA)', '100m', '35° (NE)', '풍속 Avg/SD/Gust'],
    ['A2', '풍속계 (Thies FCA)', '100m', '215° (SW)', '풍속 Avg/SD/Gust'],
    ['A3', '풍속계 (Thies FCA)', '80m', '35° (NE)', '풍속 Avg/SD/Gust'],
    ['A4', '풍속계 (Thies FCA)', '80m', '215° (SW)', '풍속 Avg/SD/Gust'],
    ['A5', '풍속계 (Thies FCA)', '60m', '35° (NE)', '풍속 Avg/SD/Gust'],
    ['A6', '풍속계 (Thies FCA)', '60m', '215° (SW)', '풍속 Avg/SD/Gust'],
    ['V1', '풍향계 (Thies)', '95m', '35° (NE)', '풍향 Avg'],
    ['V2', '풍향계 (Thies)', '75m', '35° (NE)', '풍향 Avg'],
    ['V3', '풍향계 (Thies)', '55m', '35° (NE)', '풍향 Avg'],
    ['T1', '온도계 (NRG T60)', '5m', '—', '기온 Avg'],
    ['B1', '기압계 (NRG BP65)', '2m', '—', '기압 Avg'],
    ['H1', '습도계 (NRG RH5XC)', '2m', '—', '상대습도 Avg'],
  ];
  tableData.forEach((row, ri) => {
    const ry = 1.18 + ri * 0.34;
    const bg = ri % 2 === 0 ? '0F2840' : '071117';
    sl.addShape(pres.shapes.RECTANGLE, {
      x: 5.8, y: ry, w: 3.9, h: 0.32,
      fill: { color: bg }, line: { type: 'none' }
    });
    let rx = 5.83;
    row.forEach((cell, ci) => {
      sl.addText(cell, {
        x: rx, y: ry, w: hw[ci], h: 0.32,
        fontSize: 7.8, color: ci === 0 ? CYAN : WHITE, valign: 'middle', margin: 0
      });
      rx += hw[ci];
    });
  });
}

// ═══════════════════════════════════════════
// SLIDE 3: 평균 풍속 분석
// ═══════════════════════════════════════════
{
  let sl = addSlide(pres);
  accentBar(sl, 0, H, LBLUE);

  sl.addText('평균 풍속 분석', {
    x: 0.15, y: 0.1, w: 6, h: 0.42, fontSize: 20, bold: true, color: WHITE, margin: 0
  });
  secBadge(sl, '02', 9.45, 0.1);
  sl.addText('월별 변동  &  수직 프로파일  ·  2025 연간  @ 100m', {
    x: 0.15, y: 0.52, w: 9.2, h: 0.28, fontSize: 10, color: LGRAY, italic: true
  });
  hRule(sl, 0.15, 0.78, 9.5);

  sl.addText('연간 평균 풍속 6.342 m/s(@ 100m)는 IEC 개발 권장 기준(6~7 m/s)을 충족. Wind Shear α=0.222로 허브 높이 120m에서 약 6.604 m/s, 150m에서 약 6.939 m/s 추정. 1~2월 겨울 풍속이 최대(약 8.8 m/s), 8월 여름이 최소(약 4.3 m/s)로 계절 진폭이 크며 계절 에너지 보정이 필요함.', {
    x: 0.15, y: 0.88, w: 9.5, h: 0.7,
    fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });

  sl.addImage({ data: charts.c1, x: 0.1, y: 1.62, w: 7.0, h: 3.7 });

  // Stat callouts
  const callouts = [
    ['연평균', '6.342 m/s', LBLUE],
    ['α=0.222', '120m: 6.604', YELLOW],
  ];
  callouts.forEach((c, i) => {
    const cx = 7.2, cy = 1.7 + i * 1.3;
    sl.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy, w: 2.55, h: 0.95,
      fill: { color: DARKER }, line: { color: c[2], pt: 1 }
    });
    sl.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy, w: 0.05, h: 0.95,
      fill: { color: c[2] }, line: { type: 'none' }
    });
    sl.addText(c[0], { x: cx+0.1, y: cy+0.06, w: 2.3, h: 0.25, fontSize: 9, color: LGRAY, margin: 0 });
    sl.addText(c[1], { x: cx+0.1, y: cy+0.3, w: 2.3, h: 0.45, fontSize: 17, bold: true, color: c[2], margin: 0 });
  });
  // monthly min/max note
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 7.2, y: 4.3, w: 2.55, h: 1.0,
    fill: { color: DARKER }, line: { color: GRAY, pt: 1 }
  });
  sl.addText([
    { text: '최강풍월  2월 8.832 m/s\n', options: { color: CYAN, fontSize: 9.5, bold: true } },
    { text: '최약풍월  8월 4.285 m/s\n', options: { color: ORANGE, fontSize: 9.5, bold: true } },
    { text: '계절 진폭  4.547 m/s', options: { color: LGRAY, fontSize: 8.5 } }
  ], { x: 7.25, y: 4.35, w: 2.45, h: 0.9, margin: 0 });
}

// ═══════════════════════════════════════════
// SLIDE 4: 풍향 분포
// ═══════════════════════════════════════════
{
  let sl = addSlide(pres);
  accentBar(sl, 0, H, TEAL);

  sl.addText('풍향 분포  (Wind Rose)', {
    x: 0.15, y: 0.1, w: 6, h: 0.42, fontSize: 20, bold: true, color: WHITE, margin: 0
  });
  secBadge(sl, '03', 9.45, 0.1);
  sl.addText('연간 빈도  &  에너지 풍배도  @  95m  ·  2025 연간', {
    x: 0.15, y: 0.52, w: 9.2, h: 0.28, fontSize: 10, color: LGRAY, italic: true
  });
  hRule(sl, 0.15, 0.78, 9.5);

  sl.addText('주풍향은 WNW(서북서)로 연간 빈도 24.1%, 에너지 집중도 20.4%. WNW+NW+W 편서계열이 전체 에너지의 약 40.8%를 담당. 에너지 풍배도가 빈도 풍배도보다 WNW 편중이 두드러지며, 배열 방향 최적화 시 WNW 방향을 기준축으로 설정 권장.', {
    x: 0.15, y: 0.88, w: 9.5, h: 0.65,
    fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });

  sl.addImage({ data: charts.c2, x: 0.1, y: 1.55, w: 7.2, h: 3.8 });

  // stats
  const wdata = [
    ['주풍향', 'WNW (서북서)', '빈도 24.1%', TEAL],
    ['에너지 집중도', 'WNW 20.4%', '', ORANGE],
    ['편서계열 에너지', 'WNW+NW+W', '40.8%', LBLUE],
    ['배열 권장', 'WNW 기준 최적화', '', YELLOW],
  ];
  wdata.forEach((d, i) => {
    const cy = 1.65 + i * 1.0;
    sl.addShape(pres.shapes.RECTANGLE, {
      x: 7.3, y: cy, w: 2.5, h: 0.82,
      fill: { color: DARKER }, line: { color: d[3], pt: 1 }
    });
    sl.addShape(pres.shapes.RECTANGLE, {
      x: 7.3, y: cy, w: 0.05, h: 0.82,
      fill: { color: d[3] }, line: { type: 'none' }
    });
    sl.addText(d[0], { x: 7.38, y: cy+0.04, w: 2.3, h: 0.2, fontSize: 8, color: LGRAY, margin: 0 });
    sl.addText(d[1], { x: 7.38, y: cy+0.24, w: 2.3, h: 0.32, fontSize: 13, bold: true, color: d[3], margin: 0 });
    if (d[2]) sl.addText(d[2], { x: 7.38, y: cy+0.56, w: 2.3, h: 0.2, fontSize: 8.5, color: LGRAY, margin: 0 });
  });
}

// ═══════════════════════════════════════════
// SLIDE 5: 와이블 분포
// ═══════════════════════════════════════════
{
  let sl = addSlide(pres);
  accentBar(sl, 0, H, LBLUE);

  sl.addText('와이블 분포  (Weibull Distribution)', {
    x: 0.15, y: 0.1, w: 7, h: 0.42, fontSize: 20, bold: true, color: WHITE, margin: 0
  });
  secBadge(sl, '04', 9.45, 0.1);
  sl.addText('연간·계절별 PDF  ·  k/C 파라미터  @  100m  ·  2025 연간', {
    x: 0.15, y: 0.52, w: 9.2, h: 0.28, fontSize: 10, color: LGRAY, italic: true
  });
  hRule(sl, 0.15, 0.78, 9.5);

  sl.addText('형상계수 k=2.093은 Rayleigh 분포에 근접한 중간 분산 특성. 척도계수 C=7.172 m/s는 IEC 개발 기준(C≥6.0)을 충족. 계절별로 겨울 C=9.137(최고), 여름 C=5.558(최저)로 계절 편차 3.6 m/s에 달하며, 겨울 허브 높이 120m 외삽 시 C≈9.6 m/s 이상 기대.', {
    x: 0.15, y: 0.88, w: 9.5, h: 0.65,
    fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });

  sl.addImage({ data: charts.c3, x: 0.1, y: 1.55, w: 6.9, h: 3.8 });

  // Weibull table
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 7.1, y: 1.55, w: 2.7, h: 0.32,
    fill: { color: LBLUE }, line: { type: 'none' }
  });
  sl.addText('구분', { x: 7.12, y: 1.55, w: 0.7, h: 0.32, fontSize: 8.5, bold: true, color: DARK_BG, margin: 0 });
  sl.addText('k', { x: 7.82, y: 1.55, w: 0.55, h: 0.32, fontSize: 8.5, bold: true, color: DARK_BG, margin: 0 });
  sl.addText('C (m/s)', { x: 8.37, y: 1.55, w: 0.7, h: 0.32, fontSize: 8.5, bold: true, color: DARK_BG, margin: 0 });
  sl.addText('Vave', { x: 9.07, y: 1.55, w: 0.7, h: 0.32, fontSize: 8.5, bold: true, color: DARK_BG, margin: 0 });

  const wrows = [
    ['연간', '2.093', '7.172', '6.353', CYAN],
    ['봄 (MAM)', '2.226', '7.428', '6.579', GREEN],
    ['여름 (JJA)', '2.175', '5.558', '4.923', ORANGE],
    ['가을 (SON)', '2.341', '6.572', '5.823', "FF4081"],
    ['겨울 (DJF)', '2.361', '9.137', '8.097', YELLOW],
  ];
  wrows.forEach((row, ri) => {
    const ry = 1.87 + ri * 0.37;
    const bg = ri % 2 === 0 ? '0F2840' : '071117';
    sl.addShape(pres.shapes.RECTANGLE, {
      x: 7.1, y: ry, w: 2.7, h: 0.34,
      fill: { color: bg }, line: { type: 'none' }
    });
    sl.addText(row[0], { x: 7.12, y: ry, w: 0.7, h: 0.34, fontSize: 8.5, color: row[4], valign: 'middle', bold: ri===0, margin: 0 });
    sl.addText(row[1], { x: 7.82, y: ry, w: 0.55, h: 0.34, fontSize: 8.5, color: WHITE, valign: 'middle', margin: 0 });
    sl.addText(row[2], { x: 8.37, y: ry, w: 0.7, h: 0.34, fontSize: 8.5, color: WHITE, valign: 'middle', margin: 0 });
    sl.addText(row[3], { x: 9.07, y: ry, w: 0.7, h: 0.34, fontSize: 8.5, color: WHITE, valign: 'middle', margin: 0 });
  });
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 7.1, y: 4.0, w: 2.7, h: 0.65,
    fill: { color: DARKER }, line: { color: LBLUE, pt: 1 }
  });
  sl.addText('해석  겨울 C=9.137 → 120m 외삽 시 C≈9.6 m/s.\n계절 에너지 불균형 고려 요.', {
    x: 7.15, y: 4.05, w: 2.6, h: 0.55, fontSize: 8.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.3
  });
}

// ═══════════════════════════════════════════
// SLIDE 6: Wind Shear
// ═══════════════════════════════════════════
{
  let sl = addSlide(pres);
  accentBar(sl, 0, H, YELLOW);

  sl.addText('바람 전단  (Wind Shear)', {
    x: 0.15, y: 0.1, w: 6, h: 0.42, fontSize: 20, bold: true, color: WHITE, margin: 0
  });
  secBadge(sl, '05', 9.45, 0.1);
  sl.addText('멱지수 α  ·  수직 프로파일  &  계절별  ·  2025 연간', {
    x: 0.15, y: 0.52, w: 9.2, h: 0.28, fontSize: 10, color: LGRAY, italic: true
  });
  hRule(sl, 0.15, 0.78, 9.5);

  sl.addText('연간 Wind Shear 멱지수 α=0.222로 IEC 기준(0.14~0.20) 소폭 초과. 복잡 산지 지형 영향으로 계절별 편차 존재. 겨울(α=0.189)은 IEC 기준 이내이나 여름(α=0.25)과 가을(α=0.241)이 기준 초과. 허브 높이 상향 시 풍속 증가 효과 있으며 IEC 설계 하중 계산(DLC) 시 반영 요.', {
    x: 0.15, y: 0.88, w: 9.5, h: 0.65,
    fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });

  sl.addImage({ data: charts.c4a, x: 0.1, y: 1.55, w: 7.2, h: 3.8 });

  const alphaData = [
    ['봄 (MAM)', 'α = 0.225', '✗ IEC 초과', RED],
    ['여름 (JJA)', 'α = 0.250', '✗ IEC 초과', RED],
    ['가을 (SON)', 'α = 0.241', '✗ IEC 초과', RED],
    ['겨울 (DJF)', 'α = 0.189', '✓ IEC 이내', GREEN],
    ['연간 평균', 'α = 0.222', '⚠ 소폭 초과', YELLOW],
  ];
  alphaData.forEach((d, i) => {
    const cy = 1.65 + i * 0.78;
    sl.addShape(pres.shapes.RECTANGLE, {
      x: 7.3, y: cy, w: 2.5, h: 0.65,
      fill: { color: DARKER }, line: { color: d[3], pt: 1 }
    });
    sl.addText(d[0], { x: 7.38, y: cy+0.04, w: 2.3, h: 0.2, fontSize: 8, color: LGRAY, margin: 0 });
    sl.addText(d[1], { x: 7.38, y: cy+0.22, w: 1.5, h: 0.28, fontSize: 14, bold: true, color: d[3], margin: 0 });
    sl.addText(d[2], { x: 7.38, y: cy+0.44, w: 2.3, h: 0.18, fontSize: 8.5, color: LGRAY, margin: 0 });
  });
}

// ═══════════════════════════════════════════

// ═══════════════════════════════════════════
// SLIDE 6B: Wind Shear Rose & by Sector
// ═══════════════════════════════════════════
{
  let sl = addSlide(pres);
  accentBar(sl, 0, H, YELLOW);

  sl.addText('풍향별 바람 전단  (Wind Shear by Direction)', {
    x: 0.15, y: 0.1, w: 7.5, h: 0.42, fontSize: 20, bold: true, color: WHITE, margin: 0
  });
  secBadge(sl, '05B', 9.1, 0.1);
  sl.addText('Wind Shear Rose  &  16방위 섹터별 α 멱지수  ·  2025 연간', {
    x: 0.15, y: 0.52, w: 9.2, h: 0.28, fontSize: 10, color: LGRAY, italic: true
  });
  hRule(sl, 0.15, 0.78, 9.5);

  sl.addText('풍향 섹터별 Wind Shear 멱지수 분석. SSW 방향(α=0.497)이 최대로 지형 영향 집중. 주풍향 WNW(α=0.184)는 IEC 기준 이내. ESE·NNW(α≈0.11) 구간은 매우 낮은 전단. 16방위 평균 α=0.217. 주풍 방향의 전단 특성이 설계 하중 산정에 유리함.', {
    x: 0.15, y: 0.88, w: 9.5, h: 0.65,
    fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });

  sl.addImage({ data: charts.c4b, x: 0.1, y: 1.55, w: 9.75, h: 3.85 });
}

// SLIDE 7: 난류 강도 TI
// ═══════════════════════════════════════════
{
  let sl = addSlide(pres);

  // Warning banner
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.38,
    fill: { color: RED, transparency: 20 }, line: { type: 'none' }
  });
  sl.addText('⚠  TI 경고: IEC Class A 기준(0.16) 초과  —  TI P90 @15m/s = 0.2300  (+43.8%)  →  IEC Class S 고난류 인증 기종 필수', {
    x: 0.1, y: 0, w: 9.8, h: 0.38,
    fontSize: 9.5, bold: true, color: WHITE, valign: 'middle', margin: 0
  });
  accentBar(sl, 0.38, H-0.38, RED);

  sl.addText('난류 강도  (Turbulence Intensity)  ·  P90', {
    x: 0.15, y: 0.42, w: 7, h: 0.38, fontSize: 18, bold: true, color: WHITE, margin: 0
  });
  secBadge(sl, '06', 9.45, 0.42);
  sl.addText('TI P90 곡선  &  풍속 구간별 분포  ·  2025 연간  @ 100m', {
    x: 0.15, y: 0.8, w: 9.2, h: 0.25, fontSize: 10, color: LGRAY, italic: true
  });
  hRule(sl, 0.15, 1.03, 9.5);

  sl.addText('TI P90 @ 15 m/s = 0.2300으로 IEC Class A 기준(0.16) 대비 43.8% 초과. 저풍속 구간(3~5 m/s)에서 TI가 가장 높고, 고풍속 구간에서도 Class A 기준을 지속 초과. 전 풍속 범위에 걸쳐 고난류 특성 확인. IEC Class S 고난류 인증 기종 선정 필수.', {
    x: 0.15, y: 1.1, w: 9.5, h: 0.6,
    fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });

  sl.addImage({ data: charts.c5, x: 0.1, y: 1.72, w: 9.75, h: 3.7 });
}

// ═══════════════════════════════════════════
// SLIDE 8: WPD
// ═══════════════════════════════════════════
{
  let sl = addSlide(pres);
  accentBar(sl, 0, H, ORANGE);

  sl.addText('풍력 에너지 밀도  (WPD)', {
    x: 0.15, y: 0.1, w: 6, h: 0.42, fontSize: 20, bold: true, color: WHITE, margin: 0
  });
  secBadge(sl, '07', 9.45, 0.1);
  sl.addText('월별  &  고도별  ·  ρ=1.177 kg/m³  ·  2025 연간', {
    x: 0.15, y: 0.52, w: 9.2, h: 0.28, fontSize: 10, color: LGRAY, italic: true
  });
  hRule(sl, 0.15, 0.78, 9.5);

  sl.addText('연간 WPD 281.8 W/m²(@ 100m)는 개발 기준(200 W/m²)을 40.9% 상회하여 상업 개발 가능 수준. 겨울(12~2월) 약 614 W/m², 여름(6~8월) 약 88 W/m²로 계절 변동폭이 약 7배에 달해 계절별 운영 계획 필요. α=0.222 적용 시 허브 높이 120m 설정 가능.', {
    x: 0.15, y: 0.88, w: 9.5, h: 0.65,
    fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });

  sl.addImage({ data: charts.c6, x: 0.1, y: 1.55, w: 7.2, h: 3.8 });

  const wpdStats = [
    ['연간 WPD @ 100m', '281.8 W/m²', '≥200 W/m² ✓', ORANGE],
    ['최대 WPD (겨울)', '~615 W/m²', '겨울 최고', CYAN],
    ['최소 WPD (여름)', '~88 W/m²', '여름 최저', LBLUE],
    ['허브 높이 외삽', 'α=0.222', '+4.5%/20m↑', YELLOW],
  ];
  wpdStats.forEach((d, i) => {
    const cy = 1.65 + i * 1.0;
    sl.addShape(pres.shapes.RECTANGLE, {
      x: 7.3, y: cy, w: 2.5, h: 0.82,
      fill: { color: DARKER }, line: { color: d[3], pt: 1 }
    });
    sl.addShape(pres.shapes.RECTANGLE, {
      x: 7.3, y: cy, w: 0.05, h: 0.82,
      fill: { color: d[3] }, line: { type: 'none' }
    });
    sl.addText(d[0], { x: 7.38, y: cy+0.05, w: 2.3, h: 0.2, fontSize: 7.5, color: LGRAY, margin: 0 });
    sl.addText(d[1], { x: 7.38, y: cy+0.22, w: 2.3, h: 0.35, fontSize: 15, bold: true, color: d[3], margin: 0 });
    sl.addText(d[2], { x: 7.38, y: cy+0.57, w: 2.3, h: 0.2, fontSize: 8.5, color: LGRAY, margin: 0 });
  });
}

// ═══════════════════════════════════════════
// SLIDE 9: 극한 풍속
// ═══════════════════════════════════════════
{
  let sl = addSlide(pres);
  accentBar(sl, 0, H, "FF4081");

  sl.addText('극한 풍속  (Extreme Wind Speed)', {
    x: 0.15, y: 0.1, w: 6, h: 0.42, fontSize: 20, bold: true, color: WHITE, margin: 0
  });
  secBadge(sl, '08', 9.45, 0.1);
  sl.addText('Gumbel 분포  ·  재현기간 곡선  ·  2025 연간', {
    x: 0.15, y: 0.52, w: 9.2, h: 0.28, fontSize: 10, color: LGRAY, italic: true
  });
  hRule(sl, 0.15, 0.78, 9.5);

  sl.addText('12개월 최대 풍속(총 12개 샘플)에 Gumbel 분포 적합. V50=42.30 m/s로 IEC Class III 기준(37.5 m/s) 대비 12.8% 초과하여 Class III 이내 미달. 1년 계측 데이터로 V50 신뢰도 제한적이며, 계측 기간 연장(3년 이상)을 통한 통계 신뢰도 향상이 필수적임.', {
    x: 0.15, y: 0.88, w: 9.5, h: 0.65,
    fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });

  sl.addImage({ data: charts.c7, x: 0.1, y: 1.55, w: 6.9, h: 3.8 });

  // return period table
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 7.1, y: 1.55, w: 2.7, h: 0.3,
    fill: { color: "FF4081" }, line: { type: 'none' }
  });
  sl.addText('재현기간', { x: 7.12, y: 1.55, w: 1.0, h: 0.3, fontSize: 8.5, bold: true, color: WHITE, margin: 0 });
  sl.addText('극한풍속', { x: 8.12, y: 1.55, w: 0.9, h: 0.3, fontSize: 8.5, bold: true, color: WHITE, margin: 0 });
  sl.addText('Class III', { x: 9.02, y: 1.55, w: 0.75, h: 0.3, fontSize: 8.5, bold: true, color: WHITE, margin: 0 });

  const retData = [
    ['T = 2년', '25.48', '—'],
    ['T = 5년', '30.87', '—'],
    ['T = 10년', '34.44', '—'],
    ['T = 25년', '38.96', '—'],
    ['T = 50년 (V50)', '42.30', '≤37.5 ✗'],
    ['T = 100년', '45.63', '—'],
  ];
  retData.forEach((row, ri) => {
    const ry = 1.85 + ri * 0.34;
    const bg = ri % 2 === 0 ? '0F2840' : '071117';
    sl.addShape(pres.shapes.RECTANGLE, { x: 7.1, y: ry, w: 2.7, h: 0.32, fill: { color: bg }, line: { type: 'none' } });
    const isV50 = ri === 4;
    sl.addText(row[0], { x: 7.12, y: ry, w: 1.0, h: 0.32, fontSize: 8.5, color: isV50 ? YELLOW : WHITE, bold: isV50, valign: 'middle', margin: 0 });
    sl.addText(row[1], { x: 8.12, y: ry, w: 0.9, h: 0.32, fontSize: 8.5, color: isV50 ? YELLOW : WHITE, bold: isV50, valign: 'middle', margin: 0 });
    sl.addText(row[2], { x: 9.02, y: ry, w: 0.75, h: 0.32, fontSize: 8.5, color: isV50 ? RED : WHITE, bold: isV50, valign: 'middle', margin: 0 });
  });
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 7.1, y: 4.12, w: 2.7, h: 0.65,
    fill: { color: DARKER }, line: { color: RED, pt: 1 }
  });
  sl.addText('판정  V50=42.30 m/s — Class III(37.5) 초과 ✗.\n1년 12개 Gumbel 적합. 계측 연장 필수.', {
    x: 7.15, y: 4.17, w: 2.6, h: 0.55, fontSize: 8.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.3
  });
}

// ═══════════════════════════════════════════
// SLIDE 10: 종합 분석 대시보드
// ═══════════════════════════════════════════
{
  let sl = addSlide(pres);
  accentBar(sl, 0, H, CYAN);

  sl.addText('종합 분석 대시보드', {
    x: 0.15, y: 0.1, w: 6, h: 0.42, fontSize: 20, bold: true, color: WHITE, margin: 0
  });
  secBadge(sl, '09', 9.45, 0.1);
  sl.addText('7대 항목 요약  ·  IEC 61400-1 Ed.4 기준선  ·  2025 연간', {
    x: 0.15, y: 0.52, w: 9.2, h: 0.28, fontSize: 10, color: LGRAY, italic: true
  });
  hRule(sl, 0.15, 0.78, 9.5);

  sl.addText('7대 항목 종합: WPD(281.8 W/m²)·Weibull(C=7.172)·주풍향(WNW 20.4%) 3개 항목 IEC 기준 충족(✓). 평균 풍속(6.342 m/s)은 기준 충족. Wind Shear(α=0.222) 소폭 초과(⚠). TI(P90=0.230)와 V50(42.30 m/s) 2개 항목 기준 초과(✗)로 핵심 설계 변수임.', {
    x: 0.15, y: 0.88, w: 9.5, h: 0.65,
    fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });

  // 6 big stat boxes (2x3)
  const dashStats = [
    ['연평균 풍속', '6.342 m/s', '@ 100m  ≥6 ✓', CYAN, '✓'],
    ['Weibull C', '7.172 m/s', 'k=2.093  ✓', LBLUE, '✓'],
    ['WPD', '281.8 W/m²', '≥200 W/m² ✓', ORANGE, '✓'],
    ['TI P90', '0.2300', '≤0.16 Class A ✗', RED, '✗'],
    ['Wind Shear α', '0.222', 'IEC 0.14-0.20 ⚠', YELLOW, '⚠'],
    ['극한풍속 V50', '42.30 m/s', 'Class III 37.5 ✗', RED, '✗'],
  ];
  const bw = 3.05, bh = 1.1, bgap = 0.07;
  dashStats.forEach((s, i) => {
    const col = i % 3, row = Math.floor(i / 2);
    const bx = 0.15 + col * (bw + bgap);
    const by = 1.65 + (i < 3 ? 0 : bh + bgap);
    const c = s[3];
    sl.addShape(pres.shapes.RECTANGLE, {
      x: bx, y: by, w: bw, h: bh,
      fill: { color: DARKER }, line: { color: c, pt: 1 }
    });
    sl.addShape(pres.shapes.RECTANGLE, {
      x: bx, y: by, w: bw, h: 0.05,
      fill: { color: c }, line: { type: 'none' }
    });
    const bc = s[4]==='✓' ? GREEN : (s[4]==='✗' ? RED : YELLOW);
    sl.addShape(pres.shapes.OVAL, {
      x: bx+bw-0.36, y: by+0.1, w: 0.3, h: 0.3,
      fill: { color: bc }, line: { type: 'none' }
    });
    sl.addText(s[4], {
      x: bx+bw-0.36, y: by+0.1, w: 0.3, h: 0.3,
      fontSize: 11, bold: true, color: DARK_BG, align: 'center', valign: 'middle', margin: 0
    });
    sl.addText(s[0], { x: bx+0.12, y: by+0.12, w: bw-0.55, h: 0.22, fontSize: 9, color: LGRAY, margin: 0 });
    sl.addText(s[1], { x: bx+0.12, y: by+0.33, w: bw-0.25, h: 0.5, fontSize: 22, bold: true, color: c, margin: 0 });
    sl.addText(s[2], { x: bx+0.12, y: by+0.82, w: bw-0.25, h: 0.22, fontSize: 8.5, color: LGRAY, margin: 0 });
  });

  // bottom note
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0.15, y: 5.15, w: 9.5, h: 0.3,
    fill: { color: DARKER }, line: { type: 'none' }
  });
  sl.addText('데이터 회수율: 100.0%  |  공기밀도: 1.177 kg/m³  |  분석기준: IEC 61400-1 Ed.4  |  계측기간: 2025.01.01 ~ 2025.12.31', {
    x: 0.15, y: 5.15, w: 9.5, h: 0.3,
    fontSize: 8.5, color: LGRAY, align: 'center', valign: 'middle'
  });
}

// ═══════════════════════════════════════════
// SLIDE 11: 부지 적합성 종합 평가
// ═══════════════════════════════════════════
{
  let sl = addSlide(pres);
  accentBar(sl, 0, H, GREEN);

  sl.addText('부지 적합성 종합 평가', {
    x: 0.15, y: 0.1, w: 6, h: 0.42, fontSize: 20, bold: true, color: WHITE, margin: 0
  });
  secBadge(sl, '10', 9.45, 0.1);
  sl.addText('IEC 61400-1 Ed.4 기준  ·  2025 연간 분석 결과', {
    x: 0.15, y: 0.52, w: 9.2, h: 0.28, fontSize: 10, color: LGRAY, italic: true
  });
  hRule(sl, 0.15, 0.78, 9.5);

  sl.addText('IEC 61400-1 Ed.4 기준 7개 항목 평가: 4개 충족(✓), 1개 조건부(⚠, 계측 연장 후 재평가), 2개 초과(✗). 1년 계측으로 통계 신뢰도 제한적이나, 연평균 풍속 및 WPD는 개발 기준 충족. TI·V50 기준 초과는 추가 계측·설계 대응 필요.', {
    x: 0.15, y: 0.88, w: 9.5, h: 0.65,
    fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });

  // Evaluation table
  const hdrs = ['#', '항목', '분석 결과', 'IEC 기준', '판정', '비고·권장사항'];
  const hws2 = [0.3, 1.0, 1.7, 1.3, 0.5, 4.4];
  let hx2 = 0.1;
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0.1, y: 1.62, w: 9.7, h: 0.3,
    fill: { color: GREEN }, line: { type: 'none' }
  });
  hdrs.forEach((h, i) => {
    sl.addText(h, { x: hx2, y: 1.62, w: hws2[i], h: 0.3, fontSize: 8.5, bold: true, color: DARK_BG, valign: 'middle', margin: 0 });
    hx2 += hws2[i];
  });

  const evalRows = [
    ['01', '평균 풍속', '6.342 m/s @ 100m', '≥6~7 m/s', '✓', GREEN, 'IEC 기준 충족. α=0.222 적용 120m에서 6.604 m/s.'],
    ['02', '풍향 분포', 'WNW 24.1% / 에너지 20.4%', '단일 우세', '✓', GREEN, '편서계열 에너지 40.8%. 배열 WNW 기준 최적화 권장.'],
    ['03', '와이블', 'k=2.093  C=7.172 m/s', 'C≥6.0 m/s', '✓', GREEN, 'k≈2.09 중간형. C 기준 충족. 겨울 최고, 여름 최저.'],
    ['04', 'TI P90', '0.2300 @ 15m/s', '≤0.16 (Class A)', '✗', RED, 'Class A 43.8% 초과. IEC Class S 기종 필수.'],
    ['05', 'Wind Shear α', '0.222 연간 (봄 0.225)', '0.14~0.20', '⚠', YELLOW, 'IEC 기준 소폭 초과. DLC 반영 요.'],
    ['06', 'WPD', '281.8 W/m² @ 100m', '≥200 W/m²', '✓', GREEN, '개발 기준 40.9% 상회. 계절 변동 (88~615) 고려 요.'],
    ['07', '극한풍속 V50', '42.30 m/s (Gumbel)', '≤37.5 (III)', '✗', RED, 'Class III 초과. 1년 12개 샘플 → 계측 연장 필수.'],
  ];
  evalRows.forEach((row, ri) => {
    const ry = 1.94 + ri * 0.49;
    const bg = ri % 2 === 0 ? '0F2840' : '071117';
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.1, y: ry, w: 9.7, h: 0.46, fill: { color: bg }, line: { type: 'none' } });
    const cols = [row[0], row[1], row[2], row[3], row[4], row[6]];
    let cx = 0.1;
    cols.forEach((c, ci) => {
      const txtColor = ci === 4 ? row[5] : (ci === 0 ? LGRAY : WHITE);
      sl.addText(c, { x: cx, y: ry, w: hws2[ci], h: 0.46, fontSize: ci===5 ? 7.5 : 8.5, color: txtColor, bold: ci===4, valign: 'middle', margin: 0 });
      cx += hws2[ci];
    });
  });
}

// ═══════════════════════════════════════════
// SLIDE 12: 결론 및 향후 권장 조치
// ═══════════════════════════════════════════
{
  let sl = addSlide(pres);
  accentBar(sl, 0, H, CYAN);

  sl.addText('결론 및 향후 권장 조치', {
    x: 0.15, y: 0.1, w: 7, h: 0.42, fontSize: 20, bold: true, color: WHITE, margin: 0
  });
  secBadge(sl, '11', 9.45, 0.1);
  sl.addText('1개년 분석 기반  ·  계측 연장 · 기종 선정 · 허브 높이 최적화', {
    x: 0.15, y: 0.52, w: 9.2, h: 0.28, fontSize: 10, color: LGRAY, italic: true
  });
  hRule(sl, 0.15, 0.78, 9.5);

  sl.addText('2025년 계측 분석 결과, 연평균 풍속·WPD·와이블·풍향 기준 충족. TI·V50 기준 초과 및 Wind Shear 소폭 초과가 핵심 과제. 1년 계측 데이터로 통계 신뢰도 제한적이며 추가 계측이 필수적임. 즉시 조치부터 중기 설계까지 단계별 접근으로 상업 개발 가능성 확보 가능.', {
    x: 0.15, y: 0.88, w: 9.5, h: 0.65,
    fontSize: 9.5, color: LGRAY, wrap: true, lineSpacingMultiple: 1.4
  });

  // 4 action boxes
  const actions = [
    ['🔴  즉시 조치', RED, [
      'TI·V50 기준 초과 원인 분석 — 계측기 센서 점검',
      '최저 기온 -13.5°C → LT(-20°C) 결빙 방지 시스템 검토',
      '계측 2~3년 연장 계획 수립 → V50 통계 신뢰도 향상',
    ]],
    ['🟡  단기 조치', YELLOW, [
      '포항 ASOS + ERA5 재분析 데이터 MCP 분析',
      '장기 평균 풍속 확정 후 사업성 재평가',
      'CFD 시뮬레이션: α=0.222 복잡 지형 3D 유동 해석',
    ]],
    ['🔵  중기 조치 (설계·기종)', LBLUE, [
      'IEC Class S 고난류 인증 기종 필수 (TI P90=0.230)',
      '허브 높이 최적화: 120m→6.604 m/s  /  150m→6.939 m/s',
      '저온 대응: LT(-20°C) 패키지 + 블레이드 히팅',
    ]],
    ['🟢  장기 조치 (확정 평가)', GREEN, [
      '계측 3개년 이상 확보 → V50 재산출 및 Class 판정',
      '최종 AEP 산출: P50/P90 + Wake 손실 + 공기밀도 보정',
      'IEC 인증: DLC(설계 하중 계산) + 구조 검증 완료',
    ]],
  ];
  const abw = 4.7, abh = 1.55, agap = 0.1;
  actions.forEach((a, i) => {
    const col = i % 2, row2 = Math.floor(i / 2);
    const ax = 0.1 + col * (abw + agap);
    const ay = 1.62 + row2 * (abh + agap);
    sl.addShape(pres.shapes.RECTANGLE, {
      x: ax, y: ay, w: abw, h: abh,
      fill: { color: DARKER }, line: { color: a[1], pt: 1 }
    });
    sl.addShape(pres.shapes.RECTANGLE, {
      x: ax, y: ay, w: abw, h: 0.3,
      fill: { color: a[1], transparency: 30 }, line: { type: 'none' }
    });
    sl.addText(a[0], {
      x: ax+0.1, y: ay+0.02, w: abw-0.15, h: 0.28,
      fontSize: 10, bold: true, color: WHITE, valign: 'middle', margin: 0
    });
    a[2].forEach((bullet, bi) => {
      sl.addText([{ text: '▸  ', options: { color: a[1], bold: true } }, { text: bullet, options: { color: WHITE } }], {
        x: ax+0.1, y: ay+0.38+bi*0.37, w: abw-0.2, h: 0.35,
        fontSize: 8.8, wrap: true
      });
    });
  });
}

pres.writeFile({ fileName: '/home/claude/권이리_풍황분석보고서_2025_KO.pptx' })
  .then(() => console.log('Korean PPTX saved!'));
