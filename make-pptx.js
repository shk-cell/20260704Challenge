const PptxGenJS = require('pptxgenjs');
const path = require('path');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE'; // 13.33" x 7.5"

const DIR = 'C:\\Claude\\20260704Challenge\\';
const W = 13.33, H = 7.5;

const C = {
  blue:    '0284c7',
  green:   '10b981',
  text:    '1e293b',
  muted:   '64748b',
  light:   '94a3b8',
  border:  'e2e8f0',
  bg:      'f8fafc',
  panel:   'ffffff',
  blue_bg: 'eff6ff',
  blue_bd: 'bfdbfe',
  green_bg:'f0fdf4',
  green_bd:'bbf7d0',
  red:     'ef4444',
  red_bg:  'fff1f2',
  red_bd:  'fecaca',
};

const KR   = '맑은 고딕';
const MONO = 'Courier New';

// ── helpers ──────────────────────────────
function tag(s, x, y, w, text, color = C.green) {
  s.addText(text, {
    x, y, w, h: 0.35,
    fontSize: 10, fontFace: MONO, bold: true,
    color, charSpacing: 4,
  });
}

function title(s, x, y, w, text, opts = {}) {
  s.addText(text, {
    x, y, w, h: opts.h || 0.9,
    fontSize: opts.size || 28, fontFace: KR, bold: true,
    color: opts.color || C.text, lineSpacingMultiple: 1.3,
    align: opts.align || 'left',
    ...opts,
  });
}

function body(s, x, y, w, text, opts = {}) {
  s.addText(text, {
    x, y, w, h: opts.h || 2,
    fontSize: 13, fontFace: KR,
    color: C.muted, lineSpacingMultiple: 1.65,
    ...opts,
  });
}

function card(s, x, y, w, h, opts = {}) {
  s.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h,
    fill: { color: opts.fill || C.panel },
    line: { color: opts.border || C.border, width: opts.lw || 1 },
    rectRadius: opts.r || 0.1,
  });
}

function imgBox(s, file, x, y, w, h) {
  s.addImage({ path: DIR + file, x, y, w, h });
  s.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h,
    fill: { type: 'none' },
    line: { color: C.border, width: 1 },
    rectRadius: 0.1,
  });
}

// ═══════════════════════════════════════════
// SLIDE 1 — 타이틀  (라이트)
// ═══════════════════════════════════════════
{
  const s = pptx.addSlide();
  s.background = { color: C.bg };

  // Top accent bar
  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: W, h: 0.06,
    fill: { color: C.blue },
    line: { type: 'none' },
  });

  // Badge
  s.addText('⚡  ARDUINO LEARNING PLATFORM', {
    x: 0, y: 1.0, w: W, h: 0.4,
    fontSize: 11, fontFace: MONO, bold: true,
    color: C.blue, charSpacing: 4, align: 'center',
  });

  // Logo
  s.addText([
    { text: 'Circuit', options: { color: C.blue } },
    { text: 'Lab',     options: { color: C.green } },
  ], {
    x: 0, y: 1.5, w: W, h: 1.7,
    fontSize: 88, fontFace: MONO, bold: true,
    align: 'center', charSpacing: 6,
  });

  // Subtitle
  s.addText('아두이노 회로 학습 플랫폼', {
    x: 0, y: 3.3, w: W, h: 0.5,
    fontSize: 17, fontFace: KR, color: C.muted,
    align: 'center', charSpacing: 2,
  });

  // Divider
  s.addShape(pptx.ShapeType.rect, {
    x: 4.5, y: 3.95, w: 4.3, h: 0.02,
    fill: { color: C.border }, line: { type: 'none' },
  });

  // Stats
  const stats = [
    { num: '10', lbl: 'MISSIONS',     c: C.blue },
    { num: '9',  lbl: 'COMPONENTS',   c: C.green },
    { num: 'GPT',lbl: 'AI GRADING',   c: C.blue },
    { num: 'EXE',lbl: 'ONE-CLICK',    c: C.green },
  ];
  const sw = 3.0;
  stats.forEach((st, i) => {
    const sx = (W - sw * 4) / 2 + i * sw;
    if (i > 0) {
      s.addShape(pptx.ShapeType.rect, {
        x: sx, y: 4.2, w: 0.01, h: 1.1,
        fill: { color: C.border }, line: { type: 'none' },
      });
    }
    s.addText(st.num, {
      x: sx, y: 4.15, w: sw, h: 0.7,
      fontSize: 26, fontFace: MONO, bold: true,
      color: st.c, align: 'center',
    });
    s.addText(st.lbl, {
      x: sx, y: 4.85, w: sw, h: 0.3,
      fontSize: 9, fontFace: MONO,
      color: C.light, align: 'center', charSpacing: 2,
    });
  });

  // Bottom accent bar
  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: H - 0.06, w: W, h: 0.06,
    fill: { color: C.green }, line: { type: 'none' },
  });
}

// ═══════════════════════════════════════════
// SLIDE 2 — 교육과정 배경  (라이트)
// ═══════════════════════════════════════════
{
  const s = pptx.addSlide();
  s.background = { color: C.panel };

  // Left blue sidebar
  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 0.5, h: H,
    fill: { color: C.blue }, line: { type: 'none' },
  });

  tag(s, 0.8, 0.7, 6, 'WHY CIRCUITLAB?', C.blue);
  title(s, 0.8, 1.15, 5.6, '2022 개정 교육과정,\n컴퓨팅 시스템이 1단원이 된 이유', { size: 26, h: 1.8 });
  body(s, 0.8, 3.15, 5.6,
    '피지컬 컴퓨팅, 즉 아두이노와 같은 하드웨어 제어 교육이\n이제는 선택이 아닌 필수가 되었습니다.\n\n그러나 회로 연결은 학생들이 가장 어려워하는 부분이며,\n선생님도 30명을 혼자 봐주기에 시간이 부족합니다.',
    { h: 3 });

  // Right flow cards
  const rx = 7.2, rw = 5.6;

  card(s, rx, 1.1, rw, 1.1);
  s.addText('4', { x: rx + 0.15, y: 1.18, w: 0.9, h: 0.9, fontSize: 32, fontFace: MONO, bold: true, color: C.light, align: 'center', valign: 'middle' });
  s.addText('2015 개정 교육과정', { x: rx + 1.2, y: 1.22, w: 4.2, h: 0.4, fontSize: 13, fontFace: KR, bold: true, color: C.muted });
  s.addText('컴퓨팅 시스템 — 4단원', { x: rx + 1.2, y: 1.62, w: 4.2, h: 0.35, fontSize: 12, fontFace: KR, color: C.light });

  s.addText('↓', { x: rx, y: 2.3, w: rw, h: 0.5, fontSize: 20, color: C.border, align: 'center' });

  card(s, rx, 2.9, rw, 1.1, { fill: C.green_bg, border: C.green_bd, lw: 1.5 });
  s.addText('1', { x: rx + 0.15, y: 2.98, w: 0.9, h: 0.9, fontSize: 32, fontFace: MONO, bold: true, color: C.green, align: 'center', valign: 'middle' });
  s.addText('2022 개정 교육과정', { x: rx + 1.2, y: 3.02, w: 4.2, h: 0.4, fontSize: 13, fontFace: KR, bold: true, color: C.text });
  s.addText('컴퓨팅 시스템 — 1단원으로 격상', { x: rx + 1.2, y: 3.42, w: 4.2, h: 0.35, fontSize: 12, fontFace: KR, bold: true, color: C.green });

  card(s, rx, 4.2, rw, 0.7, { fill: C.blue_bg, border: C.blue_bd });
  s.addText('그만큼 현시대에서의 중요성이 부각되었습니다 📌', {
    x: rx, y: 4.2, w: rw, h: 0.7,
    fontSize: 12, fontFace: KR, bold: true,
    color: C.blue, align: 'center', valign: 'middle',
  });
}

// ═══════════════════════════════════════════
// SLIDE 3 — 문제점  (라이트)
// ═══════════════════════════════════════════
{
  const s = pptx.addSlide();
  s.background = { color: C.bg };

  tag(s, 0, 0.55, W, 'PROBLEM', C.muted);
  title(s, 1, 1.0, W - 2, '아두이노 교실에서 무슨 일이 벌어질까요?', { size: 26, h: 0.8, align: 'center', color: C.text });

  const problems = [
    { icon: '😕', t: '학생이 왜 틀렸는지\n스스로 모릅니다', d: '회로 하나만 잘못 연결해도 전체가 작동하지 않는데, 어디가 문제인지 파악하기 어렵습니다.' },
    { icon: '👨‍🏫', t: '선생님이 30명을\n다 봐줄 수 없습니다', d: '한 명 한 명 회로를 확인해 주기에는 수업 시간이 턱없이 부족합니다.' },
    { icon: '💸', t: '실물 부품의 한계', d: '재료 구입 비용, 잘못 연결하면 부품이 파손될 위험. 자유로운 반복 실습이 어렵습니다.' },
  ];
  const cw = (W - 1.2) / 3 - 0.2;
  problems.forEach((p, i) => {
    const cx = 0.6 + i * (cw + 0.2);
    card(s, cx, 2.1, cw, 4.5, { fill: C.panel });
    s.addText(p.icon, { x: cx, y: 2.4, w: cw, h: 0.8, fontSize: 36, align: 'center' });
    s.addText(p.t, { x: cx + 0.2, y: 3.35, w: cw - 0.4, h: 1.1, fontSize: 14, fontFace: KR, bold: true, color: C.text, align: 'center', lineSpacingMultiple: 1.35 });
    s.addShape(pptx.ShapeType.rect, { x: cx + cw / 2 - 0.25, y: 4.55, w: 0.5, h: 0.04, fill: { color: C.blue }, line: { type: 'none' } });
    s.addText(p.d, { x: cx + 0.2, y: 4.7, w: cw - 0.4, h: 1.6, fontSize: 12, fontFace: KR, color: C.muted, align: 'center', lineSpacingMultiple: 1.5 });
  });
}

// ═══════════════════════════════════════════
// SLIDE 4 — 솔루션  (라이트)
// ═══════════════════════════════════════════
{
  const s = pptx.addSlide();
  s.background = { color: C.panel };

  // Top green accent bar
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.06, fill: { color: C.green }, line: { type: 'none' } });

  tag(s, 0, 0.65, W, 'SOLUTION', C.green);
  title(s, 1.5, 1.1, W - 3, '소스 코드를 보고 회로를 직접 설계하면,\nAI가 즉시 무엇이 틀렸는지 콕 집어 알려줍니다.', {
    size: 24, h: 1.5, align: 'center', color: C.text,
  });

  // Quote box
  card(s, 1.8, 2.85, W - 3.6, 1.4);
  // Blue left bar
  s.addShape(pptx.ShapeType.rect, { x: 1.8, y: 2.85, w: 0.08, h: 1.4, fill: { color: C.blue }, line: { type: 'none' } });
  s.addText('"여기 이 핀이 잘못 연결됐어요. 이 선을 GND에 연결해보세요."\n— CircuitLab AI 피드백 예시', {
    x: 2.1, y: 2.9, w: W - 4.2, h: 1.3,
    fontSize: 14, fontFace: KR, color: C.text,
    lineSpacingMultiple: 1.55, valign: 'middle',
  });

  // Pills
  const pills = ['🎓  학생은 빠르게 성장하고', '⏰  선생님은 시간을 얻고', '✨  수업은 더 풍요로워집니다'];
  const pw = 3.5, pg = 0.25;
  const pstart = (W - (pw * 3 + pg * 2)) / 2;
  pills.forEach((p, i) => {
    const px = pstart + i * (pw + pg);
    card(s, px, 4.6, pw, 0.65, { r: 0.32 });
    s.addText(p, { x: px, y: 4.6, w: pw, h: 0.65, fontSize: 13, fontFace: KR, bold: true, color: C.text, align: 'center', valign: 'middle' });
  });
}

// ═══════════════════════════════════════════
// SLIDE 5 — 설치 편의성  (라이트)
// ═══════════════════════════════════════════
{
  const s = pptx.addSlide();
  s.background = { color: C.bg };

  tag(s, 0, 0.55, W, 'INSTALLATION', C.muted);
  title(s, 1, 1.0, W - 2, '크롬 로그인? 회원가입? 전혀 필요 없습니다.', { size: 26, h: 0.8, align: 'center' });

  // Before
  card(s, 0.7, 2.1, 5.5, 4.5);
  s.addText('BEFORE — 기존 도구', { x: 1.0, y: 2.35, w: 5.0, h: 0.3, fontSize: 9, fontFace: MONO, bold: true, color: C.red, charSpacing: 3 });
  s.addText('Tinkercad 등', { x: 1.0, y: 2.75, w: 5.0, h: 0.45, fontSize: 16, fontFace: KR, bold: true, color: C.text });
  s.addShape(pptx.ShapeType.rect, { x: 1.0, y: 3.28, w: 5.0, h: 0.02, fill: { color: C.border }, line: { type: 'none' } });
  const bef = ['구글 로그인 필요', '브라우저 환경 설정', '인터넷 연결 필수', '수업 전 준비 시간 소요'];
  bef.forEach((t, i) => s.addText('·  ' + t, { x: 1.0, y: 3.45 + i * 0.58, w: 5.0, h: 0.5, fontSize: 13, fontFace: KR, color: C.muted }));

  // VS
  s.addText('VS', { x: 6.0, y: 3.8, w: 1.3, h: 0.7, fontSize: 16, fontFace: MONO, bold: true, color: C.border, align: 'center' });

  // After
  card(s, 7.1, 2.1, 5.5, 4.5, { fill: C.green_bg, border: C.green_bd, lw: 1.5 });
  s.addText('AFTER — CircuitLab', { x: 7.4, y: 2.35, w: 5.0, h: 0.3, fontSize: 9, fontFace: MONO, bold: true, color: C.green, charSpacing: 3 });
  s.addText('exe 파일 하나로 끝!', { x: 7.4, y: 2.75, w: 5.0, h: 0.45, fontSize: 16, fontFace: KR, bold: true, color: C.text });
  s.addShape(pptx.ShapeType.rect, { x: 7.4, y: 3.28, w: 5.0, h: 0.02, fill: { color: C.green_bd }, line: { type: 'none' } });
  const aft = ['로그인 불필요', '설치 후 바탕화면 아이콘 클릭', '오프라인 동작', '단 한 번 설치로 언제든지'];
  aft.forEach((t, i) => s.addText('✓  ' + t, { x: 7.4, y: 3.45 + i * 0.58, w: 5.0, h: 0.5, fontSize: 13, fontFace: KR, bold: true, color: '166534' }));
}

// ═══════════════════════════════════════════
// SLIDE 6 — 미션 목차 스크린샷  (라이트)
// ═══════════════════════════════════════════
{
  const s = pptx.addSlide();
  s.background = { color: C.panel };

  tag(s, 0, 0.5, W, 'SCREENSHOTS', C.green);
  title(s, 0, 0.95, W, '10가지 다양한 미션', { size: 26, h: 0.65, align: 'center' });
  s.addText('LED, 버튼, 초음파 센서, 서보모터, 조도 센서까지 — 체계적으로 난이도가 올라갑니다', {
    x: 1, y: 1.65, w: W - 2, h: 0.4, fontSize: 13, fontFace: KR, color: C.muted, align: 'center',
  });
  imgBox(s, 'Circuit02.png', 2.2, 2.15, 8.9, 5.0);
}

// ═══════════════════════════════════════════
// SLIDE 7 — 회로 에디터  (라이트)
// ═══════════════════════════════════════════
{
  const s = pptx.addSlide();
  s.background = { color: C.bg };

  tag(s, 0, 0.5, W, 'CIRCUIT EDITOR', C.blue);
  title(s, 0, 0.95, W, '소스 코드 분석 → 회로 직접 설계', { size: 26, h: 0.65, align: 'center' });
  s.addText('왼쪽 코드를 읽고 오른쪽 캔버스에서 부품을 드래그하고 점퍼선을 연결합니다', {
    x: 1, y: 1.65, w: W - 2, h: 0.4, fontSize: 13, fontFace: KR, color: C.muted, align: 'center',
  });

  imgBox(s, 'Circuit03.png', 0.4, 2.2, 6.15, 4.8);
  // Caption
  s.addText('미션 설명 + 소스 코드 확인', { x: 0.4, y: 7.1, w: 6.15, h: 0.3, fontSize: 11, fontFace: KR, color: C.muted, align: 'center' });

  imgBox(s, 'Circuit04.png', 6.78, 2.2, 6.15, 4.8);
  s.addText('부품 드래그로 위치 자유 변경', { x: 6.78, y: 7.1, w: 6.15, h: 0.3, fontSize: 11, fontFace: KR, color: C.muted, align: 'center' });
}

// ═══════════════════════════════════════════
// SLIDE 8 — 다양한 미션 스크린샷  (라이트)
// ═══════════════════════════════════════════
{
  const s = pptx.addSlide();
  s.background = { color: C.panel };

  tag(s, 0, 0.5, W, 'MISSION EXAMPLES', C.green);
  title(s, 0, 0.95, W, '초음파 센서, 서보모터, 조도 센서...', { size: 26, h: 0.65, align: 'center' });

  const iw = (W - 1.2) / 3 - 0.2;
  const imgs = [
    { f: 'Circuit08.png', cap: '가변저항 + 서보모터' },
    { f: 'Circuit10.png', cap: '초음파 센서 + 부저' },
    { f: 'Circuit11.png', cap: '조도 센서 + LED' },
  ];
  imgs.forEach((img, i) => {
    const ix = 0.6 + i * (iw + 0.2);
    imgBox(s, img.f, ix, 2.0, iw, 4.9);
    s.addText(img.cap, { x: ix, y: 7.0, w: iw, h: 0.35, fontSize: 11, fontFace: KR, color: C.muted, align: 'center' });
  });
}

// ═══════════════════════════════════════════
// SLIDE 9 — AI 채점 오답  (라이트)
// ═══════════════════════════════════════════
{
  const s = pptx.addSlide();
  s.background = { color: C.bg };

  // Left text
  s.addText('🤔', { x: 0.7, y: 0.6, w: 1.2, h: 1.0, fontSize: 52 });

  card(s, 0.7, 1.75, 2.3, 0.44, { fill: C.red_bg, border: C.red_bd, r: 0.22 });
  s.addText('WRONG ANSWER', { x: 0.7, y: 1.77, w: 2.3, h: 0.4, fontSize: 9, fontFace: MONO, bold: true, color: C.red, align: 'center', valign: 'middle', charSpacing: 1 });

  title(s, 0.7, 2.35, 5.7, '무엇이 왜 틀렸는지\nAI가 콕 집어 알려줍니다', { size: 24, h: 1.3, color: C.text });
  body(s, 0.7, 3.85, 5.7,
    '단순히 "오답"이 아닙니다.\n어느 핀이 잘못 연결됐는지,\n무엇을 고쳐야 하는지 구체적으로 안내합니다.\n\n초보자도 혼자서 실력을 키울 수 있습니다.',
    { h: 2.9 });

  imgBox(s, 'Circuit06.png', 6.7, 0.4, 6.2, 6.7);
}

// ═══════════════════════════════════════════
// SLIDE 10 — AI 채점 정답  (라이트)
// ═══════════════════════════════════════════
{
  const s = pptx.addSlide();
  s.background = { color: C.bg };

  imgBox(s, 'Circuit07.png', 0.4, 0.4, 6.2, 6.7);

  // Right text
  s.addText('🎉', { x: 7.2, y: 0.6, w: 1.2, h: 1.0, fontSize: 52 });

  card(s, 7.2, 1.75, 2.0, 0.44, { fill: C.green_bg, border: C.green_bd, r: 0.22 });
  s.addText('CORRECT!', { x: 7.2, y: 1.77, w: 2.0, h: 0.4, fontSize: 9, fontFace: MONO, bold: true, color: C.green, align: 'center', valign: 'middle', charSpacing: 2 });

  title(s, 7.2, 2.35, 5.7, '정답! 회로 동작 원리까지\n상세히 설명해줍니다', { size: 24, h: 1.3, color: C.text });
  body(s, 7.2, 3.85, 5.7,
    '정답일 때도 AI는 단순히 "통과"가 아닙니다.\n전류가 어떻게 흐르는지,\n왜 이 연결이 맞는지 원리까지 설명해 줍니다.\n\n학습 효과가 2배로 높아집니다.',
    { h: 2.9 });
}

// ═══════════════════════════════════════════
// SLIDE 11 — 마무리  (라이트)
// ═══════════════════════════════════════════
{
  const s = pptx.addSlide();
  s.background = { color: C.panel };

  // Top gradient bar
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W / 2, h: 0.1, fill: { color: C.blue }, line: { type: 'none' } });
  s.addShape(pptx.ShapeType.rect, { x: W / 2, y: 0, w: W / 2, h: 0.1, fill: { color: C.green }, line: { type: 'none' } });

  // Logo
  s.addText([
    { text: 'Circuit', options: { color: C.blue } },
    { text: 'Lab',     options: { color: C.green } },
  ], {
    x: 0, y: 1.0, w: W, h: 1.5,
    fontSize: 80, fontFace: MONO, bold: true,
    align: 'center', charSpacing: 6,
  });

  s.addText('더 풍요로운 Arduino 수업을 만들어 보세요.', {
    x: 1, y: 2.65, w: W - 2, h: 0.55,
    fontSize: 16, fontFace: KR, color: C.muted,
    align: 'center', charSpacing: 1,
  });

  // Divider
  s.addShape(pptx.ShapeType.rect, { x: 4.0, y: 3.35, w: 5.3, h: 0.02, fill: { color: C.border }, line: { type: 'none' } });

  // Fin pills
  const fins = ['🎓  학생은 성장하고', '⏰  선생님은 시간을 얻고', '✨  수업은 더 풍요로워집니다'];
  const fw = 3.3, fg = 0.22;
  const fstart = (W - (fw * 3 + fg * 2)) / 2;
  fins.forEach((p, i) => {
    const fx = fstart + i * (fw + fg);
    card(s, fx, 3.55, fw, 0.65, { r: 0.32 });
    s.addText(p, { x: fx, y: 3.55, w: fw, h: 0.65, fontSize: 13, fontFace: KR, bold: true, color: C.text, align: 'center', valign: 'middle' });
  });

  // Download button
  s.addShape(pptx.ShapeType.roundRect, {
    x: 4.0, y: 4.7, w: 5.3, h: 0.85,
    fill: { color: C.blue }, line: { type: 'none' }, rectRadius: 0.42,
  });
  s.addText('⬇   CircuitLab 다운로드 (Windows)', {
    x: 4.0, y: 4.7, w: 5.3, h: 0.85,
    fontSize: 14, fontFace: KR, bold: true,
    color: 'ffffff', align: 'center', valign: 'middle',
  });
  s.addText('CircuitLab.Setup.1.0.0.exe  ·  Windows 전용', {
    x: 0, y: 5.75, w: W, h: 0.3,
    fontSize: 10, fontFace: MONO, color: C.light, align: 'center',
  });
}

// ═══════════════════════════════════════════
// SAVE
// ═══════════════════════════════════════════
pptx.writeFile({ fileName: DIR + 'CircuitLab_발표.pptx' })
  .then(() => console.log('✅ CircuitLab_발표.pptx 생성 완료!'))
  .catch(e => { console.error('❌ 오류:', e); process.exit(1); });
