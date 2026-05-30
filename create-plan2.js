const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "AP 설계팀의 Hardware Debugging Process 자동화 (plan2)";

// ── Color palette ─────────────────────────────────────────────────────────────
const BG_DARK  = "05091A";
const BG_MID   = "07101F";
const BG_BLACK = "03060F";
const PANEL    = "0A1525";
const PANEL2   = "0A1220";
const BORDER   = "1E3A5F";
const BORDER2  = "1A2E50";
const CYAN     = "00D4FF";
const MINT     = "00FFD1";
const AMBER    = "FFB800";
const RED      = "FF4A4A";
const WHITE    = "FFFFFF";
const GRAY_LT  = "A8C8E8";
const GRAY_MD  = "5A7A9A";
const GRAY_DK  = "3A5A7A";
const PURPLE   = "A070FF";

function ln(color, width, transparency) {
  const obj = { color, width: width || 0.5 };
  if (transparency !== undefined) obj.transparency = transparency;
  return obj;
}
function fl(color, transparency) {
  const obj = { color };
  if (transparency !== undefined) obj.transparency = transparency;
  return obj;
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 1 — TITLE  (subtitle 변경: "기존 데이터 자산 + 도메인 이해 + 단순한 도구")
// ══════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: BG_DARK };

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.04, fill: fl(CYAN), line: ln(CYAN, 0) });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.585, w: 10, h: 0.04, fill: fl(CYAN), line: ln(CYAN, 0) });
  for (let i = 1; i < 9; i++)  sl.addShape(pres.shapes.LINE, { x: 0, y: i * 0.625, w: 10, h: 0, line: ln(CYAN, 0.3, 88) });
  for (let i = 1; i < 16; i++) sl.addShape(pres.shapes.LINE, { x: i * 0.625, y: 0, w: 0, h: 5.625, line: ln(CYAN, 0.3, 88) });

  sl.addShape(pres.shapes.LINE, { x: 0.5, y: 0.42, w: 0.83, h: 0, line: ln(CYAN, 1, 70) });
  sl.addShape(pres.shapes.LINE, { x: 0.5, y: 0.42, w: 0, h: 0.83, line: ln(CYAN, 1, 70) });
  sl.addShape(pres.shapes.LINE, { x: 8.67, y: 5.2, w: 0.83, h: 0, line: ln(CYAN, 1, 70) });
  sl.addShape(pres.shapes.LINE, { x: 9.5, y: 4.37, w: 0, h: 0.83, line: ln(CYAN, 1, 70) });

  sl.addShape(pres.shapes.RECTANGLE, { x: 0.83, y: 0.9, w: 2.7, h: 0.28, fill: fl("0A1828"), line: ln(CYAN, 0.5, 60) });
  sl.addText("AI / 자동화 활용 사례 공유", { x: 0.83, y: 0.9, w: 2.7, h: 0.28, fontSize: 9, fontFace: "Pretendard", color: CYAN, bold: true, align: "center", valign: "middle", margin: 0 });

  sl.addText([
    { text: "AP 설계팀의\n", options: { color: WHITE } },
    { text: "Hardware Debugging\n", options: { color: CYAN } },
    { text: "Process 자동화", options: { color: WHITE } },
  ], { x: 0.83, y: 1.3, w: 7.5, h: 2.1, fontSize: 36, fontFace: "Pretendard", bold: true, lineSpacingMultiple: 1.2 });

  // subtitle 변경
  sl.addText("기존 데이터 자산 + 도메인 이해 + 단순한 도구 = 전문가의 반복 작업 제거", { x: 0.83, y: 3.5, w: 7.5, h: 0.35, fontSize: 13, fontFace: "Pretendard", color: GRAY_MD });

  sl.addShape(pres.shapes.LINE, { x: 0.83, y: 4.0, w: 3.5, h: 0, line: ln(BORDER, 0.5) });
  const metaY = 4.15;
  sl.addText("TEAM",     { x: 0.83, y: metaY,       w: 2, h: 0.18, fontSize: 8,  fontFace: "Pretendard", color: "1A8899", bold: true, charSpacing: 2 });
  sl.addText("AP 설계팀 SW 담당", { x: 0.83, y: metaY+0.2, w: 2, h: 0.22, fontSize: 12, fontFace: "Pretendard", color: GRAY_LT, bold: true });
  sl.addText("DURATION", { x: 3.3, y: metaY,        w: 2, h: 0.18, fontSize: 8,  fontFace: "Pretendard", color: "1A8899", bold: true, charSpacing: 2 });
  sl.addText("60분",      { x: 3.3, y: metaY+0.2,   w: 2, h: 0.22, fontSize: 12, fontFace: "Pretendard", color: GRAY_LT, bold: true });
  sl.addText("DATE",     { x: 5.3, y: metaY,        w: 2, h: 0.18, fontSize: 8,  fontFace: "Pretendard", color: "1A8899", bold: true, charSpacing: 2 });
  sl.addText("2026",     { x: 5.3, y: metaY+0.2,   w: 2, h: 0.22, fontSize: 12, fontFace: "Pretendard", color: GRAY_LT, bold: true });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 2 — TOC  (desc 문구 변경)
// ══════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: "080D1F" };

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 2.92, h: 5.625, fill: fl("0D1A3A"), line: ln(BORDER2, 0.5) });
  sl.addText("TODAY'S STORY", { x: 0.2, y: 0.9, w: 2.5, h: 0.2, fontSize: 8, fontFace: "Pretendard", color: CYAN, bold: true, charSpacing: 1.5 });
  sl.addText("오늘의\n이야기", { x: 0.2, y: 1.15, w: 2.5, h: 1.0, fontSize: 24, fontFace: "Pretendard", color: WHITE, bold: true, lineSpacingMultiple: 1.2 });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.2, y: 2.25, w: 0.42, h: 0.03, fill: fl(CYAN), line: ln(CYAN, 0) });
  sl.addText("문제 발굴부터 솔루션까지,\n그리고 당신 팀에서의 적용까지", { x: 0.2, y: 2.35, w: 2.5, h: 0.7, fontSize: 10, fontFace: "Pretendard", color: GRAY_DK, lineSpacingMultiple: 1.4 });

  const tocItems = [
    { num: "1", title: "문제: HW 디버깅 병목", desc: "칩이 나왔다. 버그가 있다. 담당자 의존 구조의 병목", highlight: true },
    { num: "2", title: "사전 준비: 데이터 자산화", desc: "자동화 이전에 쌓아온 인프라의 가치", highlight: false },
    { num: "3", title: "솔루션: 무엇을 만들었나", desc: "아키텍처 + 라이브 데모", highlight: false },
    { num: "4", title: "프레임워크: 당신 팀에서도", desc: "자동화 기회를 발굴하는 3가지 질문", highlight: false },
  ];

  const tocX = 3.15, tocW = 6.5;
  tocItems.forEach((item, i) => {
    const y = 0.45 + i * 1.2;
    sl.addShape(pres.shapes.RECTANGLE, { x: tocX, y, w: tocW, h: 1.0, fill: fl(item.highlight ? "0A1830" : "0A1225"), line: ln(item.highlight ? "1A4A70" : BORDER, 0.5) });
    sl.addShape(pres.shapes.OVAL, { x: tocX+0.18, y: y+0.25, w: 0.45, h: 0.45, fill: fl(item.highlight ? "112233" : "0A1828"), line: ln(CYAN, 0.8) });
    sl.addText(item.num, { x: tocX+0.18, y: y+0.25, w: 0.45, h: 0.45, fontSize: 11, fontFace: "Pretendard", color: CYAN, bold: true, align: "center", valign: "middle", margin: 0 });
    sl.addText(item.title, { x: tocX+0.75, y: y+0.12, w: tocW-1.2, h: 0.32, fontSize: 13, fontFace: "Pretendard", color: "D8EAF8", bold: true });
    sl.addText(item.desc,  { x: tocX+0.75, y: y+0.48, w: tocW-1.2, h: 0.28, fontSize: 10, fontFace: "Pretendard", color: GRAY_DK });
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 3 — HW DEBUGGING  (callout 문구 변경: 담당자 의존 구조 언급)
// ══════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: BG_MID };

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.83, fill: fl("0D1B3A"), line: ln(BORDER2, 0.5) });
  sl.addText("PART 1 · 문제", { x: 0.58, y: 0.1, w: 2, h: 0.27, fontSize: 8, fontFace: "Pretendard", color: CYAN, bold: true, charSpacing: 2 });
  sl.addText("HW 디버깅이란 무엇인가", { x: 0.58, y: 0.38, w: 6, h: 0.35, fontSize: 18, fontFace: "Pretendard", color: WHITE, bold: true });

  sl.addText("칩이 처음 나왔다. 6개월, 수백억 원.\n그런데 버그가 있다. 어디서, 왜 발생했는지 찾아야 한다.", { x: 0.58, y: 1.0, w: 7.5, h: 0.55, fontSize: 12, fontFace: "Pretendard", color: GRAY_MD, lineSpacingMultiple: 1.4 });

  const steps = [
    { icon: "💡", label: "버그 발생",    sub: "Silicon 동작 이상",      boxColor: "0A1F3A", borderColor: BORDER },
    { icon: "📥", label: "Scandump",    sub: "칩 상태 덤프",           boxColor: "0A1F3A", borderColor: BORDER },
    { icon: "📄", label: "Regbank 생성", sub: "수 MB 레지스터 파일",    boxColor: "0F1A30", borderColor: BORDER },
    { icon: "🔍", label: "수동 분석",    sub: "전문가가 직접",           boxColor: "1C0A24", borderColor: "4A1A60" },
    { icon: "✅", label: "원인 파악",    sub: "디버깅 완료",             boxColor: "0A2418", borderColor: "1A5A30" },
  ];

  const flowY = 1.7, stepW = 1.3, stepH = 0.85, gapX = 0.45;
  const startX = (10 - (steps.length * stepW + (steps.length - 1) * gapX)) / 2;
  steps.forEach((step, i) => {
    const x = startX + i * (stepW + gapX);
    sl.addShape(pres.shapes.RECTANGLE, { x, y: flowY, w: stepW, h: stepH, fill: fl(step.boxColor), line: ln(step.borderColor, 0.8) });
    sl.addText(step.icon,  { x, y: flowY+0.08, w: stepW, h: 0.4,  fontSize: 22, fontFace: "Pretendard", align: "center", valign: "middle", margin: 0 });
    sl.addText(step.label, { x, y: flowY+0.52, w: stepW, h: 0.2,  fontSize: 10, fontFace: "Pretendard", color: GRAY_LT, bold: true, align: "center", margin: 0 });
    sl.addText(step.sub,   { x, y: flowY+0.72, w: stepW, h: 0.18, fontSize: 8,  fontFace: "Pretendard", color: GRAY_DK, align: "center", margin: 0 });
    if (i < steps.length - 1) {
      sl.addShape(pres.shapes.LINE, { x: x+stepW+0.05, y: flowY+stepH/2, w: gapX-0.1, h: 0, line: ln(BORDER, 1) });
      sl.addText("▶", { x: x+stepW+gapX-0.2, y: flowY+stepH/2-0.12, w: 0.2, h: 0.25, fontSize: 10, fontFace: "Pretendard", color: BORDER, align: "center", margin: 0 });
    }
  });

  // 변경된 callout 문구
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.58, y: 3.05, w: 8.84, h: 0.65, fill: fl(RED, 94), line: ln(RED, 0.5, 80) });
  sl.addText("⚠️  담당자가 각자 보유한 스크립트로 분석해야 한다.\n       그런데 이슈가 담당자에게 할당되어야 비로소 분석이 시작된다. 이 구조가 병목이다.", { x: 0.65, y: 3.07, w: 8.7, h: 0.6, fontSize: 11, fontFace: "Pretendard", color: "F0A0A0", lineSpacingMultiple: 1.4 });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 4 — 디버깅 병목: 담당자 의존 구조 (BEFORE / AFTER)  ★ 핵심 변경
// ══════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: BG_MID };

  // Header
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.83, fill: fl("0D1B3A"), line: ln(BORDER2, 0.5) });
  sl.addText("PART 1 · 문제", { x: 0.58, y: 0.1, w: 2, h: 0.27, fontSize: 8, fontFace: "Pretendard", color: "FF6B6B", bold: true, charSpacing: 2 });
  sl.addText("디버깅 병목: 담당자 의존 구조", { x: 0.58, y: 0.36, w: 7, h: 0.35, fontSize: 18, fontFace: "Pretendard", color: WHITE, bold: true });

  // ── 세로 구분선
  sl.addShape(pres.shapes.LINE, { x: 5.0, y: 0.95, w: 0, h: 4.5, line: ln(BORDER, 0.8) });

  // ── BEFORE 컬럼 (좌)
  const bx = 0.42, bw = 4.35;
  sl.addText("⚠️  BEFORE — 기존 구조", { x: bx, y: 0.98, w: bw, h: 0.28, fontSize: 10, fontFace: "Pretendard", color: "FF6B6B", bold: true, charSpacing: 0.5 });

  // 스텝 데이터: bottleneck 여부
  const beforeSteps = [
    { icon: "🐛", title: "이슈 발생",         desc: "버그가 리포트된다",                           isBottleneck: false },
    { icon: "🔎", title: "담당자 찾기",        desc: "어느 블록 담당자에게 할당해야 하나?",          isBottleneck: true,  tag: "병목 ①" },
    { icon: "⏳", title: "담당자 대기",        desc: "담당자가 시간 내야 분석 시작 가능",            isBottleneck: true,  tag: "병목 ②" },
    { icon: "📜", title: "개인 스크립트 실행", desc: "각 담당자마다 다른 스크립트 보유",             isBottleneck: false },
  ];

  beforeSteps.forEach((s, i) => {
    const y = 1.38 + i * 0.95;
    const bgColor   = s.isBottleneck ? "1A0808" : "0A1220";
    const brdColor  = s.isBottleneck ? "5A1A1A" : BORDER;
    // left accent bar
    sl.addShape(pres.shapes.RECTANGLE, { x: bx, y, w: 0.05, h: 0.75, fill: fl(s.isBottleneck ? "FF4A4A" : BORDER), line: ln(BORDER, 0) });
    // card
    sl.addShape(pres.shapes.RECTANGLE, { x: bx+0.05, y, w: bw-0.05, h: 0.75, fill: fl(bgColor), line: ln(brdColor, 0.5) });
    sl.addText(s.icon, { x: bx+0.12, y: y+0.12, w: 0.4, h: 0.4, fontSize: 18, fontFace: "Pretendard" });
    sl.addText(s.title, { x: bx+0.58, y: y+0.08, w: 2.6, h: 0.25, fontSize: 11, fontFace: "Pretendard", color: s.isBottleneck ? "F0A0A0" : "D8EAF8", bold: true });
    sl.addText(s.desc,  { x: bx+0.58, y: y+0.35, w: 2.6, h: 0.28, fontSize: 9,  fontFace: "Pretendard", color: GRAY_DK });
    if (s.isBottleneck) {
      sl.addShape(pres.shapes.RECTANGLE, { x: bx+3.3, y: y+0.2, w: 0.88, h: 0.28, fill: fl("FF4A4A", 91), line: ln("FF4A4A", 0.5, 55) });
      sl.addText(s.tag, { x: bx+3.3, y: y+0.2, w: 0.88, h: 0.28, fontSize: 9, fontFace: "Pretendard", color: "FF6B6B", bold: true, align: "center", valign: "middle", margin: 0 });
    }
    // down arrow between steps
    if (i < beforeSteps.length - 1) {
      sl.addText("↓", { x: bx + bw/2 - 0.15, y: y+0.75, w: 0.3, h: 0.2, fontSize: 13, fontFace: "Pretendard", color: "2A4A6A", align: "center", margin: 0 });
    }
  });

  // ── AFTER 컬럼 (우)
  const ax = 5.2, aw = 4.4;
  sl.addText("✅  AFTER — 자동화 시스템", { x: ax, y: 0.98, w: aw, h: 0.28, fontSize: 10, fontFace: "Pretendard", color: MINT, bold: true, charSpacing: 0.5 });

  // After step 1: 이슈 발생
  sl.addShape(pres.shapes.RECTANGLE, { x: ax, y: 1.38, w: aw, h: 0.75, fill: fl(PANEL2), line: ln(BORDER, 0.5) });
  sl.addText("🐛", { x: ax+0.1, y: 1.44, w: 0.4, h: 0.4, fontSize: 18, fontFace: "Pretendard" });
  sl.addText("이슈 발생", { x: ax+0.55, y: 1.44, w: aw-0.7, h: 0.25, fontSize: 11, fontFace: "Pretendard", color: "D8EAF8", bold: true });
  sl.addText("버그가 리포트된다", { x: ax+0.55, y: 1.71, w: aw-0.7, h: 0.22, fontSize: 9, fontFace: "Pretendard", color: GRAY_DK });

  sl.addText("↓", { x: ax + aw/2 - 0.15, y: 2.13, w: 0.3, h: 0.2, fontSize: 13, fontFace: "Pretendard", color: MINT, align: "center", margin: 0 });

  // After step 2: Web UI 바로 확인 (강조)
  sl.addShape(pres.shapes.RECTANGLE, { x: ax, y: 2.35, w: aw, h: 0.88, fill: fl("001E14"), line: ln(MINT, 1, 55) });
  sl.addText("🖥️", { x: ax+0.1, y: 2.41, w: 0.5, h: 0.5, fontSize: 24, fontFace: "Pretendard" });
  sl.addText("Web UI에서 바로 확인", { x: ax+0.68, y: 2.41, w: aw-0.82, h: 0.3, fontSize: 13, fontFace: "Pretendard", color: MINT, bold: true });
  sl.addText("담당자 할당 없이 누구나 즉시 분석 결과 열람", { x: ax+0.68, y: 2.73, w: aw-0.82, h: 0.38, fontSize: 9.5, fontFace: "Pretendard", color: "3A7A6A", lineSpacingMultiple: 1.3 });

  // Result box
  sl.addShape(pres.shapes.RECTANGLE, { x: ax, y: 3.42, w: aw, h: 1.6, fill: fl("001A14"), line: ln(MINT, 1.5, 55) });
  sl.addText("누구나 · 즉시", { x: ax, y: 3.58, w: aw, h: 0.65, fontSize: 30, fontFace: "Pretendard", color: MINT, bold: true, align: "center" });
  sl.addText("담당자를 찾고 대기하는 과정 없이\n이슈 발생 즉시 분석 결과 확인 가능", { x: ax+0.2, y: 4.28, w: aw-0.4, h: 0.6, fontSize: 10, fontFace: "Pretendard", color: "3A7A6A", align: "center", lineSpacingMultiple: 1.4 });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 5 — 전제조건: Spec 데이터 자산화  ("3년" 표현 제거, reason items 구조)
// ══════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: BG_DARK };

  // Amber top bar
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.04, fill: fl(AMBER), line: ln(AMBER, 0) });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.04, w: 10, h: 0.79, fill: fl("1A130A"), line: ln("3A2A0E", 0.5) });
  sl.addText("PART 2 · 자산화", { x: 0.58, y: 0.1, w: 2.5, h: 0.25, fontSize: 8, fontFace: "Pretendard", color: AMBER, bold: true, charSpacing: 2 });
  sl.addText("전제조건: Spec 데이터 자산화", { x: 0.58, y: 0.36, w: 7, h: 0.35, fontSize: 18, fontFace: "Pretendard", color: WHITE, bold: true });

  // Left: intro + 3 reason items
  sl.addText("이번 시스템이 가능했던 이유는\n설계팀 Spec 데이터를 미리 DB로 자산화해두었기 때문이다.", { x: 0.58, y: 0.98, w: 4.7, h: 0.6, fontSize: 11, fontFace: "Pretendard", color: GRAY_MD, lineSpacingMultiple: 1.5 });

  const reasons = [
    { icon: "🗄️", title: "Spec 데이터 DB 구축", desc: "레지스터 명세, 필드 정의, 기대값을\n구조화된 DB로 축적" },
    { icon: "⏳", title: "당시엔 활용처가 불명확했다",  desc: "바로 쓸 곳은 없었지만,\n꾸준히 쌓아왔다" },
    { icon: "⚡", title: "이 자산이 없었다면 불가능",   desc: "Regbank와 대조할 기준 데이터가 없으면\n자동화 자체가 성립하지 않는다" },
  ];
  reasons.forEach((r, i) => {
    const y = 1.72 + i * 1.18;
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.42, y, w: 0.05, h: 0.98, fill: fl(AMBER), line: ln(AMBER, 0) });
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.5,  y, w: 4.6,  h: 0.98, fill: fl("0A1220"), line: ln(BORDER, 0.5) });
    sl.addText(r.icon, { x: 0.58, y: y+0.12, w: 0.5, h: 0.45, fontSize: 20, fontFace: "Pretendard" });
    sl.addText(r.title, { x: 1.15, y: y+0.08, w: 3.8, h: 0.26, fontSize: 11, fontFace: "Pretendard", color: "D0C0A0", bold: true });
    sl.addText(r.desc,  { x: 1.15, y: y+0.36, w: 3.8, h: 0.52, fontSize: 9.5, fontFace: "Pretendard", color: GRAY_DK, lineSpacingMultiple: 1.35 });
  });

  // Right: insight box
  sl.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 1.0, w: 4.2, h: 1.65, fill: fl("1A1200"), line: ln(AMBER, 0.8, 80) });
  sl.addText("💡  핵심 인사이트", { x: 5.55, y: 1.1, w: 3.9, h: 0.28, fontSize: 10, fontFace: "Pretendard", color: AMBER, bold: true, charSpacing: 1 });
  sl.addText("지금 당장 쓸모없어 보이는 데이터 수집이\n미래의 자동화 가능성을 결정한다.\n\nAI Agent / 자동화를 원한다면,\n먼저 데이터를 자산화해야 한다.", { x: 5.55, y: 1.42, w: 3.9, h: 1.1, fontSize: 10, fontFace: "Pretendard", color: "C0A870", lineSpacingMultiple: 1.45 });

  // Right: DB stats
  sl.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 2.8, w: 4.2, h: 1.1, fill: fl("0A0F1E"), line: ln(BORDER, 0.5) });
  sl.addText("SPEC DB + REGBANK 연결", { x: 5.55, y: 2.9, w: 3.9, h: 0.22, fontSize: 8, fontFace: "Pretendard", color: GRAY_DK, bold: true, charSpacing: 1 });
  [{ num: "Spec DB", lbl: "기대값/필드 정의", color: AMBER },
   { num: "+",       lbl: "",                color: GRAY_DK },
   { num: "Regbank", lbl: "실제 덤프 데이터", color: CYAN },
   { num: "=",       lbl: "",                color: GRAY_DK },
   { num: "자동화",   lbl: "가능성",          color: MINT }].forEach((s, i) => {
    const x = 5.55 + i * 0.82;
    sl.addText(s.num, { x, y: 3.15, w: 0.78, h: 0.35, fontSize: s.num.length <= 2 ? 18 : 12, fontFace: "Pretendard", color: s.color, bold: true, align: "center" });
    if (s.lbl) sl.addText(s.lbl, { x, y: 3.52, w: 0.78, h: 0.2, fontSize: 7.5, fontFace: "Pretendard", color: GRAY_DK, align: "center" });
  });

  // Right: key message
  sl.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 4.05, w: 4.2, h: 1.1, fill: fl("0A1220"), line: ln(BORDER, 0.5) });
  sl.addText("HW 팀은 Regbank를 항상 생성하고 있었다\n→  두 자산을 연결하는 것이 이번 시스템의 핵심", { x: 5.55, y: 4.18, w: 3.9, h: 0.85, fontSize: 10, fontFace: "Pretendard", color: GRAY_MD, lineSpacingMultiple: 1.5 });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 6 — PREREQUISITE (동일)
// ══════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: BG_BLACK };

  sl.addText("AI / 자동화의 전제조건", { x: 1, y: 0.55, w: 8, h: 0.25, fontSize: 9, fontFace: "Pretendard", color: CYAN, bold: true, charSpacing: 3, align: "center" });
  sl.addText([
    { text: "자동화에 꼭 ", options: { color: WHITE } },
    { text: "AI/ML이\n", options: { color: CYAN } },
    { text: "필요하지 않다", options: { color: WHITE } },
  ], { x: 1, y: 0.85, w: 8, h: 1.4, fontSize: 34, fontFace: "Pretendard", bold: true, align: "center", lineSpacingMultiple: 1.25 });

  const eqItems = [
    { text: "기존 데이터 자산", isBox: true, color: GRAY_LT, bgColor: "0A1828", borderColor: BORDER },
    { text: "+", isBox: false, color: GRAY_DK },
    { text: "도메인 이해", isBox: true, color: GRAY_LT, bgColor: "0A1828", borderColor: BORDER },
    { text: "+", isBox: false, color: GRAY_DK },
    { text: "단순한 도구", isBox: true, color: GRAY_LT, bgColor: "0A1828", borderColor: BORDER },
    { text: "=", isBox: false, color: GRAY_DK },
    { text: "전문가의\n반복 작업 제거", isBox: true, color: CYAN, bgColor: "002233", borderColor: "006688" },
  ];
  const eqWidths = [1.5, 0.3, 1.3, 0.3, 1.3, 0.3, 1.7];
  let ex = (10 - (eqWidths.reduce((a,b)=>a+b,0) + 0.15*(eqWidths.length-1))) / 2;
  eqItems.forEach((item, i) => {
    const w = eqWidths[i];
    if (item.isBox) {
      sl.addShape(pres.shapes.RECTANGLE, { x: ex, y: 2.55, w, h: 0.75, fill: fl(item.bgColor), line: ln(item.borderColor, 0.5) });
      sl.addText(item.text, { x: ex, y: 2.55, w, h: 0.75, fontSize: 11, fontFace: "Pretendard", color: item.color, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.3, margin: 0 });
    } else {
      sl.addText(item.text, { x: ex, y: 2.55, w, h: 0.75, fontSize: 18, fontFace: "Pretendard", color: item.color, bold: true, align: "center", valign: "middle", margin: 0 });
    }
    ex += w + 0.15;
  });
  sl.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: 3.6, w: 7, h: 0.45, fill: fl("1A0A0A"), line: ln("3A1010", 0.5) });
  sl.addText("⚠️  데이터가 없으면 LLM도, 파싱도, 어떤 자동화도 할 수 없다", { x: 1.5, y: 3.6, w: 7, h: 0.45, fontSize: 11, fontFace: "Pretendard", color: "8A5A5A", align: "center", valign: "middle", margin: 0 });
  sl.addShape(pres.shapes.LINE, { x: 3, y: 4.3, w: 4, h: 0, line: ln(BORDER, 0.5) });
  sl.addText("데이터 자산화 → 자동화 가능성의 전제조건", { x: 1, y: 4.4, w: 8, h: 0.25, fontSize: 10, fontFace: "Pretendard", color: GRAY_DK, align: "center" });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 7 — TWO ASSETS ("3년간" → "자산화된")
// ══════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: BG_MID };

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.83, fill: fl("0D1B3A"), line: ln(BORDER2, 0.5) });
  sl.addText("PART 2 · 자산화", { x: 0.58, y: 0.1, w: 2.5, h: 0.25, fontSize: 8, fontFace: "Pretendard", color: MINT, bold: true, charSpacing: 2 });
  sl.addText("두 자산의 연결", { x: 0.58, y: 0.36, w: 6, h: 0.35, fontSize: 18, fontFace: "Pretendard", color: WHITE, bold: true });

  const boxY = 1.35, boxH = 2.5;
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: boxY, w: 2.3, h: boxH, fill: fl("0A1525"), line: ln(AMBER, 1, 55) });
  sl.addText("🗄️", { x: 0.5, y: boxY+0.3, w: 2.3, h: 0.5, fontSize: 32, fontFace: "Pretendard", align: "center" });
  sl.addText("Spec DB", { x: 0.5, y: boxY+0.92, w: 2.3, h: 0.3, fontSize: 14, fontFace: "Pretendard", color: "FFD070", bold: true, align: "center" });
  sl.addText("자산화된\n레지스터 명세 데이터\n기대값, 필드 정의 포함", { x: 0.55, y: boxY+1.3, w: 2.2, h: 0.95, fontSize: 9, fontFace: "Pretendard", color: GRAY_DK, align: "center", lineSpacingMultiple: 1.4 });

  sl.addShape(pres.shapes.LINE, { x: 2.85, y: boxY+boxH/2, w: 0.6, h: 0, line: ln(AMBER, 1.5, 40) });
  sl.addText("연결", { x: 3.5, y: boxY+boxH/2-0.15, w: 0.7, h: 0.3, fontSize: 9, fontFace: "Pretendard", color: MINT, bold: true, align: "center" });
  sl.addShape(pres.shapes.LINE, { x: 4.25, y: boxY+boxH/2, w: 0.6, h: 0, line: ln(CYAN, 1.5, 40) });

  sl.addShape(pres.shapes.RECTANGLE, { x: 4.88, y: boxY, w: 2.3, h: boxH, fill: fl("0A1525"), line: ln(CYAN, 1, 55) });
  sl.addText("📄", { x: 4.88, y: boxY+0.3, w: 2.3, h: 0.5, fontSize: 32, fontFace: "Pretendard", align: "center" });
  sl.addText("Regbank 파일", { x: 4.88, y: boxY+0.92, w: 2.3, h: 0.3, fontSize: 14, fontFace: "Pretendard", color: "70D4FF", bold: true, align: "center" });
  sl.addText("디버그 세션마다\nHW 팀이 항상 생성하는\n실제 레지스터 덤프", { x: 4.93, y: boxY+1.3, w: 2.2, h: 0.95, fontSize: 9, fontFace: "Pretendard", color: GRAY_DK, align: "center", lineSpacingMultiple: 1.4 });

  sl.addShape(pres.shapes.LINE, { x: 7.22, y: boxY+boxH/2, w: 0.3, h: 0, line: ln(MINT, 1.5, 40) });
  sl.addText("▶", { x: 7.52, y: boxY+boxH/2-0.15, w: 0.25, h: 0.3, fontSize: 12, fontFace: "Pretendard", color: MINT, align: "center" });

  sl.addShape(pres.shapes.RECTANGLE, { x: 7.8, y: boxY, w: 1.8, h: boxH, fill: fl("001A1A"), line: ln(MINT, 1.5, 55) });
  sl.addText("⚡", { x: 7.8, y: boxY+0.35, w: 1.8, h: 0.5, fontSize: 32, fontFace: "Pretendard", align: "center" });
  sl.addText("자동화\n가능성", { x: 7.8, y: boxY+0.95, w: 1.8, h: 0.55, fontSize: 14, fontFace: "Pretendard", color: MINT, bold: true, align: "center", lineSpacingMultiple: 1.2 });
  sl.addText("두 자산을\n연결하면 수동\n분석을 자동화\n할 수 있다", { x: 7.85, y: boxY+1.6, w: 1.7, h: 1.0, fontSize: 9, fontFace: "Pretendard", color: "3A7A6A", align: "center", lineSpacingMultiple: 1.4 });

  sl.addText("HW 팀은 Regbank를  항상 생성하고 있었다  —  연결만 하면 됐다", { x: 0.5, y: 4.25, w: 9, h: 0.35, fontSize: 11, fontFace: "Pretendard", color: GRAY_DK, align: "center" });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 8 — ARCHITECTURE (Spec Data DB + Regbank Storage 명칭 반영)
// ══════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: BG_DARK };

  for (let i = 1; i < 9; i++) sl.addShape(pres.shapes.LINE, { x: 0, y: i*0.625, w: 10, h: 0, line: ln(CYAN, 0.3, 94) });

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.83, fill: fl("0D1B3A"), line: ln(BORDER2, 0.5) });
  sl.addText("PART 3 · 솔루션", { x: 0.58, y: 0.1, w: 2.5, h: 0.25, fontSize: 8, fontFace: "Pretendard", color: CYAN, bold: true, charSpacing: 2 });
  sl.addText("시스템 아키텍처", { x: 0.58, y: 0.36, w: 6, h: 0.35, fontSize: 18, fontFace: "Pretendard", color: WHITE, bold: true });

  sl.addText("INPUT",  { x: 0.5, y: 1.0, w: 2.2, h: 0.22, fontSize: 8, fontFace: "Pretendard", color: GRAY_DK, bold: true, charSpacing: 1.5, align: "center" });
  sl.addText("ENGINE", { x: 3.8, y: 1.0, w: 2.5, h: 0.22, fontSize: 8, fontFace: "Pretendard", color: GRAY_DK, bold: true, charSpacing: 1.5, align: "center" });
  sl.addText("OUTPUT", { x: 7.1, y: 1.0, w: 2.4, h: 0.22, fontSize: 8, fontFace: "Pretendard", color: GRAY_DK, bold: true, charSpacing: 1.5, align: "center" });

  sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.27, w: 2.5, h: 0.95, fill: fl(PANEL), line: ln(AMBER, 0.8, 55) });
  sl.addText("🗄️", { x: 0.55, y: 1.32, w: 0.5, h: 0.45, fontSize: 20, fontFace: "Pretendard", valign: "middle" });
  sl.addText("Spec Data DB", { x: 1.1, y: 1.34, w: 1.8, h: 0.25, fontSize: 11, fontFace: "Pretendard", color: "FFD070", bold: true });
  sl.addText("레지스터 명세 / 기대값", { x: 1.1, y: 1.6, w: 1.8, h: 0.55, fontSize: 8.5, fontFace: "Pretendard", color: GRAY_DK });

  sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.35, w: 2.5, h: 0.95, fill: fl(PANEL), line: ln(CYAN, 0.8, 55) });
  sl.addText("📁", { x: 0.55, y: 2.4, w: 0.5, h: 0.45, fontSize: 20, fontFace: "Pretendard", valign: "middle" });
  sl.addText("Regbank Storage", { x: 1.1, y: 2.42, w: 1.8, h: 0.25, fontSize: 11, fontFace: "Pretendard", color: "70D4FF", bold: true });
  sl.addText("이슈별 Regbank 보관", { x: 1.1, y: 2.68, w: 1.8, h: 0.55, fontSize: 8.5, fontFace: "Pretendard", color: GRAY_DK });

  sl.addShape(pres.shapes.LINE, { x: 3.04, y: 2.3, w: 0.72, h: 0, line: ln(CYAN, 1.5, 50) });
  sl.addText("파싱 ▶", { x: 3.05, y: 2.12, w: 0.7, h: 0.2, fontSize: 8, fontFace: "Pretendard", color: GRAY_DK, align: "center" });

  sl.addShape(pres.shapes.RECTANGLE, { x: 3.8, y: 1.27, w: 2.5, h: 2.1, fill: fl("0A1F3A"), line: ln(CYAN, 1.5, 55) });
  sl.addText("⚙️", { x: 3.8, y: 1.4, w: 2.5, h: 0.5, fontSize: 28, fontFace: "Pretendard", align: "center" });
  sl.addText("파싱 / 분석 엔진", { x: 3.85, y: 1.95, w: 2.4, h: 0.3, fontSize: 12, fontFace: "Pretendard", color: CYAN, bold: true, align: "center" });
  sl.addText("Regbank 파싱\nSpec과 대조 분석\n이상값 / 정상값 분류", { x: 3.85, y: 2.28, w: 2.4, h: 0.7, fontSize: 9.5, fontFace: "Pretendard", color: GRAY_DK, align: "center", lineSpacingMultiple: 1.4 });

  sl.addShape(pres.shapes.LINE, { x: 6.34, y: 2.3, w: 0.72, h: 0, line: ln(MINT, 1.5, 40) });
  sl.addText("결과 ▶", { x: 6.35, y: 2.12, w: 0.7, h: 0.2, fontSize: 8, fontFace: "Pretendard", color: GRAY_DK, align: "center" });

  sl.addShape(pres.shapes.RECTANGLE, { x: 7.1, y: 1.27, w: 2.4, h: 2.1, fill: fl("0A2A1A"), line: ln(MINT, 1.5, 55) });
  sl.addText("🖥️", { x: 7.1, y: 1.4, w: 2.4, h: 0.5, fontSize: 30, fontFace: "Pretendard", align: "center" });
  sl.addText("Web UI", { x: 7.15, y: 1.95, w: 2.3, h: 0.3, fontSize: 14, fontFace: "Pretendard", color: MINT, bold: true, align: "center" });
  sl.addText("분석 결과 시각화\n누구나 즉시 확인\n이상값 하이라이트", { x: 7.15, y: 2.28, w: 2.3, h: 0.7, fontSize: 9.5, fontFace: "Pretendard", color: GRAY_DK, align: "center", lineSpacingMultiple: 1.4 });

  sl.addText("기술 복잡도는 최소화 — 결과물의 가치에 집중", { x: 0.5, y: 4.5, w: 9, h: 0.3, fontSize: 11, fontFace: "Pretendard", color: GRAY_DK, align: "center" });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 9 — LIVE DEMO INTRO (동일)
// ══════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: BG_BLACK };

  sl.addShape(pres.shapes.RECTANGLE, { x: 3.5, y: 0.65, w: 3, h: 0.38, fill: fl(MINT, 91), line: ln(MINT, 0.8, 55) });
  sl.addShape(pres.shapes.OVAL, { x: 3.72, y: 0.77, w: 0.15, h: 0.15, fill: fl(MINT), line: ln(MINT, 0) });
  sl.addText("LIVE DEMO", { x: 3.92, y: 0.65, w: 2.5, h: 0.38, fontSize: 12, fontFace: "Pretendard", color: MINT, bold: true, charSpacing: 2, valign: "middle" });
  sl.addText("직접 보여드리겠습니다", { x: 1, y: 1.2, w: 8, h: 0.8, fontSize: 36, fontFace: "Pretendard", color: WHITE, bold: true, align: "center" });
  sl.addText("Regbank 파일  →  분석 엔진  →  Web UI 결과", { x: 1, y: 2.1, w: 8, h: 0.35, fontSize: 13, fontFace: "Pretendard", color: GRAY_DK, align: "center" });

  const demoSteps = ["Regbank 파일\n업로드", "처리 과정\n시연", "Web UI 결과\n확인", "수동 대비\n시간 비교"];
  const demoStartX = (10 - (demoSteps.length * 2.0 + (demoSteps.length-1) * 0.3)) / 2;
  demoSteps.forEach((step, i) => {
    const x = demoStartX + i * 2.3;
    sl.addShape(pres.shapes.OVAL, { x: x+0.75, y: 2.72, w: 0.5, h: 0.5, fill: fl("0A1828"), line: ln(MINT, 0.8, 55) });
    sl.addText(String(i+1), { x: x+0.75, y: 2.72, w: 0.5, h: 0.5, fontSize: 13, fontFace: "Pretendard", color: MINT, bold: true, align: "center", valign: "middle", margin: 0 });
    sl.addText(step, { x: x+0.3, y: 3.32, w: 1.4, h: 0.6, fontSize: 10, fontFace: "Pretendard", color: GRAY_MD, align: "center", lineSpacingMultiple: 1.3 });
    if (i < demoSteps.length - 1) sl.addText("›", { x: x+1.95, y: 2.72, w: 0.4, h: 0.5, fontSize: 16, fontFace: "Pretendard", color: "1A4A3A", align: "center", valign: "middle", margin: 0 });
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 10 — DEMO RESULT (3번째 하이라이트 변경: "담당자 없이 누구나")
// ══════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: BG_MID };

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.75, fill: fl("0D1B3A"), line: ln(BORDER2, 0.5) });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.42, y: 0.18, w: 1.05, h: 0.38, fill: fl(MINT, 91), line: ln(MINT, 0.5, 67) });
  sl.addShape(pres.shapes.OVAL, { x: 0.52, y: 0.30, w: 0.12, h: 0.12, fill: fl(MINT), line: ln(MINT, 0) });
  sl.addText("LIVE DEMO", { x: 0.67, y: 0.19, w: 0.78, h: 0.35, fontSize: 8, fontFace: "Pretendard", color: MINT, bold: true, charSpacing: 1, valign: "middle" });
  sl.addText("분석 결과 하이라이트", { x: 1.6, y: 0.2, w: 6, h: 0.35, fontSize: 17, fontFace: "Pretendard", color: WHITE, bold: true, valign: "middle" });

  const uiX = 0.42, uiY = 0.88, uiW = 6.3, uiH = 4.45;
  sl.addShape(pres.shapes.RECTANGLE, { x: uiX, y: uiY, w: uiW, h: uiH, fill: fl("0A1220"), line: ln(BORDER, 0.8) });
  sl.addShape(pres.shapes.RECTANGLE, { x: uiX, y: uiY, w: uiW, h: 0.38, fill: fl("0D1A30"), line: ln(BORDER2, 0.5) });
  sl.addShape(pres.shapes.OVAL, { x: uiX+0.15, y: uiY+0.11, w: 0.14, h: 0.14, fill: fl("FF5F56"), line: ln("FF5F56", 0) });
  sl.addShape(pres.shapes.OVAL, { x: uiX+0.35, y: uiY+0.11, w: 0.14, h: 0.14, fill: fl("FEBC2E"), line: ln("FEBC2E", 0) });
  sl.addShape(pres.shapes.OVAL, { x: uiX+0.55, y: uiY+0.11, w: 0.14, h: 0.14, fill: fl("28C840"), line: ln("28C840", 0) });
  sl.addText("HW Debugging Analysis — Regbank Viewer", { x: uiX+0.8, y: uiY+0.06, w: 5, h: 0.26, fontSize: 9, fontFace: "Pretendard", color: GRAY_DK, valign: "middle" });

  const rows = [
    { name: "SYS_CTRL_REG",   val: "0x0000_0001  →  Expected: 0x0000_0001", status: "NORMAL", statusColor: MINT,  borderColor: MINT  },
    { name: "PWR_STATE_REG",  val: "0x0000_00FF  →  Expected: 0x0000_0000", status: "ERROR",  statusColor: RED,   borderColor: RED   },
    { name: "CLK_DOMAIN_REG", val: "0x0000_0003  →  Expected: 0x0000_0001", status: "WARN",   statusColor: AMBER, borderColor: AMBER },
    { name: "INT_STATUS_REG", val: "0x0000_0000  →  Expected: 0x0000_0000", status: "NORMAL", statusColor: MINT,  borderColor: MINT  },
    { name: "MEM_CTRL_REG",   val: "0x0000_DEAD  →  Expected: 0x0000_0001", status: "ERROR",  statusColor: RED,   borderColor: RED   },
  ];
  rows.forEach((row, i) => {
    const ry = uiY + 0.52 + i * 0.73;
    sl.addShape(pres.shapes.RECTANGLE, { x: uiX+0.1, y: ry, w: uiW-0.2, h: 0.62, fill: fl("0D1A2A"), line: ln(BORDER2, 0.3) });
    sl.addShape(pres.shapes.RECTANGLE, { x: uiX+0.1, y: ry, w: 0.055, h: 0.62, fill: fl(row.borderColor), line: ln(row.borderColor, 0) });
    sl.addText(row.name, { x: uiX+0.28, y: ry+0.08, w: 2.2, h: 0.22, fontSize: 9.5, fontFace: "Courier New", color: GRAY_LT, bold: true });
    sl.addText(row.val,  { x: uiX+0.28, y: ry+0.32, w: 4.2, h: 0.18, fontSize: 8.5, fontFace: "Courier New", color: GRAY_DK });
    sl.addShape(pres.shapes.RECTANGLE, { x: uiX+uiW-1.1, y: ry+0.16, w: 0.88, h: 0.28, fill: fl("0D1A2A"), line: ln(row.statusColor, 0.5) });
    sl.addText(row.status, { x: uiX+uiW-1.1, y: ry+0.16, w: 0.88, h: 0.28, fontSize: 8, fontFace: "Pretendard", color: row.statusColor, bold: true, align: "center", valign: "middle", margin: 0 });
  });

  // 3번째 하이라이트 변경
  const hlItems = [
    { icon: "⚡", label: "즉각적 결과",     desc: "파일 업로드 즉시 분석 완료. 수동 검토 없음" },
    { icon: "🎯", label: "이상값 자동 감지", desc: "Spec 기대값과 비교해 ERROR / WARN 자동 분류" },
    { icon: "👥", label: "담당자 없이 누구나", desc: "Web UI에서 바로 확인 — 할당 대기 없음" },
  ];
  hlItems.forEach((h, i) => {
    const y = 0.88 + i * 1.5;
    sl.addShape(pres.shapes.RECTANGLE, { x: 7.0, y, w: 2.58, h: 1.3, fill: fl(PANEL), line: ln(BORDER, 0.5) });
    sl.addText(h.icon, { x: 7.1, y: y+0.1, w: 0.4, h: 0.3, fontSize: 18, fontFace: "Pretendard" });
    sl.addText(h.label, { x: 7.1, y: y+0.42, w: 2.3, h: 0.25, fontSize: 10, fontFace: "Pretendard", color: CYAN, bold: true });
    sl.addText(h.desc,  { x: 7.1, y: y+0.68, w: 2.3, h: 0.5,  fontSize: 9,  fontFace: "Pretendard", color: GRAY_DK, lineSpacingMultiple: 1.4 });
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 11 — COLLABORATION ("축적된 데이터" — "3년간" 삭제)
// ══════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: BG_MID };

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.83, fill: fl("0D1B3A"), line: ln(BORDER2, 0.5) });
  sl.addText("PART 4 · 여정", { x: 0.58, y: 0.1, w: 2, h: 0.25, fontSize: 8, fontFace: "Pretendard", color: PURPLE, bold: true, charSpacing: 2 });
  sl.addText("SW ↔ HW 협업 여정", { x: 0.58, y: 0.36, w: 7, h: 0.35, fontSize: 18, fontFace: "Pretendard", color: WHITE, bold: true });

  const collabs = [
    { icon: "🔧", title: "HW 엔지니어의 역할", desc: "도메인 문제 정의 — 어떤 레지스터가 중요한가?\n어떤 기대값이어야 하는가? 디버깅 프로세스의 맥락", borderColor: AMBER, titleColor: AMBER },
    { icon: "💻", title: "SW 담당자의 역할",   desc: "기술 구현 — HW 디버깅 도메인 지식 없이 시작.\n도메인은 HW에게, 구현은 SW가 담당",              borderColor: CYAN,  titleColor: CYAN  },
    { icon: "🤝", title: "핵심: 협업이 없었다면 불가능", desc: "도메인 지식 없이 기술만으로는\n올바른 자동화를 만들 수 없다",           borderColor: PURPLE, titleColor: PURPLE },
  ];
  collabs.forEach((c, i) => {
    const y = 1.0 + i * 1.45;
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.42, y, w: 3.8, h: 1.25, fill: fl(PANEL2), line: ln(c.borderColor, 0.8, 55) });
    sl.addText(`${c.icon}  ${c.title}`, { x: 0.58, y: y+0.1,  w: 3.5, h: 0.28, fontSize: 10, fontFace: "Pretendard", color: c.titleColor, bold: true });
    sl.addText(c.desc,                  { x: 0.58, y: y+0.42, w: 3.5, h: 0.73, fontSize: 9.5, fontFace: "Pretendard", color: GRAY_DK, lineSpacingMultiple: 1.4 });
  });

  const challenges = [
    { icon: "🧩", title: "디버깅 프로세스 도메인 이해", desc: "SW 담당자가 HW 디버깅 프로세스를 이해하는 데 예상보다 시간이 걸렸다" },
    { icon: "🔄", title: "Spec 데이터 정제 및 구조화",  desc: "축적된 데이터도 그대로 쓸 수 없었다. 정제와 구조화 작업 필요" },  // "3년간" 삭제
    { icon: "⚡", title: "엣지 케이스 처리",             desc: "예외적인 레지스터 상태, 특수 케이스들을 처리하는 과정이 까다로웠다" },
  ];
  challenges.forEach((c, i) => {
    const y = 1.0 + i * 1.0;
    sl.addShape(pres.shapes.RECTANGLE, { x: 4.52, y, w: 0.06, h: 0.82, fill: fl(PURPLE, 67), line: ln(PURPLE, 0, 67) });
    sl.addShape(pres.shapes.RECTANGLE, { x: 4.6,  y, w: 5.0,  h: 0.82, fill: fl("0C1828"), line: ln(BORDER, 0.4) });
    sl.addText(c.icon,  { x: 4.72, y: y+0.08, w: 0.35, h: 0.38, fontSize: 16, fontFace: "Pretendard" });
    sl.addText(c.title, { x: 5.12, y: y+0.06, w: 4.35, h: 0.25, fontSize: 11, fontFace: "Pretendard", color: "C0A8E8", bold: true });
    sl.addText(c.desc,  { x: 5.12, y: y+0.34, w: 4.35, h: 0.4,  fontSize: 9.5, fontFace: "Pretendard", color: GRAY_DK, lineSpacingMultiple: 1.3 });
  });
  sl.addShape(pres.shapes.RECTANGLE, { x: 4.6, y: 4.1, w: 5.0, h: 0.95, fill: fl(PANEL), line: ln(BORDER, 0.5) });
  sl.addText("💡  기술적 단순함은 의도적 선택 — 복잡하지 않게 만든 것이 강점.\n도메인 문제에 집중하기 위해 기술 복잡도를 의도적으로 낮게 유지했다.", { x: 4.75, y: 4.17, w: 4.7, h: 0.8, fontSize: 10, fontFace: "Pretendard", color: "3A6A8A", lineSpacingMultiple: 1.5 });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 12 — EVOLUTION (동일)
// ══════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: BG_MID };

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.83, fill: fl("0D1B3A"), line: ln(BORDER2, 0.5) });
  sl.addText("PART 4 · 여정", { x: 0.58, y: 0.1, w: 2, h: 0.25, fontSize: 8, fontFace: "Pretendard", color: MINT, bold: true, charSpacing: 2 });
  sl.addText("현재 한계 + 진화 가능성", { x: 0.58, y: 0.36, w: 7, h: 0.35, fontSize: 18, fontFace: "Pretendard", color: WHITE, bold: true });

  sl.addShape(pres.shapes.RECTANGLE, { x: 0.42, y: 1.0, w: 4.0, h: 4.0, fill: fl(PANEL2), line: ln(CYAN, 1, 67) });
  sl.addText("CURRENT", { x: 0.58, y: 1.12, w: 3.7, h: 0.22, fontSize: 9, fontFace: "Pretendard", color: CYAN, bold: true, charSpacing: 2 });
  sl.addText("파싱 / 분석\n자동화 수준", { x: 0.58, y: 1.38, w: 3.7, h: 0.65, fontSize: 16, fontFace: "Pretendard", color: "D8EAF8", bold: true, lineSpacingMultiple: 1.2 });
  ["Regbank 파싱 자동화", "Spec 기대값과 대조", "이상값 자동 분류 및 표시", "Web UI로 누구나 확인", "파일럿 단계 — 측정 데이터 축적 중"].forEach((item, i) => {
    const y = 2.15 + i * 0.52;
    sl.addShape(pres.shapes.OVAL, { x: 0.62, y: y+0.06, w: 0.12, h: 0.12, fill: fl(i<4 ? CYAN : "1A3A5F"), line: ln(i<4 ? CYAN : "1A3A5F", 0) });
    sl.addText(item, { x: 0.85, y, w: 3.4, h: 0.32, fontSize: 10.5, fontFace: "Pretendard", color: i<4 ? GRAY_MD : "2A4A6A" });
  });

  sl.addShape(pres.shapes.LINE, { x: 4.85, y: 2.2, w: 0, h: 1.5, line: ln(MINT, 1, 55) });
  sl.addText("도메인\n이해 ↑", { x: 4.55, y: 2.35, w: 0.6, h: 0.6, fontSize: 8.5, fontFace: "Pretendard", color: "2A5A4A", align: "center", lineSpacingMultiple: 1.3 });
  sl.addText("▼", { x: 4.65, y: 3.8, w: 0.4, h: 0.3, fontSize: 14, fontFace: "Pretendard", color: MINT, align: "center" });

  sl.addShape(pres.shapes.RECTANGLE, { x: 5.42, y: 1.0, w: 4.16, h: 4.0, fill: fl("0A1A1A"), line: ln(MINT, 1.2, 55) });
  sl.addText("FUTURE", { x: 5.58, y: 1.12, w: 3.8, h: 0.22, fontSize: 9, fontFace: "Pretendard", color: MINT, bold: true, charSpacing: 2 });
  sl.addText("더 진화된\nAgent", { x: 5.58, y: 1.38, w: 3.8, h: 0.65, fontSize: 16, fontFace: "Pretendard", color: "D8EAF8", bold: true, lineSpacingMultiple: 1.2 });
  ["디버깅 패턴 학습 및 추론", "문제 원인 자동 추론", "유사 과거 케이스 매칭", "진단 리포트 자동 생성", "도메인 지식이 쌓일수록 가능성 확대"].forEach((item, i) => {
    const y = 2.15 + i * 0.52;
    sl.addShape(pres.shapes.OVAL, { x: 5.62, y: y+0.06, w: 0.12, h: 0.12, fill: fl(MINT), line: ln(MINT, 0) });
    sl.addText(item, { x: 5.85, y, w: 3.6, h: 0.32, fontSize: 10.5, fontFace: "Pretendard", color: "4A8A7A" });
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 13 — FRAMEWORK (동일)
// ══════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: BG_DARK };

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.06, h: 5.625, fill: fl(CYAN), line: ln(CYAN, 0) });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.06, y: 0, w: 9.94, h: 0.83, fill: fl("0D1B3A"), line: ln(BORDER2, 0.5) });
  sl.addText("PART 5 · 프레임워크", { x: 0.7, y: 0.1, w: 2.8, h: 0.25, fontSize: 8, fontFace: "Pretendard", color: CYAN, bold: true, charSpacing: 2 });
  sl.addText("당신 팀에서도 찾을 수 있다", { x: 0.7, y: 0.36, w: 7, h: 0.35, fontSize: 18, fontFace: "Pretendard", color: WHITE, bold: true });
  sl.addText("자동화 기회를 발굴하는 3가지 질문 — 지금 5분, 자기 팀에 적용해보세요", { x: 0.7, y: 0.95, w: 9, h: 0.28, fontSize: 11, fontFace: "Pretendard", color: GRAY_DK });

  const questions = [
    { num: "1", numColor: CYAN,  numBg: "001A2A", numBdr: "006688", q: "우리 팀에 반복적으로 수행하는 전문가 작업이 있는가?",            hint: "시니어만 할 수 있고, 자주 반복되며, 시간이 많이 걸리는 작업을 찾는다" },
    { num: "2", numColor: AMBER, numBg: "1A1000", numBdr: "886600", q: "그 작업에 사용되는 데이터가 이미 어딘가에 구조화되어 있는가?",    hint: "DB, 파일, 시스템 어딘가에 이미 존재하는 데이터를 발굴한다" },
    { num: "3", numColor: MINT,  numBg: "001A14", numBdr: "006650", q: "두 데이터를 연결하면 자동화할 수 있는가?",                       hint: "데이터 연결 = 자동화 가능성. AI 없이도 단순한 도구로 시작할 수 있다" },
  ];
  questions.forEach((q, i) => {
    const y = 1.35 + i * 1.3;
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.7, y, w: 8.9, h: 1.1, fill: fl(PANEL), line: ln(BORDER, 0.8) });
    sl.addShape(pres.shapes.OVAL, { x: 0.88, y: y+0.3, w: 0.5, h: 0.5, fill: fl(q.numBg), line: ln(q.numBdr, 0.8) });
    sl.addText(q.num, { x: 0.88, y: y+0.3, w: 0.5, h: 0.5, fontSize: 16, fontFace: "Pretendard", color: q.numColor, bold: true, align: "center", valign: "middle", margin: 0 });
    sl.addText(q.q,    { x: 1.55, y: y+0.1,  w: 7.8, h: 0.38, fontSize: 13, fontFace: "Pretendard", color: "D8EAF8", bold: true });
    sl.addText(q.hint, { x: 1.55, y: y+0.58, w: 7.8, h: 0.42, fontSize: 10, fontFace: "Pretendard", color: GRAY_DK, lineSpacingMultiple: 1.3 });
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 14 — CLOSING (동일)
// ══════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: BG_DARK };

  for (let i = 1; i < 9; i++) sl.addShape(pres.shapes.LINE, { x: 0, y: i*0.625, w: 10, h: 0, line: ln(CYAN, 0.3, 91) });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.04, fill: fl(CYAN), line: ln(CYAN, 0) });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.585, w: 10, h: 0.04, fill: fl(MINT), line: ln(MINT, 0) });

  sl.addText("CLOSING", { x: 0.7, y: 0.55, w: 6, h: 0.22, fontSize: 8, fontFace: "Pretendard", color: CYAN, bold: true, charSpacing: 3 });
  sl.addText([
    { text: "함께\n", options: { color: WHITE } },
    { text: "고민", options: { color: CYAN } },
    { text: "해요", options: { color: WHITE } },
  ], { x: 0.7, y: 0.82, w: 5.5, h: 1.3, fontSize: 40, fontFace: "Pretendard", bold: true, lineSpacingMultiple: 1.15 });

  sl.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 2.3, w: 0.06, h: 1.6, fill: fl(CYAN, 80), line: ln(CYAN, 0, 80) });
  sl.addText('"자동화와 Agent를 더 발전시키고 싶은데,\n어떤 Framework로 설계하면 좋을지 아직 고민 중입니다.\n같은 고민을 하고 계신 분이 있다면\n발표 후 함께 이야기 나눠보고 싶습니다."', { x: 0.88, y: 2.3, w: 5.8, h: 1.6, fontSize: 11, fontFace: "Pretendard", color: GRAY_MD, lineSpacingMultiple: 1.6 });

  const pills = ["문제 발굴", "데이터 자산화", "두 자산 연결", "SW ↔ HW 협업", "단순한 도구"];
  let px = 0.7;
  pills.forEach((p) => {
    const w = p.length * 0.115 + 0.6;
    sl.addShape(pres.shapes.RECTANGLE, { x: px, y: 4.2, w, h: 0.33, fill: fl("0A1828"), line: ln(BORDER, 0.5) });
    sl.addText(p, { x: px, y: 4.2, w, h: 0.33, fontSize: 9, fontFace: "Pretendard", color: GRAY_DK, align: "center", valign: "middle", margin: 0 });
    px += w + 0.12;
  });

  sl.addShape(pres.shapes.RECTANGLE, { x: 7.0, y: 0.7, w: 2.6, h: 2.5, fill: fl("001818"), line: ln(MINT, 1, 67) });
  sl.addText("🤝", { x: 7.0, y: 0.85, w: 2.6, h: 0.6, fontSize: 30, fontFace: "Pretendard", align: "center" });
  sl.addText("발표 후\n함께 이야기해요", { x: 7.1, y: 1.5, w: 2.4, h: 0.65, fontSize: 13, fontFace: "Pretendard", color: MINT, bold: true, align: "center", lineSpacingMultiple: 1.25 });
  sl.addText("Agent Framework 설계,\n자동화 기회 발굴,\n비슷한 경험이 있으신 분 환영합니다", { x: 7.1, y: 2.2, w: 2.4, h: 0.9, fontSize: 9, fontFace: "Pretendard", color: "3A7A6A", align: "center", lineSpacingMultiple: 1.4 });

  sl.addShape(pres.shapes.RECTANGLE, { x: 7.0, y: 3.4, w: 2.6, h: 1.2, fill: fl(PANEL2), line: ln(BORDER, 0.8) });
  sl.addText("💬", { x: 7.0, y: 3.52, w: 2.6, h: 0.4, fontSize: 22, fontFace: "Pretendard", align: "center" });
  sl.addText("Q & A", { x: 7.0, y: 3.95, w: 2.6, h: 0.5, fontSize: 18, fontFace: "Pretendard", color: "5A8AB0", bold: true, align: "center" });
}

// ── Write ──────────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "plan2.pptx" })
  .then(() => console.log("✅ plan2.pptx created successfully"))
  .catch((err) => console.error("❌ Error:", err));
