// Builds AI-Staff-Room.pptx — 17-slide, 30-min deck for vocational HS leaders
// Run:  node scripts/build-ai-staff-room-deck.js
//
// Theme: "Teal Trust × Warm Accent"
//   Navy     #1E3A5F  (titles, dark sections)
//   Teal     #2E8A9C  (primary accent)
//   Amber    #E9A24B  (highlight, callouts)
//   Sand     #F4EEE3  (warm light bg)
//   Ink      #1F2937  (body text)
//   Mute     #6B7280  (secondary text)

const pptxgen = require("pptxgenjs");
const path = require("path");

const C = {
  navy:  "1E3A5F",
  teal:  "2E8A9C",
  amber: "E9A24B",
  sand:  "F4EEE3",
  ink:   "1F2937",
  mute:  "6B7280",
  white: "FFFFFF",
  line:  "D9D2C2",
};

const FONT_KO = "맑은 고딕";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";   // 13.3" × 7.5"
pres.author = "Justin Yoo";
pres.company = "Microsoft";
pres.title = "AI 교무실 — 직업계고 멀티에이전트 활용";

const W = 13.3, H = 7.5;

// ---------- helpers ----------
function bg(slide, color) {
  slide.background = { color };
}

function titleBar(slide, text, sub = null) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 1.1, fill: { color: C.navy }, line: { color: C.navy },
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 1.1, w: W, h: 0.08, fill: { color: C.amber }, line: { color: C.amber },
  });
  slide.addText(text, {
    x: 0.6, y: 0.18, w: W - 1.2, h: sub ? 0.55 : 0.75,
    fontSize: 28, bold: true, color: C.white, fontFace: FONT_KO, valign: "middle", margin: 0,
  });
  if (sub) {
    slide.addText(sub, {
      x: 0.6, y: 0.7, w: W - 1.2, h: 0.4,
      fontSize: 14, color: C.sand, fontFace: FONT_KO, valign: "top", margin: 0,
    });
  }
}

function footer(slide, idx, total) {
  slide.addText(`AI 교무실 · 직업계고 멀티에이전트 활용`, {
    x: 0.5, y: H - 0.4, w: 8, h: 0.3,
    fontSize: 10, color: C.mute, fontFace: FONT_KO, margin: 0,
  });
  slide.addText(`${idx} / ${total}`, {
    x: W - 1.3, y: H - 0.4, w: 0.8, h: 0.3,
    fontSize: 10, color: C.mute, fontFace: FONT_KO, align: "right", margin: 0,
  });
}

function card(slide, x, y, w, h, opts = {}) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    fill: { color: opts.fill || C.white },
    line: { color: opts.border || C.line, width: 1 },
    rectRadius: 0.12,
    shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.08 },
  });
}

function chip(slide, x, y, w, h, text, fill, color) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.08,
  });
  slide.addText(text, {
    x, y, w, h, fontSize: 12, bold: true, color, fontFace: FONT_KO,
    align: "center", valign: "middle", margin: 0,
  });
}

function arrow(slide, x, y, w) {
  slide.addShape(pres.shapes.RIGHT_TRIANGLE, {
    x: x + w - 0.18, y: y + 0.12, w: 0.18, h: 0.26,
    fill: { color: C.amber }, line: { color: C.amber }, rotate: 90,
  });
  slide.addShape(pres.shapes.LINE, {
    x, y: y + 0.25, w: w - 0.15, h: 0,
    line: { color: C.amber, width: 3 },
  });
}

const TOTAL = 17;
let idx = 0;
const nextIdx = () => ++idx;

// =========================================================
// Slide 1 — Cover
// =========================================================
{
  const s = pres.addSlide();
  bg(s, C.navy);
  // accent bar
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: H - 0.4, w: W, h: 0.4, fill: { color: C.amber }, line: { color: C.amber } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.35, h: H, fill: { color: C.teal }, line: { color: C.teal } });

  s.addText("AI 교무실", {
    x: 1.0, y: 1.8, w: 11, h: 1.5,
    fontSize: 80, bold: true, color: C.white, fontFace: FONT_KO, margin: 0,
  });
  s.addText("우리 학교에 8명의 AI 동료가 출근한다면", {
    x: 1.0, y: 3.3, w: 11, h: 0.8,
    fontSize: 28, color: C.sand, fontFace: FONT_KO, margin: 0,
  });
  s.addText("직업계고 현장에서 멀티에이전트 활용하기", {
    x: 1.0, y: 4.2, w: 11, h: 0.6,
    fontSize: 18, color: C.amber, fontFace: FONT_KO, italic: true, margin: 0,
  });

  s.addText("발표 30분 · 데모 4건 · 도입 로드맵 포함", {
    x: 1.0, y: H - 1.2, w: 11, h: 0.4,
    fontSize: 14, color: C.sand, fontFace: FONT_KO, margin: 0,
  });
}
nextIdx();

// =========================================================
// Slide 2 — Hook: 한 학기 작성 문서
// =========================================================
{
  const s = pres.addSlide();
  bg(s, C.white);
  titleBar(s, "선생님 한 분이 한 학기에 작성하는 문서", "직업계고 담임·전공 교사 기준 평균치");

  const docs = [
    "NCS 기반 실습 수업안",
    "안전 점검 체크리스트",
    "현장실습 표준협약·일지",
    "수행평가 루브릭",
    "행동특성 및 종합의견",
    "취업 추천서·자기소개서 코칭",
    "진로 상담 기록",
    "학부모 안내·민원 응대",
    "기능경기·자격증 지도 자료",
  ];

  // Big number left
  s.addText("9", {
    x: 0.8, y: 2.0, w: 4.5, h: 3.2,
    fontSize: 220, bold: true, color: C.amber, fontFace: FONT_KO,
    align: "center", valign: "middle", margin: 0,
  });
  s.addText("종", {
    x: 4.5, y: 2.6, w: 1.5, h: 1.5,
    fontSize: 56, bold: true, color: C.navy, fontFace: FONT_KO,
    align: "left", valign: "middle", margin: 0,
  });
  s.addText("한 사람이 한 학기에", {
    x: 0.8, y: 5.3, w: 4.5, h: 0.5,
    fontSize: 18, color: C.ink, fontFace: FONT_KO, align: "center", margin: 0,
  });

  // Right column: doc list as chips
  const startX = 6.5, startY = 1.7, chipW = 3.0, chipH = 0.55, gap = 0.15;
  docs.forEach((d, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    chip(s, startX + col * (chipW + 0.2), startY + row * (chipH + gap),
      chipW, chipH, d, C.sand, C.ink);
  });

  s.addText("➤ 이 모든 것을 “한 명의 슈퍼맨 교사”가 다 해내고 있습니다.", {
    x: 0.8, y: H - 1.0, w: W - 1.6, h: 0.4,
    fontSize: 16, italic: true, color: C.teal, fontFace: FONT_KO, margin: 0,
  });

  footer(s, idx + 1, TOTAL);
}
nextIdx();

// =========================================================
// Slide 3 — Anti-pattern: 단일 AI의 한계
// =========================================================
{
  const s = pres.addSlide();
  bg(s, C.white);
  titleBar(s, "AI에 “그냥 한 번 묻기”, 왜 부족한가",
    "ChatGPT 한 번 호출 = 책임소재 없는 1인 답변");

  const items = [
    { t: "일관성 부족", d: "같은 질문에 매번 다른 답.\n반·학년·학교 간 표준화 어려움." },
    { t: "안전·법규 검토 누락", d: "현장실습 표준협약·산업안전·\n개인정보 보호 시각이 비어 있음." },
    { t: "NCS 매핑 미흡", d: "능력단위·수행준거와 자동 연결되지 않음.\n검토 단계 없이 그대로 사용 위험." },
    { t: "책임소재 불명확", d: "출처·근거 트랙이 한 줄.\n사후 감사·이의 제기 시 추적 불가." },
  ];

  const x0 = 0.7, y0 = 1.6, w = (W - 1.4 - 0.6) / 2, h = 2.5, gap = 0.3;
  items.forEach((it, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = x0 + col * (w + gap);
    const y = y0 + row * (h + gap);
    card(s, x, y, w, h);
    // number badge
    s.addShape(pres.shapes.OVAL, {
      x: x + 0.25, y: y + 0.25, w: 0.55, h: 0.55,
      fill: { color: C.amber }, line: { color: C.amber },
    });
    s.addText(String(i + 1), {
      x: x + 0.25, y: y + 0.25, w: 0.55, h: 0.55,
      fontSize: 20, bold: true, color: C.white, fontFace: FONT_KO,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(it.t, {
      x: x + 0.95, y: y + 0.25, w: w - 1.1, h: 0.55,
      fontSize: 20, bold: true, color: C.navy, fontFace: FONT_KO, valign: "middle", margin: 0,
    });
    s.addText(it.d, {
      x: x + 0.25, y: y + 0.95, w: w - 0.5, h: h - 1.1,
      fontSize: 14, color: C.ink, fontFace: FONT_KO, margin: 0,
    });
  });

  footer(s, idx + 1, TOTAL);
}
nextIdx();

// =========================================================
// Slide 4 — Frame: 슈퍼맨 AI vs AI 교무실
// =========================================================
{
  const s = pres.addSlide();
  bg(s, C.white);
  titleBar(s, "슈퍼맨 AI ❌  →  AI 교무실 ✅",
    "한 명에게 다 시키지 말고, 역할이 나뉜 팀을 두자");

  // Left card — 슈퍼맨 AI
  const lx = 0.7, ly = 1.6, lw = 5.6, lh = 4.8;
  card(s, lx, ly, lw, lh, { fill: "F7E8E0", border: "C97A5A" });
  s.addText("슈퍼맨 AI 한 명", {
    x: lx + 0.3, y: ly + 0.3, w: lw - 0.6, h: 0.5,
    fontSize: 22, bold: true, color: "8E3A1E", fontFace: FONT_KO, margin: 0,
  });
  s.addText([
    { text: "한 챗봇이 모든 걸 받아 처리", options: { bullet: true, breakLine: true } },
    { text: "검토·승인 단계 없음", options: { bullet: true, breakLine: true } },
    { text: "출처·책임 추적 불가", options: { bullet: true, breakLine: true } },
    { text: "프롬프트 의존도가 너무 높음", options: { bullet: true, breakLine: true } },
    { text: "교사 1인 = 위험·노동의 1인 부담", options: { bullet: true } },
  ], {
    x: lx + 0.3, y: ly + 1.0, w: lw - 0.6, h: lh - 1.2,
    fontSize: 16, color: C.ink, fontFace: FONT_KO, paraSpaceAfter: 6,
  });

  // Right card — AI 교무실
  const rx = lx + lw + 0.4, ry = ly, rw = lw, rh = lh;
  card(s, rx, ry, rw, rh, { fill: "E3F1F1", border: C.teal });
  s.addText("AI 교무실 (역할 분담된 팀)", {
    x: rx + 0.3, y: ry + 0.3, w: rw - 0.6, h: 0.5,
    fontSize: 22, bold: true, color: C.teal, fontFace: FONT_KO, margin: 0,
  });
  s.addText([
    { text: "능력단위처럼 역할이 쪼개진 AI 동료들", options: { bullet: true, breakLine: true } },
    { text: "단계별 검토·종합 → 자연스러운 감사선", options: { bullet: true, breakLine: true } },
    { text: "출처가 트랙별로 분리 → 사후 추적 가능", options: { bullet: true, breakLine: true } },
    { text: "프롬프트·데이터셋을 시도 단위로 표준화", options: { bullet: true, breakLine: true } },
    { text: "교사는 “감독·결정”에 집중", options: { bullet: true } },
  ], {
    x: rx + 0.3, y: ry + 1.0, w: rw - 0.6, h: rh - 1.2,
    fontSize: 16, color: C.ink, fontFace: FONT_KO, paraSpaceAfter: 6,
  });

  s.addText("교무실에 부장교사·담임·전공교사·산학협력부장이 따로 있는 이유 — AI도 똑같습니다.", {
    x: 0.7, y: H - 0.9, w: W - 1.4, h: 0.4,
    fontSize: 14, italic: true, color: C.mute, fontFace: FONT_KO, margin: 0, align: "center",
  });

  footer(s, idx + 1, TOTAL);
}
nextIdx();

// =========================================================
// Slide 5 — Agenda: 4 patterns mapped to school scenes
// =========================================================
{
  const s = pres.addSlide();
  bg(s, C.white);
  titleBar(s, "오늘 만나 볼 4명의 AI 교무실 동료들",
    "각 패턴 = 학교 일과의 익숙한 장면");

  const scenes = [
    { p: "Sequential", k: "결재 라인", d: "전공 실습 수업·평가 자료 제작", c: C.teal },
    { p: "Concurrent", k: "학년부 회의", d: "행동특성 및 종합의견 작성", c: C.amber },
    { p: "Handoff", k: "민원 전화 응대", d: "학생·학부모 통합 안내 챗봇", c: "7BA7BC" },
    { p: "Group Chat", k: "진로 협의회", d: "졸업 후 진로 결정 시뮬레이션", c: "C97A5A" },
  ];

  const x0 = 0.6, y0 = 1.6, w = (W - 1.2 - 0.6) / 4, h = 4.8, gap = 0.2;
  scenes.forEach((it, i) => {
    const x = x0 + i * (w + gap);
    card(s, x, y0, w, h);
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: y0, w, h: 0.6, fill: { color: it.c }, line: { color: it.c },
    });
    s.addText(it.p, {
      x, y: y0, w, h: 0.6,
      fontSize: 16, bold: true, color: C.white, fontFace: FONT_KO,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(it.k, {
      x: x + 0.2, y: y0 + 1.0, w: w - 0.4, h: 0.8,
      fontSize: 24, bold: true, color: C.navy, fontFace: FONT_KO,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("같은 장면", {
      x: x + 0.2, y: y0 + 1.8, w: w - 0.4, h: 0.3,
      fontSize: 11, color: C.mute, fontFace: FONT_KO, align: "center", margin: 0,
    });
    s.addShape(pres.shapes.LINE, {
      x: x + 0.5, y: y0 + 2.4, w: w - 1.0, h: 0,
      line: { color: C.line, width: 1 },
    });
    s.addText(it.d, {
      x: x + 0.2, y: y0 + 2.7, w: w - 0.4, h: 1.8,
      fontSize: 14, color: C.ink, fontFace: FONT_KO, align: "center", valign: "top", margin: 0,
    });
  });

  footer(s, idx + 1, TOTAL);
}
nextIdx();

// =========================================================
// Helper: scenario flow slide
// =========================================================
function scenarioFlowSlide(opts) {
  const s = pres.addSlide();
  bg(s, C.white);
  titleBar(s, opts.title, opts.sub);

  // Situation box
  card(s, 0.7, 1.5, W - 1.4, 1.4, { fill: C.sand, border: C.line });
  s.addText("현장 상황", {
    x: 0.9, y: 1.65, w: 2.0, h: 0.4,
    fontSize: 14, bold: true, color: C.amber, fontFace: FONT_KO, margin: 0,
  });
  s.addText(opts.situation, {
    x: 0.9, y: 2.05, w: W - 1.8, h: 0.85,
    fontSize: 15, color: C.ink, fontFace: FONT_KO, margin: 0,
  });

  // Flow visualization
  s.addText("AI 교무실 분업 흐름", {
    x: 0.7, y: 3.15, w: 6, h: 0.4,
    fontSize: 14, bold: true, color: C.teal, fontFace: FONT_KO, margin: 0,
  });

  const nodes = opts.nodes;
  const flowY = 3.7, flowH = 1.6;
  const totalGap = 0.4;
  const nodeW = (W - 1.4 - (nodes.length - 1) * totalGap) / nodes.length;

  nodes.forEach((n, i) => {
    const x = 0.7 + i * (nodeW + totalGap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: flowY, w: nodeW, h: flowH,
      fill: { color: n.fill || C.teal }, line: { color: n.fill || C.teal },
      rectRadius: 0.1,
      shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.12 },
    });
    s.addText(n.label, {
      x: x + 0.15, y: flowY + 0.15, w: nodeW - 0.3, h: 0.5,
      fontSize: 13, bold: true, color: C.white, fontFace: FONT_KO,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(n.role, {
      x: x + 0.15, y: flowY + 0.7, w: nodeW - 0.3, h: flowH - 0.85,
      fontSize: 11, color: C.white, fontFace: FONT_KO,
      align: "center", valign: "top", margin: 0,
    });

    if (i < nodes.length - 1 && opts.arrows !== false) {
      s.addShape(pres.shapes.RIGHT_TRIANGLE, {
        x: x + nodeW + 0.05, y: flowY + flowH / 2 - 0.12, w: 0.3, h: 0.24,
        fill: { color: C.amber }, line: { color: C.amber }, rotate: 90,
      });
    }
  });

  // Bottom callout
  s.addText(opts.takeaway, {
    x: 0.7, y: 5.7, w: W - 1.4, h: 1.0,
    fontSize: 16, italic: true, color: C.navy, fontFace: FONT_KO,
    valign: "middle", margin: 0, align: "center",
  });

  footer(s, idx + 1, TOTAL);
  return s;
}

function scenarioResultSlide(opts) {
  const s = pres.addSlide();
  bg(s, C.white);
  titleBar(s, opts.title, opts.sub);

  // Big stat card on left
  card(s, 0.7, 1.5, 4.5, 5.3, { fill: C.navy, border: C.navy });
  s.addText(opts.statValue, {
    x: 0.7, y: 1.8, w: 4.5, h: 2.0,
    fontSize: 92, bold: true, color: C.amber, fontFace: FONT_KO,
    align: "center", valign: "middle", margin: 0,
  });
  s.addText(opts.statLabel, {
    x: 0.9, y: 3.9, w: 4.1, h: 0.6,
    fontSize: 18, bold: true, color: C.white, fontFace: FONT_KO,
    align: "center", valign: "middle", margin: 0,
  });
  s.addText(opts.statSub, {
    x: 0.9, y: 4.5, w: 4.1, h: 2.0,
    fontSize: 13, color: C.sand, fontFace: FONT_KO,
    align: "center", valign: "top", margin: 0,
  });

  // Right side: bullets + takeaway
  s.addText("산출물", {
    x: 5.6, y: 1.5, w: 4, h: 0.4,
    fontSize: 14, bold: true, color: C.teal, fontFace: FONT_KO, margin: 0,
  });
  s.addText(opts.outputs.map((t, i) => ({
    text: t, options: { bullet: true, breakLine: i !== opts.outputs.length - 1 },
  })), {
    x: 5.6, y: 1.95, w: W - 6.3, h: 2.4,
    fontSize: 15, color: C.ink, fontFace: FONT_KO, paraSpaceAfter: 4,
  });

  s.addText("교무실에 던지는 한 줄", {
    x: 5.6, y: 4.5, w: 4, h: 0.4,
    fontSize: 14, bold: true, color: C.amber, fontFace: FONT_KO, margin: 0,
  });
  card(s, 5.6, 4.95, W - 6.3, 1.8, { fill: C.sand, border: C.amber });
  s.addText(opts.takeaway, {
    x: 5.8, y: 5.1, w: W - 6.7, h: 1.5,
    fontSize: 16, bold: true, color: C.navy, fontFace: FONT_KO,
    valign: "middle", margin: 0,
  });

  footer(s, idx + 1, TOTAL);
  return s;
}

// =========================================================
// Slide 6 — Scenario 1 flow (Sequential)
// =========================================================
scenarioFlowSlide({
  title: "장면 ① Sequential — “결재 라인처럼”",
  sub: "전공 실습 수업·평가 자료 제작 파이프라인",
  situation:
    "특성화고 바이오공정과 1학년 첫 학기, 첫 실습 “무균 조작과 미생물 배양 기초” 100분 블록.\n" +
    "신임 전공 교사가 NCS 능력단위·SOP·안전 수칙·평가 루브릭을 한꺼번에 만들어야 합니다.",
  nodes: [
    { label: "ncs-mapping-agent", role: "NCS 능력단위 매핑\n수행준거 도출", fill: C.teal },
    { label: "lesson-plan-agent", role: "차시 설계\n100분 블록 흐름", fill: C.teal },
    { label: "sop-author-agent",  role: "SOP · 체크리스트\n안전 수칙", fill: C.teal },
    { label: "rubric-reviewer-agent", role: "수행평가 루브릭\n안전·자료 검토", fill: C.teal },
  ],
  takeaway:
    "한 사람이 처음부터 끝까지 쓰던 4단계를, AI 4명이 결재 라인처럼 차례대로 검토합니다.",
});
nextIdx();

// =========================================================
// Slide 7 — Scenario 1 outcome
// =========================================================
scenarioResultSlide({
  title: "장면 ① 결과 — 신임 교사의 첫 실습 설계가 30분 만에",
  sub: "Sequential 패턴 — 직업계고 NCS 기반 실습 자료 제작",
  statValue: "1.5h",
  statLabel: "수업안 + 루브릭 + 안전점검",
  statSub: "기존 1~2일 분량의 첫 실습 설계를\n초안 단계까지 단축",
  outputs: [
    "NCS “화학·바이오 > 바이오공정 > 미생물 배양·관리” 능력단위 매핑표",
    "클린벤치 사용·배지 제조·도말 안전 수칙 SOP",
    "학생 자기점검 체크리스트 (2인 1조)",
    "산출물(접종 페트리디시 + 24h 콜로니 관찰) 채점 루브릭",
  ],
  takeaway:
    "“신규 전공 교사가 가장 두려워하는 첫 실습 설계” 부담을\n교사가 검토자·확정자 역할로 옮겨 줍니다.",
});
nextIdx();

// =========================================================
// Slide 8 — Scenario 2 flow (Concurrent)
// =========================================================
{
  const s = pres.addSlide();
  bg(s, C.white);
  titleBar(s, "장면 ② Concurrent — “학년부 회의처럼”",
    "행동특성 및 종합의견 / 산업체 추천 자료 동시 검토");

  card(s, 0.7, 1.5, W - 1.4, 1.3, { fill: C.sand, border: C.line });
  s.addText("현장 상황", {
    x: 0.9, y: 1.6, w: 2, h: 0.4, fontSize: 14, bold: true, color: C.amber, fontFace: FONT_KO, margin: 0,
  });
  s.addText(
    "2학기 말, 담임이 학급 24명의 행종을 작성해야 합니다.\n" +
    "전문교과·NCS / 출결·현장실습 출근 / 자격증·기능경기·창체 / 현장실습 평가·독서 — 4개 자료를 매번 따로 확인해야 합니다.",
    { x: 0.9, y: 1.95, w: W - 1.8, h: 0.8, fontSize: 14, color: C.ink, fontFace: FONT_KO, margin: 0 }
  );

  // Diagram: 4 parallel agents → 1 aggregator
  const agents = [
    { name: "subject-skill-agent",          role: "전문교과 · NCS 역량" },
    { name: "attendance-agent",             role: "출결 · 현장실습 출근" },
    { name: "certification-activity-agent", role: "자격증 · 기능경기 · 창체" },
    { name: "field-training-eval-agent",    role: "현장실습 평가 · 독서활동" },
  ];

  const colX = 0.9, colY = 3.1, nodeW = 2.7, nodeH = 1.1, gap = 0.18;
  agents.forEach((a, i) => {
    const x = colX + i * (nodeW + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: colY, w: nodeW, h: nodeH,
      fill: { color: C.teal }, line: { color: C.teal }, rectRadius: 0.1,
      shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.12 },
    });
    s.addText(a.name, {
      x: x + 0.1, y: colY + 0.12, w: nodeW - 0.2, h: 0.4,
      fontSize: 12, bold: true, color: C.white, fontFace: FONT_KO, align: "center", margin: 0,
    });
    s.addText(a.role, {
      x: x + 0.1, y: colY + 0.55, w: nodeW - 0.2, h: 0.5,
      fontSize: 11, color: C.white, fontFace: FONT_KO, align: "center", valign: "top", margin: 0,
    });
    // arrow down
    s.addShape(pres.shapes.LINE, {
      x: x + nodeW / 2, y: colY + nodeH + 0.05, w: 0, h: 0.6,
      line: { color: C.amber, width: 3 },
    });
    s.addShape(pres.shapes.RIGHT_TRIANGLE, {
      x: x + nodeW / 2 - 0.12, y: colY + nodeH + 0.55, w: 0.24, h: 0.2,
      fill: { color: C.amber }, line: { color: C.amber }, rotate: 180,
    });
  });

  // Aggregator
  const aggX = 4.0, aggY = colY + nodeH + 0.95, aggW = 5.3, aggH = 1.0;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: aggX, y: aggY, w: aggW, h: aggH,
    fill: { color: C.amber }, line: { color: C.amber }, rectRadius: 0.1,
    shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.18 },
  });
  s.addText("student-report-agent", {
    x: aggX, y: aggY + 0.1, w: aggW, h: 0.4,
    fontSize: 14, bold: true, color: C.white, fontFace: FONT_KO, align: "center", margin: 0,
  });
  s.addText("4개 트랙 종합 → 행종 또는 취업 추천 자료 초안", {
    x: aggX, y: aggY + 0.5, w: aggW, h: 0.4,
    fontSize: 12, color: C.white, fontFace: FONT_KO, align: "center", margin: 0,
  });

  s.addText(
    "병렬로 4개 자료 트랙을 동시 점검 → 종합 에이전트가 한 문단으로 묶어 줍니다.",
    {
      x: 0.7, y: 6.6, w: W - 1.4, h: 0.5,
      fontSize: 16, italic: true, color: C.navy, fontFace: FONT_KO, align: "center", margin: 0,
    }
  );

  footer(s, idx + 1, TOTAL);
}
nextIdx();

// =========================================================
// Slide 9 — Scenario 2 outcome
// =========================================================
scenarioResultSlide({
  title: "장면 ② 결과 — 학년부 회의가 24/7 돌아간다",
  sub: "Concurrent 패턴 — 행종·취업 추천 자료 종합",
  statValue: "×4",
  statLabel: "병렬 자료 트랙",
  statSub: "담임 1인이 직렬로 보던 자료를\n4개 AI 동료가 동시에 점검",
  outputs: [
    "학생 1명당 500자 행종 초안 (편입·취업 양립 표현 가능)",
    "자료 출처가 4축으로 분리 → 사후 검증·이의 제기 추적 가능",
    "학기 말 야근이 “감독·확정” 시간으로 전환",
    "산업체 제출용 자기PR + 담임 추천 분리 작성 옵션",
  ],
  takeaway:
    "담임의 야근이 아니라, 4개 자료 트랙을 동시에 돌리는\n학년부 회의가 24시간 켜져 있는 상태가 됩니다.",
});
nextIdx();

// =========================================================
// Slide 10 — Scenario 3 flow (Handoff)
// =========================================================
{
  const s = pres.addSlide();
  bg(s, C.white);
  titleBar(s, "장면 ③ Handoff — “민원 전화 응대처럼”",
    "학생·학부모 통합 안내 챗봇 — 자동 라우팅");

  card(s, 0.7, 1.5, W - 1.4, 1.3, { fill: C.sand, border: C.line });
  s.addText("현장 상황", {
    x: 0.9, y: 1.6, w: 2, h: 0.4, fontSize: 14, bold: true, color: C.amber, fontFace: FONT_KO, margin: 0,
  });
  s.addText(
    "“자격증 시험 출석인정”, “현장실습 부당 야근”, “선취업 후 재직자 특별전형” — 서로 다른 질문이 한 번호로 옵니다.\n" +
    "퇴근·주말에도 부모님 전화가 오지만, 모든 담당자가 24시간 대기할 수는 없습니다.",
    { x: 0.9, y: 1.95, w: W - 1.8, h: 0.8, fontSize: 14, color: C.ink, fontFace: FONT_KO, margin: 0 }
  );

  // Triage on top, 3 specialists below
  const triX = 5.5, triY = 3.1, triW = 2.4, triH = 0.9;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: triX, y: triY, w: triW, h: triH,
    fill: { color: C.navy }, line: { color: C.navy }, rectRadius: 0.1,
    shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.15 },
  });
  s.addText("triage-agent", {
    x: triX, y: triY + 0.1, w: triW, h: 0.35,
    fontSize: 14, bold: true, color: C.white, fontFace: FONT_KO, align: "center", margin: 0,
  });
  s.addText("질문 분류 · 라우팅", {
    x: triX, y: triY + 0.45, w: triW, h: 0.35,
    fontSize: 11, color: C.sand, fontFace: FONT_KO, align: "center", margin: 0,
  });

  const specialists = [
    { name: "academic-support-agent",      role: "학사 · 행정\n출석인정 · 기숙사 · NCS 수업", fill: C.teal },
    { name: "field-training-safety-agent", role: "현장실습 안전 · 권익\n부당 노동 · 표준협약 위반",  fill: "C97A5A" },
    { name: "career-pathway-agent",        role: "취업 · 진학\n협약기업 · 재직자 전형 · 일학습병행", fill: "7BA7BC" },
  ];

  const specY = 5.0, specW = 3.7, specH = 1.5, specGap = 0.3;
  const specStart = (W - (specW * 3 + specGap * 2)) / 2;
  specialists.forEach((sp, i) => {
    const x = specStart + i * (specW + specGap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: specY, w: specW, h: specH,
      fill: { color: sp.fill }, line: { color: sp.fill }, rectRadius: 0.1,
      shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.12 },
    });
    s.addText(sp.name, {
      x: x + 0.15, y: specY + 0.15, w: specW - 0.3, h: 0.4,
      fontSize: 13, bold: true, color: C.white, fontFace: FONT_KO, align: "center", margin: 0,
    });
    s.addText(sp.role, {
      x: x + 0.15, y: specY + 0.55, w: specW - 0.3, h: specH - 0.7,
      fontSize: 11, color: C.white, fontFace: FONT_KO, align: "center", valign: "top", margin: 0,
    });
    // arrow from triage to specialist
    s.addShape(pres.shapes.LINE, {
      x: triX + triW / 2, y: triY + triH,
      w: (x + specW / 2) - (triX + triW / 2), h: specY - (triY + triH),
      line: { color: C.amber, width: 2 },
    });
  });

  s.addText(
    "들어오는 모든 질문을 트리아지가 1차 분류 → 적합한 전문 에이전트에게 자동 위임.",
    {
      x: 0.7, y: 6.7, w: W - 1.4, h: 0.4,
      fontSize: 15, italic: true, color: C.navy, fontFace: FONT_KO, align: "center", margin: 0,
    }
  );

  footer(s, idx + 1, TOTAL);
}
nextIdx();

// =========================================================
// Slide 11 — Scenario 3 outcome
// =========================================================
scenarioResultSlide({
  title: "장면 ③ 결과 — 24/7 1차 안내 + 위험 사안 즉시 격리",
  sub: "Handoff 패턴 — 학생·학부모 통합 안내",
  statValue: "24/7",
  statLabel: "통합 안내 가능 시간",
  statSub: "퇴근·주말·방학에도\n1차 응대 + 위험 사안 알림",
  outputs: [
    "출석인정·기숙사·NCS 수업 같은 일반 질문은 자동 응대",
    "현장실습 부당 노동·안전 사안은 별도 트랙으로 즉시 격리·알림",
    "취업·진학(재직자 전형, 일학습병행, 산업기능요원) 통합 안내",
    "교사·교감은 “격리된 사안”만 직접 확인하면 됨",
  ],
  takeaway:
    "직업계고 공통 민원 패턴을 시도 교육청 단위로 표준화할 수 있는\n명확한 후보 — 장학사가 가장 먼저 봐야 할 패턴입니다.",
});
nextIdx();

// =========================================================
// Slide 12 — Scenario 4 flow (Group Chat)
// =========================================================
scenarioFlowSlide({
  title: "장면 ④ Group Chat — “진로 협의회처럼”",
  sub: "졸업 후 진로 결정 시뮬레이션 — 네 명의 시각이 동시에",
  situation:
    "바이오공정과 3학년, 협약 식품·바이오 기업 채용 vs 식품공학과·생명공학과 진학 사이.\n" +
    "본인·가정·학교의 의견이 엇갈리고, 협약 기업·재직자 전형·일학습병행을 어떤 비율로 묶을지 결정이 어렵습니다.",
  nodes: [
    { label: "homeroom-teacher-agent",     role: "담임교사\n학생 성향·전반",        fill: C.teal },
    { label: "major-instructor-agent",     role: "전공 실습 교사\n현장 역량·자격",   fill: "7BA7BC" },
    { label: "industry-partnership-agent", role: "산학협력 부장\n협약 기업·일학습병행", fill: C.amber },
    { label: "parent-agent",               role: "학부모\n가정 재정·우려",            fill: "C97A5A" },
  ],
  arrows: false,
  takeaway:
    "4명이 회의 형태로 토론·합의 → “취업 60% + 1~3년 후 재직자 전형 진학 40%”처럼\n비율과 분기 시점이 명시된 합의안이 나옵니다.",
});
nextIdx();

// =========================================================
// Slide 13 — Scenario 4 outcome
// =========================================================
scenarioResultSlide({
  title: "장면 ④ 결과 — 진로 협의회를 미리 한 번 돌려 보고 가기",
  sub: "Group Chat 패턴 — 졸업 후 진로 결정",
  statValue: "4 in 1",
  statLabel: "관점 동시 시뮬레이션",
  statSub: "담임 · 전공교사 · 산학협력 · 학부모를\n한 자리에 앉히는 효과",
  outputs: [
    "취업·진학·일학습병행·재직자 전형을 비율로 묶은 합의안",
    "1~2년 내 분기 시점(예: “2년 후 재직자 전형 출원 시점”) 명시",
    "본인·가정·학교의 우려가 표로 정리된 회의 요약",
    "실제 협의회 전에 사전 시뮬레이션 가능 → 회의 시간 단축",
  ],
  takeaway:
    "실제 협의회 전에 시나리오를 미리 돌려 보고 들어갈 수 있습니다.\n학부모와의 갈등도 “감정” 대신 “비율과 시점”으로 대화하게 됩니다.",
});
nextIdx();

// =========================================================
// Slide 14 — Governance: 단일 챗봇 vs AI 교무실
// =========================================================
{
  const s = pres.addSlide();
  bg(s, C.white);
  titleBar(s, "왜 “학교 차원의 AI 교무실”이 더 안전한가",
    "단일 ChatGPT 사용 vs 분업된 AI 팀 — 거버넌스 관점");

  const headers = ["관점", "단일 챗봇 사용", "AI 교무실 (분업)"];
  const rows = [
    ["일관성",     "교사마다 다른 프롬프트 · 다른 답",       "역할별 표준 지시문 · 일관된 산출"],
    ["검토 단계", "검토 없음. 결과를 그대로 사용",            "단계별 검토 · 종합 = 자연스러운 감사선"],
    ["출처 추적", "한 줄 응답. 어디서 나왔는지 알기 어려움",  "트랙별 출처 분리 · 사후 추적 가능"],
    ["개인정보", "외부 챗봇에 학생 정보 입력 위험",            "학교용 게이트웨이 통해서만 호출"],
    ["확장",       "교사 개인 노하우에 의존",                  "시도 단위 표준 프롬프트·데이터셋으로 확산"],
  ];

  const tx = 0.7, ty = 1.6, tw = W - 1.4;
  const colW = [2.0, (tw - 2.0) / 2, (tw - 2.0) / 2];
  const rowH = 0.85;

  // header
  s.addShape(pres.shapes.RECTANGLE, {
    x: tx, y: ty, w: tw, h: rowH,
    fill: { color: C.navy }, line: { color: C.navy },
  });
  let cx = tx;
  headers.forEach((h, i) => {
    s.addText(h, {
      x: cx + 0.15, y: ty, w: colW[i] - 0.3, h: rowH,
      fontSize: 16, bold: true, color: C.white, fontFace: FONT_KO,
      align: "center", valign: "middle", margin: 0,
    });
    cx += colW[i];
  });

  rows.forEach((r, ri) => {
    const ry = ty + rowH * (ri + 1);
    s.addShape(pres.shapes.RECTANGLE, {
      x: tx, y: ry, w: tw, h: rowH,
      fill: { color: ri % 2 === 0 ? C.sand : C.white },
      line: { color: C.line, width: 0.5 },
    });
    let cxr = tx;
    r.forEach((cell, ci) => {
      s.addText(cell, {
        x: cxr + 0.15, y: ry, w: colW[ci] - 0.3, h: rowH,
        fontSize: 13, color: ci === 0 ? C.navy : C.ink, bold: ci === 0,
        fontFace: FONT_KO, align: ci === 0 ? "center" : "left", valign: "middle", margin: 0,
      });
      cxr += colW[ci];
    });
  });

  s.addText("➤ 분업 = 자연스러운 감사선. 거버넌스를 “덧붙이는 비용”이 아니라 “설계의 부산물”로 얻습니다.", {
    x: 0.7, y: H - 0.95, w: W - 1.4, h: 0.4,
    fontSize: 14, italic: true, color: C.teal, fontFace: FONT_KO, align: "center", margin: 0,
  });

  footer(s, idx + 1, TOTAL);
}
nextIdx();

// =========================================================
// Slide 15 — 3-step roadmap
// =========================================================
{
  const s = pres.addSlide();
  bg(s, C.white);
  titleBar(s, "학교가 오늘부터 시작하는 3단계", "이번 학기 안에 실현 가능한 로드맵");

  const steps = [
    {
      n: "1",
      t: "가장 반복되는 문서 1종 선정",
      d: "행종 · 실습 수업안 · 표준협약 안내문 중 하나. 학교마다 가장 시간 잡아먹는 항목을 정합니다.",
      sub: "약 2주",
    },
    {
      n: "2",
      t: "2~3명 규모의 AI 팀으로 분업",
      d: "Sequential 또는 Concurrent 패턴부터. 교사 1~2명이 “감독·확정” 역할로 시범 운영합니다.",
      sub: "약 4주",
    },
    {
      n: "3",
      t: "시도 단위 표준 프롬프트로 확산",
      d: "NCS 매핑 데이터셋·공동 프롬프트 라이브러리·학생정보 보호 게이트웨이를 시도 교육청 단위로 표준화.",
      sub: "1학기",
    },
  ];

  const x0 = 0.7, y0 = 1.7, w = (W - 1.4 - 0.6) / 3, h = 4.5, gap = 0.3;
  steps.forEach((st, i) => {
    const x = x0 + i * (w + gap);
    card(s, x, y0, w, h);

    // big number
    s.addText(st.n, {
      x, y: y0 + 0.2, w, h: 1.8,
      fontSize: 110, bold: true, color: C.amber, fontFace: FONT_KO,
      align: "center", valign: "middle", margin: 0,
    });
    s.addShape(pres.shapes.LINE, {
      x: x + 0.6, y: y0 + 2.05, w: w - 1.2, h: 0,
      line: { color: C.amber, width: 2 },
    });
    s.addText(st.t, {
      x: x + 0.3, y: y0 + 2.2, w: w - 0.6, h: 0.9,
      fontSize: 18, bold: true, color: C.navy, fontFace: FONT_KO,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(st.d, {
      x: x + 0.3, y: y0 + 3.15, w: w - 0.6, h: 1.0,
      fontSize: 13, color: C.ink, fontFace: FONT_KO, align: "center", valign: "top", margin: 0,
    });
    s.addText(st.sub, {
      x: x + 0.3, y: y0 + h - 0.55, w: w - 0.6, h: 0.4,
      fontSize: 12, bold: true, color: C.teal, fontFace: FONT_KO, align: "center", margin: 0,
    });
  });

  s.addText("작게 시작 → 빠르게 검증 → 시도 단위로 확산.", {
    x: 0.7, y: H - 0.9, w: W - 1.4, h: 0.4,
    fontSize: 16, italic: true, color: C.navy, fontFace: FONT_KO, align: "center", margin: 0,
  });

  footer(s, idx + 1, TOTAL);
}
nextIdx();

// =========================================================
// Slide 16 — Policy implications (장학사 향)
// =========================================================
{
  const s = pres.addSlide();
  bg(s, C.white);
  titleBar(s, "시도 교육청 · 관리자에게 — 정책 시사점",
    "학교 단위 “AI 교무실”이 모이면 시도 단위 표준이 됩니다");

  const items = [
    {
      t: "① 표준 NCS 매핑 데이터셋",
      d: "전공별 능력단위·수행준거를 시도 단위로 한 번만 정리하면, 모든 학교의 ncs-mapping-agent가 같은 기준으로 동작합니다.",
    },
    {
      t: "② 공동 프롬프트 라이브러리",
      d: "행종·표준협약·진로 시뮬레이션 등 패턴별 표준 지시문을 시도 자산으로 관리. 학교는 미세 조정만.",
    },
    {
      t: "③ 학생정보 보호 게이트웨이",
      d: "교사 개인이 외부 챗봇에 학생 정보를 입력하지 않도록, “학교용 AI 게이트웨이”를 통해서만 호출.",
    },
    {
      t: "④ 모범 사례 공유 채널",
      d: "패턴 4종 × 전공 × 학년의 작동 사례를 시도 단위 모범 사례로 축적·공유.",
    },
  ];

  const x0 = 0.7, y0 = 1.7, w = (W - 1.4 - 0.4) / 2, h = 2.4, gap = 0.4;
  items.forEach((it, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = x0 + col * (w + gap), y = y0 + row * (h + 0.25);
    card(s, x, y, w, h);
    s.addText(it.t, {
      x: x + 0.3, y: y + 0.2, w: w - 0.6, h: 0.6,
      fontSize: 18, bold: true, color: C.teal, fontFace: FONT_KO, valign: "middle", margin: 0,
    });
    s.addText(it.d, {
      x: x + 0.3, y: y + 0.85, w: w - 0.6, h: h - 1.0,
      fontSize: 13, color: C.ink, fontFace: FONT_KO, valign: "top", margin: 0,
    });
  });

  footer(s, idx + 1, TOTAL);
}
nextIdx();

// =========================================================
// Slide 17 — Closing
// =========================================================
{
  const s = pres.addSlide();
  bg(s, C.navy);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.35, h: H, fill: { color: C.teal }, line: { color: C.teal } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: H - 0.4, w: W, h: 0.4, fill: { color: C.amber }, line: { color: C.amber } });

  s.addText("우리 학교에", {
    x: 1.0, y: 1.6, w: 11, h: 0.8,
    fontSize: 32, color: C.sand, fontFace: FONT_KO, margin: 0,
  });
  s.addText("첫 번째로 출근시킬", {
    x: 1.0, y: 2.4, w: 11, h: 0.8,
    fontSize: 32, color: C.sand, fontFace: FONT_KO, margin: 0,
  });
  s.addText("AI 동료는 누구입니까?", {
    x: 1.0, y: 3.2, w: 11, h: 1.5,
    fontSize: 60, bold: true, color: C.amber, fontFace: FONT_KO, margin: 0,
  });

  s.addText("Q & A", {
    x: 1.0, y: 5.4, w: 11, h: 0.8,
    fontSize: 36, bold: true, color: C.white, fontFace: FONT_KO, margin: 0,
  });
  s.addText("워크샵 자료 · 코드: aka.ms/agentframework/workshop/multiagent", {
    x: 1.0, y: H - 1.1, w: 11, h: 0.4,
    fontSize: 14, color: C.sand, fontFace: FONT_KO, margin: 0,
  });
}

// ---------- write ----------
const out = path.join(__dirname, "..", "AI-Staff-Room.pptx");
pres.writeFile({ fileName: out }).then(f => console.log("Wrote:", f));
