// (주)토인 회사소개서 리뉴얼 — KR / EN 공통 생성 스크립트
const pptxgen = require("pptxgenjs");
const path = require("path");
const { applyTheme } = require("./apply_theme.js");
const DIMS = require("./dims.json");
const IMG = (f) => path.join(__dirname, "img", f);

const HEX = {
  dk1: "1A2230", lt1: "FFFFFF", dk2: "0B1E3D", lt2: "F1F4F8",
  accent1: "005AAB", accent2: "6E7F96", accent3: "B86E00", accent4: "DCE6F2",
  accent5: "5B6676", accent6: "8FB8EA",
};

async function build(L) {
  const KO = L === "ko";
  const t = (ko, en) => (KO ? ko : en);
  const FONT = KO ? "Malgun Gothic" : "Arial";
  const THEME = {
    name: "TOIN 2027",
    headFontFace: FONT, bodyFontFace: FONT,
    colors: { ...HEX, hlink: "005AAB", folHlink: "5B6676" },
  };
  const NEED = t("입력 필요", "TBD");

  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  pres.title = t("(주)토인 회사소개서", "TOIN International Company Profile");
  pres.company = "TOIN International Co., Ltd.";
  const C = pres.SchemeColor;
  const S = pres.shapes;
  const W = 13.333;
  // 아이콘 (react-icons → PNG)
  const React = require("react"), RDS = require("react-dom/server"), FA = require("react-icons/fa"), sharp = require("sharp");
  const IC = {};
  for (const n of ["FaCar", "FaPlane", "FaMicrochip", "FaCogs", "FaRobot", "FaRulerHorizontal", "FaClock"]) {
    const svg = RDS.renderToStaticMarkup(React.createElement(FA[n], { color: "#005AAB", size: 256 }));
    IC[n] = "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64");
  }
  const iconCircle = (sl, n, x, y, d) => { sl.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: "FFFFFF" }, line: { color: "DCE6F2", width: 1 } }); sl.addImage({ data: IC[n], x: x + d * 0.24, y: y + d * 0.24, w: d * 0.52, h: d * 0.52 }); };

  // 발표자 노트에 확정 일정(2027-01-05 배포) 반영: 제작 메모마다 담당 마감일, 표지 노트에 전체 일정
  const DUE = { 1: 12, 2: 11, 3: 11, 4: 11, 5: 11, 6: 11, 7: 11, 21: 11, 23: 11 }; // 11 = 11/21(경영관리), 그 외 12/12
  const dueOf = (n) => (DUE[n] === 11 ? t("11/21", "Nov 21") : t("12/12", "Dec 12"));
  const SCHEDULE = t(
    "\n\n[일정] 사실관계 확정 11/17–11/21 → 자료 수집·사진 11/17–12/12 → v2 반영 12/14–12/19 → 검수 12/21–12/26 → 배포 2027-01-05 (확정)",
    "\n\n[Schedule] Facts confirmed Nov 17–21 → materials & photos due Dec 12 → v2 Dec 14–19 → review Dec 21–26 → release Jan 5, 2027 (confirmed)");
  let slideNo = 0;
  const _addSlide = pres.addSlide.bind(pres);
  pres.addSlide = (o) => {
    const sl = _addSlide(o), n = ++slideNo, _notes = sl.addNotes.bind(sl);
    sl.addNotes = (txt) => _notes(txt.replace("[제작 메모]", `[제작 메모 · 마감 ${dueOf(n)}]`).replace("[Production note]", `[Production note · due ${dueOf(n)}]`) + (n === 1 ? SCHEDULE : ""));
    return sl;
  };

  // ---------- layouts ----------
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: C.background1 },
    objects: [
      { placeholder: { options: { name: "eyebrow", type: "body", x: 0.6, y: 0.42, w: 9, h: 0.3, fontSize: 11, bold: true, color: C.accent1, charSpacing: 2, margin: 0, valign: "top" }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.72, w: 12.1, h: 0.62, fontSize: 28, bold: true, color: C.text1, margin: 0, valign: "middle", align: "left" }, text: "" } },
      { placeholder: { options: { name: "lead", type: "body", x: 0.6, y: 1.36, w: 12.1, h: 0.36, fontSize: 14, color: C.accent5, margin: 0, valign: "top" }, text: "" } },
      ...(KO ? [{ image: { x: 11.43, y: 6.98, w: 1.3, h: 1.3 * 303 / 1190, path: IMG("logo_full.png") } }] : [{ image: { x: 11.98, y: 6.96, w: 0.36, h: 0.3, path: IMG("logo_sym.png") } }, { text: { text: "TOIN", options: { x: 12.38, y: 6.96, w: 0.6, h: 0.3, fontSize: 14, bold: true, color: C.accent1, margin: 0, valign: "middle" } } }]),
      { text: { text: "TOIN International Co., Ltd.", options: { x: 1.0, y: 7.0, w: 4, h: 0.3, fontSize: 9, color: C.accent2, margin: 0, valign: "middle" } } },
    ],
    slideNumber: { x: 0.6, y: 7.0, w: 0.4, h: 0.3, fontSize: 9, color: C.accent1, bold: true, margin: 0 },
  });
  pres.defineSlideMaster({ title: "DARK", background: { path: IMG("bg_section.jpg") }, objects: [] });
  pres.defineSlideMaster({ title: "COVER", background: { path: IMG("bg_cover.jpg") }, objects: [] });

  // ---------- helpers ----------
  const shadow = () => ({ type: "outer", color: "0B1E3D", opacity: 0.12, blur: 8, offset: 2, angle: 90 });
  const card = (s, x, y, w, h, o = {}) =>
    s.addShape(S.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: o.r ?? 0.1, fill: { color: o.fill ?? C.background2 }, line: o.line ? { color: o.line, width: 1, dashType: o.dash ?? "solid" } : { type: "none" }, shadow: o.shadow ? shadow() : undefined, objectName: o.name });
  const txt = (s, text, o) => s.addText(text, { isTextBox: true, margin: 0, valign: "top", color: C.text1, fontSize: 14, ...o });
  const need = (label = NEED) => ({ text: label, options: { color: C.accent3, bold: true } });
  function img(s, file, x, y, w, h, mode = "contain", o = {}) {
    const [iw, ih] = DIMS[file];
    if (mode === "cover") {
      return s.addImage({ path: IMG(file), x, y, w: iw / 100, h: ih / 100, sizing: { type: "cover", w, h }, rounding: false, ...o });
    }
    const r = Math.min(w / iw, h / ih);
    const ww = iw * r, hh = ih * r;
    return s.addImage({ path: IMG(file), x: x + (w - ww) / 2, y: y + (h - hh) / 2, w: ww, h: hh, ...o });
  }
  function photoSlot(s, x, y, w, h, label) { // 자료 입력용 사진 자리
    card(s, x, y, w, h, { fill: C.background2, line: C.accent3, dash: "dash", r: 0.08 });
    txt(s, [{ text: t("사진 자리", "Photo slot"), options: { bold: true, breakLine: true } }, { text: label, options: {} }], { x: x + 0.15, y, w: w - 0.3, h, align: "center", valign: "middle", color: C.accent3, fontSize: 12 });
  }
  function head(s, eyebrow, title, lead) {
    s.addText(eyebrow, { placeholder: "eyebrow" });
    s.addText(title, { placeholder: "title" });
    s.addText(lead, { placeholder: "lead" });
  }
  const content = (section) => pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  function chevrons(s, labels, x, y, w, h, o = {}) {
    const gap = 0.06, n = labels.length, cw = (w - gap * (n - 1)) / n;
    labels.forEach((lb, i) => {
      s.addText(lb, { shape: i === 0 ? S.PENTAGON : S.CHEVRON, x: x + i * (cw + gap), y, w: cw, h, fill: { color: o.fill ?? C.accent1 }, color: C.background1, fontSize: o.fontSize ?? 13, bold: true, align: "center", valign: "middle", margin: [0, 4, 0, 14] });
    });
    return cw;
  }

  const SEC = {
    intro: t("표지·인사", "Opening"), about: t("회사 소개", "About TOIN"),
    sheet: t("판금기계사업부", "Sheet Metal"), tool: t("공작기계사업부", "Machine Tools"),
    auto: t("자동차부품사업부", "Auto Parts"), svc: t("서비스·네트워크", "Service & Network"), end: t("마무리", "Closing"),
  };

  // ===== 1. COVER =====
  pres.addSection({ title: SEC.intro });
  let s = pres.addSlide({ masterName: "COVER", sectionTitle: SEC.intro });
  if (KO) img(s, "logo_full_white.png", 0.7, 0.6, 2.8, 2.8 * 303 / 1190);
  else { img(s, "logo_sym_white.png", 0.7, 0.55, 0.9, 0.75); txt(s, "TOIN", { x: 1.7, y: 0.55, w: 3, h: 0.75, fontSize: 36, bold: true, color: C.background1, valign: "middle" }); }
  txt(s, "COMPANY PROFILE", { x: 0.7, y: 2.05, w: 6, h: 0.35, fontSize: 13, bold: true, color: C.accent6, charSpacing: 4 });
  txt(s, [{ text: "LEAD TO A NEW", options: { breakLine: true } }, { text: "PARADIGM" }], { x: 0.7, y: 2.5, w: 7, h: 1.9, fontSize: 54, bold: true, color: C.background1 });
  txt(s, t("새로운 패러다임으로 미래를 선도하는 기업, 토인", "Leading the future with a new paradigm"), { x: 0.7, y: 4.5, w: 6.6, h: 0.45, fontSize: 20, color: C.background1 });
  txt(s, t("판금기계 · 공작기계 · 자동차부품  |  Since 1992", "Sheet Metal · Machine Tools · Auto Parts  |  Since 1992"), { x: 0.7, y: 5.15, w: 6.6, h: 0.35, fontSize: 14, color: C.accent6 });
  txt(s, "www.toinint.com", { x: 0.7, y: 6.6, w: 4, h: 0.3, fontSize: 12, color: C.accent6 });
  card(s, 7.75, 0.75, 4.9, 6.0, { fill: C.background1, r: 0.12 });
  img(s, "punch.jpg", 7.95, 1.05, 4.5, 3.3);
  txt(s, t("MURATEC 터렛 펀치 프레스", "MURATEC turret punch press"), { x: 7.95, y: 4.45, w: 4.5, h: 0.3, fontSize: 12, color: C.accent5, align: "center" });
  img(s, "lathe.png", 8.4, 4.85, 3.6, 1.7);
  s.addNotes(t(
    "[발표] 안녕하십니까, (주)토인입니다. 저희 슬로건 'LEAD TO A NEW PARADIGM'처럼, 토인은 새로운 패러다임으로 고객의 생산 현장을 이끄는 기업이 되고자 합니다. 약 15분간 회사 소개를 드리겠습니다.\n\n[제작 메모] 대표 장비 고해상도 사진(현재 이미지는 저해상도) 교체 필요.\n[제작 메모 · 마감 11/21] 로고 원본(ai/eps) 수령 후 흰색 로고 교체 (경영관리부).",
    "[Talk] Good morning, we are TOIN International. As our slogan says — LEAD TO A NEW PARADIGM — we aim to lead our customers' production floors with a new paradigm. This introduction takes about 15 minutes.\n\n[Production note] Replace equipment photos with high-resolution originals.\n[Production note · due Nov 21] Replace the white logo with the official vector file (Management)."));

  // ===== 2. CEO MESSAGE =====
  s = content(SEC.intro);
  head(s, "CEO MESSAGE", t("대표 인사말", "Message from the CEO"), t("정밀 · 혁신 · 신뢰로 고객의 생산성을 높이는 파트너", "A partner that raises customer productivity through precision, innovation and trust"));
  photoSlot(s, 0.6, 1.95, 3.4, 4.0, t("대표이사 사진\n(배경 제거본)", "CEO portrait\n(background removed)"));
  txt(s, [{ text: t("대표이사  ", "CEO  "), options: { color: C.accent5 } }, { text: t("박호복", "Park Ho-bok"), options: { bold: true } }], { x: 0.6, y: 6.1, w: 3.4, h: 0.4, fontSize: 16, align: "center" });
  txt(s, KO ? [
    { text: "안녕하십니까. (주)토인 대표이사 박호복입니다.", options: { bold: true, breakLine: true, paraSpaceAfter: 8 } },
    { text: "1992년 토인무역상사로 첫발을 내디딘 이래, 토인은 일본·유럽·중국의 우수한 산업장비를 국내 제조 현장에 소개하고 설치부터 유지보수까지 장비의 전 생애를 책임져 왔습니다.", options: { breakLine: true, paraSpaceAfter: 8 } },
    { text: "현재 판금기계와 공작기계 3,300여 대가 전국 현장에서 가동 중이며, 대구 국가산업단지의 자동화 공장에서는 자동차 변속기 핵심 부품을 생산하고 있습니다.", options: { breakLine: true, paraSpaceAfter: 8 } },
    { text: "LEAD TO A NEW PARADIGM — 토인은 고객의 생산성을 높이는 새로운 기준을 만들어 가겠습니다. 감사합니다.", options: {} },
  ] : [
    { text: "Thank you for your interest in TOIN International.", options: { bold: true, breakLine: true, paraSpaceAfter: 8 } },
    { text: "Since our founding as Toin Trading in 1992, we have introduced leading industrial equipment from Japan, Europe and China to Korean manufacturers — and supported it throughout its service life, from installation to maintenance.", options: { breakLine: true, paraSpaceAfter: 8 } },
    { text: "More than 3,300 of our sheet metal and machine tool units run on production floors across Korea, and our automated plant in the Daegu National Industrial Complex produces core transmission parts.", options: { breakLine: true, paraSpaceAfter: 8 } },
    { text: "LEAD TO A NEW PARADIGM — we will keep setting new standards for our customers' productivity.", options: {} },
  ], { x: 4.5, y: 1.95, w: 8.2, h: 2.95, fontSize: KO ? 15 : 14, lineSpacingMultiple: 1.15 });
  txt(s, t("※ 초안 — 대표이사 검토 후 확정 (400~600자)", "※ Draft — to be confirmed by the CEO"), { x: 4.5, y: 4.9, w: 8.2, h: 0.28, fontSize: 11, color: C.accent3, bold: true });
  [["Precision", t("정밀", "Accuracy"), t("가공 정밀도와 품질로 증명", "Proven in every part we machine")], ["Innovation", t("혁신", "Progress"), t("자동화·스마트 공정 선도", "Leading in factory automation")], ["Trust", t("신뢰", "Reliability"), t("판매 이후까지 책임", "Responsible beyond the sale")]].forEach((v, i) => {
    const x = 4.5 + i * 2.8;
    card(s, x, 5.35, 2.6, 1.35, { fill: C.text2 });
    txt(s, [{ text: v[0], options: { bold: true, fontSize: 18, color: C.background1, breakLine: true } }, ...(KO ? [{ text: v[1], options: { fontSize: 12, color: C.accent6, breakLine: true } }] : []), { text: v[2], options: { fontSize: 12, color: C.background1 } }], { x: x + 0.2, y: 5.47, w: 2.25, h: 1.15 });
  });
  s.addNotes(t(
    "[발표] 대표 인사말을 대신 전해 드립니다. (인사말 낭독 또는 요약)\n\n[제작 메모] 인사말은 초안입니다. 대표이사 검토 후 확정. 대표 사진(배경 제거본) 필요.",
    "[Talk] A message from our CEO.\n\n[Production note] Draft text; CEO approval and portrait photo required."));

  // ===== 3. OVERVIEW =====
  pres.addSection({ title: SEC.about });
  s = content(SEC.about);
  head(s, "ABOUT US", t("회사 개요", "Company Overview"), t("1992년 설립 이래 글로벌 산업장비의 수입·판매·서비스와 자동차부품 제조를 함께 해 온 기업", "Importing, selling and servicing global industrial equipment — and manufacturing auto parts — since 1992"));
  const stats = [
    ["1992", t("설립", "Founded")],
    [t("430억", "₩43bn"), t("연매출 (2025, 잠정)", "Revenue (2025, prelim.)")],
    [t("89명", "89"), t("임직원 (2025)", "Employees (2025)")],
    ["3,300+", t("전국 가동 장비 (판금 1,500 · 공작 1,800)", "Machines installed (sheet metal 1,500 · lathes 1,800)")],
    [t("3개", "3"), t("사업부", "Business divisions")],
    [t("2곳", "2"), t("사업장 (시흥 본사 · 대구지사)", "Sites (Siheung HQ · Daegu)")],
  ];
  stats.forEach((st, i) => {
    const x = 0.6 + (i % 3) * 2.42, y = 1.95 + Math.floor(i / 3) * 2.45;
    card(s, x, y, 2.27, 2.25);
    txt(s, st[0], { x: x + 0.22, y: y + 0.3, w: 1.9, h: 0.8, fontSize: 34, bold: true, color: C.accent1 });
    txt(s, st[1], { x: x + 0.22, y: y + 1.2, w: 1.9, h: 0.9, fontSize: 13, color: C.accent5 });
  });
  const info = [
    [t("회사명", "Company"), t("주식회사 토인 (TOIN International Co., Ltd.)", "TOIN International Co., Ltd.")],
    [t("대표이사", "CEO"), t("박호복", "Park Ho-bok")],
    [t("설립", "Founded"), t("1992년 2월", "February 1992")],
    [t("업종", "Business"), t("무역(기계·전자부품), 자동화 설비, 자동차 부품, 금형", "Machinery & electronic parts trading, automation equipment, auto parts, molds")],
    [t("본사", "HQ"), t("경기도 시흥시 시화벤처로 152, 시화MTV 1사 501호", "#501, Sihwa MTV 1-sa, 152 Sihwaventure-ro, Siheung-si, Gyeonggi-do")],
    [t("대구지사", "Daegu"), t("대구광역시 달성군 구지면 국가산단서로40길 28", "28 Gukgasandanseo-ro 40-gil, Guji-myeon, Dalseong-gun, Daegu")],
  ];
  s.addTable(info.map((r) => [
    { text: r[0], options: { bold: true, color: C.accent1, fontSize: 12 } },
    { text: r[1], options: { fontSize: 12, color: C.text1 } },
  ]), { x: 8.05, y: 1.95, w: 4.68, colW: [1.05, 3.63], rowH: 0.7, border: { type: "solid", pt: 0.75, color: "DCE6F2" }, valign: "middle", fontFace: FONT });
  s.addNotes(t(
    "[발표] 토인은 1992년 설립되어 판금기계, 공작기계, 자동차부품의 3개 사업부를 운영하고 있습니다. 2025년 연매출 약 430억 원, 임직원 89명이며, 판금·공작기계 합계 3,300여 대가 전국에서 가동 중입니다.\n\n[제작 메모] 2025년 확정 매출, 2026년 현재 임직원 수로 갱신. 설치 대수(1,500/1,800)는 2025.03 자료 기준 — 최신화 확인.",
    "[Talk] Founded in 1992, TOIN runs three divisions. 2025 revenue was about KRW 43bn with 89 employees; more than 3,300 of our machines run across Korea.\n\n[Production note] Update with final 2025 revenue and current headcount; confirm installed-base figures."));

  // ===== 4. HISTORY =====
  s = content(SEC.about);
  head(s, "HISTORY", t("연혁", "History"), t("무역상사에서 시작해 판매·서비스·제조를 아우르는 기업으로", "From a trading company to a sales, service and manufacturing group"));
  const eras = [
    ["1992–2002", t("창립과 기반", "Foundation"), [["1992.02", t("토인무역상사 설립 · Murata Machinery 판매 대리점", "Toin Trading founded · Murata Machinery dealer")], ["1993.07", t("갑류 무역대리업 등록", "Registered as trade agent")], ["1997.07", t("토인무역㈜ 법인 전환", "Incorporated as Toin Trading Co.")], ["2002.09", t("상호 변경 ㈜토인 · 시화공단 신축 이전", "Renamed TOIN Co., Ltd. · new building in Sihwa")]]],
    ["2004–2012", t("전국 네트워크", "Nationwide network"), [["2004.03", t("영남지사 설립 (경주)", "Yeongnam branch (Gyeongju)")], ["2005.01", t("대구지사로 확장 이전", "Branch relocated and expanded to Daegu")], ["2007.03", t("본사 신축 이전 · ISO 9001*", "New HQ building · ISO 9001*")], ["2012.12", t("대구지사 확장 이전 (성서공단)", "Daegu branch expanded (Seongseo)")]]],
    ["2013–2017", t("제조 도약", "Manufacturing leap"), [["2013", t("기업부설연구소 설립 · IBK 우수거래기업*", "Corporate R&D center · IBK Excellent Customer*")], ["2016.09", t("본사 시화MTV 신축 이전", "HQ moved to Sihwa MTV")], ["2017.08", t("대구 국가산단 신공장 (대지 5,000평)", "New Daegu plant (16,529 m²)")]]],
    ["2020–", t("품질과 신뢰", "Quality & trust"), [["2020.12", t("위험성평가 우수사업장 인정", "Certified excellent risk-assessment workplace")], ["2021.06", t("기술역량 우수기업 인증", "Excellent Technology Competence certification (KED)")], [NEED, t("IATF 16949 · ISO 14001 · T-SQ 인증", "IATF 16949 · ISO 14001 · T-SQ")]]],
  ];
  eras.forEach((e, i) => {
    const x = 0.6 + i * 3.08, y = 1.95;
    card(s, x, y, 2.9, 4.65);
    txt(s, e[0], { x: x + 0.22, y: y + 0.22, w: 2.5, h: 0.45, fontSize: 22, bold: true, color: C.accent1 });
    txt(s, e[1], { x: x + 0.22, y: y + 0.7, w: 2.5, h: 0.32, fontSize: 14, bold: true });
    const runs = [];
    e[2].forEach((it, k) => {
      runs.push({ text: it[0], options: { bold: true, color: it[0] === NEED ? C.accent3 : C.accent1, breakLine: true } });
      runs.push({ text: it[1], options: { breakLine: k < e[2].length - 1, paraSpaceAfter: 9 } });
    });
    txt(s, runs, { x: x + 0.22, y: y + 1.2, w: 2.5, h: 3.3, fontSize: 13 });
  });
  txt(s, t("* 증빙 확인 후 확정", "* To be confirmed against records"), { x: 0.6, y: 6.66, w: 6, h: 0.25, fontSize: 10, color: C.accent3 });
  s.addNotes(t(
    "[발표] 1992년 Murata 기계 대리점으로 출발해 2002년 ㈜토인으로 상호를 변경했고, 2004년 영남지사를 시작으로 전국 서비스망을 구축했습니다. 2017년에는 대구 국가산업단지에 5,000평 규모 신공장을 세워 자동차부품 대량생산 체계를 갖췄습니다.\n\n[제작 메모] 설립 연도 1992로 확정(2026.11). 1993.07 갑류 무역대리업 등록 연도 재확인. 2007 ISO 9001, 2013 연구소·IBK 수상은 증빙 확인. IATF/ISO 14001/T-SQ 취득 연도 입력.",
    "[Talk] TOIN began in 1992 as a Murata Machinery dealer, became TOIN in 2002, built a nationwide network from 2004, and opened its Daegu plant in 2017.\n\n[Production note] Founding year fixed as 1992. Verify starred items and add certification years."));

  // ===== 5. GROWTH =====
  s = content(SEC.about);
  head(s, "FINANCIAL", t("성장 지표", "Growth"), t("최근 5년간 매출 1.7배 성장", "Revenue up 1.7x from 2020 to 2024"));
  const rev = [253, 293, 329, 446, 438, 430];
  s.addChart(pres.charts.BAR, [{ name: t("매출액(억원)", "Revenue (KRW bn)"), labels: ["2020", "2021", "2022", "2023", "2024", t("2025(잠정)", "2025(prelim.)")], values: KO ? rev : rev.map((v) => v / 10) }], {
    x: 0.6, y: 1.95, w: 7.6, h: 4.75, barDir: "col", chartColors: ["005AAB"], barGapWidthPct: 55,
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: "1A2230", dataLabelFontSize: 13, dataLabelFontFace: "+mn-lt", dataLabelFontBold: true,
    dataLabelFormatCode: KO ? "#,##0" : "0.0",
    catAxisLabelColor: "5B6676", catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt", catAxisLineShow: false,
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
    showTitle: true, title: t("연도별 매출액 (단위: 억원)", "Annual revenue (KRW bn)"), titleFontSize: 13, titleColor: "5B6676", titleFontFace: "+mn-lt",
  });
  [["+73%", t("2020 → 2024 매출 증가", "Revenue growth 2020 → 2024")], ["14.7%", t("연평균 성장률 (2020–2024)", "CAGR 2020–2024")], [t("446억", "₩44.6bn"), t("최고 매출 (2023)", "Peak revenue (2023)")]].forEach((k, i) => {
    const y = 1.95 + i * 1.62;
    card(s, 8.65, y, 4.08, 1.45);
    txt(s, k[0], { x: 8.9, y: y + 0.2, w: 3.6, h: 0.65, fontSize: 32, bold: true, color: C.accent1 });
    txt(s, k[1], { x: 8.9, y: y + 0.88, w: 3.6, h: 0.35, fontSize: 13, color: C.accent5 });
  });
  s.addNotes(t(
    "[발표] 2020년 253억 원에서 2024년 438억 원으로 4년간 73% 성장했으며, 연평균 성장률은 약 15%입니다.\n\n[제작 메모] 2025년 430억은 3/25 초안 수치(잠정) — 결산 확정치로 교체. 2025.03 자료에는 '430억(2024 기준)'과 그래프 438억이 혼재했음.",
    "[Talk] Revenue grew from KRW 25.3bn in 2020 to 43.8bn in 2024 — up 73%, about 15% a year.\n\n[Production note] Replace preliminary 2025 figure with audited result."));

  // ===== 6. ORGANIZATION =====
  s = content(SEC.about);
  head(s, "ORGANIZATION", t("조직 구조", "Organization"), t("본사(시흥)와 대구지사, 3개 사업부와 기업부설연구소 (임직원 2025년 기준)", "Siheung HQ and Daegu branch — three divisions and an R&D center (headcount as of 2025)"));
  const line = (x, y, w, h) => s.addShape(S.LINE, { x, y, w, h, line: { color: "6E7F96", width: 1.25 } });
  const box = (x, y, w, h, text, o = {}) => s.addText(text, { shape: S.ROUNDED_RECTANGLE, rectRadius: 0.06, x, y, w, h, fill: { color: o.fill ?? C.accent1 }, line: o.line ? { color: o.line, width: 1.25 } : { type: "none" }, color: o.color ?? C.background1, fontSize: o.fs ?? 14, bold: o.bold ?? true, align: "center", valign: "middle", margin: 2 });
  const colW = 2.25, gapC = 0.22, cx = (i) => 0.6 + i * (colW + gapC) + colW / 2;
  box(W / 2 - 1.2, 1.9, 2.4, 0.5, t("대표이사", "CEO"), { fill: C.text2 });
  line(W / 2 + 1.2, 2.15, 0.9, 0); box(W / 2 + 2.1, 1.9, 2.5, 0.5, t("기업부설연구소", "R&D Center"), { fill: C.background1, line: C.accent1, color: C.accent1 });
  const hqC = (cx(0) + cx(1)) / 2, dgC = cx(3);
  line(W / 2, 2.4, 0, 0.25); line(hqC, 2.65, dgC - hqC, 0); line(hqC, 2.65, 0, 0.2); line(dgC, 2.65, 0, 0.2);
  box(hqC - 1.3, 2.85, 2.6, 0.48, t("본사 · 시흥 (23명)", "HQ · Siheung (23)"), { fill: C.text2, fs: 13 });
  box(dgC - 1.3, 2.85, 2.6, 0.48, t("대구지사 (66명)", "Daegu Branch (66)"), { fill: C.text2, fs: 13 });
  line(hqC, 3.33, 0, 0.2); line(dgC, 3.33, 0, 0.2);
  line(cx(0), 3.53, cx(1) - cx(0), 0); line(cx(2), 3.53, cx(4) - cx(2), 0);
  const divs = [
    [t("경영관리부", "Management"), [t("인사총무팀", "HR & General Affairs"), t("회계팀", "Accounting"), t("경리팀", "Finance")]],
    [t("판금기계사업부", "Sheet Metal Div."), [t("영업관리팀", "Sales Admin"), t("영업팀", "Sales"), t("기술지원1팀", "Tech Support 1")]],
    [t("공작기계사업부", "Machine Tool Div."), [t("영업팀", "Sales"), t("가공팀", "Machining"), t("개조팀", "Retrofit")]],
    [t("자동차부품사업부", "Auto Parts Div."), [t("영업관리팀", "Sales Admin"), t("생산관리팀", "Production Control"), t("품질관리팀", "Quality Control"), t("설비보전팀", "Maintenance"), t("생산1팀", "Production 1"), t("생산2팀", "Production 2")]],
    [t("기술지원2팀", "Tech Support 2"), [t("판금기계사업부 소속", "Part of Sheet Metal Div.")]],
  ];
  divs.forEach((d, i) => {
    const x = 0.6 + i * (colW + gapC);
    line(cx(i), 3.53, 0, 0.17);
    box(x, 3.7, colW, 0.48, d[0], { fs: 13 });
    d[1].forEach((tm, k) => box(x + 0.15, 4.3 + k * 0.42, colW - 0.3, 0.34, tm, { fill: C.background2, color: C.text1, bold: false, fs: 12 }));
  });
  s.addNotes(t(
    "[발표] 본사는 경기도 시흥에서 경영관리부와 판금기계사업부를, 대구지사는 국가산업단지에서 공작기계사업부와 자동차부품사업부를 운영합니다. 대구지사에는 판금기계사업부 기술지원2팀이 상주해 영남권 서비스를 담당합니다.\n\n[제작 메모] 최신 조직도·인원수(2026) 반영. 3/25본의 조직명(경영관리부)과 2025.03본(경영지원부) 표기 통일.",
    "[Talk] Siheung HQ houses management and the sheet metal division; the Daegu branch runs machine tools and auto parts, plus Tech Support 2 for the southeast region.\n\n[Production note] Update with the 2026 org chart and headcount."));

  // ===== 7. BUSINESS STRUCTURE =====
  s = content(SEC.about);
  head(s, "BUSINESS DIVISIONS", t("사업 구조", "Business Structure"), t("세 개의 사업부와, 이를 받치는 하나의 서비스 체계", "Three divisions, one service platform"));
  const bd = [
    ["01", t("판금기계사업부", "Sheet Metal Division"), "Sheet Metal Division", "punch.jpg", "contain", t("절단·펀칭·절곡·패널벤딩·자동화 토털 솔루션", "Cutting, punching, bending, panel bending and automation"), "MURATEC · HSG · ADIRA · KYOWA · AIZAWA", t("1,500+ 대 가동", "1,500+ units")],
    ["02", t("공작기계사업부", "Machine Tool Division"), "Machine Tool Division", "lathe.png", "contain", t("MURATEC CNC 선반 — 자동차·항공·전자 정밀 가공", "MURATEC CNC lathes for automotive, aerospace and electronics"), "MURATEC", t("1,800+ 대 가동", "1,800+ units")],
    ["03", t("자동차부품사업부", "Auto Parts Division"), "Auto Parts Division", "factory67.jpg", "cover", t("변속기 핵심 부품 무인 자동화 생산", "Automated production of transmission parts"), t("경창산업 등 현대·기아 협력사 납품", "Supplier to Hyundai-Kia tier-1s"), t("대지 5,000평 공장", "16,529 m² plant")],
  ];
  bd.forEach((b, i) => {
    const x = 0.6 + i * 4.1, y = 1.95;
    card(s, x, y, 3.9, 3.95, { fill: C.background1, shadow: true });
    if (b[4] === "cover") img(s, b[3], x, y, 3.9, 1.55, "cover"); else img(s, b[3], x + 0.2, y + 0.1, 3.5, 1.4);
    txt(s, [{ text: b[0] + "  ", options: { color: C.accent1, bold: true } }, { text: b[1], options: { bold: true } }], { x: x + 0.25, y: y + 1.7, w: 3.45, h: 0.4, fontSize: 18 });
    if (KO) txt(s, b[2], { x: x + 0.25, y: y + 2.1, w: 3.45, h: 0.28, fontSize: 11, color: C.accent2 });
    txt(s, b[5], { x: x + 0.25, y: y + 2.45, w: 3.45, h: 0.62, fontSize: 13 });
    txt(s, b[6], { x: x + 0.25, y: y + 3.1, w: 3.45, h: 0.32, fontSize: 11, color: C.accent5 });
    txt(s, b[7], { x: x + 0.25, y: y + 3.45, w: 3.45, h: 0.35, fontSize: 14, bold: true, color: C.accent1 });
  });
  card(s, 0.6, 6.08, 12.1, 0.68, { fill: C.text2 });
  txt(s, [{ text: t("공통 서비스 플랫폼   ", "Shared service platform   "), options: { bold: true, color: C.accent6 } }, { text: t("기술 컨설팅 · 설치/시운전 · 운전 교육 · 정기점검/예방정비 · 신속 A/S · 순정 부품", "Consulting · Installation · Training · Preventive maintenance · After-sales service · Genuine parts"), options: { color: C.background1 } }], { x: 0.9, y: 6.08, w: 11.6, h: 0.68, fontSize: 14, valign: "middle" });
  s.addNotes(t(
    "[발표] 토인의 사업은 판금기계, 공작기계, 자동차부품 세 축이며, 모든 사업부를 설치·교육·정비·A/S의 공통 서비스 체계가 뒷받침합니다.\n\n[제작 메모] 자동차부품 고객 표현: 3/25본 '현대·기아 1차 협력'은 실제 거래 구조(경창산업 납품) 확인 후 확정.",
    "[Talk] Three divisions, backed by one service platform.\n\n[Production note] Confirm the customer tier wording for auto parts."));

  // ===== section divider helper =====
  function divider(sec, no, name, en, desc, statsArr, imgFile, mode, notes) {
    pres.addSection({ title: sec });
    const d = pres.addSlide({ masterName: "DARK", sectionTitle: sec });
    txt(d, no, { x: 0.7, y: 0.9, w: 3, h: 1.2, fontSize: 72, bold: true, color: C.accent6 });
    txt(d, name, { x: 0.7, y: 2.2, w: 6.2, h: 0.75, fontSize: 36, bold: true, color: C.background1 });
    if (KO) txt(d, en, { x: 0.7, y: 2.95, w: 6.2, h: 0.4, fontSize: 16, color: C.accent6 });
    txt(d, desc, { x: 0.7, y: 3.55, w: 5.9, h: 0.95, fontSize: 15, color: C.background1 });
    statsArr.forEach((st, i) => {
      const x = 0.7 + i * 2.0;
      txt(d, st[0], { x, y: 4.95, w: 1.9, h: 0.65, fontSize: 30, bold: true, color: C.background1 });
      txt(d, st[1], { x, y: 5.6, w: 1.85, h: 0.65, fontSize: 12, color: C.accent6 });
    });
    card(d, 7.3, 0.75, 5.35, 6.0, { fill: C.background1, r: 0.12 });
    if (mode === "cover") img(d, imgFile, 7.3, 0.75, 5.35, 6.0, "cover"); else img(d, imgFile, 7.55, 1.0, 4.85, 5.5);
    d.addNotes(notes);
  }

  // ===== 8. SHEET METAL DIVIDER =====
  divider(SEC.sheet, "01", t("판금기계사업부", "Sheet Metal Division"), "Sheet Metal Division",
    t("절단부터 펀칭·절곡·패널벤딩·자동화까지, 판금 전 공정을 하나의 파트너로", "From cutting to punching, bending, panel bending and automation — one partner for every step"),
    [["1,500+", t("전국 가동 장비", "Units in operation")], ["5", t("글로벌 브랜드", "Global brands")], ["4", t("해외 판매국 (中·越·印·美)", "Export markets")]],
    "punch.jpg", "contain",
    t("[발표] 첫 번째, 판금기계사업부입니다. 일본·유럽·중국의 판금 장비를 공급하며 전국 1,500여 대가 가동 중이고, 중국·베트남·인도·미국에도 장비와 부품을 판매하고 있습니다.\n\n[제작 메모] 가동 대수·해외 판매국 최신화. 섹션 대표 사진은 레이저/전시장 고해상도 컷 권장.",
      "[Talk] Our sheet metal division supplies equipment from Japan, Europe and China, with 1,500+ units running in Korea and sales to China, Vietnam, India and the US."));

  // ===== 9. PROCESS MAP =====
  s = content(SEC.sheet);
  head(s, "SHEET METAL SOLUTIONS", t("판금 공정별 솔루션", "Solutions by Process"), t("공정마다 최적의 글로벌 브랜드 장비를 조합해 제안합니다", "The right global brand for every step of your line"));
  const procs = [
    [t("레이저 절단", "Laser"), ["HSG", "MURATEC"]],
    [t("파이프 절단", "Tube"), ["HSG"]],
    [t("펀칭·복합", "Punch"), ["MURATEC"]],
    [t("절곡", "Bend"), ["ADIRA", "AIZAWA", "MURATEC"]],
    [t("패널 벤딩", "Panel bend"), ["KYOWA"]],
    [t("샤링", "Shear"), ["ADIRA", "AIZAWA"]],
    [t("용접", "Weld"), ["HSG"]],
    [t("자동화", "Automate"), ["MURATEC", "KYOWA", "AIZAWA"]],
  ];
  const pw = chevrons(s, procs.map((p) => p[0]), 0.6, 2.0, 12.13, 0.75, { fontSize: KO ? 12 : 10 });
  procs.forEach((p, i) => {
    const x = 0.6 + i * (pw + 0.06);
    p[1].forEach((b, k) => {
      s.addText(b, { shape: S.ROUNDED_RECTANGLE, rectRadius: 0.06, x: x + 0.12, y: 3.0 + k * 0.55, w: pw - 0.3, h: 0.42, fill: { color: C.accent4 }, line: { type: "none" }, color: C.accent1, bold: true, fontSize: 11, align: "center", valign: "middle", margin: 0 });
    });
  });
  card(s, 0.6, 5.0, 12.1, 1.55);
  [[t("한 파트너, 전 공정", "One partner, every step"), t("공정별 장비를 따로 찾을 필요 없이 라인 전체를 설계·공급", "Design and supply the whole line, not single machines")],
   [t("검증된 글로벌 브랜드", "Proven global brands"), t("일본·유럽·중국 제조사와 직접 파트너십", "Direct partnerships with makers in Japan, Europe and China")],
   [t("설치 후까지 책임", "Support after installation"), t("설치·시운전·교육·A/S까지 원스톱", "Installation, commissioning, training and after-sales service in one stop")]].forEach((m, i) => {
    const x = 0.9 + i * 4.0;
    txt(s, [{ text: m[0], options: { bold: true, fontSize: 16, color: C.accent1, breakLine: true } }, { text: m[1], options: { fontSize: 13 } }], { x, y: 5.0, w: 3.6, h: 1.55, valign: "middle" });
  });
  s.addNotes(t(
    "[발표] 판금 가공의 전 공정 — 레이저 절단, 파이프 절단, 펀칭, 절곡, 패널벤딩, 샤링, 용접, 자동화 — 에 대해 공정별로 가장 적합한 브랜드 장비를 조합해 제안합니다.\n\n[제작 메모] 공정–브랜드 매핑은 3/25본·기본안 기준. 판금 담당자 검수 필요(MURATEC 절곡기·자동화 범위 등).",
    "[Talk] For every step of sheet metal processing we combine the most suitable brand.\n\n[Production note] Process–brand mapping to be reviewed by the sheet metal team."));

  // ===== 10. BRAND LINEUP =====
  s = content(SEC.sheet);
  head(s, "BRAND LINEUP", t("취급 브랜드", "Brand Lineup"), t("일본 · 유럽 · 중국의 판금 전문 제조사", "Sheet metal specialists from Japan, Europe and China"));
  const brands = [
    ["MURATEC", t("일본", "Japan"), [t("터렛 펀치 프레스 (MOTORUM)", "Turret punch press (MOTORUM)"), t("파이버 레이저 (LS3015GC)", "Fiber laser (LS3015GC)"), t("듀얼 드라이브 절곡기 (BH)", "Dual-drive press brake (BH)")], t("1994년 세계 최초 서보 펀치 개발", "World's first servo punch press, 1994")],
    ["HSG", t("중국", "China"), [t("평판 파이버 레이저 (G · GT Pro)", "Flat fiber laser (G · GT Pro)"), t("파이프 레이저", "Tube laser"), t("핸드 레이저 용접기", "Handheld laser welder")], t("2006년 설립 · 15,000대+ 보급", "Founded 2006 · 15,000+ systems")],
    ["ADIRA", t("포르투갈", "Portugal"), [t("절곡기 PA·PH·PF 시리즈", "Press brakes PA/PH/PF"), t("절단기 GH·GV 시리즈", "Shears GH/GV")], t("유럽 절곡기·절단기 제조사", "European press brake & shear maker")],
    ["KYOWA", t("일본", "Japan"), [t("AC 서보 패널벤더 (KMP)", "AC servo panel bender (KMP)"), "KMP 13FT · 20FS · 25FT · 30FT"], t("전축 AC 서보 구동 · 유압 미사용", "All-axis AC servo, no hydraulics")],
    ["AIZAWA", t("일본", "Japan"), [t("프레스 브레이크", "Press brake"), t("정밀 절단·절곡", "Precision shearing & bending")], t("일본 판금기계 제조사", "Japanese sheet metal machinery")],
  ];
  brands.forEach((b, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.6 + col * 4.1 + (row === 1 ? 2.05 : 0), y = 1.95 + row * 2.42;
    card(s, x, y, 3.9, 2.25);
    txt(s, b[0], { x: x + 0.25, y: y + 0.2, w: 2.3, h: 0.45, fontSize: 20, bold: true, color: C.accent1 });
    s.addText(b[1], { shape: S.ROUNDED_RECTANGLE, rectRadius: 0.1, x: x + 2.65, y: y + 0.24, w: 1.0, h: 0.36, fill: { color: C.text2 }, color: C.background1, fontSize: 11, bold: true, align: "center", valign: "middle", margin: 0 });
    txt(s, b[2].map((p, k) => ({ text: p, options: { bullet: true, breakLine: k < b[2].length - 1 } })), { x: x + 0.25, y: y + 0.75, w: 3.45, h: 1.0, fontSize: 13 });
    txt(s, b[3], { x: x + 0.25, y: y + 1.8, w: 3.45, h: 0.3, fontSize: 12, bold: true, color: C.accent5 });
  });
  txt(s, t("출처: 회사 보유 MURATEC 종합카탈로그 · HSG GT Pro 카탈로그 · KYOWA KMP 카탈로그 · 브랜드 로고 원본 필요", "Sources: MURATEC general catalog, HSG GT Pro catalog, KYOWA KMP catalog (company copies)"), { x: 0.6, y: 6.66, w: 11, h: 0.25, fontSize: 10, color: C.accent5 });
  s.addNotes(t(
    "[발표] 판금기계사업부의 주력 브랜드입니다. 일본 MURATEC·KYOWA·AIZAWA, 포르투갈 ADIRA, 중국 HSG의 장비를 공급합니다. MURATEC은 1994년 세계 최초로 서보모터 펀치프레스(MOTORUM)를 개발했고, KYOWA 패널벤더는 전축을 AC 서보로 구동해 유압을 쓰지 않습니다.\n\n[제작 메모] 근거: MURATEC 종합카탈로그 p01-02(1994 세계 최초 서보 펀치), HSG GT Pro 카탈로그(2006년 설립, 15,000대+ 보급, 2018 브로셔 기준), KYOWA KMP 카탈로그(전축 AC 서보). 3/25 초안의 '유럽 60년'(ADIRA)·'100년·40,000대'(AIZAWA)·'세계 최초'(KYOWA)는 근거 자료가 없어 삭제. ADIRA·AIZAWA 카탈로그는 회사 카탈로그 폴더에 없음 — 확보 필요. 'DGDY'·KYOKKO·CONIC 표기 여부 결정.",
    "[Talk] Our sheet metal brands: MURATEC, KYOWA and AIZAWA from Japan, ADIRA from Portugal and HSG from China. MURATEC developed the world's first servo-motor punch press (MOTORUM) in 1994; KYOWA panel benders are driven by AC servos on every axis, with no hydraulics.\n\n[Production note] Sources: MURATEC general catalog p01-02, HSG GT Pro catalog (2018 figures), KYOWA KMP catalog. Unsupported claims from the March draft were removed. ADIRA and AIZAWA catalogs still needed."));

  // ===== 11. SPEC TABLE (sheet metal) =====
  s = content(SEC.sheet);
  head(s, "KEY SPECIFICATIONS", t("대표 장비 사양", "Key Equipment Specifications"), t("브랜드별 대표 모델 — 전체 사양은 QR 카탈로그로 제공", "Flagship models by brand — full specs via the QR catalog"));
  const H = (x) => ({ text: x, options: { bold: true, color: "FFFFFF", fill: { color: "005AAB" }, align: "center" } });
  const N = () => ({ text: NEED, options: { color: "B86E00", align: "center" } });
  const V = (x, o = {}) => ({ text: x, options: { align: "center", ...o } });
  const spec = [
    ["MURATEC", t("터렛 펀치 프레스", "Turret punch press"), "MOTORUM M3048TG", "1,250 × 2,500 mm", "300 kN (30 t)", "6.35", t("서보 램, 전력 85%↓", "Servo ram, −85% power")],
    ["MURATEC", t("파이버 레이저", "Fiber laser"), "LS3015GC", "3,070 × 1,550 mm", "2.5 / 4.0 / 6.0 kW", t("연강 25 · SUS ~14", "MS 25 · SS ~14"), t("고속 모드 15% 단축", "−15% cut time")],
    ["MURATEC", t("듀얼 드라이브 절곡기", "Dual-drive press brake"), "BH13530", t("절곡 길이 3,100 mm", "Bend length 3,100 mm"), "1,323 kN (135 t)", "—", t("볼스크류+유압 구동", "Ball screw + hydraulic")],
    ["HSG", t("평판 파이버 레이저", "Flat fiber laser"), "HS-G3015A (XLASE)", "3,000 × 1,500 mm", "4 kW (IPG)", t("연강 22 · SUS 12", "MS 22 · SS 12"), t("팔레트 교환 8초", "8 s pallet change")],
    ["HSG", t("파이프 레이저", "Tube laser")],
    ["KYOWA", t("AC 서보 패널벤더", "AC servo panel bender"), "KMP3023LFT", "3,050 × 1,525 mm", "588 kN (60 t)", t("SPCC 2.3 · SUS 1.5", "SPCC 2.3 · SS 1.5"), t("전축 AC 서보 구동", "All-axis AC servo")],
    ["ADIRA", t("절곡기", "Press brake"), "PA·PH·PF"],
    ["ADIRA", t("절단기", "Shear"), "GH·GV"],
    ["AIZAWA", t("프레스 브레이크", "Press brake")],
  ];
  const cell = (v) => (v == null ? N() : V(v));
  s.addTable([
    [H(t("브랜드", "Brand")), H(t("장비", "Machine")), H(t("대표 모델", "Model")), H(t("가공 범위", "Work range")), H(t("출력 / 톤수", "Power / tonnage")), H(t("최대 두께 (mm)", "Max. thick. (mm)")), H(t("특징", "Key feature"))],
    ...spec.map((r) => [V(r[0], { bold: true, color: "005AAB" }), V(r[1]), cell(r[2]), cell(r[3]), cell(r[4]), cell(r[5]), cell(r[6])]),
  ], { x: 0.6, y: 1.95, w: 12.13, colW: [1.15, 1.95, 1.85, 1.85, 1.75, 1.65, 1.93], rowH: 0.4, fontSize: 11, fontFace: FONT, color: "1A2230", valign: "middle", border: { type: "solid", pt: 0.75, color: "DCE6F2" }, fill: { color: "FFFFFF" } });
  txt(s, t("출처: 회사 보유 메이커 카탈로그 (MURATEC · HSG · KYOWA) · ADIRA · AIZAWA · HSG 파이프 레이저는 카탈로그 확보 후 입력", "Sources: makers' catalogs held by TOIN (MURATEC, HSG, KYOWA). ADIRA, AIZAWA and HSG tube laser to follow."), { x: 0.6, y: 6.66, w: 12, h: 0.25, fontSize: 10, color: C.accent5 });
  s.addNotes(t(
    "[발표] 브랜드별 대표 모델의 주요 사양입니다. 전체 라인업과 상세 사양은 마지막 페이지의 QR 카탈로그에서 확인하실 수 있습니다.\n\n[제작 메모] 출처(구글 드라이브 '카탈로그' 폴더): M3048TG_M3058TG.pdf 사양 페이지 · LS3015GC 카타로그(한).pdf p07 · 종합카다로그(MURATA).pdf p21-22(BH13530) · G3015A-4KW Machine 카탈로그.pdf · KMP-카달로그(영문).pdf. LS3015GC 최대 두께는 카탈로그 '가공 가능 두께' 그래프 판독치(6 kW 기준, 2.5/4.0/6.0 kW 순): 연강 약 19·25·25, 스테인리스 약 10·12·14, 알루미늄 약 10·12·14, 황동·구리 약 5·8·10 — 정확한 수치표가 있으면 교체. 남은 칸: M3048TG 최대 두께의 기준 재질, HSG 파이프 레이저(카탈로그 PDF 용량 초과로 미확인), ADIRA·AIZAWA(카탈로그 없음). BH13530은 현재 판매 모델(2026.11 확인). M2048HL(CO₂ 레이저·펀치 복합기)은 현재 판매 모델이 아니라 제외. Muratec-TOYO(TOYOKOKI) HPB·HYB·APB 절곡기와 HSS 절단기는 단종 모델이라 제외(2026.11 확인).",
    "[Talk] Key specifications of our flagship models; full specs are in the QR catalog.\n\n[Production note] Values from makers' catalogs in TOIN's Drive folder. LS3015GC thickness read from the catalog chart (6 kW: MS 25, SS ~14, Al ~14, brass/copper ~10). Still TBD: HSG tube laser, ADIRA, AIZAWA. BH13530 confirmed as a current model. M2048HL (CO₂ laser-punch combo) is not sold and was removed. Muratec-TOYO (TOYOKOKI) HPB/HYB/APB press brakes and HSS shears are discontinued and excluded."));

  // ===== 12. SHOWROOM & EXPORT =====
  s = content(SEC.sheet);
  head(s, "SHOWROOM & GLOBAL SALES", t("본사 전시장과 해외 판매", "Showroom & Global Sales"), t("직접 보고 시연한 뒤 결정하실 수 있습니다", "See it run before you decide"));
  img(s, "showroom.jpg", 0.6, 1.95, 7.3, 4.0, "cover");
  txt(s, [{ text: t("시흥 본사 전시장  ", "Siheung HQ showroom  "), options: { bold: true } }, { text: t("평판 레이저 · MURATEC 절곡기 · 협동로봇 상시 시연", "Flat laser · MURATEC press brake · cobot demos"), options: { color: C.accent5 } }], { x: 0.6, y: 6.1, w: 7.3, h: 0.5, fontSize: 13 });
  card(s, 8.25, 1.95, 4.48, 4.75);
  img(s, "network.png", 8.4, 2.05, 4.18, 2.6);
  txt(s, [{ text: t("해외 판매  ", "Export markets  "), options: { bold: true, color: C.accent1, fontSize: 16 } }, { text: t("장비 · 부품 공급", "equipment & parts"), options: { fontSize: 13, color: C.accent5 } }], { x: 8.5, y: 4.82, w: 4.0, h: 0.4, valign: "middle" });
  [t("중국", "China"), t("베트남", "Vietnam"), t("인도", "India"), t("미국", "USA")].forEach((c, i) => {
    s.addText(c, { shape: S.ROUNDED_RECTANGLE, rectRadius: 0.08, x: 8.5 + (i % 2) * 2.06, y: 5.35 + Math.floor(i / 2) * 0.62, w: 1.94, h: 0.5, fill: { color: C.background1 }, line: { color: "DCE6F2", width: 1 }, color: C.accent1, fontSize: 14, bold: true, align: "center", valign: "middle", margin: 0 });
  });
  s.addNotes(t(
    "[발표] 시흥 본사 전시장에는 평판 레이저, 무라타 절곡기, 협동로봇이 전시되어 있어 직접 시연을 보실 수 있습니다. 또한 중국·베트남·인도·미국 등 해외에도 장비와 부품을 판매하고 있습니다.\n\n[제작 메모] 전시장 최신 사진(현재 전시 장비 기준), 해외 납품 사례 1~2건 추가 권장.",
    "[Talk] Our Siheung showroom runs live demos of a flat laser, a MURATEC press brake and a cobot. We also export to China, Vietnam, India and the US."));

  // ===== 13. MACHINE TOOL DIVIDER =====
  divider(SEC.tool, "02", t("공작기계사업부", "Machine Tool Division"), "Machine Tool Division",
    t("MURATEC CNC 선반으로 자동차·항공·전자 부품의 정밀 가공과 생산성을 높입니다", "MURATEC CNC lathes that raise precision and productivity for automotive, aerospace and electronics parts"),
    [["1,800+", t("전국 가동 장비", "Units in operation")], ["6", t("MURATEC 시리즈", "MURATEC series")], [t("4개", "4"), t("주요 적용 산업", "Key industries")]],
    "lathe.png", "contain",
    t("[발표] 두 번째, 공작기계사업부입니다. 일본 MURATEC의 CNC 선반을 공급하며 현재 전국 1,800여 대가 가동 중입니다.\n\n[제작 메모] 공작기계 메이커는 MURATEC으로 확정(2026.11) — 3/25본의 'MIYANO(현 CITIZEN) 총판' 문구 삭제. 대표 장비 고해상도 사진 필요.",
      "[Talk] Our machine tool division supplies MURATEC CNC lathes, with 1,800+ units running in Korea."));

  // ===== 14. LATHE LINEUP =====
  s = content(SEC.tool);
  head(s, "MURATEC CNC LATHES", t("CNC 선반 라인업", "CNC Lathe Lineup"), t("소형 정밀 부품부터 샤프트류까지, 다품종·대량 생산에 맞춘 구성", "From small precision parts to shafts — configured for high-mix and high-volume production"));
  const lathes = [
    [t("트윈 스핀들 CNC 처커", "Twin-spindle CNC chucker"), "MW", "MW120", "8\" (ø210 mm)", "6,000", t("2주축 동시 가공 · 갠트리 로더 일체", "Two spindles + built-in gantry loader")],
    [t("싱글 스핀들 CNC 처커", "Single-spindle CNC chucker"), "MS", "MS20 · MS50 · MS60 · MS90 · MS100", "5\" – 10\"", null, t("부품 크기별 모델 선택", "Model range by part size")],
    [t("Y축 트윈 스핀들 복합기", "Twin spindle with Y-axis"), "MWR", "MWR120", null, null, t("Y축 ±60 mm · 회전공구 5.5 kW", "Y-axis ±60 mm, 5.5 kW live tool")],
    [t("회전공구 트윈 스핀들 터닝센터", "Twin-spindle live-tool turning center"), "MD", "MD120 · MD200", "8\" – 10\"", t("3,500 (MD200)", "3,500 (MD200)"), t("전 터렛 회전공구 · 갠트리 로더", "Live tool on every station + loader")],
    [t("CNC 샤프트 선반 (심압대)", "CNC shaft lathe (tailstock)"), "MJ", "MJ60 · MJ100 · MJ200", "8\" – 10\"", "3,500 – 6,000", t("샤프트류 자동화 · 유압 심압대", "Shaft work, hydraulic tailstock")],
    [t("대향 트윈 스핀들 복합기", "Opposed twin-spindle multitasker"), "MT", "MT12 · MT25 · MT100", "6\" – 10\"", "3,500 – 5,000", t("대향 2주축 · 2~3 터렛 복합 가공", "Opposed spindles, 2–3 turrets")],
  ];
  const lc = (v, o = {}) => (v == null ? N() : V(v, o));
  s.addTable([
    [H(t("장비 유형", "Type")), H(t("시리즈", "Series")), H(t("대표 모델", "Models")), H(t("최대 척", "Max. chuck")), H(t("최대 주축 회전수 (min⁻¹)", "Max. spindle speed (min⁻¹)")), H(t("특징", "Key feature"))],
    ...lathes.map((r) => [V(r[0], { bold: true, align: "left" }), V(r[1], { bold: true, color: "005AAB" }), lc(r[2]), lc(r[3]), lc(r[4]), lc(r[5], { align: "left" })]),
  ], { x: 0.6, y: 1.95, w: 12.13, colW: [2.85, 0.9, 2.55, 1.45, 1.7, 2.68], rowH: 0.6, fontSize: 12, fontFace: FONT, color: "1A2230", valign: "middle", border: { type: "solid", pt: 0.75, color: "DCE6F2" }, fill: { color: "FFFFFF" } });
  txt(s, t("출처: MURATEC 홈페이지 (muratec.com · muratec-usa.com) 기준 잠정치 · 선반 카탈로그 확보 후 최종 확인", "Source: MURATEC websites (muratec.com, muratec-usa.com), provisional until checked against the catalog"), { x: 0.6, y: 6.4, w: 12, h: 0.3, fontSize: 11, color: C.accent3, bold: true });
  s.addNotes(t(
    "[발표] MURATEC CNC 선반은 트윈 스핀들 처커(MW), 싱글 스핀들 처커(MS), Y축 복합기(MWR), 회전공구 터닝센터(MD), 샤프트 선반(MJ), 대향 트윈 스핀들 복합기(MT)로 구성되어 다품종 소량부터 대량 생산까지 대응합니다.\n\n[제작 메모] 선반 카탈로그가 회사 자료에 없어 MURATEC 홈페이지 검색 결과로 만든 잠정치(2026.11). 홈페이지 안에서도 버전별 수치가 다름: MW120 척 6\"/8\", 주축 3,000~6,000 min⁻¹; MJ100 터렛 10/12 스테이션. ML 시리즈는 확인되지 않아 제외(muratec-usa에 ML400 페이지 흔적만 있음). 최대 가공경은 홈페이지에 없어 '최대 척 크기'로 대체. 빈칸: MS 주축 회전수, MWR120 척·회전수. 공작기계사업부가 국내 판매 사양으로 확인 필요. 참고 페이지: muratec.com/mt/products/mw.html · mj.html · mt.html, muratec.net/mt/products/ms.html · md.html · mwr.html, muratec-usa.com 제품 페이지.",
    "[Talk] The MURATEC lathe lineup: twin-spindle chuckers (MW), single-spindle chuckers (MS), Y-axis twin spindle (MWR), live-tool turning centers (MD), shaft lathes (MJ) and opposed twin-spindle multitaskers (MT).\n\n[Production note] Provisional values from MURATEC websites (Nov 2026); no lathe catalog on file. Max chuck size replaces max turning diameter. Still TBD: MS spindle speed, MWR120 chuck and speed. ML series not confirmed and left out. Machine tool team to confirm against Korean sales specs."));

  // ===== 15. MACHINE TOOL CUSTOMERS =====
  s = content(SEC.tool);
  head(s, "KEY CLIENTS", t("공작기계 주요 고객", "Machine Tool Clients"), t("자동차 부품사를 중심으로 공정 개선과 불량률 최소화를 함께 합니다", "We help auto-parts makers improve processes and cut defect rates"));
  const logos = ["iljin", "ss", "oto", "phc_valeo", "samick_thk", "kcage", "busan", "kc_sanup", "kc_jeonggong", "kapec", "pha", "dic", "yuchang", "ss_mtech"];
  logos.forEach((lg, i) => {
    const col = i % 5, row = Math.floor(i / 5);
    const x = 0.6 + col * 1.5 + (row === 2 ? 0.75 : 0), y = 1.95 + row * 1.4;
    card(s, x, y, 1.35, 1.2, { fill: C.background1, line: "DCE6F2" });
    img(s, `logo_${lg}.png`, x + 0.12, y + 0.18, 1.11, 0.84);
  });
  txt(s, t("현대·기아 협력 부품사를 중심으로 MURATEC CNC 선반 공급 (로고 게재 동의 확인 필요)", "MURATEC CNC lathes supplied mainly to Hyundai-Kia parts makers"), { x: 0.6, y: 6.25, w: 7.35, h: 0.35, fontSize: 12, color: C.accent5 });
  card(s, 8.6, 1.95, 4.13, 2.6);
  txt(s, t("적용 산업", "Industries served"), { x: 8.85, y: 2.1, w: 3.7, h: 0.4, fontSize: 16, bold: true, color: C.accent1 });
  [["FaCar", t("자동차", "Automotive")], ["FaPlane", t("항공", "Aerospace")], ["FaMicrochip", t("전자", "Electronics")], ["FaCogs", t("일반기계", "Machinery")]].forEach((it, i) => {
    const x = 8.85 + (i % 2) * 1.88, y = 2.62 + Math.floor(i / 2) * 0.9;
    iconCircle(s, it[0], x, y, 0.62);
    txt(s, it[1], { x: x + 0.72, y, w: 1.1, h: 0.62, fontSize: 13, bold: true, valign: "middle" });
  });
  card(s, 8.6, 4.75, 4.13, 1.95, { fill: C.background2, line: C.accent3, dash: "dash" });
  txt(s, [{ text: t("공정 개선 사례", "Process improvement case"), options: { bold: true, fontSize: 16, color: C.accent1, breakLine: true } }, { text: t("고객사 · 적용 장비 · 사이클 타임 / 불량률 개선 수치", "Client · machine · cycle-time / defect-rate gains"), options: { fontSize: 13, breakLine: true } }, need()], { x: 8.85, y: 4.92, w: 3.7, h: 1.65 });
  s.addNotes(t(
    "[발표] 공작기계사업부는 현대·기아 협력 부품사를 중심으로 MURATEC CNC 선반을 공급하며, 공정 개선과 불량률 최소화로 생산성을 높이고 있습니다.\n\n[제작 메모] 고객사 로고는 2025.03본에서 가져옴 — 현재 거래처 기준으로 갱신, 로고 게재 동의 확인. 공정 개선 사례 1건(수치 포함) 작성.",
    "[Talk] We supply MURATEC lathes mainly to Hyundai-Kia parts makers.\n\n[Production note] Update client logos and confirm consent; add one improvement case with numbers."));

  // ===== 16. AUTO PARTS DIVIDER =====
  divider(SEC.auto, "03", t("자동차부품사업부", "Auto Parts Division"), "Auto Parts Division",
    t("대구 국가산업단지의 무인 자동화 공장에서 변속기 핵심 부품을 생산합니다", "Core transmission parts made in our automated plant in the Daegu National Industrial Complex"),
    [[t("5,000평", "16,529㎡"), t("공장 대지", "Plant site")], [t("2,815평", "9,306㎡"), t("건축 면적", "Building area")], [t("40명", "40"), t("생산 인력 (2025)", "Production staff (2025)")]],
    "daegu_aerial.jpg", "cover",
    t("[발표] 세 번째, 자동차부품사업부입니다. 2017년 대구 국가산업단지에 대지 5,000평 규모 공장을 신축해 로봇·갠트리 기반의 무인 자동화 생산 체계를 운영하고 있습니다.\n\n[제작 메모] 항공사진이 대구공장인지 확인. 인원·면적 최신화.",
      "[Talk] Our auto parts division runs an automated plant opened in 2017 in Daegu.\n\n[Production note] Confirm the aerial photo shows the Daegu plant."));

  // ===== 17. PLANT OVERVIEW =====
  s = content(SEC.auto);
  head(s, "DAEGU PLANT", t("대구 공장 개요", "Daegu Plant"), t("대구광역시 달성군 구지면 국가산단서로40길 28", "28 Gukgasandanseo-ro 40-gil, Guji-myeon, Dalseong-gun, Daegu"));
  card(s, 0.6, 1.95, 6.2, 4.75, { fill: C.background1, line: "DCE6F2" });
  img(s, KO ? "layout.png" : "layout_en.png", 0.75, 2.05, 5.9, 4.55);
  s.addTable([
    [H(t("구역", "Building")), H(t("1층", "1F")), H(t("2층", "2F")), H(t("3층", "3F"))],
    [V(t("B동", "Bldg B"), { bold: true }), V(t("2,238㎡", "2,238 m²")), V(t("430.8㎡", "430.8 m²")), V(t("252㎡", "252 m²"))],
    [V(t("C동", "Bldg C"), { bold: true }), V(t("2,238㎡", "2,238 m²")), V(t("430.8㎡", "430.8 m²")), V(t("252㎡", "252 m²"))],
    [V(t("D동", "Bldg D"), { bold: true }), N(), N(), N()],
  ], { x: 7.2, y: 1.95, w: 5.53, colW: [1.3, 1.45, 1.4, 1.38], rowH: 0.5, fontSize: 12, fontFace: FONT, color: "1A2230", valign: "middle", border: { type: "solid", pt: 0.75, color: "DCE6F2" }, fill: { color: "FFFFFF" } });
  [[t("5,000평", "16,529 m²"), t("총 대지면적 16,529㎡", "Total site area")], [t("2,815평", "9,306 m²"), t("건축 면적", "Building area")]].forEach((k, i) => {
    const x = 7.2 + i * 2.84;
    card(s, x, 4.15, 2.69, 1.15);
    txt(s, k[0], { x: x + 0.2, y: 4.27, w: 2.35, h: 0.5, fontSize: 18, bold: true, color: C.accent1 });
    txt(s, k[1], { x: x + 0.2, y: 4.8, w: 2.35, h: 0.35, fontSize: 12, color: C.accent5 });
  });
  img(s, "factory66.jpg", 7.2, 5.45, 5.53, 1.25, "cover");
  s.addNotes(t(
    "[발표] 대구 공장은 총 대지 5,000평, 건평 2,815평으로 B·C·D동에 생산라인을 최적 배치했습니다.\n\n[제작 메모] D동 면적 입력. 배치도는 2025.03본 이미지 — 최신 레이아웃(라인 구성) 반영본으로 교체 권장. 영문판 건축면적 9,306㎡는 2,815평 환산값.",
    "[Talk] The Daegu plant covers 16,529 m² with production lines arranged across buildings B, C and D.\n\n[Production note] Add building D area; 9,306 m² converted from 2,815 pyeong."));

  // ===== 18. PRODUCTS =====
  s = content(SEC.auto);
  head(s, "PRODUCTS", t("주요 생산 품목", "Main Products"), t("자동변속기·무단변속기용 정밀 선삭 부품", "Precision-turned parts for automatic and continuously variable transmissions"));
  card(s, 0.6, 1.95, 3.6, 4.75, { fill: C.background1, line: "DCE6F2" });
  img(s, "car.png", 0.75, 2.2, 3.3, 2.6);
  txt(s, [{ text: t("구동계 핵심 부품", "Drivetrain core parts"), options: { bold: true, color: C.accent1, fontSize: 16, breakLine: true } }, { text: t("경창산업·경창정공 등 현대·기아 변속기 부품 협력사에 납품", "Supplied to Hyundai-Kia transmission suppliers, including Kyungchang Industrial"), options: { fontSize: 13 } }], { x: 0.85, y: 5.0, w: 3.15, h: 1.55 });
  const prods = [["hub37r.jpg", "HUB 37R CLUTCH"], ["hub46.jpg", "HUB 46 CLUTCH"], ["gamma.jpg", t("감마 IVT 부품", "GAMMA IVT part")], ["plate_mid_sun_conn.jpg", "PLATE MID SUN CONN"], ["plate_sun_mid.jpg", "PLATE SUN MID"], [null, t("KAPPA CVT 부품", "KAPPA CVT part")]];
  prods.forEach((p, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 4.5 + col * 2.78, y = 1.95 + row * 2.42;
    if (p[0]) { img(s, p[0], x, y, 2.6, 1.8, "cover"); } else { photoSlot(s, x, y, 2.6, 1.8, t("품목 사진", "Product photo")); }
    txt(s, p[1], { x, y: y + 1.88, w: 2.6, h: 0.3, fontSize: 12, bold: true });
  });
  s.addNotes(t(
    "[발표] 자동변속기 클러치 허브, 감마 IVT 부품, 플레이트류 등 변속기 핵심 선삭 부품을 생산해 경창산업 등 현대·기아 협력사에 납품하고 있습니다.\n\n[제작 메모] 품목명이 자료마다 다름(3/25: HUB&SHAFT·KAPPA CVT·U/D BREAK / 2025.03: HUB 37R·46, 감마 IVT, PLATE류) — 현재 양산 품목으로 확정. 사진은 2025.03본 화면 캡처에서 잘라낸 것이라 원본 촬영 필요. 품목별 규격 범위·월 생산 능력 추가 권장.",
    "[Talk] We make clutch hubs, IVT parts and plates for transmissions.\n\n[Production note] Confirm current product names; reshoot product photos."));

  // ===== 19. AUTOMATED PRODUCTION =====
  s = content(SEC.auto);
  head(s, "AUTOMATED PRODUCTION", t("무인 자동화 생산 체계", "Automated Production"), t("로더 · 로봇 · 컨베이어로 소재 투입부터 출하까지 자동화", "Loaders, robots and conveyors automate everything from material input to shipping"));
  chevrons(s, [t("소재 투입", "Material input"), t("CNC 가공", "CNC machining"), t("로봇 이송", "Robot transfer"), t("검사·측정", "Inspection"), t("자동 출하", "Shipping")], 0.6, 1.95, 12.13, 0.65, { fontSize: 14 });
  ["factory68.jpg", "factory70.jpg", "factory72.jpg", "factory73.jpg"].forEach((f, i) => img(s, f, 0.6 + i * 2.1, 2.85, 1.95, 1.75, "cover"));
  txt(s, t("로봇·갠트리 자동화 라인 (대구 공장)", "Robot and gantry lines, Daegu plant"), { x: 0.6, y: 4.68, w: 8.2, h: 0.3, fontSize: 11, color: C.accent5 });
  card(s, 0.6, 5.15, 8.25, 1.3);
  [["FaRobot", t("전 공정 자동화", "Full automation"), t("로봇·갠트리 로더", "Robot & gantry loaders")], ["FaRulerHorizontal", "±0.01mm*", t("정밀도 관리", "Tolerance control")], ["FaClock", t("24시간", "24h"), t("연속 무인 가동*", "Lights-out operation*")]].forEach((k, i) => {
    const x = 0.8 + i * 2.75;
    iconCircle(s, k[0], x, 5.45, 0.7);
    txt(s, [{ text: k[1], options: { bold: true, fontSize: 17, color: C.accent1, breakLine: true } }, { text: k[2], options: { fontSize: 12, color: C.accent5 } }], { x: x + 0.82, y: 5.15, w: 1.85, h: 1.3, valign: "middle" });
  });
  card(s, 9.15, 2.85, 3.58, 3.85, { fill: C.background2, line: C.accent3, dash: "dash" });
  txt(s, [{ text: t("보유 설비", "Equipment list"), options: { bold: true, fontSize: 16, color: C.accent1, breakLine: true } },
    { text: t("CNC 선반  ○○대", "CNC lathes  ○○"), options: { bullet: true, breakLine: true } }, { text: t("연삭기  ○○대", "Grinders  ○○"), options: { bullet: true, breakLine: true } }, { text: t("로봇·갠트리  ○○기", "Robots / gantries  ○○"), options: { bullet: true, breakLine: true } }, { text: t("측정·검사 장비  ○○대", "Measuring equipment  ○○"), options: { bullet: true, breakLine: true } }, need()],
    { x: 9.4, y: 3.05, w: 3.1, h: 3.5, fontSize: 13, paraSpaceAfter: 4 });
  txt(s, t("* 수치 근거 확인 후 확정", "* To be confirmed"), { x: 0.6, y: 6.75, w: 5, h: 0.2, fontSize: 10, color: C.accent3 });
  s.addNotes(t(
    "[발표] 소재 투입부터 CNC 가공, 로봇 이송, 검사·측정, 출하까지 전 공정을 로더·로봇·컨베이어로 자동화했습니다.\n\n[제작 메모] 보유 설비 대수 입력. ±0.01mm(3/25본), 24시간 무인 가동 표현 근거 확인.",
    "[Talk] Every step from material input to shipping is automated.\n\n[Production note] Fill in the equipment list; verify starred claims."));

  // ===== 20. QUALITY & CERTS =====
  s = content(SEC.auto);
  head(s, "QUALITY & CERTIFICATIONS", t("품질 · 인증", "Quality & Certifications"), t("자동차 산업 품질경영 체계와 고객사 품질 인증", "Automotive quality management and customer quality certifications"));
  const certs = [["cert_iatf.png", "IATF 16949", t("자동차 품질경영 (LRQA)", "Automotive QMS (LRQA)")], ["cert_iso14001.png", "ISO 14001", t("환경경영시스템", "Environmental management")], ["cert_tsq.png", "T-SQ", t("현대트랜시스 협력사 품질", "Hyundai Transys supplier quality")], ["cert_psq.jpg", "P-SQ", t("협력사 품질보증", "Supplier quality assurance")], ["cert_ssq.jpg", "SSQ", t("SEMES 협력사 품질", "SEMES supplier quality")]];
  certs.forEach((c, i) => {
    const x = 0.6 + i * 2.46;
    card(s, x, 1.95, 2.3, 3.45, { fill: C.background2 });
    img(s, c[0], x + 0.25, 2.1, 1.8, 2.5, "contain", { shadow: shadow() });
    txt(s, c[1], { x: x + 0.1, y: 4.7, w: 2.1, h: 0.3, fontSize: 15, bold: true, color: C.accent1, align: "center" });
    txt(s, c[2], { x: x + 0.1, y: 5.0, w: 2.1, h: 0.35, fontSize: 11, color: C.accent5, align: "center" });
  });
  card(s, 0.6, 5.6, 12.1, 1.1);
  txt(s, [{ text: t("품질 관리 체계   ", "Quality system   "), options: { bold: true, color: C.accent1 } }, { text: t("수입검사 → 공정검사(계측) → 출하검사 → 이력 추적 관리", "Incoming → in-process (measurement) → outgoing inspection → full traceability"), options: {} }, { text: t("   (공정 흐름도 · 검사 장비 사진 추가)", "   (add flow chart and inspection photos)"), options: { color: C.accent3, bold: true } }], { x: 0.9, y: 5.6, w: 11.6, h: 1.1, fontSize: 14, valign: "middle" });
  s.addNotes(t(
    "[발표] IATF 16949, ISO 14001과 현대트랜시스 T-SQ, P-SQ, SEMES SSQ 등 고객사 품질 인증을 보유하고 있습니다.\n\n[제작 메모] 인증서는 2025.03본 이미지 — 유효기간 확인 및 최신 사본 스캔. ISO 9001 보유 시 추가. 검사 공정 흐름도 작성.",
    "[Talk] We hold IATF 16949, ISO 14001 and customer certifications including T-SQ, P-SQ and SSQ.\n\n[Production note] Check certificate validity; add ISO 9001 if held."));

  // ===== 21. SERVICE =====
  pres.addSection({ title: SEC.svc });
  s = content(SEC.svc);
  head(s, "CUSTOMER SUPPORT", t("고객지원 서비스", "Customer Support"), t("구매 이후가 진짜 시작입니다 — 장비 수명 전체를 책임집니다", "The sale is only the beginning — we look after the machine's whole life"));
  chevrons(s, [t("상담·컨설팅", "Consulting"), t("설치·시운전", "Installation"), t("운전 교육", "Training"), t("정기 점검", "Inspection"), t("A/S·부품", "Service & parts")], 0.6, 1.95, 12.13, 0.6, { fontSize: 14 });
  const svc = [
    ["T", t("기술 교육 & 컨설팅", "Training & consulting"), t("전문 인력 현장 파견, 조작·유지보수 실무 교육, 공정 최적화 컨설팅", "On-site engineers, hands-on operation and maintenance training, process consulting")],
    ["S", t("설치 & 시운전", "Installation & commissioning"), t("전문 엔지니어 현장 설치, 정밀 셋업·시운전, 초기 가동 안정화", "Installation, precise setup and commissioning, early-stage stabilization")],
    ["M", t("정기 점검 & 예방 정비", "Inspection & preventive care"), t("정기 방문 점검, 예방 정비로 고장률 최소화, 장비 수명 연장", "Scheduled visits and preventive maintenance to cut breakdowns")],
    ["A", t("신속 A/S & 부품 공급", "Fast service & parts"), t("전국 신속 출동(시흥·대구 2거점), 순정 부품 상시 보유, 원격 진단", "Nationwide dispatch from two bases, genuine parts in stock, remote diagnostics")],
  ];
  svc.forEach((v, i) => {
    const x = 0.6 + (i % 2) * 6.13, y = 2.85 + Math.floor(i / 2) * 1.55;
    card(s, x, y, 5.97, 1.4);
    s.addText(v[0], { shape: S.OVAL, x: x + 0.25, y: y + 0.3, w: 0.75, h: 0.75, fill: { color: C.accent1 }, color: C.background1, fontSize: 22, bold: true, align: "center", valign: "middle", margin: 0 });
    txt(s, [{ text: v[1], options: { bold: true, fontSize: 16, breakLine: true } }, { text: v[2], options: { fontSize: 13, color: C.accent5 } }], { x: x + 1.25, y: y + 0.2, w: 4.5, h: 1.05 });
  });
  card(s, 0.6, 6.0, 12.1, 0.7, { fill: C.text2 });
  txt(s, [{ text: t("A/S 접수   ", "Service hotline   "), options: { bold: true, color: C.accent6 } }, { text: t("본사 031-434-2031  ·  대구 053-382-7039  ·  이메일 ", "HQ +82-31-434-2031  ·  Daegu +82-53-382-7039  ·  E-mail "), options: { color: C.background1 } }, { text: NEED, options: { color: "F2B85C", bold: true } }], { x: 0.9, y: 6.0, w: 11.6, h: 0.7, fontSize: 14, valign: "middle" });
  s.addNotes(t(
    "[발표] 토인은 상담부터 설치, 교육, 정기 점검, A/S까지 장비의 전 생애를 책임집니다. 시흥과 대구 두 거점에서 전국에 신속 출동합니다.\n\n[제작 메모] A/S 대표 이메일, 평균 출동 시간 등 대응 기준(SLA)이 있으면 추가. 교육 커리큘럼 요약 별첨 검토.",
    "[Talk] From consulting to after-sales service, we look after the machine's whole life, dispatching engineers from Siheung and Daegu.\n\n[Production note] Add the service e-mail and a response-time standard if available."));

  // ===== 22. GLOBAL PARTNERS =====
  s = content(SEC.svc);
  head(s, "GLOBAL NETWORK", t("글로벌 파트너 네트워크", "Global Partner Network"), t("제조사와의 직접 파트너십으로 신뢰할 수 있는 장비와 부품을 공급합니다", "Direct partnerships with manufacturers for reliable equipment and parts"));
  card(s, 0.6, 1.95, 6.6, 4.75, { fill: C.background1, line: "DCE6F2" });
  img(s, "network.png", 0.75, 2.05, 6.3, 4.55);
  [[t("일본", "Japan"), "MURATEC · KYOWA · AIZAWA", t("KYOKKO · CONIC (부품)*", "KYOKKO · CONIC (parts)*")], [t("유럽 · 포르투갈", "Europe · Portugal"), "ADIRA", t("절곡기 · 절단기", "Press brakes · shears")], [t("중국", "China"), "HSG", t("파이버 레이저 · 파이프 레이저", "Fiber and tube lasers")]].forEach((r, i) => {
    const y = 1.95 + i * 1.32;
    card(s, 7.55, y, 5.18, 1.17);
    txt(s, [{ text: r[0], options: { bold: true, fontSize: 12, color: C.accent5, breakLine: true } }, { text: r[1], options: { bold: true, fontSize: 17, color: C.accent1, breakLine: true } }, { text: r[2], options: { fontSize: 12 } }], { x: 7.8, y: y + 0.12, w: 4.8, h: 0.98 });
  });
  card(s, 7.55, 5.91, 5.18, 0.79, { fill: C.text2 });
  txt(s, t("해외 직수입 · 기술 컨설팅 · 판매 · 설치 · A/S", "Import · Consulting · Sales · Installation · Service"), { x: 7.8, y: 5.91, w: 4.8, h: 0.79, fontSize: 14, bold: true, color: C.background1, valign: "middle" });
  txt(s, t("* 현재 거래 여부 확인", "* Confirm current partnership"), { x: 0.6, y: 6.75, w: 5, h: 0.2, fontSize: 10, color: C.accent3 });
  s.addNotes(t(
    "[발표] 일본·유럽·중국의 제조사와 직접 파트너십을 맺고 장비를 수입해, 컨설팅부터 A/S까지 제공합니다.\n\n[제작 메모] 네트워크 지도는 2025.03본 이미지 — 브랜드 확정 후 새로 제작 권장. 파트너사 로고 원본 수집.",
    "[Talk] We import directly from partners in Japan, Europe and China.\n\n[Production note] Redraw the network map after the brand list is final."));

  // ===== 23. BACK COVER =====
  pres.addSection({ title: SEC.end });
  s = pres.addSlide({ masterName: "COVER", sectionTitle: SEC.end });
  if (KO) img(s, "logo_full_white.png", 0.7, 0.6, 2.4, 2.4 * 303 / 1190);
  else { img(s, "logo_sym_white.png", 0.7, 0.55, 0.78, 0.65); txt(s, "TOIN", { x: 1.58, y: 0.55, w: 3, h: 0.65, fontSize: 30, bold: true, color: C.background1, valign: "middle" }); }
  txt(s, "LEAD TO A NEW PARADIGM", { x: 0.7, y: 1.55, w: 8, h: 0.6, fontSize: 30, bold: true, color: C.background1 });
  txt(s, t("감사합니다", "Thank you"), { x: 0.7, y: 2.2, w: 8, h: 0.5, fontSize: 20, color: C.accent6 });
  const offices = [
    [t("본사 HEAD OFFICE", "HEAD OFFICE"), t("경기도 시흥시 시화벤처로 152, 시화MTV 1사 501호 (정왕동 2646)", "#501, Sihwa MTV 1-sa, 152 Sihwaventure-ro, Siheung-si, Gyeonggi-do, Korea"), t("판금기계사업부 · 경영관리부", "Sheet Metal Division · Management"), t("T. 031-434-2031", "T. +82-31-434-2031")],
    [t("대구지사 DAEGU BRANCH", "DAEGU BRANCH"), t("대구광역시 달성군 구지면 국가산단서로40길 28 (응암리 1244-9)", "28 Gukgasandanseo-ro 40-gil, Guji-myeon, Dalseong-gun, Daegu, Korea"), t("공작기계사업부 · 자동차부품사업부", "Machine Tool · Auto Parts Divisions"), t("T. 053-382-7039", "T. +82-53-382-7039")],
  ];
  offices.forEach((o, i) => {
    const x = 0.7 + i * 4.6;
    card(s, x, 3.2, 4.4, 2.55, { fill: C.background1 });
    txt(s, [{ text: o[0], options: { bold: true, color: C.accent1, fontSize: 15, breakLine: true, paraSpaceAfter: 6 } }, { text: o[1], options: { fontSize: 13, breakLine: true, paraSpaceAfter: 6 } }, { text: o[2], options: { fontSize: 12, color: C.accent5, breakLine: true, paraSpaceAfter: 6 } }, { text: o[3], options: { fontSize: 13, bold: true } }], { x: x + 0.25, y: 3.4, w: 3.9, h: 2.2 });
  });
  card(s, 10.15, 3.2, 2.5, 2.55, { fill: C.background1 });
  s.addShape(S.RECTANGLE, { x: 10.6, y: 3.4, w: 1.6, h: 1.6, fill: { color: C.background2 }, line: { color: C.accent3, width: 1, dashType: "dash" } });
  txt(s, "QR", { x: 10.6, y: 3.4, w: 1.6, h: 1.6, fontSize: 20, bold: true, color: C.accent3, align: "center", valign: "middle" });
  txt(s, t("홈페이지 · 카탈로그", "Website · Catalog"), { x: 10.25, y: 5.15, w: 2.3, h: 0.4, fontSize: 12, align: "center", color: C.text1 });
  txt(s, [{ text: "www.toinint.com", options: { bold: true } }, { text: t("   |   주식회사 토인 · TOIN International Co., Ltd.", "   |   TOIN International Co., Ltd."), options: {} }], { x: 0.7, y: 6.35, w: 11.9, h: 0.4, fontSize: 13, color: C.background1 });
  s.addNotes(t(
    "[발표] 이상으로 (주)토인의 회사 소개를 마치겠습니다. 감사합니다.\n\n[제작 메모] QR 코드(홈페이지/카탈로그) 생성, 대표 이메일·팩스 추가, 오시는 길 약도는 배포용 PDF에 별첨 검토.",
    "[Talk] Thank you.\n\n[Production note] Generate the QR code; add e-mail and fax."));

  const out = path.join(__dirname, KO ? "TOIN_회사소개서_2027_KR.pptx" : "TOIN_Company_Profile_2027_EN.pptx");
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log("wrote", out);
}

(async () => { await build("ko"); await build("en"); })();
