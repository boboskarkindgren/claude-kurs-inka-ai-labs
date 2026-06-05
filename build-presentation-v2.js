const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const md = require("react-icons/md");

// Colors
const C = {
  orange:   "D95818",
  darkRed:  "640D02",
  offWhite: "FDFAF1",
  yellow:   "FFF2D9",
  offBlack: "2B2B2B",
  white:    "FFFFFF",
  midGray:  "8A8A8A",
  lightGray:"E0D8CC",
  dark4A:   "4A4A4A",
  dark3A:   "3A3A3A",
};

const makeShadow = () => ({ type: "outer", color: "000000", blur: 10, offset: 4, angle: 135, opacity: 0.12 });

async function iconPng(IconComp, color = "#FFFFFF", size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComp, { color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = "Reinholt Logistik × INKA AI Labs";

  // Pre-render icons
  const iCog      = await iconPng(fa.FaCog,          "#FFFFFF");
  const iHeadset  = await iconPng(fa.FaHeadset,      "#FFFFFF");
  const iChart    = await iconPng(fa.FaChartLine,    "#FFFFFF");
  const iCompass  = await iconPng(fa.FaCompass,      "#D95818");
  const iTools    = await iconPng(fa.FaTools,        "#FFFFFF");
  const iUsers    = await iconPng(fa.FaUsers,        "#D95818");
  const iCheck    = await iconPng(fa.FaCheckCircle,  "#D95818");
  const iArrow    = await iconPng(fa.FaArrowRight,   "#FFFFFF");
  const iTruck    = await iconPng(fa.FaTruck,        "#D95818");
  const iRocket   = await iconPng(fa.FaRocket,       "#FFFFFF");
  const iShield   = await iconPng(fa.FaShieldAlt,    "#FFFFFF");

  // ─────────────────────────────────────────────
  // SLIDE 1 – TITLE
  // ─────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offBlack };

    // Orange left panel
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 3.8, h: 5.625, fill: { color: C.orange }, line: { color: C.orange } });

    // INKA AI LABS label in panel
    s.addText("INKA AI LABS", {
      x: 0.25, y: 0.35, w: 3.3, h: 0.4,
      fontSize: 13, fontFace: "Calibri", bold: true, charSpacing: 4,
      color: C.white, margin: 0,
    });

    // Truck icon in panel
    s.addImage({ data: iTruck, x: 0.3, y: 1.0, w: 0.55, h: 0.55 });

    // Tagline in panel
    s.addText("Från AI-kaos\ntill verklig\naffärsnytta", {
      x: 0.25, y: 1.7, w: 3.3, h: 1.8,
      fontSize: 26, fontFace: "Calibri", bold: true,
      color: C.white, margin: 0,
    });

    // Decorative line in panel
    s.addShape(pres.shapes.RECTANGLE, { x: 0.25, y: 3.6, w: 1.2, h: 0.06, fill: { color: C.white }, line: { color: C.white } });

    s.addText("13 juni 2026", {
      x: 0.25, y: 3.8, w: 3.3, h: 0.35,
      fontSize: 13, fontFace: "Calibri", color: "FFD4B8", margin: 0,
    });

    // Right side – customer name
    s.addText("Reinholt\nLogistik AB", {
      x: 4.2, y: 0.9, w: 5.5, h: 2.6,
      fontSize: 54, fontFace: "Calibri", bold: true,
      color: C.white, margin: 0,
    });

    s.addText("Strategisk AI-plan för nästa tillväxtfas", {
      x: 4.2, y: 3.55, w: 5.5, h: 0.5,
      fontSize: 16, fontFace: "Calibri", italic: true,
      color: C.midGray, margin: 0,
    });

    s.addText("Presenterat av INKA AI Labs  ·  Anna Reinholt, COO", {
      x: 4.2, y: 4.9, w: 5.5, h: 0.35,
      fontSize: 11, fontFace: "Calibri", color: "606060", margin: 0,
    });

    s.addNotes("Välkomna! Vi på INKA AI Labs har lyssnat på vad ni berättat och ser direkt att Reinholt Logistik har rätt förutsättningar att ta AI från testfas till konkret affärsnytta. Idag visar vi hur vi tänker och hur vi tänker jobba med er.");
  }

  // ─────────────────────────────────────────────
  // SLIDE 2 – VI FÖRSTÅR ER SITUATION
  // ─────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };

    s.addText("Vi förstår er situation", {
      x: 0.45, y: 0.28, w: 9.1, h: 0.6,
      fontSize: 32, fontFace: "Calibri", bold: true, color: C.offBlack, margin: 0,
    });

    // Pain point cards (left column)
    const pains = [
      "AI används splittrat – ingen röd tråd",
      "Medarbetare provar verktyg på egen hand",
      "Svårt att prioritera vad som ger mest ROI",
      "Tillväxten kräver mer kapacitet med samma resurser",
      "Potentialen finns – men inte hur ni tar den",
    ];

    pains.forEach((text, i) => {
      const y = 1.05 + i * 0.82;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.45, y, w: 5.4, h: 0.65,
        fill: { color: C.white }, line: { color: C.lightGray, width: 1 },
        shadow: makeShadow(),
      });
      // Orange dot
      s.addShape(pres.shapes.OVAL, {
        x: 0.62, y: y + 0.22, w: 0.2, h: 0.2,
        fill: { color: C.orange }, line: { color: C.orange },
      });
      s.addText(text, {
        x: 0.95, y: y + 0.08, w: 4.8, h: 0.48,
        fontSize: 14, fontFace: "Calibri", color: C.dark4A,
        valign: "middle", margin: 0,
      });
    });

    // Right – two stat boxes stacked
    const stats = [
      { num: "280", label: "anställda\ni tillväxtfas" },
      { num: "2 nya", label: "storkundsavtal\nsom kräver skalbarhet" },
    ];

    stats.forEach((st, i) => {
      const y = 1.05 + i * 2.3;
      const isTop = i === 0;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 6.25, y, w: 3.3, h: 2.0,
        fill: { color: isTop ? C.orange : C.offBlack },
        line: { color: isTop ? C.orange : C.offBlack },
        shadow: makeShadow(),
      });
      s.addText(st.num, {
        x: 6.25, y: y + 0.2, w: 3.3, h: 1.0,
        fontSize: 56, fontFace: "Calibri", bold: true,
        color: C.white, align: "center", margin: 0,
      });
      s.addText(st.label, {
        x: 6.35, y: y + 1.15, w: 3.1, h: 0.7,
        fontSize: 13, fontFace: "Calibri",
        color: isTop ? "FFE8D8" : C.midGray, align: "center", margin: 0,
      });
    });

    s.addNotes("Ni berättade att AI-användningen idag är splittrad – och det är precis det vi ser hos de flesta företag i er fas. Nu handlar det om att sätta riktning. Tillväxten med de nya storkunderna skapar ett naturligt tryck att göra mer med samma resurser.");
  }

  // ─────────────────────────────────────────────
  // SLIDE 3 – VARFÖR DET ÄR SVÅRT
  // ─────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };

    s.addText("Varför det är svårt", {
      x: 0.45, y: 0.28, w: 9.1, h: 0.6,
      fontSize: 32, fontFace: "Calibri", bold: true, color: C.offBlack, margin: 0,
    });
    s.addText("De flesta företag fastnar i samma tre hinder", {
      x: 0.45, y: 0.85, w: 9.1, h: 0.35,
      fontSize: 15, fontFace: "Calibri", italic: true, color: C.midGray, margin: 0,
    });

    const obstacles = [
      {
        icon: iCompass, title: "För brett",
        body: "AI kan göra allt – men inte allt på en gång. Utan prioritering sprids energin tunt och ingenting blir klart.",
      },
      {
        icon: iChart, title: "För abstrakt",
        body: "Det är lätt att prata om AI-potential. Det svåra är att koppla det till era specifika processer och faktisk ROI.",
      },
      {
        icon: iUsers, title: "För ensamt",
        body: "Att driva AI-implementering internt utan rätt kompetens tar för lång tid och ger osäkra resultat.",
      },
    ];

    obstacles.forEach((obs, i) => {
      const x = 0.4 + i * 3.1;
      s.addShape(pres.shapes.RECTANGLE, {
        x, y: 1.35, w: 2.85, h: 3.95,
        fill: { color: C.white }, line: { color: C.lightGray, width: 1 },
        shadow: makeShadow(),
      });

      // Icon circle
      s.addShape(pres.shapes.OVAL, {
        x: x + 0.9, y: 1.55, w: 1.05, h: 1.05,
        fill: { color: C.orange }, line: { color: C.orange },
      });
      s.addImage({ data: obs.icon, x: x + 1.05, y: 1.7, w: 0.75, h: 0.75 });

      s.addText(obs.title, {
        x: x + 0.15, y: 2.75, w: 2.55, h: 0.55,
        fontSize: 20, fontFace: "Calibri", bold: true, color: C.offBlack,
        align: "center", margin: 0,
      });
      s.addText(obs.body, {
        x: x + 0.18, y: 3.35, w: 2.5, h: 1.8,
        fontSize: 13, fontFace: "Calibri", color: C.dark4A, margin: 0,
      });
    });

    s.addNotes("Tre hinder återkommer gång på gång: man försöker göra för mycket på en gång, man pratar abstrakt istället för konkret, och man försöker göra det ensam. Det vi gör hos INKA AI Labs är att ta er igenom dessa tre hinder strukturerat.");
  }

  // ─────────────────────────────────────────────
  // SLIDE 4 – VÅR METODIK
  // ─────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };

    s.addText("Vår metodik", {
      x: 0.45, y: 0.28, w: 9.1, h: 0.6,
      fontSize: 32, fontFace: "Calibri", bold: true, color: C.offBlack, margin: 0,
    });
    s.addText("Tre sammanlänkade pelare – från strategi till skalning", {
      x: 0.45, y: 0.85, w: 9.1, h: 0.35,
      fontSize: 15, fontFace: "Calibri", italic: true, color: C.midGray, margin: 0,
    });

    const pillars = [
      {
        num: "01", title: "AI-Strategi", sub: "Kartläggning och riktning",
        icon: iCompass, iconColor: "#D95818",
        points: ["AI-profil och mognadsbedömning", "Prioriterade use cases", "AI-policy och governance", "Mätbara mål och KPI:er"],
        highlight: false,
      },
      {
        num: "02", title: "AI-Toolbox", sub: "Implementering och verktyg",
        icon: iTools, iconColor: "#FFFFFF",
        points: ["Assistenter och chattbotar", "Automatisering av processer", "Innehållsproduktion med AI", "Analys och datainsikter"],
        highlight: true,
      },
      {
        num: "03", title: "AI-Konsult", sub: "Partnerskap och skalning",
        icon: iUsers, iconColor: "#D95818",
        points: ["Löpande strategisk rådgivning", "Utbildning och förankring", "Iterativ förbättring", "Skalning av vad som fungerar"],
        highlight: false,
      },
    ];

    pillars.forEach((p, i) => {
      const x = 0.4 + i * 3.1;
      const bg = p.highlight ? C.orange : C.white;
      const tc = p.highlight ? C.white : C.offBlack;
      const bc = p.highlight ? "FFE8D8" : C.dark4A;
      const sc = p.highlight ? "FFD4B8" : C.midGray;
      const nc = p.highlight ? "FFD4B8" : C.orange;

      s.addShape(pres.shapes.RECTANGLE, {
        x, y: 1.35, w: 2.85, h: 4.0,
        fill: { color: bg },
        line: { color: p.highlight ? C.orange : C.lightGray, width: 1 },
        shadow: makeShadow(),
      });

      s.addText(p.num, {
        x: x + 0.18, y: 1.52, w: 1.5, h: 0.32,
        fontSize: 11, fontFace: "Calibri", bold: true, color: nc, margin: 0,
      });
      s.addText(p.title, {
        x: x + 0.18, y: 1.85, w: 2.5, h: 0.55,
        fontSize: 22, fontFace: "Calibri", bold: true, color: tc, margin: 0,
      });
      s.addText(p.sub, {
        x: x + 0.18, y: 2.38, w: 2.5, h: 0.35,
        fontSize: 12, fontFace: "Calibri", italic: true, color: sc, margin: 0,
      });

      s.addText(
        p.points.map((pt, idx) => ({ text: pt, options: { bullet: true, breakLine: idx < p.points.length - 1 } })),
        {
          x: x + 0.18, y: 2.9, w: 2.5, h: 2.2,
          fontSize: 12, fontFace: "Calibri", color: bc, paraSpaceAfter: 6, margin: 0,
        }
      );
    });

    s.addNotes("Vår metodik bygger på tre pelare som hänger ihop. Vi börjar alltid med strategin – utan riktning blir verktyg bara leksaker. AI-Toolbox är där vi implementerar konkreta lösningar. Och AI-Konsult är det löpande partnerskapet som säkerställer att ni faktiskt skalar det som fungerar.");
  }

  // ─────────────────────────────────────────────
  // SLIDE 5 – VAD VI FÖRESLÅR
  // ─────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };

    s.addText("Vad vi föreslår för Reinholt", {
      x: 0.45, y: 0.28, w: 9.1, h: 0.6,
      fontSize: 32, fontFace: "Calibri", bold: true, color: C.offBlack, margin: 0,
    });
    s.addText("Tre prioriterade use cases – konkreta, lönsamma, genomförbara", {
      x: 0.45, y: 0.85, w: 9.1, h: 0.35,
      fontSize: 15, fontFace: "Calibri", italic: true, color: C.midGray, margin: 0,
    });

    const cases = [
      {
        num: "1", icon: iCog, title: "Admin-automatisering",
        desc: "Automatisera repetitiva uppgifter: orderhantering, fakturering, intern kommunikation och rapportering.",
        impact: "15–25% tidsbesparring per medarbetare",
      },
      {
        num: "2", icon: iHeadset, title: "Kundservice-AI",
        desc: "AI-driven ärendehantering som svarar på vanliga frågor, eskalerar rätt ärenden och minskar svarstider.",
        impact: "40–60% av ärenden hanteras automatiskt",
      },
      {
        num: "3", icon: iChart, title: "Resursplanering",
        desc: "AI-stöd för att optimera bemanning och kapacitet baserat på historiska mönster och prognoser.",
        impact: "Bättre kapacitetsutnyttjande inför tillväxt",
      },
    ];

    cases.forEach((uc, i) => {
      const x = 0.4 + i * 3.1;

      s.addShape(pres.shapes.RECTANGLE, {
        x, y: 1.35, w: 2.85, h: 4.0,
        fill: { color: C.white }, line: { color: C.lightGray, width: 1 },
        shadow: makeShadow(),
      });

      // Icon circle
      s.addShape(pres.shapes.OVAL, {
        x: x + 0.9, y: 1.52, w: 1.05, h: 1.05,
        fill: { color: C.orange }, line: { color: C.orange },
      });
      s.addImage({ data: uc.icon, x: x + 1.05, y: 1.67, w: 0.75, h: 0.75 });

      s.addText(uc.title, {
        x: x + 0.15, y: 2.72, w: 2.55, h: 0.52,
        fontSize: 17, fontFace: "Calibri", bold: true,
        color: C.offBlack, align: "center", margin: 0,
      });
      s.addText(uc.desc, {
        x: x + 0.18, y: 3.3, w: 2.5, h: 1.5,
        fontSize: 13, fontFace: "Calibri", color: C.dark4A, margin: 0,
      });

      // Impact box
      s.addShape(pres.shapes.RECTANGLE, {
        x: x + 0.1, y: 4.85, w: 2.65, h: 0.38,
        fill: { color: C.yellow }, line: { color: "E8D8B0", width: 1 },
      });
      s.addText(uc.impact, {
        x: x + 0.15, y: 4.87, w: 2.55, h: 0.34,
        fontSize: 11, fontFace: "Calibri", bold: true,
        color: C.darkRed, align: "center", margin: 0,
      });
    });

    s.addNotes("Vi har tittat specifikt på er verksamhet och identifierat tre use cases som passar er fas och era prioriteringar. Admin-automatisering ger snabbast synliga resultat. Kundservice-AI hanterar ökad kundvolym utan nyrekrytering. Och resursplanering är extra relevant nu när ni skalar.");
  }

  // ─────────────────────────────────────────────
  // SLIDE 6 – REFERENSCASE
  // ─────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };

    s.addText("Så här kan det se ut i praktiken", {
      x: 0.45, y: 0.28, w: 9.1, h: 0.6,
      fontSize: 32, fontFace: "Calibri", bold: true, color: C.offBlack, margin: 0,
    });

    // Left – case card
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.45, y: 1.05, w: 5.4, h: 4.28,
      fill: { color: C.white }, line: { color: C.lightGray, width: 1 },
      shadow: makeShadow(),
    });

    // Orange header strip in card
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.45, y: 1.05, w: 5.4, h: 0.55,
      fill: { color: C.orange }, line: { color: C.orange },
    });
    s.addText("Liknande uppdrag: Distributions- och e-handelslogistik", {
      x: 0.6, y: 1.1, w: 5.1, h: 0.42,
      fontSize: 13, fontFace: "Calibri", bold: true, color: C.white, margin: 0,
    });

    const caseRows = [
      { label: "Situation", text: "Medelstort logistikbolag med hög administrativ börda och ökande kundvolym." },
      { label: "Uppdrag", text: "AI-strategi + implementering av automatiserad ärendehantering och kundservice-chattbot." },
      { label: "Resultat", text: "52% av kundärenden hanterades automatiskt inom 3 månader. 2 heltidstjänster frigjordes för mer kvalificerat arbete." },
    ];

    caseRows.forEach((row, i) => {
      const y = 1.78 + i * 1.15;
      s.addText(row.label, {
        x: 0.62, y, w: 1.2, h: 0.35,
        fontSize: 12, fontFace: "Calibri", bold: true, color: C.orange, margin: 0,
      });
      s.addText(row.text, {
        x: 0.62, y: y + 0.32, w: 5.1, h: 0.72,
        fontSize: 13, fontFace: "Calibri", color: C.dark4A, margin: 0,
      });
    });

    // Right – result stats
    const results = [
      { num: "3 mån", label: "till första mätbara resultat", highlight: false },
      { num: "52%", label: "av ärenden automatiserat", highlight: true },
      { num: "2 tjänster", label: "frigjordes för värdeskapande arbete", highlight: false },
    ];

    results.forEach((r, i) => {
      const y = 1.05 + i * 1.43;
      const bg = r.highlight ? C.orange : C.white;
      const nc = r.highlight ? C.white : C.orange;
      const lc = r.highlight ? "FFE8D8" : C.midGray;

      s.addShape(pres.shapes.RECTANGLE, {
        x: 6.2, y, w: 3.35, h: 1.25,
        fill: { color: bg },
        line: { color: r.highlight ? C.orange : C.lightGray, width: 1 },
        shadow: makeShadow(),
      });
      s.addText(r.num, {
        x: 6.2, y: y + 0.08, w: 3.35, h: 0.7,
        fontSize: 32, fontFace: "Calibri", bold: true,
        color: nc, align: "center", margin: 0,
      });
      s.addText(r.label, {
        x: 6.3, y: y + 0.78, w: 3.15, h: 0.38,
        fontSize: 12, fontFace: "Calibri",
        color: lc, align: "center", margin: 0,
      });
    });

    s.addNotes("Vi vill inte bara presentera teorier – vi vill visa vad som faktiskt är möjligt. Det här exemplet från ett liknande uppdrag visar att vi inom 3 månader kan leverera konkreta, mätbara resultat.");
  }

  // ─────────────────────────────────────────────
  // SLIDE 7 – TIDSPLAN
  // ─────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };

    s.addText("Tidsplan", {
      x: 0.45, y: 0.28, w: 9.1, h: 0.6,
      fontSize: 32, fontFace: "Calibri", bold: true, color: C.offBlack, margin: 0,
    });
    s.addText("Realistisk i tre faser – vi börjar smalt och skalar brett", {
      x: 0.45, y: 0.85, w: 9.1, h: 0.35,
      fontSize: 15, fontFace: "Calibri", italic: true, color: C.midGray, margin: 0,
    });

    // Connecting line
    s.addShape(pres.shapes.RECTANGLE, {
      x: 1.35, y: 1.88, w: 7.3, h: 0.06,
      fill: { color: C.lightGray }, line: { color: C.lightGray },
    });

    const phases = [
      { num: "Fas 1", period: "Augusti 2026", title: "Kartläggning", active: false, points: ["AI-mognadsanalys", "Processgenomgång", "Prioritering av use cases", "Beslutsunderlag"] },
      { num: "Fas 2", period: "Sep–Okt 2026", title: "Pilotprojekt", active: true,  points: ["Implementering use case #1", "Intern utbildning", "Mätning och uppföljning", "Justering och förbättring"] },
      { num: "Fas 3", period: "Nov 2026 –",   title: "Skalning",     active: false, points: ["Rulla ut fler use cases", "Bygga intern AI-kompetens", "Löpande partnerskap", "Kontinuerlig optimering"] },
    ];

    phases.forEach((ph, i) => {
      const x = 0.4 + i * 3.1;
      const bg = ph.active ? C.orange : C.white;
      const tc = ph.active ? C.white : C.offBlack;
      const bc = ph.active ? "FFE8D8" : C.dark4A;
      const sc = ph.active ? "FFD4B8" : C.midGray;
      const nc = ph.active ? "FFD4B8" : C.orange;

      // Timeline dot
      s.addShape(pres.shapes.OVAL, {
        x: x + 1.15, y: 1.72, w: 0.36, h: 0.36,
        fill: { color: ph.active ? C.orange : C.lightGray },
        line: { color: ph.active ? C.orange : C.lightGray },
      });

      s.addShape(pres.shapes.RECTANGLE, {
        x, y: 2.18, w: 2.85, h: 3.18,
        fill: { color: bg },
        line: { color: ph.active ? C.orange : C.lightGray, width: 1 },
        shadow: makeShadow(),
      });

      s.addText(ph.num, { x: x + 0.15, y: 2.3, w: 2.55, h: 0.28, fontSize: 11, fontFace: "Calibri", bold: true, color: nc, margin: 0 });
      s.addText(ph.period, { x: x + 0.15, y: 2.55, w: 2.55, h: 0.28, fontSize: 11, fontFace: "Calibri", color: sc, margin: 0 });
      s.addText(ph.title, { x: x + 0.15, y: 2.85, w: 2.55, h: 0.52, fontSize: 20, fontFace: "Calibri", bold: true, color: tc, margin: 0 });

      s.addText(
        ph.points.map((pt, idx) => ({ text: pt, options: { bullet: true, breakLine: idx < ph.points.length - 1 } })),
        { x: x + 0.15, y: 3.45, w: 2.55, h: 1.75, fontSize: 12, fontFace: "Calibri", color: bc, paraSpaceAfter: 5, margin: 0 }
      );
    });

    s.addNotes("Tidsplanen är medvetet realistisk. Kartläggning i augusti – fokuserat arbete där vi förstår er verksamhet. Pilotprojekt under hösten – ett tydligt avgränsat use case med mätbara resultat. Fungerar det, skalar vi brett från november.");
  }

  // ─────────────────────────────────────────────
  // SLIDE 8 – VAD NI KAN FÖRVÄNTA ER
  // ─────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };

    s.addText("Vad ni kan förvänta er", {
      x: 0.45, y: 0.28, w: 9.1, h: 0.6,
      fontSize: 32, fontFace: "Calibri", bold: true, color: C.offBlack, margin: 0,
    });

    // Top stat callouts
    const stats = [
      { num: "3–6 mån", label: "till mätbara resultat från piloten" },
      { num: "20–40%", label: "effektivitetsökning i automatiserade processer" },
      { num: "100%", label: "transparens och kontroll under hela resan" },
    ];

    stats.forEach((st, i) => {
      const x = 0.4 + i * 3.1;
      s.addShape(pres.shapes.RECTANGLE, {
        x, y: 1.05, w: 2.85, h: 1.52,
        fill: { color: C.yellow }, line: { color: "E0CC9A", width: 1 },
      });
      s.addText(st.num, {
        x, y: 1.1, w: 2.85, h: 0.82,
        fontSize: 28, fontFace: "Calibri", bold: true,
        color: C.orange, align: "center", margin: 0,
      });
      s.addText(st.label, {
        x: x + 0.1, y: 1.9, w: 2.65, h: 0.58,
        fontSize: 11, fontFace: "Calibri",
        color: C.offBlack, align: "center", margin: 0,
      });
    });

    // Bottom left – included
    s.addText("Vad ingår i samarbetet:", {
      x: 0.45, y: 2.82, w: 4.5, h: 0.4,
      fontSize: 16, fontFace: "Calibri", bold: true, color: C.offBlack, margin: 0,
    });
    s.addText(
      [
        { text: "En dedikerad AI-konsult som känner er verksamhet",  options: { bullet: true, breakLine: true } },
        { text: "Tydlig roadmap med prioriterade use cases",         options: { bullet: true, breakLine: true } },
        { text: "Implementering och utbildning i er organisation",   options: { bullet: true, breakLine: true } },
        { text: "Löpande mätning och rapportering av resultat",      options: { bullet: true } },
      ],
      {
        x: 0.45, y: 3.28, w: 4.5, h: 2.1,
        fontSize: 14, fontFace: "Calibri", color: C.dark4A, paraSpaceAfter: 10, margin: 0,
      }
    );

    // Bottom right – quote
    s.addShape(pres.shapes.RECTANGLE, {
      x: 5.3, y: 2.78, w: 4.25, h: 2.6,
      fill: { color: C.offBlack }, line: { color: C.offBlack },
      shadow: makeShadow(),
    });
    s.addText(
      "“De flesta företag har börjat använda AI – men väldigt få har omvandlat det till verklig affärsnytta.”",
      {
        x: 5.55, y: 2.98, w: 3.8, h: 1.8,
        fontSize: 15, fontFace: "Calibri", italic: true, color: C.white, margin: 0,
      }
    );
    s.addText("– INKA AI Labs", {
      x: 5.55, y: 4.82, w: 3.8, h: 0.3,
      fontSize: 12, fontFace: "Calibri", bold: true, color: C.orange, margin: 0,
    });

    s.addNotes("Vi lovar inte magi – vi lovar ett strukturerat samarbete med tydliga leverabler och mätbara mål. 3-6 månader till resultat är realistiskt. 20-40% effektivitetsökning är vad vi sett i liknande uppdrag. Och 100% transparens är ett löfte om att ni alltid vet vad vi gör och varför.");
  }

  // ─────────────────────────────────────────────
  // SLIDE 9 – NÄSTA STEG
  // ─────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offBlack };

    // Orange left accent
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 0.12, h: 5.625,
      fill: { color: C.orange }, line: { color: C.orange },
    });

    s.addText("Nästa steg", {
      x: 0.35, y: 0.3, w: 9.3, h: 0.68,
      fontSize: 36, fontFace: "Calibri", bold: true, color: C.white, margin: 0,
    });
    s.addText("Vad händer om vi bestämmer oss för att gå vidare?", {
      x: 0.35, y: 0.95, w: 8.0, h: 0.38,
      fontSize: 15, fontFace: "Calibri", italic: true, color: C.midGray, margin: 0,
    });

    const steps = [
      "Ni bekräftar intresset och vi bokar ett uppstartsmöte (ca 2 timmar)",
      "Vi skickar en enkel behovsanalys som ni fyller i inför mötet",
      "Vi presenterar ett konkret uppdragsförslag med scope, budget och tidsplan",
      "Kartläggningsfasen startar – senast 1 augusti 2026",
    ];

    steps.forEach((st, i) => {
      const y = 1.55 + i * 0.95;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.35, y, w: 0.58, h: 0.58,
        fill: { color: C.orange }, line: { color: C.orange },
      });
      s.addText(`0${i + 1}`, {
        x: 0.35, y, w: 0.58, h: 0.58,
        fontSize: 16, fontFace: "Calibri", bold: true,
        color: C.white, align: "center", valign: "middle", margin: 0,
      });
      s.addImage({ data: iArrow, x: 1.08, y: y + 0.14, w: 0.3, h: 0.3 });
      s.addText(st, {
        x: 1.5, y: y + 0.06, w: 8.1, h: 0.48,
        fontSize: 15, fontFace: "Calibri", color: C.white, valign: "middle", margin: 0,
      });
    });

    s.addNotes("Nu vill vi veta: ser ni samma potential som vi ser? Om svaret är ja, är nästa steg väldigt konkret och utan risk. Ingen lång upphandling, inget abstrakt – bara konkret arbete.");
  }

  // ─────────────────────────────────────────────
  // SLIDE 10 – OM INKA AI LABS
  // ─────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };

    s.addText("Om INKA AI Labs", {
      x: 0.45, y: 0.28, w: 9.1, h: 0.6,
      fontSize: 32, fontFace: "Calibri", bold: true, color: C.offBlack, margin: 0,
    });
    s.addText("AI-byrå i Malmö – del av INKA Interactive", {
      x: 0.45, y: 0.85, w: 5.5, h: 0.38,
      fontSize: 16, fontFace: "Calibri", bold: true, color: C.orange, margin: 0,
    });

    s.addText(
      "INKA AI Labs hjälper medelstora till stora svenska företag att omsätta AI till verklig affärsnytta – inte bara testa, utan faktiskt implementera och skala.\n\nVi bygger på INKA Interactives 20+ år av digital erfarenhet kombinerat med AI-spetskompetens – ett helhetsperspektiv från strategi till implementation som de flesta AI-bolag saknar.",
      {
        x: 0.45, y: 1.35, w: 5.5, h: 2.45,
        fontSize: 14, fontFace: "Calibri", color: C.dark4A, margin: 0,
      }
    );

    s.addText("Nuvarande kunder:", {
      x: 0.45, y: 3.92, w: 2.0, h: 0.32,
      fontSize: 13, fontFace: "Calibri", bold: true, color: C.offBlack, margin: 0,
    });
    s.addText("HydX  ·  Forsen  ·  Jape  ·  Classicum", {
      x: 0.45, y: 4.28, w: 5.5, h: 0.32,
      fontSize: 13, fontFace: "Calibri", color: C.midGray, margin: 0,
    });

    // Right – fact cards
    const facts = [
      { num: "20+",   label: "års digital erfarenhet" },
      { num: "3",     label: "tjänstepelare i helhet" },
      { num: "Malmö", label: "lokal förankring, nationell räckvidd" },
    ];

    facts.forEach((f, i) => {
      const y = 1.05 + i * 1.52;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 6.3, y, w: 3.25, h: 1.3,
        fill: { color: C.white }, line: { color: C.lightGray, width: 1 },
        shadow: makeShadow(),
      });
      s.addShape(pres.shapes.RECTANGLE, {
        x: 6.3, y, w: 0.1, h: 1.3,
        fill: { color: C.orange }, line: { color: C.orange },
      });
      s.addText(f.num, {
        x: 6.52, y: y + 0.1, w: 2.95, h: 0.68,
        fontSize: 30, fontFace: "Calibri", bold: true, color: C.orange, margin: 0,
      });
      s.addText(f.label, {
        x: 6.52, y: y + 0.75, w: 2.95, h: 0.45,
        fontSize: 12, fontFace: "Calibri", color: C.midGray, margin: 0,
      });
    });

    s.addNotes("INKA AI Labs är inte ett nystartat bolag. Vi bygger på 20 års digital erfarenhet och det märks i hur vi jobbar. Vi är inte här för att sälja er det senaste verktyget – vi är här för att hjälpa er bygga en varaktig AI-förmåga.");
  }

  // ─────────────────────────────────────────────
  // SLIDE 11 – TACK / KONTAKT
  // ─────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offBlack };

    // Orange left panel
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 3.8, h: 5.625,
      fill: { color: C.orange }, line: { color: C.orange },
    });

    s.addText("Från\ntanke till\nhandling.", {
      x: 0.25, y: 0.7, w: 3.3, h: 2.8,
      fontSize: 34, fontFace: "Calibri", bold: true, italic: true,
      color: C.white, margin: 0,
    });

    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.25, y: 3.65, w: 1.2, h: 0.06,
      fill: { color: C.white }, line: { color: C.white },
    });

    s.addText("INKA AI Labs", {
      x: 0.25, y: 3.85, w: 3.3, h: 0.38,
      fontSize: 14, fontFace: "Calibri", bold: true, color: C.white, margin: 0,
    });
    s.addText("Malmö, Sverige", {
      x: 0.25, y: 4.22, w: 3.3, h: 0.3,
      fontSize: 12, fontFace: "Calibri", color: "FFE8D8", margin: 0,
    });

    // Right – thank you + contact
    s.addText("Tack!", {
      x: 4.2, y: 0.65, w: 5.5, h: 1.4,
      fontSize: 72, fontFace: "Calibri", bold: true, color: C.white, margin: 0,
    });
    s.addText("Vi ser fram emot att hjälpa\nReinholt Logistik ta nästa steg.", {
      x: 4.2, y: 2.1, w: 5.5, h: 0.9,
      fontSize: 18, fontFace: "Calibri", italic: true, color: C.midGray, margin: 0,
    });

    // Contact box
    s.addShape(pres.shapes.RECTANGLE, {
      x: 4.2, y: 3.2, w: 5.5, h: 2.1,
      fill: { color: "3A3A3A" }, line: { color: "3A3A3A" },
    });
    s.addText("Kontakta oss", {
      x: 4.4, y: 3.35, w: 5.1, h: 0.38,
      fontSize: 13, fontFace: "Calibri", bold: true, color: C.orange, margin: 0,
    });
    s.addText("info@inkaailabs.se\ninkaailabs.se", {
      x: 4.4, y: 3.78, w: 5.1, h: 1.35,
      fontSize: 15, fontFace: "Calibri", color: C.white, margin: 0,
    });

    s.addNotes("Tack för er tid idag. Vi hoppas att ni känner er en bit längre på vägen mot att ta AI från testfas till verklig affärsnytta. Vi är redo att starta – frågan är om ni är det.");
  }

  // ─────────────────────────────────────────────
  // SAVE
  // ─────────────────────────────────────────────
  const outPath = "C:\\Users\\kindg\\OneDrive\\Dokument\\Claude kurs - INKA AI Labs\\reinholt-logistik-v2.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log("Presentation sparad:", outPath);
}

build().catch(err => { console.error(err); process.exit(1); });
