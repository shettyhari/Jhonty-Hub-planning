/**
 * JONTYS HUB – LUXURY MEN'S FASHION GROWTH BLUEPRINT
 * Editable 14-Slide Widescreen (16:9) Deck Generator via PptxGenJS
 *
 * Theme: Quiet-Luxury Men's Tailoring & Editorial Atelier
 * Palette:
 *   Obsidian Black:     #0F1115
 *   Midnight Velvet:    #141923
 *   Warm Atelier Ivory: #F6F2EA
 *   Champagne Gold:     #C9A96E
 *   Antique Bronze:     #9E7F4B
 *   Muted Stone:        #6B7280
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

const ASSETS_DIR = path.join(__dirname, 'assets');
const OUTPUT_PPTX = path.join(__dirname, 'Jontys_Hub_Premium_Growth_Strategy.pptx');

const COLORS = {
  obsidian: '0F1115',
  obsidianCard: '171A21',
  velvet: '141923',
  velvetCard: '1C2230',
  ivory: 'F6F2EA',
  ivoryCard: 'FFFFFF',
  ivorySoft: 'EDE7DB',
  gold: 'C9A96E',
  bronze: '9E7F4B',
  stone: '5E6573',
  lightMuted: '9CA3AF',
  lightText: 'E5E7EB',
  borderLight: 'DFD8CA',
  borderDark: '293040',
  borderObsidian: '262A34',
};

// Georgia is natively installed on Windows PowerPoint -> crisp luxury serif headlines with zero font-fallback bugs
const FONTS = {
  editorial: 'Georgia',
  ui: 'Montserrat',
};

const W = 13.333;
const H = 7.5;
const MARGIN_X = 0.68;
const CONTENT_W = W - MARGIN_X * 2; // 11.973

const img = (name) => path.join(ASSETS_DIR, name);

function addSlideChrome(slide, opts) {
  const {
    theme = 'light', // 'light' | 'obsidian' | 'velvet'
    step = 'STEP 01',
    kicker = '',
    title = '',
    slideNum = 2,
    totalSlides = 16,
  } = opts;

  const isDark = theme !== 'light';
  const bg = theme === 'velvet' ? COLORS.velvet : theme === 'obsidian' ? COLORS.obsidian : COLORS.ivory;
  const titleColor = isDark ? COLORS.ivory : COLORS.obsidian;
  const footerText = isDark ? COLORS.lightMuted : COLORS.stone;
  const dividerColor = isDark
    ? theme === 'velvet'
      ? COLORS.borderDark
      : COLORS.borderObsidian
    : COLORS.borderLight;

  slide.background = { color: bg };

  // Subtle step badge + category kicker
  slide.addShape('rect', {
    x: MARGIN_X,
    y: 0.36,
    w: 1.04,
    h: 0.25,
    fill: { color: isDark ? COLORS.obsidianCard : COLORS.ivorySoft },
    line: { color: COLORS.gold, width: 0.65 },
  });

  slide.addText(step.toUpperCase(), {
    x: MARGIN_X,
    y: 0.36,
    w: 1.04,
    h: 0.25,
    fontFace: FONTS.ui,
    fontSize: 8.5,
    bold: true,
    color: COLORS.gold,
    align: 'center',
    valign: 'middle',
    charSpacing: 1.2,
    margin: 0,
  });

  slide.addText(kicker.toUpperCase(), {
    x: MARGIN_X + 1.18,
    y: 0.36,
    w: 6.5,
    h: 0.25,
    fontFace: FONTS.ui,
    fontSize: 9.5,
    bold: true,
    color: isDark ? COLORS.lightMuted : COLORS.bronze,
    valign: 'middle',
    charSpacing: 2.0,
    margin: 0,
  });

  // Client Brand Logo in Header
  slide.addImage({
    path: img('jh_logo.png'),
    x: W - MARGIN_X - 4.45,
    y: 0.32,
    w: 0.22,
    h: 0.28,
  });

  slide.addText('JHONTY HUB  |  MENSWEAR ATELIER', {
    x: W - MARGIN_X - 4.15,
    y: 0.36,
    w: 4.15,
    h: 0.25,
    fontFace: FONTS.ui,
    fontSize: 9,
    color: footerText,
    align: 'right',
    charSpacing: 1.6,
    valign: 'middle',
    margin: 0,
  });

  // Luxury Editorial Serif Headline
  slide.addText(title, {
    x: MARGIN_X,
    y: 0.68,
    w: CONTENT_W,
    h: 0.58,
    fontFace: FONTS.editorial,
    fontSize: 27,
    bold: true,
    color: titleColor,
    valign: 'top',
    margin: 0,
  });

  // Champagne hairline rule
  slide.addShape('line', {
    x: MARGIN_X,
    y: 1.32,
    w: CONTENT_W,
    h: 0,
    line: { color: dividerColor, width: 0.75 },
  });

  // Footer hairline & clean location anchor
  slide.addShape('line', {
    x: MARGIN_X,
    y: 7.02,
    w: CONTENT_W,
    h: 0,
    line: { color: dividerColor, width: 0.6 },
  });

  slide.addText('Jontys Hub • Beside B.D. Reddy Gardens, B.N. Reddy Nagar, LB Nagar, Hyderabad – 500079', {
    x: MARGIN_X,
    y: 7.08,
    w: 9.8,
    h: 0.28,
    fontFace: FONTS.ui,
    fontSize: 9,
    color: footerText,
    valign: 'middle',
    margin: 0,
  });

  const numStr = String(slideNum).padStart(2, '0');
  slide.addText(`${numStr}  /  ${totalSlides}`, {
    x: W - MARGIN_X - 1.4,
    y: 7.08,
    w: 1.4,
    h: 0.28,
    fontFace: FONTS.ui,
    fontSize: 9.5,
    bold: true,
    color: COLORS.gold,
    align: 'right',
    valign: 'middle',
    margin: 0,
  });
}

function addCard(slide, x, y, w, h, opts = {}) {
  const {
    fill = COLORS.ivoryCard,
    border = COLORS.borderLight,
    accentTop = false,
    accentLeft = false,
    accentColor = COLORS.gold,
  } = opts;

  slide.addShape('rect', {
    x,
    y,
    w,
    h,
    fill: { color: fill },
    line: { color: border, width: 0.65 },
  });

  if (accentTop) {
    slide.addShape('rect', {
      x,
      y,
      w,
      h: 0.035,
      fill: { color: accentColor },
      line: { type: 'none' },
    });
  }
  if (accentLeft) {
    slide.addShape('rect', {
      x,
      y,
      w: 0.045,
      h,
      fill: { color: accentColor },
      line: { type: 'none' },
    });
  }
}

async function buildPresentation() {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'Jontys Hub Growth Advisory';
  pptx.company = 'Jontys Hub';
  pptx.subject = "Luxury Men's Fashion Growth & B2C Remarketing Strategy";
  pptx.title = "Jontys Hub – Luxury Men's Fashion Growth Blueprint";

  // =========================================================================
  // SLIDE 01 – COVER (Obsidian Black)
  // =========================================================================
  {
    const s = pptx.addSlide();
    s.background = { color: COLORS.obsidian };

    s.addImage({
      path: img('slide01_cover.jpg'),
      x: 7.733,
      y: 0,
      w: 5.60,
      h: 7.5,
    });

    s.addShape('line', {
      x: 7.733,
      y: 0.60,
      w: 0,
      h: 6.30,
      line: { color: COLORS.gold, width: 0.9 },
    });

    s.addImage({
      path: img('jh_logo.png'),
      x: 0.78,
      y: 0.52,
      w: 0.54,
      h: 0.65,
    });

    s.addText('JHONTY HUB', {
      x: 1.46,
      y: 0.54,
      w: 5.5,
      h: 0.36,
      fontFace: FONTS.editorial,
      fontSize: 26,
      bold: true,
      color: COLORS.ivory,
      charSpacing: 2.0,
      margin: 0,
    });

    s.addText('MENSWEAR ATELIER • HYDERABAD', {
      x: 1.48,
      y: 0.90,
      w: 5.5,
      h: 0.22,
      fontFace: FONTS.ui,
      fontSize: 8.8,
      bold: true,
      color: COLORS.gold,
      charSpacing: 2.0,
      margin: 0,
    });

    s.addShape('rect', {
      x: 0.78,
      y: 1.45,
      w: 0.55,
      h: 0.035,
      fill: { color: COLORS.gold },
      line: { type: 'none' },
    });

    s.addText("LUXURY MEN'S FASHION • B2C GROWTH BLUEPRINT", {
      x: 0.78,
      y: 1.58,
      w: 6.5,
      h: 0.26,
      fontFace: FONTS.ui,
      fontSize: 9.5,
      bold: true,
      color: COLORS.gold,
      charSpacing: 2.0,
      margin: 0,
    });

    s.addText("Elevating a Dedicated Men's Fashion Store Into a High-Intent Luxury Brand.", {
      x: 0.78,
      y: 1.96,
      w: 6.45,
      h: 1.05,
      fontFace: FONTS.editorial,
      fontSize: 23,
      bold: true,
      color: COLORS.ivory,
      lineSpacingMultiple: 1.25,
      margin: 0,
    });

    s.addShape('rect', {
      x: 0.78,
      y: 3.14,
      w: 1.10,
      h: 0.035,
      fill: { color: COLORS.gold },
      line: { type: 'none' },
    });

    s.addText(
      'A disciplined growth roadmap turning client-shot in-store video into high-intent walk-ins through sequential remarketing, quality education, and VIP community retention.',
      {
        x: 0.78,
        y: 3.32,
        w: 6.10,
        h: 0.85,
        fontFace: FONTS.ui,
        fontSize: 13,
        color: COLORS.lightMuted,
        lineSpacingMultiple: 1.35,
        margin: 0,
      }
    );

    s.addShape('line', {
      x: 0.78,
      y: 6.08,
      w: 6.10,
      h: 0,
      line: { color: COLORS.borderObsidian, width: 0.75 },
    });

    s.addText('LB Nagar Menswear Store  |  Hyderabad', {
      x: 0.78,
      y: 6.24,
      w: 6.10,
      h: 0.32,
      fontFace: FONTS.ui,
      fontSize: 13,
      bold: true,
      color: COLORS.gold,
      margin: 0,
    });

    s.addText('Beside B.D. Reddy Gardens, B.N. Reddy Nagar, Hyderabad – 500079', {
      x: 0.78,
      y: 6.58,
      w: 6.10,
      h: 0.28,
      fontFace: FONTS.ui,
      fontSize: 10,
      color: COLORS.lightMuted,
      margin: 0,
    });

    s.addText('01  /  16', {
      x: 0.78,
      y: 7.04,
      w: 1.5,
      h: 0.25,
      fontFace: FONTS.ui,
      fontSize: 9.5,
      bold: true,
      color: COLORS.bronze,
      margin: 0,
    });
  }

  // =========================================================================
  // SLIDE 02 – STEP 01: BRAND FOUNDATION (Warm Atelier Ivory)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'light',
      step: 'Step 01',
      kicker: 'Brand Foundation',
      title: 'A Dedicated Luxury Menswear Destination.',
      slideNum: 2,
    });

    s.addImage({
      path: img('slide02_store_display.jpg'),
      x: MARGIN_X,
      y: 1.56,
      w: 6.26,
      h: 3.50,
    });

    addCard(s, MARGIN_X, 5.20, 6.26, 1.58, {
      fill: COLORS.ivoryCard,
      border: COLORS.borderLight,
      accentLeft: true,
    });

    s.addText('ATELIER POSITIONING', {
      x: MARGIN_X + 0.24,
      y: 5.36,
      w: 5.8,
      h: 0.22,
      fontFace: FONTS.ui,
      fontSize: 9.5,
      bold: true,
      color: COLORS.bronze,
      charSpacing: 1.6,
      margin: 0,
    });

    s.addText(
      'Exclusively curated for discerning men in LB Nagar — built around elevated fit, fabric drape, and complete wardrobing rather than discounting.',
      {
        x: MARGIN_X + 0.24,
        y: 5.66,
        w: 5.8,
        h: 0.62,
        fontFace: FONTS.ui,
        fontSize: 12,
        color: COLORS.obsidian,
        lineSpacingMultiple: 1.28,
        margin: 0,
      }
    );

    s.addText('B.N. Reddy Nagar, LB Nagar, Hyderabad – 500079', {
      x: MARGIN_X + 0.24,
      y: 6.36,
      w: 5.8,
      h: 0.28,
      fontFace: FONTS.ui,
      fontSize: 10,
      color: COLORS.stone,
      margin: 0,
    });

    const rx = 7.22;
    const rw = CONTENT_W - (rx - MARGIN_X); // 5.433

    const kpis = [
      { label: 'AVG PRODUCT PRICE', val: '₹2,000', sub: 'Premium quality tier' },
      { label: 'AVG CUSTOMER BILL', val: '₹5,000', sub: 'Multi-item outfit basket' },
      { label: 'GROSS MARGIN', val: '< 25%', sub: 'Protected via high intent' },
    ];
    const kw = (rw - 0.28) / 3;
    kpis.forEach((k, i) => {
      const kx = rx + i * (kw + 0.14);
      addCard(s, kx, 1.56, kw, 1.62, {
        fill: COLORS.obsidian,
        border: COLORS.obsidian,
        accentTop: true,
        accentColor: COLORS.gold,
      });
      s.addText(k.label, {
        x: kx + 0.14,
        y: 1.72,
        w: kw - 0.28,
        h: 0.26,
        fontFace: FONTS.ui,
        fontSize: 8.5,
        bold: true,
        color: COLORS.lightMuted,
        charSpacing: 0.9,
        margin: 0,
      });
      s.addText(k.val, {
        x: kx + 0.14,
        y: 2.04,
        w: kw - 0.28,
        h: 0.56,
        fontFace: FONTS.editorial,
        fontSize: 25,
        bold: true,
        color: COLORS.gold,
        margin: 0,
      });
      s.addText(k.sub, {
        x: kx + 0.14,
        y: 2.68,
        w: kw - 0.28,
        h: 0.36,
        fontFace: FONTS.ui,
        fontSize: 9.8,
        color: COLORS.lightText,
        lineSpacingMultiple: 1.15,
        margin: 0,
      });
    });

    addCard(s, rx, 3.38, rw, 3.40, {
      fill: COLORS.ivoryCard,
      border: COLORS.borderLight,
    });

    s.addText('FIVE CORE MENSWEAR PILLARS', {
      x: rx + 0.24,
      y: 3.56,
      w: rw - 0.48,
      h: 0.26,
      fontFace: FONTS.ui,
      fontSize: 10,
      bold: true,
      color: COLORS.bronze,
      charSpacing: 1.5,
      margin: 0,
    });

    const cats = [
      { name: 'Casual Shirts', desc: 'Textured cottons, smart-casual & weekend weaves' },
      { name: 'Formal Shirts', desc: 'Tailored office polish & evening collar precision' },
      { name: 'T-Shirts', desc: 'Heavyweight crew-necks & structured minimalism' },
      { name: 'Jeans', desc: 'Clean luxury washes, modern tapered & stretch fits' },
      { name: 'Trousers', desc: 'Tailored dress trousers & refined daily chinos' },
    ];

    cats.forEach((c, idx) => {
      const cy = 3.96 + idx * 0.54;
      s.addText(`0${idx + 1}`, {
        x: rx + 0.24,
        y: cy + 0.04,
        w: 0.34,
        h: 0.34,
        fontFace: FONTS.editorial,
        fontSize: 13,
        bold: true,
        color: COLORS.gold,
        valign: 'middle',
        margin: 0,
      });
      s.addText(c.name, {
        x: rx + 0.64,
        y: cy + 0.02,
        w: 1.55,
        h: 0.36,
        fontFace: FONTS.ui,
        fontSize: 13,
        bold: true,
        color: COLORS.obsidian,
        valign: 'middle',
        margin: 0,
      });
      s.addText(c.desc, {
        x: rx + 2.22,
        y: cy + 0.02,
        w: 3.00,
        h: 0.38,
        fontFace: FONTS.ui,
        fontSize: 11,
        color: COLORS.stone,
        valign: 'middle',
        margin: 0,
      });
      if (idx < cats.length - 1) {
        s.addShape('line', {
          x: rx + 0.24,
          y: cy + 0.46,
          w: rw - 0.48,
          h: 0,
          line: { color: COLORS.borderLight, width: 0.5 },
        });
      }
    });
  }

  // =========================================================================
  // SLIDE 03 – STEP 02: MARKET SHIFT (Midnight Velvet)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'velvet',
      step: 'Step 02',
      kicker: 'Consumer Buying Shift',
      title: 'Why Luxury Men Choose Video Discovery & Store Trust.',
      slideNum: 3,
    });

    const shifts = [
      {
        num: '01',
        title: 'Video-First Inspiration',
        body: 'Style reels on Instagram and YouTube shape outfit desire long before shoppers leave home.',
      },
      {
        num: '02',
        title: 'Education Beats Discounting',
        body: 'Showing fabric drape, collar finish, and fit builds ₹2,000+ perceived value without price cuts.',
      },
      {
        num: '03',
        title: 'Specialist Store Preference',
        body: 'Discerning men trust a focused menswear boutique over crowded general stores for fit assurance.',
      },
      {
        num: '04',
        title: 'Community-Driven Loyalty',
        body: 'Exclusive WhatsApp new-stock drops & referral perks turn first-time buyers into repeat patrons.',
      },
    ];

    shifts.forEach((d, i) => {
      const dy = 1.58 + i * 1.30;
      addCard(s, MARGIN_X, dy, 6.72, 1.14, {
        fill: COLORS.velvetCard,
        border: COLORS.borderDark,
        accentLeft: true,
        accentColor: COLORS.gold,
      });
      s.addText(d.num, {
        x: MARGIN_X + 0.22,
        y: dy + 0.20,
        w: 0.50,
        h: 0.50,
        fontFace: FONTS.editorial,
        fontSize: 20,
        bold: true,
        color: COLORS.gold,
        margin: 0,
      });
      s.addText(d.title, {
        x: MARGIN_X + 0.78,
        y: dy + 0.18,
        w: 5.72,
        h: 0.30,
        fontFace: FONTS.ui,
        fontSize: 15,
        bold: true,
        color: COLORS.ivory,
        margin: 0,
      });
      s.addText(d.body, {
        x: MARGIN_X + 0.78,
        y: dy + 0.52,
        w: 5.72,
        h: 0.48,
        fontFace: FONTS.ui,
        fontSize: 12,
        color: COLORS.lightText,
        lineSpacingMultiple: 1.25,
        margin: 0,
      });
    });

    const rx = 7.68;
    const rw = 4.97;
    s.addImage({
      path: img('slide03_shopping.jpg'),
      x: rx,
      y: 1.58,
      w: rw,
      h: 3.72,
    });

    addCard(s, rx, 5.44, rw, 1.30, {
      fill: COLORS.obsidian,
      border: COLORS.gold,
      accentTop: true,
    });

    s.addText('LUXURY INSIGHT', {
      x: rx + 0.22,
      y: 5.58,
      w: rw - 0.44,
      h: 0.22,
      fontFace: FONTS.ui,
      fontSize: 9.5,
      bold: true,
      color: COLORS.gold,
      charSpacing: 1.6,
      margin: 0,
    });

    s.addText(
      'When digital storytelling proves quality first, customers visit ready to experience complete outfits—not hunt for bargains.',
      {
        x: rx + 0.22,
        y: 5.86,
        w: rw - 0.44,
        h: 0.76,
        fontFace: FONTS.ui,
        fontSize: 12,
        color: COLORS.ivory,
        lineSpacingMultiple: 1.3,
        margin: 0,
      }
    );
  }

  // =========================================================================
  // SLIDE 04 – STEP 03: STRATEGIC POSITION (Warm Atelier Ivory)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'light',
      step: 'Step 03',
      kicker: 'Strategic Position',
      title: 'Where We Stand Today & What We Unlock Next.',
      slideNum: 4,
    });

    // Left Dark Atelier Column: Current Strengths & Constraint
    const lw = 4.65;
    addCard(s, MARGIN_X, 1.58, lw, 5.18, {
      fill: COLORS.obsidian,
      border: COLORS.obsidian,
      accentTop: true,
      accentColor: COLORS.gold,
    });

    s.addText('CORE FOUNDATION TODAY', {
      x: MARGIN_X + 0.28,
      y: 1.82,
      w: lw - 0.56,
      h: 0.26,
      fontFace: FONTS.ui,
      fontSize: 10,
      bold: true,
      color: COLORS.gold,
      charSpacing: 1.6,
      margin: 0,
    });

    const strengths = [
      { k: 'Focused Menswear Brand', v: '100% men-only specialty store building authority across 5 categories.' },
      { k: 'High Basket Value', v: '₹5,000 average customer bill shows strong multi-piece styling potential.' },
      { k: 'Prime LB Nagar Anchor', v: 'Accessible physical boutique in B.N. Reddy Nagar built for personal trial.' },
      { k: 'Margin Guardrail (< 25%)', v: 'Requires high-intent buyers & remarketing discipline instead of broad discounts.' },
    ];

    strengths.forEach((st, i) => {
      const sy = 2.24 + i * 1.08;
      s.addText(st.k, {
        x: MARGIN_X + 0.28,
        y: sy,
        w: lw - 0.56,
        h: 0.28,
        fontFace: FONTS.ui,
        fontSize: 13.5,
        bold: true,
        color: COLORS.ivory,
        margin: 0,
      });
      s.addText(st.v, {
        x: MARGIN_X + 0.28,
        y: sy + 0.30,
        w: lw - 0.56,
        h: 0.58,
        fontFace: FONTS.ui,
        fontSize: 11.2,
        color: COLORS.lightMuted,
        lineSpacingMultiple: 1.24,
        margin: 0,
      });
      if (i < strengths.length - 1) {
        s.addShape('line', {
          x: MARGIN_X + 0.28,
          y: sy + 0.94,
          w: lw - 0.56,
          h: 0,
          line: { color: COLORS.borderObsidian, width: 0.5 },
        });
      }
    });

    // Right Column: 4 Clean Growth Priorities
    const gx = MARGIN_X + lw + 0.32;
    const gw = CONTENT_W - lw - 0.32; // 7.003

    const priorities = [
      {
        num: '01',
        title: 'Build Luxury Brand Positioning',
        desc: 'Elevate Instagram, Facebook & Google Maps to reflect a refined, trusted menswear atelier.',
      },
      {
        num: '02',
        title: 'Educate With Video-View Remarketing',
        desc: 'Pool 50%+ video ad viewers and serve sequential fabric, fit & styling ads.',
      },
      {
        num: '03',
        title: 'Attract High-Intent Store Visitors',
        desc: 'Target warm, educated buyers with Maps pins, sizing checks & fresh-stock store hooks.',
      },
      {
        num: '04',
        title: 'Retention via WhatsApp Community',
        desc: 'Share daily/weekly new-stock drops & structured referral discounts to grow repeat bills.',
      },
    ];

    priorities.forEach((p, i) => {
      const py = 1.58 + i * 1.32;
      addCard(s, gx, py, gw, 1.18, {
        fill: COLORS.ivoryCard,
        border: COLORS.borderLight,
        accentLeft: true,
        accentColor: i === 2 ? COLORS.obsidian : COLORS.gold,
      });
      s.addText(p.num, {
        x: gx + 0.24,
        y: py + 0.22,
        w: 0.50,
        h: 0.50,
        fontFace: FONTS.editorial,
        fontSize: 20,
        bold: true,
        color: COLORS.gold,
        margin: 0,
      });
      s.addText(p.title, {
        x: gx + 0.80,
        y: py + 0.20,
        w: gw - 1.04,
        h: 0.30,
        fontFace: FONTS.ui,
        fontSize: 15,
        bold: true,
        color: COLORS.obsidian,
        margin: 0,
      });
      s.addText(p.desc, {
        x: gx + 0.80,
        y: py + 0.54,
        w: gw - 1.04,
        h: 0.48,
        fontFace: FONTS.ui,
        fontSize: 12,
        color: COLORS.stone,
        lineSpacingMultiple: 1.25,
        margin: 0,
      });
    });
  }

  // =========================================================================
  // SLIDE 05 – STEP 04: IDEAL CUSTOMER PROFILES (Obsidian Black)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'obsidian',
      step: 'Step 04',
      kicker: 'Target Audience Architecture',
      title: 'Four High-Value Men We Attract to Store.',
      slideNum: 5,
    });

    const personas = [
      {
        img: 'icp1_young_fashion.jpg',
        title: 'Young Style Leaders',
        age: '20 – 30 YRS',
        focus: 'Casual Shirts, T-Shirts & Slim Jeans',
        hook: 'Trendy Reels & fresh weekly stock alerts.',
      },
      {
        img: 'icp2_working_pro.jpg',
        title: 'Working Professionals',
        age: '25 – 40 YRS',
        focus: 'Formal Shirts & Tailored Trousers',
        hook: 'Fabric quality & all-day office comfort videos.',
      },
      {
        img: 'icp3_aspirational.jpg',
        title: 'Wardrobe Upgraders',
        age: '25 – 40 YRS',
        focus: 'Layered Casuals, Tees & Trousers',
        hook: 'Complete outfit pairing & styling education.',
      },
      {
        img: 'icp4_occasion.jpg',
        title: 'Occasion Shoppers',
        age: '23 – 45 YRS',
        focus: 'Evening Formals & Curated Outfits',
        hook: 'VIP in-store trial & peer referral privileges.',
      },
    ];

    const cw = (CONTENT_W - 0.24 * 3) / 4; // 2.813
    personas.forEach((p, i) => {
      const px = MARGIN_X + i * (cw + 0.24);
      const py = 1.58;
      const ch = 5.20;

      addCard(s, px, py, cw, ch, {
        fill: COLORS.obsidianCard,
        border: COLORS.borderObsidian,
        accentTop: true,
        accentColor: COLORS.gold,
      });

      // Exact 3:4 Portrait centered cleanly at top of card
      const imgW = cw - 0.24; // 2.573
      const imgH = +(imgW / 0.747).toFixed(3); // 3.444? Wait: height limit = 3.25" at w=2.42"
      const pImgW = 2.40;
      const pImgH = 3.20; // 2.40 / 3.20 = 0.75 exact 3:4!
      const pImgX = px + (cw - pImgW) / 2;

      s.addImage({
        path: img(p.img),
        x: pImgX,
        y: py + 0.16,
        w: pImgW,
        h: pImgH,
      });

      // Subtle dark age pill overlaid neatly bottom-left of portrait
      s.addShape('rect', {
        x: pImgX + 0.10,
        y: py + 0.26,
        w: 0.92,
        h: 0.26,
        fill: { color: COLORS.obsidian },
        line: { color: COLORS.gold, width: 0.6 },
      });
      s.addText(p.age, {
        x: pImgX + 0.10,
        y: py + 0.26,
        w: 0.92,
        h: 0.26,
        fontFace: FONTS.ui,
        fontSize: 8.5,
        bold: true,
        color: COLORS.gold,
        align: 'center',
        valign: 'middle',
        margin: 0,
      });

      s.addText(p.title, {
        x: px + 0.18,
        y: py + 3.48,
        w: cw - 0.36,
        h: 0.34,
        fontFace: FONTS.editorial,
        fontSize: 14.5,
        bold: true,
        color: COLORS.ivory,
        margin: 0,
      });

      s.addShape('line', {
        x: px + 0.18,
        y: py + 3.84,
        w: cw - 0.36,
        h: 0,
        line: { color: COLORS.borderObsidian, width: 0.6 },
      });

      s.addText('STAPLE FOCUS', {
        x: px + 0.18,
        y: py + 3.92,
        w: cw - 0.36,
        h: 0.18,
        fontFace: FONTS.ui,
        fontSize: 8,
        bold: true,
        color: COLORS.gold,
        charSpacing: 1.0,
        margin: 0,
      });
      s.addText(p.focus, {
        x: px + 0.18,
        y: py + 4.10,
        w: cw - 0.36,
        h: 0.34,
        fontFace: FONTS.ui,
        fontSize: 10.5,
        color: COLORS.lightText,
        margin: 0,
      });

      s.addText('VIDEO & REMARKETING HOOK', {
        x: px + 0.18,
        y: py + 4.52,
        w: cw - 0.36,
        h: 0.18,
        fontFace: FONTS.ui,
        fontSize: 8,
        bold: true,
        color: COLORS.bronze,
        charSpacing: 0.8,
        margin: 0,
      });
      s.addText(p.hook, {
        x: px + 0.18,
        y: py + 4.70,
        w: cw - 0.36,
        h: 0.38,
        fontFace: FONTS.ui,
        fontSize: 10.5,
        color: COLORS.lightMuted,
        margin: 0,
      });
    });
  }

  // =========================================================================
  // SLIDE 06 – STEP 05: LOCAL CATCHMENT FOCUS (Midnight Velvet)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'velvet',
      step: 'Step 05',
      kicker: 'Hyper-Local Catchment Focus',
      title: 'Step-by-Step Radius Expansion Around LB Nagar.',
      slideNum: 6,
    });

    const mx = MARGIN_X;
    const my = 1.56;
    const mw = 7.25;
    const mh = 5.20;

    addCard(s, mx, my, mw, mh, {
      fill: '0B0E14',
      border: COLORS.borderDark,
    });

    s.addText('SOUTH-EAST HYDERABAD RETAIL CATCHMENT', {
      x: mx + 0.22,
      y: my + 0.16,
      w: mw - 0.44,
      h: 0.22,
      fontFace: FONTS.ui,
      fontSize: 9,
      bold: true,
      color: COLORS.gold,
      charSpacing: 1.2,
      margin: 0,
    });

    const cx = mx + 3.40;
    const cy = my + 3.22;

    s.addShape('ellipse', {
      x: cx - 2.90,
      y: cy - 2.50,
      w: 5.80,
      h: 4.42,
      fill: { color: '131A29', transparency: 35 },
      line: { color: COLORS.stone, width: 1.0, dashType: 'dash' },
    });

    s.addShape('ellipse', {
      x: cx - 2.02,
      y: cy - 1.72,
      w: 4.04,
      h: 3.16,
      fill: { color: '182236', transparency: 25 },
      line: { color: COLORS.bronze, width: 1.25 },
    });

    s.addShape('ellipse', {
      x: cx - 1.14,
      y: cy - 0.94,
      w: 2.28,
      h: 1.78,
      fill: { color: COLORS.gold, transparency: 82 },
      line: { color: COLORS.gold, width: 1.75 },
    });

    s.addShape('line', {
      x: mx + 2.80,
      y: my + 0.56,
      w: 0.56,
      h: 3.86,
      line: { color: '2B3648', width: 1.8 },
    });
    s.addShape('line', {
      x: mx + 0.70,
      y: my + 1.22,
      w: 5.88,
      h: 0.45,
      flipV: true,
      line: { color: '2B3648', width: 1.8 },
    });

    s.addText('NH 65 / Vijayawada Hwy', {
      x: mx + 4.82,
      y: my + 0.96,
      w: 1.9,
      h: 0.22,
      fontFace: FONTS.ui,
      fontSize: 8.5,
      color: COLORS.lightMuted,
      margin: 0,
    });
    s.addText('Nagarjuna Sagar Rd', {
      x: mx + 1.38,
      y: my + 4.14,
      w: 1.65,
      h: 0.22,
      fontFace: FONTS.ui,
      fontSize: 8.5,
      color: COLORS.lightMuted,
      margin: 0,
    });

    const mapPin = (px, py, label, tierColor, isAnchor = false) => {
      const sz = isAnchor ? 0.26 : 0.16;
      s.addShape('ellipse', {
        x: px - sz / 2,
        y: py - sz / 2,
        w: sz,
        h: sz,
        fill: { color: tierColor },
        line: { color: COLORS.ivory, width: isAnchor ? 1.5 : 0.75 },
      });
      s.addText(label, {
        x: px - 0.85,
        y: py + (isAnchor ? 0.16 : 0.10),
        w: 1.70,
        h: 0.28,
        fontFace: FONTS.ui,
        fontSize: isAnchor ? 9.5 : 9.2,
        bold: true,
        color: isAnchor ? COLORS.gold : COLORS.ivory,
        align: 'center',
        margin: 0,
      });
    };

    mapPin(cx, cy, 'JONTYS HUB\nB.N. Reddy Nagar', COLORS.gold, true);
    mapPin(cx + 0.84, cy - 0.30, 'Vanasthalipuram', COLORS.gold);
    mapPin(cx - 0.78, cy + 0.48, 'Hastinapuram', COLORS.gold);
    mapPin(cx - 0.10, cy - 1.32, 'LB Nagar', COLORS.bronze);
    mapPin(cx + 1.44, cy - 1.04, 'Chintalkunta', COLORS.bronze);
    mapPin(cx - 1.52, cy - 2.10, 'Kothapet', COLORS.lightMuted);

    s.addShape('rect', {
      x: mx + 0.22,
      y: my + mh - 0.52,
      w: mw - 0.44,
      h: 0.36,
      fill: { color: COLORS.velvet },
      line: { color: COLORS.borderDark, width: 0.6 },
    });
    const legendItems = [
      { label: 'Store Anchor', color: COLORS.gold, x: mx + 0.38 },
      { label: 'Zone 01 Primary', color: COLORS.gold, x: mx + 1.95 },
      { label: 'Zone 02 Secondary', color: COLORS.bronze, x: mx + 3.65 },
      { label: 'Zone 03 Expand', color: COLORS.lightMuted, x: mx + 5.45 },
    ];
    legendItems.forEach((l) => {
      s.addShape('ellipse', {
        x: l.x,
        y: my + mh - 0.40,
        w: 0.12,
        h: 0.12,
        fill: { color: l.color },
      });
      s.addText(l.label, {
        x: l.x + 0.18,
        y: my + mh - 0.44,
        w: 1.4,
        h: 0.22,
        fontFace: FONTS.ui,
        fontSize: 9,
        color: COLORS.lightText,
        margin: 0,
      });
    });

    const rx = mx + mw + 0.30;
    const rw = CONTENT_W - mw - 0.30; // 4.423

    const tiers = [
      {
        tier: 'ZONE 01 • PRIMARY RADIUS',
        places: 'B.N. Reddy Nagar, Vanasthalipuram, Hastinapuram',
        role: 'High-frequency video ads, Google Maps visibility & local store-visit remarketing.',
        accent: COLORS.gold,
      },
      {
        tier: 'ZONE 02 • SECONDARY RADIUS',
        places: 'LB Nagar & Chintalkunta',
        role: 'Educated lookbooks & Instagram/YouTube retargeting for nearby style buyers.',
        accent: COLORS.bronze,
      },
      {
        tier: 'ZONE 03 • EXPANSION TEST',
        places: 'Kothapet',
        role: 'Selective budget expansion once Primary store-visit payback is verified.',
        accent: COLORS.stone,
      },
    ];

    tiers.forEach((t, i) => {
      const ty = 1.56 + i * 1.32;
      addCard(s, rx, ty, rw, 1.18, {
        fill: COLORS.velvetCard,
        border: COLORS.borderDark,
        accentLeft: true,
        accentColor: t.accent,
      });
      s.addText(t.tier, {
        x: rx + 0.20,
        y: ty + 0.12,
        w: rw - 0.34,
        h: 0.20,
        fontFace: FONTS.ui,
        fontSize: 9,
        bold: true,
        color: t.accent === COLORS.stone ? COLORS.lightMuted : t.accent,
        charSpacing: 1.2,
        margin: 0,
      });
      s.addText(t.places, {
        x: rx + 0.20,
        y: ty + 0.34,
        w: rw - 0.34,
        h: 0.30,
        fontFace: FONTS.ui,
        fontSize: 13,
        bold: true,
        color: COLORS.ivory,
        margin: 0,
      });
      s.addText(t.role, {
        x: rx + 0.20,
        y: ty + 0.66,
        w: rw - 0.34,
        h: 0.44,
        fontFace: FONTS.ui,
        fontSize: 11,
        color: COLORS.lightText,
        lineSpacingMultiple: 1.2,
        margin: 0,
      });
    });

    addCard(s, rx, 5.54, rw, 1.22, {
      fill: COLORS.obsidian,
      border: COLORS.gold,
    });
    s.addText('FLAGSHIP STORE ADDRESS', {
      x: rx + 0.20,
      y: 5.68,
      w: rw - 0.40,
      h: 0.22,
      fontFace: FONTS.ui,
      fontSize: 9,
      bold: true,
      color: COLORS.gold,
      charSpacing: 1.3,
      margin: 0,
    });
    s.addText('Beside B.D. Reddy Gardens, B.N. Reddy Nagar, LB Nagar, Hyderabad – 500079', {
      x: rx + 0.20,
      y: 5.96,
      w: rw - 0.40,
      h: 0.56,
      fontFace: FONTS.ui,
      fontSize: 11.5,
      color: COLORS.ivory,
      lineSpacingMultiple: 1.25,
      margin: 0,
    });
  }

  // =========================================================================
  // SLIDE 07 – STEP 06: 5-STAGE ACQUISITION FUNNEL (Warm Atelier Ivory)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'light',
      step: 'Step 06',
      kicker: 'Conversion Architecture',
      title: 'Step-by-Step Journey From Scroll to Store Purchase.',
      slideNum: 7,
    });

    const stages = [
      {
        num: '01',
        name: 'VIDEO DISCOVERY',
        channel: 'Instagram Reels • Facebook Video • YouTube',
        objective: 'Capture local attention & pool engaged video viewers.',
        action: 'Style-led video ads reach LB Nagar men; 50%+ video viewers enter the warm remarketing audience.',
        w: 5.15,
        bg: COLORS.obsidian,
        txt: COLORS.ivory,
      },
      {
        num: '02',
        name: 'EDUCATE & POSITION',
        channel: 'Video-View Remarketing • Meta & Display',
        objective: 'Educate buyers & build luxury brand positioning.',
        action: 'Sequential ads highlight fabric drape, collar finish & styling—proving value before store entry.',
        w: 4.50,
        bg: '181D27',
        txt: COLORS.ivory,
      },
      {
        num: '03',
        name: 'HIGH-INTENT FILTER',
        channel: 'WhatsApp Community • Google Business Profile',
        objective: 'Identify serious buyers ready to visit in person.',
        action: 'Maps directions, sizing checks & fresh-stock previews isolate genuine local shoppers.',
        w: 3.85,
        bg: '232A38',
        txt: COLORS.ivory,
      },
      {
        num: '04',
        name: 'GOAL: STORE VISIT',
        channel: 'In-Store Trial • Outfit Styling',
        objective: 'Convert high-intent walk-ins into ₹5,000+ bills.',
        action: 'Educated guests arrive pre-sold on quality—staff guide complete multi-item outfit trials.',
        w: 3.20,
        bg: '2D3340',
        txt: COLORS.ivory,
      },
      {
        num: '05',
        name: 'RETENTION',
        channel: 'WhatsApp Drops • Referrals',
        objective: 'Turn buyers into repeat members & referrers.',
        action: 'Daily/weekly stock drops & margin-safe referral perks bring customers back repeatedly.',
        w: 2.65,
        bg: COLORS.gold,
        txt: COLORS.obsidian,
      },
    ];

    const funnelCenterX = MARGIN_X + 2.60;
    stages.forEach((st, i) => {
      const sy = 1.60 + i * 1.02;
      const sh = 0.88;
      const fx = funnelCenterX - st.w / 2;

      s.addShape('roundRect', {
        x: fx,
        y: sy,
        w: st.w,
        h: sh,
        rectRadius: 0.08,
        fill: { color: st.bg },
        line: { color: i === 4 ? COLORS.bronze : COLORS.gold, width: 0.75 },
      });

      s.addText(`${st.num}   ${st.name}`, {
        x: fx + 0.14,
        y: sy + 0.14,
        w: st.w - 0.28,
        h: 0.34,
        fontFace: FONTS.ui,
        fontSize: 12.2,
        bold: true,
        color: i === 4 ? COLORS.obsidian : COLORS.gold,
        align: 'center',
        margin: 0,
      });

      s.addText(st.channel, {
        x: fx + 0.14,
        y: sy + 0.48,
        w: st.w - 0.28,
        h: 0.26,
        fontFace: FONTS.ui,
        fontSize: 9.8,
        color: st.txt,
        align: 'center',
        margin: 0,
      });

      if (i < stages.length - 1) {
        s.addShape('line', {
          x: funnelCenterX,
          y: sy + sh,
          w: 0,
          h: 0.14,
          line: { color: COLORS.gold, width: 1.5 },
        });
      }

      const dx = MARGIN_X + 5.55;
      const dw = CONTENT_W - 5.55;
      s.addShape('line', {
        x: fx + st.w + 0.08,
        y: sy + sh / 2,
        w: dx - (fx + st.w + 0.08),
        h: 0,
        line: { color: COLORS.borderLight, width: 0.9, dashType: 'dot' },
      });

      addCard(s, dx, sy, dw, sh, {
        fill: COLORS.ivoryCard,
        border: COLORS.borderLight,
        accentLeft: true,
        accentColor: i === 4 ? COLORS.obsidian : COLORS.gold,
      });

      s.addText(st.objective, {
        x: dx + 0.22,
        y: sy + 0.12,
        w: dw - 0.38,
        h: 0.28,
        fontFace: FONTS.ui,
        fontSize: 12.2,
        bold: true,
        color: COLORS.obsidian,
        margin: 0,
      });

      s.addText(st.action, {
        x: dx + 0.22,
        y: sy + 0.44,
        w: dw - 0.38,
        h: 0.36,
        fontFace: FONTS.ui,
        fontSize: 11,
        color: COLORS.stone,
        margin: 0,
      });
    });
  }

  // =========================================================================
  // SLIDE 08 – STEP 07: VIDEO-VIEW REMARKETING ENGINE (Obsidian Black)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'obsidian',
      step: 'Step 07',
      kicker: 'Video-View Remarketing Strategy',
      title: 'Educate Video Viewers Into High-Intent Store Visitors.',
      slideNum: 8,
    });

    addCard(s, MARGIN_X, 1.52, CONTENT_W, 0.84, {
      fill: COLORS.velvet,
      border: COLORS.gold,
      accentLeft: true,
      accentColor: COLORS.gold,
    });

    s.addText('STRATEGIC PRINCIPLE', {
      x: MARGIN_X + 0.24,
      y: 1.60,
      w: 3.8,
      h: 0.20,
      fontFace: FONTS.ui,
      fontSize: 9,
      bold: true,
      color: COLORS.gold,
      charSpacing: 1.4,
      margin: 0,
    });

    s.addText(
      'Anyone watching 50%+ of our Reels or YouTube ads becomes our custom Remarketing Audience — educated step-by-step on quality and positioning until they visit our LB Nagar store.',
      {
        x: MARGIN_X + 0.24,
        y: 1.84,
        w: CONTENT_W - 0.48,
        h: 0.44,
        fontFace: FONTS.ui,
        fontSize: 11.5,
        color: COLORS.ivory,
        lineSpacingMultiple: 1.18,
        margin: 0,
      }
    );

    const rSteps = [
      {
        num: '01',
        tag: 'CAPTURE AUDIENCE',
        title: 'Build Video-View\nRemarketing Pool',
        highlight: '50%+ Video Watchers',
        body: 'Style-led Instagram Reels, Facebook Video & YouTube Ads reach local men. Engaged viewers automatically enter a warm custom audience.',
        goal: 'Signal: Real Interest',
        fill: COLORS.obsidianCard,
        accent: COLORS.gold,
      },
      {
        num: '02',
        tag: 'EDUCATE BUYERS',
        title: 'Show Fabric, Fit\n& Styling Value',
        highlight: 'Educate Before Selling',
        body: 'Sequential remarketing explains fabric breathability, collar precision & complete outfits—justifying ₹2,000+ craftsmanship.',
        goal: 'Outcome: Educated Buyer',
        fill: COLORS.obsidianCard,
        accent: COLORS.gold,
      },
      {
        num: '03',
        tag: 'CREATE POSITIONING',
        title: 'Establish Luxury\nStore Authority',
        highlight: 'Men-Only Style Hub',
        body: 'Consistent Meta & Display creative positions Jontys Hub as B.N. Reddy Nagar’s specialist menswear destination.',
        goal: 'Outcome: Brand Preference',
        fill: COLORS.obsidianCard,
        accent: COLORS.bronze,
      },
      {
        num: '04',
        tag: 'HIGH-INTENT GOAL',
        title: 'Drive Qualified\nStore Visits',
        highlight: 'In-Store Trial & ₹5,000+ Bills',
        body: 'Deliver Maps pins, sizing checks & fresh-stock drops exclusively to warm viewers—converting high-intent store walk-ins.',
        goal: 'Goal: Visit the Store',
        fill: COLORS.velvet,
        accent: COLORS.gold,
        isGoal: true,
      },
    ];

    const gap = 0.30;
    const sw = (CONTENT_W - gap * 3) / 4; // 2.768
    const sy = 2.54;
    const sh = 3.24;

    rSteps.forEach((st, idx) => {
      const sx = MARGIN_X + idx * (sw + gap);

      if (idx < rSteps.length - 1) {
        s.addShape('rightArrow', {
          x: sx + sw + 0.05,
          y: sy + sh / 2 - 0.11,
          w: gap - 0.10,
          h: 0.22,
          fill: { color: COLORS.gold },
          line: { type: 'none' },
        });
      }

      addCard(s, sx, sy, sw, sh, {
        fill: st.fill,
        border: st.isGoal ? COLORS.gold : COLORS.borderObsidian,
        accentTop: true,
        accentColor: st.accent,
      });

      s.addText(`${st.num} • ${st.tag}`, {
        x: sx + 0.18,
        y: sy + 0.16,
        w: sw - 0.36,
        h: 0.20,
        fontFace: FONTS.ui,
        fontSize: 8.5,
        bold: true,
        color: st.accent,
        charSpacing: 1.0,
        margin: 0,
      });

      s.addText(st.title, {
        x: sx + 0.18,
        y: sy + 0.42,
        w: sw - 0.36,
        h: 0.54,
        fontFace: FONTS.editorial,
        fontSize: 15,
        bold: true,
        color: COLORS.ivory,
        lineSpacingMultiple: 1.15,
        margin: 0,
      });

      s.addShape('rect', {
        x: sx + 0.18,
        y: sy + 1.06,
        w: sw - 0.36,
        h: 0.32,
        fill: { color: st.isGoal ? COLORS.gold : COLORS.obsidian },
        line: { color: COLORS.gold, width: 0.6 },
      });

      s.addText(st.highlight.toUpperCase(), {
        x: sx + 0.22,
        y: sy + 1.06,
        w: sw - 0.44,
        h: 0.32,
        fontFace: FONTS.ui,
        fontSize: 8.5,
        bold: true,
        color: st.isGoal ? COLORS.obsidian : COLORS.gold,
        align: 'center',
        valign: 'middle',
        charSpacing: 0.5,
        margin: 0,
      });

      s.addText(st.body, {
        x: sx + 0.18,
        y: sy + 1.50,
        w: sw - 0.36,
        h: 1.16,
        fontFace: FONTS.ui,
        fontSize: 10.6,
        color: COLORS.lightText,
        lineSpacingMultiple: 1.24,
        margin: 0,
      });

      s.addShape('line', {
        x: sx + 0.18,
        y: sy + 2.78,
        w: sw - 0.36,
        h: 0,
        line: { color: COLORS.borderObsidian, width: 0.6 },
      });

      s.addText(st.goal, {
        x: sx + 0.18,
        y: sy + 2.86,
        w: sw - 0.36,
        h: 0.26,
        fontFace: FONTS.ui,
        fontSize: 10.5,
        bold: true,
        color: st.isGoal ? COLORS.gold : COLORS.lightMuted,
        margin: 0,
      });
    });

    addCard(s, MARGIN_X, 5.96, CONTENT_W, 0.82, {
      fill: COLORS.obsidianCard,
      border: COLORS.borderObsidian,
      accentLeft: true,
      accentColor: COLORS.bronze,
    });

    s.addText('WHY THIS PROTECTS OUR < 25% MARGIN', {
      x: MARGIN_X + 0.24,
      y: 6.06,
      w: 3.8,
      h: 0.20,
      fontFace: FONTS.ui,
      fontSize: 9,
      bold: true,
      color: COLORS.gold,
      charSpacing: 1.2,
      margin: 0,
    });

    s.addText(
      'Instead of burning budget chasing cold clicks, video ads educate locals affordably — focusing store-visit spend strictly on pre-sold, high-intent shoppers.',
      {
        x: MARGIN_X + 0.24,
        y: 6.30,
        w: CONTENT_W - 0.48,
        h: 0.38,
        fontFace: FONTS.ui,
        fontSize: 11.2,
        color: COLORS.lightText,
        margin: 0,
      }
    );
  }

  // =========================================================================
  // SLIDE 09 – STEP 08: OMNICHANNEL ECOSYSTEM (Midnight Velvet)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'velvet',
      step: 'Step 08',
      kicker: 'Channel Architecture',
      title: 'Six Connected Channels Driving High-Intent Walk-Ins.',
      slideNum: 9,
    });

    const hubW = 3.10;
    const hubH = 2.00;
    const hubX = W / 2 - hubW / 2;
    const hubY = 4.20 - hubH / 2;

    const nodes = [
      {
        side: 'left',
        row: 0,
        name: 'Instagram Ads',
        logo: 'assets/logos/instagram.svg',
        role: 'VIDEO DISCOVERY & POOL',
        desc: 'Luxury reels capturing 50%+ video viewers for warm remarketing.',
      },
      {
        side: 'left',
        row: 1,
        name: 'Facebook Ads',
        logo: 'assets/logos/facebook.svg',
        role: 'EDUCATIONAL REMARKETING',
        desc: 'Retarget warm LB Nagar viewers with fabric, fit & outfit education.',
      },
      {
        side: 'left',
        row: 2,
        name: 'YouTube Ads',
        logo: 'assets/logos/youtube.svg',
        role: 'POSITIONING STORYTELLING',
        desc: 'Shorts & in-stream videos building luxury authority & buyer intent.',
      },
      {
        side: 'right',
        row: 0,
        name: 'Google Display Ads',
        logo: 'assets/logos/google_display.svg',
        role: 'VISUAL RECALL',
        desc: 'Maintain polished brand recall across apps until prospects visit.',
      },
      {
        side: 'right',
        row: 1,
        name: 'Google Business Profile',
        logo: 'assets/logos/google_business.svg',
        role: 'GOAL: STORE VISIT',
        desc: 'Maps pins, store visuals & reviews turning intent into walk-ins.',
      },
      {
        side: 'right',
        row: 2,
        name: 'WhatsApp Community',
        logo: 'assets/logos/whatsapp.svg',
        role: 'STOCK DROPS & REFERRALS',
        desc: 'Daily/weekly new-stock updates & referral perks for repeat visits.',
      },
    ];

    const nw = 3.62;
    const nh = 1.48;
    const leftX = MARGIN_X;
    const rightX = W - MARGIN_X - nw;

    nodes.forEach((n) => {
      const ny = 1.62 + n.row * 1.76;
      const startX = n.side === 'left' ? leftX + nw : rightX;
      const startY = ny + nh / 2;
      const endX = n.side === 'left' ? hubX : hubX + hubW;
      const endY = hubY + hubH / 2;

      s.addShape('line', {
        x: Math.min(startX, endX),
        y: Math.min(startY, endY),
        w: Math.abs(endX - startX),
        h: Math.abs(endY - startY),
        flipV: (startX - endX) * (startY - endY) < 0,
        line: { color: COLORS.gold, width: 1.0, dashType: 'dash' },
      });
    });

    s.addShape('roundRect', {
      x: hubX,
      y: hubY,
      w: hubW,
      h: hubH,
      rectRadius: 0.12,
      fill: { color: COLORS.obsidian },
      line: { color: COLORS.gold, width: 2.0 },
    });

    s.addImage({
      path: img('jh_logo.png'),
      x: hubX + hubW / 2 - 0.22,
      y: hubY + 0.15,
      w: 0.44,
      h: 0.55,
    });

    s.addText('CENTRAL GROWTH ENGINE', {
      x: hubX + 0.20,
      y: hubY + 0.74,
      w: hubW - 0.40,
      h: 0.20,
      fontFace: FONTS.ui,
      fontSize: 8.5,
      bold: true,
      color: COLORS.gold,
      align: 'center',
      charSpacing: 1.5,
      margin: 0,
    });

    s.addText('JHONTY HUB\nB2C COMMUNITY', {
      x: hubX + 0.20,
      y: hubY + 0.96,
      w: hubW - 0.40,
      h: 0.56,
      fontFace: FONTS.editorial,
      fontSize: 16.5,
      bold: true,
      color: COLORS.ivory,
      align: 'center',
      valign: 'middle',
      margin: 0,
    });

    s.addText('Store Visits • Stock Drops • Referrals', {
      x: hubX + 0.15,
      y: hubY + 1.56,
      w: hubW - 0.30,
      h: 0.28,
      fontFace: FONTS.ui,
      fontSize: 8.8,
      color: COLORS.lightText,
      align: 'center',
      margin: 0,
    });

    nodes.forEach((n) => {
      const nx = n.side === 'left' ? leftX : rightX;
      const ny = 1.62 + n.row * 1.76;

      addCard(s, nx, ny, nw, nh, {
        fill: COLORS.velvetCard,
        border: COLORS.borderDark,
        accentTop: true,
        accentColor: COLORS.gold,
      });

      s.addText(n.role, {
        x: nx + 0.20,
        y: ny + 0.14,
        w: nw - 0.40,
        h: 0.20,
        fontFace: FONTS.ui,
        fontSize: 8.8,
        bold: true,
        color: COLORS.gold,
        charSpacing: 1.2,
        margin: 0,
      });

      s.addImage({
        path: path.join(__dirname, n.logo),
        x: nx + 0.20,
        y: ny + 0.36,
        w: 0.32,
        h: 0.32,
      });

      s.addText(n.name, {
        x: nx + 0.60,
        y: ny + 0.34,
        w: nw - 0.76,
        h: 0.34,
        fontFace: FONTS.editorial,
        fontSize: 16,
        bold: true,
        color: COLORS.ivory,
        margin: 0,
      });

      s.addText(n.desc, {
        x: nx + 0.20,
        y: ny + 0.74,
        w: nw - 0.40,
        h: 0.60,
        fontFace: FONTS.ui,
        fontSize: 11,
        color: COLORS.lightText,
        lineSpacingMultiple: 1.24,
        margin: 0,
      });
    });

    addCard(s, hubX, 1.62, hubW, 1.30, {
      fill: COLORS.velvetCard,
      border: COLORS.borderDark,
    });
    s.addText('VIDEO TO STORE LOOP', {
      x: hubX + 0.20,
      y: 1.76,
      w: hubW - 0.40,
      h: 0.22,
      fontFace: FONTS.ui,
      fontSize: 9,
      bold: true,
      color: COLORS.gold,
      align: 'center',
      charSpacing: 1.3,
      margin: 0,
    });
    s.addText('Video ads warm local shoppers; education & Maps convert them into store walk-ins.', {
      x: hubX + 0.20,
      y: 2.04,
      w: hubW - 0.40,
      h: 0.70,
      fontFace: FONTS.ui,
      fontSize: 11.2,
      color: COLORS.lightText,
      align: 'center',
      lineSpacingMultiple: 1.25,
      margin: 0,
    });

    addCard(s, hubX, 5.44, hubW, 1.18, {
      fill: COLORS.velvetCard,
      border: COLORS.borderDark,
    });
    s.addText('REPEAT ENGINE', {
      x: hubX + 0.20,
      y: 5.56,
      w: hubW - 0.40,
      h: 0.22,
      fontFace: FONTS.ui,
      fontSize: 9,
      bold: true,
      color: COLORS.bronze,
      align: 'center',
      charSpacing: 1.3,
      margin: 0,
    });
    s.addText('Buyers join WhatsApp Community, driving repeat visits & referrals at near-zero CAC.', {
      x: hubX + 0.20,
      y: 5.84,
      w: hubW - 0.40,
      h: 0.64,
      fontFace: FONTS.ui,
      fontSize: 11,
      color: COLORS.lightText,
      align: 'center',
      lineSpacingMultiple: 1.22,
      margin: 0,
    });
  }

  // =========================================================================
  // SLIDE 10 – STEP 09: LUXURY CREATIVE DIRECTION (Warm Atelier Ivory)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'light',
      step: 'Step 09',
      kicker: 'Merchandising & Lookbook Direction',
      title: 'Luxury Visual Stories That Educate Before They Visit.',
      slideNum: 10,
    });

    const products = [
      {
        img: 'cat1_casual_shirts.jpg',
        cat: 'CASUAL SHIRTS',
        tagline: 'Effortless Refined Style.',
        concept: 'Weekend transition reels showcasing breathable cotton weaves & sharp rolled-cuff silhouettes.',
      },
      {
        img: 'cat2_formal_shirts.jpg',
        cat: 'FORMAL SHIRTS',
        tagline: 'Boardroom & Evening Polish.',
        concept: 'Macro fabric, collar & cuff close-ups educating buyers on structure and precision tailoring.',
      },
      {
        img: 'cat3_tshirts.jpg',
        cat: 'T-SHIRTS',
        tagline: 'Elevated Everyday Minimal.',
        concept: 'Clean studio lookbooks highlighting heavyweight drape, colour depth & collar hold.',
      },
      {
        img: 'cat4_jeans.jpg',
        cat: 'JEANS',
        tagline: 'Tailored Denim Comfort.',
        concept: '360-degree fit-check videos proving tapered cuts, luxury washes & day-to-night ease.',
      },
      {
        img: 'cat5_trousers.jpg',
        cat: 'TROUSERS',
        tagline: 'Complete Outfit Authority.',
        concept: 'Full-outfit pairing reels elevating both casual & formal shirts into ₹5,000+ looks.',
      },
    ];

    const cw = (CONTENT_W - 0.20 * 4) / 5; // 2.235
    const imgW = cw;
    const imgH = +(cw / 0.747).toFixed(3); // 2.992

    products.forEach((p, i) => {
      const px = MARGIN_X + i * (cw + 0.20);
      const py = 1.54;
      const ch = 5.26;

      addCard(s, px, py, cw, ch, {
        fill: COLORS.ivoryCard,
        border: COLORS.borderLight,
      });

      s.addImage({
        path: img(p.img),
        x: px,
        y: py,
        w: imgW,
        h: imgH,
      });

      s.addShape('rect', {
        x: px,
        y: py + imgH,
        w: cw,
        h: 0.04,
        fill: { color: COLORS.gold },
        line: { type: 'none' },
      });

      s.addText(p.cat, {
        x: px + 0.14,
        y: py + imgH + 0.14,
        w: cw - 0.28,
        h: 0.22,
        fontFace: FONTS.ui,
        fontSize: 8.8,
        bold: true,
        color: COLORS.bronze,
        charSpacing: 1.2,
        margin: 0,
      });

      s.addText(p.tagline, {
        x: px + 0.14,
        y: py + imgH + 0.40,
        w: cw - 0.28,
        h: 0.50,
        fontFace: FONTS.editorial,
        fontSize: 13,
        bold: true,
        color: COLORS.obsidian,
        lineSpacingMultiple: 1.15,
        margin: 0,
      });

      s.addText(p.concept, {
        x: px + 0.14,
        y: py + imgH + 0.96,
        w: cw - 0.28,
        h: 1.10,
        fontFace: FONTS.ui,
        fontSize: 10.2,
        color: COLORS.stone,
        lineSpacingMultiple: 1.25,
        margin: 0,
      });
    });
  }

  // =========================================================================
  // SLIDE 11 – STEP 10: MARGIN-SAFE BASKET GROWTH (Obsidian Black)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'obsidian',
      step: 'Step 10',
      kicker: 'Commercial & Margin Discipline',
      title: 'Scale Store Sales Without Compromising < 25% Margin.',
      slideNum: 11,
    });

    const kpiW = 1.90;
    const baselines = [
      { label: 'AVG PRODUCT PRICE', val: '₹2,000' },
      { label: 'AVG CUSTOMER BILL', val: '₹5,000' },
      { label: 'GROSS MARGIN', val: '< 25%' },
    ];
    baselines.forEach((b, i) => {
      const bx = MARGIN_X + i * (kpiW + 0.14);
      addCard(s, bx, 1.56, kpiW, 1.08, {
        fill: COLORS.obsidianCard,
        border: COLORS.gold,
        accentTop: true,
      });
      s.addText(b.label, {
        x: bx + 0.12,
        y: 1.70,
        w: kpiW - 0.24,
        h: 0.22,
        fontFace: FONTS.ui,
        fontSize: 8.2,
        bold: true,
        color: COLORS.lightMuted,
        charSpacing: 0.8,
        margin: 0,
      });
      s.addText(b.val, {
        x: bx + 0.12,
        y: 1.96,
        w: kpiW - 0.24,
        h: 0.50,
        fontFace: FONTS.editorial,
        fontSize: 22,
        bold: true,
        color: COLORS.gold,
        margin: 0,
      });
    });

    addCard(s, MARGIN_X, 2.80, 5.98, 3.98, {
      fill: COLORS.obsidianCard,
      border: COLORS.borderObsidian,
    });

    s.addText('BASKET VALUE PROGRESSION (₹)', {
      x: MARGIN_X + 0.20,
      y: 2.94,
      w: 5.58,
      h: 0.24,
      fontFace: FONTS.ui,
      fontSize: 9.5,
      bold: true,
      color: COLORS.gold,
      charSpacing: 1.1,
      margin: 0,
    });

    s.addText('Curated outfit styling unlocks higher gross-profit headroom per store visit.', {
      x: MARGIN_X + 0.20,
      y: 3.20,
      w: 5.58,
      h: 0.34,
      fontFace: FONTS.ui,
      fontSize: 10.5,
      color: COLORS.lightText,
      margin: 0,
    });

    s.addChart(
      pptx.ChartType.bar,
      [
        {
          name: 'Order Value (₹)',
          labels: ['Single Item\n(₹2,000 Base)', 'Current Avg Bill\n(2.5 Items)', 'Complete Outfit*\n(3-Item Styling)'],
          values: [2000, 5000, 7000],
        },
      ],
      {
        x: MARGIN_X + 0.22,
        y: 3.62,
        w: 5.54,
        h: 2.70,
        barDir: 'col',
        chartColors: [COLORS.gold],
        showLegend: false,
        showTitle: false,
        valAxisHidden: false,
        valAxisMinVal: 0,
        valAxisMaxVal: 8000,
        valAxisMajorUnit: 2000,
        valAxisLabelColor: COLORS.lightMuted,
        valAxisLabelFontSize: 9,
        catAxisLabelColor: COLORS.ivory,
        catAxisLabelFontSize: 9.5,
        catGridLine: { style: 'none' },
        valGridLine: { color: COLORS.borderObsidian, size: 0.5 },
        dataLabelColor: COLORS.ivory,
        dataLabelFontSize: 10.5,
        dataLabelFontBold: true,
        showValue: true,
        dataLabelFormatCode: '₹#,##0',
        plotArea: { fill: { color: COLORS.obsidianCard } },
      }
    );

    s.addText('*Illustrative 3-piece outfit pairing (e.g. Shirt + Trouser/Jeans + Tee).', {
      x: MARGIN_X + 0.20,
      y: 6.44,
      w: 5.58,
      h: 0.24,
      fontFace: FONTS.ui,
      fontSize: 8.5,
      italic: true,
      color: COLORS.lightMuted,
      margin: 0,
    });

    const rx = MARGIN_X + 6.26;
    const rw = CONTENT_W - 6.26;

    const principles = [
      {
        title: 'Filter High-Intent Viewers First',
        desc: 'With < 25% gross margin, focus conversion budget strictly on warm 50%+ video viewers ready to visit.',
      },
      {
        title: 'Lead With Complete Outfit Styling',
        desc: 'Educate buyers on Shirt + Trouser/Jeans pairings so walk-ins average ₹5,000+ baskets over single items.',
      },
      {
        title: 'Zero Blanket Storewide Discounts',
        desc: 'Protect luxury brand positioning—never train local buyers to wait for clearance sales.',
      },
      {
        title: 'WhatsApp Community & Referral Perks',
        desc: 'Weekly new-stock updates & member-only referral rewards drive repeat visits at near-zero ad cost.',
      },
    ];

    principles.forEach((pr, i) => {
      const py = 1.56 + i * 1.32;
      addCard(s, rx, py, rw, 1.18, {
        fill: COLORS.velvet,
        border: COLORS.borderDark,
        accentLeft: true,
        accentColor: i === 2 ? COLORS.bronze : COLORS.gold,
      });
      s.addText(pr.title, {
        x: rx + 0.22,
        y: py + 0.16,
        w: rw - 0.38,
        h: 0.28,
        fontFace: FONTS.editorial,
        fontSize: 15,
        bold: true,
        color: COLORS.gold,
        margin: 0,
      });
      s.addText(pr.desc, {
        x: rx + 0.22,
        y: py + 0.50,
        w: rw - 0.38,
        h: 0.56,
        fontFace: FONTS.ui,
        fontSize: 11.5,
        color: COLORS.lightText,
        lineSpacingMultiple: 1.24,
        margin: 0,
      });
    });
  }

  // =========================================================================
  // SLIDE 12 – STEP 11: 90-DAY STEP-BY-STEP EXECUTION (Warm Atelier Ivory)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'light',
      step: 'Step 11',
      kicker: '90-Day Rollout Plan',
      title: 'Step-by-Step Execution From Launch to Scale.',
      slideNum: 12,
    });

    s.addShape('line', {
      x: MARGIN_X + 0.4,
      y: 2.02,
      w: CONTENT_W - 0.8,
      h: 0,
      line: { color: COLORS.gold, width: 2.0 },
    });

    const phases = [
      {
        days: 'DAYS 01 – 15',
        title: 'Positioning & Video Shoot',
        themeBg: COLORS.obsidian,
        themeTxt: COLORS.ivory,
        activities: [
          'Refresh Instagram, Facebook & Google Business Profile aesthetics',
          'Set up Meta Pixel & custom Video-View Remarketing audiences',
          'Shoot fabric, fit & complete outfit reels across 5 categories',
        ],
        milestone: 'Luxury assets & tracking ready',
      },
      {
        days: 'DAYS 16 – 30',
        title: 'Video Capture & Education',
        themeBg: COLORS.ivoryCard,
        themeTxt: COLORS.obsidian,
        activities: [
          'Launch local Instagram Reels, Facebook & YouTube video ads',
          'Pool 50%+ video viewers and serve fabric/fit education ads',
          'Begin weekly WhatsApp Community new-stock drop alerts',
        ],
        milestone: 'Warm educated audience live',
      },
      {
        days: 'DAYS 31 – 60',
        title: 'High-Intent Store Visits',
        themeBg: COLORS.ivoryCard,
        themeTxt: COLORS.obsidian,
        activities: [
          'Retarget educated viewers with Maps pins & store-visit hooks',
          'Onboard walk-in buyers into the WhatsApp VIP Community',
          'Track in-store bills back to remarketing & stock alerts',
        ],
        milestone: 'High-intent walk-ins flowing',
      },
      {
        days: 'DAYS 61 – 90',
        title: 'Community & Referral Scale',
        themeBg: COLORS.velvet,
        themeTxt: COLORS.ivory,
        activities: [
          'Scale top store-visit & margin-performing video creatives',
          'Launch structured customer referral rewards via WhatsApp',
          'Expand winning pools into Secondary & Kothapet zones',
        ],
        milestone: 'Self-compounding B2C engine',
      },
    ];

    const pw = (CONTENT_W - 0.22 * 3) / 4; // 2.828
    phases.forEach((ph, i) => {
      const px = MARGIN_X + i * (pw + 0.22);

      s.addShape('ellipse', {
        x: px + pw / 2 - 0.16,
        y: 1.86,
        w: 0.32,
        h: 0.32,
        fill: { color: COLORS.gold },
        line: { color: COLORS.ivory, width: 2.0 },
      });

      s.addText(`0${i + 1}`, {
        x: px + pw / 2 - 0.16,
        y: 1.86,
        w: 0.32,
        h: 0.32,
        fontFace: FONTS.ui,
        fontSize: 9.5,
        bold: true,
        color: COLORS.obsidian,
        align: 'center',
        valign: 'middle',
        margin: 0,
      });

      const py = 2.38;
      const phH = 4.40;
      addCard(s, px, py, pw, phH, {
        fill: ph.themeBg,
        border: ph.themeBg === COLORS.ivoryCard ? COLORS.borderLight : ph.themeBg,
        accentTop: true,
        accentColor: COLORS.gold,
      });

      s.addText(ph.days, {
        x: px + 0.18,
        y: py + 0.16,
        w: pw - 0.36,
        h: 0.22,
        fontFace: FONTS.ui,
        fontSize: 9,
        bold: true,
        color: COLORS.gold,
        charSpacing: 1.2,
        margin: 0,
      });

      s.addText(ph.title, {
        x: px + 0.18,
        y: py + 0.40,
        w: pw - 0.36,
        h: 0.44,
        fontFace: FONTS.editorial,
        fontSize: 15,
        bold: true,
        color: ph.themeTxt,
        valign: 'middle',
        margin: 0,
      });

      s.addShape('line', {
        x: px + 0.18,
        y: py + 0.90,
        w: pw - 0.36,
        h: 0,
        line: { color: ph.themeBg === COLORS.ivoryCard ? COLORS.borderLight : COLORS.borderDark, width: 0.6 },
      });

      const actBullets = ph.activities.map((a) => ({
        text: a,
        options: { bullet: { code: '25AA', color: COLORS.gold } },
      }));

      s.addText(actBullets, {
        x: px + 0.18,
        y: py + 1.04,
        w: pw - 0.36,
        h: 2.38,
        fontFace: FONTS.ui,
        fontSize: 10.6,
        color: ph.themeBg === COLORS.ivoryCard ? COLORS.stone : COLORS.lightText,
        lineSpacingMultiple: 1.22,
        paraSpaceAfter: 8,
        valign: 'top',
        margin: 0,
      });

      s.addShape('rect', {
        x: px + 0.16,
        y: py + 3.56,
        w: pw - 0.32,
        h: 0.68,
        fill: { color: ph.themeBg === COLORS.ivoryCard ? COLORS.ivorySoft : '1D2330' },
        line: { color: COLORS.gold, width: 0.6 },
      });

      s.addText('MILESTONE', {
        x: px + 0.26,
        y: py + 3.63,
        w: pw - 0.52,
        h: 0.18,
        fontFace: FONTS.ui,
        fontSize: 7.5,
        bold: true,
        color: COLORS.bronze,
        charSpacing: 1.0,
        margin: 0,
      });

      s.addText(ph.milestone, {
        x: px + 0.26,
        y: py + 3.83,
        w: pw - 0.52,
        h: 0.34,
        fontFace: FONTS.ui,
        fontSize: 10.5,
        bold: true,
        color: ph.themeBg === COLORS.ivoryCard ? COLORS.obsidian : COLORS.ivory,
        margin: 0,
      });
    });
  }

  // =========================================================================
  // SLIDE 13 – STEP 12: MEASUREMENT LOOP (Midnight Velvet)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'velvet',
      step: 'Step 12',
      kicker: 'Performance Measurement',
      title: 'Measuring What Grows Store Gross Profit.',
      slideNum: 13,
    });

    addCard(s, MARGIN_X, 1.54, CONTENT_W, 1.04, {
      fill: COLORS.obsidian,
      border: COLORS.gold,
      accentLeft: true,
      accentColor: COLORS.gold,
    });

    s.addText('NORTH STAR METRIC • TRACKED WEEKLY', {
      x: MARGIN_X + 0.26,
      y: 1.68,
      w: 6.5,
      h: 0.22,
      fontFace: FONTS.ui,
      fontSize: 9.5,
      bold: true,
      color: COLORS.gold,
      charSpacing: 1.5,
      margin: 0,
    });

    s.addText('Marketing-Generated Gross Profit vs. Total Advertising Spend', {
      x: MARGIN_X + 0.26,
      y: 1.98,
      w: CONTENT_W - 0.52,
      h: 0.44,
      fontFace: FONTS.editorial,
      fontSize: 20.5,
      bold: true,
      color: COLORS.ivory,
      margin: 0,
    });

    const colW = (CONTENT_W - 0.26 * 2) / 3; // 3.817
    const groups = [
      {
        stage: 'STAGE 01 • VIDEO & REMARKETING POOL',
        color: COLORS.stone,
        metrics: [
          { name: 'Video Viewers (50%+)', desc: 'Local men entering the warm educational remarketing audience.' },
          { name: 'High-Intent Enquiries', desc: 'WhatsApp & Instagram DMs requesting fit, sizing, or fresh arrivals.' },
        ],
      },
      {
        stage: 'STAGE 02 • HIGH-INTENT STORE VISITS',
        color: COLORS.bronze,
        metrics: [
          { name: 'Maps & Store Directions', desc: 'Route taps driven by educated remarketing & Google Business Profile.' },
          { name: 'Educated Store Walk-Ins', desc: 'Pre-sold local buyers visiting to try complete outfits in store.' },
        ],
      },
      {
        stage: 'STAGE 03 • COMMERCIAL PROFITABILITY',
        color: COLORS.gold,
        metrics: [
          { name: 'Average Bill & Gross Profit', desc: 'Bills tracked against ₹5,000 benchmark and < 25% margin health.' },
          { name: 'Repeat & Referral Purchases', desc: 'Zero-CAC repeat bills from WhatsApp stock drops & member referrals.' },
        ],
      },
    ];

    groups.forEach((g, colIdx) => {
      const gx = MARGIN_X + colIdx * (colW + 0.26);
      const gy = 2.82;
      const gh = 3.94;

      addCard(s, gx, gy, colW, gh, {
        fill: COLORS.velvetCard,
        border: COLORS.borderDark,
        accentTop: true,
        accentColor: g.color === COLORS.stone ? COLORS.lightMuted : g.color,
      });

      s.addText(g.stage, {
        x: gx + 0.22,
        y: gy + 0.20,
        w: colW - 0.44,
        h: 0.26,
        fontFace: FONTS.ui,
        fontSize: 8.8,
        bold: true,
        color: g.color === COLORS.stone ? COLORS.lightMuted : g.color,
        charSpacing: 0.9,
        margin: 0,
      });

      const slotH = 1.54;
      g.metrics.forEach((m, mIdx) => {
        const my = gy + 0.62 + mIdx * (slotH + 0.14);
        s.addShape('rect', {
          x: gx + 0.22,
          y: my,
          w: colW - 0.44,
          h: slotH,
          fill: { color: COLORS.obsidian },
          line: { color: COLORS.borderObsidian, width: 0.65 },
        });
        s.addShape('rect', {
          x: gx + 0.22,
          y: my,
          w: 0.05,
          h: slotH,
          fill: { color: g.color === COLORS.stone ? COLORS.lightMuted : g.color },
          line: { type: 'none' },
        });
        s.addText(m.name, {
          x: gx + 0.40,
          y: my + 0.22,
          w: colW - 0.76,
          h: 0.34,
          fontFace: FONTS.editorial,
          fontSize: 15,
          bold: true,
          color: COLORS.ivory,
          margin: 0,
        });
        s.addText(m.desc, {
          x: gx + 0.40,
          y: my + 0.64,
          w: colW - 0.76,
          h: 0.72,
          fontFace: FONTS.ui,
          fontSize: 11.2,
          color: COLORS.lightMuted,
          lineSpacingMultiple: 1.28,
          margin: 0,
        });
      });
    });
  }

  // =========================================================================
  // SLIDE 14 – STEP 13: COMMERCIAL INVESTMENT & PACKAGE SCOPE (Obsidian Black)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'obsidian',
      step: 'Step 13',
      kicker: 'Commercial Investment & Package Scope',
      title: 'All-Inclusive Growth Retainer & Scope of Deliverables.',
      slideNum: 14,
      totalSlides: 16,
    });

    const colW = (CONTENT_W - 0.35) / 2; // 5.811
    const leftX = MARGIN_X;
    const rightX = MARGIN_X + colW + 0.35;
    const topY = 1.56;

    // LEFT COLUMN: Retainer Pricing Banner + 5 Core Services Included
    addCard(s, leftX, topY, colW, 1.05, {
      fill: COLORS.velvet,
      border: COLORS.gold,
      accentLeft: true,
      accentColor: COLORS.gold,
    });

    s.addText('ALL-INCLUSIVE GROWTH RETAINER', {
      x: leftX + 0.25,
      y: topY + 0.14,
      w: 3.5,
      h: 0.22,
      fontFace: FONTS.ui,
      fontSize: 8.5,
      bold: true,
      color: COLORS.gold,
      charSpacing: 1.2,
      margin: 0,
    });

    s.addText('Monthly Performance & Creative Management', {
      x: leftX + 0.25,
      y: topY + 0.38,
      w: 3.5,
      h: 0.48,
      fontFace: FONTS.editorial,
      fontSize: 16.5,
      bold: true,
      color: COLORS.ivory,
      margin: 0,
    });

    s.addText('₹35,000', {
      x: leftX + colW - 2.1,
      y: topY + 0.16,
      w: 1.85,
      h: 0.50,
      fontFace: FONTS.editorial,
      fontSize: 27,
      bold: true,
      color: COLORS.gold,
      align: 'right',
      margin: 0,
    });

    s.addText('PER MONTH', {
      x: leftX + colW - 2.1,
      y: topY + 0.66,
      w: 1.85,
      h: 0.22,
      fontFace: FONTS.ui,
      fontSize: 8.8,
      bold: true,
      color: COLORS.lightMuted,
      align: 'right',
      charSpacing: 1.0,
      margin: 0,
    });

    // 5 Included Services Card
    const cardY = topY + 1.20;
    const cardH = 4.10;
    addCard(s, leftX, cardY, colW, cardH, {
      fill: COLORS.obsidianCard,
      border: COLORS.borderObsidian,
    });

    s.addText('FIVE CORE SERVICES INCLUDED IN THIS PACKAGE', {
      x: leftX + 0.22,
      y: cardY + 0.14,
      w: colW - 0.44,
      h: 0.22,
      fontFace: FONTS.ui,
      fontSize: 8.5,
      bold: true,
      color: COLORS.bronze,
      charSpacing: 1.2,
      margin: 0,
    });

    const deliverables = [
      {
        num: '01',
        title: 'Client Reel Post-Production & Ad Optimization',
        desc: 'Raw video footage supplied by client. Agency provides weekly luxury shot-list guidelines, styling scripts, high-end color grading, captioning, trend audio pairing, and ad cutdowns across all 5 menswear categories.',
      },
      {
        num: '02',
        title: 'Meta & YouTube Video-View Remarketing',
        desc: 'Audience pooling of 50%+ video watchers, hyper-local Zone 1 radius targeting & sequential education ads driving store visits.',
      },
      {
        num: '03',
        title: 'Google Maps & Display Retargeting',
        desc: 'Google Business Profile optimization, weekly fresh arrival posts, Maps pin directions & display retargeting banners.',
      },
      {
        num: '04',
        title: 'WhatsApp VIP Community Management',
        desc: 'Weekly Tuesday stock drops, Friday outfit styling guides, billing counter onboarding script & margin-safe referral incentives.',
      },
      {
        num: '05',
        title: 'Commercial Analytics & Margin Guardrail',
        desc: 'Weekly scorecards tracking 50%+ video viewers, store direction taps & Marketing-Generated Gross Profit vs. Ad Spend.',
      },
    ];

    const delItemH = 0.68;
    deliverables.forEach((item, idx) => {
      const iy = cardY + 0.42 + idx * delItemH;
      if (idx > 0) {
        s.addShape('line', {
          x: leftX + 0.22,
          y: iy - 0.05,
          w: colW - 0.44,
          h: 0,
          line: { color: COLORS.borderObsidian, width: 0.5 },
        });
      }

      s.addText(item.num, {
        x: leftX + 0.22,
        y: iy,
        w: 0.35,
        h: 0.26,
        fontFace: FONTS.editorial,
        fontSize: 14,
        bold: true,
        color: COLORS.gold,
        margin: 0,
      });

      s.addText(
        [
          { text: item.title + ': ', options: { bold: true, color: COLORS.ivory } },
          { text: item.desc, options: { color: COLORS.lightText } },
        ],
        {
          x: leftX + 0.60,
          y: iy,
          w: colW - 0.82,
          h: 0.58,
          fontFace: FONTS.ui,
          fontSize: 9.8,
          lineSpacingMultiple: 1.25,
          margin: 0,
        }
      );
    });

    // RIGHT COLUMN: Commercial Transparency & Chargeable Items + ROI Break-even Card
    const rightH = 5.30;
    addCard(s, rightX, topY, colW, rightH, {
      fill: COLORS.velvet,
      border: COLORS.borderDark,
      accentTop: true,
      accentColor: COLORS.bronze,
    });

    s.addText('COMMERCIAL TRANSPARENCY', {
      x: rightX + 0.25,
      y: topY + 0.16,
      w: colW - 0.50,
      h: 0.20,
      fontFace: FONTS.ui,
      fontSize: 8.5,
      bold: true,
      color: COLORS.gold,
      charSpacing: 1.2,
      margin: 0,
    });

    s.addText('Note on Chargeable & External Services', {
      x: rightX + 0.25,
      y: topY + 0.38,
      w: colW - 0.50,
      h: 0.34,
      fontFace: FONTS.editorial,
      fontSize: 16.5,
      bold: true,
      color: COLORS.ivory,
      margin: 0,
    });

    const chargeable = [
      {
        num: '1',
        title: 'Paid Ad Spend (Paid at Actuals)',
        color: COLORS.gold,
        desc: 'Platform ad budgets for Meta (Instagram/Facebook) and Google/YouTube are paid directly by the client to ad networks at actuals. (Recommended test budget: ₹15,000 – ₹25,000/mo).',
      },
      {
        num: '2',
        title: 'Physical Print Materials',
        color: COLORS.bronze,
        desc: 'In-store physical QR counter standees, printed tent-cards, and promotional flyers are billed separately at vendor print costs.',
      },
      {
        num: '3',
        title: 'Third-Party Software / WhatsApp API',
        color: COLORS.stone,
        desc: 'If high broadcast volume requires an automated WhatsApp Business API gateway (e.g. Wati/Interakt), subscription fees are paid directly to the service provider.',
      },
      {
        num: '4',
        title: 'Optional Agency Camera Crew / On-Site Shoots',
        color: COLORS.stone,
        desc: 'The ₹35,000 retainer operates on raw video reels supplied directly by the client. Professional on-site video shooting/camera crew by agency or external model shoots are optional add-ons charged separately upon client approval.',
      },
    ];

    const cItemH = 0.82;
    chargeable.forEach((c, cIdx) => {
      const cy = topY + 0.80 + cIdx * (cItemH + 0.10);
      s.addShape('rect', {
        x: rightX + 0.22,
        y: cy,
        w: colW - 0.44,
        h: cItemH,
        fill: { color: COLORS.obsidian },
        line: { color: COLORS.borderObsidian, width: 0.65 },
      });
      s.addShape('rect', {
        x: rightX + 0.22,
        y: cy,
        w: 0.045,
        h: cItemH,
        fill: { color: c.color },
        line: { type: 'none' },
      });

      s.addText(
        [
          { text: `${c.num}. ${c.title}\n`, options: { bold: true, color: c.color, fontSize: 10 } },
          { text: c.desc, options: { color: COLORS.lightText, fontSize: 9 } },
        ],
        {
          x: rightX + 0.35,
          y: cy + 0.08,
          w: colW - 0.65,
          h: cItemH - 0.14,
          fontFace: FONTS.ui,
          lineSpacingMultiple: 1.25,
          margin: 0,
        }
      );
    });

    // ROI Break-even box at bottom of right column
    const roiY = topY + 4.54;
    const roiH = 0.64;
    s.addShape('rect', {
      x: rightX + 0.22,
      y: roiY,
      w: colW - 0.44,
      h: roiH,
      fill: { color: COLORS.obsidianCard },
      line: { color: COLORS.gold, width: 0.8 },
    });

    s.addText('ROI BREAK-EVEN BENCHMARK', {
      x: rightX + 0.34,
      y: roiY + 0.08,
      w: colW - 1.2,
      h: 0.18,
      fontFace: FONTS.ui,
      fontSize: 8.0,
      bold: true,
      color: COLORS.gold,
      charSpacing: 1.0,
      margin: 0,
    });

    s.addText('At ₹5,000 average customer bill, just 7 additional walk-in customers per month fully covers this growth retainer.', {
      x: rightX + 0.34,
      y: roiY + 0.28,
      w: colW - 1.2,
      h: 0.30,
      fontFace: FONTS.ui,
      fontSize: 9.0,
      color: COLORS.ivory,
      margin: 0,
    });

    s.addImage({
      path: img('jh_logo.png'),
      x: rightX + colW - 0.76,
      y: roiY + 0.10,
      w: 0.36,
      h: 0.45,
    });
  }

  // =========================================================================
  // SLIDE 15 – STEP 14: ONBOARDING & KICKOFF CHECKLIST (Midnight Velvet)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'velvet',
      step: 'Step 14',
      kicker: 'Client Kickoff Requirements',
      title: 'Required Documents, Ad Accounts & Creative Assets to Launch.',
      slideNum: 15,
      totalSlides: 16,
    });

    const colGap = 0.24;
    const colW = (CONTENT_W - colGap * 2) / 3; // ~3.83
    const topY = 1.54;
    const cardH = 5.25;

    // COL 1: Documents & Ad Accounts
    const col1X = MARGIN_X;
    addCard(s, col1X, topY, colW, cardH, {
      fill: COLORS.velvetCard,
      border: COLORS.borderDark,
      accentTop: true,
      accentColor: COLORS.gold,
    });

    s.addText('PILLAR 01 • OFFICIAL SETUP', {
      x: col1X + 0.22,
      y: topY + 0.16,
      w: colW - 0.44,
      h: 0.20,
      fontFace: FONTS.ui,
      fontSize: 8.5,
      bold: true,
      color: COLORS.gold,
      charSpacing: 1.2,
      margin: 0,
    });

    s.addText('Documents & Ad Accounts', {
      x: col1X + 0.22,
      y: topY + 0.38,
      w: colW - 0.44,
      h: 0.32,
      fontFace: FONTS.editorial,
      fontSize: 16.5,
      bold: true,
      color: COLORS.ivory,
      margin: 0,
    });

    const col1Items = [
      {
        num: '1',
        title: 'Business Verification Proofs',
        color: COLORS.gold,
        desc: 'GST Certificate / MSME Udyam registration & Business PAN for Meta Business Portfolio verification.',
      },
      {
        num: '2',
        title: 'Store Address & ID Proof',
        color: COLORS.bronze,
        desc: 'Store electricity bill / rental deed for B.N. Reddy Nagar premises + authorized proprietor ID (Aadhaar/PAN).',
      },
      {
        num: '3',
        title: 'Ad Spend Payment Method',
        color: COLORS.stone,
        desc: 'Business credit/debit card enabled for international / online auto-pay billing on Meta Ads & Google Ads.',
      },
      {
        num: '4',
        title: 'GSTIN for Tax Credit',
        color: COLORS.stone,
        desc: 'Added to ad billing profiles to claim 18% GST input tax credit (ITC) on all advertising invoices.',
      },
    ];

    const itemH = 0.88;
    col1Items.forEach((it, idx) => {
      const iy = topY + 0.78 + idx * (itemH + 0.10);
      s.addShape('rect', {
        x: col1X + 0.20,
        y: iy,
        w: colW - 0.40,
        h: itemH,
        fill: { color: COLORS.obsidian },
        line: { color: COLORS.borderObsidian, width: 0.65 },
      });
      s.addShape('rect', {
        x: col1X + 0.20,
        y: iy,
        w: 0.045,
        h: itemH,
        fill: { color: it.color },
        line: { type: 'none' },
      });

      s.addText(
        [
          { text: `${it.num}. ${it.title}\n`, options: { bold: true, color: it.color, fontSize: 9.8 } },
          { text: it.desc, options: { color: COLORS.lightText, fontSize: 8.8 } },
        ],
        {
          x: col1X + 0.32,
          y: iy + 0.06,
          w: colW - 0.56,
          h: itemH - 0.10,
          fontFace: FONTS.ui,
          lineSpacingMultiple: 1.22,
          margin: 0,
        }
      );
    });

    s.addShape('rect', {
      x: col1X + 0.20,
      y: topY + cardH - 0.48,
      w: colW - 0.40,
      h: 0.34,
      fill: { color: COLORS.obsidian },
      line: { color: COLORS.borderObsidian, width: 0.65 },
    });
    s.addText('Agency Role: We handle complete Portfolio & Account setup.', {
      x: col1X + 0.26,
      y: topY + cardH - 0.46,
      w: colW - 0.52,
      h: 0.30,
      fontFace: FONTS.ui,
      fontSize: 8.5,
      color: COLORS.lightMuted,
      margin: 0,
    });

    // COL 2: Raw Reels From Client
    const col2X = col1X + colW + colGap;
    addCard(s, col2X, topY, colW, cardH, {
      fill: COLORS.obsidianCard,
      border: COLORS.borderObsidian,
      accentTop: true,
      accentColor: COLORS.gold,
    });

    s.addText('PILLAR 02 • INITIAL CREATIVE BATCH', {
      x: col2X + 0.22,
      y: topY + 0.16,
      w: colW - 0.44,
      h: 0.20,
      fontFace: FONTS.ui,
      fontSize: 8.5,
      bold: true,
      color: COLORS.gold,
      charSpacing: 1.2,
      margin: 0,
    });

    s.addText('Raw Reels From Client', {
      x: col2X + 0.22,
      y: topY + 0.38,
      w: colW - 0.44,
      h: 0.32,
      fontFace: FONTS.editorial,
      fontSize: 16.5,
      bold: true,
      color: COLORS.ivory,
      margin: 0,
    });

    const col2Items = [
      {
        num: '1',
        title: 'Casual & Formal Shirts (4 clips)',
        color: COLORS.gold,
        desc: '9:16 vertical clips showing fabric weave, collar stiffness, rolled sleeves & trial drape.',
      },
      {
        num: '2',
        title: 'Jeans & Trousers (4 clips)',
        color: COLORS.gold,
        desc: 'Movement clips showing denim stretch, waistband quality, ankle taper & fit checks.',
      },
      {
        num: '3',
        title: 'T-Shirts & Everyday Outfits (2 clips)',
        color: COLORS.bronze,
        desc: 'Clean lookbook clips highlighting heavyweight cotton drop & casual pairings.',
      },
      {
        num: '4',
        title: 'Store Exterior & Walk-In Reel (1–2 clips)',
        color: COLORS.bronze,
        desc: 'Facade, entrance near B.D. Reddy Gardens, interior lighting & welcome experience.',
      },
    ];

    col2Items.forEach((it, idx) => {
      const iy = topY + 0.78 + idx * (itemH + 0.10);
      s.addShape('rect', {
        x: col2X + 0.20,
        y: iy,
        w: colW - 0.40,
        h: itemH,
        fill: { color: COLORS.obsidian },
        line: { color: COLORS.borderObsidian, width: 0.65 },
      });
      s.addShape('rect', {
        x: col2X + 0.20,
        y: iy,
        w: 0.045,
        h: itemH,
        fill: { color: it.color },
        line: { type: 'none' },
      });

      s.addText(
        [
          { text: `${it.title}\n`, options: { bold: true, color: it.color, fontSize: 9.8 } },
          { text: it.desc, options: { color: COLORS.lightText, fontSize: 8.8 } },
        ],
        {
          x: col2X + 0.32,
          y: iy + 0.06,
          w: colW - 0.56,
          h: itemH - 0.10,
          fontFace: FONTS.ui,
          lineSpacingMultiple: 1.22,
          margin: 0,
        }
      );
    });

    s.addShape('rect', {
      x: col2X + 0.20,
      y: topY + cardH - 0.48,
      w: colW - 0.40,
      h: 0.34,
      fill: { color: COLORS.velvet },
      line: { color: COLORS.gold, width: 0.65 },
    });
    s.addText('Agency Role: Mobile filming guide, luxury grade & ad cutdowns.', {
      x: col2X + 0.26,
      y: topY + cardH - 0.46,
      w: colW - 0.52,
      h: 0.30,
      fontFace: FONTS.ui,
      fontSize: 8.5,
      color: COLORS.ivory,
      margin: 0,
    });

    // COL 3: Access & In-Store Setup
    const col3X = col2X + colW + colGap;
    addCard(s, col3X, topY, colW, cardH, {
      fill: COLORS.velvetCard,
      border: COLORS.borderDark,
      accentTop: true,
      accentColor: COLORS.bronze,
    });

    s.addText('PILLAR 03 • PLATFORM DELEGATION', {
      x: col3X + 0.22,
      y: topY + 0.16,
      w: colW - 0.44,
      h: 0.20,
      fontFace: FONTS.ui,
      fontSize: 8.5,
      bold: true,
      color: COLORS.bronze,
      charSpacing: 1.2,
      margin: 0,
    });

    s.addText('Access & In-Store Setup', {
      x: col3X + 0.22,
      y: topY + 0.38,
      w: colW - 0.44,
      h: 0.32,
      fontFace: FONTS.editorial,
      fontSize: 16.5,
      bold: true,
      color: COLORS.ivory,
      margin: 0,
    });

    const col3Items = [
      {
        num: '1',
        title: 'Meta & Instagram Admin Access',
        color: COLORS.bronze,
        desc: 'Invite agency email as Partner / Admin in Meta Business Portfolio and connect Instagram account.',
      },
      {
        num: '2',
        title: 'Google Business Profile (Maps)',
        color: COLORS.gold,
        desc: 'Grant Manager access to optimize LB Nagar store location pins, timings, photos & reviews.',
      },
      {
        num: '3',
        title: 'WhatsApp In-Store Phone/SIM',
        color: COLORS.stone,
        desc: 'Designate store smartphone for VIP Community broadcasts & customer enquiry management.',
      },
      {
        num: '4',
        title: 'Billing Counter Scripting',
        color: COLORS.stone,
        desc: 'Store cash-counter team brief on collecting customer numbers during ₹5,000+ billing.',
      },
    ];

    col3Items.forEach((it, idx) => {
      const iy = topY + 0.78 + idx * (itemH + 0.10);
      s.addShape('rect', {
        x: col3X + 0.20,
        y: iy,
        w: colW - 0.40,
        h: itemH,
        fill: { color: COLORS.obsidian },
        line: { color: COLORS.borderObsidian, width: 0.65 },
      });
      s.addShape('rect', {
        x: col3X + 0.20,
        y: iy,
        w: 0.045,
        h: itemH,
        fill: { color: it.color },
        line: { type: 'none' },
      });

      s.addText(
        [
          { text: `${it.num}. ${it.title}\n`, options: { bold: true, color: it.color, fontSize: 9.8 } },
          { text: it.desc, options: { color: COLORS.lightText, fontSize: 8.8 } },
        ],
        {
          x: col3X + 0.32,
          y: iy + 0.06,
          w: colW - 0.56,
          h: itemH - 0.10,
          fontFace: FONTS.ui,
          lineSpacingMultiple: 1.22,
          margin: 0,
        }
      );
    });

    s.addShape('rect', {
      x: col3X + 0.20,
      y: topY + cardH - 0.48,
      w: colW - 0.40,
      h: 0.34,
      fill: { color: COLORS.obsidian },
      line: { color: COLORS.gold, width: 0.65 },
    });
    s.addText('Timeline: First campaigns live in 3–5 business days.', {
      x: col3X + 0.26,
      y: topY + cardH - 0.46,
      w: colW - 0.52,
      h: 0.30,
      fontFace: FONTS.ui,
      fontSize: 8.5,
      bold: true,
      color: COLORS.gold,
      margin: 0,
    });
  }

  // =========================================================================
  // SLIDE 16 – CLOSING (Obsidian Black)
  // =========================================================================
  {
    const s = pptx.addSlide();
    s.background = { color: COLORS.obsidian };

    s.addImage({
      path: img('slide14_closing.jpg'),
      x: 7.733,
      y: 0,
      w: 5.60,
      h: 7.5,
    });

    s.addShape('line', {
      x: 7.733,
      y: 0.60,
      w: 0,
      h: 6.30,
      line: { color: COLORS.gold, width: 0.9 },
    });

    s.addImage({
      path: img('jh_logo.png'),
      x: 0.78,
      y: 0.48,
      w: 0.48,
      h: 0.58,
    });

    s.addText('JHONTY HUB', {
      x: 1.36,
      y: 0.48,
      w: 4.5,
      h: 0.32,
      fontFace: FONTS.editorial,
      fontSize: 22,
      bold: true,
      color: COLORS.ivory,
      charSpacing: 2.0,
      margin: 0,
    });

    s.addText('PARTNERSHIP NEXT STEPS', {
      x: 1.38,
      y: 0.80,
      w: 4.5,
      h: 0.22,
      fontFace: FONTS.ui,
      fontSize: 8.8,
      bold: true,
      color: COLORS.gold,
      charSpacing: 2.0,
      margin: 0,
    });

    s.addText("Let's Build Jhonty Hub's\nNext Luxury Chapter.", {
      x: 0.78,
      y: 1.28,
      w: 6.50,
      h: 1.15,
      fontFace: FONTS.editorial,
      fontSize: 27,
      bold: true,
      color: COLORS.ivory,
      lineSpacingMultiple: 1.2,
      margin: 0,
    });

    const stages = [
      {
        label: 'UNDERSTAND',
        desc: 'Align on store margins, category priorities & local male buyers.',
      },
      {
        label: 'BUILD',
        desc: 'Provide weekly creative guidelines for client reels, launch remarketing & WhatsApp tracking.',
      },
      {
        label: 'GROW',
        desc: 'Convert high-intent walk-ins into ₹5,000+ outfits & repeat referrals.',
      },
    ];

    stages.forEach((st, i) => {
      const sy = 2.68 + i * 0.94;
      addCard(s, 0.78, sy, 6.35, 0.80, {
        fill: COLORS.obsidianCard,
        border: COLORS.borderObsidian,
        accentLeft: true,
        accentColor: COLORS.gold,
      });

      s.addText(`0${i + 1}`, {
        x: 0.98,
        y: sy + 0.18,
        w: 0.44,
        h: 0.44,
        fontFace: FONTS.editorial,
        fontSize: 18,
        bold: true,
        color: COLORS.gold,
        valign: 'middle',
        margin: 0,
      });

      s.addText(st.label, {
        x: 1.50,
        y: sy + 0.10,
        w: 5.40,
        h: 0.26,
        fontFace: FONTS.ui,
        fontSize: 13,
        bold: true,
        color: COLORS.ivory,
        charSpacing: 1.2,
        margin: 0,
      });

      s.addText(st.desc, {
        x: 1.50,
        y: sy + 0.38,
        w: 5.40,
        h: 0.34,
        fontFace: FONTS.ui,
        fontSize: 11.2,
        color: COLORS.lightText,
        margin: 0,
      });
    });

    s.addShape('line', {
      x: 0.78,
      y: 5.76,
      w: 6.35,
      h: 0,
      line: { color: COLORS.gold, width: 0.8 },
    });

    s.addText(
      '"Our goal is clear: educate local men through compelling video, position Jhonty Hub as LB Nagar’s luxury destination, and turn high intent into store visits."',
      {
        x: 0.78,
        y: 5.92,
        w: 6.35,
        h: 0.88,
        fontFace: FONTS.editorial,
        fontSize: 12.5,
        italic: true,
        color: COLORS.ivory,
        lineSpacingMultiple: 1.3,
        margin: 0,
      }
    );

    s.addText('Jhonty Hub • Beside B.D. Reddy Gardens, B.N. Reddy Nagar, LB Nagar, Hyderabad – 500079', {
      x: 0.78,
      y: 7.00,
      w: 5.2,
      h: 0.26,
      fontFace: FONTS.ui,
      fontSize: 9,
      color: COLORS.lightMuted,
      margin: 0,
    });

    s.addText('16  /  16', {
      x: 6.15,
      y: 7.00,
      w: 0.98,
      h: 0.26,
      fontFace: FONTS.ui,
      fontSize: 9.5,
      bold: true,
      color: COLORS.gold,
      align: 'right',
      margin: 0,
    });
  }

  await pptx.writeFile({ fileName: OUTPUT_PPTX });
  console.log(`Generated: ${OUTPUT_PPTX}`);
}

buildPresentation().catch((err) => {
  console.error(err);
  process.exit(1);
});
