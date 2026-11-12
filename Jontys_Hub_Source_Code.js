/**
 * JONTYS HUB – PREMIUM BUSINESS GROWTH & PERFORMANCE MARKETING PRESENTATION
 * Editable 14-Slide Widescreen (16:9) Deck Generator via PptxGenJS
 *
 * Design System: PREMIUM BUSINESS GROWTH
 * Palette:
 *   Deep Midnight Navy: #111827
 *   Charcoal Black:     #171717
 *   Warm Ivory:         #F7F5F0
 *   Champagne Gold:     #C6A66B
 *   Muted Bronze:       #A88B5A
 *   Slate Grey:         #667085
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

const ASSETS_DIR = path.join(__dirname, 'assets');
const OUTPUT_PPTX = path.join(__dirname, 'Jontys_Hub_Premium_Growth_Strategy.pptx');

// Design Tokens
const COLORS = {
  navy: '111827',
  navyCard: '1B2436',
  charcoal: '171717',
  charcoalCard: '222222',
  ivory: 'F7F5F0',
  ivoryCard: 'FFFFFF',
  ivorySoft: 'EFECE4',
  gold: 'C6A66B',
  bronze: 'A88B5A',
  slate: '667085',
  lightMuted: '9CA3AF',
  lightText: 'E5E7EB',
  borderLight: 'E0DBD0',
  borderDark: '2D3748',
  borderCharcoal: '333333',
  successMuted: '3F6A58',
  warnMuted: '9A4A3D',
};

const FONTS = {
  heading: 'Montserrat',
  body: 'Montserrat',
};

const W = 13.333;
const H = 7.5;
const MARGIN_X = 0.65;
const CONTENT_W = W - MARGIN_X * 2; // 12.033

const img = (name) => path.join(ASSETS_DIR, name);

function addSlideChrome(slide, opts) {
  const {
    theme = 'light', // 'light' | 'navy' | 'charcoal'
    kicker = '',
    title = '',
    slideNum = 2,
    totalSlides = 15,
  } = opts;

  const isDark = theme !== 'light';
  const bg = theme === 'navy' ? COLORS.navy : theme === 'charcoal' ? COLORS.charcoal : COLORS.ivory;
  const titleColor = isDark ? COLORS.ivory : COLORS.navy;
  const footerText = isDark ? COLORS.lightMuted : COLORS.slate;
  const dividerColor = isDark ? (theme === 'navy' ? COLORS.borderDark : COLORS.borderCharcoal) : COLORS.borderLight;

  slide.background = { color: bg };

  // Top subtle accent bar on left margin
  slide.addShape('rect', {
    x: MARGIN_X,
    y: 0.42,
    w: 0.36,
    h: 0.035,
    fill: { color: COLORS.gold },
    line: { type: 'none' },
  });

  // Kicker label
  slide.addText(kicker.toUpperCase(), {
    x: MARGIN_X + 0.48,
    y: 0.31,
    w: 8.5,
    h: 0.25,
    fontFace: FONTS.body,
    fontSize: 10,
    bold: true,
    color: COLORS.gold,
    charSpacing: 2.2,
    valign: 'middle',
    margin: 0,
  });

  // Top right brand mark
  slide.addText('JONTYS HUB  |  GROWTH STRATEGY', {
    x: W - MARGIN_X - 4.2,
    y: 0.31,
    w: 4.2,
    h: 0.25,
    fontFace: FONTS.body,
    fontSize: 9.5,
    color: footerText,
    align: 'right',
    charSpacing: 1.5,
    valign: 'middle',
    margin: 0,
  });

  // Main slide headline
  slide.addText(title, {
    x: MARGIN_X,
    y: 0.64,
    w: CONTENT_W,
    h: 0.62,
    fontFace: FONTS.heading,
    fontSize: 27,
    bold: true,
    color: titleColor,
    valign: 'top',
    margin: 0,
  });

  // Thin structural hairline below header
  slide.addShape('line', {
    x: MARGIN_X,
    y: 1.32,
    w: CONTENT_W,
    h: 0,
    line: { color: dividerColor, width: 0.75 },
  });

  // Footer hairline & metadata
  slide.addShape('line', {
    x: MARGIN_X,
    y: 7.02,
    w: CONTENT_W,
    h: 0,
    line: { color: dividerColor, width: 0.6 },
  });

  slide.addText('Jontys Hub  •  B.N. Reddy Nagar, LB Nagar, Hyderabad – 500079  •  Confidential Client Proposal', {
    x: MARGIN_X,
    y: 7.08,
    w: 9.5,
    h: 0.28,
    fontFace: FONTS.body,
    fontSize: 9.5,
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
    fontFace: FONTS.heading,
    fontSize: 10,
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
    line: { color: border, width: 0.75 },
  });

  if (accentTop) {
    slide.addShape('rect', {
      x,
      y,
      w,
      h: 0.045,
      fill: { color: accentColor },
      line: { type: 'none' },
    });
  }
  if (accentLeft) {
    slide.addShape('rect', {
      x,
      y,
      w: 0.055,
      h,
      fill: { color: accentColor },
      line: { type: 'none' },
    });
  }
}

async function buildPresentation() {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'Strategic Growth & Performance Marketing Advisory';
  pptx.company = 'Jontys Hub';
  pptx.subject = 'Premium Business Growth & Performance Marketing Strategy';
  pptx.title = 'Jontys Hub – Premium Business Growth Strategy';

  // =========================================================================
  // SLIDE 01 – COVER (Charcoal Black)
  // =========================================================================
  {
    const s = pptx.addSlide();
    s.background = { color: COLORS.charcoal };

    // Right editorial portrait (exact 3:4 aspect ratio -> 5.60 x 7.50)
    s.addImage({
      path: img('slide01_cover.jpg'),
      x: 7.733,
      y: 0,
      w: 5.60,
      h: 7.5,
    });

    // Subtle vertical gold hairline framing the portrait seam
    s.addShape('line', {
      x: 7.733,
      y: 0.65,
      w: 0,
      h: 6.2,
      line: { color: COLORS.gold, width: 1.0 },
    });

    // Left dark editorial zone
    s.addShape('rect', {
      x: 0.75,
      y: 0.85,
      w: 0.55,
      h: 0.04,
      fill: { color: COLORS.gold },
      line: { type: 'none' },
    });

    s.addText('CLIENT ACQUISITION & GROWTH PROPOSAL', {
      x: 0.75,
      y: 1.05,
      w: 6.4,
      h: 0.3,
      fontFace: FONTS.body,
      fontSize: 10.5,
      bold: true,
      color: COLORS.gold,
      charSpacing: 2.5,
      margin: 0,
    });

    s.addText('JONTYS HUB', {
      x: 0.75,
      y: 1.85,
      w: 6.5,
      h: 0.95,
      fontFace: FONTS.heading,
      fontSize: 48,
      bold: true,
      color: COLORS.ivory,
      charSpacing: 1.5,
      margin: 0,
    });

    s.addShape('rect', {
      x: 0.75,
      y: 2.95,
      w: 1.25,
      h: 0.045,
      fill: { color: COLORS.gold },
      line: { type: 'none' },
    });

    s.addText(
      "Building a Premium Men's Fashion Brand\nThrough Performance Marketing.",
      {
        x: 0.75,
        y: 3.30,
        w: 6.55,
        h: 1.10,
        fontFace: FONTS.heading,
        fontSize: 20.5,
        bold: true,
        color: COLORS.ivory,
        lineSpacingMultiple: 1.25,
        margin: 0,
      }
    );

    s.addText(
      'A commercially disciplined partnership blueprint engineered to elevate brand authority, drive qualified local footfall, and protect store-level gross margins.',
      {
        x: 0.75,
        y: 4.65,
        w: 6.1,
        h: 0.95,
        fontFace: FONTS.body,
        fontSize: 14,
        color: COLORS.lightMuted,
        lineSpacingMultiple: 1.38,
        margin: 0,
      }
    );

    // Location & metadata block at bottom left
    s.addShape('line', {
      x: 0.75,
      y: 6.08,
      w: 6.1,
      h: 0,
      line: { color: COLORS.borderCharcoal, width: 0.75 },
    });

    s.addText('Strategic Growth Proposal  |  LB Nagar, Hyderabad', {
      x: 0.75,
      y: 6.24,
      w: 6.1,
      h: 0.34,
      fontFace: FONTS.heading,
      fontSize: 13.5,
      bold: true,
      color: COLORS.gold,
      margin: 0,
    });

    s.addText('Beside B.D. Reddy Gardens, B.N. Reddy Nagar, Hyderabad, Telangana – 500079', {
      x: 0.75,
      y: 6.60,
      w: 6.1,
      h: 0.3,
      fontFace: FONTS.body,
      fontSize: 10.5,
      color: COLORS.lightMuted,
      margin: 0,
    });

    s.addText('01  /  15', {
      x: 0.75,
      y: 7.05,
      w: 1.5,
      h: 0.25,
      fontFace: FONTS.heading,
      fontSize: 10,
      bold: true,
      color: COLORS.bronze,
      margin: 0,
    });
  }

  // =========================================================================
  // SLIDE 02 – BUSINESS OVERVIEW (Warm Ivory)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'light',
      kicker: 'Business Foundation',
      title: "A Premium Men's Fashion Destination.",
      slideNum: 2,
    });

    // Top-left photography (16:9 image -> 6.30 x 3.52)
    s.addImage({
      path: img('slide02_store_display.jpg'),
      x: MARGIN_X,
      y: 1.58,
      w: 6.30,
      h: 3.52,
    });

    // Subtle caption bar under image
    addCard(s, MARGIN_X, 5.22, 6.30, 1.56, {
      fill: COLORS.ivoryCard,
      border: COLORS.borderLight,
      accentLeft: true,
    });

    s.addText('BRAND POSITIONING & STORE ANCHOR', {
      x: MARGIN_X + 0.24,
      y: 5.36,
      w: 5.8,
      h: 0.25,
      fontFace: FONTS.body,
      fontSize: 10,
      bold: true,
      color: COLORS.bronze,
      charSpacing: 1.6,
      margin: 0,
    });

    s.addText(
      "Dedicated exclusively to premium men's apparel in LB Nagar — engineered around curated wardrobing rather than mass-market discounting.",
      {
        x: MARGIN_X + 0.24,
        y: 5.64,
        w: 5.8,
        h: 0.56,
        fontFace: FONTS.body,
        fontSize: 11.8,
        color: COLORS.navy,
        lineSpacingMultiple: 1.25,
        margin: 0,
      }
    );

    s.addText('Location: Beside B.D. Reddy Gardens, B.N. Reddy Nagar, LB Nagar, Hyderabad – 500079', {
      x: MARGIN_X + 0.24,
      y: 6.32,
      w: 5.8,
      h: 0.34,
      fontFace: FONTS.body,
      fontSize: 10,
      color: COLORS.slate,
      margin: 0,
    });

    // Right column: 3 Prominent Commercial KPI Cards + 5 Product Categories
    const rx = 7.22;
    const rw = 5.46;

    // 3 Metric Cards side-by-side at top right
    const kpis = [
      { label: 'AVG PRODUCT PRICE', val: '₹2,000', sub: 'Premium entry price point' },
      { label: 'AVG CUSTOMER BILL', val: '₹5,000', sub: 'Multi-item basket strength' },
      { label: 'GROSS MARGIN', val: '< 25%', sub: 'Strict CAC & discount discipline' },
    ];
    const kw = (rw - 0.30) / 3;
    kpis.forEach((k, i) => {
      const kx = rx + i * (kw + 0.15);
      addCard(s, kx, 1.58, kw, 1.62, {
        fill: COLORS.navy,
        border: COLORS.navy,
        accentTop: true,
        accentColor: COLORS.gold,
      });
      s.addText(k.label, {
        x: kx + 0.14,
        y: 1.74,
        w: kw - 0.28,
        h: 0.28,
        fontFace: FONTS.body,
        fontSize: 9,
        bold: true,
        color: COLORS.lightMuted,
        charSpacing: 1.0,
        margin: 0,
      });
      s.addText(k.val, {
        x: kx + 0.14,
        y: 2.08,
        w: kw - 0.28,
        h: 0.56,
        fontFace: FONTS.heading,
        fontSize: 25,
        bold: true,
        color: COLORS.gold,
        margin: 0,
      });
      s.addText(k.sub, {
        x: kx + 0.14,
        y: 2.70,
        w: kw - 0.28,
        h: 0.38,
        fontFace: FONTS.body,
        fontSize: 10,
        color: COLORS.lightText,
        lineSpacingMultiple: 1.15,
        margin: 0,
      });
    });

    // Core Product Assortment Card
    addCard(s, rx, 3.40, rw, 3.38, {
      fill: COLORS.ivoryCard,
      border: COLORS.borderLight,
    });

    s.addText('FIVE CORE MENSWEAR CATEGORIES', {
      x: rx + 0.25,
      y: 3.58,
      w: rw - 0.5,
      h: 0.28,
      fontFace: FONTS.body,
      fontSize: 10.5,
      bold: true,
      color: COLORS.bronze,
      charSpacing: 1.5,
      margin: 0,
    });

    const cats = [
      { name: 'Casual Shirts', desc: 'Weekend, smart-casual & textured contemporary weaves' },
      { name: 'Formal Shirts', desc: 'Tailored crisp workwear & occasion silhouettes' },
      { name: 'T-Shirts', desc: 'Elevated everyday crew-necks & structured basics' },
      { name: 'Jeans', desc: 'Clean-wash, modern tapered & stretch denim fits' },
      { name: 'Trousers', desc: 'Tailored office, chino & dress trouser staples' },
    ];

    cats.forEach((c, idx) => {
      const cy = 3.96 + idx * 0.54;
      s.addShape('rect', {
        x: rx + 0.25,
        y: cy + 0.08,
        w: 0.28,
        h: 0.28,
        fill: { color: COLORS.ivorySoft },
        line: { color: COLORS.gold, width: 0.75 },
      });
      s.addText(`0${idx + 1}`, {
        x: rx + 0.25,
        y: cy + 0.08,
        w: 0.28,
        h: 0.28,
        fontFace: FONTS.heading,
        fontSize: 9.5,
        bold: true,
        color: COLORS.navy,
        align: 'center',
        valign: 'middle',
        margin: 0,
      });
      s.addText(c.name, {
        x: rx + 0.65,
        y: cy + 0.02,
        w: 1.65,
        h: 0.36,
        fontFace: FONTS.heading,
        fontSize: 13.5,
        bold: true,
        color: COLORS.navy,
        valign: 'middle',
        margin: 0,
      });
      s.addText(c.desc, {
        x: rx + 2.30,
        y: cy + 0.02,
        w: 2.95,
        h: 0.38,
        fontFace: FONTS.body,
        fontSize: 11.5,
        color: COLORS.slate,
        valign: 'middle',
        margin: 0,
      });
      if (idx < cats.length - 1) {
        s.addShape('line', {
          x: rx + 0.25,
          y: cy + 0.46,
          w: rw - 0.5,
          h: 0,
          line: { color: COLORS.borderLight, width: 0.5 },
        });
      }
    });
  }

  // =========================================================================
  // SLIDE 03 – MARKET OPPORTUNITY (Deep Midnight Navy)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'navy',
      kicker: 'Category & Consumer Shifts',
      title: "The Future of Men's Fashion Is Style, Quality & Experience.",
      slideNum: 3,
    });

    // Left 5 structured market opportunity drivers
    const drivers = [
      {
        num: '01',
        title: 'Video-First Discovery',
        body: 'Local men watch short-form Instagram Reels & YouTube ads to shortlist styles before ever stepping into a store.',
      },
      {
        num: '02',
        title: 'Product Education Wins',
        body: 'Showing fabric quality, collar structure & fit builds premium value far faster than price-led banners.',
      },
      {
        num: '03',
        title: 'Dedicated Men-Only Trust',
        body: 'Discerning buyers prefer focused menswear stores for personalised sizing and curated outfit pairings.',
      },
      {
        num: '04',
        title: 'High-Intent Local Walk-Ins',
        body: 'Remarketing educated viewers with Maps directions turns casual video watchers into store visitors.',
      },
      {
        num: '05',
        title: 'Community-Led Retention',
        body: 'WhatsApp stock updates & referral discounts keep buyers returning without costly paid re-acquisition.',
      },
    ];

    drivers.forEach((d, i) => {
      const dy = 1.56 + i * 1.04;
      addCard(s, MARGIN_X, dy, 6.75, 0.92, {
        fill: COLORS.navyCard,
        border: COLORS.borderDark,
        accentLeft: true,
        accentColor: COLORS.gold,
      });
      s.addText(d.num, {
        x: MARGIN_X + 0.20,
        y: dy + 0.16,
        w: 0.45,
        h: 0.4,
        fontFace: FONTS.heading,
        fontSize: 16,
        bold: true,
        color: COLORS.gold,
        margin: 0,
      });
      s.addText(d.title, {
        x: MARGIN_X + 0.72,
        y: dy + 0.12,
        w: 5.8,
        h: 0.28,
        fontFace: FONTS.heading,
        fontSize: 14,
        bold: true,
        color: COLORS.ivory,
        margin: 0,
      });
      s.addText(d.body, {
        x: MARGIN_X + 0.72,
        y: dy + 0.42,
        w: 5.8,
        h: 0.42,
        fontFace: FONTS.body,
        fontSize: 11.5,
        color: COLORS.lightText,
        lineSpacingMultiple: 1.2,
        margin: 0,
      });
    });

    // Right column: 4:3 Editorial photo (5.00 x 3.74) + Strategic Takeaway Card
    const rx = 7.68;
    const rw = 5.00;
    s.addImage({
      path: img('slide03_shopping.jpg'),
      x: rx,
      y: 1.56,
      w: rw,
      h: 3.74,
    });

    addCard(s, rx, 5.46, rw, 1.30, {
      fill: COLORS.charcoal,
      border: COLORS.gold,
      accentTop: true,
    });

    s.addText('STRATEGIC MARKET INSIGHT', {
      x: rx + 0.22,
      y: 5.60,
      w: rw - 0.44,
      h: 0.24,
      fontFace: FONTS.body,
      fontSize: 9.5,
      bold: true,
      color: COLORS.gold,
      charSpacing: 1.6,
      margin: 0,
    });

    s.addText(
      'Local menswear shoppers already buy across ₹2,000+ price tiers when digital inspiration smoothly guides them to a high-trust physical try-on experience.',
      {
        x: rx + 0.22,
        y: 5.88,
        w: rw - 0.44,
        h: 0.76,
        fontFace: FONTS.body,
        fontSize: 12,
        color: COLORS.ivory,
        lineSpacingMultiple: 1.28,
        margin: 0,
      }
    );
  }

  // =========================================================================
  // SLIDE 04 – CURRENT BUSINESS SITUATION (Warm Ivory)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'light',
      kicker: 'Diagnostic & Growth Gaps',
      title: 'Connecting Digital Discovery With Physical Store Sales.',
      slideNum: 4,
    });

    // Left panel: Verified Business Facts (clearly separated)
    addCard(s, MARGIN_X, 1.56, 3.65, 5.20, {
      fill: COLORS.navy,
      border: COLORS.navy,
      accentTop: true,
    });

    s.addText('VERIFIED BUSINESS FACTS', {
      x: MARGIN_X + 0.24,
      y: 1.78,
      w: 3.17,
      h: 0.26,
      fontFace: FONTS.body,
      fontSize: 10,
      bold: true,
      color: COLORS.gold,
      charSpacing: 1.5,
      margin: 0,
    });

    s.addText('Confirmed operational baseline prior to Phase-1 audit:', {
      x: MARGIN_X + 0.24,
      y: 2.08,
      w: 3.17,
      h: 0.44,
      fontFace: FONTS.body,
      fontSize: 11.5,
      color: COLORS.lightMuted,
      margin: 0,
    });

    const facts = [
      { k: 'Store Format', v: "Pure-play Premium Men's Fashion Retail" },
      { k: 'Catchment Anchor', v: 'B.N. Reddy Nagar, LB Nagar, Hyderabad' },
      { k: 'Merchandise Mix', v: 'Casual & Formal Shirts, T-shirts, Jeans, Trousers' },
      { k: 'Avg Product Price', v: '₹2,000 per piece' },
      { k: 'Avg Customer Bill', v: '₹5,000 (strong multi-piece basket)' },
      { k: 'Margin Profile', v: 'Below 25% gross margin constraint' },
    ];

    facts.forEach((f, i) => {
      const fy = 2.68 + i * 0.65;
      s.addText(f.k.toUpperCase(), {
        x: MARGIN_X + 0.24,
        y: fy,
        w: 3.17,
        h: 0.20,
        fontFace: FONTS.body,
        fontSize: 9,
        bold: true,
        color: COLORS.gold,
        charSpacing: 0.8,
        margin: 0,
      });
      s.addText(f.v, {
        x: MARGIN_X + 0.24,
        y: fy + 0.21,
        w: 3.17,
        h: 0.34,
        fontFace: FONTS.body,
        fontSize: 11.5,
        color: COLORS.ivory,
        margin: 0,
      });
      if (i < facts.length - 1) {
        s.addShape('line', {
          x: MARGIN_X + 0.24,
          y: fy + 0.59,
          w: 3.17,
          h: 0,
          line: { color: COLORS.borderDark, width: 0.5 },
        });
      }
    });

    // Right panel header: Preliminary Opportunities to Audit & Build
    const gx = MARGIN_X + 3.92;
    const gw = CONTENT_W - 3.92; // 8.113

    s.addShape('rect', {
      x: gx,
      y: 1.56,
      w: gw,
      h: 0.42,
      fill: { color: COLORS.ivorySoft },
      line: { color: COLORS.gold, width: 0.75 },
    });

    s.addText('PRELIMINARY STRATEGIC IMPROVEMENT AREAS (TO VALIDATE IN DAYS 1–15 AUDIT)', {
      x: gx + 0.18,
      y: 1.56,
      w: gw - 0.36,
      h: 0.42,
      fontFace: FONTS.body,
      fontSize: 10,
      bold: true,
      color: COLORS.navy,
      charSpacing: 1.2,
      valign: 'middle',
      margin: 0,
    });

    const areas = [
      {
        num: '01',
        title: 'Distinct Brand Positioning',
        desc: 'Position Jontys Hub on Instagram, Facebook & Maps as B.N. Reddy Nagar’s specialist menswear destination.',
      },
      {
        num: '02',
        title: 'Educate Through Video',
        desc: 'Use short Reels & YouTube ads to highlight fabric quality, tailoring, and complete outfit pairings.',
      },
      {
        num: '03',
        title: 'Video-View Remarketing',
        desc: 'Retarget engaged video viewers across Meta & Display to nurture warm, high-intent buying audiences.',
      },
      {
        num: '04',
        title: 'Goal: High-Intent Store Visits',
        desc: 'Guide educated prospects straight to the store with Maps pins, styling checks & fresh-stock hooks.',
      },
      {
        num: '05',
        title: 'WhatsApp Community & Referrals',
        desc: 'Share daily/weekly new-stock drops and margin-safe referral perks to spark repeat walk-ins.',
      },
      {
        num: '06',
        title: 'Profit-Led Measurement',
        desc: 'Track every ad Rupee against store visits, ₹5,000+ bills, and net gross profit contribution.',
      },
    ];

    const cw = (gw - 0.24) / 2;
    const ch = 1.44;
    areas.forEach((a, idx) => {
      const col = idx % 2;
      const row = Math.floor(idx / 2);
      const ax = gx + col * (cw + 0.24);
      const ay = 2.14 + row * (ch + 0.16);

      addCard(s, ax, ay, cw, ch, {
        fill: COLORS.ivoryCard,
        border: COLORS.borderLight,
        accentLeft: true,
        accentColor: COLORS.bronze,
      });

      s.addText(a.num, {
        x: ax + 0.20,
        y: ay + 0.16,
        w: 0.4,
        h: 0.28,
        fontFace: FONTS.heading,
        fontSize: 13,
        bold: true,
        color: COLORS.gold,
        margin: 0,
      });
      s.addText(a.title, {
        x: ax + 0.62,
        y: ay + 0.16,
        w: cw - 0.78,
        h: 0.28,
        fontFace: FONTS.heading,
        fontSize: 14,
        bold: true,
        color: COLORS.navy,
        margin: 0,
      });
      s.addText(a.desc, {
        x: ax + 0.20,
        y: ay + 0.54,
        w: cw - 0.38,
        h: 0.76,
        fontFace: FONTS.body,
        fontSize: 11.5,
        color: COLORS.slate,
        lineSpacingMultiple: 1.25,
        margin: 0,
      });
    });
  }

  // =========================================================================
  // SLIDE 05 – SWOT ANALYSIS (Charcoal Black)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'charcoal',
      kicker: 'Strategic Position',
      title: 'SWOT Analysis: Building on Strengths, Protecting Margins.',
      slideNum: 5,
    });

    const qw = (CONTENT_W - 0.28) / 2; // 5.876
    const qh = 2.60;

    const quadrants = [
      {
        col: 0,
        row: 0,
        tag: 'INTERNAL  •  POSITIVE',
        title: 'STRENGTHS',
        accent: COLORS.gold,
        items: [
          "100% dedicated men's fashion focus — clear specialist brand positioning",
          'Five complementary categories built for complete outfit cross-selling',
          'Established B.N. Reddy Nagar store in a prime LB Nagar corridor',
          'Strong ₹5,000 average bill (~2.5x ₹2,000 avg product price)',
        ],
      },
      {
        col: 1,
        row: 0,
        tag: 'INTERNAL  •  CONSTRAINTS',
        title: 'WEAKNESSES',
        accent: COLORS.bronze,
        items: [
          'Gross margin under 25% forbids unfocused, low-intent ad spend',
          'Digital brand positioning & video storytelling need structured launch',
          'No active video-view remarketing pool capturing warm local viewers',
          'No store-visit attribution linking online impressions to bills',
        ],
      },
      {
        col: 0,
        row: 1,
        tag: 'EXTERNAL  •  UPSIDE',
        title: 'OPPORTUNITIES',
        accent: COLORS.gold,
        items: [
          'Capture 50%+ video-ad viewers as warm local remarketing audiences',
          'Educate buyers on fit & fabric to attract high-intent store visits',
          'WhatsApp community for daily/weekly new-stock alerts & referrals',
          'Outfit styling hooks to expand ₹5,000+ multi-item store baskets',
        ],
      },
      {
        col: 1,
        row: 1,
        tag: 'EXTERNAL  •  RISKS',
        title: 'THREATS',
        accent: COLORS.slate,
        items: [
          'Competing retail chains & nearby menswear outlets',
          'Market-wide discounting habits that compress net margins',
          'Rising CPMs if ads target low-intent cold traffic without retargeting',
          'Price resistance when brand positioning & quality are not explained',
        ],
      },
    ];

    quadrants.forEach((q) => {
      const qx = MARGIN_X + q.col * (qw + 0.28);
      const qy = 1.50 + q.row * (qh + 0.16);

      addCard(s, qx, qy, qw, qh, {
        fill: COLORS.charcoalCard,
        border: COLORS.borderCharcoal,
        accentTop: true,
        accentColor: q.accent,
      });

      s.addText(q.tag, {
        x: qx + 0.24,
        y: qy + 0.14,
        w: 3.5,
        h: 0.22,
        fontFace: FONTS.body,
        fontSize: 9,
        bold: true,
        color: q.accent,
        charSpacing: 1.5,
        margin: 0,
      });

      s.addText(q.title, {
        x: qx + 0.24,
        y: qy + 0.36,
        w: 4.0,
        h: 0.34,
        fontFace: FONTS.heading,
        fontSize: 17.5,
        bold: true,
        color: COLORS.ivory,
        margin: 0,
      });

      const bulletText = q.items.map((txt) => ({
        text: txt,
        options: { bullet: { code: '25AA', color: q.accent }, indentLevel: 0 },
      }));

      s.addText(bulletText, {
        x: qx + 0.24,
        y: qy + 0.76,
        w: qw - 0.48,
        h: 1.72,
        fontFace: FONTS.body,
        fontSize: 10.4,
        color: COLORS.lightText,
        lineSpacingMultiple: 1.18,
        paraSpaceAfter: 3,
        valign: 'top',
        margin: 0,
      });
    });
  }

  // =========================================================================
  // SLIDE 06 – IDEAL CUSTOMER PROFILE (Warm Ivory)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'light',
      kicker: 'Target Audience Architecture',
      title: 'Reaching the Right Customers.',
      slideNum: 6,
    });

    // Validation badge banner
    s.addShape('rect', {
      x: MARGIN_X,
      y: 1.46,
      w: CONTENT_W,
      h: 0.34,
      fill: { color: COLORS.ivorySoft },
      line: { color: COLORS.gold, width: 0.75 },
    });
    s.addText('PROPOSED AUDIENCE SEGMENTS TO VALIDATE  •  ALL PERSONAS REPRESENT HYPOTHESES FOR PHASE-1 CAMPAIGN TESTING', {
      x: MARGIN_X + 0.16,
      y: 1.46,
      w: CONTENT_W - 0.32,
      h: 0.34,
      fontFace: FONTS.body,
      fontSize: 9.5,
      bold: true,
      color: COLORS.navy,
      charSpacing: 1.1,
      valign: 'middle',
      margin: 0,
    });

    const personas = [
      {
        img: 'icp1_young_fashion.jpg',
        title: 'Young Fashion-Conscious Men',
        age: 'AGE 20 – 30',
        motivation: 'Trend-aligned styling, sharp silhouettes & social confidence.',
        interests: 'Casual Shirts, T-Shirts, Slim Jeans',
        comms: 'Short-form Reels, outfit styling hooks & fresh drop alerts.',
      },
      {
        img: 'icp2_working_pro.jpg',
        title: 'Working Professionals',
        age: 'AGE 25 – 40',
        motivation: 'All-day comfort, crisp polish & reliable office-to-evening versatility.',
        interests: 'Formal Shirts, Tailored Trousers',
        comms: 'Fit & fabric quality focus via Meta, Display & YouTube ads.',
      },
      {
        img: 'icp3_aspirational.jpg',
        title: 'Aspirational Fashion Buyers',
        age: 'AGE 25 – 40',
        motivation: 'Upgrading personal wardrobe from basic wear to elevated essentials.',
        interests: 'Layered Casuals, T-Shirts, Trousers',
        comms: 'Complete lookbook pairings & weekly WhatsApp stock alerts.',
      },
      {
        img: 'icp4_occasion.jpg',
        title: 'Occasion-Based Shoppers',
        age: 'AGE 23 – 45',
        motivation: 'Ready-to-wear elevated outfits for celebrations, dinners & events.',
        interests: 'Dark Formal Shirts, Crisp Trousers',
        comms: 'Event outfit videos, WhatsApp previews & referral invitations.',
      },
    ];

    const cw = (CONTENT_W - 0.22 * 3) / 4; // 2.843
    const imgW = cw;
    const imgH = 2.50; // Framed carefully inside top of card — wait: maintain 3:4 or centered framing!
    // Let's keep exact 3:4 aspect ratio to prevent any stretching: w = 1.95", h = 2.61" centered OR full card width with exact 3:4?
    // At w = 2.843, 3:4 height would be 3.80" which leaves only 1.1" for text.
    // Instead, let's place a clean 1.88" x 2.52" portrait (ratio 0.746) on left or centered, OR horizontal split inside each tall card!
    // Even better: horizontal portrait frame? No, 1.88 x 2.52 portrait on top-center with dark/ivory mat, OR portrait left? No, 4 columns at 2.84" wide:
    // If image is w = 1.80", h = 2.41" (exact 0.747 ratio!) inside an architectural charcoal/ivory top band, OR w = 2.84", h = 2.45" using PptxGenJS crop?
    // Wait: let's check if PptxGenJS `sizing: { type: 'crop', ... }` or exact 3:4 `w=1.80, h=2.41` looks cleaner. Exact 3:4 (`w=1.86, h=2.49`) with a side editorial badge looks super editorial AND never distorts! Let's check: 2.84 - 1.86 = 0.98" next to the portrait for the Age Pill + Persona Number (`01`..`04`)! That looks like a bespoke magazine profile card!

    personas.forEach((p, i) => {
      const px = MARGIN_X + i * (cw + 0.22);
      const py = 1.94;
      const ch = 4.90;

      addCard(s, px, py, cw, ch, {
        fill: COLORS.ivoryCard,
        border: COLORS.borderLight,
        accentTop: true,
        accentColor: COLORS.gold,
      });

      // Top zone: 3:4 portrait on left (1.82 x 2.44) + editorial index & age rail on right
      s.addImage({
        path: img(p.img),
        x: px + 0.14,
        y: py + 0.16,
        w: 1.76,
        h: 2.36, // 1.76 / 2.36 = 0.746 exact match!
      });

      s.addText(`0${i + 1}`, {
        x: px + 1.96,
        y: py + 0.18,
        w: 0.78,
        h: 0.44,
        fontFace: FONTS.heading,
        fontSize: 22,
        bold: true,
        color: COLORS.gold,
        align: 'right',
        margin: 0,
      });

      s.addShape('rect', {
        x: px + 1.98,
        y: py + 0.72,
        w: 0.74,
        h: 0.42,
        fill: { color: COLORS.navy },
        line: { type: 'none' },
      });
      s.addText(p.age, {
        x: px + 1.98,
        y: py + 0.72,
        w: 0.74,
        h: 0.42,
        fontFace: FONTS.body,
        fontSize: 8.5,
        bold: true,
        color: COLORS.ivory,
        align: 'center',
        valign: 'middle',
        margin: 0,
      });

      s.addText('SEGMENT\nTO TEST', {
        x: px + 1.98,
        y: py + 1.24,
        w: 0.74,
        h: 0.42,
        fontFace: FONTS.body,
        fontSize: 7.5,
        bold: true,
        color: COLORS.slate,
        align: 'right',
        margin: 0,
      });

      // Persona Name
      s.addText(p.title, {
        x: px + 0.14,
        y: py + 2.62,
        w: cw - 0.28,
        h: 0.48,
        fontFace: FONTS.heading,
        fontSize: 13.5,
        bold: true,
        color: COLORS.navy,
        valign: 'top',
        margin: 0,
      });

      s.addShape('line', {
        x: px + 0.14,
        y: py + 3.12,
        w: cw - 0.28,
        h: 0,
        line: { color: COLORS.borderLight, width: 0.6 },
      });

      // Details
      const fields = [
        { label: 'MOTIVATION', val: p.motivation, y: py + 3.20 },
        { label: 'PRODUCT FOCUS', val: p.interests, y: py + 3.80 },
        { label: 'COMMUNICATION', val: p.comms, y: py + 4.30 },
      ];
      fields.forEach((f) => {
        s.addText(f.label, {
          x: px + 0.14,
          y: f.y,
          w: cw - 0.28,
          h: 0.16,
          fontFace: FONTS.body,
          fontSize: 8,
          bold: true,
          color: COLORS.bronze,
          charSpacing: 0.8,
          margin: 0,
        });
        s.addText(f.val, {
          x: px + 0.14,
          y: f.y + 0.17,
          w: cw - 0.28,
          h: 0.38,
          fontFace: FONTS.body,
          fontSize: 10,
          color: COLORS.slate,
          lineSpacingMultiple: 1.15,
          margin: 0,
        });
      });
    });
  }

  // =========================================================================
  // SLIDE 07 – LOCAL MARKET STRATEGY (Deep Midnight Navy)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'navy',
      kicker: 'Hyper-Local Catchment Targeting',
      title: "Winning the Local Men's Fashion Market.",
      slideNum: 7,
    });

    // Left 7.4" Custom Editable Vector Catchment Map of LB Nagar / B.N. Reddy Nagar
    const mx = MARGIN_X;
    const my = 1.56;
    const mw = 7.35;
    const mh = 5.20;

    addCard(s, mx, my, mw, mh, {
      fill: '0B1220',
      border: COLORS.borderDark,
    });

    s.addText('HYDERABAD SOUTH-EAST CATCHMENT TOPOLOGY  (NO INVENTED DEMOGRAPHIC ESTIMATES)', {
      x: mx + 0.20,
      y: my + 0.14,
      w: mw - 0.40,
      h: 0.22,
      fontFace: FONTS.body,
      fontSize: 8.0,
      bold: true,
      color: COLORS.lightMuted,
      charSpacing: 0.7,
      margin: 0,
    });

    // Store center coordinates inside map box
    const cx = mx + 3.45;
    const cy = my + 3.25;

    // Outer Testing Ring (Kothapet reach)
    s.addShape('ellipse', {
      x: cx - 2.95,
      y: cy - 2.55,
      w: 5.90,
      h: 4.50,
      fill: { color: '131E33', transparency: 35 },
      line: { color: COLORS.slate, width: 1.0, dashType: 'dash' },
    });

    // Secondary Ring (LB Nagar & Chintalkunta)
    s.addShape('ellipse', {
      x: cx - 2.05,
      y: cy - 1.75,
      w: 4.10,
      h: 3.20,
      fill: { color: '192640', transparency: 25 },
      line: { color: COLORS.bronze, width: 1.25 },
    });

    // Primary Ring (B.N. Reddy Nagar, Vanasthalipuram, Hastinapuram)
    s.addShape('ellipse', {
      x: cx - 1.15,
      y: cy - 0.95,
      w: 2.30,
      h: 1.80,
      fill: { color: COLORS.gold, transparency: 82 },
      line: { color: COLORS.gold, width: 1.75 },
    });

    // Stylized Major Arterial Roads (Nagarjuna Sagar Road / NH65 / Inner Ring Rd)
    s.addShape('line', {
      x: mx + 2.85,
      y: my + 0.55,
      w: 0.56,
      h: 3.90,
      line: { color: '334155', width: 2.0 },
    });
    s.addShape('line', {
      x: mx + 0.70,
      y: my + 1.20,
      w: 5.95,
      h: 0.45,
      flipV: true,
      line: { color: '334155', width: 2.0 },
    });

    // Road labels
    s.addText('NH 65 / Vijayawada Hwy', {
      x: mx + 4.85,
      y: my + 0.96,
      w: 1.9,
      h: 0.22,
      fontFace: FONTS.body,
      fontSize: 8.5,
      color: COLORS.lightMuted,
      margin: 0,
    });
    s.addText('Nagarjuna Sagar Rd', {
      x: mx + 1.38,
      y: my + 4.18,
      w: 1.65,
      h: 0.22,
      fontFace: FONTS.body,
      fontSize: 8.5,
      color: COLORS.lightMuted,
      margin: 0,
    });

    // Map Nodes
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
        h: 0.26,
        fontFace: isAnchor ? FONTS.heading : FONTS.body,
        fontSize: isAnchor ? 10 : 9.5,
        bold: true,
        color: isAnchor ? COLORS.gold : COLORS.ivory,
        align: 'center',
        margin: 0,
      });
    };

    // Anchor Store
    mapPin(cx, cy, 'JONTYS HUB\nB.N. Reddy Nagar', COLORS.gold, true);
    // Primary Zone
    mapPin(cx + 0.85, cy - 0.30, 'Vanasthalipuram', COLORS.gold);
    mapPin(cx - 0.78, cy + 0.50, 'Hastinapuram', COLORS.gold);
    // Secondary Zone
    mapPin(cx - 0.10, cy - 1.35, 'LB Nagar', COLORS.bronze);
    mapPin(cx + 1.45, cy - 1.05, 'Chintalkunta', COLORS.bronze);
    // Testing Zone
    mapPin(cx - 1.55, cy - 2.15, 'Kothapet', COLORS.lightMuted);

    // Map Legend Bar inside bottom of map
    s.addShape('rect', {
      x: mx + 0.20,
      y: my + mh - 0.52,
      w: mw - 0.40,
      h: 0.36,
      fill: { color: COLORS.navy },
      line: { color: COLORS.borderDark, width: 0.6 },
    });
    const legendItems = [
      { label: 'Store Anchor', color: COLORS.gold, x: mx + 0.35 },
      { label: 'Primary Zone', color: COLORS.gold, x: mx + 1.95 },
      { label: 'Secondary Zone', color: COLORS.bronze, x: mx + 3.65 },
      { label: 'Testing Zone', color: COLORS.lightMuted, x: mx + 5.45 },
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
        fontFace: FONTS.body,
        fontSize: 9.5,
        color: COLORS.lightText,
        margin: 0,
      });
    });

    // Right column: 3 Tier Cards + Store Anchor Card
    const rx = mx + mw + 0.28;
    const rw = CONTENT_W - mw - 0.28; // 4.403

    const tiers = [
      {
        tier: 'TIER 01  •  PRIMARY ZONE',
        places: 'B.N. Reddy Nagar, Vanasthalipuram, Hastinapuram',
        role: 'Core always-on local reach, Google Maps visibility, and high-frequency store-visit creative.',
        accent: COLORS.gold,
      },
      {
        tier: 'TIER 02  •  SECONDARY ZONE',
        places: 'LB Nagar & Chintalkunta',
        role: 'Category-led Instagram, Facebook, YouTube & Display ads for nearby B2C shoppers.',
        accent: COLORS.bronze,
      },
      {
        tier: 'TIER 03  •  TESTING ZONE',
        places: 'Kothapet',
        role: 'Controlled budget expansion tests to measure travel willingness and incremental store-visit economics.',
        accent: COLORS.slate,
      },
    ];

    tiers.forEach((t, i) => {
      const ty = 1.56 + i * 1.32;
      addCard(s, rx, ty, rw, 1.18, {
        fill: COLORS.navyCard,
        border: COLORS.borderDark,
        accentLeft: true,
        accentColor: t.accent,
      });
      s.addText(t.tier, {
        x: rx + 0.20,
        y: ty + 0.12,
        w: rw - 0.34,
        h: 0.20,
        fontFace: FONTS.body,
        fontSize: 9,
        bold: true,
        color: t.accent,
        charSpacing: 1.2,
        margin: 0,
      });
      s.addText(t.places, {
        x: rx + 0.20,
        y: ty + 0.34,
        w: rw - 0.34,
        h: 0.30,
        fontFace: FONTS.heading,
        fontSize: 13.5,
        bold: true,
        color: COLORS.ivory,
        margin: 0,
      });
      s.addText(t.role, {
        x: rx + 0.20,
        y: ty + 0.66,
        w: rw - 0.34,
        h: 0.44,
        fontFace: FONTS.body,
        fontSize: 11,
        color: COLORS.lightText,
        lineSpacingMultiple: 1.2,
        margin: 0,
      });
    });

    addCard(s, rx, 5.54, rw, 1.22, {
      fill: COLORS.charcoal,
      border: COLORS.gold,
    });
    s.addText('STORE LOCATION ANCHOR', {
      x: rx + 0.20,
      y: 5.66,
      w: rw - 0.40,
      h: 0.22,
      fontFace: FONTS.body,
      fontSize: 9,
      bold: true,
      color: COLORS.gold,
      charSpacing: 1.3,
      margin: 0,
    });
    s.addText('Beside B.D. Reddy Gardens, B.N. Reddy Nagar, LB Nagar, Hyderabad – 500079', {
      x: rx + 0.20,
      y: 5.92,
      w: rw - 0.40,
      h: 0.42,
      fontFace: FONTS.body,
      fontSize: 11.5,
      color: COLORS.ivory,
      margin: 0,
    });
    s.addText('Budget expands outward only after Primary Zone unit economics are proven.', {
      x: rx + 0.20,
      y: 6.38,
      w: rw - 0.40,
      h: 0.26,
      fontFace: FONTS.body,
      fontSize: 10,
      italic: true,
      color: COLORS.lightMuted,
      margin: 0,
    });
  }

  // =========================================================================
  // SLIDE 08 – CUSTOMER ACQUISITION FUNNEL (Warm Ivory)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'light',
      kicker: 'Conversion Architecture',
      title: 'From Fashion Discovery to Store Purchase.',
      slideNum: 8,
    });

    const stages = [
      {
        num: '01',
        name: 'VIDEO DISCOVERY',
        channel: 'Instagram Reels • Facebook Video • YouTube',
        objective: 'Capture attention & pool engaged local video viewers.',
        action: 'Style-led video ads reach LB Nagar men; 50%+ video viewers enter the custom remarketing pool.',
        w: 5.20,
        bg: COLORS.navy,
        txt: COLORS.ivory,
        badge: COLORS.gold,
      },
      {
        num: '02',
        name: 'EDUCATE & POSITION',
        channel: 'Video-View Remarketing • Meta & Display',
        objective: 'Educate warm viewers & establish premium positioning.',
        action: 'Sequential ads teach fabric feel, collar fit & complete styling—building authority before store entry.',
        w: 4.55,
        bg: '1E293B',
        txt: COLORS.ivory,
        badge: COLORS.gold,
      },
      {
        num: '03',
        name: 'HIGH-INTENT FILTER',
        channel: 'WhatsApp Community • Google Business Profile',
        objective: 'Qualify high-intent buyers ready to visit the store.',
        action: 'Maps directions, sizing previews & fresh-stock highlights isolate serious local shoppers.',
        w: 3.90,
        bg: '2A3447',
        txt: COLORS.ivory,
        badge: COLORS.gold,
      },
      {
        num: '04',
        name: 'STORE VISIT GOAL',
        channel: 'In-Store Trial • Outfit Styling',
        objective: 'Convert high-intent walk-ins into ₹5,000+ bills.',
        action: 'Educated buyers arrive pre-sold on quality—staff guide complete multi-item outfit trials.',
        w: 3.25,
        bg: COLORS.charcoal,
        txt: COLORS.ivory,
        badge: COLORS.gold,
      },
      {
        num: '05',
        name: 'RETENTION',
        channel: 'WhatsApp Drops • Referrals',
        objective: 'Turn buyers into repeat members and peer referrers.',
        action: 'Daily/weekly new-stock updates & margin-safe referral discounts drive repeat store visits.',
        w: 2.65,
        bg: COLORS.gold,
        txt: COLORS.navy,
        badge: COLORS.navy,
      },
    ];

    // Left Column: Geometric 5-Stage Funnel Visual + Right Column: Stage Detail Rows
    const funnelCenterX = MARGIN_X + 2.65;
    stages.forEach((st, i) => {
      const sy = 1.60 + i * 1.02;
      const sh = 0.88;

      // Trapezoid / Tapered bar centered on funnelCenterX
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
        x: fx + 0.15,
        y: sy + 0.14,
        w: st.w - 0.30,
        h: 0.34,
        fontFace: FONTS.heading,
        fontSize: 12.5,
        bold: true,
        color: i === 4 ? COLORS.navy : COLORS.gold,
        align: 'center',
        margin: 0,
      });

      s.addText(st.channel, {
        x: fx + 0.15,
        y: sy + 0.48,
        w: st.w - 0.30,
        h: 0.26,
        fontFace: FONTS.body,
        fontSize: 10,
        color: st.txt,
        align: 'center',
        margin: 0,
      });

      // Connector arrow down between stages
      if (i < stages.length - 1) {
        s.addShape('line', {
          x: funnelCenterX,
          y: sy + sh,
          w: 0,
          h: 0.14,
          line: { color: COLORS.gold, width: 1.5 },
        });
      }

      // Horizontal bridge line to Detail Card
      const dx = MARGIN_X + 5.60;
      const dw = CONTENT_W - 5.60; // 6.433
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
        accentColor: i === 4 ? COLORS.navy : COLORS.gold,
      });

      s.addText(st.objective, {
        x: dx + 0.22,
        y: sy + 0.12,
        w: dw - 0.38,
        h: 0.28,
        fontFace: FONTS.heading,
        fontSize: 12.5,
        bold: true,
        color: COLORS.navy,
        margin: 0,
      });

      s.addText(st.action, {
        x: dx + 0.22,
        y: sy + 0.44,
        w: dw - 0.38,
        h: 0.34,
        fontFace: FONTS.body,
        fontSize: 11,
        color: COLORS.slate,
        margin: 0,
      });
    });

    // Bottom Bridge Callout strip inside whitespace? Actually 5 rows take up 1.60 to 6.56 — crisp and well-balanced!
  }

  // =========================================================================
  // SLIDE 09 – VIDEO-VIEW REMARKETING & POSITIONING (Deep Midnight Navy)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'navy',
      kicker: 'Remarketing & Positioning Architecture',
      title: 'Turn Video Viewers Into High-Intent Store Visitors.',
      slideNum: 9,
    });

    // Core Thesis Highlight Banner
    addCard(s, MARGIN_X, 1.50, CONTENT_W, 0.86, {
      fill: COLORS.charcoal,
      border: COLORS.gold,
      accentLeft: true,
      accentColor: COLORS.gold,
    });

    s.addText('CORE REMARKETING STRATEGY', {
      x: MARGIN_X + 0.24,
      y: 1.59,
      w: 4.5,
      h: 0.22,
      fontFace: FONTS.body,
      fontSize: 9,
      bold: true,
      color: COLORS.gold,
      charSpacing: 1.4,
      margin: 0,
    });

    s.addText(
      'Anyone who watches our Reels or YouTube video ads enters a custom Remarketing Audience — educated sequentially to build brand positioning and drive high-intent walk-ins to our LB Nagar store.',
      {
        x: MARGIN_X + 0.24,
        y: 1.84,
        w: CONTENT_W - 0.48,
        h: 0.44,
        fontFace: FONTS.body,
        fontSize: 11.5,
        color: COLORS.ivory,
        lineSpacingMultiple: 1.18,
        margin: 0,
      }
    );

    // 4 Sequential Step Cards with connecting gold chevrons/arrows
    const rSteps = [
      {
        num: '01',
        tag: 'AUDIENCE CAPTURE',
        title: 'Build Video-View\nRemarketing Pool',
        highlight: '50%+ Video Watchers',
        body: 'Launch style-led Instagram Reels, Facebook Video & YouTube Ads locally. Engaged viewers (50%+ watch time) automatically build a warm custom audience.',
        goal: 'Signal: Real Interest',
        fill: COLORS.navyCard,
        accent: COLORS.gold,
      },
      {
        num: '02',
        tag: 'CUSTOMER EDUCATION',
        tagColor: COLORS.gold,
        title: 'Educate on Fit,\nFabric & Quality',
        highlight: 'Show Value Before Price',
        body: 'Retarget warm viewers with close-up fabric weaves, collar structure & 5-category outfit pairings — proving why ₹2,000+ garments hold sharper drape.',
        goal: 'Outcome: Educated Buyer',
        fill: COLORS.navyCard,
        accent: COLORS.gold,
      },
      {
        num: '03',
        tag: 'BRAND POSITIONING',
        title: 'Build Distinct Store\nPositioning',
        highlight: 'Men-Only Style Authority',
        body: 'Consistent Meta & Display creative positions Jontys Hub as B.N. Reddy Nagar’s specialist men’s destination — eliminating bargain-shopper perception.',
        goal: 'Outcome: Brand Preference',
        fill: COLORS.navyCard,
        accent: COLORS.bronze,
      },
      {
        num: '04',
        tag: 'ULTIMATE GOAL',
        title: 'Convert High-Intent\nStore Visits',
        highlight: 'In-Store Trial & ₹5,000+ Bills',
        body: 'Serve high-intent viewers Google Maps directions, fresh-stock previews & WhatsApp sizing checks — turning pre-sold locals into walk-in buyers.',
        goal: 'Goal: Visit the Store',
        fill: COLORS.charcoal,
        accent: COLORS.gold,
        isGoal: true,
      },
    ];

    const gap = 0.30;
    const sw = (CONTENT_W - gap * 3) / 4; // 2.783
    const sy = 2.54;
    const sh = 3.24;

    rSteps.forEach((st, idx) => {
      const sx = MARGIN_X + idx * (sw + gap);

      // Connecting arrow between columns
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
        border: st.isGoal ? COLORS.gold : COLORS.borderDark,
        accentTop: true,
        accentColor: st.accent,
      });

      s.addText(`${st.num}  •  ${st.tag}`, {
        x: sx + 0.18,
        y: sy + 0.16,
        w: sw - 0.36,
        h: 0.22,
        fontFace: FONTS.body,
        fontSize: 8.5,
        bold: true,
        color: st.accent,
        charSpacing: 1.0,
        margin: 0,
      });

      s.addText(st.title, {
        x: sx + 0.18,
        y: sy + 0.44,
        w: sw - 0.36,
        h: 0.56,
        fontFace: FONTS.heading,
        fontSize: 14,
        bold: true,
        color: COLORS.ivory,
        lineSpacingMultiple: 1.15,
        margin: 0,
      });

      // Highlight pill inside card
      s.addShape('rect', {
        x: sx + 0.18,
        y: sy + 1.08,
        w: sw - 0.36,
        h: 0.32,
        fill: { color: st.isGoal ? COLORS.gold : COLORS.navy },
        line: { color: COLORS.gold, width: 0.6 },
      });

      s.addText(st.highlight.toUpperCase(), {
        x: sx + 0.22,
        y: sy + 1.08,
        w: sw - 0.44,
        h: 0.32,
        fontFace: FONTS.body,
        fontSize: 8.5,
        bold: true,
        color: st.isGoal ? COLORS.navy : COLORS.gold,
        align: 'center',
        valign: 'middle',
        charSpacing: 0.6,
        margin: 0,
      });

      s.addText(st.body, {
        x: sx + 0.18,
        y: sy + 1.52,
        w: sw - 0.36,
        h: 1.18,
        fontFace: FONTS.body,
        fontSize: 10.6,
        color: COLORS.lightText,
        lineSpacingMultiple: 1.24,
        margin: 0,
      });

      // Footer tag
      s.addShape('line', {
        x: sx + 0.18,
        y: sy + 2.80,
        w: sw - 0.36,
        h: 0,
        line: { color: COLORS.borderDark, width: 0.6 },
      });

      s.addText(st.goal, {
        x: sx + 0.18,
        y: sy + 2.88,
        w: sw - 0.36,
        h: 0.26,
        fontFace: FONTS.heading,
        fontSize: 10.5,
        bold: true,
        color: st.isGoal ? COLORS.gold : COLORS.lightMuted,
        margin: 0,
      });
    });

    // Bottom Strategic Impact Bar
    addCard(s, MARGIN_X, 5.96, CONTENT_W, 0.84, {
      fill: COLORS.navyCard,
      border: COLORS.borderDark,
      accentLeft: true,
      accentColor: COLORS.bronze,
    });

    s.addText('WHY THIS PROTECTS OUR < 25% MARGIN', {
      x: MARGIN_X + 0.24,
      y: 6.06,
      w: 3.8,
      h: 0.22,
      fontFace: FONTS.body,
      fontSize: 9,
      bold: true,
      color: COLORS.gold,
      charSpacing: 1.2,
      margin: 0,
    });

    s.addText(
      'We avoid wasting store-visit budget on uninterested cold scrollers. Broad video reaches locals cheaply; remarketing budget concentrates purely on educated, high-intent men ready to visit B.N. Reddy Nagar.',
      {
        x: MARGIN_X + 0.24,
        y: 6.32,
        w: CONTENT_W - 0.48,
        h: 0.40,
        fontFace: FONTS.body,
        fontSize: 11,
        color: COLORS.lightText,
        margin: 0,
      }
    );
  }

  // =========================================================================
  // SLIDE 10 – PERFORMANCE MARKETING ECOSYSTEM (Charcoal Black)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'charcoal',
      kicker: 'Channel Architecture',
      title: 'A Connected Digital Growth Strategy.',
      slideNum: 10,
    });

    // Central Hub: JONTYS HUB STOREFRONT & ATTRIBUTION CORE
    const hubW = 3.10;
    const hubH = 2.00;
    const hubX = W / 2 - hubW / 2;
    const hubY = 4.20 - hubH / 2;

    // Channels positioned symmetrically around the hub (3 Left, 3 Right)
    const nodes = [
      // Left 3
      {
        side: 'left',
        row: 0,
        name: 'Instagram Ads',
        role: 'VIDEO DISCOVERY & POOL',
        desc: 'Style reels that capture 50%+ video viewers for warm remarketing & local positioning.',
      },
      {
        side: 'left',
        row: 1,
        name: 'Facebook Ads',
        role: 'LOCAL REMARKETING',
        desc: 'Educate warm video viewers across LB Nagar on fabric, fit & complete outfit pairings.',
      },
      {
        side: 'left',
        row: 2,
        name: 'YouTube Ads',
        role: 'EDUCATION & STORYTELLING',
        desc: 'Shorts & in-stream videos building brand authority and feeding high-intent remarketing.',
      },
      // Right 3
      {
        side: 'right',
        row: 0,
        name: 'Google Display Ads',
        role: 'SEQUENTIAL RECALL',
        desc: 'Keep Jontys Hub top-of-mind with educated local viewers until they visit the store.',
      },
      {
        side: 'right',
        row: 1,
        name: 'Google Business Profile',
        role: 'GOAL: STORE VISIT',
        desc: 'Maps pins, photos & reviews converting high-intent prospects into store walk-ins.',
      },
      {
        side: 'right',
        row: 2,
        name: 'WhatsApp Community',
        role: 'STOCK DROPS & REFERRALS',
        desc: 'Daily/weekly new stock alerts & margin-safe referral perks driving repeat bills.',
      },
    ];

    const nw = 3.65;
    const nh = 1.48;
    const leftX = MARGIN_X;
    const rightX = W - MARGIN_X - nw;

    // Draw connector lines FIRST so they sit cleanly behind cards
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
        line: {
          color: COLORS.gold,
          width: 1.0,
          dashType: 'dash',
        },
      });
    });

    // Draw Central Hub
    s.addShape('roundRect', {
      x: hubX,
      y: hubY,
      w: hubW,
      h: hubH,
      rectRadius: 0.12,
      fill: { color: COLORS.navy },
      line: { color: COLORS.gold, width: 2.0 },
    });

    s.addText('CENTRAL GROWTH ENGINE', {
      x: hubX + 0.20,
      y: hubY + 0.24,
      w: hubW - 0.40,
      h: 0.24,
      fontFace: FONTS.body,
      fontSize: 9,
      bold: true,
      color: COLORS.gold,
      align: 'center',
      charSpacing: 1.5,
      margin: 0,
    });

    s.addText('JONTYS HUB\nB2C COMMUNITY', {
      x: hubX + 0.20,
      y: hubY + 0.56,
      w: hubW - 0.40,
      h: 0.72,
      fontFace: FONTS.heading,
      fontSize: 19,
      bold: true,
      color: COLORS.ivory,
      align: 'center',
      valign: 'middle',
      margin: 0,
    });

    s.addText('Store Visits  •  Stock Drops  •  Referrals', {
      x: hubX + 0.15,
      y: hubY + 1.42,
      w: hubW - 0.30,
      h: 0.34,
      fontFace: FONTS.body,
      fontSize: 9.5,
      color: COLORS.lightText,
      align: 'center',
      margin: 0,
    });

    // Draw 6 Channel Cards
    nodes.forEach((n) => {
      const nx = n.side === 'left' ? leftX : rightX;
      const ny = 1.62 + n.row * 1.76;

      addCard(s, nx, ny, nw, nh, {
        fill: COLORS.charcoalCard,
        border: COLORS.borderCharcoal,
        accentTop: true,
        accentColor: COLORS.gold,
      });

      s.addText(n.role, {
        x: nx + 0.20,
        y: ny + 0.14,
        w: nw - 0.40,
        h: 0.20,
        fontFace: FONTS.body,
        fontSize: 9,
        bold: true,
        color: COLORS.gold,
        charSpacing: 1.3,
        margin: 0,
      });

      s.addText(n.name, {
        x: nx + 0.20,
        y: ny + 0.36,
        w: nw - 0.40,
        h: 0.34,
        fontFace: FONTS.heading,
        fontSize: 16,
        bold: true,
        color: COLORS.ivory,
        margin: 0,
      });

      s.addText(n.desc, {
        x: nx + 0.20,
        y: ny + 0.74,
        w: nw - 0.40,
        h: 0.62,
        fontFace: FONTS.body,
        fontSize: 11,
        color: COLORS.lightText,
        lineSpacingMultiple: 1.22,
        margin: 0,
      });
    });

    // Top-center & Bottom-center ecosystem notes
    addCard(s, hubX, 1.62, hubW, 1.30, {
      fill: COLORS.charcoalCard,
      border: COLORS.borderCharcoal,
    });
    s.addText('OMNICHANNEL LOOP', {
      x: hubX + 0.20,
      y: 1.76,
      w: hubW - 0.40,
      h: 0.22,
      fontFace: FONTS.body,
      fontSize: 9,
      bold: true,
      color: COLORS.gold,
      align: 'center',
      charSpacing: 1.3,
      margin: 0,
    });
    s.addText('Every Meta, YouTube & Display impression feeds store visits or the WhatsApp community.', {
      x: hubX + 0.20,
      y: 2.04,
      w: hubW - 0.40,
      h: 0.72,
      fontFace: FONTS.body,
      fontSize: 11.5,
      color: COLORS.lightText,
      align: 'center',
      lineSpacingMultiple: 1.25,
      margin: 0,
    });

    addCard(s, hubX, 5.44, hubW, 1.18, {
      fill: COLORS.charcoalCard,
      border: COLORS.borderCharcoal,
    });
    s.addText('MARGIN DISCIPLINE', {
      x: hubX + 0.20,
      y: 5.56,
      w: hubW - 0.40,
      h: 0.22,
      fontFace: FONTS.body,
      fontSize: 9,
      bold: true,
      color: COLORS.bronze,
      align: 'center',
      charSpacing: 1.3,
      margin: 0,
    });
    s.addText('Paid ads build our WhatsApp community, lowering future CAC via stock drops & referrals.', {
      x: hubX + 0.20,
      y: 5.84,
      w: hubW - 0.40,
      h: 0.66,
      fontFace: FONTS.body,
      fontSize: 11,
      color: COLORS.lightText,
      align: 'center',
      lineSpacingMultiple: 1.22,
      margin: 0,
    });
  }

  // =========================================================================
  // SLIDE 11 – PREMIUM CONTENT STRATEGY (Warm Ivory)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'light',
      kicker: 'Creative & Merchandising Direction',
      title: 'Make Every Product Worth Discovering.',
      slideNum: 11,
    });

    const products = [
      {
        img: 'cat1_casual_shirts.jpg',
        cat: 'CASUAL SHIRTS',
        tagline: 'Upgrade Your Everyday Style.',
        concept: 'Weekend transition reels pairing textured cottons with rolled cuffs & effortless fits.',
      },
      {
        img: 'cat2_formal_shirts.jpg',
        cat: 'FORMAL SHIRTS',
        tagline: 'Premium Style. Everyday Confidence.',
        concept: 'Sharp collar, cuff & fabric close-ups built for office leaders and formal occasions.',
      },
      {
        img: 'cat3_tshirts.jpg',
        cat: 'T-SHIRTS',
        tagline: 'Dress Better. Feel Better.',
        concept: 'Minimalist studio lookbooks highlighting heavyweight drape, neckline hold & clean tones.',
      },
      {
        img: 'cat4_jeans.jpg',
        cat: 'JEANS',
        tagline: 'Find Your Perfect Fit.',
        concept: '360-degree fit-check videos demonstrating tapered leg cuts, wash depth & all-day comfort.',
      },
      {
        img: 'cat5_trousers.jpg',
        cat: 'TROUSERS',
        tagline: 'Redefine Your Wardrobe.',
        concept: 'Full-outfit styling showing how tailored trousers elevate both casual & formal shirts.',
      },
    ];

    const cw = (CONTENT_W - 0.20 * 4) / 5; // 2.246
    const imgW = cw;
    const imgH = +(cw / 0.747).toFixed(3); // 3.007" exact 3:4 aspect ratio!

    products.forEach((p, i) => {
      const px = MARGIN_X + i * (cw + 0.20);
      const py = 1.54;
      const ch = 5.26;

      addCard(s, px, py, cw, ch, {
        fill: COLORS.ivoryCard,
        border: COLORS.borderLight,
      });

      // Full-width 3:4 Editorial Photo
      s.addImage({
        path: img(p.img),
        x: px,
        y: py,
        w: imgW,
        h: imgH,
      });

      // Gold accent bar below photo
      s.addShape('rect', {
        x: px,
        y: py + imgH,
        w: cw,
        h: 0.045,
        fill: { color: COLORS.gold },
        line: { type: 'none' },
      });

      s.addText(p.cat, {
        x: px + 0.14,
        y: py + imgH + 0.14,
        w: cw - 0.28,
        h: 0.22,
        fontFace: FONTS.body,
        fontSize: 9,
        bold: true,
        color: COLORS.bronze,
        charSpacing: 1.2,
        margin: 0,
      });

      s.addText(`"${p.tagline}"`, {
        x: px + 0.12,
        y: py + imgH + 0.40,
        w: cw - 0.24,
        h: 0.52,
        fontFace: FONTS.heading,
        fontSize: 11.5,
        bold: true,
        color: COLORS.navy,
        lineSpacingMultiple: 1.18,
        margin: 0,
      });

      s.addText(p.concept, {
        x: px + 0.12,
        y: py + imgH + 0.98,
        w: cw - 0.24,
        h: 1.08,
        fontFace: FONTS.body,
        fontSize: 10.2,
        color: COLORS.slate,
        lineSpacingMultiple: 1.25,
        margin: 0,
      });
    });
  }

  // =========================================================================
  // SLIDE 12 – REVENUE AND PROFITABILITY (Deep Midnight Navy)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'navy',
      kicker: 'Commercial & Unit-Economics Discipline',
      title: 'Grow Sales Without Compromising Margins.',
      slideNum: 12,
    });

    // Top Left 3 Verified Baseline Cards
    const kpiW = 1.92;
    const baselines = [
      { label: 'AVG PRODUCT PRICE', val: '₹2,000' },
      { label: 'AVG CUSTOMER BILL', val: '₹5,000' },
      { label: 'GROSS MARGIN', val: '< 25%' },
    ];
    baselines.forEach((b, i) => {
      const bx = MARGIN_X + i * (kpiW + 0.14);
      addCard(s, bx, 1.56, kpiW, 1.10, {
        fill: COLORS.navyCard,
        border: COLORS.gold,
        accentTop: true,
      });
      s.addText(b.label, {
        x: bx + 0.12,
        y: 1.70,
        w: kpiW - 0.24,
        h: 0.24,
        fontFace: FONTS.body,
        fontSize: 8.5,
        bold: true,
        color: COLORS.lightMuted,
        charSpacing: 0.9,
        margin: 0,
      });
      s.addText(b.val, {
        x: bx + 0.12,
        y: 1.98,
        w: kpiW - 0.24,
        h: 0.52,
        fontFace: FONTS.heading,
        fontSize: 22,
        bold: true,
        color: COLORS.gold,
        margin: 0,
      });
    });

    // Left Bottom Native Editable Chart: Illustrative Basket Value Comparison
    addCard(s, MARGIN_X, 2.82, 6.04, 3.96, {
      fill: COLORS.navyCard,
      border: COLORS.borderDark,
    });

    s.addText('ILLUSTRATIVE BASKET SCENARIOS (₹)  —  FOR PLANNING ONLY', {
      x: MARGIN_X + 0.20,
      y: 2.96,
      w: 5.64,
      h: 0.24,
      fontFace: FONTS.body,
      fontSize: 9.5,
      bold: true,
      color: COLORS.gold,
      charSpacing: 1.1,
      margin: 0,
    });

    s.addText('Complete outfit styling multiplies gross profit headroom per ad click compared to single-item sales.', {
      x: MARGIN_X + 0.20,
      y: 3.22,
      w: 5.64,
      h: 0.38,
      fontFace: FONTS.body,
      fontSize: 10.5,
      color: COLORS.lightText,
      margin: 0,
    });

    s.addChart(
      pptx.ChartType.bar,
      [
        {
          name: 'Order Value (₹)',
          labels: ['Single Item\n(Baseline ₹2,000)', 'Current Avg Bill\n(2.5 Items)', 'Outfit Pairing*\n(Illustrative)'],
          values: [2000, 5000, 7000],
        },
      ],
      {
        x: MARGIN_X + 0.22,
        y: 3.66,
        w: 5.60,
        h: 2.66,
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
        valGridLine: { color: COLORS.borderDark, size: 0.5 },
        dataLabelColor: COLORS.ivory,
        dataLabelFontSize: 10.5,
        dataLabelFontBold: true,
        showValue: true,
        dataLabelFormatCode: '₹#,##0',
        plotArea: { fill: { color: COLORS.navyCard } },
      }
    );

    s.addText('*Illustrative 3-item outfit scenario (e.g. Shirt + Trouser/Jeans + T-Shirt). Not a revenue forecast.', {
      x: MARGIN_X + 0.20,
      y: 6.42,
      w: 5.64,
      h: 0.26,
      fontFace: FONTS.body,
      fontSize: 8.5,
      italic: true,
      color: COLORS.lightMuted,
      margin: 0,
    });

    // Right Column: 5 Profitability Principles
    const rx = MARGIN_X + 6.30;
    const rw = CONTENT_W - 6.30; // 5.733

    const principles = [
      {
        title: 'Revenue Is Not Equal to Profit',
        desc: 'With < 25% gross margin, scaling clicks without high buyer intent erodes net store cash.',
      },
      {
        title: 'Remarket Educated Video Viewers',
        desc: 'Focus conversion spend on warm 50%+ video viewers who already understand our quality & fit.',
      },
      {
        title: 'Complete Outfit Recommendations',
        desc: 'Pair Shirts + Jeans/Trousers so each store visit converts at ₹5,000+ rather than ₹2,000.',
      },
      {
        title: 'Structured Referrals Over Blanket Discounts',
        desc: 'Protect positioning—use referral perks only on verified repeat/referred bills, never storewide cuts.',
      },
      {
        title: 'WhatsApp Community & Repeat Growth',
        desc: 'Daily/weekly stock drops bring high-intent community members back at near-zero CAC.',
      },
    ];

    principles.forEach((pr, i) => {
      const py = 1.56 + i * 1.06;
      addCard(s, rx, py, rw, 0.94, {
        fill: COLORS.charcoal,
        border: COLORS.borderDark,
        accentLeft: true,
        accentColor: i === 0 || i === 3 ? COLORS.bronze : COLORS.gold,
      });
      s.addText(pr.title, {
        x: rx + 0.22,
        y: py + 0.12,
        w: rw - 0.36,
        h: 0.26,
        fontFace: FONTS.heading,
        fontSize: 13.5,
        bold: true,
        color: COLORS.gold,
        margin: 0,
      });
      s.addText(pr.desc, {
        x: rx + 0.22,
        y: py + 0.42,
        w: rw - 0.36,
        h: 0.44,
        fontFace: FONTS.body,
        fontSize: 11,
        color: COLORS.lightText,
        lineSpacingMultiple: 1.2,
        margin: 0,
      });
    });
  }

  // =========================================================================
  // SLIDE 13 – 90-DAY EXECUTION ROADMAP (Warm Ivory)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'light',
      kicker: 'Implementation Plan',
      title: 'From Strategy to Measurable Growth.',
      slideNum: 13,
    });

    // Horizontal timeline spine
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
        title: 'Foundation & Positioning',
        themeBg: COLORS.navy,
        themeTxt: COLORS.ivory,
        activities: [
          'Audit Instagram, Facebook & Google Business Profile',
          'Set custom Video-View Remarketing audiences & Meta Pixel',
          'Define <25% margin guardrails & referral reward caps',
          'Film fabric, fit & styling videos across 5 categories',
        ],
        milestone: 'Positioning, tracking & video assets ready',
      },
      {
        days: 'DAYS 16 – 30',
        title: 'Video Capture & Education',
        themeBg: COLORS.ivoryCard,
        themeTxt: COLORS.navy,
        activities: [
          'Launch local Instagram Reels, Facebook & YouTube video ads',
          'Pool 50%+ video viewers into warm remarketing lists',
          'Run fabric & fit education ads via Meta & Display',
          'Start weekly WhatsApp new-stock update broadcasts',
        ],
        milestone: 'Warm video-view remarketing audience built',
      },
      {
        days: 'DAYS 31 – 60',
        title: 'High-Intent Store Visits',
        themeBg: COLORS.ivoryCard,
        themeTxt: COLORS.navy,
        activities: [
          'Retarget educated viewers with Maps & Store-Visit CTAs',
          'Filter high-intent buyers across 4 male shopper personas',
          'Onboard store walk-ins into the WhatsApp Community',
          'Verify in-store bills from remarketing & stock alerts',
        ],
        milestone: 'High-intent walk-ins & validated visit cost',
      },
      {
        days: 'DAYS 61 – 90',
        title: 'Scale & Referral Loops',
        themeBg: COLORS.charcoal,
        themeTxt: COLORS.ivory,
        activities: [
          'Shift spend to top store-visit & gross-profit creatives',
          'Expand winning remarketing into LB Nagar & Kothapet',
          'Activate WhatsApp customer referral perks & daily drops',
          'Deliver 90-day gross profit vs. ad spend review',
        ],
        milestone: 'Repeatable, community-led B2C growth',
      },
    ];

    const pw = (CONTENT_W - 0.22 * 3) / 4; // 2.843
    phases.forEach((ph, i) => {
      const px = MARGIN_X + i * (pw + 0.22);

      // Timeline node circle
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
        fontFace: FONTS.heading,
        fontSize: 9.5,
        bold: true,
        color: COLORS.navy,
        align: 'center',
        valign: 'middle',
        margin: 0,
      });

      // Phase Card
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
        y: py + 0.14,
        w: pw - 0.36,
        h: 0.22,
        fontFace: FONTS.body,
        fontSize: 9,
        bold: true,
        color: COLORS.gold,
        charSpacing: 1.2,
        margin: 0,
      });

      s.addText(ph.title, {
        x: px + 0.18,
        y: py + 0.38,
        w: pw - 0.36,
        h: 0.44,
        fontFace: FONTS.heading,
        fontSize: 14,
        bold: true,
        color: ph.themeTxt,
        valign: 'middle',
        margin: 0,
      });

      s.addShape('line', {
        x: px + 0.18,
        y: py + 0.86,
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
        y: py + 0.96,
        w: pw - 0.36,
        h: 2.52,
        fontFace: FONTS.body,
        fontSize: 9.8,
        color: ph.themeBg === COLORS.ivoryCard ? COLORS.slate : COLORS.lightText,
        lineSpacingMultiple: 1.16,
        paraSpaceAfter: 4,
        valign: 'top',
        margin: 0,
      });

      // Milestone footer strip inside card
      s.addShape('rect', {
        x: px + 0.16,
        y: py + 3.56,
        w: pw - 0.32,
        h: 0.68,
        fill: { color: ph.themeBg === COLORS.ivoryCard ? COLORS.ivorySoft : '1F2937' },
        line: { color: COLORS.gold, width: 0.6 },
      });

      s.addText('EXPECTED MILESTONE', {
        x: px + 0.26,
        y: py + 3.63,
        w: pw - 0.52,
        h: 0.18,
        fontFace: FONTS.body,
        fontSize: 7.5,
        bold: true,
        color: COLORS.bronze,
        charSpacing: 1.0,
        margin: 0,
      });

      s.addText(ph.milestone, {
        x: px + 0.26,
        y: py + 3.82,
        w: pw - 0.52,
        h: 0.36,
        fontFace: FONTS.heading,
        fontSize: 10,
        bold: true,
        color: ph.themeBg === COLORS.ivoryCard ? COLORS.navy : COLORS.ivory,
        margin: 0,
      });
    });
  }

  // =========================================================================
  // SLIDE 14 – PERFORMANCE MEASUREMENT (Charcoal Black)
  // =========================================================================
  {
    const s = pptx.addSlide();
    addSlideChrome(s, {
      theme: 'charcoal',
      kicker: 'Executive BI Dashboard Framework',
      title: 'Measuring What Actually Matters.',
      slideNum: 14,
    });

    // North Star Highlight Banner
    addCard(s, MARGIN_X, 1.54, CONTENT_W, 1.04, {
      fill: COLORS.navy,
      border: COLORS.gold,
      accentLeft: true,
      accentColor: COLORS.gold,
    });

    s.addText('NORTH STAR PROFITABILITY METRIC  •  TRACKED WEEKLY', {
      x: MARGIN_X + 0.26,
      y: 1.66,
      w: 6.5,
      h: 0.24,
      fontFace: FONTS.body,
      fontSize: 9.5,
      bold: true,
      color: COLORS.gold,
      charSpacing: 1.5,
      margin: 0,
    });

    s.addText('FRAMEWORK ONLY — NO INVENTED CAMPAIGN RESULTS', {
      x: W - MARGIN_X - 4.4,
      y: 1.66,
      w: 4.15,
      h: 0.24,
      fontFace: FONTS.body,
      fontSize: 8.5,
      bold: true,
      color: COLORS.gold,
      align: 'right',
      valign: 'middle',
      margin: 0,
    });

    s.addText('Marketing-Generated Gross Profit Versus Advertising Spend', {
      x: MARGIN_X + 0.26,
      y: 1.98,
      w: CONTENT_W - 0.52,
      h: 0.44,
      fontFace: FONTS.heading,
      fontSize: 19.5,
      bold: true,
      color: COLORS.ivory,
      margin: 0,
    });

    // 3 Columns of Dashboard Metrics (Covering all 10 required KPIs cleanly)
    const colW = (CONTENT_W - 0.26 * 2) / 3; // 3.837
    const groups = [
      {
        stage: 'STAGE 01  •  VIDEO & REMARKETING POOL',
        color: COLORS.slate,
        metrics: [
          { name: 'Local Video Reach', desc: 'Unique men reached in LB Nagar radius.' },
          { name: 'Video Viewers (50%+)', desc: 'Warm audience captured for sequential education.' },
          { name: 'High-Intent Enquiries', desc: 'WhatsApp & DM sizing/stock questions.' },
        ],
      },
      {
        stage: 'STAGE 02  •  GOAL: HIGH-INTENT WALK-INS',
        color: COLORS.bronze,
        metrics: [
          { name: 'Store Directions', desc: 'Google Maps & remarketing route taps to store.' },
          { name: 'Educated Store Visits', desc: 'Walk-in shoppers pre-sold via video positioning.' },
          { name: 'Verified Purchases', desc: 'In-store bills attributed to remarketing & drops.' },
        ],
      },
      {
        stage: 'STAGE 03  •  COMMERCIAL PROFITABILITY',
        color: COLORS.gold,
        metrics: [
          { name: 'Average Bill Value', desc: 'Tracked against ₹5,000 store benchmark.' },
          { name: 'Marketing-Attributed Revenue', desc: 'Total sales verified from tracked visitors.' },
          { name: 'Gross Profit Contribution', desc: 'Net gross profit earned after product cost.' },
          { name: 'Repeat & Referral Purchases', desc: 'Repeat visits & referred bills from WhatsApp community.' },
        ],
      },
    ];

    groups.forEach((g, colIdx) => {
      const gx = MARGIN_X + colIdx * (colW + 0.26);
      const gy = 2.80;
      const gh = 3.98;

      addCard(s, gx, gy, colW, gh, {
        fill: COLORS.charcoalCard,
        border: COLORS.borderCharcoal,
        accentTop: true,
        accentColor: g.color,
      });

      s.addText(g.stage, {
        x: gx + 0.20,
        y: gy + 0.18,
        w: colW - 0.40,
        h: 0.26,
        fontFace: FONTS.body,
        fontSize: 8.8,
        bold: true,
        color: g.color === COLORS.slate ? COLORS.lightMuted : g.color,
        charSpacing: 0.9,
        margin: 0,
      });

      const slotH = g.metrics.length === 4 ? 0.82 : 1.06;
      g.metrics.forEach((m, mIdx) => {
        const my = gy + 0.56 + mIdx * slotH;
        s.addShape('rect', {
          x: gx + 0.20,
          y: my,
          w: colW - 0.40,
          h: slotH - 0.12,
          fill: { color: COLORS.charcoal },
          line: { color: COLORS.borderCharcoal, width: 0.6 },
        });
        s.addShape('rect', {
          x: gx + 0.20,
          y: my,
          w: 0.05,
          h: slotH - 0.12,
          fill: { color: g.color },
          line: { type: 'none' },
        });
        s.addText(m.name, {
          x: gx + 0.36,
          y: my + 0.08,
          w: colW - 0.66,
          h: 0.26,
          fontFace: FONTS.heading,
          fontSize: 13,
          bold: true,
          color: COLORS.ivory,
          margin: 0,
        });
        s.addText(m.desc, {
          x: gx + 0.36,
          y: my + 0.36,
          w: colW - 0.66,
          h: slotH - 0.52,
          fontFace: FONTS.body,
          fontSize: 10,
          color: COLORS.lightMuted,
          margin: 0,
        });
      });
    });
  }

  // =========================================================================
  // SLIDE 15 – CLOSING (Deep Midnight Navy)
  // =========================================================================
  {
    const s = pptx.addSlide();
    s.background = { color: COLORS.navy };

    // Right striking full-height editorial portrait (exact 3:4 -> 5.60 x 7.50)
    s.addImage({
      path: img('slide14_closing.jpg'),
      x: 7.733,
      y: 0,
      w: 5.60,
      h: 7.5,
    });

    s.addShape('line', {
      x: 7.733,
      y: 0.65,
      w: 0,
      h: 6.2,
      line: { color: COLORS.gold, width: 1.0 },
    });

    // Left content column
    s.addShape('rect', {
      x: 0.75,
      y: 0.72,
      w: 0.45,
      h: 0.04,
      fill: { color: COLORS.gold },
      line: { type: 'none' },
    });

    s.addText('PARTNERSHIP ROADMAP', {
      x: 1.32,
      y: 0.61,
      w: 4.5,
      h: 0.26,
      fontFace: FONTS.body,
      fontSize: 10,
      bold: true,
      color: COLORS.gold,
      charSpacing: 2.0,
      margin: 0,
    });

    s.addText("LET'S BUILD THE NEXT GROWTH\nCHAPTER OF JONTYS HUB.", {
      x: 0.75,
      y: 1.06,
      w: 6.55,
      h: 1.25,
      fontFace: FONTS.heading,
      fontSize: 24.5,
      bold: true,
      color: COLORS.ivory,
      lineSpacingMultiple: 1.20,
      margin: 0,
    });

    // 3 Partnership Stages: UNDERSTAND, BUILD, GROW
    const stages = [
      {
        label: 'UNDERSTAND',
        desc: 'Understand the business, margin structure and local male shoppers.',
      },
      {
        label: 'BUILD',
        desc: 'Develop the digital marketing, visual content and tracking foundation.',
      },
      {
        label: 'GROW',
        desc: 'Optimise profitable customer acquisition, store visits and repeat retention.',
      },
    ];

    stages.forEach((st, i) => {
      const sy = 2.74 + i * 0.96;
      addCard(s, 0.75, sy, 6.40, 0.82, {
        fill: COLORS.navyCard,
        border: COLORS.borderDark,
        accentLeft: true,
        accentColor: COLORS.gold,
      });

      s.addText(`0${i + 1}`, {
        x: 0.96,
        y: sy + 0.18,
        w: 0.42,
        h: 0.44,
        fontFace: FONTS.heading,
        fontSize: 16,
        bold: true,
        color: COLORS.gold,
        valign: 'middle',
        margin: 0,
      });

      s.addText(st.label, {
        x: 1.46,
        y: sy + 0.10,
        w: 5.45,
        h: 0.26,
        fontFace: FONTS.heading,
        fontSize: 13.5,
        bold: true,
        color: COLORS.ivory,
        charSpacing: 1.0,
        margin: 0,
      });

      s.addText(st.desc, {
        x: 1.46,
        y: sy + 0.38,
        w: 5.45,
        h: 0.34,
        fontFace: FONTS.body,
        fontSize: 11.5,
        color: COLORS.lightText,
        margin: 0,
      });
    });

    // Closing statement box
    s.addShape('line', {
      x: 0.75,
      y: 5.74,
      w: 6.40,
      h: 0,
      line: { color: COLORS.gold, width: 0.8 },
    });

    s.addText(
      "Our objective is to help Jontys Hub reach the right customers, attract more shoppers and build a stronger premium men's fashion business through measurable performance marketing.",
      {
        x: 0.75,
        y: 5.88,
        w: 6.40,
        h: 0.90,
        fontFace: FONTS.body,
        fontSize: 12.5,
        italic: true,
        color: COLORS.ivory,
        lineSpacingMultiple: 1.28,
        margin: 0,
      }
    );

    s.addText('Jontys Hub  •  B.N. Reddy Nagar, LB Nagar, Hyderabad – 500079', {
      x: 0.75,
      y: 7.00,
      w: 5.2,
      h: 0.26,
      fontFace: FONTS.body,
      fontSize: 9.5,
      color: COLORS.lightMuted,
      margin: 0,
    });

    s.addText('15  /  15', {
      x: 6.15,
      y: 7.00,
      w: 1.0,
      h: 0.26,
      fontFace: FONTS.heading,
      fontSize: 10,
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
