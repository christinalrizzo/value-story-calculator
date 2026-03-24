/**
 * Pelago Value Story Generator — Google Apps Script Backend
 *
 * Serves the calculator web app, runs calculations, duplicates the
 * template slide, and populates every placeholder via the Slides API.
 */

// ── Configuration ──────────────────────────────────────────────────
const CONFIG = {
  // Template presentation (the master copy with X placeholders)
  TEMPLATE_ID: '1eHYLMaM1WnUHWc-AP1v2_QPpEOQ4QHw9',

  // Partner config spreadsheet
  SHEET_ID: '1YnoKU-WW7J8dfbID-ee8_lG6g2h-Wdbk3xmJsOWewJk',

  // Standard constants
  SUD_PREVALENCE: 0.20,
  ENGAGEMENT_RATE: 0.10,
  SUPPORT_PCT: 0.56,
  MANAGE_PCT: 0.37,
  TREAT_PCT: 0.06,
  AVG_SAVINGS: 11289,

  // Pricing models
  PRICING: {
    Hawaii:   { support: 995,  manage: 2995, treat: 3995 },
    Zanzibar: { support: 495,  manage: 2995, treat: 3495 },
  },

  // Substance tag text on the template slide (to selectively delete)
  SUBSTANCE_TAGS: {
    tud: '8% of ELs TUD risk',
    aud: '10% of ELs AUD risk',
    cud: '6% of ELs CUD risk',
    oud: '2% of ELs OUD risk',
  },
};


// ── Web App Entry Points ───────────────────────────────────────────

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Pelago Value Story Generator')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ── Client-callable Functions ──────────────────────────────────────

/**
 * Returns the list of partners from the config sheet.
 */
function getPartners() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const partners = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && String(data[i][2]).toLowerCase() === 'active') {
      partners.push({
        name: data[i][0],
        pricingModel: data[i][1],
        notes: data[i][3] || '',
      });
    }
  }
  return partners;
}


/**
 * Runs the calculation. Called from the front end to preview numbers.
 */
function calculate(eligibleLives, pricingModel) {
  const prices = CONFIG.PRICING[pricingModel];
  if (!prices) throw new Error('Unknown pricing model: ' + pricingModel);

  const atRisk = Math.round(eligibleLives * CONFIG.SUD_PREVALENCE);
  const members = Math.round(atRisk * CONFIG.ENGAGEMENT_RATE);

  const supM = Math.round(members * CONFIG.SUPPORT_PCT);
  const manM = Math.round(members * CONFIG.MANAGE_PCT);
  const trtM = Math.round(members * CONFIG.TREAT_PCT);

  const supC = supM * prices.support;
  const manC = manM * prices.manage;
  const trtC = trtM * prices.treat;

  const spend = supC + manC + trtC;
  const savings = members * CONFIG.AVG_SAVINGS;
  const roi = spend > 0 ? Math.round((savings / spend) * 10) / 10 : 0;

  return {
    eligibleLives, pricingModel,
    atRisk, members,
    supM, manM, trtM,
    supC, manC, trtC,
    spend, savings, roi,
  };
}


/**
 * The main action: duplicate template, populate values, return the new URL.
 *
 * @param {Object} params - { prospectName, partnerName, eligibleLives, pricingModel, substances }
 * @returns {Object} - { url, title }
 */
function generateSlide(params) {
  const { prospectName, partnerName, eligibleLives, pricingModel, substances } = params;

  // 1. Run calculations
  const c = calculate(eligibleLives, pricingModel);

  // 2. Duplicate the template
  const templateFile = DriveApp.getFileById(CONFIG.TEMPLATE_ID);
  const slideTitle = partnerName + ' — ' + prospectName + ' Value Story';
  const copy = templateFile.makeCopy(slideTitle);
  const newId = copy.getId();

  // 3. Open the new presentation
  const pres = SlidesApp.openById(newId);
  const slide = pres.getSlides()[0];

  // 4. Build replacement map
  const replacements = buildReplacementMap(c, prospectName);

  // 5. Replace all text placeholders
  for (const [find, replace] of Object.entries(replacements)) {
    slide.replaceAllText(find, replace);
  }

  // 6. Remove unselected substance tags
  removeUnselectedSubstances(slide, substances);

  // 7. Save and return
  pres.saveAndClose();

  return {
    url: 'https://docs.google.com/presentation/d/' + newId + '/edit',
    title: slideTitle,
    fileId: newId,
  };
}


// ── Helpers ────────────────────────────────────────────────────────

function buildReplacementMap(c, prospectName) {
  const fmt = (n) => n.toLocaleString('en-US');
  const fmtD = (n) => '$' + n.toLocaleString('en-US');
  const fmtM = (n) => '$' + (n / 1e6).toFixed(1) + 'M';

  // The template has these placeholder patterns.
  // We replace the most specific / longest patterns first by using
  // replaceAllText which handles them independently.
  return {
    // Title bar — full placeholder title
    // We replace key fragments that appear in the title
    'XXX Members Engaged': fmt(c.members) + ' Members Engaged',
    '$X.XM Saved': fmtM(c.savings) + ' Saved',
    'X.X : 1': c.roi.toFixed(1) + ' : 1',

    // Top-left: eligible lives
    'XXXXX': fmt(c.eligibleLives),

    // Top-left: prospect name
    'PWGA': prospectName,

    // Top-center: prevalence
    'XXXX (20%)': fmt(c.atRisk) + ' (20%)',

    // Program members (bottom-left big number)
    // This is tricky — "XXX" appears in many places.
    // The program members box has "XXX" as a standalone large number.
    // We handle it via the tier-specific patterns below.

    // ROI box
    'X.X : 1 NET ROI': c.roi.toFixed(1) + ' : 1 NET ROI',
    'XXX engaged': fmt(c.members) + ' engaged',
    '$XXM in savings': fmtM(c.savings) + ' in savings',

    // Program spend (center box)
    '$XXXXXXX': fmtD(c.spend),

    // Tier dollar amounts — these are $XXXXXX (6 X's)
    // Since all three tiers have the same placeholder, we need a different approach.
    // We'll handle these with the shape-level replacement below.
  };
}


/**
 * Remove substance risk tags from the slide for unselected substances.
 */
function removeUnselectedSubstances(slide, selectedSubs) {
  const allSubs = ['tud', 'aud', 'cud', 'oud'];

  allSubs.forEach(sub => {
    if (!selectedSubs.includes(sub)) {
      // Find and delete the shape containing this substance's tag text
      const tagText = CONFIG.SUBSTANCE_TAGS[sub];
      const shapes = slide.getShapes();
      shapes.forEach(shape => {
        try {
          const text = shape.getText().asString().trim();
          if (text === tagText || text.includes(tagText)) {
            shape.remove();
          }
        } catch (e) {
          // skip non-text shapes
        }
      });
    }
  });
}


/**
 * Advanced replacement: walks each shape to replace tier-specific values
 * that share the same placeholder pattern ($XXXXXX).
 *
 * Call this after the basic replaceAllText pass.
 */
function replaceTierValues(slide, c) {
  const fmt = (n) => n.toLocaleString('en-US');
  const fmtD = (n) => '$' + n.toLocaleString('en-US');

  const shapes = slide.getShapes();
  shapes.forEach(shape => {
    try {
      const text = shape.getText().asString();

      // Support tier box: contains "Support" and member count
      if (text.includes('Support') && text.includes('56%')) {
        shape.getText().replaceAllText('XXX', fmt(c.supM));
        shape.getText().replaceAllText('$XXXXXX', fmtD(c.supC));
      }
      // Manage tier box
      else if (text.includes('Manage') && text.includes('37%')) {
        shape.getText().replaceAllText('XXX', fmt(c.manM));
        shape.getText().replaceAllText('$XXXXXX', fmtD(c.manC));
      }
      // Treat tier box
      else if (text.includes('Treat') && text.includes('6%')) {
        shape.getText().replaceAllText('XX', fmt(c.trtM));
        shape.getText().replaceAllText('$XXXXXX', fmtD(c.trtC));
      }
      // Program members box (big "XXX" number)
      else if (text.includes('Pelago program members')) {
        shape.getText().replaceAllText('XXX', fmt(c.members));
      }
      // Program spend box
      else if (text.includes('program spend') && text.includes('$XXXXXX')) {
        shape.getText().replaceAllText('$XXXXXX', fmtD(c.spend));
      }
    } catch (e) {
      // skip non-text shapes
    }
  });
}


/**
 * Enhanced generateSlide that uses shape-level replacement for ambiguous placeholders.
 */
function generateSlideV2(params) {
  const { prospectName, partnerName, eligibleLives, pricingModel, substances } = params;

  // 1. Run calculations
  const c = calculate(eligibleLives, pricingModel);
  c.eligibleLives = eligibleLives;

  // 2. Duplicate the template
  const templateFile = DriveApp.getFileById(CONFIG.TEMPLATE_ID);
  const slideTitle = partnerName + ' — ' + prospectName + ' Value Story';
  const copy = templateFile.makeCopy(slideTitle);
  const newId = copy.getId();

  // 3. Open the new presentation
  const pres = SlidesApp.openById(newId);
  const slide = pres.getSlides()[0];

  const fmt = (n) => n.toLocaleString('en-US');
  const fmtD = (n) => '$' + n.toLocaleString('en-US');
  const fmtM = (n) => '$' + (n / 1e6).toFixed(1) + 'M';

  // 4. Safe global replacements (unique patterns only)
  slide.replaceAllText('X.X : 1 NET ROI', c.roi.toFixed(1) + ' : 1 NET ROI');
  slide.replaceAllText('$XXM in savings', fmtM(c.savings) + ' in savings');
  slide.replaceAllText('$X.XM Saved', fmtM(c.savings) + ' Saved');
  slide.replaceAllText('XXXX (20%)', fmt(c.atRisk) + ' (20%)');

  // 5. Shape-level replacements (context-aware for ambiguous patterns)
  const shapes = slide.getShapes();
  shapes.forEach(shape => {
    try {
      const text = shape.getText().asString();

      // Title bar
      if (text.includes('Members Engaged')) {
        shape.getText().replaceAllText('XXX Members Engaged', fmt(c.members) + ' Members Engaged');
        shape.getText().replaceAllText('X.X : 1', c.roi.toFixed(1) + ' : 1');
      }
      // Eligible lives box
      else if (text.includes('Eligible lives') || text.includes('full population')) {
        shape.getText().replaceAllText('XXXXX', fmt(c.eligibleLives));
        shape.getText().replaceAllText('PWGA', prospectName);
      }
      // Program members box
      else if (text.includes('Pelago program members')) {
        shape.getText().replaceAllText('XXX', fmt(c.members));
      }
      // Support tier
      else if (text.includes('Support') && text.includes('56%')) {
        shape.getText().replaceAllText('XXX', fmt(c.supM));
        shape.getText().replaceAllText('$XXXXXX', fmtD(c.supC));
      }
      // Manage tier
      else if (text.includes('Manage') && text.includes('37%')) {
        shape.getText().replaceAllText('XXX', fmt(c.manM));
        shape.getText().replaceAllText('$XXXXXX', fmtD(c.manC));
      }
      // Treat tier
      else if (text.includes('Treat') && text.includes('6%')) {
        shape.getText().replaceAllText('XX', fmt(c.trtM));
        shape.getText().replaceAllText('$XXXXXX', fmtD(c.trtC));
      }
      // Program spend box
      else if (text.includes('program spend')) {
        shape.getText().replaceAllText('$XXXXXX', fmtD(c.spend));
      }
      // ROI box — XXX engaged
      else if (text.includes('engaged') && text.includes('savings')) {
        shape.getText().replaceAllText('XXX engaged', fmt(c.members) + ' engaged');
      }
    } catch (e) {
      // skip non-text shapes
    }
  });

  // 6. Remove unselected substance tags
  removeUnselectedSubstances(slide, substances);

  // 7. Save and return
  pres.saveAndClose();

  return {
    url: 'https://docs.google.com/presentation/d/' + newId + '/edit',
    title: slideTitle,
    fileId: newId,
  };
}
