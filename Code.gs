/**
 * Pelago Value Story Generator — Google Apps Script Backend
 *
 * Serves the calculator web app, runs calculations, duplicates the
 * template slide, and populates every placeholder via the Slides API.
 */

// ── Configuration ──────────────────────────────────────────────────
const CONFIG = {
  // Template presentation (the master copy with X placeholders)
  // Update this to your template's file ID (the part between /d/ and /edit in the URL)
  TEMPLATE_ID: '1eHYLMaM1WnUHWc-AP1v2_QPpEOQ4QHw9',

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

  // Substance tag text on the template slide (used to selectively delete shapes)
  SUBSTANCE_TAGS: {
    tud: '8% of ELs TUD risk',
    aud: '10% of ELs AUD risk',
    cud: '6% of ELs CUD risk',
    oud: '2% of ELs OUD risk',
  },
};


// ── Web App Entry Point ────────────────────────────────────────────

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Pelago Value Story Generator')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ── Client-callable Functions ──────────────────────────────────────

/**
 * Runs the value story calculation.
 * Called from the front end for the number preview.
 *
 * @param {number} eligibleLives
 * @param {string} pricingModel - "Hawaii" or "Zanzibar"
 * @returns {Object} All calculated values
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
 * Duplicates the template, populates all values, removes unselected
 * substance tags, and returns the new presentation URL.
 *
 * This is the main action — called when the user clicks "Generate Slide."
 *
 * @param {Object} params - { prospectName, partnerName, eligibleLives, pricingModel, substances }
 * @returns {Object} - { url, title, fileId }
 */
function generateSlide(params) {
  const { prospectName, partnerName, eligibleLives, pricingModel, substances } = params;

  // 1. Run calculations
  const c = calculate(eligibleLives, pricingModel);
  c.eligibleLives = eligibleLives;

  // 2. Duplicate the template into the user's Drive
  const templateFile = DriveApp.getFileById(CONFIG.TEMPLATE_ID);
  const slideTitle = partnerName + ' \u2014 ' + prospectName + ' Value Story';
  const copy = templateFile.makeCopy(slideTitle);
  const newId = copy.getId();

  // 3. Open the new presentation
  const pres = SlidesApp.openById(newId);
  const slide = pres.getSlides()[0];

  // 4. Format helpers
  const fmt  = (n) => n.toLocaleString('en-US');
  const fmtD = (n) => '$' + n.toLocaleString('en-US');
  const fmtM = (n) => '$' + (n / 1e6).toFixed(1) + 'M';

  // 5. Global replacements — unique patterns only (safe across the whole slide)
  slide.replaceAllText('X.X : 1 NET ROI', c.roi.toFixed(1) + ' : 1 NET ROI');
  slide.replaceAllText('$XXM in savings', fmtM(c.savings) + ' in savings');
  slide.replaceAllText('$X.XM Saved', fmtM(c.savings) + ' Saved');
  slide.replaceAllText('XXXX (20%)', fmt(c.atRisk) + ' (20%)');

  // 6. Shape-level replacements — context-aware for ambiguous patterns
  //    (e.g. "XXX" and "$XXXXXX" appear in multiple text boxes with different values)
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
      // ROI box — "XXX engaged" line
      else if (text.includes('engaged') && text.includes('savings')) {
        shape.getText().replaceAllText('XXX engaged', fmt(c.members) + ' engaged');
      }
    } catch (e) {
      // skip non-text shapes (images, connectors, etc.)
    }
  });

  // 7. Remove substance tags the user unchecked
  removeUnselectedSubstances_(slide, substances);

  // 8. Save and return the new presentation URL
  pres.saveAndClose();

  return {
    url: 'https://docs.google.com/presentation/d/' + newId + '/edit',
    title: slideTitle,
    fileId: newId,
  };
}


// ── Private Helpers ────────────────────────────────────────────────

/**
 * Deletes shape elements from the slide for any substance the user deselected.
 * Matches shapes by their text content against CONFIG.SUBSTANCE_TAGS.
 */
function removeUnselectedSubstances_(slide, selectedSubs) {
  const allSubs = ['tud', 'aud', 'cud', 'oud'];

  allSubs.forEach(sub => {
    if (!selectedSubs.includes(sub)) {
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
