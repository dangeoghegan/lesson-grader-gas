/**
 * ============================================================================
 * CONSOLIDATED ASSESSMENT SYSTEM (YEAR 9 JEWELLERY DESIGN)
 * FULLY CORRECTED — Config now exports SHEET_SUBMISSIONS correctly, and the
 * Classroom import no longer crashes on invalid API parameters or missing
 * sheets. Section 8 (HtmlTemplates — the large embedded Studio HTML/CSS/JS)
 * is UNCHANGED from your existing file and is NOT reproduced below — keep
 * your current getAssessmentStudioHtml(), getCourseworkPickerHtml(),
 * getClassBatchGradingHtml(), and getVoiceDictationHtml() functions exactly
 * as they are. Replace everything else (Sections 1–7 and Section 9) with
 * the code below.
 * ============================================================================
 */

/* ----------------------------------------------------------------------------
 * BLOCK A — ADD new module: GradeScaleService.
 * Place this as a new top-level module, anywhere after Config and before
 * AssessmentService (e.g. directly above "var AssessmentService = ...").
 *
 * ASSUMED SETUP SHEET LAYOUT (please confirm/correct):
 *   A row containing the text "GRADE SCALE" marks the start of that table.
 *   Each following row has a letter (A-E) in one cell and a percentage
 *   number in another cell on the same row (e.g. "A" | 100%, or "A" | 100).
 *   The table ends at the first fully blank row after the header.
 *   The "OVERALL GRADE BANDS" table works the same way, but with a letter
 *   plus TWO numbers per row (min and max), e.g. "A" | 85 | 100.
 * ---------------------------------------------------------------------------- */
var GradeScaleService = (function() {
  var DEFAULT_WEIGHTS = { A: 1.0, B: 0.875, C: 0.70, D: 0.575, E: 0.25 };
  var DEFAULT_BANDS = [
    { letter: 'A', min: 85, max: 100 },
    { letter: 'B', min: 75, max: 84.99 },
    { letter: 'C', min: 65, max: 74.99 },
    { letter: 'D', min: 50, max: 64.99 },
    { letter: 'E', min: 0, max: 49.99 }
  ];

  function findHeaderRow_(data, keyword) {
    for (var r = 0; r < data.length; r++) {
      var rowStr = data[r].join(' ').toUpperCase();
      if (rowStr.indexOf(keyword) !== -1) return r;
    }
    return -1;
  }

  function extractLetterAndNumbers_(row) {
    var letter = null;
    var nums = [];
    for (var c = 0; c < row.length; c++) {
      var cell = row[c];
      if (letter === null && typeof cell === 'string' && /^[A-E]$/i.test(cell.trim())) {
        letter = cell.trim().toUpperCase();
        continue;
      }
      if (typeof cell === 'number') {
        nums.push(cell);
      } else if (typeof cell === 'string' && /^-?\d+(\.\d+)?%?$/.test(cell.trim())) {
        nums.push(parseFloat(cell.replace('%', '')));
      }
    }
    return { letter: letter, nums: nums };
  }

  function getGradeWeights() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(Config.SHEET_SETUP);
      if (!sheet) return DEFAULT_WEIGHTS;
      var data = sheet.getDataRange().getValues();
      var headerRow = findHeaderRow_(data, 'GRADE SCALE');
      if (headerRow === -1) return DEFAULT_WEIGHTS;

      var weights = {};
      for (var r = headerRow + 1; r < data.length; r++) {
        var rowJoined = data[r].join('').trim();
        if (!rowJoined) break;
        var parsed = extractLetterAndNumbers_(data[r]);
        if (parsed.letter && parsed.nums.length > 0) {
          var val = parsed.nums[0];
          weights[parsed.letter] = val > 1 ? val / 100 : val;
        }
        if (Object.keys(weights).length >= 5) break;
      }
      var complete = ['A', 'B', 'C', 'D', 'E'].every(function(l) { return weights[l] !== undefined; });
      return complete ? weights : DEFAULT_WEIGHTS;
    } catch (err) {
      Logging.logError('GradeScaleService.getGradeWeights', err);
      return DEFAULT_WEIGHTS;
    }
  }

  function getOverallGradeBands() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(Config.SHEET_SETUP);
      if (!sheet) return DEFAULT_BANDS;
      var data = sheet.getDataRange().getValues();
      var headerRow = findHeaderRow_(data, 'OVERALL GRADE BAND');
      if (headerRow === -1) return DEFAULT_BANDS;

      var bands = [];
      for (var r = headerRow + 1; r < data.length; r++) {
        var rowJoined = data[r].join('').trim();
        if (!rowJoined) break;
        var parsed = extractLetterAndNumbers_(data[r]);
        if (parsed.letter && parsed.nums.length >= 2) {
          bands.push({ letter: parsed.letter, min: parsed.nums[0], max: parsed.nums[1] });
        } else if (parsed.letter && parsed.nums.length === 1) {
          bands.push({ letter: parsed.letter, min: parsed.nums[0], max: 100 });
        }
        if (bands.length >= 5) break;
      }
      if (bands.length === 0) return DEFAULT_BANDS;
      bands.sort(function(a, b) { return b.min - a.min; });
      return bands;
    } catch (err) {
      Logging.logError('GradeScaleService.getOverallGradeBands', err);
      return DEFAULT_BANDS;
    }
  }

  /* Round-half-up letter selection: picks the nearest weight; an exact
     midpoint tie rounds UP to the higher letter. */
  function nearestGradeLetter(score, weights) {
    var letters = ['A', 'B', 'C', 'D', 'E'];
    var sorted = letters.filter(function(l) { return weights[l] !== undefined; })
      .sort(function(a, b) { return weights[b] - weights[a]; });
    if (!sorted.length) return 'E';
    for (var i = 0; i < sorted.length - 1; i++) {
      var upper = sorted[i], lower = sorted[i + 1];
      var midpoint = (weights[upper] + weights[lower]) / 2;
      if (score >= midpoint - 1e-9) return upper;
    }
    return sorted[sorted.length - 1];
  }

  function getOverallGradeLetter(totalPercent, bands) {
    bands = bands || getOverallGradeBands();
    for (var i = 0; i < bands.length; i++) {
      if (totalPercent >= bands[i].min) return bands[i].letter;
    }
    return bands.length ? bands[bands.length - 1].letter : 'E';
  }

  return {
    getGradeWeights: getGradeWeights,
    getOverallGradeBands: getOverallGradeBands,
    nearestGradeLetter: nearestGradeLetter,
    getOverallGradeLetter: getOverallGradeLetter
  };
})();



/* ============================================================================
 * SECTION 1: CONFIGURATION & CRITERIA DEFINITIONS
 * ============================================================================ */
var Config = (function() {
  var SHEET_SETUP = 'Setup';
  var SHEET_MARKING = '9DAT1 Marking';
  var SHEET_RUBRIC = 'Rubric';
  var SHEET_SUMMARY = 'Summary';

  var SHEET_CLASSROOM_CONFIG = 'ClassroomConfig';
  var SHEET_SUBMISSIONS = 'Submissions';
  var SHEET_SUBMISSION_FILES = 'SubmissionFiles';
  var SHEET_CRITERIA_CONFIG = 'CriteriaConfig';
  var SHEET_CRITERION_ASSESSMENTS = 'CriterionAssessments';
  var SHEET_AI_ASSESSMENTS = 'AIAssessments';
  var SHEET_ASSESSMENT_HISTORY = 'AssessmentHistory';
  var SHEET_ERROR_LOG = 'ErrorLog';

  var SCRIPT_PROP_GEMINI_KEY = 'GEMINI_API_KEY';
  var PRIMARY_MODEL = 'gemini-2.5-flash';
  var FALLBACK_MODELS = ['gemini-2.0-flash', 'gemini-1.5-flash'];
  var GEMINI_API_VERSION = 'v1beta';

  var GRADE_WEIGHTS = { 'A': 1.000, 'B': 0.875, 'C': 0.700, 'D': 0.575, 'E': 0.250 };
  var STATUS = {
    NEW: 'New', IN_REVIEW: 'InReview', AI_ASSESSED: 'AIAssessed',
    TEACHER_REVIEWED: 'TeacherReviewed', APPROVED: 'Approved',
    LOCKED: 'Locked', SUPERSEDED: 'Superseded', SYNC_FAILED: 'SyncFailed'
  };

  var CRITERIA_DEFINITIONS = [
    {
      criterionId: "C01", title: "Identification & Exploration of Need", part: "Part A", section: "Section 1: Project Proposal & Management", maxMarks: 5, outcome: "DT5-2",
      mdOverride: { id: "C01-MD", text: "Missing / Distinct: clear evidence beyond listed observables at an A standard." },
      bands: {
        E: [["C01-E-1", "Design brief present."], ["C01-E-2", "Identifies design opportunity/need."], ["C01-E-3", "At least two success criteria listed."], ["C01-E-4", "Understandable and on-task."]],
        D: [["C01-D-1", "At least three success criteria outlined."], ["C01-D-2", "Most criteria relevant to need."], ["C01-D-3", "Criteria specific enough to check."], ["C01-D-4", "Basic consideration of user/context."]],
        C: [["C01-C-1", "At least four sound success criteria outlined."], ["C01-C-2", "Specific enough to guide decisions."], ["C01-C-3", "Describes design opportunity with some detail."], ["C01-C-4", "Considers user and situation."]],
        B: [["C01-B-1", "At least five mostly relevant success criteria outlined."], ["C01-B-2", "Clear explanation of design opportunity."], ["C01-B-3", "Considers constraints (time/cost/safety)."], ["C01-B-4", "Proposed solution characteristics well-aligned."]],
        A: [["C01-A-1", "SIX highly relevant and accurate success criteria."], ["C01-A-2", "Comprehensive design brief explaining solution clearly."], ["C01-A-3", "Reflects deep understanding of designer considerations."], ["C01-A-4", "Addresses constraints and real wearer needs thoughtfully."]]
      }
    },
    {
      criterionId: "C02", title: "Inspiration (Annotated Mood Board)", part: "Part A", section: "Section 1: Project Proposal & Management", maxMarks: 10, outcome: "DT5-2",
      mdOverride: { id: "C02-MD", text: "Missing / Distinct: exceptionally curated thematic visual inquiry." },
      bands: {
        E: [["C02-E-1", "Mood board present."], ["C02-E-2", "3-4 jewellery images."], ["C02-E-3", "Basic relevance identifiable."], ["C02-E-4", "Minimal labelling."]],
        D: [["C02-D-1", "Basic inspiration collection (5+ sources)."], ["C02-D-2", "Exploration of jewellery styles/materials."], ["C02-D-3", "Brief annotations identify images."], ["C02-D-4", "Basic connection to initial ideas."]],
        C: [["C02-C-1", "Sound range of 6-8 diverse inspirations."], ["C02-C-2", "Annotations explain basic relevance & design influence."], ["C02-C-3", "Identifies specific design elements (colour palette, motifs)."], ["C02-C-4", "Relevant to theme and target market."]],
        B: [["C02-B-1", "Thorough range of 8+ well-curated inspirations."], ["C02-B-2", "Clear annotations explaining relevance and influence."], ["C02-B-3", "Analyses links to materials and wearer style."], ["C02-B-4", "Structured, cohesive emerging aesthetic direction."]],
        A: [["C02-A-1", "Extensive, highly diverse inspiration sources."], ["C02-A-2", "Insightful, detailed annotations explaining aesthetic influence."], ["C02-A-3", "Explicit links to planned components/techniques."], ["C02-A-4", "Exceptional visual presentation and creative vision."]]
      }
    },
    {
      criterionId: "C03", title: "Action Plan / Gantt Chart", part: "Part A", section: "Section 1: Project Proposal & Management", maxMarks: 5, outcome: "DT5-10",
      mdOverride: { id: "C03-MD", text: "Missing / Distinct: dynamic digital tracking with milestone adaptation." },
      bands: {
        E: [["C03-E-1", "Action plan/timeline present."], ["C03-E-2", "Lists 3-4 basic tasks."], ["C03-E-3", "Indication of dates or order."], ["C03-E-4", "Loosely related to jewellery project."]],
        D: [["C03-D-1", "Outlines 5 key tasks for Part A & B."], ["C03-D-2", "Estimated timeframes or weeks assigned."], ["C03-D-3", "Sequential progression shown."], ["C03-D-4", "Due dates identified."]],
        C: [["C03-C-1", "Action plan completed soundly with key details."], ["C03-C-2", "Differentiates research, design, and making."], ["C03-C-3", "Logical and realistic sequencing."], ["C03-C-4", "Notes resources and checkpoints."]],
        B: [["C03-B-1", "Action plan completed thoroughly across stages."], ["C03-B-2", "Clear durations and estimated vs actual tracking."], ["C03-B-3", "Reflects workshop and curing constraints."], ["C03-B-4", "Evidence of ongoing monitoring and feedback."]],
        A: [["C03-A-1", "Action plan completed extensively with wide details."], ["C03-A-2", "Precise milestones, dependencies, and WHS checkpoints."], ["C03-A-3", "Integrated evaluation checkpoints and contingency time."], ["C03-A-4", "Rigorous management of time and equipment."]]
      }
    },
    {
      criterionId: "C04", title: "Application of Experiments of Materials (2 Techniques)", part: "Part A", section: "Section 2: Project Development & Realisation", maxMarks: 10, outcome: "DT5-9",
      mdOverride: { id: "C04-MD", text: "Missing / Distinct: unusually sophisticated experimentation or innovative testing." },
      bands: {
        E: [["C04-E-1", "Evidence of 1 material/technique test piece."], ["C04-E-2", "Identifiable material or technique."], ["C04-E-3", "Basic photo/note documentation."], ["C04-E-4", "Loosely connected to jewellery task."]],
        D: [["C04-D-1", "One OR two techniques researched with basic samples."], ["C04-D-2", "Basic engagement rather than token attempt."], ["C04-D-3", "Simple reference samples included."], ["C04-D-4", "Basic reflection on what worked/failed."]],
        C: [["C04-C-1", "Two jewellery techniques soundly researched."], ["C04-C-2", "Reasonable quality reference samples included."], ["C04-C-3", "Original evidence of development of both techniques."], ["C04-C-4", "Comparison of technique suitability."]],
        B: [["C04-B-1", "Two techniques thoroughly researched with high quality samples."], ["C04-B-2", "Original evidence of development of both techniques."], ["C04-B-3", "Thoughtful parameter testing (tension, curing, joins)."], ["C04-B-4", "Explains influence on final design decisions."]],
        A: [["C04-A-1", "Two techniques extensively researched with outstanding samples."], ["C04-A-2", "Original multi-stage development evidence beyond basic class templates."], ["C04-A-3", "Detailed technical reasoning linking tests to final choices."], ["C04-A-4", "High-level understanding of material properties."]]
      }
    },
    {
      criterionId: "C05", title: "Description of Inspirational Australian Jeweller", part: "Part A", section: "Section 2: Project Development & Realisation", maxMarks: 5, outcome: "DT5-4",
      mdOverride: { id: "C05-MD", text: "Missing / Distinct: primary correspondence or critical contemporary market analysis." },
      bands: {
        E: [["C05-E-1", "Australian jeweller identified."], ["C05-E-2", "At least 1 image of work/biography."], ["C05-E-3", "Minimal biographical details."], ["C05-E-4", "Information related to jewellery."]],
        D: [["C05-D-1", "Basic outline of biographical information."], ["C05-D-2", "Covers a few details about their background."], ["C05-D-3", "Shows examples of distinctive work."], ["C05-D-4", "Basic interest/connection stated."]],
        C: [["C05-C-1", "Sound description covering key biographical details."], ["C05-C-2", "Describes style, philosophy, and materials."], ["C05-C-3", "Multiple images with descriptive captions."], ["C05-C-4", "Explains what inspires the student."]],
        B: [["C05-B-1", "Thorough description covering a range of details."], ["C05-B-2", "Detailed examination of signature techniques and aesthetic."], ["C05-B-3", "Analyses market trends and customer focus."], ["C05-B-4", "Links designer practice to student's work."]],
        A: [["C05-A-1", "Extensive biographical case study covering wide details."], ["C05-A-2", "Critical analysis of technical mastery and design identity."], ["C05-A-3", "Insightful parallels drawn to student's collection brief."], ["C05-A-4", "Synthesized in own words with proper attribution."]]
      }
    },
    {
      criterionId: "C06", title: "Explanations of Design Elements & Principles", part: "Part A", section: "Section 2: Project Development & Realisation", maxMarks: 10, outcome: "DT5-4",
      mdOverride: { id: "C06-MD", text: "Missing / Distinct: visual semiotic analysis or advanced design theory critique." },
      bands: {
        E: [["C06-E-1", "1-2 elements/principles named."], ["C06-E-2", "Minimal descriptive definitions."], ["C06-E-3", "At least one illustrative image."], ["C06-E-4", "Recognisable design terminology."]],
        D: [["C06-D-1", "Outlines and basic information regarding elements/principles."], ["C06-D-2", "Identifies 3 distinct elements/principles."], ["C06-D-3", "Connects to simple jewellery examples."], ["C06-D-4", "Reasonable accuracy of basic terms."]],
        C: [["C06-C-1", "Mostly descriptive explanations of design elements & principles."], ["C06-C-2", "Covers 4-5 relevant elements/principles."], ["C06-C-3", "Applies terminology accurately to jewellery designs."], ["C06-C-4", "Explains visual interest enhancements."]],
        B: [["C06-B-1", "Explanations and detailed information regarding elements/principles."], ["C06-B-2", "Analyses interaction of balance, contrast, and form."], ["C06-B-3", "Strong application to planned jewellery collection."], ["C06-B-4", "Consistently high standard of design language."]],
        A: [["C06-A-1", "Extensive explanations and detailed theoretical analysis."], ["C06-A-2", "Deep insight into balance, scale, and proportion."], ["C06-A-3", "Synthesizes theory with practical fabrication."], ["C06-A-4", "Professional annotations displaying mastery."]]
      }
    },
    {
      criterionId: "C07", title: "Analysis of Factor Affecting Design", part: "Part A", section: "Section 2: Project Development & Realisation", maxMarks: 10, outcome: "DT5-4",
      mdOverride: { id: "C07-MD", text: "Missing / Distinct: life-cycle assessment or comprehensive ergonomics inquiry." },
      bands: {
        E: [["C07-E-1", "Piece identified for analysis."], ["C07-E-2", "Mentions at least 1 factor (cost/function)."], ["C07-E-3", "Minimal comments on factor."], ["C07-E-4", "Loosely relevant design analysis."]],
        D: [["C07-D-1", "Basic outline of factors affecting one chosen piece."], ["C07-D-2", "Basic description of factor impacts."], ["C07-D-3", "Outlines user/situational requirements."], ["C07-D-4", "Simple cause-and-effect reasoning."]],
        C: [["C07-C-1", "Sound description of factors affecting one chosen piece."], ["C07-C-2", "Analyses 2-3 factors (function/ethics/cost/form)."], ["C07-C-3", "Considers wearability and durability."], ["C07-C-4", "Draws relevant conclusions for own project."]],
        B: [["C07-B-1", "Thorough explanation of factors affecting one chosen piece."], ["C07-B-2", "Detailed interplay of cost, ethics, manufacturing, and form."], ["C07-B-3", "Analyses how constraints were resolved."], ["C07-B-4", "Explicitly links findings to own parameters."]],
        A: [["C07-A-1", "Extensive analysis of chosen piece by designer of choice."], ["C07-A-2", "Critical evaluation of cultural, ergonomic, and environmental factors."], ["C07-A-3", "Sophisticated vocabulary regarding sustainability."], ["C07-A-4", "Deep synthesis guiding structural choices."]]
      }
    },
    {
      criterionId: "C08", title: "Evidence of Creativity (3 Sketches + Final Designs)", part: "Part B", section: "Section 2: Evidence of Creativity", maxMarks: 5, outcome: "DT5-2",
      mdOverride: { id: "C08-MD", text: "Missing / Distinct: highly original design development or innovative visual presentation." },
      bands: {
        E: [["C08-E-1", "Very few or incomplete initial sketches."], ["C08-E-2", "Final designs underdeveloped and unclear."], ["C08-E-3", "Minimal annotations present."], ["C08-E-4", "Loosely related to jewellery task."]],
        D: [["C08-D-1", "Limited initial sketches with minimal variation."], ["C08-D-2", "Final designs show limited refinement."], ["C08-D-3", "Basic presentation with minimal colour/notes."], ["C08-D-4", "Clear connection to jewellery brief."]],
        C: [["C08-C-1", "Clear initial sketches that meet the brief with some variation."], ["C08-C-2", "Final designs show basic development."], ["C08-C-3", "Neat presentation with some colour."], ["C08-C-4", "Simple annotations of materials/dimensions."]],
        B: [["C08-B-1", "Three relevant initial sketches showing clear ideas."], ["C08-B-2", "Final designs well-presented with appropriate rendering/colour."], ["C08-B-3", "Annotations explain key features and construction."], ["C08-B-4", "Visible development and purposeful creativity."]],
        A: [["C08-A-1", "Three highly creative, detailed initial sketches with strong variation."], ["C08-A-2", "Refined into three final designs to professional standard."], ["C08-A-3", "Accurate rendering of proportion, texture, and colour."], ["C08-A-4", "Detailed annotations explaining design decisions."]]
      }
    },
    {
      criterionId: "C09", title: "Order of Construction (Steps + Materials & Tools)", part: "Part B", section: "Section 2: Evidence of Creativity", maxMarks: 10, outcome: "DT5-10",
      mdOverride: { id: "C09-MD", text: "Missing / Distinct: production flowchart with risk matrix and quality checks." },
      bands: {
        E: [["C09-E-1", "Attempts an Order of Construction but unclear or incomplete."], ["C09-E-2", "Resource list minimal or missing essential items."], ["C09-E-3", "Mentions some tools/materials."], ["C09-E-4", "Broadly relates to making jewellery."]],
        D: [["C09-D-1", "Simple Order of Construction with limited sequencing."], ["C09-D-2", "Basic list of materials and tools with some relevant items."], ["C09-D-3", "Missing key resources or detail."], ["C09-D-4", "Basic safety notes included."]],
        C: [["C09-C-1", "Generally clear Order of Construction with logical sequencing."], ["C09-C-2", "Suitable list of materials and tools."], ["C09-C-3", "Steps describe specific actions and techniques."], ["C09-C-4", "Relevant workshop WHS precautions noted."]],
        B: [["C09-B-1", "Clear and logical Order of Construction with appropriate sequencing."], ["C09-B-2", "Detailed list of materials and tools mostly comprehensive."], ["C09-B-3", "Technical notes explaining exact joining/shaping steps."], ["C09-B-4", "WHS and PPE integrated into procedure."]],
        A: [["C09-A-1", "Extensively constructs logical, clear Order of Construction."], ["C09-A-2", "Well-sequenced method to complete jewellery piece."], ["C09-A-3", "Extensive, clearly documented list of materials and tools."], ["C09-A-4", "Comprehensive safety management embedded."]]
      }
    },
    {
      criterionId: "C10", title: "Final Evaluation", part: "Part B", section: "Section 3: Project Evaluation", maxMarks: 10, outcome: "DT5-2",
      mdOverride: { id: "C10-MD", text: "Missing / Distinct: authentic consumer market testing data or lifecycle analysis." },
      bands: {
        E: [["C10-E-1", "Limited information on final product range/process."], ["C10-E-2", "Minimal relation to design brief."], ["C10-E-3", "States basic personal satisfaction."], ["C10-E-4", "Relevant to finished project."]],
        D: [["C10-D-1", "Outlines final product range and process."], ["C10-D-2", "Relates to design brief and limitations at a basic level."], ["C10-D-3", "Identifies a challenge and simple modification."], ["C10-D-4", "Suggests basic future improvement."]],
        C: [["C10-C-1", "Describes final product range and process soundly."], ["C10-C-2", "Relates to design brief and limitations addressing key components."], ["C10-C-3", "Reflects on material/time difficulties."], ["C10-C-4", "Meaningful suggestions for improvement."]],
        B: [["C10-B-1", "Explains final product range and process thoroughly."], ["C10-B-2", "Relates to design brief and limitations addressing many components."], ["C10-B-3", "Critical reflection on aesthetic cohesion and function."], ["C10-B-4", "Constructive proposals for refinement."]],
        A: [["C10-A-1", "Critically evaluates final product range and process."], ["C10-A-2", "Relates comprehensively to design brief and limitations."], ["C10-A-3", "Deep reflection on material limitations and personal skills."], ["C10-A-4", "Insightful recommendations for scalable production."]]
      }
    },
    {
      criterionId: "C11", title: "Construction of Jewellery (Practical Skills)", part: "Part B", section: "Section 2: Practical Project", maxMarks: 10, outcome: "DT5-9",
      mdOverride: { id: "C11-MD", text: "Missing / Distinct: exceptionally refined finishing, complex joins, or innovative construction." },
      bands: {
        E: [["C11-E-1", "Limited evidence of uniformity and quality in techniques."], ["C11-E-2", "Piece basically intact."], ["C11-E-3", "Wearable/usable rather than falling apart."], ["C11-E-4", "Photos show basic construction."]],
        D: [["C11-D-1", "Basic evidence of uniformity and quality in required techniques."], ["C11-D-2", "Demonstrates basic skill development."], ["C11-D-3", "Sharp ends managed; item functions."], ["C11-D-4", "Intentionally made rather than accidental."]],
        C: [["C11-C-1", "Sound evidence of uniformity and quality in some techniques."], ["C11-C-2", "Demonstrates reasonable skill development."], ["C11-C-3", "Joins secure, components aligned, and sturdy."], ["C11-C-4", "Reasonable consistency in repeated links/beads."]],
        B: [["C11-B-1", "Thorough evidence of uniformity and quality in most techniques."], ["C11-B-2", "Demonstrates high standard skill development."], ["C11-B-3", "High-quality finish with clean joins and neat edges."], ["C11-B-4", "Careful material handling and minimal visible flaws."]],
        A: [["C11-A-1", "Extensive evidence of uniformity and quality in all required techniques."], ["C11-A-2", "Demonstrating outstanding skill development."], ["C11-A-3", "Consistently neat, secure, flawless joins."], ["C11-A-4", "Polished, professional surface finishing."]]
      }
    },
    {
      criterionId: "C12", title: "Jewellery Aesthetics (Design Quality & Range Cohesion)", part: "Part B", section: "Section 2: Practical Project", maxMarks: 10, outcome: "DT5-9",
      mdOverride: { id: "C12-MD", text: "Missing / Distinct: unified couture collection exhibiting gallery presentation." },
      bands: {
        E: [["C12-E-1", "Limited displays of jewellery design displaying minimal creativity."], ["C12-E-2", "At least 2 pieces share basic theme/colour."], ["C12-E-3", "Aesthetics discernible."], ["C12-E-4", "Identifiable personal adornment set."]],
        D: [["C12-D-1", "Basic displays of jewellery design displaying creativity to audience."], ["C12-D-2", "Some cohesion across pieces (similar beads/colours)."], ["C12-D-3", "Basic consideration of target tastes."], ["C12-D-4", "Neat presentation communicating style."]],
        C: [["C12-C-1", "Sound displays of jewellery design matching reasonably."], ["C12-C-2", "Displays reasonable level of creativity and appropriateness."], ["C12-C-3", "Materials and equipment used are suitable."], ["C12-C-4", "Clear thematic links across 3 pieces."]],
        B: [["C12-B-1", "High level displays where pieces mostly match perfectly."], ["C12-B-2", "Displays thorough levels of creativity and appropriateness."], ["C12-B-3", "Very unique pieces suitable for purpose."], ["C12-B-4", "High presentation quality and styled photography."]],
        A: [["C12-A-1", "Outstanding displays where pieces match perfectly and are very unique."], ["C12-A-2", "Extensive level of creativity and appropriateness to target audience."], ["C12-A-3", "Materials and equipment very suitable for design purpose."], ["C12-A-4", "Exhibition quality presentation and brand identity."]]
      }
    }
  ];

  function getSerializedRubricPrompt() {
    var out = [];
    for (var i = 0; i < CRITERIA_DEFINITIONS.length; i++) {
      var c = CRITERIA_DEFINITIONS[i];
      var s = "CRITERION " + c.criterionId + ": " + c.title + " (" + c.part + ", /" + c.maxMarks + " Marks)\n";
      var bands = ['A', 'B', 'C', 'D', 'E'];
      for (var b = 0; b < bands.length; b++) {
        var bn = bands[b];
        var items = c.bands[bn] || [];
        var itemArr = [];
        for (var k = 0; k < items.length; k++) {
          itemArr.push("[" + items[k][0] + "] " + items[k][1]);
        }
        s += "  - Band " + bn + ": " + itemArr.join("; ") + "\n";
      }
      out.push(s);
    }
    return out.join("\n");
  }

  return {
    /* Modern (underscore) names — used throughout this codebase */
    SHEET_SETUP: SHEET_SETUP, SHEET_MARKING: SHEET_MARKING, SHEET_RUBRIC: SHEET_RUBRIC, SHEET_SUMMARY: SHEET_SUMMARY,
    SHEET_CLASSROOM_CONFIG: SHEET_CLASSROOM_CONFIG, SHEET_SUBMISSIONS: SHEET_SUBMISSIONS, SHEET_SUBMISSION_FILES: SHEET_SUBMISSION_FILES,
    SHEET_CRITERIA_CONFIG: SHEET_CRITERIA_CONFIG, SHEET_CRITERION_ASSESSMENTS: SHEET_CRITERION_ASSESSMENTS,
    SHEET_AI_ASSESSMENTS: SHEET_AI_ASSESSMENTS, SHEET_ASSESSMENT_HISTORY: SHEET_ASSESSMENT_HISTORY, SHEET_ERROR_LOG: SHEET_ERROR_LOG,
    SCRIPT_PROP_GEMINI_KEY: SCRIPT_PROP_GEMINI_KEY, PRIMARY_MODEL: PRIMARY_MODEL, FALLBACK_MODELS: FALLBACK_MODELS,
    GEMINI_API_VERSION: GEMINI_API_VERSION, GRADE_WEIGHTS: GRADE_WEIGHTS, STATUS: STATUS, CRITERIA_DEFINITIONS: CRITERIA_DEFINITIONS,

    /* Legacy (no-underscore) aliases — kept so any old references never see "undefined" again */
    SHEETSETUP: SHEET_SETUP, SHEETMARKING: SHEET_MARKING, SHEETRUBRIC: SHEET_RUBRIC, SHEETSUMMARY: SHEET_SUMMARY,
    SHEETCLASSROOMCONFIG: SHEET_CLASSROOM_CONFIG, SHEETSUBMISSIONS: SHEET_SUBMISSIONS, SHEETSUBMISSIONFILES: SHEET_SUBMISSION_FILES,
    SHEETCRITERIACONFIG: SHEET_CRITERIA_CONFIG, SHEETCRITERIONASSESSMENTS: SHEET_CRITERION_ASSESSMENTS,
    SHEETAIASSESSMENTS: SHEET_AI_ASSESSMENTS, SHEETASSESSMENTHISTORY: SHEET_ASSESSMENT_HISTORY, SHEETERRORLOG: SHEET_ERROR_LOG,
    SCRIPTPROPGEMINIKEY: SCRIPT_PROP_GEMINI_KEY, PRIMARYMODEL: PRIMARY_MODEL, FALLBACKMODELS: FALLBACK_MODELS,
    GEMINIAPIVERSION: GEMINI_API_VERSION, GRADEWEIGHTS: GRADE_WEIGHTS, CRITERIADEFINITIONS: CRITERIA_DEFINITIONS,

    getSerializedRubricPrompt: getSerializedRubricPrompt
  };
})();

/* ============================================================================
 * SECTION 2: UTILITIES & AUDIT LOGGING (SAFE USER RESOLUTION)
 * ============================================================================ */
var Utils = (function() {
  function generateUuid() { return Utilities.getUuid(); }
  function formatDate(d) {
    if (!d) return '';
    if (typeof d === 'string') return d;
    return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Australia/Sydney', "yyyy-MM-dd'T'HH:mm:ssXXX");
  }
  function safeJsonParse(str, fallback) {
    if (!str) return fallback;
    if (typeof str === 'object') return str;
    try { return JSON.parse(str); } catch (e) { return fallback; }
  }
  function safeJsonStringify(obj) {
    try { return JSON.stringify(obj); } catch (e) { return '{}'; }
  }
  function sanitizeError(err) {
    if (!err) return 'Unknown error';
    var msg = err.message || String(err);
    return msg.replace(/key=[a-zA-Z0-9_-]+/gi, 'key=REDACTED');
  }
  function getSafeUserEmail() {
    try {
      return Session.getActiveUser().getEmail() || 'Teacher';
    } catch (e) {
      return 'Teacher';
    }
  }
  return {
    generateUuid: generateUuid,
    formatDate: formatDate,
    safeJsonParse: safeJsonParse,
    safeJsonStringify: safeJsonStringify,
    sanitizeError: sanitizeError,
    getSafeUserEmail: getSafeUserEmail
  };
})();

var Logging = (function() {
  function logHistory(submissionRecordId, assessmentId, action, criterionId, previousState, newState, notes) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(Config.SHEET_ASSESSMENT_HISTORY);
      if (!sheet) return;
      sheet.appendRow([
        Utils.generateUuid(), submissionRecordId || '', assessmentId || '', action || '', criterionId || '',
        Utils.safeJsonStringify(previousState || {}), Utils.safeJsonStringify(newState || {}),
        Utils.getSafeUserEmail(), Utils.formatDate(new Date()), notes || ''
      ]);
    } catch (e) { console.error('History log error: ' + e.message); }
  }

  function logError(context, errorObj, metadata) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(Config.SHEET_ERROR_LOG);
      if (!sheet) return;
      sheet.appendRow([
        Utils.generateUuid(), Utils.formatDate(new Date()), context || 'General',
        Utils.sanitizeError(errorObj), (errorObj && errorObj.stack) ? Utils.sanitizeError(errorObj.stack) : '',
        Utils.safeJsonStringify(metadata || {}), Utils.getSafeUserEmail(), 'Logged'
      ]);
    } catch (e) { console.error('ErrorLog write failed: ' + e.message); }
  }
  return { logHistory: logHistory, logError: logError };
})();

/* ============================================================================
 * SECTION 3: SHEET MANAGEMENT & FORMULAS
 * ============================================================================ */
var Sheets = (function() {
  function setupOrMigrateAssessmentSystem() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var schemas = [
      { name: Config.SHEET_CLASSROOM_CONFIG, headers: ['ConfigID', 'CourseID', 'CourseName', 'CourseSection', 'CourseWorkID', 'AssignmentTitle', 'MaxPoints', 'SavedAt', 'SavedBy', 'Active'] },
      { name: Config.SHEET_SUBMISSIONS, headers: ['SubmissionRecordID', 'ClassroomCourseID', 'ClassroomCourseWorkID', 'ClassroomSubmissionID', 'StudentUserID', 'StudentName', 'StudentEmail', 'Class', 'Task', 'SubmissionVersion', 'SourceType', 'ClassroomState', 'TurnedInTime', 'UpdateTime', 'Late', 'AttachmentSummary', 'AttachmentFileIDsJSON', 'AttachmentMetadataJSON', 'DriveFolderID', 'Status', 'ParentSubmissionRecordID', 'CurrentOfficial', 'LockedAt', 'LockedBy', 'ApprovedAt', 'ApprovedBy', 'ClassroomAssignedGrade', 'LastClassroomSyncAt', 'LastSyncResult', 'Notes'] },
      { name: Config.SHEET_SUBMISSION_FILES, headers: ['SubmissionRecordID', 'FileRecordID', 'SourceType', 'DriveFileID', 'FileName', 'MimeType', 'AlternateLink', 'ThumbnailUrl', 'FileSize', 'EligibleForAI', 'AIReviewStatus', 'AIExtractedText', 'Limitations', 'CreatedAt'] },
      { name: Config.SHEET_CRITERIA_CONFIG, headers: ['CriterionID', 'CriterionTitle', 'Part', 'Section', 'MaxMarks', 'Outcome', 'Band', 'ThresholdType', 'RequiredCount', 'CheckboxID', 'CheckboxLabel', 'EvidenceFocus', 'IsMissingDistinctOverride'] },
      { name: Config.SHEET_CRITERION_ASSESSMENTS, headers: ['AssessmentID', 'SubmissionRecordID', 'CriterionID', 'AIProposedGrade', 'DerivedGrade', 'FinalApprovedGrade', 'DistinctOverrideSelected', 'CheckboxesJSON', 'AIEvidenceJSON', 'TeacherAdjusted', 'TeacherAudioTranscript', 'TeacherWrittenNote', 'Confidence', 'MinimumEvidenceIncomplete', 'Status', 'CreatedAt', 'UpdatedAt', 'ApprovedAt', 'ApprovedBy'] },
      { name: Config.SHEET_AI_ASSESSMENTS, headers: ['AssessmentID', 'SubmissionRecordID', 'RunType', 'RunTimestamp', 'ModelUsed', 'PromptVersion', 'InputFilesJSON', 'PromptSummary', 'ResponseJSON', 'ExecutionStatus', 'ErrorMessage', 'LatencySeconds', 'TriggeredBy', 'TeacherOutcome'] },
      { name: Config.SHEET_ASSESSMENT_HISTORY, headers: ['HistoryID', 'SubmissionRecordID', 'AssessmentID', 'Action', 'CriterionID', 'PreviousStateJSON', 'NewStateJSON', 'ActorEmail', 'Timestamp', 'Notes'] },
      { name: Config.SHEET_ERROR_LOG, headers: ['ErrorID', 'Timestamp', 'Context', 'ErrorMessage', 'StackTrace', 'MetadataJSON', 'UserEmail', 'ResolutionStatus'] }
    ];

    for (var i = 0; i < schemas.length; i++) {
      var sDef = schemas[i];
      var sheet = ss.getSheetByName(sDef.name);
      if (!sheet) sheet = ss.insertSheet(sDef.name);
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(sDef.headers);
        var headerRange = sheet.getRange(1, 1, 1, sDef.headers.length);
        headerRange.setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
        sheet.setFrozenRows(1);
      }
    }

    populateCriteriaConfig(ss);
    verifyMarkingSheetFormulas(ss);
    return { success: true, message: 'All tables and configurations verified.' };
  }

  function populateCriteriaConfig(ss) {
    var cSheet = ss.getSheetByName(Config.SHEET_CRITERIA_CONFIG);
    if (!cSheet || cSheet.getLastRow() > 1) return;
    var defs = Config.CRITERIA_DEFINITIONS;
    var rows = [];

    for (var i = 0; i < defs.length; i++) {
      var c = defs[i];
      rows.push([c.criterionId, c.title, c.part, c.section, c.maxMarks, c.outcome, 'A', 'Override', 1, c.mdOverride.id, c.mdOverride.text, 'Distinct Evidence', true]);
      var bands = ['A', 'B', 'C', 'D', 'E'];
      for (var b = 0; b < bands.length; b++) {
        var bName = bands[b];
        var items = c.bands[bName] || [];
        var req = (bName === 'A') ? items.length : Math.ceil(items.length * 2 / 3);
        for (var k = 0; k < items.length; k++) {
          rows.push([c.criterionId, c.title, c.part, c.section, c.maxMarks, c.outcome, bName, (bName === 'A' ? 'All' : 'TwoThirds'), req, items[k][0], items[k][1], 'Observable', false]);
        }
      }
    }
    if (rows.length > 0) cSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }

  function verifyMarkingSheetFormulas(ss) {
    var sheet = ss.getSheetByName(Config.SHEET_MARKING);
    if (!sheet) return;
    var numRows = sheet.getLastRow();
    if (numRows < 2) return;

    var W = GradeScaleService.getGradeWeights();
    var bands = GradeScaleService.getOverallGradeBands();
    var defs = Config.CRITERIA_DEFINITIONS;
    var colLetters = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N'];

    function ifChain_(col, maxMarks, r) {
      var wA = Math.round(maxMarks * W.A * 1000) / 1000;
      var wB = Math.round(maxMarks * W.B * 1000) / 1000;
      var wC = Math.round(maxMarks * W.C * 1000) / 1000;
      var wD = Math.round(maxMarks * W.D * 1000) / 1000;
      var wE = Math.round(maxMarks * W.E * 1000) / 1000;
      return 'IF(' + col + r + '="A",' + wA +
        ',IF(' + col + r + '="B",' + wB +
        ',IF(' + col + r + '="C",' + wC +
        ',IF(' + col + r + '="D",' + wD + ',' + wE + '))))';
    }

    var ifsArgs = [];
    for (var bi = 0; bi < bands.length; bi++) {
      if (bi === bands.length - 1) {
        ifsArgs.push('TRUE,"' + bands[bi].letter + '"');
      } else {
        ifsArgs.push('{R}>=' + bands[bi].min + ',"' + bands[bi].letter + '"');
      }
    }

    for (var r = 2; r <= numRows; r++) {
      if (!sheet.getRange(r, 2).getValue()) continue;

      var partATerms = [];
      var partBTerms = [];
      for (var c = 0; c < colLetters.length; c++) {
        var def = defs[c];
        if (!def) continue;
        var term = ifChain_(colLetters[c], def.maxMarks, r);
        if (def.part === 'Part A') partATerms.push(term);
        else partBTerms.push(term);
      }

      sheet.getRange(r, 15).setFormula('=SUM(' + partATerms.join(',') + ')');
      sheet.getRange(r, 16).setFormula('=SUM(' + partBTerms.join(',') + ')');
      sheet.getRange(r, 17).setFormula('=O' + r + '+P' + r);

      var rowIfs = 'IFS(' + ifsArgs.join(',').replace(/\{R\}/g, 'Q' + r) + ')';
      sheet.getRange(r, 18).setFormula('=' + rowIfs);
    }
  }

  function syncApprovedGradesToMarkingSheet(studentName, gradesObj, feedbackObj) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(Config.SHEET_MARKING);
    if (!sheet) throw new Error('Sheet ' + Config.SHEET_MARKING + ' not found.');

    var data = sheet.getDataRange().getValues();
    var targetRow = -1;
    for (var r = 1; r < data.length; r++) {
      if (data[r][1] && String(data[r][1]).trim().toLowerCase() === String(studentName).trim().toLowerCase()) {
        targetRow = r + 1;
        break;
      }
    }
    if (targetRow === -1) {
      targetRow = sheet.getLastRow() + 1;
      sheet.getRange(targetRow, 2).setValue(studentName);
    }

    var crits = ['C01', 'C02', 'C03', 'C04', 'C05', 'C06', 'C07', 'C08', 'C09', 'C10', 'C11', 'C12'];
    var rowGrades = [];
    for (var i = 0; i < crits.length; i++) rowGrades.push(gradesObj[crits[i]] || 'E');
    /* Columns C-N (3-14): the 12 letter grades. Unchanged, already correct. */
    sheet.getRange(targetRow, 3, 1, 12).setValues([rowGrades]);

    if (feedbackObj) {
      /* Columns AE, AF, AG = 31, 32, 33. This was previously 19, 20, 21
         (S, T, U), which collided with pre-existing numeric-mark formulas
         for criteria C05-C07 and corrupted them. */
      if (feedbackObj.whatWentWell) {
        sheet.getRange(targetRow, 31).setRichTextValue(Utils.buildRichTextFromMarkdown(feedbackObj.whatWentWell));
      }
      if (feedbackObj.areasForImprovement) {
        sheet.getRange(targetRow, 32).setRichTextValue(Utils.buildRichTextFromMarkdown(feedbackObj.areasForImprovement));
      }
      if (feedbackObj.goalsForNextAssessment) {
        sheet.getRange(targetRow, 33).setRichTextValue(Utils.buildRichTextFromMarkdown(feedbackObj.goalsForNextAssessment));
      }
    }
    /* verifyMarkingSheetFormulas only touches columns O, P, Q, R (15-18),
       which are already correct on your sheet — safe to leave as-is. */
    verifyMarkingSheetFormulas(ss);
  }


  function getSheetDataAsObjects(sheetName) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var headers = data[0];
    var results = [];
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var obj = {};
      for (var c = 0; c < headers.length; c++) obj[headers[c]] = row[c];
      results.push(obj);
    }
    return results;
  }

  return {
    setupOrMigrateAssessmentSystem: setupOrMigrateAssessmentSystem,
    syncApprovedGradesToMarkingSheet: syncApprovedGradesToMarkingSheet,
    getSheetDataAsObjects: getSheetDataAsObjects,
    verifyMarkingSheetFormulas: verifyMarkingSheetFormulas
  };
})();

/* ============================================================================
 * SECTION 4: ASSESSMENT ENGINE & DETERMINISTIC INDEPENDENT GRADES
 * ============================================================================ */
var AssessmentService = (function() {

  function deriveGradeFromChecks(criterionId, checksMap) {
    var defs = Config.CRITERIA_DEFINITIONS;
    var crit = null;
    for (var i = 0; i < defs.length; i++) {
      if (defs[i].criterionId === criterionId) { crit = defs[i]; break; }
    }
    var W = GradeScaleService.getGradeWeights();

    if (!crit) {
      return { grade: 'E', score: W.E, explanation: 'Criterion not configured.', incompleteMinimumEvidence: true, bandBreakdown: {} };
    }

    checksMap = checksMap || {};

    if (checksMap[crit.mdOverride.id] === true || checksMap[crit.mdOverride.id] === 'true') {
      return { grade: 'A', score: W.A, explanation: 'Distinct A evidence selected \u2014 full marks awarded.', incompleteMinimumEvidence: false, distinctOverride: true, bandBreakdown: {} };
    }

    var bandCalcs = {};
    var bandNames = ['A', 'B', 'C', 'D', 'E'];
    for (var b = 0; b < bandNames.length; b++) {
      var bandName = bandNames[b];
      var items = crit.bands[bandName] || [];
      var ticked = 0;
      for (var k = 0; k < items.length; k++) {
        if (checksMap[items[k][0]] === true || checksMap[items[k][0]] === 'true') ticked++;
      }
      bandCalcs[bandName] = { ticked: ticked, total: items.length, fraction: items.length > 0 ? (ticked / items.length) : 0 };
    }

    var score = W.E;
    score += bandCalcs.D.fraction * (W.D - W.E);
    score += bandCalcs.C.fraction * (W.C - W.D);
    score += bandCalcs.B.fraction * (W.B - W.C);
    score += bandCalcs.A.fraction * (W.A - W.B);
    if (score > W.A) score = W.A;
    if (score < W.E) score = W.E;

    var grade = GradeScaleService.nearestGradeLetter(score, W);

    var totalTicked = bandCalcs.A.ticked + bandCalcs.B.ticked + bandCalcs.C.ticked + bandCalcs.D.ticked + bandCalcs.E.ticked;
    var incompleteMinimumEvidence = (totalTicked === 0);

    var explanation = 'Weighted score: ' + Math.round(score * 100) + '% (A ' + bandCalcs.A.ticked + '/' + bandCalcs.A.total +
      ', B ' + bandCalcs.B.ticked + '/' + bandCalcs.B.total +
      ', C ' + bandCalcs.C.ticked + '/' + bandCalcs.C.total +
      ', D ' + bandCalcs.D.ticked + '/' + bandCalcs.D.total +
      ', E ' + bandCalcs.E.ticked + '/' + bandCalcs.E.total + ') \u2192 nominal grade ' + grade + '.';

    return { grade: grade, score: score, explanation: explanation, incompleteMinimumEvidence: incompleteMinimumEvidence, distinctOverride: false, bandBreakdown: bandCalcs };
  }

  function getAssessmentStudioBootstrapData(targetSubId) {
    try {
      var submissions = Sheets.getSheetDataAsObjects(Config.SHEET_SUBMISSIONS);
      var configs = Sheets.getSheetDataAsObjects(Config.SHEET_CLASSROOM_CONFIG);
      var activeConfig = null;
      for (var c = 0; c < configs.length; c++) {
        if (configs[c].Active === true || configs[c].Active === 'true' || configs[c].Active === 'TRUE') {
          activeConfig = configs[c];
          break;
        }
      }
      var selectedId = targetSubId || (submissions.length > 0 ? submissions[0].SubmissionRecordID : null);
      var subDetail = selectedId ? getSubmissionForMarking(selectedId) : null;
      return {
        criteria: Config.CRITERIA_DEFINITIONS,
        submissions: submissions,
        activeConfig: activeConfig,
        selectedSubmissionId: selectedId,
        submissionDetail: subDetail,
        gradeWeights: GradeScaleService.getGradeWeights(),
        overallGradeBands: GradeScaleService.getOverallGradeBands()
      };
    } catch (err) {
      Logging.logError('getAssessmentStudioBootstrapData', err);
      return {
        criteria: Config.CRITERIA_DEFINITIONS, submissions: [], activeConfig: null, selectedSubmissionId: null, submissionDetail: null, error: err.message,
        gradeWeights: GradeScaleService.getGradeWeights(),
        overallGradeBands: GradeScaleService.getOverallGradeBands()
      };
    }
  }


  function getSubmissionForMarking(submissionRecordId) {
    var subRows = Sheets.getSheetDataAsObjects(Config.SHEET_SUBMISSIONS);
    var submission = null;
    for (var s = 0; s < subRows.length; s++) {
      if (subRows[s].SubmissionRecordID === submissionRecordId) {
        submission = subRows[s];
        break;
      }
    }
    if (!submission) throw new Error('Submission record ' + submissionRecordId + ' not found.');

    var allFiles = Sheets.getSheetDataAsObjects(Config.SHEET_SUBMISSION_FILES);
    var files = [];
    for (var f = 0; f < allFiles.length; f++) {
      if (allFiles[f].SubmissionRecordID === submissionRecordId) files.push(allFiles[f]);
    }

    var allAssess = Sheets.getSheetDataAsObjects(Config.SHEET_CRITERION_ASSESSMENTS);
    var criteriaMap = {}, assessmentId = null;
    for (var a = 0; a < allAssess.length; a++) {
      if (allAssess[a].SubmissionRecordID === submissionRecordId) {
        var rec = allAssess[a];
        assessmentId = rec.AssessmentID;
        criteriaMap[rec.CriterionID] = {
          assessmentId: rec.AssessmentID, criterionId: rec.CriterionID,
          aiProposedGrade: rec.AIProposedGrade || '', derivedGrade: rec.DerivedGrade || 'E',
          finalApprovedGrade: rec.FinalApprovedGrade || '',
          distinctOverrideSelected: (rec.DistinctOverrideSelected === true || rec.DistinctOverrideSelected === 'TRUE'),
          checkboxes: Utils.safeJsonParse(rec.CheckboxesJSON, {}),
          aiEvidence: Utils.safeJsonParse(rec.AIEvidenceJSON, []),
          teacherAdjusted: (rec.TeacherAdjusted === true || rec.TeacherAdjusted === 'TRUE'),
          teacherAudioTranscript: rec.TeacherAudioTranscript || '', teacherWrittenNote: rec.TeacherWrittenNote || '',
          confidence: rec.Confidence || 0, minimumEvidenceIncomplete: (rec.MinimumEvidenceIncomplete === true || rec.MinimumEvidenceIncomplete === 'TRUE'),
          status: rec.Status || Config.STATUS.NEW
        };
      }
    }
    if (!assessmentId) assessmentId = Utils.generateUuid();

    var defs = Config.CRITERIA_DEFINITIONS;
    for (var c = 0; c < defs.length; c++) {
      var cid = defs[c].criterionId;
      if (!criteriaMap[cid]) {
        criteriaMap[cid] = { assessmentId: assessmentId, criterionId: cid, aiProposedGrade: '', derivedGrade: 'E', finalApprovedGrade: '', distinctOverrideSelected: false, checkboxes: {}, aiEvidence: [], teacherAdjusted: false, teacherAudioTranscript: '', teacherWrittenNote: '', confidence: 0, minimumEvidenceIncomplete: true, status: Config.STATUS.NEW };
      }
      var calc = deriveGradeFromChecks(cid, criteriaMap[cid].checkboxes);
      criteriaMap[cid].explanation = calc.explanation;
      criteriaMap[cid].bandCalculations = calc.bandCalculations;
      criteriaMap[cid].derivedGrade = calc.grade;
      criteriaMap[cid].minimumEvidenceIncomplete = calc.incompleteMinimumEvidence;
    }

    var allAiRuns = Sheets.getSheetDataAsObjects(Config.SHEET_AI_ASSESSMENTS);
    var latestFeedback = null;
    for (var r = allAiRuns.length - 1; r >= 0; r--) {
      if (allAiRuns[r].SubmissionRecordID === submissionRecordId && allAiRuns[r].RunType === 'FeedbackGeneration') {
        var parsed = Utils.safeJsonParse(allAiRuns[r].ResponseJSON, null);
        if (parsed) { latestFeedback = parsed; break; }
      }
    }
    return { submission: submission, files: files, assessmentId: assessmentId, criteriaMap: criteriaMap, feedback: latestFeedback || { whatWentWell: '', areasForImprovement: '', goalsForNextAssessment: '' } };
  }

  function saveCriterionDraft(assessmentId, submissionRecordId, criterionId, state) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(Config.SHEET_CRITERION_ASSESSMENTS);
    var checksMap = state.checkboxes || {};
    var calc = deriveGradeFromChecks(criterionId, checksMap);

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var r = 1; r < data.length; r++) {
      if (data[r][0] === assessmentId && data[r][2] === criterionId) { rowIndex = r + 1; break; }
    }
    var now = Utils.formatDate(new Date());
    var rowValues = [
      assessmentId, submissionRecordId, criterionId, state.aiProposedGrade || '', calc.grade, calc.grade,
      !!checksMap[criterionId + '-MD'], Utils.safeJsonStringify(checksMap), Utils.safeJsonStringify(state.aiEvidence || []),
      !!state.teacherAdjusted, state.teacherAudioTranscript || '', state.teacherWrittenNote || '', state.confidence || 1.0,
      calc.incompleteMinimumEvidence, Config.STATUS.IN_REVIEW, now, now, '', ''
    ];

    if (rowIndex !== -1) sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    else sheet.appendRow(rowValues);

    Logging.logHistory(submissionRecordId, assessmentId, 'DraftSaved', criterionId, {}, state, 'Updated draft for ' + criterionId);
    return { success: true, derivedGrade: calc.grade, explanation: calc.explanation, minimumEvidenceIncomplete: calc.incompleteMinimumEvidence, bandCalculations: calc.bandCalculations };
  }

  function saveAssessmentDraft(assessmentId, submissionRecordId, criteriaMap, feedbackObj) {
    var lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (e) { return { success: false, message: 'Could not obtain lock.' }; }
    try {
      for (var cid in criteriaMap) saveCriterionDraft(assessmentId, submissionRecordId, cid, criteriaMap[cid]);
      if (feedbackObj) {
        var aiSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Config.SHEET_AI_ASSESSMENTS);
        if (aiSheet) {
          aiSheet.appendRow([assessmentId, submissionRecordId, 'FeedbackGeneration', Utils.formatDate(new Date()), Config.PRIMARY_MODEL, 'TeacherDraft', '[]', 'Teacher edited feedback', Utils.safeJsonStringify(feedbackObj), 'Success', '', 1.0, Utils.getSafeUserEmail(), 'Accepted']);
        }
      }
      updateSubmissionStatus(submissionRecordId, Config.STATUS.TEACHER_REVIEWED);
      return { success: true, message: 'Draft assessment saved successfully.' };
    } catch (err) {
      Logging.logError('saveAssessmentDraft', err);
      return { success: false, message: 'Failed to save: ' + Utils.sanitizeError(err) };
    } finally { lock.releaseLock(); }
  }

  function approveAssessment(assessmentId, submissionRecordId, feedbackObj) {
    var detail = getSubmissionForMarking(submissionRecordId);
    var defs = Config.CRITERIA_DEFINITIONS;
    var grades = {};
    for (var i = 0; i < defs.length; i++) {
      var cid = defs[i].criterionId;
      var cd = detail.criteriaMap[cid];
      if (!cd || !cd.derivedGrade) return { success: false, message: 'Cannot approve: missing grade on ' + cid };
      grades[cid] = cd.derivedGrade;
    }

    var now = Utils.formatDate(new Date());
    var userEmail = Utils.getSafeUserEmail();
    Sheets.syncApprovedGradesToMarkingSheet(detail.submission.StudentName, grades, feedbackObj);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cSheet = ss.getSheetByName(Config.SHEET_CRITERION_ASSESSMENTS);
    if (cSheet) {
      var cData = cSheet.getDataRange().getValues();
      for (var r = 1; r < cData.length; r++) {
        if (cData[r][0] === assessmentId) {
          cSheet.getRange(r + 1, 6).setValue(cData[r][4]);
          cSheet.getRange(r + 1, 15).setValue(Config.STATUS.APPROVED);
          cSheet.getRange(r + 1, 18).setValue(now);
          cSheet.getRange(r + 1, 19).setValue(userEmail);
        }
      }
    }

    var sSheet = ss.getSheetByName(Config.SHEET_SUBMISSIONS);
    if (sSheet) {
      var sData = sSheet.getDataRange().getValues();
      for (var sr = 1; sr < sData.length; sr++) {
        if (sData[sr][0] === submissionRecordId) {
          sSheet.getRange(sr + 1, 20).setValue(Config.STATUS.APPROVED);
          sSheet.getRange(sr + 1, 24).setValue(now);
          sSheet.getRange(sr + 1, 25).setValue(userEmail);
          break;
        }
      }
    }
    Logging.logHistory(submissionRecordId, assessmentId, 'Approved', '', {}, grades, 'Approved for ' + detail.submission.StudentName);
    return { success: true, message: 'Assessment approved. Letter grades (fluidly derived from Setup grade scale) synced to ' + Config.SHEET_MARKING + '.' };
  }

  function lockAssessment(assessmentId, submissionRecordId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sSheet = ss.getSheetByName(Config.SHEET_SUBMISSIONS);
    var now = Utils.formatDate(new Date());
    var user = Utils.getSafeUserEmail();
    var sData = sSheet.getDataRange().getValues();
    var row = -1;
    for (var r = 1; r < sData.length; r++) {
      if (sData[r][0] === submissionRecordId) { row = r + 1; break; }
    }
    if (row === -1) return { success: false, message: 'Submission not found.' };
    if (sSheet.getRange(row, 20).getValue() !== Config.STATUS.APPROVED) return { success: false, message: 'Assessment must be approved before locking.' };

    sSheet.getRange(row, 20).setValue(Config.STATUS.LOCKED);
    sSheet.getRange(row, 22).setValue(now);
    sSheet.getRange(row, 23).setValue(user);

    var cSheet = ss.getSheetByName(Config.SHEET_CRITERION_ASSESSMENTS);
    if (cSheet) {
      var cData = cSheet.getDataRange().getValues();
      for (var cr = 1; cr < cData.length; cr++) {
        if (cData[cr][0] === assessmentId) cSheet.getRange(cr + 1, 15).setValue(Config.STATUS.LOCKED);
      }
    }
    Logging.logHistory(submissionRecordId, assessmentId, 'Locked', '', {}, {}, 'Locked by ' + user);
    return { success: true, message: 'Assessment permanently locked.' };
  }

  function createReassessmentFromSubmission(submissionRecordId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sSheet = ss.getSheetByName(Config.SHEET_SUBMISSIONS);
    var rows = Sheets.getSheetDataAsObjects(Config.SHEET_SUBMISSIONS);
    var parent = null;
    for (var p = 0; p < rows.length; p++) {
      if (rows[p].SubmissionRecordID === submissionRecordId) { parent = rows[p]; break; }
    }
    if (!parent) return { success: false, message: 'Parent submission not found.' };

    var newSubId = Utils.generateUuid();
    var newVer = (parseInt(parent.SubmissionVersion, 10) || 1) + 1;
    var now = Utils.formatDate(new Date());

    sSheet.appendRow([
      newSubId, parent.ClassroomCourseID, parent.ClassroomCourseWorkID, parent.ClassroomSubmissionID,
      parent.StudentUserID, parent.StudentName, parent.StudentEmail, parent.Class, parent.Task,
      newVer, 'Reassessment', parent.ClassroomState, parent.TurnedInTime, now, parent.Late,
      parent.AttachmentSummary, parent.AttachmentFileIDsJSON, parent.AttachmentMetadataJSON,
      parent.DriveFolderID, Config.STATUS.NEW, submissionRecordId, true, '', '', '', '', '', '', '', 'Reassessment created'
    ]);

    var allFiles = Sheets.getSheetDataAsObjects(Config.SHEET_SUBMISSION_FILES);
    var fSheet = ss.getSheetByName(Config.SHEET_SUBMISSION_FILES);
    if (fSheet) {
      for (var f = 0; f < allFiles.length; f++) {
        if (allFiles[f].SubmissionRecordID === submissionRecordId) {
          var fo = allFiles[f];
          fSheet.appendRow([newSubId, Utils.generateUuid(), fo.SourceType, fo.DriveFileID, fo.FileName, fo.MimeType, fo.AlternateLink, fo.ThumbnailUrl, fo.FileSize, fo.EligibleForAI, 'Pending', '', '', now]);
        }
      }
    }
    Logging.logHistory(newSubId, '', 'ReassessmentCreated', '', {}, { parentId: submissionRecordId, version: newVer }, 'Reassessment v' + newVer);
    return { success: true, message: 'Reassessment version ' + newVer + ' created successfully.', newSubmissionRecordId: newSubId };
  }

  function updateSubmissionStatus(submissionRecordId, status) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sSheet = ss.getSheetByName(Config.SHEET_SUBMISSIONS);
    if (!sSheet) return;
    var data = sSheet.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      if (data[r][0] === submissionRecordId) { sSheet.getRange(r + 1, 20).setValue(status); break; }
    }
  }

  return {
    deriveGradeFromChecks: deriveGradeFromChecks, getAssessmentStudioBootstrapData: getAssessmentStudioBootstrapData,
    getSubmissionForMarking: getSubmissionForMarking, saveCriterionDraft: saveCriterionDraft,
    saveAssessmentDraft: saveAssessmentDraft, approveAssessment: approveAssessment,
    lockAssessment: lockAssessment, createReassessmentFromSubmission: createReassessmentFromSubmission,
    updateSubmissionStatus: updateSubmissionStatus
  };
})();

/* ============================================================================
 * SECTION 5: GOOGLE DRIVE, ATTACHMENTS & PDF EXPORTS
 * ============================================================================ */
var DriveService = (function() {
  var SUPPORTED_IMAGE_MIMES = ['image/jpeg', 'image/jpg', 'image/png', 'image/webp', 'image/gif', 'image/heic'];
  var SUPPORTED_DOC_MIMES = ['application/vnd.google-apps.document', 'application/vnd.google-apps.presentation', 'application/vnd.google-apps.drawing', 'application/pdf'];

  function isEligibleForAi(mimeType) {
    if (!mimeType) return false;
    mimeType = mimeType.toLowerCase();
    return SUPPORTED_IMAGE_MIMES.indexOf(mimeType) !== -1 || SUPPORTED_DOC_MIMES.indexOf(mimeType) !== -1;
  }

  /* Reads ONE file from Drive and returns it prepared for Gemini, or a
     status explaining why it could not be used. This is the function that
     genuinely touches Drive — call it once per file, one at a time. */
  function prepareSingleFileForAi(fileObj) {
    var fileId = fileObj.DriveFileID;
    var fileName = fileObj.FileName || 'Attachment';
    var mime = (fileObj.MimeType || '').toLowerCase();

    if (!isEligibleForAi(mime)) {
      return { fileId: fileId, fileName: fileName, status: 'unsupported' };
    }
    try {
      var file = DriveApp.getFileById(fileId);
      var blob = (mime.indexOf('google-apps') !== -1) ? file.getAs('application/pdf') : file.getBlob();
      var bytes = blob.getBytes();
      if (bytes.length / (1024 * 1024) > 15) {
        return { fileId: fileId, fileName: fileName, status: 'partial', limitations: ['Exceeds 15MB limit'] };
      }
      var effectiveMime = (mime.indexOf('google-apps') !== -1) ? 'application/pdf' : (blob.getContentType() || mime);
      return {
        fileId: fileId,
        fileName: fileName,
        status: 'reviewed',
        mimeType: effectiveMime,
        inlineData: { mimeType: effectiveMime, data: Utilities.base64Encode(bytes) }
      };
    } catch (err) {
      return { fileId: fileId, fileName: fileName, status: 'unreadable' };
    }
  }

  /* Bulk variant retained for the separate Class-Wide AI Grading Runner,
     which still processes one whole student at a time (already genuinely
     sequential at the student level) and does not need per-file granularity. */
  function prepareSubmissionFilesForAi(submissionRecordId) {
    var allFiles = Sheets.getSheetDataAsObjects(Config.SHEET_SUBMISSION_FILES);
    var targetFiles = [];
    for (var a = 0; a < allFiles.length; a++) {
      if (allFiles[a].SubmissionRecordID === submissionRecordId) targetFiles.push(allFiles[a]);
    }
    var prepared = [], metadata = [];
    for (var j = 0; j < targetFiles.length; j++) {
      var result = prepareSingleFileForAi(targetFiles[j]);
      metadata.push({ fileId: result.fileId, fileName: result.fileName, status: result.status, limitations: result.limitations });
      if (result.status === 'reviewed') {
        prepared.push({ fileId: result.fileId, fileName: result.fileName, mimeType: result.mimeType, inlineData: result.inlineData });
      }
    }
    return { files: prepared, metadata: metadata };
  }

  function exportDriveFileAsPdf(driveFileId) {
    try {
      var file = DriveApp.getFileById(driveFileId);
      var mime = file.getMimeType();
      var blob;
      if (mime.indexOf('google-apps') !== -1) {
        blob = file.getAs('application/pdf');
      } else if (mime === 'application/pdf') {
        blob = file.getBlob();
      } else {
        blob = file.getBlob();
      }
      var cleanName = file.getName().replace(/\.[^/.]+$/, '') + '.pdf';
      var base64 = Utilities.base64Encode(blob.getBytes());
      return {
        success: true,
        fileName: cleanName,
        dataUrl: 'data:application/pdf;base64,' + base64
      };
    } catch (e) {
      return { success: false, message: 'Could not export Drive file: ' + Utils.sanitizeError(e) };
    }
  }

  return {
    isEligibleForAi: isEligibleForAi,
    prepareSingleFileForAi: prepareSingleFileForAi,
    prepareSubmissionFilesForAi: prepareSubmissionFilesForAi,
    exportDriveFileAsPdf: exportDriveFileAsPdf
  };
})();

/* ============================================================================
 * SECTION 6: GOOGLE CLASSROOM INTEGRATION
 * FIXED: importOrRefreshClassroomSubmissions no longer passes courseWorkStates
 * (invalid for studentSubmissions.list) and now paginates through all results.
 * ============================================================================ */
var ClassroomService = (function() {
  function listTeacherCourses() {
    try {
      var res = Classroom.Courses.list({ teacherId: 'me', courseStates: ['ACTIVE'], pageSize: 50 });
      var courses = [];
      var raw = res.courses || [];
      for (var i = 0; i < raw.length; i++) {
        courses.push({ id: raw[i].id, name: raw[i].name, section: raw[i].section || '', room: raw[i].room || '' });
      }
      return { success: true, courses: courses };
    } catch (err) {
      Logging.logError('listTeacherCourses', err);
      return { success: false, message: 'Classroom courses error: ' + Utils.sanitizeError(err) };
    }
  }

  function listCourseWork(courseId) {
    try {
      var res = Classroom.Courses.CourseWork.list(courseId, { courseWorkStates: ['PUBLISHED'], pageSize: 50 });
      var cwList = [];
      var raw = res.courseWork || [];
      for (var i = 0; i < raw.length; i++) {
        var cw = raw[i];
        var due = cw.dueDate ? (cw.dueDate.year + '-' + (cw.dueDate.month < 10 ? '0' : '') + cw.dueDate.month + '-' + (cw.dueDate.day < 10 ? '0' : '') + cw.dueDate.day) : '';
        cwList.push({ id: cw.id, title: cw.title, maxPoints: cw.maxPoints || 100, dueDate: due, state: cw.state });
      }
      return { success: true, coursework: cwList };
    } catch (err) {
      Logging.logError('listCourseWork', err);
      return { success: false, message: 'CourseWork error: ' + Utils.sanitizeError(err) };
    }
  }

  function saveClassroomSelection(courseId, courseWorkId) {
    try {
      var course = Classroom.Courses.get(courseId);
      var cw = Classroom.Courses.CourseWork.get(courseId, courseWorkId);
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(Config.SHEET_CLASSROOM_CONFIG);
      if (!sheet) { Sheets.setupOrMigrateAssessmentSystem(); sheet = ss.getSheetByName(Config.SHEET_CLASSROOM_CONFIG); }

      var data = sheet.getDataRange().getValues();
      for (var r = 1; r < data.length; r++) sheet.getRange(r + 1, 10).setValue(false);

      sheet.appendRow([Utils.generateUuid(), courseId, course.name || '', course.section || '', courseWorkId, cw.title || '', cw.maxPoints || 100, Utils.formatDate(new Date()), Utils.getSafeUserEmail(), true]);
      return { success: true, message: 'Saved assignment: ' + cw.title };
    } catch (err) {
      Logging.logError('saveClassroomSelection', err);
      return { success: false, message: 'Save selection error: ' + Utils.sanitizeError(err) };
    }
  }

  function getActiveClassroomConfig() {
    var configs = Sheets.getSheetDataAsObjects(Config.SHEET_CLASSROOM_CONFIG);
    for (var i = 0; i < configs.length; i++) {
      if (configs[i].Active === true || configs[i].Active === 'TRUE' || configs[i].Active === 'true') return configs[i];
    }
    return null;
  }

  /* Paginates through ALL student submissions instead of a single 100-item page. */
  function listAllStudentSubmissions(courseId, courseWorkId) {
    var all = [];
    var pageToken = null;
    do {
      var params = { pageSize: 100 };
      if (pageToken) params.pageToken = pageToken;
      /* IMPORTANT: courseWorkStates is NOT a valid parameter for studentSubmissions.list */
      var response = Classroom.Courses.CourseWork.StudentSubmissions.list(String(courseId), String(courseWorkId), params);
      if (response && response.studentSubmissions) all = all.concat(response.studentSubmissions);
      pageToken = (response && response.nextPageToken) ? response.nextPageToken : null;
    } while (pageToken);
    return all;
  }

  function importOrRefreshClassroomSubmissions() {
    var cfg = getActiveClassroomConfig();
    if (!cfg) return { success: false, message: 'No active Classroom assignment configured. Please select one via Classroom menu.' };
    try {
      var submissions = listAllStudentSubmissions(cfg.CourseID, cfg.CourseWorkID);
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sSheet = ss.getSheetByName(Config.SHEET_SUBMISSIONS);
      var fSheet = ss.getSheetByName(Config.SHEET_SUBMISSION_FILES);

      if (!sSheet || !fSheet) {
        Sheets.setupOrMigrateAssessmentSystem();
        sSheet = ss.getSheetByName(Config.SHEET_SUBMISSIONS);
        fSheet = ss.getSheetByName(Config.SHEET_SUBMISSION_FILES);
      }
      if (!sSheet) return { success: false, message: 'Required sheet "' + Config.SHEET_SUBMISSIONS + '" could not be created or found.' };

      var existing = Sheets.getSheetDataAsObjects(Config.SHEET_SUBMISSIONS);
      var subMap = {};
      for (var e = 0; e < existing.length; e++) {
        var row = existing[e];
        if (!row.StudentUserID) continue;
        if (!subMap[row.StudentUserID]) subMap[row.StudentUserID] = [];
        subMap[row.StudentUserID].push(row);
      }

      var stats = { total: submissions.length, newCount: 0, updated: 0, newVersion: 0, skipped: 0 };
      var now = Utils.formatDate(new Date());

      for (var i = 0; i < submissions.length; i++) {
        var cSub = submissions[i];
        var uid = cSub.userId;
        var sName = 'Student ' + (i + 1);
        var sEmail = '';

        try {
          var p = Classroom.UserProfiles.get(uid);
          if (p && p.name && p.name.fullName) sName = p.name.fullName;
          if (p && p.emailAddress) sEmail = p.emailAddress;
        } catch (pe) {
          sName = 'Student ' + (i + 1);
        }

        var attachments = [], fileIds = [];
        if (cSub.assignmentSubmission && cSub.assignmentSubmission.attachments) {
          var rawAtts = cSub.assignmentSubmission.attachments;
          for (var a = 0; a < rawAtts.length; a++) {
            var att = rawAtts[a];
            if (att.driveFile) {
              attachments.push({ sourceType: 'driveFile', id: att.driveFile.id, title: att.driveFile.title, alternateLink: att.driveFile.alternateLink, thumbnailUrl: att.driveFile.thumbnailUrl || '' });
              fileIds.push(att.driveFile.id);
            }
          }
        }

        var userHistory = subMap[uid] || [];
        userHistory.sort(function(a, b) { return (parseInt(b.SubmissionVersion, 10) || 1) - (parseInt(a.SubmissionVersion, 10) || 1); });
        var latest = userHistory.length > 0 ? userHistory[0] : null;

        if (!latest) {
          var newId = Utils.generateUuid();
          sSheet.appendRow([
            newId, cfg.CourseID, cfg.CourseWorkID, cSub.id, uid, sName, sEmail, cfg.CourseName || '9DAT1', cfg.AssignmentTitle || 'Task 2: Jewellery Design',
            1, 'Draft', cSub.state || 'NEW', cSub.creationTime || now, cSub.updateTime || now, cSub.late || false,
            attachments.length + ' attachment(s)', Utils.safeJsonStringify(fileIds), Utils.safeJsonStringify(attachments),
            '', Config.STATUS.NEW, '', true, '', '', '', '', cSub.assignedGrade || '', '', '', 'Initial import'
          ]);
          insertFiles(fSheet, newId, attachments);
          stats.newCount++;
        } else {
          var prevIds = Utils.safeJsonParse(latest.AttachmentFileIDsJSON, []);
          var changed = (JSON.stringify(prevIds.sort()) !== JSON.stringify(fileIds.sort())) || (cSub.updateTime && cSub.updateTime !== latest.UpdateTime);
          if (changed) {
            if (latest.Status === Config.STATUS.LOCKED || latest.Status === Config.STATUS.APPROVED) {
              var nVer = (parseInt(latest.SubmissionVersion, 10) || 1) + 1;
              var vId = Utils.generateUuid();
              sSheet.appendRow([
                vId, cfg.CourseID, cfg.CourseWorkID, cSub.id, uid, sName, sEmail, cfg.CourseName || '9DAT1', cfg.AssignmentTitle || 'Task 2: Jewellery Design',
                nVer, 'Reassessment', cSub.state || 'UPDATED', cSub.creationTime || now, cSub.updateTime || now, cSub.late || false,
                attachments.length + ' attachment(s)', Utils.safeJsonStringify(fileIds), Utils.safeJsonStringify(attachments),
                '', Config.STATUS.NEW, latest.SubmissionRecordID, true, '', '', '', '', cSub.assignedGrade || '', '', '', 'Classroom update reassessment'
              ]);
              insertFiles(fSheet, vId, attachments);
              stats.newVersion++;
            } else {
              updateSubRow(sSheet, latest.SubmissionRecordID, cSub, attachments, fileIds);
              stats.updated++;
            }
          } else {
            stats.skipped++;
          }
        }
      }
      return {
        success: true,
        message: 'Classroom sync completed: ' + stats.newCount + ' new, ' + stats.updated + ' updated, ' + stats.newVersion + ' new versions, ' + stats.skipped + ' unchanged.',
        stats: stats
      };
    } catch (err) {
      Logging.logError('importOrRefreshClassroomSubmissions', err);
      return { success: false, message: 'Import failed: ' + Utils.sanitizeError(err) };
    }
  }

  function insertFiles(fSheet, subId, attachments) {
    if (!fSheet || !attachments.length) return;
    var rows = [], now = Utils.formatDate(new Date());
    for (var i = 0; i < attachments.length; i++) {
      var att = attachments[i], mime = '', size = 0;
      if (att.id) {
        try { var df = DriveApp.getFileById(att.id); mime = df.getMimeType(); size = df.getSize(); } catch (e) {}
      }
      var eligible = DriveService.isEligibleForAi(mime);
      rows.push([subId, Utils.generateUuid(), att.sourceType, att.id || '', att.title, mime, att.alternateLink || '', att.thumbnailUrl || '', size, eligible, eligible ? 'Pending' : 'Unsupported', '', '', now]);
    }
    if (rows.length) fSheet.getRange(fSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  function updateSubRow(sSheet, recordId, cSub, attachments, fileIds) {
    var data = sSheet.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      if (data[r][0] === recordId) {
        sSheet.getRange(r + 1, 12).setValue(cSub.state || 'UPDATED');
        sSheet.getRange(r + 1, 14).setValue(cSub.updateTime || Utils.formatDate(new Date()));
        sSheet.getRange(r + 1, 16).setValue(attachments.length + ' attachment(s)');
        sSheet.getRange(r + 1, 17).setValue(Utils.safeJsonStringify(fileIds));
        sSheet.getRange(r + 1, 18).setValue(Utils.safeJsonStringify(attachments));
        break;
      }
    }
  }

  function syncAssessmentToClassroom(assessmentId, submissionRecordId) {
    var detail = AssessmentService.getSubmissionForMarking(submissionRecordId);
    var sub = detail.submission;
    if (sub.Status !== Config.STATUS.APPROVED && sub.Status !== Config.STATUS.LOCKED) {
      return { success: false, message: 'Only Approved or Locked assessments can be synced.' };
    }
    try {
      var cw = Classroom.Courses.CourseWork.get(sub.ClassroomCourseID, sub.ClassroomCourseWorkID);
      var maxPoints = cw.maxPoints || 100;
      var total = 0, defs = Config.CRITERIA_DEFINITIONS;
      for (var i = 0; i < defs.length; i++) {
        var g = detail.criteriaMap[defs[i].criterionId] ? detail.criteriaMap[defs[i].criterionId].derivedGrade : 'E';
        total += defs[i].maxMarks * (Config.GRADE_WEIGHTS[g] || 0.25);
      }
      var assignedGrade = (maxPoints === 100) ? total : Math.round((total / 100) * maxPoints * 10) / 10;
      Classroom.Courses.CourseWork.StudentSubmissions.patch({ assignedGrade: assignedGrade, draftGrade: assignedGrade }, sub.ClassroomCourseID, sub.ClassroomCourseWorkID, sub.ClassroomSubmissionID, { updateMask: 'assignedGrade,draftGrade' });

      var now = Utils.formatDate(new Date());
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sSheet = ss.getSheetByName(Config.SHEET_SUBMISSIONS);
      var sData = sSheet.getDataRange().getValues();
      for (var r = 1; r < sData.length; r++) {
        if (sData[r][0] === submissionRecordId) {
          sSheet.getRange(r + 1, 27).setValue(assignedGrade);
          sSheet.getRange(r + 1, 28).setValue(now);
          sSheet.getRange(r + 1, 29).setValue('Synced ' + assignedGrade + '/' + maxPoints);
          break;
        }
      }
      return { success: true, message: 'Synced grade ' + assignedGrade + ' / ' + maxPoints + ' to Classroom for ' + sub.StudentName + '.' };
    } catch (err) {
      Logging.logError('syncAssessmentToClassroom', err);
      return { success: false, message: 'Classroom sync failed: ' + Utils.sanitizeError(err) };
    }
  }

  return { listTeacherCourses: listTeacherCourses, listCourseWork: listCourseWork, saveClassroomSelection: saveClassroomSelection, getActiveClassroomConfig: getActiveClassroomConfig, importOrRefreshClassroomSubmissions: importOrRefreshClassroomSubmissions, syncAssessmentToClassroom: syncAssessmentToClassroom };
})();

/* ============================================================================
 * SECTION 7: HIGH-ACCURACY GEMINI ENGINE (INDEPENDENT BAND TICKING)
 * ============================================================================ */
var GeminiService = (function() {
  function getApiKey() { return PropertiesService.getScriptProperties().getProperty(Config.SCRIPT_PROP_GEMINI_KEY); }
  function setGeminiApiKey(key) {
    if (!key || key.trim().length < 10) return { success: false, message: 'Invalid API key.' };
    PropertiesService.getScriptProperties().setProperty(Config.SCRIPT_PROP_GEMINI_KEY, key.trim());
    return { success: true, message: 'Gemini API key stored securely.' };
  }

  function testGeminiConnection() {
    var key = getApiKey();
    if (!key) return { success: false, message: 'No Gemini key set in Script Properties.' };
    try {
      var res = callGeminiWithFallback([{ role: 'user', parts: [{ text: 'Ping' }] }], false);
      return res.success ? { success: true, message: 'Connected to Gemini successfully using model: ' + res.modelUsed } : { success: false, message: 'Gemini error: ' + res.message };
    } catch (e) { return { success: false, message: 'Connection error: ' + Utils.sanitizeError(e) }; }
  }

  function callGeminiWithFallback(contents, requireJson) {
    var key = getApiKey();
    if (!key) throw new Error('Gemini API key is not configured.');

    var modelsToTry = [Config.PRIMARY_MODEL].concat(Config.FALLBACK_MODELS);
    var lastError = null;

    for (var m = 0; m < modelsToTry.length; m++) {
      var model = modelsToTry[m];
      var url = 'https://generativelanguage.googleapis.com/' + Config.GEMINI_API_VERSION + '/models/' + model + ':generateContent?key=' + key;
      var payload = { contents: contents };
      if (requireJson) {
        payload.generationConfig = { responseMimeType: 'application/json', temperature: 0.2 };
      }

      for (var attempt = 1; attempt <= 3; attempt++) {
        try {
          var resp = UrlFetchApp.fetch(url, {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
          });

          var code = resp.getResponseCode();
          var text = resp.getContentText();

          if (code === 200) {
            var parsed = JSON.parse(text);
            var contentPart = parsed.candidates && parsed.candidates[0] && parsed.candidates[0].content && parsed.candidates[0].content.parts && parsed.candidates[0].content.parts[0] ? parsed.candidates[0].content.parts[0].text : '{}';
            return { success: true, text: contentPart, modelUsed: model };
          }

          if (code === 503 || code === 429) {
            lastError = 'Model ' + model + ' busy (code ' + code + '). Attempt ' + attempt + ' failed.';
            Utilities.sleep(attempt * 1500);
            continue;
          }

          lastError = 'Gemini error ' + code + ': ' + text;
          break;
        } catch (fetchErr) {
          lastError = fetchErr.message;
          Utilities.sleep(1500);
        }
      }
    }

    return { success: false, message: lastError || 'All models exhausted.' };
  }

  function runAiAssessmentWithPreparedFiles(submissionRecordId, preparedFiles) {
    if (!preparedFiles || !preparedFiles.length) {
      return { success: false, message: 'No eligible files available for AI analysis.' };
    }

    var rubricContext = Config.getSerializedRubricPrompt();

    var sysPrompt =
      "You are an experienced, encouraging secondary Technological and Applied Studies (TAS) teacher in NSW, Australia assessing Year 9 Design & Technology: Jewellery Design.\n" +
      "Assess student work against the 12 criteria (C01 to C12) strictly using the official rubric below.\n\n" +
      "=== OFFICIAL RUBRIC & OBSERVABLE DEFINITIONS ===\n" +
      rubricContext + "\n" +
      "================================================\n\n" +
      "NON-NEGOTIABLE ASSESSMENT RULES:\n" +
      "1. TICK EVERY OBSERVABLE THE EVIDENCE SUPPORTS, ACROSS ALL BANDS:\n" +
      "   - Evidence for a single criterion is often uneven. A student may satisfy an A-level observable while ALSO only satisfying a D-level observable elsewhere in the same criterion.\n" +
      "   - Tick EVERY individual observable statement across Bands A, B, C, D, and E wherever the submission provides genuine evidence for it.\n" +
      "2. IGNORE BLANK / UNUSED DUPLICATE TEMPLATE SLIDES.\n" +
      "3. VISION - HAND-DRAWN SKETCH DETECTION (C08): tick relevant observables if hand-drawn sketches with dimensions/labels are present.\n" +
      "4. VISION - PRACTICAL PHYSICAL MAKING PHOTOS (C11 & C12): tick relevant observables if photographed finished jewellery is present.\n\n" +
      "Respond ONLY in valid JSON matching this schema:\n" +
      "{\n" +
      "  \"criteria\": [\n" +
      "    {\n" +
      "      \"criterionId\": \"C01\",\n" +
      "      \"tickedCheckboxes\": [\"C01-A-1\", \"C01-D-4\"],\n" +
      "      \"evidenceNotes\": [\n" +
      "        {\"checkboxId\": \"C01-A-1\", \"fileName\": \"folio.pdf\", \"note\": \"Six success criteria stated on page 4-5.\"}\n" +
      "      ]\n" +
      "    }\n" +
      "  ]\n" +
      "}";

    var parts = [{ text: sysPrompt }];
    for (var p = 0; p < preparedFiles.length; p++) {
      parts.push({ inlineData: { mimeType: preparedFiles[p].mimeType, data: preparedFiles[p].inlineData.data } });
    }

    var geminiCall = callGeminiWithFallback([{ role: 'user', parts: parts }], true);
    if (!geminiCall.success) {
      return { success: false, message: 'Gemini service busy. Details: ' + geminiCall.message };
    }

    try {
      var parsed = JSON.parse(geminiCall.text);
      var aId = Utils.generateUuid();

      var cList = parsed.criteria || [];
      for (var j = 0; j < cList.length; j++) {
        var cRes = cList[j];
        var map = {};
        var tList = cRes.tickedCheckboxes || [];
        for (var t = 0; t < tList.length; t++) map[tList[t]] = true;
        var calc = AssessmentService.deriveGradeFromChecks(cRes.criterionId, map);
        AssessmentService.saveCriterionDraft(aId, submissionRecordId, cRes.criterionId, {
          aiProposedGrade: calc.grade,
          checkboxes: map,
          aiEvidence: cRes.evidenceNotes || [],
          confidence: 0.85
        });
      }

      var fbRes = generateTeacherReviewedFeedback(aId, submissionRecordId);

      return {
        success: true,
        message: 'AI assessment & feedback completed using ' + geminiCall.modelUsed + '.',
        assessmentId: aId,
        feedback: fbRes.success ? fbRes.feedback : null
      };
    } catch (err) {
      Logging.logError('runAiAssessmentWithPreparedFiles', err);
      return { success: false, message: 'AI response parsing error: ' + Utils.sanitizeError(err) };
    }
  }

  function runInitialAiAssessment(submissionRecordId) {
    var prep = DriveService.prepareSubmissionFilesForAi(submissionRecordId);
    if (!prep.files.length) return { success: false, message: 'No eligible files available for AI analysis.' };

    var rubricContext = Config.getSerializedRubricPrompt();

    var sysPrompt =
      "You are an experienced, encouraging secondary Technological and Applied Studies (TAS) teacher in NSW, Australia assessing Year 9 Design & Technology: Jewellery Design.\n" +
      "Assess student work against the 12 criteria (C01 to C12) strictly using the official rubric below.\n\n" +
      "=== OFFICIAL RUBRIC & OBSERVABLE DEFINITIONS ===\n" +
      rubricContext + "\n" +
      "================================================\n\n" +
      "NON-NEGOTIABLE ASSESSMENT RULES:\n" +
      "1. INDEPENDENT BAND TICKING (DO NOT TICK REDUNDANT LOWER BANDS):\n" +
      "   - Ticking the observables in Band B awards a Grade B directly without needing C, D, or E ticked.\n" +
      "   - Only tick the checkboxes for the specific band level the student demonstrated.\n" +
      "2. IGNORE BLANK / UNUSED DUPLICATE TEMPLATE SLIDES:\n" +
      "   - Judge the student's work strictly from the completed pages where content was actually provided.\n" +
      "3. VISION - HAND-DRAWN SKETCH DETECTION (C08):\n" +
      "   - Look for embedded photographs of hand-drawn sketches from physical sketchbooks. Do NOT award Band E for sketches if hand drawings are present!\n" +
      "4. VISION - PRACTICAL PHYSICAL MAKING PHOTOS (C11 & C12):\n" +
      "   - Inspect photographs of student-made physical jewellery. Do NOT default to Band E if physical items are documented!\n\n" +
      "Respond ONLY in valid JSON matching this schema:\n" +
      "{\n" +
      "  \"criteria\": [\n" +
      "    {\n" +
      "      \"criterionId\": \"C01\",\n" +
      "      \"suggestedBand\": \"B\",\n" +
      "      \"tickedCheckboxes\": [\"C01-B-1\", \"C01-B-2\", \"C01-B-3\", \"C01-B-4\"],\n" +
      "      \"evidenceNotes\": [\n" +
      "        {\"checkboxId\": \"C01-B-1\", \"fileName\": \"folio.pdf\", \"note\": \"5 criteria stated on page 4-5 with measurements.\"}\n" +
      "      ]\n" +
      "    }\n" +
      "  ]\n" +
      "}";

    var parts = [{ text: sysPrompt }];
    for (var p = 0; p < prep.files.length; p++) {
      parts.push({ inlineData: { mimeType: prep.files[p].mimeType, data: prep.files[p].inlineData.data } });
    }

    var geminiCall = callGeminiWithFallback([{ role: 'user', parts: parts }], true);
    if (!geminiCall.success) {
      return { success: false, message: 'Gemini service busy. Details: ' + geminiCall.message };
    }

    try {
      var parsed = JSON.parse(geminiCall.text);
      var aId = Utils.generateUuid();

      var cList = parsed.criteria || [];
      for (var j = 0; j < cList.length; j++) {
        var cRes = cList[j];
        var map = {};
        var tList = cRes.tickedCheckboxes || [];
        for (var t = 0; t < tList.length; t++) map[tList[t]] = true;
        var calc = AssessmentService.deriveGradeFromChecks(cRes.criterionId, map);
        AssessmentService.saveCriterionDraft(aId, submissionRecordId, cRes.criterionId, {
          aiProposedGrade: cRes.suggestedBand || calc.grade,
          checkboxes: map,
          aiEvidence: cRes.evidenceNotes || [],
          confidence: 0.85
        });
      }

      var fbRes = generateTeacherReviewedFeedback(aId, submissionRecordId);

      return {
        success: true,
        message: 'AI assessment & feedback completed using ' + geminiCall.modelUsed + '.',
        assessmentId: aId,
        feedback: fbRes.success ? fbRes.feedback : null
      };
    } catch (err) {
      Logging.logError('runInitialAiAssessment', err);
      return { success: false, message: 'AI response parsing error: ' + Utils.sanitizeError(err) };
    }
  }

  function generateTeacherReviewedFeedback(assessmentId, submissionRecordId) {
    var detail = AssessmentService.getSubmissionForMarking(submissionRecordId);
    var summary = [];
    Config.CRITERIA_DEFINITIONS.forEach(function(c) {
      var cd = detail.criteriaMap[c.criterionId] || {};
      summary.push(c.criterionId + ' (' + c.title + '): Grade ' + (cd.derivedGrade || 'E') + ' ' + (cd.teacherWrittenNote ? '| ' + cd.teacherWrittenNote : ''));
    });

    var prompt =
      "You are a supportive, knowledgeable Australian high school TAS teacher writing warm, natural, and constructive report comments for Year 9 Design & Technology student " + detail.submission.StudentName + ".\n\n" +
      "Below is the marking breakdown for their Jewellery Design project:\n" + summary.join('\n') + "\n\n" +
      "CRITICAL VOICE & TONE GUIDELINES:\n" +
      "- Speak naturally like a real teacher talking directly to a Year 9 student.\n" +
      "- Do NOT use robotic, corporate AI jargon (avoid 'commendable', 'testament to', 'pivotal', 'delve into', 'unprecedented').\n" +
      "- Reference real, specific materials and techniques from their portfolio (e.g. FIMO polymer clay, wire wrapping, silver or gold leaf, pliers, bead threading).\n" +
      "- Provide genuine praise for what worked, and clear, practical workshop advice on what to improve next time.\n\n" +
      "Return ONLY JSON matching this structure:\n" +
      "{\n" +
      "  \"whatWentWell\": \"...\",\n" +
      "  \"areasForImprovement\": \"...\",\n" +
      "  \"goalsForNextAssessment\": \"1. ...\\n2. ...\"\n" +
      "}";

    var res = callGeminiWithFallback([{ role: 'user', parts: [{ text: prompt }] }], true);
    if (!res.success) return { success: false, message: 'Feedback generation error: ' + res.message };

    try {
      var out = JSON.parse(res.text);
      var aiSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Config.SHEET_AI_ASSESSMENTS);
      if (aiSheet) {
        aiSheet.appendRow([
          assessmentId, submissionRecordId, 'FeedbackGeneration', Utils.formatDate(new Date()),
          Config.PRIMARY_MODEL, 'v1', '[]', 'Auto-generated feedback', res.text, 'Success', '', 1.0,
          Utils.getSafeUserEmail(), 'Accepted'
        ]);
      }
      return { success: true, feedback: out };
    } catch (err) {
      return { success: false, message: 'Invalid feedback JSON returned.' };
    }
  }

  return {
    setGeminiApiKey: setGeminiApiKey,
    testGeminiConnection: testGeminiConnection,
    runInitialAiAssessment: runInitialAiAssessment,
    runAiAssessmentWithPreparedFiles: runAiAssessmentWithPreparedFiles,
    generateTeacherReviewedFeedback: generateTeacherReviewedFeedback
  };
})();

/* ============================================================================
 * SECTION 8: EMBEDDED HTML TEMPLATES
 * REBUILT FROM SCRATCH — paste this entire block into Code.gs between the
 * end of Section 7 (GeminiService) and the start of Section 9 (onOpen).
 * This replaces the missing "var HtmlTemplates = (function() {...})();"
 * and resolves "ReferenceError: HtmlTemplates is not defined".
 * ============================================================================ */
var HtmlTemplates = (function() {

  function getAssessmentPdfReportHtml(detail) {
    var sub = detail.submission;
    var cMap = detail.criteriaMap;
    var fb = detail.feedback || {};
    var defs = Config.CRITERIA_DEFINITIONS;

    var partA = 0, partB = 0;
    var rowsHtml = '';

    for (var i = 0; i < defs.length; i++) {
      var c = defs[i];
      var cd = cMap[c.criterionId] || {};
      var grade = cd.derivedGrade || 'E';
      var weight = Config.GRADE_WEIGHTS[grade] || 0.25;
      var mark = (c.maxMarks * weight).toFixed(2);

      if (c.part === 'Part A') partA += parseFloat(mark);
      else partB += parseFloat(mark);

      var note = cd.teacherWrittenNote || cd.teacherAudioTranscript || cd.explanation || 'Thresholds verified.';

      rowsHtml += '<tr>' +
        '<td style="font-weight:bold;padding:6px;border:1px solid #cbd5e1;">' + c.criterionId + '</td>' +
        '<td style="padding:6px;border:1px solid #cbd5e1;"><b>' + c.title + '</b><br/><span style="font-size:10px;color:#64748b;">' + c.part + ' &bull; /' + c.maxMarks + ' Marks</span></td>' +
        '<td style="text-align:center;font-weight:bold;padding:6px;border:1px solid #cbd5e1;">' + grade + '</td>' +
        '<td style="text-align:center;padding:6px;border:1px solid #cbd5e1;">' + mark + ' / ' + c.maxMarks + '</td>' +
        '<td style="padding:6px;font-size:11px;border:1px solid #cbd5e1;">' + note + '</td>' +
      '</tr>';
    }

    var total = (partA + partB).toFixed(1);
    var overallGrade = GradeScaleService.getOverallGradeLetter(parseFloat(total), 
    GradeScaleService.getOverallGradeBands());

    return '<!DOCTYPE html><html><head><meta charset="utf-8">' +
      '<style>' +
      'body{font-family:Arial,Helvetica,sans-serif;margin:20px;color:#0f172a;line-height:1.3;font-size:12px;}' +
      'h1{margin:0;font-size:18px;color:#1e3a8a;}' +
      '.header-box{display:flex;justify-content:space-between;border-bottom:2px solid #1e3a8a;padding-bottom:10px;margin-bottom:14px;}' +
      '.score-card{background:#0f172a;color:#fff;padding:8px 14px;border-radius:6px;text-align:right;}' +
      '.details-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;background:#f8fafc;border:1px solid #cbd5e1;padding:8px 12px;border-radius:6px;margin-bottom:14px;font-size:11px;}' +
      'table{width:100%;border-collapse:collapse;margin-bottom:16px;}' +
      'th{background:#0f172a;color:#fff;padding:6px;font-size:11px;border:1px solid #0f172a;text-align:left;}' +
      '.fb-box{background:#f8fafc;border:1px solid #2563eb;border-radius:6px;padding:10px 14px;margin-top:10px;}' +
      '.fb-title{font-weight:bold;color:#1e3a8a;margin-bottom:4px;font-size:11px;}' +
      '</style></head><body>' +
      '<div class="header-box">' +
        '<div>' +
          '<h1>NSW STAGE 5 DESIGN &amp; TECHNOLOGY</h1>' +
          '<div style="font-size:13px;font-weight:bold;color:#2563eb;margin-top:2px;">Year 9 Jewellery Design &mdash; Assessment &amp; Folio Report</div>' +
        '</div>' +
        '<div class="score-card">' +
          '<div style="font-size:10px;color:#94a3b8;">OVERALL RESULT</div>' +
          '<div style="font-size:18px;font-weight:bold;color:#38bdf8;">' + total + ' / 100 (' + overallGrade + ')</div>' +
        '</div>' +
      '</div>' +
      '<div class="details-grid">' +
        '<div><b>Student:</b> ' + sub.StudentName + '</div>' +
        '<div><b>Class:</b> ' + (sub.Class || '9DAT1') + '</div>' +
        '<div><b>Submission Version:</b> v' + sub.SubmissionVersion + '</div>' +
        '<div><b>Part A (Proposal &amp; Realisation):</b> ' + partA.toFixed(2) + ' / 55</div>' +
        '<div><b>Part B (Creativity &amp; Making):</b> ' + partB.toFixed(2) + ' / 45</div>' +
        '<div><b>Assessment Date:</b> ' + Utils.formatDate(new Date()).split('T')[0] + '</div>' +
      '</div>' +
      '<table>' +
        '<thead><tr><th>Code</th><th>Criterion Title &amp; Component</th><th style="text-align:center;">Band</th><th style="text-align:center;">Mark</th><th>Observables &amp; Evidence Audit</th></tr></thead>' +
        '<tbody>' + rowsHtml + '</tbody>' +
      '</table>' +
      '<div class="fb-box">' +
        '<div style="font-size:13px;font-weight:bold;color:#0f172a;margin-bottom:8px;border-bottom:1px solid #cbd5e1;padding-bottom:4px;">Teacher Feedback &amp; Future Goals</div>' +
        '<div class="fb-title">What You Did Well:</div>' +
        '<div style="margin-bottom:8px;">' + (fb.whatWentWell || 'Demonstrated consistent design engagement and practical skill development.') + '</div>' +
        '<div class="fb-title">Areas for Improvement:</div>' +
        '<div style="margin-bottom:8px;">' + (fb.areasForImprovement || 'Deepen reflection on material experimentation and technical joining methods.') + '</div>' +
        '<div class="fb-title">Goals for Next Assessment:</div>' +
        '<div>' + (fb.goalsForNextAssessment || '1. Strengthen time tracking with annotated Gantt milestones.\n2. Refine surface finishing on practical models.') + '</div>' +
      '</div>' +
      '</body></html>';
  }

  function getCourseworkPickerHtml() {
    return '<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8">' +
      '<style>body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;margin:0;padding:20px;background:#f8fafc;color:#1e293b;}' +
      'h2{margin-top:0;font-size:18px;color:#0f172a;}.form-group{margin-bottom:16px;}label{display:block;font-weight:600;font-size:13px;margin-bottom:6px;}' +
      'select{width:100%;padding:10px;border:1px solid #cbd5e1;border-radius:6px;background:#fff;font-size:14px;}' +
      '.btn-group{margin-top:24px;display:flex;justify-content:flex-end;gap:10px;}button{padding:8px 16px;border-radius:6px;font-weight:600;font-size:13px;cursor:pointer;border:none;}' +
      '.btn-primary{background:#2563eb;color:#fff;}.btn-secondary{background:#e2e8f0;color:#334155;}.status{margin-top:12px;font-size:13px;font-weight:500;}' +
      '</style></head><body><h2>Select Classroom Course &amp; Assignment</h2>' +
      '<div class="form-group"><label>1. Course:</label><select id="courseSelect" onchange="onCourseChanged()"><option value="">Loading courses...</option></select></div>' +
      '<div class="form-group"><label>2. Assignment:</label><select id="cwSelect" disabled><option value="">Select course first</option></select></div>' +
      '<div id="status" class="status"></div>' +
      '<div class="btn-group"><button class="btn-secondary" onclick="google.script.host.close()">Cancel</button><button id="saveBtn" class="btn-primary" onclick="saveSelection()" disabled>Save Assignment</button></div>' +
      '<script>' +
      'window.onload=function(){google.script.run.withSuccessHandler(function(r){var s=document.getElementById("courseSelect");s.innerHTML="<option value=\\"\\">-- Select Course --</option>";var cl=r.courses||[];for(var i=0;i<cl.length;i++){var c=cl[i];var o=document.createElement("option");o.value=c.id;o.textContent=c.name+(c.section?" ("+c.section+")":"");s.appendChild(o);}}).apiListTeacherCourses();};' +
      'function onCourseChanged(){var cid=document.getElementById("courseSelect").value;var cws=document.getElementById("cwSelect");if(!cid){cws.disabled=true;return;}cws.disabled=true;cws.innerHTML="<option>Loading coursework...</option>";google.script.run.withSuccessHandler(function(r){cws.innerHTML="<option value=\\"\\">-- Select Assignment --</option>";var cwl=r.coursework||[];for(var j=0;j<cwl.length;j++){var cw=cwl[j];var o=document.createElement("option");o.value=cw.id;o.textContent=cw.title+" ("+cw.maxPoints+" pts)";cws.appendChild(o);}cws.disabled=false;document.getElementById("saveBtn").disabled=false;}).apiListCourseWork(cid);}' +
      'function saveSelection(){var cid=document.getElementById("courseSelect").value;var cwid=document.getElementById("cwSelect").value;document.getElementById("status").textContent="Saving...";google.script.run.withSuccessHandler(function(r){alert(r.message);google.script.host.close();}).apiSaveClassroomSelection(cid, cwid);}' +
      '<\/script></body></html>';
  }

  function getClassBatchGradingHtml() {
    return '<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8">' +
      '<title>Class AI Grading Runner</title>' +
      '<style>body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;margin:0;padding:24px;background:#f8fafc;color:#0f172a;}' +
      '.bar-wrap{background:#e2e8f0;border-radius:8px;height:24px;width:100%;overflow:hidden;margin:16px 0;position:relative;}' +
      '.bar-fill{background:linear-gradient(90deg,#7c3aed,#2563eb);height:100%;width:0%;transition:width 0.4s ease;}' +
      '.bar-text{position:absolute;top:0;left:0;right:0;bottom:0;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;color:#fff;text-shadow:0 1px 2px rgba(0,0,0,0.5);}' +
      '.status-box{background:#fff;border:1px solid #cbd5e1;border-radius:6px;padding:14px;font-size:13px;min-height:90px;display:flex;flex-direction:column;justify-content:center;}' +
      '.log-box{margin-top:14px;height:120px;overflow-y:auto;background:#0f172a;color:#38bdf8;padding:10px;border-radius:6px;font-family:monospace;font-size:11px;}' +
      'button{padding:8px 18px;border-radius:6px;font-size:13px;font-weight:700;cursor:pointer;border:none;}' +
      '.btn-start{background:#7c3aed;color:#fff;}.btn-cancel{background:#e2e8f0;color:#334155;}' +
      '</style></head><body>' +
      '<h2 style="margin-top:0;">Class-Wide AI Assessment Runner</h2>' +
      '<p style="font-size:13px;color:#475569;">Sequentially grades student submissions with high-accuracy vision checks, derives grades, and generates feedback drafts with 2.5s pacing intervals.</p>' +
      '<div class="bar-wrap"><div id="bar" class="bar-fill"></div><div id="barTxt" class="bar-text">0%</div></div>' +
      '<div class="status-box"><div id="curStudent" style="font-weight:700;">Ready to start</div><div id="subDetail" style="font-size:12px;color:#64748b;margin-top:4px;">Click Start to begin grading all students.</div></div>' +
      '<div id="log" class="log-box">&gt; System ready.</div>' +
      '<div style="margin-top:18px;display:flex;justify-content:flex-end;gap:10px;">' +
        '<button class="btn-cancel" onclick="stopOrClose()">Close</button>' +
        '<button id="btnStart" class="btn-start" onclick="startBatch()">Start Class Grading</button>' +
      '</div>' +
      '<script>' +
      'var list = [], currentIdx = 0, isRunning = false;' +
      'window.onload = function(){ google.script.run.withSuccessHandler(function(subs){ list = subs || []; document.getElementById("subDetail").textContent = list.length + " submissions available for grading."; }).apiGetClassSubmissions(); };' +
      'function log(msg){ var el = document.getElementById("log"); el.innerHTML += "<div>&gt; " + msg + "</div>"; el.scrollTop = el.scrollHeight; }' +
      'function startBatch(){ if(list.length === 0){ alert("No submissions found."); return; } isRunning = true; document.getElementById("btnStart").disabled = true; processNext(); }' +
      'function processNext(){' +
        'if(!isRunning || currentIdx >= list.length){ finishBatch(); return; }' +
        'var s = list[currentIdx];' +
        'var pct = Math.round(((currentIdx + 1) / list.length) * 100);' +
        'document.getElementById("bar").style.width = pct + "%";' +
        'document.getElementById("barTxt").textContent = pct + "% (" + (currentIdx + 1) + "/" + list.length + ")";' +
        'document.getElementById("curStudent").textContent = "Grading: " + s.StudentName + " (v" + s.SubmissionVersion + ")";' +
        'document.getElementById("subDetail").textContent = "Auditing Drive files & generating feedback...";' +
        'log("Assessing " + s.StudentName + "...");' +
        'google.script.run.withSuccessHandler(function(res){' +
          'if(res.success){ log("Completed " + s.StudentName + ": " + res.message); }' +
          'else { log(s.StudentName + " error: " + res.message); }' +
          'currentIdx++;' +
          'setTimeout(processNext, 2500);' +
        '}).withFailureHandler(function(err){' +
          'log("Script error on " + s.StudentName + ": " + err.message);' +
          'currentIdx++;' +
          'setTimeout(processNext, 2500);' +
        '}).apiGradeSingleStudentBatch(s.SubmissionRecordID);' +
      '}' +
      'function finishBatch(){ isRunning = false; document.getElementById("curStudent").textContent = "Class AI Grading Complete!"; document.getElementById("subDetail").textContent = "All " + list.length + " students graded."; log("Batch run complete. Please review marks in Assessment Studio."); document.getElementById("btnStart").style.display = "none"; }' +
      'function stopOrClose(){ isRunning = false; google.script.host.close(); }' +
      '<\/script></body></html>';
  }

  function getVoiceDictationHtml(criterionId) {
    return '<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8">' +
      '<style>body{font-family:sans-serif;margin:0;padding:16px;background:#f8fafc;color:#0f172a;}textarea{width:100%;height:110px;padding:8px;box-sizing:border-box;border-radius:4px;border:1px solid #cbd5e1;}' +
      '.btn-group{margin-top:10px;display:flex;justify-content:space-between;}button{padding:8px 14px;border-radius:4px;border:none;font-weight:600;cursor:pointer;}' +
      '.btn-mic{background:#dc2626;color:#fff;}.btn-send{background:#2563eb;color:#fff;}.btn-cancel{background:#e2e8f0;}' +
      '</style></head><body><h3>Dictate Notes (' + criterionId + ')</h3>' +
      '<textarea id="txt" placeholder="Click Start Dictating and speak..."></textarea>' +
      '<div class="btn-group"><button id="btnMic" class="btn-mic" onclick="toggle()">Start Dictating</button><div><button class="btn-send" onclick="send()">Insert Note</button> <button class="btn-cancel" onclick="google.script.host.close()">Close</button></div></div>' +
      '<script>' +
      'var cid="' + criterionId + '", rec=null, run=false;' +
      'window.onload=function(){var SR=window.SpeechRecognition||window.webkitSpeechRecognition;if(!SR){alert("Speech API not supported in this browser.");return;}rec=new SR();rec.continuous=true;rec.interimResults=true;rec.lang="en-AU";rec.onresult=function(e){var t="";for(var i=e.resultIndex;i<e.results.length;++i)t+=e.results[i][0].transcript;document.getElementById("txt").value=t;};rec.onend=function(){run=false;document.getElementById("btnMic").textContent="Start Dictating";};};' +
      'function toggle(){if(run){rec.stop();run=false;document.getElementById("btnMic").textContent="Start Dictating";}else{rec.start();run=true;document.getElementById("btnMic").textContent="Stop Dictating";}}' +
      'function send(){if(window.opener){window.opener.postMessage({type:"VOICE_TRANSCRIPT",criterionId:cid,transcript:document.getElementById("txt").value},"*");}google.script.host.close();}' +
      '<\/script></body></html>';
  }

  function getAssessmentStudioHtml() {
    var head =
      '<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8">' +
      '<title>AI Assessment Studio</title>' +
      '<style>' +
      '*{box-sizing:border-box;}' +
      'html,body{margin:0;padding:0;width:100%;height:100%;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;background:#f8fafc;color:#0f172a;overflow:hidden;}' +
      'header{background:#fff;border-bottom:1px solid #e2e8f0;padding:8px 16px;display:flex;flex-wrap:wrap;justify-content:space-between;align-items:center;gap:10px;box-shadow:0 1px 3px rgba(0,0,0,0.06);z-index:10;}' +
      '.header-left{display:flex;align-items:center;gap:10px;flex-wrap:wrap;}' +
      'select{font-size:13px;font-weight:600;padding:6px 10px;border-radius:6px;border:1px solid #cbd5e1;background:#f1f5f9;max-width:260px;}' +
      '.badge{font-size:11px;font-weight:700;padding:3px 8px;border-radius:12px;background:#e2e8f0;color:#334155;}' +
      '.score-banner{display:flex;align-items:center;gap:14px;background:#0f172a;color:#fff;padding:6px 14px;border-radius:8px;}' +
      '.score-item{display:flex;flex-direction:column;align-items:center;line-height:1.1;}' +
      '.score-label{font-size:9px;text-transform:uppercase;color:#94a3b8;font-weight:600;}' +
      '.score-val{font-size:15px;font-weight:700;}' +
      '.score-grade{font-size:17px;font-weight:800;color:#38bdf8;}' +
      '.header-actions{display:flex;align-items:center;gap:6px;flex-wrap:wrap;}' +
      'button{padding:7px 12px;border-radius:6px;font-size:12px;font-weight:700;cursor:pointer;border:none;}' +
      'button:disabled{opacity:0.5;cursor:not-allowed;}' +
      '.btn-ai{background:#7c3aed;color:#fff;}.btn-save{background:#0284c7;color:#fff;}.btn-approve{background:#16a34a;color:#fff;}' +
      '.btn-lock{background:#dc2626;color:#fff;}.btn-sync{background:#059669;color:#fff;}.btn-sec{background:#e2e8f0;color:#334155;}' +
      '.btn-pdf{background:#b91c1c;color:#fff;}' +
      '.studio-body{display:flex;height:calc(100vh - 60px);overflow:hidden;}' +
      '.left-panel{width:48%;min-width:400px;background:#1e293b;display:flex;flex-direction:column;overflow:hidden;}' +
      '.viewer-bar{background:#334155;color:#fff;padding:6px 10px;font-size:11px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:6px;}' +
      '.file-tabs{display:flex;overflow-x:auto;gap:4px;padding:6px 10px;background:#0f172a;}' +
      '.file-tab{padding:5px 10px;font-size:11px;font-weight:600;background:#334155;color:#e2e8f0;border-radius:4px;cursor:pointer;white-space:nowrap;}' +
      '.file-tab.active{background:#2563eb;color:#fff;}' +
      '.viewer-content{flex:1;overflow-y:auto;overflow-x:hidden;background:#0f172a;display:flex;flex-direction:column;align-items:center;padding:14px;gap:12px;}' +
      '.embed-frame{width:100%;min-height:75vh;border:0;background:#fff;border-radius:4px;}' +
      '.image-viewer{max-width:100%;border-radius:4px;}' +
      '.right-panel{width:52%;flex:1;overflow-y:auto;padding:14px 18px 140px;}' +
      '.criterion-card{background:#fff;border:1px solid #e2e8f0;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,0.05);padding:14px;margin-bottom:14px;}' +
      '.crit-header{display:flex;justify-content:space-between;align-items:flex-start;border-bottom:1px solid #f1f5f9;padding-bottom:8px;margin-bottom:8px;}' +
      '.crit-id-badge{font-size:11px;font-weight:700;background:#0f172a;color:#fff;padding:2px 7px;border-radius:4px;margin-right:6px;}' +
      '.crit-name{font-size:14px;font-weight:700;}' +
      '.crit-meta{font-size:11px;color:#64748b;margin-top:2px;}' +
      '.grade-badge{font-size:14px;font-weight:800;width:30px;height:30px;border-radius:50%;display:flex;align-items:center;justify-content:center;border:2px solid;flex-shrink:0;}' +
      '.grade-A{background:#dcfce7;color:#15803d;border-color:#15803d;}.grade-B{background:#dbeafe;color:#1d4ed8;border-color:#1d4ed8;}' +
      '.grade-C{background:#fef3c7;color:#b45309;border-color:#b45309;}.grade-D{background:#ffedd5;color:#c2410c;border-color:#ea580c;}' +
      '.grade-E{background:#fee2e2;color:#b91c1c;border-color:#b91c1c;}' +
      '.explanation-box{font-size:11px;background:#f8fafc;border:1px solid #e2e8f0;padding:6px 10px;border-radius:6px;color:#334155;margin-bottom:8px;}' +
      '.md-box{background:#fffbeb;border:2px dashed #f59e0b;border-radius:6px;padding:8px 10px;display:flex;align-items:flex-start;gap:8px;margin-bottom:8px;font-size:12px;}' +
      '.band-block{border:1px solid #e2e8f0;border-radius:6px;margin-bottom:6px;overflow:hidden;}' +
      '.band-header{background:#f8fafc;padding:6px 10px;font-size:11px;font-weight:700;color:#334155;}' +
      '.band-items{padding:6px 10px;background:#fff;}' +
      '.obs-row{display:flex;align-items:flex-start;gap:6px;font-size:11px;line-height:1.4;margin-bottom:5px;}' +
      '.teacher-textarea{width:100%;height:55px;padding:6px;font-size:11px;border:1px solid #cbd5e1;border-radius:4px;font-family:inherit;margin-top:6px;}' +
      '.feedback-dock{position:fixed;bottom:0;left:0;right:0;background:#fff;border-top:2px solid #0f172a;padding:10px 18px;z-index:20;}' +
      '.feedback-dock-body{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-top:8px;}' +
      '.fb-text{width:100%;height:60px;padding:6px;font-size:11px;border:1px solid #cbd5e1;border-radius:4px;font-family:inherit;}' +
      '.ai-hud{position:fixed;top:70px;right:20px;background:#0f172a;color:#fff;padding:14px 18px;border-radius:8px;font-size:12px;z-index:30;display:none;min-width:260px;box-shadow:0 4px 12px rgba(0,0,0,0.3);}' +
      '</style></head><body>';

    var headerHtml =
      '<header>' +
        '<div class="header-left">' +
          '<label style="font-size:11px;font-weight:700;color:#475569;">STUDENT</label>' +
          '<select id="stuSel" onchange="onStudentChanged()"><option value="">Loading submissions...</option></select>' +
          '<span id="versionBadge" class="badge">v1</span>' +
          '<span id="statusBadge" class="badge">New</span>' +
        '</div>' +
        '<div class="score-banner">' +
          '<div class="score-item"><span class="score-label">Part A</span><span id="scA" class="score-val">0.0</span></div>' +
          '<div class="score-item"><span class="score-label">Part B</span><span id="scB" class="score-val">0.0</span></div>' +
          '<div class="score-item"><span class="score-label">Total /100</span><span id="scTot" class="score-val">0.0</span></div>' +
          '<div class="score-item"><span class="score-label">Grade</span><span id="scG" class="score-grade">E</span></div>' +
        '</div>' +
        '<div class="header-actions">' +
          '<button id="btnAiRun" class="btn-ai" onclick="runAi()">Run AI</button>' +
          '<button class="btn-sec" onclick="openClassGrading()">Grade Class</button>' +
          '<button class="btn-pdf" onclick="downloadReportPdf()">Report PDF</button>' +
          '<button class="btn-save" onclick="saveDraft()">Save Draft</button>' +
          '<button class="btn-approve" onclick="approve()">Approve</button>' +
          '<button class="btn-lock" onclick="lockFinal()">Lock</button>' +
          '<button class="btn-sync" onclick="syncClassroom()">Sync Classroom</button>' +
        '</div>' +
      '</header>';

    var bodyHtml =
      '<div class="studio-body">' +
        '<div class="left-panel">' +
          '<div class="viewer-bar">' +
            '<span id="fileCountBadge">0 files</span>' +
            '<span id="slideCounter"></span>' +
          '</div>' +
          '<div id="fileTabs" class="file-tabs"></div>' +
          '<div id="vContent" class="viewer-content"><div style="color:#94a3b8;margin-top:40px;">No file selected</div></div>' +
        '</div>' +
        '<div class="right-panel">' +
          '<div id="critContainer"></div>' +
        '</div>' +
      '</div>' +
      '<div class="ai-hud" id="aiHud">' +
        '<div id="hudStatus" style="font-weight:700;">Scanning...</div>' +
        '<div id="hudDetail" style="font-size:11px;color:#94a3b8;margin-top:6px;font-family:monospace;"></div>' +
      '</div>' +
      '<div class="feedback-dock">' +
        '<div style="display:flex;justify-content:space-between;align-items:center;">' +
          '<span style="font-size:13px;font-weight:700;">Student Feedback Preview</span>' +
          '<button class="btn-sec" onclick="genFb()" style="font-size:11px;">Generate AI Feedback</button>' +
        '</div>' +
        '<div class="feedback-dock-body">' +
          '<div><div style="font-size:11px;font-weight:700;margin-bottom:4px;">What went well</div><textarea id="fbW" class="fb-text"></textarea></div>' +
          '<div><div style="font-size:11px;font-weight:700;margin-bottom:4px;">Areas for improvement</div><textarea id="fbI" class="fb-text"></textarea></div>' +
          '<div><div style="font-size:11px;font-weight:700;margin-bottom:4px;">Goals for next assessment</div><textarea id="fbG" class="fb-text"></textarea></div>' +
        '</div>' +
      '</div>';

    var scriptHtml =
      '<script>' +
      'var state = { criteria: [], submissions: [], currentSubId: null, detail: null, files: [], fileIdx: 0, slideNum: 1 };' +
      'var GRADE_WEIGHTS = { A: 1.0, B: 0.875, C: 0.70, D: 0.575, E: 0.25 };' +

      'window.onload = function() {' +
        'window.addEventListener("message", function(event) {' +
          'if (event.data && event.data.type === "VOICE_TRANSCRIPT") {' +
            'var t = document.getElementById("note-" + event.data.criterionId);' +
            'if (t) { t.value = (t.value ? t.value + " " : "") + event.data.transcript; onCritNoteChanged(event.data.criterionId); }' +
          '}' +
        '});' +
        'loadBootstrapData(null);' +
      '};' +

      'function loadBootstrapData(targetSubId) {' +
        'google.script.run.withSuccessHandler(function(res) {' +
          'if (!res) { showStudioError("No response from server."); return; }' +
          'state.criteria = res.criteria || [];' +
          'state.submissions = res.submissions || [];' +
          'state.gradeWeights = res.gradeWeights || { A: 1.0, B: 0.875, C: 0.70, D: 0.575, E: 0.25 };' +
        'state.overallGradeBands = res.overallGradeBands || [{ letter: "A", min: 85 }, { letter: "B", min: 75 }, { letter: "C", min: 65 }, { letter: "D", min: 50 }, { letter: "E", min: 0 }];' +
          'var sel = document.getElementById("stuSel");' +
          'sel.innerHTML = "";' +
          'if (state.submissions.length === 0) {' +
            'sel.innerHTML = "<option value=\\"\\">-- No submissions imported --</option>";' +
            'document.getElementById("critContainer").innerHTML = "<div style=\\"padding:30px;text-align:center;color:#64748b;\\"><h3>No Student Submissions Found</h3><p>Close this window, then run Classroom &gt; Import/Refresh Submissions.</p></div>";' +
            'return;' +
          '}' +
          'state.submissions.forEach(function(s) {' +
            'var o = document.createElement("option");' +
            'o.value = s.SubmissionRecordID;' +
            'o.textContent = s.StudentName + " (v" + s.SubmissionVersion + " - " + s.Status + ")";' +
            'if (s.SubmissionRecordID === res.selectedSubmissionId) o.selected = true;' +
            'sel.appendChild(o);' +
          '});' +
          'state.currentSubId = sel.value;' +
          'if (res.submissionDetail) { renderSubmission(res.submissionDetail); }' +
        '}).withFailureHandler(function(err) {' +
          'showStudioError(err && err.message ? err.message : String(err));' +
        '}).apiGetBootstrapData(targetSubId);' +
      '}' +

      'function showStudioError(msg) {' +
        'var el = document.getElementById("critContainer");' +
        'if (el) el.innerHTML = "<div style=\\"padding:20px;color:#b91c1c;background:#fee2e2;border-radius:8px;\\"><b>Error loading studio data</b><br>" + msg + "</div>";' +
      '}' +

      'function onStudentChanged() {' +
        'state.currentSubId = document.getElementById("stuSel").value;' +
        'state.fileIdx = 0; state.slideNum = 1;' +
        'google.script.run.withSuccessHandler(renderSubmission).withFailureHandler(function(err){ showStudioError(err.message); }).apiGetSubmissionDetail(state.currentSubId);' +
      '}' +

      'function renderSubmission(detail) {' +
        'state.detail = detail;' +
        'state.files = detail.files || [];' +
        'state.fileIdx = 0;' +
        'document.getElementById("versionBadge").textContent = "v" + detail.submission.SubmissionVersion;' +
        'document.getElementById("statusBadge").textContent = detail.submission.Status;' +
        'renderFileTabs();' +
        'renderCriteriaCards();' +
        'var fb = detail.feedback || {};' +
        'document.getElementById("fbW").value = fb.whatWentWell || "";' +
        'document.getElementById("fbI").value = fb.areasForImprovement || "";' +
        'document.getElementById("fbG").value = fb.goalsForNextAssessment || "";' +
        'recalculateOverallScore();' +
      '}' +

      'function renderFileTabs() {' +
        'var tabs = document.getElementById("fileTabs");' +
        'tabs.innerHTML = "";' +
        'var files = state.files;' +
        'document.getElementById("fileCountBadge").textContent = files.length + " file(s)";' +
        'if (!files.length) {' +
          'document.getElementById("vContent").innerHTML = "<div style=\\"color:#94a3b8;margin-top:40px;\\">No files attached</div>";' +
          'return;' +
        '}' +
        'files.forEach(function(f, idx) {' +
          'var tab = document.createElement("div");' +
          'tab.className = "file-tab" + (idx === state.fileIdx ? " active" : "");' +
          'tab.textContent = f.FileName || ("File " + (idx + 1));' +
          'tab.onclick = function() { state.fileIdx = idx; state.slideNum = 1; renderFileTabs(); };' +
          'tabs.appendChild(tab);' +
        '});' +
        'renderActiveFile();' +
      '}' +

      'function renderActiveFile() {' +
        'var viewer = document.getElementById("vContent");' +
        'var f = state.files[state.fileIdx];' +
        'if (!f || !f.DriveFileID) { viewer.innerHTML = "<div style=\\"color:#94a3b8;margin-top:40px;\\">File unavailable</div>"; return; }' +
        'var mime = (f.MimeType || "").toLowerCase();' +
        'document.getElementById("slideCounter").textContent = "";' +
        'if (mime.indexOf("presentation") !== -1) {' +
          'viewer.innerHTML = "<iframe class=\\"embed-frame\\" src=\\"https://docs.google.com/presentation/d/" + f.DriveFileID + "/embed?rm=minimal\\"></iframe>";' +
        '} else if (mime.indexOf("document") !== -1) {' +
          'viewer.innerHTML = "<iframe class=\\"embed-frame\\" src=\\"https://docs.google.com/document/d/" + f.DriveFileID + "/preview\\"></iframe>";' +
        '} else if (mime.indexOf("pdf") !== -1) {' +
          'viewer.innerHTML = "<iframe class=\\"embed-frame\\" src=\\"https://drive.google.com/file/d/" + f.DriveFileID + "/preview\\"></iframe>";' +
        '} else if (mime.indexOf("image") !== -1) {' +
          'viewer.innerHTML = "<img class=\\"image-viewer\\" src=\\"https://drive.google.com/thumbnail?id=" + f.DriveFileID + "&sz=w1600\\">";' +
        '} else {' +
          'viewer.innerHTML = "<div style=\\"color:#fff;\\">Unsupported preview &mdash; <a href=\\"" + (f.AlternateLink || "#") + "\\" target=\\"_blank\\" style=\\"color:#38bdf8;\\">Open in Drive</a></div>";' +
        '}' +
      '}' +

      'function getCriterion(cid) {' +
        'for (var i = 0; i < state.criteria.length; i++) { if (state.criteria[i].criterionId === cid) return state.criteria[i]; }' +
        'return null;' +
      '}' +

      'function renderCriteriaCards() {' +
        'var container = document.getElementById("critContainer");' +
        'container.innerHTML = "";' +
        'state.criteria.forEach(function(crit) {' +
          'var cid = crit.criterionId;' +
          'var cd = (state.detail && state.detail.criteriaMap && state.detail.criteriaMap[cid]) ? state.detail.criteriaMap[cid] : { checkboxes: {}, derivedGrade: "E", explanation: "" };' +
          'var card = document.createElement("div");' +
          'card.className = "criterion-card";' +
          'var html = "";' +
          'html += "<div class=\\"crit-header\\"><div><span class=\\"crit-id-badge\\">" + cid + "</span><span class=\\"crit-name\\">" + crit.title + "</span><div class=\\"crit-meta\\">" + crit.part + " &bull; /" + crit.maxMarks + " Marks</div></div>";' +
          'html += "<div class=\\"grade-badge grade-" + (cd.derivedGrade || "E") + "\\" id=\\"badge-" + cid + "\\">" + (cd.derivedGrade || "E") + "</div></div>";' +
          'html += "<div class=\\"explanation-box\\" id=\\"exp-" + cid + "\\">" + (cd.explanation || "") + "</div>";' +
          'var mdChecked = !!(cd.checkboxes && cd.checkboxes[crit.mdOverride.id]);' +
          'html += "<div class=\\"md-box\\"><input type=\\"checkbox\\" id=\\"chk-" + crit.mdOverride.id + "\\" " + (mdChecked ? "checked" : "") + " onchange=\\"onCritCheckToggled(this,\'" + cid + "\',\'" + crit.mdOverride.id + "\')\\"><div><b>Distinct A Override</b> " + crit.mdOverride.text + "</div></div>";' +
          '["A","B","C","D","E"].forEach(function(bandName) {' +
            'var items = crit.bands[bandName] || [];' +
            'html += "<div class=\\"band-block\\"><div class=\\"band-header\\">" + bandName + " Band</div><div class=\\"band-items\\">";' +
            'items.forEach(function(it) {' +
              'var checked = !!(cd.checkboxes && cd.checkboxes[it[0]]);' +
              'html += "<div class=\\"obs-row\\"><input type=\\"checkbox\\" id=\\"chk-" + it[0] + "\\" " + (checked ? "checked" : "") + " onchange=\\"onCritCheckToggled(this,\'" + cid + "\',\'" + it[0] + "\')\\"><span>" + it[1] + "</span></div>";' +
            '});' +
            'html += "</div></div>";' +
          '});' +
          'html += "<div style=\\"display:flex;justify-content:space-between;align-items:center;margin-top:6px;\\"><label style=\\"font-size:11px;font-weight:700;\\">Observations / Notes</label><button class=\\"btn-sec\\" style=\\"font-size:11px;\\" onclick=\\"openVoice(\'" + cid + "\')\\">Dictate</button></div>";' +
          'html += "<textarea id=\\"note-" + cid + "\\" class=\\"teacher-textarea\\" oninput=\\"onCritNoteChanged(\'" + cid + "\')\\">" + (cd.teacherWrittenNote || cd.teacherAudioTranscript || "") + "</textarea>";' +
          'card.innerHTML = html;' +
          'container.appendChild(card);' +
        '});' +
      '}' +

'function nearestGradeLetterClient(score, weights) {' +
        'var letters = ["A", "B", "C", "D", "E"];' +
        'var sorted = letters.filter(function(l) { return weights[l] !== undefined; })' +
          '.sort(function(a, b) { return weights[b] - weights[a]; });' +
        'if (!sorted.length) return "E";' +
        'for (var i = 0; i < sorted.length - 1; i++) {' +
          'var upper = sorted[i], lower = sorted[i + 1];' +
          'var midpoint = (weights[upper] + weights[lower]) / 2;' +
          'if (score >= midpoint - 1e-9) return upper;' +
        '}' +
        'return sorted[sorted.length - 1];' +
      '}' +

      'function deriveGradeLocal(crit, checksMap) {' +
        'var W = state.gradeWeights || { A: 1.0, B: 0.875, C: 0.70, D: 0.575, E: 0.25 };' +
        'if (checksMap[crit.mdOverride.id]) return { grade: "A", score: W.A, explanation: "Distinct A evidence selected \\u2014 full marks awarded." };' +
        'function frac(bandName) {' +
          'var items = crit.bands[bandName] || [];' +
          'if (!items.length) return 0;' +
          'var ticked = 0;' +
          'items.forEach(function(it) { if (checksMap[it[0]]) ticked++; });' +
          'return ticked / items.length;' +
        '}' +
        'var score = W.E;' +
        'score += frac("D") * (W.D - W.E);' +
        'score += frac("C") * (W.C - W.D);' +
        'score += frac("B") * (W.B - W.C);' +
        'score += frac("A") * (W.A - W.B);' +
        'if (score > W.A) score = W.A;' +
        'var grade = nearestGradeLetterClient(score, W);' +
        'return { grade: grade, score: score, explanation: "Weighted score: " + Math.round(score * 100) + "% \\u2192 nominal grade " + grade + "." };' +
      '}' +

      'function onCritCheckToggled(el, cid, checkId) {' +
        'if (!state.detail || !state.detail.criteriaMap) return;' +
        'var cd = state.detail.criteriaMap[cid];' +
        'if (!cd) return;' +
        'cd.checkboxes = cd.checkboxes || {};' +
        'cd.checkboxes[checkId] = el.checked;' +
        'var crit = getCriterion(cid);' +
        'if (!crit) return;' +
        'var calc = deriveGradeLocal(crit, cd.checkboxes);' +
        'cd.derivedGrade = calc.grade;' +
        'cd.explanation = calc.explanation;' +
        'var badge = document.getElementById("badge-" + cid);' +
        'if (badge) { badge.textContent = calc.grade; badge.className = "grade-badge grade-" + calc.grade; }' +
        'var expEl = document.getElementById("exp-" + cid);' +
        'if (expEl) expEl.textContent = calc.explanation;' +
        'recalculateOverallScore();' +
      '}' +

      'function onCritNoteChanged(cid) {' +
        'if (!state.detail || !state.detail.criteriaMap) return;' +
        'var cd = state.detail.criteriaMap[cid];' +
        'if (!cd) return;' +
        'var el = document.getElementById("note-" + cid);' +
        'if (el) cd.teacherWrittenNote = el.value;' +
      '}' +

      'function recalculateOverallScore() {' +
        'if (!state.detail || !state.detail.criteriaMap) return;' +
        'var W = state.gradeWeights || { A: 1.0, B: 0.875, C: 0.70, D: 0.575, E: 0.25 };' +
        'var pA = 0, pB = 0;' +
        'state.criteria.forEach(function(c) {' +
          'var cd = state.detail.criteriaMap[c.criterionId];' +
          'var score = (cd && cd.derivedScore !== undefined) ? cd.derivedScore : W.E;' +
          'var marks = c.maxMarks * score;' +
          'if (c.part === "Part A") pA += marks; else pB += marks;' +
        '});' +
        'var tot = pA + pB;' +
        'var bands = state.overallGradeBands || [{ letter: "A", min: 85 }, { letter: "B", min: 75 }, { letter: "C", min: 65 }, { letter: "D", min: 50 }, { letter: "E", min: 0 }];' +
        'var og = bands.length ? bands[bands.length - 1].letter : "E";' +
        'for (var i = 0; i < bands.length; i++) { if (tot >= bands[i].min) { og = bands[i].letter; break; } }' +
        'document.getElementById("scA").textContent = pA.toFixed(2);' +
        'document.getElementById("scB").textContent = pB.toFixed(2);' +
        'document.getElementById("scTot").textContent = tot.toFixed(1);' +
        'document.getElementById("scG").textContent = og;' +
      '}' +
      
      'function openVoice(cid) { google.script.run.openVoiceDictation(cid); }' +
      'function openClassGrading() { google.script.run.openClassAiGradingRunner(); }' +

      'function saveDraft() {' +
        'if (!state.detail) return;' +
        'var fb = { whatWentWell: document.getElementById("fbW").value, areasForImprovement: document.getElementById("fbI").value, goalsForNextAssessment: document.getElementById("fbG").value };' +
        'google.script.run.withSuccessHandler(function(r) { alert(r.message); }).apiSaveAssessmentDraft(state.detail.assessmentId, state.currentSubId, state.detail.criteriaMap, fb);' +
      '}' +

      'function approve() {' +
        'if (!state.detail) return;' +
        'if (!confirm("Approve assessment and sync to marking sheet?")) return;' +
        'var fb = { whatWentWell: document.getElementById("fbW").value, areasForImprovement: document.getElementById("fbI").value, goalsForNextAssessment: document.getElementById("fbG").value };' +
        'google.script.run.withSuccessHandler(function(r) { alert(r.message); loadBootstrapData(state.currentSubId); }).apiApproveAssessment(state.detail.assessmentId, state.currentSubId, fb);' +
      '}' +

      'function lockFinal() {' +
        'if (!state.detail) return;' +
        'if (!confirm("Permanently lock this assessment? Reassessment will require a new version.")) return;' +
        'google.script.run.withSuccessHandler(function(r) { alert(r.message); loadBootstrapData(state.currentSubId); }).apiLockAssessment(state.detail.assessmentId, state.currentSubId);' +
      '}' +

      'function syncClassroom() {' +
        'if (!state.detail) return;' +
        'if (!confirm("Sync calculated marks to Google Classroom?")) return;' +
        'google.script.run.withSuccessHandler(function(r) { alert(r.message); }).apiSyncToClassroom(state.detail.assessmentId, state.currentSubId);' +
      '}' +

  'function runAi() {' +
        'if (!state.detail) return;' +
        'if (!confirm("Run Gemini AI evidence pass on this submission?")) return;' +
        'var hud = document.getElementById("aiHud");' +
        'var btn = document.getElementById("btnAiRun");' +
        'hud.style.display = "block"; btn.disabled = true; btn.textContent = "Scanning...";' +

        'var MIN_STEP_MS = 550;' +
        'var filesToScan = (state.files || []).filter(function(f) { return f.DriveFileID; });' +
        'var preparedFiles = [];' +
        'var idx = 0;' +

        'if (!filesToScan.length) {' +
          'document.getElementById("hudStatus").textContent = "No files to scan";' +
          'document.getElementById("hudDetail").textContent = "This submission has no eligible Drive attachments.";' +
        '}' +

        'function scanNext() {' +
          'if (idx >= filesToScan.length) {' +
            'document.getElementById("hudStatus").textContent = "Cross-checking rubric observables...";' +
            'document.getElementById("hudDetail").textContent = "Sending " + preparedFiles.length + " reviewed file(s) to Gemini for analysis...";' +
            'google.script.run.withSuccessHandler(function(r) {' +
              'document.getElementById("hudStatus").textContent = "Audit complete!";' +
              'document.getElementById("hudDetail").textContent = r.message;' +
              'setTimeout(function() {' +
                'hud.style.display = "none"; btn.disabled = false; btn.textContent = "Run AI";' +
                'loadBootstrapData(state.currentSubId);' +
              '}, 900);' +
            '}).withFailureHandler(function(err) {' +
              'hud.style.display = "none"; btn.disabled = false; btn.textContent = "Run AI";' +
              'alert("AI Error: " + err.message);' +
            '}).apiRunAiAssessmentWithFiles(state.currentSubId, preparedFiles);' +
            'return;' +
          '}' +

          'var f = filesToScan[idx];' +
          'document.getElementById("hudStatus").textContent = "Scanning file " + (idx + 1) + " of " + filesToScan.length;' +
          'document.getElementById("hudDetail").textContent = "Reading from Drive: " + (f.FileName || ("attachment " + (idx + 1))) + "...";' +

          'setTimeout(function() {' +
            'var startedAt = Date.now();' +
            'google.script.run.withSuccessHandler(function(result) {' +
              'if (result && result.status === "reviewed") {' +
                'preparedFiles.push(result);' +
                'document.getElementById("hudDetail").textContent = "Reviewed: " + result.fileName;' +
              '} else {' +
                'document.getElementById("hudDetail").textContent = (result ? result.fileName : f.FileName) + " skipped (" + (result ? result.status : "error") + ").";' +
              '}' +
              'var elapsed = Date.now() - startedAt;' +
              'var wait = Math.max(0, MIN_STEP_MS - elapsed);' +
              'setTimeout(function() { idx++; scanNext(); }, wait);' +
            '}).withFailureHandler(function(err) {' +
              'document.getElementById("hudDetail").textContent = "Error reading " + f.FileName + ": " + err.message;' +
              'var elapsed = Date.now() - startedAt;' +
              'var wait = Math.max(0, MIN_STEP_MS - elapsed);' +
              'setTimeout(function() { idx++; scanNext(); }, wait);' +
            '}).apiPrepareSingleFileForAi(state.currentSubId, f.DriveFileID);' +
          '}, 30);' +
        '}' +

        'scanNext();' +
      '}' +

      'function genFb() {' +
        'if (!state.detail) return;' +
        'google.script.run.withSuccessHandler(function(r) {' +
          'if (r.success && r.feedback) {' +
            'document.getElementById("fbW").value = r.feedback.whatWentWell || "";' +
            'document.getElementById("fbI").value = r.feedback.areasForImprovement || "";' +
            'document.getElementById("fbG").value = r.feedback.goalsForNextAssessment || "";' +
          '} else { alert(r.message); }' +
        '}).apiGenerateFeedback(state.detail.assessmentId, state.currentSubId);' +
      '}' +

      'function downloadReportPdf() {' +
        'if (!state.currentSubId) return;' +
        'var btn = event.target; btn.disabled = true; btn.textContent = "Generating...";' +
        'google.script.run.withSuccessHandler(function(res) {' +
          'btn.disabled = false; btn.textContent = "Report PDF";' +
          'if (res.success) {' +
            'var a = document.createElement("a"); a.href = res.dataUrl; a.download = res.fileName;' +
            'document.body.appendChild(a); a.click(); document.body.removeChild(a);' +
          '} else { alert("PDF Error: " + res.message); }' +
        '}).withFailureHandler(function(err) {' +
          'btn.disabled = false; btn.textContent = "Report PDF";' +
          'alert("PDF generation error: " + err.message);' +
        '}).apiGenerateAssessmentPdf(state.currentSubId);' +
      '}' +
      '<\/script></body></html>';

    return head + headerHtml + bodyHtml + scriptHtml;
  }

  return {
    getAssessmentPdfReportHtml: getAssessmentPdfReportHtml,
    getCourseworkPickerHtml: getCourseworkPickerHtml,
    getClassBatchGradingHtml: getClassBatchGradingHtml,
    getVoiceDictationHtml: getVoiceDictationHtml,
    getAssessmentStudioHtml: getAssessmentStudioHtml
  };
})();


/* ============================================================================
 * SECTION 9: TOP-LEVEL ENTRY POINTS & RPC API ENDPOINTS
 * ============================================================================ */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Assessment System')
    .addItem('Open Assessment Studio (Fullscreen)', 'openAssessmentStudio')
    .addItem('Run Class AI Grading (All Students)', 'openClassAiGradingRunner')
    .addItem('Setup / Migrate Assessment System', 'setupOrMigrateAssessmentSystemUi')
    .addItem('Refresh Marking Sheet', 'refreshMarkingSheetUi')
    .addSeparator()
    .addSubMenu(ui.createMenu('Classroom')
      .addItem('Select Course and Assignment', 'openClassroomPicker')
      .addItem('Import / Refresh Submissions', 'importOrRefreshClassroomSubmissionsUi')
      .addItem('Sync Approved Assessment to Classroom', 'syncApprovedAssessmentToClassroomUi'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Gemini')
      .addItem('Set Gemini API Key', 'setGeminiApiKeyDialog')
      .addItem('Test Gemini Connection', 'testGeminiConnectionUi'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Administration')
      .addItem('Create Reassessment', 'createReassessmentUi')
      .addItem('View Error Log', 'viewErrorLogUi')
      .addItem('Verify Config & Sheets (Diagnostic)', 'verifySubmissionSystem')
      .addItem('Repair System & Re-Import (Diagnostic)', 'repairAssessmentSystemAndImport'))
    .addToUi();
}

function openAssessmentStudio() {
  var html = HtmlService.createHtmlOutput(HtmlTemplates.getAssessmentStudioHtml())
    .setWidth(1600)
    .setHeight(1000)
    .setTitle('AI Assessment Studio — Year 9 Jewellery Design');
  SpreadsheetApp.getUi().showModalDialog(html, 'AI Assessment Studio — Year 9 Jewellery Design');
}

function openClassAiGradingRunner() {
  var html = HtmlService.createHtmlOutput(HtmlTemplates.getClassBatchGradingHtml())
    .setWidth(680)
    .setHeight(420)
    .setTitle('Class-Wide AI Grading Runner');
  SpreadsheetApp.getUi().showModalDialog(html, 'Class-Wide AI Grading Runner');
}

function openClassroomPicker() {
  var html = HtmlService.createHtmlOutput(HtmlTemplates.getCourseworkPickerHtml())
    .setWidth(620)
    .setHeight(480)
    .setTitle('Select Classroom Course and Assignment');
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Classroom Course and Assignment');
}

function openVoiceDictation(criterionId) {
  var html = HtmlService.createHtmlOutput(HtmlTemplates.getVoiceDictationHtml(criterionId || 'C01'))
    .setWidth(420)
    .setHeight(360)
    .setTitle('Voice Dictation');
  SpreadsheetApp.getUi().showModalDialog(html, 'Dictate Observation Notes');
}

function setupOrMigrateAssessmentSystemUi() {
  var res = Sheets.setupOrMigrateAssessmentSystem();
  SpreadsheetApp.getUi().alert('Setup Status', res.message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function refreshMarkingSheetUi() {
  Sheets.verifyMarkingSheetFormulas(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Success', 'Formulas refreshed on 9DAT1 Marking.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function setGeminiApiKeyDialog() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Set Gemini API Key', 'Paste your Google AI Studio Gemini API Key', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() === ui.Button.OK) {
    var out = GeminiService.setGeminiApiKey(res.getResponseText());
    ui.alert(out.success ? 'Success' : 'Error', out.message, ui.ButtonSet.OK);
  }
}

function testGeminiConnectionUi() {
  var out = GeminiService.testGeminiConnection();
  SpreadsheetApp.getUi().alert(out.success ? 'Success' : 'Failed', out.message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/* This now displays the REAL result message from the fixed import function. */
function importOrRefreshClassroomSubmissionsUi() {
  var res = ClassroomService.importOrRefreshClassroomSubmissions();
  SpreadsheetApp.getUi().alert(
    res.success ? 'Classroom Import Complete' : 'Classroom Import Failed',
    res.message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function syncApprovedAssessmentToClassroomUi() {
  SpreadsheetApp.getUi().alert('Classroom Sync', 'Please open Assessment Studio and click "Sync" for the desired student.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function createReassessmentUi() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Create Reassessment', 'Enter parent SubmissionRecordID:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() === ui.Button.OK) {
    var out = AssessmentService.createReassessmentFromSubmission(res.getResponseText().trim());
    ui.alert(out.success ? 'Success' : 'Error', out.message, ui.ButtonSet.OK);
  }
}

function viewErrorLogUi() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(Config.SHEET_ERROR_LOG);
  if (sh) { sh.showSheet(); ss.setActiveSheet(sh); }
}

/* --- Diagnostic helpers (safe to keep permanently) --- */
function verifySubmissionSystem() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var report = {
    spreadsheet: ss.getName(),
    resolvedSubmissionsSheetName: Config.SHEET_SUBMISSIONS,
    submissionsSheetExists: !!ss.getSheetByName(Config.SHEET_SUBMISSIONS),
    submissionRows: ss.getSheetByName(Config.SHEET_SUBMISSIONS)
      ? Math.max(0, ss.getSheetByName(Config.SHEET_SUBMISSIONS).getLastRow() - 1)
      : 0,
    activeClassroomConfig: ClassroomService.getActiveClassroomConfig()
  };
  Logger.log(JSON.stringify(report, null, 2));
  SpreadsheetApp.getUi().alert('Diagnostic Report', JSON.stringify(report, null, 2), SpreadsheetApp.getUi().ButtonSet.OK);
  return report;
}

function repairAssessmentSystemAndImport() {
  Sheets.setupOrMigrateAssessmentSystem();
  var result = ClassroomService.importOrRefreshClassroomSubmissions();
  Logger.log(JSON.stringify(result, null, 2));
  SpreadsheetApp.getUi().alert(
    result.success ? 'Repair & Import Complete' : 'Repair & Import Failed',
    result.message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return result;
}

/* --- Client-Callable RPC Endpoints (google.script.run) --- */
function apiGetBootstrapData(id) { return AssessmentService.getAssessmentStudioBootstrapData(id); }
function apiGetSubmissionDetail(id) { return AssessmentService.getSubmissionForMarking(id); }
function apiSaveCriterionDraft(aId, sId, cId, st) { return AssessmentService.saveCriterionDraft(aId, sId, cId, st); }
function apiSaveAssessmentDraft(aId, sId, cMap, fb) { return AssessmentService.saveAssessmentDraft(aId, sId, cMap, fb); }
function apiRunInitialAiAssessment(sId) { return GeminiService.runInitialAiAssessment(sId); }
function apiGenerateFeedback(aId, sId) { return GeminiService.generateTeacherReviewedFeedback(aId, sId); }
function apiApproveAssessment(aId, sId, fb) { return AssessmentService.approveAssessment(aId, sId, fb); }
function apiLockAssessment(aId, sId) { return AssessmentService.lockAssessment(aId, sId); }
function apiCreateReassessment(sId) { return AssessmentService.createReassessmentFromSubmission(sId); }
function apiSyncToClassroom(aId, sId) { return ClassroomService.syncAssessmentToClassroom(aId, sId); }
function apiListTeacherCourses() { return ClassroomService.listTeacherCourses(); }
function apiListCourseWork(cid) { return ClassroomService.listCourseWork(cid); }
function apiSaveClassroomSelection(cid, cwid) { return ClassroomService.saveClassroomSelection(cid, cwid); }
function apiImportClassroomSubmissions() { return ClassroomService.importOrRefreshClassroomSubmissions(); }

function apiPrepareSingleFileForAi(submissionRecordId, driveFileId) {
  var allFiles = Sheets.getSheetDataAsObjects(Config.SHEET_SUBMISSION_FILES);
  for (var i = 0; i < allFiles.length; i++) {
    if (allFiles[i].SubmissionRecordID === submissionRecordId && allFiles[i].DriveFileID === driveFileId) {
      return DriveService.prepareSingleFileForAi(allFiles[i]);
    }
  }
  return { fileId: driveFileId, fileName: '', status: 'unreadable' };
}

function apiRunAiAssessmentWithFiles(submissionRecordId, preparedFiles) {
  return GeminiService.runAiAssessmentWithPreparedFiles(submissionRecordId, preparedFiles);
}

/* --- Batch Class AI Endpoints --- */
function apiGetClassSubmissions() {
  var subs = Sheets.getSheetDataAsObjects(Config.SHEET_SUBMISSIONS);
  var result = [];
  for (var i = 0; i < subs.length; i++) {
    if (subs[i].Status !== Config.STATUS.LOCKED && subs[i].Status !== Config.STATUS.APPROVED) {
      result.push(subs[i]);
    }
  }
  return result;
}
function apiGradeSingleStudentBatch(submissionRecordId) { return GeminiService.runInitialAiAssessment(submissionRecordId); }

/* --- PDF Export Endpoints --- */
function apiGenerateAssessmentPdf(submissionRecordId) {
  try {
    var detail = AssessmentService.getSubmissionForMarking(submissionRecordId);
    var htmlContent = HtmlTemplates.getAssessmentPdfReportHtml(detail);
    var blob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf');
    var cleanStudentName = detail.submission.StudentName.replace(/[^a-zA-Z0-9-]/g, '_');
    var fileName = cleanStudentName + '_Year9JewelleryAssessmentReport.pdf';
    blob.setName(fileName);
    return {
      success: true,
      fileName: fileName,
      dataUrl: 'data:application/pdf;base64,' + Utilities.base64Encode(blob.getBytes())
    };
  } catch (err) {
    return { success: false, message: 'Report PDF generation failed: ' + Utils.sanitizeError(err) };
  }
}
function apiExportSubmittedDriveFile(driveFileId) { return DriveService.exportDriveFileAsPdf(driveFileId); }
