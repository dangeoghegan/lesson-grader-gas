/* Teacher access and Gemini testing controls. Do not put credentials in Git,
 * HTML, query strings, logs or API responses. Apps Script deployments MUST
 * run as the accessing user and be restricted to the school's audience.
 * LESSON_GRADER_ALLOWED_EMAILS is a comma-separated Script Property set by
 * the owner in Apps Script Project Settings. Missing config fails closed. */
var WebAccess = (function () {
  var ALLOWLIST_PROPERTY = 'LESSON_GRADER_ALLOWED_EMAILS';

  function activeEmail() {
    try { return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase(); }
    catch (e) { return ''; }
  }

  function allowed() {
    var email = activeEmail();
    if (!email) return false;
    var csv = PropertiesService.getScriptProperties().getProperty(ALLOWLIST_PROPERTY) || '';
    return csv.split(',').some(function (entry) { return entry.trim().toLowerCase() === email; });
  }

  function requireTeacher() {
    if (!allowed()) throw new Error('Access denied. Sign in with an allowed teacher account, or ask the owner to configure teacher access.');
    return activeEmail();
  }

  function status() {
    return {
      success: true,
      data: {
        authorised: allowed(),
        email: activeEmail(),
        keyConfigured: allowed() && !!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')
      }
    };
  }

  function setKey(value) {
    requireTeacher();
    var key = String(value || '').trim();
    if (key.length < 20 || key.length > 512 || /\s/.test(key)) {
      return { success: false, message: 'Enter a valid Gemini API key without spaces.' };
    }
    PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
    return { success: true, message: 'Key saved to Apps Script Script Properties. It is not returned to your browser.' };
  }

  function testKey() {
    requireTeacher();
    /* No student files, names or rubric data in this request. The same
       transport (x-goog-api-key header) is used for actual AI proposals. */
    try {
      var result = GeminiService.testGeminiConnection();
      return result && result.success ? { success: true, message: result.message } :
        { success: false, message: result && result.message || 'Gemini connection could not be verified.' };
    } catch (e) {
      return { success: false, message: 'Gemini connection could not be verified. Check the key, API access and quotas.' };
    }
  }

  return { requireTeacher: requireTeacher, status: status, setKey: setKey, testKey: testKey };
})();

/* This endpoint returns only setup state; never the stored key or allowlist. */
function apiWebAuthStatus() { return WebAccess.status(); }
function apiWebSetGeminiKey(key) { return WebAccess.setKey(key); }
function apiWebTestGeminiKey() { return WebAccess.testKey(); }
