const FALLBACK_CURATORS = [
  { id: "fallback-1", display_name: "Tanmay Mukherji", email: "tanmay@greenruraleconomy.in" },
  { id: "fallback-2", display_name: "Phaneesh K", email: "agri@greenruraleconomy.in" },
  { id: "fallback-3", display_name: "Swati Singh", email: "solution@greenruraleconomy.in" },
  { id: "fallback-4", display_name: "Shaifali Nagar", email: "help@greenruraleconomy.in" },
];

const state = {
  view: "overview",
  sharedFormMode: new URLSearchParams(window.location.search).get("sharedForm") === "solution"
    ? "solution"
    : new URLSearchParams(window.location.search).get("sharedForm") === "need"
      ? "need-intake"
      : "",
  selectedNeedId: null,
  queueNeedsScrollIntoView: false,
  overviewPage: 1,
  matchPage: 1,
  matchCache: new Map(),
  mapGeocodeCache: new Map(),
  mapplsSdkLoaded: false,
  mapplsLoadPromise: null,
  caseMapRequestToken: 0,
  caseMap: null,
  caseMapMarkers: [],
  activeMapGroupKey: "",
  lgdSearchToken: 0,
  showClosedNeeds: false,
  overviewFilters: {
    metric: [],
    pipeline: [],
    curator: [],
    state: [],
    category: [],
  },
  userToken: localStorage.getItem("gre-mis-user-token") || localStorage.getItem("gre-mis-admin-token") || "",
  adminToken: "",
  userSession: null,
  adminSession: null,
  submissionReviewId: "",
  localSolutionReviewId: "",
  solutionTags: [],
  solutionGeographies: [],
  adminUserRoleTab: "admin",
  localSolutionFilters: {
    search: "",
    provider: "all",
  },
    puterModelsLoaded: false,
    puterModels: [],
    puterRecommendations: {},
    filters: {
    status: "all",
    curator: "all",
    state: "all",
    search: "",
  },
  data: {
      curators: [],
      traders: [],
      offeringMaster: {
        valuechains: [],
        applications: [],
        tags: [],
        languages: [],
        geographies: [],
      },
      options: [],
      needs: [],
      needUpdates: [],
      pendingNeeds: [],
      pendingUpdates: [],
      pendingFormSubmissions: [],
      aiReviewNeeds: [],
      users: [],
      localSolutions: [],
  },
};

const OFFERING_CATEGORY_OPTIONS = [
  { value: "Service offerings", label: "Service" },
  { value: "Product offerings", label: "Product" },
  { value: "Knowledge offerings", label: "Knowledge" },
];

const OFFERING_TYPE_OPTIONS = {
  "Product offerings": [
    { value: "Machinery", label: "Machinery" },
    { value: "Plant setup", label: "Plant Setup" },
    { value: "Product bought", label: "Product/Raw Material Bought" },
    { value: "Raw material", label: "Raw Material Supply" },
  ],
  "Service offerings": [
    { value: "Training", label: "Training" },
    { value: "Consulting", label: "Consulting / Mentoring" },
    { value: "Financial support", label: "Financial Support" },
    { value: "Market support", label: "Market Support" },
    { value: "Tech transfer", label: "Technology Transfer" },
  ],
  "Knowledge offerings": [
    { value: "Blogs", label: "Blogs" },
    { value: "Market reports", label: "Market Reports" },
    { value: "Sop manuals", label: "SOP / Manuals" },
    { value: "Videos", label: "Videos" },
  ],
};

const DEFAULT_LANGUAGE_OPTIONS = [
  "ENG",
  "HIN",
  "ASSAMESE",
  "BENGALI",
  "BODO",
  "DOGRI",
  "GUJARATI",
  "HINDI",
  "KANNADA",
  "KASHMIRI",
  "KONKANI",
  "MAITHILI",
  "MALAYALAM",
  "MANIPURI",
  "MARATHI",
  "NEPALI",
  "ODIA",
  "PUNJABI",
  "SANSKRIT",
  "SANTALI",
  "SINDHI",
  "TAMIL",
  "TELUGU",
  "URDU",
];
const MAX_EMBEDDED_FILE_BYTES = 5 * 1024 * 1024;

const MATCH_STOPWORDS = new Set([
  "about", "across", "after", "also", "been", "being", "between", "could", "does", "from", "have", "into",
  "more", "need", "needs", "only", "other", "problem", "seeker", "should", "solution", "solutions", "their",
  "there", "these", "they", "this", "through", "under", "want", "where", "which", "with", "would", "rural",
  "green", "economy", "help", "looking", "support", "required", "request",
]);

const SIX_M_LABELS = ["Manpower", "Method", "Machine", "Material", "Market", "Money"];
const GENERIC_THEMATIC_TERMS = new Set([
    "infrastructure",
    "technology",
    "vendor",
    "connect collaborate",
    "business",
    "business consultation",
    "business mentoring",
    "business development",
    "training",
    "capacity building",
    "advisory",
    "consulting",
    "consultancy",
    "entrepreneurship",
    "business training",
    "finance",
    "funding",
    "investments",
  ]);
const DOMAIN_MATCH_STOPWORDS = new Set([
  "project",
  "projects",
  "village",
  "villages",
  "support",
  "farmers",
  "farmer",
  "guidance",
  "implement",
  "implementation",
  "funding",
  "under",
  "need",
  "needs",
  "management",
  "going",
]);
const SIX_M_RULES = [
  { label: "Manpower", patterns: ["training", "capacity building"] },
  { label: "Method", patterns: ["consulting", "consultancy", "business consultation", "mentoring", "business mentoring", "technology transfer", "technology", "video", "videos", "manual", "manuals", "sop", "sops", "blog", "blogs", "connect collaborate"] },
  { label: "Machine", patterns: ["machinery", "machine", "plant setup"] },
  { label: "Material", patterns: ["raw material", "raw materials", "material supply"] },
  { label: "Market", patterns: ["products bought", "market support", "market reports", "market report", "business development", "branding", "packaging"] },
  { label: "Money", patterns: ["financial support", "finance", "funding", "investment", "investments", "credit"] },
];
const SERVICE_PHRASES = [
  "market linkage",
  "capacity building",
  "knowledge product",
  "consulting",
  "consultancy",
  "training",
  "advisory",
  "service",
  "services",
  "product",
  "products",
  "finance",
  "financing",
  "credit",
  "equipment",
  "machinery",
  "technology",
  "knowledge",
];
const NEED_THEME_RULES = [
  { label: "dairy", patterns: ["dairy", "milk", "milching", "cow", "cows", "livestock", "fodder"] },
  { label: "solar", patterns: ["solar", "street light", "street lights", "streetlight"] },
  { label: "wild mango", patterns: ["wild mango", "mango", "amchur", "ntfp", "forest produce"] },
  { label: "goatery", patterns: ["goat", "goatery", "goat farming"] },
  { label: "poultry", patterns: ["poultry", "chicken", "broiler", "layer"] },
  { label: "fisheries", patterns: ["fishery", "fisheries", "aquaculture", "fish farming"] },
  { label: "makhana", patterns: ["makhana", "fox nut"] },
  { label: "soap", patterns: ["soap", "soaps", "detergent"] },
  { label: "branding", patterns: ["branding", "logo", "packaging", "marketplace onboarding"] },
  { label: "business planning", patterns: ["business plan", "business development plan", "costing", "pricing"] },
];
const INDIA_FALLBACK_CENTER = { lat: 22.9734, lng: 78.6569 };
const STATE_MAP_CENTERS = {
  "andaman and nicobar islands": { lat: 11.7401, lng: 92.6586 },
  "andhra pradesh": { lat: 15.9129, lng: 79.74 },
  "arunachal pradesh": { lat: 28.218, lng: 94.7278 },
  assam: { lat: 26.2006, lng: 92.9376 },
  bihar: { lat: 25.0961, lng: 85.3131 },
  chandigarh: { lat: 30.7333, lng: 76.7794 },
  chattisgarh: { lat: 21.2787, lng: 81.8661 },
  chhattisgarh: { lat: 21.2787, lng: 81.8661 },
  "dadra and nagar haveli and daman and diu": { lat: 20.3974, lng: 72.8328 },
  delhi: { lat: 28.7041, lng: 77.1025 },
  goa: { lat: 15.2993, lng: 74.124 },
  gujarat: { lat: 22.2587, lng: 71.1924 },
  haryana: { lat: 29.0588, lng: 76.0856 },
  "himachal pradesh": { lat: 31.1048, lng: 77.1734 },
  "jammu and kashmir": { lat: 33.7782, lng: 76.5762 },
  jharkhand: { lat: 23.6102, lng: 85.2799 },
  karnataka: { lat: 15.3173, lng: 75.7139 },
  kerala: { lat: 10.8505, lng: 76.2711 },
  ladakh: { lat: 34.1526, lng: 77.5771 },
  lakshadweep: { lat: 10.5667, lng: 72.6417 },
  "madhya pradesh": { lat: 22.9734, lng: 78.6569 },
  maharashtra: { lat: 19.7515, lng: 75.7139 },
  manipur: { lat: 24.6637, lng: 93.9063 },
  meghalaya: { lat: 25.467, lng: 91.3662 },
  mizoram: { lat: 23.1645, lng: 92.9376 },
  nagaland: { lat: 26.1584, lng: 94.5624 },
  odisha: { lat: 20.9517, lng: 85.0985 },
  orissa: { lat: 20.9517, lng: 85.0985 },
  puducherry: { lat: 11.9416, lng: 79.8083 },
  punjab: { lat: 31.1471, lng: 75.3412 },
  rajasthan: { lat: 27.0238, lng: 74.2179 },
  sikkim: { lat: 27.533, lng: 88.5122 },
  "tamil nadu": { lat: 11.1271, lng: 78.6569 },
  telangana: { lat: 18.1124, lng: 79.0193 },
  tripura: { lat: 23.9408, lng: 91.9882 },
  "uttar pradesh": { lat: 26.8467, lng: 80.9462 },
  uttarakhand: { lat: 30.0668, lng: 79.0193 },
  "west bengal": { lat: 22.9868, lng: 87.855 },
};

const PIPELINE_SEGMENTS = [
  {
    id: "new",
    label: "New",
    note: "Fresh approved needs waiting for assignment.",
    match: (need) => need.status === "New",
  },
  {
    id: "accepted",
    label: "Accepted",
    note: "Needs allocated but still waiting for movement.",
    match: (need) => need.status === "Accepted",
  },
  {
    id: "in_progress",
    label: "In progress",
    note: "Needs in live curation and matching flow.",
    match: (need) => need.status === "In progress",
  },
  {
    id: "closed",
    label: "Closed",
    note: "Resolved or completed needs.",
    match: (need) => need.status === "Closed",
  },
  {
    id: "stuck",
    label: "Blocked / Stalled",
    note: "Needs requiring escalation.",
    match: (need) => ["Blocked", "Stalled"].includes(need.internal_status),
  },
  {
    id: "broadcast",
    label: "Broadcast Suggested",
    note: "Needs that could benefit from wider ecosystem response.",
    match: (need) => Boolean(need.demand_broadcast_needed),
  },
];

function esc(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function byId(id) {
  return document.getElementById(id);
}

function escAttr(value) {
  return esc(value).replaceAll("'", "&#39;");
}

function toast(message) {
  window.alert(message);
}

function setText(id, message = "") {
  const el = byId(id);
  if (el) el.textContent = message;
}

function safeAsync(handler) {
  return async (...args) => {
    try {
      await handler(...args);
    } catch (error) {
      console.error(error);
      const loginStatus = byId("loginStatus");
      const sessionStatus = byId("sessionStatus");
      const formId = args[0]?.target?.id;
      if (loginStatus && formId === "adminLoginForm") {
        loginStatus.textContent = error?.message || "Admin login failed.";
      } else if (formId === "userLoginForm") {
        setText("userLoginStatus", error?.message || "Sign in failed.");
      } else if (formId === "registerUserForm") {
        setText("registerStatus", error?.message || "Registration failed.");
      } else if (formId === "requestResetForm" || formId === "resetPasswordForm") {
        setText("resetStatus", error?.message || "Password reset flow failed.");
      } else if (formId === "changePasswordForm") {
        setText("changePasswordStatus", error?.message || "Password change failed.");
      } else if (sessionStatus) {
        sessionStatus.textContent = error?.message || "Something went wrong.";
      }
      toast(error?.message || "Something went wrong. Please try again.");
    }
  };
}

function normalizeText(value) {
  return String(value ?? "").trim();
}

function syncSessionState(user, token = state.userToken) {
  state.userSession = user || null;
  state.userToken = token || "";
  const isAdmin = user?.role === "admin";
  state.adminToken = isAdmin ? state.userToken : "";
  state.adminSession = isAdmin
    ? { username: user?.username || "admin", email: user?.email || "" }
    : null;

  if (state.userToken) {
    localStorage.setItem("gre-mis-user-token", state.userToken);
    if (isAdmin) localStorage.setItem("gre-mis-admin-token", state.userToken);
    else localStorage.removeItem("gre-mis-admin-token");
  } else {
    localStorage.removeItem("gre-mis-user-token");
    localStorage.removeItem("gre-mis-admin-token");
  }
}

function isLoggedIn() {
  return Boolean(state.userSession);
}

function isSharedFormMode() {
  return state.sharedFormMode === "solution" || state.sharedFormMode === "need-intake";
}

function isAdminUser() {
  return state.userSession?.role === "admin";
}

function isCuratorUser() {
  return state.userSession?.role === "curator";
}

function canAccessOperationsDesk() {
  return isLoggedIn();
}

function canSeeCurationDetails() {
  return isAdminUser() || isCuratorUser();
}

function parseArray(value) {
  if (Array.isArray(value)) return value.filter(Boolean).map((item) => normalizeText(item)).filter(Boolean);
  if (typeof value === "string") {
    return value
      .replace(/^[\[\{"]+|[\]\}"]+$/g, "")
      .split(/[;,|]/)
      .map((item) => normalizeText(item))
      .filter(Boolean);
  }
  return [];
}

function parseDelimitedList(value, delimiter = ",") {
  return normalizeText(value)
    .split(delimiter)
    .map((item) => normalizeText(item))
    .filter(Boolean);
}

function parseCommaList(value) {
  return parseDelimitedList(value, ",");
}

function ensureList(value) {
  return Array.isArray(value) ? value : [];
}

function parseNumber(value, fallback = 0) {
  const normalized = String(value ?? "").replace(/[^0-9.-]/g, "").trim();
  if (!normalized) return fallback;
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function parseWorkbookDate(value) {
  const text = normalizeText(value);
  if (!text) return null;
  const parts = text.match(/^(\d{1,2})-(\d{1,2})-(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (parts) {
    const [, day, month, year, hour = "0", minute = "0", second = "0"] = parts;
    return new Date(Date.UTC(Number(year), Number(month) - 1, Number(day), Number(hour), Number(minute), Number(second))).toISOString();
  }
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed.toISOString();
}

function parseWorkbookBoolean(value) {
  const text = normalizeText(value).toLowerCase();
  if (!text || text === "null") return null;
  if (["yes", "true", "1"].includes(text)) return true;
  if (["no", "false", "0"].includes(text)) return false;
  return null;
}

async function parseInboundWorkbookFile(file) {
  if (!file) throw new Error("Choose an inbound workbook first.");
  if (!window.XLSX) throw new Error("Workbook parser is not available.");
  const buffer = await file.arrayBuffer();
  const workbook = window.XLSX.read(buffer, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = window.XLSX.utils.sheet_to_json(sheet, { defval: "" });
  return rows.map((row) => ({
    request_id: normalizeText(row["Request Id"]),
    organization_name: normalizeText(row["Seeker Organisation"]),
    website: normalizeText(row.Website),
    contact_person: normalizeText(row["Contact Person"]),
    designation: normalizeText(row.Designation),
    seeker_phone: normalizeText(row["Phone Number"]),
    seeker_email: normalizeText(row.Email).toLowerCase(),
    requested_on: parseWorkbookDate(row["Requested On"]),
    problem_statement: normalizeText(row.Request),
    state: normalizeText(row["Solution Needed in State"]),
    district: normalizeText(row["Solution Needed in District"]),
    status: normalizeText(row.Status),
    internal_status: normalizeText(row["Internal Status"]),
    curator_name: normalizeText(row["Curator Assigned"]),
    curation_call_date: parseWorkbookDate(row["Curation Call Date"])?.slice(0, 10) || null,
    curation_age_days: parseNumber(row["Curation Age"], 0),
    curation_call_details: normalizeText(row["Curation Call Details"]),
    curated_need: parseArray(String(row["Curated Need of Service Seeker"] || "").replaceAll(",", ";")),
    demand_broadcast_needed: parseWorkbookBoolean(row["Demand Broadcast Needed"]),
    solutions_shared_count: parseNumber(row["Solutions Shared Count"], 0),
    solutions_shared: normalizeText(row["Solutions Shared"]),
    invited_providers_count: parseNumber(row["Invited Providers Count"], 0),
    invited_providers: normalizeText(row["Invited Providers"]),
  })).filter((row) => row.request_id);
}

function downloadBase64Workbook(download) {
  if (!download?.base64 || !download?.fileName) throw new Error("Download payload is missing.");
  const mimeType = download.mimeType || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const binary = atob(download.base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  const blob = new Blob([bytes], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = download.fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function formatDate(value) {
  if (!value) return "Not set";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? String(value) : date.toLocaleString("en-IN", { dateStyle: "medium", timeStyle: "short" });
}

function parseCoordinate(value) {
  if (value === null || value === undefined || value === "") return null;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : null;
}

function slugStatus(value) {
  return normalizeText(value).toLowerCase().replaceAll(/\s+/g, "-");
}

function badgeTone(value) {
  const text = normalizeText(value).toLowerCase();
  if (["closed", "approved", "accepted", "connection made"].includes(text)) return "good";
  if (["blocked", "stalled", "rejected", "overdue"].includes(text)) return "bad";
  if (["pending_admin", "need solution providers", "new", "pending"].includes(text)) return "warn";
  return "info";
}

function formatAiEnrichmentStatus(value) {
  const text = normalizeText(value);
  const lowered = text.toLowerCase();
  if (!text) return "not enriched";
  if (lowered.includes("not configured")) return "Legacy AI refresh unavailable";
  if (lowered === "rules_only") return "Rule-based enrichment";
  return text;
}

function tokenizeText(value, minimumLength = 4) {
  return normalizeText(value)
    .toLowerCase()
    .split(/[^a-z0-9]+/)
    .map((token) => token.trim())
    .filter((token) => token.length >= minimumLength && !MATCH_STOPWORDS.has(token));
}

function uniq(items) {
  return [...new Set(items.filter(Boolean))];
}

function getEffectiveNeedThematicArea(need) {
  return normalizeText(need?.override_thematic_area || need?.ai_thematic_area);
}

function getEffectiveNeedApplicationArea(need) {
  return normalizeText(need?.override_application_area || need?.ai_application_area);
}

function getEffectiveNeedKind(need) {
  return normalizeText(need?.override_need_kind || need?.ai_need_kind);
}

function getEffectiveNeedServiceKind(need) {
  return normalizeText(need?.override_service_kind || need?.ai_service_kind);
}

function getEffectiveNeedKeywords(need) {
  return uniq([
    ...parseArray(need?.override_keywords).map((item) => item.toLowerCase()),
    ...parseArray(need?.ai_keywords).map((item) => item.toLowerCase()),
  ]);
}

function getEffectiveNeed6MSignals(need) {
  return uniq([
    ...parseArray(need?.override_6m_signals),
    ...parseArray(need?.ai_6m_signals),
  ]);
}

function clipText(value, length = 180) {
  const text = normalizeText(value);
  return text.length > length ? `${text.slice(0, length)}...` : text;
}

function extractUrls(value) {
  return String(value || "").match(/https?:\/\/[^\s)]+/gi) || [];
}

function stripUrls(value) {
  return String(value || "").replace(/https?:\/\/[^\s)]+/gi, "").replace(/\s{2,}/g, " ").trim();
}

function extractSharedSolutionHints(value) {
  const text = normalizeText(value);
  if (!text) return { phrases: [], ids: { offeringIds: [], solutionIds: [], traderIds: [] } };
  const urls = extractUrls(text);
  const ids = { offeringIds: [], solutionIds: [], traderIds: [] };

  urls.forEach((url) => {
    const offeringId = url.match(/[?&]productSkuId=(\d+)/i)?.[1];
    const solutionId = url.match(/[?&]solutionId=(\d+)/i)?.[1];
    const traderId = url.match(/[?&]traderId=(\d+)/i)?.[1];
    if (offeringId) ids.offeringIds.push(offeringId);
    if (solutionId) ids.solutionIds.push(solutionId);
    if (traderId) ids.traderIds.push(traderId);
  });

  const phrases = [];
  text.split(/\n+/).forEach((line) => {
    const cleaned = line.replace(/https?:\/\/[^\s)]+/gi, "").trim();
    const parts = cleaned.split(":").map((item) => normalizeText(item));
    if (parts.length >= 2 && parts[1]) {
      phrases.push(parts[1].toLowerCase());
    }
  });

  return {
    phrases: uniq(phrases),
    ids: {
      offeringIds: uniq(ids.offeringIds),
      solutionIds: uniq(ids.solutionIds),
      traderIds: uniq(ids.traderIds),
    },
  };
}

function extractPuterText(response) {
  if (typeof response === "string") return response.trim();
  const candidates = [
    response?.message?.content,
    response?.content,
    response?.text,
    response?.result,
    response?.message,
  ];
  for (const candidate of candidates) {
    if (typeof candidate === "string" && candidate.trim()) return candidate.trim();
    if (Array.isArray(candidate)) {
      const joined = candidate
        .map((item) => {
          if (typeof item === "string") return item;
          if (typeof item?.text === "string") return item.text;
          if (typeof item?.content === "string") return item.content;
          return "";
        })
        .filter(Boolean)
        .join("\n")
        .trim();
      if (joined) return joined;
    }
    if (candidate && typeof candidate === "object") {
      const nested = String(candidate.text || candidate.content || candidate.message || "").trim();
      if (nested) return nested;
    }
  }
  return JSON.stringify(response || {});
}

function stripCodeFences(value) {
  return String(value || "").replace(/^```(?:json)?/i, "").replace(/```$/i, "").trim();
}

function parseJsonObject(text) {
  const cleaned = stripCodeFences(text);
  try {
    return JSON.parse(cleaned);
  } catch {
    const match = cleaned.match(/\{[\s\S]*\}/);
    if (!match) throw new Error("AI response did not contain valid JSON.");
    return JSON.parse(match[0]);
  }
}

function normalizePuterModelEntries(items) {
  return (Array.isArray(items) ? items : [])
    .map((item) => {
      if (typeof item === "string") return { id: item, name: item };
      const id = String(item?.id || item?.model || item?.name || "").trim();
      const name = String(item?.name || item?.label || item?.id || id).trim();
      return id ? { id, name } : null;
    })
    .filter(Boolean);
}

function extractProblemPhrases(value) {
  const tokens = String(value || "")
    .toLowerCase()
    .split(/[^a-z0-9]+/)
    .map((token) => token.trim())
    .filter((token) => token.length >= 4 && !MATCH_STOPWORDS.has(token));
  const phrases = [];
  for (let index = 0; index < tokens.length - 1; index += 1) {
    if (!MATCH_STOPWORDS.has(tokens[index]) && !MATCH_STOPWORDS.has(tokens[index + 1])) {
      phrases.push(`${tokens[index]} ${tokens[index + 1]}`);
    }
    if (index < tokens.length - 2) {
      if (!MATCH_STOPWORDS.has(tokens[index + 2])) {
        phrases.push(`${tokens[index]} ${tokens[index + 1]} ${tokens[index + 2]}`);
      }
    }
  }
  return uniq(phrases).slice(0, 18);
}

function extractCategoryParts(value) {
  const raw = normalizeText(value).toLowerCase();
  if (!raw) return { raw: "", thematic: "", service: "" };
  const matchedService = [...SERVICE_PHRASES].sort((a, b) => b.length - a.length).find((phrase) => raw.endsWith(phrase));
  if (!matchedService) return { raw, thematic: raw, service: "" };
  const thematic = raw.slice(0, raw.length - matchedService.length).trim();
  return {
    raw,
    thematic: thematic || raw,
    service: matchedService,
  };
}

function extractNeedRuleThemes(need) {
  const text = [
    normalizeText(need?.problem_statement),
    normalizeText(need?.curation_notes).replace(/\bnull\b/gi, " "),
    parseArray(need?.curated_need).join(" "),
  ].join(" | ").toLowerCase();
  return uniq(
    NEED_THEME_RULES
      .filter((rule) => rule.patterns.some((pattern) => text.includes(pattern)))
      .map((rule) => rule.label),
  );
}

function getNeedThemeSignals(need) {
  const themes = [];
  parseArray(need?.curated_need).forEach((item) => {
    const parts = extractCategoryParts(item);
    if (parts.thematic) themes.push(parts.thematic);
  });
  return uniq(themes);
}

function getNeedSixMSignals(need) {
  const explicit = uniq([
    ...parseArray(need?.override_6m_signals),
    ...parseArray(need?.six_m_signals),
  ]);
  if (explicit.length) return uniq(explicit);

  const haystack = [
    ...parseArray(need?.curated_need),
    normalizeText(need?.problem_statement),
    normalizeText(need?.curation_notes),
  ]
    .join(" | ")
    .toLowerCase();

  return SIX_M_RULES
    .filter((rule) => rule.patterns.some((pattern) => haystack.includes(pattern)))
    .map((rule) => rule.label);
}

function categoryFilterMatches(category, need) {
  if (!need) return false;
  const normalized = normalizeText(category).toLowerCase();
  const curated = parseArray(need.curated_need);
  if (SIX_M_LABELS.map((item) => item.toLowerCase()).includes(normalized)) {
    return getNeedSixMSignals(need).map((item) => item.toLowerCase()).includes(normalized);
  }
  return getNeedThemeSignals(need).includes(normalized) || curated.some((item) => item.toLowerCase() === normalized);
}

function buildNeedMatchProfile(need) {
  const categories = uniq(parseArray(need.curated_need).map((item) => item.toLowerCase()));
  const categoryParts = categories.map((item) => extractCategoryParts(item));
  const problemThemeSignals = extractNeedRuleThemes(need);
  const ruleThemes = uniq(parseArray(need.rule_thematic_hints).map((item) => item.toLowerCase()).filter(Boolean));
  const ruleServices = uniq(parseArray(need.rule_service_hints).map((item) => item.toLowerCase()).filter(Boolean));
  const ruleKeywords = uniq(parseArray(need.rule_keywords).map((item) => item.toLowerCase()).filter(Boolean));
  const ruleNeedKind = normalizeText(need.rule_need_kind).toLowerCase();
  const aiKeywords = getEffectiveNeedKeywords(need);
  const effectiveThematicArea = getEffectiveNeedThematicArea(need).toLowerCase();
  const effectiveApplicationArea = getEffectiveNeedApplicationArea(need).toLowerCase();
  const aiSignals = uniq([
    effectiveThematicArea,
    effectiveApplicationArea,
    getEffectiveNeedKind(need).toLowerCase(),
    getEffectiveNeedServiceKind(need).toLowerCase(),
    ...aiKeywords,
  ].filter(Boolean));
  const aiNeedKind = getEffectiveNeedKind(need).toLowerCase();
  const aiServiceKind = getEffectiveNeedServiceKind(need).toLowerCase();
  const categoryThematicAreas = uniq(
    categoryParts
      .map((item) => item.thematic)
      .filter((item) => item && !GENERIC_THEMATIC_TERMS.has(item)),
  );
  const serviceTerms = uniq([...categoryParts.map((item) => item.service).filter(Boolean), ...ruleServices]);
  const problemTokens = uniq([...ruleKeywords, ...tokenizeText(need.problem_statement, 7)]);
  const notesTokens = uniq([...ruleKeywords, ...tokenizeText(need.curation_notes, 5)]);
  const geographyTokens = tokenizeText(`${need.state || ""} ${need.district || ""}`, 3);
  const serviceTokens = serviceTerms.flatMap((item) => tokenizeText(item, 3));
  const sharedSolutionHints = extractSharedSolutionHints(need.curation_notes);
  const problemPhrases = extractProblemPhrases(need.problem_statement);
  const explicitThemeTokens = uniq([
    ...ruleThemes.flatMap((item) => tokenizeText(item, 3)),
    ...ruleKeywords,
    ...tokenizeText(need.problem_statement, 4),
    ...tokenizeText(need.curation_notes, 4),
  ]).slice(0, 16);
  const domainFocusTokens = uniq([
    ...problemThemeSignals.flatMap((item) => tokenizeText(item, 3)),
    ...ruleThemes.flatMap((item) => tokenizeText(item, 3)),
    ...ruleKeywords,
    ...explicitThemeTokens,
  ])
    .filter((token) => !DOMAIN_MATCH_STOPWORDS.has(token));
  const thematicAreas = uniq([
    ...sharedSolutionHints.phrases,
    ...problemThemeSignals,
    ...ruleThemes,
    ...categoryThematicAreas,
    ...aiSignals,
    ...(categoryThematicAreas.length ? [] : problemPhrases.slice(0, 8)),
  ]);
  const categoryTokens = thematicAreas.flatMap((item) => tokenizeText(item, 3));
  const primaryTerms = uniq([...categoryTokens, ...serviceTokens, ...problemTokens.slice(0, 8), ...notesTokens.slice(0, 4)]);
  const phrases = uniq(
    thematicAreas
      .filter((item) => item.includes(" "))
      .concat(
        normalizeText(need.problem_statement)
          .split(/[.;]/)
          .map((item) => item.trim())
          .filter((item) => item.split(/\s+/).length <= 5 && item.length >= 12)
          .slice(0, 2)
          .map((item) => item.toLowerCase()),
      ),
  );
  const requiresServiceMatch =
    !ruleThemes.length &&
    !categoryThematicAreas.length &&
    (Boolean(aiServiceKind || ruleServices[0]) || (aiNeedKind || ruleNeedKind) === "service");
  const resolvedNeedKind = (
    aiNeedKind ||
    ((ruleThemes.length && ruleServices.length && !["product", "knowledge", "finance"].includes(ruleNeedKind)) ? "service" : "") ||
    ruleNeedKind ||
    (ruleServices.length ? "service" : "")
  );
  const preferredOfferingKinds = ["service", "product", "knowledge"];

  return {
    categories,
    categoryParts,
    categoryThematicAreas,
    problemThemeSignals,
    thematicAreas,
    serviceTerms,
    requiresServiceMatch,
    categoryTokens: uniq(categoryTokens),
    serviceTokens: uniq(serviceTokens),
    problemTokens,
    notesTokens,
    geographyTokens,
    phrases,
    primaryTerms,
    explicitThemeTokens,
    domainFocusTokens,
    sharedSolutionHints,
    resolvedNeedKind,
    preferredOfferingKinds,
    hasStrongTheme: Boolean(problemThemeSignals.length || ruleThemes.length || categoryThematicAreas.length || effectiveThematicArea),
  };
}

function getOfferingKind(offering) {
  const offeringGroup = normalizeText(offering?.offering_group).toLowerCase();
  const offeringType = normalizeText(offering?.offering_type).toLowerCase();
  const category = normalizeText(offering?.offering_category).toLowerCase();
  const aiKind = normalizeText(offering?.ai_offering_kind).toLowerCase();
  if (offeringGroup.includes("knowledge")) return "knowledge";
  if (offeringGroup.includes("service")) return "service";
  if (offeringGroup.includes("product")) return "product";
  if (offeringType.includes("manual") || offeringType.includes("video") || offeringType.includes("sop")) return "knowledge";
  if (aiKind) return aiKind;
  if (category.includes("service")) return "service";
  if (category.includes("product")) return "product";
  if (category.includes("knowledge")) return "knowledge";
  return "";
}

function getPipelineSegmentNeeds(segmentId) {
  const segment = PIPELINE_SEGMENTS.find((item) => item.id === segmentId) || PIPELINE_SEGMENTS[0];
  return state.data.needs
    .filter(segment.match)
    .sort((a, b) => Number(b.curation_age_days || 0) - Number(a.curation_age_days || 0));
}

function isOverviewFocus(kind, id) {
  return (state.overviewFilters?.[kind] || []).includes(id);
}

function setOverviewFocus(kind, id) {
  const active = new Set(state.overviewFilters?.[kind] || []);
  if (active.has(id)) {
    active.delete(id);
  } else {
    active.add(id);
  }
  state.overviewFilters[kind] = [...active];
  state.overviewPage = 1;
}

function sortNeedsByUrgency(needs) {
  return [...needs].sort((a, b) => Number(b.curation_age_days || 0) - Number(a.curation_age_days || 0));
}

async function ensureMapplsSdk() {
  if (window.mappls?.Map && state.mapplsSdkLoaded) return true;
  if (state.mapplsLoadPromise) return state.mapplsLoadPromise;
  const key = String(window.APP_CONFIG?.MAPPLS_MAP_KEY || window.APP_CONFIG?.MAPMYINDIA_MAP_KEY || "").trim();
  if (!key) return false;

  if (!document.getElementById("mappls-web-sdk-css")) {
    const link = document.createElement("link");
    link.id = "mappls-web-sdk-css";
    link.rel = "stylesheet";
    link.href = "https://apis.mappls.com/vector_map/assets/v3.5/mappls-glob.css";
    document.head.appendChild(link);
  }

  const urls = [
    `https://sdk.mappls.com/map/sdk/web?v=3.0&access_token=${encodeURIComponent(key)}`,
    `https://sdk.mappls.com/map/sdk/web?v=3.0&layer=vector&access_token=${encodeURIComponent(key)}`,
    `https://apis.mappls.com/advancedmaps/api/${encodeURIComponent(key)}/map_sdk?layer=vector&v=3.0`,
  ];

  state.mapplsLoadPromise = (async () => {
    for (const src of urls) {
      try {
        await new Promise((resolve, reject) => {
          document.querySelectorAll("script[data-mappls-sdk='true']").forEach((node) => node.remove());
          const script = document.createElement("script");
          script.src = src;
          script.async = true;
          script.defer = true;
          script.dataset.mapplsSdk = "true";
          script.onload = () => (window.mappls?.Map ? resolve(true) : reject(new Error("Mappls SDK unavailable")));
          script.onerror = reject;
          document.head.appendChild(script);
        });
        state.mapplsSdkLoaded = true;
        return true;
      } catch (error) {
        console.warn("Mappls SDK load failed for", src, error);
      }
    }
    state.mapplsSdkLoaded = false;
    return false;
  })();

  const loaded = await state.mapplsLoadPromise;
  if (!loaded) state.mapplsLoadPromise = null;
  return loaded;
}

function getNeedMapLocationLabel(need) {
  return [normalizeText(need.district), normalizeText(need.state)].filter(Boolean).join(", ") || "Location pending";
}

function hashText(value) {
  const text = normalizeText(value);
  let hash = 0;
  for (let index = 0; index < text.length; index += 1) {
    hash = ((hash << 5) - hash) + text.charCodeAt(index);
    hash |= 0;
  }
  return Math.abs(hash);
}

function getFallbackNeedPoint(need) {
  const stateName = normalizeText(need?.state).toLowerCase();
  const base = STATE_MAP_CENTERS[stateName] || INDIA_FALLBACK_CENTER;
  const seed = hashText(`${need?.id || ""}|${need?.organization_name || ""}|${need?.district || ""}|${need?.state || ""}`);
  const latOffset = ((seed % 21) - 10) * 0.035;
  const lngOffset = (((Math.floor(seed / 21)) % 21) - 10) * 0.035;
  return {
    lat: Number((base.lat + latOffset).toFixed(5)),
    lng: Number((base.lng + lngOffset).toFixed(5)),
    label: getNeedMapLocationLabel(need),
    derived: true,
  };
}

function getNeedMapQueries(need) {
  const district = normalizeText(need?.district);
  const stateName = normalizeText(need?.state);
  const organization = normalizeText(need?.organization_name);
  const queries = [
    [district, stateName, "India"].filter(Boolean).join(", "),
    [organization, district, stateName, "India"].filter(Boolean).join(", "),
    [stateName, "India"].filter(Boolean).join(", "),
  ];
  return uniq(queries.filter(Boolean));
}

async function geocodeNeedForMap(need) {
  const existingLat = parseCoordinate(need?.latitude);
  const existingLng = parseCoordinate(need?.longitude);
  if (existingLat !== null && existingLng !== null) {
    return {
      lat: existingLat,
      lng: existingLng,
      label: normalizeText(need?.geocoded_label) || getNeedMapLocationLabel(need),
    };
  }

  const queries = getNeedMapQueries(need);
  for (const query of queries) {
    if (!state.mapGeocodeCache.has(query)) {
      state.mapGeocodeCache.set(query, (async () => {
        try {
          const response = await fetch(`https://nominatim.openstreetmap.org/search?format=jsonv2&limit=1&q=${encodeURIComponent(query)}`, {
            headers: { Accept: "application/json" },
          });
          const data = await response.json().catch(() => null);
          const match = Array.isArray(data) ? data[0] : null;
          const lat = parseCoordinate(match?.lat);
          const lng = parseCoordinate(match?.lon);
          if (lat === null || lng === null) return null;
          return {
            lat,
            lng,
            label: normalizeText(match?.display_name) || query,
          };
        } catch (error) {
          console.warn("Map geocode failed for", query, error);
          return null;
        }
      })());
    }
    const point = await state.mapGeocodeCache.get(query);
    if (point) return point;
  }
  return getFallbackNeedPoint(need);
}

async function getNeedMapGroups(needs) {
  const groups = new Map();
  const resolvedNeeds = await Promise.all(
    ensureList(needs).map(async (need) => ({
      need,
      point: await geocodeNeedForMap(need),
    })),
  );
  resolvedNeeds.forEach(({ need, point }) => {
    const lat = parseCoordinate(point?.lat);
    const lng = parseCoordinate(point?.lng);
    if (lat === null || lng === null) return;
    const key = `${lat.toFixed(5)}|${lng.toFixed(5)}`;
    if (!groups.has(key)) {
      groups.set(key, {
        key,
        lat,
        lng,
        label: normalizeText(point?.label) || getNeedMapLocationLabel(need),
        needs: [],
      });
    }
    groups.get(key).needs.push(need);
  });
  return [...groups.values()];
}

function renderCaseMapLocationPanel(group) {
  const panel = byId("caseMapLocationPanel");
  if (!panel) return;
  if (!group) {
    panel.classList.add("hidden");
    panel.innerHTML = "";
    state.activeMapGroupKey = "";
    return;
  }
  state.activeMapGroupKey = group.key;
  panel.classList.remove("hidden");
  panel.innerHTML = `
    <div class="case-map-location-panel-shell">
      <div class="case-map-location-panel-head">
        <div>
          <p class="eyebrow">Location Cases</p>
          <h4>${esc(group.label)}</h4>
        </div>
        <div class="case-map-location-panel-actions">
          <span class="status-pill info">${esc(group.needs.length)} cases</span>
          <button class="btn btn-secondary btn-small" type="button" data-close-map-panel="true">Close</button>
        </div>
      </div>
      <div class="case-map-location-grid">
      ${group.needs
        .map(
          (need) => `
            <article class="case-map-location-card">
              <div class="status-row">
                <span class="status-pill ${badgeTone(need.status)}">${esc(need.status)}</span>
                <span class="status-pill ${badgeTone(need.internal_status)}">${esc(need.internal_status)}</span>
              </div>
              <h4>${esc(need.organization_name)}</h4>
              <p class="helper-text">${esc(clipText(need.problem_statement, 150))}</p>
              <div class="card-actions">
                <span class="meta-text">${esc(getNeedMapLocationLabel(need))}</span>
                <button class="btn btn-secondary" data-open-need-id="${escAttr(need.id)}">View Details</button>
              </div>
            </article>
          `,
        )
        .join("")}
      </div>
    </div>
  `;
}

async function focusNeedFromMap(needId) {
  state.selectedNeedId = needId;
  state.queueNeedsScrollIntoView = true;
  state.matchPage = 1;
  switchView("operations");
  renderQueue();
  renderNeedDetail();
  renderWorkbench();
  await renderMatches();
}

async function renderCaseMap(needs) {
  const mapCanvas = byId("categoryCasesMap");
  if (!mapCanvas) return;
  const overlayPanel = byId("caseMapLocationPanel");
  if (overlayPanel) {
    overlayPanel.classList.add("hidden");
    overlayPanel.innerHTML = "";
  }
  const requestToken = Date.now();
  state.caseMapRequestToken = requestToken;
  mapCanvas.innerHTML = `<div class="case-map-empty">Mapping visible cases...</div><div id="caseMapLocationPanel" class="case-map-location-panel hidden"></div>`;
  const groups = await getNeedMapGroups(needs);
  if (state.caseMapRequestToken !== requestToken) return;
  if (!groups.length) {
    mapCanvas.innerHTML = `<div class="case-map-empty">No map points could be derived yet for the visible cases.</div><div id="caseMapLocationPanel" class="case-map-location-panel hidden"></div>`;
    renderCaseMapLocationPanel(null);
    state.caseMap = null;
    state.caseMapMarkers = [];
    return;
  }

  const sdkReady = await ensureMapplsSdk();
  if (state.caseMapRequestToken !== requestToken) return;
  if (!sdkReady || !window.mappls) {
    mapCanvas.innerHTML = `<div class="case-map-empty">Mappls map is waiting for a valid SDK key in <code>config.js</code>.</div><div id="caseMapLocationPanel" class="case-map-location-panel hidden"></div>`;
    renderCaseMapLocationPanel(null);
    return;
  }

  try {
    state.caseMapMarkers.forEach((marker) => marker?.remove?.());
    state.caseMapMarkers = [];
    state.caseMap?.remove?.();
    state.caseMap = null;

    const canvasId = `categoryCasesMapCanvas-${requestToken}`;
    mapCanvas.innerHTML = `<div id="${canvasId}" class="case-map-canvas"></div><div id="caseMapLocationPanel" class="case-map-location-panel hidden"></div>`;
    state.caseMap = new window.mappls.Map(canvasId, {
      center: { lat: groups[0].lat, lng: groups[0].lng },
      zoom: groups.length > 1 ? 4 : 7,
      zoomControl: true,
      geolocation: false,
      location: false,
    });

    groups.forEach((group) => {
      const markerHtml = group.needs.length > 1
        ? `<div class="case-map-marker-ring" data-count="${escAttr(group.needs.length)}"></div>`
        : `<div class="case-map-marker"></div>`;
      const marker = new window.mappls.Marker({
        map: state.caseMap,
        position: { lat: group.lat, lng: group.lng },
        html: markerHtml,
        popupHtml: `<div><strong>${esc(group.label)}</strong><br/>${esc(group.needs.length)} cases</div>`,
        popupOptions: { autoClose: true },
        fitbounds: false,
      });
      marker?.on?.("click", async () => {
        renderCaseMapLocationPanel(group);
      });
      marker?.addListener?.("click", async () => {
        renderCaseMapLocationPanel(group);
      });
      state.caseMapMarkers.push(marker);
    });

    if (groups.length > 1 && state.caseMap?.fitBounds) {
      const bounds = groups.map((group) => [group.lng, group.lat]);
      state.caseMap.fitBounds(bounds, { padding: 50, maxZoom: 7 });
    }

  } catch (error) {
    console.error("Case map could not be rendered.", error);
    mapCanvas.innerHTML = `<div class="case-map-empty">The MapmyIndia map could not be loaded on this page yet.</div><div id="caseMapLocationPanel" class="case-map-location-panel hidden"></div>`;
    renderCaseMapLocationPanel(null);
    state.caseMap = null;
    state.caseMapMarkers = [];
  }
}

function getCaseNeed(item) {
  if (item.type === "need" || item.type === "pendingNeed") return item.need;
  if (item.type === "pendingUpdate") return item.need;
  return null;
}

function metricMatchesNeed(metricId, need) {
  if (!need) return false;
  if (metricId === "approved") return need.approval_status === "approved";
  if (metricId === "in_progress") return need.status === "In progress";
  if (metricId === "need_providers") return need.internal_status === "Need solution providers";
  if (metricId === "connection_made") return need.internal_status === "Connection made";
  if (metricId === "stuck") return Number(need.curation_age_days || 0) >= 7;
  return false;
}

function metricMatchesItem(metricId, item) {
  if (metricId === "admin_queue") return item.type === "pendingNeed" || item.type === "pendingUpdate";
  return metricMatchesNeed(metricId, getCaseNeed(item));
}

function pipelineMatchesNeed(pipelineId, need) {
  if (!need) return false;
  const segment = PIPELINE_SEGMENTS.find((item) => item.id === pipelineId);
  return segment ? segment.match(need) : false;
}

function buildOverviewCases() {
  const needCases = getDisplayNeeds().map((need) => ({ type: "need", need }));
  const pendingNeedCases = ensureList(state.data.pendingNeeds).map((need) => ({ type: "pendingNeed", need }));
  const pendingUpdateCases = ensureList(state.data.pendingUpdates).map((request) => ({
    type: "pendingUpdate",
    request,
    need: getNeedById(request.need_id) || null,
  }));
  return [...needCases, ...pendingNeedCases, ...pendingUpdateCases];
}

function buildNeedCaseCard(need, extraMeta = "") {
  return `
    <article class="stack-card">
      <div class="status-row">
        <span class="status-pill ${badgeTone(need.status)}">${esc(need.status)}</span>
        <span class="status-pill ${badgeTone(need.internal_status)}">${esc(need.internal_status)}</span>
        <span class="status-pill info">${esc(need.curation_age_days || 0)} days</span>
      </div>
      <h4>${esc(need.organization_name)}</h4>
      <p class="helper-text">${esc(clipText(need.problem_statement, 165))}</p>
      <div class="card-actions">
        <span class="meta-text">${esc(`${need.state || "Unknown state"}${need.district ? ` / ${need.district}` : ""}${extraMeta ? ` • ${extraMeta}` : ""}`)}</span>
        <button class="btn btn-secondary" data-open-need-id="${esc(need.id)}">Open Need</button>
      </div>
    </article>
  `;
}

function buildPendingRequestCaseCard(request) {
  const linkedNeed = getNeedById(request.need_id);
  return `
    <article class="stack-card">
      <div class="status-row">
        <span class="status-pill warn">Pending update</span>
        <span class="status-pill info">${esc(request.submitted_by_curator_name || request.submitted_by_curator_email || "Curator")}</span>
      </div>
      <h4>${esc(linkedNeed?.organization_name || `Need #${request.need_id}`)}</h4>
      <p class="helper-text">${esc(clipText(request.proposed_curation_notes || request.proposed_next_action || "Curator update waiting for admin review.", 165))}</p>
      <div class="card-actions">
        <span class="meta-text">${esc(linkedNeed ? `${linkedNeed.state || "Unknown state"}${linkedNeed.district ? ` / ${linkedNeed.district}` : ""}` : `Need #${request.need_id}`)}</span>
        ${linkedNeed ? `<button class="btn btn-secondary" data-open-need-id="${esc(linkedNeed.id)}">Open Need</button>` : `<span></span>`}
      </div>
    </article>
  `;
}

function getOverviewFocusPayload() {
  const filters = state.overviewFilters || {};
  const activeMetrics = ensureList(filters.metric);
  const activePipelines = ensureList(filters.pipeline);
  const activeCurators = ensureList(filters.curator);
  const activeStates = ensureList(filters.state);
  const activeCategories = ensureList(filters.category);
  const activeFilterCount = [activeMetrics, activePipelines, activeCurators, activeStates, activeCategories].reduce((sum, list) => sum + list.length, 0);

  const allCases = buildOverviewCases();
  const filteredCases = allCases.filter((item) => {
    const need = getCaseNeed(item);
    if (activeMetrics.length && !activeMetrics.some((metricId) => metricMatchesItem(metricId, item))) return false;
    if (activePipelines.length && !activePipelines.some((pipelineId) => pipelineMatchesNeed(pipelineId, need))) return false;
    if (
      activeCurators.length &&
      !activeCurators.some((curatorId) => (curatorId === "unassigned" ? !need?.curator_id : need?.curator_id === curatorId))
    ) return false;
    if (activeStates.length && !activeStates.includes(normalizeText(need?.state))) return false;
    if (activeCategories.length && !activeCategories.some((category) => categoryFilterMatches(category, need))) return false;
    return true;
  });

  const orderedCases = filteredCases.sort((left, right) => {
    const leftNeed = getCaseNeed(left);
    const rightNeed = getCaseNeed(right);
    return Number(rightNeed?.curation_age_days || 0) - Number(leftNeed?.curation_age_days || 0);
  });

  const cards = orderedCases.map((item) => {
    if (item.type === "pendingNeed") return buildNeedCaseCard(item.need, "Pending intake");
    if (item.type === "pendingUpdate") return buildPendingRequestCaseCard(item.request);
    return buildNeedCaseCard(item.need);
  });

  const labels = [];
  activeMetrics.forEach((metricId) => {
    const metricLabels = {
      approved: "Approved Needs",
      in_progress: "In Progress",
      need_providers: "Need Providers",
      connection_made: "Connection Made",
      stuck: "Stuck 7+ Days",
      admin_queue: "Admin Queue",
    };
    labels.push(metricLabels[metricId] || metricId);
  });
  activePipelines.forEach((pipelineId) => {
    const segment = PIPELINE_SEGMENTS.find((item) => item.id === pipelineId);
    labels.push(segment?.label || pipelineId);
  });
  activeCurators.forEach((curatorId) =>
    labels.push(`Curator: ${curatorId === "unassigned" ? "Unassigned" : getCuratorById(curatorId)?.display_name || curatorId}`),
  );
  activeStates.forEach((stateName) => labels.push(`State: ${stateName}`));
  activeCategories.forEach((category) => labels.push(`Category: ${category}`));

  return {
    label: activeFilterCount ? labels.join(" | ") : "All Cases",
    tone: activeMetrics.includes("stuck") || activePipelines.includes("stuck") ? "bad" : activeMetrics.includes("admin_queue") ? "warn" : "info",
    items: orderedCases,
    cards,
    emptyText: activeMetrics.includes("admin_queue") && !isAdminUser()
      ? "Sign in as admin to inspect pending approvals and curator requests."
      : "No cases match the current combination of filters.",
  };
}

function scoreOfferingMatch(need, profile, offering) {
  const tags = parseArray(offering.tags).map((item) => item.toLowerCase());
  const geographies = parseArray(offering.geographies).map((item) => item.toLowerCase());
  const category = normalizeText(offering.offering_category).toLowerCase();
  const name = normalizeText(offering.offering_name).toLowerCase();
  const primaryApplication = normalizeText(offering.primary_application).toLowerCase();
  const primaryValuechain = normalizeText(offering.primary_valuechain).toLowerCase();
  const offeringGroup = normalizeText(offering.offering_group).toLowerCase();
  const offeringType = normalizeText(offering.offering_type).toLowerCase();
  const domain6m = normalizeText(offering.domain_6m).toLowerCase();
  const applications = parseArray(offering.applications).map((item) => item.toLowerCase());
  const valuechains = parseArray(offering.valuechains).map((item) => item.toLowerCase());
  const about = normalizeText(offering.about_offering_text || offering.solution?.about_solution_text).toLowerCase();
  const solutionName = normalizeText(offering.solution?.solution_name).toLowerCase();
  const offeringKind = getOfferingKind(offering);
  const aiOfferingTheme = normalizeText(offering.ai_thematic_area).toLowerCase();
  const aiOfferingApplication = normalizeText(offering.ai_application_area).toLowerCase();
  const aiOfferingService = normalizeText(offering.ai_service_kind).toLowerCase();
  const aiOfferingKeywords = parseArray(offering.ai_keywords).map((item) => item.toLowerCase());
  const joined = [
    name,
    category,
    primaryApplication,
    primaryValuechain,
    offeringGroup,
    offeringType,
    domain6m,
    applications.join(" "),
    valuechains.join(" "),
    about,
    solutionName,
    tags.join(" "),
    geographies.join(" "),
    aiOfferingTheme,
    aiOfferingApplication,
    aiOfferingService,
    aiOfferingKeywords.join(" "),
  ].join(" ");

  let score = 0;
  const reasons = [];
  let thematicMatched = false;
  let serviceMatched = false;
  let primaryThematicMatched = false;
  let problemThemeMatched = false;
  const matchedSharedLink =
    profile.sharedSolutionHints.ids.offeringIds.includes(String(offering.offering_id || "")) ||
    profile.sharedSolutionHints.ids.solutionIds.includes(String(offering.solution_id || "")) ||
    profile.sharedSolutionHints.ids.traderIds.includes(String(offering.trader_id || ""));

  if (matchedSharedLink) {
    thematicMatched = true;
    serviceMatched = true;
    score += 60;
    reasons.push("Already shared with seeker");
  }

  profile.problemThemeSignals.forEach((phrase) => {
    if (
      primaryApplication.includes(phrase) ||
      primaryValuechain.includes(phrase) ||
      applications.some((item) => item.includes(phrase)) ||
      valuechains.some((item) => item.includes(phrase)) ||
      name.includes(phrase) ||
      solutionName.includes(phrase) ||
      tags.some((tag) => tag.includes(phrase))
    ) {
      thematicMatched = true;
      primaryThematicMatched = true;
      problemThemeMatched = true;
      score += 22;
      reasons.push(`Theme:${phrase}`);
    } else if (about.includes(phrase) || aiOfferingApplication.includes(phrase) || aiOfferingTheme.includes(phrase)) {
      thematicMatched = true;
      problemThemeMatched = true;
      score += 10;
      reasons.push(`Theme:${phrase}`);
    }
  });

  profile.thematicAreas.forEach((phrase) => {
    if (joined.includes(phrase)) {
      thematicMatched = true;
      score += 18;
      reasons.push(phrase);
      if (
        tags.some((tag) => tag.includes(phrase)) ||
        primaryApplication.includes(phrase) ||
        primaryValuechain.includes(phrase) ||
        applications.some((item) => item.includes(phrase)) ||
        valuechains.some((item) => item.includes(phrase)) ||
        name.includes(phrase) ||
        solutionName.includes(phrase)
      ) {
        primaryThematicMatched = true;
      }
    }
  });

  if (aiOfferingTheme) {
    profile.thematicAreas.forEach((phrase) => {
      if (aiOfferingTheme.includes(phrase) || aiOfferingApplication.includes(phrase)) {
        thematicMatched = true;
        primaryThematicMatched = true;
        score += 18;
        reasons.push(`AI:${phrase}`);
      }
    });
  }

  profile.categoryTokens.forEach((token) => {
    if (tags.some((tag) => tag.includes(token))) {
      thematicMatched = true;
      primaryThematicMatched = true;
      score += 8;
      reasons.push(token);
    } else if (primaryApplication.includes(token) || primaryValuechain.includes(token)) {
      thematicMatched = true;
      primaryThematicMatched = true;
      score += 9;
      reasons.push(token);
    } else if (applications.some((item) => item.includes(token)) || valuechains.some((item) => item.includes(token))) {
      thematicMatched = true;
      primaryThematicMatched = true;
      score += 8;
      reasons.push(token);
    } else if (category.includes(token)) {
      thematicMatched = true;
      score += 7;
      reasons.push(token);
    } else if (name.includes(token) || solutionName.includes(token)) {
      thematicMatched = true;
      score += 6;
      reasons.push(token);
    } else if (about.includes(token)) {
      thematicMatched = true;
      score += 4;
      reasons.push(token);
    }
  });

  profile.serviceTerms.forEach((service) => {
    if (joined.includes(service)) {
      serviceMatched = true;
      score += 10;
      reasons.push(service);
    }
  });
  if (aiOfferingService && profile.serviceTerms.some((service) => aiOfferingService.includes(service) || service.includes(aiOfferingService))) {
    serviceMatched = true;
    score += 12;
    reasons.push(`AI:${aiOfferingService}`);
  }

  profile.serviceTokens.forEach((token) => {
    if (joined.includes(token)) {
      serviceMatched = true;
      score += 4;
      reasons.push(token);
    }
  });

  profile.problemTokens.slice(0, 8).forEach((token) => {
    if ((thematicMatched || !profile.categoryThematicAreas.length) && (
      tags.some((tag) => tag.includes(token)) ||
      name.includes(token) ||
      solutionName.includes(token) ||
      primaryApplication.includes(token) ||
      primaryValuechain.includes(token)
    )) {
      if (!profile.categoryThematicAreas.length) thematicMatched = true;
      score += 5;
      reasons.push(token);
    } else if ((thematicMatched || !profile.categoryThematicAreas.length) && about.includes(token)) {
      if (!profile.categoryThematicAreas.length) thematicMatched = true;
      score += 2;
      reasons.push(token);
    }
  });

  const domainMatched = profile.domainFocusTokens?.some((token) => (
    tags.some((tag) => tag.includes(token)) ||
    name.includes(token) ||
    solutionName.includes(token) ||
    primaryApplication.includes(token) ||
    primaryValuechain.includes(token) ||
    applications.some((item) => item.includes(token)) ||
    valuechains.some((item) => item.includes(token)) ||
    about.includes(token)
  )) || false;

  if (profile.domainFocusTokens?.length && !domainMatched) {
    score -= 18;
  }
  if (profile.problemThemeSignals.length && !problemThemeMatched) {
    score -= 16;
  }
  if (profile.domainFocusTokens?.length && thematicMatched && !primaryThematicMatched) {
    score -= 10;
  }

  const preferenceIndex = profile.preferredOfferingKinds.indexOf(offeringKind);
  if (preferenceIndex === 0) {
    score += 14;
    reasons.push(`${offeringKind} fit`);
  } else if (preferenceIndex === 1) {
    score += 5;
  } else if (preferenceIndex === 2) {
    score -= 8;
  }

  if ((profile.resolvedNeedKind === "service" || (profile.resolvedNeedKind === "mixed" && profile.serviceTerms.length)) && offeringKind === "knowledge") {
    score -= 14;
  }
  if (profile.resolvedNeedKind === "product" && offeringKind === "knowledge") {
    score -= 16;
  }
  if (profile.resolvedNeedKind === "knowledge" && offeringKind === "knowledge") {
    score += 10;
  }

  profile.geographyTokens.forEach((token) => {
    if (geographies.some((item) => item.includes(token)) || about.includes(token)) {
      score += 2;
      reasons.push(token);
    }
  });

  if (!thematicMatched) score -= 25;
  if (profile.requiresServiceMatch && profile.serviceTerms.length && !serviceMatched) score -= 12;
  if (!profile.categoryTokens.length && !profile.problemTokens.length) score += 1;
  if (!reasons.length) score -= 8;

  return {
    score,
    thematicMatched,
    primaryThematicMatched,
    serviceMatched,
    domainMatched,
    offeringKind,
    reasons: uniq(reasons).slice(0, 4),
  };
}

class GreMisStore {
  constructor() {
    this.config = window.APP_CONFIG || {};
    this.client = null;
    this.fallbackMode = false;
  }

  getClient() {
    if (this.client) return this.client;
    if (!this.config.SUPABASE_URL || !this.config.SUPABASE_ANON_KEY || !window.supabase?.createClient) {
      this.fallbackMode = true;
      return null;
    }
    this.client = window.supabase.createClient(this.config.SUPABASE_URL, this.config.SUPABASE_ANON_KEY);
    return this.client;
  }

  async loadBaseData() {
      const client = this.getClient();
      if (!client) {
        state.data.curators = FALLBACK_CURATORS;
        state.data.traders = [];
        state.data.offeringMaster = {
          valuechains: [],
          applications: [],
          tags: [],
          languages: [...DEFAULT_LANGUAGE_OPTIONS],
          geographies: [],
        };
        state.data.options = [];
        state.data.needs = [];
        state.data.needUpdates = [];
        state.matchCache.clear();
        return;
      }

    const [curators, traders, offerings, options, needs, updates] = await Promise.all([
        client.from("gre_mis_curators").select("id, display_name, email").eq("is_active", true).order("display_name"),
        client.from("traders").select("trader_id,trader_name,organisation_name,email,website,association_status").order("organisation_name"),
        client.from("offerings").select("primary_valuechain,primary_application,valuechains,applications,tags,languages,geographies").range(0, 19999),
        client.from("gre_mis_options").select("id, option_type, label, sort_order").eq("is_active", true).order("sort_order"),
        client.from("gre_mis_needs").select("*").eq("approval_status", "approved").order("requested_on", { ascending: false }),
        client.from("gre_mis_need_updates").select("*").order("created_at", { ascending: false }),
      ]);

      if (curators.error || traders.error || offerings.error || options.error || needs.error || updates.error) {
        throw new Error(curators.error?.message || traders.error?.message || offerings.error?.message || options.error?.message || needs.error?.message || updates.error?.message || "Could not load live dashboard data.");
      }

      state.data.curators = ensureList(curators.data);
      state.data.traders = ensureList(traders.data);
      state.data.offeringMaster = buildOfferingMasterData(ensureList(offerings.data));
      state.data.options = ensureList(options.data);
      state.data.needs = ensureList(needs.data).map((need) => ({
        ...need,
        curated_need: parseArray(need.curated_need),
      }));
    state.data.needUpdates = ensureList(updates.data);
    state.matchCache.clear();
  }

  async callAdmin(action, body = {}, requireAdmin = false) {
    const headers = {
      "Content-Type": "application/json",
      apikey: String(this.config.SUPABASE_ANON_KEY || ""),
      Authorization: `Bearer ${String(this.config.SUPABASE_ANON_KEY || "")}`,
    };
    if (state.userToken) headers["x-gre-user-session"] = state.userToken;
    if (state.adminToken) headers["x-gre-admin-session"] = state.adminToken;
    const response = await fetch(`${this.config.SUPABASE_URL}/functions/v1/${this.config.ADMIN_FUNCTION || "gre-mis-admin"}`, {
      method: "POST",
      headers,
      body: JSON.stringify({
        action,
        ...body,
        userSessionToken: state.userToken || undefined,
        adminSessionToken: state.adminToken || undefined,
      }),
    });
    const rawText = await response.text().catch(() => "");
    let data = {};
    try {
      data = rawText ? JSON.parse(rawText) : {};
    } catch {}
    if (!response.ok) throw new Error(data.error || rawText || `Request failed (${response.status}).`);
    if (requireAdmin && !isAdminUser()) throw new Error("Admin login required.");
    return data;
  }

  async userLogin(identifier, password) {
    const data = await this.callAdmin("userLogin", { identifier, password });
    syncSessionState(data.user, data.token);
    return data;
  }

  async validateUserSession() {
    if (!state.userToken) return false;
    try {
      const data = await this.callAdmin("validateUserSession");
      syncSessionState({
        id: data.user.id,
        username: data.user.username,
        email: data.user.email,
        first_name: data.user.firstName,
        full_name: data.user.fullName,
        phone: data.user.phone,
        role: data.user.role,
        must_change_password: data.user.mustChangePassword,
      });
      return true;
    } catch {
      syncSessionState(null, "");
      return false;
    }
  }

  async userLogout() {
    if (state.userToken) {
      await this.callAdmin("userLogout");
    }
    syncSessionState(null, "");
  }

  async registerUser(payload) {
    return this.callAdmin("registerUser", payload);
  }

  async requestPasswordReset(email) {
    return this.callAdmin("requestPasswordReset", { email });
  }

  async resetPassword(email, code, newPassword) {
    return this.callAdmin("resetPassword", { email, code, newPassword });
  }

  async changePassword(currentPassword, newPassword) {
    return this.callAdmin("changePassword", { currentPassword, newPassword });
  }

  async adminLogin(username, password) {
    return this.userLogin(username, password);
  }

  async validateAdminSession() {
    return this.validateUserSession();
  }

  async adminLogout() {
    return this.userLogout();
  }

  async loadAdminSnapshot() {
    if (!isAdminUser()) {
      state.data.pendingNeeds = [];
      state.data.pendingUpdates = [];
      state.data.pendingFormSubmissions = [];
      state.data.aiReviewNeeds = [];
      state.data.users = [];
      state.data.localSolutions = [];
      return;
    }
    const data = await this.callAdmin("adminSnapshot", {}, true);
    state.data.pendingNeeds = ensureList(data.pendingNeeds);
    state.data.pendingUpdates = ensureList(data.pendingUpdates);
    state.data.pendingFormSubmissions = ensureList(data.pendingFormSubmissions);
    state.data.aiReviewNeeds = ensureList(data.aiReviewNeeds);
    state.data.users = ensureList(data.users);
    state.data.localSolutions = ensureList(data.localSolutions);
  }

  async refreshUserDirectory() {
    const data = await this.callAdmin("refreshUserDirectory", {}, true);
    state.data.users = ensureList(data.users);
    return data;
  }

  async applyNeedOverride(needId, patch, conflictNote = "", resolveConflict = false) {
    return this.callAdmin("applyNeedOverride", { needId, patch, conflictNote, resolveConflict }, true);
  }

  async updateLocalSolution(offeringId, payload) {
    return this.callAdmin("updateLocalSolution", { offeringId, payload }, true);
  }

  async deleteLocalSolution(offeringId) {
    return this.callAdmin("deleteLocalSolution", { offeringId }, true);
  }

  async createNeed(payload) {
    const client = this.getClient();
    if (!client) throw new Error("Supabase is not configured.");

    const row = {
      id: String(Date.now()),
      organization_name: normalizeText(payload.organization_name),
      contact_person: normalizeText(payload.contact_person),
      seeker_email: normalizeText(payload.seeker_email),
      seeker_phone: normalizeText(payload.seeker_phone),
      state: normalizeText(payload.state),
      district: normalizeText(payload.district),
      problem_statement: normalizeText(payload.problem_statement),
      status: "New",
      internal_status: "Need solution providers",
      curated_need: normalizeText(payload.categories).split(",").map((item) => item.trim()).filter(Boolean),
      approval_status: "pending_admin",
      source_kind: "portal_submission",
      imported_from_batch: null,
      next_action: "Schedule curation call",
      requested_on: new Date().toISOString(),
      last_status_change_at: new Date().toISOString(),
    };

    const { error } = await client.from("gre_mis_needs").insert(row);
    if (error) throw new Error(error.message);
  }

  async assignCurator(needId, curatorId) {
    await this.callAdmin("assignCurator", { needId, curatorId }, true);
  }

  async approveNeed(needId, decision, reviewNotes = "") {
    await this.callAdmin("approveNeed", { needId, decision, reviewNotes }, true);
  }

  async submitUpdateRequest(payload) {
    await this.callAdmin("submitUpdateRequest", payload);
  }

  async reviewUpdateRequest(requestId, decision, reviewNotes = "") {
    await this.callAdmin("reviewUpdateRequest", { requestId, decision, reviewNotes }, true);
  }

  async upsertOption(optionType, label) {
    await this.callAdmin("upsertOption", { optionType, label }, true);
  }

  async sendProviderIntro(needId, providerEmail) {
    return this.callAdmin("sendProviderIntro", { needId, providerEmail });
  }

  async sendCuratorMessage(needId, message) {
    return this.callAdmin("sendCuratorMessage", { needId, message });
  }

  async directCuratorUpdate(payload) {
    return this.callAdmin("directCuratorUpdate", payload);
  }

  async updateUserRole(userId, role) {
    return this.callAdmin("updateUserRole", { userId, role }, true);
  }

  async completeUserRoleActivation(userId, otp) {
    return this.callAdmin("completeUserRoleActivation", { userId, otp }, true);
  }

  async removeManagedUser(userId, removalMode = "org_only") {
    return this.callAdmin("removeManagedUser", { userId, removalMode }, true);
  }

  async submitSharedForm(submissionType, payload) {
    return this.callAdmin("submitSharedForm", { submissionType, payload, sourceMode: "shared_link" });
  }

  async submitSignedInForm(submissionType, payload) {
    return this.callAdmin("submitSignedInForm", { submissionType, payload });
  }

  async reviewFormSubmission(submissionId, decision, reviewNotes = "") {
    return this.callAdmin("reviewFormSubmission", { submissionId, decision, reviewNotes }, true);
  }

  async updateFormSubmission(submissionId, update) {
    return this.callAdmin("updateFormSubmission", { submissionId, update }, true);
  }

  async suggestSolutionTags(payload) {
    return this.callAdmin("suggestSolutionTags", { payload }, false);
  }

  async searchLgdGeographies(query) {
    const client = this.getClient();
    const search = normalizeText(query);
    if (!client || search.length < 2) return [];
    const wildcard = `%${search.replace(/[%_]/g, " ").trim()}%`;
    const { data, error } = await client
      .from("lgd_geography_directory")
      .select("display_label,location_kind,block_name,district_name,state_name,gram_panchayat_name,village_name")
      .or(`search_text.ilike.${wildcard},display_label.ilike.${wildcard},block_name.ilike.${wildcard},district_name.ilike.${wildcard},state_name.ilike.${wildcard}`)
      .limit(20);
    if (error) {
      console.warn("LGD geography lookup failed", error);
      return [];
    }
    return buildRankedLgdSuggestions(ensureList(data), search);
  }

  async importInboundWorkbook(fileName, rows, aiProvider) {
    return this.callAdmin("importInboundWorkbook", { fileName, rows, aiProvider }, true);
  }

  async syncGreLiveInbounds(aiProvider) {
    return this.callAdmin("syncGreLiveInbounds", { aiProvider }, true);
  }

  async syncGreChatbotData(aiProvider) {
    return this.callAdmin("syncGreChatbotData", { aiProvider }, true);
  }

  async downloadGreChatbotReport(reportKind) {
    return this.callAdmin("downloadGreChatbotReport", { reportKind }, true);
  }

  async refreshNeedIntelligence(aiProvider) {
    return this.callAdmin("refreshNeedIntelligence", { aiProvider }, true);
  }

  async searchMatchesForNeed(need) {
    const client = this.getClient();
    if (!client || !need) return [];
    const profile = buildNeedMatchProfile(need);
    const offeringSelect = "offering_id,solution_id,trader_id,offering_name,offering_category,offering_group,offering_type,primary_application,primary_valuechain,applications,valuechains,domain_6m,tags,geographies,about_offering_text,contact_details,gre_link,ai_thematic_area,ai_application_area,ai_offering_kind,ai_service_kind,ai_keywords";
    const domainTerms = uniq([
      ...profile.thematicAreas.filter((item) => item && !GENERIC_THEMATIC_TERMS.has(item)),
      ...profile.domainFocusTokens,
      ...profile.phrases,
    ]);
    const searchTerms = uniq(
      profile.hasStrongTheme
        ? [...domainTerms, ...profile.serviceTerms.slice(0, 2)]
        : [...domainTerms, ...profile.serviceTerms, ...profile.primaryTerms, ...profile.explicitThemeTokens],
    ).slice(0, 14);
    const queries = [];

    if (searchTerms.length) {
      queries.push(
        client
          .from("offerings")
          .select(offeringSelect)
          .textSearch("search_document", searchTerms.join(" | "), {
            config: "simple",
            type: "websearch",
          })
          .limit(72),
      );
    }

    searchTerms.slice(0, 6).forEach((term) => {
      const safeTerm = term.replaceAll("%", "").replaceAll(",", " ").trim();
      if (!safeTerm) return;
      queries.push(
        client
          .from("offerings")
          .select(offeringSelect)
          .or(`offering_name.ilike.%${safeTerm}%,offering_category.ilike.%${safeTerm}%,about_offering_text.ilike.%${safeTerm}%,primary_application.ilike.%${safeTerm}%,primary_valuechain.ilike.%${safeTerm}%`)
          .limit(42),
      );
    });

    if (profile.sharedSolutionHints.ids.offeringIds.length) {
      queries.push(
        client
          .from("offerings")
          .select(offeringSelect)
          .in("offering_id", profile.sharedSolutionHints.ids.offeringIds)
          .limit(24),
      );
    }
    if (profile.sharedSolutionHints.ids.solutionIds.length) {
      queries.push(
        client
          .from("offerings")
          .select(offeringSelect)
          .in("solution_id", profile.sharedSolutionHints.ids.solutionIds)
          .limit(24),
      );
    }
    if (profile.sharedSolutionHints.ids.traderIds.length) {
      queries.push(
        client
          .from("offerings")
          .select(offeringSelect)
          .in("trader_id", profile.sharedSolutionHints.ids.traderIds)
          .limit(36),
      );
    }

    if (!queries.length) {
      queries.push(client.from("offerings").select(offeringSelect).limit(72));
    }

    const responses = await Promise.all(queries);
    const offerings = [];
    const seenOfferingIds = new Set();
    responses.forEach((response) => {
      (response.data || []).forEach((row) => {
        if (!seenOfferingIds.has(row.offering_id)) {
          seenOfferingIds.add(row.offering_id);
          offerings.push(row);
        }
      });
    });

    if (!offerings.length && profile.thematicAreas.length) {
      const thematicFallbackTerms = uniq([
        ...profile.thematicAreas,
        ...profile.domainFocusTokens,
      ]).slice(0, 4);
      const fallbackResponses = await Promise.all(
        thematicFallbackTerms.map((term) => {
          const safeTerm = term.replaceAll("%", "").replaceAll(",", " ").trim();
          if (!safeTerm) return Promise.resolve({ data: [], error: null });
          return client
            .from("offerings")
            .select(offeringSelect)
            .or(`offering_name.ilike.%${safeTerm}%,offering_category.ilike.%${safeTerm}%,about_offering_text.ilike.%${safeTerm}%,primary_application.ilike.%${safeTerm}%,primary_valuechain.ilike.%${safeTerm}%`)
            .limit(60);
        }),
      );

      fallbackResponses.forEach((response) => {
        (response.data || []).forEach((row) => {
          if (!seenOfferingIds.has(row.offering_id)) {
            seenOfferingIds.add(row.offering_id);
            offerings.push(row);
          }
        });
      });
    }

    if (!offerings.length) return [];
    const traderIds = [...new Set(offerings.map((item) => item.trader_id).filter(Boolean))];
    const solutionIds = [...new Set(offerings.map((item) => item.solution_id).filter(Boolean))];

    const [tradersRes, solutionsRes] = await Promise.all([
      traderIds.length
        ? client.from("traders").select("trader_id,trader_name,organisation_name,email,website").in("trader_id", traderIds)
        : Promise.resolve({ data: [], error: null }),
      solutionIds.length
        ? client.from("solutions").select("solution_id,solution_name,about_solution_text").in("solution_id", solutionIds)
        : Promise.resolve({ data: [], error: null }),
    ]);

    const traderMap = new Map((tradersRes.data || []).map((row) => [row.trader_id, row]));
    const solutionMap = new Map((solutionsRes.data || []).map((row) => [row.solution_id, row]));

    const ranked = offerings
      .map((offering) => {
        const enriched = {
          ...offering,
          trader: traderMap.get(offering.trader_id) || null,
          solution: solutionMap.get(offering.solution_id) || null,
        };
        const matchMeta = scoreOfferingMatch(need, profile, enriched);
      return {
        ...enriched,
        matchScore: matchMeta.score,
        thematicMatched: matchMeta.thematicMatched,
        primaryThematicMatched: matchMeta.primaryThematicMatched,
        serviceMatched: matchMeta.serviceMatched,
        domainMatched: matchMeta.domainMatched,
        offeringKind: matchMeta.offeringKind,
        matchReasons: matchMeta.reasons,
      };
    })
      .sort((a, b) => b.matchScore - a.matchScore);

    const strictMatches = ranked.filter((item) => {
      if (!item.thematicMatched) return false;
      if (profile.hasStrongTheme && !item.primaryThematicMatched) return false;
      if (profile.domainFocusTokens.length && !item.domainMatched) return false;
      if (profile.requiresServiceMatch && !item.serviceMatched) return false;
      if ((profile.resolvedNeedKind === "service" || (profile.resolvedNeedKind === "mixed" && profile.serviceTerms.length)) && item.offeringKind === "knowledge") return false;
      if (profile.resolvedNeedKind === "product" && item.offeringKind === "knowledge") return false;
      return item.matchScore >= 8;
    });

    if (strictMatches.length) return strictMatches;

    return ranked.filter((item) => {
      if (item.matchScore < 10) return false;
      if (!(item.domainMatched || item.thematicMatched)) return false;
      if (profile.hasStrongTheme && !item.primaryThematicMatched) return false;
      if ((profile.resolvedNeedKind === "service" || (profile.resolvedNeedKind === "mixed" && profile.serviceTerms.length)) && item.offeringKind === "knowledge") return false;
      return true;
    });
  }
}

const store = new GreMisStore();

function getCuratorById(id) {
  return state.data.curators.find((curator) => curator.id === id) || null;
}

function getCuratorByUser(user = state.userSession) {
  if (!user) return null;
  return state.data.curators.find(
    (curator) =>
      curator.user_id === user.id ||
      String(curator.email || "").toLowerCase() === String(user.email || "").toLowerCase(),
  ) || null;
}

function canEditNeedCuration(need) {
  if (!need || !state.userSession) return false;
  if (isAdminUser()) return true;
  if (!isCuratorUser()) return false;
  const myCurator = getCuratorByUser();
  return Boolean(myCurator && need.curator_id === myCurator.id);
}

function getDisplayNeeds() {
  return ensureList(state.data.needs).filter((need) => state.showClosedNeeds || need.status !== "Closed");
}

function getNeedById(id) {
  return getDisplayNeeds().find((need) => need.id === id) || null;
}

function getNeedMatchCacheKey(need) {
  if (!need) return "";
  return [
    need.id,
    need.last_status_change_at,
    need.updated_at,
    parseArray(need.curated_need).join("|"),
    normalizeText(need.problem_statement),
    normalizeText(need.curation_notes),
    need.state,
    need.district,
  ].join("::");
}

function getVisibleNeeds() {
  return [...getDisplayNeeds()]
    .filter((need) => state.filters.status === "all" || need.status === state.filters.status)
    .filter((need) => {
      if (state.filters.curator === "all") return true;
      if (state.filters.curator === "unassigned") return !need.curator_id;
      return need.curator_id === state.filters.curator;
    })
    .filter((need) => state.filters.state === "all" || normalizeText(need.state) === state.filters.state)
    .filter((need) => {
      if (!state.filters.search) return true;
      const haystack = [
        need.organization_name,
        need.problem_statement,
        parseArray(need.curated_need).join(" "),
        need.internal_status,
      ]
        .join(" ")
        .toLowerCase();
      return haystack.includes(state.filters.search.toLowerCase());
    })
    .sort((a, b) => Number(b.curation_age_days || 0) - Number(a.curation_age_days || 0));
}

function getNeedUpdates(needId) {
  return state.data.needUpdates.filter((update) => update.need_id === needId).sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
}

function countBy(items, pick) {
  return items.reduce((acc, item) => {
    const key = pick(item);
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
}

function topEntries(mapLike, limit = 6) {
  return Object.entries(mapLike).sort((a, b) => b[1] - a[1]).slice(0, limit);
}

function getSupplierOptions() {
  return ensureList(state.data.traders)
    .filter((row) => (row.association_status || "").toLowerCase() !== "rejected")
    .sort((a, b) => String(a.organisation_name || a.trader_name || "").localeCompare(String(b.organisation_name || b.trader_name || "")));
}

function getTraderById(traderId) {
  return getSupplierOptions().find((row) => String(row.trader_id) === String(traderId)) || null;
}

function findSupplierByName(name) {
  const normalized = normalizeText(name).toLowerCase();
  if (!normalized) return null;
  return getSupplierOptions().find((row) => {
    const org = normalizeText(row.organisation_name).toLowerCase();
    const trader = normalizeText(row.trader_name).toLowerCase();
    return org === normalized || trader === normalized;
  }) || null;
}

function buildSupplierOptionsHtml(selectedId = "", placeholder = "Choose approved GRE supplier") {
  const options = [`<option value="">${esc(placeholder)}</option>`];
  getSupplierOptions().forEach((row) => {
    const label = row.organisation_name || row.trader_name || `Supplier ${row.trader_id}`;
    options.push(`<option value="${esc(row.trader_id)}" ${String(row.trader_id) === String(selectedId) ? "selected" : ""}>${esc(label)}</option>`);
  });
  return options.join("");
}

function fillSupplierSelect(selectId, orgInputId) {
  const select = byId(selectId);
  const orgInput = byId(orgInputId);
  if (!select) return;
  const currentValue = select.value;
  const placeholder = selectId === "solutionTraderSelect"
    ? "Choose existing organisation (optional)"
    : "Choose approved GRE supplier";
  select.innerHTML = buildSupplierOptionsHtml(currentValue, placeholder);
  if (orgInput && !orgInput.value && select.value) {
    const trader = getTraderById(select.value);
    if (trader) orgInput.value = trader.organisation_name || trader.trader_name || "";
  }
}

function openMissingOrgDialog(orgName = "") {
  const dialog = byId("missingOrgDialog");
  const text = byId("missingOrgText");
  if (text) {
    text.textContent = orgName
      ? `${orgName} is not in the approved GRE supplier list yet. Please create the organisation account on the GRE platform first, then return here to submit the form.`
      : "This organisation is not in the approved GRE supplier list yet. Please create the organisation account on the GRE platform first, then return here to submit the form.";
  }
  dialog?.showModal();
}

function getShareUrl(mode) {
  const url = new URL(window.location.href);
  url.searchParams.set("sharedForm", mode === "solution" ? "solution" : "need");
  return url.toString();
}

function buildOfferingMasterData(rows) {
  const lists = {
    valuechains: new Set(),
    applications: new Set(),
    tags: new Set(),
    languages: new Set(DEFAULT_LANGUAGE_OPTIONS),
    geographies: new Set(),
  };
  const normalizedRows = [];
  ensureList(rows).forEach((row) => {
    const primaryValuechain = normalizeText(row.primary_valuechain);
    const primaryApplication = normalizeText(row.primary_application);
    const valuechains = parseArray(row.valuechains).filter(Boolean);
    const applications = parseArray(row.applications).filter(Boolean);
    [
      primaryValuechain,
      ...valuechains,
    ].filter(Boolean).forEach((item) => lists.valuechains.add(item));
    [
      primaryApplication,
      ...applications,
    ].filter(Boolean).forEach((item) => lists.applications.add(item));
    parseArray(row.tags).filter(Boolean).forEach((item) => lists.tags.add(item));
    parseArray(row.languages).filter(Boolean).forEach((item) => lists.languages.add(item));
    parseArray(row.geographies).filter(Boolean).forEach((item) => lists.geographies.add(item));
    normalizedRows.push({
      primaryValuechain,
      primaryApplication,
      valuechains,
      applications,
    });
  });
  return {
    valuechains: [...lists.valuechains].sort((a, b) => a.localeCompare(b)),
    applications: [...lists.applications].sort((a, b) => a.localeCompare(b)),
    tags: [...lists.tags].sort((a, b) => a.localeCompare(b)),
    languages: [...lists.languages].sort((a, b) => a.localeCompare(b)),
    geographies: [...lists.geographies].sort((a, b) => a.localeCompare(b)),
    rows: normalizedRows,
  };
}

function formatLgdGeographyLabel(row) {
  if (!row || typeof row !== "object") return "";
  const country = "India";
  const block = normalizeText(row.block_name || row.gram_panchayat_name || row.village_name);
  const city = normalizeText(row.district_name);
  const state = normalizeText(row.state_name);
  const kind = normalizeText(row.location_kind).toLowerCase();

  if (block && city && state) {
    return uniq([block, city, state, country]).join(", ");
  }
  if (city && state) {
    return uniq([city, state, country]).join(", ");
  }
  if (state) {
    return uniq([state, country]).join(", ");
  }
  if (kind === "country") {
    return country;
  }
  return normalizeText(row.display_label) || country;
}

function buildLgdGeographyVariants(row) {
  if (!row || typeof row !== "object") return [];
  const country = "India";
  const block = normalizeText(row.block_name || row.gram_panchayat_name || row.village_name);
  const city = normalizeText(row.district_name);
  const state = normalizeText(row.state_name);
  const variants = [];
  if (block && city && state) variants.push(uniq([block, city, state, country]).join(", "));
  if (city && state) variants.push(uniq([city, state, country]).join(", "));
  if (state) variants.push(uniq([state, country]).join(", "));
  if (normalizeText(row.location_kind).toLowerCase() === "country" || !variants.length) variants.push(country);
  const display = normalizeText(row.display_label);
  if (display) variants.push(display);
  return uniq(variants.filter(Boolean));
}

function scoreLgdSuggestion(label, query) {
  const normalizedLabel = normalizeText(label).toLowerCase();
  const normalizedQuery = normalizeText(query).toLowerCase();
  if (!normalizedLabel || !normalizedQuery) return -1;
  if (normalizedLabel === normalizedQuery) return 400;
  if (normalizedLabel.startsWith(`${normalizedQuery},`)) return 320;
  if (normalizedLabel.startsWith(normalizedQuery)) return 280;
  const parts = normalizedLabel.split(",").map((item) => item.trim());
  if (parts[0] === normalizedQuery) return 260;
  if (parts.some((part) => part === normalizedQuery)) return 220;
  if (parts.some((part) => part.startsWith(normalizedQuery))) return 180;
  if (normalizedLabel.includes(normalizedQuery)) return 120;
  return 0;
}

function buildRankedLgdSuggestions(rows, query) {
  const seen = new Set();
  return ensureList(rows)
    .flatMap((row) => buildLgdGeographyVariants(row))
    .filter((label) => {
      const key = normalizeText(label).toLowerCase();
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .map((label) => ({ label, score: scoreLgdSuggestion(label, query) }))
    .filter((item) => item.score > 0)
    .sort((left, right) => {
      if (right.score !== left.score) return right.score - left.score;
      return left.label.localeCompare(right.label);
    })
    .slice(0, 20)
    .map((item) => item.label);
}

function setSelectOptions(selectId, values, allowBlank = true) {
  const select = byId(selectId);
  if (!select) return;
  const previous = select.value;
  const items = ensureList(values);
  select.innerHTML = [
    allowBlank ? `<option value="">Select</option>` : "",
    ...items.map((item) => {
      const value = typeof item === "string" ? item : item.value;
      const label = typeof item === "string" ? item : item.label;
      return `<option value="${esc(value)}">${esc(label)}</option>`;
    }),
  ].join("");
  if (items.some((item) => (typeof item === "string" ? item : item.value) === previous)) {
    select.value = previous;
  }
}

function setMultiSelectOptions(selectId, values) {
  const select = byId(selectId);
  if (!select) return;
  const previous = [...select.selectedOptions].map((option) => option.value);
  select.innerHTML = ensureList(values)
    .map((item) => `<option value="${esc(item)}">${esc(item)}</option>`)
    .join("");
  [...select.options].forEach((option) => {
    option.selected = previous.includes(option.value);
  });
}

function setCheckboxGroupOptions(groupId, inputName, values) {
  const group = byId(groupId);
  if (!group) return;
  const previous = [...group.querySelectorAll(`input[name="${inputName}"]:checked`)].map((input) => input.value);
  group.innerHTML = ensureList(values)
    .map((item) => `
      <label>
        <input type="checkbox" name="${escAttr(inputName)}" value="${escAttr(item)}" ${previous.includes(item) ? "checked" : ""} />
        <span>${esc(item)}</span>
      </label>
    `)
    .join("");
}

function setDatalistOptions(listId, values) {
  const list = byId(listId);
  if (!list) return;
  list.innerHTML = ensureList(values)
    .map((value) => `<option value="${esc(value)}"></option>`)
    .join("");
}

function getCheckedValues(form, name) {
  return [...form.querySelectorAll(`input[name="${name}"]:checked`)].map((input) => normalizeText(input.value)).filter(Boolean);
}

function normalizeMultiTextValue(value) {
  if (Array.isArray(value)) return uniq(value.map((item) => normalizeText(item)).filter(Boolean));
  return uniq(
    String(value || "")
      .split(",")
      .map((item) => normalizeText(item))
      .filter(Boolean),
  );
}

function normalizeSemicolonSeparatedValue(value) {
  if (Array.isArray(value)) return uniq(value.map((item) => normalizeText(item)).filter(Boolean));
  return uniq(
    String(value || "")
      .split(/[;\n]+/)
      .map((item) => normalizeText(item))
      .filter(Boolean),
  );
}

function parseLooseListInput(value) {
  if (Array.isArray(value)) return uniq(value.map((item) => normalizeText(item)).filter(Boolean));
  return uniq(
    String(value || "")
      .split(/[,\n;]+/)
      .map((item) => normalizeText(item))
      .filter(Boolean),
  );
}

function syncProductAudienceFromSolution() {
  const solutionValues = getCheckedValues(byId("solutionSubmissionForm"), "solution_audience");
  document.querySelectorAll('input[name="product_audience"]').forEach((input) => {
    input.checked = solutionValues.includes(normalizeText(input.value));
  });
}

function updateSolutionOfferingForm() {
  const categorySelect = byId("solutionOfferingCategory");
  if (!categorySelect) return;
  const category = categorySelect.value || "Service offerings";
  setSelectOptions("solutionOfferingType", OFFERING_TYPE_OPTIONS[category] || [], true);
  document.querySelectorAll(".conditional-offering-section").forEach((section) => {
    section.classList.toggle("hidden", section.dataset.offeringKind !== category);
  });
  const productFields = [...document.querySelectorAll('#productOfferingFields input, #productOfferingFields textarea, #productOfferingFields select')];
  const serviceFields = [...document.querySelectorAll('#serviceOfferingFields input, #serviceOfferingFields textarea, #serviceOfferingFields select')];
  const knowledgeFields = [...document.querySelectorAll('#knowledgeOfferingFields input, #knowledgeOfferingFields textarea, #knowledgeOfferingFields select')];
  productFields.forEach((field) => { field.disabled = category !== "Product offerings"; });
  serviceFields.forEach((field) => { field.disabled = category !== "Service offerings"; });
  knowledgeFields.forEach((field) => { field.disabled = category !== "Knowledge offerings"; });
}

function renderSolutionReferenceInputs() {
  const master = state.data.offeringMaster || buildOfferingMasterData([]);
  setCheckboxGroupOptions("solutionLanguagesGroup", "languages", master.languages);
  setDatalistOptions("solutionTagOptions", master.tags);
  updateSolutionOfferingForm();
}

function getPrimaryApplicationOptions(primaryValuechain) {
  const master = state.data.offeringMaster || buildOfferingMasterData([]);
  if (!primaryValuechain) return [];
  const matches = master.rows
    .filter((row) => row.primaryValuechain === primaryValuechain)
    .flatMap((row) => [row.primaryApplication, ...row.applications])
    .filter(Boolean);
  return uniq(matches).sort((a, b) => a.localeCompare(b));
}

function getSecondaryApplicationOptions(primaryValuechain, secondaryValuechain) {
  const master = state.data.offeringMaster || buildOfferingMasterData([]);
  if (!primaryValuechain || !secondaryValuechain) return [];
  const primaryOptions = new Set(getPrimaryApplicationOptions(primaryValuechain));
  const matches = master.rows
    .filter((row) => {
      if (row.primaryValuechain !== primaryValuechain) return false;
      return row.valuechains.includes(secondaryValuechain) || row.primaryValuechain === secondaryValuechain;
    })
    .flatMap((row) => row.applications)
    .filter((item) => item && !primaryOptions.has(item));
  if (!matches.length) {
    return master.rows
      .filter((row) => row.primaryValuechain === secondaryValuechain || row.valuechains.includes(secondaryValuechain))
      .flatMap((row) => [row.primaryApplication, ...row.applications])
      .filter((item) => item && !primaryOptions.has(item))
      .filter(Boolean)
      .filter(Boolean)
      .reduce((acc, item) => (acc.includes(item) ? acc : [...acc, item]), [])
      .sort((a, b) => a.localeCompare(b));
  }
  return uniq(matches)
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b));
}

function updateSolutionApplicationOptions() {
  const primaryValuechain = normalizeText(byId("solutionPrimaryValuechain")?.value || "");
  const secondaryValuechain = normalizeText(byId("solutionSecondaryValuechain")?.value || "");
  setSelectOptions("solutionPrimaryApplication", getPrimaryApplicationOptions(primaryValuechain), true);
  setSelectOptions("solutionSecondaryApplication", getSecondaryApplicationOptions(primaryValuechain, secondaryValuechain), true);
}

function appendSubmissionEntry(target, key, value) {
  if (value === undefined || value === null || value === "") return;
  if (Object.prototype.hasOwnProperty.call(target, key)) {
    target[key] = Array.isArray(target[key]) ? [...target[key], value] : [target[key], value];
    return;
  }
  target[key] = value;
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error(`Could not read file ${file.name}.`));
    reader.readAsDataURL(file);
  });
}

async function collectSingleFileAttachment(form, fieldName) {
  const input = form.querySelector(`input[name="${fieldName}"]`);
  const file = input?.files?.[0];
  if (!file || !file.name) return null;
  if (file.size > MAX_EMBEDDED_FILE_BYTES) {
    throw new Error(`${file.name} is larger than 5 MB. Please use a smaller file for admin review.`);
  }
  return {
    name: file.name,
    type: file.type || "application/octet-stream",
    size: file.size,
    dataUrl: await readFileAsDataUrl(file),
  };
}

async function updateSolutionGeographySuggestions(rawValue) {
  const input = byId("solutionGeographyEntry");
  if (!input) return;
  const query = normalizeText(String(rawValue || ""));
  const requestToken = Date.now();
  state.lgdSearchToken = requestToken;
  if (query.length < 2) {
    setDatalistOptions("solutionGeographyOptions", []);
    return;
  }
  const suggestions = await store.searchLgdGeographies(query);
  if (state.lgdSearchToken !== requestToken) return;
  setDatalistOptions("solutionGeographyOptions", suggestions);
}

function renderSolutionGeographyChips() {
  const target = byId("solutionGeographyChips");
  const hidden = byId("solutionGeographies");
  if (!target || !hidden) return;
  hidden.value = state.solutionGeographies.join("; ");
  target.innerHTML = state.solutionGeographies.length
    ? state.solutionGeographies
        .map(
          (location) => `
            <span class="tag-chip">
              ${esc(location)}
              <button type="button" class="tag-chip-remove" data-remove-solution-geo="${escAttr(location)}" aria-label="Remove ${escAttr(location)}">&times;</button>
            </span>
          `,
        )
        .join("")
    : `<span class="helper-text">No locations added yet.</span>`;
}

function addSolutionGeography(rawValue) {
  const location = normalizeText(rawValue);
  if (!location) return false;
  if (!state.solutionGeographies.includes(location)) state.solutionGeographies.push(location);
  renderSolutionGeographyChips();
  return true;
}

function removeSolutionGeography(location) {
  state.solutionGeographies = state.solutionGeographies.filter((item) => item !== location);
  renderSolutionGeographyChips();
}

function renderSolutionTagChips() {
  const target = byId("solutionTagChips");
  const hidden = byId("solutionTagsInput");
  if (!target || !hidden) return;
  hidden.value = state.solutionTags.join(", ");
  target.innerHTML = state.solutionTags.length
    ? state.solutionTags
        .map(
          (tag) => `
            <span class="tag-chip">
              ${esc(tag)}
              <button type="button" class="tag-chip-remove" data-remove-solution-tag="${escAttr(tag)}" aria-label="Remove ${escAttr(tag)}">&times;</button>
            </span>
          `,
        )
        .join("")
    : `<span class="helper-text">No tags added yet.</span>`;
}

function addSolutionTag(rawValue) {
  const tag = normalizeText(rawValue);
  if (!tag) return false;
  if (!state.solutionTags.includes(tag)) state.solutionTags.push(tag);
  renderSolutionTagChips();
  return true;
}

function removeSolutionTag(tag) {
  state.solutionTags = state.solutionTags.filter((item) => item !== tag);
  renderSolutionTagChips();
}

function deriveLocalSolutionTags(draft) {
  const text = [
    normalizeText(draft.offering_name),
    normalizeText(draft.offering_type),
    normalizeText(draft.about_offering_text),
    normalizeText(draft.organization_name),
  ].join(" ").toLowerCase();
  const hintMap = [
    ["dairy", ["dairy", "milk", "milking", "fodder", "cattle", "cow", "livestock"]],
    ["goatery", ["goat", "goatery"]],
    ["poultry", ["poultry", "broiler", "layer", "chicken"]],
    ["soap", ["soap", "detergent"]],
    ["solar", ["solar", "street light", "street lights"]],
    ["training", ["training", "capacity building"]],
    ["consulting", ["consulting", "consultancy", "advisory"]],
    ["mentoring", ["mentoring"]],
    ["technology transfer", ["technology transfer"]],
    ["market support", ["market support", "market linkage", "branding", "packaging"]],
    ["financial support", ["financial support", "finance", "credit", "loan"]],
  ];
  const tags = hintMap
    .filter(([, patterns]) => patterns.some((pattern) => text.includes(pattern)))
    .map(([label]) => label);
  const tokenTags = uniq(
    text
      .split(/[^a-z0-9]+/)
      .map((token) => token.trim())
      .filter((token) => token.length >= 4 && !["with", "from", "that", "this", "service", "solution", "support", "offering"].includes(token)),
  ).slice(0, 8);
  return uniq([...tags, ...tokenTags]).slice(0, 12);
}

function getSubmissionReviewFieldConfig(submissionType, payload = {}) {
  if (submissionType === "need") {
    return [
      { key: "contact_person", label: "Contact Person" },
      { key: "seeker_email", label: "Email" },
      { key: "seeker_phone", label: "Phone" },
      { key: "state", label: "State" },
      { key: "district", label: "District" },
      { key: "need_kind", label: "Need Kind" },
      { key: "service_kind", label: "Service Type" },
      { key: "thematic_area", label: "Thematic Area" },
      { key: "application_area", label: "Application Area" },
      { key: "curated_need", label: "Curated Need / Categories", multiline: true, list: true },
      { key: "six_m_signals", label: "6M Signals", multiline: true, list: true },
      { key: "problem_statement", label: "Problem Statement", multiline: true, wide: true },
    ];
  }
  const category = normalizeText(payload?.offering_category) || "Service offerings";
  const baseFields = [
    { key: "submitter_name", label: "Contact Person" },
    { key: "submitter_email", label: "Contact Email" },
    { key: "submitter_phone", label: "Contact Phone" },
    { key: "offering_category", label: "Offering Category" },
    { key: "offering_type", label: "Offering Type" },
    { key: "offering_name", label: "Offering Name" },
    { key: "about_offering_text", label: "Offering Description", multiline: true, wide: true },
  ];
  const productFields = [
    { key: "grade_capacity", label: "Grade / Capacity" },
    { key: "product_cost", label: "Product Cost" },
    { key: "product_cost_quote_on_scope", label: "Quote on Scope" },
    { key: "lead_time", label: "Lead Time" },
    { key: "support_details", label: "Support Services", multiline: true },
    { key: "product_contact_details", label: "Product Contact Details", multiline: true },
  ];
  const serviceFields = [
    { key: "trainer_name", label: "Facilitator Name" },
    { key: "trainer_email", label: "Facilitator Email" },
    { key: "trainer_phone", label: "Facilitator Phone" },
    { key: "trainer_details_text", label: "Facilitator Details", multiline: true, wide: true },
    { key: "languages", label: "Languages", multiline: true, list: true },
    { key: "geographies", label: "Geographies", multiline: true, list: true },
    { key: "duration", label: "Duration" },
    { key: "duration_unit", label: "Duration Units" },
    { key: "prerequisites", label: "Prerequisites", multiline: true },
    { key: "location_availability", label: "Location Availability", multiline: true, list: true },
    { key: "service_cost", label: "Service Cost" },
    { key: "service_cost_unit", label: "Service Cost Units" },
    { key: "cost_remarks", label: "Cost Remarks", multiline: true },
    { key: "support_post_service", label: "Support Post Service" },
    { key: "support_post_service_cost", label: "Support Post Service Cost" },
    { key: "delivery_mode", label: "Delivery Mode" },
    { key: "certification_offered", label: "Certification Offered" },
  ];
  const knowledgeFields = [
    { key: "knowledge_content_link", label: "Knowledge Content Link" },
    { key: "contact_details", label: "Knowledge Contact Details", multiline: true },
  ];
  const trailingFields = [
    { key: "tags", label: "Tags", multiline: true, list: true },
  ];
  if (category === "Product offerings") return [...baseFields, ...productFields, ...trailingFields];
  if (category === "Knowledge offerings") return [...baseFields, ...knowledgeFields, ...trailingFields];
  return [...baseFields, ...serviceFields, ...trailingFields];
}

function renderSubmissionReviewFields(submissionType, payload) {
  const target = byId("submissionReviewFields");
  const title = byId("submissionStructuredTitle");
  if (!target) return;
  if (title) title.textContent = submissionType === "need" ? "Need Submission Details" : "Solution Submission Details";
  const config = getSubmissionReviewFieldConfig(submissionType, payload);
  target.innerHTML = config
    .map((field) => {
      const rawValue = payload?.[field.key];
      const value = field.list ? parseLooseListInput(rawValue).join(", ") : normalizeText(rawValue);
      if (field.multiline) {
        return `
          <label class="${field.wide ? "wide" : ""}">
            <span>${esc(field.label)}</span>
            <textarea data-review-field="${escAttr(field.key)}" rows="${field.wide ? 4 : 3}">${esc(value)}</textarea>
          </label>
        `;
      }
      return `
        <label class="${field.wide ? "wide" : ""}">
          <span>${esc(field.label)}</span>
          <input data-review-field="${escAttr(field.key)}" value="${escAttr(value)}" />
        </label>
      `;
    })
    .join("");
}

function getSubmissionAttachmentConfigs(submissionType, payload = {}) {
  if (submissionType !== "solution") return [];
  const category = normalizeText(payload?.offering_category) || "Service offerings";
  const configs = [];
  if (category === "Product offerings") {
    configs.push({
      key: "product_brochure_attachment",
      label: "Product Brochure",
      inputName: "review_product_brochure_attachment",
      accept: ".pdf,.doc,.docx,.ppt,.pptx,.jpg,.jpeg,.png",
    });
  }
  if (category === "Service offerings") {
    configs.push({
      key: "service_brochure_attachment",
      label: "Service Brochure",
      inputName: "review_service_brochure_attachment",
      accept: ".pdf,.doc,.docx,.ppt,.pptx,.jpg,.jpeg,.png",
    });
  }
  if (category === "Knowledge offerings") {
    configs.push({
      key: "knowledge_content_attachment",
      label: "Knowledge File",
      inputName: "review_knowledge_content_attachment",
      accept: ".pdf,.doc,.docx,.ppt,.pptx,.xls,.xlsx,.csv,.jpg,.jpeg,.png,.mp4",
    });
  }
  configs.push({
    key: "offering_image_attachment",
    label: "Offering Image",
    inputName: "review_offering_image_attachment",
    accept: ".jpg,.jpeg,.png,.webp",
  });
  return configs;
}

function getAttachmentDisplayInfo(value) {
  if (!value || typeof value !== "object") return null;
  const dataUrl = normalizeText(value.dataUrl || value.url || "");
  const name = normalizeText(value.name || value.fileName || "Attached file");
  if (!dataUrl) return name ? { name, href: "" } : null;
  return { name, href: dataUrl };
}

function renderSubmissionReviewAttachments(submissionType, payload) {
  const target = byId("submissionReviewAttachments");
  if (!target) return;
  const configs = getSubmissionAttachmentConfigs(submissionType, payload);
  if (!configs.length) {
    target.innerHTML = `<div class="empty-state">No file attachments are relevant for this submission type.</div>`;
    return;
  }
  target.innerHTML = configs.map((config) => {
    const current = getAttachmentDisplayInfo(payload?.[config.key]);
    return `
      <article class="attachment-review-card">
        <div class="attachment-review-copy">
          <strong>${esc(config.label)}</strong>
          ${
            current
              ? current.href
                ? `<a href="${esc(current.href)}" target="_blank" rel="noreferrer">${esc(current.name)}</a>`
                : `<span>${esc(current.name)}</span>`
              : `<span class="helper-text">No file attached yet.</span>`
          }
        </div>
        <label>
          <span>Replace File</span>
          <input name="${escAttr(config.inputName)}" type="file" accept="${escAttr(config.accept)}" />
        </label>
      </article>
    `;
  }).join("");
}

function getFormSubmissionById(submissionId) {
  return ensureList(state.data.pendingFormSubmissions).find((submission) => submission.id === submissionId) || null;
}

function fillSubmissionReviewTraderSelect(selectedId = "") {
  const select = byId("submissionReviewTraderSelect");
  if (!select) return;
  select.innerHTML = buildSupplierOptionsHtml(selectedId);
}

async function buildSubmissionReviewPatch() {
  const form = byId("submissionReviewForm");
  if (!form) throw new Error("Submission review form is unavailable.");
  const submissionId = normalizeText(form.querySelector('[name="submission_id"]')?.value || "");
  const organizationName = normalizeText(form.querySelector('[name="organization_name"]')?.value || "");
  const existingTraderId = normalizeText(form.querySelector('[name="existing_trader_id"]')?.value || "");
  const adminReviewNotes = normalizeText(form.querySelector('[name="admin_review_notes"]')?.value || "");
  const submission = getFormSubmissionById(submissionId);
  const isSolution = submission?.submission_type === "solution";
  const payloadBase = submission?.payload && typeof submission.payload === "object" ? { ...submission.payload } : {};
  const payload = payloadBase;
  form.querySelectorAll("[data-review-field]").forEach((field) => {
    const key = field.dataset.reviewField;
    if (!key) return;
    if (field.tagName === "TEXTAREA") {
      payload[key] = field.value;
    } else {
      payload[key] = field.value;
    }
  });
  const attachmentConfigs = getSubmissionAttachmentConfigs(submission?.submission_type || "", payload);
  for (const config of attachmentConfigs) {
    const attachment = await collectSingleFileAttachment(form, config.inputName);
    if (attachment) payload[config.key] = attachment;
  }
  const payloadField = form.querySelector('[name="payload_json"]');
  const trader = getTraderById(existingTraderId);
  if (!trader && !isSolution) throw new Error("Please choose a valid GRE supplier before saving the submission.");
  payload.organization_name = organizationName
    || payload.organization_name
    || trader?.organisation_name
    || trader?.trader_name
    || "";
  payload.existing_trader_id = trader?.trader_id || "";
  payload.existing_trader_name = trader?.organisation_name || trader?.trader_name || "";
  if (payloadField) payloadField.value = JSON.stringify(payload, null, 2);
  return {
    submissionId,
    update: {
      organizationName: payload.organization_name,
      existingTraderId: trader?.trader_id || "",
      existingTraderName: payload.existing_trader_name,
      adminReviewNotes,
      payload,
    },
  };
}

function openSubmissionReviewDialog(submissionId) {
  const submission = getFormSubmissionById(submissionId);
  if (!submission) {
    toast("This submission is no longer available in the admin queue.");
    return;
  }
  state.submissionReviewId = submissionId;
  const dialog = byId("submissionReviewDialog");
  const form = byId("submissionReviewForm");
  const title = byId("submissionReviewTitle");
  const meta = byId("submissionReviewMeta");
  const status = byId("submissionReviewStatus");
  if (!dialog || !form || !title || !meta) return;
  const payload = submission.payload && typeof submission.payload === "object" ? submission.payload : {};
  title.textContent = `${submission.submission_type === "solution" ? "Solution" : "Need"} Submission Review`;
  if (status) status.textContent = "";
  const submissionIdField = form.querySelector('[name="submission_id"]');
  const orgField = form.querySelector('[name="organization_name"]');
  const notesField = form.querySelector('[name="admin_review_notes"]');
  const payloadField = form.querySelector('[name="payload_json"]');
  if (submissionIdField) submissionIdField.value = submission.id;
  if (orgField) orgField.value = submission.organization_name || payload.organization_name || "";
  if (notesField) notesField.value = submission.admin_review_notes || "";
  if (payloadField) payloadField.value = JSON.stringify(payload, null, 2);
  fillSubmissionReviewTraderSelect(submission.existing_trader_id || payload.existing_trader_id || "");
  renderSubmissionReviewFields(submission.submission_type, payload);
  renderSubmissionReviewAttachments(submission.submission_type, payload);
  meta.innerHTML = `
    <div><strong>Submission Type:</strong> ${esc(submission.submission_type)}</div>
    <div><strong>Source:</strong> ${esc(submission.source_mode || "shared_link")}</div>
    <div><strong>Submitter:</strong> ${esc(submission.submitter_name || submission.submitter_email || "Anonymous")}</div>
    <div><strong>Current GRE Sync:</strong> ${esc(submission.gre_sync_status ? submission.gre_sync_status.replaceAll("_", " ") : "not started")}</div>
  `;
  dialog.showModal();
}

function renderSharePanel(targetId, mode) {
  const target = byId(targetId);
  if (!target) return;
  const label = mode === "solution" ? "Add Solution" : "Solution Need Form";
  const shareUrl = getShareUrl(mode);
  const mailto = `mailto:?subject=${encodeURIComponent(`GRE ${label} intake link`)}&body=${encodeURIComponent(`Please use this GRE MIS form link to share your ${label.toLowerCase()} input:\n\n${shareUrl}`)}`;
  target.innerHTML = `
    <article class="stack-card share-card">
      <p class="helper-text">Use this public link to collect ${esc(label.toLowerCase())} inputs from other contributors. Their submission will go to admin review first.</p>
      <div class="share-link-box">${esc(shareUrl)}</div>
      <div class="card-actions">
        <button type="button" class="btn btn-secondary" data-copy-share-url="${escAttr(shareUrl)}">Copy Link</button>
        <a class="btn btn-primary" href="${esc(mailto)}">Mail This Link</a>
      </div>
    </article>
  `;
}

function populateSubmissionDefaults(form) {
  if (!form || !state.userSession) return;
  [
    ["submitter_name", state.userSession.full_name || state.userSession.first_name || ""],
    ["submitter_email", state.userSession.email || ""],
    ["submitter_phone", state.userSession.phone || ""],
    ["contact_person", state.userSession.full_name || state.userSession.first_name || ""],
    ["seeker_email", state.userSession.email || ""],
    ["seeker_phone", state.userSession.phone || ""],
  ].forEach(([name, value]) => {
    const input = form.querySelector(`[name="${name}"]`);
    if (input && !input.value && value) input.value = value;
  });
}

function renderSubmissionViews() {
  fillSupplierSelect("solutionTraderSelect", "solutionOrgName");
  fillSupplierSelect("needTraderSelect", "needOrgName");
  renderSolutionReferenceInputs();
  state.solutionTags = [];
  state.solutionGeographies = [];
  renderSolutionTagChips();
  renderSolutionGeographyChips();
  renderSharePanel("solutionSharePanel", "solution");
  renderSharePanel("needSharePanel", "need");
  populateSubmissionDefaults(byId("solutionSubmissionForm"));
  populateSubmissionDefaults(byId("needSubmissionForm"));

  const solutionTitle = document.querySelector('#solutionView h3');
  const needTitle = document.querySelector('#need-intakeView h3');
  if (solutionTitle) {
    solutionTitle.textContent = isSharedFormMode() ? "Share a Solution for GRE Review" : "Add Solution to Admin Review Queue";
  }
  if (needTitle) {
    needTitle.textContent = isSharedFormMode() ? "Share a Need with GRE" : "Add Need to GRE Review Queue";
  }
}

function renderMetrics() {
  const metricsGrid = byId("metricsGrid");
  if (!metricsGrid) return;
  const needs = getDisplayNeeds();
  const metrics = [
      ["approved", "Approved Needs", needs.length, "Live inbound needs currently visible in the MIS."],
      ["in_progress", "In Progress", needs.filter((need) => need.status === "In progress").length, "Needs under active curation or provider search."],
      ["need_providers", "Need Providers", needs.filter((need) => need.internal_status === "Need solution providers").length, "Needs waiting on live solution/provider matching."],
      ["connection_made", "Connection Made", needs.filter((need) => need.internal_status === "Connection made").length, "Needs where the solution side has begun to move."],
      ["stuck", "Stuck 7+ Days", needs.filter((need) => Number(need.curation_age_days || 0) >= 7).length, "Aging needs that need intervention during the day."],
      ["admin_queue", "Admin Queue", state.data.pendingNeeds.length + state.data.pendingUpdates.length, "Pending intake approvals and curator change requests."],
    ];
  if (byId("adminView")) {
    metrics.push(["ai_review", "Match QA", state.data.aiReviewNeeds.length, "Approved needs flagged for conflict review, missing validation, or weak classification."]);
  }
  metricsGrid.style.setProperty("--metric-count", String(metrics.length));

  metricsGrid.innerHTML = metrics
    .map(
      ([id, label, value, note]) => `
        <article class="metric-card ${isOverviewFocus("metric", id) ? "active" : ""}" data-overview-kind="metric" data-overview-id="${esc(id)}">
          <p class="eyebrow">${esc(label)}</p>
          <strong>${esc(value)}</strong>
          <p>${esc(note)}</p>
        </article>
      `,
    )
    .join("");
}

function renderOverview() {
  if (!byId("overviewView")) return;
  const needs = getDisplayNeeds();
  const headline = byId("datasetHeadline");
  const subline = byId("datasetSubline");
  if (headline) headline.textContent = `${needs.length} Active Inbound Needs loaded from GRE Operations Data`;
  if (subline) {
    const adminStats = [
      ["Admin Action Awaited", state.data.pendingNeeds.length + state.data.pendingUpdates.length],
      ["Intake Records", state.data.pendingNeeds.length],
      ["Curator Updates", state.data.pendingUpdates.length],
      ["Closed Display", state.showClosedNeeds ? "Visible" : "Hidden"],
    ];
    subline.innerHTML = adminStats
      .map(
        ([label, value]) => `
          <article class="dataset-stat-card">
            <span>${esc(label)}</span>
            <strong>${esc(value)}</strong>
          </article>
        `,
      )
      .join("");
  }
  const closedToggleBtn = byId("closedToggleBtn");
  if (closedToggleBtn) closedToggleBtn.textContent = `Show Closed: ${state.showClosedNeeds ? "On" : "Off"}`;

  const stageData = PIPELINE_SEGMENTS.map((segment) => [
    segment.id,
    segment.label,
    needs.filter(segment.match).length,
    segment.note,
  ]);

  byId("pipelineBoard").innerHTML = stageData
    .map(
      ([id, label, value, note]) => `
        <article class="stage-card ${isOverviewFocus("pipeline", id) ? "active" : ""}" data-overview-kind="pipeline" data-overview-id="${esc(id)}">
          <p class="eyebrow">${esc(label)}</p>
          <strong>${esc(value)}</strong>
          <p class="helper-text">${esc(note)}</p>
        </article>
      `,
    )
    .join("");

  const workload = ensureList(state.data.curators).map((curator) => ({
    label: curator.display_name,
    value: needs.filter((need) => need.curator_id === curator.id).length,
    focusKind: "curator",
    focusId: curator.id,
  }));
  workload.push({
    label: "Unassigned",
    value: needs.filter((need) => !need.curator_id).length,
    focusKind: "curator",
    focusId: "unassigned",
  });
  renderBarList("workloadChart", workload, "good");

  const stateCounts = topEntries(countBy(needs, (need) => normalizeText(need.state) || "Unknown"));
  renderBarList(
    "stateChart",
    stateCounts.map(([label, value]) => ({ label, value, focusKind: "state", focusId: label })),
    "info",
  );

  const sixMCounts = Object.fromEntries(SIX_M_LABELS.map((label) => [label, 0]));
  const themeCounts = {};
  needs.forEach((need) => {
    getNeedSixMSignals(need).forEach((label) => {
      sixMCounts[label] = (sixMCounts[label] || 0) + 1;
    });
    getNeedThemeSignals(need).forEach((theme) => {
      themeCounts[theme] = (themeCounts[theme] || 0) + 1;
    });
  });

  byId("categoryChart").innerHTML = `
    <div class="signal-section">
      <p class="signal-heading">6M View</p>
      <div class="tag-cloud">
        ${SIX_M_LABELS.map(
          (label) =>
            `<button class="${isOverviewFocus("category", label) ? "active" : ""}" data-overview-kind="category" data-overview-id="${esc(label)}">${esc(label)} (${esc(sixMCounts[label] || 0)})</button>`,
        ).join("")}
      </div>
    </div>
    <div class="signal-section">
      <p class="signal-heading">Most Common Problem Themes</p>
      <div class="tag-cloud">
        ${topEntries(themeCounts, 12)
          .map(
            ([label, value]) =>
              `<button class="${isOverviewFocus("category", label) ? "active" : ""}" data-overview-kind="category" data-overview-id="${esc(label)}">${esc(label)} (${esc(value)})</button>`,
          )
          .join("")}
      </div>
    </div>
  `;

  const focusPayload = getOverviewFocusPayload();
  const pageSize = 12;
  const totalPages = Math.max(1, Math.ceil(focusPayload.cards.length / pageSize));
  if (state.overviewPage > totalPages) state.overviewPage = totalPages;
  const pageStart = (state.overviewPage - 1) * pageSize;
  const visibleCards = focusPayload.cards.slice(pageStart, pageStart + pageSize);
  const focusNeeds = focusPayload.items.map((item) => getCaseNeed(item)).filter(Boolean);
  byId("pipelineDrilldown").innerHTML = `
    <div class="pipeline-drilldown-head">
      <div>
        <p class="eyebrow">Category Cases</p>
        <h4>${esc(focusPayload.label)}</h4>
      </div>
      <span class="status-pill ${focusPayload.tone}">${esc(focusPayload.cards.length)} cases</span>
    </div>
    <div class="pipeline-drilldown-split">
      <div class="pipeline-drilldown-sidebar">
        <div class="pipeline-drilldown-list">
          ${focusPayload.cards.length
            ? visibleCards.join("")
            : `<div class="empty-state">${esc(focusPayload.emptyText || "No cases are currently sitting in this selection.")}</div>`}
        </div>
        ${
          focusPayload.cards.length
            ? `<div class="pipeline-pagination pipeline-pagination-left">
                <span class="meta-text">Showing ${esc(pageStart + 1)}-${esc(Math.min(pageStart + pageSize, focusPayload.cards.length))} of ${esc(focusPayload.cards.length)}</span>
                <div class="pipeline-pagination-actions">
                  <button class="btn btn-secondary" data-page-action="prev" ${state.overviewPage <= 1 ? "disabled" : ""}>Prev</button>
                  <span class="meta-text">Page ${esc(state.overviewPage)} of ${esc(totalPages)}</span>
                  <button class="btn btn-secondary" data-page-action="next" ${state.overviewPage >= totalPages ? "disabled" : ""}>Next</button>
                </div>
              </div>`
            : ""
        }
      </div>
      <div id="categoryCasesMap" class="pipeline-drilldown-map">
        <div id="caseMapLocationPanel" class="case-map-location-panel hidden"></div>
      </div>
    </div>
  `;
  renderCaseMap(focusNeeds);
}

function renderBarList(targetId, items, tone) {
  const target = byId(targetId);
  if (!target) return;
  const safeItems = ensureList(items);
  const max = Math.max(...safeItems.map((item) => Number(item.value || 0)), 1);
  target.innerHTML = safeItems
    .map(
      (item) => `
        <button class="bar-row bar-row-button ${item.focusKind && isOverviewFocus(item.focusKind, item.focusId) ? "active" : ""}" ${
          item.focusKind ? `data-overview-kind="${esc(item.focusKind)}" data-overview-id="${esc(item.focusId)}"` : ""
        }>
          <span>${esc(item.label)}</span>
          <div class="bar-track"><div class="bar-fill ${tone === "warn" ? "warn" : tone === "bad" ? "bad" : ""}" style="width:${(Number(item.value || 0) / max) * 100}%"></div></div>
          <strong>${esc(item.value)}</strong>
        </button>
      `,
    )
    .join("");
}

function renderFilters() {
  if (!byId("statusFilter")) return;
  const displayNeeds = getDisplayNeeds();
  const statusOptions = ["all", ...new Set(displayNeeds.map((need) => need.status).filter(Boolean))];
  const stateOptions = ["all", ...new Set(displayNeeds.map((need) => normalizeText(need.state)).filter(Boolean))];

  document.getElementById("statusFilter").innerHTML = statusOptions
    .map((value) => `<option value="${esc(value)}">${esc(value === "all" ? "All statuses" : value)}</option>`)
    .join("");
  document.getElementById("curatorFilter").innerHTML = [
    `<option value="all">All curators</option>`,
    `<option value="unassigned">Unassigned</option>`,
    ...ensureList(state.data.curators).map((curator) => `<option value="${esc(curator.id)}">${esc(curator.display_name)}</option>`),
  ].join("");
  document.getElementById("stateFilter").innerHTML = stateOptions
    .map((value) => `<option value="${esc(value)}">${esc(value === "all" ? "All states" : value)}</option>`)
    .join("");

  document.getElementById("statusFilter").value = state.filters.status;
  document.getElementById("curatorFilter").value = state.filters.curator;
  document.getElementById("stateFilter").value = state.filters.state;
  document.getElementById("searchFilter").value = state.filters.search;
}

function renderQueue() {
  const queue = byId("needsQueue");
  if (!queue) return;
  const visible = getVisibleNeeds();
  if (!state.selectedNeedId && visible[0]) state.selectedNeedId = visible[0].id;

  queue.innerHTML = visible
    .map((need) => {
      const curator = getCuratorById(need.curator_id);
      return `
        <article class="queue-item ${state.selectedNeedId === need.id ? "active" : ""}" data-need-id="${esc(need.id)}">
          <div class="status-row">
            <span class="status-pill ${badgeTone(need.status)}">${esc(need.status)}</span>
            <span class="status-pill ${badgeTone(need.internal_status)}">${esc(need.internal_status)}</span>
          </div>
          <h4>${esc(need.organization_name)}</h4>
          <p class="helper-text">${esc(need.problem_statement.slice(0, 120))}${need.problem_statement.length > 120 ? "..." : ""}</p>
          <p class="meta-text">${esc(need.state)}${need.district ? ` / ${esc(need.district)}` : ""} • Curator: ${esc(curator?.display_name || "Unassigned")} • Age: ${esc(need.curation_age_days || 0)} days</p>
        </article>
      `;
    })
    .join("");

  if (state.queueNeedsScrollIntoView) {
    const activeCard = queue.querySelector(`[data-need-id="${CSS.escape(String(state.selectedNeedId || ""))}"]`);
    if (activeCard) {
      activeCard.scrollIntoView({ inline: "center", block: "nearest", behavior: "smooth" });
    }
    state.queueNeedsScrollIntoView = false;
  }
}

function renderNeedDetail() {
  const need = getNeedById(state.selectedNeedId);
  const detailEl = byId("needDetail");
  if (!detailEl) return;
  if (!need) {
    detailEl.innerHTML = `<div class="empty-state">No need selected.</div>`;
    return;
  }

  const curator = getCuratorById(need.curator_id);
  const solutionLinks = extractUrls(need.curation_notes);
  const noteText = normalizeText(stripUrls(need.curation_notes)).replace(/\bnull\b/gi, "").replace(/\s+/g, " ").trim();
  const canInspectCuration = canSeeCurationDetails();
  const summaryBadges = [
    { label: need.status, tone: badgeTone(need.status) },
    { label: need.internal_status, tone: badgeTone(need.internal_status) },
    { label: `${need.curation_age_days || 0} days old`, tone: Number(need.curation_age_days || 0) >= 7 ? "bad" : "info" },
    { label: need.approval_status, tone: badgeTone(need.approval_status) },
  ];
  detailEl.innerHTML = `
    <div class="need-detail-stack selected-need-layout">
      <section class="need-overview-card">
      <div class="need-overview-head">
        <div>
          <p class="eyebrow">Need #${esc(need.id)}</p>
          <h3>${esc(need.organization_name)}</h3>
          <p class="helper-text">${esc(`${need.state || "Unknown state"}${need.district ? ` / ${need.district}` : ""}`)}${curator ? ` • ${esc(curator.display_name)}` : " • Unassigned curator"}</p>
        </div>
        <div class="status-row">
          ${summaryBadges.map((badge) => `<span class="status-pill ${badge.tone}">${esc(badge.label)}</span>`).join("")}
        </div>
      </div>
      <div class="need-statement">
        <h4>Problem Statement</h4>
        <p>${esc(need.problem_statement || "No problem statement provided.")}</p>
      </div>
      </section>

      <div class="selected-need-grid">
        <article class="detail-card detail-stack-card">
          <div class="detail-card-subhead">
            <h4>Seeker Details</h4>
            <span class="status-pill info">${esc(formatDate(need.requested_on))}</span>
          </div>
          <div class="kv-grid">
            <div><span>Contact Person</span><strong>${esc(need.contact_person || "Not set")}</strong></div>
            <div><span>Email</span><strong>${esc(need.seeker_email || "Not set")}</strong></div>
            <div><span>Phone</span><strong>${esc(need.seeker_phone || "Not set")}</strong></div>
            <div><span>Website</span><strong>${esc(need.website || "Not set")}</strong></div>
            <div><span>Designation</span><strong>${esc(need.designation || "Not set")}</strong></div>
            <div><span>Requested On</span><strong>${esc(formatDate(need.requested_on))}</strong></div>
          </div>
        </article>

        ${canInspectCuration
          ? `<article class="detail-card detail-stack-card">
              <h4>Curation Snapshot</h4>
              <div class="kv-grid">
                <div><span>Assigned Curator</span><strong>${esc(curator?.display_name || "Unassigned")}</strong></div>
                <div><span>Curator Email</span><strong>${esc(curator?.email || "Not set")}</strong></div>
                <div><span>Next Action</span><strong>${esc(need.next_action || "Not set")}</strong></div>
                <div><span>Curation Call</span><strong>${esc(need.curation_call_date || "Not set")}</strong></div>
                <div><span>Broadcast Needed</span><strong>${need.demand_broadcast_needed ? "Yes" : "No"}</strong></div>
                <div><span>Solutions Shared</span><strong>${esc(need.solutions_shared_count || 0)}</strong></div>
                <div><span>Invited Providers</span><strong>${esc(need.invited_providers_count || 0)}</strong></div>
                <div><span>Current Stage</span><strong>${esc(need.internal_status || "Not set")}</strong></div>
              </div>
            </article>`
          : `<article class="detail-card detail-stack-card">
              <h4>Curator Handling</h4>
              <div class="kv-grid">
                <div><span>Assigned Curator</span><strong>${esc(curator?.display_name || "Unassigned")}</strong></div>
                <div><span>Curator Email</span><strong>${esc(curator?.email || "Not set")}</strong></div>
                <div><span>Status</span><strong>${esc(need.status || "Not set")}</strong></div>
                <div><span>Current Stage</span><strong>${esc(need.internal_status || "Not set")}</strong></div>
              </div>
              <p class="helper-text">Detailed curation notes and internal actioning are visible only to curators and admins.</p>
            </article>`}

        <article class="detail-card detail-stack-card">
          <h4>Categories and Signals</h4>
          <div class="tag-row">${parseArray(need.curated_need).map((item) => `<span>${esc(item)}</span>`).join("") || `<span>Unclassified</span>`}</div>
          <div class="kv-grid">
            <div><span>State</span><strong>${esc(need.state || "Not set")}</strong></div>
            <div><span>District</span><strong>${esc(need.district || "Not set")}</strong></div>
            <div><span>Approval Status</span><strong>${esc(need.approval_status || "Not set")}</strong></div>
            <div><span>Curation Age</span><strong>${esc(need.curation_age_days || 0)} days</strong></div>
          </div>
        </article>
      </div>

      ${canInspectCuration
        ? `<article class="detail-card detail-stack-card">
            <h4>Curation Notes</h4>
            ${noteText ? `<p class="detail-note">${esc(noteText)}</p>` : `<p class="detail-note">No curation notes have been recorded yet.</p>`}
            ${
              solutionLinks.length
                ? `<div class="detail-links">
                    <p class="signal-heading">Solution Links Shared with Seeker</p>
                    <ol class="detail-link-list">
                      ${solutionLinks
                        .map(
                          (link, index) =>
                            `<li><a href="${esc(link)}" target="_blank" rel="noreferrer">Solution Link ${index + 1}</a></li>`,
                        )
                        .join("")}
                    </ol>
                  </div>`
                : ""
            }
          </article>`
        : ""}

      <div class="selected-need-actions">
        ${isLoggedIn() ? `<button type="button" id="openWorkbenchBtn" class="btn btn-primary">Open Action Workbench</button>` : ""}
      </div>
    </div>
  `;
}

function buildMatchDetailHtml(match) {
  if (!match) return `<div class="empty-state">Match details are not available.</div>`;
  const fullText = normalizeText(match.about_offering_text || match.solution?.about_solution_text || "No detailed offering summary is available yet.");
  return `
    <section class="need-overview-card">
      <div class="need-overview-head">
        <div>
          <p class="eyebrow">${esc(match.offering_category || "GRE Offering")}</p>
          <h3>${esc(match.offering_name || match.solution?.solution_name || "Unnamed match")}</h3>
          <p class="helper-text">${esc(match.trader?.organisation_name || match.trader?.trader_name || "Provider not mapped")}</p>
        </div>
        <div class="status-row">
          <span class="status-pill match-score-pill">${esc(`Relevance Score ${match.matchScore}`)}</span>
          <span class="status-pill info">${esc((match.offeringKind || match.ai_offering_kind || match.offering_group || "Offering").toString())}</span>
        </div>
      </div>
      <div class="tag-row">
        ${parseArray(match.tags).map((tag) => `<span>${esc(tag)}</span>`).join("")}
      </div>
    </section>
    <div class="selected-need-grid">
      <article class="detail-card detail-stack-card">
        <h4>Provider</h4>
        <div class="kv-grid">
          <div><span>Organisation</span><strong>${esc(match.trader?.organisation_name || match.trader?.trader_name || "Not set")}</strong></div>
          <div><span>Email</span><strong>${esc(match.trader?.email || "Not set")}</strong></div>
          <div><span>Website</span><strong>${esc(match.trader?.website || "Not set")}</strong></div>
          <div><span>GRE Link</span><strong>${match.gre_link ? `<a href="${esc(match.gre_link)}" target="_blank" rel="noreferrer">Open on GRE</a>` : "Not available"}</strong></div>
        </div>
      </article>
      <article class="detail-card detail-stack-card">
        <h4>Offering Structure</h4>
        <div class="kv-grid">
          <div><span>Offering Group</span><strong>${esc(match.offering_group || "Not set")}</strong></div>
          <div><span>Offering Type</span><strong>${esc(match.offering_type || "Not set")}</strong></div>
          <div><span>Primary Application</span><strong>${esc(match.primary_application || "Not set")}</strong></div>
          <div><span>Primary Value Chain</span><strong>${esc(match.primary_valuechain || "Not set")}</strong></div>
        </div>
      </article>
      <article class="detail-card detail-stack-card">
        <h4>Why It Matched</h4>
        <div class="kv-grid">
          <div><span>Reasons</span><strong>${esc((match.matchReasons || []).join(", ") || "Closest live GRE keyword overlap")}</strong></div>
          <div><span>Geographies</span><strong>${esc(parseArray(match.geographies).join(", ") || "Not listed")}</strong></div>
          <div><span>Solution</span><strong>${esc(match.solution?.solution_name || "Not mapped")}</strong></div>
          <div><span>Application Area</span><strong>${esc(match.ai_application_area || "Not set")}</strong></div>
        </div>
      </article>
    </div>
    <article class="detail-card detail-stack-card">
      <h4>Full Description</h4>
      <p class="detail-note">${esc(fullText)}</p>
    </article>
  `;
}

function openMatchDetailDialog(match) {
  const dialog = byId("matchDetailDialog");
  const title = byId("matchDetailTitle");
  const body = byId("matchDetailBody");
  if (!dialog || !title || !body) return;
  title.textContent = match?.offering_name || match?.solution?.solution_name || "Match Details";
  body.innerHTML = buildMatchDetailHtml(match);
  dialog.showModal();
}

function openMatchDetailPage(match) {
  if (!match?.offering_id) return;
  try {
    sessionStorage.setItem(`gre-match-${match.offering_id}`, JSON.stringify(match));
  } catch {}
  window.open(`./offering-detail.html?offering_id=${encodeURIComponent(match.offering_id)}`, "_blank", "noopener");
}

function getCachedMatchByOfferingId(offeringId) {
  const need = getNeedById(state.selectedNeedId);
  const matches = state.matchCache.get(getNeedMatchCacheKey(need)) || [];
  return matches.find((item) => String(item.offering_id) === String(offeringId)) || null;
}

async function renderMatches() {
  const need = getNeedById(state.selectedNeedId);
  const matchesEl = byId("matchResults");
  if (!matchesEl) return;
  if (!need) {
    matchesEl.innerHTML = `<div class="empty-state">Select a need to search the GRE solutions and providers catalog.</div>`;
    return;
  }

  matchesEl.innerHTML = `<div class="empty-state">Searching live GRE offerings and solution providers…</div>`;
  const cacheKey = getNeedMatchCacheKey(need);
  let matches = state.matchCache.get(cacheKey);
  if (!matches) {
    matches = await store.searchMatchesForNeed(need);
    state.matchCache.set(cacheKey, matches);
  }
  const pageSize = 6;
  const totalPages = Math.max(1, Math.ceil(matches.length / pageSize));
  if (state.matchPage > totalPages) state.matchPage = totalPages;
  const pageStart = (state.matchPage - 1) * pageSize;
  const visibleMatches = matches.slice(pageStart, pageStart + pageSize);

  matchesEl.innerHTML = matches.length
    ? visibleMatches
        .map((match) => {
          const email = match.trader?.email || "";
          const matchSummary = normalizeText(match.about_offering_text || match.solution?.about_solution_text || "");
          const truncatedSummary = matchSummary.slice(0, 280);
          const hasMore = matchSummary.length > 280;
          return `
            <article class="match-card ${match.matchScore >= 100 ? "match-score-high" : "match-score-medium"}">
              <div class="match-head">
                <div>
                  <p class="eyebrow">${esc(match.offering_category || "GRE Offering")}</p>
                  <h4>${esc(match.offering_name || match.solution?.solution_name || "Unnamed match")}</h4>
                </div>
                <span class="status-pill match-score-pill">${esc(`Relevance Score ${match.matchScore}`)}</span>
              </div>
              <div class="tag-row">
                ${parseArray(match.tags).slice(0, 5).map((tag) => `<span>${esc(tag)}</span>`).join("")}
              </div>
              <div class="match-meta-grid">
                <div>
                  <span>Provider</span>
                  <strong>${esc(match.trader?.organisation_name || match.trader?.trader_name || "Provider not mapped")}</strong>
                </div>
                <div>
                  <span>Why This Match</span>
                  <strong>${esc(match.matchReasons?.length ? match.matchReasons.join(", ") : "Closest live GRE keyword overlap")}</strong>
                </div>
              </div>
              <div class="match-copy">
                <p>${esc(truncatedSummary)}${hasMore ? "..." : ""}</p>
                ${hasMore ? `<button class="match-readmore" type="button" data-action="read-more-match" data-offering-id="${esc(match.offering_id)}">Read more</button>` : ""}
              </div>
              <p class="meta-text">${esc(parseArray(match.geographies).slice(0, 3).join(", ") || "Geography not listed")}</p>
              <div class="card-actions match-actions">
                <button class="btn btn-secondary" type="button" data-action="view-match-details" data-offering-id="${esc(match.offering_id)}">View Details</button>
                ${match.gre_link ? `<a class="btn btn-secondary" href="${esc(match.gre_link)}" target="_blank" rel="noreferrer">Open GRE Link</a>` : `<span></span>`}
                ${window.APP_CONFIG?.ENABLE_EMAIL_ACTIONS && (isAdminUser() || isCuratorUser()) && email
                  ? `<button class="btn btn-primary" data-action="email-provider" data-provider-email="${esc(email)}">Email This Provider</button>`
                  : `<button class="btn btn-primary" type="button" disabled title="Provider outreach is available only to signed-in curators and admins.">Email This Provider</button>`}
              </div>
            </article>
          `;
        })
        .join("") + `
      <div class="pipeline-pagination">
        <span class="meta-text">Showing ${esc(pageStart + 1)}-${esc(Math.min(pageStart + pageSize, matches.length))} of ${esc(matches.length)} matches</span>
        <div class="pipeline-pagination-actions">
          <button class="btn btn-secondary" data-match-page-action="prev" ${state.matchPage <= 1 ? "disabled" : ""}>Prev</button>
          <span class="meta-text">Page ${esc(state.matchPage)} of ${esc(totalPages)}</span>
          <button class="btn btn-secondary" data-match-page-action="next" ${state.matchPage >= totalPages ? "disabled" : ""}>Next</button>
        </div>
      </div>`
    : `<div class="empty-state">No direct live catalog match was found for this need yet. Try refining categories or curator notes.</div>`;
}

function renderWorkbench() {
  const need = getNeedById(state.selectedNeedId);
  const workbench = byId("actionWorkbench");
  if (!workbench) return;
  if (!need) {
    workbench.innerHTML = `<div class="empty-state">Select a need to take action.</div>`;
    return;
  }

  const curator = getCuratorById(need.curator_id);
  const statusOptions = state.data.options.filter((item) => item.option_type === "status");
  const internalOptions = state.data.options.filter((item) => item.option_type === "internal_status");
  const nextActionOptions = state.data.options.filter((item) => item.option_type === "next_action");
  const currentStatusLabel = normalizeText(need.status) || "Not set";
  const currentInternalStatusLabel = normalizeText(need.internal_status) || "Not set";
  const currentNextActionLabel = normalizeText(need.next_action) || "Not set";
  const curatorOptions = [`<option value="">Unassigned</option>`, ...state.data.curators.map((item) => `<option value="${esc(item.id)}" ${item.id === need.curator_id ? "selected" : ""}>${esc(item.display_name)}</option>`)].join("");
  const canEdit = canEditNeedCuration(need);
  const isCurator = isCuratorUser();
  const isAdmin = isAdminUser();
  const canMessageCurator = isLoggedIn() && !canEdit;

  workbench.innerHTML = `
    ${isAdmin ? `
      <article class="action-card">
        <p class="eyebrow">Admin Assignment</p>
        <h4>Curator Allocation</h4>
        <label>
          <span>Assigned Curator</span>
          <select id="assignCuratorSelect">${curatorOptions}</select>
        </label>
        <button class="btn btn-secondary" id="assignCuratorBtn">Save Curator Assignment</button>
        <p class="helper-text">Admin can rebalance or assign needs directly.</p>
      </article>` : ""}

    ${(canEdit || isAdmin) ? `
      <article class="action-card">
        <p class="eyebrow">${isAdmin ? "Admin Curation Update" : "Curator Update"}</p>
        <h4>${isAdmin ? "Update Any Curation Directly" : "Update Your Assigned Need"}</h4>
        <form id="directUpdateForm" class="action-form">
          <label>
            <span>Curator</span>
            <input name="curator_name" value="${esc(curator?.display_name || state.userSession?.full_name || "")}" readonly />
          </label>
          <label>
            <span>Curator Email</span>
            <input name="curator_email" value="${esc(curator?.email || state.userSession?.email || "")}" readonly />
          </label>
          <label>
            <span>Updated Status</span>
            <select name="proposed_status">
              <option value="">Current: ${esc(currentStatusLabel)}</option>
              ${statusOptions.map((item) => `<option value="${esc(item.label)}">${esc(item.label)}</option>`).join("")}
            </select>
          </label>
          <label>
            <span>Updated Internal Status</span>
            <select name="proposed_internal_status">
              <option value="">Current: ${esc(currentInternalStatusLabel)}</option>
              ${internalOptions.map((item) => `<option value="${esc(item.label)}">${esc(item.label)}</option>`).join("")}
            </select>
          </label>
          <label>
            <span>Updated Next Action</span>
            <select name="proposed_next_action">
              <option value="">Current: ${esc(currentNextActionLabel)}</option>
              ${nextActionOptions.map((item) => `<option value="${esc(item.label)}">${esc(item.label)}</option>`).join("")}
            </select>
          </label>
          <label>
            <span>Curation Call Date</span>
            <input name="proposed_curation_call_date" type="date" />
          </label>
          <label class="wide">
            <span>Curation Notes</span>
            <textarea name="proposed_curation_notes" rows="4" placeholder="Describe what changed and why."></textarea>
          </label>
          <label>
            <span>Broadcast Needed</span>
            <select name="proposed_demand_broadcast_needed">
              <option value="">Current: ${need.demand_broadcast_needed ? "Yes" : "No"}</option>
              <option value="true">Yes</option>
              <option value="false">No</option>
            </select>
          </label>
          <div class="wide">
            <button class="btn btn-primary" type="submit">${isAdmin ? "Save Curation Update" : "Save and Sync to GRE"}</button>
            <p class="helper-text">${isAdmin ? "Admin can directly update any curator's curation details." : "Your curation update will sync to the GRE website immediately."}</p>
          </div>
        </form>
      </article>` : ""}

    ${(isCurator && !canEdit) ? `
      <article class="action-card">
        <p class="eyebrow">Curator Access</p>
        <h4>Need Owned by Another Curator</h4>
        <p class="helper-text">You can inspect the need and mail the assigned curator or a provider, but only the assigned curator can edit curation here.</p>
      </article>` : ""}

    ${(isAdmin || isCurator) ? `
      <article class="action-card">
        <p class="eyebrow">Provider Outreach</p>
        <h4>Send Need Details by Email</h4>
        <form id="manualProviderForm" class="action-form">
          <label>
            <span>Provider Email</span>
            <input name="provider_email" type="email" placeholder="provider@example.org" required />
          </label>
          <div class="wide">
            <button class="btn btn-secondary" type="submit">Send Need Details</button>
          </div>
        </form>
        <p class="helper-text">The mail is sent from help@greenruraleconomy.in, with the seeker and actual curator copied.</p>
      </article>` : ""}

    ${canMessageCurator ? `
      <article class="action-card">
        <p class="eyebrow">Message Curator</p>
        <h4>Send a Custom Message</h4>
        <form id="curatorMessageForm" class="action-form">
          <label class="wide">
            <span>Message to ${esc(curator?.display_name || "the curator")}</span>
            <textarea name="message" rows="4" placeholder="Write a message to the curator about this need." required></textarea>
          </label>
          <div class="wide">
            <button class="btn btn-secondary" type="submit">Mail Curator</button>
          </div>
        </form>
      </article>` : ""}
  `;
}

function getChosenPuterModel() {
  return String(byId("puterModelSelect")?.value || "").trim() || null;
}

async function ensurePuterModelsLoaded(forceRefresh = false) {
  if (!window.puter?.ai) {
    throw new Error("Puter AI is not available on this page.");
  }
  if (state.puterModelsLoaded && !forceRefresh) return state.puterModels;
  const status = byId("puterStatus");
  if (status) status.textContent = "Loading Puter models...";
  const result = await window.puter.ai.listModels();
  const models = normalizePuterModelEntries(result);
  state.puterModels = models;
  state.puterModelsLoaded = true;
  const select = byId("puterModelSelect");
  if (select) {
    const previous = getChosenPuterModel();
    select.innerHTML = `<option value="">Default Puter model</option>`;
    models.slice(0, 200).forEach((model) => {
      const option = document.createElement("option");
      option.value = model.id;
      option.textContent = model.name;
      select.appendChild(option);
    });
    if (previous && models.some((model) => model.id === previous)) {
      select.value = previous;
    }
  }
  if (status) status.textContent = models.length ? `Loaded ${models.length} Puter model options.` : "No Puter models were returned.";
  return models;
}

async function runPuterChat(prompt) {
  await ensurePuterModelsLoaded(false);
  const model = getChosenPuterModel();
  const options = model ? { model } : {};
  const response = await window.puter.ai.chat(prompt, options);
  return extractPuterText(response);
}

function buildPuterNeedReviewPrompt(need) {
  return `
You are reviewing a rural-economy inbound need for classification quality control.
Return only valid JSON.

Need record:
${JSON.stringify({
  id: need.id,
  organization_name: need.organization_name,
  state: need.state,
  district: need.district,
  curated_need: parseArray(need.curated_need),
  problem_statement: need.problem_statement,
  curation_notes: need.curation_notes,
  current_rule_thematic_hints: parseArray(need.rule_thematic_hints),
  current_ai_thematic_area: need.ai_thematic_area,
  current_ai_application_area: need.ai_application_area,
  current_ai_need_kind: need.ai_need_kind,
  current_ai_service_kind: need.ai_service_kind,
  current_ai_6m_signals: parseArray(need.ai_6m_signals),
  current_ai_keywords: parseArray(need.ai_keywords),
}, null, 2)}

Classify the need into:
- thematic_area: the core theme like dairy, soap, solar, goatery
- application_area: narrower operational area
- need_kind: one of service, product, knowledge, finance, mixed
- service_kind: only if need_kind is service
- six_m_signals: any of Manpower, Method, Machine, Material, Market, Money
- keywords: 6 to 12 precise search keywords
- summary: one concise operational summary
- conflict_reason: explain briefly why current stored classification may be weak or wrong

Rules:
- prioritize the core thematic area over generic secondary needs like branding or packaging
- identify deployable need form clearly
- output arrays for six_m_signals and keywords
- output empty string for service_kind if not service

JSON schema:
{
  "thematic_area": "",
  "application_area": "",
  "need_kind": "",
  "service_kind": "",
  "six_m_signals": [],
  "keywords": [],
  "summary": "",
  "conflict_reason": ""
}`.trim();
}

function normalizeNeedReviewSuggestion(payload) {
  return {
    thematic_area: normalizeText(payload?.thematic_area),
    application_area: normalizeText(payload?.application_area),
    need_kind: normalizeText(payload?.need_kind).toLowerCase(),
    service_kind: normalizeText(payload?.service_kind).toLowerCase(),
    six_m_signals: uniq(parseArray(payload?.six_m_signals)),
    keywords: uniq(parseArray(payload?.keywords).map((item) => item.toLowerCase())).slice(0, 12),
    summary: normalizeText(payload?.summary),
    conflict_reason: normalizeText(payload?.conflict_reason),
  };
}

function getNeedRecommendation(needId) {
  return state.puterRecommendations?.[needId] || null;
}

async function runPuterNeedReview(needId) {
  const need = state.data.aiReviewNeeds.find((item) => item.id === needId) || state.data.needs.find((item) => item.id === needId);
  if (!need) throw new Error("Need not found for Puter review.");
  const text = await runPuterChat(buildPuterNeedReviewPrompt(need));
  const suggestion = normalizeNeedReviewSuggestion(parseJsonObject(text));
  state.puterRecommendations[needId] = suggestion;
  renderAdminQueue();
  return suggestion;
}

function renderAdminQueue() {
  const pendingNeedsList = byId("pendingNeedsList");
  const pendingUpdatesList = byId("pendingUpdatesList");
  const pendingFormSubmissionsList = byId("pendingFormSubmissionsList");
  const aiReviewList = byId("aiReviewList");
  if (!pendingNeedsList || !pendingUpdatesList || !pendingFormSubmissionsList || !aiReviewList) return;

  pendingNeedsList.innerHTML = isAdminUser()
    ? state.data.pendingNeeds.length
      ? state.data.pendingNeeds
          .map(
            (need) => `
              <article class="approval-card">
                <div class="status-row">
                  <span class="status-pill warn">Pending intake</span>
                  <span class="status-pill info">${esc(need.source_kind || "manual")}</span>
                </div>
                <h4>${esc(need.organization_name)}</h4>
                <p class="helper-text">${esc(need.problem_statement.slice(0, 180))}${need.problem_statement.length > 180 ? "..." : ""}</p>
                <p class="meta-text">${esc(formatDate(need.requested_on))}</p>
                <div class="card-actions">
                  <button class="btn btn-primary" data-action="approve-need" data-need-id="${esc(need.id)}">Approve</button>
                  <button class="btn btn-danger" data-action="reject-need" data-need-id="${esc(need.id)}">Reject</button>
                </div>
              </article>
            `,
          )
          .join("")
      : `<div class="empty-state">No intake records are waiting for admin approval.</div>`
    : `<div class="empty-state">Login as admin to view the intake approval queue.</div>`;

  pendingUpdatesList.innerHTML = isAdminUser()
    ? state.data.pendingUpdates.length
      ? state.data.pendingUpdates
          .map(
            (request) => `
              <article class="approval-card">
                <div class="status-row">
                  <span class="status-pill warn">Pending status update</span>
                  <span class="status-pill info">${esc(request.submitted_by_curator_name || request.submitted_by_curator_email)}</span>
                </div>
                <h4>Need #${esc(request.need_id)}</h4>
                <div class="detail-list">
                  ${request.proposed_status ? `<div><strong>Status:</strong> ${esc(request.proposed_status)}</div>` : ""}
                  ${request.proposed_internal_status ? `<div><strong>Internal status:</strong> ${esc(request.proposed_internal_status)}</div>` : ""}
                  ${request.proposed_next_action ? `<div><strong>Next action:</strong> ${esc(request.proposed_next_action)}</div>` : ""}
                  ${request.proposed_curation_notes ? `<div><strong>Notes:</strong> ${esc(request.proposed_curation_notes)}</div>` : ""}
                </div>
                <div class="card-actions">
                  <button class="btn btn-primary" data-action="approve-update" data-request-id="${esc(request.id)}">Approve</button>
                  <button class="btn btn-danger" data-action="reject-update" data-request-id="${esc(request.id)}">Reject</button>
                </div>
              </article>
            `,
          )
          .join("")
      : `<div class="empty-state">No curator update requests are waiting for approval.</div>`
      : `<div class="empty-state">Login as admin to view curator-submitted status changes.</div>`;

  pendingFormSubmissionsList.innerHTML = isAdminUser()
    ? ensureList(state.data.pendingFormSubmissions).length
      ? ensureList(state.data.pendingFormSubmissions)
            .map((submission) => {
              const payload = submission.payload && typeof submission.payload === "object" ? submission.payload : {};
              const isNeed = submission.submission_type === "need";
              const summaryText = isNeed
                ? normalizeText(payload.problem_statement || "")
                : normalizeText(payload.about_offering_text || payload.about_solution_text || payload.solution_name || payload.offering_name || "");
              const categoryLine = isNeed
                ? [payload.need_kind, payload.service_kind, payload.thematic_area, payload.application_area].filter(Boolean).join(" / ")
                : [
                    payload.offering_category,
                    payload.offering_type,
                    payload.primary_valuechain,
                    payload.primary_application,
                  ].filter(Boolean).join(" / ");
              return `
                <article class="approval-card">
                <div class="status-row">
                  <span class="status-pill warn">${esc(isNeed ? "Need submission" : "Solution submission")}</span>
                  <span class="status-pill info">${esc(submission.source_mode || "shared_link")}</span>
                </div>
                <h4>${esc(submission.organization_name || payload.organization_name || "Untitled submission")}</h4>
                  <div class="detail-list">
                    <div><strong>Submitter:</strong> ${esc(submission.submitter_name || submission.submitter_email || "Anonymous shared link")}</div>
                    <div><strong>GRE Supplier:</strong> ${esc(submission.existing_trader_name || "Not selected")}</div>
                    ${!isNeed && payload.solution_name ? `<div><strong>Solution:</strong> ${esc(payload.solution_name)}</div>` : ""}
                    ${!isNeed && payload.offering_name ? `<div><strong>Offering:</strong> ${esc(payload.offering_name)}</div>` : ""}
                    ${categoryLine ? `<div><strong>Classification:</strong> ${esc(categoryLine)}</div>` : ""}
                    ${summaryText ? `<div><strong>Summary:</strong> ${esc(clipText(summaryText, 220))}</div>` : ""}
                    ${!isNeed ? `<div><strong>GRE Sync:</strong> ${esc(submission.gre_sync_status ? submission.gre_sync_status.replaceAll("_", " ") : "Pending admin review")}</div>` : ""}
                  </div>
                <div class="card-actions">
                  <button class="btn btn-secondary" data-action="edit-form-submission" data-submission-id="${esc(submission.id)}">View / Edit</button>
                  <button class="btn btn-primary" data-action="approve-form-submission" data-submission-id="${esc(submission.id)}">Approve</button>
                  <button class="btn btn-danger" data-action="reject-form-submission" data-submission-id="${esc(submission.id)}">Reject</button>
                </div>
              </article>
            `;
          })
          .join("")
      : `<div class="empty-state">No shared solution or need forms are waiting for admin review.</div>`
    : `<div class="empty-state">Login as admin to review shared form submissions.</div>`;

  aiReviewList.innerHTML = isAdminUser()
    ? state.data.aiReviewNeeds.length
      ? state.data.aiReviewNeeds
          .map(
            (need) => {
              const recommendation = getNeedRecommendation(need.id);
              const currentTheme = getEffectiveNeedThematicArea(need) || "Not set";
              const currentApplication = getEffectiveNeedApplicationArea(need) || "Not set";
              const currentNeedKind = getEffectiveNeedKind(need) || "Not set";
              const currentServiceKind = getEffectiveNeedServiceKind(need);
              const current6M = getEffectiveNeed6MSignals(need).join(", ") || parseArray(need.rule_6m_signals).join(", ") || "None";
              const overrideBadge = need.override_updated_at ? `<span class="status-pill good">Override active</span>` : "";
              const recommendationBlock = recommendation
                ? `
                  <div class="qa-suggestion-card">
                    <div class="status-row">
                      <span class="status-pill info">Puter suggestion</span>
                      ${recommendation.conflict_reason ? `<span class="helper-text">${esc(recommendation.conflict_reason)}</span>` : ""}
                    </div>
                    <div class="detail-list qa-suggestion-list">
                      <div><strong>Theme:</strong> ${esc(recommendation.thematic_area || "Not set")} <button class="btn btn-secondary btn-mini" data-action="accept-need-override" data-need-id="${esc(need.id)}" data-override-field="thematic_area">Accept</button></div>
                      <div><strong>Application:</strong> ${esc(recommendation.application_area || "Not set")} <button class="btn btn-secondary btn-mini" data-action="accept-need-override" data-need-id="${esc(need.id)}" data-override-field="application_area">Accept</button></div>
                      <div><strong>Need Kind:</strong> ${esc(recommendation.need_kind || "Not set")}${recommendation.service_kind ? ` / ${esc(recommendation.service_kind)}` : ""} <button class="btn btn-secondary btn-mini" data-action="accept-need-override" data-need-id="${esc(need.id)}" data-override-field="need_kind">Accept</button></div>
                      <div><strong>6M:</strong> ${esc(recommendation.six_m_signals.join(", ") || "None")} <button class="btn btn-secondary btn-mini" data-action="accept-need-override" data-need-id="${esc(need.id)}" data-override-field="six_m_signals">Accept</button></div>
                      <div><strong>Keywords:</strong> ${esc(recommendation.keywords.join(", ") || "None")} <button class="btn btn-secondary btn-mini" data-action="accept-need-override" data-need-id="${esc(need.id)}" data-override-field="keywords">Accept</button></div>
                    </div>
                    <p class="helper-text">${esc(recommendation.summary || "")}</p>
                    <div class="card-actions">
                      <button class="btn btn-primary" data-action="accept-need-override" data-need-id="${esc(need.id)}" data-override-field="all">Accept All</button>
                    </div>
                  </div>
                `
                : `<div class="card-actions"><button class="btn btn-secondary" data-action="run-puter-need-review" data-need-id="${esc(need.id)}">Puter Suggest Review</button></div>`;
              return `
                <article class="approval-card">
                  <div class="status-row">
                    <span class="status-pill ${need.ai_validation_status === "flagged" ? "warn" : "info"}">${esc(need.ai_validation_status || "needs review")}</span>
                    <span class="status-pill info">${esc(`Confidence ${need.ai_confidence ?? 0}`)}</span>
                    <span class="status-pill ${badgeTone(need.ai_enrichment_status)}">${esc(formatAiEnrichmentStatus(need.ai_enrichment_status))}</span>
                    ${overrideBadge}
                  </div>
                  <h4>${esc(need.organization_name)}</h4>
                  <p class="helper-text">${esc(clipText(need.problem_statement, 170))}</p>
                  <div class="detail-list">
                    <div><strong>Rule Themes:</strong> ${esc(parseArray(need.rule_thematic_hints).join(", ") || "None")}</div>
                    <div><strong>Current Theme:</strong> ${esc(currentTheme)}</div>
                    <div><strong>Current Application:</strong> ${esc(currentApplication)}</div>
                    <div><strong>Current Need Kind:</strong> ${esc(currentNeedKind)}${currentServiceKind ? ` / ${esc(currentServiceKind)}` : ""}</div>
                    <div><strong>Current 6M:</strong> ${esc(current6M)}</div>
                    <div><strong>Flags:</strong> ${esc(parseArray(need.ai_validation_flags).join(", ") || "None")}</div>
                  </div>
                  ${recommendationBlock}
                </article>
              `;
            },
          )
          .join("")
      : `<div class="empty-state">No approved needs are currently flagged for AI review.</div>`
    : `<div class="empty-state">Login as admin to view the AI review queue.</div>`;

}

function renderUserManagement() {
  const list = byId("userManagementList");
  if (!list) return;
  if (!isAdminUser()) {
    list.innerHTML = `<div class="empty-state">Login as admin to manage users.</div>`;
    return;
  }
  const tabs = [
    ["admin", "Admin"],
    ["curator", "Curator"],
    ["user", "Users"],
  ];
  const filteredUsers = ensureList(state.data.users).filter((user) => normalizeText(user.role).toLowerCase() === state.adminUserRoleTab);
  list.innerHTML = `
    <div class="admin-role-tabs">
      ${tabs.map(([role, label]) => `
        <button class="btn ${state.adminUserRoleTab === role ? "btn-primary" : "btn-secondary"}" type="button" data-action="set-user-role-tab" data-role-tab="${esc(role)}">
          ${esc(label)}
        </button>
      `).join("")}
    </div>
    ${
      filteredUsers.length
        ? `<div class="user-management-grid">${filteredUsers.map((user) => `
          <article class="approval-card user-management-card">
            <div class="status-row">
              <span class="status-pill ${badgeTone(user.role)}">${esc(user.role)}</span>
              <span class="status-pill info">${esc(user.is_active ? "Active" : "Inactive")}</span>
              ${user.gre_sync_status ? `<span class="status-pill ${badgeTone(user.gre_sync_status)}">${esc(user.gre_sync_status.replaceAll("_", " "))}</span>` : ""}
              ${user.gre_pending_role ? `<span class="status-pill warning">Pending ${esc(user.gre_pending_role)}</span>` : ""}
            </div>
            <h4>${esc(user.full_name || user.first_name || user.username)}</h4>
            <div class="detail-list user-detail-grid">
              <div><strong>Username:</strong> ${esc(user.username)}</div>
              <div><strong>Email:</strong> ${esc(user.email || "Not set")}</div>
              <div><strong>Phone:</strong> ${esc(user.phone || "Not set")}</div>
              <div><strong>Role:</strong> ${esc(user.role || "Not set")}</div>
            </div>
            <div class="detail-list user-detail-grid user-sync-grid">
              <div><strong>GRE Login:</strong> ${esc(user.gre_login_name || "Not mapped")}</div>
              <div><strong>GRE User ID:</strong> ${esc(user.gre_user_id || "Not mapped")}</div>
              <div><strong>GRE Sync:</strong> ${esc(user.gre_sync_status ? user.gre_sync_status.replaceAll("_", " ") : "Not attempted")}</div>
              <div><strong>Synced At:</strong> ${esc(user.gre_synced_at ? formatDate(user.gre_synced_at) : "Not synced")}</div>
            </div>
            ${user.gre_pending_role ? `
              <div class="detail-list user-detail-grid user-sync-grid">
                <div><strong>Pending GRE Role:</strong> ${esc(user.gre_pending_role)}</div>
                <div><strong>Activation Requested:</strong> ${esc(user.gre_activation_requested_at ? formatDate(user.gre_activation_requested_at) : "Not started")}</div>
                <div><strong>OTP Status:</strong> ${esc(user.gre_sync_status === "awaiting_gre_activation" ? "Awaiting OTP" : "Not required")}</div>
                <div><strong>GRE Temp Password:</strong> ${esc(user.gre_pending_role ? "gre@1234" : "-")}</div>
              </div>
            ` : ""}
            ${user.gre_sync_message ? `<p class="helper-text">${esc(user.gre_sync_message)}</p>` : ""}
            <div class="card-actions">
              <label class="inline-select">
                <span class="helper-text">Assign role</span>
                <select data-role-select data-user-id="${esc(user.id)}">
                  ${["user", "curator", "admin"].map((role) => `<option value="${role}" ${user.role === role ? "selected" : ""}>${role[0].toUpperCase()}${role.slice(1)}</option>`).join("")}
                </select>
              </label>
              <button class="btn btn-primary" data-action="update-user-role" data-user-id="${esc(user.id)}">Update Role</button>
              <button class="btn btn-secondary" data-action="remove-user" data-user-id="${esc(user.id)}">Remove User...</button>
            </div>
            ${user.gre_sync_status === "awaiting_gre_activation" ? `
              <div class="card-actions activation-actions">
                <label class="inline-select inline-input">
                  <span class="helper-text">Enter GRE OTP</span>
                  <input type="text" maxlength="8" placeholder="Enter OTP" data-user-otp data-user-id="${esc(user.id)}" />
                </label>
                <button class="btn btn-secondary" data-action="complete-user-role-activation" data-user-id="${esc(user.id)}">Complete GRE Activation</button>
              </div>
            ` : ""}
          </article>
        `).join("")}</div>`
        : `<div class="empty-state">No ${esc(state.adminUserRoleTab)} records found.</div>`
    }
  `;
}

function getLocalSolutionByOfferingId(offeringId) {
  return ensureList(state.data.localSolutions).find((item) => String(item.offering_id) === String(offeringId)) || null;
}

function getLocalSolutionAttachmentConfigs(solution) {
  const category = normalizeText(solution?.offering_category || "");
  const configs = [];
  if (category === "Product offerings") {
    configs.push({
      label: "Product Brochure",
      url: solution?.product_brochure_url || "",
      inputName: "local_product_brochure_attachment",
      accept: ".pdf,.doc,.docx,.ppt,.pptx,.jpg,.jpeg,.png",
    });
  }
  if (category === "Service offerings") {
    configs.push({
      label: "Service Brochure",
      url: solution?.service_brochure_url || "",
      inputName: "local_service_brochure_attachment",
      accept: ".pdf,.doc,.docx,.ppt,.pptx,.jpg,.jpeg,.png",
    });
  }
  if (category === "Knowledge offerings") {
    configs.push({
      label: "Knowledge Content",
      url: solution?.knowledge_content_url || "",
      inputName: "local_knowledge_content_attachment",
      accept: ".pdf,.doc,.docx,.ppt,.pptx,.xls,.xlsx,.csv,.jpg,.jpeg,.png,.mp4",
    });
  }
  configs.push({
    label: "Offering Image",
    url: solution?.solution?.solution_image_url || "",
    inputName: "local_offering_image_attachment",
    accept: ".jpg,.jpeg,.png,.webp",
  });
  return configs;
}

function renderLocalSolutionReviewAttachments(solution) {
  const target = byId("localSolutionReviewAttachments");
  if (!target) return;
  const configs = getLocalSolutionAttachmentConfigs(solution);
  target.innerHTML = configs.map((config) => `
    <article class="attachment-review-card">
      <div class="attachment-review-copy">
        <strong>${esc(config.label)}</strong>
        ${
          config.url
            ? `<a href="${esc(config.url)}" target="_blank" rel="noreferrer">Open current file</a>`
            : `<span class="helper-text">No file attached yet.</span>`
        }
      </div>
      <label>
        <span>Replace File</span>
        <input name="${escAttr(config.inputName)}" type="file" accept="${escAttr(config.accept)}" />
      </label>
    </article>
  `).join("");
}

function getLocalSolutionProviders() {
  return uniq(
    ensureList(state.data.localSolutions)
      .map((item) => normalizeText(item?.trader?.organisation_name || item?.trader?.trader_name || ""))
      .filter(Boolean),
  ).sort((a, b) => a.localeCompare(b));
}

function getFilteredLocalSolutions() {
  const search = normalizeText(state.localSolutionFilters.search).toLowerCase();
  const provider = normalizeText(state.localSolutionFilters.provider);
  return ensureList(state.data.localSolutions).filter((item) => {
    const providerName = normalizeText(item?.trader?.organisation_name || item?.trader?.trader_name || "");
    const haystack = [
      normalizeText(item.offering_name),
      normalizeText(item?.solution?.solution_name),
      providerName,
      ...parseArray(item.tags),
      normalizeText(item.about_offering_text),
    ].join(" ").toLowerCase();
    const providerMatch = provider === "all" || providerName === provider;
    const searchMatch = !search || haystack.includes(search);
    return providerMatch && searchMatch;
  });
}

function openLocalSolutionReviewDialog(offeringId) {
  const solution = getLocalSolutionByOfferingId(offeringId);
  if (!solution) {
    toast("This local solution is no longer available.");
    return;
  }
  state.localSolutionReviewId = offeringId;
  const dialog = byId("localSolutionReviewDialog");
  const form = byId("localSolutionReviewForm");
  if (!dialog || !form) return;
  const provider = solution.trader || {};
  const linkedSolution = solution.solution || {};
  const meta = byId("localSolutionReviewMeta");
  if (meta) {
    meta.innerHTML = `
      <div><strong>Offering ID:</strong> ${esc(solution.offering_id)}</div>
      <div><strong>Solution ID:</strong> ${esc(solution.solution_id)}</div>
      <div><strong>Provider:</strong> ${esc(provider.organisation_name || provider.trader_name || "Not set")}</div>
      <div><strong>Status:</strong> ${esc(solution.publish_status || "Not set")}</div>
    `;
  }
  const setValue = (name, value) => {
    const field = form.querySelector(`[name="${name}"]`);
    if (field) field.value = value || "";
  };
  setValue("offering_id", solution.offering_id);
  setValue("organization_name", provider.organisation_name || provider.trader_name || "");
  setValue("submitter_name", provider.poc_name || "");
  setValue("submitter_email", provider.email || "");
  setValue("submitter_phone", provider.mobile || "");
  setValue("solution_name", linkedSolution.solution_name || "");
  setValue("about_solution_text", linkedSolution.about_solution_text || "");
  setValue("offering_category", solution.offering_category || "");
  setValue("offering_type", solution.offering_type || "");
  setValue("offering_name", solution.offering_name || "");
  setValue("about_offering_text", solution.about_offering_text || "");
  setValue("trainer_name", solution.trainer_name || "");
  setValue("trainer_email", solution.trainer_email || "");
  setValue("trainer_phone", solution.trainer_phone || "");
  setValue("trainer_details_text", solution.trainer_details_text || "");
  setValue("languages", parseArray(solution.languages).join(", "));
  setValue("geographies", parseArray(solution.geographies).join("; "));
  setValue("service_cost", solution.service_cost || "");
  setValue("product_cost", solution.product_cost || "");
  setValue("lead_time", solution.lead_time || "");
  setValue("support_details", solution.support_details || "");
  setValue("contact_details", solution.contact_details || "");
  setValue("tags", parseArray(solution.tags).join(", "));
  renderLocalSolutionReviewAttachments(solution);
  byId("localSolutionReviewStatus").textContent = "";
  dialog.showModal();
}

function renderLocalSolutionManagement() {
  const list = byId("localSolutionsList");
  const providerSelect = byId("localSolutionProviderFilter");
  const searchInput = byId("localSolutionSearch");
  if (!list || !providerSelect || !searchInput) return;
  if (!isAdminUser()) {
    list.innerHTML = `<div class="empty-state">Login as admin to manage local Supabase solutions.</div>`;
    providerSelect.innerHTML = `<option value="all">All solution providers</option>`;
    searchInput.value = "";
    return;
  }
  const providers = getLocalSolutionProviders();
  providerSelect.innerHTML = [`<option value="all">All solution providers</option>`, ...providers.map((item) => `<option value="${esc(item)}">${esc(item)}</option>`)].join("");
  providerSelect.value = state.localSolutionFilters.provider || "all";
  searchInput.value = state.localSolutionFilters.search || "";
  const rows = getFilteredLocalSolutions();
  list.innerHTML = rows.length
    ? `<div class="local-solutions-grid">${rows.map((item) => {
        const providerName = item?.trader?.organisation_name || item?.trader?.trader_name || "Unmapped provider";
        return `
          <article class="approval-card local-solution-card">
            <div class="status-row">
              <span class="status-pill info">${esc(item.offering_category || "Offering")}</span>
              <span class="status-pill good">${esc(item.publish_status || "MIS Published")}</span>
            </div>
            <h4>${esc(item.offering_name || item?.solution?.solution_name || "Untitled offering")}</h4>
            <div class="detail-list">
              <div><strong>Solution:</strong> ${esc(item?.solution?.solution_name || "Not set")}</div>
              <div><strong>Provider:</strong> ${esc(providerName)}</div>
              <div><strong>Type:</strong> ${esc(item.offering_type || "Not set")}</div>
              <div><strong>Tags:</strong> ${esc(parseArray(item.tags).join(", ") || "None")}</div>
            </div>
            <p class="helper-text">${esc(clipText(item.about_offering_text || "", 180) || "No summary available.")}</p>
            <div class="card-actions">
              <button class="btn btn-secondary" type="button" data-action="edit-local-solution" data-offering-id="${esc(item.offering_id)}">Edit</button>
              <button class="btn btn-danger" type="button" data-action="delete-local-solution" data-offering-id="${esc(item.offering_id)}">Delete</button>
            </div>
          </article>
        `;
      }).join("")}</div>`
    : `<div class="empty-state">No local Supabase-only solutions match the current filters.</div>`;
}

function chooseUserRemovalMode(name) {
  const choice = window.prompt(
    `Choose removal mode for ${name}:\n1 = Remove from GRE organisation only and delete from MIS\n2 = Remove GRE account completely and delete from MIS\n\nType 1 or 2.`,
    "1",
  );
  if (choice == null) return null;
  const normalized = choice.trim();
  if (normalized === "1") return "org_only";
  if (normalized === "2") return "full_account";
  toast("Enter 1 or 2 to continue.");
  return null;
}

function renderAuthState() {
  const authGate = byId("authGate");
  const appShell = byId("appShell");
  const masthead = document.querySelector(".masthead");
  const heroStrip = byId("heroStrip");
  const overviewTab = document.querySelector('.tab[data-view="overview"]');
  const operationsTab = document.querySelector('.tab[data-view="operations"]');
  const solutionTab = document.querySelector('.tab[data-view="solution"]');
  const needTab = document.querySelector('.tab[data-view="need-intake"]');
  const roleTab = document.querySelector('.tab[data-view="admin"]');
  const newNeedBtn = byId("newNeedBtn");
  const accountName = byId("accountName");
  const accountRoleChip = byId("accountRoleChip");
  const accountSummary = byId("accountSummary");
  const sharedMode = isSharedFormMode();
  const loggedIn = isLoggedIn();
  const showShell = loggedIn || sharedMode;
  const isAdminDatasetVisible = isAdminUser() && state.view === "admin";

  if (authGate) authGate.classList.toggle("hidden", showShell);
  if (appShell) appShell.classList.toggle("hidden", !showShell);
  masthead?.classList.toggle("hidden", sharedMode);
  heroStrip?.classList.toggle("hidden", sharedMode || !isAdminDatasetVisible);
  overviewTab?.classList.toggle("hidden", sharedMode);
  operationsTab?.classList.toggle("hidden", sharedMode || !canAccessOperationsDesk());
  solutionTab?.classList.toggle("hidden", sharedMode && state.sharedFormMode !== "solution");
  needTab?.classList.toggle("hidden", sharedMode && state.sharedFormMode !== "need-intake");
  roleTab?.classList.toggle("hidden", sharedMode || !isAdminUser());
  if (newNeedBtn) newNeedBtn.classList.toggle("hidden", !loggedIn);
  if (sharedMode) {
    state.view = state.sharedFormMode;
  } else if (!loggedIn) {
    state.view = "overview";
  } else if (!isAdminUser() && state.view === "admin") {
    state.view = canAccessOperationsDesk() ? "operations" : "overview";
  } else if (!canAccessOperationsDesk() && state.view === "operations") {
    state.view = "overview";
  }

  if (accountName) {
    accountName.textContent = state.userSession
      ? (state.userSession.full_name || state.userSession.first_name || state.userSession.username || "GRE MIS User")
      : "GRE MIS User";
  }
  if (accountRoleChip) {
    accountRoleChip.textContent = state.userSession ? state.userSession.role : "User";
    accountRoleChip.className = `chip ${isAdminUser() ? "good" : isCuratorUser() ? "info" : "muted"}`;
  }
  if (accountSummary) {
    accountSummary.textContent = state.userSession
      ? `${state.userSession.email || ""}${state.userSession.must_change_password ? " • Please change your default password." : ""}`
      : "";
  }
  if (!isLoggedIn()) {
    setText("changePasswordStatus", "");
  }
  document.querySelectorAll(".tab").forEach((button) => button.classList.toggle("active", button.dataset.view === state.view));
  document.querySelectorAll(".view").forEach((section) => section.classList.toggle("active", section.id === `${state.view}View`));
}

function renderAdminState() {
  const chip = byId("adminStatusChip");
  const text = byId("adminStatusText");
  const loginStatus = byId("loginStatus");
  const sessionPanel = byId("sessionPanel");
  if (!chip || !text) return;

  if (isAdminUser()) {
    chip.textContent = "Unlocked";
    chip.className = "chip good";
    text.textContent = `Admin session active for ${state.userSession?.full_name || state.userSession?.username}. Intake approvals, sync controls, and user management are available below.`;
    sessionPanel?.classList.remove("hidden");
  } else {
    chip.textContent = "Locked";
    chip.className = "chip muted";
    text.textContent = "Admin approval is required for new intake records and curator-submitted status updates.";
    sessionPanel?.classList.add("hidden");
  }

  if (loginStatus && !isAdminUser() && !loginStatus.textContent) {
    loginStatus.textContent = "";
  }
}

function switchView(nextView) {
  if (isSharedFormMode()) {
    nextView = state.sharedFormMode;
  }
  if (nextView === "operations" && !canAccessOperationsDesk()) {
    nextView = "overview";
  }
  if (nextView === "admin" && !isAdminUser()) {
    nextView = canAccessOperationsDesk() ? "operations" : "overview";
  }
  state.view = nextView;
  renderAuthState();
  if (nextView === "operations") {
    renderMatches();
  }
}

async function rerender(options = {}) {
  const { includeMatches = state.view === "operations" } = options;
  renderAuthState();
  renderSubmissionViews();
  renderMetrics();
  renderOverview();
  renderFilters();
  renderQueue();
  renderNeedDetail();
  renderWorkbench();
  renderAdminQueue();
  renderUserManagement();
  renderLocalSolutionManagement();
  renderAdminState();
  if (!byId("overviewView") && byId("adminView")) {
    const headline = byId("datasetHeadline");
    const subline = byId("datasetSubline");
      if (headline) headline.textContent = "Admin sync workspace for approvals, GRE refresh, and chatbot data sync";
    if (subline) subline.textContent = `${state.data.pendingNeeds.length} intake records and ${state.data.pendingUpdates.length} curator updates are waiting for review.`;
  }
  if (includeMatches) {
    await renderMatches();
  }
}

async function refreshAll() {
  await Promise.all([
    store.loadBaseData(),
      isAdminUser()
        ? store.loadAdminSnapshot().catch(() => {
            state.data.pendingNeeds = [];
            state.data.pendingUpdates = [];
            state.data.pendingFormSubmissions = [];
            state.data.aiReviewNeeds = [];
            state.data.users = [];
            state.data.localSolutions = [];
          })
        : Promise.resolve().then(() => {
            state.data.pendingNeeds = [];
            state.data.pendingUpdates = [];
            state.data.pendingFormSubmissions = [];
            state.data.aiReviewNeeds = [];
            state.data.users = [];
            state.data.localSolutions = [];
        }),
  ]);
  const displayNeeds = getDisplayNeeds();
  if (!displayNeeds.find((need) => need.id === state.selectedNeedId)) {
    state.selectedNeedId = displayNeeds[0]?.id || null;
  }
  await rerender();
}

function buildNeedOverridePatch(needId, field) {
  const recommendation = getNeedRecommendation(needId);
  if (!recommendation) throw new Error("No recommendation available for this need yet.");
  if (field === "thematic_area") {
    return { override_thematic_area: recommendation.thematic_area || null };
  }
  if (field === "application_area") {
    return { override_application_area: recommendation.application_area || null };
  }
  if (field === "need_kind") {
    return {
      override_need_kind: recommendation.need_kind || null,
      override_service_kind: recommendation.need_kind === "service" ? recommendation.service_kind || null : null,
    };
  }
  if (field === "six_m_signals") {
    return { override_6m_signals: recommendation.six_m_signals || [] };
  }
  if (field === "keywords") {
    return { override_keywords: recommendation.keywords || [] };
  }
  if (field === "all") {
    return {
      override_thematic_area: recommendation.thematic_area || null,
      override_application_area: recommendation.application_area || null,
      override_need_kind: recommendation.need_kind || null,
      override_service_kind: recommendation.need_kind === "service" ? recommendation.service_kind || null : null,
      override_6m_signals: recommendation.six_m_signals || [],
      override_keywords: recommendation.keywords || [],
      override_summary: recommendation.summary || null,
    };
  }
  throw new Error("Unsupported override field.");
}

function resetDashboardSelections() {
  state.overviewFilters = {
    metric: [],
    pipeline: [],
    curator: [],
    state: [],
    category: [],
  };
  state.overviewPage = 1;
  state.matchPage = 1;
  state.filters = {
    status: "all",
    curator: "all",
    state: "all",
    search: "",
  };
}

async function collectSubmissionPayload(form, submissionType = "need") {
  const formData = new FormData(form);
  const entries = {};
  for (const [key, value] of formData.entries()) {
    if (value instanceof File) continue;
    appendSubmissionEntry(entries, key, value);
  }
  if (submissionType !== "solution") {
    const traderId = normalizeText(entries.existing_trader_id);
    const trader = getTraderById(traderId);
    if (!trader) {
      openMissingOrgDialog(entries.organization_name || "");
      throw new Error("Please select an approved GRE supplier before submitting.");
    }
    return {
      ...entries,
      existing_trader_id: trader.trader_id,
      existing_trader_name: trader.organisation_name || trader.trader_name || entries.organization_name || "",
      organization_name: entries.organization_name || trader.organisation_name || trader.trader_name || "",
      keywords: parseArray(entries.keywords),
      six_m_signals: parseArray(entries.six_m_signals),
      curated_need: parseArray(entries.curated_need),
    };
  }

  const traderId = normalizeText(entries.existing_trader_id);
  const selectedTrader = getTraderById(traderId);
  const nameMatchedTrader = findSupplierByName(entries.organization_name || "");
  const trader = selectedTrader || nameMatchedTrader || null;
  const offeringCategory = normalizeText(entries.offering_category) || "Service offerings";
  const offeringGroup = offeringCategory.replace(/\s+offerings$/i, "").trim() || "Service";
  const defaultAudience = ["Individuals", "Groups", "SHGs", "Organisations"];
  const serviceLanguages = normalizeMultiTextValue(entries.languages);
  const tags = uniq([
    ...parseLooseListInput(entries.tags),
    ...normalizeMultiTextValue(entries.keywords),
  ]);
  const geographies = normalizeSemicolonSeparatedValue(entries.geographies);
  const offeringName = normalizeText(entries.offering_name);
  const offeringDescription = normalizeText(entries.about_offering_text);
  const organizationName = normalizeText(entries.organization_name) || trader?.organisation_name || trader?.trader_name || "";
  if (!organizationName) throw new Error("Organisation name is required.");
  if (!normalizeText(entries.submitter_email)) throw new Error("Contact email is required.");
  if (!offeringName) throw new Error("Offering name is required.");
  if (!offeringDescription) throw new Error("Offering description is required.");
  if (!tags.length) throw new Error("Please add at least one tag before submitting.");

  const productBrochure = await collectSingleFileAttachment(form, "product_brochure_upload");
  const serviceBrochure = await collectSingleFileAttachment(form, "service_brochure_upload");
  const knowledgeContent = await collectSingleFileAttachment(form, "knowledge_content_upload");
  const offeringImage = await collectSingleFileAttachment(form, "offering_image_upload");
  const productCostQuoteOnScope = normalizeText(entries.product_cost_quote_on_scope).toLowerCase() === "yes";
  const normalizedProductCost = productCostQuoteOnScope && !normalizeText(entries.product_cost)
    ? "Can be quoted after finalising scope"
    : normalizeText(entries.product_cost);
  const contactDetails = normalizeText(entries.product_contact_details || entries.contact_details || "");

  return {
    ...entries,
    existing_trader_id: trader?.trader_id || "",
    existing_trader_name: trader?.organisation_name || trader?.trader_name || "",
    organization_name: organizationName,
    submitter_email: normalizeText(entries.submitter_email).toLowerCase(),
    offering_category: offeringCategory,
    offering_group: offeringGroup,
    offering_kind: offeringCategory.replace(/\s+offerings$/i, "").toLowerCase(),
    service_kind: offeringCategory === "Service offerings" ? normalizeText(entries.offering_type) : "",
    solution_name: offeringName,
    about_solution_text: offeringDescription,
    about_solution_html: offeringDescription,
    about_offering_html: offeringDescription,
    solution_audience: defaultAudience,
    audience: defaultAudience,
    languages: serviceLanguages,
    geographies,
    location_availability: getCheckedValues(form, "location_availability"),
    tags,
    keywords: tags,
    primary_valuechain: "",
    primary_application: "",
    secondary_valuechain: "",
    secondary_application: "",
    valuechains: [],
    applications: [],
    certification_offered: normalizeText(entries.certification_offered || "Not Provided"),
    product_cost_quote_on_scope: productCostQuoteOnScope,
    product_cost: normalizedProductCost,
    contact_details: contactDetails,
    product_brochure_attachment: productBrochure,
    service_brochure_attachment: serviceBrochure,
    knowledge_content_attachment: knowledgeContent,
    offering_image_attachment: offeringImage,
    knowledge_content_url: normalizeText(entries.knowledge_content_link || ""),
  };
}

function bindStaticEvents() {
  document.querySelectorAll(".tab").forEach((button) => {
    button.addEventListener("click", () => switchView(button.dataset.view));
  });

  byId("statusFilter")?.addEventListener("change", (event) => {
    state.filters.status = event.target.value;
    renderQueue();
    renderNeedDetail();
    renderWorkbench();
    renderMatches();
    renderOverview();
  });
  byId("curatorFilter")?.addEventListener("change", (event) => {
    state.filters.curator = event.target.value;
    renderQueue();
    renderNeedDetail();
    renderWorkbench();
    renderOverview();
  });
  byId("stateFilter")?.addEventListener("change", (event) => {
    state.filters.state = event.target.value;
    renderQueue();
    renderNeedDetail();
    renderWorkbench();
    renderOverview();
  });
  byId("searchFilter")?.addEventListener("input", (event) => {
    state.filters.search = event.target.value.trim();
    renderQueue();
    renderNeedDetail();
    renderWorkbench();
    renderOverview();
  });

  byId("needsQueue")?.addEventListener("click", async (event) => {
    const card = event.target.closest("[data-need-id]");
    if (!card) return;
    state.selectedNeedId = card.dataset.needId;
    state.queueNeedsScrollIntoView = true;
    state.matchPage = 1;
    renderQueue();
    renderNeedDetail();
    renderWorkbench();
    await renderMatches();
  });

  byId("overviewView")?.addEventListener("click", (event) => {
    const card = event.target.closest("[data-overview-kind]");
    if (!card) return;
    setOverviewFocus(card.dataset.overviewKind, card.dataset.overviewId);
    renderMetrics();
    renderOverview();
  });

  byId("pipelineDrilldown")?.addEventListener("click", async (event) => {
    const pageButton = event.target.closest("[data-page-action]");
    if (pageButton) {
      if (pageButton.dataset.pageAction === "prev" && state.overviewPage > 1) state.overviewPage -= 1;
      if (pageButton.dataset.pageAction === "next") state.overviewPage += 1;
      renderOverview();
      return;
    }
    const button = event.target.closest("[data-open-need-id]");
    if (!button) return;
    state.selectedNeedId = button.dataset.openNeedId;
    state.queueNeedsScrollIntoView = true;
    state.matchPage = 1;
    switchView("operations");
    renderQueue();
    renderNeedDetail();
    renderWorkbench();
    await renderMatches();
  });

  byId("pipelineDrilldown")?.addEventListener("click", async (event) => {
    const closeButton = event.target.closest("[data-close-map-panel]");
    if (closeButton) {
      renderCaseMapLocationPanel(null);
      return;
    }
    const button = event.target.closest("[data-open-need-id]");
    if (!button) return;
    await focusNeedFromMap(button.dataset.openNeedId);
  });

  byId("matchResults")?.addEventListener("click", safeAsync(async (event) => {
    const pageButton = event.target.closest("[data-match-page-action]");
    if (pageButton) {
      if (pageButton.dataset.matchPageAction === "prev" && state.matchPage > 1) state.matchPage -= 1;
      if (pageButton.dataset.matchPageAction === "next") state.matchPage += 1;
      await renderMatches();
      return;
    }
  }));

  byId("refreshBtn")?.addEventListener("click", safeAsync(async () => {
    resetDashboardSelections();
    await rerender({ includeMatches: false });
    if (byId("adminView") && isAdminUser()) {
      const provider = byId("aiProviderSelect")?.value || "openrouter";
      const syncStatus = byId("syncStatus");
      if (syncStatus) syncStatus.textContent = "Refreshing AI need intelligence...";
      const result = await store.refreshNeedIntelligence(provider);
      if (syncStatus) syncStatus.textContent = result.message || "AI need intelligence refreshed.";
    }
    await refreshAll();
  }));

  byId("closedToggleBtn")?.addEventListener("click", safeAsync(async () => {
    state.showClosedNeeds = !state.showClosedNeeds;
    state.overviewPage = 1;
    state.matchPage = 1;
    const displayNeeds = getDisplayNeeds();
    if (!displayNeeds.find((need) => need.id === state.selectedNeedId)) {
      state.selectedNeedId = displayNeeds[0]?.id || null;
    }
    await rerender();
  }));

  byId("refreshPuterModelsBtn")?.addEventListener("click", safeAsync(async () => {
    await ensurePuterModelsLoaded(true);
  }));

  const dialog = byId("needDialog");
  const missingOrgDialog = byId("missingOrgDialog");
  const registerDialog = byId("registerDialog");
  const passwordHelpDialog = byId("passwordHelpDialog");
  const workbenchDialog = byId("workbenchDialog");
  const matchDetailDialog = byId("matchDetailDialog");
  const submissionReviewDialog = byId("submissionReviewDialog");
  const localSolutionReviewDialog = byId("localSolutionReviewDialog");
  byId("newNeedBtn")?.addEventListener("click", () => dialog?.showModal());
  byId("closeNeedDialog")?.addEventListener("click", () => dialog?.close());
  byId("closeMissingOrgDialog")?.addEventListener("click", () => missingOrgDialog?.close());
  byId("openRegisterDialogBtn")?.addEventListener("click", () => {
    setText("registerStatus", "");
    registerDialog?.showModal();
  });
  byId("closeRegisterDialog")?.addEventListener("click", () => registerDialog?.close());
  byId("openPasswordHelpDialogBtn")?.addEventListener("click", () => {
    setText("resetStatus", "");
    passwordHelpDialog?.showModal();
  });
  byId("closePasswordHelpDialog")?.addEventListener("click", () => passwordHelpDialog?.close());
  byId("closeWorkbenchDialog")?.addEventListener("click", () => workbenchDialog?.close());
  byId("closeMatchDetailDialog")?.addEventListener("click", () => matchDetailDialog?.close());
  byId("closeSubmissionReviewDialog")?.addEventListener("click", () => submissionReviewDialog?.close());
  byId("closeLocalSolutionReviewDialog")?.addEventListener("click", () => localSolutionReviewDialog?.close());
  byId("submissionReviewTraderSelect")?.addEventListener("change", (event) => {
    const trader = getTraderById(event.target.value);
    const orgInput = byId("submissionReviewForm")?.querySelector('[name="organization_name"]');
    if (trader && orgInput) {
      orgInput.value = trader.organisation_name || trader.trader_name || "";
    }
  });
  ["solution", "need"].forEach((kind) => {
    const selectId = kind === "solution" ? "solutionTraderSelect" : "needTraderSelect";
    const orgInputId = kind === "solution" ? "solutionOrgName" : "needOrgName";
    byId(selectId)?.addEventListener("change", (event) => {
      const trader = getTraderById(event.target.value);
      const orgInput = byId(orgInputId);
      if (orgInput && trader) orgInput.value = trader.organisation_name || trader.trader_name || "";
    });
    byId(orgInputId)?.addEventListener("blur", (event) => {
      const orgName = normalizeText(event.target.value);
      const select = byId(selectId);
      if (!orgName) return;
      const matchingSupplier = findSupplierByName(orgName);
      if (matchingSupplier && select) {
        select.value = matchingSupplier.trader_id;
        event.target.value = matchingSupplier.organisation_name || matchingSupplier.trader_name || orgName;
      } else if (kind !== "solution" && !matchingSupplier && !(select && select.value)) {
        openMissingOrgDialog(orgName);
      }
    });
  });

  byId("needForm")?.addEventListener("submit", safeAsync(async (event) => {
    event.preventDefault();
    const form = new FormData(event.target);
    await store.createNeed(Object.fromEntries(form.entries()));
    event.target.reset();
    dialog?.close();
    await refreshAll();
    toast("Need submitted to the admin approval queue.");
  }));

  byId("solutionSubmissionForm")?.addEventListener("submit", safeAsync(async (event) => {
      event.preventDefault();
      const status = byId("solutionSubmissionStatus");
      if (status) status.textContent = "Submitting solution for admin review...";
      const payload = await collectSubmissionPayload(event.target, "solution");
      const result = isLoggedIn()
        ? await store.submitSignedInForm("solution", payload)
        : await store.submitSharedForm("solution", payload);
      event.target.reset();
      state.solutionTags = [];
      state.solutionGeographies = [];
      renderSolutionTagChips();
      renderSolutionGeographyChips();
      if (isSharedFormMode()) {
        event.target.innerHTML = `<div class="empty-state">Solution Sent for Approval</div>`;
      } else {
        fillSupplierSelect("solutionTraderSelect", "solutionOrgName");
        renderSolutionReferenceInputs();
        if (status) status.textContent = result.message || "Solution submission sent to admin review.";
      }
      if (isAdminUser()) await refreshAll();
      toast("Solution submission queued for admin review.");
    }));

  byId("needSubmissionForm")?.addEventListener("submit", safeAsync(async (event) => {
      event.preventDefault();
      const status = byId("needSubmissionStatus");
      if (status) status.textContent = "Submitting need for admin review...";
      const payload = await collectSubmissionPayload(event.target, "need");
      const result = isLoggedIn()
        ? await store.submitSignedInForm("need", payload)
        : await store.submitSharedForm("need", payload);
    event.target.reset();
    fillSupplierSelect("needTraderSelect", "needOrgName");
    if (status) status.textContent = result.message || "Need submission sent to admin review.";
    if (isAdminUser()) await refreshAll();
    toast("Need submission queued for admin review.");
  }));

  byId("shareSolutionFormBtn")?.addEventListener("click", safeAsync(async () => {
    const shareUrl = getShareUrl("solution");
    await navigator.clipboard?.writeText(shareUrl);
    toast("Solution form link copied.");
  }));

  byId("shareNeedFormBtn")?.addEventListener("click", safeAsync(async () => {
      const shareUrl = getShareUrl("need");
      await navigator.clipboard?.writeText(shareUrl);
      toast("Need form link copied.");
    }));

  byId("solutionOfferingCategory")?.addEventListener("change", () => updateSolutionOfferingForm());
  byId("solutionPrimaryValuechain")?.addEventListener("change", () => updateSolutionApplicationOptions());
  byId("solutionSecondaryValuechain")?.addEventListener("change", () => updateSolutionApplicationOptions());
  byId("solutionGeographyEntry")?.addEventListener("input", safeAsync(async (event) => {
    await updateSolutionGeographySuggestions(event.target.value);
  }));
  byId("addSolutionGeographyBtn")?.addEventListener("click", () => {
    const input = byId("solutionGeographyEntry");
    if (!input) return;
    if (addSolutionGeography(input.value)) {
      input.value = "";
      setDatalistOptions("solutionGeographyOptions", []);
    }
  });
  byId("addSolutionTagBtn")?.addEventListener("click", () => {
    const input = byId("solutionTagEntry");
    if (!input) return;
    if (addSolutionTag(input.value)) input.value = "";
  });
  byId("generateSolutionTagsBtn")?.addEventListener("click", safeAsync(async () => {
    const status = byId("solutionSubmissionStatus");
    const form = byId("solutionSubmissionForm");
    if (!form) return;
    if (status) status.textContent = "Generating AI tags from the current offering draft...";
    const draft = {
      organization_name: normalizeText(form.querySelector('[name="organization_name"]')?.value || ""),
      offering_category: normalizeText(form.querySelector('[name="offering_category"]')?.value || ""),
      offering_type: normalizeText(form.querySelector('[name="offering_type"]')?.value || ""),
      offering_name: normalizeText(form.querySelector('[name="offering_name"]')?.value || ""),
      about_offering_text: normalizeText(form.querySelector('[name="about_offering_text"]')?.value || ""),
      trainer_details_text: normalizeText(form.querySelector('[name="trainer_details_text"]')?.value || ""),
      contact_details: normalizeText(form.querySelector('[name="contact_details"]')?.value || form.querySelector('[name="product_contact_details"]')?.value || ""),
      tags: [...state.solutionTags],
    };
    if (!draft.offering_name || !draft.about_offering_text) {
      throw new Error("Please add the offering name and description before generating tags.");
    }
    const result = await store.suggestSolutionTags(draft);
    const generatedTags = parseArray(result.tags);
    const tagsToAdd = generatedTags.length ? generatedTags : deriveLocalSolutionTags(draft);
    tagsToAdd.forEach((tag) => addSolutionTag(tag));
    if (status) status.textContent = generatedTags.length
      ? result.message || "AI tags added. You can remove any tag before submitting."
      : "AI suggestions were unavailable, so local smart tags were added instead.";
  }));
  byId("solutionTagEntry")?.addEventListener("keydown", (event) => {
    if (event.key !== "Enter" && event.key !== ",") return;
    event.preventDefault();
    const input = event.target;
    if (addSolutionTag(input.value)) input.value = "";
  });
  byId("solutionGeographyEntry")?.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") return;
    event.preventDefault();
    const input = event.target;
    if (addSolutionGeography(input.value)) {
      input.value = "";
      setDatalistOptions("solutionGeographyOptions", []);
    }
  });
    document.querySelectorAll('input[name="solution_audience"]').forEach((input) => {
      input.addEventListener("change", () => {
        if ((byId("solutionOfferingCategory")?.value || "Service offerings") === "Product offerings") {
          syncProductAudienceFromSolution();
        }
      });
    });

  byId("appShell")?.addEventListener("click", safeAsync(async (event) => {
    const removeTagButton = event.target.closest("[data-remove-solution-tag]");
    if (removeTagButton) {
      removeSolutionTag(removeTagButton.dataset.removeSolutionTag);
      return;
    }
    const removeGeoButton = event.target.closest("[data-remove-solution-geo]");
    if (removeGeoButton) {
      removeSolutionGeography(removeGeoButton.dataset.removeSolutionGeo);
      return;
    }
    const button = event.target.closest("[data-copy-share-url]");
    if (!button) return;
    await navigator.clipboard?.writeText(button.dataset.copyShareUrl);
    toast("Form link copied.");
  }));

  byId("userLoginForm")?.addEventListener("submit", safeAsync(async (event) => {
    event.preventDefault();
    const form = new FormData(event.target);
    setText("userLoginStatus", "Signing in...");
    await store.userLogin(form.get("identifier"), form.get("password"));
    event.target.reset();
    await refreshAll();
    setText("userLoginStatus", isLoggedIn() ? "Signed in successfully." : "");
  }));

  byId("registerUserForm")?.addEventListener("submit", safeAsync(async (event) => {
    event.preventDefault();
    const form = new FormData(event.target);
    setText("registerStatus", "Creating account...");
    await store.registerUser({
      firstName: form.get("first_name"),
      fullName: form.get("full_name"),
      email: form.get("email"),
      phone: form.get("phone"),
      password: form.get("password"),
    });
    event.target.reset();
    setText("registerStatus", "Account created. You can sign in now.");
    byId("registerDialog")?.close();
  }));

  byId("requestResetForm")?.addEventListener("submit", safeAsync(async (event) => {
    event.preventDefault();
    const form = new FormData(event.target);
    setText("resetStatus", "Sending reset code...");
    const result = await store.requestPasswordReset(form.get("email"));
    setText("resetStatus", result.message || "If the email exists, a reset code has been sent.");
  }));

  byId("resetPasswordForm")?.addEventListener("submit", safeAsync(async (event) => {
    event.preventDefault();
    const form = new FormData(event.target);
    setText("resetStatus", "Resetting password...");
    await store.resetPassword(form.get("email"), form.get("code"), form.get("new_password"));
    event.target.reset();
    setText("resetStatus", "Password reset complete. Please sign in.");
  }));

  byId("changePasswordForm")?.addEventListener("submit", safeAsync(async (event) => {
    event.preventDefault();
    const form = new FormData(event.target);
    const status = byId("changePasswordStatus");
    if (status) status.textContent = "Updating password...";
    await store.changePassword(form.get("current_password"), form.get("new_password"));
    event.target.reset();
    if (status) status.textContent = "Password changed successfully.";
  }));

  byId("localSolutionSearch")?.addEventListener("input", (event) => {
    state.localSolutionFilters.search = event.target.value || "";
    renderLocalSolutionManagement();
  });

  byId("localSolutionProviderFilter")?.addEventListener("change", (event) => {
    state.localSolutionFilters.provider = event.target.value || "all";
    renderLocalSolutionManagement();
  });

  byId("localSolutionReviewForm")?.addEventListener("submit", safeAsync(async (event) => {
    event.preventDefault();
    const form = new FormData(event.target);
    const status = byId("localSolutionReviewStatus");
    if (status) status.textContent = "Saving local solution changes...";
    const offeringId = form.get("offering_id");
    const productBrochure = await collectSingleFileAttachment(event.target, "local_product_brochure_attachment");
    const serviceBrochure = await collectSingleFileAttachment(event.target, "local_service_brochure_attachment");
    const knowledgeContent = await collectSingleFileAttachment(event.target, "local_knowledge_content_attachment");
    const offeringImage = await collectSingleFileAttachment(event.target, "local_offering_image_attachment");
    const payload = {
      organization_name: form.get("organization_name"),
      submitter_name: form.get("submitter_name"),
      submitter_email: form.get("submitter_email"),
      submitter_phone: form.get("submitter_phone"),
      solution_name: form.get("solution_name"),
      about_solution_text: form.get("about_solution_text"),
      offering_category: form.get("offering_category"),
      offering_type: form.get("offering_type"),
      offering_name: form.get("offering_name"),
      about_offering_text: form.get("about_offering_text"),
      trainer_name: form.get("trainer_name"),
      trainer_email: form.get("trainer_email"),
      trainer_phone: form.get("trainer_phone"),
      trainer_details_text: form.get("trainer_details_text"),
      languages: parseCommaList(form.get("languages")),
      geographies: parseDelimitedList(form.get("geographies"), ";"),
      service_cost: form.get("service_cost"),
      product_cost: form.get("product_cost"),
      lead_time: form.get("lead_time"),
      support_details: form.get("support_details"),
      contact_details: form.get("contact_details"),
      tags: parseCommaList(form.get("tags")),
      product_brochure_attachment: productBrochure,
      service_brochure_attachment: serviceBrochure,
      knowledge_content_attachment: knowledgeContent,
      offering_image_attachment: offeringImage,
    };
    const result = await store.updateLocalSolution(offeringId, payload);
    await refreshAll();
    if (status) status.textContent = result.message || "Local solution saved.";
    localSolutionReviewDialog?.close();
    toast(result.message || "Local solution updated.");
  }));

  byId("userLogoutBtn")?.addEventListener("click", safeAsync(async () => {
    await store.userLogout();
    await refreshAll();
  }));

  byId("adminLoginForm")?.addEventListener("submit", safeAsync(async (event) => {
    event.preventDefault();
    const form = new FormData(event.target);
    const loginStatus = byId("loginStatus");
    if (loginStatus) loginStatus.textContent = "Signing in...";
    await store.adminLogin(form.get("username"), form.get("password"));
    event.target.reset();
    await refreshAll();
    if (document.querySelector('.tab[data-view="admin"]')) switchView("admin");
    if (loginStatus) loginStatus.textContent = "Signed in successfully.";
    toast("Admin access unlocked.");
  }));

  byId("adminLogoutBtn")?.addEventListener("click", safeAsync(async () => {
    await store.adminLogout();
    await refreshAll();
    const loginStatus = byId("loginStatus");
    if (loginStatus) loginStatus.textContent = "";
    toast("Admin session closed.");
  }));

  byId("saveOptionBtn")?.addEventListener("click", safeAsync(async () => {
    if (!isAdminUser()) {
      toast("Login as admin to manage taxonomy options.");
      return;
    }
    const optionType = byId("optionType").value;
    const optionLabel = byId("optionLabel").value.trim();
    if (!optionLabel) {
      toast("Enter an option label.");
      return;
    }
    await store.upsertOption(optionType, optionLabel);
    byId("optionLabel").value = "";
    await refreshAll();
    toast("Option saved.");
  }));

  byId("uploadInboundBtn")?.addEventListener("click", safeAsync(async () => {
    if (!isAdminUser()) {
      toast("Login as admin first.");
      return;
    }
    const fileInput = byId("inboundWorkbookFile");
    const syncStatus = byId("syncStatus");
    const file = fileInput?.files?.[0];
    if (syncStatus) syncStatus.textContent = "Reading inbound workbook...";
    const rows = await parseInboundWorkbookFile(file);
    const provider = byId("aiProviderSelect")?.value || "openrouter";
    if (syncStatus) syncStatus.textContent = `Syncing ${rows.length} inbound rows and refreshing AI intelligence...`;
    const result = await store.importInboundWorkbook(file.name, rows, provider);
    await refreshAll();
    if (syncStatus) {
      syncStatus.textContent = `Imported ${result.insertedCount || 0} new needs, updated ${result.updatedCount || 0} existing needs, refreshed AI for ${result.aiUpdatedCount || 0} needs.`;
    }
    toast("Inbound workbook synced.");
  }));

  byId("syncGreInboundsBtn")?.addEventListener("click", safeAsync(async () => {
    if (!isAdminUser()) {
      toast("Login as admin first.");
      return;
    }
    const syncStatus = byId("syncStatus");
    const provider = byId("aiProviderSelect")?.value || "openrouter";
    if (syncStatus) syncStatus.textContent = "Fetching the latest inbound report from the GRE website...";
    const result = await store.syncGreLiveInbounds(provider);
    await refreshAll();
    if (syncStatus) {
      syncStatus.textContent =
        `Live GRE sync complete. Imported ${result.insertedCount || 0} new needs, updated ${result.updatedCount || 0} existing needs, and refreshed AI for ${result.aiUpdatedCount || 0} needs.`;
    }
    toast("Latest GRE inbound data synced.");
  }));

  byId("refreshAiBtn")?.addEventListener("click", safeAsync(async () => {
    if (!isAdminUser()) {
      toast("Login as admin first.");
      return;
    }
    const syncStatus = byId("syncStatus");
    const provider = byId("aiProviderSelect")?.value || "openrouter";
    if (syncStatus) syncStatus.textContent = "Refreshing AI need intelligence...";
    const result = await store.refreshNeedIntelligence(provider);
    await refreshAll();
    if (syncStatus) syncStatus.textContent = result.message || "AI need intelligence refreshed.";
    toast("AI signals refreshed.");
  }));

  byId("refreshUserDirectoryBtn")?.addEventListener("click", safeAsync(async () => {
    if (!isAdminUser()) {
      toast("Login as admin first.");
      return;
    }
    const result = await store.refreshUserDirectory();
    await rerender(false);
    toast(result.message || "User roles refreshed from GRE.");
  }));

  byId("syncGreChatbotBtn")?.addEventListener("click", safeAsync(async () => {
    if (!isAdminUser()) {
      toast("Login as admin first.");
      return;
    }
    const statusEl = byId("chatbotSyncStatus");
    const provider = byId("aiProviderSelect")?.value || "openrouter";
    if (statusEl) statusEl.textContent = "Fetching live trader and solution exports from GRE and updating the chatbot dataset...";
    const result = await store.syncGreChatbotData(provider);
    if (statusEl) {
      statusEl.textContent =
        `Chatbot dataset refreshed. Traders: ${result.summary?.traders || 0}, solutions: ${result.summary?.solutions || 0}, offerings: ${result.summary?.offerings || 0}, offering AI refreshed: ${result.summary?.offeringAiUpdated || 0}.`;
    }
    toast("GRE Chatbot dataset refreshed.");
  }));

  byId("downloadGreTraderBtn")?.addEventListener("click", safeAsync(async () => {
    if (!isAdminUser()) {
      toast("Login as admin first.");
      return;
    }
    const statusEl = byId("chatbotSyncStatus");
    if (statusEl) statusEl.textContent = "Downloading live GRE trader workbook...";
    const result = await store.downloadGreChatbotReport("trader");
    downloadBase64Workbook(result.download);
    if (statusEl) statusEl.textContent = `Downloaded ${result.download?.fileName || "trader workbook"}.`;
  }));

  byId("downloadGreSolutionBtn")?.addEventListener("click", safeAsync(async () => {
    if (!isAdminUser()) {
      toast("Login as admin first.");
      return;
    }
    const statusEl = byId("chatbotSyncStatus");
    if (statusEl) statusEl.textContent = "Downloading live GRE solution workbook...";
    const result = await store.downloadGreChatbotReport("solution");
    downloadBase64Workbook(result.download);
    if (statusEl) statusEl.textContent = `Downloaded ${result.download?.fileName || "solution workbook"}.`;
  }));

  byId("actionWorkbench")?.addEventListener("click", safeAsync(async (event) => {
    const button = event.target.closest("button");
    if (!button) return;

    if (button.id === "assignCuratorBtn") {
      if (!isAdminUser()) {
        toast("Login as admin to assign curators.");
        return;
      }
      await store.assignCurator(state.selectedNeedId, byId("assignCuratorSelect").value || null);
      await refreshAll();
      toast("Curator assignment updated.");
    }
  }));

  byId("needDetail")?.addEventListener("click", (event) => {
    const button = event.target.closest("#openWorkbenchBtn");
    if (!button) return;
    workbenchDialog?.showModal();
  });

  byId("actionWorkbench")?.addEventListener("submit", safeAsync(async (event) => {
    if (event.target.id === "directUpdateForm") {
      event.preventDefault();
      const need = getNeedById(state.selectedNeedId);
      const form = new FormData(event.target);
      await store.directCuratorUpdate({
        needId: need.id,
        proposedStatus: form.get("proposed_status"),
        proposedInternalStatus: form.get("proposed_internal_status"),
        proposedNextAction: form.get("proposed_next_action"),
        proposedCurationNotes: form.get("proposed_curation_notes"),
        proposedCurationCallDate: form.get("proposed_curation_call_date"),
        proposedDemandBroadcastNeeded:
          form.get("proposed_demand_broadcast_needed") === ""
            ? null
            : form.get("proposed_demand_broadcast_needed") === "true",
      });
      event.target.reset();
      await refreshAll();
      toast("Curation update saved and synced to GRE.");
      return;
    }
    if (event.target.id === "curatorMessageForm") {
      event.preventDefault();
      const form = new FormData(event.target);
      const result = await store.sendCuratorMessage(state.selectedNeedId, form.get("message"));
      event.target.reset();
      toast(result.message || "Message sent to curator.");
      return;
    }
    if (event.target.id === "manualProviderForm") {
      event.preventDefault();
      const form = new FormData(event.target);
      const result = await store.sendProviderIntro(state.selectedNeedId, form.get("provider_email"));
      event.target.reset();
      toast(result.message || "Provider outreach email triggered.");
    }
  }));

  byId("saveSubmissionReviewBtn")?.addEventListener("click", safeAsync(async () => {
    const status = byId("submissionReviewStatus");
    if (status) status.textContent = "Saving submission changes...";
    const patch = await buildSubmissionReviewPatch();
    const result = await store.updateFormSubmission(patch.submissionId, patch.update);
    await refreshAll();
    openSubmissionReviewDialog(patch.submissionId);
    if (status) status.textContent = result.message || "Submission changes saved.";
    toast(result.message || "Submission changes saved.");
  }));

  byId("approveSubmissionReviewBtn")?.addEventListener("click", safeAsync(async () => {
    const status = byId("submissionReviewStatus");
    if (status) status.textContent = "Saving changes and approving...";
    const patch = await buildSubmissionReviewPatch();
    await store.updateFormSubmission(patch.submissionId, patch.update);
    const result = await store.reviewFormSubmission(patch.submissionId, "approve", patch.update.adminReviewNotes || "");
    submissionReviewDialog?.close();
    await refreshAll();
    toast(result.message || (result.targetNeedId ? `Submission approved as Need ${result.targetNeedId}.` : "Form submission approved."));
  }));

  byId("rejectSubmissionReviewBtn")?.addEventListener("click", safeAsync(async () => {
    const status = byId("submissionReviewStatus");
    if (status) status.textContent = "Saving changes and rejecting...";
    const patch = await buildSubmissionReviewPatch();
    await store.updateFormSubmission(patch.submissionId, patch.update);
    await store.reviewFormSubmission(patch.submissionId, "reject", patch.update.adminReviewNotes || "");
    submissionReviewDialog?.close();
    await refreshAll();
    toast("Form submission rejected.");
  }));

  byId("adminView")?.addEventListener("click", safeAsync(async (event) => {
    const button = event.target.closest("[data-action]");
    if (!button) return;
    if (!isAdminUser()) {
      toast("Login as admin first.");
      return;
    }
    if (button.dataset.action === "approve-need") {
      await store.approveNeed(button.dataset.needId, "approve");
      await refreshAll();
      toast("Need approved.");
    }
    if (button.dataset.action === "reject-need") {
      await store.approveNeed(button.dataset.needId, "reject");
      await refreshAll();
      toast("Need rejected.");
    }
    if (button.dataset.action === "approve-update") {
      await store.reviewUpdateRequest(button.dataset.requestId, "approve");
      await refreshAll();
      toast("Curator update approved.");
    }
    if (button.dataset.action === "reject-update") {
      await store.reviewUpdateRequest(button.dataset.requestId, "reject");
      await refreshAll();
      toast("Curator update rejected.");
    }
    if (button.dataset.action === "edit-form-submission") {
      openSubmissionReviewDialog(button.dataset.submissionId);
      return;
    }
    if (button.dataset.action === "approve-form-submission") {
      const result = await store.reviewFormSubmission(button.dataset.submissionId, "approve");
      await refreshAll();
      toast(result.message || (result.targetNeedId ? `Submission approved as Need ${result.targetNeedId}.` : "Form submission approved."));
    }
    if (button.dataset.action === "reject-form-submission") {
      await store.reviewFormSubmission(button.dataset.submissionId, "reject");
      await refreshAll();
      toast("Form submission rejected.");
    }
    if (button.dataset.action === "run-puter-need-review") {
      const status = byId("puterStatus");
      if (status) status.textContent = `Running Puter review for Need ${button.dataset.needId}...`;
      await runPuterNeedReview(button.dataset.needId);
      if (status) status.textContent = `Puter recommendation ready for Need ${button.dataset.needId}.`;
      toast("Puter recommendation prepared.");
    }
    if (button.dataset.action === "accept-need-override") {
      const needId = button.dataset.needId;
      const field = button.dataset.overrideField;
      const patch = buildNeedOverridePatch(needId, field);
      const recommendation = getNeedRecommendation(needId);
      await store.applyNeedOverride(needId, patch, recommendation?.conflict_reason || "", field === "all");
      await refreshAll();
      toast(field === "all" ? "Recommendation accepted." : "Field override accepted.");
    }
    if (button.dataset.action === "set-user-role-tab") {
      state.adminUserRoleTab = button.dataset.roleTab || "admin";
      renderUserManagement();
      return;
    }
    if (button.dataset.action === "update-user-role") {
      const userId = button.dataset.userId;
      const card = button.closest(".approval-card");
      const select = card?.querySelector("[data-role-select]");
      const nextRole = select?.value || "";
      const result = await store.updateUserRole(userId, nextRole);
      await refreshAll();
      toast(result.message || "User role updated.");
    }
    if (button.dataset.action === "complete-user-role-activation") {
      const userId = button.dataset.userId;
      const card = button.closest(".approval-card");
      const otpField = card?.querySelector("[data-user-otp]");
      const otp = otpField?.value?.trim() || "";
      const result = await store.completeUserRoleActivation(userId, otp);
      await refreshAll();
      toast(result.message || "GRE activation completed.");
    }
    if (button.dataset.action === "remove-user") {
      const userId = button.dataset.userId;
      const card = button.closest(".approval-card");
      const name = card?.querySelector("h4")?.textContent?.trim() || "this user";
      const removalMode = chooseUserRemovalMode(name);
      if (!removalMode) return;
      const confirmMessage = removalMode === "full_account"
        ? `This will try to remove ${name} from MIS and completely remove the mapped GRE workforce account for Green Rural Economy. Continue?`
        : `This will remove ${name} from MIS and clear their GRE organisation access for Green Rural Economy. Continue?`;
      if (!window.confirm(confirmMessage)) return;
      const result = await store.removeManagedUser(userId, removalMode);
      await refreshAll();
      toast(result.message || "User removed.");
    }
    if (button.dataset.action === "edit-local-solution") {
      openLocalSolutionReviewDialog(button.dataset.offeringId);
      return;
    }
    if (button.dataset.action === "delete-local-solution") {
      const offeringId = button.dataset.offeringId;
      const card = button.closest(".approval-card");
      const name = card?.querySelector("h4")?.textContent?.trim() || "this solution";
      if (!window.confirm(`Delete ${name} from local Supabase solution records? This cannot be undone.`)) return;
      const result = await store.deleteLocalSolution(offeringId);
      await refreshAll();
      toast(result.message || "Local solution deleted.");
    }
  }));

  byId("matchResults")?.addEventListener("click", safeAsync(async (event) => {
    const button = event.target.closest("[data-action]");
    if (!button) return;
    const match = getCachedMatchByOfferingId(button.dataset.offeringId);

    if (button.dataset.action === "read-more-match") {
      openMatchDetailDialog(match);
      return;
    }
    if (button.dataset.action === "view-match-details") {
      openMatchDetailPage(match || { offering_id: button.dataset.offeringId });
      return;
    }
    if (button.dataset.action === "email-provider") {
      if (!(isAdminUser() || isCuratorUser())) {
        toast("Login as curator or admin to send provider outreach from the GRE mailbox.");
        return;
      }
      const result = await store.sendProviderIntro(state.selectedNeedId, button.dataset.providerEmail);
      toast(result.message || "Provider outreach email triggered.");
    }
  }));
}

async function init() {
  setText("userLoginStatus", "");
  setText("registerStatus", "");
  setText("resetStatus", "");
  setText("changePasswordStatus", "");
  bindStaticEvents();
  await store.validateUserSession();
  await refreshAll();
}

init().catch((error) => {
  console.error(error);
  toast(error.message || "Dashboard could not load.");
});
