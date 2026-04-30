const FALLBACK_CURATORS = [
  { id: "fallback-1", display_name: "Tanmay Mukherji", email: "tanmay@greenruraleconomy.in" },
  { id: "fallback-2", display_name: "Phaneesh K", email: "phaneesh@greenruraleconomy.in" },
  { id: "fallback-3", display_name: "Swati Singh", email: "swati@greenruraleconomy.in" },
  { id: "fallback-4", display_name: "Shaifali Nagar", email: "shaifali@greenruraleconomy.in" },
];

const state = {
  view: "overview",
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
  showClosedNeeds: false,
  overviewFilters: {
    metric: [],
    pipeline: [],
    curator: [],
    state: [],
    category: [],
  },
  adminToken: localStorage.getItem("gre-mis-admin-token") || "",
  adminSession: null,
  filters: {
    status: "all",
    curator: "all",
    state: "all",
    search: "",
  },
  data: {
    curators: [],
    options: [],
    needs: [],
    needUpdates: [],
    pendingNeeds: [],
    pendingUpdates: [],
  },
};

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
  "finance",
  "funding",
  "investments",
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

function safeAsync(handler) {
  return async (...args) => {
    try {
      await handler(...args);
    } catch (error) {
      console.error(error);
      const loginStatus = byId("loginStatus");
      const sessionStatus = byId("sessionStatus");
      if (loginStatus && args[0]?.target?.id === "adminLoginForm") {
        loginStatus.textContent = error?.message || "Admin login failed.";
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

function getNeedThemeSignals(need) {
  const themes = [];
  parseArray(need?.curated_need).forEach((item) => {
    const parts = extractCategoryParts(item);
    if (parts.thematic) themes.push(parts.thematic);
  });
  return uniq(themes);
}

function getNeedSixMSignals(need) {
  const explicit = parseArray(need?.six_m_signals);
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
  const aiKeywords = uniq(parseArray(need.ai_keywords).map((item) => item.toLowerCase()));
  const aiSignals = uniq([
    normalizeText(need.ai_thematic_area).toLowerCase(),
    normalizeText(need.ai_application_area).toLowerCase(),
    normalizeText(need.ai_need_kind).toLowerCase(),
    normalizeText(need.ai_service_kind).toLowerCase(),
    ...aiKeywords,
  ].filter(Boolean));
  const aiNeedKind = normalizeText(need.ai_need_kind).toLowerCase();
  const aiServiceKind = normalizeText(need.ai_service_kind).toLowerCase();
  const categoryThematicAreas = uniq(
    categoryParts
      .map((item) => item.thematic)
      .filter((item) => item && !GENERIC_THEMATIC_TERMS.has(item)),
  );
  const serviceTerms = uniq(categoryParts.map((item) => item.service).filter(Boolean));
  const problemTokens = tokenizeText(need.problem_statement, 5);
  const notesTokens = tokenizeText(need.curation_notes, 5);
  const geographyTokens = tokenizeText(`${need.state || ""} ${need.district || ""}`, 3);
  const serviceTokens = serviceTerms.flatMap((item) => tokenizeText(item, 3));
  const sharedSolutionHints = extractSharedSolutionHints(need.curation_notes);
  const problemPhrases = extractProblemPhrases(need.problem_statement);
  const thematicAreas = uniq([
    ...sharedSolutionHints.phrases,
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
  const requiresServiceMatch = Boolean(aiServiceKind) || (aiNeedKind === "service" && categoryThematicAreas.length === 0);

  return {
    categories,
    categoryParts,
    categoryThematicAreas,
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
    sharedSolutionHints,
  };
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

    mapCanvas.innerHTML = `<div id="categoryCasesMapCanvas" class="case-map-canvas"></div><div id="caseMapLocationPanel" class="case-map-location-panel hidden"></div>`;
    state.caseMap = new window.mappls.Map("categoryCasesMapCanvas", {
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
    if (activeCurators.length && !activeCurators.includes(need?.curator_id || "")) return false;
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
  activeCurators.forEach((curatorId) => labels.push(`Curator: ${getCuratorById(curatorId)?.display_name || curatorId}`));
  activeStates.forEach((stateName) => labels.push(`State: ${stateName}`));
  activeCategories.forEach((category) => labels.push(`Category: ${category}`));

  return {
    label: activeFilterCount ? labels.join(" | ") : "All Cases",
    tone: activeMetrics.includes("stuck") || activePipelines.includes("stuck") ? "bad" : activeMetrics.includes("admin_queue") ? "warn" : "info",
    items: orderedCases,
    cards,
    emptyText: activeMetrics.includes("admin_queue") && !state.adminToken
      ? "Sign in through Admin Sync to inspect pending approvals and curator requests."
      : "No cases match the current combination of filters.",
  };
}

function scoreOfferingMatch(need, profile, offering) {
  const tags = parseArray(offering.tags).map((item) => item.toLowerCase());
  const geographies = parseArray(offering.geographies).map((item) => item.toLowerCase());
  const category = normalizeText(offering.offering_category).toLowerCase();
  const name = normalizeText(offering.offering_name).toLowerCase();
  const about = normalizeText(offering.about_offering_text || offering.solution?.about_solution_text).toLowerCase();
  const solutionName = normalizeText(offering.solution?.solution_name).toLowerCase();
  const joined = [name, category, about, solutionName, tags.join(" "), geographies.join(" ")].join(" ");

  let score = 0;
  const reasons = [];
  let thematicMatched = false;
  let serviceMatched = false;
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

  profile.thematicAreas.forEach((phrase) => {
    if (joined.includes(phrase)) {
      thematicMatched = true;
      score += 18;
      reasons.push(phrase);
    }
  });

  profile.categoryTokens.forEach((token) => {
    if (tags.some((tag) => tag.includes(token))) {
      thematicMatched = true;
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

  profile.serviceTokens.forEach((token) => {
    if (joined.includes(token)) {
      serviceMatched = true;
      score += 4;
      reasons.push(token);
    }
  });

  profile.problemTokens.slice(0, 8).forEach((token) => {
    if ((thematicMatched || !profile.categoryThematicAreas.length) && (tags.some((tag) => tag.includes(token)) || name.includes(token) || solutionName.includes(token))) {
      if (!profile.categoryThematicAreas.length) thematicMatched = true;
      score += 4;
      reasons.push(token);
    } else if ((thematicMatched || !profile.categoryThematicAreas.length) && about.includes(token)) {
      if (!profile.categoryThematicAreas.length) thematicMatched = true;
      score += 2;
      reasons.push(token);
    }
  });

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
    serviceMatched,
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
      state.data.options = [];
      state.data.needs = [];
      state.data.needUpdates = [];
      state.matchCache.clear();
      return;
    }

    const [curators, options, needs, updates] = await Promise.all([
      client.from("gre_mis_curators").select("id, display_name, email").eq("is_active", true).order("display_name"),
      client.from("gre_mis_options").select("id, option_type, label, sort_order").eq("is_active", true).order("sort_order"),
      client.from("gre_mis_needs").select("*").eq("approval_status", "approved").order("requested_on", { ascending: false }),
      client.from("gre_mis_need_updates").select("*").order("created_at", { ascending: false }),
    ]);

    if (curators.error || options.error || needs.error || updates.error) {
      throw new Error(curators.error?.message || options.error?.message || needs.error?.message || updates.error?.message || "Could not load live dashboard data.");
    }

    state.data.curators = ensureList(curators.data);
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
    if (state.adminToken) headers["x-gre-admin-session"] = state.adminToken;
    const response = await fetch(`${this.config.SUPABASE_URL}/functions/v1/${this.config.ADMIN_FUNCTION || "gre-mis-admin"}`, {
      method: "POST",
      headers,
      body: JSON.stringify({
        action,
        ...body,
        adminSessionToken: state.adminToken || undefined,
      }),
    });
    const rawText = await response.text().catch(() => "");
    let data = {};
    try {
      data = rawText ? JSON.parse(rawText) : {};
    } catch {}
    if (!response.ok) throw new Error(data.error || rawText || `Request failed (${response.status}).`);
    if (requireAdmin && !state.adminToken) throw new Error("Admin login required.");
    return data;
  }

  async adminLogin(username, password) {
    const data = await this.callAdmin("adminLogin", { username, password });
    state.adminToken = data.token;
    state.adminSession = { username: data.username, email: "tanmay@greenruraleconomy.in" };
    localStorage.setItem("gre-mis-admin-token", data.token);
    return data;
  }

  async validateAdminSession() {
    if (!state.adminToken) return false;
    try {
      const data = await this.callAdmin("validateAdminSession");
      state.adminSession = { username: data.username, email: data.email };
      return true;
    } catch {
      state.adminToken = "";
      state.adminSession = null;
      localStorage.removeItem("gre-mis-admin-token");
      return false;
    }
  }

  async adminLogout() {
    if (state.adminToken) {
      await this.callAdmin("adminLogout");
    }
    state.adminToken = "";
    state.adminSession = null;
    localStorage.removeItem("gre-mis-admin-token");
  }

  async loadAdminSnapshot() {
    if (!state.adminToken) {
      state.data.pendingNeeds = [];
      state.data.pendingUpdates = [];
      return;
    }
    const data = await this.callAdmin("adminSnapshot", {}, true);
    state.data.pendingNeeds = ensureList(data.pendingNeeds);
    state.data.pendingUpdates = ensureList(data.pendingUpdates);
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
    return this.callAdmin("sendProviderIntro", { needId, providerEmail }, true);
  }

  async importInboundWorkbook(fileName, rows, aiProvider) {
    return this.callAdmin("importInboundWorkbook", { fileName, rows, aiProvider }, true);
  }

  async refreshNeedIntelligence(aiProvider) {
    return this.callAdmin("refreshNeedIntelligence", { aiProvider }, true);
  }

  async searchMatchesForNeed(need) {
    const client = this.getClient();
    if (!client || !need) return [];
    const profile = buildNeedMatchProfile(need);
    const offeringSelect = "offering_id,solution_id,trader_id,offering_name,offering_category,tags,geographies,about_offering_text,contact_details,gre_link";
    const searchTerms = uniq([...profile.thematicAreas, ...profile.serviceTerms, ...profile.phrases, ...profile.primaryTerms]).slice(0, 8);
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
          .limit(48),
      );
    }

    searchTerms.slice(0, 3).forEach((term) => {
      const safeTerm = term.replaceAll("%", "").replaceAll(",", " ").trim();
      if (!safeTerm) return;
      queries.push(
        client
          .from("offerings")
          .select(offeringSelect)
          .or(`offering_name.ilike.%${safeTerm}%,offering_category.ilike.%${safeTerm}%,about_offering_text.ilike.%${safeTerm}%`)
          .limit(30),
      );
    });

    if (!queries.length) {
      queries.push(client.from("offerings").select(offeringSelect).limit(48));
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

    return offerings
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
          serviceMatched: matchMeta.serviceMatched,
          matchReasons: matchMeta.reasons,
        };
      })
      .filter((item) => item.thematicMatched && (profile.requiresServiceMatch ? item.serviceMatched : true) && item.matchScore >= 12)
      .sort((a, b) => b.matchScore - a.matchScore);
  }
}

const store = new GreMisStore();

function getCuratorById(id) {
  return state.data.curators.find((curator) => curator.id === id) || null;
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
    .filter((need) => state.filters.curator === "all" || (need.curator_id || "unassigned") === state.filters.curator)
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
  if (headline) headline.textContent = `${needs.length} active inbound needs loaded from GRE operations data`;
  if (subline) subline.textContent = `${state.data.pendingNeeds.length} intake records and ${state.data.pendingUpdates.length} curator updates are waiting for admin action.${state.showClosedNeeds ? " Closed needs are currently visible." : " Closed needs are currently hidden."}`;
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
      <div class="pipeline-drilldown-list">
        ${focusPayload.cards.length
          ? visibleCards.join("")
          : `<div class="empty-state">${esc(focusPayload.emptyText || "No cases are currently sitting in this selection.")}</div>`}
      </div>
      <div id="categoryCasesMap" class="pipeline-drilldown-map">
        <div id="caseMapLocationPanel" class="case-map-location-panel hidden"></div>
      </div>
    </div>
    ${
      focusPayload.cards.length
        ? `<div class="pipeline-pagination">
            <span class="meta-text">Showing ${esc(pageStart + 1)}-${esc(Math.min(pageStart + pageSize, focusPayload.cards.length))} of ${esc(focusPayload.cards.length)}</span>
            <div class="pipeline-pagination-actions">
              <button class="btn btn-secondary" data-page-action="prev" ${state.overviewPage <= 1 ? "disabled" : ""}>Prev</button>
              <span class="meta-text">Page ${esc(state.overviewPage)} of ${esc(totalPages)}</span>
              <button class="btn btn-secondary" data-page-action="next" ${state.overviewPage >= totalPages ? "disabled" : ""}>Next</button>
            </div>
          </div>`
        : ""
    }
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
          <p class="helper-text">${esc(need.problem_statement.slice(0, 155))}${need.problem_statement.length > 155 ? "..." : ""}</p>
          <p class="meta-text">${esc(need.state)}${need.district ? ` / ${esc(need.district)}` : ""} • Curator: ${esc(curator?.display_name || "Unassigned")} • Age: ${esc(need.curation_age_days || 0)} days</p>
        </article>
      `;
    })
    .join("");

  if (state.queueNeedsScrollIntoView) {
    const activeCard = queue.querySelector(`[data-need-id="${CSS.escape(String(state.selectedNeedId || ""))}"]`);
    if (activeCard) {
      activeCard.scrollIntoView({ block: "center", behavior: "smooth" });
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
  const noteText = stripUrls(need.curation_notes);
  const summaryBadges = [
    { label: need.status, tone: badgeTone(need.status) },
    { label: need.internal_status, tone: badgeTone(need.internal_status) },
    { label: `${need.curation_age_days || 0} days old`, tone: Number(need.curation_age_days || 0) >= 7 ? "bad" : "info" },
    { label: need.approval_status, tone: badgeTone(need.approval_status) },
  ];
  detailEl.innerHTML = `
    <div class="need-detail-stack">
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

      <article class="detail-card detail-stack-card">
        <h4>Curation Snapshot</h4>
        <div class="kv-grid">
          <div><span>Assigned Curator</span><strong>${esc(curator?.display_name || "Unassigned")}</strong></div>
          <div><span>Next Action</span><strong>${esc(need.next_action || "Not set")}</strong></div>
          <div><span>Curation Call</span><strong>${esc(need.curation_call_date || "Not set")}</strong></div>
          <div><span>Broadcast Needed</span><strong>${need.demand_broadcast_needed ? "Yes" : "No"}</strong></div>
          <div><span>Solutions Shared</span><strong>${esc(need.solutions_shared_count || 0)}</strong></div>
          <div><span>Invited Providers</span><strong>${esc(need.invited_providers_count || 0)}</strong></div>
        </div>
      </article>

      <article class="detail-card detail-stack-card">
        <h4>Categories</h4>
        <div class="tag-row">${parseArray(need.curated_need).map((item) => `<span>${esc(item)}</span>`).join("") || `<span>Unclassified</span>`}</div>
      </article>

      <article class="detail-card detail-stack-card">
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
      </article>

    </div>
  `;
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
          return `
            <article class="match-card">
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
              <p>${esc((match.about_offering_text || match.solution?.about_solution_text || "").slice(0, 320))}${(match.about_offering_text || match.solution?.about_solution_text || "").length > 320 ? "..." : ""}</p>
              <p class="meta-text">${esc(parseArray(match.geographies).slice(0, 3).join(", ") || "Geography not listed")}</p>
              <div class="card-actions">
                ${match.gre_link ? `<a class="btn btn-secondary" href="${esc(match.gre_link)}" target="_blank" rel="noreferrer">Open GRE Link</a>` : `<span></span>`}
                ${state.adminToken && email ? `<button class="btn btn-primary" data-action="email-provider" data-provider-email="${esc(email)}">Email This Provider</button>` : ""}
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
  const curatorOptions = [`<option value="">Unassigned</option>`, ...state.data.curators.map((item) => `<option value="${esc(item.id)}" ${item.id === need.curator_id ? "selected" : ""}>${esc(item.display_name)}</option>`)].join("");

  workbench.innerHTML = `
    <article class="action-card">
      <p class="eyebrow">Admin Assignment</p>
      <h4>Curator Allocation</h4>
      <label>
        <span>Assigned Curator</span>
        <select id="assignCuratorSelect">${curatorOptions}</select>
      </label>
      <button class="btn btn-secondary" id="assignCuratorBtn" ${state.adminToken ? "" : "disabled"}>Save Curator Assignment</button>
      <p class="helper-text">${state.adminToken ? "Admin can rebalance or assign needs directly." : "Login as admin to change curator allocation."}</p>
    </article>
    <article class="action-card">
      <p class="eyebrow">Curator Update Request</p>
      <h4>Submit Status Change for Approval</h4>
      <form id="updateRequestForm" class="action-form">
        <label>
          <span>Curator</span>
          <input name="curator_name" value="${esc(curator?.display_name || "")}" ${curator ? "readonly" : "readonly"} />
        </label>
        <label>
          <span>Curator Email</span>
          <input name="curator_email" value="${esc(curator?.email || "")}" ${curator ? "readonly" : "readonly"} />
        </label>
        <label>
          <span>Proposed Status</span>
          <select name="proposed_status">
            <option value="">No change</option>
            ${statusOptions.map((item) => `<option value="${esc(item.label)}">${esc(item.label)}</option>`).join("")}
          </select>
        </label>
        <label>
          <span>Proposed Internal Status</span>
          <select name="proposed_internal_status">
            <option value="">No change</option>
            ${internalOptions.map((item) => `<option value="${esc(item.label)}">${esc(item.label)}</option>`).join("")}
          </select>
        </label>
        <label>
          <span>Proposed Next Action</span>
          <select name="proposed_next_action">
            <option value="">No change</option>
            ${nextActionOptions.map((item) => `<option value="${esc(item.label)}">${esc(item.label)}</option>`).join("")}
          </select>
        </label>
        <label>
          <span>Curation Call Date</span>
          <input name="proposed_curation_call_date" type="date" />
        </label>
        <label class="wide">
          <span>Curator Notes</span>
          <textarea name="proposed_curation_notes" rows="4" placeholder="Describe what changed and why."></textarea>
        </label>
        <label>
          <span>Broadcast Needed</span>
          <select name="proposed_demand_broadcast_needed">
            <option value="">No change</option>
            <option value="true">Yes</option>
            <option value="false">No</option>
          </select>
        </label>
        <label>
          <span>Solutions Shared Count</span>
          <input name="proposed_solutions_shared_count" type="number" min="0" />
        </label>
        <label>
          <span>Invited Providers Count</span>
          <input name="proposed_invited_providers_count" type="number" min="0" />
        </label>
        <div class="wide">
          <button class="btn btn-primary" type="submit" ${curator ? "" : "disabled"}>Submit Update for Admin Approval</button>
        </div>
      </form>
      <p class="helper-text">${curator ? "Only the currently assigned curator can submit an update request for this need." : "Assign a curator first so status changes can flow through approval."}</p>
    </article>
  `;
}

function renderAdminQueue() {
  const pendingNeedsList = byId("pendingNeedsList");
  const pendingUpdatesList = byId("pendingUpdatesList");
  const optionsList = byId("optionsList");
  if (!pendingNeedsList || !pendingUpdatesList || !optionsList) return;

  pendingNeedsList.innerHTML = state.adminToken
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

  pendingUpdatesList.innerHTML = state.adminToken
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

  const grouped = ensureList(state.data.options).reduce((acc, option) => {
    acc[option.option_type] ||= [];
    acc[option.option_type].push(option);
    return acc;
  }, {});

  optionsList.innerHTML = Object.entries(grouped)
    .map(
      ([type, items]) => `
        <article class="stack-card">
          <h4>${esc(type.replaceAll("_", " "))}</h4>
          <div class="tag-row">${items.map((item) => `<span>${esc(item.label)}</span>`).join("")}</div>
        </article>
      `,
    )
    .join("");
}

function renderAdminState() {
  const chip = byId("adminStatusChip");
  const text = byId("adminStatusText");
  const logout = byId("adminLogoutBtn");
  const loginStatus = byId("loginStatus");
  const sessionPanel = byId("sessionPanel");
  if (!chip || !text || !logout) return;

  if (state.adminSession) {
    chip.textContent = "Unlocked";
    chip.className = "chip good";
    text.textContent = `Admin session active for ${state.adminSession.username}. Intake approvals and curator update approvals are available below.`;
    logout.classList.remove("hidden");
    sessionPanel?.classList.remove("hidden");
  } else {
    chip.textContent = "Locked";
    chip.className = "chip muted";
    text.textContent = "Admin approval is required for new intake records and curator-submitted status updates.";
    logout.classList.add("hidden");
    sessionPanel?.classList.add("hidden");
  }

  if (loginStatus && !state.adminSession && !loginStatus.textContent) {
    loginStatus.textContent = "";
  }
}

function switchView(nextView) {
  state.view = nextView;
  document.querySelectorAll(".tab").forEach((button) => button.classList.toggle("active", button.dataset.view === nextView));
  document.querySelectorAll(".view").forEach((section) => section.classList.toggle("active", section.id === `${nextView}View`));
  if (nextView === "operations") {
    renderMatches();
  }
}

async function rerender(options = {}) {
  const { includeMatches = state.view === "operations" } = options;
  renderMetrics();
  renderOverview();
  renderFilters();
  renderQueue();
  renderNeedDetail();
  renderWorkbench();
  renderAdminQueue();
  renderAdminState();
  if (!byId("overviewView") && byId("adminView")) {
    const headline = byId("datasetHeadline");
    const subline = byId("datasetSubline");
    if (headline) headline.textContent = "Admin sync workspace for approval and taxonomy maintenance";
    if (subline) subline.textContent = `${state.data.pendingNeeds.length} intake records and ${state.data.pendingUpdates.length} curator updates are waiting for review.`;
  }
  if (includeMatches) {
    await renderMatches();
  }
}

async function refreshAll() {
  await Promise.all([
    store.loadBaseData(),
    state.adminToken
      ? store.loadAdminSnapshot().catch(() => {
          state.data.pendingNeeds = [];
          state.data.pendingUpdates = [];
        })
      : Promise.resolve().then(() => {
          state.data.pendingNeeds = [];
          state.data.pendingUpdates = [];
        }),
  ]);
  const displayNeeds = getDisplayNeeds();
  if (!displayNeeds.find((need) => need.id === state.selectedNeedId)) {
    state.selectedNeedId = displayNeeds[0]?.id || null;
  }
  await rerender();
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
  });
  byId("curatorFilter")?.addEventListener("change", (event) => {
    state.filters.curator = event.target.value;
    renderQueue();
  });
  byId("stateFilter")?.addEventListener("change", (event) => {
    state.filters.state = event.target.value;
    renderQueue();
  });
  byId("searchFilter")?.addEventListener("input", (event) => {
    state.filters.search = event.target.value.trim();
    renderQueue();
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
    if (byId("adminView") && state.adminToken) {
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

  const dialog = byId("needDialog");
  byId("newNeedBtn")?.addEventListener("click", () => dialog?.showModal());
  byId("closeNeedDialog")?.addEventListener("click", () => dialog?.close());

  byId("needForm")?.addEventListener("submit", safeAsync(async (event) => {
    event.preventDefault();
    const form = new FormData(event.target);
    await store.createNeed(Object.fromEntries(form.entries()));
    event.target.reset();
    dialog?.close();
    await refreshAll();
    toast("Need submitted to the admin approval queue.");
  }));

  byId("adminLoginForm")?.addEventListener("submit", safeAsync(async (event) => {
    event.preventDefault();
    const form = new FormData(event.target);
    const loginStatus = byId("loginStatus");
    if (loginStatus) loginStatus.textContent = "Signing in...";
    await store.adminLogin(form.get("username"), form.get("password"));
    event.target.reset();
    await refreshAll();
    if (byId("adminView")) switchView("admin");
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
    if (!state.adminToken) {
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
    if (!state.adminToken) {
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

  byId("refreshAiBtn")?.addEventListener("click", safeAsync(async () => {
    if (!state.adminToken) {
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

  byId("actionWorkbench")?.addEventListener("click", safeAsync(async (event) => {
    const button = event.target.closest("button");
    if (!button) return;

    if (button.id === "assignCuratorBtn") {
      if (!state.adminToken) {
        toast("Login as admin to assign curators.");
        return;
      }
      await store.assignCurator(state.selectedNeedId, byId("assignCuratorSelect").value || null);
      await refreshAll();
      toast("Curator assignment updated.");
    }
  }));

  byId("actionWorkbench")?.addEventListener("submit", safeAsync(async (event) => {
    if (event.target.id !== "updateRequestForm") return;
    event.preventDefault();
    const need = getNeedById(state.selectedNeedId);
    const form = new FormData(event.target);
    await store.submitUpdateRequest({
      needId: need.id,
      curatorName: form.get("curator_name"),
      curatorEmail: form.get("curator_email"),
      proposedStatus: form.get("proposed_status"),
      proposedInternalStatus: form.get("proposed_internal_status"),
      proposedNextAction: form.get("proposed_next_action"),
      proposedCurationNotes: form.get("proposed_curation_notes"),
      proposedCurationCallDate: form.get("proposed_curation_call_date"),
      proposedDemandBroadcastNeeded:
        form.get("proposed_demand_broadcast_needed") === ""
          ? null
          : form.get("proposed_demand_broadcast_needed") === "true",
      proposedSolutionsSharedCount:
        form.get("proposed_solutions_shared_count") === "" ? null : Number(form.get("proposed_solutions_shared_count")),
      proposedInvitedProvidersCount:
        form.get("proposed_invited_providers_count") === "" ? null : Number(form.get("proposed_invited_providers_count")),
    });
    event.target.reset();
    await refreshAll();
    toast("Curator update submitted for admin approval.");
  }));

  byId("adminView")?.addEventListener("click", safeAsync(async (event) => {
    const button = event.target.closest("[data-action]");
    if (!button) return;
    if (!state.adminToken) {
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
  }));

  byId("matchResults")?.addEventListener("click", safeAsync(async (event) => {
    const button = event.target.closest("[data-action='email-provider']");
    if (!button) return;
    if (!state.adminToken) {
      toast("Login as admin to send provider outreach from the GRE mailbox.");
      return;
    }
    const result = await store.sendProviderIntro(state.selectedNeedId, button.dataset.providerEmail);
    toast(result.message || "Provider outreach email triggered.");
  }));
}

async function init() {
  bindStaticEvents();
  await store.validateAdminSession();
  await refreshAll();
}

init().catch((error) => {
  console.error(error);
  toast(error.message || "Dashboard could not load.");
});
