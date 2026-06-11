import { createClient } from "npm:@supabase/supabase-js@2";
import * as XLSX from "npm:xlsx@0.18.5";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type, x-gre-admin-session, x-gre-user-session",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

const supabaseUrl = Deno.env.get("SUPABASE_URL") ?? "";
const serviceRoleKey =
  Deno.env.get("GRE_MIS_SERVICE_ROLE_KEY") ??
  Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ??
  Deno.env.get("SELCO_VENDOR_SERVICE_ROLE_KEY") ??
  "";
const gmailClientId = Deno.env.get("GMAIL_CLIENT_ID") ?? "";
const gmailClientSecret = Deno.env.get("GMAIL_CLIENT_SECRET") ?? "";
const gmailRefreshToken = Deno.env.get("GMAIL_REFRESH_TOKEN") ?? "";
const gmailSenderEmail = Deno.env.get("GMAIL_SENDER_EMAIL") ?? "help@greenruraleconomy.in";
const helpGmailClientId = Deno.env.get("HELP_GMAIL_CLIENT_ID") ?? gmailClientId;
const helpGmailClientSecret = Deno.env.get("HELP_GMAIL_CLIENT_SECRET") ?? gmailClientSecret;
const helpGmailRefreshToken = Deno.env.get("HELP_GMAIL_REFRESH_TOKEN") ?? gmailRefreshToken;
const helpGmailSenderEmail = Deno.env.get("HELP_GMAIL_SENDER_EMAIL") ?? "help@greenruraleconomy.in";
const solutionGmailClientId = Deno.env.get("SOLUTION_GMAIL_CLIENT_ID") ?? gmailClientId;
const solutionGmailClientSecret = Deno.env.get("SOLUTION_GMAIL_CLIENT_SECRET") ?? gmailClientSecret;
const solutionGmailRefreshToken = Deno.env.get("SOLUTION_GMAIL_REFRESH_TOKEN") ?? "";
const solutionGmailSenderEmail = Deno.env.get("SOLUTION_GMAIL_SENDER_EMAIL") ?? "solution@greenruraleconomy.in";
const openRouterApiKey = Deno.env.get("OPENROUTER_API_KEY") ?? Deno.env.get("GRE_MIS_OPENROUTER_API_KEY") ?? "";
const geminiApiKey = Deno.env.get("GEMINI_API_KEY") ?? Deno.env.get("GRE_MIS_GEMINI_API_KEY") ?? "";
const deepSeekApiKey = Deno.env.get("DEEPSEEK_API_KEY") ?? Deno.env.get("GRE_MIS_DEEPSEEK_API_KEY") ?? "";
const openAiApiKey = Deno.env.get("OPENAI_API_KEY") ?? Deno.env.get("GRE_MIS_OPENAI_API_KEY") ?? "";
const defaultAiProvider = (Deno.env.get("GRE_MIS_AI_PROVIDER") ?? "openrouter").toLowerCase();
const defaultOpenRouterModel = Deno.env.get("GRE_MIS_OPENROUTER_MODEL") ?? "openai/gpt-4.1-mini";
const defaultGeminiModel = Deno.env.get("GRE_MIS_GEMINI_MODEL") ?? "gemini-2.0-flash";
const defaultDeepSeekModel = Deno.env.get("GRE_MIS_DEEPSEEK_MODEL") ?? "deepseek-chat";
const defaultOpenAiModel = Deno.env.get("GRE_MIS_OPENAI_MODEL") ?? "gpt-4.1-mini";
const aiPromptVersion = "2026-05-01.gre-mis.v1";
const aiSchemaVersion = "gre-mis-need-intelligence.v1";
const mapplsAccessToken = Deno.env.get("MAPPLS_ACCESS_TOKEN") ?? Deno.env.get("GRE_MIS_MAPPLS_ACCESS_TOKEN") ?? "";
const greLoginBaseUrl = Deno.env.get("GRE_LOGIN_BASE_URL") ?? "https://login.platformcommons.org/gateway";
const greCtldBaseUrl = Deno.env.get("GRE_CTLD_BASE_URL") ?? "https://login.platformcommons.org/ctld";
const grePtldBaseUrl = Deno.env.get("GRE_PTLD_BASE_URL") ?? "https://login.platformcommons.org/ptld";
const greSiteOrigin = Deno.env.get("GRE_SITE_ORIGIN") ?? "https://greenruraleconomy.in";
const greMasterLogin = Deno.env.get("GRE_MASTER_LOGIN") ?? "";
const greMasterTenant = Deno.env.get("GRE_MASTER_TENANT") ?? "green_rural_economy";
const greMasterPassword = Deno.env.get("GRE_MASTER_PASSWORD") ?? "";
const greTemporaryUserPassword = Deno.env.get("GRE_TEMP_USER_PASSWORD") ?? "gre@1234";
const greMarketId = Deno.env.get("GRE_MARKET_ID") ?? "5";
const greChannelId = Deno.env.get("GRE_CHANNEL_ID") ?? "4";
const greInboundReportName = Deno.env.get("GRE_INBOUND_REPORT_NAME") ?? "MARKIFY_GRE.REQUEST_DETAILS_REPORT";
const greTraderReportName = Deno.env.get("GRE_TRADER_REPORT_NAME") ?? "MARKIFY_REPORT.MARKET_TRADER_BASIC_DETAILS";
const greSolutionReportName =
  Deno.env.get("GRE_SOLUTION_REPORT_NAME") ?? "MARKIFY_REPORT.GRE_SOLUTION_WITH_OFFERINGS_DETAILS";
const githubAssetToken =
  Deno.env.get("GRE_MIS_GITHUB_TOKEN") ??
  Deno.env.get("GITHUB_TOKEN") ??
  Deno.env.get("GITHUB_ACTIONS_TOKEN") ??
  Deno.env.get("GRE_GITHUB_TOKEN") ??
  "";
const githubAssetRepo = Deno.env.get("GRE_MIS_GITHUB_ASSET_REPO") ?? "tanmaymukherji/gre-chatbot";
const githubAssetBranch = Deno.env.get("GRE_MIS_GITHUB_ASSET_BRANCH") ?? "main";
const githubAssetRoot = Deno.env.get("GRE_MIS_GITHUB_ASSET_ROOT") ?? "public/uploads/local-offerings";
const askGreBaseUrl = Deno.env.get("ASKGRE_BASE_URL") ?? "https://askgre.grameee.org";
const greMisBaseUrl = Deno.env.get("GRE_MIS_BASE_URL") ?? "https://gre.grameee.org";
const libreTranslateApiUrl = Deno.env.get("LIBRETRANSLATE_API_URL") ?? "https://libretranslate.com/translate";
const libreTranslateApiKey = Deno.env.get("LIBRETRANSLATE_API_KEY") ?? "";

const adminClient = createClient(supabaseUrl, serviceRoleKey, {
  auth: { persistSession: false },
});

function jsonResponse(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      ...corsHeaders,
      "Content-Type": "application/json",
    },
  });
}

function requireString(value: unknown) {
  return typeof value === "string" ? value.trim() : "";
}

function normalizeEmailList(value: unknown): string[] {
  const list = Array.isArray(value) ? value : String(value || "").split(/[;,]/);
  return list
    .map((item: unknown) => String(item ?? "").trim().toLowerCase())
    .filter((item: string) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(item));
}

function formatLgdGeographyLabel(row: Record<string, unknown>) {
  const country = "India";
  const block = requireString(row.block_name || row.gram_panchayat_name || row.village_name);
  const city = requireString(row.district_name);
  const state = requireString(row.state_name);
  const kind = requireString(row.location_kind).toLowerCase();
  if (block && city && state) return [block, city, state, country].filter(Boolean).join(", ");
  if (city && state) return [city, state, country].filter(Boolean).join(", ");
  if (state) return [state, country].filter(Boolean).join(", ");
  if (kind === "country") return country;
  return requireString(row.display_label) || country;
}

function buildLgdGeographyVariants(row: Record<string, unknown>) {
  const country = "India";
  const block = requireString(row.block_name || row.gram_panchayat_name || row.village_name);
  const city = requireString(row.district_name);
  const state = requireString(row.state_name);
  const display = requireString(row.display_label);
  const variants = new Set<string>();
  if (block && city && state) variants.add([block, city, state, country].filter(Boolean).join(", "));
  if (city && state) variants.add([city, state, country].filter(Boolean).join(", "));
  if (state) variants.add([state, country].filter(Boolean).join(", "));
  variants.add(country);
  if (display) variants.add(display);
  return [...variants].filter(Boolean);
}

function scoreLgdSuggestion(label: string, query: string) {
  const normalizedLabel = requireString(label).toLowerCase();
  const normalizedQuery = requireString(query).toLowerCase();
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

function buildRankedLgdSuggestions(rows: Record<string, unknown>[], query: string) {
  const seen = new Set<string>();
  return rows
    .flatMap((row) => buildLgdGeographyVariants(row))
    .filter((label) => {
      const key = requireString(label).toLowerCase();
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .map((label) => ({ label, score: scoreLgdSuggestion(label, query) }))
    .filter((item) => item.score > 0)
    .sort((left, right) => right.score - left.score || left.label.localeCompare(right.label))
    .slice(0, 20)
    .map((item) => item.label);
}

function decodeJwtPayload(token: string) {
  const normalized = requireString(token);
  if (!normalized) return null;
  const parts = normalized.split(".");
  if (parts.length < 2) return null;
  try {
    const base64 = parts[1].replace(/-/g, "+").replace(/_/g, "/");
    const padded = base64.padEnd(Math.ceil(base64.length / 4) * 4, "=");
    const decoded = atob(padded);
    return JSON.parse(decoded) as Record<string, unknown>;
  } catch {
    return null;
  }
}

async function resolveGrameeeAuthUser(grameeeAccessToken: string) {
  const normalizedToken = requireString(grameeeAccessToken);
  if (!normalizedToken) return null;

  const { data: authData, error: authError } = await adminClient.auth.getUser(normalizedToken);
  if (!authError && authData?.user) {
    return authData.user;
  }

  const jwtPayload = decodeJwtPayload(normalizedToken);
  if (!jwtPayload) return null;

  const appMetadata = (jwtPayload.app_metadata || {}) as Record<string, unknown>;
  const userMetadata = (jwtPayload.user_metadata || jwtPayload.raw_user_meta_data || {}) as Record<string, unknown>;

  return {
    email: requireString(jwtPayload.email),
    app_metadata: appMetadata,
    user_metadata: userMetadata,
    raw_user_meta_data: userMetadata,
  } as unknown as { email?: string | null; app_metadata?: Record<string, unknown>; user_metadata?: Record<string, unknown>; raw_user_meta_data?: Record<string, unknown> };
}

function resolveGrameeeSummaryFallback(summary: unknown) {
  const record = summary && typeof summary === "object" ? summary as Record<string, unknown> : null;
  if (!record) return null;
  const email = requireString(record.email).toLowerCase();
  if (!email) return null;
  return {
    email,
    user_metadata: {
      full_name: requireString(record.fullName || record.full_name),
      first_name: requireString(record.firstName || record.first_name),
      username: requireString(record.username),
      phone: requireString(record.phone),
    },
    raw_user_meta_data: {
      full_name: requireString(record.fullName || record.full_name),
      first_name: requireString(record.firstName || record.first_name),
      username: requireString(record.username),
      phone: requireString(record.phone),
    },
    app_metadata: {},
  } as unknown as { email?: string | null; app_metadata?: Record<string, unknown>; user_metadata?: Record<string, unknown>; raw_user_meta_data?: Record<string, unknown> };
}

function escapeHtml(value: unknown) {
  return requireString(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function asStringArray(value: unknown) {
  if (Array.isArray(value)) return value.map((item) => requireString(item)).filter(Boolean);
  if (typeof value === "string") {
    return value.split(/[;,|]/).map((item) => item.trim()).filter(Boolean);
  }
  return [];
}

function normalizeLanguageLabel(value: unknown) {
  const text = requireString(value);
  if (!text) return "";
  const normalized = text.toLowerCase();
  if (["eng", "english"].includes(normalized)) return "ENGLISH";
  if (["hin", "hindi"].includes(normalized)) return "HINDI";
  if (["odia", "oriya", "odiya", "od"].includes(normalized)) return "ODIA";
  return text.toUpperCase();
}

function normalizeLanguageArray(value: unknown) {
  return uniqueStrings(asStringArray(value).map((item) => normalizeLanguageLabel(item)).filter(Boolean));
}

function parseBoolean(value: unknown) {
  if (typeof value === "boolean") return value;
  const normalized = requireString(value).toLowerCase();
  if (!normalized || normalized === "null") return null;
  if (["yes", "true", "1"].includes(normalized)) return true;
  if (["no", "false", "0"].includes(normalized)) return false;
  return null;
}

function parseNumber(value: unknown, fallback = 0) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const normalized = requireString(value).replace(/[^0-9.-]/g, "");
  if (!normalized) return fallback;
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function normalizeCell(value: unknown) {
  if (value === null || value === undefined) return null;
  const text = String(value).trim();
  return text.length ? text : null;
}

function rowValue(row: Record<string, unknown>, aliases: string[]) {
  for (const alias of aliases) {
    if (Object.prototype.hasOwnProperty.call(row, alias)) {
      const value = requireString(row[alias]);
      if (value) return value;
    }
  }
  const normalizedAliases = aliases.map((alias) => alias.toLowerCase().replace(/[^a-z0-9]+/g, ""));
  for (const [key, value] of Object.entries(row)) {
    const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]+/g, "");
    if (normalizedAliases.includes(normalizedKey)) {
      const text = requireString(value);
      if (text) return text;
    }
  }
  return "";
}

function rowDateValue(row: Record<string, unknown>, aliases: string[]) {
  for (const alias of aliases) {
    if (Object.prototype.hasOwnProperty.call(row, alias)) {
      const parsed = parseWorkbookDate(row[alias])?.slice(0, 10) || requireString(row[alias]);
      if (parsed) return parsed;
    }
  }
  const normalizedAliases = aliases.map((alias) => alias.toLowerCase().replace(/[^a-z0-9]+/g, ""));
  for (const [key, value] of Object.entries(row)) {
    const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]+/g, "");
    if (normalizedAliases.includes(normalizedKey)) {
      const parsed = parseWorkbookDate(value)?.slice(0, 10) || requireString(value);
      if (parsed) return parsed;
    }
  }
  return "";
}

function stripHtml(html: string | null) {
  if (!html) return "";
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/\s+/g, " ")
    .trim();
}

function uniqueStrings(values: string[]) {
  return [...new Set(values.filter(Boolean))];
}

function getStoredPayloadRecord(value: unknown) {
  if (!value || typeof value !== "object") return {} as Record<string, unknown>;
  const record = value as Record<string, unknown>;
  if (record.payload && typeof record.payload === "object") {
    return record.payload as Record<string, unknown>;
  }
  return record;
}

async function invalidateAskGreSearchCache() {
  const target = `${askGreBaseUrl.replace(/\/+$/, "")}/api/admin/cache`;
  try {
    const response = await fetch(target, {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "cache-control": "no-store",
      },
      body: JSON.stringify({ source: "gre-mis-admin" }),
    });
    if (!response.ok) {
      const text = await response.text();
      console.warn("AskGRE cache invalidation failed:", response.status, text);
    }
  } catch (error) {
    console.warn("AskGRE cache invalidation request failed:", error instanceof Error ? error.message : String(error));
  }
}

function tokenizeLooseText(value: unknown, minimumLength = 4) {
  return uniqueStrings(
    requireString(value)
      .toLowerCase()
      .split(/[^a-z0-9]+/)
      .map((entry) => entry.trim())
      .filter((entry) => entry.length >= minimumLength),
  );
}

function normalizeNeedTag(tag: unknown) {
  return requireString(tag)
    .toLowerCase()
    .replace(/[_/]+/g, " ")
    .replace(/\bstreet\s+lights?\b/g, "street lighting")
    .replace(/\bdecentralised\b/g, "decentralized")
    .replace(/\blow[\s-]+cost\b/g, "low cost")
    .replace(/\bvillage[\s-]+level\b/g, "village level")
    .replace(/\blast[\s-]+mile\b/g, "last mile")
    .replace(/\s+/g, " ")
    .trim();
}

function isWeakNeedTag(tag: unknown) {
  const normalized = normalizeNeedTag(tag);
  if (!normalized || normalized.length < 3) return true;
  if (needTagStopwords.has(normalized)) return true;
  const parts = normalized.split(" ").filter(Boolean);
  if (!parts.length) return true;
  if (parts.every((part) => needTagStopwords.has(part) || /^\d+$/.test(part))) return true;
  if (parts.length === 1 && parts[0].length < 5 && !["solar", "rural", "laser"].includes(parts[0])) return true;
  return false;
}

function extractNeedPriorityPhrases(value: unknown) {
  const text = requireString(value).toLowerCase();
  return uniqueStrings(
    needPriorityPhraseSignals
      .filter((signal) => signal.patterns.some((pattern) => text.includes(pattern)))
      .map((signal) => signal.label),
  );
}

function extractNeedWindowPhrases(value: unknown) {
  const text = requireString(value)
    .toLowerCase()
    .replace(/[\r\n]+/g, " ")
    .replace(/\s+/g, " ");
  const phrases: string[] = [];
  const windowPatterns = [
    /seeking\s+([^.;:]+)/g,
    /looking for\s+([^.;:]+)/g,
    /solutions should be\s+([^.;:]+)/g,
    /we need\s+([^.;:]+)/g,
    /required to be\s+([^.;:]+)/g,
  ];
  for (const pattern of windowPatterns) {
    for (const match of text.matchAll(pattern)) {
      const fragment = requireString(match[1]);
      if (!fragment) continue;
      for (const piece of fragment.split(/,| and /g)) {
        const normalized = normalizeNeedTag(piece);
        if (!isWeakNeedTag(normalized) && normalized.split(" ").length <= 4) phrases.push(normalized);
      }
    }
  }
  return uniqueStrings(phrases);
}

const genericThematicTerms = new Set([
  "business",
  "business consultation",
  "business mentoring",
  "business development",
  "training",
  "capacity building",
  "advisory",
  "consulting",
  "consultancy",
  "technology",
  "infrastructure",
  "vendor",
  "connect collaborate",
  "entrepreneurship",
  "business training",
  "finance",
  "funding",
  "investments",
]);

const domainStopwords = new Set([
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
  "need",
  "needs",
  "management",
  "going",
  "under",
  "null",
  "shared",
  "solutions",
  "solution",
  "business",
  "consultation",
  "mentoring",
]);

const needTagStopwords = new Set([
  "a",
  "an",
  "and",
  "are",
  "as",
  "at",
  "be",
  "been",
  "being",
  "by",
  "can",
  "could",
  "demands",
  "different",
  "for",
  "from",
  "functional",
  "government",
  "have",
  "in",
  "into",
  "is",
  "it",
  "its",
  "need",
  "needs",
  "of",
  "offerings",
  "our",
  "over",
  "raised",
  "scheme",
  "schemes",
  "service",
  "services",
  "should",
  "solution",
  "solutions",
  "support",
  "that",
  "the",
  "their",
  "there",
  "these",
  "this",
  "to",
  "transfer",
  "under",
  "we",
  "while",
  "with",
]);

const needPriorityPhraseSignals = [
  { label: "street lighting", patterns: ["street lighting", "street lights", "street light", "streetlight", "streetlights"] },
  { label: "rural street lighting", patterns: ["rural street lighting", "rural street lights", "rural street light"] },
  { label: "solar street lighting", patterns: ["solar street lighting", "solar street lights", "solar street light"] },
  { label: "community lighting", patterns: ["community lighting", "lighting solutions for the community"] },
  { label: "decentralized", patterns: ["decentralized", "decentralised"] },
  { label: "low cost", patterns: ["low cost", "low-cost", "cost effective", "cost-effective"] },
  { label: "affordable", patterns: ["affordable"] },
  { label: "maintainable", patterns: ["maintainable", "easy to maintain", "easy maintenance"] },
  { label: "village level", patterns: ["village level", "village-level"] },
  { label: "scalable", patterns: ["scalable", "scale across", "scale-up"] },
  { label: "small habitations", patterns: ["small habitations", "small settlements", "small villages"] },
  { label: "remote settlements", patterns: ["remote settlements", "remote settlement", "remote habitations", "remote villages"] },
  { label: "last mile delivery", patterns: ["last-mile delivery", "last mile delivery"] },
  { label: "public lighting", patterns: ["public lighting", "community street lighting"] },
  { label: "off-grid lighting", patterns: ["off-grid lighting", "off grid lighting"] },
  { label: "community infrastructure", patterns: ["community infrastructure", "shared infrastructure"] },
];

const ruleThemeSignals = [
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

const ruleNeedKindPatterns = [
  { label: "finance", patterns: ["finance", "funding", "credit", "loan", "financial support"] },
  { label: "product", patterns: ["machine", "machinery", "equipment", "plant", "street light", "product", "unit"] },
  { label: "knowledge", patterns: ["manual", "video", "sop", "blog", "knowledge product", "guide"] },
  { label: "service", patterns: ["training", "consulting", "consultancy", "mentoring", "technology transfer", "advisory"] },
];

const ruleServiceKindPatterns = [
  { label: "training", patterns: ["training", "capacity building"] },
  { label: "consulting", patterns: ["consulting", "consultancy", "business consultation"] },
  { label: "mentoring", patterns: ["mentoring", "business mentoring"] },
  { label: "technology transfer", patterns: ["technology transfer"] },
  { label: "advisory", patterns: ["advisory"] },
];

function splitLooseList(value: string | null, separators = [",", ";", "\n"]) {
  if (!value) return [];
  let working = value;
  separators.forEach((separator) => {
    working = working.split(separator).join("|");
  });
  return uniqueStrings(
    working
      .split("|")
      .map((entry) => entry.trim())
      .filter(Boolean),
  );
}

function splitGeographies(value: string | null) {
  if (!value) return [];
  if (value.includes(";") || value.includes("\n") || value.includes("|")) {
    return splitLooseList(value, [";", "\n", "|"]);
  }
  return [value.trim()].filter(Boolean);
}

function buildSearchDocument(parts: Array<string | null | string[]>) {
  return parts
    .flatMap((part) => Array.isArray(part) ? part : [part])
    .filter(Boolean)
    .join(" | ")
    .replace(/\s+/g, " ")
    .trim();
}

function chunkArray<T>(items: T[], size = 250) {
  const out: T[][] = [];
  for (let index = 0; index < items.length; index += size) {
    out.push(items.slice(index, index + size));
  }
  return out;
}

function parseWorkbookDate(value: unknown) {
  const text = requireString(value);
  if (!text) return null;
  const parts = text.match(/^(\d{1,2})-(\d{1,2})-(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (parts) {
    const [, day, month, year, hour = "0", minute = "0", second = "0"] = parts;
    return new Date(Date.UTC(Number(year), Number(month) - 1, Number(day), Number(hour), Number(minute), Number(second))).toISOString();
  }
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed.toISOString();
}

function readWorkbookRows(buffer: ArrayBuffer) {
  const workbook = XLSX.read(buffer, { type: "array" });
  const firstSheet = workbook.SheetNames[0];
  return XLSX.utils.sheet_to_json<Record<string, unknown>>(workbook.Sheets[firstSheet], {
    defval: "",
    raw: false,
  });
}

function bytesToBase64(bytes: Uint8Array) {
  let binary = "";
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    binary += String.fromCharCode(...bytes.slice(index, index + chunkSize));
  }
  return btoa(binary);
}

async function createWorkbookDownloadPayload(buffer: ArrayBuffer, fileName: string, mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
  return {
    fileName,
    mimeType,
    base64: bytesToBase64(new Uint8Array(buffer)),
  };
}

function escapeExcelHtml(value: unknown) {
  return requireString(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/\n/g, "<br>");
}

function escapeExcelXml(value: unknown) {
  const str = value === null || value === undefined ? "" : String(value);
  return str
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, " ")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function buildSeekerSolutionEntries(notes: unknown) {
  const text = requireString(notes).replace(/\r/g, "\n");
  if (!text) return "";
  const section = text.match(/solutions\s+shared\s*:\s*([\s\S]*?)(?:\n\s*solution\s+links\s+shared\s+with\s+seeker|$)/i)?.[1] || "";
  const entries = section.split(/,|;|\n/).map((s) => s.trim()).filter(Boolean);
  return entries.map((entry) => {
    const parts = entry.split(/\s*:\s*/);
    if (parts.length >= 3) {
      const name = parts[0];
      const provider = parts[1];
      const url = parts.slice(2).join(": ");
      return `${escapeExcelXml(name)} | ${escapeExcelXml(provider)}\n${escapeExcelXml(url)}`;
    }
    if (parts.length === 2) {
      return `${escapeExcelXml(parts[0])}\n${escapeExcelXml(parts[1])}`;
    }
    return escapeExcelXml(entry);
  }).join("\n\n");
}

function buildSeekerStatusStyle(styleId: string, fontColor: string, fillColor: string) {
  return `<Style ss:ID="${styleId}"><Alignment ss:Horizontal="Left" ss:Vertical="Top" ss:WrapText="1"/><Font ss:FontName="Calibri" ss:Size="11" ss:Color="${fontColor}"/><Interior ss:Color="${fillColor}" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/></Borders></Style>`;
}

function buildStyledTrackerHtml({
  seekerLabel,
  totalNeeds,
  solutionProviderNeeded,
  seekerResponsePending,
  solutionsImplemented,
  needs,
}: {
  seekerLabel: string;
  totalNeeds: number;
  solutionProviderNeeded: number;
  seekerResponsePending: number;
  solutionsImplemented: number;
  needs: Array<Record<string, unknown>>;
}) {
  const generatedOn = new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata" });
  const emptyCell = () => '<Cell ss:StyleID="Empty"/>';
  const labelCell = (value: unknown, styleId: string) =>
    `<Cell ss:StyleID="${styleId}"><Data ss:Type="String">${escapeExcelXml(value)}</Data></Cell>`;
  const valueCell = (value: unknown, styleId: string, type = "String") =>
    value === "" || value === undefined || value === null
      ? emptyCell()
      : `<Cell ss:StyleID="${styleId}"><Data ss:Type="${type}">${escapeExcelXml(value)}</Data></Cell>`;
  const dataCell = (value: unknown, styleId: string) => {
    const text = requireString(value);
    if (!text) return emptyCell();
    return `<Cell ss:StyleID="${styleId}"><Data ss:Type="String">${escapeExcelXml(value)}</Data></Cell>`;
  };

  const dStyle = (need: Record<string, unknown>) => {
    const s = requireString(need.status).toLowerCase();
    if (s === "accepted") return "StAccepted";
    if (s === "closed") return "StClosed";
    return "BodyBordered";
  };
  const eStyle = (need: Record<string, unknown>) => {
    const s = requireString(need.internal_status).toLowerCase();
    if (s.includes("connection made")) return "StConnMade";
    if (s.includes("need solution")) return "StNeedSol";
    if (s.includes("blocked")) return "StBlocked";
    if (s === "complete") return "StClosed";
    return "BodyBordered";
  };
  const fStyle = (need: Record<string, unknown>) => {
    const s = requireString(need.next_action).toLowerCase();
    if (s === "follow up with seeker") return "StFollowUp";
    return "StOtherAction";
  };
  const gStyle = (need: Record<string, unknown>) => {
    const s = requireString(need.seeker_provider_agreement).toLowerCase();
    if (!s) return "StNoAgree";
    if (s === "agreement completed") return "StAgreeDone";
    if (s.includes("no agreement")) return "StNoAgree";
    return "StAgreeOther";
  };

  const safeStr = (v: unknown) => v === null || v === undefined || v === "" ? null : requireString(v);
  const makeCell = (col: number, value: unknown, styleId: string, type = "String") => {
    if (value === null || value === undefined || value === "") {
      return `<Cell ss:Index="${col}" ss:StyleID="Empty"/>`;
    }
    return `<Cell ss:Index="${col}" ss:StyleID="${styleId}"><Data ss:Type="${type}">${escapeExcelXml(value)}</Data></Cell>`;
  };

  const dataRows = needs.map((need, index) => {
    const solutions = buildSeekerSolutionEntries(need.curation_notes);
    return `
    <Row ss:AutoFitHeight="0" ss:Height="80">
      ${makeCell(1, index + 1, "Serial", "Number")}
      ${makeCell(2, safeStr(need.problem_statement), "BodyBordered")}
      ${makeCell(3, safeStr(solutions), "BodyBordered")}
      ${makeCell(4, safeStr(need.status), dStyle(need))}
      ${makeCell(5, safeStr(need.internal_status), eStyle(need))}
      ${makeCell(6, safeStr(need.next_action), fStyle(need))}
      ${makeCell(7, safeStr(need.seeker_provider_agreement), gStyle(need))}
      ${makeCell(8, null, "BodyNoBorder")}
    </Row>`;
  }).join("");

  return `<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
  <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
    <Title>Request Tracker Sheet - ${escapeExcelXml(seekerLabel)}</Title>
    <Author>GRE MIS</Author>
    <Created>${new Date().toISOString()}</Created>
  </DocumentProperties>
  <Styles>
    <Style ss:ID="Default" ss:Name="Normal">
      <Alignment ss:Vertical="Top"/>
      <Font ss:FontName="Calibri" ss:Size="11" ss:Color="#333333"/>
    </Style>
    <Style ss:ID="Empty"/>
    <Style ss:ID="HeaderCenter">
      <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
      <Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="#333333"/>
      <Interior ss:Color="#F2F2F2" ss:Pattern="Solid"/>
      <Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/></Borders>
    </Style>
    <Style ss:ID="TitleLabel">
      <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
      <Font ss:FontName="Calibri" ss:Size="12" ss:Bold="1" ss:Color="#333333"/>
    </Style>
    <Style ss:ID="TitleValue">
      <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
      <Font ss:FontName="Calibri" ss:Size="12" ss:Color="#333333"/>
    </Style>
    <Style ss:ID="SummaryGrey">
      <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
      <Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="#333333"/>
      <Interior ss:Color="#F2F2F2" ss:Pattern="Solid"/>
      <Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/></Borders>
    </Style>
    <Style ss:ID="SummaryValueGrey">
      <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
      <Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="#333333"/>
      <Interior ss:Color="#F2F2F2" ss:Pattern="Solid"/>
      <Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/></Borders>
    </Style>
    <Style ss:ID="SummaryBlue">
      <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
      <Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="#333333"/>
      <Interior ss:Color="#D6E4F0" ss:Pattern="Solid"/>
      <Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/></Borders>
    </Style>
    <Style ss:ID="SummaryValueBlue">
      <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
      <Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="#333333"/>
      <Interior ss:Color="#D6E4F0" ss:Pattern="Solid"/>
      <Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/></Borders>
    </Style>
    <Style ss:ID="SummaryOrange">
      <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
      <Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="#333333"/>
      <Interior ss:Color="#FCE4D6" ss:Pattern="Solid"/>
      <Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/></Borders>
    </Style>
    <Style ss:ID="SummaryValueOrange">
      <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
      <Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="#333333"/>
      <Interior ss:Color="#FCE4D6" ss:Pattern="Solid"/>
      <Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/></Borders>
    </Style>
    <Style ss:ID="SummaryGreen">
      <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
      <Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="#333333"/>
      <Interior ss:Color="#E2EFDA" ss:Pattern="Solid"/>
      <Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/></Borders>
    </Style>
    <Style ss:ID="SummaryValueGreen">
      <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
      <Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="#333333"/>
      <Interior ss:Color="#E2EFDA" ss:Pattern="Solid"/>
      <Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/></Borders>
    </Style>
    <Style ss:ID="BodyBordered">
      <Alignment ss:Horizontal="Left" ss:Vertical="Top" ss:WrapText="1"/>
      <Font ss:FontName="Calibri" ss:Size="11" ss:Color="#333333"/>
      <Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/></Borders>
    </Style>
    <Style ss:ID="BodyNoBorder">
      <Alignment ss:Horizontal="Left" ss:Vertical="Top" ss:WrapText="1"/>
      <Font ss:FontName="Calibri" ss:Size="11" ss:Color="#333333"/>
    </Style>
    <Style ss:ID="Serial">
      <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
      <Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="#333333"/>
      <Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D9D9"/></Borders>
    </Style>
    ${buildSeekerStatusStyle("StAccepted", "#333333", "#FFF2CC")}
    ${buildSeekerStatusStyle("StClosed", "#333333", "#E2EFDA")}
    ${buildSeekerStatusStyle("StConnMade", "#333333", "#D6E4F0")}
    ${buildSeekerStatusStyle("StNeedSol", "#333333", "#FFF2CC")}
    ${buildSeekerStatusStyle("StBlocked", "#333333", "#FCE4EC")}
    ${buildSeekerStatusStyle("StFollowUp", "#333333", "#FFF2CC")}
    ${buildSeekerStatusStyle("StOtherAction", "#333333", "#FCE4EC")}
    ${buildSeekerStatusStyle("StAgreeDone", "#333333", "#E2EFDA")}
    ${buildSeekerStatusStyle("StNoAgree", "#333333", "#FCE4EC")}
    ${buildSeekerStatusStyle("StAgreeOther", "#333333", "#FFF2CC")}
  </Styles>
  <Worksheet ss:Name="Tracker">
    <Table ss:ExpandedColumnCount="8" ss:ExpandedRowCount="${needs.length + 10}" x:FullColumns="1" x:FullRows="1"
           ss:DefaultRowHeight="16">
      <Column ss:Width="30.48"/>
      <Column ss:Width="376.63"/>
      <Column ss:Width="298.37"/>
      <Column ss:Width="108.11"/>
      <Column ss:Width="144.37"/>
      <Column ss:Width="138.63"/>
      <Column ss:Width="152"/>
      <Column ss:Width="188.26"/>
      <Row ss:Height="24">
        ${emptyCell()}
        ${labelCell("Request Tracker Sheet", "TitleLabel")}
        ${labelCell(seekerLabel, "TitleValue")}
      </Row>
      <Row ss:Height="20">
        ${emptyCell()}
        ${labelCell("Generated On", "TitleLabel")}
        ${labelCell(generatedOn, "TitleValue")}
      </Row>
      <Row ss:Height="10"><Cell ss:StyleID="Empty"/></Row>
      <Row ss:Height="24">
        ${emptyCell()}
        ${labelCell("Total Needs Logged", "SummaryGrey")}
        ${valueCell(totalNeeds, "SummaryValueGrey", "Number")}
      </Row>
      <Row ss:Height="24">
        ${emptyCell()}
        ${labelCell("Solution Provider Needed", "SummaryBlue")}
        ${valueCell(solutionProviderNeeded, "SummaryValueBlue", "Number")}
      </Row>
      <Row ss:Height="24">
        ${emptyCell()}
        ${labelCell("Solution Provider Connected, Seeker Response Pending", "SummaryOrange")}
        ${valueCell(seekerResponsePending, "SummaryValueOrange", "Number")}
      </Row>
      <Row ss:Height="24">
        ${emptyCell()}
        ${labelCell("Solutions Implemented", "SummaryGreen")}
        ${valueCell(solutionsImplemented, "SummaryValueGreen", "Number")}
      </Row>
      <Row ss:Height="10"><Cell ss:StyleID="Empty"/></Row>
      <Row ss:Height="28">
        ${labelCell("Sr", "HeaderCenter")}
        ${labelCell("Request", "HeaderCenter")}
        ${labelCell("Solution Shared", "HeaderCenter")}
        ${labelCell("Status", "HeaderCenter")}
        ${labelCell("Internal Status", "HeaderCenter")}
        ${labelCell("Next Action", "HeaderCenter")}
        ${labelCell("Seeker-Provider Agreement", "HeaderCenter")}
        ${labelCell("Comments", "HeaderCenter")}
      </Row>
      ${dataRows}
    </Table>
    <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
      <FreezePanes/>
      <FrozenNoSplit/>
      <SplitHorizontal>9</SplitHorizontal>
      <TopRowBottomPane>9</TopRowBottomPane>
      <ActivePane>2</ActivePane>
    </WorksheetOptions>
  </Worksheet>
</Workbook>`;
}

async function loginToGre() {
  if (!greMasterLogin || !greMasterPassword || !greMasterTenant) {
    throw new Error("GRE master credentials are not configured in Supabase secrets.");
  }

  const response = await fetch(`${greLoginBaseUrl}/commons-iam-service/api/v1/obo/cross/login`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Origin: greSiteOrigin,
      Referer: `${greSiteOrigin}/`,
    },
    body: JSON.stringify({
      userLogin: greMasterLogin,
      tenantLogin: greMasterTenant,
      password: greMasterPassword,
    }),
  });
  const data = await response.json().catch(() => null);
  const sessionId = requireString(data?.sessionId);
  if (!response.ok || !sessionId) {
    throw new Error(data?.error?.message || data?.message || "GRE login failed.");
  }
  return sessionId;
}

async function fetchGreReportWorkbook(reportName: string, limitRowCount: number, filePrefix: string) {
  const sessionId = await loginToGre();
  const params = `MARKET_ID=${greMarketId}--LIMIT_OFFSET=0--LIMIT_ROW_COUNT=${limitRowCount}`;
  const response = await fetch(
    `${greLoginBaseUrl}/commons-report-service/api/v1/datasets/name/${encodeURIComponent(reportName)}/execute/download?fileType=EXCEL&param=${encodeURIComponent(params)}`,
    {
      headers: {
        "x-sessionid": sessionId,
        Accept: "application/octet-stream, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, */*",
        Origin: greSiteOrigin,
        Referer: `${greSiteOrigin}/`,
      },
    },
  );
  if (!response.ok) {
    const message = await response.text().catch(() => "");
    throw new Error(message || `GRE report download failed for ${reportName}.`);
  }
  const buffer = await response.arrayBuffer();
  return {
    buffer,
    fileName: `${filePrefix}_${Date.now()}.xlsx`,
  };
}

function normalizeComparable(value: unknown) {
  return requireString(value).toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

const greRefDataCache = new Map<string, Record<string, unknown>[]>();
const greProductRefDataCache = new Map<string, Record<string, unknown>[]>();
const greProductRefHierarchyCache = new Map<string, Record<string, unknown>[]>();
let greTenantUsersCache: Record<string, unknown>[] | null = null;

const greServiceSpecTemplate = [
  { id: 74, sequence: 1, paramName: "Trainer Name", specificationCode: "SO_RESOURCE_PERSON_NAME", groupCode: "SPECIFICATION_GROUP.RESOURCE_PERSON_CONTACT", dataType: "TEXT", isMandatory: true, isMultiValued: false },
  { id: 75, sequence: 2, paramName: "Trainer Email Address", specificationCode: "SO_RESOURCE_PERSON_EMAIL", groupCode: "SPECIFICATION_GROUP.RESOURCE_PERSON_CONTACT", dataType: "TEXT", isMandatory: true, isMultiValued: false },
  { id: 72, sequence: 3, paramName: "Languages", specificationCode: "SOLUTION_LANGUAGE", groupCode: "SPECIFICATION_GROUP.SOLUTION", dataType: "CHIP", isMandatory: true, isMultiValued: true, customRefClass: "CLASS.TRAINING_LANGUAGE" },
  { id: 73, sequence: 4, paramName: "Geographies", specificationCode: "SOLUTION_GEOGRAPHIES", groupCode: "SPECIFICATION_GROUP.SOLUTION", dataType: "LOCATION", isMandatory: true, isMultiValued: true },
  { id: 76, sequence: 5, paramName: "Trainer Phone Number", specificationCode: "SO_RESOURCE_PERSON_PHONE", groupCode: "SPECIFICATION_GROUP.RESOURCE_PERSON_CONTACT", dataType: "TEXT", isMandatory: true, isMultiValued: false },
  { id: 77, sequence: 6, paramName: "Trainer Details", specificationCode: "SO_TRAINER_DETAILS", groupCode: "SPECIFICATION_GROUP.RESOURCE_PERSON_CONTACT", dataType: "RICH_TEXT", isMandatory: true, isMultiValued: false },
  { id: 78, sequence: 7, paramName: "Duration", specificationCode: "SO_DURATION", groupCode: "SPECIFICATION_GROUP.SERVICE_OFFERINGS", dataType: "DURATION", isMandatory: true, isMultiValued: false, customRefClass: "CLASS.TRAINING_DURATION" },
  { id: 79, sequence: 8, paramName: "Prerequisites - Participants and Training", specificationCode: "SO_PR_PARTICIPANTS_AND_TRAINING_LOCATION", groupCode: "SPECIFICATION_GROUP.SERVICE_OFFERINGS", dataType: "TEXT_AREA", isMandatory: true, isMultiValued: false },
  { id: 95, sequence: 9, paramName: "Location Availability", specificationCode: "SO_LOCATION_AVAILABILITY", groupCode: "SPECIFICATION_GROUP.SERVICE_OFFERINGS", dataType: "CHIP", isMandatory: true, isMultiValued: true, customRefClass: "CLASS.SERVICE_LOCATION_AVAILABILITY" },
  { id: 82, sequence: 10, paramName: "Cost", specificationCode: "SO_COST", groupCode: "SPECIFICATION_GROUP.SERVICE_OFFERINGS", dataType: "COST_WITH_SPEC", isMandatory: true, isMultiValued: false, customRefClass: "CLASS.COST_SPEC" },
  { id: 94, sequence: 11, paramName: "Remarks on Cost", specificationCode: "SO_COST_REMARKS", groupCode: "SPECIFICATION_GROUP.SERVICE_OFFERINGS", dataType: "TEXT", isMandatory: false, isMultiValued: false },
  { id: 83, sequence: 12, paramName: "Support post Service", specificationCode: "SO_SUPPORT_POST_SERVICE", groupCode: "SPECIFICATION_GROUP.SUPPORT_POST_SERVICE_AND_COST", dataType: "CHIP", isMandatory: true, isMultiValued: false, customRefClass: "CLASS.POST_SERVICE" },
  { id: 84, sequence: 13, paramName: "Support post Service Cost", specificationCode: "SO_SUPPORT_POST_SERVICE_COST", groupCode: "SPECIFICATION_GROUP.SUPPORT_POST_SERVICE_AND_COST", dataType: "CHIP", isMandatory: true, isMultiValued: false, customRefClass: "CLASS.POST_SERVICE_COST" },
  { id: 85, sequence: 14, paramName: "Is it offered - Online or Offline?", specificationCode: "SO_OFFERED_ONLINE_OR_OFFLINE", groupCode: "SPECIFICATION_GROUP.SERVICE_OFFERINGS", dataType: "CHIP", isMandatory: true, isMultiValued: false, customRefClass: "CLASS.SESSION_AVAILABILITY" },
  { id: 86, sequence: 15, paramName: "Certification Offered", specificationCode: "SO_CERTIFICATION_OFFERED", groupCode: "SPECIFICATION_GROUP.SERVICE_OFFERINGS", dataType: "CHIP", isMandatory: true, isMultiValued: false, customRefClass: "CLASS.CERTIFICATION" },
  { id: 87, sequence: 16, paramName: "Service offering Brochure", specificationCode: "SO_SERVICE_OFFERING_BROCHURE", groupCode: "SPECIFICATION_GROUP.SERVICE_OFFERINGS", dataType: "ATTACHMENT", isMandatory: false, isMultiValued: true },
];

const greServiceOfferingSubtypeFallbacks: Record<string, string> = {
  "training": "OFFERINGS_CATEGORY.TRAINING",
  "consulting": "OFFERINGS_CATEGORY.CONSULTING",
  "consulting mentoring": "OFFERINGS_CATEGORY.CONSULTING",
  "consulting mentoring.": "OFFERINGS_CATEGORY.CONSULTING",
  "mentoring": "OFFERINGS_CATEGORY.MENTORING",
  "financial support": "OFFERINGS_CATEGORY.FINANCIAL_SUPPORT",
  "market support": "OFFERINGS_CATEGORY.MARKET_SUPPORT",
  "technology transfer": "OFFERINGS_CATEGORY.TECHNOLOGY_TRANSFER",
};

const greTrainingLanguageFallbacks: Record<string, string> = {
  "eng": "TRAINING_LANGUAGE.ENG",
  "english": "TRAINING_LANGUAGE.ENG",
  "hin": "TRAINING_LANGUAGE.HIN",
  "hindi": "TRAINING_LANGUAGE.HIN",
  "bengali": "TRAINING_LANGUAGE.BENGALI",
  "bangla": "TRAINING_LANGUAGE.BENGALI",
  "marathi": "TRAINING_LANGUAGE.MARATHI",
  "kannada": "TRAINING_LANGUAGE.KANNADA",
  "malayalam": "TRAINING_LANGUAGE.MALAYALAM",
  "tamil": "TRAINING_LANGUAGE.TAMIL",
  "telugu": "TRAINING_LANGUAGE.TELUGU",
  "odia": "TRAINING_LANGUAGE.ODIA",
  "oriya": "TRAINING_LANGUAGE.ODIA",
  "gujarati": "TRAINING_LANGUAGE.GUJARATI",
  "punjabi": "TRAINING_LANGUAGE.PUNJABI",
  "assamese": "TRAINING_LANGUAGE.ASSAMESE",
  "urdu": "TRAINING_LANGUAGE.URDU",
};

type GreUserResolution = {
  status: string;
  greUserId: number | null;
  greLoginName: string;
  message: string;
  profile?: Record<string, unknown> | null;
};

async function fetchGreJson(path: string, sessionId?: string) {
  const resolvedSessionId = sessionId || await loginToGre();
  const response = await fetch(`${greLoginBaseUrl}${path}`, {
    headers: {
      "x-sessionid": resolvedSessionId,
      Accept: "application/json, text/plain, */*",
      Origin: greSiteOrigin,
      Referer: `${greSiteOrigin}/`,
    },
  });
  const data = await response.json().catch(() => null);
  if (!response.ok) {
    throw new Error(data?.error?.message || data?.message || `GRE request failed for ${path}.`);
  }
  return { sessionId: resolvedSessionId, data };
}

async function requestGreGatewayJson(
  path: string,
  method = "GET",
  body?: unknown,
  sessionId?: string,
) {
  const resolvedSessionId = sessionId || await loginToGre();
  const payload = body === undefined ? undefined : JSON.stringify(body);
  const response = await fetch(`${greLoginBaseUrl}${path}`, {
    method,
    headers: {
      "x-sessionid": resolvedSessionId,
      Accept: "application/json, text/plain, */*",
      ...(payload ? { "Content-Type": "application/json" } : {}),
      Origin: greSiteOrigin,
      Referer: `${greSiteOrigin}/`,
    },
    body: payload,
  });
  const text = await response.text().catch(() => "");
  let data: unknown = null;
  if (text) {
    try {
      data = JSON.parse(text);
    } catch {
      data = text;
    }
  }
  if (!response.ok) {
    const parsed = (data && typeof data === "object") ? data as Record<string, unknown> : null;
    throw new Error(
      requireString(parsed?.error?.message) ||
      requireString(parsed?.errorMessage) ||
      requireString(parsed?.message) ||
      requireString(data) ||
      `GRE ${method} failed for ${path}.`,
    );
  }
  return { sessionId: resolvedSessionId, data };
}

async function fetchGreProductRefData(classCode: string, sessionId?: string) {
  if (greProductRefDataCache.has(classCode)) return greProductRefDataCache.get(classCode) || [];
  const { data } = await requestGreGatewayJson(
    `/commons-product-service/api/v1/ref-data/class/${encodeURIComponent(classCode)}?direction=ASC&fetchContextDataOnly=false&languageCode=ENG&page=0&size=100`,
    "GET",
    undefined,
    sessionId,
  );
  const elements = Array.isArray((data as Record<string, unknown>)?.elements)
    ? ((data as Record<string, unknown>).elements as Record<string, unknown>[])
    : Array.isArray(data)
      ? data as Record<string, unknown>[]
      : [];
  greProductRefDataCache.set(classCode, elements);
  return elements;
}

async function fetchGreProductRefHierarchy(parentCode: string, sessionId?: string) {
  if (greProductRefHierarchyCache.has(parentCode)) return greProductRefHierarchyCache.get(parentCode) || [];
  const { data } = await requestGreGatewayJson(
    `/commons-product-service/api/v1/ref-data-hierarchy/parent?childClassCode=${encodeURIComponent("CLASS.OFFERINGS_CATEGORY")}&direction=ASC&fetchContextDataOnly=false&languageCode=ENG&page=0&size=100&parentCode=${encodeURIComponent(parentCode)}`,
    "GET",
    undefined,
    sessionId,
  );
  const elements = Array.isArray((data as Record<string, unknown>)?.elements)
    ? ((data as Record<string, unknown>).elements as Record<string, unknown>[])
    : Array.isArray(data)
      ? data as Record<string, unknown>[]
      : [];
  greProductRefHierarchyCache.set(parentCode, elements);
  return elements;
}

async function fetchGreTenantUsers(forceRefresh = false) {
  if (!forceRefresh && greTenantUsersCache) return greTenantUsersCache;
  const { data: countData } = await fetchGreJson(
    `/commons-report-service/api/v1/datasets/name/GET_USER_LIST_MARKIFY_COUNT/execute?params=${encodeURIComponent("ROLECODE=-1")}`,
  );
  const totalCount = Math.max(
    parseNumber((Array.isArray(countData) ? countData[0]?.totalCount : countData?.totalCount) ?? 0, 0),
    25,
  );
  const params = `ROLECODE=-1--SORT_BY=u.id--SORT_ORDER=desc--LIMIT_OFFSET=0--LIMIT_ROW_COUNT=${Math.min(totalCount + 10, 500)}`;
  const { data } = await fetchGreJson(
    `/commons-report-service/api/v1/datasets/name/GET_USER_LIST_MARKIFY/execute?params=${encodeURIComponent(params)}`,
  );
  greTenantUsersCache = Array.isArray(data) ? data as Record<string, unknown>[] : [];
  return greTenantUsersCache;
}

async function callGreCtld(path: string, method = "POST", body: unknown = {}) {
  const sessionId = await loginToGre();
  const payload = method === "POST" ? JSON.stringify(body) : undefined;
  const response = await fetch(`${greCtldBaseUrl}${path}`, {
    method,
    headers: {
      "x-sessionid": sessionId,
      ...(payload ? { "Content-Type": "application/json" } : {}),
      Accept: "application/json, text/plain, */*",
      Origin: greSiteOrigin,
      Referer: `${greSiteOrigin}/`,
    },
    body: payload,
  });
  const text = await response.text().catch(() => "");
  let data: Record<string, unknown> | null = null;
  if (text) {
    try {
      data = JSON.parse(text);
    } catch {
      data = { message: text };
    }
  }
  if (!response.ok) {
    throw new Error(data?.errorMessage || data?.message || data?.error?.message || `GRE CTLD ${method} failed for ${path}.`);
  }
  return data;
}

async function callGrePtld(path: string, method = "POST", body: unknown = {}) {
  const sessionId = await loginToGre();
  const payload = method === "POST" || method === "PUT" ? JSON.stringify(body) : undefined;
  const response = await fetch(`${grePtldBaseUrl}${path}`, {
    method,
    headers: {
      "x-sessionid": sessionId,
      ...(payload ? { "Content-Type": "application/json" } : {}),
      Accept: "application/json, text/plain, */*",
      Origin: greSiteOrigin,
      Referer: `${greSiteOrigin}/`,
    },
    body: payload,
  });
  const text = await response.text().catch(() => "");
  let data: Record<string, unknown> | null = null;
  if (text) {
    try {
      data = JSON.parse(text);
    } catch {
      data = { message: text };
    }
  }
  if (!response.ok) {
    throw new Error(data?.errorMessage || data?.message || data?.error?.message || `GRE PTLD ${method} failed for ${path}.`);
  }
  return {
    data,
    text,
    headers: response.headers,
  };
}

async function checkGreIdentity(loginName: string) {
  const normalized = requireString(loginName);
  if (!normalized) return "NOT_FOUND";
  const { data } = await fetchGreJson(
    `/commons-iam-service/api/v1/obo/register/check?tenantName=${encodeURIComponent(greMasterTenant)}&loginName=${encodeURIComponent(normalized)}`,
  );
  return requireString((data as Record<string, unknown>)?.message || data || "NOT_FOUND").toUpperCase();
}

function getRoleCodeForMisRole(role: string) {
  if (role === "admin") return `role.${greMasterTenant}.admin`;
  if (role === "curator") return `role.${greMasterTenant}.curator`;
  if (role === "moderator") return `role.${greMasterTenant}.user`;
  return `role.${greMasterTenant}.user`;
}

function isAdminLikeMisRole(role: string) {
  return ["admin", "moderator"].includes(requireString(role).toLowerCase());
}

function isModeratorMisRole(role: string) {
  return requireString(role).toLowerCase() === "moderator";
}

function normalizePhoneLogin(value: unknown) {
  return requireString(value).replace(/\D+/g, "");
}

function getUserNameParts(user: Record<string, unknown>) {
  const fullName = requireString(user.full_name) || requireString(user.first_name) || requireString(user.username);
  const pieces = fullName.split(/\s+/).filter(Boolean);
  const firstName = requireString(user.first_name) || pieces[0] || "GRE";
  const lastName = pieces.length > 1 ? pieces.slice(1).join(" ") : "User";
  return { fullName: fullName || `${firstName} ${lastName}`.trim(), firstName, lastName };
}

function getGreLoginCandidates(user: Record<string, unknown>) {
  const candidates = [
    requireString(user.gre_login_name),
    normalizePhoneLogin(user.phone),
    requireString(user.email).toLowerCase(),
  ].filter(Boolean);
  return uniqueStrings(candidates);
}

function getGreDirectoryMatchScore(user: Record<string, unknown>, profile: Record<string, unknown>, candidates: string[]) {
  let score = 0;
  const expectedRoleCode = getRoleCodeForMisRole(requireString(user.role).toLowerCase() || "user");
  const profileLogin = requireString(profile.Login).replace(/\D+/g, "");
  const profileContact = requireString(profile.Contact).replace(/\D+/g, "");
  const profileEmail = requireString(profile.Email).toLowerCase();
  const profileName = normalizeComparable(profile.UserName || profile.firstName || "");
  const profileRoles = requireString(profile.Code);

  for (const candidate of candidates) {
    const normalizedCandidate = candidate.toLowerCase();
    if (normalizedCandidate && normalizedCandidate === profileEmail) score += 6;
    if (candidate.replace(/\D+/g, "") && candidate.replace(/\D+/g, "") === profileLogin) score += 7;
    if (candidate.replace(/\D+/g, "") && candidate.replace(/\D+/g, "") === profileContact) score += 5;
  }

  const fullName = normalizeComparable(user.full_name || user.first_name || user.username || "");
  if (fullName && fullName === profileName) score += 3;
  if (profileRoles.includes(expectedRoleCode)) score += 4;
  return score;
}

function findGreDirectoryProfile(user: Record<string, unknown>, profiles: Record<string, unknown>[]) {
  const candidates = getGreLoginCandidates(user);
  let best: Record<string, unknown> | null = null;
  let bestScore = 0;
  for (const profile of profiles) {
    const score = getGreDirectoryMatchScore(user, profile, candidates);
    if (score > bestScore) {
      best = profile;
      bestScore = score;
    }
  }
  return bestScore >= 6 ? best : null;
}

function buildGreMappedPatch(user: Record<string, unknown>, profile: Record<string, unknown>) {
  const misRole = requireString(user.role).toLowerCase() || "user";
  const liveRole = getMisRoleFromGreProfile(profile);
  const expectedRoleCode = getRoleCodeForMisRole(misRole);
  const roleCodes = requireString(profile.Code)
    .split(",")
    .map((entry) => entry.trim())
    .filter(Boolean);
  const roleMatches = roleCodes.includes(expectedRoleCode);
  const patch: Record<string, unknown> = {
    role: liveRole,
    gre_user_id: parseNumber(profile.id, 0) || null,
    gre_login_name: requireString(profile.Login) || requireString(profile.Contact) || requireString(user.gre_login_name) || null,
    gre_sync_status: roleMatches && liveRole === misRole ? "synced" : "mapped_role_mismatch",
    gre_sync_message: roleMatches && liveRole === misRole
      ? `GRE account mapped and ${misRole} role verified.`
      : `GRE account mapped, and MIS role has been refreshed to ${liveRole}. GRE roles are currently ${requireString(profile.RoleName) || liveRole || "different from MIS"}.`,
    gre_synced_at: new Date().toISOString(),
    gre_pending_role: null,
    gre_activation_mod_key: null,
    gre_activation_requested_at: null,
  };
  const email = requireString(profile.Email).toLowerCase();
  const phone = requireString(profile.Contact) || requireString(profile.Login);
  const firstName = requireString(profile.firstName);
  const fullName = requireString(profile.UserName) || firstName;
  if (email) patch.email = email;
  if (phone) patch.phone = phone;
  if (firstName) patch.first_name = firstName;
  if (fullName) patch.full_name = fullName;
  return patch;
}

function getMisRoleFromGreProfile(profile: Record<string, unknown>) {
  const codes = requireString(profile.Code).toLowerCase();
  const roleName = requireString(profile.RoleName).toLowerCase();
  if (codes.includes(`role.${greMasterTenant}.admin`) || /\badmin\b/.test(roleName)) return "admin";
  if (codes.includes(`role.${greMasterTenant}.curator`) || /\bcurator\b/.test(roleName)) return "curator";
  return "user";
}

function getGreRoleSetFromProfile(profile: Record<string, unknown>) {
  const codes = requireString(profile.Code).toLowerCase();
  const roleName = requireString(profile.RoleName).toLowerCase();
  const roles = new Set<string>();
  if (codes.includes(`role.${greMasterTenant}.admin`) || /\badmin\b/.test(roleName)) roles.add("admin");
  if (codes.includes(`role.${greMasterTenant}.curator`) || /\bcurator\b/.test(roleName)) roles.add("curator");
  if (codes.includes(`role.${greMasterTenant}.user`) || /\buser\b/.test(roleName) || roles.size === 0) roles.add("user");
  return roles;
}

function buildUniqueUsername(base: string, existingUsers: Record<string, unknown>[], greUserId: number) {
  const normalizedBase = requireString(base).toLowerCase().replace(/[^a-z0-9._-]/g, "") || `greuser${greUserId}`;
  const existing = new Set(existingUsers.map((user) => requireString(user.username).toLowerCase()).filter(Boolean));
  if (!existing.has(normalizedBase)) return normalizedBase;
  const withId = `${normalizedBase}${greUserId}`;
  if (!existing.has(withId)) return withId;
  let counter = 2;
  while (existing.has(`${normalizedBase}${counter}`)) counter += 1;
  return `${normalizedBase}${counter}`;
}

function findLocalUserForGreProfile(profile: Record<string, unknown>, users: Record<string, unknown>[]) {
  const greUserId = parseNumber(profile.id, 0);
  const email = requireString(profile.Email).toLowerCase();
  const phone = normalizePhoneLogin(profile.Contact || profile.Login);
  const fullName = normalizeComparable(profile.UserName || profile.firstName || "");

  return users.find((user) => {
    if (greUserId > 0 && Number(user.gre_user_id || 0) === greUserId) return true;
    if (email && requireString(user.email).toLowerCase() === email) return true;
    if (phone && normalizePhoneLogin(user.phone) === phone) return true;
    if (fullName && normalizeComparable(user.full_name || user.first_name || user.username || "") === fullName) return true;
    return false;
  }) || null;
}

async function createLocalUserFromGreProfile(profile: Record<string, unknown>, existingUsers: Record<string, unknown>[]) {
  const greRoles = getGreRoleSetFromProfile(profile);
  const role = greRoles.has("admin") ? "admin" : greRoles.has("curator") ? "curator" : "user";
  const greUserId = parseNumber(profile.id, 0);
  const firstName = requireString(profile.firstName) || requireString(profile.UserName).split(/\s+/).filter(Boolean)[0] || "GRE";
  const fullName = requireString(profile.UserName) || firstName;
  const email = requireString(profile.Email).toLowerCase() || `${normalizePhoneLogin(profile.Contact || profile.Login) || `gre${greUserId}`}@greenruraleconomy.in`;
  const phone = normalizePhoneLogin(profile.Contact || profile.Login) || null;
  const usernameBase =
    requireString(profile.firstName) ||
    requireString(profile.UserName).split(/\s+/).filter(Boolean)[0] ||
    email.split("@")[0] ||
    `greuser${greUserId}`;
  const username = buildUniqueUsername(usernameBase, existingUsers, greUserId);

  const { data: userId, error } = await adminClient.rpc("gre_mis_register_user", {
    p_username: username,
    p_first_name: firstName,
    p_full_name: fullName,
    p_email: email,
    p_phone: phone,
    p_password: greTemporaryUserPassword,
  });
  if (error) throw new Error(error.message);

  const { error: passwordError } = await adminClient.rpc("gre_mis_user_set_password", {
    p_user_id: userId,
    p_password: greTemporaryUserPassword,
    p_must_change_password: true,
  });
  if (passwordError) throw new Error(passwordError.message);

  const patch = buildGreMappedPatch(
    {
      id: userId,
      username,
      first_name: firstName,
      full_name: fullName,
      email,
      phone,
      role,
    },
    profile,
  );

  const { data: insertedUser, error: updateError } = await adminClient
    .from("gre_mis_users")
    .update({
      ...patch,
      role,
      must_change_password: true,
      is_active: true,
      updated_at: new Date().toISOString(),
    })
    .eq("id", userId)
    .select("id, username, first_name, full_name, email, phone, role, is_active, must_change_password, last_login_at, created_at, gre_user_id, gre_login_name, gre_sync_status, gre_sync_message, gre_synced_at, gre_pending_role, gre_activation_mod_key, gre_activation_requested_at")
    .single();
  if (updateError || !insertedUser) throw new Error(updateError?.message || "New GRE user could not be loaded into MIS.");

  if (greRoles.has("curator")) {
    await syncCanonicalCuratorRowForUser(
      insertedUser as Record<string, unknown>,
      greRoles.has("admin")
        ? "Created from GRE refresh with combined admin and curator access."
        : "Created from GRE refresh with curator access.",
    );
  }

  return insertedUser as Record<string, unknown>;
}

async function ensureMissingGreDirectoryUsers(users: Record<string, unknown>[], profiles: Record<string, unknown>[]) {
  const localUsers = [...users];
  let createdAny = false;

  for (const profile of profiles) {
    const greRoles = getGreRoleSetFromProfile(profile);
    if (!greRoles.has("curator") && !greRoles.has("admin")) continue;
    if (findLocalUserForGreProfile(profile, localUsers)) continue;
    const inserted = await createLocalUserFromGreProfile(profile, localUsers);
    localUsers.push(inserted);
    createdAny = true;
  }

  if (!createdAny) return localUsers;

  const { data: refreshedUsers, error } = await adminClient
    .from("gre_mis_users")
    .select("id, username, first_name, full_name, email, phone, role, is_active, must_change_password, last_login_at, created_at, gre_user_id, gre_login_name, gre_sync_status, gre_sync_message, gre_synced_at, gre_pending_role, gre_activation_mod_key, gre_activation_requested_at")
    .order("role", { ascending: true })
    .order("first_name", { ascending: true });
  if (error) throw new Error(error.message);
  return (refreshedUsers || []) as Record<string, unknown>[];
}

async function updateUserGreSyncFields(
  userId: string,
  patch: {
    gre_user_id?: number | null;
    gre_login_name?: string | null;
    gre_sync_status: string;
    gre_sync_message: string;
    gre_synced_at?: string | null;
    gre_pending_role?: string | null;
    gre_activation_mod_key?: string | null;
    gre_activation_requested_at?: string | null;
  },
) {
  const { error } = await adminClient.from("gre_mis_users").update({
    ...patch,
    updated_at: new Date().toISOString(),
  }).eq("id", userId);
  if (error) throw new Error(error.message);
}

async function resolveExistingGreUser(user: Record<string, unknown>): Promise<GreUserResolution> {
  if (
    requireString(user.gre_sync_status) === "awaiting_gre_activation" &&
    requireString(user.gre_pending_role) &&
    requireString(user.gre_activation_mod_key)
  ) {
    return {
      status: "awaiting_gre_activation",
      greUserId: parseNumber(user.gre_user_id, 0) || null,
      greLoginName: requireString(user.gre_login_name) || normalizePhoneLogin(user.phone) || requireString(user.email).toLowerCase(),
      message: requireString(user.gre_sync_message) || "GRE activation is pending OTP completion.",
    };
  }

  const storedGreUserId = Number(user.gre_user_id || 0);
  const directoryProfiles = await fetchGreTenantUsers();
  if (storedGreUserId > 0) {
    const liveProfile = directoryProfiles.find((profile) => parseNumber(profile.id, 0) === storedGreUserId) || null;
    if (liveProfile) {
      return {
        status: "mapped",
        greUserId: storedGreUserId,
        greLoginName: requireString(liveProfile.Login) || requireString(liveProfile.Contact) || requireString(user.gre_login_name) || normalizePhoneLogin(user.phone) || requireString(user.email).toLowerCase(),
        profile: liveProfile,
        message: "Using live GRE directory mapping.",
      };
    }
  }

  if (storedGreUserId > 0) {
    return {
      status: "mapped",
      greUserId: storedGreUserId,
      greLoginName: requireString(user.gre_login_name) || normalizePhoneLogin(user.phone) || requireString(user.email).toLowerCase(),
      message: "Using stored GRE user mapping.",
    };
  }

  const directoryProfile = findGreDirectoryProfile(user, directoryProfiles);
  if (directoryProfile) {
    const greUserId = parseNumber(directoryProfile.id, 0);
    const greLoginName = requireString(directoryProfile.Login) || requireString(directoryProfile.Contact) || getGreLoginCandidates(user)[0] || "";
    if (greUserId > 0) {
      return {
        status: "mapped",
        greUserId,
        greLoginName,
        profile: directoryProfile,
        message: `Resolved GRE user mapping for ${greLoginName}.`,
      };
    }
  }

  const candidates = getGreLoginCandidates(user);
  for (const candidate of candidates) {
    const state = await checkGreIdentity(candidate);
    if (state === "ACTIVE_USER" || state === "INACTIVE_USER") {
      return {
        status: "needs_mapping",
        greUserId: null,
        greLoginName: candidate,
        message: `A GRE account exists for ${candidate}, but its GRE user id is not mapped in MIS yet.`,
      };
    }
  }

  return {
    status: "needs_gre_account_creation",
    greUserId: null,
    greLoginName: candidates[0] || "",
    message: "No existing GRE user could be resolved for this MIS user. GRE account creation is still pending.",
  };
}

function buildGreUserCreatePayload(user: Record<string, unknown>, targetRole: string) {
  const loginName = normalizePhoneLogin(user.phone);
  const email = requireString(user.email).toLowerCase();
  const { firstName, lastName } = getUserNameParts(user);

  if (loginName.length < 10) {
    throw new Error("A valid phone number is required before this user can be created on GRE.");
  }
  if (!email) {
    throw new Error("A valid email is required before this user can be created on GRE.");
  }

  return {
    userDTO: {
      firstName,
      lastName,
      login: loginName,
    },
    password: greTemporaryUserPassword,
    userRoleCode: getRoleCodeForMisRole(targetRole),
    functionName: "CREATE",
    userAddressDTOList: [
      {
        primaryAddress: true,
        address: {
          line1: "Green Rural Economy",
          village: "NA",
          country: { id: 1 },
          state: { id: 1 },
          district: { id: 1 },
        },
      },
    ],
    userContactDTOList: [
      {
        primaryContact: true,
        contact: {
          contactType: { dataCode: "CONTACT_TYPE.MOBILE" },
          contactValue: loginName,
        },
      },
      {
        primaryContact: false,
        contact: {
          contactType: { dataCode: "CONTACT_TYPE.MAIL" },
          contactValue: email,
        },
      },
    ],
  };
}

function extractGreActivationModKey(payload: Record<string, unknown> | null, text: string, headers: Headers) {
  const headerValue = requireString(headers.get("modKey") || headers.get("modkey"));
  if (headerValue) return headerValue;
  const payloadValue = requireString(payload?.modKey || payload?.modkey || payload?.data?.modKey || payload?.data?.modkey);
  if (payloadValue) return payloadValue;
  const textMatch = text.match(/"modKey"\s*:\s*"([^"]+)"/i) || text.match(/modKey[:=]\s*([A-Za-z0-9_-]+)/i);
  return textMatch?.[1] || "";
}

function extractGreUserId(payload: Record<string, unknown> | null, text: string) {
  const direct = parseNumber(payload?.userId || payload?.id || payload?.userDTO?.id || payload?.data?.userId || payload?.data?.id || payload?.data?.userDTO?.id, 0);
  if (direct > 0) return direct;
  const textMatch = text.match(/"userId"\s*:\s*(\d+)/i) || text.match(/"id"\s*:\s*(\d+)/i);
  return textMatch ? Number(textMatch[1]) : 0;
}

async function createGreUserForRolePromotion(user: Record<string, unknown>, targetRole: string) {
  const loginName = normalizePhoneLogin(user.phone);
  const payload = buildGreUserCreatePayload(user, targetRole);
  try {
    const result = await callGrePtld("/api/workforce/userWrapper/v1", "POST", payload);
    const modKey = extractGreActivationModKey(result.data, result.text, result.headers);
    const greUserId = extractGreUserId(result.data, result.text);
    return {
      loginName,
      greUserId: greUserId > 0 ? greUserId : null,
      modKey: modKey || null,
      message: requireString(result.data?.message || result.data?.successMessage || result.text || `GRE account created for ${loginName}.`),
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (/already exists/i.test(message)) {
      return {
        loginName,
        greUserId: null,
        modKey: null,
        message,
      };
    }
    throw error;
  }
}

async function activateGreUserWithOtp(greUserId: number, otp: string, modKey: string) {
  const sessionId = await loginToGre();
  const response = await fetch(
    `${greLoginBaseUrl}/commons-iam-service/api/v1/obo/activate-user?userId=${encodeURIComponent(String(greUserId))}`,
    {
      method: "POST",
      headers: {
        "x-sessionid": sessionId,
        modKey,
        otp,
        tenantName: greMasterTenant,
        Accept: "application/json, text/plain, */*",
        Origin: greSiteOrigin,
        Referer: `${greSiteOrigin}/`,
      },
      body: "{}",
    },
  );
  const text = await response.text().catch(() => "");
  if (!response.ok) {
    let parsed: Record<string, unknown> | null = null;
    if (text) {
      try {
        parsed = JSON.parse(text);
      } catch {
        parsed = { message: text };
      }
    }
    throw new Error(parsed?.error?.message || parsed?.errorMessage || parsed?.message || "GRE OTP activation failed.");
  }
  return text;
}

async function activateInactiveGreUser(loginName: string) {
  const { data } = await fetchGreJson(
    `/commons-iam-service/api/v1/obo/activate-inactive-user?tenantName=${encodeURIComponent(greMasterTenant)}&userName=${encodeURIComponent(loginName)}`,
  );
  return data;
}

async function reconcileGreUserMappings(users: Record<string, unknown>[]) {
  const nextUsers: Record<string, unknown>[] = [];
  for (const user of users) {
    const resolution = await resolveExistingGreUser(user);
    if (resolution.status === "mapped" && resolution.greUserId) {
      const patch = resolution.profile
        ? buildGreMappedPatch(user, resolution.profile as Record<string, unknown>)
        : {
          gre_user_id: resolution.greUserId,
          gre_login_name: resolution.greLoginName || null,
          gre_sync_status: "synced",
          gre_sync_message: requireString(user.gre_sync_message) || "GRE mapping is stored and ready for role sync.",
          gre_synced_at: new Date().toISOString(),
          gre_pending_role: null,
          gre_activation_mod_key: null,
          gre_activation_requested_at: null,
        };
      const needsUpdate =
        Number(user.gre_user_id || 0) !== Number(patch.gre_user_id || 0) ||
        requireString(user.gre_login_name) !== requireString(patch.gre_login_name) ||
        requireString(user.gre_sync_status) !== requireString(patch.gre_sync_status) ||
        (patch.email && requireString(user.email).toLowerCase() !== requireString(patch.email).toLowerCase()) ||
        (patch.phone && requireString(user.phone) !== requireString(patch.phone)) ||
        (patch.full_name && requireString(user.full_name) !== requireString(patch.full_name)) ||
        (patch.first_name && requireString(user.first_name) !== requireString(patch.first_name));
      if (needsUpdate) {
        const { error } = await adminClient.from("gre_mis_users").update({
          ...patch,
          updated_at: new Date().toISOString(),
        }).eq("id", requireString(user.id));
        if (error) throw new Error(error.message);
      }
        nextUsers.push({
          ...user,
          ...patch,
          __gre_roles: resolution.profile ? [...getGreRoleSetFromProfile(resolution.profile as Record<string, unknown>)].join(",") : "",
        });
        continue;
      }

    if (resolution.status !== requireString(user.gre_sync_status)) {
      await updateUserGreSyncFields(requireString(user.id), {
        gre_user_id: null,
        gre_login_name: resolution.greLoginName || null,
        gre_sync_status: resolution.status,
        gre_sync_message: resolution.message,
        gre_synced_at: null,
      });
    }
    nextUsers.push({
      ...user,
      gre_user_id: null,
      gre_login_name: resolution.greLoginName || null,
      gre_sync_status: resolution.status,
      gre_sync_message: resolution.message,
      gre_synced_at: null,
    });
  }
  return nextUsers;
}

async function addGreRole(greUserId: number, roleCode: string) {
  await callGreCtld(
    `/api/tenant/user/role/v3?roleCode=${encodeURIComponent(roleCode)}&userId=${encodeURIComponent(String(greUserId))}&returnIfExists=true&removeExistingRoles=false&removeExistingUserRoleHierarchy=false&reason=${encodeURIComponent("role hierarchy")}`,
    "POST",
    {},
  );
}

async function removeGreRole(greUserId: number, roleCode: string) {
  await callGreCtld(
    `/api/tenant/user/role/v1?roleCode=${encodeURIComponent(roleCode)}&userId=${encodeURIComponent(String(greUserId))}`,
    "DELETE",
  );
}

async function syncGreRoleAssignment(greUserId: number, targetRole: string) {
  const targetRoleCode = getRoleCodeForMisRole(targetRole);
  const rolesToRemove = [...new Set(["admin", "moderator", "curator", "user"]
    .filter((role) => role !== targetRole)
    .map((role) => getRoleCodeForMisRole(role)))];

  await addGreRole(greUserId, targetRoleCode);
  for (const roleCode of rolesToRemove) {
    try {
      await removeGreRole(greUserId, roleCode);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      if (!/could not find user role|not found|user.*exists.*userrolemapid/i.test(message)) {
        throw error;
      }
    }
  }
}

async function removeGreTenantRoles(greUserId: number) {
  for (const roleCode of [...new Set(["admin", "moderator", "curator", "user"].map((role) => getRoleCodeForMisRole(role)))]) {
    try {
      await removeGreRole(greUserId, roleCode);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      if (!/could not find user role|not found|user.*exists.*userrolemapid/i.test(message)) {
        throw error;
      }
    }
  }
}

function buildGreWorkforceDeactivationPayload(profile: Record<string, unknown>) {
  const userDTO = (profile.userDTO || {}) as Record<string, unknown>;
  const person = (userDTO.person || {}) as Record<string, unknown>;
  const personProfile = (person.personProfile || {}) as Record<string, unknown>;
  const personContacts = Array.isArray(person.personContacts) ? person.personContacts : [];
  const mobileContact = personContacts.find((entry) => {
    const contact = (entry?.contact || {}) as Record<string, unknown>;
    return requireString((contact.contactType || {})?.dataCode).toUpperCase() === "CONTACT_TYPE.MOBILE";
  }) as Record<string, unknown> | undefined;
  const mailContact = personContacts.find((entry) => {
    const contact = (entry?.contact || {}) as Record<string, unknown>;
    return requireString((contact.contactType || {})?.dataCode).toUpperCase() === "CONTACT_TYPE.MAIL";
  }) as Record<string, unknown> | undefined;

  return [
    {
      ...profile,
      functionName: "UPDATE",
      userDTO: {
        ...userDTO,
        isActive: false,
        person: {
          ...person,
          isActive: false,
          personProfile: {
            ...personProfile,
            isActive: false,
          },
        },
      },
      userAddressDTOList: [
        {
          primaryAddress: true,
          address: {
            line1: "Green Rural Economy",
            village: "NA",
            country: { id: 1 },
            state: { id: 1 },
            district: { id: 1 },
          },
        },
      ],
      userContactDTOList: [
        {
          primaryContact: true,
          contact: {
            contactType: { dataCode: "CONTACT_TYPE.MOBILE" },
            contactValue: requireString(((mobileContact?.contact || {}) as Record<string, unknown>).contactValue || userDTO.login),
          },
        },
        {
          primaryContact: false,
          contact: {
            contactType: { dataCode: "CONTACT_TYPE.MAIL" },
            contactValue: requireString(((mailContact?.contact || {}) as Record<string, unknown>).contactValue),
          },
        },
      ].filter((entry) => requireString((entry.contact as Record<string, unknown>).contactValue)),
    },
  ];
}

async function attemptDeactivateGreWorkforceUser(greUserId: number) {
  const response = await callGrePtld(`/api/workforce/userWrapper/v1?userId=${encodeURIComponent(String(greUserId))}`, "GET");
  const profile = (response.data || {}) as Record<string, unknown>;
  if (!Object.keys(profile).length) {
    throw new Error("GRE workforce profile could not be loaded for complete account removal.");
  }

  const payload = buildGreWorkforceDeactivationPayload(profile);
  try {
    await callGrePtld("/api/workforce/userWrapper/v1", "PUT", payload);
    return {
      ok: true,
      mode: "full_account",
      message: "GRE workforce account was marked inactive.",
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(
      `GRE complete account removal is not yet fully verified for this tenant. The workforce deactivation attempt was rejected by GRE: ${message}`,
    );
  }
}

async function sendGreJson(path: string, method: string, payload: unknown, sessionId?: string) {
  const resolvedSessionId = sessionId || await loginToGre();
  const response = await fetch(`${greLoginBaseUrl}${path}`, {
    method,
    headers: {
      "x-sessionid": resolvedSessionId,
      "Content-Type": "application/json",
      Accept: "application/json, text/plain, */*",
      Origin: greSiteOrigin,
      Referer: `${greSiteOrigin}/`,
    },
    body: JSON.stringify(payload),
  });
  const text = await response.text().catch(() => "");
  let data: Record<string, unknown> | null = null;
  if (text) {
    try {
      data = JSON.parse(text);
    } catch {
      data = { message: text };
    }
  }
  if (!response.ok) {
    throw new Error(
      requireString(data?.error?.message || data?.errorMessage || data?.message || data?.successMessage) ||
      `GRE ${method} failed for ${path}.`,
    );
  }
  return data;
}

async function patchGreJson(path: string, payload: unknown, sessionId?: string) {
  return await sendGreJson(path, "PATCH", payload, sessionId);
}

async function updateGreRequestJson(path: string, payload: unknown, sessionId?: string) {
  try {
    return await sendGreJson(path, "PATCH", payload, sessionId);
  } catch (patchError) {
    const patchMessage = patchError instanceof Error ? patchError.message : String(patchError);
    try {
      return await sendGreJson(path, "PUT", payload, sessionId);
    } catch (putError) {
      const putMessage = putError instanceof Error ? putError.message : String(putError);
      throw new Error(`GRE request update failed. PATCH: ${patchMessage} PUT: ${putMessage}`);
    }
  }
}

async function getGreRefData(classCode: string) {
  if (greRefDataCache.has(classCode)) return greRefDataCache.get(classCode) || [];
  const { data } = await fetchGreJson(
    `/commons-request-management-service/api/v1/ref-data/class/${encodeURIComponent(classCode)}?direction=ASC&fetchContextDataOnly=false&languageCode=ENG&page=0&size=100`,
  );
  const elements = Array.isArray(data?.elements) ? data.elements : [];
  greRefDataCache.set(classCode, elements);
  return elements as Record<string, unknown>[];
}

function normalizeRichText(value: unknown) {
  const text = requireString(value);
  if (!text) return "<p></p>";
  if (/<[a-z][\s\S]*>/i.test(text)) return text;
  return `<p>${text.replace(/\n+/g, "<br/>")}</p>`;
}

function buildGreTextList(text: string) {
  return [{ id: 0, languageCode: "ENG", text: requireString(text) }];
}

function firstArrayObject(value: unknown) {
  return Array.isArray(value) && value.length && value[0] && typeof value[0] === "object"
    ? value[0] as Record<string, unknown>
    : null;
}

function normalizeGreLabel(value: unknown) {
  return normalizeComparable(value);
}

function extractGreEntityName(entity: Record<string, unknown> | null) {
  if (!entity) return "";
  const direct = requireString(entity.text || entity.name || entity.label || entity.displayLabel);
  if (direct) return direct;
  const nameEntry = firstArrayObject(entity.name);
  return requireString(nameEntry?.text);
}

function buildGreSolutionAudienceValues(values: string[]) {
  return values.map((value, index) => ({
    id: 0,
    sequence: index,
    value: buildGreTextList(value),
  }));
}

function buildGreSpecificationEntry(
  template: Record<string, unknown>,
  value: string,
  specificationValues: unknown[] = [],
) {
  return {
    id: 0,
    sequence: parseNumber(template.sequence, 0),
    specificationMaster: parseNumber(template.id, 0),
    SpecificationMasterDTO: {
      id: parseNumber(template.id, 0),
      paramName: requireString(template.paramName),
      specificationCode: requireString(template.specificationCode),
      groupCode: requireString(template.groupCode),
      dataType: requireString(template.dataType),
      isMandatory: Boolean(template.isMandatory),
      isMultiValued: Boolean(template.isMultiValued),
      customRefClass: requireString(template.customRefClass) || undefined,
      value,
      specificationLabel: [],
      specificationDisplayLabel: [],
      classificationList: [],
      isActive: true,
      sequence: parseNumber(template.sequence, 0),
    },
    value,
    specificationValues,
  };
}

function extractGreResponseRecord(data: unknown) {
  if (Array.isArray(data)) return firstArrayObject(data);
  if (data && typeof data === "object") return data as Record<string, unknown>;
  return null;
}

function normalizeGreLocationLabel(item: Record<string, unknown>) {
  const label = requireString(item.displayLabel || item.display_label || item.label || item.addressLabel || item.address_label);
  if (label) return label;
  const textValue = firstArrayObject(item.value);
  if (textValue?.text) return requireString(textValue.text);
  return [
    requireString(item.blockName || item.block_name || item.subDistrictName),
    requireString(item.districtName || item.district_name || item.cityName),
    requireString(item.stateName || item.state_name),
    requireString(item.countryName || item.country_name || "India"),
  ].filter(Boolean).join(", ");
}

function extractGreLocationCode(item: Record<string, unknown>) {
  const locationType = requireString(item.locationType || item.location_type || item.subType || item.sub_type).toUpperCase();
  if (locationType === "CITY") {
    return requireString(item.cityCode || item.city_code || item.code || item.locationCode || item.location_code);
  }
  if (locationType === "STATE") {
    return requireString(item.stateCode || item.state_code || item.code || item.locationCode || item.location_code);
  }
  if (locationType === "COUNTRY") {
    return requireString(item.countryCode || item.country_code || item.code || item.locationCode || item.location_code);
  }
  return requireString(item.code || item.locationCode || item.location_code || item.cityCode || item.stateCode || item.countryCode);
}

function extractGreLocationSubtype(item: Record<string, unknown>) {
  const locationType = requireString(item.locationType || item.location_type);
  if (locationType) return locationType.toUpperCase();
  return requireString(item.subType || item.sub_type || item.locationKind || item.location_kind || item.type).toUpperCase();
}

function extractGreTagEntry(entity: Record<string, unknown> | null, sequence: number) {
  if (!entity) return null;
  const tagCode = requireString(entity.tagCode || entity.tagIdentifier);
  const tagId = parseNumber(entity.tagId, 0);
  if (!tagCode || !tagId) return null;
  return {
    id: 0,
    tagCode,
    tagId,
    tagSequence: sequence,
    isActive: true,
  };
}

async function fetchGreTagsByCodes(tagCodes: string[], sessionId: string) {
  const uniqueCodes = uniqueStrings(tagCodes.map((item) => requireString(item)).filter(Boolean));
  if (!uniqueCodes.length) return [];
  const query = uniqueCodes.join(",");
  const { data } = await requestGreGatewayJson(
    `/commons-domain-service/api/v1/tag/codes?codes=${encodeURIComponent(query)}`,
    "GET",
    undefined,
    sessionId,
  );
  if (Array.isArray(data)) return data as Record<string, unknown>[];
  if (Array.isArray((data as Record<string, unknown>)?.value)) {
    return (data as Record<string, unknown>).value as Record<string, unknown>[];
  }
  if (Array.isArray((data as Record<string, unknown>)?.data)) {
    return (data as Record<string, unknown>).data as Record<string, unknown>[];
  }
  if (data && typeof data === "object") {
    return [data as Record<string, unknown>];
  }
  return [];
}

async function buildGreTagEntriesFromCodes(tagCodes: string[], sessionId: string) {
  const tags = await fetchGreTagsByCodes(tagCodes, sessionId);
  return uniqueStrings(tagCodes).map((tagCode, index) => {
    const matched = tags.find((item) =>
      requireString(item.code || item.tagCode || item.identifier) === tagCode
    );
    if (!matched) return null;
    const tagId = parseNumber(matched.id || matched.tagId, 0);
    if (!tagId) return null;
    return {
      id: 0,
      tagCode,
      tagId,
      tagSequence: index,
      isActive: true,
    };
  }).filter(Boolean) as Record<string, unknown>[];
}

function buildGreFallbackVariety(applicationId: string, applicationName: string) {
  return {
    id: parseNumber(applicationId, 0),
    varietyCode: `GRE_MIS.${applicationId || "VARIETY"}`,
    name: buildGreTextList(applicationName || "Application"),
    description: buildGreTextList(""),
    tagIdentifier: "",
  };
}

function buildGreTagIdentifierFromCode(code: string) {
  const normalized = requireString(code);
  return normalized ? `TAG.${normalized}` : "";
}

function buildGreVarietyTagIdentifier(genericProductCode: string, applicationName: string) {
  const base = requireString(genericProductCode);
  const suffix = requireString(applicationName)
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
  return base && suffix ? `TAG.${base}.${suffix}` : "";
}

function buildGreNameMl(text: string) {
  return buildGreTextList(text);
}

function prettifyGreCodeTail(code: string) {
  return requireString(code)
    .split(".")
    .pop()
    ?.replace(/_/g, " ")
    .toLowerCase()
    .replace(/\b\w/g, (character) => character.toUpperCase()) || "";
}

function buildGreClassificationStub(tagMeta: Record<string, unknown> | null) {
  if (!tagMeta) return [];
  const subCategoryCode = requireString(tagMeta.subCategoryCode);
  const context = requireString(tagMeta.context) || "COMMONS.GRE";
  if (!subCategoryCode) return [];
  return [
    {
      id: 241,
      code: subCategoryCode,
      context,
      name: buildGreNameMl(prettifyGreCodeTail(subCategoryCode)),
      description: [],
      productType: "VALUE_CHAIN",
      iconPic: "",
      logoURL: "",
    },
  ];
}

async function resolveGreHierarchyMatch(
  primaryValuechain: string,
  primaryApplication: string,
) {
  const normalizedValuechain = normalizeComparable(primaryValuechain);
  const normalizedApplication = normalizeComparable(primaryApplication);
  if (!normalizedValuechain) throw new Error("Primary Value Chain is required.");
  if (!normalizedApplication) throw new Error("Primary Application is required.");

  const { data, error } = await adminClient
    .from("offerings")
    .select("primary_valuechain_id, primary_valuechain, primary_application_id, primary_application")
    .ilike("primary_valuechain", `%${primaryValuechain}%`)
    .limit(500);
  if (error) throw new Error(error.message);

  const rows = Array.isArray(data) ? data as Record<string, unknown>[] : [];
  const exactApplication = rows.find((row) =>
    normalizeComparable(row.primary_valuechain) === normalizedValuechain &&
    normalizeComparable(row.primary_application) === normalizedApplication &&
    requireString(row.primary_valuechain_id) &&
    requireString(row.primary_application_id)
  );
  if (exactApplication) return exactApplication;

  const exactValuechain = rows.find((row) =>
    normalizeComparable(row.primary_valuechain) === normalizedValuechain &&
    requireString(row.primary_valuechain_id)
  );
  if (exactValuechain && normalizedApplication === normalizeComparable(exactValuechain.primary_application)) return exactValuechain;

  throw new Error(`Could not map GRE hierarchy for "${primaryValuechain}" -> "${primaryApplication}".`);
}

async function fetchGreGenericProductDetail(genericProductId: string, sessionId: string) {
  const { data } = await requestGreGatewayJson(
    `/commons-product-service/api/v1/generic-product/detail?genericProductIds=${encodeURIComponent(genericProductId)}`,
    "GET",
    undefined,
    sessionId,
  );
  const record = firstArrayObject(data) || extractGreResponseRecord(data);
  if (!record) throw new Error(`GRE generic product detail was not found for id ${genericProductId}.`);
  return record;
}

async function fetchGreGenericProductApplications(genericProductId: string, sessionId: string) {
  const { data } = await requestGreGatewayJson(
    `/commons-product-service/api/v1/solution/generic-product?genericProductId=${encodeURIComponent(genericProductId)}&marketId=${encodeURIComponent(greMarketId)}&page=0&size=200`,
    "GET",
    undefined,
    sessionId,
  );
  if (Array.isArray((data as Record<string, unknown>)?.elements)) {
    return (data as Record<string, unknown>).elements as Record<string, unknown>[];
  }
  if (Array.isArray(data)) return data as Record<string, unknown>[];
  return [];
}

function resolveGreVarietyFromProductDetail(
  genericProduct: Record<string, unknown>,
  applicationId: string,
  applicationName: string,
) {
  const varieties = Array.isArray(genericProduct.genericProductVarietyList)
    ? genericProduct.genericProductVarietyList as Record<string, unknown>[]
    : [];
  const normalizedApplicationName = normalizeComparable(applicationName);
  return varieties.find((item) =>
    String(item.id ?? "") === applicationId ||
    normalizeGreLabel(extractGreEntityName(item)) === normalizedApplicationName
  ) || null;
}

async function resolveGreSolutionHierarchy(payload: Record<string, unknown>, sessionId: string) {
  const primaryValuechain = requireString(payload.primary_valuechain);
  const primaryApplication = requireString(payload.primary_application);
  const secondaryValuechain = requireString(payload.secondary_valuechain);
  const secondaryApplication = requireString(payload.secondary_application);

  const primaryMatch = await resolveGreHierarchyMatch(primaryValuechain, primaryApplication);
  const genericProductId = requireString(primaryMatch.primary_valuechain_id);
  const primaryApplicationId = requireString(primaryMatch.primary_application_id);
  const genericProduct = await fetchGreGenericProductDetail(genericProductId, sessionId);
  const primaryTagCode = buildGreTagIdentifierFromCode(requireString(genericProduct.genericProductCode || genericProduct.code));
  const primaryVarietyTagCode = buildGreVarietyTagIdentifier(
    requireString(genericProduct.genericProductCode || genericProduct.code),
    primaryApplication,
  );
  const primaryTagRecords = await fetchGreTagsByCodes(
    [primaryTagCode, primaryVarietyTagCode].filter(Boolean),
    sessionId,
  );
  const primaryProductTag = primaryTagRecords.find((item) => requireString(item.code) === primaryTagCode) || null;
  const primaryVarietyTag = primaryTagRecords.find((item) => requireString(item.code) === primaryVarietyTagCode) || null;
  if (!requireString(genericProduct.id)) {
    genericProduct.id = parseNumber(genericProduct.genericProductId, 0) || parseNumber(genericProductId, 0);
  }
  if (!requireString(genericProduct.tagIdentifier) && requireString(genericProduct.genericProductCode || genericProduct.code)) {
    genericProduct.tagIdentifier = buildGreTagIdentifierFromCode(requireString(genericProduct.genericProductCode || genericProduct.code));
  }
  if (!requireString(genericProduct.code) && requireString(genericProduct.genericProductCode)) {
    genericProduct.code = requireString(genericProduct.genericProductCode);
  }
  if (!Array.isArray(genericProduct.name) && Array.isArray(genericProduct.genericProductName)) {
    genericProduct.name = genericProduct.genericProductName;
  }
  if (!Array.isArray(genericProduct.name)) {
    genericProduct.name = buildGreNameMl(primaryValuechain);
  }
  if (!Array.isArray(genericProduct.description)) {
    genericProduct.description = Array.isArray(genericProduct.genericProductDescription) ? genericProduct.genericProductDescription : [];
  }
  if (!Array.isArray(genericProduct.classificationList) || !genericProduct.classificationList.length) {
    genericProduct.classificationList = buildGreClassificationStub(primaryProductTag);
  }
  if (genericProduct.marketId === undefined) genericProduct.marketId = parseNumber(greMarketId, 0);
  if (genericProduct.isPrivate === undefined) genericProduct.isPrivate = false;
  if (!requireString(genericProduct.type)) genericProduct.type = "VALUE_CHAIN";
  const primaryApplications = await fetchGreGenericProductApplications(genericProductId, sessionId);
  const genericProductVariety = resolveGreVarietyFromProductDetail(genericProduct, primaryApplicationId, primaryApplication)
    || primaryApplications.find((item) =>
      String(item.id ?? "") === primaryApplicationId ||
      normalizeGreLabel(extractGreEntityName(item)) === normalizeComparable(primaryApplication)
    );
  const resolvedPrimaryVariety = genericProductVariety || buildGreFallbackVariety(primaryApplicationId, primaryApplication);
  if (!requireString(resolvedPrimaryVariety.id)) {
    resolvedPrimaryVariety.id = parseNumber(primaryApplicationId, 0);
  }
  if (!requireString(resolvedPrimaryVariety.tagIdentifier)) {
    resolvedPrimaryVariety.tagIdentifier = buildGreVarietyTagIdentifier(
      requireString(genericProduct.code || genericProduct.genericProductCode),
      primaryApplication,
    );
  }
  if (!requireString(resolvedPrimaryVariety.varietyCode)) {
    resolvedPrimaryVariety.varietyCode = requireString(resolvedPrimaryVariety.tagIdentifier).replace(/^TAG\./, "") || `GRE_MIS.${primaryApplicationId}`;
  }
  if (!Array.isArray(resolvedPrimaryVariety.name)) {
    resolvedPrimaryVariety.name = buildGreNameMl(primaryApplication);
  }
  if (!Array.isArray(resolvedPrimaryVariety.description)) {
    resolvedPrimaryVariety.description = [];
  }
  if (!Array.isArray(genericProduct.genericProductVarietyList) || !genericProduct.genericProductVarietyList.length) {
    genericProduct.genericProductVarietyList = [resolvedPrimaryVariety];
  }

  const tagCodes = [
    requireString(genericProduct.tagIdentifier),
    requireString(resolvedPrimaryVariety?.tagIdentifier),
  ].filter(Boolean);

  if (secondaryValuechain && secondaryApplication) {
    try {
      const secondaryMatch = await resolveGreHierarchyMatch(secondaryValuechain, secondaryApplication);
      const secondaryGenericProduct = await fetchGreGenericProductDetail(requireString(secondaryMatch.primary_valuechain_id), sessionId);
      const secondaryTagCode = buildGreTagIdentifierFromCode(requireString(secondaryGenericProduct.genericProductCode || secondaryGenericProduct.code));
      const secondaryVarietyTagCode = buildGreVarietyTagIdentifier(
        requireString(secondaryGenericProduct.genericProductCode || secondaryGenericProduct.code),
        secondaryApplication,
      );
      if (!requireString(secondaryGenericProduct.tagIdentifier) && requireString(secondaryGenericProduct.genericProductCode || secondaryGenericProduct.code)) {
        secondaryGenericProduct.tagIdentifier = buildGreTagIdentifierFromCode(requireString(secondaryGenericProduct.genericProductCode || secondaryGenericProduct.code));
      }
      if (!requireString(secondaryGenericProduct.id)) {
        secondaryGenericProduct.id = parseNumber(secondaryGenericProduct.genericProductId, 0) || parseNumber(secondaryMatch.primary_valuechain_id, 0);
      }
      if (!requireString(secondaryGenericProduct.code) && requireString(secondaryGenericProduct.genericProductCode)) {
        secondaryGenericProduct.code = requireString(secondaryGenericProduct.genericProductCode);
      }
      if (!Array.isArray(secondaryGenericProduct.name) && Array.isArray(secondaryGenericProduct.genericProductName)) {
        secondaryGenericProduct.name = secondaryGenericProduct.genericProductName;
      }
      if (!Array.isArray(secondaryGenericProduct.name)) {
        secondaryGenericProduct.name = buildGreNameMl(secondaryValuechain);
      }
      if (!Array.isArray(secondaryGenericProduct.description)) {
        secondaryGenericProduct.description = Array.isArray(secondaryGenericProduct.genericProductDescription) ? secondaryGenericProduct.genericProductDescription : [];
      }
      if (secondaryGenericProduct.marketId === undefined) secondaryGenericProduct.marketId = parseNumber(greMarketId, 0);
      if (secondaryGenericProduct.isPrivate === undefined) secondaryGenericProduct.isPrivate = false;
      if (!requireString(secondaryGenericProduct.type)) secondaryGenericProduct.type = "VALUE_CHAIN";
      const secondaryTagRecords = await fetchGreTagsByCodes(
        [secondaryTagCode, secondaryVarietyTagCode].filter(Boolean),
        sessionId,
      );
      const secondaryProductTag = secondaryTagRecords.find((item) => requireString(item.code) === secondaryTagCode) || null;
      if (!Array.isArray(secondaryGenericProduct.classificationList) || !secondaryGenericProduct.classificationList.length) {
        secondaryGenericProduct.classificationList = buildGreClassificationStub(secondaryProductTag);
      }
      const secondaryApplications = await fetchGreGenericProductApplications(requireString(secondaryMatch.primary_valuechain_id), sessionId);
      const secondaryVariety = resolveGreVarietyFromProductDetail(
        secondaryGenericProduct,
        requireString(secondaryMatch.primary_application_id),
        secondaryApplication,
      ) || secondaryApplications.find((item) =>
        String(item.id ?? "") === requireString(secondaryMatch.primary_application_id) ||
        normalizeGreLabel(extractGreEntityName(item)) === normalizeComparable(secondaryApplication)
      ) || null;
      if (secondaryVariety && !requireString(secondaryVariety.tagIdentifier)) {
        secondaryVariety.tagIdentifier = buildGreVarietyTagIdentifier(
          requireString(secondaryGenericProduct.code || secondaryGenericProduct.genericProductCode),
          secondaryApplication,
        );
      }
      if (secondaryVariety) {
        if (!requireString(secondaryVariety.id)) secondaryVariety.id = parseNumber(secondaryMatch.primary_application_id, 0);
        if (!requireString(secondaryVariety.varietyCode)) {
          secondaryVariety.varietyCode = requireString(secondaryVariety.tagIdentifier).replace(/^TAG\./, "") || `GRE_MIS.${requireString(secondaryMatch.primary_application_id)}`;
        }
        if (!Array.isArray(secondaryVariety.name)) secondaryVariety.name = buildGreNameMl(secondaryApplication);
        if (!Array.isArray(secondaryVariety.description)) secondaryVariety.description = [];
      }
      [secondaryGenericProduct, secondaryVariety].forEach((entry) => {
        const code = requireString(entry?.tagIdentifier);
        if (code) tagCodes.push(code);
      });
    } catch {
      // Secondary hierarchy is optional for write-back.
    }
  }

  const solutionTagList = await buildGreTagEntriesFromCodes(tagCodes, sessionId);

  return {
    genericProductId,
    primaryApplicationId,
    genericProduct,
    genericProductVariety: resolvedPrimaryVariety,
    solutionTagList,
  };
}

async function resolveGreOfferingTypeAndSubType(payload: Record<string, unknown>, sessionId: string) {
  const offeringCategory = requireString(payload.offering_category).toLowerCase();
  if (offeringCategory !== "service offerings") {
    throw new Error("GRE write-back is currently verified only for Service offerings. Product and Knowledge create flows still need explicit HAR capture.");
  }
  const offeringGroupCode = "OFFERINGS.SERVICE_OFFERINGS";
  const subTypes = await fetchGreProductRefHierarchy(offeringGroupCode, sessionId);
  const target = requireString(payload.offering_type);
  const normalizedTarget = normalizeComparable(target);
  const matched = findGreRefDataOption(subTypes, target, [
    target.replace(/\//g, " "),
    target.replace(/\s*&\s*/g, " "),
    target.replace(/\s+/g, " "),
  ]);
  if (!matched) {
    const fallbackCode = greServiceOfferingSubtypeFallbacks[normalizedTarget];
    if (fallbackCode) {
      return {
        offeringGroupCode,
        offeringSubTypeCode: fallbackCode,
      };
    }
    throw new Error(`Could not map GRE offering type for "${target}".`);
  }
  return {
    offeringGroupCode,
    offeringSubTypeCode: requireString(matched.code),
  };
}

async function resolveGreChipValues(
  classCode: string,
  labels: string[],
  sessionId: string,
  aliasMap: Record<string, string[]> = {},
) {
  const options = await fetchGreProductRefData(classCode, sessionId);
  return labels.map((label, index) => {
    const normalizedLabel = normalizeComparable(label);
    const fallbackCode = classCode === "CLASS.TRAINING_LANGUAGE"
      ? greTrainingLanguageFallbacks[normalizedLabel]
      : "";
    let matched = findGreRefDataOption(options, label, aliasMap[label] || aliasMap[normalizedLabel] || []);
    if (!matched && fallbackCode) {
      matched = options.find((option) => requireString(option.code) === fallbackCode) || null;
    }
    if (!matched && fallbackCode) {
      return {
        code: fallbackCode,
        id: 0,
        sequence: index,
      };
    }
    if (!matched) throw new Error(`Could not map GRE option "${label}" for ${classCode}.`);
    return {
      code: requireString(matched.code),
      id: 0,
      sequence: index,
    };
  });
}

async function resolveGreLocationValues(geographies: string[], sessionId: string) {
  const resolved: Record<string, unknown>[] = [];
  for (const geography of geographies) {
    const query = requireString(geography);
    if (!query) continue;
    const { data } = await requestGreGatewayJson(
      `/commons-search-service/api/v1/commons-location/search?keyword=${encodeURIComponent(query)}&page=0&size=10&cityData=true`,
      "GET",
      undefined,
      sessionId,
    );
    const elements = Array.isArray((data as Record<string, unknown>)?.elements)
      ? ((data as Record<string, unknown>).elements as Record<string, unknown>[])
      : Array.isArray(data)
        ? data as Record<string, unknown>[]
        : [];
    const match = elements.find((item) => normalizeComparable(normalizeGreLocationLabel(item)) === normalizeComparable(query))
      || elements.find((item) => normalizeComparable(normalizeGreLocationLabel(item)).includes(normalizeComparable(query)))
      || elements[0];
    if (!match) throw new Error(`Could not map GRE geography for "${query}".`);
    const code = extractGreLocationCode(match);
    const subType = extractGreLocationSubtype(match) || "DISTRICT";
    const label = normalizeGreLocationLabel(match) || query;
    if (!code) throw new Error(`GRE geography code is missing for "${query}".`);
    resolved.push({
      code,
      id: 0,
      subType,
      type: "LOCATION",
      value: buildGreTextList(label),
    });
  }
  return resolved;
}

function buildGreSolutionCreatePayload(
  payload: Record<string, unknown>,
  hierarchy: {
    genericProduct: Record<string, unknown>;
    genericProductVariety: Record<string, unknown>;
    solutionTagList: Record<string, unknown>[];
  },
  traderId: string,
) {
  const title = requireString(payload.solution_name);
  const description = requireString(payload.about_solution_html || payload.about_solution_text);
  const audience = asStringArray(payload.solution_audience);
  if (!title) throw new Error("Solution Name is required.");
  if (!description) throw new Error("Solution Description is required.");
  if (!audience.length) throw new Error("At least one 'Who Can avail it' choice is required.");

  return {
    channelId: greChannelId,
    marketId: greMarketId,
    traderId: Number(traderId),
    status: "TRADER_CENTER_SOLUTION_STATUS.DRAFT",
    solutionDTO: {
      id: 0,
      title: buildGreTextList(title),
      description: buildGreTextList(description),
      isGenericSolution: false,
      statusUpdatedOnDate: Date.now(),
      genericProduct: hierarchy.genericProduct,
      genericProductVariety: hierarchy.genericProductVariety,
      solutionTagList: hierarchy.solutionTagList,
      marketId: greMarketId,
      solutionSpecificationList: [
        {
          id: 0,
          sequence: 1,
          specificationMaster: {
            id: 71,
            paramName: "Who Can avail it?",
            specificationCode: "SOLUTION_WHO_CAN_AVAIL",
            customRefClass: "CLASS.COMMUNITY_ENTITIES",
            groupCode: "SPECIFICATION_GROUP.SOLUTION",
            dataType: "CHIP",
            isMandatory: true,
            isMultiValued: true,
            specificationLabel: [],
            specificationDisplayLabel: [],
            classificationList: [],
            isActive: true,
            sequence: 1,
          },
          specificationValues: buildGreSolutionAudienceValues(audience),
        },
      ],
      status: "SOLUTION_STATUS.PENDING_REVIEW",
      solutionClass: "SOLUTION_CLASS.VALUE_CHAIN",
    },
  };
}

async function buildGreServicePublishPayload(
  payload: Record<string, unknown>,
  solutionId: number,
  traderId: string,
  hierarchy: {
    genericProductId: string;
    primaryApplicationId: string;
    solutionTagList: Record<string, unknown>[];
  },
  offeringType: {
    offeringGroupCode: string;
    offeringSubTypeCode: string;
  },
  sessionId: string,
) {
  const offeringName = requireString(payload.offering_name);
  const offeringDescription = normalizeRichText(payload.about_offering_html || payload.about_offering_text);
  const trainerName = requireString(payload.trainer_name);
  const trainerEmail = requireString(payload.trainer_email).toLowerCase();
  const trainerPhone = requireString(payload.trainer_phone);
  const trainerDetails = normalizeRichText(payload.trainer_details_text || payload.trainer_details);
  const languages = normalizeLanguageArray(payload.languages);
  const geographies = asStringArray(payload.geographies);
  const duration = requireString(payload.duration);
  const durationUnit = requireString(payload.duration_unit || "Days");
  const prerequisites = requireString(payload.prerequisites);
  const locationAvailability = asStringArray(payload.location_availability);
  const serviceCost = requireString(payload.service_cost || "0");
  const serviceCostUnit = requireString(payload.service_cost_unit || "Can be quoted after finalising scope");
  const supportPostService = requireString(payload.support_post_service);
  const supportPostServiceCost = requireString(payload.support_post_service_cost);
  const deliveryMode = requireString(payload.delivery_mode);
  const certificationOffered = requireString(payload.certification_offered || "Not Provided");
  const costRemarks = requireString(payload.cost_remarks);

  if (!offeringName) throw new Error("Offering Name is required.");
  if (!trainerName || !trainerEmail || !trainerPhone) throw new Error("Trainer name, email, and phone are required.");
  if (!languages.length) throw new Error("At least one language is required.");
  if (!geographies.length) throw new Error("At least one geography is required.");
  if (!duration) throw new Error("Duration is required.");
  if (!prerequisites) throw new Error("Prerequisites are required.");
  if (!locationAvailability.length) throw new Error("At least one location availability choice is required.");
  if (!supportPostService || !supportPostServiceCost || !deliveryMode) {
    throw new Error("Support post service, support post service cost, and delivery mode are required.");
  }

  const languageValues = await resolveGreChipValues("CLASS.TRAINING_LANGUAGE", languages, sessionId);
  const locationValues = await resolveGreChipValues(
    "CLASS.SERVICE_LOCATION_AVAILABILITY",
    locationAvailability,
    sessionId,
    {
      "at service provider": ["provider", "provider end"],
      "at service seeker": ["seeker", "seeker end"],
      "others": ["other"],
    },
  );
  const geographyValues = await resolveGreLocationValues(geographies, sessionId);
  const supportPostServiceOption = findGreRefDataOption(
    await fetchGreProductRefData("CLASS.POST_SERVICE", sessionId),
    supportPostService,
  );
  const supportPostServiceCostOption = findGreRefDataOption(
    await fetchGreProductRefData("CLASS.POST_SERVICE_COST", sessionId),
    supportPostServiceCost,
  );
  const deliveryModeOption = findGreRefDataOption(
    await fetchGreProductRefData("CLASS.SESSION_AVAILABILITY", sessionId),
    deliveryMode,
  );
  const certificationOption = findGreRefDataOption(
    await fetchGreProductRefData("CLASS.CERTIFICATION", sessionId),
    certificationOffered,
  );
  if (!supportPostServiceOption || !supportPostServiceCostOption || !deliveryModeOption || !certificationOption) {
    throw new Error("One or more GRE service dropdown values could not be mapped.");
  }

  const generatedSuffix = Date.now();
  const productCode = `MARKIFY_PRODUCT._${generatedSuffix}`;
  const skuCode = `MARKIFY_SKU._${generatedSuffix}`;
  const skuSpecs = [
    buildGreSpecificationEntry(greServiceSpecTemplate[0], trainerName),
    buildGreSpecificationEntry(greServiceSpecTemplate[1], trainerEmail),
    buildGreSpecificationEntry(greServiceSpecTemplate[2], "", languageValues),
    buildGreSpecificationEntry(greServiceSpecTemplate[3], "", geographyValues),
    buildGreSpecificationEntry(greServiceSpecTemplate[4], trainerPhone),
    buildGreSpecificationEntry(greServiceSpecTemplate[5], trainerDetails),
    buildGreSpecificationEntry(greServiceSpecTemplate[6], `${duration} ${durationUnit}`.trim()),
    buildGreSpecificationEntry(greServiceSpecTemplate[7], prerequisites),
    buildGreSpecificationEntry(greServiceSpecTemplate[8], "", locationValues),
    buildGreSpecificationEntry(greServiceSpecTemplate[9], `${serviceCost} ${serviceCostUnit}`.trim()),
    buildGreSpecificationEntry(greServiceSpecTemplate[10], costRemarks),
    buildGreSpecificationEntry(greServiceSpecTemplate[11], requireString(supportPostServiceOption.value || supportPostServiceOption.name)),
    buildGreSpecificationEntry(greServiceSpecTemplate[12], requireString(supportPostServiceCostOption.value || supportPostServiceCostOption.name)),
    buildGreSpecificationEntry(greServiceSpecTemplate[13], requireString(deliveryModeOption.value || deliveryModeOption.name)),
    buildGreSpecificationEntry(greServiceSpecTemplate[14], requireString(certificationOption.value || certificationOption.name)),
    buildGreSpecificationEntry(greServiceSpecTemplate[15], "", []),
  ];

  return {
    customProductDTOS: [
      {
        channelId: greChannelId,
        code: productCode,
        customProductSKUDTOs: [
          {
            code: skuCode,
            description: buildGreTextList(offeringDescription),
            id: 0,
            isDefaultSKU: true,
            name: buildGreTextList(offeringName),
            productSkuKind: "PRODUCT_SKU_KIND.PACKAGED",
            discountPercentage: 0,
            mrp: {
              currency: "INR",
              id: 0,
              value: parseNumber(serviceCost, 0).toString(),
            },
            quantity: {
              id: 0,
              uoM: "Units",
              value: 0,
            },
            productSKUTagList: hierarchy.solutionTagList,
            productSkuSpecificationList: skuSpecs,
            skuSubType: offeringType.offeringSubTypeCode,
            skuType: offeringType.offeringGroupCode,
          },
        ],
        description: buildGreTextList(offeringDescription),
        genericProductId: parseNumber(hierarchy.genericProductId, 0),
        id: 0,
        isMarketChannelProduct: true,
        name: buildGreTextList(offeringName),
        productKind: "PRODUCT_KIND.PACKAGED",
        solutionId,
        traderId: Number(traderId),
        subType: offeringType.offeringSubTypeCode,
        type: offeringType.offeringGroupCode,
        varietyId: parseNumber(hierarchy.primaryApplicationId, 0),
      },
    ],
    solutionId,
    marketId: greMarketId,
    traderId: Number(traderId),
    channelId: greChannelId,
  };
}

async function syncApprovedSolutionToGre(payload: Record<string, unknown>) {
  const sessionId = await loginToGre();
  const traderId = requireString(payload.existing_trader_id);
  if (!traderId) throw new Error("Approved GRE supplier id is required for solution sync.");

  const hierarchy = await resolveGreSolutionHierarchy(payload, sessionId);
  const offeringType = await resolveGreOfferingTypeAndSubType(payload, sessionId);
  const solutionPayload = buildGreSolutionCreatePayload(payload, hierarchy, traderId);
  const { data: solutionData } = await requestGreGatewayJson(
    "/commons-market-service/api/v1/tma-channel-solution",
    "POST",
    solutionPayload,
    sessionId,
  );
  const solutionRecord = extractGreResponseRecord(solutionData);
  const solutionId = parseNumber(solutionRecord?.id || (solutionRecord?.solutionDTO as Record<string, unknown>)?.id, 0);
  if (!solutionId) throw new Error("GRE solution create did not return a solution id.");

  const publishPayload = await buildGreServicePublishPayload(payload, solutionId, traderId, hierarchy, offeringType, sessionId);
  const { data: publishData } = await requestGreGatewayJson(
    "/commons-market-service/api/v2/trader-center-product/publish-solution-product",
    "POST",
    publishPayload,
    sessionId,
  );
  const publishRecord = extractGreResponseRecord(publishData);
  const firstProduct = Array.isArray(publishRecord?.customProductDTOS)
    ? firstArrayObject(publishRecord.customProductDTOS)
    : Array.isArray(publishRecord?.customProductDTOs)
      ? firstArrayObject(publishRecord.customProductDTOs)
      : publishRecord?.customProductDTO && typeof publishRecord.customProductDTO === "object"
        ? publishRecord.customProductDTO as Record<string, unknown>
        : null;
  const firstSku = firstArrayObject((firstProduct?.customProductSKUDTOs as unknown[]) || []);

  return {
    solutionId,
    offeringId: parseNumber(firstProduct?.id, 0) || null,
    skuId: parseNumber(firstSku?.id, 0) || null,
  };
}

function findGreRefDataOption(
  options: Record<string, unknown>[],
  value: string,
  aliases: string[] = [],
) {
  const normalized = normalizeComparable(value);
  const probes = [normalized, ...aliases.map((alias) => normalizeComparable(alias))].filter(Boolean);
  return options.find((option) => {
    const names = [
      option?.name,
      option?.code,
      ...(Array.isArray(option?.labels) ? option.labels.map((label: { text?: string }) => label?.text || "") : []),
    ].map((entry) => normalizeComparable(entry));
    return probes.some((probe) => names.includes(probe));
  }) || null;
}

function buildGreDateString(value: string) {
  const normalized = requireString(value);
  if (!normalized) return null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(normalized)) return `${normalized}T00:00:00.000+05:30`;
  const parsed = new Date(normalized);
  return Number.isNaN(parsed.getTime()) ? null : parsed.toISOString();
}

function normalizeFieldKey(value: unknown) {
  return requireString(value).toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function findGreFieldPath(root: unknown, aliases: string[], depth = 0, path: (string | number)[] = []): (string | number)[] | null {
  if (!root || typeof root !== "object" || depth > 5) return null;
  const normalizedAliases = aliases.map(normalizeFieldKey).filter(Boolean);
  if (Array.isArray(root)) {
    for (let index = 0; index < root.length; index += 1) {
      const nested = findGreFieldPath(root[index], aliases, depth + 1, [...path, index]);
      if (nested) return nested;
    }
    return null;
  }
  const record = root as Record<string, unknown>;
  for (const key of Object.keys(record)) {
    if (normalizedAliases.includes(normalizeFieldKey(key))) return [...path, key];
  }
  for (const key of Object.keys(record)) {
    const nested = findGreFieldPath(record[key], aliases, depth + 1, [...path, key]);
    if (nested) return nested;
  }
  return null;
}

function getPathValue(root: Record<string, unknown>, path: (string | number)[]) {
  return path.reduce((current: unknown, key) => {
    if (!current || typeof current !== "object") return undefined;
    return (current as Record<string | number, unknown>)[key];
  }, root as unknown);
}

function setPathValue(root: Record<string, unknown>, path: (string | number)[], value: unknown) {
  let current: Record<string | number, unknown> | null = root;
  for (let index = 0; index < path.length - 1; index += 1) {
    const key = path[index];
    const next = current?.[key];
    if (!next || typeof next !== "object") return false;
    current = next as Record<string | number, unknown>;
  }
  if (!current) return false;
  current[path[path.length - 1]] = value;
  return true;
}

async function buildGreCompatibleFieldValue(existing: unknown, nextValue: string) {
  const normalized = requireString(nextValue);
  if (!normalized) return null;
  if (existing && typeof existing === "object" && !Array.isArray(existing)) {
    const record = existing as Record<string, unknown>;
    const classCode = requireString(record.classCode);
    if (classCode) {
      const mapped = findGreRefDataOption(await getGreRefData(classCode), normalized);
      if (mapped) return mapped;
    }
    return {
      ...record,
      name: normalized,
      value: normalized,
    };
  }
  return normalized;
}

async function setGreFieldIfPresent(root: Record<string, unknown>, aliases: string[], value: unknown) {
  const normalized = requireString(value);
  if (!normalized) return false;
  const path = findGreFieldPath(root, aliases);
  if (!path) return false;
  const existing = getPathValue(root, path);
  return setPathValue(root, path, await buildGreCompatibleFieldValue(existing, normalized));
}

async function syncApprovedUpdateToGre(requestRow: Record<string, any>) {
  const needId = requireString(requestRow.need_id);
  if (!needId) throw new Error("Need id is required for GRE sync.");

  const { sessionId, data: greRequest } = await fetchGreJson(`/commons-request-management-service/api/v1/request?id=${encodeURIComponent(needId)}`);
  const requestStatusOptions = await getGreRefData("CLASS.REQUEST_STATUS");
  const internalStatusOptions = await getGreRefData("CLASS.REQUEST_INTERNAL_STATUS");
  const demandBroadcastOptions = await getGreRefData("CLASS.DEMAND_BROADCAST_NEEDED");

  const payload = structuredClone(greRequest);
  const curationList = Array.isArray(payload.curationList) ? payload.curationList : [];
  const selectedCurationIndex = selectGreCurationIndex(curationList, requestRow);
  const primaryCuration = curationList[selectedCurationIndex] || {
    requestId: Number(needId),
    curatorUserId: null,
    curatorUserName: null,
    curatedNeedOfServiceSeekers: [],
    demandBroadcastNeed: null,
  };

  const unsupportedFields: string[] = [];

  if (requestRow.proposed_status) {
    const mapped = findGreRefDataOption(requestStatusOptions, requestRow.proposed_status, [
      requestRow.proposed_status === "Accepted" ? "Request Accepted" : "",
      requestRow.proposed_status === "Closed" ? "Request Closed" : "",
      requestRow.proposed_status === "New" ? "Request Submitted" : "",
      requestRow.proposed_status === "In progress" ? "Work in Progress" : "",
    ]);
    if (!mapped) throw new Error(`Could not map GRE request status for "${requestRow.proposed_status}".`);
    payload.status = mapped;
  }

  if (requestRow.proposed_internal_status) {
    const mapped = findGreRefDataOption(internalStatusOptions, requestRow.proposed_internal_status);
    if (!mapped) throw new Error(`Could not map GRE internal status for "${requestRow.proposed_internal_status}".`);
    payload.internalStatus = mapped;
  }

  if (requestRow.proposed_curation_notes) {
    primaryCuration.callDetails = requestRow.proposed_curation_notes;
  }

  if (requestRow.proposed_curation_call_date) {
    primaryCuration.callDate = buildGreDateString(requestRow.proposed_curation_call_date);
  }

  if (Array.isArray(requestRow.proposed_curated_need)) {
    const curatedNeedValues = asStringArray(requestRow.proposed_curated_need);
    if (curatedNeedValues.length) {
      const curatedNeedClass = await getGreRefData("CLASS.CURATED_NEED_OF_SERVICE_SEEKER");
      const curatedCodeMap = new Map<string, Record<string, unknown>>();
      for (const opt of curatedNeedClass as Record<string, unknown>[]) {
        const code = String(opt.code || "");
        const name = String(opt.name || "");
        if (name) curatedCodeMap.set(name.toLowerCase(), opt);
        if (code) {
          const shortCode = code.includes(".") ? code.split(".").pop()?.toLowerCase() || "" : code;
          if (shortCode) curatedCodeMap.set(shortCode, opt);
        }
      }
      primaryCuration.curatedNeedOfServiceSeekers = curatedNeedValues.map((v) => {
        const match = curatedCodeMap.get(v.toLowerCase());
        if (!match) throw new Error(`Curated need "${v}" not found in GRE reference data class CLASS.CURATED_NEED_OF_SERVICE_SEEKER.`);
        return {
          code: String(match.code),
          name: String(match.name),
          classCode: "CLASS.CURATED_NEED_OF_SERVICE_SEEKER",
        };
      });
    } else {
      primaryCuration.curatedNeedOfServiceSeekers = [];
    }
  }

  if (typeof requestRow.proposed_demand_broadcast_needed === "boolean") {
    const mapped = findGreRefDataOption(
      demandBroadcastOptions,
      requestRow.proposed_demand_broadcast_needed ? "Yes" : "No",
    );
    if (!mapped) throw new Error("Could not map GRE demand broadcast option.");
    primaryCuration.demandBroadcastNeed = mapped;
  }

  if (requestRow.proposed_next_action) {
    const synced = await setGreFieldIfPresent(payload, ["nextAction", "next_action", "nextActionStatus"], requestRow.proposed_next_action);
    if (!synced) unsupportedFields.push("next_action");
  }
  const outcomeSyncs: Array<[string, string[], unknown]> = [
    ["funding_mechanism", ["fundingMechanism", "funding_mechanism", "funding", "outcomeFundingMechanism"], requestRow.proposed_funding_mechanism],
    ["seeker_provider_agreement", ["seekerProviderAgreement", "seeker_provider_agreement", "seekerAgreement", "providerAgreement"], requestRow.proposed_seeker_provider_agreement],
    ["solution_deployment_status", ["solutionDeploymentStatus", "solution_deployment_status", "deploymentStatus"], requestRow.proposed_solution_deployment_status],
    ["closure_date", ["closureDate", "closure_date", "closedOn"], buildGreDateString(requireString(requestRow.proposed_closure_date)) || requestRow.proposed_closure_date],
    ["feedback_about_seeker", ["feedbackAboutSeeker", "feedback_about_seeker", "seekerFeedback"], requestRow.proposed_feedback_about_seeker],
    ["feedback_about_provider", ["feedbackAboutProvider", "feedback_about_provider", "providerFeedback"], requestRow.proposed_feedback_about_provider],
  ];
  for (const [fieldName, aliases, value] of outcomeSyncs) {
    if (!requireString(value)) continue;
    const syncedOnRequest = await setGreFieldIfPresent(payload, aliases, value);
    const syncedOnCuration = await setGreFieldIfPresent(primaryCuration, aliases, value);
    if (!syncedOnRequest && !syncedOnCuration) unsupportedFields.push(fieldName);
  }
  if (Number.isInteger(requestRow.proposed_solutions_shared_count)) unsupportedFields.push("solutions_shared_count");
  if (Number.isInteger(requestRow.proposed_invited_providers_count)) unsupportedFields.push("invited_providers_count");

  const nextCurationList = [...curationList];
  nextCurationList[selectedCurationIndex] = primaryCuration;
  payload.curationList = nextCurationList;
  await updateGreRequestJson("/commons-request-management-service/api/v1/request", payload, sessionId);
  const { data: verifiedRequest } = await fetchGreJson(
    `/commons-request-management-service/api/v1/request?id=${encodeURIComponent(needId)}`,
    sessionId,
  );
  const verifiedPayload = (verifiedRequest && typeof verifiedRequest === "object") ? verifiedRequest as Record<string, unknown> : {};
  const verifiedCurationList = Array.isArray(verifiedPayload.curationList) ? verifiedPayload.curationList : [];
  const verifiedPrimaryCuration = findVerifiedGreCurationEntry(
    verifiedCurationList,
    primaryCuration,
    selectedCurationIndex,
    requestRow,
  );

  if (requestRow.proposed_status) {
    const nextCode = requireString(((verifiedPayload.status || {}) as Record<string, unknown>).code);
    const expectedCode = requireString((payload.status || {}).code);
    if (expectedCode && nextCode !== expectedCode) {
      throw new Error(`GRE verification failed for status. Expected ${expectedCode}, found ${nextCode || "blank"}.`);
    }
  }
  if (requestRow.proposed_internal_status) {
    const nextCode = requireString(((verifiedPayload.internalStatus || {}) as Record<string, unknown>).code);
    const expectedCode = requireString((payload.internalStatus || {}).code);
    if (expectedCode && nextCode !== expectedCode) {
      throw new Error(`GRE verification failed for internal status. Expected ${expectedCode}, found ${nextCode || "blank"}.`);
    }
  }
  if (requestRow.proposed_curation_notes) {
    const verifiedNotes = requireString(verifiedPrimaryCuration.callDetails);
    const expectedNotes = requireString(primaryCuration.callDetails);
    const normalizedVerifiedNotes = normalizeComparable(verifiedNotes);
    const normalizedExpectedNotes = normalizeComparable(expectedNotes);
    const notesMatch =
      !normalizedExpectedNotes ||
      normalizedVerifiedNotes === normalizedExpectedNotes ||
      normalizedVerifiedNotes.includes(normalizedExpectedNotes) ||
      normalizedExpectedNotes.includes(normalizedVerifiedNotes);
    if (!notesMatch) {
      throw new Error(
        `GRE verification failed for curation notes. Expected snippet "${expectedNotes.slice(0, 80)}", found "${verifiedNotes.slice(0, 80)}".`,
      );
    }
  }
  if (requestRow.proposed_curation_call_date) {
    const verifiedDate = requireString(verifiedPrimaryCuration.callDate);
    const expectedDate = requireString(primaryCuration.callDate);
    if (expectedDate && verifiedDate !== expectedDate) {
      throw new Error(`GRE verification failed for curation call date. Expected ${expectedDate}, found ${verifiedDate || "blank"}.`);
    }
  }
  if (typeof requestRow.proposed_demand_broadcast_needed === "boolean") {
    const nextCode = requireString((((verifiedPrimaryCuration.demandBroadcastNeed || {}) as Record<string, unknown>).code));
    const expectedCode = requireString((((primaryCuration.demandBroadcastNeed || {}) as Record<string, unknown>).code));
    if (expectedCode && nextCode !== expectedCode) {
      throw new Error(`GRE verification failed for demand broadcast flag. Expected ${expectedCode}, found ${nextCode || "blank"}.`);
    }
  }
  return { ok: true, unsupportedFields };
}

function selectGreCurationIndex(curationList: unknown[], requestRow: Record<string, unknown>) {
  if (!Array.isArray(curationList) || !curationList.length) return 0;
  const wantsNotes = Boolean(requireString(requestRow.proposed_curation_notes));
  const wantsDate = Boolean(requireString(requestRow.proposed_curation_call_date));
  if (wantsNotes || wantsDate) {
    const withExistingDetail = curationList.findIndex((entry) => {
      const row = (entry && typeof entry === "object") ? entry as Record<string, unknown> : {};
      return Boolean(requireString(row.callDetails) || requireString(row.callDate));
    });
    if (withExistingDetail >= 0) return withExistingDetail;
  }
  return 0;
}

function findVerifiedGreCurationEntry(
  curationList: unknown[],
  expectedCuration: Record<string, unknown>,
  selectedCurationIndex: number,
  requestRow: Record<string, unknown>,
) {
  const rows = Array.isArray(curationList)
    ? curationList.map((entry) => (entry && typeof entry === "object") ? entry as Record<string, unknown> : {})
    : [];
  if (!rows.length) return {};

  const selectedRow = rows[selectedCurationIndex] || {};
  const expectedNotes = normalizeComparable(expectedCuration.callDetails);
  const expectedDate = requireString(expectedCuration.callDate);
  const expectedCuratorId = requireString(expectedCuration.curatorUserId);
  const expectedCuratorName = normalizeComparable(expectedCuration.curatorUserName);
  const wantsNotes = Boolean(requireString(requestRow.proposed_curation_notes));
  const wantsDate = Boolean(requireString(requestRow.proposed_curation_call_date));

  const scoreRow = (row: Record<string, unknown>) => {
    let score = 0;
    const rowNotes = normalizeComparable(row.callDetails);
    const rowDate = requireString(row.callDate);
    const rowCuratorId = requireString(row.curatorUserId);
    const rowCuratorName = normalizeComparable(row.curatorUserName);

    if (expectedCuratorId && rowCuratorId === expectedCuratorId) score += 6;
    if (expectedCuratorName && rowCuratorName === expectedCuratorName) score += 3;
    if (wantsDate && expectedDate && rowDate === expectedDate) score += 5;
    if (wantsNotes && expectedNotes) {
      if (rowNotes === expectedNotes) score += 8;
      else if (rowNotes.includes(expectedNotes) || expectedNotes.includes(rowNotes)) score += 4;
    }
    return score;
  };

  const bestRow = rows
    .map((row, index) => ({ row, index, score: scoreRow(row) }))
    .sort((a, b) => b.score - a.score || a.index - b.index)[0];

  if (bestRow && bestRow.score > 0) return bestRow.row;
  return selectedRow;
}

function stableRowSignature(value: unknown) {
  return JSON.stringify(value, Object.keys(value as Record<string, unknown>).sort());
}

const inboundOutcomeSignatureKeys = [
  "funding_mechanism",
  "seeker_provider_agreement",
  "solution_deployment_status",
  "closure_date",
  "feedback_about_seeker",
  "feedback_about_provider",
];

function stableInboundRowSignature(row: Record<string, unknown>) {
  const signatureRow: Record<string, unknown> = { ...row };
  inboundOutcomeSignatureKeys.forEach((key) => {
    const value = signatureRow[key];
    const isEmptyArray = Array.isArray(value) && value.length === 0;
    if (!requireString(value) || isEmptyArray) {
      delete signatureRow[key];
    }
  });
  return stableRowSignature(signatureRow);
}

function buildNeedIntelligencePrompt(need: Record<string, unknown>) {
  return `
Classify this GRE inbound need into structured JSON for matching. Return only valid JSON.

Rules:
- Extract the main thematic or application area as specifically as possible.
- thematic_area must be the real domain or application area, such as dairy, solar street lights, wild mango, goatery, branding, packaging.
- Choose need_kind as one of: product, service, knowledge, finance, mixed.
- Prefer service when the need is asking for implementation support, advisory, handholding, consulting, mentoring, training, deployment support, or operational guidance.
- Prefer product when the need is asking for equipment, machinery, physical units, or installed infrastructure.
- Prefer knowledge when the need is mainly asking for manuals, SOPs, videos, reports, or learning content.
- Use mixed only when the need clearly needs more than one of service/product/knowledge in the same request.
- If need_kind is service, fill service_kind with the specific service type. Otherwise return empty string.
- Pick zero or more 6M labels from exactly: Manpower, Method, Machine, Material, Market, Money.
- keywords should be a concise array of high-value match terms.
- summary should be one short sentence focused on match intent.

Need payload:
${JSON.stringify({
    organization_name: need.organization_name,
    state: need.state,
    district: need.district,
    problem_statement: need.problem_statement,
    curated_need: need.curated_need,
    curation_notes: need.curation_notes,
    funding_mechanism: need.funding_mechanism,
    seeker_provider_agreement: need.seeker_provider_agreement,
    solution_deployment_status: need.solution_deployment_status,
    feedback_about_seeker: need.feedback_about_seeker,
    feedback_about_provider: need.feedback_about_provider,
    solutions_shared_count: need.solutions_shared_count,
  })}

Return JSON with exactly:
{
  "thematic_area": "",
  "application_area": "",
  "need_kind": "",
  "service_kind": "",
  "six_m_signals": [],
  "keywords": [],
  "summary": ""
  }`.trim();
}

function buildOfferingIntelligencePrompt(input: {
  offering: Record<string, unknown>;
  solution: Record<string, unknown> | null;
  trader: Record<string, unknown> | null;
}) {
  return `
Classify this GRE solution offering into structured JSON for matching. Return only valid JSON.

Rules:
- thematic_area must be the real domain or application area, such as dairy, solar street lights, wild mango, goatery, branding, packaging.
- Choose offering_kind as one of: product, service, knowledge, finance, mixed.
- Prefer service when the offering provides implementation support, consulting, mentoring, training, technology transfer, advisory, deployment, or handholding.
- Prefer product when the offering is a physical product, machinery, equipment, raw material, or installed unit.
- Prefer knowledge when the offering is mainly manuals, SOPs, videos, reports, or static knowledge content.
- Use mixed only when the offering genuinely combines more than one kind.
- If offering_kind is service, fill service_kind with the specific service type. Otherwise return empty string.
- Pick zero or more 6M labels from exactly: Manpower, Method, Machine, Material, Market, Money.
- keywords should be a concise array of high-value domain and subdomain terms.
- summary should be one short sentence focused on deployment relevance.

Offering payload:
${JSON.stringify({
    offering_name: input.offering.offering_name,
    offering_group: input.offering.offering_group,
    offering_type: input.offering.offering_type,
    offering_category: input.offering.offering_category,
    primary_application: input.offering.primary_application,
    primary_valuechain: input.offering.primary_valuechain,
    applications: input.offering.applications,
    valuechains: input.offering.valuechains,
    tags: input.offering.tags,
    domain_6m: input.offering.domain_6m,
    about_offering_text: input.offering.about_offering_text,
    about_solution_text: input.solution?.about_solution_text,
    solution_name: input.solution?.solution_name,
    trader_name: input.trader?.organisation_name || input.trader?.trader_name,
  })}

Return JSON with exactly:
{
  "thematic_area": "",
  "application_area": "",
  "offering_kind": "",
  "service_kind": "",
  "six_m_signals": [],
  "keywords": [],
  "summary": ""
}`.trim();
}

function classifyNeedByRules(need: Record<string, unknown>) {
  const curatedNeed = asStringArray(need.curated_need).map((item) => item.toLowerCase());
  const cleanNotes = requireString(need.curation_notes).replace(/\bnull\b/gi, " ");
  const text = [
    requireString(need.problem_statement),
    cleanNotes,
    curatedNeed.join(" "),
  ].join(" | ").toLowerCase();

  const thematicHints = uniqueStrings(
    ruleThemeSignals
      .filter((rule) => rule.patterns.some((pattern) => text.includes(pattern)))
      .map((rule) => rule.label),
  );

  const serviceHints = uniqueStrings(
    ruleServiceKindPatterns
      .filter((rule) => rule.patterns.some((pattern) => text.includes(pattern) || curatedNeed.some((entry) => entry.includes(pattern))))
      .map((rule) => rule.label),
  );

  const rule6M = uniqueStrings(
    [
      text.includes("training") || text.includes("capacity building") ? "Manpower" : "",
      ["consulting", "consultancy", "mentoring", "technology transfer", "manual", "video", "sop", "blog"].some((pattern) => text.includes(pattern)) ? "Method" : "",
      ["machine", "machinery", "plant setup", "street light", "street lights", "equipment"].some((pattern) => text.includes(pattern)) ? "Machine" : "",
      ["raw material", "raw materials", "material supply"].some((pattern) => text.includes(pattern)) ? "Material" : "",
      ["products bought", "market support", "market report", "market reports", "branding", "packaging", "business development"].some((pattern) => text.includes(pattern)) ? "Market" : "",
      ["financial support", "finance", "funding", "investment", "credit", "loan"].some((pattern) => text.includes(pattern)) ? "Money" : "",
    ].filter(Boolean) as string[],
  );

  const ruleNeedKinds = uniqueStrings(
    ruleNeedKindPatterns
      .filter((rule) => rule.patterns.some((pattern) => text.includes(pattern)))
      .map((rule) => rule.label),
  );

  let needKind = ruleNeedKinds.length > 1 ? "mixed" : ruleNeedKinds[0] || "";
  if (serviceHints.length && thematicHints.length && ruleNeedKinds.includes("service")) {
    needKind = "service";
  }
  const serviceKind = needKind === "service" || needKind === "mixed" ? serviceHints[0] || "" : "";
  const domainTokens = uniqueStrings([
    ...extractNeedPriorityPhrases(text),
    ...extractNeedWindowPhrases(text),
    ...tokenizeLooseText(text, 5).filter((token) => !domainStopwords.has(token) && !needTagStopwords.has(token)),
  ]).slice(0, 18);

  return {
    thematicHints,
    serviceHints,
    sixMSignals: rule6M,
    needKind,
    serviceKind,
    keywords: uniqueStrings([...thematicHints, ...domainTokens]).filter((entry) => !isWeakNeedTag(entry)).slice(0, 18),
  };
}

function classifyOfferingByRules(
  offering: Record<string, unknown>,
  solution: Record<string, unknown> | null,
) {
  const text = [
    requireString(offering.offering_name),
    requireString(offering.offering_group),
    requireString(offering.offering_type),
    requireString(offering.offering_category),
    requireString(offering.primary_application),
    requireString(offering.primary_valuechain),
    asStringArray(offering.applications).join(" "),
    asStringArray(offering.valuechains).join(" "),
    asStringArray(offering.tags).join(" "),
    requireString(offering.domain_6m),
    requireString(offering.about_offering_text),
    requireString(solution?.solution_name),
    requireString(solution?.about_solution_text),
  ].join(" | ").toLowerCase();

  const thematicHints = uniqueStrings(
    ruleThemeSignals
      .filter((rule) => rule.patterns.some((pattern) => text.includes(pattern)))
      .map((rule) => rule.label),
  );

  const serviceHints = uniqueStrings(
    ruleServiceKindPatterns
      .filter((rule) => rule.patterns.some((pattern) => text.includes(pattern)))
      .map((rule) => rule.label),
  );

  const offeringGroup = requireString(offering.offering_group).toLowerCase();
  const offeringType = requireString(offering.offering_type).toLowerCase();
  const category = requireString(offering.offering_category).toLowerCase();

  let offeringKind = "";
  if (offeringGroup.includes("service") || category.includes("service")) offeringKind = "service";
  else if (offeringGroup.includes("product") || category.includes("product")) offeringKind = "product";
  else if (offeringGroup.includes("knowledge") || category.includes("knowledge") || ["manual", "video", "sop", "blog"].some((token) => offeringType.includes(token))) offeringKind = "knowledge";
  else {
    const kinds = uniqueStrings(
      ruleNeedKindPatterns
        .filter((rule) => rule.patterns.some((pattern) => text.includes(pattern)))
        .map((rule) => rule.label === "service" ? "service" : rule.label),
    );
    offeringKind = kinds.length > 1 ? "mixed" : kinds[0] || "";
  }

  const rule6M = uniqueStrings(
    [
      ["training", "capacity building"].some((pattern) => text.includes(pattern)) ? "Manpower" : "",
      ["consulting", "consultancy", "mentoring", "technology transfer", "manual", "video", "sop", "blog", "advisory"].some((pattern) => text.includes(pattern)) ? "Method" : "",
      ["machine", "machinery", "equipment", "plant setup", "street light"].some((pattern) => text.includes(pattern)) ? "Machine" : "",
      ["raw material", "material", "supply"].some((pattern) => text.includes(pattern)) ? "Material" : "",
      ["market", "branding", "packaging", "marketplace"].some((pattern) => text.includes(pattern)) ? "Market" : "",
      ["financial", "finance", "funding", "credit", "loan"].some((pattern) => text.includes(pattern)) ? "Money" : "",
    ].filter(Boolean) as string[],
  );

  return {
    thematicHints,
    serviceHints,
    sixMSignals: rule6M,
    offeringKind,
    serviceKind: offeringKind === "service" || offeringKind === "mixed" ? serviceHints[0] || "" : "",
    keywords: uniqueStrings([
      ...thematicHints,
      ...tokenizeLooseText(text, 4).filter((token) => !domainStopwords.has(token)),
    ]).slice(0, 18),
  };
}

function buildNeedTagDraft(input: {
  thematicArea?: unknown;
  problemStatement?: unknown;
  rules: ReturnType<typeof classifyNeedByRules>;
}) {
  const problemStatement = requireString(input.problemStatement);
  const phrases = uniqueStrings([
    ...extractNeedPriorityPhrases(problemStatement),
    ...extractNeedWindowPhrases(problemStatement),
  ]);
  const scopedTokens = tokenizeLooseText(problemStatement, 5)
    .filter((token) => !needTagStopwords.has(token) && !domainStopwords.has(token))
    .filter((token) => !/^\d+$/.test(token))
    .slice(0, 8);
  return uniqueStrings([
    normalizeNeedTag(input.thematicArea),
    ...phrases,
    ...input.rules.thematicHints.map((entry) => normalizeNeedTag(entry)),
    ...input.rules.serviceHints.map((entry) => normalizeNeedTag(entry)),
    ...input.rules.keywords.map((entry) => normalizeNeedTag(entry)),
    ...scopedTokens.map((entry) => normalizeNeedTag(entry)),
  ])
    .filter((entry) => !isWeakNeedTag(entry))
    .slice(0, 12);
}

function extractSolutionPriorityPhrases(value: unknown) {
  const text = requireString(value)
    .toLowerCase()
    .replace(/[\r\n]+/g, " ")
    .replace(/\s+/g, " ");

  const phrases: string[] = [];

  for (const match of text.matchAll(/\b([a-z][a-z\s-]{1,30})\s+processing\b/g)) {
    const phrase = normalizeNeedTag(`${requireString(match[1])} processing`);
    if (!isWeakNeedTag(phrase)) phrases.push(phrase);
  }

  if (/\bvalue added\b|\bvalue addition\b/.test(text)) phrases.push("value addition");
  if (/\bdrying\b/.test(text) && /\bmango|fruit|fruits\b/.test(text)) phrases.push("fruit drying");
  if (/\bachar\b/.test(text)) phrases.push("achar production");
  if (/\bamchur\b/.test(text)) phrases.push("amchur powder");
  if (/\bjuice\b|\bjuices\b/.test(text)) phrases.push("juice processing");
  if (/\bdrying\b|\bpreservation\b|\bpickle\b|\bachar\b|\bamchur\b/.test(text)) phrases.push("food preservation");
  if (/\bmango\b|\bfruit\b|\bfruits\b|\bagricultural\b|\bproduce\b/.test(text)) phrases.push("agricultural products");
  if (/\brural\b|\bvillage\b|\bvillages\b|\bfarmer\b|\bfarmers\b|\bsmallholder\b/.test(text)) phrases.push("rural livelihoods");
  if (/\bsmallholder\b|\bfarmers?\b/.test(text)) phrases.push("smallholder farmers");

  return uniqueStrings(
    phrases
      .map((entry) => normalizeNeedTag(entry))
      .filter((entry) => !isWeakNeedTag(entry)),
  );
}

function buildSolutionTagDraft(input: {
  offeringName?: unknown;
  offeringDescription?: unknown;
  organizationName?: unknown;
  rules: ReturnType<typeof classifyOfferingByRules>;
}) {
  const combined = [
    requireString(input.offeringName),
    requireString(input.offeringDescription),
    requireString(input.organizationName),
  ].filter(Boolean).join(" | ");
  const phrases = extractSolutionPriorityPhrases(combined);
  const scopedTokens = tokenizeLooseText(combined, 5)
    .filter((token) => !domainStopwords.has(token))
    .filter((token) => !/^\d+$/.test(token))
    .slice(0, 6);
  return uniqueStrings([
    ...phrases,
    ...input.rules.thematicHints.map((entry) => normalizeNeedTag(entry)),
    ...input.rules.serviceHints.map((entry) => normalizeNeedTag(entry)),
    ...input.rules.keywords.map((entry) => normalizeNeedTag(entry)),
    ...scopedTokens.map((entry) => normalizeNeedTag(entry)),
  ])
    .filter((entry) => !isWeakNeedTag(entry))
    .slice(0, 12);
}

function validateAiEnrichment(
  need: Record<string, unknown>,
  ai: Record<string, unknown>,
  rules: ReturnType<typeof classifyNeedByRules>,
) {
  const flags: string[] = [];
  const aiThematic = requireString(ai.thematic_area).toLowerCase();
  const aiNeedKind = requireString(ai.need_kind).toLowerCase();
  const aiServiceKind = requireString(ai.service_kind).toLowerCase();
  const aiKeywords = asStringArray(ai.keywords).map((item) => item.toLowerCase());
  const aiSixM = asStringArray(ai.six_m_signals);

  if (!aiThematic) {
    flags.push("missing_thematic_area");
  } else if (genericThematicTerms.has(aiThematic)) {
    flags.push("generic_thematic_area");
  }

  if (!aiNeedKind) {
    flags.push("missing_need_kind");
  }
  if (aiNeedKind !== "service" && aiServiceKind) {
    flags.push("service_kind_without_service_need");
  }
  if (aiNeedKind === "service" && !aiServiceKind) {
    flags.push("missing_service_kind");
  }

  if (rules.sixMSignals.length && !rules.sixMSignals.some((label) => aiSixM.includes(label))) {
    flags.push("six_m_mismatch");
  }
  if (rules.thematicHints.length && aiThematic && !rules.thematicHints.some((hint) => aiThematic.includes(hint) || hint.includes(aiThematic))) {
    flags.push("thematic_mismatch");
  }
  if (rules.keywords.length && !rules.keywords.some((keyword) => aiKeywords.some((item) => item.includes(keyword) || keyword.includes(item)))) {
    flags.push("keyword_mismatch");
  }
  if (!aiKeywords.length) {
    flags.push("missing_keywords");
  }

  let confidence = 94;
  confidence -= flags.length * 14;
  if (rules.thematicHints.length && aiThematic && rules.thematicHints.some((hint) => aiThematic.includes(hint) || hint.includes(aiThematic))) {
    confidence += 6;
  }
  if (rules.sixMSignals.length && rules.sixMSignals.some((label) => aiSixM.includes(label))) {
    confidence += 4;
  }
  confidence = Math.max(0, Math.min(100, confidence));

  return {
    flags: uniqueStrings(flags),
    status: flags.length ? "flagged" : "ready",
    confidence,
  };
}

function getConfiguredAiProviders() {
  return [
    geminiApiKey ? "gemini" : "",
    openRouterApiKey ? "openrouter" : "",
    deepSeekApiKey ? "deepseek" : "",
    openAiApiKey ? "openai" : "",
  ].filter(Boolean);
}

async function callAiJsonForProvider(resolvedProvider: string, prompt: string) {
  if (resolvedProvider === "gemini") {
    if (!geminiApiKey) throw new Error("GEMINI_API_KEY is not configured.");
    const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${defaultGeminiModel}:generateContent?key=${geminiApiKey}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { responseMimeType: "application/json" },
      }),
    });
    const data = await response.json().catch(() => null);
    const text = data?.candidates?.[0]?.content?.parts?.map((part: { text?: string }) => part.text || "").join("") || "";
    if (!response.ok || !text) throw new Error(data?.error?.message || "Gemini enrichment failed.");
    return JSON.parse(text);
  }

  if (resolvedProvider === "deepseek") {
    if (!deepSeekApiKey) throw new Error("DEEPSEEK_API_KEY is not configured.");
    const response = await fetch("https://api.deepseek.com/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${deepSeekApiKey}`,
      },
      body: JSON.stringify({
        model: defaultDeepSeekModel,
        response_format: { type: "json_object" },
        messages: [{ role: "user", content: prompt }],
      }),
    });
    const data = await response.json().catch(() => null);
    const text = data?.choices?.[0]?.message?.content || "";
    if (!response.ok || !text) throw new Error(data?.error?.message || "DeepSeek enrichment failed.");
    return JSON.parse(text);
  }

  if (resolvedProvider === "openai") {
    if (!openAiApiKey) throw new Error("OPENAI_API_KEY is not configured.");
    const response = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${openAiApiKey}`,
      },
      body: JSON.stringify({
        model: defaultOpenAiModel,
        response_format: { type: "json_object" },
        messages: [{ role: "user", content: prompt }],
      }),
    });
    const data = await response.json().catch(() => null);
    const text = data?.choices?.[0]?.message?.content || "";
    if (!response.ok || !text) throw new Error(data?.error?.message || "OpenAI enrichment failed.");
    return JSON.parse(text);
  }

  if (!openRouterApiKey) throw new Error("OPENROUTER_API_KEY is not configured.");
  const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${openRouterApiKey}`,
      "HTTP-Referer": "https://greenruraleconomy.in",
      "X-Title": "GRE MIS Dashboard",
    },
    body: JSON.stringify({
      model: defaultOpenRouterModel,
      response_format: { type: "json_object" },
      messages: [{ role: "user", content: prompt }],
    }),
  });
  const data = await response.json().catch(() => null);
  const text = data?.choices?.[0]?.message?.content || "";
  if (!response.ok || !text) throw new Error(data?.error?.message || "OpenRouter enrichment failed.");
  return JSON.parse(text);
}

async function callAiJsonWithOrder(order: string[], prompt: string) {
  const configuredProviders = getConfiguredAiProviders();
  const fallbackOrder = [
    ...order.map((item) => item.toLowerCase()).filter((candidate) => configuredProviders.includes(candidate)),
    ...configuredProviders.filter((candidate) => !order.map((item) => item.toLowerCase()).includes(candidate)),
  ];
  if (!fallbackOrder.length) {
    throw new Error("No AI provider is configured. Falling back to rule-based enrichment.");
  }

  let lastError: unknown = null;
  for (const candidate of fallbackOrder) {
    try {
      return await callAiJsonForProvider(candidate, prompt);
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError instanceof Error ? lastError : new Error("No AI provider could complete the request.");
}

async function callAiJson(providerInput: string, prompt: string) {
  const provider = (providerInput || defaultAiProvider || "openrouter").toLowerCase();

  const requestedOrder = provider === "gemini"
    ? ["gemini", "openai", "openrouter", "deepseek"]
    : provider === "deepseek"
      ? ["deepseek", "openrouter", "gemini", "openai"]
      : provider === "openai"
        ? ["openai", "gemini", "openrouter", "deepseek"]
        : ["openrouter", "gemini", "deepseek", "openai"];
  return await callAiJsonWithOrder(requestedOrder, prompt);
}

function isHindiSourceLanguage(value: unknown) {
  return requireString(value).toLowerCase() === "hi";
}

async function translateTextWithLibreTranslate(
  text: string,
  source = "auto",
  target = "en",
) {
  const normalized = requireString(text);
  if (!normalized || source === target) return normalized;

  try {
    const requestBody: Record<string, unknown> = {
      q: normalized,
      source,
      target,
      format: "text",
    };
    if (libreTranslateApiKey) requestBody.api_key = libreTranslateApiKey;

    const response = await fetch(libreTranslateApiUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(requestBody),
    });
    const data = await response.json().catch(() => null);
    if (!response.ok) {
      throw new Error(requireString(data?.error || data?.message) || "LibreTranslate request failed.");
    }
    return requireString(data?.translatedText || normalized);
  } catch (error) {
    console.warn("LibreTranslate failed, using original text:", error instanceof Error ? error.message : String(error));
    return normalized;
  }
}

async function translateStringArrayToEnglish(values: unknown) {
  const normalizedValues = uniqueStrings(asStringArray(values));
  if (!normalizedValues.length) return [];
  const translated = await Promise.all(
    normalizedValues.map((value) => translateTextWithLibreTranslate(value, "hi", "en")),
  );
  return uniqueStrings(translated);
}

async function normalizeSubmissionTranslations(
  submissionType: string,
  input: Record<string, unknown>,
) {
  const payload = { ...input };
  if (!isHindiSourceLanguage(payload.source_language)) return payload;

  if (requireString(payload.organization_name)) {
    payload.original_organization_name_hi = requireString(payload.organization_name);
    payload.organization_name = await translateTextWithLibreTranslate(requireString(payload.organization_name), "hi", "en");
  }

  if (submissionType === "solution") {
    if (requireString(payload.offering_name)) {
      payload.original_offering_name_hi = requireString(payload.offering_name);
      payload.offering_name = await translateTextWithLibreTranslate(requireString(payload.offering_name), "hi", "en");
    }
    const currentSolutionName = requireString(payload.solution_name || payload.offering_name);
    if (currentSolutionName) {
      payload.original_solution_name_hi = requireString(payload.solution_name || "");
      payload.solution_name = await translateTextWithLibreTranslate(currentSolutionName, "hi", "en");
    }
    if (requireString(payload.about_offering_text)) {
      payload.translated_about_offering_text_en = await translateTextWithLibreTranslate(
        requireString(payload.about_offering_text),
        "hi",
        "en",
      );
    }
    if (requireString(payload.about_solution_text)) {
      payload.translated_about_solution_text_en = await translateTextWithLibreTranslate(
        requireString(payload.about_solution_text),
        "hi",
        "en",
      );
    } else if (requireString(payload.translated_about_offering_text_en)) {
      payload.translated_about_solution_text_en = requireString(payload.translated_about_offering_text_en);
    }
    payload.tags = await translateStringArrayToEnglish(payload.tags);
    payload.geographies = await translateStringArrayToEnglish(payload.geographies);
  } else if (submissionType === "need") {
    if (requireString(payload.thematic_area)) {
      payload.original_thematic_area_hi = requireString(payload.thematic_area);
      payload.thematic_area = await translateTextWithLibreTranslate(requireString(payload.thematic_area), "hi", "en");
    }
    if (requireString(payload.problem_statement)) {
      payload.translated_problem_statement_en = await translateTextWithLibreTranslate(
        requireString(payload.problem_statement),
        "hi",
        "en",
      );
    }
    payload.keywords = await translateStringArrayToEnglish(payload.keywords);
    payload.deployment_locations = await translateStringArrayToEnglish(payload.deployment_locations);
  }

  return payload;
}

async function searchLgdGeographies(query: string) {
  const search = requireString(query);
  if (search.length < 2) {
    return { ok: true, suggestions: scoreLgdSuggestion("India", search) > 0 ? ["India"] : [] };
  }
  const lgdBase = "https://grameee.org/api/lgd";
  try {
    const response = await fetch(`${lgdBase}/s.php?q=${encodeURIComponent(search)}`);
    if (!response.ok) {
      console.warn(`Hostinger LGD search returned ${response.status}`);
      return { ok: true, suggestions: [] };
    }
    const data = await response.json() as Array<{ label: string }>;
    const labels = (data || []).map((item) => item.label).filter(Boolean);
    const indiaScore = scoreLgdSuggestion("India", search);
    return {
      ok: true,
      suggestions: uniqueStrings([
        ...(indiaScore > 0 ? ["India"] : []),
        ...labels,
      ]),
    };
  } catch (error) {
    console.warn("Hostinger LGD search error", error);
    return { ok: true, suggestions: [] };
  }
}

async function enrichNeedIntelligence(need: Record<string, unknown>, provider: string) {
  const rules = classifyNeedByRules(need);
  const basePatch = {
    rule_thematic_hints: rules.thematicHints,
    rule_service_hints: rules.serviceHints,
    rule_keywords: rules.keywords,
    rule_6m_signals: rules.sixMSignals,
    rule_need_kind: rules.needKind,
  };
  const nowIso = new Date().toISOString();

  let patch: Record<string, unknown>;
  try {
    const ai = await callAiJson(provider, buildNeedIntelligencePrompt(need));
    const validation = validateAiEnrichment(need, ai, rules);
    patch = {
      ...basePatch,
      ai_thematic_area: requireString(ai.thematic_area),
      ai_application_area: requireString(ai.application_area),
      ai_need_kind: requireString(ai.need_kind),
      ai_service_kind: requireString(ai.service_kind),
      ai_keywords: asStringArray(ai.keywords),
      ai_6m_signals: asStringArray(ai.six_m_signals),
      ai_summary: requireString(ai.summary),
      ai_engine: provider || defaultAiProvider,
      ai_enriched_at: nowIso,
      ai_enrichment_status: validation.status === "ready" ? "ready" : "ready_flagged",
      ai_validation_flags: validation.flags,
      ai_validation_status: validation.status,
      ai_confidence: validation.confidence,
      ai_prompt_version: aiPromptVersion,
      ai_schema_version: aiSchemaVersion,
      ai_payload: ai,
    };
  } catch (error) {
    const fallbackFlags = rules.thematicHints.length || rules.keywords.length ? [] : ["needs_review"];
    patch = {
      ...basePatch,
      ai_thematic_area: rules.thematicHints[0] || "",
      ai_application_area: "",
      ai_need_kind: rules.needKind || "",
      ai_service_kind: rules.serviceHints[0] || "",
      ai_keywords: rules.keywords,
      ai_6m_signals: rules.sixMSignals,
      ai_summary: requireString(need.problem_statement).slice(0, 500),
      ai_engine: "rules_only",
      ai_enriched_at: nowIso,
      ai_enrichment_status: "rules_only",
      ai_validation_flags: fallbackFlags,
      ai_validation_status: fallbackFlags.length ? "flagged" : "ready",
      ai_confidence: rules.thematicHints.length ? 74 : rules.keywords.length ? 61 : 28,
      ai_prompt_version: aiPromptVersion,
      ai_schema_version: aiSchemaVersion,
      ai_payload: {
        mode: "rules_only",
        provider_requested: provider || defaultAiProvider,
        reason: error instanceof Error ? error.message : "AI enrichment unavailable.",
      },
    };
  }

  const { error } = await adminClient.from("gre_mis_needs").update(patch).eq("id", requireString(need.id));
  if (error) throw new Error(error.message);
  return patch;
}

async function enrichOfferingIntelligence(
  offering: Record<string, unknown>,
  solution: Record<string, unknown> | null,
  trader: Record<string, unknown> | null,
  provider: string,
) {
  const rules = classifyOfferingByRules(offering, solution);
  const basePatch = {
    rule_thematic_hints: rules.thematicHints,
    rule_service_hints: rules.serviceHints,
    rule_keywords: rules.keywords,
    rule_6m_signals: rules.sixMSignals,
  };
  const nowIso = new Date().toISOString();

  let patch: Record<string, unknown>;
  try {
    const ai = await callAiJson(provider, buildOfferingIntelligencePrompt({ offering, solution, trader }));
    patch = {
      ...basePatch,
      ai_thematic_area: requireString(ai.thematic_area),
      ai_application_area: requireString(ai.application_area),
      ai_offering_kind: requireString(ai.offering_kind),
      ai_service_kind: requireString(ai.service_kind),
      ai_keywords: asStringArray(ai.keywords),
      ai_6m_signals: asStringArray(ai.six_m_signals),
      ai_summary: requireString(ai.summary),
      ai_engine: provider || defaultAiProvider,
      ai_enriched_at: nowIso,
      ai_enrichment_status: "ready",
      ai_prompt_version: aiPromptVersion,
      ai_schema_version: `${aiSchemaVersion}.offering`,
      ai_payload: ai,
    };
  } catch (error) {
    patch = {
      ...basePatch,
      ai_thematic_area: rules.thematicHints[0] || "",
      ai_application_area: requireString(offering.primary_application) || requireString(offering.primary_valuechain),
      ai_offering_kind: rules.offeringKind || "",
      ai_service_kind: rules.serviceKind,
      ai_keywords: rules.keywords,
      ai_6m_signals: rules.sixMSignals,
      ai_summary: requireString(offering.about_offering_text || solution?.about_solution_text).slice(0, 500),
      ai_engine: "rules_only",
      ai_enriched_at: nowIso,
      ai_enrichment_status: "rules_only",
      ai_prompt_version: aiPromptVersion,
      ai_schema_version: `${aiSchemaVersion}.offering`,
      ai_payload: {
        mode: "rules_only",
        reason: error instanceof Error ? error.message : "AI enrichment unavailable.",
      },
    };
  }

  const { error } = await adminClient.from("offerings").update(patch).eq("offering_id", requireString(offering.offering_id));
  if (error) throw new Error(error.message);
  return patch;
}

async function geocodeNeedLocation(input: {
  organization_name?: string;
  district?: string;
  state?: string;
}) {
  const query = [
    requireString(input.organization_name),
    requireString(input.district),
    requireString(input.state),
    "India",
  ].filter(Boolean).join(", ");

  if (!query) {
    return {
      latitude: null,
      longitude: null,
      geocoded_label: "",
      geocode_status: "location_missing",
    };
  }

  const parseCoordinate = (value: unknown) => {
    if (value === null || value === undefined || value === "") return null;
    const parsed = typeof value === "number" ? value : Number(value);
    return Number.isFinite(parsed) ? parsed : null;
  };

  const fallbackGeocode = async () => {
    const response = await fetch(
      `https://nominatim.openstreetmap.org/search?format=jsonv2&limit=1&q=${encodeURIComponent(query)}`,
      {
        headers: {
          Accept: "application/json",
          "User-Agent": "gre-mis-dashboard/1.0",
        },
      },
    );
    const data = await response.json().catch(() => null);
    if (!response.ok) {
      return {
        latitude: null,
        longitude: null,
        geocoded_label: "",
        geocode_status: "fallback_search_failed",
      };
    }
    const first = Array.isArray(data) ? data[0] : null;
    const latitude = parseCoordinate(first?.lat);
    const longitude = parseCoordinate(first?.lon);
    const label = requireString(first?.display_name || query);
    return {
      latitude,
      longitude,
      geocoded_label: label,
      geocode_status: latitude !== null && longitude !== null ? "ready_fallback" : "fallback_not_found",
    };
  };

  if (!mapplsAccessToken) {
    return await fallbackGeocode();
  }

  const response = await fetch(
    `https://search.mappls.com/search/places/textsearch/json?query=${encodeURIComponent(query)}&region=IND&access_token=${encodeURIComponent(mapplsAccessToken)}`,
  );
  const data = await response.json().catch(() => null);
  if (!response.ok) {
    return await fallbackGeocode();
  }

  const candidates = data?.suggestedLocations || data?.copResults || [];
  const first = Array.isArray(candidates) ? candidates[0] : candidates;
  const latitude = parseCoordinate(first?.latitude);
  const longitude = parseCoordinate(first?.longitude);
  const label = requireString(first?.placeAddress || first?.formattedAddress || query);
  if (latitude === null || longitude === null) {
    return await fallbackGeocode();
  }
  return {
    latitude,
    longitude,
    geocoded_label: label,
    geocode_status: "ready",
  };
}

async function hashToken(token: string) {
  const bytes = new TextEncoder().encode(token);
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  return Array.from(new Uint8Array(digest))
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
}

function generateToken() {
  const bytes = crypto.getRandomValues(new Uint8Array(32));
  return Array.from(bytes).map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

type GmailMailbox = "default" | "help" | "solution";

function getGmailMailboxConfig(mailbox: GmailMailbox = "default") {
  if (mailbox === "help" && helpGmailRefreshToken) {
    return {
      clientId: helpGmailClientId,
      clientSecret: helpGmailClientSecret,
      refreshToken: helpGmailRefreshToken,
      senderEmail: helpGmailSenderEmail,
    };
  }
  if (mailbox === "solution" && solutionGmailRefreshToken) {
    return {
      clientId: solutionGmailClientId,
      clientSecret: solutionGmailClientSecret,
      refreshToken: solutionGmailRefreshToken,
      senderEmail: solutionGmailSenderEmail,
    };
  }
  return {
    clientId: gmailClientId,
    clientSecret: gmailClientSecret,
    refreshToken: gmailRefreshToken,
    senderEmail: gmailSenderEmail,
  };
}

async function getGmailAccessToken(mailbox: GmailMailbox = "default") {
  const config = getGmailMailboxConfig(mailbox);
  const response = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: config.clientId,
      client_secret: config.clientSecret,
      refresh_token: config.refreshToken,
      grant_type: "refresh_token",
    }),
  });
  const data = await response.json().catch(() => null);
  if (!response.ok || !data?.access_token) {
    throw new Error(data?.error_description || data?.error || "Could not refresh Gmail access token.");
  }
  return String(data.access_token);
}

function toBase64Url(input: string) {
  return bytesToBase64(new TextEncoder().encode(input)).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

function wrapBase64Lines(value: string, lineLength = 76) {
  const lines: string[] = [];
  for (let index = 0; index < value.length; index += lineLength) {
    lines.push(value.slice(index, index + lineLength));
  }
  return lines.join("\r\n");
}

function encodeMimeHeader(value: string) {
  const text = requireString(value);
  if (!text) return "";
  if (/^[\x00-\x7F]*$/.test(text)) return text;
  return `=?UTF-8?B?${bytesToBase64(new TextEncoder().encode(text))}?=`;
}

function encodeMimeBody(value: string) {
  return wrapBase64Lines(bytesToBase64(new TextEncoder().encode(value)));
}

async function sendEmail({
  to,
  cc,
  subject,
  body,
  mailbox = "default",
}: {
  to: string;
  cc?: string;
  subject: string;
  body: string;
  mailbox?: GmailMailbox;
}) {
  const config = getGmailMailboxConfig(mailbox);
  const accessToken = await getGmailAccessToken(mailbox);
  const headers = [
    `From: ${config.senderEmail}`,
    `To: ${to}`,
    `Subject: ${encodeMimeHeader(subject)}`,
    "MIME-Version: 1.0",
    "Content-Type: text/plain; charset=UTF-8",
    "Content-Transfer-Encoding: base64",
  ];
  if (cc) headers.splice(2, 0, `Cc: ${cc}`);
  const rawMessage = `${headers.join("\r\n")}\r\n\r\n${encodeMimeBody(body)}`;

  const response = await fetch("https://gmail.googleapis.com/gmail/v1/users/me/messages/send", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ raw: toBase64Url(rawMessage) }),
  });
  const data = await response.json().catch(() => null);
  if (!response.ok) {
    throw new Error(data?.error?.message || "Gmail send failed.");
  }
  return data;
}

async function sendSolutionSubmissionConfirmationEmail(payload: Record<string, unknown>) {
  if (!solutionGmailRefreshToken) {
    throw new Error("Solution mailbox is not configured yet.");
  }
  const submitterEmail = requireString(payload.submitter_email || payload.submitterEmail).toLowerCase();
  if (!submitterEmail) return { ok: false, reason: "No submitter email provided." };

  const solutionProvider =
    requireString(payload.submitter_name || payload.submitterName) ||
    requireString(payload.contact_person || payload.contactPerson) ||
    "Solution Provider";
  const offeringName = requireString(payload.offering_name || payload.offeringName || payload.solution_name || payload.solutionName) || "your submitted solution";
  const cc = "tanmay@greenruraleconomy.in";
  const subject = `AskGRE solution received for ${offeringName}`;
  const body = [
    `Hello ${solutionProvider},`,
    "",
    `We have received your submitted Solution for ${offeringName}, on the AskGRE platform. This will be validated internally by our team and post validation get uploaded on the platform.`,
    "",
    "Thank you for sharing your Solutions with us and the wider ecosystem.",
    "",
    "Warm Regards,",
    "Team GRE",
  ].join("\n");

  await sendEmail({
    to: submitterEmail,
    cc,
    subject,
    body,
    mailbox: "solution",
  });

  await adminClient.from("gre_mis_email_log").insert({
    recipient_email: submitterEmail,
    cc_email: cc,
    subject,
    body_preview: body.slice(0, 1000),
    sent_by_email: solutionGmailSenderEmail,
  });

  return { ok: true };
}

async function sendNeedSubmissionConfirmationEmail(payload: Record<string, unknown>) {
  if (!helpGmailRefreshToken && !gmailRefreshToken) {
    throw new Error("Help mailbox is not configured yet.");
  }
  const seekerEmail = requireString(payload.seeker_email || payload.seekerEmail).toLowerCase();
  if (!seekerEmail) return { ok: false, reason: "No seeker email provided." };

  const seekerName = [
    payload.contact_person,
    payload.contactPerson,
    payload.contact_name,
    payload.contactName,
    payload.seeker_name,
    payload.seekerName,
    payload.submitter_name,
    payload.submitterName,
    payload.organization_name,
    payload.organizationName,
  ]
    .map((value) => requireString(value))
    .find(Boolean) || "Solution Seeker";
  const salutation = `Hello ${seekerName},`;
  const allowBroadcast = parseBoolean(payload.demand_broadcast_needed) ?? false;
  const broadcastLine = allowBroadcast ? "have" : "have not";
  const cc = ["help@greenruraleconomy.in", "tanmay@greenruraleconomy.in"].join(", ");
  const subject = "We have received your help request on AskGRE";
  const body = [
    salutation,
    "",
    "We have received your help request at our end and will shortly review it.",
    "",
    "Once we are able to understand the need, we will set up a call with you",
    "to fine tune the requirements and suggest possible solution providers",
    "for your need.",
    "",
    "In case we are unable to find one in our current network, we will also",
    "broadcast these needs to the wider ecosystem, if you so permit.",
    "",
    `We have noted that you ${broadcastLine} given us permission to broadcast this need`,
    "to the wider ecosystem.",
    "",
    "We thank you for reaching out to us.",
    "",
    "Regards,",
    "Team GRE",
  ].join("\n");
  await sendEmail({
    to: seekerEmail,
    cc,
    subject,
    body,
    mailbox: "help",
  });

  await adminClient.from("gre_mis_email_log").insert({
    recipient_email: seekerEmail,
    cc_email: cc,
    subject,
    body_preview: body.slice(0, 1000),
    sent_by_email: helpGmailRefreshToken ? helpGmailSenderEmail : gmailSenderEmail,
  });

  return { ok: true };
}

async function createUserSession(userId: string) {
  const token = generateToken();
  const tokenHash = await hashToken(token);
  const expiresAt = new Date(Date.now() + 1000 * 60 * 60 * 12).toISOString();

  const { error } = await adminClient.from("gre_mis_web_sessions").insert({
    user_id: userId,
    token_hash: tokenHash,
    expires_at: expiresAt,
  });
  if (error) throw new Error(error.message);

  return { token, expiresAt };
}

async function createLegacyAdminSession(username: string) {
  const token = generateToken();
  const tokenHash = await hashToken(token);
  const expiresAt = new Date(Date.now() + 1000 * 60 * 60 * 12).toISOString();

  const { error } = await adminClient.from("gre_mis_admin_web_sessions").insert({
    username,
    token_hash: tokenHash,
    expires_at: expiresAt,
  });
  if (error) throw new Error(error.message);
  return { token, expiresAt, username };
}

async function validateUserSession(token: string) {
  const normalized = requireString(token);
  if (!normalized) return null;
  const tokenHash = await hashToken(normalized);
  const { data, error } = await adminClient
    .from("gre_mis_web_sessions")
    .select("id, user_id, expires_at, gre_mis_users!inner(id, username, email, first_name, full_name, phone, role, is_active, must_change_password)")
    .eq("token_hash", tokenHash)
    .maybeSingle();

  if (error || !data) return null;

  if (new Date(data.expires_at).getTime() <= Date.now()) {
    await adminClient.from("gre_mis_web_sessions").delete().eq("id", data.id);
    return null;
  }

  await adminClient
    .from("gre_mis_web_sessions")
    .update({ last_used_at: new Date().toISOString() })
    .eq("id", data.id);

  const userRow = Array.isArray(data.gre_mis_users) ? data.gre_mis_users[0] : data.gre_mis_users;
  if (!userRow?.is_active) return null;
  return {
    id: data.id,
    user_id: data.user_id,
    expires_at: data.expires_at,
    user: userRow,
  };
}

async function validateLegacyAdminSession(token: string) {
  const normalized = requireString(token);
  if (!normalized) return null;
  const tokenHash = await hashToken(normalized);
  const { data, error } = await adminClient
    .from("gre_mis_admin_web_sessions")
    .select("id, username, expires_at")
    .eq("token_hash", tokenHash)
    .maybeSingle();
  if (error || !data) return null;
  if (new Date(data.expires_at).getTime() <= Date.now()) {
    await adminClient.from("gre_mis_admin_web_sessions").delete().eq("id", data.id);
    return null;
  }
  await adminClient
    .from("gre_mis_admin_web_sessions")
    .update({ last_used_at: new Date().toISOString() })
    .eq("id", data.id);
  return data;
}

async function requireUserSession(req: Request, bodyToken?: string) {
  const token =
    requireString(req.headers.get("x-gre-user-session")) ||
    requireString(req.headers.get("x-gre-admin-session")) ||
    requireString(bodyToken);
  const session = await validateUserSession(token);
  if (!session) throw new Error("Login required.");
  return { session, user: session.user, token };
}

async function requireUserSessionFromRequest(req: Request, payload: Record<string, unknown>, bodyToken?: string) {
  try {
    return await requireUserSession(req, bodyToken);
  } catch (sessionError) {
    const accessToken = requireString(payload.grameeeAccessToken);
    if (!accessToken) throw sessionError;
    const bridge = await bridgeGrameeeSession(accessToken);
    return await requireUserSession(req, requireString(bridge.token));
  }
}

function assertRoles(
  userCtx: { user: { role?: string } },
  roles: string[],
  message = "You do not have permission for this action.",
) {
  if (!roles.includes(requireString(userCtx.user.role))) {
    throw new Error(message);
  }
}

async function requireAdminSession(req: Request, bodyToken?: string) {
  const token =
    requireString(req.headers.get("x-gre-admin-session")) ||
    requireString(req.headers.get("x-gre-user-session")) ||
    requireString(bodyToken);
  try {
    const userCtx = await requireUserSession(req, bodyToken);
    assertRoles(userCtx, ["admin", "moderator"], "Admin or moderator login required.");
    return {
      ...userCtx,
      admin: {
        email: userCtx.user.email,
        display_name: userCtx.user.full_name || userCtx.user.first_name || userCtx.user.username,
      },
    };
  } catch {
    const legacySession = await validateLegacyAdminSession(token);
    if (!legacySession) throw new Error("Admin login required.");
    const { data: adminRow, error } = await adminClient
      .from("gre_mis_admins")
      .select("email, display_name")
      .limit(1)
      .maybeSingle();
    if (error || !adminRow) throw new Error("Admin record not found.");
    return {
      session: legacySession,
      user: { id: "", username: legacySession.username, email: adminRow.email, role: "admin" },
      admin: adminRow,
      token,
    };
  }
}

async function requireAdminSessionFromRequest(req: Request, payload: Record<string, unknown>, bodyToken?: string) {
  try {
    return await requireAdminSession(req, bodyToken);
  } catch (sessionError) {
    const accessToken = requireString(payload.grameeeAccessToken);
    if (!accessToken) throw sessionError;
    const bridge = await bridgeGrameeeSession(accessToken);
    const userCtx = await requireUserSession(req, requireString(bridge.token));
    assertRoles(userCtx, ["admin", "moderator"], "Admin or moderator login required.");
    return {
      ...userCtx,
      admin: {
        email: userCtx.user.email,
        display_name: userCtx.user.full_name || userCtx.user.first_name || userCtx.user.username,
      },
    };
  }
}

async function verifyUserPassword(identifier: string, password: string) {
  const { data, error } = await adminClient.rpc("gre_mis_user_password_matches", {
    p_identifier: identifier,
    p_password: password,
  });
  if (error) throw new Error(`Password verification failed: ${error.message}`);
  return Boolean(data);
}

async function verifyLegacyAdminPassword(username: string, password: string) {
  const { data, error } = await adminClient.rpc("grameee_admin_password_matches", {
    p_username: username,
    p_password: password,
  });
  if (error) throw new Error(`Admin password verification failed: ${error.message}`);
  return Boolean(data);
}

async function bridgeGrameeeSession(grameeeAccessToken: string) {
  const normalizedToken = requireString(grameeeAccessToken);
  if (!normalizedToken) {
    throw new Error("GramEEE login is required.");
  }

  const authUser = await resolveGrameeeAuthUser(normalizedToken);
  if (!authUser) {
    throw new Error("Your GramEEE login could not be verified. Please sign in again.");
  }

  const appMetadata = (authUser.app_metadata || {}) as Record<string, unknown>;
  const userMetadata = (authUser.user_metadata || {}) as Record<string, unknown>;
  const role = requireString(appMetadata.grameee_role).toLowerCase();

  if (!["admin", "moderator", "curator"].includes(role)) {
    throw new Error("This GramEEE login does not currently have GRE access.");
  }

  const email = requireString(authUser.email).toLowerCase();
  const username = requireString(userMetadata.username).toLowerCase();
  if (!email) {
    throw new Error("Your GramEEE login is missing an email address.");
  }

  const rawFullName = requireString(userMetadata.full_name) || username || email;
  const fullName = email === "tanmay@greenruraleconomy.in" && /admin/i.test(rawFullName)
    ? "Tanmay Mukherji"
    : rawFullName;
  const firstName = requireString(userMetadata.first_name) || fullName.split(/\s+/)[0] || fullName;
  const phone = requireString(userMetadata.phone);
  const syncedAt = new Date().toISOString();

  const { data: existingUser, error: existingError } = await adminClient
    .from("gre_mis_users")
    .select("id, username, email, first_name, full_name, phone, role, is_active, must_change_password")
    .or(`email.eq.${email},username.eq.${username || email}`)
    .limit(1)
    .maybeSingle();
  if (existingError) throw new Error(existingError.message);

  let greUser: Record<string, unknown> | null = existingUser as Record<string, unknown> | null;

  if (greUser) {
    const { data: updatedUser, error: updateError } = await adminClient
      .from("gre_mis_users")
      .update({
        username: username || requireString(greUser.username) || email,
        email,
        first_name: firstName,
        full_name: fullName,
        phone: phone || null,
        role,
        is_active: true,
        updated_at: syncedAt,
      })
      .eq("id", requireString(greUser.id))
      .select("id, username, email, first_name, full_name, phone, role, is_active, must_change_password")
      .single();
    if (updateError || !updatedUser) throw new Error(updateError?.message || "GRE MIS user could not be updated.");
    greUser = updatedUser as Record<string, unknown>;
  } else {
    const { data: insertedUser, error: insertError } = await adminClient
      .from("gre_mis_users")
      .insert({
        username: username || email,
        email,
        first_name: firstName,
        full_name: fullName,
        phone: phone || null,
        role,
        is_active: true,
        must_change_password: false,
        password_hash: "placeholder",
      })
      .select("id, username, email, first_name, full_name, phone, role, is_active, must_change_password")
      .single();
    if (insertError || !insertedUser) throw new Error(insertError?.message || "GRE MIS user could not be created.");
    greUser = insertedUser as Record<string, unknown>;
  }

  await syncRoleArtifactsForUser(greUser, role, "GRE MIS access refreshed from GramEEE session.");

  const session = await createUserSession(requireString(greUser.id));
  await adminClient
    .from("gre_mis_users")
    .update({ last_login_at: syncedAt })
    .eq("id", requireString(greUser.id));

  const { data: refreshedUser, error: refreshedError } = await adminClient
    .from("gre_mis_users")
    .select("id, username, email, first_name, full_name, phone, role, is_active, must_change_password")
    .eq("id", requireString(greUser.id))
    .single();
  if (refreshedError || !refreshedUser) throw new Error(refreshedError?.message || "GRE MIS user could not be refreshed.");

  return {
    ok: true,
    bridged: true,
    ...session,
    user: refreshedUser,
  };
}

async function ensureAdminUserFromLegacyPassword(password: string) {
  const { data: existingUser, error: existingError } = await adminClient
    .from("gre_mis_users")
    .select("id, username, email")
    .eq("username", "admin")
    .maybeSingle();
  if (existingError) throw new Error(existingError.message);

  let userId = existingUser?.id || null;
  if (!userId) {
    const { data: insertedUser, error: insertError } = await adminClient
      .from("gre_mis_users")
      .insert({
        username: "admin",
        first_name: "Tanmay",
        full_name: "Tanmay Mukherji",
        email: "tanmay@greenruraleconomy.in",
        role: "admin",
        is_active: true,
        must_change_password: false,
        password_hash: "placeholder",
      })
      .select("id")
      .single();
    if (insertError) throw new Error(insertError.message);
    userId = insertedUser.id;
  }

  const { error: passwordError } = await adminClient.rpc("gre_mis_user_set_password", {
    p_user_id: userId,
    p_password: password,
    p_must_change_password: false,
  });
  if (passwordError) throw new Error(passwordError.message);

  await adminClient
    .from("gre_mis_users")
    .update({
      first_name: "Tanmay",
      full_name: "Tanmay Mukherji",
      email: "tanmay@greenruraleconomy.in",
      role: "admin",
      is_active: true,
      updated_at: new Date().toISOString(),
    })
    .eq("id", userId);

  const { data: refreshedUser, error: refreshedError } = await adminClient
    .from("gre_mis_users")
    .select("id, username, email, first_name, full_name, phone, role, is_active, must_change_password")
    .eq("id", userId)
    .single();
  if (refreshedError || !refreshedUser) throw new Error(refreshedError?.message || "Admin user could not be refreshed.");
  return refreshedUser;
}

async function userLogin(identifier: string, password: string, requiredRole?: string) {
  if (!password) throw new Error("Password is required.");
  const normalizedIdentifier = requireString(identifier).toLowerCase();
  if (!normalizedIdentifier) throw new Error("Username or email is required.");
  let valid = await verifyUserPassword(normalizedIdentifier, password);
  let usedLegacyAdminFallback = false;
  if (!valid && normalizedIdentifier === "admin") {
    const legacyValid = await verifyLegacyAdminPassword("admin", password).catch(() => false);
    if (legacyValid) {
      await ensureAdminUserFromLegacyPassword(password);
      valid = true;
      usedLegacyAdminFallback = true;
    }
  }
  if (!valid) throw new Error("Incorrect username or password.");

  const { data: user, error } = await adminClient
    .from("gre_mis_users")
    .select("id, username, email, first_name, full_name, phone, role, is_active, must_change_password")
    .or(`username.eq.${normalizedIdentifier},email.eq.${normalizedIdentifier}`)
    .maybeSingle();
  if (error || !user) throw new Error(error?.message || "User record not found.");
  if (!user.is_active) throw new Error("This account is inactive.");
  if (requiredRole && user.role !== requiredRole) throw new Error(`${requiredRole} access is required.`);

  const session = await createUserSession(user.id);
  await adminClient.from("gre_mis_users").update({ last_login_at: new Date().toISOString() }).eq("id", user.id);
  return { ok: true, ...session, user, usedLegacyAdminFallback };
}

async function userLogout(token: string) {
  const normalized = requireString(token);
  if (!normalized) return { ok: true };
  const tokenHash = await hashToken(normalized);
  await adminClient.from("gre_mis_web_sessions").delete().eq("token_hash", tokenHash);
  return { ok: true };
}

async function adminLogin(username: string, password: string) {
  try {
    const result = await userLogin(requireString(username) || "admin", password, "admin");
    return {
      ok: true,
      token: result.token,
      expiresAt: result.expiresAt,
      username: result.user.username,
    };
  } catch (error) {
    const normalizedUsername = requireString(username) || "admin";
    const valid = await verifyLegacyAdminPassword(normalizedUsername, password);
    if (!valid) throw error instanceof Error ? error : new Error("Incorrect admin password.");
    return { ok: true, ...(await createLegacyAdminSession(normalizedUsername)) };
  }
}

async function adminLogout(token: string) {
  await userLogout(token);
  const normalized = requireString(token);
  if (normalized) {
    const tokenHash = await hashToken(normalized);
    await adminClient.from("gre_mis_admin_web_sessions").delete().eq("token_hash", tokenHash);
  }
  return { ok: true };
}

function isGreNeedsExtendedColumnError(error: unknown) {
  const message = [
    error instanceof Error ? error.message : "",
    typeof error === "object" && error ? requireString((error as Record<string, unknown>).message) : "",
    typeof error === "object" && error ? requireString((error as Record<string, unknown>).details) : "",
    typeof error === "object" && error ? requireString((error as Record<string, unknown>).hint) : "",
    String(error || ""),
  ].join(" ");
  return [
    "deployment_locations",
    "submitted_keywords",
    "submitted_thematic_area",
    "submitted_offering_category",
    "submitted_offering_type",
  ].some((column) => message.includes(`'${column}'`) || message.includes(`"${column}"`));
}

async function getAdminSnapshot() {
  const localNeedsQuery = async () => {
    return adminClient
      .from("gre_mis_needs")
      .select("*")
      .eq("approval_status", "approved")
      .eq("source_kind", "shared_form_submission")
      .order("requested_on", { ascending: false })
      .limit(500);
  };

  const [pendingNeeds, pendingUpdates, aiReviewNeeds, users, pendingFormSubmissions, approvedNeedSubmissions, localSolutions, localNeeds] = await Promise.allSettled([
    adminClient
      .from("gre_mis_needs")
      .select("id, organization_name, state, district, status, internal_status, requested_on, curator_id, problem_statement, source_kind")
      .eq("approval_status", "pending_admin")
      .order("requested_on", { ascending: false }),
    adminClient
      .from("gre_mis_update_requests")
      .select("*")
      .eq("approval_status", "pending")
      .order("created_at", { ascending: false }),
    adminClient
      .from("gre_mis_needs")
      .select("id, organization_name, state, district, problem_statement, curation_notes, curated_need, ai_thematic_area, ai_application_area, ai_need_kind, ai_service_kind, ai_keywords, ai_6m_signals, ai_validation_status, ai_validation_flags, ai_confidence, ai_enrichment_status, ai_engine, ai_enriched_at, rule_thematic_hints, rule_6m_signals, override_thematic_area, override_application_area, override_need_kind, override_service_kind, override_keywords, override_6m_signals, override_summary, override_source, override_conflict_note, override_updated_at, source_kind")
      .eq("approval_status", "approved")
      .or("ai_validation_status.is.null,ai_validation_status.eq.flagged,ai_enrichment_status.is.null")
      .order("updated_at", { ascending: false })
      .limit(60),
    adminClient
      .from("gre_mis_users")
      .select("id, username, first_name, full_name, email, phone, role, is_active, must_change_password, last_login_at, created_at, gre_user_id, gre_login_name, gre_sync_status, gre_sync_message, gre_synced_at, gre_pending_role, gre_activation_mod_key, gre_activation_requested_at")
      .order("role", { ascending: true })
      .order("first_name", { ascending: true }),
    adminClient
      .from("gre_mis_form_submissions")
      .select("*")
      .or("approval_status.eq.pending_admin,approval_status.is.null")
      .order("created_at", { ascending: false }),
    adminClient
      .from("gre_mis_form_submissions")
      .select("id, target_need_id, organization_name, payload")
      .eq("submission_type", "need")
      .eq("approval_status", "approved")
      .not("target_need_id", "is", null)
      .order("reviewed_at", { ascending: false })
      .limit(500),
    adminClient
      .from("offerings")
      .select(`
        offering_id,
        solution_id,
        trader_id,
        publish_status,
        offering_name,
        offering_category,
        offering_group,
        offering_type,
        tags,
        languages,
        geographies,
        about_offering_text,
        contact_details,
        raw_payload,
        trainer_name,
        trainer_email,
        trainer_phone,
        trainer_details_text,
        duration,
        prerequisites,
        location_availability,
        service_cost,
        product_cost,
        lead_time,
        support_details,
        support_post_service,
        support_post_service_cost,
        delivery_mode,
        certification_offered,
        cost_remarks,
        grade_capacity,
        service_brochure_url,
        product_brochure_url,
        knowledge_content_url,
        updated_at,
        solution:solutions (
          solution_id,
          solution_name,
          about_solution_text,
          solution_image_url
        ),
        trader:traders (
          trader_id,
          trader_name,
          organisation_name,
          email,
          mobile,
          poc_name,
          association_status
        )
      `)
      .eq("publish_status", "MIS Published")
      .like("offering_id", "MIS-OFFERING-%")
      .order("updated_at", { ascending: false })
      .limit(500),
    localNeedsQuery(),
  ]);

  const pendingNeedsResult = pendingNeeds.status === "fulfilled" && !pendingNeeds.value.error ? pendingNeeds.value.data || [] : [];
  const pendingUpdatesResult = pendingUpdates.status === "fulfilled" && !pendingUpdates.value.error ? pendingUpdates.value.data || [] : [];
  const aiReviewNeedsResult = aiReviewNeeds.status === "fulfilled" && !aiReviewNeeds.value.error ? aiReviewNeeds.value.data || [] : [];
  const usersResult = users.status === "fulfilled" && !users.value.error ? users.value.data || [] : [];
  const pendingFormSubmissionsResult = pendingFormSubmissions.status === "fulfilled" && !pendingFormSubmissions.value.error ? pendingFormSubmissions.value.data || [] : [];
  const approvedNeedSubmissionsResult = approvedNeedSubmissions.status === "fulfilled" && !approvedNeedSubmissions.value.error ? approvedNeedSubmissions.value.data || [] : [];
  const localSolutionsResult = localSolutions.status === "fulfilled" && !localSolutions.value.error ? localSolutions.value.data || [] : [];
  const localNeedsResult = localNeeds.status === "fulfilled" && !localNeeds.value.error ? localNeeds.value.data || [] : [];
  const approvedNeedSubmissionMap = new Map(
    approvedNeedSubmissionsResult
      .map((row) => {
        const record = row as Record<string, unknown>;
        const targetNeedId = requireString(record.target_need_id);
        if (!targetNeedId) return null;
        return [targetNeedId, record] as const;
      })
      .filter(Boolean) as readonly (readonly [string, Record<string, unknown>])[],
  );
  const hydratedLocalNeeds = localNeedsResult.map((row) => {
    const needRecord = { ...(row as Record<string, unknown>) };
    const linkedSubmission = approvedNeedSubmissionMap.get(requireString(needRecord.id));
    const payload = linkedSubmission?.payload && typeof linkedSubmission.payload === "object"
      ? linkedSubmission.payload as Record<string, unknown>
      : {};
    if (!requireString(needRecord.organization_name)) needRecord.organization_name = requireString(linkedSubmission?.organization_name || payload.organization_name);
    if (!requireString(needRecord.contact_person)) {
      needRecord.contact_person = requireString(payload.contact_person || payload.contact_name || payload.submitter_name);
    }
    if (!requireString(needRecord.seeker_email)) {
      needRecord.seeker_email = requireString(payload.seeker_email || payload.contact_email || payload.submitter_email).toLowerCase();
    }
    if (!requireString(needRecord.seeker_phone)) {
      needRecord.seeker_phone = requireString(payload.seeker_phone || payload.contact_phone || payload.submitter_phone);
    }
    if (!Array.isArray(needRecord.deployment_locations) || !needRecord.deployment_locations.length) {
      needRecord.deployment_locations = uniqueStrings(asStringArray(payload.deployment_locations));
    }
    if (!Array.isArray(needRecord.submitted_keywords) || !needRecord.submitted_keywords.length) {
      needRecord.submitted_keywords = uniqueStrings(asStringArray(payload.keywords));
    }
    if (!requireString(needRecord.submitted_thematic_area)) {
      needRecord.submitted_thematic_area = requireString(payload.thematic_area) || null;
    }
    if (!requireString(needRecord.submitted_offering_category)) {
      needRecord.submitted_offering_category = requireString(payload.offering_category) || null;
    }
    if (!requireString(needRecord.submitted_offering_type)) {
      needRecord.submitted_offering_type = requireString(payload.offering_type) || null;
    }
    return needRecord;
  });
  // Keep the dashboard snapshot as a pure read path.
  // GRE reconciliation and role sync belong on the explicit refresh action,
  // not on the page-load snapshot that powers Management Desk.
  const reconciledUsers = usersResult as Record<string, unknown>[];

  let greMailTemplates;
  try {
    greMailTemplates = await getGreMailTemplates();
  } catch (error) {
    console.error("GRE mail template load failed during admin snapshot", error);
    greMailTemplates = {
      providerIntroTemplate: DEFAULT_PROVIDER_INTRO_TEMPLATE,
      curatorForwardTemplate: DEFAULT_CURATOR_FORWARD_TEMPLATE,
      solutionSeekerTemplate: DEFAULT_SOLUTION_SEEKER_TEMPLATE,
      needSeekerTemplate: DEFAULT_NEED_SEEKER_TEMPLATE,
      inboundAutoSyncEnabled: true,
      lshContactEmails: ["subekkumar@pradan.net"],
      lshHelpCcEmails: ["help@greenruraleconomy.in"],
      lshRequestSupportTemplate: DEFAULT_LSH_MANAGEMENT_SETTINGS.requestSupportDraft,
      lshEmailProviderTemplate: DEFAULT_LSH_MANAGEMENT_SETTINGS.emailProviderDraft,
    };
  }

  let impactAuditLogs;
  try {
    impactAuditLogs = await getImpactAuditLogs();
  } catch (error) {
    console.error("Impact audit log load failed during admin snapshot", error);
    impactAuditLogs = { viewLogs: [], emailLogs: [] };
  }

  return {
    ok: true,
    pendingNeeds: pendingNeedsResult,
    pendingUpdates: pendingUpdatesResult,
    aiReviewNeeds: aiReviewNeedsResult.filter((row) => {
      const record = row as Record<string, unknown>;
      return requireString(record.source_kind) !== "shared_form_submission" && !requireString(record.id).startsWith("FORM-");
    }),
    users: reconciledUsers,
    mailTemplates: greMailTemplates,
    impactAuditLogs,
    pendingFormSubmissions: pendingFormSubmissionsResult,
    localSolutions: localSolutionsResult,
    localNeeds: hydratedLocalNeeds.filter((row) => {
      const record = row as Record<string, unknown>;
      return requireString(record.approval_status) === "approved"
        && (requireString(record.source_kind) === "shared_form_submission" || requireString(record.id).startsWith("FORM-"));
    }),
  };
}

async function getAdminUsersSnapshot() {
  const { data, error } = await adminClient
    .from("gre_mis_users")
    .select("id, username, first_name, full_name, email, phone, role, is_active, must_change_password, last_login_at, created_at, gre_user_id, gre_login_name, gre_sync_status, gre_sync_message, gre_synced_at, gre_pending_role, gre_activation_mod_key, gre_activation_requested_at")
    .order("role", { ascending: true })
    .order("first_name", { ascending: true });
  if (error) throw new Error(error.message);
  return { ok: true, users: data || [] };
}

async function getAdminMailImpactSnapshot() {
  let greMailTemplates;
  try {
    greMailTemplates = await getGreMailTemplates();
  } catch (error) {
    console.error("GRE mail template load failed during mail/impact snapshot", error);
    greMailTemplates = {
      providerIntroTemplate: DEFAULT_PROVIDER_INTRO_TEMPLATE,
      curatorForwardTemplate: DEFAULT_CURATOR_FORWARD_TEMPLATE,
      solutionSeekerTemplate: DEFAULT_SOLUTION_SEEKER_TEMPLATE,
      needSeekerTemplate: DEFAULT_NEED_SEEKER_TEMPLATE,
      inboundAutoSyncEnabled: true,
      lshContactEmails: ["subekkumar@pradan.net"],
      lshHelpCcEmails: ["help@greenruraleconomy.in"],
      lshRequestSupportTemplate: DEFAULT_LSH_MANAGEMENT_SETTINGS.requestSupportDraft,
      lshEmailProviderTemplate: DEFAULT_LSH_MANAGEMENT_SETTINGS.emailProviderDraft,
    };
  }

  let impactAuditLogs;
  try {
    impactAuditLogs = await getImpactAuditLogs();
  } catch (error) {
    console.error("Impact audit log load failed during mail/impact snapshot", error);
    impactAuditLogs = { viewLogs: [], emailLogs: [] };
  }

  return {
    ok: true,
    mailTemplates: greMailTemplates,
    impactAuditLogs,
  };
}

async function getAdminDataSyncSnapshot() {
  const { data, error } = await adminClient
    .from("gre_mis_needs")
    .select("id, organization_name, state, district, problem_statement, curation_notes, curated_need, ai_thematic_area, ai_application_area, ai_need_kind, ai_service_kind, ai_keywords, ai_6m_signals, ai_validation_status, ai_validation_flags, ai_confidence, ai_enrichment_status, ai_engine, ai_enriched_at, rule_thematic_hints, rule_6m_signals, override_thematic_area, override_application_area, override_need_kind, override_service_kind, override_keywords, override_6m_signals, override_summary, override_source, override_conflict_note, override_updated_at, source_kind")
    .eq("approval_status", "approved")
    .or("ai_validation_status.is.null,ai_validation_status.eq.flagged,ai_enrichment_status.is.null")
    .order("updated_at", { ascending: false })
    .limit(60);
  if (error) throw new Error(error.message);
  const templates = await getGreMailTemplates().catch(() => ({ inboundAutoSyncEnabled: true }));
  return {
    ok: true,
    aiReviewNeeds: (data || []).filter((row) => {
      const record = row as Record<string, unknown>;
      return requireString(record.source_kind) !== "shared_form_submission" && !requireString(record.id).startsWith("FORM-");
    }),
    inboundAutoSyncEnabled: typeof templates.inboundAutoSyncEnabled === "boolean" ? templates.inboundAutoSyncEnabled : true,
  };
}

async function getAdminApprovalsSnapshot() {
  const [pendingNeeds, pendingUpdates, pendingFormSubmissions] = await Promise.all([
    adminClient
      .from("gre_mis_needs")
      .select("id, organization_name, state, district, status, internal_status, requested_on, curator_id, problem_statement, source_kind")
      .eq("approval_status", "pending_admin")
      .order("requested_on", { ascending: false }),
    adminClient
      .from("gre_mis_update_requests")
      .select("*")
      .eq("approval_status", "pending")
      .order("created_at", { ascending: false }),
    adminClient
      .from("gre_mis_form_submissions")
      .select("id, submission_type, organization_name, source_mode, existing_trader_name, gre_sync_status, submitter_name, submitter_email, payload")
      .or("approval_status.eq.pending_admin,approval_status.is.null")
      .order("created_at", { ascending: false }),
  ]);
  if (pendingNeeds.error) throw new Error(pendingNeeds.error.message);
  if (pendingUpdates.error) throw new Error(pendingUpdates.error.message);
  if (pendingFormSubmissions.error) throw new Error(pendingFormSubmissions.error.message);
  return {
    ok: true,
    pendingNeeds: pendingNeeds.data || [],
    pendingUpdates: pendingUpdates.data || [],
    pendingFormSubmissions: pendingFormSubmissions.data || [],
  };
}

async function getAdminLocalSolutionsSnapshot() {
  const { data, error } = await adminClient
    .from("offerings")
    .select(`
      offering_id,
      solution_id,
      trader_id,
      publish_status,
      offering_name,
      offering_category,
      offering_group,
      offering_type,
      tags,
      languages,
      geographies,
      about_offering_text,
      updated_at,
      solution:solutions (
        solution_id,
        solution_name,
        about_solution_text,
        solution_image_url
      ),
      trader:traders (
        trader_id,
        trader_name,
        organisation_name,
        email,
        mobile,
        poc_name,
        association_status
      )
    `)
    .eq("publish_status", "MIS Published")
    .like("offering_id", "MIS-OFFERING-%")
    .order("updated_at", { ascending: false })
    .limit(500);
  if (error) throw new Error(error.message);
  return { ok: true, localSolutions: data || [] };
}

async function getAdminLocalNeedsSnapshot() {
  const [localNeeds, approvedNeedSubmissions] = await Promise.all([
    adminClient
      .from("gre_mis_needs")
      .select("*")
      .eq("approval_status", "approved")
      .eq("source_kind", "shared_form_submission")
      .order("requested_on", { ascending: false })
      .limit(500),
    adminClient
      .from("gre_mis_form_submissions")
      .select("id, target_need_id, organization_name, payload")
      .eq("submission_type", "need")
      .eq("approval_status", "approved")
      .not("target_need_id", "is", null)
      .order("reviewed_at", { ascending: false })
      .limit(500),
  ]);
  if (localNeeds.error) throw new Error(localNeeds.error.message);
  if (approvedNeedSubmissions.error) throw new Error(approvedNeedSubmissions.error.message);

  const approvedNeedSubmissionMap = new Map(
    (approvedNeedSubmissions.data || [])
      .map((row) => {
        const record = row as Record<string, unknown>;
        const targetNeedId = requireString(record.target_need_id);
        if (!targetNeedId) return null;
        return [targetNeedId, record] as const;
      })
      .filter(Boolean) as readonly (readonly [string, Record<string, unknown>])[],
  );

  const hydratedLocalNeeds = (localNeeds.data || []).map((row) => {
    const needRecord = { ...(row as Record<string, unknown>) };
    const linkedSubmission = approvedNeedSubmissionMap.get(requireString(needRecord.id));
    const payload = linkedSubmission?.payload && typeof linkedSubmission.payload === "object"
      ? linkedSubmission.payload as Record<string, unknown>
      : {};
    if (!requireString(needRecord.organization_name)) needRecord.organization_name = requireString(linkedSubmission?.organization_name || payload.organization_name);
    if (!requireString(needRecord.contact_person)) {
      needRecord.contact_person = requireString(payload.contact_person || payload.contact_name || payload.submitter_name);
    }
    if (!requireString(needRecord.seeker_email)) {
      needRecord.seeker_email = requireString(payload.seeker_email || payload.contact_email || payload.submitter_email).toLowerCase();
    }
    if (!requireString(needRecord.seeker_phone)) {
      needRecord.seeker_phone = requireString(payload.seeker_phone || payload.contact_phone || payload.submitter_phone);
    }
    if (!Array.isArray(needRecord.deployment_locations) || !needRecord.deployment_locations.length) {
      needRecord.deployment_locations = uniqueStrings(asStringArray(payload.deployment_locations));
    }
    if (!Array.isArray(needRecord.submitted_keywords) || !needRecord.submitted_keywords.length) {
      needRecord.submitted_keywords = uniqueStrings(asStringArray(payload.keywords));
    }
    if (!requireString(needRecord.submitted_thematic_area)) {
      needRecord.submitted_thematic_area = requireString(payload.thematic_area) || null;
    }
    if (!requireString(needRecord.submitted_offering_category)) {
      needRecord.submitted_offering_category = requireString(payload.offering_category) || null;
    }
    if (!requireString(needRecord.submitted_offering_type)) {
      needRecord.submitted_offering_type = requireString(payload.offering_type) || null;
    }
    return needRecord;
  });

  return {
    ok: true,
    localNeeds: hydratedLocalNeeds.filter((row) => {
      const record = row as Record<string, unknown>;
      return requireString(record.approval_status) === "approved"
        && (requireString(record.source_kind) === "shared_form_submission" || requireString(record.id).startsWith("FORM-"));
    }),
  };
}

async function refreshUserDirectory() {
  const { data: users, error } = await adminClient
    .from("gre_mis_users")
    .select("id, username, first_name, full_name, email, phone, role, is_active, must_change_password, last_login_at, created_at, gre_user_id, gre_login_name, gre_sync_status, gre_sync_message, gre_synced_at, gre_pending_role, gre_activation_mod_key, gre_activation_requested_at")
    .order("role", { ascending: true })
    .order("first_name", { ascending: true });
  if (error) throw new Error(error.message);
  const greProfiles = await fetchGreTenantUsers(true);
  const usersWithGreSeed = await ensureMissingGreDirectoryUsers((users || []) as Record<string, unknown>[], greProfiles);
  const reconciledUsers = await reconcileGreUserMappings(usersWithGreSeed);
  for (const refreshedUser of reconciledUsers) {
    const primaryRole = requireString(refreshedUser.role).toLowerCase() || "user";
    const greRoles = requireString((refreshedUser as Record<string, unknown>).__gre_roles)
      .split(",")
      .map((role) => role.trim().toLowerCase())
      .filter(Boolean);
    const syncMessage = requireString(refreshedUser.gre_sync_message) || "Role refreshed from GRE.";
    await syncRoleArtifactsForUser(refreshedUser, primaryRole, syncMessage);
    if (primaryRole !== "curator" && greRoles.includes("curator")) {
      await syncCanonicalCuratorRowForUser(
        refreshedUser,
        greRoles.includes("admin")
          ? "GRE refresh confirmed combined admin and curator access; curator allocations kept active."
          : syncMessage,
      );
    }
  }
  return {
    ok: true,
    users: reconciledUsers,
    message: "User roles and contact details refreshed from GRE.",
  };
}

async function registerUser(payload: Record<string, unknown>) {
  const firstName = requireString(payload.firstName);
  const fullName = requireString(payload.fullName) || firstName;
  const email = requireString(payload.email).toLowerCase();
  const phone = requireString(payload.phone);
  const password = requireString(payload.password);
  const usernameSource = requireString(payload.username) || firstName || email.split("@")[0];
  const username = usernameSource.toLowerCase().replace(/[^a-z0-9._-]/g, "");
  if (!firstName || !email || !password) throw new Error("Name, email, and password are required.");
  if (password.length < 8) throw new Error("Password must be at least 8 characters.");

  const { data: userId, error } = await adminClient.rpc("gre_mis_register_user", {
    p_username: username,
    p_first_name: firstName,
    p_full_name: fullName,
    p_email: email,
    p_phone: phone || null,
    p_password: password,
  });
  if (error) {
    const message = error.message.toLowerCase().includes("duplicate")
      ? "An account with this username or email already exists."
      : error.message;
    throw new Error(message);
  }
  return { ok: true, userId };
}

async function requestPasswordReset(email: string) {
  const normalizedEmail = requireString(email).toLowerCase();
  if (!normalizedEmail) throw new Error("Email is required.");
  const { data: user, error } = await adminClient
    .from("gre_mis_users")
    .select("id, first_name, email")
    .eq("email", normalizedEmail)
    .maybeSingle();
  if (error) throw new Error(error.message);
  if (!user) return { ok: true, message: "If the email exists, a reset code has been sent." };

  const code = String(Math.floor(100000 + Math.random() * 900000));
  const codeHash = await hashToken(code);
  const expiresAt = new Date(Date.now() + 1000 * 60 * 20).toISOString();

  await adminClient
    .from("gre_mis_password_reset_requests")
    .delete()
    .eq("email", normalizedEmail)
    .is("used_at", null);

  const { error: insertError } = await adminClient.from("gre_mis_password_reset_requests").insert({
    user_id: user.id,
    email: normalizedEmail,
    code_hash: codeHash,
    expires_at: expiresAt,
  });
  if (insertError) throw new Error(insertError.message);

  await sendEmail({
    to: normalizedEmail,
    subject: "GRE MIS password reset code",
    body: [
      `Hello ${user.first_name || "there"},`,
      "",
      `Your GRE MIS password reset code is: ${code}`,
      "This code will expire in 20 minutes.",
      "",
      "Regards,",
      "Green Rural Economy",
    ].join("\n"),
  });
  return { ok: true, message: "If the email exists, a reset code has been sent." };
}

async function resetPassword(email: string, code: string, newPassword: string) {
  const normalizedEmail = requireString(email).toLowerCase();
  const normalizedCode = requireString(code);
  if (!normalizedEmail || !normalizedCode || !newPassword) throw new Error("Email, code, and new password are required.");
  if (newPassword.length < 8) throw new Error("Password must be at least 8 characters.");

  const codeHash = await hashToken(normalizedCode);
  const { data: requestRow, error } = await adminClient
    .from("gre_mis_password_reset_requests")
    .select("id, user_id, expires_at, used_at")
    .eq("email", normalizedEmail)
    .eq("code_hash", codeHash)
    .is("used_at", null)
    .order("created_at", { ascending: false })
    .maybeSingle();
  if (error || !requestRow) throw new Error("Invalid or expired reset code.");
  if (new Date(requestRow.expires_at).getTime() <= Date.now()) {
    throw new Error("Reset code has expired.");
  }

  const { error: passwordError } = await adminClient.rpc("gre_mis_user_set_password", {
    p_user_id: requestRow.user_id,
    p_password: newPassword,
    p_must_change_password: false,
  });
  if (passwordError) throw new Error(passwordError.message);

  await adminClient
    .from("gre_mis_password_reset_requests")
    .update({ used_at: new Date().toISOString() })
    .eq("id", requestRow.id);

  return { ok: true };
}

async function changePassword(userCtx: { user: Record<string, unknown> }, currentPassword: string, newPassword: string) {
  if (!currentPassword || !newPassword) throw new Error("Current and new passwords are required.");
  if (newPassword.length < 8) throw new Error("Password must be at least 8 characters.");
  const identifier = requireString(userCtx.user.email) || requireString(userCtx.user.username);
  const valid = await verifyUserPassword(identifier, currentPassword);
  if (!valid) throw new Error("Current password is incorrect.");
  const { error } = await adminClient.rpc("gre_mis_user_set_password", {
    p_user_id: requireString(userCtx.user.id),
    p_password: newPassword,
    p_must_change_password: false,
  });
  if (error) throw new Error(error.message);
  return { ok: true };
}

async function syncCanonicalCuratorRowForUser(user: Record<string, unknown>, syncMessage: string) {
  const syncedAt = new Date().toISOString();
  const email = requireString(user.email).toLowerCase();
  const userId = requireString(user.id);
  const fullName = requireString(user.full_name) || requireString(user.first_name) || requireString(user.username);
  const firstName = requireString(user.first_name) || fullName.split(" ")[0] || fullName;

  const { data: curatorRows, error: curatorFetchError } = await adminClient
    .from("gre_mis_curators")
    .select("id, user_id, email, display_name, is_active, created_at, updated_at")
    .or(`user_id.eq.${userId},email.eq.${email}`);
  if (curatorFetchError) throw new Error(curatorFetchError.message);

  const rows = Array.isArray(curatorRows) ? curatorRows : [];
  const rowIds = rows.map((row) => requireString(row.id)).filter(Boolean);
  const needCounts = new Map<string, number>();

  if (rowIds.length) {
    const { data: linkedNeeds, error: needLinkError } = await adminClient
      .from("gre_mis_needs")
      .select("curator_id")
      .in("curator_id", rowIds);
    if (needLinkError) throw new Error(needLinkError.message);
    for (const linkedNeed of linkedNeeds || []) {
      const curatorId = requireString(linkedNeed.curator_id);
      if (!curatorId) continue;
      needCounts.set(curatorId, (needCounts.get(curatorId) || 0) + 1);
    }
  }

  const rankedRows = [...rows].sort((left, right) => {
    const leftMatch = requireString(left.email).toLowerCase() === email;
    const rightMatch = requireString(right.email).toLowerCase() === email;
    if (leftMatch !== rightMatch) return leftMatch ? -1 : 1;
    const leftCount = needCounts.get(requireString(left.id)) || 0;
    const rightCount = needCounts.get(requireString(right.id)) || 0;
    if (leftCount !== rightCount) return rightCount - leftCount;
    return new Date(requireString(left.created_at) || 0).getTime() - new Date(requireString(right.created_at) || 0).getTime();
  });

  const preferredRow = rankedRows[0] || null;
  const duplicateIds = rankedRows
    .slice(1)
    .map((row) => requireString(row.id))
    .filter(Boolean);

  if (duplicateIds.length) {
    const { error: duplicateError } = await adminClient
      .from("gre_mis_curators")
      .update({
        is_active: false,
        gre_sync_status: "synced",
        gre_sync_message: "Duplicate curator row archived during GRE sync reconciliation.",
        gre_synced_at: syncedAt,
        updated_at: syncedAt,
      })
      .in("id", duplicateIds);
    if (duplicateError) throw new Error(duplicateError.message);
  }

  if (preferredRow) {
    const { error: updateError } = await adminClient
      .from("gre_mis_curators")
      .update({
        user_id: userId,
        display_name: fullName,
        first_name: firstName,
        email,
        phone: user.phone,
        is_active: true,
        gre_sync_status: "synced",
        gre_sync_message: syncMessage,
        gre_synced_at: syncedAt,
        updated_at: syncedAt,
      })
      .eq("id", requireString(preferredRow.id));
    if (updateError) throw new Error(updateError.message);
    return requireString(preferredRow.id);
  }

  const { data: upsertedRow, error: upsertError } = await adminClient
    .from("gre_mis_curators")
    .upsert({
      user_id: userId,
      display_name: fullName,
      first_name: firstName,
      email,
      phone: user.phone,
      is_active: true,
      gre_sync_status: "synced",
      gre_sync_message: syncMessage,
      gre_synced_at: syncedAt,
      updated_at: new Date().toISOString(),
    }, { onConflict: "email" })
    .select("id")
    .single();
  if (upsertError) throw new Error(upsertError.message);
  return requireString(upsertedRow?.id);
}

async function persistMisRoleState(user: Record<string, unknown>, targetRole: string, greUserId: number | null, greLoginName: string, syncMessage: string) {
  const syncedAt = new Date().toISOString();
  const { error: userUpdateError } = await adminClient
    .from("gre_mis_users")
    .update({
      role: targetRole,
      gre_user_id: greUserId,
      gre_login_name: greLoginName || null,
      gre_sync_status: "synced",
      gre_sync_message: syncMessage,
      gre_synced_at: syncedAt,
      gre_pending_role: null,
      gre_activation_mod_key: null,
      gre_activation_requested_at: null,
      updated_at: syncedAt,
    })
    .eq("id", requireString(user.id));
  if (userUpdateError) throw new Error(userUpdateError.message);

  if (["admin", "moderator", "curator"].includes(targetRole)) {
    await syncCanonicalCuratorRowForUser(
      user,
      targetRole === "curator"
        ? "GRE curator role synced successfully."
        : `GRE role synced as ${targetRole}; curator allocations remain active on MIS.`,
    );
  } else {
    const { error: curatorError } = await adminClient
      .from("gre_mis_curators")
      .update({
        is_active: false,
        gre_sync_status: "synced",
        gre_sync_message: targetRole === "user" || targetRole === "moderator"
          ? "Curator access removed from GRE and MIS."
          : `GRE curator role removed. Current GRE role is ${targetRole}.`,
        gre_synced_at: syncedAt,
      })
      .eq("user_id", user.id);
    if (curatorError) throw new Error(curatorError.message);
  }

  if (targetRole === "admin") {
    const { error: adminError } = await adminClient.from("gre_mis_admins").upsert({
      email: requireString(user.email).toLowerCase(),
      display_name: user.full_name,
    });
    if (adminError) throw new Error(adminError.message);
  } else {
    const { error: adminError } = await adminClient
      .from("gre_mis_admins")
      .delete()
      .eq("email", requireString(user.email).toLowerCase());
    if (adminError) throw new Error(adminError.message);
  }
}

async function syncRoleArtifactsForUser(user: Record<string, unknown>, nextRole: string, syncMessage: string) {
  const syncedAt = new Date().toISOString();
  const email = requireString(user.email).toLowerCase();
  if (["admin", "moderator", "curator"].includes(nextRole)) {
    await syncCanonicalCuratorRowForUser(user, syncMessage);
  } else {
    const { error: curatorError } = await adminClient
      .from("gre_mis_curators")
      .update({
        is_active: false,
        gre_sync_status: "synced",
        gre_sync_message: syncMessage,
        gre_synced_at: syncedAt,
      })
      .or(`user_id.eq.${requireString(user.id)},email.eq.${email}`);
    if (curatorError) throw new Error(curatorError.message);
  }

  if (nextRole === "admin") {
    const { error: adminError } = await adminClient.from("gre_mis_admins").upsert({
      email,
      display_name: user.full_name,
    });
    if (adminError) throw new Error(adminError.message);
  } else {
    const { error: adminError } = await adminClient
      .from("gre_mis_admins")
      .delete()
      .eq("email", email);
    if (adminError) throw new Error(adminError.message);
  }
}

async function updateUserRole(userId: string, targetRoleInput: string) {
  const targetRole = requireString(targetRoleInput).toLowerCase();
  if (!["user", "curator", "moderator", "admin"].includes(targetRole)) throw new Error("Invalid role selected.");
  const greTargetRole = targetRole === "moderator" ? "user" : targetRole;

  const { data: user, error } = await adminClient
    .from("gre_mis_users")
    .select("id, username, first_name, full_name, email, phone, role, gre_user_id, gre_login_name, gre_sync_status, gre_pending_role, gre_activation_mod_key")
    .eq("id", userId)
    .single();
  if (error || !user) throw new Error(error?.message || "User not found.");

  if (user.role === targetRole) {
    return { ok: true, status: "unchanged", role: targetRole, message: `${user.full_name || user.username} is already ${targetRole}.` };
  }

  if (user.role === "admin" && targetRole !== "admin") {
    const { count, error: countError } = await adminClient
      .from("gre_mis_users")
      .select("id", { count: "exact", head: true })
      .eq("role", "admin")
      .eq("is_active", true)
      .neq("id", userId);
    if (countError) throw new Error(countError.message);
    if (!count) throw new Error("At least one admin must remain active in MIS.");
  }

  if (targetRole === "user" && user.role === "user") {
    return { ok: true, status: "unchanged", role: targetRole, message: `${user.full_name || user.username} is already user.` };
  }

  const greResolution = await resolveExistingGreUser(user);
  if ((targetRole === "user" || targetRole === "moderator") && (!greResolution.greUserId || greResolution.status !== "mapped")) {
    const { error: localOnlyError } = await adminClient
      .from("gre_mis_users")
      .update({
        role: targetRole,
        gre_pending_role: null,
        gre_activation_mod_key: null,
        gre_activation_requested_at: null,
        gre_sync_status: "local_only",
        gre_sync_message: "MIS user role updated locally. No GRE sync was required.",
        gre_synced_at: null,
        updated_at: new Date().toISOString(),
      })
      .eq("id", userId);
    if (localOnlyError) throw new Error(localOnlyError.message);
    await adminClient
      .from("gre_mis_curators")
      .update({
        is_active: false,
        gre_sync_status: "local_only",
        gre_sync_message: "Curator access removed in MIS. No GRE account mapping was available.",
        gre_synced_at: null,
      })
      .eq("user_id", userId);
    await adminClient
      .from("gre_mis_admins")
      .delete()
      .eq("email", requireString(user.email).toLowerCase());
    return {
      ok: true,
      status: "local_only",
      role: targetRole,
      message: `${user.full_name || user.username} is now a MIS ${targetRole} only.`,
    };
  }

  if ((targetRole === "curator" || targetRole === "admin") && (greResolution.status !== "mapped" || !greResolution.greUserId)) {
    if (greResolution.status === "needs_gre_account_creation") {
      const created = await createGreUserForRolePromotion(user, targetRole);
      await fetchGreTenantUsers(true);
      const refreshedResolution = await resolveExistingGreUser({
        ...user,
        gre_login_name: created.loginName || greResolution.greLoginName,
      });
      if (refreshedResolution.status === "mapped" && refreshedResolution.greUserId) {
        await syncGreRoleAssignment(refreshedResolution.greUserId, targetRole);
        await persistMisRoleState(
          user,
          targetRole,
          refreshedResolution.greUserId,
          refreshedResolution.greLoginName,
          `GRE account created and synced as ${targetRole}.`,
        );
        return {
          ok: true,
          status: "synced",
          role: targetRole,
          greUserId: refreshedResolution.greUserId,
          greLoginName: refreshedResolution.greLoginName,
          message: `${user.full_name || user.username} was created on GRE and synced as ${targetRole}.`,
        };
      }

      const identityState = await checkGreIdentity(created.loginName || greResolution.greLoginName || normalizePhoneLogin(user.phone));
      if (identityState === "INACTIVE_USER" && !created.modKey) {
        await activateInactiveGreUser(created.loginName || greResolution.greLoginName || normalizePhoneLogin(user.phone));
        await fetchGreTenantUsers(true);
        const activatedResolution = await resolveExistingGreUser({
          ...user,
          gre_login_name: created.loginName || greResolution.greLoginName,
        });
        if (activatedResolution.status === "mapped" && activatedResolution.greUserId) {
          await syncGreRoleAssignment(activatedResolution.greUserId, targetRole);
          await persistMisRoleState(
            user,
            targetRole,
            activatedResolution.greUserId,
            activatedResolution.greLoginName,
            `GRE account activated and synced as ${targetRole}.`,
          );
          return {
            ok: true,
            status: "synced",
            role: targetRole,
            greUserId: activatedResolution.greUserId,
            greLoginName: activatedResolution.greLoginName,
            message: `${user.full_name || user.username} was activated on GRE and synced as ${targetRole}.`,
          };
        }
      }

      if (identityState === "INACTIVE_USER" || created.modKey) {
        const activationUserId = created.greUserId || parseNumber(user.gre_user_id, 0) || null;
        await updateUserGreSyncFields(userId, {
          gre_user_id: activationUserId,
          gre_login_name: created.loginName || greResolution.greLoginName || normalizePhoneLogin(user.phone) || null,
          gre_sync_status: "awaiting_gre_activation",
          gre_sync_message: "GRE account created. Enter the OTP received on the user's phone to complete activation.",
          gre_synced_at: null,
          gre_pending_role: targetRole,
          gre_activation_mod_key: created.modKey,
          gre_activation_requested_at: new Date().toISOString(),
        });
        return {
          ok: true,
          status: "awaiting_gre_activation",
          role: user.role,
          pendingRole: targetRole,
          greUserId: activationUserId,
          greLoginName: created.loginName || greResolution.greLoginName,
          message: "GRE account created. OTP activation is pending.",
        };
      }
    }

    await updateUserGreSyncFields(userId, {
      gre_user_id: null,
      gre_login_name: greResolution.greLoginName || null,
      gre_sync_status: greResolution.status,
      gre_sync_message: greResolution.message,
      gre_synced_at: null,
      gre_pending_role: targetRole,
      gre_activation_mod_key: null,
      gre_activation_requested_at: null,
    });
    return {
      ok: true,
      status: greResolution.status,
      role: user.role,
      pendingRole: targetRole,
      greLoginName: greResolution.greLoginName,
      message: greResolution.message,
    };
  }

  await syncGreRoleAssignment(greResolution.greUserId, greTargetRole);
  await persistMisRoleState(
    user,
    targetRole,
    greResolution.greUserId,
    greResolution.greLoginName,
    `GRE role synced as ${targetRole}.`,
  );

  return {
    ok: true,
    status: "synced",
    role: targetRole,
    greUserId: greResolution.greUserId,
    greLoginName: greResolution.greLoginName,
    message: `${user.full_name || user.username} is now synced as ${targetRole} on MIS and GRE.`,
  };
}

async function completeUserRoleActivation(userId: string, otpInput: string) {
  const otp = requireString(otpInput);
  if (!otp) throw new Error("OTP is required.");

  const { data: user, error } = await adminClient
    .from("gre_mis_users")
    .select("id, username, first_name, full_name, email, phone, role, gre_user_id, gre_login_name, gre_pending_role, gre_activation_mod_key")
    .eq("id", userId)
    .single();
  if (error || !user) throw new Error(error?.message || "User not found.");

  const pendingRole = requireString(user.gre_pending_role).toLowerCase();
  if (!["curator", "admin"].includes(pendingRole)) {
    throw new Error("No pending GRE activation is currently stored for this user.");
  }

  const greUserId = parseNumber(user.gre_user_id, 0);
  const modKey = requireString(user.gre_activation_mod_key);
  if (!greUserId || !modKey) {
    throw new Error("GRE activation details are incomplete. Please retry the role update.");
  }

  await activateGreUserWithOtp(greUserId, otp, modKey);
  await fetchGreTenantUsers(true);
  const resolution = await resolveExistingGreUser(user);
  const resolvedGreUserId = resolution.greUserId || greUserId;
  const resolvedGreLogin = resolution.greLoginName || requireString(user.gre_login_name) || normalizePhoneLogin(user.phone);

  await syncGreRoleAssignment(resolvedGreUserId, pendingRole);
  await persistMisRoleState(
    user,
    pendingRole,
    resolvedGreUserId,
    resolvedGreLogin,
    `GRE account activated and synced as ${pendingRole}.`,
  );

  return {
    ok: true,
    status: "synced",
    role: pendingRole,
    greUserId: resolvedGreUserId,
    greLoginName: resolvedGreLogin,
    message: `${user.full_name || user.username} is now activated and synced as ${pendingRole} on GRE.`,
  };
}

async function removeManagedUser(userId: string, removalMode = "org_only") {
  const { data: user, error } = await adminClient
    .from("gre_mis_users")
    .select("id, username, first_name, full_name, email, phone, role, is_active, gre_user_id, gre_login_name")
    .eq("id", userId)
    .single();
  if (error || !user) throw new Error(error?.message || "User not found.");

  if (requireString(user.role) === "admin") {
    const { count, error: countError } = await adminClient
      .from("gre_mis_users")
      .select("id", { count: "exact", head: true })
      .eq("role", "admin")
      .eq("is_active", true)
      .neq("id", userId);
    if (countError) throw new Error(countError.message);
    if (!count) throw new Error("At least one admin must remain active in MIS.");
  }

  const greResolution = await resolveExistingGreUser(user);
  if (removalMode === "full_account") {
    if (!greResolution.greUserId) {
      throw new Error("This user is not mapped to a GRE account yet, so complete GRE account removal cannot be attempted.");
    }
    await attemptDeactivateGreWorkforceUser(greResolution.greUserId);
    await removeGreTenantRoles(greResolution.greUserId);
  } else if (greResolution.greUserId) {
    await removeGreTenantRoles(greResolution.greUserId);
  }

  const { data: curatorRows, error: curatorFetchError } = await adminClient
    .from("gre_mis_curators")
    .select("id")
    .or(`user_id.eq.${userId},email.eq.${requireString(user.email).toLowerCase()}`);
  if (curatorFetchError) throw new Error(curatorFetchError.message);

  const curatorIds = (Array.isArray(curatorRows) ? curatorRows : []).map((row) => requireString(row.id)).filter(Boolean);
  if (curatorIds.length) {
    const { error: needError } = await adminClient
      .from("gre_mis_needs")
      .update({ curator_id: null, updated_at: new Date().toISOString() })
      .in("curator_id", curatorIds);
    if (needError) throw new Error(needError.message);
  }

  const { error: curatorDeleteError } = await adminClient
    .from("gre_mis_curators")
    .delete()
    .or(`user_id.eq.${userId},email.eq.${requireString(user.email).toLowerCase()}`);
  if (curatorDeleteError) throw new Error(curatorDeleteError.message);

  const { error: adminDeleteError } = await adminClient
    .from("gre_mis_admins")
    .delete()
    .eq("email", requireString(user.email).toLowerCase());
  if (adminDeleteError) throw new Error(adminDeleteError.message);

  const { error: userDeleteError } = await adminClient
    .from("gre_mis_users")
    .delete()
    .eq("id", userId);
  if (userDeleteError) throw new Error(userDeleteError.message);

  await fetchGreTenantUsers(true);

  return {
    ok: true,
    removalMode,
    removedGreRoles: Boolean(greResolution.greUserId),
    message:
      removalMode === "full_account"
        ? `${user.full_name || user.username} was removed from MIS, their GRE organisation roles were cleared, and a GRE complete account removal attempt was submitted.`
        : `${user.full_name || user.username} was removed from MIS and their GRE organisation roles were cleared.`,
  };
}

function attachmentDataUrl(value: unknown) {
  if (value && typeof value === "object" && !Array.isArray(value)) {
    return requireString((value as Record<string, unknown>).dataUrl);
  }
  return "";
}

function attachmentName(value: unknown) {
  if (value && typeof value === "object" && !Array.isArray(value)) {
    return requireString((value as Record<string, unknown>).name);
  }
  return "";
}

function attachmentMimeType(value: unknown) {
  if (value && typeof value === "object" && !Array.isArray(value)) {
    return requireString((value as Record<string, unknown>).type);
  }
  return "";
}

function sanitizeFileName(value: string) {
  const cleaned = value.replace(/[^a-zA-Z0-9._-]+/g, "-").replace(/-+/g, "-").replace(/^-|-$/g, "");
  return cleaned || `asset-${Date.now()}`;
}

function parseDataUrl(value: string) {
  const match = value.match(/^data:([^;,]+)?(?:;charset=[^;,]+)?;base64,(.+)$/);
  if (!match) return null;
  return {
    mimeType: requireString(match[1]) || "application/octet-stream",
    base64: match[2],
  };
}

async function uploadAttachmentToGithub(
  attachment: unknown,
  folder: string,
  preferredName: string,
) {
  const dataUrl = attachmentDataUrl(attachment);
  if (!dataUrl) return null;
  const parsed = parseDataUrl(dataUrl);
  if (!parsed) return dataUrl;
  if (!githubAssetToken) return dataUrl;

  const sourceName = attachmentName(attachment) || preferredName;
  const safeName = sanitizeFileName(sourceName);
  const path = `${githubAssetRoot}/${folder}/${Date.now()}-${safeName}`;
  const response = await fetch(`https://api.github.com/repos/${githubAssetRepo}/contents/${encodeURIComponent(path).replace(/%2F/g, "/")}`, {
    method: "PUT",
    headers: {
      "Authorization": `Bearer ${githubAssetToken}`,
      "Accept": "application/vnd.github+json",
      "Content-Type": "application/json",
      "X-GitHub-Api-Version": "2022-11-28",
    },
    body: JSON.stringify({
      message: `Upload GRE MIS asset ${safeName}`,
      branch: githubAssetBranch,
      content: parsed.base64,
    }),
  });
  if (!response.ok) {
    const message = await response.text();
    throw new Error(`GitHub asset upload failed for ${sourceName}: ${message}`);
  }
  return `https://raw.githubusercontent.com/${githubAssetRepo}/${githubAssetBranch}/${path}`;
}

function buildLocalContactDetails(payload: Record<string, unknown>) {
  const explicit = requireString(payload.product_contact_details || payload.contact_details);
  if (explicit) return explicit;
  return [
    requireString(payload.submitter_name),
    requireString(payload.submitter_email),
    requireString(payload.submitter_phone),
  ].filter(Boolean).join(" | ");
}

async function resolveLocalTraderForSolution(payload: Record<string, unknown>, submissionId: string) {
  const requestedTraderId = requireString(payload.existing_trader_id);
  const organizationName = requireString(payload.organization_name);
  const submitterEmail = requireString(payload.submitter_email).toLowerCase();
  const submitterPhone = requireString(payload.submitter_phone);
  const submitterName = requireString(payload.submitter_name);
  const nowIso = new Date().toISOString();

  let trader: Record<string, unknown> | null = null;
  if (requestedTraderId) {
    const { data, error } = await adminClient
      .from("traders")
      .select("*")
      .eq("trader_id", requestedTraderId)
      .maybeSingle();
    if (error) throw new Error(error.message);
    trader = data;
  }

  if (trader) {
    const patch = {
      organisation_name: organizationName || requireString(trader.organisation_name || trader.trader_name),
      trader_name: requireString(trader.trader_name || organizationName),
      email: submitterEmail || requireString(trader.email),
      mobile: submitterPhone || requireString(trader.mobile),
      poc_name: submitterName || requireString(trader.poc_name),
      association_status: requireString(trader.association_status) || "Approved",
      raw_payload: {
        ...(trader.raw_payload && typeof trader.raw_payload === "object" ? trader.raw_payload as Record<string, unknown> : {}),
        mis_local_solution_contact_update: {
          submission_id: submissionId,
          organization_name: organizationName,
          submitter_name: submitterName,
          submitter_email: submitterEmail,
          submitter_phone: submitterPhone,
          updated_at: nowIso,
        },
      },
    };
    const { data: updated, error: updateError } = await adminClient
      .from("traders")
      .update(patch)
      .eq("trader_id", requireString(trader.trader_id))
      .select("*")
      .single();
    if (updateError) throw new Error(updateError.message);
    return updated as Record<string, unknown>;
  }

  const traderId = `MIS-TRADER-${crypto.randomUUID()}`;
  const row = {
    trader_id: traderId,
    trader_name: organizationName,
    organisation_name: organizationName,
    mobile: submitterPhone || null,
    email: submitterEmail || null,
    poc_name: submitterName || null,
    tenant_id: null,
    profile_id: null,
    description: null,
    short_description: null,
    tagline: null,
    website: null,
    created_at_source: nowIso,
    association_status: "MIS Local",
    raw_payload: {
      source: "gre_mis_local_solution_submission",
      submission_id: submissionId,
      organization_name: organizationName,
      submitter_name: submitterName,
      submitter_email: submitterEmail,
      submitter_phone: submitterPhone,
    },
  };
  const { data, error } = await adminClient.from("traders").insert(row).select("*").single();
  if (error) throw new Error(error.message);
  return data as Record<string, unknown>;
}

async function createLocalSolutionFromSubmission(
  submissionId: string,
  payload: Record<string, unknown>,
  actorEmail: string,
) {
  const nowIso = new Date().toISOString();
  const trader = await resolveLocalTraderForSolution(payload, submissionId);
  const traderId = requireString(trader.trader_id);
  const solutionId = `MIS-SOLUTION-${crypto.randomUUID()}`;
  const offeringId = `MIS-OFFERING-${crypto.randomUUID()}`;
  const offeringCategory = requireString(payload.offering_category) || "Service offerings";
  const offeringGroup = requireString(payload.offering_group) || offeringCategory.replace(/\s+offerings$/i, "").trim() || "Service";
  const offeringType = requireString(payload.offering_type);
  const offeringName = requireString(payload.offering_name);
  const offeringDescription = requireString(payload.about_offering_text);
  const offeringSearchDescription = requireString(payload.translated_about_offering_text_en || offeringDescription);
  const solutionName = requireString(payload.solution_name || offeringName);
  const solutionDescription = requireString(payload.about_solution_text || offeringDescription || offeringName);
  const solutionSearchDescription = requireString(payload.translated_about_solution_text_en || solutionDescription);
  const tags = uniqueStrings(asStringArray(payload.tags));
  const languages = normalizeLanguageArray(payload.languages);
  const geographies = uniqueStrings(asStringArray(payload.geographies));
  const locationAvailability = uniqueStrings(asStringArray(payload.location_availability));
  const contactDetails = buildLocalContactDetails(payload);
  const audience = "Individuals, Groups, SHGs, Organisations";
  const duration = [requireString(payload.duration), requireString(payload.duration_unit)].filter(Boolean).join(" ");
  const serviceCost = [requireString(payload.service_cost), requireString(payload.service_cost_unit)].filter(Boolean).join(" ");
  const productCost = parseBoolean(payload.product_cost_quote_on_scope) && !requireString(payload.product_cost)
    ? "Can be quoted after finalising scope"
    : requireString(payload.product_cost);
  const serviceBrochureUrl =
    await uploadAttachmentToGithub(payload.service_brochure_attachment, "service-brochures", `${offeringName || "service-offering"}-${attachmentName(payload.service_brochure_attachment) || "brochure"}`) ||
    attachmentDataUrl(payload.service_brochure_attachment);
  const productBrochureUrl =
    await uploadAttachmentToGithub(payload.product_brochure_attachment, "product-brochures", `${offeringName || "product-offering"}-${attachmentName(payload.product_brochure_attachment) || "brochure"}`) ||
    attachmentDataUrl(payload.product_brochure_attachment);
  const uploadedKnowledgeContent =
    await uploadAttachmentToGithub(payload.knowledge_content_attachment, "knowledge-content", `${offeringName || "knowledge-offering"}-${attachmentName(payload.knowledge_content_attachment) || "content"}`);
  const knowledgeContentUrl = uploadedKnowledgeContent || attachmentDataUrl(payload.knowledge_content_attachment) || requireString(payload.knowledge_content_url);
  const solutionImageUrl =
    await uploadAttachmentToGithub(payload.offering_image_attachment, "offering-images", `${offeringName || "offering"}-${attachmentName(payload.offering_image_attachment) || "image"}`) ||
    attachmentDataUrl(payload.offering_image_attachment);

  const rawPayload = {
    source: "gre_mis_local_solution_submission",
    submission_id: submissionId,
    approved_by: actorEmail,
    approved_at: nowIso,
    payload,
    files: {
      service_brochure_name: attachmentName(payload.service_brochure_attachment),
      product_brochure_name: attachmentName(payload.product_brochure_attachment),
      knowledge_content_name: attachmentName(payload.knowledge_content_attachment),
      offering_image_name: attachmentName(payload.offering_image_attachment),
    },
  };

  const solutionHtml = normalizeRichText(solutionDescription);
  const offeringHtml = normalizeRichText(offeringDescription);
  const offeringRules = classifyOfferingByRules({
    offering_name: offeringName,
    offering_group: offeringGroup,
    offering_type: offeringType,
    offering_category: offeringCategory,
    tags,
    languages,
    geographies,
    about_offering_text: offeringSearchDescription,
    contact_details: contactDetails,
  }, {
    solution_name: solutionName,
    about_solution_text: solutionSearchDescription,
  });

  const solutionRow = {
    solution_id: solutionId,
    trader_id: traderId,
    solution_name: solutionName,
    solution_status: "Approved in MIS",
    publish_status: "MIS Published",
    created_at_source: nowIso,
    about_solution_html: solutionHtml,
    about_solution_text: stripHtml(solutionHtml),
    solution_image_url: solutionImageUrl || null,
    raw_payload: rawPayload,
  };

  const offeringRow = {
    offering_id: offeringId,
    solution_id: solutionId,
    trader_id: traderId,
    publish_status: "MIS Published",
    created_at_source: nowIso,
    offering_name: offeringName,
    offering_category: offeringCategory,
    offering_group: offeringGroup,
    offering_type: offeringType,
    domain_6m: offeringRules.sixMSignals.join(", ") || null,
    primary_valuechain_id: null,
    primary_valuechain: null,
    primary_application_id: null,
    primary_application: null,
    valuechains: [] as string[],
    applications: [] as string[],
    tags,
    languages,
    geographies,
    geographies_raw: geographies.join("; "),
    about_offering_html: offeringHtml,
    about_offering_text: stripHtml(offeringHtml),
    audience,
    trainer_name: requireString(payload.trainer_name) || null,
    trainer_email: requireString(payload.trainer_email).toLowerCase() || null,
    trainer_phone: requireString(payload.trainer_phone) || null,
    trainer_details_html: normalizeRichText(payload.trainer_details_text),
    trainer_details_text: stripHtml(normalizeRichText(payload.trainer_details_text)),
    duration: duration || null,
    prerequisites: requireString(payload.prerequisites) || null,
    service_cost: serviceCost || null,
    support_post_service: requireString(payload.support_post_service) || null,
    support_post_service_cost: requireString(payload.support_post_service_cost) || null,
    delivery_mode: requireString(payload.delivery_mode) || null,
    certification_offered: requireString(payload.certification_offered) || null,
    cost_remarks: requireString(payload.cost_remarks) || null,
    location_availability: locationAvailability.join(", ") || null,
    service_brochure_url: serviceBrochureUrl || null,
    grade_capacity: requireString(payload.grade_capacity) || null,
    product_cost: productCost || null,
    lead_time: requireString(payload.lead_time) || null,
    support_details: requireString(payload.support_details) || null,
    product_brochure_url: productBrochureUrl || null,
    knowledge_content_url: knowledgeContentUrl || null,
    contact_details: contactDetails || null,
    gre_link: null,
    search_document: buildSearchDocument([
      solutionName,
      offeringName,
      offeringCategory,
      offeringGroup,
      offeringType,
      tags,
      languages,
      geographies,
      offeringDescription,
      offeringSearchDescription,
      solutionDescription,
      solutionSearchDescription,
      requireString(trader.organisation_name || trader.trader_name),
      contactDetails,
    ]),
    last_import_id: null,
    raw_payload: rawPayload,
    source_row_signature: stableRowSignature(rawPayload),
  };

  const { error: solutionError } = await adminClient.from("solutions").insert(solutionRow);
  if (solutionError) throw new Error(solutionError.message);

  const { error: offeringError } = await adminClient.from("offerings").insert(offeringRow);
  if (offeringError) throw new Error(offeringError.message);

  await enrichOfferingIntelligence(offeringRow, solutionRow, trader, defaultAiProvider || "openrouter");

  return {
    traderId,
    solutionId,
    offeringId,
  };
}

async function suggestSolutionTags(payload: Record<string, unknown>) {
  payload = await normalizeSubmissionTranslations("solution", payload);
  const sourceLanguage = requireString(payload.source_language || payload.sourceLanguage).toLowerCase();
  const normalized = {
    offering_name: requireString(payload.offering_name),
    offering_category: requireString(payload.offering_category),
    offering_type: requireString(payload.offering_type),
    about_offering_text: requireString(payload.translated_about_offering_text_en || payload.about_offering_text),
    trainer_details_text: requireString(payload.trainer_details_text),
    organization_name: requireString(payload.organization_name),
    contact_details: requireString(payload.contact_details),
  };
  const rules = classifyOfferingByRules({
    offering_name: normalized.offering_name,
    offering_category: normalized.offering_category,
    offering_group: normalized.offering_category.replace(/\s+offerings$/i, ""),
    offering_type: normalized.offering_type,
    about_offering_text: normalized.about_offering_text,
    contact_details: normalized.contact_details,
  }, {
    solution_name: normalized.offering_name,
    about_solution_text: normalized.about_offering_text,
  });
  const normalizeSolutionTag = (tag: unknown) =>
    requireString(tag)
      .toLowerCase()
      .replace(/[_/]+/g, " ")
      .replace(/\blow[\s-]+cost\b/g, "low cost")
      .replace(/\bdecentralised\b/g, "decentralized")
      .replace(/\s+/g, " ")
      .trim();
  const collapseSolutionTags = (values: unknown[]) => {
    const normalizedTags = uniqueStrings(
      values
        .map((entry) => normalizeSolutionTag(entry))
        .filter((entry) => !isWeakSolutionTag(entry)),
    );
    const phraseTags = normalizedTags.filter((tag) => tag.includes(" "));
    const compactTags = normalizedTags.filter((tag) => {
      const words = tag.split(" ").filter(Boolean);
      if (words.length > 1) return true;
      const singular = words[0]?.replace(/s$/i, "") || "";
      return !phraseTags.some((phrase) => {
        if (phrase === tag) return false;
        const phraseWords = phrase.split(" ").filter(Boolean);
        return phraseWords.includes(tag) || (singular && phraseWords.includes(singular));
      });
    });
    return compactTags.slice(0, 12);
  };
  const weakSolutionTags = new Set([
    "service",
    "services",
    "solution",
    "solutions",
    "support",
    "offering",
    "offerings",
    "provider",
    "product",
    "knowledge",
    "training",
  ]);
  const isWeakSolutionTag = (tag: unknown) => {
    const normalizedTag = normalizeSolutionTag(tag);
    if (!normalizedTag || normalizedTag.length < 3) return true;
    const parts = normalizedTag.split(" ").filter(Boolean);
    if (!parts.length) return true;
    if (weakSolutionTags.has(normalizedTag)) return true;
    if (parts.every((part) => weakSolutionTags.has(part) || /^\d+$/.test(part))) return true;
    return false;
  };
  let aiFallbackReason = "";
  try {
    const ai = await callAiJsonWithOrder(["gemini", "openai"], `You are helping classify a rural economy solution offering for AskGRE search discovery.

Return valid JSON only in this shape:
{
  "tags": ["", ""]
}

Rules:
- return 8 to 12 tags
- prefer phrase-level discovery tags, not raw token fragments
- prioritize use-case, value addition, beneficiary segment, deployment context, technical capability, commodity, and end-product language
- preserve meaningful multi-word phrases wherever possible
- prefer concrete outputs such as "mango processing", "fruit drying", "achar production", "juice processing", "food preservation"
- include sector/domain context when useful, such as "agricultural products", "rural livelihoods", "smallholder farmers"
- avoid generic filler like "service", "solution", "support", "offering", "provider", "product", "value", "addition", "processing" when they appear as weak standalone words
- avoid repeating single words from the description as separate tags when a better phrase exists
- each tag should be 1 to 4 words
- do not include numbering

Example:
Input offering name: Mango Value Addition
Input description: We offer services for Mango processing. Mainly wild mangoes. Drying and converting them into value added products like Achar, Amchur, Juices.
Good output:
{
  "tags": [
    "mango processing",
    "value addition",
    "wild mangoes",
    "fruit drying",
    "achar production",
    "amchur powder",
    "juice processing",
    "food preservation",
    "agricultural products",
    "rural livelihoods"
  ]
}

Organisation: ${normalized.organization_name}
Offering Category: ${normalized.offering_category}
Offering Type: ${normalized.offering_type}
Offering Name: ${normalized.offering_name}
Offering Description: ${normalized.about_offering_text}
Facilitator / Contact Notes: ${normalized.trainer_details_text || normalized.contact_details}`);
    const tags = collapseSolutionTags([
      ...asStringArray(ai.tags),
      ...buildSolutionTagDraft({
        offeringName: normalized.offering_name,
        offeringDescription: normalized.about_offering_text,
        organizationName: normalized.organization_name,
        rules,
      }),
    ]);
    if (tags.length) {
      const finalTags = isHindiSourceLanguage(sourceLanguage)
        ? await Promise.all(tags.map((tag) => translateTextWithLibreTranslate(tag, "en", "hi")))
        : tags;
      return {
        ok: true,
        tags: finalTags,
        message: "AI tags added. You can remove any tag before submitting.",
      };
    }
    aiFallbackReason = "AI returned no usable tags.";
  } catch (error) {
    aiFallbackReason = error instanceof Error ? error.message : String(error);
  }

  const tags = collapseSolutionTags(buildSolutionTagDraft({
    offeringName: normalized.offering_name,
    offeringDescription: [normalized.offering_type, normalized.about_offering_text].filter(Boolean).join(" | "),
    organizationName: normalized.organization_name,
    rules,
  }));
  const finalTags = isHindiSourceLanguage(sourceLanguage)
    ? await Promise.all(tags.map((tag) => translateTextWithLibreTranslate(tag, "en", "hi")))
    : tags;
  return {
    ok: true,
    tags: finalTags,
    message: `AI fallback used: ${aiFallbackReason || "No usable AI response."} A rule-based tag draft was added instead.`,
  };
}

async function suggestNeedTags(payload: Record<string, unknown>) {
  payload = await normalizeSubmissionTranslations("need", payload);
  const sourceLanguage = requireString(payload.source_language || payload.sourceLanguage).toLowerCase();
  const normalized = {
    organization_name: requireString(payload.organization_name),
    offering_category: requireString(payload.offering_category),
    offering_type: requireString(payload.offering_type),
    thematic_area: requireString(payload.thematic_area),
    problem_statement: requireString(payload.translated_problem_statement_en || payload.problem_statement),
    deployment_locations: asStringArray(payload.deployment_locations),
  };
  const rules = classifyNeedByRules({
    organization_name: normalized.organization_name,
    curated_need: [normalized.thematic_area].filter(Boolean),
    curation_notes: "",
    problem_statement: [
      normalized.offering_category,
      normalized.offering_type,
      normalized.thematic_area,
      normalized.problem_statement,
      normalized.deployment_locations.join(" "),
    ].filter(Boolean).join(" | "),
  });

  let aiFallbackReason = "";
  try {
    const ai = await callAiJson("gemini", `You are helping classify a rural economy need for solution matching.

Return valid JSON only in this shape:
{
  "tags": ["", ""]
}

Rules:
- give 8 to 12 tags
- prefer exact phrases from the problem statement when they express solution intent
- prioritize:
  1. core problem domain
  2. deployment context
  3. desired solution attributes
  4. beneficiary or operating context
- preserve strong multi-word phrases such as "street lighting", "low cost", "last mile delivery", "village level", "small habitations", "remote settlements"
- avoid generic filler, grammar fragments, or intro-statistics like "over", "demands", "raised", "service", "offerings", "have been"
- do not include broad labels like "solution" or "support" unless part of a meaningful phrase
- keep each tag to 1 to 4 words
- return plain tags only, not sentences

Organisation: ${normalized.organization_name}
Offering Category: ${normalized.offering_category}
Offering Type: ${normalized.offering_type}
Broad Need Group: ${normalized.thematic_area}
Place of Deployment: ${normalized.deployment_locations.join("; ")}
Problem Statement: ${normalized.problem_statement}`);
    const aiTags = asStringArray(ai.tags)
      .map((entry) => normalizeNeedTag(entry))
      .filter((entry) => !isWeakNeedTag(entry));
    const tags = uniqueStrings([
      ...aiTags,
      ...extractNeedPriorityPhrases(normalized.problem_statement).filter((entry) => !aiTags.includes(entry)),
    ]).slice(0, 12);
    if (tags.length) {
      const finalTags = isHindiSourceLanguage(sourceLanguage)
        ? await Promise.all(tags.map((tag) => translateTextWithLibreTranslate(tag, "en", "hi")))
        : tags;
      return {
        ok: true,
        tags: finalTags,
        message: "AI keywords added. You can remove any keyword before submitting.",
      };
    }
    aiFallbackReason = "AI returned no usable keywords.";
  } catch (error) {
    aiFallbackReason = error instanceof Error ? error.message : String(error);
  }

  const tags = buildNeedTagDraft({
    thematicArea: normalized.thematic_area,
    problemStatement: normalized.problem_statement,
    rules,
  });
  const finalTags = isHindiSourceLanguage(sourceLanguage)
    ? await Promise.all(tags.map((tag) => translateTextWithLibreTranslate(tag, "en", "hi")))
    : tags;
  return {
    ok: true,
    tags: finalTags,
    message: `AI fallback used: ${aiFallbackReason || "No usable AI response."} A rule-based keyword draft was added instead.`,
  };
}

async function submitFormSubmission(
  submissionType: string,
  payload: Record<string, unknown>,
  sourceMode = "shared_link",
  userCtx?: { user?: Record<string, unknown> | null } | null,
) {
  const normalizedType = requireString(submissionType).toLowerCase();
  if (!["need", "solution"].includes(normalizedType)) throw new Error("Unsupported submission type.");
  payload = await normalizeSubmissionTranslations(normalizedType, payload);
  const organizationName = requireString(payload.organization_name || payload.organizationName);
  const existingTraderId = requireString(payload.existing_trader_id || payload.existingTraderId);
  const existingTraderName = requireString(payload.existing_trader_name || payload.existingTraderName);
  const submitterName = requireString(payload.submitter_name || payload.submitterName || userCtx?.user?.full_name || userCtx?.user?.first_name);
  const submitterEmail = requireString(payload.submitter_email || payload.submitterEmail || userCtx?.user?.email).toLowerCase();
  const submitterPhone = requireString(payload.submitter_phone || payload.submitterPhone || userCtx?.user?.phone);
  const shareContext = requireString(payload.share_context || payload.shareContext);
  const demandBroadcastNeeded = parseBoolean(payload.demand_broadcast_needed) ?? false;

  if (!organizationName) throw new Error("Organization name is required.");
  let trader: Record<string, unknown> | null = null;
  if (normalizedType !== "solution") {
    if (existingTraderId) {
      const { data, error: traderError } = await adminClient
        .from("traders")
        .select("trader_id, organisation_name, trader_name")
        .eq("trader_id", existingTraderId)
        .maybeSingle();
      if (traderError) throw new Error(traderError.message);
      if (!data) throw new Error("Selected GRE supplier organization could not be found.");
      trader = data;
    }
  } else if (existingTraderId) {
    const { data, error: traderError } = await adminClient
      .from("traders")
      .select("trader_id, organisation_name, trader_name")
      .eq("trader_id", existingTraderId)
      .maybeSingle();
    if (traderError) throw new Error(traderError.message);
    trader = data;
  }

  const submissionRow = {
    submission_type: normalizedType,
    source_mode: sourceMode === "signed_in" ? "signed_in" : "shared_link",
    submitter_name: submitterName || null,
    submitter_email: submitterEmail || null,
    submitter_phone: submitterPhone || null,
    submitter_user_id: requireString(userCtx?.user?.id) || null,
    organization_name: organizationName,
    existing_trader_id: trader ? requireString(trader.trader_id) : existingTraderId || null,
    existing_trader_name: trader ? requireString(trader.organisation_name || trader.trader_name) : existingTraderName || null,
    org_exists_on_gre: Boolean(trader),
    payload,
    share_context: shareContext || null,
    approval_status: "pending_admin",
    gre_sync_status: "pending_admin_review",
    gre_sync_message: "Submission is waiting for admin review.",
  };

  const { data, error } = await adminClient
    .from("gre_mis_form_submissions")
    .insert(submissionRow)
    .select("id")
    .single();
  if (error) throw new Error(error.message);

  let message = normalizedType === "solution"
    ? "Solution Submitted for Approval"
    : "Need Help Submitted for Approval";
  if (normalizedType === "solution") {
    try {
      await sendSolutionSubmissionConfirmationEmail(payload);
    } catch (mailError) {
      const reason = requireString((mailError as { message?: string } | null)?.message) || "Unknown mail error.";
      message = `Solution Submitted for Approval. Confirmation email could not be sent: ${reason}`;
    }
  } else if (normalizedType === "need") {
    try {
      await sendNeedSubmissionConfirmationEmail({
        ...payload,
        demand_broadcast_needed: demandBroadcastNeeded,
      });
    } catch (mailError) {
      const reason = requireString((mailError as { message?: string } | null)?.message) || "Unknown mail error.";
      message = `Need Help Submitted for Approval. Confirmation email could not be sent: ${reason}`;
    }
  }

  return {
    ok: true,
    id: data.id,
    message,
  };
}

async function approveFormSubmission(submissionId: string, decision: string, reviewNotes: string, actorEmail: string) {
  const { data: submission, error } = await adminClient
    .from("gre_mis_form_submissions")
    .select("*")
    .eq("id", submissionId)
    .single();
  if (error || !submission) throw new Error(error?.message || "Form submission not found.");

  if (decision === "reject") {
    if (["solution", "need"].includes(requireString(submission.submission_type))) {
      const { error: deleteError } = await adminClient
        .from("gre_mis_form_submissions")
        .delete()
        .eq("id", submissionId);
      if (deleteError) throw new Error(deleteError.message);
      return { ok: true, deleted: true };
    }
    const { error: rejectError } = await adminClient
      .from("gre_mis_form_submissions")
      .update({
        approval_status: "rejected",
        admin_review_notes: reviewNotes || "Rejected by admin.",
        reviewed_by_email: actorEmail,
        reviewed_at: new Date().toISOString(),
        gre_sync_status: "rejected",
        gre_sync_message: reviewNotes || "Rejected by admin.",
      })
      .eq("id", submissionId);
    if (rejectError) throw new Error(rejectError.message);
    return { ok: true };
  }

  const payload = (submission.payload && typeof submission.payload === "object") ? submission.payload as Record<string, unknown> : {};

  if (submission.submission_type === "need") {
    const nextNeedId = `FORM-${Date.now()}`;
    const deploymentLocations = uniqueStrings(asStringArray(payload.deployment_locations));
    const submittedKeywords = uniqueStrings(asStringArray(payload.keywords));
    const curatedNeed = uniqueStrings([requireString(payload.thematic_area)].filter(Boolean));
    const submittedOfferingCategory = requireString(payload.offering_category);
    const submittedOfferingType = requireString(payload.offering_type);
    const normalizedNeedKind = submittedOfferingCategory.replace(/\s+offerings$/i, "").toLowerCase();
    const normalizedServiceKind = submittedOfferingCategory === "Service offerings" ? submittedOfferingType || null : null;
    const validationReady = Boolean(curatedNeed.length || submittedKeywords.length);
    const baseRow = {
      id: nextNeedId,
      organization_name: requireString(payload.organization_name || submission.organization_name),
      contact_person: requireString(payload.contact_person),
      seeker_email: requireString(payload.seeker_email).toLowerCase(),
      seeker_phone: requireString(payload.seeker_phone),
      state: requireString(payload.state),
      district: requireString(payload.district),
      problem_statement: requireString(payload.problem_statement),
      status: "New",
      internal_status: "Need solution providers",
      demand_broadcast_needed: parseBoolean(payload.demand_broadcast_needed) ?? false,
      curated_need: curatedNeed,
      approval_status: "approved",
      source_kind: "shared_form_submission",
      next_action: "Allocate curator and continue curation",
      requested_on: new Date().toISOString(),
      last_status_change_at: new Date().toISOString(),
      ai_thematic_area: requireString(payload.thematic_area) || null,
      ai_need_kind: normalizedNeedKind || null,
      ai_service_kind: normalizedServiceKind,
      ai_keywords: submittedKeywords,
      ai_engine: "local_form",
      ai_enriched_at: new Date().toISOString(),
      ai_enrichment_status: validationReady ? "ready" : "ready_flagged",
      ai_validation_status: validationReady ? "ready" : "flagged",
      ai_validation_flags: validationReady ? [] : ["needs_review"],
      ai_confidence: validationReady ? 82 : 46,
    };
    const extendedRow = {
      ...baseRow,
      deployment_locations: deploymentLocations,
      submitted_keywords: submittedKeywords,
      submitted_thematic_area: requireString(payload.thematic_area) || null,
      submitted_offering_category: submittedOfferingCategory || null,
      submitted_offering_type: submittedOfferingType || null,
    };
    const { error: insertError } = await adminClient.from("gre_mis_needs").insert(extendedRow);
    if (insertError) {
      if (!isGreNeedsExtendedColumnError(insertError)) throw new Error(insertError.message);
      const { error: legacyInsertError } = await adminClient.from("gre_mis_needs").insert(baseRow);
      if (legacyInsertError) throw new Error(legacyInsertError.message);
    }

    const { error: approveError } = await adminClient
      .from("gre_mis_form_submissions")
      .update({
        approval_status: "approved",
        admin_review_notes: reviewNotes || "Approved by admin.",
        reviewed_by_email: actorEmail,
        reviewed_at: new Date().toISOString(),
        target_need_id: nextNeedId,
        synced_to_gre: false,
        gre_sync_status: "synced_local_only",
        gre_sync_message: `Need was saved to local Supabase successfully as ${nextNeedId}. This record will remain MIS-local and continue through curation inside MIS.`,
      })
      .eq("id", submissionId);
    if (approveError) throw new Error(approveError.message);
    let questionsMessage = "";
    try {
      const questionResult = await generateSuggestedQuestionsForNeed(nextNeedId, actorEmail);
      questionsMessage = questionResult?.questions?.length
        ? ` Suggested questions to seeker were also prepared automatically.`
        : "";
    } catch (questionError) {
      questionsMessage = ` Suggested questions could not be prepared automatically: ${questionError instanceof Error ? questionError.message : String(questionError)}.`;
    }
    return {
      ok: true,
      targetNeedId: nextNeedId,
      message: `Need was saved to local Supabase successfully as ${nextNeedId}.${questionsMessage}`.trim(),
    };
  }

  try {
    const localWrite = await createLocalSolutionFromSubmission(submissionId, payload, actorEmail);
    const greSyncStatus = "synced_local_only";
    const greSyncMessage = `Solution and offering were saved to local Supabase successfully. Local Solution ID ${localWrite.solutionId}, Local Offering ID ${localWrite.offeringId}. These records will be discoverable in AskGRE with View Details only.`;

    const nextPayload = {
      ...payload,
      local_trader_id: localWrite.traderId,
      local_solution_id: localWrite.solutionId,
      local_offering_id: localWrite.offeringId,
    };
    const { error: approveError } = await adminClient
      .from("gre_mis_form_submissions")
      .update({
        approval_status: "approved",
        admin_review_notes: reviewNotes || "Approved by admin.",
        reviewed_by_email: actorEmail,
        reviewed_at: new Date().toISOString(),
        payload: nextPayload,
        target_solution_id: String(localWrite.solutionId),
        existing_trader_id: localWrite.traderId,
        synced_to_gre: false,
        gre_sync_status: greSyncStatus,
        gre_sync_message: greSyncMessage,
      })
      .eq("id", submissionId);
    if (approveError) throw new Error(approveError.message);
    await invalidateAskGreSearchCache();
    return {
      ok: true,
      greSyncStatus,
      message: greSyncMessage,
      solutionId: localWrite.solutionId,
      offeringId: localWrite.offeringId,
    };
  } catch (syncError) {
    const errorMessage = syncError instanceof Error ? syncError.message : String(syncError);
    const { error: patchError } = await adminClient
      .from("gre_mis_form_submissions")
      .update({
        approval_status: "pending_admin",
        admin_review_notes: reviewNotes || submission.admin_review_notes || null,
        gre_sync_status: "failed",
        gre_sync_message: errorMessage,
      })
      .eq("id", submissionId);
    if (patchError) throw new Error(`${errorMessage} (Also failed to store MIS sync failure: ${patchError.message})`);
    throw new Error(errorMessage);
  }
}

async function updateLocalNeed(needId: string, payload: Record<string, unknown>) {
  const { data: existing, error } = await adminClient
    .from("gre_mis_needs")
    .select("*")
    .eq("id", needId)
    .eq("source_kind", "shared_form_submission")
    .maybeSingle();
  if (error) throw new Error(error.message);
  if (!existing) throw new Error("Local need not found.");

  const submittedOfferingCategory = requireString(payload.submitted_offering_category || existing.submitted_offering_category);
  const submittedOfferingType = requireString(payload.submitted_offering_type || existing.submitted_offering_type);
  const submittedThematicArea = requireString(payload.submitted_thematic_area || existing.submitted_thematic_area);
  const submittedKeywords = uniqueStrings(asStringArray(payload.submitted_keywords || existing.submitted_keywords));
  const deploymentLocations = uniqueStrings(asStringArray(payload.deployment_locations || existing.deployment_locations));
  const stateFromDeployment = deploymentLocations.length
    ? deploymentLocations
      .map((entry) => entry.split(",").map((item) => requireString(item)).filter(Boolean))
      .map((parts) => parts.length >= 2 ? parts[parts.length - 2] : "")
      .find(Boolean)
    : "";
  const districtFromDeployment = deploymentLocations.length
    ? deploymentLocations
      .map((entry) => entry.split(",").map((item) => requireString(item)).filter(Boolean))
      .map((parts) => parts.length >= 3 ? parts[parts.length - 3] : parts[0] || "")
      .find(Boolean)
    : "";

  const basePatch = {
    organization_name: requireString(payload.organization_name || existing.organization_name),
    contact_person: requireString(payload.contact_person || existing.contact_person),
    seeker_email: requireString(payload.seeker_email || existing.seeker_email).toLowerCase(),
    seeker_phone: requireString(payload.seeker_phone || existing.seeker_phone),
    problem_statement: requireString(payload.problem_statement || existing.problem_statement),
    curated_need: uniqueStrings([submittedThematicArea].filter(Boolean)),
    state: stateFromDeployment || requireString(existing.state),
    district: districtFromDeployment || requireString(existing.district),
    updated_at: new Date().toISOString(),
  };
  const validationReady = Boolean(submittedThematicArea || submittedKeywords.length);

  const patch = {
    ...basePatch,
    deployment_locations: deploymentLocations,
    submitted_keywords: submittedKeywords,
    submitted_thematic_area: submittedThematicArea || null,
    submitted_offering_category: submittedOfferingCategory || null,
    submitted_offering_type: submittedOfferingType || null,
    ai_thematic_area: submittedThematicArea || null,
    ai_need_kind: submittedOfferingCategory.replace(/\s+offerings$/i, "").toLowerCase() || null,
    ai_service_kind: submittedOfferingCategory === "Service offerings" ? submittedOfferingType || null : null,
    ai_keywords: submittedKeywords,
    ai_engine: "local_form",
    ai_enriched_at: new Date().toISOString(),
    ai_enrichment_status: validationReady ? "ready" : "ready_flagged",
    ai_validation_status: validationReady ? "ready" : "flagged",
    ai_validation_flags: validationReady ? [] : ["needs_review"],
    ai_confidence: validationReady ? 82 : 46,
  };

  const { error: updateError } = await adminClient
    .from("gre_mis_needs")
    .update(patch)
    .eq("id", needId);
  if (updateError) {
    if (!isGreNeedsExtendedColumnError(updateError)) throw new Error(updateError.message);
    const legacyPatch = {
      ...basePatch,
      ai_thematic_area: submittedThematicArea || null,
      ai_need_kind: submittedOfferingCategory.replace(/\s+offerings$/i, "").toLowerCase() || null,
      ai_service_kind: submittedOfferingCategory === "Service offerings" ? submittedOfferingType || null : null,
      ai_keywords: submittedKeywords,
      ai_engine: "local_form",
      ai_enriched_at: new Date().toISOString(),
      ai_enrichment_status: validationReady ? "ready" : "ready_flagged",
      ai_validation_status: validationReady ? "ready" : "flagged",
      ai_validation_flags: validationReady ? [] : ["needs_review"],
      ai_confidence: validationReady ? 82 : 46,
    };
    const { error: legacyUpdateError } = await adminClient
      .from("gre_mis_needs")
      .update(legacyPatch)
      .eq("id", needId);
    if (legacyUpdateError) throw new Error(legacyUpdateError.message);
  }

  const { error: submissionUpdateError } = await adminClient
    .from("gre_mis_form_submissions")
    .update({
      payload: {
        ...(existing.raw_payload && typeof existing.raw_payload === "object" ? existing.raw_payload as Record<string, unknown> : {}),
        organization_name: patch.organization_name,
        contact_person: patch.contact_person,
        seeker_email: patch.seeker_email,
        seeker_phone: patch.seeker_phone,
        problem_statement: patch.problem_statement,
        deployment_locations: deploymentLocations,
        keywords: submittedKeywords,
        thematic_area: submittedThematicArea,
        offering_category: submittedOfferingCategory,
        offering_type: submittedOfferingType,
      },
      organization_name: patch.organization_name,
      existing_trader_name: patch.organization_name,
    })
    .eq("target_need_id", needId)
    .eq("submission_type", "need");
  if (submissionUpdateError) throw new Error(submissionUpdateError.message);

  try {
    await enrichNeedIntelligence({ ...existing, ...patch }, defaultAiProvider || "openrouter");
  } catch {}

  return { ok: true, message: "Local need updated successfully." };
}

async function deleteLocalNeed(needId: string) {
  const { data: need, error } = await adminClient
    .from("gre_mis_needs")
    .select("id, source_kind")
    .eq("id", needId)
    .eq("source_kind", "shared_form_submission")
    .maybeSingle();
  if (error) throw new Error(error.message);
  if (!need) throw new Error("Local need not found.");

  const { error: deleteNeedError } = await adminClient
    .from("gre_mis_needs")
    .delete()
    .eq("id", needId);
  if (deleteNeedError) throw new Error(deleteNeedError.message);

  await adminClient
    .from("gre_mis_form_submissions")
    .update({
      gre_sync_status: "deleted_local",
      gre_sync_message: "Local need deleted from MIS by admin.",
    })
    .eq("target_need_id", needId)
    .eq("submission_type", "need");

  return { ok: true, message: "Local need deleted successfully." };
}

async function updateFormSubmission(submissionId: string, update: Record<string, unknown>, actorEmail: string) {
  const { data: submission, error } = await adminClient
    .from("gre_mis_form_submissions")
    .select("*")
    .eq("id", submissionId)
    .single();
  if (error || !submission) throw new Error(error?.message || "Form submission not found.");
  if (requireString(submission.approval_status) !== "pending_admin") {
    throw new Error("Only pending submissions can be edited from the admin review queue.");
  }

  const payload = (update.payload && typeof update.payload === "object") ? update.payload as Record<string, unknown> : null;
  if (!payload) throw new Error("A valid submission payload is required.");

  const existingTraderId = requireString(update.existingTraderId || payload.existing_trader_id || submission.existing_trader_id);
  const isSolution = requireString(submission.submission_type) === "solution";
  let trader: Record<string, unknown> | null = null;
  if (existingTraderId) {
    const { data, error: traderError } = await adminClient
      .from("traders")
      .select("trader_id, trader_name, organisation_name")
      .eq("trader_id", existingTraderId)
      .maybeSingle();
    if (traderError) throw new Error(traderError.message);
    trader = data;
  }
  if (!trader && !isSolution) throw new Error("Please select a valid GRE supplier before saving the submission.");

  const organizationName =
    requireString(update.organizationName) ||
    requireString(payload.organization_name) ||
    requireString(trader?.organisation_name || trader?.trader_name);
  payload.organization_name = organizationName;
  payload.existing_trader_id = requireString(trader?.trader_id);
  payload.existing_trader_name = requireString(trader?.organisation_name || trader?.trader_name);

  const patch = {
    organization_name: organizationName,
    existing_trader_id: requireString(trader?.trader_id),
    existing_trader_name: requireString(trader?.organisation_name || trader?.trader_name),
    payload,
    admin_review_notes: requireString(update.adminReviewNotes) || submission.admin_review_notes || null,
  };

  const { error: updateError } = await adminClient
    .from("gre_mis_form_submissions")
    .update(patch)
    .eq("id", submissionId);
  if (updateError) throw new Error(updateError.message);

  return {
    ok: true,
    message: "Submission draft updated for admin review.",
  };
}

async function updateLocalSolution(offeringId: string, payload: Record<string, unknown>, actorEmail: string) {
  const normalizedOfferingId = requireString(offeringId);
  if (!normalizedOfferingId) throw new Error("Offering ID is required.");

  const { data: offering, error: offeringError } = await adminClient
    .from("offerings")
    .select("*, solution:solutions(*), trader:traders(*)")
    .eq("offering_id", normalizedOfferingId)
    .maybeSingle();
  if (offeringError) throw new Error(offeringError.message);
  if (!offering) throw new Error("Local solution offering not found.");
  if (requireString(offering.publish_status) !== "MIS Published") {
    throw new Error("Only locally managed MIS offerings can be edited here.");
  }

  const traderRow = Array.isArray(offering.trader) ? offering.trader[0] : offering.trader;
  const solutionRow = Array.isArray(offering.solution) ? offering.solution[0] : offering.solution;
  const offeringPayload = getStoredPayloadRecord(offering.raw_payload);
  const solutionPayload = getStoredPayloadRecord(solutionRow?.raw_payload);
  const originalPayload = Object.keys(offeringPayload).length ? offeringPayload : solutionPayload;
  const nowIso = new Date().toISOString();

  const organizationName = requireString(payload.organization_name || traderRow?.organisation_name || traderRow?.trader_name);
  const submitterName = requireString(payload.submitter_name || originalPayload.submitter_name || traderRow?.poc_name);
  const submitterEmail = requireString(payload.submitter_email || originalPayload.submitter_email || traderRow?.email).toLowerCase();
  const submitterPhone = requireString(payload.submitter_phone || originalPayload.submitter_phone || traderRow?.mobile);
  const offeringName = requireString(payload.offering_name || offering.offering_name);
  const offeringDescription = requireString(payload.about_offering_text || offering.about_offering_text);
  const offeringSearchDescription = requireString(payload.translated_about_offering_text_en || offeringDescription);
  const offeringCategory = requireString(payload.offering_category || offering.offering_category || "Service offerings");
  const offeringGroup = requireString(payload.offering_group || offering.offering_group || offeringCategory.replace(/\s+offerings$/i, "").trim() || "Service");
  const offeringType = requireString(payload.offering_type || offering.offering_type);
  const solutionName = requireString(payload.solution_name || offeringName || solutionRow?.solution_name);
  const solutionDescription = requireString(payload.about_solution_text || offeringDescription || solutionRow?.about_solution_text);
  const solutionSearchDescription = requireString(payload.translated_about_solution_text_en || solutionDescription);
  const tags = uniqueStrings(asStringArray(payload.tags || originalPayload.tags || offering.tags));
  if (!tags.length) throw new Error("At least one tag is required.");
  const languages = normalizeLanguageArray(payload.languages || originalPayload.languages || offering.languages);
  const geographies = uniqueStrings(asStringArray(payload.geographies || originalPayload.geographies || offering.geographies));
  const locationAvailability = uniqueStrings(asStringArray(payload.location_availability || originalPayload.location_availability)).join(", ") || requireString(offering.location_availability);
  const productCost = requireString(payload.product_cost || offering.product_cost);
  const serviceCost = requireString(payload.service_cost || offering.service_cost);
  const contactDetails = buildLocalContactDetails({
    ...payload,
    submitter_name: submitterName,
    submitter_email: submitterEmail,
    submitter_phone: submitterPhone,
      contact_details: payload.contact_details || payload.product_contact_details || originalPayload.contact_details || offering.contact_details,
  });
  const offeringHtml = normalizeRichText(offeringDescription);
  const solutionHtml = normalizeRichText(solutionDescription);
  const nextServiceBrochureUrl =
    await uploadAttachmentToGithub(payload.service_brochure_attachment, "service-brochures", `${offeringName || "service-offering"}-${attachmentName(payload.service_brochure_attachment) || "brochure"}`) ||
    attachmentDataUrl(payload.service_brochure_attachment) ||
    requireString(offering.service_brochure_url);
  const nextProductBrochureUrl =
    await uploadAttachmentToGithub(payload.product_brochure_attachment, "product-brochures", `${offeringName || "product-offering"}-${attachmentName(payload.product_brochure_attachment) || "brochure"}`) ||
    attachmentDataUrl(payload.product_brochure_attachment) ||
    requireString(offering.product_brochure_url);
  const nextKnowledgeContentUrl =
    await uploadAttachmentToGithub(payload.knowledge_content_attachment, "knowledge-content", `${offeringName || "knowledge-offering"}-${attachmentName(payload.knowledge_content_attachment) || "content"}`) ||
    attachmentDataUrl(payload.knowledge_content_attachment) ||
    requireString(payload.knowledge_content_url || offering.knowledge_content_url);
  const nextSolutionImageUrl =
    await uploadAttachmentToGithub(payload.offering_image_attachment, "offering-images", `${offeringName || "offering"}-${attachmentName(payload.offering_image_attachment) || "image"}`) ||
    attachmentDataUrl(payload.offering_image_attachment) ||
    requireString(payload.solution_image_url || solutionRow?.solution_image_url);
  const rawPayload = {
    ...(offering.raw_payload && typeof offering.raw_payload === "object" ? offering.raw_payload as Record<string, unknown> : {}),
    last_manual_edit: {
      updated_by: actorEmail,
      updated_at: nowIso,
      payload,
      manual_fields: [],
    },
  };

  const traderPatch = {
    organisation_name: organizationName || null,
    trader_name: requireString(traderRow?.trader_name || organizationName) || null,
    email: submitterEmail || null,
    mobile: submitterPhone || null,
    poc_name: submitterName || null,
  };
  const { error: traderUpdateError } = await adminClient
    .from("traders")
    .update(traderPatch)
    .eq("trader_id", requireString(offering.trader_id));
  if (traderUpdateError) throw new Error(traderUpdateError.message);

  const solutionPatch = {
    solution_name: solutionName,
    about_solution_html: solutionHtml,
    about_solution_text: stripHtml(solutionHtml),
    solution_image_url: nextSolutionImageUrl || null,
    raw_payload: rawPayload,
  };
  const { error: solutionUpdateError } = await adminClient
    .from("solutions")
    .update(solutionPatch)
    .eq("solution_id", requireString(offering.solution_id));
  if (solutionUpdateError) throw new Error(solutionUpdateError.message);

  const offeringPatch = {
    offering_name: offeringName,
    offering_category: offeringCategory,
    offering_group: offeringGroup,
    offering_type: offeringType,
    domain_6m: classifyOfferingByRules({
      offering_name: offeringName,
      offering_group: offeringGroup,
      offering_type: offeringType,
      offering_category: offeringCategory,
      tags,
      languages,
      geographies,
      about_offering_text: offeringSearchDescription,
      contact_details: contactDetails,
    }, {
      solution_name: solutionName,
      about_solution_text: solutionSearchDescription,
    }).sixMSignals.join(", ") || null,
    tags,
    languages,
    geographies,
    geographies_raw: geographies.join("; "),
    about_offering_html: offeringHtml,
    about_offering_text: stripHtml(offeringHtml),
    trainer_name: requireString(payload.trainer_name || originalPayload.trainer_name || offering.trainer_name) || null,
    trainer_email: requireString(payload.trainer_email || originalPayload.trainer_email || offering.trainer_email).toLowerCase() || null,
    trainer_phone: requireString(payload.trainer_phone || originalPayload.trainer_phone || offering.trainer_phone) || null,
    trainer_details_html: normalizeRichText(payload.trainer_details_text || originalPayload.trainer_details_text || offering.trainer_details_text),
    trainer_details_text: stripHtml(normalizeRichText(payload.trainer_details_text || originalPayload.trainer_details_text || offering.trainer_details_text)),
    duration: [requireString(payload.duration || originalPayload.duration), requireString(payload.duration_unit || originalPayload.duration_unit)].filter(Boolean).join(" ") || requireString(offering.duration) || null,
    prerequisites: requireString(payload.prerequisites || originalPayload.prerequisites || offering.prerequisites) || null,
    service_cost: serviceCost || null,
    support_post_service: requireString(payload.support_post_service || originalPayload.support_post_service || offering.support_post_service) || null,
    support_post_service_cost: requireString(payload.support_post_service_cost || originalPayload.support_post_service_cost || offering.support_post_service_cost) || null,
    delivery_mode: requireString(payload.delivery_mode || originalPayload.delivery_mode || offering.delivery_mode) || null,
    certification_offered: requireString(payload.certification_offered || originalPayload.certification_offered || offering.certification_offered) || null,
    cost_remarks: requireString(payload.cost_remarks || originalPayload.cost_remarks || offering.cost_remarks) || null,
    location_availability: locationAvailability || null,
    grade_capacity: requireString(payload.grade_capacity || originalPayload.grade_capacity || offering.grade_capacity) || null,
    product_cost: productCost || null,
    lead_time: requireString(payload.lead_time || originalPayload.lead_time || offering.lead_time) || null,
    support_details: requireString(payload.support_details || originalPayload.support_details || offering.support_details) || null,
    service_brochure_url: nextServiceBrochureUrl || null,
    product_brochure_url: nextProductBrochureUrl || null,
    knowledge_content_url: nextKnowledgeContentUrl || null,
    gre_link: requireString(payload.gre_link || offering.gre_link) || null,
    contact_details: contactDetails || requireString(originalPayload.contact_details || offering.contact_details) || null,
    search_document: buildSearchDocument([
      solutionName,
      offeringName,
      offeringCategory,
      offeringGroup,
      offeringType,
      tags,
      languages,
      geographies,
      offeringDescription,
      offeringSearchDescription,
      solutionDescription,
      solutionSearchDescription,
      organizationName,
      contactDetails,
    ]),
    raw_payload: rawPayload,
    source_row_signature: stableRowSignature(rawPayload),
  };
  const manualFields = Object.keys(offeringPatch).filter((key) => {
    if (key === "raw_payload" || key === "source_row_signature" || key === "search_document" || key === "domain_6m") return false;
    return String(offeringPatch[key as keyof typeof offeringPatch] ?? "") !== String(offering[key as keyof typeof offering] ?? "");
  });
  if (manualFields.length) {
    rawPayload.last_manual_edit = {
      ...(rawPayload.last_manual_edit as Record<string, unknown>),
      manual_fields: manualFields,
    };
    offeringPatch.source_row_signature = stableRowSignature(rawPayload);
  }
  const { error: offeringUpdateError } = await adminClient
    .from("offerings")
    .update(offeringPatch)
    .eq("offering_id", normalizedOfferingId);
  if (offeringUpdateError) throw new Error(offeringUpdateError.message);

  const refreshedOffering = {
    ...offering,
    ...offeringPatch,
    offering_id: normalizedOfferingId,
    solution_id: requireString(offering.solution_id),
    trader_id: requireString(offering.trader_id),
  };
  const refreshedSolution = {
    ...(solutionRow && typeof solutionRow === "object" ? solutionRow : {}),
    ...solutionPatch,
    solution_id: requireString(offering.solution_id),
    trader_id: requireString(offering.trader_id),
  };
  const refreshedTrader = {
    ...(traderRow && typeof traderRow === "object" ? traderRow : {}),
    ...traderPatch,
    trader_id: requireString(offering.trader_id),
  };
  await enrichOfferingIntelligence(refreshedOffering, refreshedSolution, refreshedTrader, defaultAiProvider || "openrouter");
  await invalidateAskGreSearchCache();

  return {
    ok: true,
    message: "Local solution updated in Supabase.",
  };
}

async function getLocalSolutionDetail(offeringId: string) {
  const normalizedOfferingId = requireString(offeringId);
  if (!normalizedOfferingId) throw new Error("Offering ID is required.");

  const { data: offering, error } = await adminClient
    .from("offerings")
    .select("*, solution:solutions(*), trader:traders(*)")
    .eq("offering_id", normalizedOfferingId)
    .maybeSingle();
  if (error) throw new Error(error.message);
  if (!offering) throw new Error("Local solution offering not found.");
  if (requireString(offering.publish_status) !== "MIS Published") {
    throw new Error("Only locally managed MIS offerings can be opened here.");
  }
  if (!normalizedOfferingId.startsWith("MIS-OFFERING-")) {
    throw new Error("This offering is not a local platform offering.");
  }

  return { ok: true, offering };
}

async function getLocalNeedDetail(needId: string) {
  const normalizedNeedId = requireString(needId);
  if (!normalizedNeedId) throw new Error("Need ID is required.");

  const { data: need, error } = await adminClient
    .from("gre_mis_needs")
    .select("*")
    .eq("id", normalizedNeedId)
    .maybeSingle();
  if (error) throw new Error(error.message);
  if (!need) throw new Error("Local need not found.");
  if (requireString(need.source_kind) !== "shared_form_submission") {
    throw new Error("This need is not a local platform need.");
  }

  return { ok: true, need };
}

async function deleteLocalSolution(offeringId: string) {
  const normalizedOfferingId = requireString(offeringId);
  if (!normalizedOfferingId) throw new Error("Offering ID is required.");
  const { data: offering, error } = await adminClient
    .from("offerings")
    .select("offering_id, solution_id, publish_status")
    .eq("offering_id", normalizedOfferingId)
    .maybeSingle();
  if (error) throw new Error(error.message);
  if (!offering) throw new Error("Local solution offering not found.");
  if (requireString(offering.publish_status) !== "MIS Published") {
    throw new Error("Only locally managed MIS offerings can be deleted here.");
  }

  const solutionId = requireString(offering.solution_id);
  const { error: deleteOfferingError } = await adminClient
    .from("offerings")
    .delete()
    .eq("offering_id", normalizedOfferingId);
  if (deleteOfferingError) throw new Error(deleteOfferingError.message);

  const { count, error: countError } = await adminClient
    .from("offerings")
    .select("offering_id", { count: "exact", head: true })
    .eq("solution_id", solutionId);
  if (countError) throw new Error(countError.message);
  if (!count) {
    const { error: deleteSolutionError } = await adminClient
      .from("solutions")
      .delete()
      .eq("solution_id", solutionId);
    if (deleteSolutionError) throw new Error(deleteSolutionError.message);
  }
  await invalidateAskGreSearchCache();

  return {
    ok: true,
    message: "Local solution deleted from Supabase.",
  };
}

async function applyNeedOverride(
  needId: string,
  patchInput: Record<string, unknown>,
  conflictNote: string,
  resolveConflict: boolean,
  actorEmail: string,
) {
  const patch: Record<string, unknown> = {};
  const allowed = new Set([
    "override_thematic_area",
    "override_application_area",
    "override_need_kind",
    "override_service_kind",
    "override_keywords",
    "override_6m_signals",
    "override_summary",
  ]);

  Object.entries(patchInput || {}).forEach(([key, value]) => {
    if (!allowed.has(key)) return;
    if (key === "override_keywords" || key === "override_6m_signals") {
      patch[key] = asStringArray(value);
    } else {
      patch[key] = requireString(value) || null;
    }
  });

  if (!Object.keys(patch).length) throw new Error("No override values were provided.");

  patch.override_source = "puter_review";
  patch.override_conflict_note = requireString(conflictNote) || null;
  patch.override_updated_at = new Date().toISOString();
  patch.override_updated_by = actorEmail;
  patch.ai_validation_status = resolveConflict ? "ready" : "flagged";
  patch.ai_validation_flags = resolveConflict ? [] : ["manual_override_partial"];
  patch.ai_enrichment_status = resolveConflict ? "ready" : "ready_flagged";

  const { error } = await adminClient
    .from("gre_mis_needs")
    .update(patch)
    .eq("id", needId);
  if (error) throw new Error(error.message);

  await adminClient.from("gre_mis_need_updates").insert({
    need_id: needId,
    update_type: "manual_classification_override",
    note: requireString(conflictNote) || `Manual classification override applied by ${actorEmail}.`,
    created_by_email: actorEmail,
  });

  return { ok: true };
}

async function resolveCuratorId(curatorName: string) {
  const normalized = requireString(curatorName);
  if (!normalized) return null;
  const { data } = await adminClient
    .from("gre_mis_curators")
    .select("id, display_name")
    .ilike("display_name", normalized)
    .maybeSingle();
  return data?.id || null;
}

function buildNeedNotes(curationCallDetails: string, solutionsShared: string) {
  return [curationCallDetails, solutionsShared ? `Solutions Shared: ${solutionsShared}` : ""]
    .filter(Boolean)
    .join("\n\n")
    .trim();
}

const suggestedQuestionsHeading = "Suggested Questions to Seeker";

function buildSuggestedQuestionSection(questions: string[]) {
  const rows = questions
    .map((question) => requireString(question).replace(/\s+/g, " ").trim())
    .filter(Boolean)
    .slice(0, 5)
    .map((question, index) => `${index + 1}. ${question}`);
  if (!rows.length) return "";
  return `${suggestedQuestionsHeading}\n${rows.join("\n")}`.trim();
}

function mergeSuggestedQuestionSection(existingNotes: string, suggestedSection: string) {
  const cleanExisting = requireString(existingNotes).replace(/\r/g, "").trim();
  const cleanSuggested = requireString(suggestedSection).trim();
  if (!cleanSuggested) return cleanExisting;
  const pattern = new RegExp(`${suggestedQuestionsHeading}[\\s\\S]*?(?=\\n\\n[^\\n]|$)`, "i");
  if (pattern.test(cleanExisting)) {
    return cleanExisting.replace(pattern, cleanSuggested).trim();
  }
  return [cleanSuggested, cleanExisting].filter(Boolean).join("\n\n").trim();
}

function removeSuggestedQuestionSection(existingNotes: string) {
  const cleanExisting = requireString(existingNotes).replace(/\r/g, "").trim();
  if (!cleanExisting) return "";
  const pattern = new RegExp(`${suggestedQuestionsHeading}[\\s\\S]*?(?=\\n\\n[^\\n]|$)`, "i");
  return cleanExisting.replace(pattern, "").replace(/\n{3,}/g, "\n\n").trim();
}

function appendSharedSolutionToCurationNotes(existingNotes: string, offeringName: string, providerName: string, viewLink: string) {
  const cleanExisting = requireString(existingNotes).replace(/\r/g, "").trim();
  const normalizedOffering = requireString(offeringName);
  const normalizedProvider = requireString(providerName);
  const normalizedLink = requireString(viewLink);
  const entryText = [normalizedOffering, normalizedProvider, normalizedLink].filter(Boolean).join(" : ");
  if (!normalizedOffering && !normalizedProvider) return cleanExisting;
  if (cleanExisting.includes(entryText)) return cleanExisting;

  const header = "Solutions Shared";
  const linePattern = new RegExp(`${header}\\n((?:\\d+\\. .*\\n?)*)`, "i");
  const match = cleanExisting.match(linePattern);
  let existingLines: string[] = [];
  let rest = "";

  if (match) {
    existingLines = match[1].split("\n").map((l) => l.trim()).filter(Boolean);
    rest = cleanExisting.replace(match[0], "").trim();
  }

  const isDuplicate = existingLines.some((l) => l.toLowerCase().includes(entryText.toLowerCase()));
  if (!isDuplicate) {
    existingLines.push(`${existingLines.length + 1}. ${entryText}`);
  }

  const section = `${header}\n${existingLines.join("\n")}`;
  const other = match ? rest : cleanExisting;
  return [other, section].filter(Boolean).join("\n\n").trim();
}

function buildRuleBasedSuggestedQuestions(need: Record<string, unknown>) {
  const thematicArea = requireString((need as Record<string, unknown>).submitted_thematic_area) || requireString(need.ai_thematic_area) || "this need";
  const needType = requireString((need as Record<string, unknown>).submitted_offering_type) || requireString(need.ai_service_kind) || "support";
  const deployment = asStringArray((need as Record<string, unknown>).deployment_locations).join(", ") || [requireString(need.district), requireString(need.state)].filter(Boolean).join(", ") || "the target location";
  return uniqueStrings([
    `Which exact locations in ${deployment} should we prioritize first, and how many sites or habitations are involved?`,
    `What is the current pain point with ${thematicArea}, and what outcome would make this need feel solved for the community?`,
    `What affordability range or budget constraint should we keep in mind before we shortlist ${needType} options?`,
    `Who will operate and maintain the solution locally, and what village-level capacity already exists for upkeep?`,
    `What scale do you want to reach in the next 3 to 12 months, and what constraints could slow adoption across smaller settlements?`,
  ]).slice(0, 5);
}

async function saveSuggestedQuestionsForNeed(
  needId: string,
  questionsInput: unknown,
  actorEmail: string,
  sourceLabel: string,
) {
  const { data: need, error } = await adminClient
    .from("gre_mis_needs")
    .select("*")
    .eq("id", needId)
    .single();
  if (error || !need) throw new Error(error?.message || "Need not found.");
  if (requireString(need.approval_status) === "pending_admin") {
    throw new Error("Suggested questions can be saved only for approved needs.");
  }

  const questions = uniqueStrings(asStringArray(questionsInput))
    .map((question) => requireString(question).replace(/^\d+[\).\s-]*/, "").trim())
    .filter((question) => question.length >= 12)
    .slice(0, 5);
  if (!questions.length) {
    throw new Error("No valid suggested questions were provided.");
  }

  const suggestedSection = buildSuggestedQuestionSection(questions);
  const mergedNotes = mergeSuggestedQuestionSection(requireString(need.curation_notes), suggestedSection);
  const { error: updateError } = await adminClient
    .from("gre_mis_needs")
    .update({
      curation_notes: mergedNotes,
      updated_at: new Date().toISOString(),
    })
    .eq("id", needId);
  if (updateError) throw new Error(updateError.message);

  await adminClient.from("gre_mis_need_updates").insert({
    need_id: needId,
    update_type: "suggested_questions_generated",
    note: `Suggested curator questions generated by ${sourceLabel} flow for ${actorEmail}.`,
    created_by_email: actorEmail,
  });

  return {
    ok: true,
    questions,
    curationNotes: mergedNotes,
    message: sourceLabel === "puter"
      ? "Suggested questions from Puter were added to the curation notes."
      : "Suggested questions were added to the curation notes.",
  };
}

async function rejectSuggestedQuestionsForNeed(
  needId: string,
  actorEmail: string,
) {
  const { data: need, error } = await adminClient
    .from("gre_mis_needs")
    .select("*")
    .eq("id", needId)
    .single();
  if (error || !need) throw new Error(error?.message || "Need not found.");

  const nextNotes = removeSuggestedQuestionSection(requireString(need.curation_notes));
  const { error: updateError } = await adminClient
    .from("gre_mis_needs")
    .update({
      curation_notes: nextNotes || null,
      updated_at: new Date().toISOString(),
    })
    .eq("id", needId);
  if (updateError) throw new Error(updateError.message);

  await adminClient.from("gre_mis_need_updates").insert({
    need_id: needId,
    update_type: "suggested_questions_rejected",
    note: `Suggested curator questions were removed by ${actorEmail}.`,
    created_by_email: actorEmail,
  });

  return {
    ok: true,
    curationNotes: nextNotes,
    message: "Suggested questions removed from the curation notes.",
  };
}

async function generateSuggestedQuestionsForNeed(
  needId: string,
  actorEmail: string,
) {
  const { data: need, error } = await adminClient
    .from("gre_mis_needs")
    .select("*")
    .eq("id", needId)
    .single();
  if (error || !need) throw new Error(error?.message || "Need not found.");
  if (requireString(need.approval_status) === "pending_admin") {
    throw new Error("Suggested questions can be generated only for approved needs.");
  }

  const prompt = `You are preparing a curator for a need-assessment call with a rural organisation seeker.

Return valid JSON only in this shape:
{
  "questions": ["", "", "", "", ""]
}

Rules:
- return exactly 5 probing questions
- each question should help clarify the need fast and improve solution matching
- prioritize: deployment context, current pain point, budget/affordability, maintenance/operations, scale and adoption constraints
- keep each question concise and practical
- avoid repeating the need statement back
- do not include numbering in the JSON

Need record:
${JSON.stringify({
    organization_name: need.organization_name,
    state: need.state,
    district: need.district,
    thematic_area: requireString((need as Record<string, unknown>).submitted_thematic_area) || need.ai_thematic_area,
    need_category: requireString((need as Record<string, unknown>).submitted_offering_category) || need.ai_need_kind,
    need_type: requireString((need as Record<string, unknown>).submitted_offering_type) || need.ai_service_kind,
    keywords: asStringArray(need.submitted_keywords),
    curated_need: asStringArray(need.curated_need),
    problem_statement: need.problem_statement,
  }, null, 2)}
`;

  let questions: string[] = [];
  let generationMode = "ai";
  try {
    const ai = await callAiJsonWithOrder(["gemini", "openai"], prompt);
    questions = uniqueStrings(asStringArray(ai.questions))
      .map((question) => requireString(question).replace(/^\d+[\).\s-]*/, "").trim())
      .filter((question) => question.length >= 12)
      .slice(0, 5);
    if (questions.length < 5) {
      throw new Error("AI did not return enough suggested questions.");
    }
  } catch {
    questions = buildRuleBasedSuggestedQuestions(need);
    generationMode = "rules";
  }

  const saved = await saveSuggestedQuestionsForNeed(
    needId,
    questions,
    actorEmail,
    generationMode,
  );
  return {
    ...saved,
    message: generationMode === "ai"
      ? "Suggested questions were added to the curation notes."
      : "Suggested questions were added using the fallback question generator.",
  };
}

function parseGreInboundWorkbookRows(buffer: ArrayBuffer) {
  return readWorkbookRows(buffer)
    .map((row) => ({
      request_id: requireString(row["Request Id"]),
      organization_name: requireString(row["Seeker Organisation"]),
      website: requireString(row.Website),
      contact_person: requireString(row["Contact Person"]),
      designation: requireString(row.Designation),
      seeker_phone: requireString(row["Phone Number"]),
      seeker_email: requireString(row.Email).toLowerCase(),
      requested_on: parseWorkbookDate(row["Requested On"]),
      problem_statement: requireString(row.Request),
      state: requireString(row["Solution Needed in State"]),
      district: requireString(row["Solution Needed in District"]),
      status: requireString(row.Status),
      internal_status: requireString(row["Internal Status"]),
      curator_name: requireString(row["Curator Assigned"]),
      curation_call_date: parseWorkbookDate(row["Curation Call Date"])?.slice(0, 10) || null,
      curation_age_days: parseNumber(row["Curation Age"], 0),
      curation_call_details: requireString(row["Curation Call Details"]),
      curated_need: asStringArray(String(row["Curated Need of Service Seeker"] || "").replaceAll(",", ";")),
      demand_broadcast_needed: parseBoolean(row["Demand Broadcast Needed"]),
      solutions_shared_count: parseNumber(row["Solutions Shared Count"], 0),
      solutions_shared: requireString(row["Solutions Shared"]),
      invited_providers_count: parseNumber(row["Invited Providers Count"], 0),
      invited_providers: requireString(row["Invited Providers"]),
      funding_mechanism: rowValue(row, [
        "Funding Mechanism",
        "Capture Outcome - Funding Mechanism",
        "Outcome Funding Mechanism",
        "Funding mechanism",
      ]),
      seeker_provider_agreement: rowValue(row, [
        "Seeker / Provider Agreement",
        "Seeker/Provider Agreement",
        "Capture Outcome - Seeker / Provider Agreement",
        "Seeker Provider Agreement",
      ]),
      solution_deployment_status: rowValue(row, [
        "Solution Deployment Status",
        "Capture Outcome - Solution Deployment Status",
        "Deployment Status",
      ]),
      closure_date: rowDateValue(row, [
        "Closure Date",
        "Capture Outcome - Closure Date",
        "Closed On",
      ]),
      feedback_about_seeker: rowValue(row, [
        "Feedback about Seeker",
        "Feedback About Seeker",
        "Seeker Feedback",
      ]),
      feedback_about_provider: rowValue(row, [
        "Feedback about Provider",
        "Feedback About Provider",
        "Provider Feedback",
      ]),
    }))
    .filter((row) => row.request_id);
}

function normalizeTraderRowForChatbot(row: Record<string, unknown>) {
  return {
    trader_id: String(row.TraderId || "").trim(),
    trader_name: normalizeCell(row.TraderName),
    organisation_name: normalizeCell(row.TraderOrganisation),
    mobile: normalizeCell(row.TraderMobile),
    email: normalizeCell(row.TraderMail),
    poc_name: normalizeCell(row.TraderPOC),
    tenant_id: normalizeCell(row.TraderTenantId),
    profile_id: normalizeCell(row.TraderProfileId),
    description: normalizeCell(row.TraderDescription),
    short_description: normalizeCell(row.TraderShortDescription),
    tagline: normalizeCell(row.TraderTagLine),
    website: normalizeCell(row.TraderWebsite),
    created_at_source: normalizeCell(row.TraderCreatedDate),
    association_status: normalizeCell(row.TraderAssociationStatus),
    raw_payload: row,
  };
}

function normalizeSolutionRowForChatbot(row: Record<string, unknown>) {
  return {
    solution_id: String(row.SolutionId || "").trim(),
    trader_id: normalizeCell(row.TraderId),
    solution_name: normalizeCell(row.SolutionName),
    solution_status: normalizeCell(row.SolutionStatus),
    publish_status: normalizeCell(row.SolutionPublishStatus),
    created_at_source: normalizeCell(row.SolutionCreationDate),
    about_solution_html: normalizeCell(row.AboutSolution),
    about_solution_text: stripHtml(normalizeCell(row.AboutSolution)),
    solution_image_url: rowValue(row, ["SolutionImage", "Solution Image", "Solution Image URL", "Solution Image Url"]) || null,
    raw_payload: row,
  };
}

function normalizeOfferingRowForChatbot(row: Record<string, unknown>) {
  const aboutOfferingHtml = normalizeCell(row.AboutOffering);
  const trainerDetailsHtml = normalizeCell(row["Trainer Details"]);
  const valuechains = splitLooseList(normalizeCell(row.Valuechains));
  const applications = splitLooseList(normalizeCell(row.Applications));
  const tags = splitLooseList(normalizeCell(row.Tags));
  const languages = normalizeLanguageArray(splitLooseList(normalizeCell(row.Languages)));
  const geographiesRaw = normalizeCell(row.Geographies);
  const geographies = splitGeographies(geographiesRaw);

  return {
    offering_id: String(row.OfferingId || "").trim().replace(/\.0$/, ""),
    solution_id: normalizeCell(row.SolutionId),
    trader_id: normalizeCell(row.TraderId),
    offering_name: normalizeCell(row.OfferingName),
    publish_status: normalizeCell(row.OfferingPublishStatus),
    created_at_source: normalizeCell(row.OfferingCreationDate),
    offering_category: normalizeCell(row.OfferingCategory),
    offering_group: normalizeCell(row.OfferingGroup),
    offering_type: normalizeCell(row.OfferingType),
    domain_6m: normalizeCell(row["6M"]),
    primary_valuechain_id: normalizeCell(row.PrimaryValuechainId),
    primary_valuechain: normalizeCell(row.PrimaryValuechain),
    primary_application_id: normalizeCell(row.PrimaryApplicationId),
    primary_application: normalizeCell(row.PrimaryApplication),
    valuechains,
    applications,
    tags,
    languages,
    geographies,
    geographies_raw: geographiesRaw,
    about_offering_html: aboutOfferingHtml,
    about_offering_text: stripHtml(aboutOfferingHtml),
    audience: normalizeCell(row["Who Can avail it"]),
    trainer_name: normalizeCell(row["Trainer Name"]),
    trainer_email: normalizeCell(row["Trainer Email Address"]),
    trainer_phone: normalizeCell(row["Trainer Phone Number"]),
    trainer_details_html: trainerDetailsHtml,
    trainer_details_text: stripHtml(trainerDetailsHtml),
    duration: normalizeCell(row.Duration),
    prerequisites: normalizeCell(row["Prerequisites - Participants and Training"]),
    service_cost: normalizeCell(row["Cost (Service)"]),
    support_post_service: normalizeCell(row["Support post Service"]),
    support_post_service_cost: normalizeCell(row["Support post Service Cost"]),
    delivery_mode: normalizeCell(row["Is it offered - Online or Offline"]),
    certification_offered: normalizeCell(row["Certification Offered"]),
    cost_remarks: normalizeCell(row["Remarks on Cost"]),
    location_availability: normalizeCell(row["Location Availability"]),
    service_brochure_url: rowValue(row, ["Service offering Brochure", "Service Offering Brochure", "Service Brochure", "ServiceBrochure"]) || null,
    grade_capacity: normalizeCell(row["Grade/Capacity"]),
    product_cost: normalizeCell(row["Cost (Product)"]),
    lead_time: normalizeCell(row["Lead Time"]),
    support_details: normalizeCell(row.Support),
    product_brochure_url: rowValue(row, ["Product Brochure", "Product Brochure URL", "Product Brochure Url", "ProductBrochure"]) || null,
    knowledge_content_url: rowValue(row, ["Knowledge Offering Content", "Knowledge Content URL", "Knowledge Content Url", "Knowledge Offering Content URL", "Knowledge Offering Content Url", "KnowledgeContent"]) || null,
    contact_details: normalizeCell(row["Contact Details"]),
    gre_link: rowValue(row, ["Offering Link on GRE", "GRE Link", "Offering GRE Link", "Offering Link", "OfferingLinkOnGRE"]) || null,
    source_row_signature: stableRowSignature(row),
    raw_payload: row,
  };
}

function dedupeRowsById<T extends Record<string, unknown>>(rows: T[], key: keyof T) {
  const map = new Map<string, T>();
  rows.forEach((row) => {
    const id = String(row[key] ?? "").trim();
    if (id) map.set(id, row);
  });
  return [...map.values()];
}

function buildChatbotImportBundle(solutionBuffer: ArrayBuffer, traderBuffer: ArrayBuffer) {
  const solutionRows = readWorkbookRows(solutionBuffer);
  const traderRows = readWorkbookRows(traderBuffer);
  return {
    traders: dedupeRowsById(traderRows.map(normalizeTraderRowForChatbot), "trader_id"),
    solutions: dedupeRowsById(solutionRows.map(normalizeSolutionRowForChatbot), "solution_id"),
    offerings: dedupeRowsById(solutionRows.map(normalizeOfferingRowForChatbot), "offering_id"),
    stats: {
      solutionRows: solutionRows.length,
      traderRows: traderRows.length,
    },
  };
}

async function applyChatbotImportBundle(
  bundle: {
    traders: Record<string, unknown>[];
    solutions: Record<string, unknown>[];
    offerings: Record<string, unknown>[];
    stats: { solutionRows: number; traderRows: number };
  },
  fileNames: { solutionFileName: string; traderFileName: string },
  provider: string,
  skipEnrichment = false,
) {
  const { data: importRow, error: importError } = await adminClient
    .from("data_imports")
    .insert({
      solution_file_name: fileNames.solutionFileName,
      trader_file_name: fileNames.traderFileName,
      status: "running",
      source_solution_rows: bundle.stats.solutionRows,
      source_trader_rows: bundle.stats.traderRows,
    })
    .select("id")
    .single();
  if (importError) throw new Error(importError.message);

  const importId = importRow.id;
  const offeringIds = bundle.offerings.map((row) => requireString(row.offering_id)).filter(Boolean);
  const { data: existingOfferings, error: existingOfferingsError } = offeringIds.length
    ? await adminClient
      .from("offerings")
      .select("offering_id, source_row_signature, ai_enriched_at, raw_payload")
      .in("offering_id", offeringIds)
    : { data: [], error: null };
  if (existingOfferingsError) throw new Error(existingOfferingsError.message);
  const existingOfferingMap = new Map((existingOfferings || []).map((row) => [requireString(row.offering_id), row]));
  const changedOfferingIds = bundle.offerings
    .filter((row) => {
      const existing = existingOfferingMap.get(requireString(row.offering_id)) as Record<string, unknown> | undefined;
      if (!existing) return true;
      return requireString(existing.source_row_signature) !== requireString(row.source_row_signature) || !existing.ai_enriched_at;
    })
    .map((row) => requireString(row.offering_id));

  try {
    for (const rows of chunkArray(bundle.traders)) {
      const { error } = await adminClient.from("traders").upsert(rows, { onConflict: "trader_id" });
      if (error) throw new Error(error.message);
    }

    for (const rows of chunkArray(bundle.solutions)) {
      const rowsCleaned = rows.map((row) => {
        const cleanRow = { ...row };
        const rawPayload = cleanRow.raw_payload as Record<string, unknown> | undefined;
        if (rawPayload && typeof rawPayload === "object") {
          const hasAnyColumn = (aliases: string[]) =>
            aliases.some((alias) => Object.prototype.hasOwnProperty.call(rawPayload, alias));
          if (!hasAnyColumn(["SolutionImage", "Solution Image", "Solution Image URL", "Solution Image Url"]))
            delete cleanRow.solution_image_url;
        }
        return cleanRow;
      });
      const { error } = await adminClient.from("solutions").upsert(rowsCleaned, { onConflict: "solution_id" });
      if (error) throw new Error(error.message);
    }

    const changedOfferings = bundle.offerings.filter((row) => {
      const id = requireString(row.offering_id);
      return id && changedOfferingIds.includes(id);
    });
    if (changedOfferings.length) {
      for (const rows of chunkArray(changedOfferings)) {
        const rowsWithImport = rows.map((row) => {
          const cleanRow = { ...row };
          const rawPayload = cleanRow.raw_payload as Record<string, unknown> | undefined;
          if (rawPayload && typeof rawPayload === "object") {
            const hasAnyColumn = (aliases: string[]) =>
              aliases.some((alias) => Object.prototype.hasOwnProperty.call(rawPayload, alias));
            if (!hasAnyColumn(["Service offering Brochure", "Service Offering Brochure", "Service Brochure", "ServiceBrochure"]))
              delete cleanRow.service_brochure_url;
            if (!hasAnyColumn(["Product Brochure", "Product Brochure URL", "Product Brochure Url", "ProductBrochure"]))
              delete cleanRow.product_brochure_url;
            if (!hasAnyColumn(["Knowledge Offering Content", "Knowledge Content URL", "Knowledge Content Url", "Knowledge Offering Content URL", "Knowledge Offering Content Url", "KnowledgeContent"]))
              delete cleanRow.knowledge_content_url;
            if (!hasAnyColumn(["Offering Link on GRE", "GRE Link", "Offering GRE Link", "Offering Link", "OfferingLinkOnGRE"]))
              delete cleanRow.gre_link;
          }
          const existing = existingOfferingMap.get(requireString(cleanRow.offering_id)) as Record<string, unknown> | undefined;
          if (existing) {
            const existingRawPayload = existing.raw_payload as Record<string, unknown> | undefined;
            const manualFields = (existingRawPayload?.last_manual_edit as Record<string, unknown> | undefined)?.manual_fields as string[] | undefined;
            if (manualFields?.length) {
              for (const field of manualFields) {
                delete cleanRow[field];
              }
            }
          }
          const rp = cleanRow.raw_payload as Record<string, unknown> | undefined;
          cleanRow.search_document = rp ? buildSearchDocument([
            normalizeCell(rp.SolutionName),
            normalizeCell(rp.OfferingName),
            normalizeCell(rp.OfferingCategory),
            normalizeCell(rp.OfferingGroup),
            normalizeCell(rp.OfferingType),
            normalizeCell(rp["6M"]),
            normalizeCell(rp.PrimaryValuechain),
            normalizeCell(rp.PrimaryApplication),
            splitLooseList(normalizeCell(rp.Valuechains)),
            splitLooseList(normalizeCell(rp.Applications)),
            splitLooseList(normalizeCell(rp.Tags)),
            splitLooseList(normalizeCell(rp.Languages)),
            splitGeographies(normalizeCell(rp.Geographies)),
            stripHtml(normalizeCell(rp.AboutSolution)),
            stripHtml(normalizeCell(rp.AboutOffering)),
            normalizeCell(rp.TraderOrganisation),
          ]) : null;
          return { ...cleanRow, last_import_id: importId };
        });
        const { error } = await adminClient.from("offerings").upsert(rowsWithImport, { onConflict: "offering_id" });
        if (error) throw new Error(error.message);
      }
    }
    const unchangedOfferingIds = bundle.offerings
      .filter((row) => {
        const id = requireString(row.offering_id);
        return id && !changedOfferingIds.includes(id);
      })
      .map((row) => requireString(row.offering_id));
    if (unchangedOfferingIds.length) {
      for (const ids of chunkArray(unchangedOfferingIds)) {
        const { error } = await adminClient
          .from("offerings")
          .update({ last_import_id: importId })
          .in("offering_id", ids);
        if (error) throw new Error(error.message);
      }
    }

    if (changedOfferingIds.length && !skipEnrichment) {
      const { data: offeringRows, error: offeringError } = await adminClient
        .from("offerings")
        .select("*")
        .in("offering_id", changedOfferingIds);
      if (offeringError) throw new Error(offeringError.message);

      const solutionIds = uniqueStrings((offeringRows || []).map((row) => requireString(row.solution_id)));
      const traderIds = uniqueStrings((offeringRows || []).map((row) => requireString(row.trader_id)));
      const [{ data: solutionRows, error: solutionError }, { data: traderRows, error: traderError }] = await Promise.all([
        solutionIds.length
          ? adminClient.from("solutions").select("*").in("solution_id", solutionIds)
          : Promise.resolve({ data: [], error: null }),
        traderIds.length
          ? adminClient.from("traders").select("*").in("trader_id", traderIds)
          : Promise.resolve({ data: [], error: null }),
      ]);
      if (solutionError) throw new Error(solutionError.message);
      if (traderError) throw new Error(traderError.message);

      const solutionMap = new Map((solutionRows || []).map((row) => [requireString(row.solution_id), row]));
      const traderMap = new Map((traderRows || []).map((row) => [requireString(row.trader_id), row]));

      for (const offeringRow of offeringRows || []) {
        await enrichOfferingIntelligence(
          offeringRow,
          solutionMap.get(requireString(offeringRow.solution_id)) || null,
          traderMap.get(requireString(offeringRow.trader_id)) || null,
          provider,
        );
      }
    }

    const { error: completeError } = await adminClient
      .from("data_imports")
      .update({
        status: "completed",
        inserted_traders: bundle.traders.length,
        inserted_solutions: bundle.solutions.length,
        inserted_offerings: bundle.offerings.length,
        completed_at: new Date().toISOString(),
      })
      .eq("id", importId);
    if (completeError) throw new Error(completeError.message);

    await invalidateAskGreSearchCache();

    return {
      importId,
      traders: bundle.traders.length,
      solutions: bundle.solutions.length,
      offerings: bundle.offerings.length,
      offeringAiUpdated: changedOfferingIds.length,
    };
  } catch (error) {
    await adminClient
      .from("data_imports")
      .update({
        status: "failed",
        completed_at: new Date().toISOString(),
        error_message: error instanceof Error ? error.message : "Unknown import error",
      })
      .eq("id", importId);
    throw error;
  }
}

async function normalizeInboundRow(row: Record<string, unknown>, curatorId: string | null, fileName: string) {
  const requestId = requireString(row.request_id);
  const problemStatement = requireString(row.problem_statement);
  const curationNotes = buildNeedNotes(requireString(row.curation_call_details), requireString(row.solutions_shared));
  const curatedNeed = asStringArray(row.curated_need);
  const geocode = await geocodeNeedLocation({
    organization_name: requireString(row.organization_name),
    district: requireString(row.district),
    state: requireString(row.state),
  });
  return {
    id: requestId,
    organization_name: requireString(row.organization_name),
    website: requireString(row.website),
    contact_person: requireString(row.contact_person),
    designation: requireString(row.designation),
    seeker_phone: requireString(row.seeker_phone),
    seeker_email: requireString(row.seeker_email).toLowerCase(),
    requested_on: requireString(row.requested_on) || new Date().toISOString(),
    problem_statement: problemStatement,
    state: requireString(row.state),
    district: requireString(row.district),
    status: requireString(row.status) || "New",
    internal_status: requireString(row.internal_status) || "Need solution providers",
    curator_id: curatorId,
    curation_call_date: requireString(row.curation_call_date) || null,
    curation_age_days: parseNumber(row.curation_age_days, 0),
    curation_notes: curationNotes || null,
    curated_need: curatedNeed,
    demand_broadcast_needed: parseBoolean(row.demand_broadcast_needed) ?? false,
    solutions_shared_count: parseNumber(row.solutions_shared_count, 0),
    invited_providers_count: parseNumber(row.invited_providers_count, 0),
    next_action: requireString(row.status).toLowerCase() === "closed" ? "Closed" : "Follow up with seeker",
    funding_mechanism: requireString(row.funding_mechanism) || null,
    seeker_provider_agreement: requireString(row.seeker_provider_agreement) || null,
    solution_deployment_status: requireString(row.solution_deployment_status) || null,
    closure_date: requireString(row.closure_date) || null,
    feedback_about_seeker: requireString(row.feedback_about_seeker) || null,
    feedback_about_provider: requireString(row.feedback_about_provider) || null,
    approval_status: "approved",
    imported_from_batch: fileName,
    source_kind: "website_inbound_snapshot",
    last_status_change_at: new Date().toISOString(),
    last_synced_at: new Date().toISOString(),
    latitude: geocode.latitude,
    longitude: geocode.longitude,
    geocoded_label: geocode.geocoded_label || null,
    geocode_status: geocode.geocode_status,
    geocoded_at: new Date().toISOString(),
  };
}

async function importInboundWorkbook(rowsInput: unknown, fileName: string, actorEmail: string, provider: string) {
  const rows = Array.isArray(rowsInput) ? rowsInput as Record<string, unknown>[] : [];
  if (!rows.length) throw new Error("No inbound rows were received.");

  const ids = rows.map((row) => requireString(row.request_id)).filter(Boolean);
  const { data: existingRows, error: existingError } = await adminClient
    .from("gre_mis_needs")
    .select("id, source_row_signature, curator_id")
    .in("id", ids);
  if (existingError) throw new Error(existingError.message);

  const existingMap = new Map((existingRows || []).map((row) => [String(row.id), row]));
  const toInsert: Record<string, unknown>[] = [];
  const toUpdate: Record<string, unknown>[] = [];
  const changedNeedIds: string[] = [];

  for (const row of rows) {
    const requestId = requireString(row.request_id);
    if (!requestId) continue;
    const sourceRowSignature = stableInboundRowSignature(row);
    const existing = existingMap.get(requestId);

    if (existing && existing.source_row_signature === sourceRowSignature) {
      continue;
    }

    const curatorId = await resolveCuratorId(requireString(row.curator_name));
    const patch = await normalizeInboundRow(row, curatorId, fileName);

    if (!existing) {
      toInsert.push({
        ...patch,
        source_row_signature: sourceRowSignature,
      });
      changedNeedIds.push(requestId);
      continue;
    }

    if (existing.source_row_signature !== sourceRowSignature) {
      toUpdate.push({
        ...patch,
        source_row_signature: sourceRowSignature,
      });
      changedNeedIds.push(requestId);
    }
  }

  if (toInsert.length) {
    const { error } = await adminClient.from("gre_mis_needs").insert(toInsert);
    if (error) throw new Error(error.message);
  }

  for (const row of toUpdate) {
    const { id, ...patch } = row;
    const { error } = await adminClient.from("gre_mis_needs").update(patch).eq("id", id);
    if (error) throw new Error(error.message);
  }

  let aiUpdatedCount = 0;
  if (changedNeedIds.length) {
    const { data: changedNeeds, error } = await adminClient
      .from("gre_mis_needs")
      .select("*")
      .in("id", changedNeedIds);
    if (error) throw new Error(error.message);

    for (const need of changedNeeds || []) {
      try {
        await enrichNeedIntelligence(need, provider);
        aiUpdatedCount += 1;
      } catch (error) {
        await adminClient
          .from("gre_mis_needs")
          .update({
            ai_engine: provider,
            ai_enriched_at: new Date().toISOString(),
            ai_enrichment_status: error instanceof Error ? `error: ${error.message}` : "error",
          })
          .eq("id", need.id);
      }
    }
  }

  await adminClient.from("gre_mis_import_runs").insert({
    file_name: fileName,
    imported_by_email: actorEmail,
    total_rows: rows.length,
    inserted_count: toInsert.length,
    updated_count: toUpdate.length,
    ai_updated_count: aiUpdatedCount,
  });

  return {
    ok: true,
    insertedCount: toInsert.length,
    updatedCount: toUpdate.length,
    aiUpdatedCount,
  };
}

async function syncGreLiveInbounds(actorEmail: string, provider: string) {
  const report = await fetchGreReportWorkbook(greInboundReportName, 1000, "request_details_report");
  const rows = parseGreInboundWorkbookRows(report.buffer);
  const summary = await importInboundWorkbook(rows, report.fileName, actorEmail, provider);
  return {
    ...summary,
    downloadedFileName: report.fileName,
  };
}

async function syncGreChatbotData(provider: string) {
  const [traderReport, solutionReport] = await Promise.all([
    fetchGreReportWorkbook(greTraderReportName, 1000, "trader_data"),
    fetchGreReportWorkbook(greSolutionReportName, 5000, "solution_data"),
  ]);
  const bundle = buildChatbotImportBundle(solutionReport.buffer, traderReport.buffer);
  const summary = await applyChatbotImportBundle(bundle, {
    solutionFileName: solutionReport.fileName,
    traderFileName: traderReport.fileName,
  }, provider);
  return {
    ok: true,
    summary,
    files: {
      trader: { fileName: traderReport.fileName },
      solution: { fileName: solutionReport.fileName },
    },
  };
}

async function uploadChatbotWorkbooks(
  solutionBase64: string,
  traderBase64: string,
  solutionFileName: string,
  traderFileName: string,
  provider: string,
) {
  if (!solutionBase64 && !traderBase64) {
    throw new Error("No solution or trader workbook data was received.");
  }
  const solutionBuffer = solutionBase64 ? decodeBase64ToArrayBuffer(solutionBase64) : new ArrayBuffer(0);
  const traderBuffer = traderBase64 ? decodeBase64ToArrayBuffer(traderBase64) : new ArrayBuffer(0);
  if (!solutionBuffer.byteLength && !traderBuffer.byteLength) {
    throw new Error("Could not decode workbook data.");
  }
  const bundle = buildChatbotImportBundle(solutionBuffer, traderBuffer);
  const summary = await applyChatbotImportBundle(bundle, {
    solutionFileName: solutionFileName || "uploaded_solutions.xlsx",
    traderFileName: traderFileName || "uploaded_traders.xlsx",
  }, provider);
  return { ok: true, summary };
}

async function uploadChatbotNormalized(
  traders: Record<string, unknown>[],
  solutions: Record<string, unknown>[],
  offerings: Record<string, unknown>[],
  solutionFileName: string,
  traderFileName: string,
  provider: string,
) {
  const bundle: ImportBundle = {
    traders,
    solutions,
    offerings: offerings.map((row) => {
      const rp = row.raw_payload as Record<string, unknown> | undefined;
      return {
        ...row,
        search_document: rp ? buildSearchDocument([
          normalizeCell(rp.SolutionName),
          normalizeCell(rp.OfferingName),
          normalizeCell(rp.OfferingCategory),
          normalizeCell(rp.OfferingGroup),
          normalizeCell(rp.OfferingType),
          normalizeCell(rp["6M"]),
          normalizeCell(rp.PrimaryValuechain),
          normalizeCell(rp.PrimaryApplication),
          splitLooseList(normalizeCell(rp.Valuechains)),
          splitLooseList(normalizeCell(rp.Applications)),
          splitLooseList(normalizeCell(rp.Tags)),
          splitLooseList(normalizeCell(rp.Languages)),
          splitGeographies(normalizeCell(rp.Geographies)),
          stripHtml(normalizeCell(rp.AboutSolution)),
          stripHtml(normalizeCell(rp.AboutOffering)),
          normalizeCell(rp.TraderOrganisation),
        ]) : null,
      };
    }),
    stats: { solutionRows: solutions.length, traderRows: traders.length },
  };
  const summary = await applyChatbotImportBundle(bundle, {
    solutionFileName: solutionFileName || "normalized_solutions.json",
    traderFileName: traderFileName || "normalized_traders.json",
  }, provider, true);
  return { ok: true, summary };
}

function decodeBase64ToArrayBuffer(base64: string): ArrayBuffer {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes.buffer;
}

async function downloadGreInboundReport() {
  const report = await fetchGreReportWorkbook(greInboundReportName, 1000, "request_details_report");
  return {
    ok: true,
    download: await createWorkbookDownloadPayload(report.buffer, report.fileName),
  };
}

async function downloadGreChatbotReport(reportKind: string) {
  const normalized = requireString(reportKind).toLowerCase();
  if (normalized === "trader") {
    const report = await fetchGreReportWorkbook(greTraderReportName, 1000, "trader_data");
    return {
      ok: true,
      kind: "trader",
      download: await createWorkbookDownloadPayload(report.buffer, report.fileName),
    };
  }
  if (normalized === "solution") {
    const report = await fetchGreReportWorkbook(greSolutionReportName, 5000, "solution_data");
    return {
      ok: true,
      kind: "solution",
      download: await createWorkbookDownloadPayload(report.buffer, report.fileName),
    };
  }
  throw new Error("Unsupported GRE chatbot report kind.");
}

function formatDdmmyy(date = new Date()) {
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = String(date.getFullYear()).slice(-2);
  return `${day}${month}${year}`;
}

function safeFilePart(value: unknown) {
  return (requireString(value) || "Seeker")
    .replace(/[\\/:*?"<>|]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 80) || "Seeker";
}


async function downloadSeekerRequestTracker(seekerKey: string, includeClosed = false) {
  const key = requireString(seekerKey);
  if (!key) throw new Error("Select a solution seeker first.");
  const isOrgKey = key.startsWith("org:");
  let query = adminClient
    .from("gre_mis_needs")
    .select("id, organization_name, contact_person, seeker_email, problem_statement, curation_notes, status, internal_status, next_action, requested_on, seeker_provider_agreement")
    .eq("approval_status", "approved")
    .order("requested_on", { ascending: true });
  if (!includeClosed) {
    query = query.neq("status", "Closed");
  }
  const { data, error } = isOrgKey
    ? await query.eq("organization_name", key.slice(4))
    : await query.eq("seeker_email", key.toLowerCase());
  if (error) throw new Error(error.message);
  const needs = data || [];
  if (!needs.length) throw new Error("No needs were found for this seeker.");

  const seekerLabel =
    requireString(needs[0]?.organization_name) ||
    requireString(needs[0]?.contact_person) ||
    requireString(needs[0]?.seeker_email) ||
    "Solution Seeker";
  const totalNeeds = needs.length;
  const solutionProviderNeeded = needs.filter((need) => requireString(need.internal_status).toLowerCase() === "need solution providers").length;
  const seekerResponsePending = needs.filter((need) => {
    const iStatus = requireString(need.internal_status).toLowerCase();
    const nAction = requireString(need.next_action).toLowerCase();
    return iStatus.includes("connection made") && (!nAction || nAction === "follow up with seeker");
  }).length;
  const solutionsImplemented = needs.filter((need) =>
    requireString(need.status).toLowerCase() === "closed" ||
    requireString(need.seeker_provider_agreement).toLowerCase() === "agreement completed"
  ).length;

  const html = buildStyledTrackerHtml({
    seekerLabel,
    totalNeeds,
    solutionProviderNeeded,
    seekerResponsePending,
    solutionsImplemented,
    needs,
  });
  const buffer = new TextEncoder().encode(html).buffer;
  return {
    ok: true,
    download: await createWorkbookDownloadPayload(
      buffer,
      `Solution Seeker Status - ${safeFilePart(seekerLabel)} - ${formatDdmmyy()}.xls`,
      "application/vnd.ms-excel",
    ),
  };
}

async function refreshNeedIntelligence(actorEmail: string, provider: string) {
  const { data: needs, error } = await adminClient
    .from("gre_mis_needs")
    .select("*")
    .eq("approval_status", "approved")
    .order("updated_at", { ascending: false })
    .limit(200);

  if (error) throw new Error(error.message);

  const staleNeeds = (needs || [])
    .filter((need) =>
      !need.ai_enriched_at ||
      !need.ai_validation_status ||
      need.ai_validation_status === "flagged" ||
      new Date(need.updated_at).getTime() >= new Date(need.ai_enriched_at).getTime(),
    )
    .slice(0, 250);

  let aiUpdatedCount = 0;
  for (const need of staleNeeds) {
    try {
      if (typeof need.latitude !== "number" || typeof need.longitude !== "number") {
        try {
          const geocode = await geocodeNeedLocation({
            organization_name: requireString(need.organization_name),
            district: requireString(need.district),
            state: requireString(need.state),
          });
          await adminClient
            .from("gre_mis_needs")
            .update({
              latitude: geocode.latitude,
              longitude: geocode.longitude,
              geocoded_label: geocode.geocoded_label || null,
              geocode_status: geocode.geocode_status,
              geocoded_at: new Date().toISOString(),
            })
            .eq("id", need.id);
        } catch {
          await adminClient
            .from("gre_mis_needs")
            .update({
              geocode_status: "geocode_failed",
              geocoded_at: new Date().toISOString(),
            })
            .eq("id", need.id);
        }
      }
      await enrichNeedIntelligence(need, provider);
      aiUpdatedCount += 1;
    } catch (error) {
      const rules = classifyNeedByRules(need);
      await adminClient
        .from("gre_mis_needs")
        .update({
          rule_thematic_hints: rules.thematicHints,
          rule_service_hints: rules.serviceHints,
          rule_keywords: rules.keywords,
          rule_6m_signals: rules.sixMSignals,
          rule_need_kind: rules.needKind,
          ai_thematic_area: rules.thematicHints[0] || null,
          ai_need_kind: rules.needKind || null,
          ai_service_kind: rules.serviceHints[0] || null,
          ai_keywords: rules.keywords,
          ai_6m_signals: rules.sixMSignals,
          ai_summary: requireString(need.problem_statement).slice(0, 500),
          ai_engine: "rules_only",
          ai_enriched_at: new Date().toISOString(),
          ai_enrichment_status: "rules_only",
          ai_validation_status: rules.thematicHints.length || rules.keywords.length ? "ready" : "flagged",
          ai_validation_flags: rules.thematicHints.length || rules.keywords.length ? [] : ["needs_review"],
          ai_confidence: rules.thematicHints.length ? 74 : rules.keywords.length ? 61 : 28,
          ai_payload: {
            mode: "rules_only",
            reason: error instanceof Error ? error.message : "Rule fallback applied.",
          },
        })
        .eq("id", need.id);
      aiUpdatedCount += 1;
    }
  }

  return {
    ok: true,
    aiUpdatedCount,
    message: `AI need intelligence refreshed for ${aiUpdatedCount} needs.`,
  };
}

async function refreshChatbotIntelligence(provider: string) {
  const { data: existing, error } = await adminClient
    .from("offerings")
    .select("offering_id, solution_id, trader_id, ai_enriched_at")
    .filter("ai_enriched_at", "is", null)
    .limit(250);

  if (error) throw new Error(error.message);
  const toEnrich = existing || [];
  if (!toEnrich.length) return { ok: true, aiUpdatedCount: 0, message: "All offerings already have AI enrichment." };

  const solutionIds = uniqueStrings(toEnrich.map((r) => requireString(r.solution_id)));
  const traderIds = uniqueStrings(toEnrich.map((r) => requireString(r.trader_id)));
  const [{ data: solutionRows }, { data: traderRows }] = await Promise.all([
    solutionIds.length ? adminClient.from("solutions").select("*").in("solution_id", solutionIds) : Promise.resolve({ data: [] }),
    traderIds.length ? adminClient.from("traders").select("*").in("trader_id", traderIds) : Promise.resolve({ data: [] }),
  ]);
  const solutionMap = new Map((solutionRows || []).map((r) => [requireString(r.solution_id), r]));
  const traderMap = new Map((traderRows || []).map((r) => [requireString(r.trader_id), r]));

  let count = 0;
  for (const offeringRow of toEnrich) {
    try {
      const fullRow = offeringRow as Record<string, unknown>;
      await enrichOfferingIntelligence(
        fullRow,
        solutionMap.get(requireString(fullRow.solution_id)) || null,
        traderMap.get(requireString(fullRow.trader_id)) || null,
        provider,
      );
      count += 1;
    } catch {
      // skip individual failures
    }
  }
  return { ok: true, aiUpdatedCount: count, message: `AI chatbot intelligence refreshed for ${count} offerings.` };
}

async function assignCurator(needId: string, curatorId: string | null, actorEmail: string) {
  const { data: needRow, error: needFetchError } = await adminClient
    .from("gre_mis_needs")
    .select("id, source_kind")
    .eq("id", needId)
    .single();
  if (needFetchError || !needRow) throw new Error(needFetchError?.message || "Need not found.");

  const sourceKind = requireString(needRow.source_kind);
  const isLocalOnlyNeed = sourceKind === "shared_form_submission" || requireString(needRow.id).startsWith("FORM-");
  let resolvedCuratorId: string | null = null;

  if (curatorId) {
    const normalizedSelection = requireString(curatorId);
    if (normalizedSelection.startsWith("curator:")) {
      resolvedCuratorId = normalizedSelection.slice("curator:".length);
    } else if (normalizedSelection.startsWith("user:")) {
      const userId = normalizedSelection.slice("user:".length);
      const { data: userRow, error: userError } = await adminClient
        .from("gre_mis_users")
        .select("id, username, first_name, full_name, email, phone, role, gre_user_id, gre_sync_status, is_active")
        .eq("id", userId)
        .single();
      if (userError || !userRow) throw new Error(userError?.message || "Selected user could not be found.");
      if (userRow.is_active === false) throw new Error("Only active users can be assigned.");

      const userRole = requireString(userRow.role).toLowerCase();
      if (!["admin", "moderator", "curator"].includes(userRole)) {
        throw new Error("Only admin, moderator, or curator users can be assigned.");
      }
      if (!isLocalOnlyNeed) {
        if (!["admin", "moderator", "curator"].includes(userRole)) {
          throw new Error("GRE-synced needs can be assigned only to GRE-registered admins, moderators, or curators.");
        }
        if (!requireString(userRow.gre_user_id)) {
          throw new Error("This user is not registered on the GRE website, so the assignment cannot sync back.");
        }
      }

      const email = requireString(userRow.email).toLowerCase();
      const { data: existingCurator, error: curatorFetchError } = await adminClient
        .from("gre_mis_curators")
        .select("id")
        .or(`user_id.eq.${userId},email.eq.${email}`)
        .order("created_at", { ascending: true })
        .limit(1)
        .maybeSingle();
      if (curatorFetchError) throw new Error(curatorFetchError.message);

      if (existingCurator?.id) {
        const { error: updateCuratorError } = await adminClient
          .from("gre_mis_curators")
          .update({
            user_id: userId,
            display_name: requireString(userRow.full_name) || requireString(userRow.first_name) || requireString(userRow.username),
            first_name: requireString(userRow.first_name) || requireString(userRow.full_name).split(" ")[0] || requireString(userRow.username),
            email,
            phone: requireString(userRow.phone) || null,
            is_active: true,
            gre_sync_status: isLocalOnlyNeed ? "local_only" : requireString(userRow.gre_sync_status) || "synced",
            gre_sync_message: isLocalOnlyNeed
              ? "Local GramEEE curator assignment enabled for shared-form needs."
              : "GRE-linked curator assignment ready.",
            gre_synced_at: new Date().toISOString(),
            updated_at: new Date().toISOString(),
          })
          .eq("id", existingCurator.id);
        if (updateCuratorError) throw new Error(updateCuratorError.message);
        resolvedCuratorId = existingCurator.id;
      } else {
        const { data: upsertedCurator, error: upsertCuratorError } = await adminClient
          .from("gre_mis_curators")
          .upsert({
            user_id: userId,
            display_name: requireString(userRow.full_name) || requireString(userRow.first_name) || requireString(userRow.username),
            first_name: requireString(userRow.first_name) || requireString(userRow.full_name).split(" ")[0] || requireString(userRow.username),
            email,
            phone: requireString(userRow.phone) || null,
            is_active: true,
            gre_sync_status: isLocalOnlyNeed ? "local_only" : requireString(userRow.gre_sync_status) || "synced",
            gre_sync_message: isLocalOnlyNeed
              ? "Local GramEEE curator assignment enabled for shared-form needs."
              : "GRE-linked curator assignment ready.",
            gre_synced_at: new Date().toISOString(),
            updated_at: new Date().toISOString(),
          }, { onConflict: "email" })
          .select("id")
          .single();
        if (upsertCuratorError || !upsertedCurator) throw new Error(upsertCuratorError?.message || "Could not create the local curator record.");
        resolvedCuratorId = requireString(upsertedCurator.id);
      }
    } else {
      resolvedCuratorId = normalizedSelection;
    }
  }

  const { error } = await adminClient
    .from("gre_mis_needs")
    .update({
      curator_id: resolvedCuratorId,
      status: "Accepted",
      updated_at: new Date().toISOString(),
    })
    .eq("id", needId);
  if (error) throw new Error(error.message);

  await adminClient.from("gre_mis_need_updates").insert({
    need_id: needId,
    update_type: "curator_assignment",
    note: resolvedCuratorId ? `Curator assigned by ${actorEmail}.` : `Curator unassigned by ${actorEmail}.`,
    created_by_email: actorEmail,
  });

  if (/^\d+$/.test(requireString(needRow.id)) && sourceKind !== "shared_form_submission") {
    const { sessionId, data: greRequest } = await fetchGreJson(`/commons-request-management-service/api/v1/request?id=${encodeURIComponent(needId)}`);
    const payload = structuredClone(greRequest);
    const curationList = Array.isArray(payload.curationList) ? payload.curationList : [];
    const primaryCuration = curationList[0] || { requestId: Number(needId) };
    if (resolvedCuratorId) {
      const { data: curatorRow } = await adminClient
        .from("gre_mis_curators")
        .select("display_name, user_id")
        .eq("id", resolvedCuratorId)
        .maybeSingle();
      let greCuratorUserId: number | null = null;
      let greCuratorName = curatorRow?.display_name || primaryCuration.curatorUserName || null;
      const mappedUserId = requireString(curatorRow?.user_id);
      if (mappedUserId) {
        const { data: mappedUser } = await adminClient
          .from("gre_mis_users")
          .select("full_name, first_name, gre_user_id")
          .eq("id", mappedUserId)
          .maybeSingle();
        greCuratorUserId = parseNumber(mappedUser?.gre_user_id, 0) || null;
        greCuratorName =
          requireString(mappedUser?.full_name) ||
          requireString(mappedUser?.first_name) ||
          greCuratorName;
      }
      primaryCuration.curatorUserId = greCuratorUserId;
      primaryCuration.curatorUserName = greCuratorName || null;
    } else {
      primaryCuration.curatorUserName = null;
      primaryCuration.curatorUserId = null;
    }
    payload.curationList = [primaryCuration, ...curationList.slice(1)];
    try {
      await updateGreRequestJson("/commons-request-management-service/api/v1/request", payload, sessionId);
      const { data: verifiedRequest } = await fetchGreJson(
        `/commons-request-management-service/api/v1/request?id=${encodeURIComponent(needId)}`,
        sessionId,
      );
      const verifiedPayload = (verifiedRequest && typeof verifiedRequest === "object") ? verifiedRequest as Record<string, unknown> : {};
      const verifiedCurationList = Array.isArray(verifiedPayload.curationList) ? verifiedPayload.curationList : [];
      const verifiedPrimaryCuration = (verifiedCurationList[0] && typeof verifiedCurationList[0] === "object")
        ? verifiedCurationList[0] as Record<string, unknown>
        : {};
      const expectedCuratorId = parseNumber(primaryCuration.curatorUserId, 0) || null;
      const actualCuratorId = parseNumber(verifiedPrimaryCuration.curatorUserId, 0) || null;
      const expectedCuratorName = requireString(primaryCuration.curatorUserName);
      const actualCuratorName = requireString(verifiedPrimaryCuration.curatorUserName);
      if (curatorId && ((expectedCuratorId && actualCuratorId !== expectedCuratorId) || (expectedCuratorName && actualCuratorName !== expectedCuratorName))) {
        throw new Error(
          `GRE curator verification failed. Expected ${expectedCuratorName || expectedCuratorId || "assigned curator"}, found ${actualCuratorName || actualCuratorId || "blank"}.`,
        );
      }
    } catch (syncError) {
      await adminClient.from("gre_mis_need_updates").insert({
        need_id: needId,
        update_type: "curator_assignment_sync_warning",
        note: `Curator assignment saved in MIS, but GRE sync needs review: ${syncError instanceof Error ? syncError.message : "unknown error"}`,
        created_by_email: actorEmail,
      });
    }
  }

  return { ok: true };
}

async function approveNeed(needId: string, decision: string, reviewNotes: string, actorEmail: string) {
  const normalizedDecision = decision === "reject" ? "rejected" : "approved";
  const { error } = await adminClient
    .from("gre_mis_needs")
    .update({
      approval_status: normalizedDecision,
      updated_at: new Date().toISOString(),
    })
    .eq("id", needId);
  if (error) throw new Error(error.message);

  await adminClient.from("gre_mis_need_updates").insert({
    need_id: needId,
    update_type: normalizedDecision === "approved" ? "admin_approved" : "admin_rejected",
    note: reviewNotes || `Need ${normalizedDecision} by admin.`,
    created_by_email: actorEmail,
  });

  return { ok: true };
}

async function submitUpdateRequest(payload: Record<string, unknown>) {
  const needId = requireString(payload.needId);
  const curatorEmail = requireString(payload.curatorEmail).toLowerCase();
  const curatorName = requireString(payload.curatorName);
  if (!needId || !curatorEmail) throw new Error("Need and curator identity are required.");

  const { data: need, error: needError } = await adminClient
    .from("gre_mis_needs")
    .select("id, approval_status, curator_id, gre_mis_curators:curator_id(email, display_name)")
    .eq("id", needId)
    .single();

  if (needError || !need) throw new Error(needError?.message || "Need not found.");
  if (need.approval_status !== "approved") throw new Error("Only approved needs can receive curator updates.");

  const assignedEmail = Array.isArray(need.gre_mis_curators)
    ? String(need.gre_mis_curators[0]?.email || "").toLowerCase()
    : String((need.gre_mis_curators as { email?: string } | null)?.email || "").toLowerCase();

  if (!assignedEmail || assignedEmail !== curatorEmail) {
    throw new Error("Only the assigned curator can submit a status update for this need.");
  }

  const row = {
    need_id: needId,
    submitted_by_curator_name: curatorName || assignedEmail,
    submitted_by_curator_email: curatorEmail,
    proposed_status: requireString(payload.proposedStatus) || null,
    proposed_internal_status: requireString(payload.proposedInternalStatus) || null,
    proposed_next_action: requireString(payload.proposedNextAction) || null,
    proposed_curation_notes: requireString(payload.proposedCurationNotes) || null,
    proposed_curation_call_date: requireString(payload.proposedCurationCallDate) || null,
    proposed_curated_need: asStringArray(payload.proposedCuratedNeed),
    proposed_funding_mechanism: requireString(payload.proposedFundingMechanism) || null,
    proposed_seeker_provider_agreement: requireString(payload.proposedSeekerProviderAgreement) || null,
    proposed_solution_deployment_status: requireString(payload.proposedSolutionDeploymentStatus) || null,
    proposed_closure_date: requireString(payload.proposedClosureDate) || null,
    proposed_feedback_about_seeker: requireString(payload.proposedFeedbackAboutSeeker) || null,
    proposed_feedback_about_provider: requireString(payload.proposedFeedbackAboutProvider) || null,
    proposed_demand_broadcast_needed:
      typeof payload.proposedDemandBroadcastNeeded === "boolean" ? payload.proposedDemandBroadcastNeeded : null,
    proposed_solutions_shared_count:
      typeof payload.proposedSolutionsSharedCount === "number" ? payload.proposedSolutionsSharedCount : null,
    proposed_invited_providers_count:
      typeof payload.proposedInvitedProvidersCount === "number" ? payload.proposedInvitedProvidersCount : null,
    approval_status: "pending",
  };

  const { error } = await adminClient.from("gre_mis_update_requests").insert(row);
  if (error) throw new Error(error.message);
  return { ok: true };
}

function normalizeBooleanInput(value: unknown) {
  if (typeof value === "boolean") return value;
  const normalized = requireString(value).toLowerCase();
  if (!normalized) return null;
  if (["true", "yes", "1"].includes(normalized)) return true;
  if (["false", "no", "0"].includes(normalized)) return false;
  return null;
}

async function applyApprovedNeedPatch(requestRow: Record<string, unknown>, actorEmail: string) {
  const greSyncResult = await syncApprovedUpdateToGre(requestRow);

  const nextNeedPatch: Record<string, unknown> = {
    updated_at: new Date().toISOString(),
  };
  const changeParts: string[] = [];

  if (requestRow.proposed_status) {
    nextNeedPatch.status = requestRow.proposed_status;
    nextNeedPatch.last_status_change_at = new Date().toISOString();
    changeParts.push(`Status -> ${requestRow.proposed_status}`);
  }
  if (requestRow.proposed_internal_status) {
    nextNeedPatch.internal_status = requestRow.proposed_internal_status;
    nextNeedPatch.last_status_change_at = new Date().toISOString();
    changeParts.push(`Internal status -> ${requestRow.proposed_internal_status}`);
  }
  if (requestRow.proposed_next_action) {
    nextNeedPatch.next_action = requestRow.proposed_next_action;
    changeParts.push(`Next action -> ${requestRow.proposed_next_action}`);
  }
  if (requestRow.proposed_curation_notes) nextNeedPatch.curation_notes = requestRow.proposed_curation_notes;
  if (requestRow.proposed_curation_call_date) nextNeedPatch.curation_call_date = requestRow.proposed_curation_call_date;
  if (Array.isArray(requestRow.proposed_curated_need)) {
    nextNeedPatch.curated_need = asStringArray(requestRow.proposed_curated_need);
    const curatedLabel = asStringArray(requestRow.proposed_curated_need).length
      ? asStringArray(requestRow.proposed_curated_need).join(", ")
      : "(cleared)";
    changeParts.push(`Curated need -> ${curatedLabel}`);
  }
  if (requestRow.proposed_funding_mechanism) {
    nextNeedPatch.funding_mechanism = requestRow.proposed_funding_mechanism;
    changeParts.push(`Funding mechanism -> ${requestRow.proposed_funding_mechanism}`);
  }
  if (requestRow.proposed_seeker_provider_agreement) {
    nextNeedPatch.seeker_provider_agreement = requestRow.proposed_seeker_provider_agreement;
    changeParts.push(`Seeker/provider agreement -> ${requestRow.proposed_seeker_provider_agreement}`);
  }
  if (requestRow.proposed_solution_deployment_status) {
    nextNeedPatch.solution_deployment_status = requestRow.proposed_solution_deployment_status;
    changeParts.push(`Deployment status -> ${requestRow.proposed_solution_deployment_status}`);
  }
  if (requestRow.proposed_closure_date) {
    nextNeedPatch.closure_date = requestRow.proposed_closure_date;
    changeParts.push(`Closure date -> ${requestRow.proposed_closure_date}`);
  }
  if (requestRow.proposed_feedback_about_seeker) nextNeedPatch.feedback_about_seeker = requestRow.proposed_feedback_about_seeker;
  if (requestRow.proposed_feedback_about_provider) nextNeedPatch.feedback_about_provider = requestRow.proposed_feedback_about_provider;
  if (requestRow.proposed_demand_broadcast_needed !== null && requestRow.proposed_demand_broadcast_needed !== undefined) {
    nextNeedPatch.demand_broadcast_needed = requestRow.proposed_demand_broadcast_needed;
  }
  if (Number.isInteger(requestRow.proposed_solutions_shared_count)) {
    nextNeedPatch.solutions_shared_count = requestRow.proposed_solutions_shared_count;
  }
  if (Number.isInteger(requestRow.proposed_invited_providers_count)) {
    nextNeedPatch.invited_providers_count = requestRow.proposed_invited_providers_count;
  }

  const { error: patchError } = await adminClient.from("gre_mis_needs").update(nextNeedPatch).eq("id", requestRow.need_id);
  if (patchError) throw new Error(patchError.message);

  await adminClient.from("gre_mis_need_updates").insert({
    need_id: requestRow.need_id,
    update_type: "curation_updated",
    note: [
      changeParts.join(" | ") || "Curation update saved.",
      "Synced to GRE website.",
      greSyncResult.unsupportedFields?.length ? `Not synced to GRE: ${greSyncResult.unsupportedFields.join(", ")}.` : "",
    ].filter(Boolean).join(" "),
    created_by_email: actorEmail,
  });

  return greSyncResult;
}

async function directCuratorUpdate(payload: Record<string, unknown>, userCtx: { user: Record<string, unknown> }) {
  const needId = requireString(payload.needId);
  if (!needId) throw new Error("Need ID is required.");
  const actorEmail = requireString(userCtx.user.email).toLowerCase();
  const actorRole = requireString(userCtx.user.role).toLowerCase();
  const isAdminActor = isAdminLikeMisRole(actorRole);

  const { data: need, error: needError } = await adminClient
    .from("gre_mis_needs")
    .select("id, approval_status, curator_id, source_kind, gre_mis_curators:curator_id(user_id, email)")
    .eq("id", needId)
    .single();
  if (needError || !need) throw new Error(needError?.message || "Need not found.");
  if (need.approval_status !== "approved") throw new Error("Only approved needs can be updated.");

  const curatorRow = Array.isArray(need.gre_mis_curators) ? need.gre_mis_curators[0] : need.gre_mis_curators;
  const assignedUserId = requireString(curatorRow?.user_id);
  const assignedEmail = requireString(curatorRow?.email).toLowerCase();
  if (!isAdminActor && assignedUserId !== requireString(userCtx.user.id) && assignedEmail !== actorEmail) {
    throw new Error("You can edit curation only for needs assigned to you.");
  }

  const requestRow = {
    need_id: needId,
    proposed_status: requireString(payload.proposedStatus) || null,
    proposed_internal_status: requireString(payload.proposedInternalStatus) || null,
    proposed_next_action: requireString(payload.proposedNextAction) || null,
    proposed_curation_notes: requireString(payload.proposedCurationNotes) || null,
    proposed_curation_call_date: requireString(payload.proposedCurationCallDate) || null,
    proposed_curated_need: asStringArray(payload.proposedCuratedNeed),
    proposed_funding_mechanism: requireString(payload.proposedFundingMechanism) || null,
    proposed_seeker_provider_agreement: requireString(payload.proposedSeekerProviderAgreement) || null,
    proposed_solution_deployment_status: requireString(payload.proposedSolutionDeploymentStatus) || null,
    proposed_closure_date: requireString(payload.proposedClosureDate) || null,
    proposed_feedback_about_seeker: requireString(payload.proposedFeedbackAboutSeeker) || null,
    proposed_feedback_about_provider: requireString(payload.proposedFeedbackAboutProvider) || null,
    proposed_demand_broadcast_needed: normalizeBooleanInput(payload.proposedDemandBroadcastNeeded),
    proposed_solutions_shared_count: null,
    proposed_invited_providers_count: null,
  };

  const isLocalOnlyNeed = requireString(need.source_kind) === "shared_form_submission" || requireString(need.id).startsWith("FORM-");
  if (isLocalOnlyNeed) {
    const nextNeedPatch: Record<string, unknown> = {
      updated_at: new Date().toISOString(),
    };
    const changeParts: string[] = [];

    if (requestRow.proposed_status) {
      nextNeedPatch.status = requestRow.proposed_status;
      nextNeedPatch.last_status_change_at = new Date().toISOString();
      changeParts.push(`Status -> ${requestRow.proposed_status}`);
    }
    if (requestRow.proposed_internal_status) {
      nextNeedPatch.internal_status = requestRow.proposed_internal_status;
      nextNeedPatch.last_status_change_at = new Date().toISOString();
      changeParts.push(`Internal status -> ${requestRow.proposed_internal_status}`);
    }
    if (requestRow.proposed_next_action) {
      nextNeedPatch.next_action = requestRow.proposed_next_action;
      changeParts.push(`Next action -> ${requestRow.proposed_next_action}`);
    }
    if (requestRow.proposed_curation_notes) nextNeedPatch.curation_notes = requestRow.proposed_curation_notes;
    if (requestRow.proposed_curation_call_date) nextNeedPatch.curation_call_date = requestRow.proposed_curation_call_date;
    if (Array.isArray(requestRow.proposed_curated_need)) {
      nextNeedPatch.curated_need = asStringArray(requestRow.proposed_curated_need);
      const curatedLabel = asStringArray(requestRow.proposed_curated_need).length
        ? asStringArray(requestRow.proposed_curated_need).join(", ")
        : "(cleared)";
      changeParts.push(`Curated need -> ${curatedLabel}`);
    }
    if (requestRow.proposed_funding_mechanism) {
      nextNeedPatch.funding_mechanism = requestRow.proposed_funding_mechanism;
      changeParts.push(`Funding mechanism -> ${requestRow.proposed_funding_mechanism}`);
    }
    if (requestRow.proposed_seeker_provider_agreement) {
      nextNeedPatch.seeker_provider_agreement = requestRow.proposed_seeker_provider_agreement;
      changeParts.push(`Seeker/provider agreement -> ${requestRow.proposed_seeker_provider_agreement}`);
    }
    if (requestRow.proposed_solution_deployment_status) {
      nextNeedPatch.solution_deployment_status = requestRow.proposed_solution_deployment_status;
      changeParts.push(`Deployment status -> ${requestRow.proposed_solution_deployment_status}`);
    }
    if (requestRow.proposed_closure_date) {
      nextNeedPatch.closure_date = requestRow.proposed_closure_date;
      changeParts.push(`Closure date -> ${requestRow.proposed_closure_date}`);
    }
    if (requestRow.proposed_feedback_about_seeker) nextNeedPatch.feedback_about_seeker = requestRow.proposed_feedback_about_seeker;
    if (requestRow.proposed_feedback_about_provider) nextNeedPatch.feedback_about_provider = requestRow.proposed_feedback_about_provider;
    if (requestRow.proposed_demand_broadcast_needed !== null && requestRow.proposed_demand_broadcast_needed !== undefined) {
      nextNeedPatch.demand_broadcast_needed = requestRow.proposed_demand_broadcast_needed;
    }

    const { error: patchError } = await adminClient.from("gre_mis_needs").update(nextNeedPatch).eq("id", needId);
    if (patchError) throw new Error(patchError.message);

    await adminClient.from("gre_mis_need_updates").insert({
      need_id: needId,
      update_type: "curation_updated_local",
      note: [
        changeParts.join(" | ") || "Local GramEEE curation updated.",
        "This update remains on GramEEE only and was not synced back to the GRE website.",
      ].filter(Boolean).join(" "),
      created_by_email: actorEmail,
    });

    return {
      ok: true,
      message: "Local GramEEE curation update saved.",
      greSync: {
        synced: false,
        mode: "local_only",
        unsupportedFields: [],
      },
    };
  }

  const greSyncResult = await applyApprovedNeedPatch(requestRow, actorEmail);
  return {
    ok: true,
    message: "Curation update saved and synced to GRE.",
    greSync: {
      synced: true,
      unsupportedFields: greSyncResult.unsupportedFields || [],
    },
  };
}

async function debugGreNeedSync(needId: string) {
  const normalizedNeedId = requireString(needId);
  if (!normalizedNeedId) throw new Error("Need ID is required.");
  const { data: misNeed, error: misError } = await adminClient
    .from("gre_mis_needs")
    .select("id, organization_name, status, internal_status, next_action, curation_notes, curation_call_date, demand_broadcast_needed, updated_at, source_kind")
    .eq("id", normalizedNeedId)
    .maybeSingle();
  if (misError) throw new Error(misError.message);
  const { sessionId, data: greRequest } = await fetchGreJson(`/commons-request-management-service/api/v1/request?id=${encodeURIComponent(normalizedNeedId)}`);
  const payload = (greRequest && typeof greRequest === "object") ? greRequest as Record<string, unknown> : {};
  const curationList = Array.isArray(payload.curationList) ? payload.curationList : [];
  const primaryCuration = (curationList[0] && typeof curationList[0] === "object") ? curationList[0] as Record<string, unknown> : {};
  return {
    ok: true,
    needId: normalizedNeedId,
    mis: misNeed || null,
    gre: {
      id: requireString(payload.id || payload.requestId),
      status: payload.status || null,
      internalStatus: payload.internalStatus || null,
      nextAction: payload.nextAction || null,
      curationListCount: curationList.length,
      primaryCuration: {
        curatorUserId: primaryCuration.curatorUserId ?? null,
        curatorUserName: primaryCuration.curatorUserName ?? null,
        callDate: primaryCuration.callDate ?? null,
        callDetails: primaryCuration.callDetails ?? null,
        demandBroadcastNeed: primaryCuration.demandBroadcastNeed ?? null,
      },
      curationListSummary: curationList.map((entry, index) => {
        const row = (entry && typeof entry === "object") ? entry as Record<string, unknown> : {};
        return {
          index,
          id: row.id ?? row.curationId ?? null,
          curatorUserId: row.curatorUserId ?? null,
          curatorUserName: row.curatorUserName ?? null,
          callDate: row.callDate ?? null,
          hasCallDetails: Boolean(requireString(row.callDetails)),
          demandBroadcastNeed: row.demandBroadcastNeed ?? null,
          updatedAt: row.updatedAt ?? row.lastUpdatedAt ?? null,
        };
      }),
    },
    sessionActive: Boolean(sessionId),
  };
}

async function debugGreNeedSyncService(needId: string) {
  const normalizedNeedId = requireString(needId);
  if (!normalizedNeedId) throw new Error("Need ID is required.");
  const { data: misNeed, error: misError } = await adminClient
    .from("gre_mis_needs")
    .select("*")
    .eq("id", normalizedNeedId)
    .maybeSingle();
  if (misError) throw new Error(misError.message);
  const { sessionId, data: greRequest } = await fetchGreJson(`/commons-request-management-service/api/v1/request?id=${encodeURIComponent(normalizedNeedId)}`);
  const payload = (greRequest && typeof greRequest === "object") ? greRequest as Record<string, unknown> : {};
  const curationList = Array.isArray(payload.curationList) ? payload.curationList : [];
  return {
    ok: true,
    needId: normalizedNeedId,
    mis: misNeed || null,
    greRaw: payload,
    greCurationList: curationList,
    sessionActive: Boolean(sessionId),
  };
}

async function reviewUpdateRequest(requestId: string, decision: string, reviewNotes: string, actorEmail: string) {
  const { data: requestRow, error: requestError } = await adminClient
    .from("gre_mis_update_requests")
    .select("*")
    .eq("id", requestId)
    .single();
  if (requestError || !requestRow) throw new Error(requestError?.message || "Update request not found.");

  if (decision === "reject") {
    const { error } = await adminClient
      .from("gre_mis_update_requests")
      .update({
        approval_status: "rejected",
        review_notes: reviewNotes || "Rejected by admin.",
        reviewed_by_email: actorEmail,
        reviewed_at: new Date().toISOString(),
      })
      .eq("id", requestId);
    if (error) throw new Error(error.message);
    return { ok: true };
  }

  const greSyncResult = await applyApprovedNeedPatch(requestRow, actorEmail);

  const { error: reviewError } = await adminClient
    .from("gre_mis_update_requests")
    .update({
      approval_status: "approved",
      review_notes: reviewNotes || "Approved by admin.",
      reviewed_by_email: actorEmail,
      reviewed_at: new Date().toISOString(),
    })
    .eq("id", requestId);
  if (reviewError) throw new Error(reviewError.message);

  return {
    ok: true,
    greSync: {
      synced: true,
      unsupportedFields: greSyncResult.unsupportedFields || [],
    },
  };
}

async function upsertOption(optionType: string, label: string) {
  const normalizedType = requireString(optionType);
  const normalizedLabel = requireString(label);
  if (!normalizedType || !normalizedLabel) throw new Error("Option type and label are required.");

  const { data: existing } = await adminClient
    .from("gre_mis_options")
    .select("id")
    .eq("option_type", normalizedType)
    .ilike("label", normalizedLabel)
    .maybeSingle();

  if (existing?.id) {
    const { error } = await adminClient.from("gre_mis_options").update({ is_active: true }).eq("id", existing.id);
    if (error) throw new Error(error.message);
    return { ok: true, reused: true };
  }

  const { data: current } = await adminClient
    .from("gre_mis_options")
    .select("sort_order")
    .eq("option_type", normalizedType)
    .order("sort_order", { ascending: false })
    .limit(1);

  const nextSortOrder = Number(current?.[0]?.sort_order || 0) + 1;
  const { error } = await adminClient.from("gre_mis_options").insert({
    option_type: normalizedType,
    label: normalizedLabel,
    sort_order: nextSortOrder,
  });
  if (error) throw new Error(error.message);
  return { ok: true };
}

const DEFAULT_PROVIDER_INTRO_TEMPLATE = `Hello {{providerName}},

We are reaching out to you from GRE platform to connect you with {{seekerLabel}}, marked in copy of this mail. They have a need for {{problemStatement}} and your solution of {{viewLink}} may be of interest to them. We would suggest you to connect mutually and take this forward. Do reach out to us if you would like us to help facilitate the conversation.

Regards,
Team GRE`;

const DEFAULT_CURATOR_FORWARD_TEMPLATE = `Hello {{assignedCuratorName}},

For the needs of {{seekerLabel}} for {{problemStatement}}, i feel the solution {{viewLink}} by {{providerName}} may be of interest to you. Kindly review and do the needful.

Regards,
{{actorName}}`;

const DEFAULT_SOLUTION_SEEKER_TEMPLATE = `Hello {{seekerLabel}},

Greetings from Team GRE.

We have identified against your stated need of

"{{problemStatement}}",

{{providerName}} offers a solution of {{solutionName}},

Link to solution {{viewLink}}.

Do let us know if this solution is of interest to you enabling us to help coordinate a meeting.

Regards,

Team GRE`;

const DEFAULT_NEED_SEEKER_TEMPLATE = `Hello {{seekerLabel}},

I have reviewed your broadcasted need for {{thematicArea}} on the GramEEE GRE Platform. We feel we can offer you a solution for this and would like to connect with you.

Regards,
{{actorName}}
{{actorPhone}}`;
const LSH_SETTINGS_KEY = "lsh_management";
const DEFAULT_LSH_MANAGEMENT_SETTINGS = {
  contactEmails: ["subekkumar@pradan.net"],
  helpCcEmails: ["help@greenruraleconomy.in"],
  requestSupportDraft: `Hello Team LSH,

We are looking at some of our needs in the Livestock domain and would like to connect with your team. Do reach out to us.

Regards,
{{name}}
{{organisation}}`,
  emailProviderDraft: `Hello Team LSH,

We would like to know more about {{offeringName}} and request a follow-up from your team.

Regards,
{{name}}
{{organisation}}`,
};
const IMPACT_VIEW_AUDIT_LOG_KEY = "impact_view_audit_log";
const IMPACT_EMAIL_AUDIT_LOG_KEY = "impact_email_audit_log";
const MAX_IMPACT_AUDIT_ENTRIES = 2500;

function normalizeGreMailTemplateLegacyPlaceholders(template: string, kind: "provider" | "curator" | "needSeeker") {
  const normalized = requireString(template)
    .replace(/\{\{\s*solutionSeeker\s*\}\}/g, "{{seekerLabel}}")
    .replace(/\{\{\s*solutionProvider\s*\}\}/g, "{{providerName}}")
    .replace(/\{\{\s*curatorName\s*\}\}/g, "{{actorName}}")
    .replace(/\{\{\s*needTheme\s*\}\}/g, "{{thematicArea}}")
    .replace(/\{\{\s*senderName\s*\}\}/g, "{{actorName}}")
    .replace(/\{\{\s*senderPhone\s*\}\}/g, "{{actorPhone}}");
  if (kind === "curator") {
    return normalized.replace(/\{\{\s*providerLabel\s*\}\}/g, "{{providerName}}");
  }
  return normalized;
}

function renderGreMailTemplateText(template: string, replacements: Record<string, unknown>) {
  return requireString(template).replace(/\{\{\s*([a-zA-Z0-9_]+)\s*\}\}/g, (_, key) => {
    const value = replacements[key];
    return value === undefined || value === null ? "" : String(value);
  });
}

async function getGreMisSetting(key: string) {
  const { data, error } = await adminClient
    .from("gre_mis_settings")
    .select("key, value_text, value_json, updated_by_email, updated_at")
    .eq("key", key)
    .maybeSingle();
  if (error) throw new Error(error.message);
  return data as Record<string, unknown> | null;
}

async function upsertGreMisSetting(key: string, valueText: string | null, valueJson: Record<string, unknown>, updatedByEmail: string) {
  const { error } = await adminClient
    .from("gre_mis_settings")
    .upsert({
      key,
      value_text: valueText,
      value_json: valueJson,
      updated_by_email: updatedByEmail || null,
      updated_at: new Date().toISOString(),
    });
  if (error) throw new Error(error.message);
}

async function appendGreMisAuditEntry(
  key: string,
  entry: Record<string, unknown>,
  updatedByEmail: string,
) {
  const current = await getGreMisSetting(key);
  const currentValue = current?.value_json && typeof current.value_json === "object"
    ? current.value_json as Record<string, unknown>
    : {};
  const currentEntries = Array.isArray(currentValue.entries) ? currentValue.entries : [];
  const entries = [entry, ...currentEntries].slice(0, MAX_IMPACT_AUDIT_ENTRIES);
  await upsertGreMisSetting(key, null, { entries }, updatedByEmail || "system:audit");
}

async function getImpactAuditLogs() {
  const [viewRow, emailRow] = await Promise.all([
    getGreMisSetting(IMPACT_VIEW_AUDIT_LOG_KEY),
    getGreMisSetting(IMPACT_EMAIL_AUDIT_LOG_KEY),
  ]);
  return {
    viewLogs: Array.isArray((viewRow?.value_json as Record<string, unknown> | null)?.entries)
      ? ((viewRow?.value_json as Record<string, unknown>).entries as unknown[])
      : [],
    emailLogs: Array.isArray((emailRow?.value_json as Record<string, unknown> | null)?.entries)
      ? ((emailRow?.value_json as Record<string, unknown>).entries as unknown[])
      : [],
  };
}

async function getGreMailTemplates() {
  const [providerRow, curatorRow, solutionSeekerRow, needSeekerRow, inboundSyncRow, lshManagementRow] = await Promise.all([
    getGreMisSetting("provider_intro_template"),
    getGreMisSetting("curator_forward_template"),
    getGreMisSetting("solution_seeker_intro_template"),
    getGreMisSetting("need_seeker_intro_template"),
    getGreMisSetting("inbound_auto_sync"),
    getGreMisSetting(LSH_SETTINGS_KEY),
  ]);
  const inboundEnabledValue = (inboundSyncRow?.value_json as Record<string, unknown> | null)?.enabled;
  const inboundEnabled = typeof inboundEnabledValue === "boolean" ? Boolean(inboundEnabledValue) : true;
  const lshValue = lshManagementRow?.value_json && typeof lshManagementRow.value_json === "object"
    ? lshManagementRow.value_json as Record<string, unknown>
    : {};
  if (!inboundEnabled) {
    await upsertGreMisSetting(
      "inbound_auto_sync",
      null,
      { enabled: true },
      "system:auto-sync-default",
    );
  }

  return {
    providerIntroTemplate: normalizeGreMailTemplateLegacyPlaceholders(
      requireString(providerRow?.value_text) || DEFAULT_PROVIDER_INTRO_TEMPLATE,
      "provider",
    ),
    curatorForwardTemplate: normalizeGreMailTemplateLegacyPlaceholders(
      requireString(curatorRow?.value_text) || DEFAULT_CURATOR_FORWARD_TEMPLATE,
      "curator",
    ),
    solutionSeekerTemplate: normalizeGreMailTemplateLegacyPlaceholders(
      requireString(solutionSeekerRow?.value_text) || DEFAULT_SOLUTION_SEEKER_TEMPLATE,
      "needSeeker",
    ),
    needSeekerTemplate: normalizeGreMailTemplateLegacyPlaceholders(
      requireString(needSeekerRow?.value_text) || DEFAULT_NEED_SEEKER_TEMPLATE,
      "needSeeker",
    ),
    inboundAutoSyncEnabled: true,
    lshContactEmails: normalizeEmailList(lshValue.contactEmails).length
      ? normalizeEmailList(lshValue.contactEmails)
      : DEFAULT_LSH_MANAGEMENT_SETTINGS.contactEmails,
    lshHelpCcEmails: normalizeEmailList(lshValue.helpCcEmails).length
      ? normalizeEmailList(lshValue.helpCcEmails)
      : DEFAULT_LSH_MANAGEMENT_SETTINGS.helpCcEmails,
    lshRequestSupportTemplate: requireString(lshValue.requestSupportDraft) || DEFAULT_LSH_MANAGEMENT_SETTINGS.requestSupportDraft,
    lshEmailProviderTemplate: requireString(lshValue.emailProviderDraft) || DEFAULT_LSH_MANAGEMENT_SETTINGS.emailProviderDraft,
  };
}

async function saveGreMailTemplates(
  actorEmail: string,
  providerIntroTemplate: string,
  curatorForwardTemplate: string,
  solutionSeekerTemplate: string,
  needSeekerTemplate: string,
  inboundAutoSyncEnabled: boolean | null,
  lshContactEmails: unknown,
  lshHelpCcEmails: unknown,
  lshRequestSupportTemplate: string,
  lshEmailProviderTemplate: string,
) {
  await upsertGreMisSetting(
    "provider_intro_template",
    normalizeGreMailTemplateLegacyPlaceholders(
      requireString(providerIntroTemplate) || DEFAULT_PROVIDER_INTRO_TEMPLATE,
      "provider",
    ),
    {},
    actorEmail,
  );
  await upsertGreMisSetting(
    "curator_forward_template",
    normalizeGreMailTemplateLegacyPlaceholders(
      requireString(curatorForwardTemplate) || DEFAULT_CURATOR_FORWARD_TEMPLATE,
      "curator",
    ),
    {},
    actorEmail,
  );
  await upsertGreMisSetting(
    "solution_seeker_intro_template",
    normalizeGreMailTemplateLegacyPlaceholders(
      requireString(solutionSeekerTemplate) || DEFAULT_SOLUTION_SEEKER_TEMPLATE,
      "needSeeker",
    ),
    {},
    actorEmail,
  );
  await upsertGreMisSetting(
    "need_seeker_intro_template",
    normalizeGreMailTemplateLegacyPlaceholders(
      requireString(needSeekerTemplate) || DEFAULT_NEED_SEEKER_TEMPLATE,
      "needSeeker",
    ),
    {},
    actorEmail,
  );
  if (typeof inboundAutoSyncEnabled === "boolean") {
    await upsertGreMisSetting(
      "inbound_auto_sync",
      null,
      { enabled: inboundAutoSyncEnabled },
      actorEmail,
    );
  }
  await upsertGreMisSetting(
    LSH_SETTINGS_KEY,
    null,
    {
      contactEmails: normalizeEmailList(lshContactEmails).length
        ? normalizeEmailList(lshContactEmails)
        : DEFAULT_LSH_MANAGEMENT_SETTINGS.contactEmails,
      helpCcEmails: normalizeEmailList(lshHelpCcEmails).length
        ? normalizeEmailList(lshHelpCcEmails)
        : DEFAULT_LSH_MANAGEMENT_SETTINGS.helpCcEmails,
      requestSupportDraft: requireString(lshRequestSupportTemplate) || DEFAULT_LSH_MANAGEMENT_SETTINGS.requestSupportDraft,
      emailProviderDraft: requireString(lshEmailProviderTemplate) || DEFAULT_LSH_MANAGEMENT_SETTINGS.emailProviderDraft,
    },
    actorEmail,
  );
  return {
    ok: true,
    ...(await getGreMailTemplates()),
  };
}

async function sendProviderIntro(
  needId: string,
  providerEmail: string,
  actor: { id: string; name: string; email: string; role: string },
  providerName = "",
  offeringId = "",
  mailDraft: Record<string, unknown> | null = null,
) {
  const { data: need, error: needError } = await adminClient
    .from("gre_mis_needs")
    .select("id, organization_name, seeker_email, contact_person, problem_statement, state, district, curated_need, curator_id, gre_mis_curators:curator_id(user_id, email, display_name)")
    .eq("id", needId)
    .single();
  if (needError || !need) throw new Error(needError?.message || "Need not found.");
  const curator = Array.isArray(need.gre_mis_curators) ? need.gre_mis_curators[0] : need.gre_mis_curators;
  const curatorEmail = requireString(curator?.email).toLowerCase();
  const curatorUserId = requireString(curator?.user_id);
  const actorRole = requireString(actor.role).toLowerCase();
  const actorEmail = requireString(actor.email).toLowerCase();
  const actorName = requireString(actor.name) || actorEmail;
  const providerLabel = requireString(providerName) || "Solution Provider";
  const seekerLabel = requireString(need.organization_name) || requireString(need.contact_person) || "the solution seeker";
  const problemStatement = requireString(need.problem_statement) || "the shared problem statement";
  const viewLink = offeringId
    ? `${greMisBaseUrl}/offering-detail.html?offering_id=${encodeURIComponent(offeringId)}`
    : greMisBaseUrl;
  const isAssignedCurator = (
    (curatorUserId && curatorUserId === requireString(actor.id)) ||
    (curatorEmail && curatorEmail === actorEmail)
  );
  const shouldForwardToAssignedCurator = actorRole === "curator" || (
    ["admin", "moderator"].includes(actorRole) &&
    Boolean(curatorEmail) &&
    !isAssignedCurator
  );

  if (shouldForwardToAssignedCurator) {
    if (!curatorEmail) throw new Error("No curator is assigned to this need yet.");
    const draftTo = requireString(mailDraft?.to || curatorEmail).toLowerCase();
    const draftCc = requireString(mailDraft?.cc || actorEmail).toLowerCase();
    const subject = requireString(mailDraft?.subject) || `GRE review suggestion for ${seekerLabel}`;
    const body = requireString(mailDraft?.body) || [
      `Hello ${curator?.display_name || "Assigned Curator"},`,
      "",
      `For the needs of ${seekerLabel} for ${problemStatement}, i feel the solution ${viewLink} by ${providerLabel} may be of interest to you. Kindly review and do the needful.`,
      "",
      "Regards,",
      actorName,
    ].join("\n");

    await sendEmail({
      to: draftTo,
      cc: draftCc,
      subject,
      body,
      mailbox: "help",
    });

    await adminClient.from("gre_mis_email_log").insert({
      need_id: need.id,
      recipient_email: draftTo,
      cc_email: draftCc,
      subject,
      body_preview: body.slice(0, 1000),
      sent_by_email: actorEmail,
    });
    await appendGreMisAuditEntry(IMPACT_EMAIL_AUDIT_LOG_KEY, {
      id: crypto.randomUUID(),
      logged_at: new Date().toISOString(),
      kind: "email",
      surface: "gre-mis",
      action: "email_to_curator",
      counter_key: "connections_made",
      sender_email: actorEmail,
      sender_name: actorName,
      sender_role: actorRole,
      recipient_email: draftTo,
      cc_email: draftCc,
      reply_to: actorEmail,
      subject,
      item_id: requireString(offeringId),
      item_label: providerLabel,
      item_source: "gre-mis",
      need_id: requireString(need.id),
      provider_name: providerLabel,
      seeker_name: seekerLabel,
    }, actorEmail);

    await adminClient.from("gre_mis_need_updates").insert({
      need_id: need.id,
      update_type: "provider_intro_curator_forwarded",
      note: `Provider suggestion for ${providerEmail} was forwarded to the assigned curator ${draftTo}.`,
      created_by_email: actorEmail,
    });

    return { ok: true, message: `Assigned curator notified from ${helpGmailSenderEmail}.` };
  }

  const subject = requireString(mailDraft?.subject) || `GRE introduction: ${seekerLabel} challenge for your consideration`;
  const body = requireString(mailDraft?.body) || [
    `Hello ${providerLabel},`,
    "",
    `We are reaching out to you from GRE platform to connect you with ${seekerLabel}, marked in copy of this mail. They have a need for ${problemStatement} and your solution of ${viewLink} may be of interest to them. We would suggest you to connect mutually and take this forward. Do reach out to us if you would like us to help facilitate the conversation.`,
    "",
    "Regards,",
    "Team GRE",
  ].join("\n");
  const ccEmail = requireString(mailDraft?.cc) || [requireString(need.seeker_email).toLowerCase(), curatorEmail].filter(Boolean).join(", ");
  const toEmail = requireString(mailDraft?.to || providerEmail).toLowerCase();

  await sendEmail({
    to: toEmail,
    cc: ccEmail,
    subject,
    body,
    mailbox: "help",
  });

  await adminClient.from("gre_mis_email_log").insert({
    need_id: need.id,
    recipient_email: toEmail,
    cc_email: ccEmail,
    subject,
    body_preview: body.slice(0, 1000),
    sent_by_email: actorEmail,
  });
  await appendGreMisAuditEntry(IMPACT_EMAIL_AUDIT_LOG_KEY, {
    id: crypto.randomUUID(),
    logged_at: new Date().toISOString(),
    kind: "email",
    surface: "gre-mis",
    action: "email_to_provider",
    counter_key: "connections_made",
    sender_email: actorEmail,
    sender_name: actorName,
    sender_role: actorRole,
    recipient_email: toEmail,
    cc_email: ccEmail,
    reply_to: actorEmail,
    subject,
    item_id: requireString(offeringId),
    item_label: providerLabel,
    item_source: "gre-mis",
    need_id: requireString(need.id),
    provider_name: providerLabel,
    seeker_name: seekerLabel,
  }, actorEmail);

  await adminClient.from("gre_mis_need_updates").insert({
    need_id: need.id,
    update_type: "provider_intro_sent",
    note: `Provider introduction sent to ${providerEmail} with seeker and assigned curator copied.`,
    created_by_email: actorEmail,
  });

  return { ok: true, message: `Provider introduction email sent from ${helpGmailSenderEmail}.` };
}

async function sendSolutionSeekerIntro(
  needId: string,
  providerEmail: string,
  actor: { id: string; name: string; email: string; role: string },
  providerName = "",
  offeringId = "",
  mailDraft: Record<string, unknown> | null = null,
) {
  const { data: need, error: needError } = await adminClient
    .from("gre_mis_needs")
    .select("id, organization_name, seeker_email, contact_person, problem_statement, curation_notes, solutions_shared_count, curator_id, gre_mis_curators:curator_id(user_id, email, display_name)")
    .eq("id", needId)
    .single();
  if (needError || !need) throw new Error(needError?.message || "Need not found.");

  const seekerEmail = requireString(need.seeker_email).toLowerCase();
  if (!seekerEmail) throw new Error("This need does not yet have a seeker email address.");

  const curator = Array.isArray(need.gre_mis_curators) ? need.gre_mis_curators[0] : need.gre_mis_curators;
  const curatorEmail = requireString(curator?.email).toLowerCase();
  const actorEmail = requireString(actor.email).toLowerCase();
  const seekerLabel = requireString(need.contact_person) || requireString(need.organization_name) || "Seeker";
  const problemStatement = requireString(need.problem_statement) || "the stated need";
  const providerLabel = requireString(providerName) || "Solution Provider";
  const solutionName = requireString(mailDraft?.solutionName) || providerLabel;
  const viewLink = offeringId
    ? `${greMisBaseUrl}/offering-detail.html?offering_id=${encodeURIComponent(offeringId)}`
    : greMisBaseUrl;
  const templates = await getGreMailTemplates();
  const subject = requireString(mailDraft?.subject) || `GRE solution identified for ${seekerLabel}`;
  const body = requireString(mailDraft?.body) || normalizeGreMailTemplateLegacyPlaceholders(
    renderGreMailTemplateText(templates.solutionSeekerTemplate || DEFAULT_SOLUTION_SEEKER_TEMPLATE, {
      seekerLabel,
      problemStatement,
      providerName: providerLabel,
      solutionName,
      viewLink,
    }),
    "needSeeker",
  );
  const ccEmail = requireString(mailDraft?.cc) || ["help@greenruraleconomy.in", curatorEmail].filter(Boolean).join(", ");
  const toEmail = requireString(mailDraft?.to || seekerEmail).toLowerCase();

  await sendEmail({
    to: toEmail,
    cc: ccEmail,
    subject,
    body,
    mailbox: "solution",
  });

  await adminClient.from("gre_mis_email_log").insert({
    need_id: need.id,
    recipient_email: toEmail,
    cc_email: ccEmail,
    subject,
    body_preview: body.slice(0, 1000),
    sent_by_email: actorEmail,
  });
  await appendGreMisAuditEntry(IMPACT_EMAIL_AUDIT_LOG_KEY, {
    id: crypto.randomUUID(),
    logged_at: new Date().toISOString(),
    kind: "email",
    surface: "gre-mis",
    action: "email_to_seeker",
    counter_key: "connections_made",
    sender_email: actorEmail,
    sender_name: actor.name,
    sender_role: actor.role,
    recipient_email: toEmail,
    cc_email: ccEmail,
    reply_to: actorEmail,
    subject,
    item_id: requireString(offeringId),
    item_label: solutionName,
    item_source: "gre-mis",
    need_id: requireString(need.id),
    provider_name: providerLabel,
    seeker_name: seekerLabel,
  }, actorEmail);

  const nextNotes = appendSharedSolutionToCurationNotes(
    requireString(need.curation_notes),
    solutionName,
    providerLabel,
    viewLink,
  );
  const nextSharedCount = Math.max(parseNumber(need.solutions_shared_count, 0), 1);
  await adminClient
    .from("gre_mis_needs")
    .update({
      curation_notes: nextNotes || null,
      solutions_shared_count: nextSharedCount,
      updated_at: new Date().toISOString(),
    })
    .eq("id", need.id);

  await adminClient.from("gre_mis_need_updates").insert({
    need_id: need.id,
    update_type: "solution_shared_with_seeker",
    note: `Solution ${solutionName} by ${providerLabel} was shared with seeker ${toEmail}.`,
    created_by_email: actorEmail,
  });

  return { ok: true, message: `Seeker outreach email sent from ${solutionGmailSenderEmail}.` };
}

async function buildNeedSeekerIntroDraft(needId: string, grameeeAccessToken: string, grameeeUserSummary: unknown = null) {
  const normalizedToken = requireString(grameeeAccessToken);
  const authUser = normalizedToken
    ? await resolveGrameeeAuthUser(normalizedToken)
    : resolveGrameeeSummaryFallback(grameeeUserSummary);
  if (!authUser) {
    throw new Error("Your GramEEE login could not be verified. Please sign in again.");
  }
  if (authUser.user_metadata?.grameee_removed === true || authUser.app_metadata?.grameee_removed === true) {
    throw new Error("This GramEEE account is no longer active.");
  }

  const userMetadata = (authUser.user_metadata || authUser.raw_user_meta_data || {}) as Record<string, unknown>;
  const actorEmail = requireString(authUser.email).toLowerCase();
  if (!actorEmail) throw new Error("Your GramEEE login is missing an email address.");
  const actorName =
    requireString(userMetadata.full_name) ||
    requireString(userMetadata.first_name) ||
    requireString(userMetadata.username) ||
    actorEmail;
  const actorPhone = requireString(userMetadata.phone);

  const { data: need, error: needError } = await adminClient
    .from("gre_mis_needs")
    .select("id, organization_name, contact_person, seeker_email, state, district, problem_statement, ai_thematic_area, override_thematic_area, curator_id, gre_mis_curators:curator_id(email, display_name)")
    .eq("id", needId)
    .single();
  if (needError || !need) throw new Error(needError?.message || "Need not found.");

  const seekerEmail = requireString(need.seeker_email).toLowerCase();
  if (!seekerEmail) throw new Error("This need does not yet have a seeker email address.");

  const curator = Array.isArray(need.gre_mis_curators) ? need.gre_mis_curators[0] : need.gre_mis_curators;
  const curatorEmail = requireString(curator?.email).toLowerCase();
  const seekerLabel = requireString(need.contact_person) || requireString(need.organization_name) || "Solution Seeker";
  const thematicArea =
    requireString((need as Record<string, unknown>).override_thematic_area) ||
    requireString((need as Record<string, unknown>).ai_thematic_area) ||
    "this need";
  const templates = await getGreMailTemplates();
  const subject = `Response to your GramEEE GRE need on ${thematicArea}`;
  const body = normalizeGreMailTemplateLegacyPlaceholders(
    renderGreMailTemplateText(templates.needSeekerTemplate || DEFAULT_NEED_SEEKER_TEMPLATE, {
      seekerLabel,
      thematicArea,
      actorName,
      actorPhone,
    }),
    "needSeeker",
  );
  const ccEmail = [actorEmail, "solution@greenruraleconomy.in", curatorEmail]
    .filter(Boolean)
    .filter((value, index, list) => list.indexOf(value) === index)
    .join(", ");

  return {
    needId: requireString(need.id),
    actorEmail,
    actorName,
    seekerEmail,
    seekerLabel,
    thematicArea,
    subject,
    body,
    ccEmail,
    from: `Team GRE <${helpGmailSenderEmail}>`,
    replyTo: actorEmail,
  };
}

async function sendNeedSeekerIntro(needId: string, grameeeAccessToken: string, grameeeUserSummary: unknown = null) {
  const draft = await buildNeedSeekerIntroDraft(needId, grameeeAccessToken, grameeeUserSummary);

  await sendEmail({
    to: draft.seekerEmail,
    cc: draft.ccEmail,
    subject: draft.subject,
    body: draft.body,
    mailbox: "help",
  });

  await adminClient.from("gre_mis_email_log").insert({
    need_id: draft.needId,
    recipient_email: draft.seekerEmail,
    cc_email: draft.ccEmail,
    subject: draft.subject,
    body_preview: draft.body.slice(0, 1000),
    sent_by_email: draft.actorEmail,
  });
  await appendGreMisAuditEntry(IMPACT_EMAIL_AUDIT_LOG_KEY, {
    id: crypto.randomUUID(),
    logged_at: new Date().toISOString(),
    kind: "email",
    surface: "needs-map",
    action: "reach_out_to_seeker",
    counter_key: "connections_made",
    sender_email: draft.actorEmail,
    sender_name: draft.actorName,
    sender_role: "user",
    recipient_email: draft.seekerEmail,
    cc_email: draft.ccEmail,
    reply_to: draft.actorEmail,
    subject: draft.subject,
    item_id: "",
    item_label: draft.thematicArea,
    item_source: "needs-map",
    need_id: requireString(draft.needId),
    seeker_name: draft.seekerLabel,
  }, draft.actorEmail);

  await adminClient.from("gre_mis_need_updates").insert({
    need_id: draft.needId,
    update_type: "needs_map_seeker_intro_sent",
    note: `Needs Map outreach mail sent to seeker ${draft.seekerEmail} by ${draft.actorEmail}.`,
    created_by_email: draft.actorEmail,
  });

  return { ok: true, message: `Seeker introduction email sent from ${helpGmailSenderEmail}.` };
}

async function sendCuratorMessage(
  needId: string,
  actorName: string,
  actorEmail: string,
  message: string,
) {
  const { data: need, error } = await adminClient
    .from("gre_mis_needs")
    .select("id, organization_name, seeker_email, contact_person, problem_statement, state, district, gre_mis_curators:curator_id(email, display_name)")
    .eq("id", needId)
    .single();
  if (error || !need) throw new Error(error?.message || "Need not found.");
  const curator = Array.isArray(need.gre_mis_curators) ? need.gre_mis_curators[0] : need.gre_mis_curators;
  const curatorEmail = requireString(curator?.email).toLowerCase();
  if (!curatorEmail) throw new Error("No curator is assigned to this need yet.");

  const subject = `GRE MIS message on Need ${need.id}: ${need.organization_name}`;
  const body = [
    `Hello ${curator?.display_name || "Curator"},`,
    "",
    `${actorName || actorEmail} has sent a message from the GRE MIS dashboard regarding the need below.`,
    "",
    `Organization: ${need.organization_name}`,
    `Contact person: ${need.contact_person || "Not provided"}`,
    `Location: ${need.state || "Not provided"}${need.district ? ` / ${need.district}` : ""}`,
    "",
    "Problem statement:",
    `${need.problem_statement}`,
    "",
    "Message:",
    message || "No additional message provided.",
    "",
    `Reply to: ${actorEmail}`,
    "",
    "Regards,",
    "Green Rural Economy",
  ].join("\n");

  await sendEmail({
    to: curatorEmail,
    cc: actorEmail,
    subject,
    body,
  });

  await adminClient.from("gre_mis_email_log").insert({
    need_id: need.id,
    recipient_email: curatorEmail,
    cc_email: actorEmail,
    subject,
    body_preview: body.slice(0, 1000),
    sent_by_email: actorEmail,
  });
  await appendGreMisAuditEntry(IMPACT_EMAIL_AUDIT_LOG_KEY, {
    id: crypto.randomUUID(),
    logged_at: new Date().toISOString(),
    kind: "email",
    surface: "gre-mis",
    action: "curator_message",
    counter_key: "connections_made",
    sender_email: actorEmail,
    sender_name: actorName,
    sender_role: "curator",
    recipient_email: curatorEmail,
    cc_email: actorEmail,
    reply_to: actorEmail,
    subject,
    item_source: "gre-mis",
    need_id: requireString(need.id),
    seeker_name: requireString(need.contact_person) || requireString(need.organization_name),
  }, actorEmail);

  return { ok: true, message: `Message sent to ${curator?.display_name || "the curator"}.` };
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  try {
    const payload = (await req.json().catch(() => ({}))) as Record<string, unknown>;
    const action = requireString(payload.action);

    if (action === "searchLgdGeographies") {
      return jsonResponse(await searchLgdGeographies(requireString(payload.query)));
    }

    if (action === "translateTextBatch") {
      const texts = asStringArray(payload.texts);
      const source = requireString(payload.source) || "en";
      const target = requireString(payload.target) || "hi";
      return jsonResponse({
        ok: true,
        translations: await Promise.all(
          texts.map(async (text) => ({
            sourceText: text,
            translatedText: await translateTextWithLibreTranslate(text, source, target),
          })),
        ),
      });
    }

    if (action === "suggestSolutionTags") {
      return jsonResponse(await suggestSolutionTags(
        (payload.payload && typeof payload.payload === "object") ? payload.payload as Record<string, unknown> : {},
      ));
    }

    if (action === "suggestNeedTags") {
      return jsonResponse(await suggestNeedTags(
        (payload.payload && typeof payload.payload === "object") ? payload.payload as Record<string, unknown> : {},
      ));
    }

    if (action === "userLogin") {
      return jsonResponse(await userLogin(requireString(payload.identifier), requireString(payload.password)));
    }

    if (action === "bridgeGrameeeSession") {
      return jsonResponse(await bridgeGrameeeSession(requireString(payload.grameeeAccessToken)));
    }

    if (action === "registerUser") {
      return jsonResponse(await registerUser(payload));
    }

    if (action === "requestPasswordReset") {
      return jsonResponse(await requestPasswordReset(requireString(payload.email)));
    }

      if (action === "submitSharedForm") {
        return jsonResponse(await submitFormSubmission(
          requireString(payload.submissionType),
          (payload.payload && typeof payload.payload === "object") ? payload.payload as Record<string, unknown> : {},
          requireString(payload.sourceMode) || "shared_link",
          null,
        ));
      }

    if (action === "resetPassword") {
      return jsonResponse(await resetPassword(requireString(payload.email), requireString(payload.code), requireString(payload.newPassword)));
    }

    if (action === "adminLogin") {
      return jsonResponse(await adminLogin(requireString(payload.username), requireString(payload.password)));
    }

    if (action === "validateUserSession") {
      const userCtx = await requireUserSession(req, requireString(payload.userSessionToken));
      return jsonResponse({
        ok: true,
        user: {
          id: userCtx.user.id,
          username: userCtx.user.username,
          email: userCtx.user.email,
          firstName: userCtx.user.first_name,
          fullName: userCtx.user.full_name,
          phone: userCtx.user.phone,
          role: userCtx.user.role,
          mustChangePassword: userCtx.user.must_change_password,
        },
      });
    }

    if (action === "validateAdminSession") {
      const session = await requireAdminSession(req, requireString(payload.adminSessionToken));
      return jsonResponse({ ok: true, username: session.user.username, email: session.admin.email });
    }

    if (action === "userLogout") {
      return jsonResponse(await userLogout(
        requireString(payload.userSessionToken) ||
        requireString(payload.adminSessionToken) ||
        requireString(req.headers.get("x-gre-user-session")) ||
        requireString(req.headers.get("x-gre-admin-session")),
      ));
    }

    if (action === "adminLogout") {
      return jsonResponse(await adminLogout(requireString(payload.adminSessionToken) || requireString(req.headers.get("x-gre-admin-session"))));
    }

    if (action === "submitUpdateRequest") {
      return jsonResponse(await submitUpdateRequest(payload));
    }

    if (action === "sendNeedSeekerIntro") {
      return jsonResponse(await sendNeedSeekerIntro(
        requireString(payload.needId),
        requireString(payload.grameeeAccessToken),
        payload.grameeeUserSummary,
      ));
    }

    if (action === "previewNeedSeekerIntro") {
      return jsonResponse(await buildNeedSeekerIntroDraft(
        requireString(payload.needId),
        requireString(payload.grameeeAccessToken),
        payload.grameeeUserSummary,
      ));
    }

    if (action === "__debugNeedSyncService") {
      return jsonResponse(await debugGreNeedSyncService(requireString(payload.needId)));
    }

    const userCtx = await requireUserSessionFromRequest(
      req,
      payload,
      requireString(payload.userSessionToken) || requireString(payload.adminSessionToken),
    );
    const actorEmail = requireString(userCtx.user.email).toLowerCase();
    const actorName = requireString(userCtx.user.full_name) || requireString(userCtx.user.first_name) || actorEmail;

    if (action === "changePassword") {
      return jsonResponse(await changePassword(userCtx, requireString(payload.currentPassword), requireString(payload.newPassword)));
    }

    if (action === "sendCuratorMessage") {
      return jsonResponse(await sendCuratorMessage(
        requireString(payload.needId),
        actorName,
        actorEmail,
        requireString(payload.message),
      ));
    }

    if (action === "submitSignedInForm") {
      return jsonResponse(await submitFormSubmission(
        requireString(payload.submissionType),
        (payload.payload && typeof payload.payload === "object") ? payload.payload as Record<string, unknown> : {},
        "signed_in",
        userCtx,
      ));
    }

    if (action === "directCuratorUpdate") {
      assertRoles(userCtx, ["curator", "moderator", "admin"], "Curator, moderator, or admin access is required.");
      return jsonResponse(await directCuratorUpdate(payload, userCtx));
    }

    if (action === "sendProviderIntro") {
      assertRoles(userCtx, ["curator", "moderator", "admin"], "Curator, moderator, or admin access is required.");
      return jsonResponse(await sendProviderIntro(
        requireString(payload.needId),
        requireString(payload.providerEmail),
        {
          id: requireString(userCtx.user.id),
          name: actorName,
          email: actorEmail,
          role: requireString(userCtx.user.role),
        },
        requireString(payload.providerName),
        requireString(payload.offeringId),
        (payload.mailDraft && typeof payload.mailDraft === "object") ? payload.mailDraft as Record<string, unknown> : null,
      ));
    }

    if (action === "sendSolutionSeekerIntro") {
      assertRoles(userCtx, ["curator", "moderator", "admin"], "Curator, moderator, or admin access is required.");
      return jsonResponse(await sendSolutionSeekerIntro(
        requireString(payload.needId),
        requireString(payload.providerEmail),
        {
          id: requireString(userCtx.user.id),
          name: actorName,
          email: actorEmail,
          role: requireString(userCtx.user.role),
        },
        requireString(payload.providerName),
        requireString(payload.offeringId),
        (payload.mailDraft && typeof payload.mailDraft === "object") ? payload.mailDraft as Record<string, unknown> : null,
      ));
    }

    if (action === "rejectNeedSuggestedQuestions") {
      assertRoles(userCtx, ["curator", "moderator", "admin"], "Curator, moderator, or admin access is required.");
      return jsonResponse(await rejectSuggestedQuestionsForNeed(requireString(payload.needId), actorEmail));
    }

    if (action === "downloadSeekerRequestTracker") {
      assertRoles(userCtx, ["curator", "moderator", "admin"], "Curator, moderator, or admin access is required.");
      return jsonResponse(await downloadSeekerRequestTracker(requireString(payload.seekerKey), Boolean(payload.includeClosed)));
    }

    if (action === "getMailTemplates") {
      assertRoles(userCtx, ["moderator", "admin"], "Moderator or admin access is required.");
      return jsonResponse(await getGreMailTemplates());
    }

    if (action === "saveMailTemplates") {
      assertRoles(userCtx, ["admin"], "Admin login required.");
      return jsonResponse(await saveGreMailTemplates(
        actorEmail,
        requireString(payload.providerIntroTemplate),
        requireString(payload.curatorForwardTemplate),
        requireString(payload.solutionSeekerTemplate),
        requireString(payload.needSeekerTemplate),
        typeof payload.inboundAutoSyncEnabled === "boolean" ? payload.inboundAutoSyncEnabled : null,
        payload.lshContactEmails,
        payload.lshHelpCcEmails,
        requireString(payload.lshRequestSupportTemplate),
        requireString(payload.lshEmailProviderTemplate),
      ));
    }

    if (action === "updateUserRole") {
      assertRoles(userCtx, ["admin"], "Admin login required.");
      return jsonResponse(await updateUserRole(requireString(payload.userId), requireString(payload.role)));
    }

    if (action === "completeUserRoleActivation") {
      assertRoles(userCtx, ["admin"], "Admin login required.");
      return jsonResponse(await completeUserRoleActivation(requireString(payload.userId), requireString(payload.otp)));
    }

    if (action === "refreshUserDirectory") {
      assertRoles(userCtx, ["admin"], "Admin login required.");
      return jsonResponse(await refreshUserDirectory());
    }

    if (action === "removeManagedUser") {
      assertRoles(userCtx, ["admin"], "Admin login required.");
      return jsonResponse(await removeManagedUser(requireString(payload.userId), requireString(payload.removalMode) || "org_only"));
    }

    const adminCtx = await requireAdminSessionFromRequest(req, payload, requireString(payload.adminSessionToken));
    const adminActorEmail = adminCtx.admin.email;
    const adminActorRole = requireString(adminCtx.user.role).toLowerCase();

    if (action === "adminSnapshot") {
      return jsonResponse(await getAdminSnapshot());
    }

    if (action === "adminUsersSnapshot") {
      return jsonResponse(await getAdminUsersSnapshot());
    }

    if (action === "adminMailImpactSnapshot") {
      return jsonResponse(await getAdminMailImpactSnapshot());
    }

    if (action === "adminDataSyncSnapshot") {
      return jsonResponse(await getAdminDataSyncSnapshot());
    }

    if (action === "adminApprovalsSnapshot") {
      return jsonResponse(await getAdminApprovalsSnapshot());
    }

    if (action === "adminLocalSolutionsSnapshot") {
      return jsonResponse(await getAdminLocalSolutionsSnapshot());
    }

    if (action === "adminLocalNeedsSnapshot") {
      return jsonResponse(await getAdminLocalNeedsSnapshot());
    }

    if (action === "assignCurator") {
      return jsonResponse(await assignCurator(requireString(payload.needId), requireString(payload.curatorId) || null, adminActorEmail));
    }

    if (action === "approveNeed") {
      if (isModeratorMisRole(adminActorRole) && requireString(payload.decision).toLowerCase() === "reject") {
        throw new Error("Moderators cannot reject needs.");
      }
      return jsonResponse(await approveNeed(requireString(payload.needId), requireString(payload.decision), requireString(payload.reviewNotes), adminActorEmail));
    }

    if (action === "reviewUpdateRequest") {
      if (isModeratorMisRole(adminActorRole) && requireString(payload.decision).toLowerCase() === "reject") {
        throw new Error("Moderators cannot reject curator updates.");
      }
      return jsonResponse(await reviewUpdateRequest(requireString(payload.requestId), requireString(payload.decision), requireString(payload.reviewNotes), adminActorEmail));
    }

    if (action === "debugGreNeedSync") {
      return jsonResponse(await debugGreNeedSync(requireString(payload.needId)));
    }

    if (action === "reviewFormSubmission") {
      if (isModeratorMisRole(adminActorRole) && requireString(payload.decision).toLowerCase() === "reject") {
        throw new Error("Moderators cannot reject form submissions.");
      }
      return jsonResponse(await approveFormSubmission(
        requireString(payload.submissionId),
        requireString(payload.decision),
        requireString(payload.reviewNotes),
        adminActorEmail,
      ));
    }

    if (action === "updateFormSubmission") {
      return jsonResponse(await updateFormSubmission(
        requireString(payload.submissionId),
        (payload.update && typeof payload.update === "object") ? payload.update as Record<string, unknown> : {},
        adminActorEmail,
      ));
    }

    if (action === "updateLocalSolution") {
      return jsonResponse(await updateLocalSolution(
        requireString(payload.offeringId),
        (payload.payload && typeof payload.payload === "object") ? payload.payload as Record<string, unknown> : {},
        adminActorEmail,
      ));
    }

    if (action === "getLocalSolutionDetail") {
      return jsonResponse(await getLocalSolutionDetail(requireString(payload.offeringId)));
    }

    if (action === "deleteLocalSolution") {
      if (isModeratorMisRole(adminActorRole)) {
        throw new Error("Moderators cannot delete local solutions.");
      }
      return jsonResponse(await deleteLocalSolution(requireString(payload.offeringId)));
    }

    if (action === "updateLocalNeed") {
      return jsonResponse(await updateLocalNeed(
        requireString(payload.needId),
        (payload.payload && typeof payload.payload === "object") ? payload.payload as Record<string, unknown> : {},
      ));
    }

    if (action === "getLocalNeedDetail") {
      return jsonResponse(await getLocalNeedDetail(requireString(payload.needId)));
    }

    if (action === "deleteLocalNeed") {
      if (isModeratorMisRole(adminActorRole)) {
        throw new Error("Moderators cannot delete local needs.");
      }
      return jsonResponse(await deleteLocalNeed(requireString(payload.needId)));
    }

    if (action === "upsertOption") {
      return jsonResponse(await upsertOption(requireString(payload.optionType), requireString(payload.label)));
    }

    if (action === "importInboundWorkbook") {
      return jsonResponse(
        await importInboundWorkbook(payload.rows, requireString(payload.fileName) || "GRE inbound workbook", adminActorEmail, requireString(payload.aiProvider) || defaultAiProvider),
      );
    }

    if (action === "syncGreLiveInbounds") {
      return jsonResponse(await syncGreLiveInbounds(adminActorEmail, requireString(payload.aiProvider) || defaultAiProvider));
    }

    if (action === "syncGreChatbotData") {
      return jsonResponse(await syncGreChatbotData(requireString(payload.aiProvider) || defaultAiProvider));
    }

    if (action === "downloadGreChatbotReport") {
      return jsonResponse(await downloadGreChatbotReport(requireString(payload.reportKind)));
    }

    if (action === "downloadGreInboundReport") {
      return jsonResponse(await downloadGreInboundReport());
    }

    if (action === "uploadChatbotWorkbooks") {
      return jsonResponse(await uploadChatbotWorkbooks(
        requireString(payload.solutionBase64) || "",
        requireString(payload.traderBase64) || "",
        requireString(payload.solutionFileName) || "uploaded_solutions.xlsx",
        requireString(payload.traderFileName) || "uploaded_traders.xlsx",
        requireString(payload.aiProvider) || defaultAiProvider,
      ));
    }

    if (action === "uploadChatbotNormalized") {
      return jsonResponse(await uploadChatbotNormalized(
        Array.isArray(payload.traders) ? payload.traders as Record<string, unknown>[] : [],
        Array.isArray(payload.solutions) ? payload.solutions as Record<string, unknown>[] : [],
        Array.isArray(payload.offerings) ? payload.offerings as Record<string, unknown>[] : [],
        requireString(payload.solutionFileName) || "normalized_solutions.json",
        requireString(payload.traderFileName) || "normalized_traders.json",
        requireString(payload.aiProvider) || defaultAiProvider,
      ));
    }

      if (action === "refreshNeedIntelligence") {
        return jsonResponse(await refreshNeedIntelligence(adminActorEmail, requireString(payload.aiProvider) || defaultAiProvider));
      }

      if (action === "refreshChatbotIntelligence") {
        return jsonResponse(await refreshChatbotIntelligence(requireString(payload.aiProvider) || defaultAiProvider));
      }

      if (action === "generateNeedSuggestedQuestions") {
        return jsonResponse(await generateSuggestedQuestionsForNeed(requireString(payload.needId), adminActorEmail));
      }

      if (action === "rejectNeedSuggestedQuestions") {
        return jsonResponse(await rejectSuggestedQuestionsForNeed(requireString(payload.needId), adminActorEmail));
      }

      if (action === "saveNeedSuggestedQuestions") {
        return jsonResponse(await saveSuggestedQuestionsForNeed(
          requireString(payload.needId),
          payload.questions,
          adminActorEmail,
          requireString(payload.sourceLabel) || "manual",
        ));
      }

      if (action === "applyNeedOverride") {
        return jsonResponse(await applyNeedOverride(
          requireString(payload.needId),
          (payload.patch && typeof payload.patch === "object") ? payload.patch as Record<string, unknown> : {},
          requireString(payload.conflictNote),
          Boolean(payload.resolveConflict),
          adminActorEmail,
        ));
      }

      return jsonResponse({ error: "Unsupported action." }, 400);
  } catch (error) {
    const message = error instanceof Error ? error.message : "Function failed.";
    return jsonResponse({ error: message }, 400);
  }
});
