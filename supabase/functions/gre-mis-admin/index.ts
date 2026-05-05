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
const greSiteOrigin = Deno.env.get("GRE_SITE_ORIGIN") ?? "https://greenruraleconomy.in";
const greMasterLogin = Deno.env.get("GRE_MASTER_LOGIN") ?? "";
const greMasterTenant = Deno.env.get("GRE_MASTER_TENANT") ?? "green_rural_economy";
const greMasterPassword = Deno.env.get("GRE_MASTER_PASSWORD") ?? "";
const greMarketId = Deno.env.get("GRE_MARKET_ID") ?? "5";
const greInboundReportName = Deno.env.get("GRE_INBOUND_REPORT_NAME") ?? "MARKIFY_GRE.REQUEST_DETAILS_REPORT";
const greTraderReportName = Deno.env.get("GRE_TRADER_REPORT_NAME") ?? "MARKIFY_REPORT.MARKET_TRADER_BASIC_DETAILS";
const greSolutionReportName =
  Deno.env.get("GRE_SOLUTION_REPORT_NAME") ?? "MARKIFY_REPORT.GRE_SOLUTION_WITH_OFFERINGS_DETAILS";

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

function asStringArray(value: unknown) {
  if (Array.isArray(value)) return value.map((item) => requireString(item)).filter(Boolean);
  if (typeof value === "string") {
    return value.split(/[;,|]/).map((item) => item.trim()).filter(Boolean);
  }
  return [];
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

function tokenizeLooseText(value: unknown, minimumLength = 4) {
  return uniqueStrings(
    requireString(value)
      .toLowerCase()
      .split(/[^a-z0-9]+/)
      .map((entry) => entry.trim())
      .filter((entry) => entry.length >= minimumLength),
  );
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

async function createWorkbookDownloadPayload(buffer: ArrayBuffer, fileName: string) {
  return {
    fileName,
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    base64: bytesToBase64(new Uint8Array(buffer)),
  };
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

async function fetchGreJson(path: string) {
  const sessionId = await loginToGre();
  const response = await fetch(`${greLoginBaseUrl}${path}`, {
    headers: {
      "x-sessionid": sessionId,
      Accept: "application/json, text/plain, */*",
      Origin: greSiteOrigin,
      Referer: `${greSiteOrigin}/`,
    },
  });
  const data = await response.json().catch(() => null);
  if (!response.ok) {
    throw new Error(data?.error?.message || data?.message || `GRE request failed for ${path}.`);
  }
  return { sessionId, data };
}

async function patchGreJson(path: string, payload: unknown, sessionId?: string) {
  const resolvedSessionId = sessionId || await loginToGre();
  const response = await fetch(`${greLoginBaseUrl}${path}`, {
    method: "PATCH",
    headers: {
      "x-sessionid": resolvedSessionId,
      "Content-Type": "application/json",
      Accept: "application/json, text/plain, */*",
      Origin: greSiteOrigin,
      Referer: `${greSiteOrigin}/`,
    },
    body: JSON.stringify(payload),
  });
  const data = await response.json().catch(() => null);
  if (!response.ok) {
    throw new Error(data?.error?.message || data?.message || `GRE patch failed for ${path}.`);
  }
  return data;
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

async function syncApprovedUpdateToGre(requestRow: Record<string, any>) {
  const needId = requireString(requestRow.need_id);
  if (!needId) throw new Error("Need id is required for GRE sync.");

  const { sessionId, data: greRequest } = await fetchGreJson(`/commons-request-management-service/api/v1/request?id=${encodeURIComponent(needId)}`);
  const requestStatusOptions = await getGreRefData("CLASS.REQUEST_STATUS");
  const internalStatusOptions = await getGreRefData("CLASS.REQUEST_INTERNAL_STATUS");
  const demandBroadcastOptions = await getGreRefData("CLASS.DEMAND_BROADCAST_NEEDED");

  const payload = structuredClone(greRequest);
  const curationList = Array.isArray(payload.curationList) ? payload.curationList : [];
  const primaryCuration = curationList[0] || {
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

  if (typeof requestRow.proposed_demand_broadcast_needed === "boolean") {
    const mapped = findGreRefDataOption(
      demandBroadcastOptions,
      requestRow.proposed_demand_broadcast_needed ? "Yes" : "No",
    );
    if (!mapped) throw new Error("Could not map GRE demand broadcast option.");
    primaryCuration.demandBroadcastNeed = mapped;
  }

  if (requestRow.proposed_next_action) unsupportedFields.push("next_action");
  if (Number.isInteger(requestRow.proposed_solutions_shared_count)) unsupportedFields.push("solutions_shared_count");
  if (Number.isInteger(requestRow.proposed_invited_providers_count)) unsupportedFields.push("invited_providers_count");

  payload.curationList = [primaryCuration, ...curationList.slice(1)];
  await patchGreJson("/commons-request-management-service/api/v1/request", payload, sessionId);
  return { ok: true, unsupportedFields };
}

function stableRowSignature(value: unknown) {
  return JSON.stringify(value, Object.keys(value as Record<string, unknown>).sort());
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
  const domainTokens = tokenizeLooseText(text, 4).filter((token) => !domainStopwords.has(token)).slice(0, 18);

  return {
    thematicHints,
    serviceHints,
    sixMSignals: rule6M,
    needKind,
    serviceKind,
    keywords: uniqueStrings([...thematicHints, ...domainTokens]).slice(0, 18),
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

async function callAiJson(providerInput: string, prompt: string) {
  const provider = (providerInput || defaultAiProvider || "openrouter").toLowerCase();
  const configuredProviders = [
    geminiApiKey ? "gemini" : "",
    openRouterApiKey ? "openrouter" : "",
    deepSeekApiKey ? "deepseek" : "",
    openAiApiKey ? "openai" : "",
  ].filter(Boolean);
  const tryProvider = async (resolvedProvider: string) => {
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
  };

  const requestedOrder = provider === "gemini"
    ? ["gemini", "openrouter", "deepseek", "openai"]
    : provider === "deepseek"
      ? ["deepseek", "openrouter", "gemini", "openai"]
      : provider === "openai"
        ? ["openai", "openrouter", "gemini", "deepseek"]
        : ["openrouter", "gemini", "deepseek", "openai"];
  const fallbackOrder = [
    ...requestedOrder.filter((candidate) => configuredProviders.includes(candidate)),
    ...configuredProviders.filter((candidate) => !requestedOrder.includes(candidate)),
  ];

  if (!fallbackOrder.length) {
    throw new Error("No AI provider is configured. Falling back to rule-based enrichment.");
  }

  let lastError: unknown = null;
  for (const candidate of fallbackOrder) {
    try {
      return await tryProvider(candidate);
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError instanceof Error ? lastError : new Error("No AI provider could enrich the need.");
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

async function getGmailAccessToken() {
  const response = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: gmailClientId,
      client_secret: gmailClientSecret,
      refresh_token: gmailRefreshToken,
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
  return btoa(input).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

async function sendEmail({
  to,
  cc,
  subject,
  body,
}: {
  to: string;
  cc?: string;
  subject: string;
  body: string;
}) {
  const accessToken = await getGmailAccessToken();
  const rawMessage = [
    `From: ${gmailSenderEmail}`,
    `To: ${to}`,
    cc ? `Cc: ${cc}` : "",
    `Subject: ${subject}`,
    "Content-Type: text/plain; charset=UTF-8",
    "",
    body,
  ]
    .filter(Boolean)
    .join("\r\n");

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
    assertRoles(userCtx, ["admin"], "Admin login required.");
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
        first_name: "Admin",
        full_name: "GRE Admin",
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

async function getAdminSnapshot() {
  const [pendingNeeds, pendingUpdates, aiReviewNeeds, users, pendingFormSubmissions] = await Promise.all([
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
        .select("id, organization_name, state, district, problem_statement, curation_notes, curated_need, ai_thematic_area, ai_application_area, ai_need_kind, ai_service_kind, ai_keywords, ai_6m_signals, ai_validation_status, ai_validation_flags, ai_confidence, ai_enrichment_status, ai_engine, ai_enriched_at, rule_thematic_hints, rule_6m_signals, override_thematic_area, override_application_area, override_need_kind, override_service_kind, override_keywords, override_6m_signals, override_summary, override_source, override_conflict_note, override_updated_at")
        .eq("approval_status", "approved")
        .or("ai_validation_status.is.null,ai_validation_status.eq.flagged,ai_enrichment_status.is.null")
        .order("updated_at", { ascending: false })
        .limit(60),
      adminClient
        .from("gre_mis_users")
        .select("id, username, first_name, full_name, email, phone, role, is_active, must_change_password, last_login_at, created_at")
        .order("role", { ascending: true })
        .order("first_name", { ascending: true }),
      adminClient
        .from("gre_mis_form_submissions")
        .select("*")
        .eq("approval_status", "pending_admin")
        .order("created_at", { ascending: false }),
    ]);

  if (pendingNeeds.error) throw new Error(pendingNeeds.error.message);
  if (pendingUpdates.error) throw new Error(pendingUpdates.error.message);
  if (aiReviewNeeds.error) throw new Error(aiReviewNeeds.error.message);
  if (users.error) throw new Error(users.error.message);
  if (pendingFormSubmissions.error) throw new Error(pendingFormSubmissions.error.message);

  return {
    ok: true,
    pendingNeeds: pendingNeeds.data || [],
    pendingUpdates: pendingUpdates.data || [],
    aiReviewNeeds: aiReviewNeeds.data || [],
    users: users.data || [],
    pendingFormSubmissions: pendingFormSubmissions.data || [],
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

async function promoteUserToCurator(userId: string) {
  const { data: user, error } = await adminClient
    .from("gre_mis_users")
    .select("id, first_name, full_name, email, phone")
    .eq("id", userId)
    .single();
  if (error || !user) throw new Error(error?.message || "User not found.");

  const { error: userUpdateError } = await adminClient
    .from("gre_mis_users")
    .update({ role: "curator", updated_at: new Date().toISOString() })
    .eq("id", userId);
  if (userUpdateError) throw new Error(userUpdateError.message);

  const { error: curatorError } = await adminClient.from("gre_mis_curators").upsert({
    user_id: user.id,
    display_name: user.full_name,
    first_name: user.first_name,
    email: user.email,
    phone: user.phone,
    is_active: true,
    gre_sync_status: "pending",
    gre_sync_message: "Awaiting GRE curator-directory sync mapping.",
    gre_synced_at: null,
  }, { onConflict: "email" });
  if (curatorError) throw new Error(curatorError.message);

  return { ok: true };
}

async function submitFormSubmission(
  submissionType: string,
  payload: Record<string, unknown>,
  sourceMode = "shared_link",
  userCtx?: { user?: Record<string, unknown> | null } | null,
) {
  const normalizedType = requireString(submissionType).toLowerCase();
  if (!["need", "solution"].includes(normalizedType)) throw new Error("Unsupported submission type.");
  const organizationName = requireString(payload.organization_name || payload.organizationName);
  const existingTraderId = requireString(payload.existing_trader_id || payload.existingTraderId);
  const existingTraderName = requireString(payload.existing_trader_name || payload.existingTraderName);
  const submitterName = requireString(payload.submitter_name || payload.submitterName || userCtx?.user?.full_name || userCtx?.user?.first_name);
  const submitterEmail = requireString(payload.submitter_email || payload.submitterEmail || userCtx?.user?.email).toLowerCase();
  const submitterPhone = requireString(payload.submitter_phone || payload.submitterPhone || userCtx?.user?.phone);
  const shareContext = requireString(payload.share_context || payload.shareContext);

  if (!organizationName) throw new Error("Organization name is required.");
  if (!existingTraderId) throw new Error("Please select an existing GRE supplier organization first.");

  const { data: trader, error: traderError } = await adminClient
    .from("traders")
    .select("trader_id, organisation_name, trader_name")
    .eq("trader_id", existingTraderId)
    .maybeSingle();
  if (traderError) throw new Error(traderError.message);
  if (!trader) throw new Error("Selected GRE supplier organization could not be found.");

  const submissionRow = {
    submission_type: normalizedType,
    source_mode: sourceMode === "signed_in" ? "signed_in" : "shared_link",
    submitter_name: submitterName || null,
    submitter_email: submitterEmail || null,
    submitter_phone: submitterPhone || null,
    submitter_user_id: requireString(userCtx?.user?.id) || null,
    organization_name: organizationName,
    existing_trader_id: trader.trader_id,
    existing_trader_name: trader.organisation_name || trader.trader_name || existingTraderName || organizationName,
    org_exists_on_gre: true,
    payload,
    share_context: shareContext || null,
  };

  const { data, error } = await adminClient
    .from("gre_mis_form_submissions")
    .insert(submissionRow)
    .select("id")
    .single();
  if (error) throw new Error(error.message);

  return {
    ok: true,
    id: data.id,
    message: "Submission saved for admin review.",
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
    const curatedNeed = asStringArray(payload.curated_need || payload.categories);
    const row = {
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
      curated_need: curatedNeed,
      approval_status: "approved",
      source_kind: "shared_form_submission",
      next_action: "Allocate curator and sync to GRE",
      requested_on: new Date().toISOString(),
      last_status_change_at: new Date().toISOString(),
    };
    const { error: insertError } = await adminClient.from("gre_mis_needs").insert(row);
    if (insertError) throw new Error(insertError.message);

    const { error: approveError } = await adminClient
      .from("gre_mis_form_submissions")
      .update({
        approval_status: "approved",
        admin_review_notes: reviewNotes || "Approved by admin.",
        reviewed_by_email: actorEmail,
        reviewed_at: new Date().toISOString(),
        target_need_id: nextNeedId,
        gre_sync_status: "pending_gre_create_verification",
        gre_sync_message: "Local need created. GRE request create API still needs explicit create-flow verification.",
      })
      .eq("id", submissionId);
    if (approveError) throw new Error(approveError.message);
    return { ok: true, targetNeedId: nextNeedId };
  }

  const { error: approveError } = await adminClient
    .from("gre_mis_form_submissions")
    .update({
      approval_status: "approved",
      admin_review_notes: reviewNotes || "Approved by admin.",
      reviewed_by_email: actorEmail,
      reviewed_at: new Date().toISOString(),
      gre_sync_status: "pending_gre_solution_create_verification",
      gre_sync_message: "Queued for GRE solution create after API flow verification.",
    })
    .eq("id", submissionId);
  if (approveError) throw new Error(approveError.message);
  return { ok: true };
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
    solution_image_url: normalizeCell(row.SolutionImage),
    raw_payload: row,
  };
}

function normalizeOfferingRowForChatbot(row: Record<string, unknown>) {
  const aboutOfferingHtml = normalizeCell(row.AboutOffering);
  const trainerDetailsHtml = normalizeCell(row["Trainer Details"]);
  const valuechains = splitLooseList(normalizeCell(row.Valuechains));
  const applications = splitLooseList(normalizeCell(row.Applications));
  const tags = splitLooseList(normalizeCell(row.Tags));
  const languages = splitLooseList(normalizeCell(row.Languages));
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
    service_brochure_url: normalizeCell(row["Service offering Brochure"]),
    grade_capacity: normalizeCell(row["Grade/Capacity"]),
    product_cost: normalizeCell(row["Cost (Product)"]),
    lead_time: normalizeCell(row["Lead Time"]),
    support_details: normalizeCell(row.Support),
    product_brochure_url: normalizeCell(row["Product Brochure"]),
    knowledge_content_url: normalizeCell(row["Knowledge Offering Content"]),
    contact_details: normalizeCell(row["Contact Details"]),
    gre_link: normalizeCell(row["Offering Link on GRE"]),
    source_row_signature: stableRowSignature(row),
    search_document: buildSearchDocument([
      normalizeCell(row.SolutionName),
      normalizeCell(row.OfferingName),
      normalizeCell(row.OfferingCategory),
      normalizeCell(row.OfferingGroup),
      normalizeCell(row.OfferingType),
      normalizeCell(row["6M"]),
      normalizeCell(row.PrimaryValuechain),
      normalizeCell(row.PrimaryApplication),
      valuechains,
      applications,
      tags,
      languages,
      geographies,
      stripHtml(normalizeCell(row.AboutSolution)),
      stripHtml(aboutOfferingHtml),
      normalizeCell(row.TraderOrganisation),
    ]),
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
      .select("offering_id, source_row_signature, ai_enriched_at")
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
      const { error } = await adminClient.from("solutions").upsert(rows, { onConflict: "solution_id" });
      if (error) throw new Error(error.message);
    }

    for (const rows of chunkArray(bundle.offerings)) {
      const rowsWithImport = rows.map((row) => ({ ...row, last_import_id: importId }));
      const { error } = await adminClient.from("offerings").upsert(rowsWithImport, { onConflict: "offering_id" });
      if (error) throw new Error(error.message);
    }

    if (changedOfferingIds.length) {
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
    const curatorId = await resolveCuratorId(requireString(row.curator_name));
    const patch = await normalizeInboundRow(row, curatorId, fileName);
    const sourceRowSignature = stableRowSignature(row);
    const existing = existingMap.get(requestId);

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
  const report = await fetchGreReportWorkbook(greInboundReportName, 5000, "request_details_report");
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

async function assignCurator(needId: string, curatorId: string | null, actorEmail: string) {
  const { data: needRow, error: needFetchError } = await adminClient
    .from("gre_mis_needs")
    .select("id, source_kind")
    .eq("id", needId)
    .single();
  if (needFetchError || !needRow) throw new Error(needFetchError?.message || "Need not found.");

  const { error } = await adminClient
    .from("gre_mis_needs")
    .update({
      curator_id: curatorId,
      status: "Accepted",
      updated_at: new Date().toISOString(),
    })
    .eq("id", needId);
  if (error) throw new Error(error.message);

  await adminClient.from("gre_mis_need_updates").insert({
    need_id: needId,
    update_type: "curator_assignment",
    note: curatorId ? `Curator assigned by ${actorEmail}.` : `Curator unassigned by ${actorEmail}.`,
    created_by_email: actorEmail,
  });

  if (/^\d+$/.test(requireString(needRow.id)) && needRow.source_kind !== "shared_form_submission") {
    const { sessionId, data: greRequest } = await fetchGreJson(`/commons-request-management-service/api/v1/request?id=${encodeURIComponent(needId)}`);
    const payload = structuredClone(greRequest);
    const curationList = Array.isArray(payload.curationList) ? payload.curationList : [];
    const primaryCuration = curationList[0] || { requestId: Number(needId) };
    if (curatorId) {
      const { data: curatorRow } = await adminClient
        .from("gre_mis_curators")
        .select("display_name")
        .eq("id", curatorId)
        .maybeSingle();
      primaryCuration.curatorUserName = curatorRow?.display_name || primaryCuration.curatorUserName || null;
    } else {
      primaryCuration.curatorUserName = null;
      primaryCuration.curatorUserId = null;
    }
    payload.curationList = [primaryCuration, ...curationList.slice(1)];
    try {
      await patchGreJson("/commons-request-management-service/api/v1/request", payload, sessionId);
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

  const { data: need, error: needError } = await adminClient
    .from("gre_mis_needs")
    .select("id, approval_status, curator_id, gre_mis_curators:curator_id(user_id, email)")
    .eq("id", needId)
    .single();
  if (needError || !need) throw new Error(needError?.message || "Need not found.");
  if (need.approval_status !== "approved") throw new Error("Only approved needs can be updated.");

  const curatorRow = Array.isArray(need.gre_mis_curators) ? need.gre_mis_curators[0] : need.gre_mis_curators;
  const assignedUserId = requireString(curatorRow?.user_id);
  const assignedEmail = requireString(curatorRow?.email).toLowerCase();
  if (assignedUserId !== requireString(userCtx.user.id) && assignedEmail !== actorEmail) {
    throw new Error("You can edit curation only for needs assigned to you.");
  }

  const requestRow = {
    need_id: needId,
    proposed_status: requireString(payload.proposedStatus) || null,
    proposed_internal_status: requireString(payload.proposedInternalStatus) || null,
    proposed_next_action: requireString(payload.proposedNextAction) || null,
    proposed_curation_notes: requireString(payload.proposedCurationNotes) || null,
    proposed_curation_call_date: requireString(payload.proposedCurationCallDate) || null,
    proposed_demand_broadcast_needed: normalizeBooleanInput(payload.proposedDemandBroadcastNeeded),
    proposed_solutions_shared_count: null,
    proposed_invited_providers_count: null,
  };

  const greSyncResult = await applyApprovedNeedPatch(requestRow, actorEmail);
  return {
    ok: true,
    greSync: {
      synced: true,
      unsupportedFields: greSyncResult.unsupportedFields || [],
    },
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

async function sendProviderIntro(needId: string, providerEmail: string, actorEmail: string) {
  const { data: need, error: needError } = await adminClient
    .from("gre_mis_needs")
    .select("id, organization_name, seeker_email, contact_person, problem_statement, state, district, curated_need, gre_mis_curators:curator_id(email, display_name)")
    .eq("id", needId)
    .single();
  if (needError || !need) throw new Error(needError?.message || "Need not found.");
  const curator = Array.isArray(need.gre_mis_curators) ? need.gre_mis_curators[0] : need.gre_mis_curators;
  const curatorEmail = requireString(curator?.email).toLowerCase();

  const subject = `GRE introduction: ${need.organization_name} challenge for your consideration`;
  const body = [
    `Hello,`,
    "",
    "We are reaching out from Green Rural Economy regarding a challenge shared on the GRE platform.",
    "",
    `Organization: ${need.organization_name}`,
    `Contact person: ${need.contact_person || "Not provided"}`,
    `Location: ${need.state || "Not provided"}${need.district ? ` / ${need.district}` : ""}`,
    `Need categories: ${(need.curated_need || []).join(", ") || "Not yet classified"}`,
    "",
    "Problem statement:",
    `${need.problem_statement}`,
    "",
    "If this is relevant to your work, please reply to this email so we can coordinate next steps with the seeker.",
    "",
    "Regards,",
    "GRE Team",
  ].join("\n");

  await sendEmail({
    to: providerEmail,
    cc: [need.seeker_email, curatorEmail].filter(Boolean).join(", "),
    subject,
    body,
  });

  await adminClient.from("gre_mis_email_log").insert({
    need_id: need.id,
    recipient_email: providerEmail,
    cc_email: [need.seeker_email, curatorEmail].filter(Boolean).join(", "),
    subject,
    body_preview: body.slice(0, 1000),
    sent_by_email: actorEmail,
  });

  await adminClient.from("gre_mis_need_updates").insert({
    need_id: need.id,
    update_type: "provider_intro_sent",
    note: `Problem statement emailed to ${providerEmail} with seeker copied.`,
    created_by_email: actorEmail,
  });

  return { ok: true, message: `Provider introduction email sent from ${gmailSenderEmail}.` };
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

  return { ok: true, message: `Message sent to ${curator?.display_name || "the curator"}.` };
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  try {
    const payload = (await req.json().catch(() => ({}))) as Record<string, unknown>;
    const action = requireString(payload.action);

    if (action === "userLogin") {
      return jsonResponse(await userLogin(requireString(payload.identifier), requireString(payload.password)));
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

    const userCtx = await requireUserSession(req, requireString(payload.userSessionToken) || requireString(payload.adminSessionToken));
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
      assertRoles(userCtx, ["curator", "admin"], "Curator or admin access is required.");
      return jsonResponse(await directCuratorUpdate(payload, userCtx));
    }

    if (action === "sendProviderIntro") {
      assertRoles(userCtx, ["curator", "admin"], "Curator or admin access is required.");
      return jsonResponse(await sendProviderIntro(requireString(payload.needId), requireString(payload.providerEmail), actorEmail));
    }

    if (action === "promoteUserToCurator") {
      assertRoles(userCtx, ["admin"], "Admin login required.");
      return jsonResponse(await promoteUserToCurator(requireString(payload.userId)));
    }

    const adminCtx = await requireAdminSession(req, requireString(payload.adminSessionToken));
    const adminActorEmail = adminCtx.admin.email;

    if (action === "adminSnapshot") {
      return jsonResponse(await getAdminSnapshot());
    }

    if (action === "assignCurator") {
      return jsonResponse(await assignCurator(requireString(payload.needId), requireString(payload.curatorId) || null, adminActorEmail));
    }

    if (action === "approveNeed") {
      return jsonResponse(await approveNeed(requireString(payload.needId), requireString(payload.decision), requireString(payload.reviewNotes), adminActorEmail));
    }

    if (action === "reviewUpdateRequest") {
      return jsonResponse(await reviewUpdateRequest(requireString(payload.requestId), requireString(payload.decision), requireString(payload.reviewNotes), adminActorEmail));
    }

    if (action === "reviewFormSubmission") {
      return jsonResponse(await approveFormSubmission(
        requireString(payload.submissionId),
        requireString(payload.decision),
        requireString(payload.reviewNotes),
        adminActorEmail,
      ));
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

      if (action === "refreshNeedIntelligence") {
        return jsonResponse(await refreshNeedIntelligence(adminActorEmail, requireString(payload.aiProvider) || defaultAiProvider));
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
