const FALLBACK_CURATORS = [
  { id: "fallback-1", display_name: "Tanmay Mukherji", email: "tanmay@greenruraleconomy.in" },
  { id: "fallback-2", display_name: "Phaneesh K", email: "phaneesh@greenruraleconomy.in" },
  { id: "fallback-3", display_name: "Swati Singh", email: "swati@greenruraleconomy.in" },
  { id: "fallback-4", display_name: "Shaifali Nagar", email: "shaifali@greenruraleconomy.in" },
];

const state = {
  view: "overview",
  selectedNeedId: null,
  pipelineFocus: "stuck",
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

function toast(message) {
  window.alert(message);
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

function formatDate(value) {
  if (!value) return "Not set";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? String(value) : date.toLocaleString("en-IN", { dateStyle: "medium", timeStyle: "short" });
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

function buildNeedMatchProfile(need) {
  const categories = uniq(parseArray(need.curated_need).map((item) => item.toLowerCase()));
  const categoryTokens = categories.flatMap((item) => tokenizeText(item, 3));
  const problemTokens = tokenizeText(need.problem_statement, 5);
  const notesTokens = tokenizeText(need.curation_notes, 5);
  const geographyTokens = tokenizeText(`${need.state || ""} ${need.district || ""}`, 3);
  const primaryTerms = uniq([...categoryTokens, ...problemTokens.slice(0, 8), ...notesTokens.slice(0, 4)]);
  const phrases = uniq(
    categories
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

  return {
    categories,
    categoryTokens: uniq(categoryTokens),
    problemTokens,
    notesTokens,
    geographyTokens,
    phrases,
    primaryTerms,
  };
}

function getPipelineSegmentNeeds(segmentId) {
  const segment = PIPELINE_SEGMENTS.find((item) => item.id === segmentId) || PIPELINE_SEGMENTS[0];
  return state.data.needs
    .filter(segment.match)
    .sort((a, b) => Number(b.curation_age_days || 0) - Number(a.curation_age_days || 0));
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

  profile.categories.forEach((phrase) => {
    if (joined.includes(phrase)) {
      score += 10;
      reasons.push(phrase);
    }
  });

  profile.categoryTokens.forEach((token) => {
    if (tags.some((tag) => tag.includes(token))) {
      score += 7;
      reasons.push(token);
    } else if (category.includes(token)) {
      score += 6;
      reasons.push(token);
    } else if (name.includes(token) || solutionName.includes(token)) {
      score += 5;
      reasons.push(token);
    } else if (about.includes(token)) {
      score += 3;
      reasons.push(token);
    }
  });

  profile.problemTokens.slice(0, 8).forEach((token) => {
    if (tags.some((tag) => tag.includes(token)) || name.includes(token) || solutionName.includes(token)) {
      score += 4;
      reasons.push(token);
    } else if (about.includes(token)) {
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

  if (!profile.categoryTokens.length && !profile.problemTokens.length) score += 1;
  if (!reasons.length) score -= 8;
  if (profile.categoryTokens.length && !profile.categoryTokens.some((token) => joined.includes(token))) score -= 4;

  return {
    score,
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

    state.data.curators = curators.data || [];
    state.data.options = options.data || [];
    state.data.needs = (needs.data || []).map((need) => ({
      ...need,
      curated_need: parseArray(need.curated_need),
    }));
    state.data.needUpdates = updates.data || [];
  }

  async callAdmin(action, body = {}, requireAdmin = false) {
    const client = this.getClient();
    const headers = {
      "Content-Type": "application/json",
      apikey: this.config.SUPABASE_ANON_KEY,
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
    const data = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(data.error || "Request failed.");
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
    state.data.pendingNeeds = data.pendingNeeds || [];
    state.data.pendingUpdates = data.pendingUpdates || [];
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

  async searchMatchesForNeed(need) {
    const client = this.getClient();
    if (!client || !need) return [];
    const profile = buildNeedMatchProfile(need);
    const offeringSelect = "offering_id,solution_id,trader_id,offering_name,offering_category,tags,geographies,about_offering_text,contact_details,gre_link";
    const searchTerms = uniq([...profile.categories, ...profile.phrases, ...profile.primaryTerms]).slice(0, 6);
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
          .limit(18),
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
          .limit(12),
      );
    });

    if (!queries.length) {
      queries.push(client.from("offerings").select(offeringSelect).limit(18));
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
          matchReasons: matchMeta.reasons,
        };
      })
      .filter((item) => item.matchScore >= 6)
      .sort((a, b) => b.matchScore - a.matchScore)
      .slice(0, 6);
  }
}

const store = new GreMisStore();

function getCuratorById(id) {
  return state.data.curators.find((curator) => curator.id === id) || null;
}

function getNeedById(id) {
  return state.data.needs.find((need) => need.id === id) || null;
}

function getVisibleNeeds() {
  return [...state.data.needs]
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
  const needs = state.data.needs;
  const metrics = [
    ["Approved Needs", needs.length, "Live inbound needs currently visible in the MIS."],
    ["In Progress", needs.filter((need) => need.status === "In progress").length, "Needs under active curation or provider search."],
    ["Need Providers", needs.filter((need) => need.internal_status === "Need solution providers").length, "Needs waiting on live solution/provider matching."],
    ["Connection Made", needs.filter((need) => need.internal_status === "Connection made").length, "Needs where the solution side has begun to move."],
    ["Stuck 7+ Days", needs.filter((need) => Number(need.curation_age_days || 0) >= 7).length, "Aging needs that need intervention during the day."],
    ["Admin Queue", state.data.pendingNeeds.length + state.data.pendingUpdates.length, "Pending intake approvals and curator change requests."],
  ];

  metricsGrid.innerHTML = metrics
    .map(
      ([label, value, note]) => `
        <article class="metric-card">
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
  const needs = state.data.needs;
  const headline = byId("datasetHeadline");
  const subline = byId("datasetSubline");
  if (headline) headline.textContent = `${needs.length} approved inbound needs loaded from GRE operations data`;
  if (subline) subline.textContent = `${state.data.pendingNeeds.length} intake records and ${state.data.pendingUpdates.length} curator updates are waiting for admin action.`;

  const stageData = PIPELINE_SEGMENTS.map((segment) => [
    segment.id,
    segment.label,
    needs.filter(segment.match).length,
    segment.note,
  ]);

  byId("pipelineBoard").innerHTML = stageData
    .map(
      ([id, label, value, note]) => `
        <article class="stage-card ${state.pipelineFocus === id ? "active" : ""}" data-pipeline-segment="${esc(id)}">
          <p class="eyebrow">${esc(label)}</p>
          <strong>${esc(value)}</strong>
          <p class="helper-text">${esc(note)}</p>
        </article>
      `,
    )
    .join("");

  const pipelineNeeds = getPipelineSegmentNeeds(state.pipelineFocus);
  const focusMeta = PIPELINE_SEGMENTS.find((segment) => segment.id === state.pipelineFocus) || PIPELINE_SEGMENTS[0];
  const focusTone = state.pipelineFocus === "stuck" ? "bad" : state.pipelineFocus === "broadcast" ? "warn" : "info";
  byId("pipelineDrilldown").innerHTML = `
    <div class="pipeline-drilldown-head">
      <div>
        <p class="eyebrow">Category Cases</p>
        <h4>${esc(focusMeta.label)}</h4>
      </div>
      <span class="status-pill ${focusTone}">${esc(pipelineNeeds.length)} cases</span>
    </div>
    <div class="pipeline-drilldown-list">
      ${pipelineNeeds.length
        ? pipelineNeeds
            .slice(0, 8)
            .map(
              (need) => `
                <article class="stack-card">
                  <div class="status-row">
                    <span class="status-pill ${badgeTone(need.status)}">${esc(need.status)}</span>
                    <span class="status-pill ${badgeTone(need.internal_status)}">${esc(need.internal_status)}</span>
                    <span class="status-pill info">${esc(need.curation_age_days || 0)} days</span>
                  </div>
                  <h4>${esc(need.organization_name)}</h4>
                  <p class="helper-text">${esc(clipText(need.problem_statement, 165))}</p>
                  <div class="card-actions">
                    <span class="meta-text">${esc(`${need.state || "Unknown state"}${need.district ? ` / ${need.district}` : ""}`)}</span>
                    <button class="btn btn-secondary" data-open-need-id="${esc(need.id)}">Open Need</button>
                  </div>
                </article>
              `,
            )
            .join("")
        : `<div class="empty-state">No cases are currently sitting in this pipeline category.</div>`}
    </div>
  `;

  const priorityNeeds = needs
    .filter((need) => Number(need.curation_age_days || 0) >= 7 || ["Blocked", "Stalled"].includes(need.internal_status))
    .sort((a, b) => Number(b.curation_age_days || 0) - Number(a.curation_age_days || 0))
    .slice(0, 8);

  byId("priorityList").innerHTML = priorityNeeds.length
    ? priorityNeeds
        .map(
          (need) => `
            <article class="stack-card">
              <div class="status-row">
                <span class="status-pill ${badgeTone(need.internal_status)}">${esc(need.internal_status)}</span>
                <span class="status-pill bad">${esc(need.curation_age_days || 0)} days</span>
              </div>
              <h4>${esc(need.organization_name)}</h4>
              <p class="helper-text">${esc(need.problem_statement.slice(0, 170))}${need.problem_statement.length > 170 ? "..." : ""}</p>
            </article>
          `,
        )
        .join("")
    : `<div class="empty-state">No urgent aging needs right now.</div>`;

  const workload = state.data.curators.map((curator) => ({
    label: curator.display_name,
    value: needs.filter((need) => need.curator_id === curator.id).length,
  }));
  renderBarList("workloadChart", workload, "good");

  const stateCounts = topEntries(countBy(needs, (need) => normalizeText(need.state) || "Unknown"));
  renderBarList(
    "stateChart",
    stateCounts.map(([label, value]) => ({ label, value })),
    "info",
  );

  const categoryCounts = {};
  needs.forEach((need) =>
    parseArray(need.curated_need).forEach((item) => {
      categoryCounts[item] = (categoryCounts[item] || 0) + 1;
    }),
  );

  byId("categoryChart").innerHTML = topEntries(categoryCounts, 12)
    .map(([label, value]) => `<span>${esc(label)} (${esc(value)})</span>`)
    .join("");

  byId("approvalSummary").innerHTML = `
    <article class="stack-card">
      <div class="status-row">
        <span class="status-pill warn">${esc(state.data.pendingNeeds.length)} pending needs</span>
        <span class="status-pill info">${esc(state.data.pendingUpdates.length)} pending updates</span>
      </div>
      <p class="helper-text">New submissions go to admin approval first. Assigned curators submit status updates that are only applied after admin review.</p>
    </article>
  `;
}

function renderBarList(targetId, items, tone) {
  const target = byId(targetId);
  if (!target) return;
  const max = Math.max(...items.map((item) => Number(item.value || 0)), 1);
  target.innerHTML = items
    .map(
      (item) => `
        <div class="bar-row">
          <span>${esc(item.label)}</span>
          <div class="bar-track"><div class="bar-fill ${tone === "warn" ? "warn" : tone === "bad" ? "bad" : ""}" style="width:${(Number(item.value || 0) / max) * 100}%"></div></div>
          <strong>${esc(item.value)}</strong>
        </div>
      `,
    )
    .join("");
}

function renderFilters() {
  if (!byId("statusFilter")) return;
  const statusOptions = ["all", ...new Set(state.data.needs.map((need) => need.status).filter(Boolean))];
  const stateOptions = ["all", ...new Set(state.data.needs.map((need) => normalizeText(need.state)).filter(Boolean))];

  document.getElementById("statusFilter").innerHTML = statusOptions
    .map((value) => `<option value="${esc(value)}">${esc(value === "all" ? "All statuses" : value)}</option>`)
    .join("");
  document.getElementById("curatorFilter").innerHTML = [
    `<option value="all">All curators</option>`,
    `<option value="unassigned">Unassigned</option>`,
    ...state.data.curators.map((curator) => `<option value="${esc(curator.id)}">${esc(curator.display_name)}</option>`),
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
        <p class="detail-note">${esc(need.curation_notes || "No curation notes have been recorded yet.")}</p>
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
  const matches = await store.searchMatchesForNeed(need);

  matchesEl.innerHTML = matches.length
    ? matches
        .map((match) => {
          const email = match.trader?.email || "";
          return `
            <article class="match-card">
              <div class="match-head">
                <div>
                  <p class="eyebrow">${esc(match.offering_category || "GRE Offering")}</p>
                  <h4>${esc(match.offering_name || match.solution?.solution_name || "Unnamed match")}</h4>
                </div>
                <span class="status-pill match-score-pill">${esc(`Match ${match.matchScore}`)}</span>
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
        .join("")
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

  const grouped = state.data.options.reduce((acc, option) => {
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
  if (!chip || !text || !logout) return;

  if (state.adminSession) {
    chip.textContent = "Unlocked";
    chip.className = "chip good";
    text.textContent = `Admin session active for ${state.adminSession.username}. Intake approvals and curator update approvals are available below.`;
    logout.classList.remove("hidden");
  } else {
    chip.textContent = "Locked";
    chip.className = "chip muted";
    text.textContent = "Admin approval is required for new intake records and curator-submitted status updates.";
    logout.classList.add("hidden");
  }
}

function switchView(nextView) {
  state.view = nextView;
  document.querySelectorAll(".tab").forEach((button) => button.classList.toggle("active", button.dataset.view === nextView));
  document.querySelectorAll(".view").forEach((section) => section.classList.toggle("active", section.id === `${nextView}View`));
}

async function rerender() {
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
  await renderMatches();
}

async function refreshAll() {
  await store.loadBaseData();
  await store.loadAdminSnapshot().catch(() => {
    state.data.pendingNeeds = [];
    state.data.pendingUpdates = [];
  });
  if (!state.selectedNeedId && state.data.needs[0]) state.selectedNeedId = state.data.needs[0].id;
  await rerender();
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
    renderQueue();
    renderNeedDetail();
    renderWorkbench();
    await renderMatches();
  });

  byId("pipelineBoard")?.addEventListener("click", (event) => {
    const card = event.target.closest("[data-pipeline-segment]");
    if (!card) return;
    state.pipelineFocus = card.dataset.pipelineSegment;
    renderOverview();
  });

  byId("pipelineDrilldown")?.addEventListener("click", async (event) => {
    const button = event.target.closest("[data-open-need-id]");
    if (!button) return;
    state.selectedNeedId = button.dataset.openNeedId;
    switchView("operations");
    renderQueue();
    renderNeedDetail();
    renderWorkbench();
    await renderMatches();
  });

  byId("refreshBtn")?.addEventListener("click", refreshAll);

  const dialog = byId("needDialog");
  byId("newNeedBtn")?.addEventListener("click", () => dialog?.showModal());
  byId("closeNeedDialog")?.addEventListener("click", () => dialog?.close());

  byId("needForm")?.addEventListener("submit", async (event) => {
    event.preventDefault();
    const form = new FormData(event.target);
    await store.createNeed(Object.fromEntries(form.entries()));
    event.target.reset();
    dialog?.close();
    await refreshAll();
    toast("Need submitted to the admin approval queue.");
  });

  byId("adminLoginForm")?.addEventListener("submit", async (event) => {
    event.preventDefault();
    const form = new FormData(event.target);
    await store.adminLogin(form.get("username"), form.get("password"));
    event.target.reset();
    await refreshAll();
    if (byId("adminView")) switchView("admin");
    toast("Admin access unlocked.");
  });

  byId("adminLogoutBtn")?.addEventListener("click", async () => {
    await store.adminLogout();
    await refreshAll();
    toast("Admin session closed.");
  });

  byId("saveOptionBtn")?.addEventListener("click", async () => {
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
  });

  byId("actionWorkbench")?.addEventListener("click", async (event) => {
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
  });

  byId("actionWorkbench")?.addEventListener("submit", async (event) => {
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
  });

  byId("adminView")?.addEventListener("click", async (event) => {
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
  });

  byId("matchResults")?.addEventListener("click", async (event) => {
    const button = event.target.closest("[data-action='email-provider']");
    if (!button) return;
    if (!state.adminToken) {
      toast("Login as admin to send provider outreach from the GRE mailbox.");
      return;
    }
    const result = await store.sendProviderIntro(state.selectedNeedId, button.dataset.providerEmail);
    toast(result.message || "Provider outreach email triggered.");
  });
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
