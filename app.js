const sampleOptions = [
  { id: "status-new", option_type: "status", label: "New", sort_order: 1 },
  { id: "status-accepted", option_type: "status", label: "Accepted", sort_order: 2 },
  { id: "status-progress", option_type: "status", label: "In progress", sort_order: 3 },
  { id: "status-closed", option_type: "status", label: "Closed", sort_order: 4 },
  { id: "internal-providers", option_type: "internal_status", label: "Need solution providers", sort_order: 1 },
  { id: "internal-connected", option_type: "internal_status", label: "Connection made", sort_order: 2 },
  { id: "internal-blocked", option_type: "internal_status", label: "Blocked", sort_order: 3 },
  { id: "internal-stalled", option_type: "internal_status", label: "Stalled", sort_order: 4 },
  { id: "category-capacity", option_type: "category", label: "Capacity building", sort_order: 1 },
  { id: "category-training", option_type: "category", label: "Training", sort_order: 2 },
  { id: "category-business", option_type: "category", label: "Business development", sort_order: 3 },
  { id: "category-infra", option_type: "category", label: "Infrastructure", sort_order: 4 },
  { id: "action-call", option_type: "next_action", label: "Schedule curation call", sort_order: 1 },
  { id: "action-provider", option_type: "next_action", label: "Find provider match", sort_order: 2 },
  { id: "action-seeker", option_type: "next_action", label: "Follow up with seeker", sort_order: 3 },
];

const sampleCurators = [
  { id: "c1", display_name: "Tanmay Mukherji", email: "tanmay@greenruraleconomy.in" },
  { id: "c2", display_name: "Phaneesh K", email: "phaneesh@greenruraleconomy.in" },
  { id: "c3", display_name: "Swati Singh", email: "swati@greenruraleconomy.in" },
  { id: "c4", display_name: "Shaifali Nagar", email: "shaifali@greenruraleconomy.in" },
];

const sampleProviders = [
  {
    id: "p1",
    name: "Rural Livestock Capacity Lab",
    email: "connect@rurallivestocklab.org",
    organization_type: "Technical Partner",
    states_served: ["Andhra Pradesh", "Odisha"],
    solution_tags: ["Capacity building", "Training", "Livestock health"],
    notes: "Strong fit for livestock extension, market linkages, and local training of trainers.",
  },
  {
    id: "p2",
    name: "Makhana Mechanisation Network",
    email: "machinery@makhananetwork.in",
    organization_type: "Solution Provider",
    states_served: ["Bihar"],
    solution_tags: ["Business development", "Machinery", "Vendor"],
    notes: "Maintains directory of machinery suppliers, indicative pricing, and after-sales partners.",
  },
  {
    id: "p3",
    name: "FPO Enterprise Clinic",
    email: "fielddesk@fpoenterprise.org",
    organization_type: "Advisory",
    states_served: ["Odisha", "Jharkhand"],
    solution_tags: ["Business consultation", "Business mentoring", "Business development"],
    notes: "Useful for business plans, board training, and vernacular planning support.",
  },
];

const sampleNeeds = [
  {
    id: "156",
    organization_name: "RURAL RECONSTRUCTION AND DEVELOPMENT SOCIETY",
    contact_person: "GangiReddy Vutukuri",
    seeker_email: "rrds111@gmail.com",
    seeker_phone: "9989988008",
    problem_statement: "The community lacks proper knowledge and market linkages for goat, sheep, and livestock rearing. Support is needed for capacity building, marketing skills, and animal health services.",
    state: "Andhra Pradesh",
    district: "Nellore",
    status: "In progress",
    internal_status: "Connection made",
    curator_id: "c4",
    curation_call_date: "2026-04-27",
    curation_age_days: 2,
    curation_notes: "Initial curation completed. Need provider intro pack before field-level meeting.",
    curated_need: ["Capacity building", "Training"],
    demand_broadcast_needed: false,
    solutions_shared_count: 1,
    invited_providers_count: 1,
    next_action: "Follow up with seeker",
    last_updated_at: "2026-04-29T09:00:00Z",
  },
  {
    id: "155",
    organization_name: "Sarva Seva Samity Sanstha (4S-India)",
    contact_person: "Gaurav",
    seeker_email: "gaurav4sindia@gmail.com",
    seeker_phone: "8340643639",
    problem_statement: "Need solution providers for Makhana Harvesting Machine and Makhana Popping Machine, along with vendor details and pricing.",
    state: "Bihar",
    district: "Katihar",
    status: "Accepted",
    internal_status: "Need solution providers",
    curator_id: "c3",
    curation_call_date: "2026-04-27",
    curation_age_days: 2,
    curation_notes: "Need verified vendors, approximate pricing, and operational guidance.",
    curated_need: ["Business development", "Vendor"],
    demand_broadcast_needed: false,
    solutions_shared_count: 0,
    invited_providers_count: 0,
    next_action: "Find provider match",
    last_updated_at: "2026-04-29T08:15:00Z",
  },
  {
    id: "154",
    organization_name: "Centre for Youth and Social Development",
    contact_person: "Kajal Pradhan",
    seeker_email: "kajalpradhan@cysd.org",
    seeker_phone: "7608009156",
    problem_statement: "Ten FPOs in tribal Odisha need training to build business development plans in Odia for their boards of directors.",
    state: "Odisha",
    district: "Khordha",
    status: "In progress",
    internal_status: "Need solution providers",
    curator_id: "c1",
    curation_call_date: "2026-04-27",
    curation_age_days: 2,
    curation_notes: "Demand is strong. Could become a reusable learning product.",
    curated_need: ["Business consultation", "Business development", "Business mentoring"],
    demand_broadcast_needed: true,
    solutions_shared_count: 0,
    invited_providers_count: 2,
    next_action: "Find provider match",
    last_updated_at: "2026-04-29T08:45:00Z",
  },
  {
    id: "141",
    organization_name: "Shivganga Samagra Gramvikas Parishad",
    contact_person: "Field Team",
    seeker_email: "ops@shivganga.org",
    seeker_phone: "9000000000",
    problem_statement: "Need support for enterprise planning and market linkage across a cluster of rural producer groups.",
    state: "Madhya Pradesh",
    district: "Jhabua",
    status: "Accepted",
    internal_status: "Need solution providers",
    curator_id: "c1",
    curation_call_date: "2026-03-12",
    curation_age_days: 48,
    curation_notes: "Several provider conversations happened but seeker has not received a closed-loop shortlist.",
    curated_need: ["Business development", "Branding"],
    demand_broadcast_needed: false,
    solutions_shared_count: 8,
    invited_providers_count: 3,
    next_action: "Follow up with seeker",
    last_updated_at: "2026-04-28T13:30:00Z",
  },
  {
    id: "051",
    organization_name: "Feedback Advisory",
    contact_person: "Coordination Desk",
    seeker_email: "desk@feedbackadvisory.org",
    seeker_phone: "",
    problem_statement: "Need has stalled after multiple outreach attempts and requires admin escalation.",
    state: "Karnataka",
    district: "Bengaluru",
    status: "In progress",
    internal_status: "Stalled",
    curator_id: "c2",
    curation_call_date: "2025-12-04",
    curation_age_days: 146,
    curation_notes: "Seeker not responding. Provider interest exists but no confirmation loop closed.",
    curated_need: ["Infrastructure"],
    demand_broadcast_needed: false,
    solutions_shared_count: 0,
    invited_providers_count: 1,
    next_action: "Escalate to admin",
    last_updated_at: "2026-04-27T16:10:00Z",
  },
];

const sampleUpdates = {
  "156": [
    { label: "Inbound logged", at: "2026-04-27T16:00:53Z", note: "Need received from GRE website." },
    { label: "Curator assigned", at: "2026-04-27T18:20:00Z", note: "Assigned to Shaifali Nagar." },
    { label: "Connection made", at: "2026-04-29T09:00:00Z", note: "One livestock capacity partner identified." },
  ],
  "155": [
    { label: "Inbound logged", at: "2026-04-25T11:28:04Z", note: "Machine requirement captured from seeker." },
    { label: "Curation completed", at: "2026-04-27T12:00:00Z", note: "Need sharpened to vendor and pricing discovery." },
  ],
  "154": [
    { label: "Inbound logged", at: "2026-04-18T11:46:56Z", note: "Business plan training need captured." },
    { label: "Broadcast suggested", at: "2026-04-28T15:30:00Z", note: "Could benefit from wider provider response." },
  ],
  "141": [
    { label: "Need aging alert", at: "2026-04-28T13:30:00Z", note: "Open for 48 days with no closure signal." },
  ],
};

function safeText(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function formatDate(value) {
  if (!value) return "Not set";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? String(value) : date.toLocaleString("en-IN", { dateStyle: "medium", timeStyle: "short" });
}

function slugStatus(value) {
  return String(value || "new").toLowerCase().replaceAll(/\s+/g, "-");
}

class GreMisStore {
  constructor() {
    this.config = window.APP_CONFIG || {};
    this.client = null;
    this.fallbackMode = true;
    this.curators = [];
    this.options = [];
    this.providers = [];
    this.needs = [];
  }

  getClient() {
    if (this.client) return this.client;
    if (!this.config.SUPABASE_URL || !this.config.SUPABASE_ANON_KEY || !window.supabase?.createClient) {
      return null;
    }
    this.client = window.supabase.createClient(this.config.SUPABASE_URL, this.config.SUPABASE_ANON_KEY);
    this.fallbackMode = false;
    return this.client;
  }

  async loadAll() {
    const client = this.getClient();
    if (!client) {
      this.loadFallback();
      return;
    }
    try {
      const [curators, options, providers, needs] = await Promise.all([
        client.from("gre_mis_curators").select("id, display_name, email, is_active").eq("is_active", true).order("display_name"),
        client.from("gre_mis_options").select("id, option_type, label, sort_order, is_active").eq("is_active", true).order("sort_order"),
        client.from("gre_mis_solution_providers").select("id, name, email, organization_type, states_served, solution_tags, notes, is_active").eq("is_active", true).order("name"),
        client.from("gre_mis_needs").select("*").order("requested_on", { ascending: false }),
      ]);
      if (curators.error || options.error || providers.error || needs.error) {
        throw new Error(curators.error?.message || options.error?.message || providers.error?.message || needs.error?.message || "Could not load dashboard data.");
      }
      this.curators = curators.data || [];
      this.options = options.data || [];
      this.providers = providers.data || [];
      this.needs = (needs.data || []).map((row) => ({
        ...row,
        curated_need: row.curated_need || [],
      }));
    } catch (error) {
      console.error(error);
      this.loadFallback();
    }
  }

  loadFallback() {
    this.fallbackMode = true;
    this.curators = structuredClone(sampleCurators);
    this.options = structuredClone(sampleOptions);
    this.providers = structuredClone(sampleProviders);
    this.needs = structuredClone(sampleNeeds);
  }

  async callAdminFunction(payload) {
    const client = this.getClient();
    if (!client || this.fallbackMode) return null;
    const token = (await client.auth.getSession()).data.session?.access_token;
    const response = await fetch(`${this.config.SUPABASE_URL}/functions/v1/${this.config.ADMIN_FUNCTION || "gre-mis-admin"}`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        apikey: this.config.SUPABASE_ANON_KEY,
        Authorization: `Bearer ${token || this.config.SUPABASE_ANON_KEY}`,
      },
      body: JSON.stringify(payload),
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(data.error || "Admin function request failed.");
    return data;
  }

  getNeedUpdates(needId) {
    return structuredClone(sampleUpdates[needId] || []);
  }

  async saveNeed(payload) {
    const client = this.getClient();
    const need = {
      id: payload.id || `${Date.now()}`,
      organization_name: payload.organization_name,
      contact_person: payload.contact_person,
      seeker_email: payload.seeker_email,
      seeker_phone: payload.seeker_phone || "",
      state: payload.state,
      district: payload.district || "",
      problem_statement: payload.problem_statement,
      status: "New",
      internal_status: "Need solution providers",
      curated_need: String(payload.categories || "").split(",").map((item) => item.trim()).filter(Boolean),
      curator_id: null,
      curation_call_date: null,
      curation_age_days: 0,
      curation_notes: "",
      demand_broadcast_needed: false,
      solutions_shared_count: 0,
      invited_providers_count: 0,
      next_action: "Schedule curation call",
      last_updated_at: new Date().toISOString(),
      requested_on: new Date().toISOString(),
    };
    if (!client) {
      this.needs.unshift(need);
      return need;
    }
    const { data, error } = await client.from("gre_mis_needs").insert(need).select("*").single();
    if (error) throw new Error(error.message);
    this.needs.unshift(data);
    return data;
  }

  async assignCurator(needId, curatorId) {
    const need = this.needs.find((item) => item.id === needId);
    if (!need) throw new Error("Need not found.");
    need.curator_id = curatorId;
    need.status = need.status === "New" ? "Accepted" : need.status;
    need.last_updated_at = new Date().toISOString();
    if (!this.getClient() || this.fallbackMode) return need;
    await this.callAdminFunction({
      action: "assignCurator",
      needId,
      curatorId,
    });
    return need;
  }

  async updateOption(optionType, label) {
    const existing = this.options.find((item) => item.option_type === optionType && item.label.toLowerCase() === label.toLowerCase());
    if (existing) return existing;
    const row = {
      id: `local-${Date.now()}`,
      option_type: optionType,
      label,
      sort_order: this.options.filter((item) => item.option_type === optionType).length + 1,
      is_active: true,
    };
    if (!this.getClient() || this.fallbackMode) {
      this.options.push(row);
      return row;
    }
    await this.callAdminFunction({
      action: "upsertOption",
      optionType,
      label,
    });
    this.options.push(row);
    return row;
  }

  findProviderMatches(need) {
    const tags = new Set((need.curated_need || []).map((item) => item.toLowerCase()));
    return this.providers
      .map((provider) => {
        const score = (provider.solution_tags || []).reduce((total, tag) => total + (tags.has(String(tag).toLowerCase()) ? 3 : 0), 0)
          + ((provider.states_served || []).includes(need.state) ? 2 : 0);
        return { ...provider, score };
      })
      .filter((provider) => provider.score > 0)
      .sort((a, b) => b.score - a.score);
  }

  async sendProviderIntro(need, providerId) {
    const provider = this.providers.find((item) => item.id === providerId);
    if (!provider) throw new Error("Solution provider not found.");
    if (!this.getClient() || this.fallbackMode) {
      return {
        ok: true,
        preview: `Would send from tanmay@greenruraleconomy.in to ${provider.email} with CC to ${need.seeker_email}.`,
      };
    }
    return this.callAdminFunction({
      action: "sendProviderIntro",
      needId: need.id,
      providerId,
    });
  }
}

const store = new GreMisStore();

const ui = {
  role: "admin",
  selectedNeedId: null,
  filters: {
    status: "all",
    curator: "all",
    state: "all",
    search: "",
  },
};

const roleCopy = {
  admin: "Admin sees network-wide analytics, bottlenecks, option management, and approvals across the GRE platform.",
  curator: "Curator sees a personal action desk focused on assigned needs, follow-ups, seeker history, and provider matching.",
};

function getCuratorName(curatorId) {
  return store.curators.find((item) => item.id === curatorId)?.display_name || "Unassigned";
}

function computeMetrics() {
  const metrics = [
    ["New Needs", store.needs.filter((item) => item.status === "New").length, "Fresh inbound demand awaiting action."],
    ["In Curation", store.needs.filter((item) => item.status === "In progress").length, "Needs actively being worked by curators."],
    ["Need Solution Providers", store.needs.filter((item) => item.internal_status === "Need solution providers").length, "Needs requiring solution/provider discovery."],
    ["Broadcast Required", store.needs.filter((item) => item.demand_broadcast_needed).length, "Needs that may benefit from open ecosystem outreach."],
    ["Connection Made", store.needs.filter((item) => item.internal_status === "Connection made").length, "Needs where a meaningful connection already exists."],
    ["Overdue / Stuck", store.needs.filter((item) => Number(item.curation_age_days || 0) >= 7 || ["Blocked", "Stalled"].includes(item.internal_status)).length, "Cases needing escalation or movement today."],
  ];
  return metrics;
}

function renderKpis() {
  document.getElementById("kpiStrip").innerHTML = computeMetrics().map(([label, value, note]) => `
    <article class="metric-card">
      <p class="eyebrow">${safeText(label)}</p>
      <strong>${safeText(value)}</strong>
      <p class="metric-note">${safeText(note)}</p>
    </article>
  `).join("");
}

function getVisibleNeeds() {
  let needs = [...store.needs];
  if (ui.role === "curator") {
    needs = needs.filter((item) => item.curator_id === "c1" || !item.curator_id);
  }
  if (ui.filters.status !== "all") needs = needs.filter((item) => item.status === ui.filters.status);
  if (ui.filters.curator !== "all") needs = needs.filter((item) => (item.curator_id || "unassigned") === ui.filters.curator);
  if (ui.filters.state !== "all") needs = needs.filter((item) => item.state === ui.filters.state);
  if (ui.filters.search) {
    const query = ui.filters.search.toLowerCase();
    needs = needs.filter((item) => [item.organization_name, item.problem_statement, (item.curated_need || []).join(", ")].join(" ").toLowerCase().includes(query));
  }
  return needs.sort((a, b) => Number(b.curation_age_days || 0) - Number(a.curation_age_days || 0));
}

function renderFilters() {
  const statusOptions = ["all", ...new Set(store.options.filter((item) => item.option_type === "status").map((item) => item.label))];
  const states = ["all", ...new Set(store.needs.map((item) => item.state).filter(Boolean))];
  document.getElementById("statusFilter").innerHTML = statusOptions.map((value) => `<option value="${safeText(value)}">${safeText(value === "all" ? "All statuses" : value)}</option>`).join("");
  document.getElementById("curatorFilter").innerHTML = [`<option value="all">All curators</option>`, `<option value="unassigned">Unassigned</option>`, ...store.curators.map((item) => `<option value="${safeText(item.id)}">${safeText(item.display_name)}</option>`)].join("");
  document.getElementById("stateFilter").innerHTML = states.map((value) => `<option value="${safeText(value)}">${safeText(value === "all" ? "All states" : value)}</option>`).join("");
  document.getElementById("statusFilter").value = ui.filters.status;
  document.getElementById("curatorFilter").value = ui.filters.curator;
  document.getElementById("stateFilter").value = ui.filters.state;
  document.getElementById("searchFilter").value = ui.filters.search;
}

function renderQueue() {
  const visible = getVisibleNeeds();
  document.getElementById("queueSummary").innerHTML = `
    <span class="status-badge status-overdue">${visible.filter((item) => item.curation_age_days >= 7).length} aging</span>
    <span class="status-badge status-awaiting">${visible.filter((item) => !item.curator_id).length} unassigned</span>
    <span class="status-badge status-broadcast">${visible.filter((item) => item.demand_broadcast_needed).length} broadcast</span>
  `;
  document.getElementById("needsList").innerHTML = visible.map((need) => {
    const overdue = Number(need.curation_age_days || 0) >= 7 || ["Blocked", "Stalled"].includes(need.internal_status);
    return `
      <article class="need-card ${ui.selectedNeedId === need.id ? "active" : ""}" data-need-id="${safeText(need.id)}">
        <div class="status-row">
          <span class="status-badge status-${safeText(slugStatus(need.status))}">${safeText(need.status)}</span>
          <span class="status-badge status-${safeText(slugStatus(need.internal_status))}">${safeText(need.internal_status)}</span>
          ${overdue ? `<span class="status-badge status-overdue">Needs movement</span>` : ""}
        </div>
        <h4>${safeText(need.organization_name)}</h4>
        <p>${safeText(need.problem_statement.slice(0, 165))}${need.problem_statement.length > 165 ? "..." : ""}</p>
        <div class="need-meta">
          <span>${safeText(need.state)}${need.district ? ` / ${safeText(need.district)}` : ""}</span>
          <span>Curator: ${safeText(getCuratorName(need.curator_id))}</span>
          <span>Age: ${safeText(need.curation_age_days)} days</span>
        </div>
      </article>
    `;
  }).join("");
  if (!ui.selectedNeedId && visible[0]) ui.selectedNeedId = visible[0].id;
}

function renderInsights() {
  const maxCurator = Math.max(...store.curators.map((curator) => store.needs.filter((need) => need.curator_id === curator.id).length), 1);
  document.getElementById("allocationChart").innerHTML = `
    <div class="info-card">
      <h4>Curator Allocation</h4>
      <div class="bar-list">
        ${store.curators.map((curator) => {
          const count = store.needs.filter((need) => need.curator_id === curator.id).length;
          return `
            <div class="bar-row">
              <span>${safeText(curator.display_name)}</span>
              <div class="bar-track"><div class="bar-fill" style="width:${(count / maxCurator) * 100}%"></div></div>
              <strong>${count}</strong>
            </div>
          `;
        }).join("")}
      </div>
    </div>
  `;
  const agingBuckets = [
    ["0-2 days", store.needs.filter((item) => item.curation_age_days <= 2).length, ""],
    ["3-7 days", store.needs.filter((item) => item.curation_age_days > 2 && item.curation_age_days <= 7).length, "warn"],
    ["8-14 days", store.needs.filter((item) => item.curation_age_days > 7 && item.curation_age_days <= 14).length, "warn"],
    ["15+ days", store.needs.filter((item) => item.curation_age_days > 14).length, "alert"],
  ];
  const maxAging = Math.max(...agingBuckets.map(([, value]) => value), 1);
  document.getElementById("agingChart").innerHTML = `
    <div class="info-card">
      <h4>Aging of Needs</h4>
      <div class="bar-list">
        ${agingBuckets.map(([label, value, css]) => `
          <div class="bar-row">
            <span>${safeText(label)}</span>
            <div class="bar-track"><div class="bar-fill ${css}" style="width:${(value / maxAging) * 100}%"></div></div>
            <strong>${value}</strong>
          </div>
        `).join("")}
      </div>
    </div>
  `;
  const categoryCounts = {};
  store.needs.forEach((need) => (need.curated_need || []).forEach((category) => {
    categoryCounts[category] = (categoryCounts[category] || 0) + 1;
  }));
  const topCategories = Object.entries(categoryCounts).sort((a, b) => b[1] - a[1]).slice(0, 5);
  const maxCategory = Math.max(...topCategories.map(([, value]) => value), 1);
  document.getElementById("categoryChart").innerHTML = `
    <div class="info-card">
      <h4>Top Need Categories</h4>
      <div class="bar-list">
        ${topCategories.map(([label, value]) => `
          <div class="bar-row">
            <span>${safeText(label)}</span>
            <div class="bar-track"><div class="bar-fill" style="width:${(value / maxCategory) * 100}%"></div></div>
            <strong>${value}</strong>
          </div>
        `).join("")}
      </div>
    </div>
  `;
}

function renderOptions() {
  const grouped = store.options.reduce((acc, option) => {
    acc[option.option_type] ||= [];
    acc[option.option_type].push(option);
    return acc;
  }, {});
  document.getElementById("optionsList").innerHTML = Object.entries(grouped).map(([type, items]) => `
    <div class="info-card">
      <h4>${safeText(type.replaceAll("_", " "))}</h4>
      ${items.map((item) => `<div class="option-row"><span>${safeText(item.label)}</span><span class="status-badge status-accepted">Active</span></div>`).join("")}
    </div>
  `).join("");
}

function renderNeedDetail() {
  const need = store.needs.find((item) => item.id === ui.selectedNeedId);
  if (!need) {
    document.getElementById("needDetail").innerHTML = `<div class="detail-empty">No need selected.</div>`;
    document.getElementById("providerSearch").innerHTML = `<div class="detail-empty">Choose a need to search providers.</div>`;
    return;
  }
  const updates = store.getNeedUpdates(need.id);
  const curatorOptions = [`<option value="">Unassigned</option>`, ...store.curators.map((curator) => `<option value="${safeText(curator.id)}" ${curator.id === need.curator_id ? "selected" : ""}>${safeText(curator.display_name)}</option>`)].join("");
  document.getElementById("needDetail").innerHTML = `
    <div class="detail-stack">
      <div class="info-card">
        <div class="status-row">
          <span class="status-badge status-${safeText(slugStatus(need.status))}">${safeText(need.status)}</span>
          <span class="status-badge status-${safeText(slugStatus(need.internal_status))}">${safeText(need.internal_status)}</span>
          ${need.demand_broadcast_needed ? `<span class="status-badge status-broadcast">Broadcast suggested</span>` : ""}
        </div>
        <h4>${safeText(need.organization_name)}</h4>
        <p class="body-copy">${safeText(need.problem_statement)}</p>
      </div>
      <div class="detail-grid">
        <div class="detail-stack">
          <div class="info-card">
            <h4>Request Snapshot</h4>
            <div class="info-table">
              <div><small>Contact</small>${safeText(need.contact_person)}</div>
              <div><small>Email</small>${safeText(need.seeker_email)}</div>
              <div><small>Phone</small>${safeText(need.seeker_phone || "Not available")}</div>
              <div><small>Location</small>${safeText(`${need.state}${need.district ? ` / ${need.district}` : ""}`)}</div>
              <div><small>Curator</small>${safeText(getCuratorName(need.curator_id))}</div>
              <div><small>Curation Call</small>${safeText(need.curation_call_date || "Not scheduled")}</div>
              <div><small>Age</small>${safeText(need.curation_age_days)} days</div>
              <div><small>Solutions Shared</small>${safeText(need.solutions_shared_count)}</div>
            </div>
          </div>
          <div class="info-card">
            <h4>Curator Assignment</h4>
            <label>
              <span>Select curator</span>
              <select id="assignCuratorSelect">${curatorOptions}</select>
            </label>
            <div class="quick-actions">
              <button class="primary-btn" data-action="assign-curator" data-need-id="${safeText(need.id)}">Save Assignment</button>
              <button class="secondary-btn" data-action="copy-problem" data-need-id="${safeText(need.id)}">Copy Problem Statement</button>
            </div>
            <p class="body-copy">Assignment supports queue balancing. Provider outreach happens separately from the matched provider section below.</p>
          </div>
        </div>
        <div class="detail-stack">
          <div class="info-card">
            <h4>Next Best Action</h4>
            <div class="mini-metrics">
              <span class="status-badge status-awaiting">${safeText(need.next_action || "Not set")}</span>
              <span class="status-badge status-${need.curation_age_days >= 7 ? "overdue" : "accepted"}">${need.curation_age_days >= 7 ? "Escalation needed" : "On track"}</span>
            </div>
            <p>${safeText(need.curation_notes || "No curation notes added yet.")}</p>
          </div>
          <div class="info-card">
            <h4>Need Categories</h4>
            <div class="provider-tags">${(need.curated_need || []).map((item) => `<span>${safeText(item)}</span>`).join("") || "<span>Unclassified</span>"}</div>
          </div>
        </div>
      </div>
      <div class="timeline-card">
        <h4>Seeker Timeline</h4>
        <div class="timeline">
          ${updates.map((update) => `
            <div>
              <strong>${safeText(update.label)}</strong>
              <small>${safeText(formatDate(update.at))}</small>
              <p>${safeText(update.note)}</p>
            </div>
          `).join("") || "<p>No updates yet.</p>"}
        </div>
      </div>
    </div>
  `;
  renderProviderSearch(need);
}

function renderProviderSearch(need) {
  const matches = store.findProviderMatches(need);
  document.getElementById("providerSearch").innerHTML = `
    <div class="detail-stack">
      <div class="info-card">
        <h4>Recommended Matches for Need #${safeText(need.id)}</h4>
        <p>These matches use GRE solution categories and provider tags today. Once the Supabase solution and provider datasets are connected, this section can search the full directory in real time.</p>
      </div>
      <div class="provider-grid">
        ${matches.length ? matches.map((provider) => `
          <article class="provider-card">
            <header>
              <div>
                <h4>${safeText(provider.name)}</h4>
                <small>${safeText(provider.organization_type)} • ${safeText(provider.email)}</small>
              </div>
              <span class="status-badge status-accepted">Score ${safeText(provider.score)}</span>
            </header>
            <div class="provider-tags">${(provider.solution_tags || []).map((tag) => `<span>${safeText(tag)}</span>`).join("")}</div>
            <p>${safeText(provider.notes || "")}</p>
            <div class="quick-actions">
              <button class="primary-btn" data-action="send-provider" data-need-id="${safeText(need.id)}" data-provider-id="${safeText(provider.id)}">Email Problem Statement to Provider</button>
              <button class="action-btn" data-action="copy-provider-email" data-provider-id="${safeText(provider.id)}">Copy Provider Email</button>
            </div>
          </article>
        `).join("") : `<div class="detail-empty">No provider match found yet. Add solution tags or broadcast to ecosystem.</div>`}
      </div>
    </div>
  `;
}

function renderRoleSummary() {
  const suffix = store.fallbackMode ? " Demo data is currently shown because no live Supabase configuration was detected." : "";
  document.getElementById("roleSummary").textContent = `${roleCopy[ui.role]}${suffix}`;
}

function rerender() {
  renderRoleSummary();
  renderKpis();
  renderFilters();
  renderQueue();
  renderNeedDetail();
  renderInsights();
  renderOptions();
}

function toast(message) {
  window.alert(message);
}

function wireEvents() {
  document.querySelectorAll(".role-pill").forEach((button) => {
    button.addEventListener("click", () => {
      ui.role = button.dataset.role;
      document.querySelectorAll(".role-pill").forEach((item) => item.classList.toggle("active", item === button));
      rerender();
    });
  });

  document.getElementById("statusFilter").addEventListener("change", (event) => {
    ui.filters.status = event.target.value;
    rerender();
  });
  document.getElementById("curatorFilter").addEventListener("change", (event) => {
    ui.filters.curator = event.target.value;
    rerender();
  });
  document.getElementById("stateFilter").addEventListener("change", (event) => {
    ui.filters.state = event.target.value;
    rerender();
  });
  document.getElementById("searchFilter").addEventListener("input", (event) => {
    ui.filters.search = event.target.value.trim();
    rerender();
  });

  document.getElementById("needsList").addEventListener("click", (event) => {
    const card = event.target.closest("[data-need-id]");
    if (!card) return;
    ui.selectedNeedId = card.dataset.needId;
    rerender();
  });

  document.getElementById("needDetail").addEventListener("click", async (event) => {
    const button = event.target.closest("[data-action]");
    if (!button) return;
    const action = button.dataset.action;
    const need = store.needs.find((item) => item.id === button.dataset.needId);
    if (!need) return;
    if (action === "copy-problem") {
      await navigator.clipboard.writeText(need.problem_statement);
      toast("Problem statement copied.");
      return;
    }
    if (action === "assign-curator") {
      const select = document.getElementById("assignCuratorSelect");
      await store.assignCurator(need.id, select.value || null);
      toast("Curator assignment updated.");
      rerender();
    }
  });

  document.getElementById("providerSearch").addEventListener("click", async (event) => {
    const button = event.target.closest("[data-action]");
    if (!button) return;
    const need = store.needs.find((item) => item.id === button.dataset.needId);
    if (!need) return;
    if (button.dataset.action === "copy-provider-email") {
      const provider = store.providers.find((item) => item.id === button.dataset.providerId);
      await navigator.clipboard.writeText(provider?.email || "");
      toast("Provider email copied.");
      return;
    }
    if (button.dataset.action === "send-provider") {
      const result = await store.sendProviderIntro(need, button.dataset.providerId);
      toast(result.preview || "Provider introduction email triggered.");
    }
  });

  document.getElementById("saveOptionBtn").addEventListener("click", async () => {
    const optionType = document.getElementById("optionType").value;
    const optionLabel = document.getElementById("optionLabel").value.trim();
    if (!optionLabel) {
      toast("Enter an option label first.");
      return;
    }
    await store.updateOption(optionType, optionLabel);
    document.getElementById("optionLabel").value = "";
    rerender();
  });

  const dialog = document.getElementById("needDialog");
  document.getElementById("newNeedBtn").addEventListener("click", () => dialog.showModal());
  document.getElementById("closeNeedDialog").addEventListener("click", () => dialog.close());
  document.getElementById("seedDemoBtn").addEventListener("click", async () => {
    store.loadFallback();
    ui.selectedNeedId = store.needs[0]?.id || null;
    rerender();
    toast("Sample GRE MIS data loaded.");
  });
  document.getElementById("refreshBtn").addEventListener("click", async () => {
    await store.loadAll();
    rerender();
  });

  document.getElementById("needForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const form = new FormData(event.target);
    await store.saveNeed(Object.fromEntries(form.entries()));
    dialog.close();
    event.target.reset();
    rerender();
    toast("Need saved.");
  });
}

async function init() {
  await store.loadAll();
  ui.selectedNeedId = store.needs[0]?.id || null;
  renderFilters();
  wireEvents();
  rerender();
}

init().catch((error) => {
  console.error(error);
  toast(error.message || "Dashboard could not start.");
});
