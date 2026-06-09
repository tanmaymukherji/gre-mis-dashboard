const FALLBACK_CURATORS = [
  { id: "fallback-1", display_name: "Tanmay Mukherji", email: "tanmay@greenruraleconomy.in" },
  { id: "fallback-2", display_name: "Phaneesh K", email: "agri@greenruraleconomy.in" },
  { id: "fallback-3", display_name: "Swati Singh", email: "solution@greenruraleconomy.in" },
  { id: "fallback-4", display_name: "Shaifali Nagar", email: "help@greenruraleconomy.in" },
];

const bootstrapQuery = new URLSearchParams(window.location.search);
const sharedFormQuery = bootstrapQuery.get("sharedForm");
const GRAMEEE_SESSION_TRANSFER_PARAM = "grameeeAuthState";
const GRAMEEE_TRANSFER_CACHE_KEY = "grameee-last-transfer";
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
const DEFAULT_LSH_REQUEST_SUPPORT_TEMPLATE = `Hello Team LSH,

We are looking at some of our needs in the Livestock domain and would like to connect with your team. Do reach out to us.

Regards,
{{name}}
{{organisation}}`;
const DEFAULT_LSH_EMAIL_PROVIDER_TEMPLATE = `Hello Team LSH,

We would like to know more about {{offeringName}} and request a follow-up from your team.

Regards,
{{name}}
{{organisation}}`;

const state = {
  view: "overview",
  sharedFormMode: sharedFormQuery === "solution"
    ? "solution"
    : sharedFormQuery === "need"
      ? "need-intake"
      : "",
  standalonePublicFormMode: bootstrapQuery.get("publicForm") || "",
  grameeeEmbed: bootstrapQuery.get("grameeeEmbed") === "1",
  formLanguage: normalizeSharedFormLanguage(
    bootstrapQuery.get("formLang")
    || localStorage.getItem(`grameee-shared-form-lang:${sharedFormQuery || bootstrapQuery.get("publicForm") || "default"}`)
    || "en",
  ),
  embeddedContextOrigin: "",
  embeddedActor: null,
  translationCache: new Map(),
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
  localNeedReviewId: "",
  solutionTags: [],
  solutionGeographies: [],
  solutionGeographySuggestions: ["India"],
  needTags: [],
  needDeploymentLocations: [],
  needDeploymentSuggestions: ["India"],
  adminUserRoleTab: "admin",
  adminDeskTab: "data-sync",
  localSolutionFilters: {
    search: "",
    provider: "all",
  },
  localNeedFilters: {
    search: "",
    category: "all",
  },
  manualSolutionSearch: {
    provider: "",
    category: "",
    domain6m: "",
    offeringType: "",
    valuechain: "",
    application: "",
    language: "",
    geography: "",
    keyword: "",
    results: [],
    searched: false,
    loading: false,
    page: 1,
  },
  grameeeTransfer: null,
  grameeeBridgeMessage: "",
  puterModelsLoaded: false,
  puterModels: [],
    puterRecommendations: {},
    mailTemplates: {
      providerIntroTemplate: DEFAULT_PROVIDER_INTRO_TEMPLATE,
      curatorForwardTemplate: DEFAULT_CURATOR_FORWARD_TEMPLATE,
      solutionSeekerTemplate: DEFAULT_SOLUTION_SEEKER_TEMPLATE,
      needSeekerTemplate: DEFAULT_NEED_SEEKER_TEMPLATE,
      inboundAutoSyncEnabled: true,
      lshContactEmails: ["subekkumar@pradan.net"],
      lshHelpCcEmails: ["help@greenruraleconomy.in"],
      lshRequestSupportTemplate: DEFAULT_LSH_REQUEST_SUPPORT_TEMPLATE,
      lshEmailProviderTemplate: DEFAULT_LSH_EMAIL_PROVIDER_TEMPLATE,
    },
  pendingMailReview: null,
  adminAuditSearch: "",
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
      impactAuditLogs: {
        emailLogs: [],
        viewLogs: [],
      },
      users: [],
      localSolutions: [],
      localNeeds: [],
  },
};

const SHARED_FORM_TRANSLATABLE_SELECTOR = [
  ".form-section-head .eyebrow",
  ".form-section-head h4",
  "label > span",
  ".field-label",
  ".helper-text",
  "button",
  "h3",
].join(", ");

const SHARED_FORM_STATIC_TRANSLATIONS = {
  hi: {
    "Solution Upload": "समाधान अपलोड",
    "Add Solution to GRE Review Queue": "GRE के लिए समाधान साझा करें",
    "Share a Solution for GRE": "GRE के लिए समाधान साझा करें",
    "Organisation": "संस्था",
    "Organisation and Submission Contact": "संस्था और संपर्क विवरण",
    "Existing Organisation": "मौजूदा संस्था",
    "Organisation Name": "संस्था का नाम",
    "Contact Person": "संपर्क व्यक्ति",
    "Contact Email": "संपर्क ईमेल",
    "Contact Phone": "संपर्क फोन",
    "Offering": "ऑफरिंग",
    "Offering Structure": "ऑफरिंग संरचना",
    "Offering Type": "ऑफरिंग प्रकार",
    "Offering Name": "ऑफरिंग का नाम",
    "Offering Description": "ऑफरिंग का विवरण",
    "Product Offering": "उत्पाद ऑफरिंग",
    "Product Details": "उत्पाद विवरण",
    "Grade/Capacity": "ग्रेड / क्षमता",
    "Cost": "लागत",
    "Can be quoted after finalising scope": "दायरा तय होने के बाद उद्धृत किया जा सकता है",
    "Lead Time": "लीड टाइम",
    "Product Brochure": "उत्पाद पुस्तिका",
    "Support Services": "सहायता सेवाएँ",
    "Contact Details for Product": "उत्पाद हेतु संपर्क विवरण",
    "Use this if the product contact differs from the organisation contact": "यदि उत्पाद का संपर्क संस्था के संपर्क से अलग हो तो इसका उपयोग करें",
    "Service Offering": "सेवा ऑफरिंग",
    "Service Delivery Details": "सेवा वितरण विवरण",
    "Facilitator Name": "सुविधादाता का नाम",
    "Facilitator Email Address": "सुविधादाता का ईमेल",
    "Facilitator Phone Number": "सुविधादाता का फोन नंबर",
    "Facilitator Details": "सुविधादाता का विवरण",
    "Languages": "भाषाएँ",
    "Geographies": "भौगोलिक क्षेत्र",
    "Search a block, city, state, or country": "ब्लॉक, शहर, राज्य या देश खोजें",
    "Add Location": "स्थान जोड़ें",
    "Duration": "अवधि",
    "Duration Units": "अवधि इकाई",
    "Prerequisites - Participants and Training": "पूर्व आवश्यकताएँ - प्रतिभागी और प्रशिक्षण",
    "Location Availability": "सेवा उपलब्धता स्थान",
    "Remarks on Cost": "लागत पर टिप्पणी",
    "Support Post Service": "सेवा पश्चात सहायता",
    "Support Post Service Cost": "सेवा पश्चात सहायता लागत",
    "Is it Offered Online / Offline": "क्या यह ऑनलाइन / ऑफलाइन उपलब्ध है",
    "Certification Offered": "प्रमाणन उपलब्ध",
    "Service Offering Brochure": "सेवा ऑफरिंग पुस्तिका",
    "Knowledge Offering": "ज्ञान ऑफरिंग",
    "Knowledge Resource Details": "ज्ञान संसाधन विवरण",
    "Knowledge Offering Content": "ज्ञान ऑफरिंग सामग्री",
    "Knowledge Offering Content Link": "ज्ञान सामग्री लिंक",
    "Contact Details": "संपर्क विवरण",
    "Media and Tags": "मीडिया और टैग",
    "Supporting Assets": "सहायक सामग्री",
    "Offering Image": "ऑफरिंग चित्र",
    "Tags": "टैग",
    "Select or type a tag, then press Enter": "टैग चुनें या लिखें, फिर Enter दबाएँ",
    "Add Tag": "टैग जोड़ें",
    "AI Generate Tags": "AI से टैग बनाएं",
    "Use own AI with Puter for Tags": "टैग हेतु Puter के अपने AI का उपयोग करें",
    "Tags are compulsory for discovery. Generate a first draft or add them manually.": "खोज हेतु टैग अनिवार्य हैं। पहले ड्राफ्ट बनाएं या उन्हें मैन्युअली जोड़ें।",
    "Submit Solution for Approval": "स्वीकृति हेतु समाधान भेजें",
    "Share Page Link": "पेज लिंक साझा करें",
    "Sharing": "साझा करना",
    "External Submission Link": "बाहरी सबमिशन लिंक",
    "Request Solution": "समाधान अनुरोध",
    "Need Help": "मदद चाहिए",
    "Share a Need with GRE": "GRE के साथ आवश्यकता साझा करें",
    "Account and Contact": "खाता और संपर्क",
    "GRE Account": "GRE खाता",
    "Email": "ईमेल",
    "Phone": "फोन",
    "Need Detail": "आवश्यकता विवरण",
    "Classification and Matching Inputs": "वर्गीकरण और मिलान इनपुट",
    "Need Type": "आवश्यकता प्रकार",
    "Thematic Area": "विषय क्षेत्र",
    "Dairy, Solar, Soap, Wild Mango": "डेयरी, सोलर, साबुन, जंगली आम",
    "Need Statement": "आवश्यकता विवरण",
    "Place of Deployment": "कार्यान्वयन का स्थान",
    "Search village, block, city, district, state, or country": "गाँव, ब्लॉक, शहर, जिला, राज्य या देश खोजें",
    "Broadcast to Ecosystem": "इकोसिस्टम में प्रसारित करें",
    "Allow this need to be broadcast wider if needed": "यदि आवश्यक हो तो इस आवश्यकता को व्यापक रूप से प्रसारित करने की अनुमति दें",
    "Request Solution": "समाधान अनुरोध",
    "Select": "चुनें",
    "Per day": "प्रति दिन",
    "Per hour": "प्रति घंटा",
    "Per person": "प्रति व्यक्ति",
    "Not Provided": "उपलब्ध नहीं",
    "Provided, if in the Service Location": "उपलब्ध, यदि सेवा स्थान पर हो",
    "Provided, for both Service Location and Outside Service Location": "उपलब्ध, सेवा स्थान और बाहरी स्थान दोनों पर",
    "No Cost": "कोई लागत नहीं",
    "No Cost if in Service Location, Additional Cost if Outside Service Location": "सेवा स्थान पर बिना लागत, बाहर अतिरिक्त लागत",
    "Additional Cost for both Service Location and Outside Service Location": "सेवा स्थान और बाहरी स्थान दोनों के लिए अतिरिक्त लागत",
    "At Service Seeker": "सेवा प्राप्तकर्ता के स्थान पर",
    "Others": "अन्य",
    "Service": "सेवा",
    "Product": "उत्पाद",
    "Knowledge": "ज्ञान",
    "Days": "दिन",
    "Hours": "घंटे",
    "Months": "महीने",
    "Weeks": "सप्ताह",
    "Online": "ऑनलाइन",
    "Offline": "ऑफलाइन",
    "Online & Offline": "ऑनलाइन और ऑफलाइन",
    "Provided": "उपलब्ध",
    "Not Provided": "उपलब्ध नहीं",
    "At Service Provider": "सेवा प्रदाता के स्थान पर",
    "At Service Seeker": "सेवा चाहने वाले के स्थान पर",
    "Others": "अन्य",
  },
};

function buildImpactAuditEvent(baseEvent = {}) {
  const session = state.userSession || state.adminSession || {};
  return {
    ...baseEvent,
    actorEmail: normalizeText(baseEvent.actorEmail || session.email || session.user_email || ""),
    actorName: String(baseEvent.actorName || session.full_name || session.display_name || session.username || "").trim(),
    surface: normalizeText(baseEvent.surface || "gre-mis"),
  };
}

function trackImpactCounter(counterKey, delta = 1, auditEvent = null) {
  const supabaseUrl = normalizeText(window.APP_CONFIG?.SUPABASE_URL || "");
  const supabaseAnonKey = normalizeText(window.APP_CONFIG?.SUPABASE_ANON_KEY || "");

  if (!supabaseUrl || !supabaseAnonKey) {
    return Promise.resolve(null);
  }

  const accessToken = normalizeText(getCookieValue("grameee_access_token"));
  const bearer = accessToken || supabaseAnonKey;

  return fetch(`${supabaseUrl}/functions/v1/grameee-admin`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      apikey: supabaseAnonKey,
      Authorization: `Bearer ${bearer}`,
    },
    body: JSON.stringify({
      action: "incrementImpactCounter",
      counterKey,
      delta,
      auditEvent: auditEvent ? buildImpactAuditEvent(auditEvent) : undefined,
    }),
    keepalive: true,
  }).catch(() => null);
}

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

const CURATED_NEED_OPTIONS = [
  "Branding",
  "Business Consultation",
  "Business Development",
  "Business Mentoring",
  "Capacity Building",
  "Connect and Collaborate",
  "Finance",
  "Funding",
  "Infrastructure",
  "Investments",
  "Machinery",
  "Market",
  "Packaging",
  "Technology",
  "Training",
  "Vendor",
];

const DEFAULT_LANGUAGE_OPTIONS = [
  "ENGLISH",
  "HINDI",
  "ASSAMESE",
  "BENGALI",
  "BODO",
  "DOGRI",
  "GUJARATI",
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

const HINDI_LANGUAGE_LABELS = {
  ENGLISH: "अंग्रेज़ी",
  HINDI: "हिंदी",
  ASSAMESE: "असमिया",
  BENGALI: "बांग्ला",
  BODO: "बोडो",
  DOGRI: "डोगरी",
  GUJARATI: "गुजराती",
  KANNADA: "कन्नड़",
  KASHMIRI: "कश्मीरी",
  KONKANI: "कोंकणी",
  MAITHILI: "मैथिली",
  MALAYALAM: "मलयालम",
  MANIPURI: "मणिपुरी",
  MARATHI: "मराठी",
  NEPALI: "नेपाली",
  ODIA: "ओड़िया",
  PUNJABI: "पंजाबी",
  SANSKRIT: "संस्कृत",
  SANTALI: "संथाली",
  SINDHI: "सिंधी",
  TAMIL: "तमिल",
  TELUGU: "तेलुगु",
  URDU: "उर्दू",
};

const HINDI_OFFERING_CATEGORY_LABELS = {
  Service: "सेवा",
  Product: "उत्पाद",
  Knowledge: "ज्ञान",
};

const HINDI_OFFERING_TYPE_LABELS = {
  Training: "प्रशिक्षण",
  "Consulting / Mentoring": "परामर्श / मार्गदर्शन",
  "Financial Support": "वित्तीय सहायता",
  "Market Support": "बाज़ार सहायता",
  "Technology Transfer": "प्रौद्योगिकी हस्तांतरण",
  Machinery: "मशीनरी",
  "Plant Setup": "प्लांट स्थापना",
  "Product/Raw Material Bought": "उत्पाद / कच्चा माल खरीदी",
  "Raw Material Supply": "कच्चा माल आपूर्ति",
  Blogs: "ब्लॉग",
  "Market Reports": "बाज़ार रिपोर्ट",
  "SOP / Manuals": "एसओपी / मैनुअल",
  Videos: "वीडियो",
};
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

function getCookieValue(name) {
  const parts = document.cookie ? document.cookie.split("; ") : [];
  const prefix = `${name}=`;
  for (const part of parts) {
    if (part.indexOf(prefix) === 0) {
      return decodeURIComponent(part.slice(prefix.length));
    }
  }
  return "";
}

function decodeGrameeeSessionTransfer(value) {
  const encoded = normalizeText(value);
  if (!encoded) return null;
  try {
    return JSON.parse(decodeURIComponent(escape(atob(encoded))));
  } catch {
    return null;
  }
}

function decodeJwtPayload(token) {
  const encoded = normalizeText(token);
  if (!encoded) return null;
  const parts = encoded.split(".");
  if (parts.length < 2) return null;
  try {
    const normalized = parts[1].replace(/-/g, "+").replace(/_/g, "/");
    const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=");
    return JSON.parse(atob(padded));
  } catch {
    return null;
  }
}

function captureGrameeeSessionTransfer() {
  const currentUrl = new URL(window.location.href);
  const rawTransfer = currentUrl.searchParams.get(GRAMEEE_SESSION_TRANSFER_PARAM);
  let parsed = null;

  if (rawTransfer) {
    parsed = decodeGrameeeSessionTransfer(rawTransfer);
    currentUrl.searchParams.delete(GRAMEEE_SESSION_TRANSFER_PARAM);
    window.history.replaceState({}, "", currentUrl.toString());
  } else if (window.__grameeeInboundTransferPayload && typeof window.__grameeeInboundTransferPayload === "object") {
    parsed = window.__grameeeInboundTransferPayload;
  } else {
    try {
      const cached = window.sessionStorage.getItem(GRAMEEE_TRANSFER_CACHE_KEY);
      parsed = cached ? JSON.parse(cached) : null;
    } catch {
      parsed = null;
    }
  }

  if (!parsed || typeof parsed !== "object") return null;

  state.grameeeTransfer = parsed && typeof parsed === "object" ? parsed : null;
  try {
    window.sessionStorage.removeItem(GRAMEEE_TRANSFER_CACHE_KEY);
  } catch {}
  return state.grameeeTransfer;
}

function readGrameeeSummary() {
  if (state.grameeeTransfer?.summary && typeof state.grameeeTransfer.summary === "object") {
    const sharedSummary = state.grameeeTransfer.summary;
    return {
      username: normalizeText(sharedSummary.username),
      fullName: normalizeText(sharedSummary.fullName),
      email: normalizeText(sharedSummary.email).toLowerCase(),
      role: normalizeText(sharedSummary.role).toLowerCase(),
      privileges: sharedSummary.privileges && typeof sharedSummary.privileges === "object" ? sharedSummary.privileges : {},
    };
  }

  if (window.grameeeAuth?.getStoredSummary) {
    try {
      const sharedSummary = window.grameeeAuth.getStoredSummary();
      if (sharedSummary) {
        return {
          username: normalizeText(sharedSummary.username),
          fullName: normalizeText(sharedSummary.fullName),
          email: normalizeText(sharedSummary.email).toLowerCase(),
          role: normalizeText(sharedSummary.role).toLowerCase(),
          privileges: sharedSummary.privileges && typeof sharedSummary.privileges === "object" ? sharedSummary.privileges : {},
        };
      }
    } catch {}
  }

  const raw = getCookieValue("grameee_user_summary");
  if (!raw) {
    const tokenPayload = decodeJwtPayload(getCookieValue("grameee_access_token"));
    if (!tokenPayload) return null;
    return {
      username: normalizeText(tokenPayload.user_metadata?.username),
      fullName: normalizeText(tokenPayload.user_metadata?.full_name),
      email: normalizeText(tokenPayload.email).toLowerCase(),
      role: normalizeText(tokenPayload.app_metadata?.grameee_role).toLowerCase(),
      privileges: tokenPayload.app_metadata?.grameee_privileges && typeof tokenPayload.app_metadata.grameee_privileges === "object"
        ? tokenPayload.app_metadata.grameee_privileges
        : {},
    };
  }
  try {
    const parsed = JSON.parse(raw);
    return {
      username: normalizeText(parsed.username),
      fullName: normalizeText(parsed.fullName),
      email: normalizeText(parsed.email).toLowerCase(),
      role: normalizeText(parsed.role).toLowerCase(),
      privileges: parsed.privileges && typeof parsed.privileges === "object" ? parsed.privileges : {},
    };
  } catch {
    return null;
  }
}

async function getGrameeeAccessToken() {
  if (state.grameeeTransfer?.accessToken) {
    return normalizeText(state.grameeeTransfer.accessToken);
  }
  if (window.grameeeAuth?.getAccessToken) {
    return await window.grameeeAuth.getAccessToken();
  }
  return normalizeText(getCookieValue("grameee_access_token"));
}

async function waitForGrameeeAuthBootstrap(timeoutMs = 2500) {
  const startedAt = Date.now();

  while (Date.now() - startedAt < timeoutMs) {
    if (window.grameeeAuth?.getStoredSummary || window.grameeeAuth?.hydrateAuthSession) {
      return true;
    }
    await new Promise((resolve) => window.setTimeout(resolve, 80));
  }

  return false;
}

function hasImmediateGreLoginContext() {
  const summary = readGrameeeSummary();
  const sharedRole = normalizeText(summary?.role).toLowerCase();
  return Boolean(
    getCookieValue("grameee_access_token")
    && ["admin", "moderator", "curator"].includes(sharedRole),
  );
}
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
  const text = normalizeText(message);
  if (!text) return;
  let host = byId("greInlineToastHost");
  if (!host) {
    host = document.createElement("div");
    host.id = "greInlineToastHost";
    host.style.position = "fixed";
    host.style.right = "24px";
    host.style.bottom = "24px";
    host.style.zIndex = "9999";
    host.style.display = "grid";
    host.style.gap = "10px";
    host.style.maxWidth = "420px";
    document.body.appendChild(host);
  }
  const notice = document.createElement("button");
  notice.type = "button";
  notice.textContent = text;
  notice.style.border = "1px solid rgba(32,85,44,0.18)";
  notice.style.background = "#fffdf7";
  notice.style.color = "#23402c";
  notice.style.boxShadow = "0 16px 40px rgba(16, 24, 20, 0.16)";
  notice.style.borderRadius = "16px";
  notice.style.padding = "14px 16px";
  notice.style.textAlign = "left";
  notice.style.cursor = "pointer";
  notice.style.font = "inherit";
  notice.style.lineHeight = "1.45";
  notice.addEventListener("click", () => notice.remove());
  host.appendChild(notice);
  window.setTimeout(() => {
    if (notice.isConnected) notice.remove();
  }, 5000);
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

function normalizeEmailList(value) {
  return ensureList(
    Array.isArray(value)
      ? value
      : String(value || "").split(/[;,]/),
  )
    .map((item) => normalizeText(item).toLowerCase())
    .filter((item) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(item));
}

function normalizeSharedFormLanguage(value) {
  return String(value || "").toLowerCase() === "hi" ? "hi" : "en";
}

function getSharedFormLanguageStorageKey() {
  return `grameee-shared-form-lang:${state.sharedFormMode || state.standalonePublicFormMode || "default"}`;
}

function isHindiSharedFormMode() {
  return normalizeSharedFormLanguage(state.formLanguage) === "hi";
}

function uniqueStrings(list) {
  return [...new Set(list.filter(Boolean))];
}

function canonicalizeLanguageLabel(value) {
  const text = normalizeText(value);
  if (!text) return "";
  const normalized = text.toLowerCase();
  if (["eng", "english"].includes(normalized)) return "ENGLISH";
  if (["hin", "hindi"].includes(normalized)) return "HINDI";
  if (["odia", "oriya", "odiya", "od"].includes(normalized)) return "ODIA";
  return text.toUpperCase();
}

function normalizeCell(value) {
  if (value === null || value === undefined) return null;
  const text = String(value).trim();
  return text.length ? text : null;
}

function stripHtml(html) {
  if (!html) return "";
  return String(html)
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/\s+/g, " ")
    .trim();
}

function splitLooseList(value, separators) {
  if (!value) return [];
  const seps = separators || [",", ";", "\n"];
  let working = String(value);
  for (const sep of seps) {
    working = working.split(sep).join("|");
  }
  return uniqueStrings(
    working.split("|").map((entry) => entry.trim()).filter(Boolean)
  );
}

function splitGeographies(value) {
  if (!value) return [];
  const text = String(value);
  if (text.includes(";") || text.includes("\n") || text.includes("|")) {
    return splitLooseList(text, [";", "\n", "|"]);
  }
  return [text.trim()].filter(Boolean);
}

function normalizeLanguageArray(value) {
  const raw = Array.isArray(value) ? value : splitLooseList(value);
  return uniqueStrings(raw.map(canonicalizeLanguageLabel).filter(Boolean));
}

function buildSearchDocument(parts) {
  return parts
    .flatMap((part) => (Array.isArray(part) ? part : [part]))
    .filter(Boolean)
    .join(" | ")
    .replace(/\s+/g, " ")
    .trim();
}

function dedupeRowsById(rows, key) {
  const map = new Map();
  rows.forEach((row) => {
    const id = String(row[key] ?? "").trim();
    if (id) map.set(id, row);
  });
  return [...map.values()];
}

function canonicalizeLanguageArray(values) {
  return uniq(ensureList(values).map((item) => canonicalizeLanguageLabel(item)).filter(Boolean));
}

function translateLanguageLabel(label) {
  const canonical = canonicalizeLanguageLabel(label);
  if (!isHindiSharedFormMode()) return canonical || normalizeText(label);
  return HINDI_LANGUAGE_LABELS[canonical] || canonical || normalizeText(label);
}

function translateOfferingCategoryLabel(label) {
  const base = normalizeText(label);
  if (!isHindiSharedFormMode()) return base;
  return HINDI_OFFERING_CATEGORY_LABELS[base] || base;
}

function translateOfferingTypeLabel(label) {
  const base = normalizeText(label);
  if (!isHindiSharedFormMode()) return base;
  return HINDI_OFFERING_TYPE_LABELS[base] || base;
}

function syncSessionState(user, token = state.userToken) {
  state.userSession = user || null;
  state.userToken = token || "";
  const hasAdminLikeAccess = ["admin", "moderator"].includes(String(user?.role || "").toLowerCase());
  state.adminToken = hasAdminLikeAccess ? state.userToken : "";
  state.adminSession = hasAdminLikeAccess
    ? { username: user?.username || "admin", email: user?.email || "" }
    : null;

  if (state.userToken) {
    localStorage.setItem("gre-mis-user-token", state.userToken);
    if (hasAdminLikeAccess) localStorage.setItem("gre-mis-admin-token", state.userToken);
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

function isStandalonePublicFormMode(mode = "") {
  if (!state.standalonePublicFormMode) return false;
  if (!mode) return true;
  return state.standalonePublicFormMode === mode;
}

function isGrameeeEmbed() {
  return Boolean(state.grameeeEmbed);
}

function isEmbeddedSharedForm() {
  return isGrameeeEmbed() && isSharedFormMode();
}

function isAllowedGrameeeOrigin(origin) {
  try {
    const url = new URL(origin);
    return url.protocol === "https:" && (url.hostname === "grameee.org" || url.hostname.endsWith(".grameee.org"));
  } catch {
    return false;
  }
}

function normalizeEmbeddedActor(payload) {
  if (!payload || typeof payload !== "object") return null;
  const fullName = normalizeText(payload.fullName || payload.full_name || payload.name || "");
  const username = normalizeText(payload.username || payload.user_name || "");
  return {
    id: normalizeText(payload.id || ""),
    username,
    full_name: fullName,
    first_name: normalizeText(payload.firstName || payload.first_name || fullName || username || ""),
    email: normalizeText(payload.email || "").toLowerCase(),
    phone: normalizeText(payload.phone || ""),
    organization: normalizeText(payload.organization || payload.organisation || "") || "Individual",
  };
}

function getSubmissionActor() {
  return state.embeddedActor || state.userSession || null;
}

function getSubmissionActorDisplayName(actor = getSubmissionActor()) {
  return normalizeText(actor?.username || actor?.full_name || actor?.first_name || "");
}

function postEmbeddedMessage(type, payload = {}) {
  if (!isEmbeddedSharedForm() || window.parent === window) return;
  window.parent.postMessage({ type, payload }, state.embeddedContextOrigin || "*");
}

function requestEmbeddedContext() {
  postEmbeddedMessage("grameee-form-request-context", {
    mode: state.sharedFormMode,
  });
}

function applyEmbeddedContext(payload, origin = "") {
  const actor = normalizeEmbeddedActor(payload?.user || payload);
  state.embeddedContextOrigin = isAllowedGrameeeOrigin(origin) ? origin : state.embeddedContextOrigin;
  state.embeddedActor = actor;
  if (payload?.language) {
    state.formLanguage = normalizeSharedFormLanguage(payload.language);
  }
  renderSubmissionViews();
  applySharedFormLanguage().catch(() => null);
}

function requestEmbeddedLogin(mode = state.sharedFormMode || "need") {
  if (!isEmbeddedSharedForm()) return;
  postEmbeddedMessage("grameee-form-require-login", { mode });
}

function getActiveSharedFormRoot() {
  if (state.sharedFormMode === "solution") return byId("solutionView");
  if (state.sharedFormMode === "need-intake") return byId("need-intakeView");
  return null;
}

function collectSharedFormTranslatableNodes(root) {
  if (!root) return [];
  return [...root.querySelectorAll(SHARED_FORM_TRANSLATABLE_SELECTOR)]
    .filter((node) => !node.closest(".tag-cloud"))
    .filter((node) => !node.closest(".gre-inline-toast-host"))
    .filter((node) => !node.closest("#solutionLanguagesGroup"))
    .filter((node) => !node.closest("#solutionOfferingButtonGrid"))
    .filter((node) => !node.closest("#needOfferingButtonGrid"));
}

function collectSharedFormPlaceholderNodes(root) {
  if (!root) return [];
  return [...root.querySelectorAll("input[placeholder], textarea[placeholder], select option")]
    .filter((node) => !node.closest(".tag-cloud"));
}

function rememberNodeEnglishCopy(node) {
  if ("textContent" in node && !node.dataset.i18nBaseText) {
    node.dataset.i18nBaseText = normalizeText(node.textContent || "");
  }
  if ("placeholder" in node && !node.dataset.i18nBasePlaceholder) {
    node.dataset.i18nBasePlaceholder = normalizeText(node.placeholder || "");
  }
  if (node.tagName === "OPTION" && !node.dataset.i18nBaseText) {
    node.dataset.i18nBaseText = normalizeText(node.textContent || "");
  }
}

function restoreSharedFormEnglish(root) {
  collectSharedFormTranslatableNodes(root).forEach((node) => {
    rememberNodeEnglishCopy(node);
    if (node.dataset.i18nBaseText) {
      node.textContent = node.dataset.i18nBaseText;
    }
  });
  collectSharedFormPlaceholderNodes(root).forEach((node) => {
    rememberNodeEnglishCopy(node);
    if ("placeholder" in node && node.dataset.i18nBasePlaceholder) {
      node.placeholder = node.dataset.i18nBasePlaceholder;
    }
    if (node.tagName === "OPTION" && node.dataset.i18nBaseText) {
      node.textContent = node.dataset.i18nBaseText;
    }
  });
}

function queueSharedFormLanguageRefresh() {
  if (!isSharedFormMode()) return;
  setTimeout(() => {
    applySharedFormLanguage().catch(() => null);
  }, 0);
  if (typeof requestAnimationFrame === "function") {
    requestAnimationFrame(() => applySharedFormLanguage().catch(() => null));
  }
}

async function translateUiStrings(strings, source = "en", target = "hi") {
  const unique = [...new Set(strings.map((entry) => normalizeText(entry)).filter(Boolean))];
  if (!unique.length) return new Map();

  const resultMap = new Map();
  const pending = [];
  const staticMap = SHARED_FORM_STATIC_TRANSLATIONS[target] || {};

  unique.forEach((text) => {
    if (staticMap[text]) {
      resultMap.set(text, staticMap[text]);
      return;
    }
    const cacheKey = `${source}:${target}:${text}`;
    if (state.translationCache.has(cacheKey)) {
      resultMap.set(text, state.translationCache.get(cacheKey));
    } else {
      pending.push(text);
    }
  });

  if (pending.length) {
    try {
      const response = await store.translateTextBatch(pending, source, target);
      const translations = Array.isArray(response?.translations) ? response.translations : [];
      translations.forEach((entry, index) => {
        const original = pending[index];
        const translated = normalizeText(entry?.translatedText || entry?.translated_text || original);
        const cacheKey = `${source}:${target}:${original}`;
        state.translationCache.set(cacheKey, translated);
        resultMap.set(original, translated);
      });
    } catch (error) {
      console.warn("Shared form remote translation unavailable:", error instanceof Error ? error.message : String(error));
    }
    pending.forEach((text) => {
      if (!resultMap.has(text)) {
        resultMap.set(text, staticMap[text] || text);
      }
    });
  }

  return resultMap;
}

async function applySharedFormLanguage() {
  if (!isSharedFormMode()) return;
  const root = getActiveSharedFormRoot();
  if (!root) return;

  restoreSharedFormEnglish(root);
  if (!isHindiSharedFormMode()) return;

  const textNodes = collectSharedFormTranslatableNodes(root);
  const placeholderNodes = collectSharedFormPlaceholderNodes(root);

  textNodes.forEach(rememberNodeEnglishCopy);
  placeholderNodes.forEach(rememberNodeEnglishCopy);

  const textStrings = textNodes.map((node) => node.dataset.i18nBaseText || "");
  const placeholderStrings = placeholderNodes.flatMap((node) => {
    const values = [];
    if (node.dataset.i18nBasePlaceholder) values.push(node.dataset.i18nBasePlaceholder);
    if (node.tagName === "OPTION" && node.dataset.i18nBaseText) values.push(node.dataset.i18nBaseText);
    return values;
  });
  const translations = await translateUiStrings([...textStrings, ...placeholderStrings], "en", "hi");

  textNodes.forEach((node) => {
    const base = node.dataset.i18nBaseText || "";
    if (base) node.textContent = translations.get(base) || base;
  });
  placeholderNodes.forEach((node) => {
    const placeholderBase = node.dataset.i18nBasePlaceholder || "";
    if ("placeholder" in node && placeholderBase) {
      node.placeholder = translations.get(placeholderBase) || placeholderBase;
    }
    if (node.tagName === "OPTION") {
      const optionBase = node.dataset.i18nBaseText || "";
      if (optionBase) node.textContent = translations.get(optionBase) || optionBase;
    }
  });
}

function setSharedFormLanguage(language, { persist = true } = {}) {
  state.formLanguage = normalizeSharedFormLanguage(language);
  if (persist) {
    localStorage.setItem(getSharedFormLanguageStorageKey(), state.formLanguage);
  }
  if (isSharedFormMode()) {
    renderSolutionReferenceInputs();
    renderNeedReferenceInputs();
    renderSolutionGeographyChips();
    renderNeedDeploymentChips();
    ensureSharedFormInteractiveControls();
  }
  queueSharedFormLanguageRefresh();
  applySharedFormLanguage().catch((error) => {
    console.warn("Shared form translation failed:", error instanceof Error ? error.message : String(error));
  });
}

function syncSharedFormLanguageFromBootstrap({ persist = false } = {}) {
  const bootstrapLanguage =
    bootstrapQuery.get("formLang") ||
    localStorage.getItem(getSharedFormLanguageStorageKey()) ||
    state.formLanguage ||
    "en";
  setSharedFormLanguage(bootstrapLanguage, { persist });
}

function isAdminUser() {
  return state.userSession?.role === "admin";
}

function isModeratorUser() {
  return state.userSession?.role === "moderator";
}

function isCuratorUser() {
  return state.userSession?.role === "curator";
}

function hasAdminLikeAccess() {
  return isAdminUser() || isModeratorUser();
}

function hasMisAccessRole() {
  return hasAdminLikeAccess() || isCuratorUser();
}

function isLocalOnlyNeed(need) {
  if (!need) return false;
  const sourceKind = normalizeText(need.source_kind).toLowerCase();
  return sourceKind === "shared_form_submission" || String(need.id || "").startsWith("FORM-");
}

function canAccessOperationsDesk() {
  return hasMisAccessRole();
}

function canSeeCurationDetails() {
  return hasAdminLikeAccess() || isCuratorUser();
}

function canRejectRecords() {
  return isAdminUser();
}

function canDeleteRecords() {
  return isAdminUser();
}

function canManageUsers() {
  return isAdminUser();
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

function pickWorkbookValue(row, aliases) {
  for (const alias of aliases) {
    const text = normalizeText(row[alias]);
    if (text) return text;
  }
  const normalizedAliases = aliases.map((alias) => normalizeText(alias).toLowerCase().replace(/[^a-z0-9]+/g, ""));
  for (const [key, value] of Object.entries(row || {})) {
    const normalizedKey = normalizeText(key).toLowerCase().replace(/[^a-z0-9]+/g, "");
    if (normalizedAliases.includes(normalizedKey)) {
      const text = normalizeText(value);
      if (text) return text;
    }
  }
  return "";
}

function pickWorkbookDate(row, aliases) {
  for (const alias of aliases) {
    const raw = row[alias];
    const parsed = parseWorkbookDate(raw)?.slice(0, 10) || normalizeText(raw);
    if (parsed) return parsed;
  }
  const normalizedAliases = aliases.map((alias) => normalizeText(alias).toLowerCase().replace(/[^a-z0-9]+/g, ""));
  for (const [key, value] of Object.entries(row || {})) {
    const normalizedKey = normalizeText(key).toLowerCase().replace(/[^a-z0-9]+/g, "");
    if (normalizedAliases.includes(normalizedKey)) {
      const parsed = parseWorkbookDate(value)?.slice(0, 10) || normalizeText(value);
      if (parsed) return parsed;
    }
  }
  return "";
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
    funding_mechanism: pickWorkbookValue(row, [
      "Funding Mechanism",
      "Capture Outcome - Funding Mechanism",
      "Outcome Funding Mechanism",
    ]),
    seeker_provider_agreement: pickWorkbookValue(row, [
      "Seeker / Provider Agreement",
      "Seeker/Provider Agreement",
      "Capture Outcome - Seeker / Provider Agreement",
      "Seeker Provider Agreement",
    ]),
    solution_deployment_status: pickWorkbookValue(row, [
      "Solution Deployment Status",
      "Capture Outcome - Solution Deployment Status",
      "Deployment Status",
    ]),
    closure_date: pickWorkbookDate(row, [
      "Closure Date",
      "Capture Outcome - Closure Date",
      "Closed On",
    ]),
    feedback_about_seeker: pickWorkbookValue(row, [
      "Feedback about Seeker",
      "Feedback About Seeker",
      "Seeker Feedback",
    ]),
    feedback_about_provider: pickWorkbookValue(row, [
      "Feedback about Provider",
      "Feedback About Provider",
      "Provider Feedback",
    ]),
  })).filter((row) => row.request_id);
}

async function readFileAsBase64(file) {
  if (!file) return "";
  const buffer = await file.arrayBuffer();
  const bytes = new Uint8Array(buffer);
  let binary = "";
  const chunkSize = 8192;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize));
  }
  return btoa(binary);
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

function uniqBy(items, keyFn) {
  const seen = new Set();
  const results = [];
  ensureList(items).forEach((item) => {
    const key = normalizeText(keyFn(item));
    if (!key || seen.has(key)) return;
    seen.add(key);
    results.push(item);
  });
  return results;
}

function getEffectiveNeedThematicArea(need) {
  return normalizeText(need?.override_thematic_area || need?.ai_thematic_area || need?.submitted_thematic_area);
}

function normalizeNeedCategoryKind(value) {
  const normalized = normalizeText(value).toLowerCase();
  if (!normalized) return "";
  if (normalized.includes("service")) return "service";
  if (normalized.includes("product")) return "product";
  if (normalized.includes("knowledge")) return "knowledge";
  return normalized;
}

function getEffectiveNeedApplicationArea(need) {
  return normalizeText(need?.override_application_area || need?.ai_application_area);
}

function getEffectiveNeedKind(need) {
  return normalizeText(need?.override_need_kind || need?.ai_need_kind || normalizeNeedCategoryKind(need?.submitted_offering_category));
}

function getEffectiveNeedServiceKind(need) {
  const submittedCategory = normalizeNeedCategoryKind(need?.submitted_offering_category);
  return normalizeText(need?.override_service_kind || need?.ai_service_kind || (submittedCategory === "service" ? need?.submitted_offering_type : ""));
}

function getEffectiveNeedKeywords(need) {
  return uniq([
    ...parseArray(need?.submitted_keywords).map((item) => item.toLowerCase()),
    ...parseArray(need?.keywords).map((item) => item.toLowerCase()),
    ...parseArray(need?.override_keywords).map((item) => item.toLowerCase()),
    ...parseArray(need?.ai_keywords).map((item) => item.toLowerCase()),
  ]);
}

function getEffectiveNeedDeploymentLocations(need) {
  const direct = uniq([
    ...parseArray(need?.deployment_locations),
    ...parseArray(need?.submitted_deployment_locations),
  ]);
  if (direct.length) return direct;
  const district = normalizeText(need?.district);
  const stateName = normalizeText(need?.state);
  if (district && stateName) return [`${district}, ${stateName}`];
  if (stateName) return [stateName];
  return [];
}

function getEffectiveSubmittedNeedCategory(need) {
  const direct = normalizeText(need?.submitted_offering_category);
  if (direct) return direct;
  const kind = getEffectiveNeedKind(need).toLowerCase();
  if (kind === "service") return "Service offerings";
  if (kind === "product") return "Product offerings";
  if (kind === "knowledge") return "Knowledge offerings";
  return "Service offerings";
}

function getEffectiveSubmittedNeedType(need) {
  return normalizeText(need?.submitted_offering_type || need?.submitted_need_type || getEffectiveNeedServiceKind(need));
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

function normalizeLooseId(value) {
  return String(value || "").trim().replace(/\.0$/i, "");
}

function getNeedCurationAgeDays(need) {
  const importedAge = parseNumber(need?.curation_age_days, 0);
  if (importedAge > 0) return importedAge;
  const candidateDates = [
    need?.requested_on,
    need?.created_at,
    need?.updated_at,
  ]
    .map((value) => {
      const text = normalizeText(value);
      if (!text) return null;
      const parsed = new Date(text);
      return Number.isNaN(parsed.getTime()) ? null : parsed;
    })
    .filter(Boolean);
  const anchorDate = candidateDates[0];
  if (!anchorDate) return 0;
  return Math.max(0, Math.floor((Date.now() - anchorDate.getTime()) / 86400000));
}

function extractSharedSolutionHints(value) {
  const text = normalizeText(value);
  if (!text) {
    return {
      phrases: [],
      ids: { offeringIds: [], solutionIds: [], traderIds: [] },
      exactShareCount: 0,
      namedShares: [],
    };
  }
  const urls = extractUrls(text);
  const ids = { offeringIds: [], solutionIds: [], traderIds: [] };
  const exactSharedKeys = [];
  const namedShares = [];

  urls.forEach((url) => {
    const offeringId = url.match(/[?&]productSkuId=(\d+)/i)?.[1];
    const solutionId = url.match(/[?&]solutionId=(\d+)/i)?.[1];
    const traderId = url.match(/[?&]traderId=(\d+)/i)?.[1];
    if (offeringId) ids.offeringIds.push(offeringId);
    if (solutionId) ids.solutionIds.push(solutionId);
    if (traderId) ids.traderIds.push(traderId);
    if (solutionId) exactSharedKeys.push(`solution:${solutionId}`);
    else if (offeringId) exactSharedKeys.push(`offering:${offeringId}`);
  });

  const phrases = [];
  text.split(/\n+/).forEach((line) => {
    const cleaned = line.replace(/https?:\/\/[^\s)]+/gi, "").trim();
    const parts = cleaned.split(":").map((item) => normalizeText(item));
    if (parts.length >= 2 && parts[1]) {
      phrases.push(parts[1].toLowerCase());
    }
  });

  let solutionsSharedMatch = text.match(/solutions shared\n((?:\d+\. .*\n?)*)/i);
  if (!solutionsSharedMatch) {
    solutionsSharedMatch = text.match(/solutions shared\s*:\s*([\s\S]*?)(?:solution links shared with seeker|$)/i);
  }
  const solutionsSharedText = normalizeText(solutionsSharedMatch?.[1] || "");
  if (solutionsSharedText) {
    const rawEntries = [];
    let cursor = 0;
    [...solutionsSharedText.matchAll(/https?:\/\/[^\s)]+/gi)].forEach((match) => {
      const url = normalizeText(match[0]);
      const index = Number(match.index || 0);
      const prefix = solutionsSharedText.slice(cursor, index);
      const entry = normalizeText(`${prefix}${url}`);
      if (entry) rawEntries.push(entry);
      cursor = index + url.length;
    });
    if (!rawEntries.length) {
      rawEntries.push(...solutionsSharedText.split(/\s*;\s*/).map((item) => normalizeText(item)).filter(Boolean));
    }
    rawEntries.forEach((entry) => {
        const entryUrls = extractUrls(entry);
        entryUrls.forEach((url) => {
          const offeringId = normalizeLooseId(url.match(/[?&]productSkuId=(\d+)/i)?.[1]);
          const solutionId = normalizeLooseId(url.match(/[?&]solutionId=(\d+)/i)?.[1]);
          const traderId = normalizeLooseId(url.match(/[?&]traderId=(\d+)/i)?.[1]);
          if (offeringId) ids.offeringIds.push(offeringId);
          if (solutionId) ids.solutionIds.push(solutionId);
          if (traderId) ids.traderIds.push(traderId);
          if (offeringId) exactSharedKeys.push(`offering:${offeringId}`);
          else if (solutionId) exactSharedKeys.push(`solution:${solutionId}`);
        });

        const cleanedEntry = normalizeText(stripUrls(entry))
          .replace(/^[,\s]+/, "")
          .replace(/\s:\s*$/g, "")
          .replace(/^\d+\.\s*/, "");
        const providerSplitIndex = cleanedEntry.lastIndexOf(" : ");
        if (providerSplitIndex > 0) {
          const offeringName = normalizeText(cleanedEntry.slice(0, providerSplitIndex));
          const providerName = normalizeText(cleanedEntry.slice(providerSplitIndex + 3)).replace(/\s:\s*$/g, "");
          namedShares.push({
            offeringName: offeringName.toLowerCase(),
            providerName: providerName.toLowerCase(),
            url: entryUrls[0] || "",
            offeringId: normalizeLooseId(entryUrls.map((url) => url.match(/[?&]productSkuId=(\d+)/i)?.[1]).find(Boolean)),
            solutionId: normalizeLooseId(entryUrls.map((url) => url.match(/[?&]solutionId=(\d+)/i)?.[1]).find(Boolean)),
            traderId: normalizeLooseId(entryUrls.map((url) => url.match(/[?&]traderId=(\d+)/i)?.[1]).find(Boolean)),
          });
        } else {
          namedShares.push({
            offeringName: cleanedEntry.toLowerCase(),
            providerName: "",
            url: entryUrls[0] || "",
            offeringId: normalizeLooseId(entryUrls.map((url) => url.match(/[?&]productSkuId=(\d+)/i)?.[1]).find(Boolean)),
            solutionId: normalizeLooseId(entryUrls.map((url) => url.match(/[?&]solutionId=(\d+)/i)?.[1]).find(Boolean)),
            traderId: normalizeLooseId(entryUrls.map((url) => url.match(/[?&]traderId=(\d+)/i)?.[1]).find(Boolean)),
          });
        }
    });
  }

  return {
    phrases: uniq(phrases),
    ids: {
      offeringIds: uniq(ids.offeringIds),
      solutionIds: uniq(ids.solutionIds),
      traderIds: uniq(ids.traderIds),
    },
    exactShareCount: uniq(ids.offeringIds).length || ensureList(namedShares).length || uniq(exactSharedKeys).length,
    namedShares,
  };
}

function getNeedSharedSolutionDisplayEntries(need) {
  const hints = extractSharedSolutionHints(need?.curation_notes);
  const explicitEntries = ensureList(hints.namedShares).map((share) => ({
    offeringName: normalizeText(share?.offeringName),
    providerName: normalizeText(share?.providerName),
    url: normalizeText(share?.url),
    inferred: false,
  }));
  const inferredEntries = ensureList(need?._derived_inferred_shared_entries).map((entry) => ({
    offeringName: normalizeText(entry?.offeringName),
    providerName: normalizeText(entry?.providerName),
    inferred: true,
  }));
  const seen = new Set();
  return [...explicitEntries, ...inferredEntries].filter((entry) => {
    const key = `${entry.offeringName.toLowerCase()}::${entry.providerName.toLowerCase()}`;
    if (!entry.offeringName || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function applySharedSolutionAnnotations(need, matches) {
  if (!need || !Array.isArray(matches) || !matches.length) return;
  const hints = extractSharedSolutionHints(need.curation_notes);
  const targetCount = Math.max(
    parseNumber(need?.solutions_shared_count, 0),
    ensureList(hints.namedShares).length,
    parseNumber(need?._derived_shared_match_count, 0),
  );
  const sharedReason = "Already shared with seeker";
  const explicitProviderNames = new Set(
    ensureList(hints.namedShares)
      .map((share) => normalizeText(share?.providerName).toLowerCase())
      .filter(Boolean),
  );
  const explicitTraderIds = new Set(ensureList(hints.ids?.traderIds).map((item) => normalizeLooseId(item)).filter(Boolean));
  const explicitOfferingIds = new Set(ensureList(hints.ids?.offeringIds).map((item) => normalizeLooseId(item)).filter(Boolean));
  const explicitSolutionIds = new Set(ensureList(hints.ids?.solutionIds).map((item) => normalizeLooseId(item)).filter(Boolean));

  const sharedMatches = matches.filter((match) =>
    ensureList(match.matchReasons).some((reason) => normalizeText(reason).toLowerCase() === sharedReason.toLowerCase()),
  );
  if (sharedMatches.length >= targetCount) {
    need._derived_shared_match_count = sharedMatches.length;
    need._derived_inferred_shared_entries = [];
    return;
  }

  const inferredEntries = [];
  matches
    .filter((match) => !ensureList(match.matchReasons).some((reason) => normalizeText(reason).toLowerCase() === sharedReason.toLowerCase()))
    .filter((match) => {
      const offeringId = normalizeLooseId(match.offering_id);
      const solutionId = normalizeLooseId(match.solution_id);
      const traderId = normalizeLooseId(match.trader_id || match.trader?.trader_id);
      const providerName = normalizeText(match.trader?.organisation_name || match.trader?.trader_name).toLowerCase();
      return (
        (offeringId && explicitOfferingIds.has(offeringId)) ||
        (solutionId && explicitSolutionIds.has(solutionId)) ||
        (traderId && explicitTraderIds.has(traderId)) ||
        (providerName && explicitProviderNames.has(providerName))
      );
    })
    .sort((a, b) => parseNumber(b.matchScore, 0) - parseNumber(a.matchScore, 0))
    .some((match) => {
      if (sharedMatches.length + inferredEntries.length >= targetCount) return true;
      match.matchReasons = uniq([...(match.matchReasons || []), sharedReason, "GRE shared count alignment"]);
      inferredEntries.push({
        offeringName: match.offering_name || match.solution?.solution_name || "",
        providerName: match.trader?.organisation_name || match.trader?.trader_name || "",
      });
      return false;
    });

  need._derived_shared_match_count = sharedMatches.length + inferredEntries.length;
  need._derived_inferred_shared_entries = inferredEntries;
}

function getNeedSharedSolutionCount(need) {
  const importedCount = parseNumber(need?.solutions_shared_count, 0);
  const sharedHints = extractSharedSolutionHints(need?.curation_notes);
  const derivedCount = Math.max(
    sharedHints.exactShareCount,
    ensureList(sharedHints.namedShares).length,
    getNeedSharedSolutionDisplayEntries(need).length,
  );
  const matchCount = parseNumber(need?._derived_shared_match_count, 0);
  return Math.max(importedCount, derivedCount, matchCount);
}

function hasSuggestedQuestionSection(notes) {
  return /suggested questions to seeker/i.test(normalizeText(notes));
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

function normalizeNeedThemeLabel(value) {
  const raw = normalizeText(value)
    .replace(/\s+/g, " ")
    .replace(/[_-]+/g, " ")
    .trim();
  if (!raw) return "";
  return raw
    .split(" ")
    .map((part) => part.length <= 3 && part === part.toUpperCase() ? part : `${part.charAt(0).toUpperCase()}${part.slice(1).toLowerCase()}`)
    .join(" ");
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
  [
    need?.override_thematic_area,
    need?.submitted_thematic_area,
    need?.ai_thematic_area,
    need?.thematic_area,
    need?.category,
    ...parseArray(need?.rule_thematic_hints),
  ].forEach((item) => {
    const label = normalizeNeedThemeLabel(item);
    if (label && !GENERIC_THEMATIC_TERMS.has(label.toLowerCase())) themes.push(label);
  });
  parseArray(need?.curated_need).forEach((item) => {
    const parts = extractCategoryParts(item);
    const label = normalizeNeedThemeLabel(parts.thematic);
    if (label && !GENERIC_THEMATIC_TERMS.has(label.toLowerCase())) themes.push(label);
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
  return getNeedThemeSignals(need).map((item) => item.toLowerCase()).includes(normalized) || curated.some((item) => item.toLowerCase() === normalized);
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
  const submittedKeywords = uniq(parseArray(need.submitted_keywords).map((item) => item.toLowerCase()).filter(Boolean));
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
  const priorityKeywordPhrases = uniq([
    ...submittedKeywords.filter((item) => item.includes(" ")),
    ...aiKeywords.filter((item) => item.includes(" ")),
  ]);
  const problemTokens = uniq([
    ...submittedKeywords.flatMap((item) => tokenizeText(item, 4)),
    ...aiKeywords.flatMap((item) => tokenizeText(item, 4)),
    ...ruleKeywords,
    ...tokenizeText(need.problem_statement, 7),
    ...tokenizeText([
      need.solution_deployment_status,
      need.feedback_about_seeker,
      need.feedback_about_provider,
      need.funding_mechanism,
      need.seeker_provider_agreement,
    ].filter(Boolean).join(" "), 5),
  ]);
  const notesTokens = uniq([
    ...ruleKeywords,
    ...tokenizeText(need.curation_notes, 5),
    ...tokenizeText([
      need.solution_deployment_status,
      need.feedback_about_seeker,
      need.feedback_about_provider,
    ].filter(Boolean).join(" "), 5),
  ]);
  const geographyTokens = tokenizeText(`${need.state || ""} ${need.district || ""}`, 3);
  const serviceTokens = serviceTerms.flatMap((item) => tokenizeText(item, 3));
  const sharedSolutionHints = extractSharedSolutionHints(need.curation_notes);
  const problemPhrases = extractProblemPhrases(need.problem_statement);
  const explicitThemeTokens = uniq([
    ...submittedKeywords.flatMap((item) => tokenizeText(item, 4)),
    ...aiKeywords.flatMap((item) => tokenizeText(item, 4)),
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
    ...submittedKeywords,
    ...aiSignals,
    ...(categoryThematicAreas.length ? [] : problemPhrases.slice(0, 8)),
  ]);
  const categoryTokens = thematicAreas.flatMap((item) => tokenizeText(item, 3));
  const primaryTerms = uniq([
    ...submittedKeywords,
    ...categoryTokens,
    ...serviceTokens,
    ...problemTokens.slice(0, 10),
    ...notesTokens.slice(0, 4),
  ]);
  const phrases = uniq(
    priorityKeywordPhrases
      .concat(thematicAreas)
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
    priorityKeywordPhrases,
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
    .sort((a, b) => getNeedCurationAgeDays(b) - getNeedCurationAgeDays(a));
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
  return [...needs].sort((a, b) => getNeedCurationAgeDays(b) - getNeedCurationAgeDays(a));
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
  resetManualSolutionSearch();
  switchView("operations");
  renderQueue();
  renderNeedDetail();
  renderWorkbench();
  renderManualSolutionSearch();
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
  if (metricId === "stuck") return getNeedCurationAgeDays(need) >= 7;
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
  const themeLabel = getNeedThemeSignals(need)[0] || normalizeText(need.ai_thematic_area || need.override_thematic_area || "");
  return `
    <article class="stack-card">
      <div class="status-row">
        <span class="status-pill ${badgeTone(need.status)}">${esc(need.status)}</span>
        <span class="status-pill ${badgeTone(need.internal_status)}">${esc(need.internal_status)}</span>
        <span class="status-pill info">${esc(getNeedCurationAgeDays(need))} days</span>
        ${themeLabel ? `<span class="status-pill good">${esc(themeLabel)}</span>` : ""}
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
    return getNeedCurationAgeDays(rightNeed) - getNeedCurationAgeDays(leftNeed);
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
    emptyText: activeMetrics.includes("admin_queue") && !hasAdminLikeAccess()
      ? "Sign in as admin or moderator to inspect pending approvals and curator requests."
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
    profile.sharedSolutionHints.ids.offeringIds.includes(normalizeLooseId(offering.offering_id)) ||
    profile.sharedSolutionHints.ids.solutionIds.includes(normalizeLooseId(offering.solution_id));
  const matchedSharedName = ensureList(profile.sharedSolutionHints.namedShares).some((share) => {
    const shareOffering = normalizeText(share?.offeringName).toLowerCase();
    const shareProvider = normalizeText(share?.providerName).toLowerCase();
    if (!shareOffering) return false;
    const offeringNameMatch =
      name === shareOffering ||
      solutionName === shareOffering ||
      name.includes(shareOffering) ||
      solutionName.includes(shareOffering) ||
      shareOffering.includes(name) ||
      shareOffering.includes(solutionName);
    if (!offeringNameMatch) return false;
    if (!shareProvider) return true;
    const providerName = normalizeText(offering.trader?.organisation_name || offering.trader?.trader_name).toLowerCase();
    return providerName === shareProvider || providerName.includes(shareProvider) || shareProvider.includes(providerName);
  });

  if (matchedSharedLink || matchedSharedName) {
    thematicMatched = true;
    serviceMatched = true;
    score += 240;
    reasons.push("Already shared with seeker");
  }

  profile.priorityKeywordPhrases.forEach((phrase) => {
    if (
      tags.some((tag) => tag.includes(phrase)) ||
      name.includes(phrase) ||
      solutionName.includes(phrase) ||
      primaryApplication.includes(phrase) ||
      primaryValuechain.includes(phrase) ||
      applications.some((item) => item.includes(phrase)) ||
      valuechains.some((item) => item.includes(phrase))
    ) {
      thematicMatched = true;
      primaryThematicMatched = true;
      score += 20;
      reasons.push(`Keyword:${phrase}`);
    } else if (about.includes(phrase) || aiOfferingTheme.includes(phrase) || aiOfferingApplication.includes(phrase) || aiOfferingKeywords.some((item) => item.includes(phrase))) {
      thematicMatched = true;
      score += 10;
      reasons.push(`Keyword:${phrase}`);
    }
  });

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

    const [curators, traders, offerings, options, needs, updates] = await Promise.allSettled([
      client.from("gre_mis_curators").select("id, user_id, display_name, email, gre_sync_status").eq("is_active", true).order("display_name"),
      client.from("traders").select("trader_id,trader_name,organisation_name,email,website,association_status").order("organisation_name"),
      client.from("offerings").select("primary_valuechain,primary_application,valuechains,applications,tags,languages,geographies").range(0, 19999),
      client.from("gre_mis_options").select("id, option_type, label, sort_order").eq("is_active", true).order("sort_order"),
      client.from("gre_mis_needs").select("*").eq("approval_status", "approved").order("requested_on", { ascending: false }),
      client.from("gre_mis_need_updates").select("*").order("created_at", { ascending: false }),
    ]);

      const curatorsResult = curators.status === "fulfilled" && !curators.value.error ? ensureList(curators.value.data) : [];
      const tradersResult = traders.status === "fulfilled" && !traders.value.error ? ensureList(traders.value.data) : [];
      const offeringsResult = offerings.status === "fulfilled" && !offerings.value.error ? ensureList(offerings.value.data) : [];
      const optionsResult = options.status === "fulfilled" && !options.value.error ? ensureList(options.value.data) : [];
      const needsResult = needs.status === "fulfilled" && !needs.value.error ? ensureList(needs.value.data) : [];
      const updatesResult = updates.status === "fulfilled" && !updates.value.error ? ensureList(updates.value.data) : [];

      state.data.curators = curatorsResult.length ? curatorsResult : FALLBACK_CURATORS;
      state.data.traders = tradersResult;
      state.data.offeringMaster = buildOfferingMasterData(offeringsResult);
      state.data.options = optionsResult;
      state.data.needs = needsResult.map((need) => ({
        ...need,
        curated_need: parseArray(need.curated_need),
      }));
    state.data.needUpdates = updatesResult;
    state.matchCache.clear();
  }

  async callAdmin(action, body = {}, requireAdmin = false) {
    const performRequest = async () => {
      const grameeeAccessToken = await getGrameeeAccessToken().catch(() => "");
      const grameeeUserSummary = readGrameeeSummary();
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
          grameeeAccessToken: grameeeAccessToken || undefined,
          grameeeUserSummary: grameeeUserSummary || undefined,
        }),
      });
      const rawText = await response.text().catch(() => "");
      let data = {};
      try {
        data = rawText ? JSON.parse(rawText) : {};
      } catch {}
      if (!response.ok) throw new Error(data.error || rawText || `Request failed (${response.status}).`);
      return data;
    };
    const isRetryableAuthError = (error) => {
      const message = normalizeText(error instanceof Error ? error.message : error).toLowerCase();
      return [
        "login required",
        "admin or moderator login required",
        "access is required",
        "could not be verified",
        "session",
      ].some((fragment) => message.includes(fragment));
    };
    try {
      const data = await performRequest();
      if (requireAdmin && !hasAdminLikeAccess()) throw new Error("Admin or moderator login required.");
      return data;
    } catch (error) {
      if (!requireAdmin || isSharedFormMode() || !isRetryableAuthError(error)) throw error;
      const bridged = await this.bridgeGrameeeSession().catch(() => false);
      if (!bridged) throw error;
      const data = await performRequest();
      if (requireAdmin && !hasAdminLikeAccess()) throw new Error("Admin or moderator login required.");
      return data;
    }
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

  async bridgeGrameeeSession() {
    const hydratedSummary = await window.grameeeAuth?.hydrateAuthSession?.().catch(() => null);
    const accessToken = await getGrameeeAccessToken();
    const summary = hydratedSummary || readGrameeeSummary();
    const sharedRole = normalizeText(summary?.role).toLowerCase();

    if (!accessToken) {
      state.grameeeBridgeMessage = "";
      return false;
    }

    if (!["admin", "moderator", "curator"].includes(sharedRole)) {
      state.grameeeBridgeMessage = summary
        ? "This GramEEE login does not yet have GRE access. Please ask an admin to assign admin, moderator, or curator access."
        : "";
      return false;
    }

    try {
      const data = await this.callAdmin("bridgeGrameeeSession", {
        grameeeAccessToken: accessToken,
      });
      syncSessionState(data.user, data.token);
      state.grameeeBridgeMessage = "";
      return true;
    } catch (error) {
      syncSessionState(null, "");
      state.grameeeBridgeMessage = error instanceof Error
        ? error.message
        : "The GRE page could not use your GramEEE login right now.";
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
    if (!hasAdminLikeAccess()) {
      state.data.pendingNeeds = [];
      state.data.pendingUpdates = [];
      state.data.pendingFormSubmissions = [];
      state.data.aiReviewNeeds = [];
      state.data.impactAuditLogs = { emailLogs: [], viewLogs: [] };
      state.data.users = [];
        state.data.localSolutions = [];
        state.data.localNeeds = [];
      state.mailTemplates = {
        providerIntroTemplate: DEFAULT_PROVIDER_INTRO_TEMPLATE,
        curatorForwardTemplate: DEFAULT_CURATOR_FORWARD_TEMPLATE,
        solutionSeekerTemplate: DEFAULT_SOLUTION_SEEKER_TEMPLATE,
        needSeekerTemplate: DEFAULT_NEED_SEEKER_TEMPLATE,
        inboundAutoSyncEnabled: true,
        lshContactEmails: ["subekkumar@pradan.net"],
        lshHelpCcEmails: ["help@greenruraleconomy.in"],
        lshRequestSupportTemplate: DEFAULT_LSH_REQUEST_SUPPORT_TEMPLATE,
        lshEmailProviderTemplate: DEFAULT_LSH_EMAIL_PROVIDER_TEMPLATE,
      };
        return;
    }
    const data = await this.callAdmin("adminSnapshot", {}, true);
    state.data.pendingNeeds = ensureList(data.pendingNeeds);
      state.data.pendingUpdates = ensureList(data.pendingUpdates);
      state.data.pendingFormSubmissions = ensureList(data.pendingFormSubmissions);
      state.data.aiReviewNeeds = ensureList(data.aiReviewNeeds);
      state.data.impactAuditLogs = {
        emailLogs: ensureList(data.impactAuditLogs?.emailLogs),
        viewLogs: ensureList(data.impactAuditLogs?.viewLogs),
      };
      state.data.users = uniqBy(ensureList(data.users), (item) => String(item?.id || item?.email || item?.username || ""));
      state.data.localSolutions = uniqBy(ensureList(data.localSolutions), (item) => String(item?.offering_id || ""));
      state.data.localNeeds = uniqBy(ensureList(data.localNeeds), (item) => {
        const idKey = String(item?.id || "");
        if (idKey) return `id:${idKey}`;
        return [
          normalizeText(item?.organization_name),
          normalizeText(item?.problem_statement),
          normalizeText(item?.contact_person),
          normalizeText(item?.seeker_email),
          normalizeText(item?.submitted_offering_category),
          normalizeText(item?.submitted_offering_type),
        ].join("|");
      });
    state.mailTemplates = {
      providerIntroTemplate: normalizeText(data.mailTemplates?.providerIntroTemplate) || DEFAULT_PROVIDER_INTRO_TEMPLATE,
      curatorForwardTemplate: normalizeText(data.mailTemplates?.curatorForwardTemplate) || DEFAULT_CURATOR_FORWARD_TEMPLATE,
      solutionSeekerTemplate: normalizeText(data.mailTemplates?.solutionSeekerTemplate) || DEFAULT_SOLUTION_SEEKER_TEMPLATE,
      needSeekerTemplate: normalizeText(data.mailTemplates?.needSeekerTemplate) || DEFAULT_NEED_SEEKER_TEMPLATE,
      inboundAutoSyncEnabled: typeof data.mailTemplates?.inboundAutoSyncEnabled === "boolean"
        ? data.mailTemplates.inboundAutoSyncEnabled
        : true,
      lshContactEmails: ensureList(data.mailTemplates?.lshContactEmails).map((value) => normalizeText(value)).filter(Boolean),
      lshHelpCcEmails: ensureList(data.mailTemplates?.lshHelpCcEmails).map((value) => normalizeText(value)).filter(Boolean),
      lshRequestSupportTemplate: normalizeText(data.mailTemplates?.lshRequestSupportTemplate) || DEFAULT_LSH_REQUEST_SUPPORT_TEMPLATE,
      lshEmailProviderTemplate: normalizeText(data.mailTemplates?.lshEmailProviderTemplate) || DEFAULT_LSH_EMAIL_PROVIDER_TEMPLATE,
    };
  }

  async loadAdminDataSyncSnapshot() {
    if (!hasAdminLikeAccess()) return;
    const data = await this.callAdmin("adminDataSyncSnapshot", {}, true);
    state.data.aiReviewNeeds = ensureList(data.aiReviewNeeds);
    if (typeof data.inboundAutoSyncEnabled === "boolean") {
      state.mailTemplates.inboundAutoSyncEnabled = data.inboundAutoSyncEnabled;
    }
  }

  async loadAdminMailImpactSnapshot() {
    if (!hasAdminLikeAccess()) return;
    const data = await this.callAdmin("adminMailImpactSnapshot", {}, true);
    state.data.impactAuditLogs = {
      emailLogs: ensureList(data.impactAuditLogs?.emailLogs),
      viewLogs: ensureList(data.impactAuditLogs?.viewLogs),
    };
    state.mailTemplates = {
      ...state.mailTemplates,
      providerIntroTemplate: normalizeText(data.mailTemplates?.providerIntroTemplate) || DEFAULT_PROVIDER_INTRO_TEMPLATE,
      curatorForwardTemplate: normalizeText(data.mailTemplates?.curatorForwardTemplate) || DEFAULT_CURATOR_FORWARD_TEMPLATE,
      solutionSeekerTemplate: normalizeText(data.mailTemplates?.solutionSeekerTemplate) || DEFAULT_SOLUTION_SEEKER_TEMPLATE,
      needSeekerTemplate: normalizeText(data.mailTemplates?.needSeekerTemplate) || DEFAULT_NEED_SEEKER_TEMPLATE,
      inboundAutoSyncEnabled: typeof data.mailTemplates?.inboundAutoSyncEnabled === "boolean"
        ? data.mailTemplates.inboundAutoSyncEnabled
        : state.mailTemplates.inboundAutoSyncEnabled,
      lshContactEmails: ensureList(data.mailTemplates?.lshContactEmails).map((value) => normalizeText(value)).filter(Boolean),
      lshHelpCcEmails: ensureList(data.mailTemplates?.lshHelpCcEmails).map((value) => normalizeText(value)).filter(Boolean),
      lshRequestSupportTemplate: normalizeText(data.mailTemplates?.lshRequestSupportTemplate) || DEFAULT_LSH_REQUEST_SUPPORT_TEMPLATE,
      lshEmailProviderTemplate: normalizeText(data.mailTemplates?.lshEmailProviderTemplate) || DEFAULT_LSH_EMAIL_PROVIDER_TEMPLATE,
    };
  }

  async loadAdminUsersSnapshot() {
    if (!hasAdminLikeAccess()) return;
    const data = await this.callAdmin("adminUsersSnapshot", {}, true);
    state.data.users = uniqBy(ensureList(data.users), (item) => String(item?.id || item?.email || item?.username || ""));
  }

  async loadAdminLocalSolutionsSnapshot() {
    if (!hasAdminLikeAccess()) return;
    const data = await this.callAdmin("adminLocalSolutionsSnapshot", {}, true);
    state.data.localSolutions = uniqBy(ensureList(data.localSolutions), (item) => String(item?.offering_id || ""));
  }

  async loadAdminLocalNeedsSnapshot() {
    if (!hasAdminLikeAccess()) return;
    const data = await this.callAdmin("adminLocalNeedsSnapshot", {}, true);
    state.data.localNeeds = uniqBy(ensureList(data.localNeeds), (item) => {
      const idKey = String(item?.id || "");
      if (idKey) return `id:${idKey}`;
      return [
        normalizeText(item?.organization_name),
        normalizeText(item?.problem_statement),
        normalizeText(item?.contact_person),
        normalizeText(item?.seeker_email),
        normalizeText(item?.submitted_offering_category),
        normalizeText(item?.submitted_offering_type),
      ].join("|");
    });
  }

  async loadAdminApprovalsSnapshot() {
    if (!hasAdminLikeAccess()) return;
    const data = await this.callAdmin("adminApprovalsSnapshot", {}, true);
    state.data.pendingNeeds = ensureList(data.pendingNeeds);
    state.data.pendingUpdates = ensureList(data.pendingUpdates);
    state.data.pendingFormSubmissions = ensureList(data.pendingFormSubmissions);
  }

  async loadAdminDeskTabSnapshot(tab = state.adminDeskTab || "data-sync") {
    if (!hasAdminLikeAccess()) return;
    if (tab === "data-sync") return this.loadAdminDataSyncSnapshot();
    if (tab === "mail-impact") return this.loadAdminMailImpactSnapshot();
    if (tab === "users") return this.loadAdminUsersSnapshot();
    if (tab === "solutions") return this.loadAdminLocalSolutionsSnapshot();
    if (tab === "needs") return this.loadAdminLocalNeedsSnapshot();
    if (tab === "approvals") return this.loadAdminApprovalsSnapshot();
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

  async getLocalSolutionDetail(offeringId) {
    return this.callAdmin("getLocalSolutionDetail", { offeringId }, true);
  }

  async deleteLocalSolution(offeringId) {
    return this.callAdmin("deleteLocalSolution", { offeringId }, true);
  }

  async updateLocalNeed(needId, payload) {
    return this.callAdmin("updateLocalNeed", { needId, payload }, true);
  }

  async getLocalNeedDetail(needId) {
    return this.callAdmin("getLocalNeedDetail", { needId }, true);
  }

  async deleteLocalNeed(needId) {
    return this.callAdmin("deleteLocalNeed", { needId }, true);
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

  async sendProviderIntro(needId, providerEmail, providerName = "", offeringId = "", mailDraft = null) {
    return this.callAdmin("sendProviderIntro", { needId, providerEmail, providerName, offeringId, mailDraft });
  }

  async sendSolutionSeekerIntro(needId, providerEmail, providerName = "", offeringId = "", mailDraft = null) {
    return this.callAdmin("sendSolutionSeekerIntro", { needId, providerEmail, providerName, offeringId, mailDraft });
  }

  async rejectNeedSuggestedQuestions(needId) {
    return this.callAdmin("rejectNeedSuggestedQuestions", { needId });
  }

  async saveMailTemplates(
    providerIntroTemplate,
    curatorForwardTemplate,
    solutionSeekerTemplate,
    needSeekerTemplate,
    inboundAutoSyncEnabled,
    lshContactEmails,
    lshHelpCcEmails,
    lshRequestSupportTemplate,
    lshEmailProviderTemplate,
  ) {
    return this.callAdmin("saveMailTemplates", {
      providerIntroTemplate,
      curatorForwardTemplate,
      solutionSeekerTemplate,
      needSeekerTemplate,
      inboundAutoSyncEnabled,
      lshContactEmails,
      lshHelpCcEmails,
      lshRequestSupportTemplate,
      lshEmailProviderTemplate,
    }, true);
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

  async suggestNeedTags(payload) {
    return this.callAdmin("suggestNeedTags", { payload }, false);
  }

  async translateTextBatch(texts, source = "en", target = "hi") {
    return this.callAdmin("translateTextBatch", { texts, source, target }, false);
  }

  async generateNeedSuggestedQuestions(needId) {
    return this.callAdmin("generateNeedSuggestedQuestions", { needId }, true);
  }

  async saveNeedSuggestedQuestions(needId, questions, sourceLabel = "puter") {
    return this.callAdmin("saveNeedSuggestedQuestions", { needId, questions, sourceLabel }, true);
  }

  async searchLgdGeographies(query) {
    const search = normalizeText(query);
    if (search.length < 2) return [];
    const fallbackSuggestions = buildFallbackGeographySuggestions(search);
    const restSuggestions = await this.searchLgdGeographiesViaRest(search).catch((error) => {
      console.warn("REST LGD geography lookup failed", error);
      return [];
    });
    if (restSuggestions.length) {
      return uniq([...restSuggestions, ...fallbackSuggestions]);
    }
    try {
      const response = await this.callAdmin("searchLgdGeographies", { query: search }, false);
      const suggestions = ensureList(response?.suggestions).map((item) => normalizeText(item)).filter(Boolean);
      return uniq([...suggestions, ...fallbackSuggestions]);
    } catch (error) {
      console.warn("Backend LGD geography lookup failed", error);
    }

    const client = this.getClient();
    if (!client) return fallbackSuggestions;
    const manualSuggestions = scoreLgdSuggestion("India", search) > 0 ? ["India"] : [];
    const wildcard = `%${search.replace(/[%_]/g, " ").trim()}%`;
    const { data, error } = await client
      .from("lgd_geography_directory")
      .select("display_label,location_kind,block_name,district_name,state_name,gram_panchayat_name,village_name")
      .or(`search_text.ilike.${wildcard},display_label.ilike.${wildcard},block_name.ilike.${wildcard},district_name.ilike.${wildcard},state_name.ilike.${wildcard}`)
      .limit(20);
    if (error) {
      console.warn("LGD geography lookup failed", error);
      return uniq([...manualSuggestions, ...fallbackSuggestions]);
    }
    return uniq([...manualSuggestions, ...buildRankedLgdSuggestions(ensureList(data), search), ...fallbackSuggestions]);
  }

  async searchLgdGeographiesViaRest(query) {
    const search = normalizeText(query);
    if (search.length < 2) return [];
    const supabaseUrl = normalizeText(this.config.SUPABASE_URL || "");
    const anonKey = normalizeText(this.config.SUPABASE_ANON_KEY || "");
    if (!supabaseUrl || !anonKey) return [];
    const wildcard = `*${search.replace(/[%_*]/g, " ").trim()}*`;
    const url =
      `${supabaseUrl}/rest/v1/lgd_geography_directory` +
      `?select=display_label,location_kind,block_name,district_name,state_name,gram_panchayat_name,village_name` +
      `&or=(` +
      `search_text.ilike.${encodeURIComponent(wildcard)},` +
      `display_label.ilike.${encodeURIComponent(wildcard)},` +
      `block_name.ilike.${encodeURIComponent(wildcard)},` +
      `district_name.ilike.${encodeURIComponent(wildcard)},` +
      `state_name.ilike.${encodeURIComponent(wildcard)}` +
      `)&limit=20`;
    const response = await fetch(url, {
      headers: {
        apikey: anonKey,
        Authorization: `Bearer ${anonKey}`,
      },
    });
    if (!response.ok) {
      throw new Error(`LGD REST lookup failed (${response.status}).`);
    }
    const rows = await response.json().catch(() => []);
    return uniq([
      ...(scoreLgdSuggestion("India", search) > 0 ? ["India"] : []),
      ...buildRankedLgdSuggestions(ensureList(rows), search),
    ]);
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

  async downloadGreInboundReport() {
    return this.callAdmin("downloadGreInboundReport", {}, true);
  }

  async uploadChatbotWorkbooks(solutionBase64, traderBase64, solutionFileName, traderFileName, aiProvider) {
    return this.callAdmin("uploadChatbotWorkbooks", { solutionBase64, traderBase64, solutionFileName, traderFileName, aiProvider }, true);
  }

  async uploadChatbotNormalized(traders, solutions, offerings, solutionFileName, traderFileName, aiProvider) {
    return this.callAdmin("uploadChatbotNormalized", { traders, solutions, offerings, solutionFileName, traderFileName, aiProvider }, true);
  }

  async downloadSeekerRequestTracker(seekerKey, includeClosed = false) {
    return this.callAdmin("downloadSeekerRequestTracker", { seekerKey, includeClosed });
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
        ? client.from("traders").select("trader_id,trader_name,organisation_name,email,website,mobile,poc_name").in("trader_id", traderIds)
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
      const alreadySharedWithSeeker = ensureList(item.matchReasons).some(
        (reason) => normalizeText(reason).toLowerCase() === "already shared with seeker",
      );
      if (alreadySharedWithSeeker) return true;
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
      const alreadySharedWithSeeker = ensureList(item.matchReasons).some(
        (reason) => normalizeText(reason).toLowerCase() === "already shared with seeker",
      );
      if (alreadySharedWithSeeker) return true;
      if (item.matchScore < 10) return false;
      if (!(item.domainMatched || item.thematicMatched)) return false;
      if (profile.hasStrongTheme && !item.primaryThematicMatched) return false;
      if ((profile.resolvedNeedKind === "service" || (profile.resolvedNeedKind === "mixed" && profile.serviceTerms.length)) && item.offeringKind === "knowledge") return false;
      return true;
    });
  }

  async searchManualSolutionMatches(need, filters = {}) {
    const client = this.getClient();
    if (!client || !need) return [];
    const profile = buildNeedMatchProfile(need);
    const offeringSelect = "offering_id,solution_id,trader_id,offering_name,offering_category,offering_group,offering_type,primary_application,primary_valuechain,applications,valuechains,domain_6m,tags,languages,geographies,about_offering_text,contact_details,gre_link,ai_thematic_area,ai_application_area,ai_offering_kind,ai_service_kind,ai_keywords";
    const keyword = normalizeText(filters.keyword);
    let query = client.from("offerings").select(offeringSelect).limit(350);
    if (normalizeText(filters.provider)) {
      query = query.eq("trader_id", filters.provider);
    }
    if (keyword) {
      query = query.textSearch("search_document", keyword, {
        config: "simple",
        type: "websearch",
      });
    }

    let response = await query;
    if (response.error && keyword) {
      let fallbackQuery = client.from("offerings").select(offeringSelect).limit(350);
      if (normalizeText(filters.provider)) {
        fallbackQuery = fallbackQuery.eq("trader_id", filters.provider);
      }
      response = await fallbackQuery;
    }
    if (response.error) {
      throw new Error(response.error.message || "Manual solution search failed.");
    }

    const selectedValuechain = normalizeText(filters.valuechain).toLowerCase();
    const selectedApplication = normalizeText(filters.application).toLowerCase();
    const selectedLanguage = normalizeText(filters.language).toLowerCase();
    const selectedGeography = normalizeText(filters.geography).toLowerCase();
    const selectedCategory = normalizeText(filters.category).toLowerCase();
    const selectedDomain6m = normalizeText(filters.domain6m).toLowerCase();
    const selectedOfferingType = normalizeText(filters.offeringType).toLowerCase();
    const keywordTokens = keyword.toLowerCase().split(/\s+/).filter((item) => item.length >= 2);

    const offerings = ensureList(response.data).filter((row) => {
      const valuechains = [row.primary_valuechain, ...parseArray(row.valuechains)].map((item) => normalizeText(item).toLowerCase());
      const applications = [row.primary_application, ...parseArray(row.applications)].map((item) => normalizeText(item).toLowerCase());
      const languages = canonicalizeLanguageArray(parseArray(row.languages)).map((item) => normalizeText(item).toLowerCase());
      const geographies = parseArray(row.geographies).map((item) => normalizeText(item).toLowerCase());
      const domains = parseArray(row.domain_6m).map((item) => normalizeText(item).toLowerCase());
      const category = normalizeText(row.offering_category).toLowerCase();
      const offeringType = normalizeText(row.offering_type).toLowerCase();
      const haystack = [
        row.offering_name,
        row.offering_category,
        row.offering_type,
        row.domain_6m,
        row.primary_application,
        row.primary_valuechain,
        parseArray(row.tags).join(" "),
        row.about_offering_text,
      ].map((item) => normalizeText(item).toLowerCase()).join(" ");

      if (selectedCategory && category !== selectedCategory) return false;
      if (selectedDomain6m && !domains.includes(selectedDomain6m)) return false;
      if (selectedOfferingType && offeringType !== selectedOfferingType) return false;
      if (selectedValuechain && !valuechains.includes(selectedValuechain)) return false;
      if (selectedApplication && !applications.includes(selectedApplication)) return false;
      if (selectedLanguage && !languages.includes(selectedLanguage)) return false;
      if (selectedGeography && !geographies.some((item) => item.includes(selectedGeography) || selectedGeography.includes(item))) return false;
      if (keywordTokens.length && !keywordTokens.every((token) => haystack.includes(token))) return false;
      return true;
    });

    if (!offerings.length) return [];
    const traderIds = [...new Set(offerings.map((item) => item.trader_id).filter(Boolean))];
    const solutionIds = [...new Set(offerings.map((item) => item.solution_id).filter(Boolean))];
    const [tradersRes, solutionsRes] = await Promise.all([
      traderIds.length
        ? client.from("traders").select("trader_id,trader_name,organisation_name,email,website,mobile,poc_name").in("trader_id", traderIds)
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
          primaryThematicMatched: matchMeta.primaryThematicMatched,
          serviceMatched: matchMeta.serviceMatched,
          domainMatched: matchMeta.domainMatched,
          offeringKind: matchMeta.offeringKind,
          matchReasons: matchMeta.reasons,
        };
      })
      .sort((left, right) => parseNumber(right.matchScore, 0) - parseNumber(left.matchScore, 0))
      .slice(0, 36);
  }
}

const store = new GreMisStore();

function getCuratorById(id) {
  return state.data.curators.find((curator) => curator.id === id) || null;
}

function getUserById(id) {
  return ensureList(state.data.users).find((user) => String(user.id) === String(id)) || null;
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
  if (hasAdminLikeAccess()) return true;
  if (!isCuratorUser()) return false;
  const myCurator = getCuratorByUser();
  return Boolean(myCurator && need.curator_id === myCurator.id);
}

function renderTemplateText(template, replacements) {
  return normalizeText(template).replace(/\{\{\s*([a-zA-Z0-9_]+)\s*\}\}/g, (_, key) => {
    const value = replacements?.[key];
    return value === undefined || value === null ? "" : String(value);
  });
}

function getAssignedCuratorForNeed(need) {
  return getCuratorById(need?.curator_id) || null;
}

function getCurrentActorDisplayName() {
  return normalizeText(state.userSession?.full_name || state.userSession?.first_name || state.userSession?.username || state.userSession?.email || "");
}

function resolveMisDisplayName(userLike) {
  const email = normalizeText(userLike?.email).toLowerCase();
  const fullName = normalizeText(userLike?.full_name || userLike?.display_name);
  if (email === "tanmay@greenruraleconomy.in" && /admin/i.test(fullName)) {
    return "Tanmay Mukherji";
  }
  return normalizeText(fullName || userLike?.first_name || userLike?.username || userLike?.email || "");
}

function isCurrentUserMappedToNeedCurator(need) {
  if (!need || !state.userSession) return false;
  const myCurator = getCuratorByUser();
  const assignedCurator = getAssignedCuratorForNeed(need);
  if (myCurator && myCurator.id === need.curator_id) return true;
  if (!assignedCurator) return false;
  const assignedUserId = normalizeText(assignedCurator.user_id);
  const assignedEmail = normalizeText(assignedCurator.email).toLowerCase();
  const actorUserId = normalizeText(state.userSession.id);
  const actorEmail = normalizeText(state.userSession.email).toLowerCase();
  return Boolean(
    (assignedUserId && actorUserId && assignedUserId === actorUserId) ||
    (assignedEmail && actorEmail && assignedEmail === actorEmail)
  );
}

function getMatchEmailActionConfig(need, match) {
  const assignedCurator = getAssignedCuratorForNeed(need);
  const providerEmail = normalizeText(parseContactDetailParts(match?.contact_details).email || match?.trader?.email || "").toLowerCase();
  const providerName = normalizeText(match?.trader?.organisation_name || match?.trader?.trader_name || "Solution Provider");
  const seekerEmail = normalizeText(need?.seeker_email).toLowerCase();
  const assignedCuratorEmail = normalizeText(assignedCurator?.email).toLowerCase();
  const actorEmail = normalizeText(state.userSession?.email).toLowerCase();
  const actorName = getCurrentActorDisplayName();
  const offeringName = normalizeText(match?.offering_name || match?.solution?.solution_name || "this solution");
  const viewLink = match?.offering_id
    ? `${window.location.origin}/offering-detail.html?offering_id=${encodeURIComponent(match.offering_id)}`
    : normalizeText(match?.gre_link || window.location.origin);
  const problemStatement = normalizeText(need?.problem_statement || "the shared problem statement");
  const seekerLabel = normalizeText(need?.organization_name || need?.contact_person || "the solution seeker");
  const isCuratorForward = Boolean(assignedCurator) && !isCurrentUserMappedToNeedCurator(need);

  if (isCuratorForward) {
    const body = renderTemplateText(
      state.mailTemplates.curatorForwardTemplate,
      {
        assignedCuratorName: normalizeText(assignedCurator?.display_name || "Assigned Curator"),
        seekerLabel,
        solutionSeeker: seekerLabel,
        problemStatement,
        viewLink,
        providerName,
        solutionProvider: providerName,
        actorName: actorName || actorEmail,
        curatorName: actorName || actorEmail,
      },
    );
    return {
      type: "curator_forward",
      actionLabel: "Email to Curator",
      eyebrow: "Curator Forward",
      title: "Review the curator-forward email",
      from: "Team GRE <help@greenruraleconomy.in>",
      to: assignedCuratorEmail,
      cc: actorEmail,
      replyTo: actorEmail,
      subject: `GRE review suggestion for ${seekerLabel}`,
      body,
      providerEmail,
      providerName,
      offeringId: normalizeText(match?.offering_id),
    };
  }

  const body = renderTemplateText(
    state.mailTemplates.providerIntroTemplate,
    {
      providerName,
      seekerLabel,
      solutionSeeker: seekerLabel,
      problemStatement,
      viewLink,
      offeringName,
    },
  );
  return {
    type: "provider_intro",
    actionLabel: "Email to Provider",
    eyebrow: "Provider Outreach",
    title: "Review the provider outreach email",
    from: "Team GRE <help@greenruraleconomy.in>",
    to: providerEmail,
    cc: [seekerEmail, assignedCuratorEmail].filter(Boolean).join(", "),
    replyTo: actorEmail,
    subject: `GRE introduction: ${seekerLabel} challenge for your consideration`,
    body,
    providerEmail,
    providerName,
    offeringId: normalizeText(match?.offering_id),
  };
}

function getMatchSeekerEmailActionConfig(need, match) {
  const assignedCurator = getAssignedCuratorForNeed(need);
  const curatorEmail = normalizeText(assignedCurator?.email).toLowerCase();
  const actorEmail = normalizeText(state.userSession?.email).toLowerCase();
  const providerName = normalizeText(match?.trader?.organisation_name || match?.trader?.trader_name || "Solution Provider");
  const solutionName = normalizeText(match?.offering_name || match?.solution?.solution_name || "Solution");
  const seekerLabel = normalizeText(need?.contact_person || need?.organization_name || "Seeker");
  const seekerEmail = normalizeText(need?.seeker_email).toLowerCase();
  const problemStatement = normalizeText(need?.problem_statement || "the stated need");
  const viewLink = match?.offering_id
    ? `${window.location.origin}/offering-detail.html?offering_id=${encodeURIComponent(match.offering_id)}`
    : normalizeText(match?.gre_link || window.location.origin);
  const body = renderTemplateText(
    state.mailTemplates.solutionSeekerTemplate || DEFAULT_SOLUTION_SEEKER_TEMPLATE,
    {
      seekerLabel,
      problemStatement,
      providerName,
      solutionName,
      viewLink,
    },
  );
  return {
    type: "solution_seeker_intro",
    actionLabel: "Email to Seeker",
    eyebrow: "Seeker Outreach",
    title: "Review the seeker outreach email",
    from: "Team GRE <solution@greenruraleconomy.in>",
    to: seekerEmail,
    cc: ["help@greenruraleconomy.in", curatorEmail].filter(Boolean).join(", "),
    replyTo: actorEmail,
    subject: `GRE solution identified for ${seekerLabel}`,
    body,
    providerEmail: normalizeText(parseContactDetailParts(match?.contact_details).email || match?.trader?.email || "").toLowerCase(),
    providerName,
    solutionName,
    offeringId: normalizeText(match?.offering_id),
  };
}

function getAssignableUsersForNeed(need) {
  const users = ensureList(state.data.users)
    .filter((user) => user && user.is_active !== false)
    .filter((user) => ["admin", "moderator", "curator"].includes(normalizeText(user.role).toLowerCase()));

  if (isLocalOnlyNeed(need)) {
    return users;
  }

  return users.filter((user) => {
    const role = normalizeText(user.role).toLowerCase();
    return ["admin", "moderator", "curator"].includes(role) && normalizeText(user.gre_user_id);
  });
}

function getCuratorRoleLabel(curator) {
  const matchedUser = ensureList(state.data.users).find((user) =>
    String(user.id || "") === String(curator.user_id || "")
    || String(user.email || "").toLowerCase() === String(curator.email || "").toLowerCase()
  );
  const role = normalizeText(matchedUser?.role).toLowerCase();
  if (role === "admin") return "Admin";
  if (role === "moderator") return "Moderator";
  if (role === "curator") return "Curator";
  return "";
}

function getCuratorWorkloadEntries(needs) {
  const entries = [];
  const seen = new Set();

  ensureList(state.data.curators).forEach((curator) => {
    const roleLabel = getCuratorRoleLabel(curator);
    const labelBase = curator.display_name || curator.email || curator.id;
    entries.push({
      label: roleLabel ? `${labelBase} (${roleLabel})` : labelBase,
      value: needs.filter((need) => need.curator_id === curator.id).length,
      focusKind: "curator",
      focusId: curator.id,
    });
    seen.add(`curator:${curator.id}`);
    if (curator.user_id) seen.add(`user:${curator.user_id}`);
    if (curator.email) seen.add(`email:${String(curator.email).toLowerCase()}`);
  });

  ensureList(state.data.users)
    .filter((user) => user && user.is_active !== false)
    .filter((user) => ["admin", "moderator", "curator"].includes(normalizeText(user.role).toLowerCase()))
    .forEach((user) => {
      const emailKey = String(user.email || "").toLowerCase();
      if (seen.has(`user:${user.id}`) || (emailKey && seen.has(`email:${emailKey}`))) return;
      const role = normalizeText(user.role).toLowerCase();
      const roleLabel = role === "admin" ? "Admin" : role === "moderator" ? "Moderator" : "Curator";
      entries.push({
        label: `${resolveMisDisplayName(user)} (${roleLabel})`,
        value: 0,
        focusKind: "curator",
        focusId: "",
      });
    });

  entries.push({
    label: "Unassigned",
    value: needs.filter((need) => !need.curator_id).length,
    focusKind: "curator",
    focusId: "unassigned",
  });

  return entries;
}

function getCuratorAssignmentOptions(need) {
  const options = [`<option value="">Unassigned</option>`];
  const seen = new Set([""]);
  const assignableUsers = getAssignableUsersForNeed(need);

  assignableUsers.forEach((user) => {
    const role = normalizeText(user.role).toLowerCase();
    const existingCurator = ensureList(state.data.curators).find((curator) =>
      String(curator.user_id || "") === String(user.id)
      || String(curator.email || "").toLowerCase() === String(user.email || "").toLowerCase()
    );
    const optionValue = existingCurator ? `curator:${existingCurator.id}` : `user:${user.id}`;
    if (seen.has(optionValue)) return;
    seen.add(optionValue);
    const selected = existingCurator && existingCurator.id === need.curator_id;
    const label = `${resolveMisDisplayName(user)} (${role})`;
    options.push(`<option value="${esc(optionValue)}" ${selected ? "selected" : ""}>${esc(label)}</option>`);
  });

  ensureList(state.data.curators).forEach((curator) => {
    const optionValue = `curator:${curator.id}`;
    if (seen.has(optionValue)) return;
    const selected = curator.id === need.curator_id;
    options.push(`<option value="${esc(optionValue)}" ${selected ? "selected" : ""}>${esc(curator.display_name || curator.email || curator.id)}</option>`);
  });

  return options.join("");
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
    .sort((a, b) => getNeedCurationAgeDays(b) - getNeedCurationAgeDays(a));
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
  if (mode === "solution") {
    const publicSolutionFormUrl = normalizeText(window.APP_CONFIG?.PUBLIC_SOLUTION_FORM_URL || "");
    if (publicSolutionFormUrl) return publicSolutionFormUrl;
  }
  if (mode === "need") {
    const publicNeedFormUrl = normalizeText(window.APP_CONFIG?.PUBLIC_NEED_FORM_URL || "");
    if (publicNeedFormUrl) return publicNeedFormUrl;
  }
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
    canonicalizeLanguageArray(parseArray(row.languages)).forEach((item) => lists.languages.add(item));
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

function buildFallbackGeographySuggestions(query) {
  const search = normalizeText(query).toLowerCase();
  if (!search) return [];
  const master = state.data.offeringMaster || buildOfferingMasterData([]);
  const expanded = uniq(
    ensureList(master.geographies)
      .flatMap((label) => {
        const text = normalizeText(label);
        if (!text) return [];
        const parts = text.split(",").map((item) => normalizeText(item)).filter(Boolean);
        const variants = [text];
        if (parts.length >= 3) {
          variants.push(`${parts[0]}, ${parts[1]}, India`);
          variants.push(`${parts[1]}, India`);
        } else if (parts.length === 2) {
          variants.push(`${parts[0]}, ${parts[1]}, India`);
          variants.push(`${parts[1]}, India`);
        }
        variants.push("India");
        return uniq(variants);
      }),
  );
  return expanded
    .filter((label) => normalizeText(label).toLowerCase().includes(search))
    .sort((left, right) => {
      const scoreDiff = scoreLgdSuggestion(right, search) - scoreLgdSuggestion(left, search);
      if (scoreDiff) return scoreDiff;
      return left.localeCompare(right);
    })
    .slice(0, 20);
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
        <span>${esc(translateLanguageLabel(item))}</span>
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

function renderInlineSuggestions(containerId, suggestions, dataAttribute) {
  const container = byId(containerId);
  if (!container) return;
  const labels = ensureList(suggestions).map((item) => normalizeText(item)).filter(Boolean).slice(0, 12);
  container.innerHTML = labels
    .map((label) => `
      <button
        type="button"
        class="inline-suggestion-pill"
        ${dataAttribute}="${escAttr(label)}"
      >${esc(label)}</button>
    `)
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

function parseContactDetailParts(value) {
  const text = normalizeText(value);
  if (!text) return { text: "", name: "", email: "", phone: "" };
  const emailMatch = text.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  const phoneMatch = text.match(/(?:\+?\d[\d\s().-]{7,}\d)/);
  const lines = text.split(/\r?\n|[,;|]/).map((item) => normalizeText(item)).filter(Boolean);
  const firstLine = lines[0] || "";
  const looksLikeName = firstLine && !/@/.test(firstLine) && !/\d{5,}/.test(firstLine);
  return {
    text,
    name: looksLikeName ? firstLine : "",
    email: emailMatch ? emailMatch[0] : "",
    phone: phoneMatch ? phoneMatch[0].trim() : "",
  };
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

function splitCombinedUnitValue(value, knownUnits = []) {
  const text = normalizeText(value);
  if (!text) return { value: "", unit: "" };
  const normalizedUnits = ensureList(knownUnits).map((item) => normalizeText(item)).filter(Boolean);
  const matchedUnit = [...normalizedUnits]
    .sort((left, right) => right.length - left.length)
    .find((unit) => text.toLowerCase().endsWith(unit.toLowerCase()));
  if (!matchedUnit) {
    return { value: text, unit: "" };
  }
  const numeric = normalizeText(text.slice(0, text.length - matchedUnit.length));
  return {
    value: numeric || text,
    unit: matchedUnit,
  };
}

function getStoredPayloadRecord(value) {
  if (!value || typeof value !== "object") return {};
  const record = value;
  if (record.payload && typeof record.payload === "object") {
    return record.payload;
  }
  return record;
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
  renderOfferingButtonGrid("solution");
  applySharedFormLanguage().catch(() => null);
}

function updateNeedOfferingForm() {
  const categorySelect = byId("needOfferingCategory");
  if (!categorySelect) return;
  const category = categorySelect.value || "Service offerings";
  setSelectOptions("needOfferingType", OFFERING_TYPE_OPTIONS[category] || [], true);
  renderOfferingButtonGrid("need");
  applySharedFormLanguage().catch(() => null);
}

function buildOfferingButtonGridMarkup(scope, selectedCategory, selectedType) {
  return OFFERING_CATEGORY_OPTIONS.map((category) => {
    const categoryValue = category.value;
    const items = OFFERING_TYPE_OPTIONS[categoryValue] || [];
    return `
      <div class="offering-button-row" data-offering-row="${escAttr(categoryValue)}">
        <span class="offering-button-row-label">${esc(translateOfferingCategoryLabel(category.label))}</span>
        <div class="offering-button-row-actions">
          ${items.map((item) => {
            const active = selectedCategory === categoryValue && normalizeText(selectedType) === normalizeText(item.value);
            return `
              <button
                type="button"
                class="btn ${active ? "btn-primary" : "btn-secondary"} btn-compact offering-choice-btn"
                data-offering-scope="${escAttr(scope)}"
                data-offering-category="${escAttr(categoryValue)}"
                data-offering-type="${escAttr(item.value)}"
                aria-pressed="${active ? "true" : "false"}"
              >${esc(translateOfferingTypeLabel(item.label))}</button>
            `;
          }).join("")}
        </div>
      </div>
    `;
  }).join("");
}

function renderOfferingButtonGrid(scope) {
  const isSolution = scope === "solution";
  const categorySelect = byId(isSolution ? "solutionOfferingCategory" : "needOfferingCategory");
  const typeSelect = byId(isSolution ? "solutionOfferingType" : "needOfferingType");
  const host = byId(isSolution ? "solutionOfferingButtonGrid" : "needOfferingButtonGrid");
  if (!categorySelect || !typeSelect || !host) return;
  const selectedCategory = categorySelect.value || "Service offerings";
  const selectedType = typeSelect.value || "";
  host.innerHTML = buildOfferingButtonGridMarkup(scope, selectedCategory, selectedType);
}

function getSharedFormSeedGeographies() {
  return ["India"];
}

function isAllowedSharedFormGeography(location, suggestions = []) {
  const normalizedLocation = normalizeText(location);
  if (!normalizedLocation) return false;
  if (normalizedLocation === "India") return true;
  return ensureList(suggestions).some((item) => normalizeText(item) === normalizedLocation);
}

function applyOfferingButtonSelection(scope, category, type) {
  const isSolution = scope === "solution";
  const categorySelect = byId(isSolution ? "solutionOfferingCategory" : "needOfferingCategory");
  if (!categorySelect) return;
  categorySelect.value = category;
  if (isSolution) {
    updateSolutionOfferingForm();
  } else {
    updateNeedOfferingForm();
  }
  const typeSelect = byId(isSolution ? "solutionOfferingType" : "needOfferingType");
  if (typeSelect) {
    typeSelect.value = type;
  }
  renderOfferingButtonGrid(scope);
}

function renderSolutionReferenceInputs() {
  const master = state.data.offeringMaster || buildOfferingMasterData([]);
  const languageOptions = uniq([
    ...DEFAULT_LANGUAGE_OPTIONS,
    ...canonicalizeLanguageArray(master.languages || []),
  ]);
  setCheckboxGroupOptions("solutionLanguagesGroup", "languages", languageOptions);
  const seedGeographies = getSharedFormSeedGeographies();
  state.solutionGeographySuggestions = [...seedGeographies];
  setDatalistOptions("solutionGeographyOptions", seedGeographies);
  renderInlineSuggestions("solutionGeographySuggestions", seedGeographies, "data-add-solution-geo-suggestion");
  setDatalistOptions("solutionTagOptions", master.tags);
  updateSolutionOfferingForm();
}

function renderNeedReferenceInputs() {
  const master = state.data.offeringMaster || buildOfferingMasterData([]);
  const seedGeographies = getSharedFormSeedGeographies();
  state.needDeploymentSuggestions = [...seedGeographies];
  setDatalistOptions("needDeploymentOptions", seedGeographies);
  renderInlineSuggestions("needDeploymentSuggestions", seedGeographies, "data-add-need-deployment-suggestion");
  setDatalistOptions("needTagOptions", master.tags);
  updateNeedOfferingForm();
}

function ensureSharedFormInteractiveControls() {
  if (!isSharedFormMode()) return;
  const master = state.data.offeringMaster || buildOfferingMasterData([]);
  const languageOptions = uniq([
    ...DEFAULT_LANGUAGE_OPTIONS,
    ...canonicalizeLanguageArray(master.languages || []),
  ]);
  const languageGroup = byId("solutionLanguagesGroup");
  if (languageGroup && !languageGroup.querySelector('input[name="languages"]')) {
    setCheckboxGroupOptions("solutionLanguagesGroup", "languages", languageOptions);
  }
  const solutionGeographyList = byId("solutionGeographyOptions");
  if (solutionGeographyList && !solutionGeographyList.querySelector("option")) {
    const seedGeographies = getSharedFormSeedGeographies();
    state.solutionGeographySuggestions = [...seedGeographies];
    setDatalistOptions("solutionGeographyOptions", seedGeographies);
    renderInlineSuggestions("solutionGeographySuggestions", seedGeographies, "data-add-solution-geo-suggestion");
  }
  const needDeploymentList = byId("needDeploymentOptions");
  if (needDeploymentList && !needDeploymentList.querySelector("option")) {
    const seedGeographies = getSharedFormSeedGeographies();
    state.needDeploymentSuggestions = [...seedGeographies];
    setDatalistOptions("needDeploymentOptions", seedGeographies);
    renderInlineSuggestions("needDeploymentSuggestions", seedGeographies, "data-add-need-deployment-suggestion");
  }
  if (isHindiSharedFormMode()) {
    renderOfferingButtonGrid("solution");
    renderOfferingButtonGrid("need");
  }
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
    const seedSuggestions = getSharedFormSeedGeographies();
    state.solutionGeographySuggestions = [...seedSuggestions];
    setDatalistOptions("solutionGeographyOptions", seedSuggestions);
    renderInlineSuggestions("solutionGeographySuggestions", seedSuggestions, "data-add-solution-geo-suggestion");
    return;
  }
  const suggestions = await store.searchLgdGeographies(query);
  if (state.lgdSearchToken !== requestToken) return;
  const nextSuggestions = suggestions.length ? suggestions : getSharedFormSeedGeographies();
  state.solutionGeographySuggestions = [...nextSuggestions];
  setDatalistOptions("solutionGeographyOptions", nextSuggestions);
  renderInlineSuggestions("solutionGeographySuggestions", nextSuggestions, "data-add-solution-geo-suggestion");
}

async function updateNeedDeploymentSuggestions(rawValue) {
  const input = byId("needDeploymentEntry");
  if (!input) return;
  const query = normalizeText(String(rawValue || ""));
  const requestToken = Date.now();
  state.lgdSearchToken = requestToken;
  if (query.length < 2) {
    const seedSuggestions = getSharedFormSeedGeographies();
    state.needDeploymentSuggestions = [...seedSuggestions];
    setDatalistOptions("needDeploymentOptions", seedSuggestions);
    renderInlineSuggestions("needDeploymentSuggestions", seedSuggestions, "data-add-need-deployment-suggestion");
    return;
  }
  const suggestions = await store.searchLgdGeographies(query);
  if (state.lgdSearchToken !== requestToken) return;
  const nextSuggestions = suggestions.length ? suggestions : getSharedFormSeedGeographies();
  state.needDeploymentSuggestions = [...nextSuggestions];
  setDatalistOptions("needDeploymentOptions", nextSuggestions);
  renderInlineSuggestions("needDeploymentSuggestions", nextSuggestions, "data-add-need-deployment-suggestion");
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
    : `<span class="helper-text">${esc(isHindiSharedFormMode() ? "अभी तक कोई स्थान नहीं जोड़ा गया है।" : "No locations added yet.")}</span>`;
}

function renderNeedDeploymentChips() {
  const target = byId("needDeploymentChips");
  const hidden = byId("needDeploymentLocations");
  if (!target || !hidden) return;
  hidden.value = state.needDeploymentLocations.join("; ");
  target.innerHTML = state.needDeploymentLocations.length
    ? state.needDeploymentLocations
        .map(
          (location) => `
            <span class="tag-chip">
              ${esc(location)}
              <button type="button" class="tag-chip-remove" data-remove-need-deployment="${escAttr(location)}" aria-label="Remove ${escAttr(location)}">&times;</button>
            </span>
          `,
        )
        .join("")
    : `<span class="helper-text">${esc(isHindiSharedFormMode() ? "अभी तक कोई तैनाती स्थान नहीं जोड़ा गया है।" : "No deployment locations added yet.")}</span>`;
}

function addSolutionGeography(rawValue) {
  const location = normalizeText(rawValue);
  if (!location || !isAllowedSharedFormGeography(location, state.solutionGeographySuggestions)) {
    toast("Please choose a geography from the LGD list.");
    return false;
  }
  if (!state.solutionGeographies.includes(location)) state.solutionGeographies.push(location);
  renderSolutionGeographyChips();
  return true;
}

function addNeedDeploymentLocation(rawValue) {
  const location = normalizeText(rawValue);
  if (!location || !isAllowedSharedFormGeography(location, state.needDeploymentSuggestions)) {
    toast("Please choose a geography from the LGD list.");
    return false;
  }
  if (!state.needDeploymentLocations.includes(location)) state.needDeploymentLocations.push(location);
  renderNeedDeploymentChips();
  return true;
}

function removeSolutionGeography(location) {
  state.solutionGeographies = state.solutionGeographies.filter((item) => item !== location);
  renderSolutionGeographyChips();
}

function removeNeedDeploymentLocation(location) {
  state.needDeploymentLocations = state.needDeploymentLocations.filter((item) => item !== location);
  renderNeedDeploymentChips();
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

function renderNeedTagChips() {
  const target = byId("needTagChips");
  const hidden = byId("needTagsInput");
  if (!target || !hidden) return;
  hidden.value = state.needTags.join(", ");
  target.innerHTML = state.needTags.length
    ? state.needTags
        .map(
          (tag) => `
            <span class="tag-chip">
              ${esc(tag)}
              <button type="button" class="tag-chip-remove" data-remove-need-tag="${escAttr(tag)}" aria-label="Remove ${escAttr(tag)}">&times;</button>
            </span>
          `,
        )
        .join("")
    : `<span class="helper-text">No keywords added yet.</span>`;
}

function addSolutionTag(rawValue) {
  const tag = normalizeText(rawValue);
  if (!tag) return false;
  if (!state.solutionTags.includes(tag)) state.solutionTags.push(tag);
  renderSolutionTagChips();
  return true;
}

function addNeedTag(rawValue) {
  const tag = normalizeText(rawValue);
  if (!tag) return false;
  if (!state.needTags.includes(tag)) state.needTags.push(tag);
  renderNeedTagChips();
  return true;
}

function removeSolutionTag(tag) {
  state.solutionTags = state.solutionTags.filter((item) => item !== tag);
  renderSolutionTagChips();
}

function removeNeedTag(tag) {
  state.needTags = state.needTags.filter((item) => item !== tag);
  renderNeedTagChips();
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
      { key: "offering_category", label: "Need Category" },
      { key: "offering_type", label: "Need Type" },
      { key: "thematic_area", label: "Thematic Area" },
      { key: "deployment_locations", label: "Place of Deployment", multiline: true, list: true, wide: true },
      { key: "demand_broadcast_needed", label: "Broadcast to Ecosystem" },
      { key: "keywords", label: "Keywords for Need Statement", multiline: true, list: true, wide: true },
      { key: "problem_statement", label: "Need Statement", multiline: true, wide: true },
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
  const section = byId("submissionReviewAttachmentsSection");
  if (!target) return;
  const configs = getSubmissionAttachmentConfigs(submissionType, payload);
  if (!configs.length) {
    if (section) section.hidden = true;
    target.innerHTML = "";
    return;
  }
  if (section) section.hidden = false;
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

function getSubmissionReviewNeedDraft() {
  const form = byId("submissionReviewForm");
  if (!form) return null;
  return {
    organization_name: normalizeText(form.querySelector('[name="organization_name"]')?.value || ""),
    offering_category: normalizeText(form.querySelector('[data-review-field="offering_category"]')?.value || ""),
    offering_type: normalizeText(form.querySelector('[data-review-field="offering_type"]')?.value || ""),
    thematic_area: normalizeText(form.querySelector('[data-review-field="thematic_area"]')?.value || ""),
    problem_statement: normalizeText(form.querySelector('[data-review-field="problem_statement"]')?.value || ""),
    deployment_locations: parseLooseListInput(form.querySelector('[data-review-field="deployment_locations"]')?.value || ""),
    keywords: parseLooseListInput(form.querySelector('[data-review-field="keywords"]')?.value || ""),
  };
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
  const needTools = byId("submissionReviewNeedTools");
  const needToolsStatus = byId("submissionNeedToolsStatus");
  if (needTools) needTools.classList.toggle("hidden", submission.submission_type !== "need");
  if (needToolsStatus) needToolsStatus.textContent = "";
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

function populateSubmissionDefaults(form, force = false) {
  const actor = getSubmissionActor();
  if (!form || !actor) return;
  const organizationName = normalizeText(actor.organization) || "Individual";
  const displayName = getSubmissionActorDisplayName(actor);
  [
    ["organization_name", organizationName],
    ["submitter_name", displayName],
    ["submitter_email", actor.email || ""],
    ["submitter_phone", actor.phone || ""],
    ["contact_person", displayName],
    ["seeker_email", actor.email || ""],
    ["seeker_phone", actor.phone || ""],
  ].forEach(([name, value]) => {
    const input = form.querySelector(`[name="${name}"]`);
    if (input && value && (force || !input.value)) input.value = value;
  });
}

function ensureEmbeddedSubmissionActionRow(formId, mode) {
  const form = byId(formId);
  if (!form) return;
  const primarySectionGrid = form.querySelector(".form-section .form-grid");
  if (!primarySectionGrid) return;
  let actionRow = form.querySelector(`[data-on-behalf-row="${mode}"]`);
  if (!isEmbeddedSharedForm()) {
    actionRow?.remove();
    return;
  }
  if (!actionRow) {
    actionRow = document.createElement("div");
    actionRow.className = "wide card-actions";
    actionRow.dataset.onBehalfRow = mode;
    primarySectionGrid.appendChild(actionRow);
  }
  actionRow.innerHTML = mode === "solution"
    ? `
      <p class="helper-text">These contact fields are being filled from the active GramEEE login for this solution submission.</p>
    `
    : `
      <p class="helper-text">These contact fields are being filled from the active GramEEE login for this need submission.</p>
    `;
}

function configureEmbeddedSubmissionForms() {
  [
    { formId: "solutionSubmissionForm", selectId: "solutionTraderSelect", mode: "solution" },
    { formId: "needSubmissionForm", selectId: "needTraderSelect", mode: "need" },
  ].forEach(({ formId, selectId, mode }) => {
    const form = byId(formId);
    const select = byId(selectId);
    const label = select?.closest("label");
    if (select) {
      if (isEmbeddedSharedForm()) {
        select.value = "";
      }
      select.disabled = isEmbeddedSharedForm();
      select.required = !isEmbeddedSharedForm() && mode === "need";
    }
    label?.classList.toggle("hidden", isEmbeddedSharedForm());
    ensureEmbeddedSubmissionActionRow(formId, mode);
    populateSubmissionDefaults(form, isEmbeddedSharedForm());
  });
}

function bindEmbeddedBridge() {
  if (!isEmbeddedSharedForm()) return;
  window.addEventListener("message", (event) => {
    if (!isAllowedGrameeeOrigin(event.origin)) return;
    const data = event.data && typeof event.data === "object" ? event.data : {};
    if (data.type === "grameee-form-context") {
      applyEmbeddedContext(data.payload || {}, event.origin);
    }
    if (data.type === "grameee-form-add-user-created") {
      applyEmbeddedContext({ user: data.payload?.user || data.payload }, event.origin);
      toast("Form details updated for the selected organisation or person.");
    }
    if (data.type === "grameee-form-language") {
      setSharedFormLanguage(data.payload?.language || "en", { persist: true });
    }
  });
  requestEmbeddedContext();
}

function attachGrameeeAuthBridgeListener() {
  if (document.body.dataset.greAuthBridgeBound === "true") return;
  document.body.dataset.greAuthBridgeBound = "true";

  document.addEventListener("grameee:auth-updated", safeAsync(async (event) => {
    if (isSharedFormMode() || isStandalonePublicFormMode()) return;

    const sharedUser = event?.detail?.user || null;

    if (!sharedUser) {
      if (state.userSession) {
        await store.userLogout().catch(() => null);
      }
      renderAuthState();
      renderSubmissionViews();
      renderAdminState();
      return;
    }

    const sharedRole = normalizeText(sharedUser.role).toLowerCase();
    if (!["admin", "moderator", "curator"].includes(sharedRole)) {
      renderAuthState();
      renderSubmissionViews();
      renderAdminState();
      return;
    }

    const sameUser = state.userSession
      && normalizeText(state.userSession.email).toLowerCase() === normalizeText(sharedUser.email).toLowerCase()
      && normalizeText(state.userSession.role).toLowerCase() === sharedRole;

    if (!sameUser) {
      const bridged = await store.bridgeGrameeeSession();
      if (!bridged) {
        renderAuthState();
        renderSubmissionViews();
        renderAdminState();
        return;
      }
    }

    await refreshAll();
  }));
}

async function tryBridgeSharedSessionAndRefresh() {
  if (isSharedFormMode() || isStandalonePublicFormMode()) return false;
  if (state.userToken) return true;
  if (document.body.dataset.greBridgePending === "true") return false;

  const summary = readGrameeeSummary();
  const accessToken = await getGrameeeAccessToken();
  const sharedRole = normalizeText(summary?.role).toLowerCase();

  if (!accessToken || !["admin", "moderator", "curator"].includes(sharedRole)) {
    return false;
  }

  document.body.dataset.greBridgePending = "true";
  try {
    const bridged = await store.bridgeGrameeeSession();
    if (bridged) {
      await refreshAll();
      return true;
    }
    renderAuthState();
    renderSubmissionViews();
    renderAdminState();
    return false;
  } finally {
    document.body.dataset.greBridgePending = "false";
  }
}

function renderSubmissionViews() {
  fillSupplierSelect("solutionTraderSelect", "solutionOrgName");
  fillSupplierSelect("needTraderSelect", "needOrgName");
  renderSolutionReferenceInputs();
  renderNeedReferenceInputs();
  renderSolutionTagChips();
  renderSolutionGeographyChips();
  renderNeedTagChips();
  renderNeedDeploymentChips();
  renderSharePanel("solutionSharePanel", "solution");
  renderSharePanel("needSharePanel", "need");
  configureEmbeddedSubmissionForms();
  populateSubmissionDefaults(byId("solutionSubmissionForm"), isEmbeddedSharedForm());
  populateSubmissionDefaults(byId("needSubmissionForm"), isEmbeddedSharedForm());

    const solutionTitle = document.querySelector('#solutionView h3');
    const needTitle = document.querySelector('#need-intakeView h3');
    if (solutionTitle) {
      solutionTitle.textContent = isSharedFormMode() ? "Share a Solution for GRE" : "Add Solution to Admin Review Queue";
    }
  if (needTitle) {
    needTitle.textContent = isSharedFormMode() ? "Share a Need with GRE" : "Need Help";
  }
  ensureSharedFormInteractiveControls();
  queueSharedFormLanguageRefresh();
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
      ["stuck", "Stuck 7+ Days", needs.filter((need) => getNeedCurationAgeDays(need) >= 7).length, "Aging needs that need intervention during the day."],
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

  const workload = getCuratorWorkloadEntries(needs);
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
        <p class="signal-heading">Need Themes</p>
        <div class="tag-cloud">
        ${Object.entries(themeCounts)
          .sort((a, b) => Number(b[1]) - Number(a[1]) || String(a[0]).localeCompare(String(b[0])))
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

function getSeekerKey(need) {
  return normalizeText(need.seeker_email).toLowerCase()
    || `org:${normalizeText(need.organization_name).toLowerCase()}`;
}

function renderRequestTrackerControls() {
  const select = byId("requestTrackerSeekerSelect");
  if (!select) return;
  const seekers = uniqBy(
    getDisplayNeeds()
      .map((need) => ({
        key: getSeekerKey(need),
        label: normalizeText(need.organization_name) || normalizeText(need.contact_person) || normalizeText(need.seeker_email) || "Unknown seeker",
        email: normalizeText(need.seeker_email),
      }))
      .filter((item) => item.key),
    (item) => item.key,
  ).sort((left, right) => left.label.localeCompare(right.label));
  select.innerHTML = seekers.length
    ? seekers.map((item) => `<option value="${esc(item.key)}">${esc(item.label)}${item.email ? ` (${esc(item.email)})` : ""}</option>`).join("")
    : `<option value="">No seekers available</option>`;
}

function resetManualSolutionSearch() {
  state.manualSolutionSearch = {
    provider: "",
    category: "",
    domain6m: "",
    offeringType: "",
    valuechain: "",
    application: "",
    language: "",
    geography: "",
    keyword: "",
    results: [],
    searched: false,
    loading: false,
    page: 1,
  };
}

function buildManualSelectOptions(values, blankLabel, currentValue = "") {
  const current = normalizeText(currentValue);
  return [
    `<option value="">${esc(blankLabel)}</option>`,
    ...ensureList(values).map((item) => {
      const value = normalizeText(typeof item === "string" ? item : item.value);
      const label = normalizeText(typeof item === "string" ? item : item.label) || value;
      return `<option value="${esc(value)}" ${value === current ? "selected" : ""}>${esc(label)}</option>`;
    }).filter(Boolean),
  ].join("");
}

function setManualSelectOptions(id, values, blankLabel, currentValue = "") {
  const select = byId(id);
  if (select) select.innerHTML = buildManualSelectOptions(values, blankLabel, currentValue);
}

function renderManualSolutionSearch() {
  const form = byId("manualSolutionSearchForm");
  const resultsEl = byId("manualMatchResults");
  const statusEl = byId("manualSolutionSearchStatus");
  if (!form || !resultsEl) return;

  const need = getNeedById(state.selectedNeedId);
  const master = state.data.offeringMaster || buildOfferingMasterData([]);
  const filters = state.manualSolutionSearch;
  const providerOptions = ensureList(state.data.traders)
    .map((trader) => ({
      value: normalizeText(trader.trader_id),
      label: normalizeText(trader.organisation_name || trader.trader_name || trader.email || trader.trader_id),
    }))
    .filter((item) => item.value && item.label)
    .sort((left, right) => left.label.localeCompare(right.label));

  setManualSelectOptions("manualSolutionProvider", providerOptions, "All solution providers", filters.provider);
  setManualSelectOptions("manualSolutionCategory", OFFERING_CATEGORY_OPTIONS, "All categories", filters.category);
  setManualSelectOptions("manualSolutionDomain6m", SIX_M_LABELS, "All 6M domains", filters.domain6m);
  setManualSelectOptions(
    "manualSolutionOfferingType",
    Object.values(OFFERING_TYPE_OPTIONS).flat(),
    "All offering types",
    filters.offeringType,
  );
  setManualSelectOptions("manualSolutionValuechain", master.valuechains, "All value chains", filters.valuechain);
  setManualSelectOptions("manualSolutionApplication", master.applications, "All applications", filters.application);
  setManualSelectOptions("manualSolutionLanguage", master.languages, "All languages", filters.language);
  setManualSelectOptions("manualSolutionGeography", master.geographies, "All geographies", filters.geography);
  const keywordInput = byId("manualSolutionKeyword");
  if (keywordInput) keywordInput.value = filters.keyword || "";

  if (!need) {
    if (statusEl) statusEl.textContent = "";
    resultsEl.innerHTML = `<div class="empty-state">Select a need before manually searching for additional solutions.</div>`;
    return;
  }
  if (filters.loading) {
    if (statusEl) statusEl.textContent = "Searching solutions by selected parameters...";
    resultsEl.innerHTML = `<div class="empty-state">Searching additional solutions...</div>`;
    return;
  }
  if (!filters.searched) {
    if (statusEl) statusEl.textContent = "Use parameters to search beyond the automatic GRE Knowledge Match list.";
    resultsEl.innerHTML = `<div class="empty-state">No manual search has been run for this need yet.</div>`;
    return;
  }
  if (statusEl) {
    statusEl.textContent = `${filters.results.length} manual solution match${filters.results.length === 1 ? "" : "es"} found.`;
  }
  const pageSize = 4;
  const totalPages = Math.max(1, Math.ceil(filters.results.length / pageSize));
  if (filters.page > totalPages) filters.page = totalPages;
  const pageStart = ((filters.page || 1) - 1) * pageSize;
  const visibleMatches = filters.results.slice(pageStart, pageStart + pageSize);
  resultsEl.innerHTML = filters.results.length
    ? visibleMatches.map((match) => buildMatchCardHtml(match, need)).join("") + `
      <div class="pipeline-pagination">
        <span class="meta-text">Showing ${esc(pageStart + 1)}-${esc(Math.min(pageStart + pageSize, filters.results.length))} of ${esc(filters.results.length)} manual matches</span>
        <div class="pipeline-pagination-actions">
          <button class="btn btn-secondary" data-manual-match-page-action="prev" ${filters.page <= 1 ? "disabled" : ""}>Prev</button>
          <span class="meta-text">Page ${esc(filters.page || 1)} of ${esc(totalPages)}</span>
          <button class="btn btn-secondary" data-manual-match-page-action="next" ${filters.page >= totalPages ? "disabled" : ""}>Next</button>
        </div>
      </div>`
    : `<div class="empty-state">No manual solution matches found for these parameters. Try reducing filters or searching by keyword.</div>`;
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
          <p class="meta-text">${esc(need.state)}${need.district ? ` / ${esc(need.district)}` : ""} • Curator: ${esc(curator?.display_name || "Unassigned")} • Age: ${esc(getNeedCurationAgeDays(need))} days</p>
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
  const sharedSolutionEntries = getNeedSharedSolutionDisplayEntries(need);
  const noteText = String(stripUrls(need.curation_notes || ""))
    .replace(/\r/g, "")
    .replace(/\bnull\b/gi, "")
    .replace(/solutions shared\n(\d+\..*(?:\n|$))*/gi, "")
    .replace(/solutions shared\s*:.*(?:\n|$)/gi, "")
    .trim();
  const noteHtml = noteText ? esc(noteText).replace(/\n/g, "<br>") : "";
  const canInspectCuration = canSeeCurationDetails();
  const summaryBadges = [
    { label: need.status, tone: badgeTone(need.status) },
    { label: need.internal_status, tone: badgeTone(need.internal_status) },
    { label: `${getNeedCurationAgeDays(need)} days old`, tone: getNeedCurationAgeDays(need) >= 7 ? "bad" : "info" },
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
                <div><span>Solutions Shared</span><strong>${esc(getNeedSharedSolutionCount(need))}</strong></div>
                <div><span>Invited Providers</span><strong>${esc(need.invited_providers_count || 0)}</strong></div>
                <div><span>Current Stage</span><strong>${esc(need.internal_status || "Not set")}</strong></div>
                <div><span>Funding Mechanism</span><strong>${esc(need.funding_mechanism || "Not set")}</strong></div>
                <div><span>Seeker / Provider Agreement</span><strong>${esc(need.seeker_provider_agreement || "Not set")}</strong></div>
                <div><span>Deployment Status</span><strong>${esc(need.solution_deployment_status || "Not set")}</strong></div>
                <div><span>Closure Date</span><strong>${esc(need.closure_date || "Not set")}</strong></div>
              </div>
              ${(need.feedback_about_seeker || need.feedback_about_provider) ? `
                <div class="detail-list">
                  ${need.feedback_about_seeker ? `<div><strong>Feedback about Seeker</strong><span>${esc(need.feedback_about_seeker)}</span></div>` : ""}
                  ${need.feedback_about_provider ? `<div><strong>Feedback about Provider</strong><span>${esc(need.feedback_about_provider)}</span></div>` : ""}
                </div>
              ` : ""}
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
            <div><span>Curation Age</span><strong>${esc(getNeedCurationAgeDays(need))} days</strong></div>
          </div>
        </article>
      </div>

        ${canInspectCuration
          ? `<article class="detail-card detail-stack-card">
             <h4>Curation Notes</h4>
             ${noteText ? `<p class="detail-note">${noteHtml}</p>` : `<p class="detail-note">No curation notes have been recorded yet.</p>`}
            ${
              sharedSolutionEntries.length
                ? `<div class="detail-links">
                    <p class="signal-heading">Solutions Shared with Seeker (${sharedSolutionEntries.length})</p>
                    <ol class="detail-link-list">
                      ${sharedSolutionEntries
                        .map((entry) => {
                          const link = entry.url;
                          return `<li>${esc(entry.providerName || entry.offeringName)}${entry.providerName && entry.offeringName ? ` : ${esc(entry.offeringName)}` : ""}${link ? ` : <a href="${esc(link)}" target="_blank" rel="noreferrer">View Solution</a>` : ""}${entry.inferred ? ` <span class="helper-text">(GRE-aligned)</span>` : ""}</li>`;
                        })
                        .join("")}
                    </ol>
                  </div>`
                : solutionLinks.length
                  ? `<div class="detail-links">
                      <p class="signal-heading">Solutions Shared with Seeker</p>
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
                ${canEditNeedCuration(need) ? `<button type="button" class="btn btn-secondary" data-action="generate-suggested-questions" data-need-id="${esc(need.id)}">Generate Suggested Questions</button><button type="button" class="btn btn-secondary" data-action="generate-suggested-questions-puter" data-need-id="${esc(need.id)}">Use your local AI using Puter for Questions</button>${hasSuggestedQuestionSection(need.curation_notes) ? `<button type="button" class="btn btn-secondary" data-action="reject-suggested-questions" data-need-id="${esc(need.id)}">Reject Suggested Questions</button>` : ""}` : ""}
        </div>
      </div>
    `;
}

function buildMatchDetailHtml(match) {
  if (!match) return `<div class="empty-state">Match details are not available.</div>`;
  const fullText = normalizeText(match.about_offering_text || match.solution?.about_solution_text || "No detailed offering summary is available yet.");
  const contactParts = parseContactDetailParts(match.contact_details);
  const preferredEmail = contactParts.email || match.trader?.email || "";
  const preferredPhone = contactParts.phone || match.trader?.mobile || match.trader?.phone || "";
  const preferredName = contactParts.name || match.trader?.poc_name || "";
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
          <div><span>Email</span><strong>${esc(preferredEmail || "Not set")}</strong></div>
          <div><span>Phone</span><strong>${esc(preferredPhone || "Not set")}</strong></div>
          <div><span>Point of Contact</span><strong>${esc(preferredName || "Not set")}</strong></div>
          <div><span>Website</span><strong>${esc(match.trader?.website || "Not set")}</strong></div>
          <div><span>Contact Details</span><strong>${esc(match.contact_details || "Not set")}</strong></div>
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
  if (match?.offering_id) {
    body.innerHTML = `
      <div class="match-detail-frame-wrap">
        <iframe
          class="match-detail-frame"
          src="./offering-detail.html?offering_id=${encodeURIComponent(match.offering_id)}&embedded=1&v=20260521gre8"
          title="${esc(title.textContent || "Match Details")}"
          loading="lazy"
        ></iframe>
      </div>
    `;
  } else {
    body.innerHTML = buildMatchDetailHtml(match);
  }
  dialog.showModal();
}

function closeMailReviewDialog() {
  state.pendingMailReview = null;
  const dialog = byId("mailReviewDialog");
  dialog?.close();
  setText("mailReviewStatus", "");
}

function openMailReviewDialog(mailReview) {
  state.pendingMailReview = mailReview;
  setText("mailReviewEyebrow", mailReview.eyebrow || "Provider Outreach");
  setText("mailReviewTitle", mailReview.title || "Review the GRE outreach email");
  setText("mailReviewFrom", mailReview.from || "");
  setText("mailReviewTo", mailReview.to || "Not available");
  setText("mailReviewCc", mailReview.cc || "None");
  setText("mailReviewReplyTo", mailReview.replyTo || "Not available");
  const body = byId("mailReviewBody");
  if (body) body.textContent = mailReview.body || "";
  setText("mailReviewStatus", "");
  byId("mailReviewDialog")?.showModal();
}

function renderMailTemplatePanel() {
  const providerField = byId("providerIntroTemplate");
  const curatorField = byId("curatorForwardTemplate");
  const solutionSeekerField = byId("solutionSeekerTemplate");
  const seekerField = byId("needSeekerTemplate");
  const autoSyncField = byId("inboundAutoSyncEnabled");
  const lshContactEmailsField = byId("lshContactEmails");
  const lshHelpCcEmailsField = byId("lshHelpCcEmails");
  const lshRequestSupportField = byId("lshRequestSupportTemplate");
  const lshEmailProviderField = byId("lshEmailProviderTemplate");
  if (providerField) providerField.value = state.mailTemplates.providerIntroTemplate || "";
  if (curatorField) curatorField.value = state.mailTemplates.curatorForwardTemplate || "";
  if (solutionSeekerField) solutionSeekerField.value = state.mailTemplates.solutionSeekerTemplate || "";
  if (seekerField) seekerField.value = state.mailTemplates.needSeekerTemplate || "";
  if (autoSyncField) autoSyncField.checked = Boolean(state.mailTemplates.inboundAutoSyncEnabled);
  if (lshContactEmailsField) lshContactEmailsField.value = ensureList(state.mailTemplates.lshContactEmails).join(", ");
  if (lshHelpCcEmailsField) lshHelpCcEmailsField.value = ensureList(state.mailTemplates.lshHelpCcEmails).join(", ");
  if (lshRequestSupportField) lshRequestSupportField.value = state.mailTemplates.lshRequestSupportTemplate || "";
  if (lshEmailProviderField) lshEmailProviderField.value = state.mailTemplates.lshEmailProviderTemplate || "";
}

function formatAuditDateTime(value) {
  const date = value ? new Date(value) : null;
  if (!date || Number.isNaN(date.getTime())) return "Not recorded";
  return date.toLocaleString("en-IN", {
    dateStyle: "medium",
    timeStyle: "short",
  });
}

function normalizeAuditSearch(entry) {
  return [
    entry.surface,
    entry.action,
    entry.actor_email,
    entry.actor_name,
    entry.sender_email,
    entry.sender_name,
    entry.recipient_email,
    entry.cc_email,
    entry.reply_to,
    entry.subject,
    entry.item_id,
    entry.item_label,
    entry.item_source,
    entry.need_id,
    entry.provider_name,
    entry.seeker_name,
    entry.portal_url,
    entry.source_page,
  ].map((value) => normalizeText(value)).join(" ");
}

function computeAuditSearchResults(entries, query) {
  const normalizedQuery = normalizeText(query);
  if (!normalizedQuery) return entries;
  return ensureList(entries).filter((entry) => normalizeAuditSearch(entry).includes(normalizedQuery));
}

function renderImpactAuditPanels() {
  const searchInput = byId("impactAuditSearch");
  const emailMeta = byId("impactEmailLogMeta");
  const viewMeta = byId("impactViewLogMeta");
  const emailList = byId("impactEmailLogList");
  const viewList = byId("impactViewLogList");
  if (!emailList || !viewList) return;

  const query = normalizeText(searchInput?.value || state.adminAuditSearch || "");
  if (searchInput && searchInput.value !== state.adminAuditSearch) {
    state.adminAuditSearch = searchInput.value;
  }

  const emailEntries = computeAuditSearchResults(ensureList(state.data.impactAuditLogs?.emailLogs), query);
  const viewEntries = computeAuditSearchResults(ensureList(state.data.impactAuditLogs?.viewLogs), query).map((entry) => {
    const ledToEmail = ensureList(state.data.impactAuditLogs?.emailLogs).some((mail) =>
      normalizeText(mail.sender_email || mail.sent_by_email).toLowerCase() === normalizeText(entry.actor_email).toLowerCase() &&
      (
        (normalizeText(mail.item_id) && normalizeText(mail.item_id) === normalizeText(entry.item_id)) ||
        (normalizeText(mail.item_label) && normalizeText(mail.item_label) === normalizeText(entry.item_label))
      ) &&
      new Date(mail.logged_at || mail.sent_at || 0).getTime() >= new Date(entry.logged_at || 0).getTime()
    );
    return { ...entry, _ledToEmail: ledToEmail };
  });

  if (emailMeta) {
    emailMeta.textContent = `${emailEntries.length} email log${emailEntries.length === 1 ? "" : "s"}${query ? " match the current search." : " available."}`;
  }
  if (viewMeta) {
    viewMeta.textContent = `${viewEntries.length} view log${viewEntries.length === 1 ? "" : "s"}${query ? " match the current search." : " available."}`;
  }

  emailList.innerHTML = emailEntries.length
    ? emailEntries.map((entry) => `
        <article class="stack-card">
          <div class="status-row">
            <span class="status-pill good">${esc(entry.surface || "surface")}</span>
            <span class="meta-text">${esc(formatAuditDateTime(entry.logged_at || entry.sent_at))}</span>
          </div>
          <p><strong>Action:</strong> ${esc(entry.action || "email")}</p>
          <p><strong>From:</strong> ${esc(entry.sender_name || entry.sender_email || "Unknown")} | ${esc(entry.sender_email || "No email")}</p>
          <p><strong>To:</strong> ${esc(entry.recipient_email || "Not recorded")}</p>
          <p><strong>Cc:</strong> ${esc(entry.cc_email || "None")}</p>
          <p><strong>Subject:</strong> ${esc(entry.subject || "Not recorded")}</p>
          <p><strong>Item:</strong> ${esc(entry.item_label || entry.provider_name || entry.seeker_name || "Not recorded")}</p>
        </article>
      `).join("")
    : `<div class="empty-state">No email logs match the current filter.</div>`;

  viewList.innerHTML = viewEntries.length
    ? viewEntries.map((entry) => `
        <article class="stack-card">
          <div class="status-row">
            <span class="status-pill info">${esc(entry.surface || "surface")}</span>
            <span class="meta-text">${esc(formatAuditDateTime(entry.logged_at))}</span>
          </div>
          <p><strong>Action:</strong> ${esc(entry.action || "view")}</p>
          <p><strong>Viewer:</strong> ${esc(entry.actor_name || entry.actor_email || "Anonymous")} | ${esc(entry.actor_email || "No email")}</p>
          <p><strong>Solution Viewed:</strong> ${esc(entry.item_label || entry.item_id || "Not recorded")}</p>
          <p><strong>Portal / Source:</strong> ${esc(entry.item_source || entry.portal_url || "Not recorded")}</p>
          <p><strong>Led to Email:</strong> ${entry._ledToEmail ? "Yes" : "No"}</p>
        </article>
      `).join("")
    : `<div class="empty-state">No view logs match the current filter.</div>`;
}

function getCachedMatchByOfferingId(offeringId) {
  const need = getNeedById(state.selectedNeedId);
  const matches = state.matchCache.get(getNeedMatchCacheKey(need)) || [];
  return matches.find((item) => normalizeLooseId(item.offering_id) === normalizeLooseId(offeringId)) || null;
}

function getManualMatchByOfferingId(offeringId) {
  return ensureList(state.manualSolutionSearch.results)
    .find((item) => normalizeLooseId(item.offering_id) === normalizeLooseId(offeringId)) || null;
}

function getMatchForActionButton(button) {
  if (button?.closest?.("#manualMatchResults")) {
    return getManualMatchByOfferingId(button.dataset.offeringId);
  }
  return getCachedMatchByOfferingId(button.dataset.offeringId);
}

function buildMatchCardHtml(match, need) {
  const email = parseContactDetailParts(match.contact_details).email || match.trader?.email || "";
  const matchSummary = normalizeText(match.about_offering_text || match.solution?.about_solution_text || "");
  const truncatedSummary = matchSummary.slice(0, 280);
  const hasMore = matchSummary.length > 280;
  const emailAction = getMatchEmailActionConfig(need, match);
  const seekerEmailAction = getMatchSeekerEmailActionConfig(need, match);
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
        ${match.gre_link ? `<a class="btn btn-secondary" href="${esc(match.gre_link)}" target="_blank" rel="noreferrer" data-action="open-gre-link" data-offering-id="${esc(match.offering_id)}">Open GRE Link</a>` : `<span></span>`}
        ${window.APP_CONFIG?.ENABLE_EMAIL_ACTIONS && (hasAdminLikeAccess() || isCuratorUser()) && email
          ? `<button class="btn btn-primary" data-action="email-provider" data-provider-email="${esc(email)}" data-provider-name="${esc(match.trader?.organisation_name || match.trader?.trader_name || "")}" data-offering-id="${esc(match.offering_id)}">${esc(emailAction.actionLabel)}</button>`
          : `<button class="btn btn-primary" type="button" disabled title="Provider outreach is available only to signed-in curators, moderators, and admins.">Email to Provider</button>`}
        ${window.APP_CONFIG?.ENABLE_EMAIL_ACTIONS && (hasAdminLikeAccess() || isCuratorUser()) && seekerEmailAction.to
          ? `<button class="btn btn-primary" data-action="email-seeker" data-offering-id="${esc(match.offering_id)}">${esc(seekerEmailAction.actionLabel)}</button>`
          : `<button class="btn btn-primary" type="button" disabled title="Seeker outreach is available only when a seeker email is stored for this need.">Email to Seeker</button>`}
      </div>
    </article>
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
  applySharedSolutionAnnotations(need, matches);
  matches = [...matches].sort((left, right) => {
    const leftShared = ensureList(left.matchReasons).some((reason) => normalizeText(reason).toLowerCase() === "already shared with seeker");
    const rightShared = ensureList(right.matchReasons).some((reason) => normalizeText(reason).toLowerCase() === "already shared with seeker");
    if (leftShared !== rightShared) return rightShared - leftShared;
    return parseNumber(right.matchScore, 0) - parseNumber(left.matchScore, 0);
  });
  state.matchCache.set(cacheKey, matches);
  const exactSharedMatchCount = matches.filter((match) =>
    ensureList(match.matchReasons).some((reason) => normalizeText(reason).toLowerCase() === "already shared with seeker"),
  ).length;
  if (exactSharedMatchCount !== parseNumber(need._derived_shared_match_count, 0)) {
    need._derived_shared_match_count = exactSharedMatchCount;
    if (String(state.selectedNeedId) === String(need.id)) {
      renderNeedDetail();
    }
  }
  const pageSize = 6;
  const totalPages = Math.max(1, Math.ceil(matches.length / pageSize));
  if (state.matchPage > totalPages) state.matchPage = totalPages;
  const pageStart = (state.matchPage - 1) * pageSize;
  const visibleMatches = matches.slice(pageStart, pageStart + pageSize);

  matchesEl.innerHTML = matches.length
    ? visibleMatches
        .map((match) => buildMatchCardHtml(match, need))
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

function buildSelectWithCurrent(currentValue, options) {
  const current = normalizeText(currentValue);
  const values = uniq([current, ...options.map((item) => normalizeText(item))]).filter(Boolean);
  return [
    `<option value="">Current: ${esc(current || "Not set")}</option>`,
    ...values.map((value) => `<option value="${esc(value)}">${esc(value)}</option>`),
  ].join("");
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
  const currentCurationNotes = normalizeText(need.curation_notes);
  const currentCuratedNeedValues = parseArray(need.curated_need).map((item) => normalizeText(item).toLowerCase());
  const currentCuratedNeedDisplay = parseArray(need.curated_need).filter((item) => normalizeText(item));
  const curatorOptions = getCuratorAssignmentOptions(need);
  const canEdit = canEditNeedCuration(need);
  const isCurator = isCuratorUser();
  const isAdmin = isAdminUser();
  const isModerator = isModeratorUser();
  const hasAdminLike = hasAdminLikeAccess();
  const canMessageCurator = isLoggedIn() && !canEdit;
  const localOnlyNeed = isLocalOnlyNeed(need);

  workbench.innerHTML = `
    ${hasAdminLike ? `
      <article class="action-card">
        <p class="eyebrow">${isAdmin ? "Admin Assignment" : "Moderator Assignment"}</p>
        <h4>Curator Allocation</h4>
        <label>
          <span>Assigned Curator</span>
          <select id="assignCuratorSelect">${curatorOptions}</select>
        </label>
        <button class="btn btn-secondary" id="assignCuratorBtn">Save Curator Assignment</button>
        <p class="helper-text">${localOnlyNeed
          ? "This GramEEE-origin need can be assigned locally to any active admin, moderator, or curator. Its curation will remain on GramEEE only."
          : isAdmin
            ? "Only GRE-registered curators, moderators, or admins can be assigned here so sync-back to the GRE website remains intact."
            : "Moderators can assign and rebalance curator ownership, and only GRE-registered curators, moderators, or admins are valid for GRE-synced needs."}</p>
      </article>` : ""}

    ${(canEdit || hasAdminLike) ? `
      <article class="action-card">
        <p class="eyebrow">${hasAdminLike ? `${isAdmin ? "Admin" : "Moderator"} Curation Update` : "Curator Update"}</p>
        <h4>${hasAdminLike ? "Update Curation Directly" : "Update Your Assigned Need"}</h4>
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
            <span>Curated Need of Seeker</span>
            <div id="curatedNeedEditor" style="display:grid;gap:8px;">
              <div id="curatedNeedChips" class="tag-cloud tag-chip-editor" style="min-height:36px;">
                ${currentCuratedNeedDisplay.length ? currentCuratedNeedDisplay.map((item) => `
                  <span class="tag-chip" style="display:inline-flex;align-items:center;gap:6px;padding:7px 11px;border-radius:999px;background:var(--green-soft);color:var(--green);font-size:0.8rem;font-weight:700;">
                    ${esc(item)}
                    <button type="button" class="tag-chip-remove" data-remove-curated-need="${escAttr(item)}" style="border:none;background:transparent;color:inherit;cursor:pointer;padding:0;font:inherit;line-height:1;" aria-label="Remove ${escAttr(item)}">&times;</button>
                  </span>
                `).join("") : '<span class="helper-text">No curated need categories selected.</span>'}
              </div>
              <div style="display:flex;gap:8px;">
                <input id="curatedNeedEntry" list="curatedNeedOptions" placeholder="Add a curated need category..." style="flex:1;" />
                <button type="button" id="addCuratedNeedBtn" class="btn btn-secondary btn-compact">Add</button>
              </div>
              <datalist id="curatedNeedOptions">
                ${CURATED_NEED_OPTIONS.map((opt) => `<option value="${esc(opt)}">${esc(opt)}</option>`).join("")}
              </datalist>
              <div id="curatedNeedHiddenInputs" style="display:none;">
                ${currentCuratedNeedDisplay.map((item) => `<input type="hidden" name="proposed_curated_need" value="${escAttr(item)}" />`).join("")}
              </div>
              <small class="helper-text">Select one or more GRE curated need categories.</small>
            </div>
          </label>
          <label class="wide">
            <span>Curation Notes</span>
            <textarea name="proposed_curation_notes" rows="4" placeholder="Describe what changed and why.">${esc(currentCurationNotes)}</textarea>
          </label>
          <label>
            <span>Broadcast Needed</span>
            <select name="proposed_demand_broadcast_needed">
              <option value="">Current: ${need.demand_broadcast_needed ? "Yes" : "No"}</option>
              <option value="true">Yes</option>
              <option value="false">No</option>
            </select>
          </label>
          <label>
            <span>Funding Mechanism</span>
            <select name="proposed_funding_mechanism">
              ${buildSelectWithCurrent(need.funding_mechanism, ["Grant", "CSR", "Government Scheme", "Loan / Credit", "Equity / Investment", "Seeker Paid", "Provider Supported", "Not Decided"])}
            </select>
          </label>
          <label>
            <span>Seeker / Provider Agreement</span>
            <select name="proposed_seeker_provider_agreement">
              ${buildSelectWithCurrent(need.seeker_provider_agreement, ["Not Started", "Discussion Initiated", "Agreement Pending", "Agreement Completed", "No Agreement"])}
            </select>
          </label>
          <label>
            <span>Solution Deployment Status</span>
            <input name="proposed_solution_deployment_status" value="${esc(need.solution_deployment_status || "")}" placeholder="Not started, in progress, deployed..." />
          </label>
          <label>
            <span>Closure Date</span>
            <input name="proposed_closure_date" type="date" value="${esc(need.closure_date || "")}" />
          </label>
          <label class="wide">
            <span>Feedback about Seeker</span>
            <textarea name="proposed_feedback_about_seeker" rows="3" placeholder="Capture seeker-side feedback">${esc(need.feedback_about_seeker || "")}</textarea>
          </label>
          <label class="wide">
            <span>Feedback about Provider</span>
            <textarea name="proposed_feedback_about_provider" rows="3" placeholder="Capture provider-side feedback">${esc(need.feedback_about_provider || "")}</textarea>
          </label>
          <div class="wide">
            <button class="btn btn-primary" type="submit">${localOnlyNeed ? "Save Local Curation Update" : isAdmin ? "Save Curation Update" : "Save and Sync to GRE"}</button>
            <p class="helper-text">${localOnlyNeed
              ? "This need originated on GramEEE. Its curator updates stay on GramEEE and do not sync back to the GRE website."
              : isAdmin
                ? "Admin can directly update any curator's curation details."
                : "Your curation update will sync to the GRE website immediately."}</p>
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

function buildPuterSuggestedQuestionsPrompt(need) {
  return `
You are preparing a curator for a rural need-assessment call.
Return only valid JSON.

Need record:
${JSON.stringify({
  id: need.id,
  organization_name: need.organization_name,
  state: need.state,
  district: need.district,
  thematic_area: need.submitted_thematic_area || need.ai_thematic_area,
  need_category: need.submitted_offering_category || need.ai_need_kind,
  need_type: need.submitted_offering_type || need.ai_service_kind,
  keywords: parseArray(need.submitted_keywords),
  problem_statement: need.problem_statement,
  curation_notes: need.curation_notes,
}, null, 2)}

Return:
{
  "questions": ["", "", "", "", ""]
}

Rules:
- return exactly 5 probing questions
- make them practical for a curator call
- prioritize deployment context, pain point, budget, maintenance, and scale/adoption
- each question should be concise
- do not include numbering in the JSON
`.trim();
}

function buildPuterNeedKeywordPrompt(draft) {
  return `
You are helping an admin prepare strong solution-matching keywords for a rural-economy need.
Return only valid JSON.

Need draft:
${JSON.stringify({
  organization_name: draft.organization_name,
  offering_category: draft.offering_category,
  offering_type: draft.offering_type,
  thematic_area: draft.thematic_area,
  deployment_locations: draft.deployment_locations,
  problem_statement: draft.problem_statement,
}, null, 2)}

Return:
{
  "tags": ["", "", ""]
}

Rules:
- return 8 to 12 keywords
- preserve meaningful multi-word phrases
- prioritize the actual need, deployment context, constraints, and desired solution qualities
- avoid filler words, generic verbs, or introductory statistics
- each keyword should be 1 to 4 words
- do not include numbering
`.trim();
}

function buildPuterSolutionTagPrompt(draft) {
  return `
You are helping classify a solution offering for AskGRE search discovery.
Return only valid JSON.

Offering draft:
${JSON.stringify({
  organization_name: draft.organization_name,
  offering_category: draft.offering_category,
  offering_type: draft.offering_type,
  offering_name: draft.offering_name,
  about_offering_text: draft.about_offering_text,
  trainer_details_text: draft.trainer_details_text,
  contact_details: draft.contact_details,
}, null, 2)}

Return:
{
  "tags": ["", "", ""]
}

Rules:
- return 8 to 12 thematic tags
- prioritize use-case, thematic area, user segment, deployment context, technical capability, and commodity or domain
- preserve meaningful multi-word phrases
- avoid generic words like service, solution, support, offering, provider, training unless part of a more specific phrase
- avoid repeating individual words from the description as separate tags
- each tag should be 1 to 4 words
- do not include numbering
`.trim();
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

async function runPuterSuggestedQuestions(needId) {
  const need = state.data.needs.find((item) => item.id === needId) || state.data.localNeeds.find((item) => item.id === needId) || state.data.aiReviewNeeds.find((item) => item.id === needId);
  if (!need) throw new Error("Need not found for Puter question generation.");
  const text = await runPuterChat(buildPuterSuggestedQuestionsPrompt(need));
  const parsed = parseJsonObject(text);
  const questions = parseArray(parsed?.questions).map((item) => normalizeText(item)).filter((item) => item.length >= 12).slice(0, 5);
  if (!questions.length) {
    throw new Error("Puter did not return usable suggested questions.");
  }
  const result = await store.saveNeedSuggestedQuestions(needId, questions, "puter");
  await refreshAll();
  return result;
}

async function runPuterNeedKeywordAssist(draft) {
  const text = await runPuterChat(buildPuterNeedKeywordPrompt(draft));
  const parsed = parseJsonObject(text);
  const tags = parseArray(parsed?.tags)
    .map((item) => normalizeText(item).toLowerCase())
    .filter((item) => item.length >= 3)
    .slice(0, 12);
  if (!tags.length) {
    throw new Error("Puter did not return usable keywords.");
  }
  return uniq(tags);
}

async function runPuterSolutionTagAssist(draft) {
  const text = await runPuterChat(buildPuterSolutionTagPrompt(draft));
  const parsed = parseJsonObject(text);
  const tags = parseArray(parsed?.tags)
    .map((item) => normalizeText(item))
    .filter((item) => item.length >= 3)
    .filter((item) => {
      const lowered = item.toLowerCase();
      return ![
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
      ].includes(lowered);
    })
    .slice(0, 12);
  if (!tags.length) {
    throw new Error("Puter did not return usable thematic tags.");
  }
  return uniq(tags);
}

function getLocalSolutionTagDraft() {
  const form = byId("localSolutionReviewForm");
  if (!form) return null;
  return {
    organization_name: normalizeText(form.querySelector('[name="organization_name"]')?.value || ""),
    offering_category: normalizeText(form.querySelector('[name="offering_category"]')?.value || ""),
    offering_type: normalizeText(form.querySelector('[name="offering_type"]')?.value || ""),
    offering_name: normalizeText(form.querySelector('[name="offering_name"]')?.value || ""),
    solution_name: normalizeText(form.querySelector('[name="solution_name"]')?.value || ""),
    about_solution_text: normalizeText(form.querySelector('[name="about_solution_text"]')?.value || ""),
    about_offering_text: normalizeText(form.querySelector('[name="about_offering_text"]')?.value || ""),
    primary_valuechain: "",
    primary_application: "",
    geographies: normalizeText(form.querySelector('[name="geographies"]')?.value || ""),
  };
}

function getLocalNeedKeywordDraft() {
  const form = byId("localNeedReviewForm");
  if (!form) return null;
  return {
    organization_name: normalizeText(form.querySelector('[name="organization_name"]')?.value || ""),
    offering_category: normalizeText(form.querySelector('[name="submitted_offering_category"]')?.value || ""),
    offering_type: normalizeText(form.querySelector('[name="submitted_offering_type"]')?.value || ""),
    thematic_area: normalizeText(form.querySelector('[name="submitted_thematic_area"]')?.value || ""),
    problem_statement: normalizeText(form.querySelector('[name="problem_statement"]')?.value || ""),
    deployment_locations: normalizeText(form.querySelector('[name="deployment_locations"]')?.value || ""),
  };
}

function mergeCommaListValues(currentValue, nextValues) {
  return uniq([
    ...parseArray(currentValue),
    ...parseArray(nextValues),
  ].map((item) => normalizeText(item)).filter(Boolean)).join(", ");
}

function renderAdminQueue() {
  const pendingNeedsList = byId("pendingNeedsList");
  const pendingUpdatesList = byId("pendingUpdatesList");
  const pendingFormSubmissionsList = byId("pendingFormSubmissionsList");
  const aiReviewList = byId("aiReviewList");
  if (!pendingNeedsList || !pendingUpdatesList || !pendingFormSubmissionsList || !aiReviewList) return;

  pendingNeedsList.innerHTML = hasAdminLikeAccess()
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
                  ${canRejectRecords() ? `<button class="btn btn-danger" data-action="reject-need" data-need-id="${esc(need.id)}">Reject</button>` : ""}
                </div>
              </article>
            `,
          )
          .join("")
      : `<div class="empty-state">No intake records are waiting for admin approval.</div>`
    : `<div class="empty-state">Login as admin or moderator to view the intake approval queue.</div>`;

  pendingUpdatesList.innerHTML = hasAdminLikeAccess()
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
                  ${canRejectRecords() ? `<button class="btn btn-danger" data-action="reject-update" data-request-id="${esc(request.id)}">Reject</button>` : ""}
                </div>
              </article>
            `,
          )
          .join("")
      : `<div class="empty-state">No curator update requests are waiting for approval.</div>`
      : `<div class="empty-state">Login as admin or moderator to view curator-submitted status changes.</div>`;

  pendingFormSubmissionsList.innerHTML = hasAdminLikeAccess()
    ? ensureList(state.data.pendingFormSubmissions).length
      ? ensureList(state.data.pendingFormSubmissions)
            .map((submission) => {
              const payload = submission.payload && typeof submission.payload === "object" ? submission.payload : {};
              const isNeed = submission.submission_type === "need";
              const summaryText = isNeed
                ? normalizeText(payload.problem_statement || "")
                : normalizeText(payload.about_offering_text || payload.about_solution_text || payload.solution_name || payload.offering_name || "");
                const categoryLine = isNeed
                  ? [payload.offering_category, payload.offering_type, payload.thematic_area].filter(Boolean).join(" / ")
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
                  ${canRejectRecords() ? `<button class="btn btn-danger" data-action="reject-form-submission" data-submission-id="${esc(submission.id)}">Reject</button>` : ""}
                </div>
              </article>
            `;
          })
          .join("")
      : `<div class="empty-state">No shared solution or need forms are waiting for admin review.</div>`
    : `<div class="empty-state">Login as admin or moderator to review shared form submissions.</div>`;

  aiReviewList.innerHTML = hasAdminLikeAccess()
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
                        <button class="btn btn-secondary" data-action="generate-suggested-questions" data-need-id="${esc(need.id)}">Generate Questions</button>
                        <button class="btn btn-secondary" data-action="generate-suggested-questions-puter" data-need-id="${esc(need.id)}">Use Puter</button>
                        <button class="btn btn-primary" data-action="accept-need-override" data-need-id="${esc(need.id)}" data-override-field="all">Accept All</button>
                      </div>
                    </div>
                  `
                  : `<div class="card-actions"><button class="btn btn-secondary" data-action="run-puter-need-review" data-need-id="${esc(need.id)}">Puter Suggest Review</button><button class="btn btn-secondary" data-action="generate-suggested-questions" data-need-id="${esc(need.id)}">Generate Questions</button><button class="btn btn-secondary" data-action="generate-suggested-questions-puter" data-need-id="${esc(need.id)}">Use Puter</button></div>`;
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
    : `<div class="empty-state">Login as admin or moderator to view the AI review queue.</div>`;

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
    ["moderator", "Moderator"],
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
            <h4>${esc(resolveMisDisplayName(user))}</h4>
            <div class="detail-list user-detail-grid">
              <div><strong>Username</strong><span>${esc(user.username)}</span></div>
              <div><strong>Email</strong><span>${esc(user.email || "Not set")}</span></div>
              <div><strong>Phone</strong><span>${esc(user.phone || "Not set")}</span></div>
              <div><strong>Role</strong><span>${esc(user.role || "Not set")}</span></div>
            </div>
            <div class="detail-list user-detail-grid user-sync-grid">
              <div><strong>GRE Login</strong><span>${esc(user.gre_login_name || "Not mapped")}</span></div>
              <div><strong>GRE User ID</strong><span>${esc(user.gre_user_id || "Not mapped")}</span></div>
              <div><strong>GRE Sync</strong><span>${esc(user.gre_sync_status ? user.gre_sync_status.replaceAll("_", " ") : "Not attempted")}</span></div>
              <div><strong>Synced At</strong><span>${esc(user.gre_synced_at ? formatDate(user.gre_synced_at) : "Not synced")}</span></div>
            </div>
            ${user.gre_pending_role ? `
              <div class="detail-list user-detail-grid user-sync-grid">
                <div><strong>Pending GRE Role</strong><span>${esc(user.gre_pending_role)}</span></div>
                <div><strong>Activation Requested</strong><span>${esc(user.gre_activation_requested_at ? formatDate(user.gre_activation_requested_at) : "Not started")}</span></div>
                <div><strong>OTP Status</strong><span>${esc(user.gre_sync_status === "awaiting_gre_activation" ? "Awaiting OTP" : "Not required")}</span></div>
                <div><strong>GRE Temp Password</strong><span>${esc(user.gre_pending_role ? "gre@1234" : "-")}</span></div>
              </div>
            ` : ""}
            ${user.gre_sync_message ? `<p class="helper-text">${esc(user.gre_sync_message)}</p>` : ""}
            <div class="card-actions">
              <label class="inline-select">
                <span class="helper-text">Assign role</span>
                <select data-role-select data-user-id="${esc(user.id)}">
                  ${["user", "curator", "moderator", "admin"].map((role) => `<option value="${role}" ${user.role === role ? "selected" : ""}>${role[0].toUpperCase()}${role.slice(1)}</option>`).join("")}
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

async function openLocalSolutionReviewDialog(offeringId) {
  const summary = getLocalSolutionByOfferingId(offeringId);
  if (!summary) {
    toast("This local solution is no longer available.");
    return;
  }
  const detailResponse = await store.getLocalSolutionDetail(offeringId);
  const solution = detailResponse?.offering;
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
  const offeringPayload = getStoredPayloadRecord(solution.raw_payload);
  const solutionPayload = getStoredPayloadRecord(linkedSolution.raw_payload);
  const storedPayload = Object.keys(offeringPayload).length ? offeringPayload : solutionPayload;
  const durationParts = splitCombinedUnitValue(solution.duration, ["Days", "Hours", "Months", "Weeks"]);
  const serviceCostParts = splitCombinedUnitValue(solution.service_cost, [
    "Can be quoted after finalising scope",
    "Per day",
    "Per hour",
    "Per person",
  ]);
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
  setValue("organization_name", storedPayload.organization_name || provider.organisation_name || provider.trader_name || "");
  setValue("submitter_name", storedPayload.submitter_name || provider.poc_name || "");
  setValue("submitter_email", storedPayload.submitter_email || provider.email || "");
  setValue("submitter_phone", storedPayload.submitter_phone || provider.mobile || "");
  setValue("solution_name", linkedSolution.solution_name || "");
  setValue("about_solution_text", linkedSolution.about_solution_text || "");
  setValue("offering_category", solution.offering_category || "");
  setValue("offering_type", solution.offering_type || "");
  setValue("offering_name", solution.offering_name || "");
  setValue("about_offering_text", solution.about_offering_text || "");
  setValue("trainer_name", storedPayload.trainer_name || solution.trainer_name || "");
  setValue("trainer_email", storedPayload.trainer_email || solution.trainer_email || "");
  setValue("trainer_phone", storedPayload.trainer_phone || solution.trainer_phone || "");
  setValue("trainer_details_text", storedPayload.trainer_details_text || solution.trainer_details_text || "");
  setValue("languages", canonicalizeLanguageArray(parseArray(solution.languages)).join(", "));
  setValue("geographies", normalizeSemicolonSeparatedValue(storedPayload.geographies || solution.geographies).join("; "));
  setValue("duration", storedPayload.duration || durationParts.value || "");
  setValue("duration_unit", storedPayload.duration_unit || durationParts.unit || "");
  setValue("prerequisites", storedPayload.prerequisites || solution.prerequisites || "");
  setValue("location_availability", parseArray(storedPayload.location_availability || solution.location_availability).join(", "));
  setValue("service_cost", storedPayload.service_cost || serviceCostParts.value || "");
  setValue("service_cost_unit", storedPayload.service_cost_unit || serviceCostParts.unit || "");
  setValue("cost_remarks", storedPayload.cost_remarks || solution.cost_remarks || "");
  setValue("support_post_service", storedPayload.support_post_service || solution.support_post_service || "");
  setValue("support_post_service_cost", storedPayload.support_post_service_cost || solution.support_post_service_cost || "");
  setValue("delivery_mode", storedPayload.delivery_mode || solution.delivery_mode || "");
  setValue("certification_offered", storedPayload.certification_offered || solution.certification_offered || "");
  setValue("product_cost", solution.product_cost || "");
  setValue("grade_capacity", storedPayload.grade_capacity || solution.grade_capacity || "");
  setValue("lead_time", solution.lead_time || "");
  setValue("support_details", solution.support_details || "");
  setValue("contact_details", storedPayload.contact_details || solution.contact_details || "");
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
  if (!hasAdminLikeAccess()) {
    list.innerHTML = `<div class="empty-state">Login as admin or moderator to manage local Supabase solutions.</div>`;
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
              ${canDeleteRecords() ? `<button class="btn btn-danger" type="button" data-action="delete-local-solution" data-offering-id="${esc(item.offering_id)}">Delete</button>` : ""}
            </div>
          </article>
        `;
      }).join("")}</div>`
    : `<div class="empty-state">No local Supabase-only solutions match the current filters.</div>`;
}

function getLocalNeedById(needId) {
  return ensureList(state.data.localNeeds).find((item) => String(item.id) === String(needId)) || null;
}

function getLocalNeedCategories() {
  return uniq(
    ensureList(state.data.localNeeds)
      .map((item) => normalizeText(getEffectiveSubmittedNeedCategory(item)))
      .filter(Boolean),
  ).sort((a, b) => a.localeCompare(b));
}

function getFilteredLocalNeeds() {
  const search = normalizeText(state.localNeedFilters.search).toLowerCase();
  const category = normalizeText(state.localNeedFilters.category);
  return ensureList(state.data.localNeeds).filter((item) => {
    const itemCategory = normalizeText(getEffectiveSubmittedNeedCategory(item));
    const haystack = [
      normalizeText(item.organization_name),
      normalizeText(item.contact_person),
      normalizeText(item.problem_statement),
      normalizeText(item.submitted_thematic_area),
      ...getEffectiveNeedKeywords(item),
      ...getEffectiveNeedDeploymentLocations(item),
    ].join(" ").toLowerCase();
    const categoryMatch = category === "all" || itemCategory === category;
    const searchMatch = !search || haystack.includes(search);
    return categoryMatch && searchMatch;
  });
}

async function openLocalNeedReviewDialog(needId) {
  const summary = getLocalNeedById(needId);
  if (!summary) {
    toast("This local need is no longer available.");
    return;
  }
  const detailResponse = await store.getLocalNeedDetail(needId);
  const need = detailResponse?.need;
  if (!need) {
    toast("This local need is no longer available.");
    return;
  }
  state.localNeedReviewId = needId;
  const dialog = byId("localNeedReviewDialog");
  const form = byId("localNeedReviewForm");
  if (!dialog || !form) return;
  const meta = byId("localNeedReviewMeta");
  if (meta) {
    meta.innerHTML = `
      <div><strong>Need ID:</strong> ${esc(need.id)}</div>
      <div><strong>Status:</strong> ${esc(need.status || "New")}</div>
      <div><strong>Internal Status:</strong> ${esc(need.internal_status || "Need solution providers")}</div>
      <div><strong>Source:</strong> ${esc(need.source_kind || "shared_form_submission")}</div>
    `;
  }
  const setValue = (name, value) => {
    const field = form.querySelector(`[name="${name}"]`);
    if (field) field.value = value || "";
  };
  setValue("need_id", need.id);
  setValue("organization_name", need.organization_name || "");
  setValue("contact_person", need.contact_person || "");
  setValue("seeker_email", need.seeker_email || "");
  setValue("seeker_phone", need.seeker_phone || "");
  setValue("submitted_offering_category", getEffectiveSubmittedNeedCategory(need));
  setValue("submitted_offering_type", getEffectiveSubmittedNeedType(need));
  setValue("submitted_thematic_area", need.submitted_thematic_area || need.ai_thematic_area || "");
  setValue("deployment_locations", getEffectiveNeedDeploymentLocations(need).join("; "));
  setValue("submitted_keywords", getEffectiveNeedKeywords(need).join(", "));
  setValue("problem_statement", need.problem_statement || "");
  byId("localNeedReviewStatus").textContent = "";
  dialog.showModal();
}

function renderLocalNeedManagement() {
  const list = byId("localNeedsList");
  const categorySelect = byId("localNeedCategoryFilter");
  const searchInput = byId("localNeedSearch");
  if (!list || !categorySelect || !searchInput) return;
  if (!hasAdminLikeAccess()) {
    list.innerHTML = `<div class="empty-state">Login as admin or moderator to manage local Supabase needs.</div>`;
    categorySelect.innerHTML = `<option value="all">All need categories</option>`;
    searchInput.value = "";
    return;
  }
  const categories = getLocalNeedCategories();
  categorySelect.innerHTML = [`<option value="all">All need categories</option>`, ...categories.map((item) => `<option value="${esc(item)}">${esc(item)}</option>`)].join("");
  categorySelect.value = state.localNeedFilters.category || "all";
  searchInput.value = state.localNeedFilters.search || "";
  const rows = getFilteredLocalNeeds();
  list.innerHTML = rows.length
    ? `<div class="local-solutions-grid">${rows.map((item) => `
          <article class="approval-card local-solution-card">
            <div class="status-row">
              <span class="status-pill info">${esc(getEffectiveSubmittedNeedCategory(item) || "Need")}</span>
              <span class="status-pill good">${esc(item.status || "New")}</span>
            </div>
            <h4>${esc(item.organization_name || "Unnamed organisation")}</h4>
            <div class="detail-list">
              <div><strong>Need Type:</strong> ${esc(getEffectiveSubmittedNeedType(item) || "Not set")}</div>
              <div><strong>Thematic Area:</strong> ${esc(item.submitted_thematic_area || item.ai_thematic_area || "Not set")}</div>
              <div><strong>Keywords:</strong> ${esc(getEffectiveNeedKeywords(item).join(", ") || "None")}</div>
              <div><strong>Places:</strong> ${esc(getEffectiveNeedDeploymentLocations(item).join(", ") || "Not set")}</div>
            </div>
            <p class="helper-text">${esc(clipText(item.problem_statement || "", 180) || "No statement available.")}</p>
            <div class="card-actions">
              <button class="btn btn-secondary" type="button" data-action="generate-suggested-questions" data-need-id="${esc(item.id)}">Generate Questions</button>
              <button class="btn btn-secondary" type="button" data-action="generate-suggested-questions-puter" data-need-id="${esc(item.id)}">Use Puter</button>
              <button class="btn btn-secondary" type="button" data-action="edit-local-need" data-need-id="${esc(item.id)}">Edit</button>
              ${canDeleteRecords() ? `<button class="btn btn-danger" type="button" data-action="delete-local-need" data-need-id="${esc(item.id)}">Delete</button>` : ""}
            </div>
          </article>
        `).join("")}</div>`
    : `<div class="empty-state">No local Supabase-only needs match the current filters.</div>`;
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
  const loginStatus = byId("userLoginStatus");
  const grameeeSummary = readGrameeeSummary();
  const sharedMode = isSharedFormMode();
  const loggedIn = isLoggedIn();
  const hasPortalAccess = hasMisAccessRole();
  const showShell = sharedMode || (loggedIn && hasPortalAccess);
  const isAdminDatasetVisible = hasAdminLikeAccess() && state.view === "admin";
  document.body.classList.toggle("public-need-form-mode", isStandalonePublicFormMode("need"));
  document.body.classList.toggle("public-solution-form-mode", isStandalonePublicFormMode("solution"));

  if (authGate) authGate.classList.toggle("hidden", showShell);
  if (appShell) appShell.classList.toggle("hidden", !showShell);
  masthead?.classList.toggle("hidden", sharedMode);
  heroStrip?.classList.toggle("hidden", sharedMode || !isAdminDatasetVisible);
  overviewTab?.classList.toggle("hidden", sharedMode);
  operationsTab?.classList.toggle("hidden", sharedMode || !canAccessOperationsDesk());
  solutionTab?.classList.add("hidden");
  needTab?.classList.add("hidden");
  roleTab?.classList.toggle("hidden", sharedMode || !hasAdminLikeAccess());
  if (newNeedBtn) newNeedBtn.classList.toggle("hidden", !loggedIn);
  if (sharedMode) {
    state.view = state.sharedFormMode;
  } else if (!loggedIn) {
    state.view = "overview";
  } else if (!hasPortalAccess) {
    state.view = "overview";
  } else if (!hasAdminLikeAccess() && state.view === "admin") {
    state.view = canAccessOperationsDesk() ? "operations" : "overview";
  } else if (!canAccessOperationsDesk() && state.view === "operations") {
    state.view = "overview";
  }

  if (accountName) {
    accountName.textContent = state.userSession
      ? (resolveMisDisplayName(state.userSession) || "GRE MIS User")
      : "GRE MIS User";
  }
  if (accountRoleChip) {
    accountRoleChip.textContent = state.userSession ? state.userSession.role : "User";
    accountRoleChip.className = `chip ${isAdminUser() ? "good" : isModeratorUser() ? "warn" : isCuratorUser() ? "info" : "muted"}`;
  }
  if (accountSummary) {
    accountSummary.textContent = state.userSession
      ? `${state.userSession.email || ""}${state.userSession.must_change_password ? " • Please change your default password." : ""}`
      : "";
  }
  if (!isLoggedIn()) {
    setText("changePasswordStatus", "");
  }
  if (loginStatus) {
    if (state.grameeeBridgeMessage && !sharedMode) {
      loginStatus.textContent = state.grameeeBridgeMessage;
    } else if (loggedIn && !hasPortalAccess && !sharedMode) {
      loginStatus.textContent = "This account does not yet have GRE MIS access. Please ask an admin to assign curator, moderator, or admin access.";
    } else if (!loggedIn && grameeeSummary && !["admin", "moderator", "curator"].includes(grameeeSummary.role) && !sharedMode) {
      loginStatus.textContent = "This GramEEE login does not yet have GRE access. Please ask an admin to assign admin, moderator, or curator access.";
    } else if (!loginStatus.dataset.lockedMessage) {
      loginStatus.textContent = "";
    }
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

  if (hasAdminLikeAccess()) {
    chip.textContent = "Unlocked";
    chip.className = `chip ${isAdminUser() ? "good" : "warn"}`;
    text.textContent = `${isAdminUser() ? "Admin" : "Moderator"} session active for ${state.userSession?.full_name || state.userSession?.username}. Intake approvals, sync controls, and local management tools are available below.`;
    sessionPanel?.classList.remove("hidden");
  } else {
    chip.textContent = "Locked";
    chip.className = "chip muted";
    text.textContent = "Admin or moderator approval is required for new intake records and curator-submitted status updates.";
    sessionPanel?.classList.add("hidden");
  }

  if (loginStatus && !hasAdminLikeAccess() && !loginStatus.textContent) {
    loginStatus.textContent = "";
  }
}

function renderAdminDeskTabs() {
  const activeTab = state.adminDeskTab || "data-sync";
  document.querySelectorAll("[data-admin-desk-tab]").forEach((button) => {
    const isActive = button.dataset.adminDeskTab === activeTab;
    button.classList.toggle("btn-primary", isActive);
    button.classList.toggle("btn-secondary", !isActive);
    button.classList.toggle("active", isActive);
  });
  document.querySelectorAll("[data-admin-desk-panel]").forEach((panel) => {
    panel.classList.toggle("hidden", panel.dataset.adminDeskPanel !== activeTab);
  });
}

function switchView(nextView) {
  if (isSharedFormMode()) {
    nextView = state.sharedFormMode;
  }
  if (nextView === "operations" && !canAccessOperationsDesk()) {
    nextView = "overview";
  }
  if (nextView === "admin" && !hasAdminLikeAccess()) {
    nextView = canAccessOperationsDesk() ? "operations" : "overview";
  }
  state.view = nextView;
  renderAuthState();
  renderAdminDeskTabs();
  if (nextView === "operations") {
    renderManualSolutionSearch();
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
  renderRequestTrackerControls();
  renderQueue();
  renderManualSolutionSearch();
  renderNeedDetail();
  renderWorkbench();
  renderAdminQueue();
  renderUserManagement();
  renderLocalSolutionManagement();
  renderLocalNeedManagement();
  renderAdminState();
  renderAdminDeskTabs();
  renderMailTemplatePanel();
  renderImpactAuditPanels();
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

async function refreshActiveAdminDeskTab() {
  if (hasAdminLikeAccess()) {
    await store.loadAdminDeskTabSnapshot(state.adminDeskTab || "approvals");
  }
  await rerender({ includeMatches: false });
}

async function refreshAll() {
  const resetAdminState = () => {
    state.data.pendingNeeds = [];
    state.data.pendingUpdates = [];
    state.data.pendingFormSubmissions = [];
    state.data.aiReviewNeeds = [];
    state.data.impactAuditLogs = { emailLogs: [], viewLogs: [] };
    state.data.users = [];
    state.data.localSolutions = [];
    state.data.localNeeds = [];
    state.mailTemplates = {
      providerIntroTemplate: DEFAULT_PROVIDER_INTRO_TEMPLATE,
      curatorForwardTemplate: DEFAULT_CURATOR_FORWARD_TEMPLATE,
      solutionSeekerTemplate: DEFAULT_SOLUTION_SEEKER_TEMPLATE,
      needSeekerTemplate: DEFAULT_NEED_SEEKER_TEMPLATE,
      inboundAutoSyncEnabled: true,
    };
  };
  const baseResults = await Promise.allSettled([store.loadBaseData()]);
  const failed = baseResults.filter((result) => result.status === "rejected");
  if (failed.length) {
    failed.forEach((result) => console.error("Dashboard refresh segment failed", result.reason));
    toast("Some dashboard sections could not refresh fully. Loaded sections remain available.");
  }
  if (hasAdminLikeAccess()) {
    const adminResults = await Promise.allSettled([
      store.loadAdminUsersSnapshot(),
      store.loadAdminDataSyncSnapshot(),
      store.loadAdminMailImpactSnapshot(),
      store.loadAdminLocalSolutionsSnapshot(),
      store.loadAdminLocalNeedsSnapshot(),
      store.loadAdminApprovalsSnapshot(),
    ]);
    const adminFailures = adminResults.filter((result) => result.status === "rejected");
    if (adminFailures.length) {
      adminFailures.forEach((result) => console.error("Management Desk segment failed", result.reason));
      toast("Some Management Desk tabs could not refresh fully. Available tabs remain usable.");
    }
  } else {
    resetAdminState();
  }
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
  resetManualSolutionSearch();
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
  entries.source_language = isHindiSharedFormMode() ? "hi" : "en";
  if (submissionType !== "solution") {
    const traderId = normalizeText(entries.existing_trader_id);
    const trader = getTraderById(traderId);
    if (!trader && !isEmbeddedSharedForm()) {
      openMissingOrgDialog(entries.organization_name || "");
      throw new Error("Please select an approved GRE supplier before submitting.");
    }
    const deploymentLocations = normalizeSemicolonSeparatedValue(entries.deployment_locations);
    const thematicArea = normalizeText(entries.thematic_area);
    const offeringCategory = normalizeText(entries.offering_category) || "Service offerings";
    const offeringType = normalizeText(entries.offering_type);
    const stateFromDeployment = deploymentLocations.length
      ? deploymentLocations
          .map((entry) => entry.split(",").map((item) => normalizeText(item)).filter(Boolean))
          .map((parts) => parts.length >= 2 ? parts[parts.length - 2] : "")
          .find(Boolean)
      : "";
    const districtFromDeployment = deploymentLocations.length
      ? deploymentLocations
          .map((entry) => entry.split(",").map((item) => normalizeText(item)).filter(Boolean))
          .map((parts) => parts.length >= 3 ? parts[parts.length - 3] : parts[0] || "")
          .find(Boolean)
      : "";
    return {
      ...entries,
      existing_trader_id: trader?.trader_id || "",
      existing_trader_name: trader?.organisation_name || trader?.trader_name || entries.organization_name || "",
      organization_name: entries.organization_name || trader?.organisation_name || trader?.trader_name || "Individual",
      offering_category: offeringCategory,
      offering_type: offeringType,
      need_kind: offeringCategory.replace(/\s+offerings$/i, "").toLowerCase(),
      service_kind: offeringCategory === "Service offerings" ? offeringType : "",
      thematic_area: thematicArea,
      keywords: [],
      deployment_locations: deploymentLocations,
      demand_broadcast_needed: normalizeText(entries.broadcast_to_ecosystem) === "true",
      state: stateFromDeployment || "",
      district: districtFromDeployment || "",
    };
  }

  const traderId = normalizeText(entries.existing_trader_id);
  const selectedTrader = getTraderById(traderId);
  const nameMatchedTrader = findSupplierByName(entries.organization_name || "");
  const trader = selectedTrader || nameMatchedTrader || null;
  const offeringCategory = normalizeText(entries.offering_category) || "Service offerings";
  const offeringGroup = offeringCategory.replace(/\s+offerings$/i, "").trim() || "Service";
  const defaultAudience = ["Individuals", "Groups", "SHGs", "Organisations"];
  const serviceLanguages = canonicalizeLanguageArray(normalizeMultiTextValue(entries.languages));
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
    button.addEventListener("click", safeAsync(async () => {
      const nextView = button.dataset.view;
      switchView(nextView);
      if (nextView === "admin" && hasAdminLikeAccess()) {
        await store.loadAdminDeskTabSnapshot(state.adminDeskTab || "data-sync");
        await rerender({ includeMatches: false });
      }
    }));
  });

  document.querySelectorAll("[data-admin-desk-tab]").forEach((button) => {
    button.addEventListener("click", safeAsync(async () => {
      state.adminDeskTab = button.dataset.adminDeskTab || "data-sync";
      renderAdminDeskTabs();
      if (hasAdminLikeAccess()) {
        await store.loadAdminDeskTabSnapshot(state.adminDeskTab);
      }
      await rerender({ includeMatches: false });
    }));
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
    resetManualSolutionSearch();
    renderQueue();
    renderNeedDetail();
    renderWorkbench();
    renderManualSolutionSearch();
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
    resetManualSolutionSearch();
    switchView("operations");
    renderQueue();
    renderNeedDetail();
    renderWorkbench();
    renderManualSolutionSearch();
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

  byId("manualSolutionSearchForm")?.addEventListener("submit", safeAsync(async (event) => {
    event.preventDefault();
    const need = getNeedById(state.selectedNeedId);
    if (!need) {
      toast("Select a need before searching for additional solutions.");
      return;
    }
    state.manualSolutionSearch = {
      provider: byId("manualSolutionProvider")?.value || "",
      category: byId("manualSolutionCategory")?.value || "",
      domain6m: byId("manualSolutionDomain6m")?.value || "",
      offeringType: byId("manualSolutionOfferingType")?.value || "",
      valuechain: byId("manualSolutionValuechain")?.value || "",
      application: byId("manualSolutionApplication")?.value || "",
      language: byId("manualSolutionLanguage")?.value || "",
      geography: byId("manualSolutionGeography")?.value || "",
      keyword: byId("manualSolutionKeyword")?.value.trim() || "",
      results: [],
      searched: true,
      loading: true,
      page: 1,
    };
    renderManualSolutionSearch();
    const results = await store.searchManualSolutionMatches(need, state.manualSolutionSearch);
    applySharedSolutionAnnotations(need, results);
    state.manualSolutionSearch = {
      ...state.manualSolutionSearch,
      results,
      searched: true,
      loading: false,
      page: 1,
    };
    renderManualSolutionSearch();
  }));

  byId("manualMatchResults")?.addEventListener("click", (event) => {
    const pageButton = event.target.closest("[data-manual-match-page-action]");
    if (!pageButton) return;
    if (pageButton.dataset.manualMatchPageAction === "prev" && state.manualSolutionSearch.page > 1) {
      state.manualSolutionSearch.page -= 1;
    }
    if (pageButton.dataset.manualMatchPageAction === "next") {
      state.manualSolutionSearch.page += 1;
    }
    renderManualSolutionSearch();
  });

  byId("manualSolutionClearBtn")?.addEventListener("click", () => {
    resetManualSolutionSearch();
    renderManualSolutionSearch();
  });

  byId("refreshBtn")?.addEventListener("click", safeAsync(async () => {
    resetDashboardSelections();
    await rerender({ includeMatches: false });
    if (byId("adminView") && hasAdminLikeAccess()) {
      const provider = byId("aiProviderSelect")?.value || "openai";
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
      resetManualSolutionSearch();
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
  const mailReviewDialog = byId("mailReviewDialog");
  const submissionReviewDialog = byId("submissionReviewDialog");
  const localSolutionReviewDialog = byId("localSolutionReviewDialog");
  const localNeedReviewDialog = byId("localNeedReviewDialog");
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
  byId("closeMailReviewDialog")?.addEventListener("click", closeMailReviewDialog);
  byId("cancelMailReviewBtn")?.addEventListener("click", closeMailReviewDialog);
  byId("closeSubmissionReviewDialog")?.addEventListener("click", () => submissionReviewDialog?.close());
  byId("closeLocalSolutionReviewDialog")?.addEventListener("click", () => localSolutionReviewDialog?.close());
  byId("closeLocalNeedReviewDialog")?.addEventListener("click", () => localNeedReviewDialog?.close());
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
      if (isEmbeddedSharedForm()) return;
      const trader = getTraderById(event.target.value);
      const orgInput = byId(orgInputId);
      if (orgInput && trader) orgInput.value = trader.organisation_name || trader.trader_name || "";
    });
    byId(orgInputId)?.addEventListener("blur", (event) => {
      if (isEmbeddedSharedForm()) return;
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

  document.addEventListener("click", (event) => {
    const button = event.target.closest('[data-action="add-on-behalf"]');
    if (!button || !isEmbeddedSharedForm()) return;
    postEmbeddedMessage("grameee-form-open-add-user", {
      mode: button.dataset.mode || state.sharedFormMode,
    });
  });

  byId("appShell")?.addEventListener("click", (event) => {
    const offeringChoiceButton = event.target.closest("[data-offering-scope][data-offering-category][data-offering-type]");
    if (!offeringChoiceButton) return;
    applyOfferingButtonSelection(
      offeringChoiceButton.dataset.offeringScope,
      offeringChoiceButton.dataset.offeringCategory,
      offeringChoiceButton.dataset.offeringType,
    );
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
      if (isEmbeddedSharedForm() && !state.embeddedActor) {
        requestEmbeddedLogin("solution");
        toast("Please log in with GramEEE to continue with the solution form.");
        return;
      }
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
        event.target.innerHTML = `<button type="button" class="btn btn-primary" id="solutionSubmissionReloadBtn">Solution Submitted for Approval</button>`;
        byId("solutionSubmissionReloadBtn")?.addEventListener("click", () => window.location.reload());
      } else {
        fillSupplierSelect("solutionTraderSelect", "solutionOrgName");
        renderSolutionReferenceInputs();
        if (status) status.textContent = result.message || "Solution submission sent to admin review.";
      }
      if (hasAdminLikeAccess()) await refreshAll();
      if (!isSharedFormMode()) {
        toast(result.message || "Solution Submitted for Approval");
      }
    }));

  byId("needSubmissionForm")?.addEventListener("submit", safeAsync(async (event) => {
      event.preventDefault();
      if (isEmbeddedSharedForm() && !state.embeddedActor) {
        requestEmbeddedLogin("need");
        toast("Please log in with GramEEE to continue with the help form.");
        return;
      }
      const status = byId("needSubmissionStatus");
      if (status) status.textContent = "Submitting help request for admin review...";
      const payload = await collectSubmissionPayload(event.target, "need");
      const result = isLoggedIn()
        ? await store.submitSignedInForm("need", payload)
        : await store.submitSharedForm("need", payload);
      event.target.reset();
      state.needTags = [];
      state.needDeploymentLocations = [];
      renderNeedTagChips();
      renderNeedDeploymentChips();
      if (isSharedFormMode()) {
        event.target.innerHTML = `
          <div class="stack-list">
            <button type="button" class="btn btn-primary" id="needSubmissionReloadBtn">Need Help Submitted for Approval</button>
            <p class="helper-text">${esc(result.message || "Need Help Submitted for Approval")}</p>
          </div>
        `;
        byId("needSubmissionReloadBtn")?.addEventListener("click", () => window.location.reload());
      } else {
        fillSupplierSelect("needTraderSelect", "needOrgName");
        renderNeedReferenceInputs();
        if (status) status.textContent = result.message || "Need Help Submitted for Approval";
      }
      if (hasAdminLikeAccess()) await refreshAll();
      if (!isSharedFormMode()) {
        toast(result.message || "Need Help Submitted for Approval");
      }
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
  byId("needOfferingCategory")?.addEventListener("change", () => updateNeedOfferingForm());
  byId("solutionPrimaryValuechain")?.addEventListener("change", () => updateSolutionApplicationOptions());
  byId("solutionSecondaryValuechain")?.addEventListener("change", () => updateSolutionApplicationOptions());
  byId("solutionGeographyEntry")?.addEventListener("input", safeAsync(async (event) => {
    await updateSolutionGeographySuggestions(event.target.value);
  }));
  byId("solutionGeographyEntry")?.addEventListener("focus", safeAsync(async () => {
    await updateSolutionGeographySuggestions(byId("solutionGeographyEntry")?.value || "");
  }));
  byId("addSolutionGeographyBtn")?.addEventListener("click", () => {
    const input = byId("solutionGeographyEntry");
    if (!input) return;
    if (addSolutionGeography(input.value)) {
      input.value = "";
      setDatalistOptions("solutionGeographyOptions", getSharedFormSeedGeographies());
      renderInlineSuggestions("solutionGeographySuggestions", getSharedFormSeedGeographies(), "data-add-solution-geo-suggestion");
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
    if (status) status.textContent = "Generating thematic tags from the current offering draft...";
      const draft = {
        organization_name: normalizeText(form.querySelector('[name="organization_name"]')?.value || ""),
        offering_category: normalizeText(form.querySelector('[name="offering_category"]')?.value || ""),
        offering_type: normalizeText(form.querySelector('[name="offering_type"]')?.value || ""),
        offering_name: normalizeText(form.querySelector('[name="offering_name"]')?.value || ""),
        about_offering_text: normalizeText(form.querySelector('[name="about_offering_text"]')?.value || ""),
        trainer_details_text: normalizeText(form.querySelector('[name="trainer_details_text"]')?.value || ""),
        contact_details: normalizeText(form.querySelector('[name="contact_details"]')?.value || form.querySelector('[name="product_contact_details"]')?.value || ""),
        tags: [...state.solutionTags],
        source_language: isHindiSharedFormMode() ? "hi" : "en",
      };
    if (!draft.offering_name || !draft.about_offering_text) {
      throw new Error("Please add the offering name and description before generating tags.");
    }
    const result = await store.suggestSolutionTags(draft);
    const generatedTags = parseArray(result.tags);
    const message = result.message || "AI-assisted thematic tags added. You can remove any tag before submitting.";
    const tagsToAdd = generatedTags.length ? generatedTags : deriveLocalSolutionTags(draft);
    tagsToAdd.forEach((tag) => addSolutionTag(tag));
    if (status) status.textContent = generatedTags.length
      ? message || "AI tags added. You can remove any tag before submitting."
      : "AI suggestions were unavailable, so local smart tags were added instead.";
  }));
  byId("generateSolutionTagsPuterBtn")?.addEventListener("click", safeAsync(async () => {
    const status = byId("solutionSubmissionStatus");
    const form = byId("solutionSubmissionForm");
    if (!form) return;
    if (status) status.textContent = "Generating Puter-assisted thematic tags from the current offering draft...";
      const draft = {
        organization_name: normalizeText(form.querySelector('[name="organization_name"]')?.value || ""),
        offering_category: normalizeText(form.querySelector('[name="offering_category"]')?.value || ""),
        offering_type: normalizeText(form.querySelector('[name="offering_type"]')?.value || ""),
        offering_name: normalizeText(form.querySelector('[name="offering_name"]')?.value || ""),
        about_offering_text: normalizeText(form.querySelector('[name="about_offering_text"]')?.value || ""),
        trainer_details_text: normalizeText(form.querySelector('[name="trainer_details_text"]')?.value || ""),
        contact_details: normalizeText(form.querySelector('[name="contact_details"]')?.value || form.querySelector('[name="product_contact_details"]')?.value || ""),
        tags: [...state.solutionTags],
        source_language: isHindiSharedFormMode() ? "hi" : "en",
      };
    if (!draft.offering_name || !draft.about_offering_text) {
      throw new Error("Please add the offering name and description before generating tags.");
    }
    const generatedTags = await runPuterSolutionTagAssist(draft);
    const tagsToAdd = generatedTags.length ? generatedTags : deriveLocalSolutionTags(draft);
    tagsToAdd.forEach((tag) => addSolutionTag(tag));
    if (status) status.textContent = generatedTags.length
      ? "Puter-assisted thematic tags added. You can remove any tag before submitting."
      : "Puter suggestions were unavailable, so local smart tags were added instead.";
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
      setDatalistOptions("solutionGeographyOptions", getSharedFormSeedGeographies());
      renderInlineSuggestions("solutionGeographySuggestions", getSharedFormSeedGeographies(), "data-add-solution-geo-suggestion");
    }
  });
  byId("needDeploymentEntry")?.addEventListener("input", safeAsync(async (event) => {
    await updateNeedDeploymentSuggestions(event.target.value);
  }));
  byId("needDeploymentEntry")?.addEventListener("focus", safeAsync(async () => {
    await updateNeedDeploymentSuggestions(byId("needDeploymentEntry")?.value || "");
  }));
  byId("addNeedDeploymentBtn")?.addEventListener("click", () => {
    const input = byId("needDeploymentEntry");
    if (!input) return;
    if (addNeedDeploymentLocation(input.value)) {
      input.value = "";
      setDatalistOptions("needDeploymentOptions", getSharedFormSeedGeographies());
      renderInlineSuggestions("needDeploymentSuggestions", getSharedFormSeedGeographies(), "data-add-need-deployment-suggestion");
    }
  });
  byId("needDeploymentEntry")?.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") return;
    event.preventDefault();
    const input = event.target;
    if (addNeedDeploymentLocation(input.value)) {
      input.value = "";
      setDatalistOptions("needDeploymentOptions", getSharedFormSeedGeographies());
      renderInlineSuggestions("needDeploymentSuggestions", getSharedFormSeedGeographies(), "data-add-need-deployment-suggestion");
    }
  });
  byId("addNeedTagBtn")?.addEventListener("click", () => {
    const input = byId("needTagEntry");
    if (!input) return;
    if (addNeedTag(input.value)) input.value = "";
  });
  byId("generateNeedTagsBtn")?.addEventListener("click", safeAsync(async () => {
    const status = byId("needSubmissionStatus");
    const form = byId("needSubmissionForm");
    if (!form) return;
    if (status) status.textContent = "Generating AI keywords from the current need draft...";
      const draft = {
        organization_name: normalizeText(form.querySelector('[name="organization_name"]')?.value || ""),
        offering_category: normalizeText(form.querySelector('[name="offering_category"]')?.value || ""),
        offering_type: normalizeText(form.querySelector('[name="offering_type"]')?.value || ""),
        thematic_area: normalizeText(form.querySelector('[name="thematic_area"]')?.value || ""),
        problem_statement: normalizeText(form.querySelector('[name="problem_statement"]')?.value || ""),
        deployment_locations: [...state.needDeploymentLocations],
        keywords: [...state.needTags],
        source_language: isHindiSharedFormMode() ? "hi" : "en",
      };
    if (!draft.problem_statement) {
      throw new Error("Please add the problem statement before generating tags.");
    }
    const result = await store.suggestNeedTags(draft);
    const generatedTags = parseArray(result.tags);
    generatedTags.forEach((tag) => addNeedTag(tag));
    if (status) status.textContent = result.message || "AI keywords added. You can remove any keyword before submitting.";
  }));
  byId("needTagEntry")?.addEventListener("keydown", (event) => {
    if (event.key !== "Enter" && event.key !== ",") return;
    event.preventDefault();
    const input = event.target;
    if (addNeedTag(input.value)) input.value = "";
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
    const removeNeedTagButton = event.target.closest("[data-remove-need-tag]");
    if (removeNeedTagButton) {
      removeNeedTag(removeNeedTagButton.dataset.removeNeedTag);
      return;
    }
    const removeNeedDeploymentButton = event.target.closest("[data-remove-need-deployment]");
    if (removeNeedDeploymentButton) {
      removeNeedDeploymentLocation(removeNeedDeploymentButton.dataset.removeNeedDeployment);
      return;
    }
    const addSolutionGeoSuggestionButton = event.target.closest("[data-add-solution-geo-suggestion]");
    if (addSolutionGeoSuggestionButton) {
      const input = byId("solutionGeographyEntry");
      const value = addSolutionGeoSuggestionButton.dataset.addSolutionGeoSuggestion || "";
      if (addSolutionGeography(value) && input) {
        input.value = "";
      }
      renderInlineSuggestions("solutionGeographySuggestions", getSharedFormSeedGeographies(), "data-add-solution-geo-suggestion");
      return;
    }
    const addNeedDeploymentSuggestionButton = event.target.closest("[data-add-need-deployment-suggestion]");
    if (addNeedDeploymentSuggestionButton) {
      const input = byId("needDeploymentEntry");
      const value = addNeedDeploymentSuggestionButton.dataset.addNeedDeploymentSuggestion || "";
      if (addNeedDeploymentLocation(value) && input) {
        input.value = "";
      }
      renderInlineSuggestions("needDeploymentSuggestions", getSharedFormSeedGeographies(), "data-add-need-deployment-suggestion");
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

  byId("localNeedSearch")?.addEventListener("input", (event) => {
    state.localNeedFilters.search = event.target.value || "";
    renderLocalNeedManagement();
  });

  byId("localNeedCategoryFilter")?.addEventListener("change", (event) => {
    state.localNeedFilters.category = event.target.value || "all";
    renderLocalNeedManagement();
  });

  byId("impactAuditSearch")?.addEventListener("input", (event) => {
    state.adminAuditSearch = event.target.value || "";
    renderImpactAuditPanels();
  });

  byId("generateLocalSolutionTagsPuterBtn")?.addEventListener("click", safeAsync(async () => {
    const status = byId("localSolutionReviewStatus");
    const form = byId("localSolutionReviewForm");
    const draft = getLocalSolutionTagDraft();
    if (!form || !draft) return;
    if (!draft.offering_name && !draft.about_offering_text && !draft.about_solution_text) {
      throw new Error("Add solution/offering details before generating tags.");
    }
    if (status) status.textContent = "Generating Puter-assisted tags from this solution...";
    const tags = await runPuterSolutionTagAssist(draft);
    const field = form.querySelector('[name="tags"]');
    if (field) field.value = mergeCommaListValues(field.value, tags);
    if (status) status.textContent = "Puter-assisted tags added. Review and save the solution.";
  }));

  byId("generateLocalSolutionTagsAiBtn")?.addEventListener("click", safeAsync(async () => {
    const status = byId("localSolutionReviewStatus");
    const form = byId("localSolutionReviewForm");
    const draft = getLocalSolutionTagDraft();
    if (!form || !draft) return;
    if (!draft.offering_name && !draft.about_offering_text && !draft.about_solution_text) {
      throw new Error("Add solution/offering details before generating tags.");
    }
    if (status) status.textContent = "Generating backend AI tags from this solution...";
    const result = await store.suggestSolutionTags(draft);
    const field = form.querySelector('[name="tags"]');
    if (field) field.value = mergeCommaListValues(field.value, result.tags || []);
    if (status) status.textContent = result.message || "AI-assisted tags added. Review and save the solution.";
  }));

  byId("generateLocalNeedTagsPuterBtn")?.addEventListener("click", safeAsync(async () => {
    const status = byId("localNeedReviewStatus");
    const form = byId("localNeedReviewForm");
    const draft = getLocalNeedKeywordDraft();
    if (!form || !draft) return;
    if (!draft.problem_statement) {
      throw new Error("Add the need statement before generating keywords.");
    }
    if (status) status.textContent = "Generating Puter-assisted keywords from this need...";
    const tags = await runPuterNeedKeywordAssist(draft);
    const field = form.querySelector('[name="submitted_keywords"]');
    if (field) field.value = mergeCommaListValues(field.value, tags);
    if (status) status.textContent = "Puter-assisted keywords added. Review and save the need.";
  }));

  byId("generateLocalNeedTagsAiBtn")?.addEventListener("click", safeAsync(async () => {
    const status = byId("localNeedReviewStatus");
    const form = byId("localNeedReviewForm");
    const draft = getLocalNeedKeywordDraft();
    if (!form || !draft) return;
    if (!draft.problem_statement) {
      throw new Error("Add the need statement before generating keywords.");
    }
    if (status) status.textContent = "Generating backend AI keywords from this need...";
    const result = await store.suggestNeedTags(draft);
    const field = form.querySelector('[name="submitted_keywords"]');
    if (field) field.value = mergeCommaListValues(field.value, result.tags || []);
    if (status) status.textContent = result.message || "AI-assisted keywords added. Review and save the need.";
  }));

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
      languages: canonicalizeLanguageArray(parseCommaList(form.get("languages"))),
      geographies: parseDelimitedList(form.get("geographies"), ";"),
      duration: form.get("duration"),
      duration_unit: form.get("duration_unit"),
      prerequisites: form.get("prerequisites"),
      location_availability: parseCommaList(form.get("location_availability")),
      service_cost: form.get("service_cost"),
      service_cost_unit: form.get("service_cost_unit"),
      cost_remarks: form.get("cost_remarks"),
      support_post_service: form.get("support_post_service"),
      support_post_service_cost: form.get("support_post_service_cost"),
      delivery_mode: form.get("delivery_mode"),
      certification_offered: form.get("certification_offered"),
      product_cost: form.get("product_cost"),
      grade_capacity: form.get("grade_capacity"),
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

  byId("localNeedReviewForm")?.addEventListener("submit", safeAsync(async (event) => {
    event.preventDefault();
    const form = new FormData(event.target);
    const status = byId("localNeedReviewStatus");
    if (status) status.textContent = "Saving local need changes...";
    const needId = form.get("need_id");
    const payload = {
      organization_name: form.get("organization_name"),
      contact_person: form.get("contact_person"),
      seeker_email: form.get("seeker_email"),
      seeker_phone: form.get("seeker_phone"),
      submitted_offering_category: form.get("submitted_offering_category"),
      submitted_offering_type: form.get("submitted_offering_type"),
      submitted_thematic_area: form.get("submitted_thematic_area"),
      deployment_locations: parseDelimitedList(form.get("deployment_locations"), ";"),
      submitted_keywords: parseCommaList(form.get("submitted_keywords")),
      problem_statement: form.get("problem_statement"),
    };
    const result = await store.updateLocalNeed(needId, payload);
    await refreshAll();
    if (status) status.textContent = result.message || "Local need saved.";
    localNeedReviewDialog?.close();
    toast(result.message || "Local need updated.");
  }));

  byId("deleteLocalNeedBtn")?.addEventListener("click", safeAsync(async () => {
    const form = byId("localNeedReviewForm");
    const needId = normalizeText(form?.querySelector('[name="need_id"]')?.value || "");
    if (!needId) return;
    const need = getLocalNeedById(needId);
    const name = need?.organization_name || "this need";
    if (!window.confirm(`Delete ${name} from local Supabase need records? This cannot be undone.`)) return;
    const result = await store.deleteLocalNeed(needId);
    localNeedReviewDialog?.close();
    await refreshAll();
    toast(result.message || "Local need deleted.");
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
    if (!hasAdminLikeAccess()) {
      toast("Login as admin or moderator to manage taxonomy options.");
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
    if (!hasAdminLikeAccess()) {
      toast("Login as admin or moderator first.");
      return;
    }
    const fileInput = byId("inboundWorkbookFile");
    const syncStatus = byId("syncStatus");
    const file = fileInput?.files?.[0];
    if (syncStatus) syncStatus.textContent = "Reading inbound workbook...";
    const rows = await parseInboundWorkbookFile(file);
    const provider = byId("aiProviderSelect")?.value || "openai";
    if (syncStatus) syncStatus.textContent = `Syncing ${rows.length} inbound rows and refreshing AI intelligence...`;
    const result = await store.importInboundWorkbook(file.name, rows, provider);
    await refreshAll();
    if (syncStatus) {
      syncStatus.textContent = `Imported ${result.insertedCount || 0} new needs, updated ${result.updatedCount || 0} existing needs, refreshed AI for ${result.aiUpdatedCount || 0} needs.`;
    }
    toast("Inbound workbook synced.");
  }));

  byId("downloadGreInboundBtn")?.addEventListener("click", safeAsync(async () => {
    if (!hasAdminLikeAccess()) {
      toast("Login as admin or moderator first.");
      return;
    }
    const statusEl = byId("syncStatus");
    if (statusEl) statusEl.textContent = "Downloading live GRE inbound workbook...";
    const result = await store.downloadGreInboundReport();
    downloadBase64Workbook(result.download);
    if (statusEl) statusEl.textContent = `Downloaded ${result.download?.fileName || "inbound workbook"}.`;
  }));

  byId("syncGreInboundsBtn")?.addEventListener("click", safeAsync(async () => {
    if (!hasAdminLikeAccess()) {
      toast("Login as admin or moderator first.");
      return;
    }
    const syncStatus = byId("syncStatus");
    const provider = byId("aiProviderSelect")?.value || "openai";
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
    if (!hasAdminLikeAccess()) {
      toast("Login as admin or moderator first.");
      return;
    }
    const syncStatus = byId("syncStatus");
    const provider = byId("aiProviderSelect")?.value || "openai";
    if (syncStatus) syncStatus.textContent = "Refreshing AI need intelligence...";
    const result = await store.refreshNeedIntelligence(provider);
    await refreshAll();
    if (syncStatus) syncStatus.textContent = result.message || "AI need intelligence refreshed.";
    toast("AI signals refreshed.");
  }));

  byId("saveMailTemplatesBtn")?.addEventListener("click", safeAsync(async () => {
    if (!isAdminUser()) {
      toast("Only admins can update GRE mail templates.");
      return;
    }
    const providerTemplate = byId("providerIntroTemplate")?.value || "";
    const curatorTemplate = byId("curatorForwardTemplate")?.value || "";
    const solutionSeekerTemplate = byId("solutionSeekerTemplate")?.value || "";
    const seekerTemplate = byId("needSeekerTemplate")?.value || "";
    const autoSyncEnabled = Boolean(byId("inboundAutoSyncEnabled")?.checked);
    const lshContactEmails = String(byId("lshContactEmails")?.value || "")
      .split(",")
      .map((value) => normalizeText(value))
      .filter(Boolean);
    const lshHelpCcEmails = String(byId("lshHelpCcEmails")?.value || "")
      .split(",")
      .map((value) => normalizeText(value))
      .filter(Boolean);
    const lshRequestSupportTemplate = byId("lshRequestSupportTemplate")?.value || "";
    const lshEmailProviderTemplate = byId("lshEmailProviderTemplate")?.value || "";
    setText("mailTemplateStatus", "Saving GRE mail templates...");
    const result = await store.saveMailTemplates(
      providerTemplate,
      curatorTemplate,
      solutionSeekerTemplate,
      seekerTemplate,
      autoSyncEnabled,
      lshContactEmails,
      lshHelpCcEmails,
      lshRequestSupportTemplate,
      lshEmailProviderTemplate,
    );
    state.mailTemplates = {
      providerIntroTemplate: normalizeText(result.providerIntroTemplate),
      curatorForwardTemplate: normalizeText(result.curatorForwardTemplate),
      solutionSeekerTemplate: normalizeText(result.solutionSeekerTemplate) || DEFAULT_SOLUTION_SEEKER_TEMPLATE,
      needSeekerTemplate: normalizeText(result.needSeekerTemplate) || DEFAULT_NEED_SEEKER_TEMPLATE,
      inboundAutoSyncEnabled: Boolean(result.inboundAutoSyncEnabled),
      lshContactEmails: ensureList(result.lshContactEmails).map((value) => normalizeText(value)).filter(Boolean),
      lshHelpCcEmails: ensureList(result.lshHelpCcEmails).map((value) => normalizeText(value)).filter(Boolean),
      lshRequestSupportTemplate: normalizeText(result.lshRequestSupportTemplate) || DEFAULT_LSH_REQUEST_SUPPORT_TEMPLATE,
      lshEmailProviderTemplate: normalizeText(result.lshEmailProviderTemplate) || DEFAULT_LSH_EMAIL_PROVIDER_TEMPLATE,
    };
    renderMailTemplatePanel();
    setText("mailTemplateStatus", "GRE mail templates saved.");
    toast("GRE outreach templates updated.");
  }));

  byId("confirmMailReviewBtn")?.addEventListener("click", safeAsync(async () => {
    const pending = state.pendingMailReview;
    if (!pending) return;
    setText("mailReviewStatus", "Sending GRE outreach email...");
    let result = null;
    if (pending.type === "solution_seeker_intro") {
      result = await store.sendSolutionSeekerIntro(
        state.selectedNeedId,
        pending.providerEmail,
        pending.providerName,
        pending.offeringId,
        {
          to: pending.to,
          cc: pending.cc,
          subject: pending.subject,
          body: pending.body,
          solutionName: pending.solutionName,
        },
      );
    } else {
      result = await store.sendProviderIntro(
        state.selectedNeedId,
        pending.providerEmail,
        pending.providerName,
        pending.offeringId,
        {
          to: pending.to,
          cc: pending.cc,
          subject: pending.subject,
          body: pending.body,
        },
      );
    }
    await trackImpactCounter("connections_made");
    closeMailReviewDialog();
    toast(result.message || "GRE outreach email sent.");
    const supabaseClient = store?.getClient();
    if (supabaseClient) {
      try {
        const { data: freshNeed } = await supabaseClient
          .from("gre_mis_needs")
          .select("*")
          .eq("id", state.selectedNeedId)
          .single();
        if (freshNeed) {
          const idx = state.data.needs.findIndex((n) => n.id === state.selectedNeedId);
          if (idx >= 0) {
            state.data.needs[idx] = { ...freshNeed, curated_need: parseArray(freshNeed.curated_need) };
          }
        }
      } catch {}
    }
    rerender(false);
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

  byId("uploadChatbotBtn")?.addEventListener("click", safeAsync(async () => {
    if (!hasAdminLikeAccess()) {
      toast("Login as admin or moderator first.");
      return;
    }
    if (!window.XLSX) {
      toast("Excel parser (SheetJS) is not loaded. Refresh the page.");
      return;
    }
    const solutionFile = byId("chatbotSolutionFile")?.files?.[0];
    const traderFile = byId("chatbotTraderFile")?.files?.[0];
    if (!solutionFile && !traderFile) {
      toast("Select at least one workbook (solution or trader).");
      return;
    }
    const statusEl = byId("chatbotSyncStatus");
    const provider = byId("aiProviderSelect")?.value || "openai";
    const fileDesc = [solutionFile?.name, traderFile?.name].filter(Boolean).join(", ");

    let traders = [];
    let solutions = [];
    let offerings = [];

    if (traderFile) {
      if (statusEl) statusEl.textContent = `Parsing trader workbook...`;
      const traderBuf = await traderFile.arrayBuffer();
      const traderBook = window.XLSX.read(traderBuf, { type: "array" });
      const traderSheet = traderBook.Sheets[traderBook.SheetNames[0]];
      const traderRows = window.XLSX.utils.sheet_to_json(traderSheet, { defval: "", raw: false });
      traders = dedupeRowsById(traderRows.map((row) => ({
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
      })), "trader_id");
    }

    if (solutionFile) {
      if (statusEl) statusEl.textContent = `Parsing solution workbook...`;
      const solBuf = await solutionFile.arrayBuffer();
      const solBook = window.XLSX.read(solBuf, { type: "array" });
      const solSheet = solBook.Sheets[solBook.SheetNames[0]];
      const solRows = window.XLSX.utils.sheet_to_json(solSheet, { defval: "", raw: false });

      solutions = dedupeRowsById(solRows.map((row) => ({
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
      })), "solution_id");

      offerings = dedupeRowsById(solRows.map((row) => {
        const aboutOfferingHtml = normalizeCell(row.AboutOffering);
        const trainerDetailsHtml = normalizeCell(row["Trainer Details"]);
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
          valuechains: splitLooseList(normalizeCell(row.Valuechains)),
          applications: splitLooseList(normalizeCell(row.Applications)),
          tags: splitLooseList(normalizeCell(row.Tags)),
          languages: normalizeLanguageArray(normalizeCell(row.Languages)),
          geographies: splitGeographies(normalizeCell(row.Geographies)),
          geographies_raw: normalizeCell(row.Geographies),
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
          raw_payload: row,
        };
      }), "offering_id");
    }

    if (statusEl) statusEl.textContent = `Uploading ${fileDesc} (${offerings.length} offerings, ${traders.length} traders)...`;
    const result = await store.uploadChatbotNormalized(
      traders, solutions, offerings,
      solutionFile?.name || "normalized_solutions.json",
      traderFile?.name || "normalized_traders.json",
      provider,
    );
    if (statusEl) {
      statusEl.textContent =
        `Upload complete. Traders: ${result.summary?.traders || 0}, solutions: ${result.summary?.solutions || 0}, offerings: ${result.summary?.offerings || 0}, offering AI refreshed: ${result.summary?.offeringAiUpdated || 0}.`;
    }
    toast("Chatbot workbooks uploaded and synced.");
  }));

  byId("syncGreChatbotBtn")?.addEventListener("click", safeAsync(async () => {
    if (!hasAdminLikeAccess()) {
      toast("Login as admin or moderator first.");
      return;
    }
    const statusEl = byId("chatbotSyncStatus");
    const provider = byId("aiProviderSelect")?.value || "openai";
    if (statusEl) statusEl.textContent = "Fetching live trader and solution exports from GRE and updating the chatbot dataset...";
    const result = await store.syncGreChatbotData(provider);
    if (statusEl) {
      statusEl.textContent =
        `Chatbot dataset refreshed. Traders: ${result.summary?.traders || 0}, solutions: ${result.summary?.solutions || 0}, offerings: ${result.summary?.offerings || 0}, offering AI refreshed: ${result.summary?.offeringAiUpdated || 0}.`;
    }
    toast("GRE Chatbot dataset refreshed.");
  }));

  byId("downloadGreTraderBtn")?.addEventListener("click", safeAsync(async () => {
    if (!hasAdminLikeAccess()) {
      toast("Login as admin or moderator first.");
      return;
    }
    const statusEl = byId("chatbotSyncStatus");
    if (statusEl) statusEl.textContent = "Downloading live GRE trader workbook...";
    const result = await store.downloadGreChatbotReport("trader");
    downloadBase64Workbook(result.download);
    if (statusEl) statusEl.textContent = `Downloaded ${result.download?.fileName || "trader workbook"}.`;
  }));

  byId("downloadGreSolutionBtn")?.addEventListener("click", safeAsync(async () => {
    if (!hasAdminLikeAccess()) {
      toast("Login as admin or moderator first.");
      return;
    }
    const statusEl = byId("chatbotSyncStatus");
    if (statusEl) statusEl.textContent = "Downloading live GRE solution workbook...";
    const result = await store.downloadGreChatbotReport("solution");
    downloadBase64Workbook(result.download);
    if (statusEl) statusEl.textContent = `Downloaded ${result.download?.fileName || "solution workbook"}.`;
  }));

  byId("downloadRequestTrackerBtn")?.addEventListener("click", safeAsync(async () => {
    if (!hasAdminLikeAccess() && !isCuratorUser()) {
      toast("Login as curator, moderator, or admin first.");
      return;
    }
    const statusEl = byId("requestTrackerStatus");
    const seekerKey = byId("requestTrackerSeekerSelect")?.value || "";
    if (!seekerKey) {
      toast("Select a solution seeker first.");
      return;
    }
    if (statusEl) statusEl.textContent = "Preparing seeker request tracker...";
    const includeClosed = Boolean(byId("requestTrackerIncludeClosed")?.checked);
    const result = await store.downloadSeekerRequestTracker(seekerKey, includeClosed);
    downloadBase64Workbook(result.download);
    if (statusEl) statusEl.textContent = `Downloaded ${result.download?.fileName || "request tracker"}.`;
  }));

  byId("actionWorkbench")?.addEventListener("click", safeAsync(async (event) => {
    const button = event.target.closest("button");
    if (!button) return;

    if (button.id === "assignCuratorBtn") {
      if (!hasAdminLikeAccess()) {
        toast("Login as admin or moderator to assign curators.");
        return;
      }
      await store.assignCurator(state.selectedNeedId, byId("assignCuratorSelect").value || null);
      await refreshAll();
      toast("Curator assignment updated.");
    }
  }));

  byId("needDetail")?.addEventListener("click", (event) => {
    const questionButton = event.target.closest('[data-action="generate-suggested-questions"]');
    if (questionButton) {
      safeAsync(async () => {
        const result = await store.generateNeedSuggestedQuestions(questionButton.dataset.needId);
        await refreshAll();
        toast(result.message || "Suggested questions added to curation notes.");
      })();
      return;
    }
    const rejectQuestionButton = event.target.closest('[data-action="reject-suggested-questions"]');
    if (rejectQuestionButton) {
      safeAsync(async () => {
        const result = await store.rejectNeedSuggestedQuestions(rejectQuestionButton.dataset.needId);
        await refreshAll();
        toast(result.message || "Suggested questions removed from curation notes.");
      })();
      return;
    }
    const puterQuestionButton = event.target.closest('[data-action="generate-suggested-questions-puter"]');
    if (puterQuestionButton) {
      safeAsync(async () => {
        const result = await runPuterSuggestedQuestions(puterQuestionButton.dataset.needId);
        toast(result.message || "Suggested questions from Puter were added to curation notes.");
      })();
      return;
    }
    const button = event.target.closest("#openWorkbenchBtn");
    if (!button) return;
    workbenchDialog?.showModal();
  });

  byId("actionWorkbench")?.addEventListener("click", (event) => {
    const addBtn = event.target.closest("#addCuratedNeedBtn");
    if (addBtn) {
      const input = byId("curatedNeedEntry");
      if (!input) return;
      const value = normalizeText(input.value);
      if (!value) return;
      const lowered = value.toLowerCase();
      const matchedOption = CURATED_NEED_OPTIONS.find((opt) => opt.toLowerCase() === lowered);
      if (!matchedOption) {
        toast(`"${value}" is not a valid curated need category.`);
        return;
      }
      const chipContainer = byId("curatedNeedChips");
      const hiddenContainer = byId("curatedNeedHiddenInputs");
      if (!chipContainer || !hiddenContainer) return;
      if ([...chipContainer.querySelectorAll("[data-remove-curated-need]")].some((btn) => btn.getAttribute("data-remove-curated-need").toLowerCase() === lowered)) {
        toast("This category is already selected.");
        return;
      }
      const chip = document.createElement("span");
      chip.className = "tag-chip";
      chip.style.cssText = "display:inline-flex;align-items:center;gap:6px;padding:7px 11px;border-radius:999px;background:var(--green-soft);color:var(--green);font-size:0.8rem;font-weight:700;";
      chip.innerHTML = `${esc(matchedOption)} <button type="button" class="tag-chip-remove" data-remove-curated-need="${escAttr(matchedOption)}" style="border:none;background:transparent;color:inherit;cursor:pointer;padding:0;font:inherit;line-height:1;" aria-label="Remove ${escAttr(matchedOption)}">&times;</button>`;
      chipContainer.append(chip);
      chipContainer.querySelector(".helper-text")?.remove();
      const hidden = document.createElement("input");
      hidden.type = "hidden";
      hidden.name = "proposed_curated_need";
      hidden.value = matchedOption;
      hiddenContainer.appendChild(hidden);
      input.value = "";
      return;
    }
    const removeBtn = event.target.closest("[data-remove-curated-need]");
    if (removeBtn) {
      const value = removeBtn.getAttribute("data-remove-curated-need");
      removeBtn.closest(".tag-chip")?.remove();
      const hiddenContainer = byId("curatedNeedHiddenInputs");
      if (hiddenContainer) {
        const inputs = hiddenContainer.querySelectorAll(`input[value="${escAttr(value)}"]`);
        inputs.forEach((el) => el.remove());
      }
      const chipContainer = byId("curatedNeedChips");
      if (chipContainer && !chipContainer.querySelector(".tag-chip")) {
        chipContainer.innerHTML = '<span class="helper-text">No curated need categories selected.</span>';
      }
      return;
    }
  });

  byId("actionWorkbench")?.addEventListener("keydown", (event) => {
    if (event.target.id === "curatedNeedEntry" && event.key === "Enter") {
      event.preventDefault();
      const addBtn = byId("addCuratedNeedBtn");
      if (addBtn) addBtn.click();
    }
  });

  byId("actionWorkbench")?.addEventListener("submit", safeAsync(async (event) => {
    if (event.target.id === "directUpdateForm") {
      event.preventDefault();
      const need = getNeedById(state.selectedNeedId);
      const form = new FormData(event.target);
      const result = await store.directCuratorUpdate({
        needId: need.id,
        proposedStatus: form.get("proposed_status"),
        proposedInternalStatus: form.get("proposed_internal_status"),
        proposedNextAction: form.get("proposed_next_action"),
        proposedCurationNotes: form.get("proposed_curation_notes"),
        proposedCurationCallDate: form.get("proposed_curation_call_date"),
        proposedCuratedNeed: form.getAll("proposed_curated_need"),
        proposedFundingMechanism: form.get("proposed_funding_mechanism"),
        proposedSeekerProviderAgreement: form.get("proposed_seeker_provider_agreement"),
        proposedSolutionDeploymentStatus: form.get("proposed_solution_deployment_status"),
        proposedClosureDate: form.get("proposed_closure_date"),
        proposedFeedbackAboutSeeker: form.get("proposed_feedback_about_seeker"),
        proposedFeedbackAboutProvider: form.get("proposed_feedback_about_provider"),
        proposedDemandBroadcastNeeded:
          form.get("proposed_demand_broadcast_needed") === ""
            ? null
            : form.get("proposed_demand_broadcast_needed") === "true",
      });
      event.target.reset();
      await refreshAll();
      workbenchDialog?.close();
      toast(result.message || "Curation update saved.");
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
      trackImpactCounter("connections_made");
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

  byId("generateSubmissionNeedKeywordsBtn")?.addEventListener("click", safeAsync(async () => {
    const status = byId("submissionNeedToolsStatus");
    const draft = getSubmissionReviewNeedDraft();
    if (!draft) return;
    if (!draft.problem_statement) {
      throw new Error("Please add the need statement before generating keywords.");
    }
    if (status) status.textContent = "Generating AI-assisted keywords from the current need draft...";
    const result = await store.suggestNeedTags(draft);
    const keywordField = byId("submissionReviewForm")?.querySelector('[data-review-field="keywords"]');
    if (keywordField) {
      keywordField.value = parseArray(result.tags).join(", ");
    }
    if (status) status.textContent = result.message || "AI-assisted keywords added to the review draft.";
  }));

  byId("generateSubmissionNeedKeywordsPuterBtn")?.addEventListener("click", safeAsync(async () => {
    const status = byId("submissionNeedToolsStatus");
    const draft = getSubmissionReviewNeedDraft();
    if (!draft) return;
    if (!draft.problem_statement) {
      throw new Error("Please add the need statement before generating keywords.");
    }
    if (status) status.textContent = "Generating Puter-assisted keywords from the current need draft...";
    const tags = await runPuterNeedKeywordAssist(draft);
    const keywordField = byId("submissionReviewForm")?.querySelector('[data-review-field="keywords"]');
    if (keywordField) {
      keywordField.value = tags.join(", ");
    }
    if (status) status.textContent = "Puter-assisted keywords added to the review draft.";
  }));

  byId("approveSubmissionReviewBtn")?.addEventListener("click", safeAsync(async () => {
    const status = byId("submissionReviewStatus");
    if (status) status.textContent = "Saving changes and approving...";
    const patch = await buildSubmissionReviewPatch();
    await store.updateFormSubmission(patch.submissionId, patch.update);
    const result = await store.reviewFormSubmission(patch.submissionId, "approve", patch.update.adminReviewNotes || "");
    submissionReviewDialog?.close();
    await refreshActiveAdminDeskTab();
    toast(result.message || (result.targetNeedId ? `Submission approved as Need ${result.targetNeedId}.` : "Form submission approved."));
  }));

  byId("rejectSubmissionReviewBtn")?.addEventListener("click", safeAsync(async () => {
    if (!canRejectRecords()) {
      toast("Only admins can reject form submissions.");
      return;
    }
    const status = byId("submissionReviewStatus");
    if (status) status.textContent = "Saving changes and rejecting...";
    const patch = await buildSubmissionReviewPatch();
    await store.updateFormSubmission(patch.submissionId, patch.update);
    await store.reviewFormSubmission(patch.submissionId, "reject", patch.update.adminReviewNotes || "");
    submissionReviewDialog?.close();
    await refreshActiveAdminDeskTab();
    toast("Form submission rejected.");
  }));

  byId("adminView")?.addEventListener("click", safeAsync(async (event) => {
    const button = event.target.closest("[data-action]");
    if (!button) return;
    if (!hasAdminLikeAccess()) {
      toast("Login as admin or moderator first.");
      return;
    }
    if (button.dataset.action === "approve-need") {
      await store.approveNeed(button.dataset.needId, "approve");
      await refreshActiveAdminDeskTab();
      toast("Need approved.");
    }
    if (button.dataset.action === "reject-need") {
      if (!canRejectRecords()) {
        toast("Only admins can reject needs.");
        return;
      }
      await store.approveNeed(button.dataset.needId, "reject");
      await refreshActiveAdminDeskTab();
      toast("Need rejected.");
    }
    if (button.dataset.action === "approve-update") {
      await store.reviewUpdateRequest(button.dataset.requestId, "approve");
      await refreshActiveAdminDeskTab();
      toast("Curator update approved.");
    }
    if (button.dataset.action === "reject-update") {
      if (!canRejectRecords()) {
        toast("Only admins can reject curator updates.");
        return;
      }
      await store.reviewUpdateRequest(button.dataset.requestId, "reject");
      await refreshActiveAdminDeskTab();
      toast("Curator update rejected.");
    }
    if (button.dataset.action === "edit-form-submission") {
      openSubmissionReviewDialog(button.dataset.submissionId);
      return;
    }
    if (button.dataset.action === "approve-form-submission") {
      const result = await store.reviewFormSubmission(button.dataset.submissionId, "approve");
      await refreshActiveAdminDeskTab();
      toast(result.message || (result.targetNeedId ? `Submission approved as Need ${result.targetNeedId}.` : "Form submission approved."));
    }
    if (button.dataset.action === "reject-form-submission") {
      if (!canRejectRecords()) {
        toast("Only admins can reject form submissions.");
        return;
      }
      await store.reviewFormSubmission(button.dataset.submissionId, "reject");
      await refreshActiveAdminDeskTab();
      toast("Form submission rejected.");
    }
      if (button.dataset.action === "run-puter-need-review") {
        const status = byId("puterStatus");
        if (status) status.textContent = `Running Puter review for Need ${button.dataset.needId}...`;
        await runPuterNeedReview(button.dataset.needId);
        if (status) status.textContent = `Puter recommendation ready for Need ${button.dataset.needId}.`;
        toast("Puter recommendation prepared.");
      }
      if (button.dataset.action === "generate-suggested-questions") {
        const result = await store.generateNeedSuggestedQuestions(button.dataset.needId);
        await refreshAll();
        toast(result.message || "Suggested questions added to curation notes.");
      }
      if (button.dataset.action === "generate-suggested-questions-puter") {
        const result = await runPuterSuggestedQuestions(button.dataset.needId);
        toast(result.message || "Suggested questions from Puter were added to curation notes.");
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
      if (!canManageUsers()) {
        toast("Only admins can remove users.");
        return;
      }
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
      await openLocalSolutionReviewDialog(button.dataset.offeringId);
      return;
    }
    if (button.dataset.action === "delete-local-solution") {
      if (!canDeleteRecords()) {
        toast("Only admins can delete local solutions.");
        return;
      }
      const offeringId = button.dataset.offeringId;
      const card = button.closest(".approval-card");
      const name = card?.querySelector("h4")?.textContent?.trim() || "this solution";
      if (!window.confirm(`Delete ${name} from local Supabase solution records? This cannot be undone.`)) return;
      const result = await store.deleteLocalSolution(offeringId);
      await refreshAll();
      toast(result.message || "Local solution deleted.");
    }
    if (button.dataset.action === "edit-local-need") {
      await openLocalNeedReviewDialog(button.dataset.needId);
      return;
    }
    if (button.dataset.action === "delete-local-need") {
      if (!canDeleteRecords()) {
        toast("Only admins can delete local needs.");
        return;
      }
      const needId = button.dataset.needId;
      const card = button.closest(".approval-card");
      const name = card?.querySelector("h4")?.textContent?.trim() || "this need";
      if (!window.confirm(`Delete ${name} from local Supabase need records? This cannot be undone.`)) return;
      const result = await store.deleteLocalNeed(needId);
      await refreshAll();
      toast(result.message || "Local need deleted.");
    }
  }));

  [byId("matchResults"), byId("manualMatchResults")].forEach((container) => {
    container?.addEventListener("click", safeAsync(async (event) => {
      const button = event.target.closest("[data-action]");
      if (!button) return;
      const match = getMatchForActionButton(button);

      if (button.dataset.action === "read-more-match") {
        openMatchDetailDialog(match);
        return;
      }
      if (button.dataset.action === "view-match-details") {
        await trackImpactCounter("solutions_discovered", 1, {
          kind: "view",
          action: "view_details",
          itemId: button.dataset.offeringId || match?.offering_id || "",
          itemLabel: match?.offering_name || match?.solution_name || "Offering",
          itemSource: normalizeText(match?.source_slug || match?.source_label || "gre"),
        });
        openMatchDetailDialog(match || { offering_id: button.dataset.offeringId });
        return;
      }
      if (button.dataset.action === "open-gre-link") {
        event.preventDefault();
        await trackImpactCounter("solutions_discovered", 1, {
          kind: "view",
          action: "view_portal",
          itemId: button.dataset.offeringId || match?.offering_id || "",
          itemLabel: match?.offering_name || match?.solution_name || "Offering",
          itemSource: normalizeText(match?.source_slug || match?.source_label || "gre"),
          portalUrl: match?.gre_link || button.getAttribute("href") || "",
        });
        const href = button.getAttribute("href");
        if (href) {
          window.open(href, "_blank", "noopener,noreferrer");
        }
        return;
      }
      if (button.dataset.action === "email-provider") {
        if (!(hasAdminLikeAccess() || isCuratorUser())) {
          toast("Login as curator, moderator, or admin to send provider outreach from the GRE mailbox.");
          return;
        }
        const need = getNeedById(state.selectedNeedId);
        if (!need || !match) {
          toast("Need or provider context is unavailable right now.");
          return;
        }
        openMailReviewDialog(getMatchEmailActionConfig(need, match));
        return;
      }
      if (button.dataset.action === "email-seeker") {
        if (!(hasAdminLikeAccess() || isCuratorUser())) {
          toast("Login as curator, moderator, or admin to send seeker outreach from the GRE mailbox.");
          return;
        }
        const need = getNeedById(state.selectedNeedId);
        if (!need || !match) {
          toast("Need or seeker context is unavailable right now.");
          return;
        }
        openMailReviewDialog(getMatchSeekerEmailActionConfig(need, match));
      }
    }));
  });
}

async function init() {
  document.body.classList.add("gre-boot-pending");
  setText("userLoginStatus", "");
  setText("registerStatus", "");
  setText("resetStatus", "");
  setText("changePasswordStatus", "");
  bindStaticEvents();
  bindEmbeddedBridge();
  attachGrameeeAuthBridgeListener();
  captureGrameeeSessionTransfer();
  const needsGreMisSession = !isSharedFormMode() && !isStandalonePublicFormMode();

  let sessionReady = await store.validateUserSession();
  if (!sessionReady && needsGreMisSession && hasImmediateGreLoginContext()) {
    sessionReady = await store.bridgeGrameeeSession();
  }
  if (!sessionReady && needsGreMisSession && !hasImmediateGreLoginContext()) {
    await waitForGrameeeAuthBootstrap(600);
    sessionReady = await store.validateUserSession();
    if (!sessionReady) {
      sessionReady = await store.bridgeGrameeeSession();
    }
  }

  if (!sessionReady && needsGreMisSession) {
    renderAuthState();
    renderSubmissionViews();
    renderAdminState();
    document.body.classList.remove("gre-boot-pending");
    return;
  }

  renderAuthState();
  renderSubmissionViews();
  renderAdminState();
  document.body.classList.remove("gre-boot-pending");
  syncSharedFormLanguageFromBootstrap({ persist: false });
  await applySharedFormLanguage().catch(() => null);
  await refreshAll();
  syncSharedFormLanguageFromBootstrap({ persist: false });

  window.setTimeout(() => {
    tryBridgeSharedSessionAndRefresh().catch(() => null);
  }, 900);

  window.addEventListener("focus", () => {
    tryBridgeSharedSessionAndRefresh().catch(() => null);
  });

  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") {
      tryBridgeSharedSessionAndRefresh().catch(() => null);
    }
  });
}

init().catch((error) => {
  console.error(error);
  toast(error.message || "Dashboard could not load.");
});
