import { createClient } from "npm:@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type, x-gre-admin-session",
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
const gmailSenderEmail = Deno.env.get("GMAIL_SENDER_EMAIL") ?? "tanmay@greenruraleconomy.in";

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

async function createAdminSession(username: string) {
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

async function validateAdminSession(token: string) {
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

async function requireAdminSession(req: Request, bodyToken?: string) {
  const token = requireString(req.headers.get("x-gre-admin-session")) || requireString(bodyToken);
  const session = await validateAdminSession(token);
  if (!session) throw new Error("Admin login required.");

  const { data: adminRow, error } = await adminClient
    .from("gre_mis_admins")
    .select("email, display_name")
    .limit(1)
    .maybeSingle();
  if (error || !adminRow) throw new Error("Admin record not found.");
  return { session, admin: adminRow, token };
}

async function verifyAdminPassword(username: string, password: string) {
  const { data, error } = await adminClient.rpc("grameee_admin_password_matches", {
    p_username: username,
    p_password: password,
  });
  if (error) throw new Error(`Admin password verification failed: ${error.message}`);
  return Boolean(data);
}

async function adminLogin(username: string, password: string) {
  if (!password) throw new Error("Password is required.");
  const normalizedUsername = requireString(username) || "admin";
  const valid = await verifyAdminPassword(normalizedUsername, password);
  if (!valid) throw new Error("Incorrect admin password.");
  const session = await createAdminSession(normalizedUsername);
  return { ok: true, ...session };
}

async function adminLogout(token: string) {
  const normalized = requireString(token);
  if (!normalized) return { ok: true };
  const tokenHash = await hashToken(normalized);
  await adminClient.from("gre_mis_admin_web_sessions").delete().eq("token_hash", tokenHash);
  return { ok: true };
}

async function getAdminSnapshot() {
  const [pendingNeeds, pendingUpdates] = await Promise.all([
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
  ]);

  if (pendingNeeds.error) throw new Error(pendingNeeds.error.message);
  if (pendingUpdates.error) throw new Error(pendingUpdates.error.message);

  return {
    ok: true,
    pendingNeeds: pendingNeeds.data || [],
    pendingUpdates: pendingUpdates.data || [],
  };
}

async function assignCurator(needId: string, curatorId: string | null, actorEmail: string) {
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
  if (requestRow.proposed_curation_notes) {
    nextNeedPatch.curation_notes = requestRow.proposed_curation_notes;
  }
  if (requestRow.proposed_curation_call_date) {
    nextNeedPatch.curation_call_date = requestRow.proposed_curation_call_date;
  }
  if (requestRow.proposed_demand_broadcast_needed !== null) {
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

  await adminClient.from("gre_mis_need_updates").insert({
    need_id: requestRow.need_id,
    update_type: "curator_update_approved",
    note: changeParts.join(" | ") || "Curator update approved.",
    created_by_email: actorEmail,
  });

  return { ok: true };
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
    .select("id, organization_name, seeker_email, contact_person, problem_statement, state, district, curated_need")
    .eq("id", needId)
    .single();
  if (needError || !need) throw new Error(needError?.message || "Need not found.");

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
    cc: need.seeker_email,
    subject,
    body,
  });

  await adminClient.from("gre_mis_email_log").insert({
    need_id: need.id,
    recipient_email: providerEmail,
    cc_email: need.seeker_email,
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

  return { ok: true, message: "Provider introduction email sent." };
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  try {
    const payload = (await req.json().catch(() => ({}))) as Record<string, unknown>;
    const action = requireString(payload.action);

    if (action === "adminLogin") {
      return jsonResponse(await adminLogin(requireString(payload.username), requireString(payload.password)));
    }

    if (action === "validateAdminSession") {
      const session = await requireAdminSession(req, requireString(payload.adminSessionToken));
      return jsonResponse({ ok: true, username: session.session.username, email: session.admin.email });
    }

    if (action === "adminLogout") {
      return jsonResponse(await adminLogout(requireString(payload.adminSessionToken) || requireString(req.headers.get("x-gre-admin-session"))));
    }

    if (action === "submitUpdateRequest") {
      return jsonResponse(await submitUpdateRequest(payload));
    }

    const adminCtx = await requireAdminSession(req, requireString(payload.adminSessionToken));
    const actorEmail = adminCtx.admin.email;

    if (action === "adminSnapshot") {
      return jsonResponse(await getAdminSnapshot());
    }

    if (action === "assignCurator") {
      return jsonResponse(await assignCurator(requireString(payload.needId), requireString(payload.curatorId) || null, actorEmail));
    }

    if (action === "approveNeed") {
      return jsonResponse(await approveNeed(requireString(payload.needId), requireString(payload.decision), requireString(payload.reviewNotes), actorEmail));
    }

    if (action === "reviewUpdateRequest") {
      return jsonResponse(await reviewUpdateRequest(requireString(payload.requestId), requireString(payload.decision), requireString(payload.reviewNotes), actorEmail));
    }

    if (action === "upsertOption") {
      return jsonResponse(await upsertOption(requireString(payload.optionType), requireString(payload.label)));
    }

    if (action === "sendProviderIntro") {
      return jsonResponse(await sendProviderIntro(requireString(payload.needId), requireString(payload.providerEmail), actorEmail));
    }

    return jsonResponse({ error: "Unsupported action." }, 400);
  } catch (error) {
    const message = error instanceof Error ? error.message : "Function failed.";
    return jsonResponse({ error: message }, 400);
  }
});
