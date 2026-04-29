import { createClient } from "npm:@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

const supabaseUrl = Deno.env.get("SUPABASE_URL") ?? "";
const serviceRoleKey = Deno.env.get("GRE_MIS_SERVICE_ROLE_KEY") ?? "";
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

async function getGmailAccessToken() {
  const response = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
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

async function sendEmail({ to, cc, subject, body }: { to: string; cc?: string; subject: string; body: string }) {
  const accessToken = await getGmailAccessToken();
  const rawMessage = [
    `From: ${gmailSenderEmail}`,
    `To: ${to}`,
    cc ? `Cc: ${cc}` : "",
    `Subject: ${subject}`,
    "Content-Type: text/plain; charset=UTF-8",
    "",
    body,
  ].filter(Boolean).join("\r\n");

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

async function getRequesterEmail(req: Request) {
  const jwt = req.headers.get("Authorization")?.replace(/^Bearer\s+/i, "").trim();
  if (!jwt) return null;
  const authClient = createClient(supabaseUrl, Deno.env.get("SUPABASE_ANON_KEY") ?? "", {
    auth: { persistSession: false },
    global: { headers: { Authorization: `Bearer ${jwt}` } },
  });
  const { data } = await authClient.auth.getUser(jwt);
  return data.user?.email?.toLowerCase() ?? null;
}

async function assertAdminOrCurator(email: string | null) {
  if (!email) throw new Error("Authentication required.");
  const [adminResult, curatorResult] = await Promise.all([
    adminClient.from("gre_mis_admins").select("email").eq("email", email).maybeSingle(),
    adminClient.from("gre_mis_curators").select("email, is_active").eq("email", email).eq("is_active", true).maybeSingle(),
  ]);
  if (!adminResult.data && !curatorResult.data) {
    throw new Error("You do not have access to this function.");
  }
  return { isAdmin: Boolean(adminResult.data), isCurator: Boolean(curatorResult.data) };
}

async function assignCurator(needId: string, curatorId: string | null, actorEmail: string) {
  const updatePayload: Record<string, unknown> = {
    curator_id: curatorId,
    status: "Accepted",
    updated_at: new Date().toISOString(),
  };
  const { error } = await adminClient.from("gre_mis_needs").update(updatePayload).eq("id", needId);
  if (error) throw new Error(error.message);
  await adminClient.from("gre_mis_need_updates").insert({
    need_id: needId,
    update_type: "curator_assignment",
    note: curatorId ? `Curator assigned by ${actorEmail}.` : `Curator unassigned by ${actorEmail}.`,
    created_by_email: actorEmail,
  });
  return { ok: true };
}

async function upsertOption(optionType: string, label: string, actorEmail: string) {
  const { data: existing } = await adminClient
    .from("gre_mis_options")
    .select("id")
    .eq("option_type", optionType)
    .ilike("label", label)
    .maybeSingle();

  if (existing?.id) {
    const { error } = await adminClient.from("gre_mis_options").update({ is_active: true }).eq("id", existing.id);
    if (error) throw new Error(error.message);
    return { ok: true, reused: true };
  }

  const { data: current } = await adminClient
    .from("gre_mis_options")
    .select("sort_order")
    .eq("option_type", optionType)
    .order("sort_order", { ascending: false })
    .limit(1);

  const nextSortOrder = Number(current?.[0]?.sort_order || 0) + 1;
  const { error } = await adminClient.from("gre_mis_options").insert({
    option_type: optionType,
    label,
    sort_order: nextSortOrder,
  });
  if (error) throw new Error(error.message);
  return { ok: true, createdBy: actorEmail };
}

async function sendProviderIntro(needId: string, providerId: string, actorEmail: string) {
  const { data: need, error: needError } = await adminClient
    .from("gre_mis_needs")
    .select("id, organization_name, seeker_email, contact_person, problem_statement, state, district, curated_need")
    .eq("id", needId)
    .single();
  if (needError || !need) throw new Error(needError?.message || "Need not found.");

  const { data: provider, error: providerError } = await adminClient
    .from("gre_mis_solution_providers")
    .select("id, name, email, solution_tags, notes")
    .eq("id", providerId)
    .single();
  if (providerError || !provider) throw new Error(providerError?.message || "Solution provider not found.");

  const subject = `GRE introduction: ${need.organization_name} challenge for your consideration`;
  const body = [
    `Hello ${provider.name},`,
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
    to: provider.email,
    cc: need.seeker_email,
    subject,
    body,
  });

  await adminClient.from("gre_mis_need_matches").upsert({
    need_id: need.id,
    provider_id: provider.id,
    outreach_status: "emailed",
    emailed_at: new Date().toISOString(),
    emailed_by_email: actorEmail,
    seeker_cc_email: need.seeker_email,
    match_reason: "Email introduced from GRE MIS dashboard.",
  }, { onConflict: "need_id,provider_id" });

  await adminClient.from("gre_mis_email_log").insert({
    need_id: need.id,
    provider_id: provider.id,
    recipient_email: provider.email,
    cc_email: need.seeker_email,
    subject,
    body_preview: body.slice(0, 1000),
    sent_by_email: actorEmail,
  });

  await adminClient.from("gre_mis_need_updates").insert({
    need_id: need.id,
    update_type: "provider_intro_sent",
    note: `Problem statement emailed to ${provider.name} with seeker copied.`,
    created_by_email: actorEmail,
  });

  return { ok: true, message: "Provider introduction email sent." };
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  try {
    const actorEmail = await getRequesterEmail(req);
    const { isAdmin } = await assertAdminOrCurator(actorEmail);
    const payload = await req.json();
    const action = String(payload.action || "");

    if (action === "assignCurator") {
      if (!isAdmin) throw new Error("Only admins can assign curators.");
      return jsonResponse(await assignCurator(String(payload.needId || ""), payload.curatorId ? String(payload.curatorId) : null, actorEmail || ""));
    }

    if (action === "upsertOption") {
      if (!isAdmin) throw new Error("Only admins can manage options.");
      return jsonResponse(await upsertOption(String(payload.optionType || ""), String(payload.label || ""), actorEmail || ""));
    }

    if (action === "sendProviderIntro") {
      return jsonResponse(await sendProviderIntro(String(payload.needId || ""), String(payload.providerId || ""), actorEmail || ""));
    }

    return jsonResponse({ error: "Unsupported action." }, 400);
  } catch (error) {
    const message = error instanceof Error ? error.message : "Function failed.";
    return jsonResponse({ error: message }, 400);
  }
});
