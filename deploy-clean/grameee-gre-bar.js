(function () {
  const SUMMARY_COOKIE = "grameee_user_summary";
  const ACCESS_TOKEN_COOKIE = "grameee_access_token";
  const REFRESH_TOKEN_COOKIE = "grameee_refresh_token";
  const STORAGE_KEY = "grameee-user-session";
  const APP_BASE = "https://grameee.org";
  const COOKIE_DOMAIN = window.location.hostname.endsWith(".grameee.org") || window.location.hostname === "grameee.org"
    ? ".grameee.org"
    : "";

  function getCookie(name) {
    const parts = document.cookie ? document.cookie.split("; ") : [];
    const prefix = `${name}=`;
    for (const part of parts) {
      if (part.indexOf(prefix) === 0) {
        return decodeURIComponent(part.slice(prefix.length));
      }
    }
    return "";
  }

  function deleteCookie(name) {
    const options = ["path=/", "SameSite=Lax", "max-age=0"];
    if (window.location.protocol === "https:") options.push("Secure");
    if (COOKIE_DOMAIN) options.push(`domain=${COOKIE_DOMAIN}`);
    document.cookie = `${name}=; ${options.join("; ")}`;
  }

  function trim(value) {
    return String(value || "").trim();
  }

  function decodeJwtPayload(token) {
    const value = trim(token);
    if (!value) return null;
    const parts = value.split(".");
    if (parts.length < 2) return null;
    try {
      const normalized = parts[1].replace(/-/g, "+").replace(/_/g, "/");
      const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=");
      return JSON.parse(atob(padded));
    } catch {
      return null;
    }
  }

  function readSummary() {
    const raw = getCookie(SUMMARY_COOKIE);
    if (!raw) {
      const tokenPayload = decodeJwtPayload(getCookie(ACCESS_TOKEN_COOKIE));
      if (!tokenPayload) return null;
      return {
        username: trim(tokenPayload.user_metadata?.username),
        fullName: trim(tokenPayload.user_metadata?.full_name),
        role: trim(tokenPayload.app_metadata?.grameee_role).toLowerCase(),
      };
    }
    try {
      const parsed = JSON.parse(raw);
      return {
        username: trim(parsed.username),
        fullName: trim(parsed.fullName),
        role: trim(parsed.role).toLowerCase(),
      };
    } catch {
      return null;
    }
  }

  function updateLoginLinks() {
    const returnTo = encodeURIComponent(window.location.href);
    document.querySelectorAll("[data-login-link]").forEach((node) => {
      node.setAttribute("href", `${APP_BASE}/login.html?returnTo=${returnTo}`);
    });
  }

  function buildUserMenu(summary) {
    const slot = document.querySelector("[data-gre-auth-slot]");
    if (!slot || !summary) return;
    const label = summary.username || summary.fullName || "User";
    slot.hidden = false;
    slot.innerHTML = `
      <details class="grameee-user-menu">
        <summary class="grameee-user-trigger" role="button" aria-label="User menu">
          <span>${label}</span>
          <span class="grameee-user-caret">&#9662;</span>
        </summary>
        <div class="grameee-user-dropdown">
          <a class="grameee-user-action" href="${APP_BASE}/account.html">Update User Details</a>
          <button class="grameee-user-action" type="button" data-gre-logout>Logout</button>
        </div>
      </details>
    `;

    const logoutButton = slot.querySelector("[data-gre-logout]");
    logoutButton?.addEventListener("click", () => {
      window.localStorage.removeItem(STORAGE_KEY);
      deleteCookie(SUMMARY_COOKIE);
      deleteCookie(ACCESS_TOKEN_COOKIE);
      deleteCookie(REFRESH_TOKEN_COOKIE);
      window.location.reload();
    });
  }

  function init() {
    updateLoginLinks();
    const summary = readSummary();
    document.querySelectorAll("[data-login-link]").forEach((node) => {
      node.hidden = Boolean(summary);
    });
    if (summary) {
      buildUserMenu(summary);
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init, { once: true });
  } else {
    init();
  }
}());
