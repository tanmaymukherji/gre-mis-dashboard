// Fresh deploy script for solution.grameee.org
// Creates and uploads a new ZIP every time — never reuses old archives

import { createReadStream, statSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import { createRequire } from "module";

const require = createRequire(import.meta.url);
const tus = require("C:\\Users\\tmukh\\AppData\\Roaming\\npm\\node_modules\\hostinger-api-mcp\\node_modules\\tus-js-client");
const SCRIPT_DIR = dirname(fileURLToPath(import.meta.url));

const TOKEN = "tsAkruR0VW24ayQyL0TRlWTuoxXdbjUSzluPhrSS16c6ca26";
const BASE = "https://developers.hostinger.com";
const USERNAME = "u973202051";
const DOMAIN = process.env.HOSTINGER_DOMAIN || "solution.grameee.org";
const SOURCE_DIR = process.env.HOSTINGER_SOURCE_DIR || "hostinger-solution-need";
const ARCHIVE_PREFIX = process.env.HOSTINGER_ARCHIVE_PREFIX || SOURCE_DIR;

// Always create a fresh archive — never reuse old ones
const now = new Date();
const datestamp = now.getFullYear()
  + String(now.getMonth() + 1).padStart(2, "0")
  + String(now.getDate()).padStart(2, "0")
  + "_"
  + String(now.getHours()).padStart(2, "0")
  + String(now.getMinutes()).padStart(2, "0")
  + String(now.getSeconds()).padStart(2, "0");
const ARCHIVE_NAME = `${ARCHIVE_PREFIX}_${datestamp}.zip`;
const ARCHIVE_PATH = join(process.env.TEMP || "C:\\Users\\tmukh\\AppData\\Local\\Temp", ARCHIVE_NAME);

async function main() {
  // 1. Create fresh ZIP from source
  console.log(`Creating fresh archive from ${SOURCE_DIR}: ${ARCHIVE_NAME}`);
  const { execSync } = await import("child_process");
  execSync(
    `powershell -Command "Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::CreateFromDirectory('${SOURCE_DIR.replace(/'/g, "''")}', '${ARCHIVE_PATH.replace(/'/g, "''")}')"`,
    { cwd: SCRIPT_DIR, stdio: "inherit" }
  );
  console.log("Archive created.");

  // 2. Get upload credentials
  console.log("Getting upload credentials...");
  const credResp = await fetch(`${BASE}/api/hosting/v1/files/upload-urls`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${TOKEN}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ username: USERNAME, domain: DOMAIN }),
  });
  if (!credResp.ok) {
    const text = await credResp.text();
    throw new Error(`Failed to get upload credentials: ${credResp.status} ${text}`);
  }
  const creds = await credResp.json();
  const { url: uploadUrl, auth_key: authToken, rest_auth_key: authRestToken } = creds;
  console.log(`Upload URL obtained`);

  // 3. Pre-upload POST
  const cleanUploadUrl = uploadUrl.replace(/\/$/, "");
  const uploadUrlWithFile = `${cleanUploadUrl}/${ARCHIVE_NAME}?override=true`;
  const stats = statSync(ARCHIVE_PATH);

  console.log("Pre-upload POST...");
  const preResp = await fetch(uploadUrlWithFile, {
    method: "POST",
    headers: {
      "X-Auth": authToken,
      "X-Auth-Rest": authRestToken,
      "upload-length": String(stats.size),
      "upload-offset": "0",
    },
  });
  if (preResp.status !== 201) {
    const text = await preResp.text();
    throw new Error(`Pre-upload failed: ${preResp.status} ${text}`);
  }
  console.log("Pre-upload OK (201)");

  // 4. TUS upload
  console.log("Uploading via TUS...");
  await new Promise((resolve, reject) => {
    const fileStream = createReadStream(ARCHIVE_PATH);
    const upload = new tus.Upload(fileStream, {
      uploadUrl: uploadUrlWithFile,
      retryDelays: [1000, 2000, 4000, 8000],
      uploadDataDuringCreation: false,
      parallelUploads: 1,
      chunkSize: 10485760,
      headers: {
        "X-Auth": authToken,
        "X-Auth-Rest": authRestToken,
        "upload-length": String(stats.size),
        "upload-offset": "0",
      },
      removeFingerprintOnSuccess: true,
      uploadSize: stats.size,
      metadata: { filename: ARCHIVE_NAME },
      onError: (err) => reject(new Error(`TUS upload error: ${err.message}`)),
      onSuccess: () => {
        console.log("TUS upload completed");
        resolve();
      },
    });
    upload.start();
  });

  // 5. Trigger deployment
  console.log("Triggering deployment...");
  const deployResp = await fetch(
    `${BASE}/api/hosting/v1/accounts/${USERNAME}/websites/${DOMAIN}/deploy`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${TOKEN}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ archive_path: ARCHIVE_NAME }),
    }
  );
  if (!deployResp.ok) {
    const text = await deployResp.text();
    throw new Error(`Deploy failed: ${deployResp.status} ${text}`);
  }
  const deployResult = await deployResp.json();
  console.log("Deployment triggered:", JSON.stringify(deployResult));
  console.log(`Done! Files deploying to ${DOMAIN}.`);
}

main().catch((err) => {
  console.error("Error:", err.message);
  process.exit(1);
});
