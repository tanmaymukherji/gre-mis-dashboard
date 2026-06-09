import { createReadStream, statSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import { createRequire } from "module";

const require = createRequire(import.meta.url);

// TUS client is available from the hostinger MCP package
const tus = require("C:\\Users\\tmukh\\AppData\\Roaming\\npm\\node_modules\\hostinger-api-mcp\\node_modules\\tus-js-client");

const TOKEN = "tsAkruR0VW24ayQyL0TRlWTuoxXdbjUSzluPhrSS16c6ca26";
const BASE = "https://developers.hostinger.com";
const USERNAME = "u973202051";
const DOMAIN = "gre.grameee.org";
// Update these paths/names for each deploy
const ARCHIVE_PATH = "C:\\Users\\tmukh\\AppData\\Local\\Temp\\gre-mis_20260609.zip";
const ARCHIVE_NAME = "gre-mis_20260609.zip";

async function main() {
  // 1. Get upload credentials
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
  console.log(`Upload URL obtained: ${uploadUrl}`);

  // 2. Pre-upload POST
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

  // 3. TUS upload
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

  // 4. Trigger deployment
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
  console.log("Deployment triggered success:", JSON.stringify(deployResult));
  console.log("Done! Files deploying to gre.grameee.org.");
}

main().catch((err) => {
  console.error("Error:", err.message);
  process.exit(1);
});
