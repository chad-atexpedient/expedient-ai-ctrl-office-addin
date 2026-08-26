import { chromium } from "playwright";

const cliArgs = process.argv.slice(2).filter((arg) => arg !== "--");
const origin = (cliArgs[0] || "https://localhost:3000").replace(/\/+$/, "");
const browser = await chromium.launch({ headless: true });
const results = [];

for (const viewport of [
  { name: "desktop", width: 1000, height: 760 },
  { name: "narrow", width: 360, height: 640 },
]) {
  const page = await browser.newPage({ viewport: { width: viewport.width, height: viewport.height }, ignoreHTTPSErrors: true });

  // Track failed/erroring HTTP responses by URL so allowlisted static-mode
  // 404s (e.g. /api/settings when no backend is running) can be filtered
  // precisely. Console message text for resource-load failures does not
  // reliably include the failing URL across browser versions, so string
  // matching on console text alone is not a safe filter.
  const failedResponses = [];
  page.on("response", (response) => {
    if (response.status() >= 400) failedResponses.push({ url: response.url(), status: response.status() });
  });

  // Non-network console errors (script exceptions, React warnings promoted
  // to errors, etc.) are still tracked directly from console text.
  const consoleErrors = [];
  page.on("console", (message) => {
    if (message.type() !== "error") return;
    if (/^Failed to load resource/.test(message.text())) return; // covered by response tracking above
    consoleErrors.push(message.text());
  });

  await page.goto(`${origin}/`, { waitUntil: "networkidle", timeout: 30_000 });
  const setupVisible = await page.locator(".setup-note").count();
  await page.getByRole("button", { name: "Open settings" }).click();
  await page.locator(".settings-scroll select").first().selectOption("managed");
  const draftDisclosure = await page.getByText(/Mode preview only/).count();
  const managedKeyDisabled = await page.locator('input[placeholder="Configured by your organization"]').isDisabled();
  await page.getByRole("button", { name: "Close settings" }).click();

  const allowStaticApi = process.env.UI_SMOKE_ALLOW_STATIC_API === "true";
  const unexpectedFailedResponses = failedResponses.filter((entry) => {
    if (allowStaticApi && entry.status === 404 && entry.url.includes("/api/settings")) return false;
    return true;
  });
  const allErrors = [
    ...consoleErrors,
    ...unexpectedFailedResponses.map((entry) => `HTTP ${entry.status}: ${entry.url}`),
  ];

  results.push({ viewport: viewport.name, setupVisible, draftDisclosure, managedKeyDisabled, consoleErrors: allErrors });
  await page.close();
}

await browser.close();
const failures = results.filter((result) => result.draftDisclosure !== 1 || !result.managedKeyDisabled || result.consoleErrors.length);
for (const result of results) console.log(`${failures.includes(result) ? "FAIL" : "PASS"} ${result.viewport} ${JSON.stringify(result)}`);
process.exit(failures.length ? 1 : 0);
