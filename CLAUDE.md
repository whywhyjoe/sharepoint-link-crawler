# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A **SharePoint Online (SPO) Link Crawler** — a client-side tool that runs *inside*
SharePoint-hosted pages on the BMO tenant (`bmo.sharepoint.com`). It crawls SP
Classic **and Modern** pages from a start URL, extracts and verifies links
(including documents), surfaces broken/denied links, and exports results to XLSX.

It bundles **two crawlers** sharing the same page-loading/scope machinery:
- the **link crawler** (`crawlSite`) — verifies links and document existence; and
- a **text scanner** (`crawlSiteText`) — crawls the same scope looking for literal
  and regex text patterns, returning snippets with context (separate XLSX export).

Author: Joseph Zapert; FCU Creative & Digital Services (joseph.zapert@bmo.com).

## Project layout

Two files matter:

- `link-crawler_spo_2.15.js` — all crawler logic, wrapped in a single IIFE to
  avoid polluting global scope. This is where everything happens. **This is the
  current code.** (The in-file `@version`/console banner reads `2.16` — its
  changelog runs ahead of the filename; trust the file contents.)
- `link-crawler_spo_2.15.html` — the UI markup, delivered as a SharePoint
  "script-include". Loads external deps and the crawler JS, defines all the DOM
  element IDs, and contains **all the CSS inline in a `<style>` block**.
  ⚠️ **The HTML is stale relative to the JS** (still `_2.14`). When asked about
  current behavior, treat the `_2.15.js` file as the source of truth and don't
  assume the HTML wires up everything the JS now supports (e.g. the text-scanner
  module has no committed UI here).

**CSS is intentionally inlined in the HTML.** A standalone `link-crawler_SPO.css`
may exist in the directory but is **not used** — it was folded into the HTML's
`<style>` block on purpose, so there's one fewer file to upload and cache-bust on
each deploy. Make all style changes in the HTML's `<style>` block; ignore the
separate CSS file.

## No build / test / lint

This is vanilla HTML/CSS/JS with **no toolchain** — no package.json, no bundler,
no test runner. Do not look for `npm`/build commands.

**"Running" it = deploying to SharePoint.** The HTML include and JS are hosted
in SharePoint document libraries and referenced by absolute URLs. To ship a change:
1. Upload the edited file(s) to the corresponding SharePoint location.
2. **Bump the cache-bust query param** on the `<script>` URL(s) in the HTML
   (e.g. `...link-crawler_spo_2.15.js?cd=141`). SharePoint caches aggressively;
   without bumping this, changes won't load. (Inlining the CSS means style changes
   ride along with the HTML and need no separate cache-bust.)
3. Reload the SharePoint page hosting the include and watch the console — the code
   logs heavily under the `[LinkCrawler]` prefix.

Versioning is done **by filename** (`_2.15`) plus the cache-bust param, not git
tags. (The in-file `@version`/changelog may read higher than the filename — the
changelog is the most reliable record of what changed.)

## Architecture & the HTML↔JS contract

The HTML and JS are coupled **by element ID**. The JS (`document.getElementById(...)`)
reads config from inputs and writes results into specific nodes. If you rename or
remove an ID in the HTML, the JS silently breaks. Key IDs:

- Config inputs: `startUrl`, `maxDepth`, `crawlDelay`, `verifyConcurrency`
  (parallel link checks per page, optional), `urlMustStartWith` (crawl scope),
  `destMustStartWith` (link/href scope), plus auto-fill checkboxes
  `autoUrlMustStartWith` / `autoDestMustStartWith`.
- Filters: `urlMustNotContain`, `destUrlExclusion`, `externalOnly`,
  `headerFooterSelectors`, `navSelectors`, `excludeNavLinks`.
- Document mode: `documentExtensions`, `documentsOnly`, `checkViaSharePoint`.
- Controls: `startBtn`, `stopBtn`, `exportExcelBtn`, `exportDocsExcelBtn`.
- Live output: `pagesCount`, `linksCount`, `errorsCount`, `queueCount`,
  `progressFill`, `progressText`, `crawlStatus`, `allLogBody`, `errorLogBody`,
  and `*Warning` divs for field validation.

### External runtime dependencies (loaded via `<script>`, not npm)
- **PnP JS** (`pnp2.bundle.js`, exposes `$pnp`) — optional SharePoint REST helper;
  the code degrades gracefully to plain `fetch` if it's unavailable.
- **ExcelJS** (`exceljs.min.js`) — XLSX export.

### Page loading — Modern SharePoint rendering (shared, module scope)
Modern SP pages are SPFx/React apps: a plain `fetch` only sees the server *shell*
(a handful of anchors); the real content (web parts, link lists, document cards)
hydrates client-side seconds later. Two shared module-scope helpers handle this
and are used by **both** crawlers:
- `loadRenderedPageDocument(pageUrl)` — loads the page in a hidden, same-origin
  iframe, lets SP's own JS run, **scrolls every scroll region to the bottom on each
  poll** (to trip virtualized/lazy-loaded library rows), and waits for the anchor
  count to *settle* (unchanged across `STABLE_POLLS` after `MIN_WAIT`, capped at
  `MAX_WAIT` ≈ 20s). Returns a detached `Document` of the rendered DOM. Only works
  same-origin — cross-origin iframes block `contentDocument`.
- `loadCrawlablePage(pageUrl)` — fetches the shell **and** (same-origin) renders it,
  then returns the richer doc plus a de-duplicated **UNION of anchors** from shell
  + rendered DOM (`unionAnchorsByHref`), so neither load is wasted and links present
  in only one source aren't dropped. Cross-origin / failed fetches fall back to the
  shell. **Use this, not raw `fetch`, to load any page** — `fetchWithTimeout` and
  these helpers live at module scope precisely so the text scanner shares them
  (a regression earlier had the text scanner calling an out-of-scope helper and
  only ever seeing the empty shell).

### Link crawler logic (in `crawlSite()`, the IIFE core)
- BFS over a `queue` of `{url, depth}`, with `visited`/`queued` sets for dedupe and
  a `verifyCache` so each destination URL is only checked once. Pages are loaded via
  `loadCrawlablePage`; per-page link checks run in parallel through `runPool` bounded
  by `verifyConcurrency`.
- Link verification (`verify` / `verifyCached`): `fetch` (HEAD for docs, GET for
  HTML, via `fetchWithTimeout`) and, when "Check via SharePoint API" is on,
  SharePoint REST (`GetFileByServerRelativePath(decodedurl=...)`) through
  `buildSpFileApiUrl` / `spExists`.
- **SharePoint sharing links** (`/:x:/r/…`, `/:x:/s/…`) are detected by
  `isSharingLinkLikeUrl` / `isTokenSharingLinkLikeUrl` / `isSubsiteSharingLink` and
  resolved to a server-relative path first via `resolveShareLinkViaRemoteWeb`
  (`SP.RemoteWeb(@a)/web/GetFileByUrl` — a documented-but-"unsupported" fallback);
  resolved hits are labeled `Ok (SP Sharing)`. If resolution fails the code does
  **not** fall back to a raw existence check.
- `spExists` returns existence + file metadata, then **best-effort enrichment**
  (`ListItemAllFields`, `Author`, `ModifiedBy`) for author/editor/created/modified
  columns. Enrichment is non-fatal: a confirmed file must never be downgraded to
  `Error (SP)` just because enrichment failed (a 2.16 fix). `GetUserById` resolves
  author/editor IDs to name/email (cached).
- Error detection is both HTTP status **and text-based** — it matches friendly SP
  pages (`notFoundTexts`, `accessDeniedTexts`) because SP often returns 200 for
  "this item no longer exists" / "request access".
- `CRAWLER_DEFAULTS` holds the baseline config (exclusion lists, selectors,
  concurrency `verifyConcurrency`, timeouts) merged with UI values in `crawlSite`.

### Text scanner module (`crawlSiteText`, near the bottom of the IIFE)
A self-contained second crawler that reuses the scope/queue rules and the shared
modern-rendering path but does **no link verification**. It compiles literal phrases
and regex strings (`compileTextPatterns`), scans content blocks (`getScanBlocks`,
skipping header/footer/nav), normalizes text (soft-hyphen / zero-width stripping,
whitespace collapse) so matches survive HTML boundaries, and emits rows with a
`[[match]]` snippet plus a compact `cssPath` selector. Config baseline is
`TEXT_DEFAULTS`; export via `exportTextMatchesExcelJS`. It follows `.aspx`/`.html`/
`.htm` pages (matching the link crawler) rather than only extension-less URLs.
**Note:** there is no committed UI wiring for it in the stale HTML — it's currently
console-/programmatically-driven.

### Scope/validation rules (important, BMO-specific)
URL scope logic lives in `normalizeUrlForCompare`, `getStartUrlScopeInfo`,
`urlMatchesRequiredPrefix`, and `validateCrawlerUrlConstraints`:
- **Internal** = `https://bmo.sharepoint.com/sites/{name}/...` or `/teams/{name}/...`.
  Crawl + destination scope are enforced to stay within that site/team prefix.
- The tenant root and bare `/sites` / `/teams` are rejected as invalid start URLs.
- **External** URLs are allowed but **max depth is clamped to 5**.
- Prefix matching is path-boundary-safe (`/foo` must not match `/foobar`) — preserve
  this when touching scope code.

## Conventions to follow when editing

- Keep all crawler code inside the IIFE; expose nothing globally.
- Keep verbose `[LinkCrawler]` console logging — it's the primary debugging tool
  since there are no tests.
- Make all styling changes in the HTML's inline `<style>` block (not the standalone
  CSS file). The styles are layered and sometimes self-overriding
  (base → "Copilot" tweaks → "Fluent polish overrides") — later rules win, so add
  refinements at the end rather than hunting for the original rule.
- After any change to the hosted JS, bump its cache-bust query param in the HTML.

**See full coding and documentation conventions:
* `CODE-CONVENTIONS.md` - JS coding style, SPA navigation, and library usage.
* `DOC-CONVENTIONS.md` - docstring and comment style.**