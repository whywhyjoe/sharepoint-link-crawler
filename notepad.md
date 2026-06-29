# Notepad — changes & fixes log

Running log of changes/fixes as we work. Newest first.

---

## 2026-06-29 — Root-scope match bug → zero result rows

**File:** `link-crawler_spo_2.16.js`

**Symptom:** Crawl visited pages and logged "Page links collected" (e.g. 13 links),
but the UI log stayed empty and the console reported `Crawl finished {results: 0}`.

**Cause:** With `destMustStartWith` set to a bare tenant root (`https://bmo.sharepoint.com/`),
`normalizeUrlForCompare` reduces the pathname to `"/"`. The destination-scope filter then
ran `candidate.pathname.startsWith("/" + "/")` → `startsWith("//")`, which is never true for
a real path like `/sites/bmocentral`. So every link failed the `destMustStartWith` check and
was `continue`d before being pushed to `results`. Page *discovery* still worked because it
uses `urlMustStartWith` (a non-root path), which is why pages were crawled but no rows emitted.

**Fix:** Added a root-only guard to both scope-matching functions — when the required
pathname is `"/"`, match everything on the same origin:
- `hrefStartsWithScope` (~line 494)
- `urlMatchesRequiredPrefix` (~line 295)

Non-root scopes are unchanged (still path-boundary-safe).

**Also documented in:** `CLAUDE.md` scope/validation section.

**Note (not a code bug):** The "went to depth 4" observation wasn't reproduced — the logs
only reach depth 2 (correct for `maxDepth: 2`). Likely a stale cached build; the script URL
had a doubled query string (`?cd=148?v=...`). Bump a single clean `?cd=` param on re-deploy.

**Reminder:** `SP_TENANT_ORIGIN` is a placeholder (`[company]`) kept generic for the public
repo. Set it to the real tenant locally when testing or the start URL is misclassified as
external.
