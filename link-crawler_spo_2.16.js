/**
 * @fileoverview SharePoint Link Crawler + Text Scanner (Classic/Modern pages)
 *
 * A self-contained crawler utility wrapped in an IIFE to avoid global scope pollution.
 * Features:
 * - HTML page crawling with configurable depth, delay, and URL scope constraints
 * - Link extraction and verification via:
 *   - Fetch (HEAD/GET)
 *   - SharePoint REST API checks for documents (optional)
 * - Error detection (HTTP + friendly SharePoint page text detection)
 * - UI integration hooks for row streaming, stats, and warnings
 * - Export to CSV and XLSX (ExcelJS), with optional SharePoint metadata columns
 * - Optional text scanning module for literal/regex pattern discovery with snippets
 *
 * - NEW in 2.16: Sharepoint Modern crawling
 *
 * @author Joseph Zapert; [Company] Creative & Digital Services; joseph.zapert@[company].com
 * @created Created: 2026/03/10 12:56:04
 * @lastmodified Last modified: 2026/06/27 18:56:31
 * @version 2.16
 *
 * @changelog
 * - 2.16 (2026/06/27)
 *   - FIX: loadRenderedPageDocument() no longer hangs the whole crawl on a
 *     cross-origin redirect (e.g. session-expiry login). The hard-timeout
 *     capture() now catches the SecurityError thrown by iframe.contentDocument
 *     and returns null (falls back to the server shell) instead of throwing in
 *     the timer callback and leaving the page promise forever pending.
 *   - FIX: prevented "zombie" polling loops when iframe onload fires more than
 *     once (redirects). onload now bails if already settled and clears any
 *     pending poll before restarting; pollForRenderedAnchors() also returns
 *     early once settled. Stops concurrent poll chains + stranded timers that
 *     could keep the hard timeout reachable (the trigger for the hang above).
 *   - FIX: loadCrawlablePage() now re-throws AbortError instead of swallowing it
 *     and processing the server shell, so a stopped crawl short-circuits the
 *     in-flight page consistently with an aborted initial fetch.
 *   - FIX: spExists() no longer reports an existing file as "Error (SP)" when the
 *     optional Author/ModifiedBy enrichment fails — enrichment is now best-effort
 *     and the confirmed file/item metadata is always returned. (Also fixed an
 *     undefined `e` reference in that catch block.)
 *   - Modern rendering reworked into a shared module-scope loadCrawlablePage():
 *     the fetched server shell and the rendered iframe DOM are now UNIONED by
 *     absolute href, so neither load is wasted and links present in only one
 *     source are no longer dropped (previously it kept whichever DOM had more).
 *   - Rendered pages are now scrolled to the bottom (window + SharePoint content
 *     scroll-regions) on every poll so virtualized / lazy-loaded library rows
 *     and link lists are captured. MAX_WAIT raised 15s -> 20s for headroom.
 *   - Text scanner (crawlSiteText) now renders modern pages too: it was calling
 *     an out-of-scope fetchWithTimeout (would throw) and only ever saw the empty
 *     shell. fetchWithTimeout + loadRenderedPageDocument were moved to module
 *     scope and shared; the text scanner now follows .aspx/.html/.htm pages
 *     (previously only extension-less URLs were queued).
 * - 2.15: SharePoint Modern crawling (iframe-rendered DOM capture)
 *
 * @example
 * // Standard crawl:
 * // Set Start URL, depth, and filters in the UI, then click "Start Crawl".
 *
 * @example
 * // Document-only crawl with SharePoint verification:
 * // Enable "Check documents only" + "Check via SharePoint API".
 *
 * @example
 * // Export:
 * // Use "Export XLSX" for all rows or "Export Docs XLSX" for document rows only.
 */

console.log("[LinkCrawler] Loaded pre-IIFE");
console.log(`[LinkCrawler] v2.16 — Last modified: 2026/06/27 18:56:31
	`);

(() => {
	const DEBUG_SP = true;
	const LOG_PREFIX = "[LinkCrawler]";
	const log = (...args) => console.log(LOG_PREFIX, ...args);
	const warn = (...args) => console.warn(LOG_PREFIX, ...args);
	const error = (...args) => console.error(LOG_PREFIX, ...args);
	const safeClone = (obj) => {
		try { return JSON.parse(JSON.stringify(obj)); }
		catch { return obj; }
	};

	log("Script initialized", {
		page: location.href,
		version: "2.16",
		lastModifiedBanner: "2026/06/27 18:56:31"
	});

	// === SharePoint tenant configuration ===============================
	// Single source of truth for the organization's SharePoint Online
	// tenant origin. Set this to your tenant, e.g.
	// "https://contoso.sharepoint.com". Used for internal/external scope
	// detection and for building the required crawl-scope prefix. The
	// site/team name is derived dynamically from the entered Start URL, so
	// only the tenant origin needs to be configured here.
	const SP_TENANT_ORIGIN = "https://[company].sharepoint.com";

	const CRAWLER_DEFAULTS = {
		startUrl: "",
		urlMustStartWith: location.origin,
		maxDepth: 1,
		destMustStartWith: "", // catalog constraint for destination URLs
		externalOnly: false,
		crawlDelay: 100,
		verifyConcurrency: 8, // NEW: parallel link checks per page
		urlMustNotContain: [
			"_catalogs", "[company]-my", "spcommon.png", "siteicon.png",
		],
		destUrlExclusion: ["mailto:", ".ashx", "javascript:", "[company]-my", "/FR/"],
		// NOTE: the leading IDs (#companyCentralFooter, #companyPortals,
		// #company-ribbonRow) are tenant-specific chrome selectors —
		// customize them to match your SharePoint master page / branding.
		headerFooterSelectors: [
			"#DeltaPlaceHolderLeftNavBar, #companyCentralFooter, #companyPortals, #desktop-bluebar, #desktop-darkbluebar, #company-ribbonRow, #titleRow, #CommentsWrapper, #spCommandBar, #SuiteNavWrapper, #sp-appBar, [data-automationid='SiteHeader'], [id*='RecommendedItemsWebPart']",
		],
		navSelectors: ["#desktop-whitebar", ".ms-Nav"],
		documentExtensions: ["pdf", "docx", "xlsx", "pptx", "mp4"],
		enablePnP: true,
		notFoundText: "This item may no longer exist or has changed location.",
		notFoundTexts: [
			"This item may no longer exist or has changed location.",
			"The page you're looking for doesn't exist.",
		],

		accessDeniedTexts: [
			"Sorry, you don't have access",
			"Request Access",
			"You need permission to access this site",
		],
		pnpTimeout: 5000,
	};

	/* ---
	   Tiny helper utilities
	--- */



	function logSpDebug(label, data) {
		if (!DEBUG_SP) return;
		log(`SP DEBUG → ${label}`, data);
	}

	/**
	 * Normalize a URL for comparison.
	 * - lowercases origin + pathname
	 * - strips query/hash
	 * - strips trailing slash (except root)
	 * - supports relative URLs by resolving against location.origin
	 * @param {string} rawUrl
	 * @returns {{
	 *   valid: boolean,
	 *   origin: string,
	 *   pathname: string,
	 *   normalized: string,
	 *   url: URL | null
	 * }}
	 */
	function normalizeUrlForCompare(rawUrl) {
		try {
			const u = new URL((rawUrl || "").trim(), location.origin);
			const origin = (u.origin || "").toLowerCase();
			let pathname = (u.pathname || "/").toLowerCase();

			// collapse duplicate slashes
			pathname = pathname.replace(/\/{2,}/g, "/");

			// strip trailing slash unless root
			if (pathname.length > 1) {
				pathname = pathname.replace(/\/+$/, "");
			}

			return {
				valid: true,
				origin,
				pathname,
				normalized: `${origin}${pathname}`,
				url: u,
			};
		} catch {
			return {
				valid: false,
				origin: "",
				pathname: "",
				normalized: "",
				url: null,
			};
		}
	}

	/**
	 * Determine whether a start URL is:
	 * - invalid root-level [Company] SharePoint URL
	 * - internal [Company] site/team URL
	 * - external URL
	 *
	 * Internal means (SP_TENANT_ORIGIN is the configured tenant origin):
	 *   ${SP_TENANT_ORIGIN}/sites/{siteName}/...
	 *   ${SP_TENANT_ORIGIN}/teams/{teamName}/...
	 *
	 * Required prefix becomes:
	 *   ${SP_TENANT_ORIGIN}/sites/{siteName}
	 *   ${SP_TENANT_ORIGIN}/teams/{teamName}
	 *
	 * @param {string} rawStartUrl
	 * @returns {{
	 *   valid: boolean,
	 *   isInternal: boolean,
	 *   isExternal: boolean,
	 *   isInvalidRoot: boolean,
	 *   requiredPrefix: string,
	 *   message: string
	 * }}
	 */
	function getStartUrlScopeInfo(rawStartUrl) {
		const info = normalizeUrlForCompare(rawStartUrl);
		if (!info.valid) {
			return {
				valid: false,
				isInternal: false,
				isExternal: false,
				isInvalidRoot: false,
				requiredPrefix: "",
				message: "Enter a valid Start URL.",
			};
		}

		const isTenantSp = info.origin === SP_TENANT_ORIGIN;
		const path = info.pathname;

		// Explicitly invalid root/start URLs
		if (
			isTenantSp &&
			(path === "/" || path === "/sites" || path === "/teams")
		) {
			return {
				valid: true,
				isInternal: false,
				isExternal: false,
				isInvalidRoot: true,
				requiredPrefix: "",
				message: "Start URL cannot be the tenant root, /sites, or /teams. Enter a specific SharePoint site or team URL.",
			};
		}

		// Internal site/team root detection
		const internalMatch = path.match(/^\/(sites|teams)\/([^\/]+)/i);
		if (isTenantSp && internalMatch) {
			const managedPath = internalMatch[1].toLowerCase();
			const siteName = internalMatch[2];
			const requiredPrefix = `${SP_TENANT_ORIGIN}/${managedPath}/${siteName}`;

			return {
				valid: true,
				isInternal: true,
				isExternal: false,
				isInvalidRoot: false,
				requiredPrefix,
				message: "",
			};
		}

		// Anything else is treated as external per your rule
		return {
			valid: true,
			isInternal: false,
			isExternal: true,
			isInvalidRoot: false,
			requiredPrefix: "",
			message: "External link. You must have permission to crawl this page. Max depth is 5.",
		};
	}

	/**
	 * Returns true if candidate URL is equal to or underneath requiredPrefix.
	 * Comparison ignores query/hash/case/trailing slash.
	 *
	 * Examples:
	 * - prefix: /sites/foo
	 *   candidate: /sites/foo           => true
	 *   candidate: /sites/foo/page.aspx => true
	 *   candidate: /sites/foobar        => false
	 *
	 * @param {string} candidateUrl
	 * @param {string} requiredPrefix
	 * @returns {boolean}
	 */
	function urlMatchesRequiredPrefix(candidateUrl, requiredPrefix) {
		const candidate = normalizeUrlForCompare(candidateUrl);
		const required = normalizeUrlForCompare(requiredPrefix);

		if (!candidate.valid || !required.valid) return false;
		if (candidate.origin !== required.origin) return false;

		// Root-only scope ("/") matches everything on the same origin. Without
		// this, the boundary check below would test startsWith("//") and reject
		// every non-root path (e.g. a bare-tenant destMustStartWith dropped all rows).
		if (required.pathname === "/") return true;

		return (
			candidate.pathname === required.pathname ||
			candidate.pathname.startsWith(required.pathname + "/")
		);
	}

	function getRecommendedScopeFromStartUrl(rawStartUrl) {
		const scope = getStartUrlScopeInfo(rawStartUrl);
		if (!scope.valid || scope.isInvalidRoot) return "";

		try {
			const u = new URL(rawStartUrl, location.origin);
			let path = (u.pathname || "/").replace(/\/{2,}/g, "/");
			const lastSegment = path.split("/").filter(Boolean).pop() || "";

			if (/\.[a-z0-9]{1,10}$/i.test(lastSegment)) {
				path = path.slice(0, path.lastIndexOf("/") + 1) || "/";
			}

			if (path.length > 1) {
				path = path.replace(/\/+$/, "");
			}

			return `${u.origin}${path}`;
		} catch {
			return scope.requiredPrefix;
		}
	}

	function getRecommendedUrlScopeFromStartUrl(rawStartUrl) {
		return getRecommendedScopeFromStartUrl(rawStartUrl);
	}

	function getRecommendedDestinationScopeFromStartUrl(rawStartUrl) {
		return getRecommendedScopeFromStartUrl(rawStartUrl);
	}

	function isAutoFillEnabled(elementId, fallback = true) {
		const el = document.getElementById(elementId);
		return el ? !!el.checked : fallback;
	}

	function getOrCreateFieldWarningElement(elementId) {
		const existing = document.getElementById(elementId);
		if (existing) return existing;

		const fieldIdMap = {
			startUrlWarning: "startUrl",
			destMustStartWithWarning: "destMustStartWith",
			urlMustStartWithWarning: "urlMustStartWith",
		};
		const fieldId = fieldIdMap[elementId] || elementId.replace(/Warning$/, "");
		const field = document.getElementById(fieldId);
		if (!field) return null;

		const warning = document.createElement("div");
		warning.id = elementId;
		warning.className = "crawler-field-warning";
		warning.setAttribute("role", "alert");
		warning.setAttribute("aria-live", "polite");
		warning.style.display = "none";
		warning.style.marginTop = "6px";
		warning.style.fontSize = "12px";
		warning.style.color = "#b42318";
		field.insertAdjacentElement("afterend", warning);
		return warning;
	}

	/**
	 * Display field-level warning text.
	 *
	 * @param {string} elementId
	 * @param {string} message
	 */
	function showFieldWarning(elementId, message) {
		const el = getOrCreateFieldWarningElement(elementId);
		if (!el) return;

		el.textContent = message || "";
		el.style.display = message ? "block" : "none";
	}

	/**
	 * Enable/disable Start button.
	 * @param {boolean} disabled
	 */
	function setStartButtonDisabled(disabled) {
		const btn = document.getElementById("startBtn");
		if (!btn) return;
		btn.disabled = !!disabled;
	}


	/**
 * Validate Start URL / Destination URL / Max Depth constraints.
 * Call this on load + input/blur.
 */
	function validateCrawlerUrlConstraints() {
		const startUrlEl = document.getElementById("startUrl");
		const destEl = document.getElementById("destMustStartWith");
		const maxDepthEl = document.getElementById("maxDepth");
		const urlScopeEl = document.getElementById("urlMustStartWith");

		if (!startUrlEl || !maxDepthEl) return true;

		const startUrl = startUrlEl.value || "";
		const scope = getStartUrlScopeInfo(startUrl);
		const recommendedUrlScope = getRecommendedUrlScopeFromStartUrl(startUrl);
		const recommendedDestScope = getRecommendedDestinationScopeFromStartUrl(startUrl);
		const autoUrlScope = isAutoFillEnabled("autoUrlMustStartWith", true);
		const autoDestScope = isAutoFillEnabled("autoDestMustStartWith", true);

		// Clear warnings up front
		showFieldWarning("startUrlWarning", "");
		showFieldWarning("urlMustStartWithWarning", "");
		showFieldWarning("destMustStartWithWarning", "");

		// Invalid URL entirely
		if (!scope.valid) {
			showFieldWarning("startUrlWarning", scope.message);
			setStartButtonDisabled(true);
			return false;
		}

		// Invalid [Company] root URLs: disable start
		if (scope.isInvalidRoot) {
			showFieldWarning("startUrlWarning", scope.message);
			setStartButtonDisabled(true);
			return false;
		}

		// Internal SharePoint URL: enforce crawl-scope boundary
		if (scope.isInternal) {
			if (autoUrlScope && urlScopeEl && recommendedUrlScope && urlScopeEl.value !== recommendedUrlScope) {
				urlScopeEl.value = recommendedUrlScope;
			}

			if (!autoUrlScope && urlScopeEl?.value?.trim() && !urlMatchesRequiredPrefix(urlScopeEl.value, scope.requiredPrefix)) {
				showFieldWarning(
					"urlMustStartWithWarning",
					`Crawl scope must stay within ${scope.requiredPrefix}`
				);
				setStartButtonDisabled(true);
				return false;
			}

			if (autoDestScope && destEl && recommendedDestScope && destEl.value !== recommendedDestScope) {
				destEl.value = recommendedDestScope;
			}

			// Internal URL is valid
			setStartButtonDisabled(false);
			return true;
		}

		// External URL: warn + clamp max depth, but allow crawl
		if (scope.isExternal) {
			showFieldWarning("startUrlWarning", scope.message);

			if (autoUrlScope && urlScopeEl && recommendedUrlScope) {
				urlScopeEl.value = recommendedUrlScope;
			}

			const currentDepth = +maxDepthEl.value || 0;
			if (currentDepth > 5) {
				maxDepthEl.value = 5;
			}

			setStartButtonDisabled(false);
			return true;
		}

		// Fallback
		setStartButtonDisabled(false);
		return true;
	}

	// Normalize a scope URL so comparisons are path-boundary-safe.
	// Ensures trailing slash on pathname.
	function normalizeScope(scope) {
		const u = new URL(scope || location.origin, location.origin);
		const path = u.pathname.endsWith("/") ? u.pathname : u.pathname + "/";
		return u.origin + path;
	}


	function isExternalToScope(url, scope) {
		return !hrefStartsWithScope(url, scope);
	}


	// Path-boundary-safe scope check: prevents /foo matching /foobar
	function hrefStartsWithScope(href, scope) {
		const candidate = normalizeUrlForCompare(href);
		const required = normalizeUrlForCompare(scope);

		if (!candidate.valid || !required.valid) return false;
		if (candidate.origin !== required.origin) return false;

		// Root-only scope ("/") matches everything on the same origin. Without
		// this, the boundary check below would test startsWith("//") and reject
		// every non-root path.
		if (required.pathname === "/") return true;

		return (
			candidate.pathname === required.pathname ||
			candidate.pathname.startsWith(required.pathname + "/")
		);
	}


	function inClosest(el, selectors) {
		return (selectors || []).some((sel) => {
			try { return !!el.closest(sel); }
			catch { return false; }
		});
	}

	function stripHash(url) {
		try {
			const u = new URL(url, location.origin);
			u.hash = "";
			return u.href;
		} catch {
			return url;
		}
	}

	function stripRefParams(url) {
		try {
			const u = new URL(url, location.origin);
			u.searchParams.delete("refPageTitle");
			u.searchParams.delete("refHrefTitle");
			return u.href;
		} catch {
			return url;
		}
	}

	function normalizeSavedUrl(url) {
		return stripRefParams(stripHash(url));
	}

	function isSharingLinkLikeUrl(u) {
		try {
			const x = (u instanceof URL) ? u : new URL(u, location.origin);
			return /\/:[a-z]:\/[rs]\//i.test(x.pathname); // r OR s
		} catch {
			return false;
		}
	}

	function isTokenSharingLinkLikeUrl(u) {
		try {
			const x = (u instanceof URL) ? u : new URL(u, location.origin);
			return /\/:[a-z]:\/s\//i.test(x.pathname); // token-style
		} catch {
			return false;
		}
	}

	function isSubsiteSharingLink(u) {
		try {
			const x = (u instanceof URL) ? u : new URL(u, location.origin);
			const p = x.pathname.toLowerCase();

			// Matches :b:/r:/sites/<site>/<subsite>/...
			const m = p.match(/\/:b:\/r:\/sites\/[^/]+\/([^/]+)\//);
			return !!m; // has a subsite segment
		} catch {
			return false;
		}
	}

	async function resolveShareLinkViaRemoteWeb(linkUrl) {
		try {
			const u = new URL(linkUrl, location.origin);
			const host = u.origin; // hostname of the link (per MS guidance)
			const enc = encodeURIComponent(linkUrl);

			// Workaround pattern from Microsoft doc: SP.RemoteWeb(@a)/web/GetFileByUrl(@a)?@a='{encoded_link}'
			// NOTE: This is documented as "not supported", but is the recommended fallback if you cannot use Graph.
			const api =
				`${host}/_api/SP.RemoteWeb(@a)/web/GetFileByUrl(@a)` +
				`?@a='${enc}'` +
				`&$select=Name,ServerRelativeUrl,TimeCreated,TimeLastModified,Length`;


			logSpDebug("RemoteWeb request", {
				originalUrl: linkUrl,
				apiUrl: api
			});


			const res = await fetch(api, {
				method: "GET",
				headers: { Accept: "application/json;odata=nometadata" },
				credentials: "include",
			});

			if (!res.ok) return null;

			const data = await res.json();

			// With odata=nometadata, SharePoint usually returns the object directly (not wrapped in d)
			// But keep it defensive:
			const f = data?.d ?? data;

			const name = f?.Name;
			const serverRelativeUrl = f?.ServerRelativeUrl;

			if (!serverRelativeUrl) return null;

			return { name, serverRelativeUrl };
		} catch (e) {

			logSpDebug("RemoteWeb error", { linkUrl, error: String(e) });
			return null;
		}
	}


	const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
	const truncate = (txt, len) =>
		txt && txt.length > len ? txt.slice(0, len) + "…" : txt || "";

	const truncateStart = (txt, len, fromEnd = false) =>
		txt && txt.length > len
			? fromEnd
				? "…" + txt.slice(-len)
				: txt.slice(0, len) + "…"
			: txt;

	const csvEscape = (s) => `"${String(s ?? "").replace(/"/g, '""')}"`;

	const isError = (status = "") => /error|missing|timeout|denied|notfound/i.test(status);

	const arrayFromInput = (id) =>
		document
			.getElementById(id)
			.value.split(",")
			.map((s) => s.trim())
			.filter(Boolean);
	/* ---
   Simple input validation
--- */
	function validateConfig(cfg) {
		log("Validating crawl config", {
			startUrl: cfg?.startUrl,
			maxDepth: cfg?.maxDepth,
			crawlDelay: cfg?.crawlDelay
		});

		if (!cfg.startUrl) throw new Error("Start URL is required");
		try { new URL(cfg.startUrl); } catch { throw new Error("Start URL is invalid"); }

		const scope = getStartUrlScopeInfo(cfg.startUrl);
		if (!scope.valid) throw new Error(scope.message || "Start URL is invalid");
		if (scope.isInvalidRoot) throw new Error(scope.message);

		cfg.maxDepth = Number(cfg.maxDepth || 0);
		if (cfg.maxDepth < 1 || cfg.maxDepth > 20) throw new Error("Depth must be between 1 and 20");
		if (scope.isExternal && cfg.maxDepth > 5) {
			log("Clamping external crawl depth to 5", { startUrl: cfg.startUrl, requestedDepth: cfg.maxDepth });
			cfg.maxDepth = 5;
		}

		const recommendedUrlScope = getRecommendedUrlScopeFromStartUrl(cfg.startUrl);
		const recommendedDestScope = getRecommendedDestinationScopeFromStartUrl(cfg.startUrl);
		const autoUrlScope = cfg.autoUrlMustStartWith !== false;
		const autoDestScope = cfg.autoDestMustStartWith !== false;
		const rawUrlScope = (cfg.urlMustStartWith || "").trim();
		if (scope.isInternal) {
			const effectiveUrlScope = autoUrlScope
				? (recommendedUrlScope || scope.requiredPrefix)
				: (rawUrlScope || recommendedUrlScope || scope.requiredPrefix);
			if (!urlMatchesRequiredPrefix(effectiveUrlScope, scope.requiredPrefix)) {
				throw new Error(`Crawl scope must stay within ${scope.requiredPrefix}`);
			}
			cfg.urlMustStartWith = normalizeScope(effectiveUrlScope);
		} else {
			const effectiveUrlScope = autoUrlScope
				? (recommendedUrlScope || location.origin)
				: (rawUrlScope || recommendedUrlScope || location.origin);
			cfg.urlMustStartWith = normalizeScope(effectiveUrlScope);
		}

		// Normalize destMustStartWith ONLY if the user provided it.
		// Blank means "no catalog scope restriction".
		cfg.destMustStartWith = (cfg.destMustStartWith || "").trim();
		if (cfg.externalOnly) {
			cfg.destMustStartWith = "";
		} else if (scope.isInternal && autoDestScope) {
			const effectiveDestScope = recommendedDestScope || cfg.destMustStartWith;
			cfg.destMustStartWith = effectiveDestScope ? normalizeScope(effectiveDestScope) : "";
		} else if (cfg.destMustStartWith) {
			cfg.destMustStartWith = normalizeScope(cfg.destMustStartWith);
		} else {
			cfg.destMustStartWith = "";
		}

		cfg.verifyConcurrency = Math.min(20, Math.max(1, Number(cfg.verifyConcurrency || 8))); // NEW
		cfg.documentExtensions = cfg.documentExtensions.map((e) => e.replace(/^\./, "").toLowerCase());
		log("Config validated successfully");
		return cfg;
	}
	/* ---
   Visual feedback helper for the Export buttons
--- */
	function flash(btn, msg = "✓ Exported", dur = 1200) {
		const original = btn.textContent;
		btn.textContent = msg;
		btn.disabled = true;
		setTimeout(() => {
			btn.textContent = original;
			btn.disabled = false;
		}, dur);
	}

	/**
	 * Extracts the file name from a URL, removing query strings and decoding it.
	 * @param {string} url - The full destination URL.
	 * @returns {string} - The decoded file name, or an empty string if not found.
	 */
	function extractFileName(url) {
		try {
			const parsedUrl = new URL(url);
			const segs = parsedUrl.pathname.split("/");
			let last = segs.pop();
			if (last === "" && segs.length) last = segs.pop(); // only pop twice when trailing slash
			return decodeURIComponent(String(last || "").split("?")[0].split("#")[0]);
		} catch (e) {
			error("Invalid URL in extractFileName", { url, err: String(e) });
			return "";
		}
	}



	/* ---
   Shared page-loading helpers (module scope so BOTH the link crawler and the
   text scanner render modern SharePoint pages identically). Previously
   fetchWithTimeout + loadRenderedPageDocument lived inside crawlSite, which
   left the text scanner calling an out-of-scope fetchWithTimeout.
--- */

	/* Custom fetch with its own timeout, chained to an optional outer abort signal. */
	function fetchWithTimeout(url, options = {}, timeout = 15000, outerSignal) {
		const controller = new AbortController();
		const onOuterAbort = () => controller.abort();

		if (outerSignal) {
			if (outerSignal.aborted) controller.abort();
			else outerSignal.addEventListener("abort", onOuterAbort, { once: true });
		}

		const id = setTimeout(() => controller.abort(), timeout);

		return fetch(url, { ...options, signal: controller.signal }).finally(() => {
			clearTimeout(id);
			if (outerSignal) outerSignal.removeEventListener("abort", onOuterAbort);
		});
	}

	const getDocTitle = (doc) => (doc?.querySelector("title")?.textContent ?? "").trim();

	/**
	 * Union anchors from multiple documents, de-duplicated by absolute href so a
	 * link present in both the server shell and the rendered DOM is kept once.
	 * Earlier lists win on collision (pass the shell first to prefer its node).
	 */
	function unionAnchorsByHref(pageUrl, ...anchorLists) {
		const seen = new Set();
		const out = [];
		for (const list of anchorLists) {
			for (const a of list) {
				const raw = a.getAttribute("href");
				if (!raw) continue;
				let key;
				try { key = new URL(raw, pageUrl).href; } catch { key = raw; }
				if (seen.has(key)) continue;
				seen.add(key);
				out.push(a);
			}
		}
		return out;
	}

	/**
	 * Load a same-origin page in a hidden iframe and return its *fully rendered*
	 * DOM (a detached Document parsed from the iframe's live outerHTML).
	 *
	 * Modern SharePoint pages are SPFx/React apps: the server returns a shell
	 * with only a few anchors and the real content (web parts, link lists,
	 * document cards) is injected client-side seconds later. Plain `fetch`
	 * never sees that content. We let SharePoint's own JS run inside the
	 * iframe, scroll every scroll region to the bottom to trip lazy/virtualized
	 * loaders, then wait for the anchor count to *settle* — i.e. stop growing
	 * for a short window after a minimum warm-up — so we capture the final
	 * page, not the early shell.
	 *
	 * NOTE: only works same-origin; cross-origin iframes block contentDocument.
	 *
	 * @param {string} pageUrl - Same-origin page to render.
	 * @param {AbortSignal} [outerSignal] - Cancels the wait.
	 * @returns {Promise<Document>} Detached document of the rendered page.
	 */
	async function loadRenderedPageDocument(pageUrl, outerSignal) {
		// Tuning: SharePoint modern pages routinely take several seconds to
		// hydrate. Wait at least MIN_WAIT before trusting a stable count, give
		// up to MAX_WAIT total, and treat the count as "settled" once it is
		// unchanged across STABLE_POLLS consecutive checks.
		const POLL_MS = 350;
		const MIN_WAIT_MS = 1800;
		const MAX_WAIT_MS = 20000;   // headroom for virtualized libraries we scroll through
		const STABLE_POLLS = 3;

		// Modern SP content scrolls inside a region, not the window. Jump every
		// scrollable area to its bottom so IntersectionObserver / virtualized web
		// parts load deferred rows; because we re-jump each poll, a page that keeps
		// growing is chased to the end (and the settle logic waits for it to stop).
		const scrollFrameToBottom = (frameWin, frameDoc) => {
			try {
				const docH = Math.max(
					frameDoc.body?.scrollHeight || 0,
					frameDoc.documentElement?.scrollHeight || 0
				);
				frameWin?.scrollTo?.(0, docH);
				const regions = frameDoc.querySelectorAll(
					"[data-automationid='contentScrollRegion'], .SPPageChrome-content, [class*='scrollRegion'], [class*='ScrollRegion']"
				);
				regions.forEach((r) => { try { r.scrollTop = r.scrollHeight; } catch { } });
			} catch { }
		};

		return new Promise((resolve, reject) => {
			const iframe = document.createElement("iframe");
			let settled = false;
			let pollId = null;
			let timeoutId = null;
			let startTime = 0;
			let lastCount = -1;
			let stableHits = 0;

			const cleanup = () => {
				if (pollId) clearTimeout(pollId);
				if (timeoutId) clearTimeout(timeoutId);
				if (outerSignal) outerSignal.removeEventListener("abort", onAbort);
				iframe.onload = null;
				iframe.onerror = null;
				iframe.remove();
			};

			const finish = (err, doc) => {
				if (settled) return;
				settled = true;
				cleanup();
				if (err) reject(err);
				else resolve(doc);
			};

			// Capture the current rendered DOM as a detached document. Best-effort
			// fallback used both on settle and on the hard timeout, so a slow page
			// still yields whatever rendered rather than throwing it all away.
			const capture = () => {
				try {
					const frameDoc = iframe.contentDocument || iframe.contentWindow?.document;
					if (!frameDoc?.documentElement) return null;
					return new DOMParser().parseFromString(frameDoc.documentElement.outerHTML, "text/html");
				} catch (e) {
					return null;
				}
			};

			const onAbort = () => finish(new DOMException("Aborted", "AbortError"));

			const pollForRenderedAnchors = () => {
				if (settled) return;
				if (outerSignal?.aborted) { onAbort(); return; }

				try {
					const frameWin = iframe.contentWindow;
					const frameDoc = iframe.contentDocument || frameWin?.document;
					const elapsed = Date.now() - startTime;

					if (!frameDoc?.documentElement) {
						if (elapsed >= MAX_WAIT_MS) {
							finish(new Error(`Rendered DOM unavailable for ${pageUrl}`));
							return;
						}
						pollId = setTimeout(pollForRenderedAnchors, POLL_MS);
						return;
					}

					// Trip lazy/virtualized loaders before sampling the anchor count.
					scrollFrameToBottom(frameWin, frameDoc);

					const anchorCount = frameDoc.querySelectorAll("a[href]").length;
					const isReady = frameDoc.readyState === "complete";

					// Track stability: count unchanged since last poll => one stable hit.
					if (anchorCount === lastCount) stableHits++;
					else { stableHits = 0; lastCount = anchorCount; }

					const warmedUp = elapsed >= MIN_WAIT_MS;
					const hasSettled =
						warmedUp && isReady && anchorCount > 0 && stableHits >= STABLE_POLLS;

					if (hasSettled || elapsed >= MAX_WAIT_MS) {
						const doc = capture();
						if (doc) finish(null, doc);
						else finish(new Error(`Rendered DOM unavailable for ${pageUrl}`));
						return;
					}

					pollId = setTimeout(pollForRenderedAnchors, POLL_MS);
				} catch (e) {
					// Cross-origin or other access error — caller falls back to fetch.
					finish(e);
				}
			};

			iframe.style.position = "absolute";
			iframe.style.left = "-99999px";
			iframe.style.top = "0";
			iframe.style.width = "1024px";   // give web parts a realistic viewport so they render
			iframe.style.height = "768px";
			iframe.style.border = "0";
			iframe.setAttribute("aria-hidden", "true");

			if (outerSignal) outerSignal.addEventListener("abort", onAbort, { once: true });
			// Hard ceiling slightly above MAX_WAIT_MS: capture best-effort rather than reject.
			timeoutId = setTimeout(() => {
				const doc = capture();
				if (doc) finish(null, doc);
				else finish(new Error(`Rendered DOM timeout for ${pageUrl}`));
			}, MAX_WAIT_MS + 2000);
			iframe.onload = () => {
				if (settled) return;
				if (pollId) clearTimeout(pollId);
				startTime = Date.now();
				pollForRenderedAnchors();
			};
			iframe.onerror = () => finish(new Error(`Rendered DOM load failed for ${pageUrl}`));
			document.body.appendChild(iframe);
			iframe.src = pageUrl;
		});
	}

	/**
	 * Fetch a page and, for same-origin pages, additionally render it in a hidden
	 * iframe so SPFx/React content hydrates. Returns the richest available document
	 * plus a de-duplicated UNION of anchors from the shell AND the rendered DOM, so
	 * neither the fetch nor the render is wasted and no link is dropped.
	 *
	 * @param {string} pageUrl
	 * @param {AbortSignal} [outerSignal]
	 * @param {{log?:Function, warn?:Function}} [logger]
	 * @returns {Promise<{ok:boolean, status:number, doc:Document, anchors:Element[], title:string, fromRendered:boolean}>}
	 */
	async function loadCrawlablePage(pageUrl, outerSignal, logger = {}) {
		const _log = logger.log || (() => { });
		const _warn = logger.warn || (() => { });

		const res = await fetch(pageUrl, { signal: outerSignal, credentials: "include" });
		const status = res.status;
		const ok = res.ok;
		const html = ok ? await res.text() : "";
		const shellDoc = new DOMParser().parseFromString(html, "text/html");
		const shellAnchors = [...shellDoc.querySelectorAll("a[href]")];

		let sameOrigin = false;
		try { sameOrigin = new URL(pageUrl, location.origin).origin === location.origin; } catch { }

		// Cross-origin (can't read iframe content) or a failed fetch: shell is all we get.
		if (!ok || !sameOrigin) {
			return { ok, status, doc: shellDoc, anchors: shellAnchors, title: getDocTitle(shellDoc), fromRendered: false };
		}

		try {
			const renderedDoc = await loadRenderedPageDocument(pageUrl, outerSignal);
			const renderedAnchors = [...renderedDoc.querySelectorAll("a[href]")];
			const anchors = unionAnchorsByHref(pageUrl, shellAnchors, renderedAnchors);
			_log("Rendered DOM parsed", {
				pageUrl,
				shellLinks: shellAnchors.length,
				renderedLinks: renderedAnchors.length,
				unionLinks: anchors.length,
			});
			return { ok, status, doc: renderedDoc, anchors, title: getDocTitle(renderedDoc), fromRendered: true };
		} catch (e) {
			if (e?.name === "AbortError") throw e;
			_warn("Rendered DOM unavailable; using fetched HTML", { pageUrl, err: String(e) });
			return { ok, status, doc: shellDoc, anchors: shellAnchors, title: getDocTitle(shellDoc), fromRendered: false };
		}
	}

	/* ---
   Core crawler packaged as crawlSite()
--- */
	async function crawlSite(userCfg, hooks = {}) {
		const cfg = Object.assign({}, CRAWLER_DEFAULTS, userCfg);
		log("crawlSite started", safeClone({
			startUrl: normalizeSavedUrl(cfg.startUrl),
			maxDepth: cfg.maxDepth,
			crawlDelay: cfg.crawlDelay,
			documentsOnly: cfg.documentsOnly,
			checkViaSharePoint: cfg.checkViaSharePoint
		}));

		const visited = new Set();
		const queued = new Set();          // NEW: dedupe queued pages
		const verifyCache = new Map();     // NEW: dedupe destination checks
		const userCache = new Map(); // id -> Promise<{ title, email } | null>

		async function getUserById(id) {
			if (!id) return null;

			const key = String(id);
			if (!userCache.has(key)) {
				const base = new URL(cfg.urlMustStartWith, location.origin).origin + new URL(cfg.urlMustStartWith, location.origin).pathname.replace(/\/?$/, "/");
				const url = `${base}_api/web/GetUserById(${encodeURIComponent(id)})?$select=Title,Email`;

				userCache.set(
					key,
					fetch(url, {
						headers: { Accept: "application/json;odata=nometadata" },
						credentials: "include",
						signal,
					})
						.then((r) => (r.ok ? r.json() : null))
						.then((j) => (j ? { title: j.Title, email: j.Email } : null))
						.catch(() => null)
				);
			}

			return userCache.get(key);
		}

		const queue = []; // { url, depth }
		const results = [];
		const stats = {
			pages: 0,
			links: 0,
			errors: 0,
			queue: 0,
		};

		const docMode = cfg.documentsOnly;
		const docSet = new Set(
			cfg.documentExtensions.map((e) => e.toLowerCase())
		);
		/*  use one AbortController (comes from UI)  */
		const {
			signal = undefined,
			onRow = () => { },
			onStats = () => { },
			onPnPWarn = () => { },
		} = hooks;
		const aborted = () => signal?.aborted;
		/*  ##### mini helper functions #####  */
		const getExt = (url) => {
			const m = url.pathname.match(/\.([a-z0-9]+)(?:[\?#]|$)/i);
			return m ? m[1].toLowerCase() : "";
		};
		const isExternal = (url) => url.origin !== location.origin;
		const matchesAny = (str, arr) => arr.some((t) => str.includes(t));

		/* ---
	   Optional PnP.js initialisation
	--- */
		let pnpReady = false;
		async function initPnP() {
			if (!cfg.enablePnP || typeof $pnp === "undefined") {
				warn("PnP unavailable or disabled; continuing with fetch-based verification");
				return;
			}
			try {
				await $pnp.setup({
					sp: { baseUrl: location.origin, timeout: cfg.pnpTimeout }
				});
				pnpReady = true;
				log("PnP initialized successfully");
			} catch (e) {
				pnpReady = false;
				warn("PnP initialization failed; using fetch-only verification", e);
				onPnPWarn();
			}
		}



		/* fetchWithTimeout + loadRenderedPageDocument moved to module scope
		   (above crawlSite) so the text scanner can share them. */

		/**
		 * Builds a SharePoint REST API URL for file existence/metadata lookup.
		 * Normalizes sharing-link patterns (/:x:/r/, /:w:/r/, etc.) and targets the
		 * appropriate site root for _api/web/GetFileByServerRelativeUrl().
		 *
		 * @param {string} absHref - Absolute URL to a SharePoint file.
		 * @returns {string} SharePoint REST API endpoint.
		 * @throws {Error} If URL parsing fails.
		 */


		/* COPILOT-EDIT 3-12-16
		Copilot is replacing this function with a version to compensate for subweb sharing links */
		/* Old
				function buildSpFileApiUrl(absHref) {
					try {
						// Normalize “sharing” link formats (/:x:/r/, /:w:/r/, etc.) to tenant origin
						if (/\/:[a-z]:\/r\//i.test(absHref)) {
						log("Normalizing sharing link format", absHref);
						absHref = absHref.replace(/^https?:\/\/[^/]+\/:[a-z]:\/r/i, location.origin);
						}

						const u = new URL(absHref);

						// Determine the site root so this works across /sites/, /teams/, /personal/
						const siteMatch = u.pathname.match(/^\/(sites|teams|personal)\/[^/]+/i);
						const siteRoot = siteMatch ? (u.origin + siteMatch[0]) : u.origin;

						// Server-relative file path must NOT include querystring/hash
						const serverRel = decodeURIComponent(u.pathname); // const serverRel = u.pathname;

						// Use parameter form to avoid quote-escaping issues
						// (encodeURIComponent is safe here; SP will decode @p)
						const p = encodeURIComponent(serverRel);
						return `${siteRoot}/_api/web/GetFileByServerRelativeUrl(@p)?@p='${p}'`;
					} catch (e) {
						error("Failed to build SharePoint API URL", { absHref, err: String(e) });
						throw e;
					}
				}
		*/
		function buildSpFileApiUrl(absHref) {
			try {
				const u = new URL(absHref, location.origin);

				// server-relative path (decoded)
				const serverRel = decodeURIComponent(u.pathname);

				// ✅ decodedurl must preserve '/'
				const p = encodeURI(serverRel);

				// ✅ scope to site root (/sites/X or /teams/X)
				const m = u.pathname.match(/^\/(sites|teams|personal)\/[^/]+/i);
				const webRoot = m ? (u.origin + m[0]) : u.origin;

				return `${webRoot}/_api/web/GetFileByServerRelativePath(decodedurl='${p}')`;
			} catch (e) {
				error("Failed to build SharePoint API URL", { absHref, err: String(e) });
				throw e;
			}
		}


		async function spExists(absHref) {
			log("SharePoint verification started", absHref);

			try {
				const fullUrl = new URL(absHref, location.origin).href;
				const baseUrl = buildSpFileApiUrl(fullUrl);

				logSpDebug("SP Exists call (base)", {
					absHref,
					baseUrl
				});


				const fileUrl =
					`${baseUrl}?$select=Exists,Name,ServerRelativeUrl,Length,TimeCreated,TimeLastModified`;



				logSpDebug("SP Exists call (metadata)", {
					fileUrl
				});



				const res = await fetch(fileUrl, {
					headers: { Accept: "application/json;odata=nometadata" },
					credentials: "include",
					signal,
				});


				if (res.status === 403) { warn("SP denied", absHref); return { status: "Denied (SP)", code: 403 }; }
				if (res.status === 404) { warn("SP not found", absHref); return { status: "NotFound (SP)", code: 404 }; }
				if (!res.ok) {

					const text = await res.text().catch(() => "");
					logSpDebug("SP Exists ERROR", {
						absHref,
						status: res.status,
						response: text.slice(0, 300)
					});
					return { status: "Error (SP)", code: res.status };

				}


				const data = await res.json();
				if (!data.Exists) return { status: "NotFound (SP)", code: 404 };

				const v = { status: "Ok (SP)", code: 200, _sp: {} };

				v._sp.file = {
					name: data.Name,
					serverRelativeUrl: data.ServerRelativeUrl,
					sizeBytes: data.Length,
					timeCreated: data.TimeCreated,
					timeLastModified: data.TimeLastModified,
				};



				try {
					const itemUrl =
						`${baseUrl}/ListItemAllFields` +
						`?$select=Id,Created,Modified,AuthorId,EditorId`;

					logSpDebug("SP File/ListItemAllFields call", { itemUrl });

					const iRes = await fetch(itemUrl, {
						headers: { Accept: "application/json;odata=nometadata" },
						credentials: "include",
						signal,
					});

					if (iRes.ok) {
						const li = await iRes.json();
						v._sp.item = {
							id: li.Id,
							created: li.Created,
							modified: li.Modified,
							authorId: li.AuthorId,
							editorId: li.EditorId,
						};
					}
				} catch { }


				try {
					// baseUrl looks like: .../GetFileByServerRelativeUrl(@p)?@p='...'

					const call = baseUrl; // no ? anymore

					// File -> Author (SP.User)

					// Author
					const authorUrl =
						`${call}/Author?$select=Id,Title,Email,LoginName,UserPrincipalName`;

					logSpDebug("SP File/Author call", { authorUrl });

					const aRes = await fetch(authorUrl, {
						headers: { Accept: "application/json;odata=nometadata" },
						credentials: "include",
						signal,
					});


					if (aRes.ok) {
						const a = await aRes.json();
						v._sp.author = {
							id: a.Id,
							title: a.Title,
							email: a.Email,
							loginName: a.LoginName,
							upn: a.UserPrincipalName,
						};
					}


					if (!aRes.ok) {
						const t = await aRes.text().catch(() => "");
						logSpDebug("SP File/Author ERROR", {
							status: aRes.status,
							response: t.slice(0, 300)
						});

					}


					// File -> ModifiedBy (SP.User)
					const modifiedByUrl = `${call}/ModifiedBy?$select=Id,Title,Email,LoginName,UserPrincipalName`;
					const mRes = await fetch(modifiedByUrl, {
						headers: { Accept: "application/json;odata=nometadata" },
						credentials: "include",
						signal,
					});

					if (mRes.ok) {
						const m = await mRes.json();
						v._sp.modifiedBy = { id: m.Id, title: m.Title, email: m.Email, loginName: m.LoginName, upn: m.UserPrincipalName };
					}

					if (!mRes.ok) {
						const tm = await mRes.text().catch(() => "");
						logSpDebug("SP File/ModifiedBy ERROR", {
							status: mRes.status,
							response: tm.slice(0, 300)
						});

					}

				} catch (enrichErr) {
					// Author/ModifiedBy enrichment is best-effort: a failure here (network
					// blip, abort, or an endpoint that doesn't expose these) must NOT discard
					// the existence + file/item metadata already gathered in v. Fall through
					// and return v with whatever we have.
					logSpDebug("SP author/modifiedby enrich failed (non-fatal)", { absHref, err: String(enrichErr) });
				}



				log("SharePoint verification success", { absHref, code: 200 });

				return v;
			} catch (e) {
				error("SP check failed", { absHref, err: String(e) });
				return { status: "Error (SP)", code: "" };
			}
		}

		/* ---
	   Link verification (HEAD → PnP)
	--- */

		async function verify(u, ext) {
			if (aborted()) return { status: "Aborted", code: "" };

			if (isExternal(u)) {
				return { status: "External", code: "" };
			}

			const isDocument = !!ext && docSet.has(ext);

			const isSharing = isSharingLinkLikeUrl(u);       // now matches /r/ and /s/
			const isTokenShare = isTokenSharingLinkLikeUrl(u);
			const isSubsiteShare = isSubsiteSharingLink(u);


			logSpDebug("Verify routing", {
				href: u.href,
				isSharing,
				isTokenShare,
				willUseRemoteWeb: isTokenShare || isSharing // depending on your final rule
			});

			if (cfg.checkViaSharePoint && (isDocument || isSharing)) {

				// ✅ ALL SharePoint sharing links must be resolved first
				if (isSharing) {
					const resolved = await resolveShareLinkViaRemoteWeb(u.href);

					if (resolved?.serverRelativeUrl) {
						const absResolved = new URL(
							resolved.serverRelativeUrl,
							location.origin
						).href;

						const v = await spExists(absResolved);

						if (v?.status === "Ok (SP)") {
							v.status = "Ok (SP Sharing)";
						}

						// keep resolved name if needed
						if (v?._sp?.file && !v._sp.file.name && resolved.name) {
							v._sp.file.name = resolved.name;
						}

						return v;
					}

					// ✅ If resolution fails, do NOT fall back to spExists(u.href)
					return { status: "Error (SP)", code: 500 };
				}

				// ✅ Non-sharing docs (normal .pdf URLs)
				return spExists(u.href);
			}


			if (isDocument) {
				return fetchWithTimeout(u.href, { method: "HEAD" }, cfg.pnpTimeout, signal)
					.then(r => ({ status: r.ok ? "OK" : "Error", code: r.status }))
					.catch(e => ({
						status: e.name === "AbortError" ? (aborted() ? "Aborted" : "Timeout") : "Unverified",
						code: ""
					}));
			}

			// HTML (non-doc): GET and detect friendly error pages
			try {
				const r = await fetchWithTimeout(u.href, { method: "GET" }, cfg.pnpTimeout, signal);
				const code = r.status;
				if (code === 403) return { status: "Denied", code };
				if (code === 404) return { status: "NotFound", code };
				if (!r.ok) return { status: "Error", code };

				const html = await r.text();
				if (cfg.notFoundTexts?.some(txt => html.includes(txt))) return { status: "NotFound", code };
				if (cfg.accessDeniedTexts?.some(txt => html.includes(txt))) return { status: "Denied", code };
				return { status: "OK", code };
			} catch (e) {
				return e.name === "AbortError"
					? { status: aborted() ? "Aborted" : "Timeout", code: "" }
					: { status: "Unverified", code: "" };
			}
		}


		/* ---
	   Crawl a single page
	--- */


		async function runPool(items, limit, worker) {
			if (!items.length) return;
			let i = 0;
			const n = Math.max(1, Number(limit || 1));
			const workers = Array.from({ length: Math.min(n, items.length) }, async () => {
				while (true) {
					const idx = i++;
					if (idx >= items.length) break;
					await worker(items[idx], idx);
				}
			});
			await Promise.all(workers);
		}

		async function verifyCached(u, ext, isSharingLink = false) {
			const key = u.href;
			if (!verifyCache.has(key)) {
				verifyCache.set(
					key,
					verify(u, ext).catch((e) => ({
						status: e?.name === "AbortError" ? (aborted() ? "Aborted" : "Timeout") : "Unverified",
						code: "",
					}))
				);
			}
			const v = await verifyCache.get(key);

			const out = { ...v, _sp: v?._sp ? safeClone(v._sp) : undefined };


			// If it’s a sharing-link-like URL and SP succeeded, label it
			if (isSharingLink && /^Ok\s*\(SP\)$/i.test(out.status || "")) {
				out.status = "Ok (SP Sharing)";
			}

			return out;

		}

		function enqueuePage(url, depth) {
			if (depth > cfg.maxDepth) return false;
			if (visited.has(url) || queued.has(url)) return false;
			queued.add(url);
			queue.push({ url, depth });
			return true;
		}

		async function crawlPage(pageUrl, depth) {
			log("Crawling page", { pageUrl, depth, queue: queue.length, visited: visited.size });


			if (aborted() || depth > cfg.maxDepth || visited.has(pageUrl)) return;

			visited.add(pageUrl);
			stats.pages++;
			onStats({ ...stats, queue: queue.length });

			try {
				// Load the page. For same-origin pages this also renders it in a hidden
				// iframe so modern SPFx content hydrates; `links` is the de-duplicated
				// UNION of the server shell and the rendered DOM, so neither load is
				// wasted and no anchor is lost. Cross-origin pages can't expose their
				// contentDocument, so they fall back to the fetched shell.
				const page = await loadCrawlablePage(pageUrl, signal, { log, warn });
				if (!page.ok) {
					stats.errors++;
					warn("Page fetch failed", { pageUrl, code: page.status });
					return;
				}

				const links = page.anchors;
				log("Page links collected", { pageUrl, linksFound: links.length, rendered: page.fromRendered });

				const title = page.title;
				const pending = []; // NEW: collect verifications, then run in parallel

				for (const a of links) {
					if (aborted()) break;
					if (inClosest(a, cfg.headerFooterSelectors)) continue;

					const rawHref = a.getAttribute("href")?.trim();
					if (!rawHref || rawHref.startsWith("#")) continue;
					if (!rawHref || matchesAny(rawHref, cfg.destUrlExclusion)) continue;

					let abs;
					try {

						abs = new URL(rawHref, pageUrl);
						abs = new URL(normalizeSavedUrl(abs.href));

					} catch {
						continue;
					}

					const isSharing = isSharingLinkLikeUrl(abs);

					const isExternalLink = isExternalToScope(abs.href, cfg.urlMustStartWith);

					const ext = getExt(abs);


					// Treat SharePoint pages as HTML pages even though they end in .aspx
					const isHtmlPage = !ext || ext === "aspx" || ext === "html" || ext === "htm";


					const isNav = inClosest(a, cfg.navSelectors);
					if (cfg.excludeNavLinks && isNav) continue;

					// queue HTML pages once only


					const shouldQueue =
						isHtmlPage &&
						!isExternal(abs) &&
						hrefStartsWithScope(abs.href, cfg.urlMustStartWith) &&
						!matchesAny(abs.href, cfg.urlMustNotContain);



					if (shouldQueue) enqueuePage(abs.href, depth + 1);

					// Destination scope limits result rows, but should not block page discovery.
					if (cfg.destMustStartWith && !hrefStartsWithScope(abs.href, cfg.destMustStartWith)) {
						continue;
					}

					const isDocument = docSet.has(ext);
					if (docMode && !isDocument && !isSharing) continue; // in doc mode: crawl-only for non-docs

					let text = (a.textContent ?? "").trim().replace(/\s+/g, " ");
					if (!text) {
						const img = a.querySelector("img");
						if (img) {
							const src = img.getAttribute("src") ?? "";
							const parts = src.split("/");
							text = parts[parts.length - 1];
						}
					}
					if (isNav) text = "NAV: " + text;

					if (cfg.externalOnly && !isExternalLink) {
						continue; // skip internal links entirely
					}

					const row = {
						pageTitle: title,
						pageUrl,
						linkText: text,
						destUrl: normalizeSavedUrl(abs.href),
						ext,
						status: "",
						code: "",
						srcFileName: extractFileName(pageUrl),
						destFileName: extractFileName(abs.href),
					};

					pending.push({ row, abs, ext, isSharing });
				}

				let statsTick = 0;

				await runPool(pending, cfg.verifyConcurrency, async ({ row, abs, ext, isSharing }) => {
					if (aborted()) return;

					if (matchesAny(row.destUrl, cfg.urlMustNotContain)) {
						row.status = "Skipped";
					} else {
						const v = await verifyCached(abs, ext, isSharing);
						row.status = v.status;
						row.code = v.code;

						if (v._sp) {
							row.spName = v._sp.file?.name;
							row.spUrl = v._sp.file?.serverRelativeUrl;
							row.spSizeBytes = v._sp.file?.sizeBytes;
							row.spTimeCreated = v._sp.file?.timeCreated;
							row.spTimeModified = v._sp.file?.timeLastModified;
							row.itemId = v._sp.item?.id;
							row.itemCreated = v._sp.item?.created;
							row.itemModified = v._sp.item?.modified;


							if ((!row.ext || row.ext === "html" || row.ext === "htm") && (row.spName || row.spUrl)) {
								const src = row.spName || row.spUrl; // prefer name if available
								const m = String(src).match(/\.([a-z0-9]+)(?:$|[?#])/i);
								if (m) row.ext = m[1].toLowerCase();
							}



							// Keep IDs if you still want them for troubleshooting/export
							row.itemAuthorId = v._sp.item?.authorId;
							row.itemEditorId = v._sp.item?.editorId;

							// NEW: use navigation results (File/Author and File/ModifiedBy)
							row.itemAuthorName = v._sp.author?.title;
							row.itemAuthorEmail = v._sp.author?.email;

							row.itemEditorName = v._sp.modifiedBy?.title;
							row.itemEditorEmail = v._sp.modifiedBy?.email;

						}

						if (isError(row.status)) stats.errors++;
					}

					results.push(row);
					stats.links++;
					onRow(row);

					// reduce UI churn: update every 10 rows (or final)
					statsTick++;
					if (statsTick % 10 === 0 || statsTick === pending.length) {
						onStats({ ...stats, queue: queue.length });
					}
				});

				await sleep(cfg.crawlDelay);
			} catch (e) {
				stats.errors++;
				error("crawlPage exception", { pageUrl, err: String(e) });
			}
		}

		/* ---
	   Main run loop
	--- */
		await initPnP();
		enqueuePage(cfg.startUrl, 0); // NEW (instead of direct queue.push)
		while (queue.length && !aborted()) {
			const { url, depth } = queue.shift();
			await crawlPage(url, depth);
		}
		onStats({
			...stats,
			queue: queue.length,
		});
		if (aborted()) {
			throw new DOMException("Aborted", "AbortError");
		}
		return {
			stats,
			results,
		};
	}
	/* ---
   UI helpers (table, counters, CSV)
--- */



	function addRow(tbodyId, row) {
		const tbody = document.getElementById(tbodyId);

		/* Remove placeholder “no rows yet” row */
		if (tbody.children.length === 1 && tbody.children[0].children.length === 1) {
			tbody.innerHTML = "";
		}

		const tr = document.createElement("tr");

		const cls = row.status.startsWith("OK")
			? "status-ok"
			: isError(row.status)
				? "status-error"
				: row.status === "External"
					? "status-external"
					: row.status === "Skipped"
						? "status-skipped"
						: "";

		// Helper to create a <td> with safe text + optional title/class
		const td = (text, { title, className } = {}) => {
			const cell = document.createElement("td");
			if (className) cell.className = className;
			if (title != null) cell.title = title;
			cell.textContent = text ?? "";
			return cell;
		};

		// 1) Page Title
		tr.appendChild(
			td(truncate(row.pageTitle, 30), {
				title: row.pageTitle ?? "",
			})
		);

		// 2) Link Text
		tr.appendChild(
			td(truncate(row.linkText, 40), {
				title: row.linkText ?? "",
				className: "link-text",
			})
		);

		// 3) Destination URL with copy-to-clipboard
		const destTd = document.createElement("td");
		destTd.className = "dest-url";
		destTd.title = row.destUrl ?? "";

		const span = document.createElement("span");
		span.className = "copyable-url";
		span.style.cursor = "pointer";
		span.title = "Click to copy";
		span.dataset.url = row.destUrl ?? "";
		span.textContent = truncateStart(row.destUrl ?? "", 80, true);

		span.addEventListener("click", (e) => {
			const url = e.currentTarget?.dataset?.url || "";
			if (!url) return;
			// Clipboard can fail in some contexts; ignore errors quietly
			navigator.clipboard?.writeText(url).catch(() => { });
		});

		destTd.appendChild(span);
		tr.appendChild(destTd);

		// 4) Ext
		tr.appendChild(td(row.ext || ""));


		// 5) Status
		tr.appendChild(
			td(row.status ?? "", {
				className: cls,
			})
		);

		// 6) Code
		tr.appendChild(td(String(row.code ?? "")));

		tbody.appendChild(tr);

		/* Auto-scroll to bottom */
		const container = tbody.closest(".log-table-container");
		if (container) container.scrollTop = container.scrollHeight;
	}

	function updateCounters(s) {
		document.getElementById("pagesCount").textContent = s.pages;
		document.getElementById("linksCount").textContent = s.links;
		document.getElementById("errorsCount").textContent = s.errors;
		document.getElementById("queueCount").textContent = s.queue;
		/* Simple “how far through known work we are” progress */
		const total = Math.max(1, s.pages + s.queue);
		const pct = Math.min(100, (s.pages / total) * 100);
		document.getElementById("progressFill").style.width = pct + "%";
		document.getElementById(
			"progressText"
		).textContent = `${s.pages} pages processed`;
	}


	async function exportExcelJS(rows, opts = {}) {
		if (!rows?.length) { alert("Nothing to export!"); return; }

		const includeSp = !!opts.includeSp;

		const baseHeaders = [
			"Page Title", "Page URL", "Filename", "Link Text", "Destination URL",
			"Destination Filename", "File Ext", "Status", "HTTP",
		];
		const spHeaders = [
			"SP: File Name", "SP: ServerRelUrl", "SP: Size (bytes)",
			"SP: Time Created", "SP: Time Modified",
			"Item Id", "Item Created", "Item Modified",
			"Item Author", "Item Editor",
		];
		const headers = includeSp ? [...baseHeaders, ...spHeaders] : baseHeaders;

		const baseRow = (r) => ([
			r.pageTitle ?? "", r.pageUrl ?? "", r.srcFileName ?? "", r.linkText ?? "", r.destUrl ?? "",
			r.destFileName ?? "", r.ext ?? "", r.status ?? "", r.code ?? "",
		]);
		const spRow = (r) => ([
			r.spName ?? "", r.spUrl ?? "", r.spSizeBytes ?? "",
			r.spTimeCreated ?? "", r.spTimeModified ?? "",
			r.itemId ?? "", r.itemCreated ?? "", r.itemModified ?? "",
			r.itemAuthorName ?? "", r.itemEditorName ?? "",
		]);

		const tableData = rows.map(r => includeSp ? [...baseRow(r), ...spRow(r)] : baseRow(r));

		const workbook = new ExcelJS.Workbook();
		const sheet = workbook.addWorksheet("Crawl Results");

		sheet.addTable({
			name: "CrawlResults",
			ref: "A1",
			headerRow: true,
			style: { theme: "TableStyleLight9", showRowStripes: true },
			columns: headers.map(h => ({ name: h })),
			rows: tableData,
		});

		const baseWidths = [30, 40, 30, 25, 50, 30, 10, 15, 10];
		const spWidths = [30, 50, 16, 22, 22, 10, 22, 22, 18, 18];
		const widths = includeSp ? [...baseWidths, ...spWidths] : baseWidths;

		sheet.columns = widths.map(w => ({ width: w }));
		sheet.eachRow(row => row.eachCell(cell => { cell.font = { size: 10 }; }));

		const buffer = await workbook.xlsx.writeBuffer();
		const blob = new Blob([buffer], {
			type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
		});
		const ts = new Date().toISOString().replace(/[:T]/g, "-").slice(0, 19);
		const a = document.createElement("a");
		a.href = URL.createObjectURL(blob);
		a.download = `crawl-${ts}.xlsx`;
		a.click();
		URL.revokeObjectURL(a.href);
	}

	/* ---
   UI event wiring
--- */
	let abortCtrl = null;
	let crawlResults = [];
	let crawlErrors = [];
	/*  Tab switching (“All” / “Errors”)  */
	document.querySelectorAll(".log-tab").forEach((tab) => {
		tab.addEventListener("click", () => {
			document
				.querySelectorAll(".log-tab")
				.forEach((t) => t.classList.remove("active"));
			tab.classList.add("active");
			const which = tab.dataset.tab;
			document.getElementById("allLogContainer").style.display =
				which === "all" ? "block" : "none";
			document.getElementById("errorLogContainer").style.display =
				which === "errors" ? "block" : "none";
		});
	});
	/*  Start Crawl  */
	document.getElementById("startBtn").addEventListener("click", async () => {
		log("Start Crawl clicked");
		let cfg;

		try {
			cfg = validateConfig(getUIConfig());
			lastRunConfig = cfg;

			log("Start Crawl config resolved", safeClone(cfg));
		} catch (e) {
			error("Start Crawl validation error", String(e));
			document.getElementById("crawlStatus").textContent = `Error: ${e.message}`;
			return;
		}

		/*  Prepare UI  */
		document.getElementById("crawlStatus").textContent = "Starting…";
		document.getElementById("startBtn").disabled = true;
		document.getElementById("stopBtn").disabled = false;
		document.getElementById("exportExcelBtn").disabled = true;
		document.getElementById("exportDocsExcelBtn").disabled = true;
		/*  Clear old results  */
		crawlResults = [];
		crawlErrors = [];
		document.getElementById("allLogBody").innerHTML =
			'<tr><td colspan="6" style="text-align:center;padding:24px">Running…</td></tr>';
		document.getElementById("errorLogBody").innerHTML =
			'<tr><td colspan="6" style="text-align:center;padding:24px">No errors</td></tr>';
		/*  Kick off crawl  */
		abortCtrl = new AbortController();
		document.getElementById("crawlStatus").textContent = "Running…";

		try {
			const { results, stats } = await crawlSite(cfg, {
				signal: abortCtrl.signal,
				onRow: (row) => {
					addRow("allLogBody", row);
					if (isError(row.status)) {
						addRow("errorLogBody", row);
						crawlErrors.push(row);
					}
					crawlResults.push(row);
				},
				onStats: (s) => {
					updateCounters(s);
					if (!abortCtrl?.signal?.aborted) {
						document.getElementById("crawlStatus").textContent =
							`Running… ${s.pages} pages, ${s.links} links, ${s.errors} errors`;
					}
				},
				onPnPWarn: () => {
					warn("PnP warning surfaced to UI");
					document.getElementById("crawlStatus").textContent =
						"Warning: PnP.js unavailable – fetch-only verification";
				},
			});

			log("Crawl finished", { results: results.length, stats });
			document.getElementById(
				"crawlStatus"
			).textContent = `Done → ${stats.pages} pages, ${stats.links} links, ${stats.errors} errors`;
		} catch (e) {
			if (e?.name === "AbortError") {
				warn("Crawl stopped by user");
				document.getElementById("crawlStatus").textContent = "Stopped by user";
			} else {
				error("Crawl failed", String(e));
				document.getElementById("crawlStatus").textContent = `Error: ${e.message || e}`;
			}
		} finally {
			log("Crawler UI reset after run");
			document.getElementById("startBtn").disabled = false;
			document.getElementById("stopBtn").disabled = true;
			document.getElementById("exportExcelBtn").disabled = false;
			document.getElementById("exportDocsExcelBtn").disabled = false;
			abortCtrl = null;
		}
	});
	/*  Stop Crawl  */
	document.getElementById("stopBtn").addEventListener("click", () => {
		abortCtrl?.abort();
	});

	document.getElementById("exportExcelBtn").addEventListener("click", () => {
		log("Export XLSX clicked", { rows: crawlResults.length });
		const includeSp = !!(lastRunConfig?.documentsOnly && lastRunConfig?.checkViaSharePoint);
		exportExcelJS(crawlResults, { includeSp });
		flash(document.getElementById("exportExcelBtn"), "Excel Exported");
	});

	document.getElementById("exportDocsExcelBtn").addEventListener("click", () => {
		const exts = arrayFromInput("documentExtensions").map(e => e.toLowerCase());
		const docs = crawlResults.filter(r => exts.includes((r.ext ?? "").toLowerCase()));
		log("Export Docs XLSX clicked", { rows: docs.length, exts });
		const includeSp = !!(lastRunConfig?.documentsOnly && lastRunConfig?.checkViaSharePoint);
		exportExcelJS(docs, { includeSp });
		flash(document.getElementById("exportDocsExcelBtn"), "✓ Excel Exported");
	});

	/*  Helper – read all inputs into a config object  */
	let lastRunConfig = null; // track how the last crawl was configured

	/*		function getUIConfig() {
				return {
					startUrl: document.getElementById("startUrl").value || location.href,
					maxDepth: +document.getElementById("maxDepth").value || 2,
					crawlDelay: +document.getElementById("crawlDelay").value || 100,
					verifyConcurrency: +(document.getElementById("verifyConcurrency")?.value || 8), // NEW
					urlMustStartWith:
						document.getElementById("urlMustStartWith").value || location.origin,				
					destMustStartWith:
					document.getElementById("destMustStartWith")?.value || "",
					externalOnly: document.getElementById("externalOnly")?.checked,
					urlMustNotContain: arrayFromInput("urlMustNotContain"),
					destUrlExclusion: arrayFromInput("destUrlExclusion"),
					headerFooterSelectors: arrayFromInput("headerFooterSelectors"),
					navSelectors: arrayFromInput("navSelectors"),
					excludeNavLinks: document.getElementById("excludeNavLinks").checked,
					documentExtensions: arrayFromInput("documentExtensions"),
					documentsOnly: document.getElementById("documentsOnly").checked,
					checkViaSharePoint:
						document.getElementById("checkViaSharePoint").checked,
				};
			}
	*/

	function getUIConfig() {
		const startUrlEl = document.getElementById("startUrl");
		const maxDepthEl = document.getElementById("maxDepth");
		const destEl = document.getElementById("destMustStartWith");
		const urlScopeEl = document.getElementById("urlMustStartWith");

		const startUrl = startUrlEl?.value || location.href;
		const scope = getStartUrlScopeInfo(startUrl);
		const recommendedUrlScope = getRecommendedUrlScopeFromStartUrl(startUrl);
		const recommendedDestScope = getRecommendedDestinationScopeFromStartUrl(startUrl);
		const autoUrlScope = isAutoFillEnabled("autoUrlMustStartWith", true);
		const autoDestScope = isAutoFillEnabled("autoDestMustStartWith", true);
		const rawUrlScopeValue = urlScopeEl?.value?.trim() || "";
		const rawDestValue = destEl?.value?.trim() || "";
		const effectiveUrlScope = autoUrlScope
			? (recommendedUrlScope || rawUrlScopeValue || location.origin)
			: (rawUrlScopeValue || recommendedUrlScope || location.origin);
		const effectiveDestScope = autoDestScope && scope.isInternal
			? (recommendedDestScope || rawDestValue)
			: rawDestValue;

		if (autoUrlScope && urlScopeEl && effectiveUrlScope && urlScopeEl.value !== effectiveUrlScope) {
			urlScopeEl.value = effectiveUrlScope;
		}

		if (autoDestScope && scope.isInternal && destEl && effectiveDestScope && destEl.value !== effectiveDestScope) {
			destEl.value = effectiveDestScope;
		}

		// Clamp external max depth
		if (scope.isExternal && maxDepthEl) {
			const currentDepth = +maxDepthEl.value || 2;
			if (currentDepth > 5) {
				maxDepthEl.value = 5;
			}
		}

		return {
			startUrl,
			maxDepth: +maxDepthEl?.value || 2,
			crawlDelay: +document.getElementById("crawlDelay").value || 100,
			verifyConcurrency: +(document.getElementById("verifyConcurrency")?.value || 8),

			urlMustStartWith:
				effectiveUrlScope,

			destMustStartWith:
				effectiveDestScope,

			autoUrlMustStartWith: autoUrlScope,
			autoDestMustStartWith: autoDestScope,

			externalOnly: document.getElementById("externalOnly")?.checked,
			urlMustNotContain: arrayFromInput("urlMustNotContain"),
			destUrlExclusion: arrayFromInput("destUrlExclusion"),
			headerFooterSelectors: arrayFromInput("headerFooterSelectors"),
			navSelectors: arrayFromInput("navSelectors"),
			excludeNavLinks: document.getElementById("excludeNavLinks").checked,
			documentExtensions: arrayFromInput("documentExtensions"),
			documentsOnly: document.getElementById("documentsOnly").checked,
			checkViaSharePoint:
				document.getElementById("checkViaSharePoint").checked,
		};
	}



	/***** -----------------------------------------------------------
	 * TEXT SCANNER MODULE (adds text-only crawl without touching link crawler)
	 * ----------------------------------------------------------- *****/

	/** Defaults specifically for text scanning */
	const TEXT_DEFAULTS = {
		// What to scan
		includeSelectors: "", // optional CSV: restrict scan to these selectors
		excludeSelectors: "", // optional CSV: skip these selectors (in addition to header/footer/nav)
		// Targets
		textTargets: [],      // literal phrases to find
		regexTargets: [],     // regex strings (without slashes), e.g. "(anti|counter)[ -]?money[ -]?laundering"
		caseSensitive: false,
		wholeWord: false,
		contextChars: 60,     // characters of context on each side of a hit
		maxMatchesPerPage: 200,
		// Crawl scope (reuses your page queuing rules)
		startUrl: "",
		urlMustStartWith: (typeof CRAWLER_DEFAULTS !== "undefined" ? CRAWLER_DEFAULTS.urlMustStartWith : location.origin),
		maxDepth: (typeof CRAWLER_DEFAULTS !== "undefined" ? CRAWLER_DEFAULTS.maxDepth : 2),
		crawlDelay: (typeof CRAWLER_DEFAULTS !== "undefined" ? CRAWLER_DEFAULTS.crawlDelay : 100),
		// Skip large boilerplate by default using your known header/footer/nav selectors
		headerFooterSelectors: (typeof CRAWLER_DEFAULTS !== "undefined" ? CRAWLER_DEFAULTS.headerFooterSelectors : []),
		navSelectors: (typeof CRAWLER_DEFAULTS !== "undefined" ? CRAWLER_DEFAULTS.navSelectors : []),
		// Normalization
		collapseWhitespace: true,
		removeSoftHyphens: true, // removes \u00AD so “anti­money” becomes “antimoney”
		stripZeroWidth: true,    // removes ZW* characters
	};

	/** Utility: parse comma/line separated input into array (robust for UI later) */
	function asList(input) {
		if (!input) return [];
		if (Array.isArray(input)) return input.filter(Boolean);
		return String(input)
			.split(/[,|\n]/)
			.map(s => s.trim())
			.filter(Boolean);
	}

	/** Utility: normalize text to improve matching across HTML boundaries */
	function normalizeForMatch(s, cfg) {
		if (!s) return "";
		let t = String(s);
		if (cfg.removeSoftHyphens) t = t.replace(/\u00AD/g, "");        // soft hyphen
		if (cfg.stripZeroWidth) t = t.replace(/[\u200B-\u200D\uFEFF]/g, ""); // zero width chars
		if (cfg.collapseWhitespace) t = t.replace(/\s+/g, " ");
		return t;
	}

	/** Compile literals & regex strings into ready-to-use RegExp objects */
	function compileTextPatterns(cfg) {
		const flags = cfg.caseSensitive ? "g" : "gi";
		const patterns = [];

		// Literal phrases
		for (const phrase of asList(cfg.textTargets)) {
			const escaped = phrase.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
			const source = cfg.wholeWord ? `\\b${escaped}\\b` : escaped;
			patterns.push({ type: "literal", sourcePhrase: phrase, re: new RegExp(source, flags) });
		}

		// Raw regex (as strings)
		for (const rx of asList(cfg.regexTargets)) {
			const source = cfg.wholeWord ? `\\b(?:${rx})\\b` : `(?:${rx})`;
			patterns.push({ type: "regex", sourcePhrase: `/${rx}/`, re: new RegExp(source, flags) });
		}

		return patterns;
	}

	/** Decide which blocks to scan (defaults to common content blocks) */
	function getScanBlocks(doc, cfg) {
		const defaultBlocks = "main, article, section, p, li, dd, dt, td, th, caption, blockquote, pre, h1, h2, h3, h4, h5, h6";
		let rootNodes = [...doc.querySelectorAll(defaultBlocks)];

		// Optional includeSelectors: if present, restrict to those
		const includes = asList(cfg.includeSelectors).join(", ");
		if (includes) {
			rootNodes = [...doc.querySelectorAll(includes)];
		}

		// Exclude boilerplate: header/footer/nav (from config) and custom excludes
		const excludes = [
			...(cfg.headerFooterSelectors || []),
			...(cfg.navSelectors || []),
			...asList(cfg.excludeSelectors),
		];

		// Filter nodes wrapped by any excluded selector
		const filtered = rootNodes.filter(el => !inClosest(el, excludes));

		return filtered;
	}

	/** Build a simple CSS path for a node (compact, stable enough for triage) */
	function cssPath(el, maxSegments = 6) {
		if (!el || el.nodeType !== 1) return "";
		const segments = [];
		let node = el;
		while (node && segments.length < maxSegments) {
			let sel = node.nodeName.toLowerCase();
			if (node.id) {
				sel += `#${CSS.escape(node.id)}`;
				segments.unshift(sel);
				break;
			} else {
				// choose a class if present and not too noisy
				const c = (node.className || "").toString().trim().split(/\s+/).filter(Boolean)[0];
				if (c) sel += `.${CSS.escape(c)}`;
				// :nth-child
				const parent = node.parentElement;
				if (parent) {
					const idx = [...parent.children].indexOf(node) + 1;
					sel += `:nth-child(${idx})`;
				}
				segments.unshift(sel);
			}
			node = node.parentElement;
		}
		return segments.join(" > ");
	}

	/** Find matches in a single block element; returns array of match rows */
	function findMatchesInBlock(el, patterns, cfg, pageTitle, pageUrl) {
		const rawText = el.textContent || "";
		const text = normalizeForMatch(rawText, cfg);
		if (!text) return [];

		const rows = [];
		for (const { re, sourcePhrase } of patterns) {
			re.lastIndex = 0;
			let m;
			let guard = 0;
			while ((m = re.exec(text)) && guard++ < cfg.maxMatchesPerPage) {
				const start = m.index;
				const end = start + Math.max(1, (m[0] || "").length);
				const left = Math.max(0, start - cfg.contextChars);
				const right = Math.min(text.length, end + cfg.contextChars);
				const snippet =
					(left > 0 ? "…" : "") +
					text.slice(left, start) +
					"[[" + text.slice(start, end) + "]]" +
					text.slice(end, right) +
					(right < text.length ? "…" : "");

				rows.push({
					pageTitle,
					pageUrl,
					elementTag: el.tagName.toLowerCase(),
					selector: cssPath(el),
					matchText: text.slice(start, end),
					pattern: sourcePhrase,
					index: start,
					snippet,
				});

				// Prevent infinite loops on zero-width matches
				if (re.lastIndex === start) re.lastIndex++;
				if (rows.length >= cfg.maxMatchesPerPage) break;
			}
			if (rows.length >= cfg.maxMatchesPerPage) break;
		}
		return rows;
	}

	/** Core: crawl site looking for text patterns (no link verification) */
	async function crawlSiteText(userCfg, hooks = {}) {
		const cfg = Object.assign({}, TEXT_DEFAULTS, userCfg);
		const patterns = compileTextPatterns(cfg);

		if (!cfg.startUrl) throw new Error("Start URL is required for text scan");
		if (!patterns.length) throw new Error("Provide at least one text or regex target");

		const visited = new Set();
		const queued = new Set();          // NEW: dedupe queued pages
		const verifyCache = new Map();   // NEW: cache of verified URLs
		const queue = [{ url: cfg.startUrl, depth: 0 }];
		const results = [];
		const stats = { pages: 0, matches: 0, errors: 0, queue: 1 };

		const {
			signal = undefined,
			onMatch = () => { },
			onPage = () => { },
			onStats = () => { },
		} = hooks;

		const aborted = () => signal?.aborted;

		const getExt = (urlObj) => {
			const m = urlObj.pathname.match(/\.(\w+)(?:[\?#]|$)/i);
			return m ? m[1].toLowerCase() : "";
		};

		while (queue.length && !aborted()) {
			const { url, depth } = queue.shift();
			if (visited.has(url) || depth > cfg.maxDepth) {
				onStats({ ...stats, queue: queue.length });
				continue;
			}
			visited.add(url);

			try {
				// Same modern-rendering path as the link crawler: for same-origin pages
				// this renders the page so SPFx content hydrates before we scan its text,
				// otherwise text scans of modern pages only ever see the empty shell.
				const page = await loadCrawlablePage(url, signal, { log, warn });
				if (!page.ok) { stats.errors++; onStats({ ...stats, queue: queue.length }); continue; }

				const doc = page.doc;
				const title = page.title;

				// Scan blocks for text hits
				const blocks = getScanBlocks(doc, cfg);
				let pageMatches = 0;
				for (const el of blocks) {
					const blockHits = findMatchesInBlock(el, patterns, cfg, title, url);
					for (const row of blockHits) {
						results.push(row);
						stats.matches++;
						pageMatches++;
						onMatch(row);
					}
				}

				stats.pages++;
				onPage({ url, title, matches: pageMatches });
				onStats({ ...stats, queue: queue.length });

				// Discover next pages to visit by following anchors (no verification)
				// Reuse your existing constraints to stay in-scope:
				const anchors = page.anchors; // union of shell + rendered DOM
				for (const a of anchors) {
					if (inClosest(a, cfg.headerFooterSelectors || [])) continue; // skip known boilerplate
					const rawHref = a.getAttribute("href")?.trim();
					if (!rawHref || rawHref.startsWith("#")) continue;

					let abs;
					try {
						abs = new URL(rawHref, url);
						abs = new URL(normalizeSavedUrl(abs.href));
					} catch { continue; }

					const ext = getExt(abs);
					// Treat SharePoint pages as HTML even though they end in .aspx, so the
					// text scanner actually follows modern-page links (matches link crawler).
					const isHtmlPage = !ext || ext === "aspx" || ext === "html" || ext === "htm";
					const isExternal = abs.origin !== location.origin;
					const shouldQueue =
						isHtmlPage &&
						!isExternal &&
						hrefStartsWithScope(abs.href, cfg.urlMustStartWith) &&
						!visited.has(abs.href) &&
						!queue.some(q => q.url === abs.href);

					if (shouldQueue) queue.push({ url: abs.href, depth: depth + 1 });
				}

				if (cfg.crawlDelay) await new Promise(r => setTimeout(r, cfg.crawlDelay));
			} catch (e) {
				stats.errors++;
				onStats({ ...stats, queue: queue.length });
			}
		}

		if (aborted()) throw new DOMException("Aborted", "AbortError");
		return { stats, results };

	}

	/** Optional: Excel export for text matches (mirrors your ExcelJS export) */
	async function exportTextMatchesExcelJS(rows) {
		if (!rows?.length) { alert("Nothing to export!"); return; }

		const headers = [
			"Page Title",
			"Page URL",
			"Element",
			"Selector",
			"Pattern",
			"Matched Text",
			"Index",
			"Snippet",
		];

		const workbook = new ExcelJS.Workbook();
		const sheet = workbook.addWorksheet("Text Matches");

		const tableData = rows.map(r => [
			r.pageTitle ?? "",
			r.pageUrl ?? "",
			r.elementTag ?? "",
			r.selector ?? "",
			r.pattern ?? "",
			r.matchText ?? "",
			r.index ?? "",
			r.snippet ?? "",
		]);

		sheet.addTable({
			name: "TextMatches",
			ref: "A1",
			headerRow: true,
			style: { theme: "TableStyleLight9", showRowStripes: true },
			columns: headers.map(h => ({ name: h })),
			rows: tableData,
		});

		sheet.columns = [
			{ width: 36 }, // Page Title
			{ width: 50 }, // URL
			{ width: 12 }, // Element
			{ width: 50 }, // Selector
			{ width: 40 }, // Pattern
			{ width: 30 }, // Matched Text
			{ width: 10 }, // Index
			{ width: 80 }, // Snippet
		];
		sheet.eachRow(row => row.eachCell(cell => { cell.font = { size: 10 }; }));

		const buffer = await workbook.xlsx.writeBuffer();
		const blob = new Blob([buffer], {
			type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
		});

		const ts = new Date().toISOString().replace(/[:T]/g, "-").slice(0, 19);
		const a = document.createElement("a");
		a.href = URL.createObjectURL(blob);
		a.download = `crawl-text-${ts}.xlsx`;
		a.click();
		URL.revokeObjectURL(a.href);
	}

	/* ------- END TEXT SCANNER MODULE ----------- */


	/*  On first load – fill a few helpful defaults  */

	/*  -----------------------------------------------------------
	   Fill the form with CRAWLER_DEFAULTS when the inputs exist
	   ----------------------------------------------------------- */
	let defaultsWaitLogged = false;

	function initDefaultsWhenReady() {
		/* wait until our markup is actually on the page */
		if (!document.getElementById("startUrl")) {
			if (!defaultsWaitLogged) {
				log("Waiting for crawler form controls before applying defaults");
				defaultsWaitLogged = true;
			}
			requestAnimationFrame(initDefaultsWhenReady);
			return;
		}

		log("Applying default form values");

		const d = CRAWLER_DEFAULTS;

		/* one-liner setter: handles text, textarea, number, checkbox */
		const set = (id, val) => {
			const el = document.getElementById(id);
			if (!el) return;
			/* arrays → comma-string, empty array → '' */
			const v = Array.isArray(val) ? val.join(", ") : (val ?? "");
			if (el.type === "checkbox") {
				el.checked = !!v; // true / false
			} else {
				el.value = v; // text input / textarea
			}
		};

		/* ---- write each field (blanks allowed) ---- */
		set("startUrl", d.startUrl);
		set("urlMustStartWith", d.urlMustStartWith);
		set("maxDepth", d.maxDepth);
		set("crawlDelay", d.crawlDelay);
		set("urlMustNotContain", d.urlMustNotContain);
		set("destUrlExclusion", d.destUrlExclusion);
		set("headerFooterSelectors", d.headerFooterSelectors);
		set("navSelectors", d.navSelectors);
		set("documentExtensions", d.documentExtensions);
		set("enablePnP", d.enablePnP); // checkbox
		set("destMustStartWith", "");
		set("autoUrlMustStartWith", true);
		set("autoDestMustStartWith", true);

		log("Default form values applied");

		wireCrawlerValidationEvents();
	}

	function wireCrawlerValidationEvents() {
		const startUrlEl = document.getElementById("startUrl");
		const destEl = document.getElementById("destMustStartWith");
		const maxDepthEl = document.getElementById("maxDepth");
		const urlScopeEl = document.getElementById("urlMustStartWith");
		const autoUrlScopeEl = document.getElementById("autoUrlMustStartWith");
		const autoDestScopeEl = document.getElementById("autoDestMustStartWith");

		[startUrlEl, destEl, maxDepthEl, urlScopeEl, autoUrlScopeEl, autoDestScopeEl].forEach((el) => {
			if (!el) return;
			el.addEventListener("input", validateCrawlerUrlConstraints);
			el.addEventListener("blur", validateCrawlerUrlConstraints);
		});

		// Run once on load after defaults are present
		validateCrawlerUrlConstraints();
	}

	/* run once the whole SharePoint page finishes loading */

	if (document.readyState === "complete") {
		// load already happened
		log("Ready state complete detected");
		requestAnimationFrame(initDefaultsWhenReady);
	} else {
		window.addEventListener("load", () => {
			log("Window load detected");
			requestAnimationFrame(initDefaultsWhenReady);
		});
	}



	/* run once the whole SharePoint page finishes loading */
	let defaultsApplied = false;


	/* -------- CONSOLE HELPER FOR TEXT SCAN -------- */
	/* ----------------- Console-accessible API (optional) -----------------
	   Exposes minimal helpers to the global window for ad-hoc console use.
	   Keeps everything else private inside the IIFE.
	------------------------------------------------------------------------ */

})(); /*  end of IIFE wrapper  */

