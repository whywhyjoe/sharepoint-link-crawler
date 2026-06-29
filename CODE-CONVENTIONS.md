
# GitHub Copilot/Claude Instructions — JS Coding Standards (FCU / SharePoint Scripts)

The following instructions apply to script coding, but explicitly do NOT apply to HTML or CSS.

==================================================
STYLE & STRUCTURE
==================================================

- Write the code so it reads top-to-bottom like a story:
  context detection → environment readiness → main flow → edge cases → observers/hooks → cleanup

- Avoid “function soup.” Only extract helper functions if:
  (1) reused 2+ times or must be called from observers, loops, retries, or async flows
  (2) genuinely complex (10+ lines, branching, or hard-to-read logic)

- Inline one-off transforms (string replacements, simple mapping, short queries).

- Prefer a flat, procedural style. Avoid “enterprise architecture” patterns.

- Avoid unnecessary layers:
  no initEverything() → initFeatureA() → runFeature()
  unless the separation truly improves clarity.

==================================================
FUNCTION RULES
==================================================

- Prefer standard hoisted function declarations:
  `function name() {}` (not arrow functions for core logic)

- If helpers exist:
  - keep them few
  - keep them obvious
  - place them at the bottom under a clear “Helpers” heading

- Use guard clauses to reduce nesting.

- Use short section comments to label phases.

- Avoid abstraction chains (“init → setup → run → render”).
  Keep execution direct and readable.

==================================================
IIFE / GLOBALS
==================================================

- Default: DO NOT wrap the entire file in an IIFE.
- SharePoint custom script blocks are already isolated enough.
- If global safety *is* required:
  - use ONE minimal wrapper
  - keep internals flat (no nested modules or namespaces)

==================================================
SHAREPOINT ENVIRONMENT & LOADING RULES
==================================================

SharePoint pages load asynchronously and may rehydrate without full reloads.
Your script **must be resilient, but not over-engineered**.

### Library Readiness

- Do NOT assume availability timing of:
  - jQuery
  - Alpine
  - PnPjs
  - page DOM structure

- Prefer to *wait explicitly* for what you need.

- If the standard script library provides environment helpers, strongly consider using them custom MutationObservers. 

### Preferred Waiting Strategy (in order):

1. Use shared helpers from the standard script library when available and appropriate to task:
   - `waitForElement(selector, options)`
   - shared DOM-ready or hydration helpers
   - shared navigation or re-init utilities

2. Use lightweight polling or single-purpose observers **only when needed**.

3. Use MutationObserver **only when truly needed**, and scope it narrowly.
   Do not watch the entire document by default. Explain why you are making this choice when you output code.
   
4. For any custom helper, consider its reuse potential. If broadly applicable, offer to generalize it for adding to the standard script library.


### PnP / Alpine Readiness

- If using PnPjs:
  - Treat it as async-ready, not instantly available.
  - Wait until the global (`pnp2`) is defined before executing calls.

- If using Alpine:
  - Do not assume Alpine has started.
  - Prefer waiting for Alpine presence or a stable DOM node Alpine will hydrate.
  - Do NOT force Alpine.start() unless explicitly required.

==================================================
SPA NAVIGATION (IMPORTANT)
==================================================

- SharePoint uses SPA-style navigation.
- Not every script needs SPA compensation. Ask the user before starting the project. 

### Implement navigation handling if:
- Your script depends on page-specific DOM
- Your logic breaks when navigating without full refresh
- You need to re-run logic after hydration or navigation

### When navigation handling *is* needed:
- Prefer shared navigation / re-init helpers from the standard script library
- Avoid duplicating navigation logic per script
- Keep tracking minimal and idempotent
- Track whether your script has already initialized on the current page

### When navigation handling is *not* needed:
- Do not add observers “just in case”
- Keep the script simple

==================================================
READABILITY REQUIREMENTS
==================================================

- Add short (1-line) section comments to mark phases
- Prefer explicit names over abstractions
- Prefer clarity over cleverness
- Avoid defensive boilerplate (excess try/catch, framework-like utilities)
  unless required by real SharePoint behavior

==================================================
STANDARD LIBRARIES & INCLUDES
==================================================

These are known, supported locations and should be referenced instead of custom bundles when possible.

Available libraries:
https://bmo.sharepoint.com/:u:/r/sites/FCUPortal/Code/lib/jquery-3.7.1.min.js
https://bmo.sharepoint.com/:u:/r/sites/FCUPortal/Code/lib/alpine_cdn.min.js
https://bmo.sharepoint.com/:u:/r/sites/FCUPortal/Code/lib/pnp2.bundle.js


Optional includes or references (preferred for helpers. They can be referenced OR the relevant helper can be copied into the script. This is to help the coder keep track of helper methodology.):
https://bmo.sharepoint.com/:u:/r/sites/FCUPortal/Code/standard-include_2.1.html
https://bmo.sharepoint.com/:u:/r/sites/FCUPortal/Code/standard-script_2.1.js
----------
For review of includes by agents, local locations are:
'C:\Users\jzapert\BMO Financial Group\FCU Portal - Code\standard-include_2.1.html'
'C:\Users\jzapert\BMO Financial Group\FCU Portal - Code\standard-script_2.1.js'

** NEVER MODIFY THE STANDARD INCLUDE FILES.**
* IF A MODIFICATION IS CONSIDERED NECESSARY, COPY THE HELPER INTO THE CURRECT WORKING FILE AND COMMENT THE MODIFICATIONS, and NOTIFY THE USER WHEN THE FULL CURRENT TASK IS COMPLETE


==================================================
OUTPUT EXPECTATIONS
==================================================

- Write script-first code, not framework scaffolding.
- If you believe a more structured approach is “better”:
  still follow this style,
  and mention alternatives only briefly *after* the code.

