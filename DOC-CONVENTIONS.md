# GitHub Copilot/Claude Instructions — JS Commenting Standards (FCU / SharePoint Scripts)

## Purpose
Ensure all JavaScript generated or modified by GitHub Copilot follows a **consistent, enterprise-grade documentation standard** aligned to our SharePoint custom script patterns.

---

## Core Principles

- **Always document the file**
- **Always use JSDoc for public + meaningful private functions**
- **Prefer clarity over verbosity**
- **Do NOT generate fluff comments**
- **Match the tone and structure of existing scripts exactly**
- **Keep comments useful for future maintainers (including Copilot itself)**

---

## 1. File Header (MANDATORY)

```js

/** Always add an initial log to signal the script start. This must be included before the comment header in order for an auto-timestamp extension to automatically update the modification date. */
const LOG_PREFIX = "[ScriptName]";  // A short keyword identifiying the script

console.log(`${LOG_PREFIX} Loaded`);
console.log(`${LOG_PREFIX} Last modified: 2026/06/27 11:59:15
`);



/**
 * @fileoverview <Short description of what this script does>
 *
 * <1–2 lines describing the script at a high level>
 *
 * Features:
 * - <Feature 1>
 * - <Feature 2>
 * - <Feature 3>
 *
 * @author <NAME; TEAM; EMAIL>
 * @created Created: YYYY/MM/DD HH:mm:ss
 * @lastmodified Last modified: 2026/06/27 11:59:15
 *
 * @example
 * // <Example usage 1>
 *
 * @example
 * // <Example usage 2>
 *
 * @notes (optional)
 * - <Important implementation detail>
 * - <Constraints or caveats>
 */
```

## Required Fields

- @fileoverview → Always present
- Features: → Always a short bullet list
- @author → Full format:

```text
Name; Team; Email
```

- @created → Preserve original timestamp if file already exists
- @lastmodified → Only add if missing. A VSCode extension will handle the update. 
- At least 1 example block

## Formatting Rules

- Use YYYY/MM/DD HH:mm:ss
- Use - for bullets (not *)
- Keep spacing exactly aligned
- No trailing punctuation unless needed

## 3. Function-Level JSDoc (REQUIRED for all meaningful functions)

Every function must include JSDoc if it is:

- Reusable
- Non-trivial
- Performs logic beyond simple passthrough

### Template:

```js
/**
 * <Short description of what the function does>
 *
 * <Optional second line describing behavior or constraints>
 *
 * @param {Type} paramName - Description
 * @param {Type} [optionalParam] - Description (optional)
 * @returns {Type} Description of return value
 */
```

### Example

```js
/**
 * Safely deep clones an object using JSON serialization
 *
 * Falls back to returning the original object if cloning fails
 *
 * @param {Object} obj - Object to clone
 * @returns {Object} Cloned object or original fallback
 */
const safeClone = (obj) => {
  try {
    return JSON.parse(JSON.stringify(obj));
  } catch {
    return obj;
  }
};
```

## 4. Section Comments (STRUCTURAL)

Use banner-style comments to define major sections:

```js
/* -----------------------------------------------------------
   Initialization
----------------------------------------------------------- */

/* ------- TEXT SCANNER MODULE ----------- */
```


## Rules

### Use for:

- Modules
- Phases (init, execution, export)
- Logical groupings


- Keep consistent dash style
- Do NOT overuse


## 5. Inline Comments (MINIMAL + PRACTICAL)

Only add inline comments when:

- Logic is not obvious
- There is a workaround or hack
- Timing / async behavior matters

### Example:

```js
// Wait until SharePoint renders controls before applying defaults
```

### Do NOT:

- Comment obvious code
- Repeat variable names in English
- Add filler like “This function does X” (JSDoc already handles this)

## 7. Copilot Behavior Rules (STRICT)

### Copilot MUST NOT:

- ❌ Introduce new comment styles
- ❌ Use block comments instead of JSDoc for functions
- ❌ Remove existing metadata
- ❌ Reformat the header
- ❌ Add unnecessary comments

## 6. Logging. Script must ALWAYS use a logging pattern where an initializtion log appears FIRST in the script above all header comments.

```js
const LOG_PREFIX = "[ScriptName]";
console.log(LOG_PREFIX + " Loaded | Last modified: YYYY/MM/DD HH:mm:ss");
```


Script must ALWAYS use logging pattern:

```js
console.log(LOG_PREFIX + " Message");
console.warn(LOG_PREFIX + " Warning message");
console.error(LOG_PREFIX + " Error message");
```

## Rules

### ✅ Use native console methods directly:

- console.log
- console.warn
- console.error

### ❌ Do NOT:

- Create logging helper functions
- Wrap console calls in utilities
- Abstract logging

## Grouping (USE WHEN APPROPRIATE)

Use console groups for structured outputs:

```js
console.group(LOG_PREFIX + " Crawl Results");
console.log(LOG_PREFIX + " Found X links");
console.log(LOG_PREFIX + " Errors: Y");
console.groupEnd();
```

- Batch operations (crawl, scan, export)
- Debug summaries
- Multi-line related outputs

Do NOT overuse for simple logs