# GEMINI.md  
_last updated: 2025‑07‑22 JST_

> **Purpose**  
> 1. Provide Gemini CLI / Code Assist with an authoritative, machine‑parsable guideline.  
> 2. Enforce a deterministic 4‑stage development loop (PLAN → IMPLEMENT → TEST → REVIEW).  
> 3. Encode all coding conventions, architecture rules, security and CI policies in one place.  

---

## 0. Global Directives
* Language: **Japanese** (コード内コメント含む。).  
* No chain‑of‑thought exposure.  
* Always obey the latest explicit user instruction.  
* If input is ambiguous, ask. If stuck, say so.  

---

## 1. Role & Mindset
1. **Partner**: Think, design, implement, improve.  
2. **Requirement‑Driven**: Map every change to `README.md` purpose.  
3. **UX‑First**: Prioritise teacher / student clarity.  

---

## 2. Interaction Protocol
### 2.1 Four‑Stage Loop  
Each message **must contain exactly one** of the following blocks.

```md
### PLAN
- 要件: …
- 修正対象: …
- タスク:
  1. [ ]

### IMPLEMENT
```gs
// 差分のみ
````

### TEST

* 結果: PASS / FAIL
* 手順: …

### REVIEW

* 品質: ✅ / ❌
* 指摘: …

````

*Order is strict; after REVIEW, restart at PLAN.*  
*On TEST=FAIL, analyze cause then PLAN again.*

### 2.2 Formatting
| Field | Rule |
|------|------|
| Stage headers | `### {PLAN|IMPLEMENT|TEST|REVIEW}` |
| Lists / code | Markdown fenced blocks or `-` / `1.` lists |
| Questions | End with `?` |

---

## 3. Coding Conventions
* **Indent**: 2 spaces.  
* **EOL**: LF + single trailing newline.  
* **Trailing whitespace**: disallowed.  
* **Semicolons**: mandatory.  
* **Strings**: `'single'`.  
* **Trailing commas**: allowed.  
* **Guard clauses** over deep nesting.  

| Entity | Style |
|--------|-------|
| var / func | `camelCase` |
| const | `UPPER_SNAKE_CASE` |
| class | `PascalCase` |
| files | concise‑english + `.gs/.html/.css.html/.js.html` |

*JSDoc every function; inline comment “why”, not “what”.*  
*TODO: `// TODO(#ticket): …`.*

---

## 4. Architecture
### 4.1 Layering
* **Server (`.gs`)**: business only, entry `doGet|doPost`.  
* **Client (`.html/.js.html`)**: UI, async via `google.script.run`.  

### 4.2 Service Wrappers
No direct `DriveApp` / `SpreadsheetApp`. Use wrapper modules.
```gs
// example
const DriveService = {
  getFileById(id){ return DriveApp.getFileById(id); }
};
````

---

## 5. API Best Practices

* Batch (`getAllValues`, `setValues`).
* Cache: `CacheService` / client `localStorage`.
* Advanced APIs: specify `fields`.
* Retry with exponential backoff; use `LockService` for contended ops.

---

## 6. Security

* Secrets → `PropertiesService`.
* OAuth scopes = minimum.
* `HtmlService.setXssProtection(HtmlService.XssProtectionMode.V8)`.
* Validate & sanitize all user input.

---

## 7. Testing & CI

* Unit: Jest or QUnit; mock wrappers.
* GitHub Actions triggers `npm test`.
* Logging: `Logger.log` (dev) / `console.log` (prod) prefixed `[Module]`.

---

## 8. Git & Deploy

* Conventional Commits.
* Branch names: `feature/*`, `bugfix/*`.
* Small PRs with what/why/how.

---

## 9. Documentation

Keep `README.md` & `CHANGELOG.md` current.

---

## 10. TaskMaster AI Integration

* Convert chat actions to tasks:
  `@task-master-ai:add_task "説明" due:YYYY-MM-DD p{1|2}`
* Sync Task IDs inside PLAN list.
* Mark done after TEST passes.

---

```
::contentReference[oaicite:0]{index=0}
```
