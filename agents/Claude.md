# Coding Style Guide

## 1. Role & Mindset

* **Full-Stack Development Partner**
  Own the process from requirements → proposal → implementation → refinement.
* **Requirements First**
  Every feature or fix must trace back to `README.md`.
* **UX-Driven**
  Continually ask “How can teachers and students use this most intuitively?”

---

## 2. Common Coding Conventions

### 2.1 Indentation & Whitespace

* Use **2 spaces** per indent level.
* Remove trailing whitespace on every line.
* End every file with exactly one newline.

### 2.2 Naming

* **Variables & functions**: `camelCase`
* **Constants**: `UPPER_SNAKE_CASE`
* **Classes / constructors**: `PascalCase`
* **File names**: short English words; extensions: `.gs`, `.html`, `.css.html`, `.js.html`.

### 2.3 Syntax & Style

* Always use semicolons.
* Prefer single quotes (`'string'`).
* Allow trailing commas in objects/arrays for cleaner diffs.
* Use guard clauses to keep nesting shallow:

  ```js
  if (!user) return;
  // safe to use user here
  ```

### 2.4 Comments & Documentation

* Document every function with **JSDoc**:

  ```js
  /**
   * Fetches user data by ID.
   * @param {string} userId – The ID of the user.
   * @returns {Object} User data object.
   */
  function fetchUserData(userId) { … }
  ```
* Add detailed inline comments for complex logic or external-spec integrations.
* Mark todos/fixes with ticket numbers:

  ```js
  // TODO #123: update API endpoint
  ```

---

## 3. Client / Server Separation

### 3.1 Server-Side (Apps Script)

* Implement **only** business logic—no DOM/UI code.
* Entry points: `doGet(e)` / `doPost(e)` only. Delegate other logic to modules.
* All data retrieval and updates occur in server functions.

### 3.2 Client-Side (HTML/CSS/JS)

* Build UI entirely in HTML templates + external JS/CSS.
* **Load data asynchronously**; initial HTML shows placeholders:

  ```html
  <body>
    <div id="content">Loading…</div>
    <?!= include('scripts/main.js.html'); ?>
  </body>
  ```

  ```js
  // main.js.html
  <script>
    window.addEventListener('load', () => {
      google.script.run
        .withSuccessHandler(renderContent)
        .fetchInitialData();
    });
    function renderContent(data) {
      document.getElementById('content').textContent = data.text;
    }
  </script>
  ```

---

## 4. Drive API / Sheets API Guidelines

### 4.1 Service Wrapper Modules

* Never call `DriveApp` or `SpreadsheetApp` directly in business logic.
* Wrap API calls in a module for easy mocking and future changes:

  ```js
  // driveService.gs
  var DriveService = (function() {
    /**
     * Gets a file by ID.
     * @param {string} fileId
     * @returns {GoogleAppsScript.Drive.File}
     */
    function getFileById(fileId) {
      return DriveApp.getFileById(fileId);
    }
    return { getFileById };
  })();
  ```

### 4.2 Field Filtering

* When using Advanced Drive Service, always specify `fields`:

  ```js
  Drive.Files.get(fileId, { fields: 'id,name,mimeType' });
  ```

### 4.3 Batch Operations

* For multiple sheet updates, use `batchUpdate`:

  ```js
  var requests = [
    { updateCells: { /* … */ } },
    { appendDimension: { /* … */ } }
  ];
  Sheets.Spreadsheets.batchUpdate({ requests }, SPREADSHEET_ID);
  ```

### 4.4 Idempotency & Retry

* Implement exponential back-off for network errors or quota limits:

  ```js
  function retry(fn, attempts) {
    try {
      return fn();
    } catch (e) {
      if (attempts > 0) {
        Utilities.sleep(Math.pow(2, attempts) * 1000);
        return retry(fn, attempts - 1);
      }
      throw e;
    }
  }
  var values = retry(() => Sheets.Spreadsheets.Values.get(id, range), 5);
  ```

### 4.5 Caching

* **CacheService**: cache infrequently changing reference data for minutes.
* **localStorage** (client-side): cache UI data to reduce server calls.

### 4.6 Logging & Monitoring

* Log key flows and errors via `console.log()` (Stackdriver).
* Use a consistent prefix:

  ```js
  console.log('[DriveService] getFileById:', fileId);
  ```

### 4.7 Testing with Mocks

* In CI unit tests, mock service wrappers.
* **Jest example**:

  ```js
  jest.mock('../server/services/driveService', () => ({
    getFileById: jest.fn().mockReturnValue({ id: '123', name: 'Test' })
  }));
  ```

---

## 5. HTML Service Best Practices

1. **Separate HTML / CSS / JS**

   * `.css.html` files contain `<style>…</style>`.
   * `.js.html` files contain `<script>…</script>`.
   * Main templates include them via `<?!= include('…'); ?>`.

2. **HTML5 Doctype**

   ```html
   <!DOCTYPE html>
   <html lang="en"> … </html>
   ```

3. **HTTPS Only**

   * All external scripts/styles must use `https://`.

4. **Load scripts at end of `<body>`**

   * Prioritize HTML/CSS rendering.

5. **Asynchronous Data Fetching**

   * Initial UI shows placeholders; data rendered via `google.script.run`.

6. **Optional jQuery Use**

   * If needed, load via CDN:

     ```html
     <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.6.0/jquery.min.js"></script>
     ```

---

## 6. Performance Optimization

* Minimize API calls by batching reads/writes.
* Use `LockService` to prevent concurrent trigger conflicts.
* CacheService / localStorage to avoid redundant fetches.

---

## 7. Security Measures

* Store secrets in `PropertiesService`—never hard-code API keys.
* Limit OAuth scopes to the minimum required.
* Enforce Content Security Policy and sanitize all user input.

---

## 8. Testing & Debugging

* Write unit tests (QUnit / Jest) covering server-side logic.
* Mock external calls in tests.
* Use `Logger.log()` or `console.log()` appropriately.
* Integrate tests into CI (clasp + GitHub Actions) to run `npm test`.

---

## 9. Git Workflow & Pull Requests

* **Conventional Commits**:

  ```
  feat(server): add fetchData function
  fix(ui): correct modal close behavior
  test(services): add retry logic tests
  ```
* **Branch Naming**: `feature/…`, `bugfix/…`, `chore/…`.
* Keep PRs small, focused, and include clear “What”, “Why”, and “How”.

---

## 10. Trigger Management

* Register only necessary triggers.
* Remove obsolete triggers immediately.
* Specify time zones explicitly for time-based triggers.

---

## 11. Documentation

* **README.md**: Keep project overview and setup instructions up to date.
* **CHANGELOG.md**: Record major feature additions and breaking changes.
* In-code comments should explain intent and any gotchas.

---

Adhering to these guidelines ensures a consistent, maintainable, high-performance, secure, and user-friendly codebase for “Everyone’s Answer Board.” Always refer back here if in doubt.
