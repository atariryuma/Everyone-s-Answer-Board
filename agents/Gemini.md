# 🤖 Agent’s Guidebook for “Everyone’s Answer Board”

This guide ensures all AI agents and human contributors build secure, maintainable, and user-friendly educational tools, following both our project conventions and Google’s HTML Service best practices.

---

## 1. Agent’s Role & Objective

* **Role**: Full-stack development partner, not just a code generator.
* **Objective**:

  * Implement features strictly per `README.md`.
  * Adhere to the prescribed file structure and separation of concerns.
  * Propose and implement UI/UX improvements for teachers and students.

---

## 2. Core Principles

1. **Requirements First**: Every change must trace back to `README.md`.
2. **Structure Is Paramount**: Follow the `/src` directory layout exactly.
3. **Security First**: Never expose credentials or sensitive data.
4. **Test Everything**: Include unit tests for all new server logic.
5. **User-Centered Design**: Prioritize ease of use for teachers and students.

---

## 3. Development Workflow

1. **Clarify Requirements**: Review the relevant `README.md` section.
2. **Target Files**: Identify which files under `/src` to update.
3. **Implement**: Write code following the architecture and coding rules.
4. **Test**: Add or update tests under `/tests`; confirm `npm run test` passes.
5. **Self-Review**: Verify adherence to all guidelines.
6. **Commit**: Use Conventional Commits (e.g., `feat(server): add reactionService`).
7. **Pull Request**: Keep PRs focused; include a clear description.

---

## 4. Project Architecture & File Structure

All code lives under `/src`. Slashes in filenames become folders in the GAS editor.

```text
/src
├── server/
│   ├── main.gs         # doGet, routing, render/include helpers only
│   ├── database.gs     # Sheets/database operations
│   └── services/       # Business logic (e.g. reactionService.gs)
│
├── client/
│   ├── views/          # HTML templates (AdminPanel.html, Page.html, etc.)
│   ├── styles/         # CSS snippets wrapped in .css.html
│   ├── scripts/        # JS snippets wrapped in .js.html
│   └── components/     # Reusable HTML pieces (Header.html, Modals, etc.)
│
└── appsscript.json     # GAS manifest (do not modify)
```

### 4.1 Server-Side Logic (`/src/server/`)

* **`main.gs`**:

  * Contains `doGet(e)` to route requests.
  * Defines `renderPage(viewName, data)` and `include(path)` only.
  * **No business logic here**—delegate to `database.gs` or `services/`.

* **Other `.gs` files**:

  * Each feature’s logic lives in its own file under `server/` or `server/services/`.

### 4.2 Front-End Views (`/src/client/views/`)

* Pure HTML templates.
* Use scriptlets to include CSS, JS, and components:

  ```html
  <!DOCTYPE html>
  <html lang="en">
    <head>
      <base target="_top">
      <?!= include('styles/main.css.html'); ?>
      <?!= include('styles/Page.css.html'); ?>
    </head>
    <body>
      <?!= include('components/Header.html'); ?>

      <main id="app"></main>

      <?!= include('scripts/main.js.html'); ?>
      <?!= include('scripts/Page.js.html'); ?>
    </body>
  </html>
  ```

---

## 5. Google’s HTML Service Best Practices

These guidelines come directly from Google’s official documentation and ensure your HTML UI loads quickly, stays secure, and remains maintainable.

1. **Separate HTML, CSS & JavaScript**

   * **Why**: Improves readability and reuse.
   * **How**:

     * In `Code.gs`, define:

       ```js
       function include(filename) {
         return HtmlService
           .createHtmlOutputFromFile(filename)
           .getContent();
       }
       ```
     * In your main template:

       ```html
       <?!= include('styles/Page.css.html'); ?>
       <?!= include('scripts/Page.js.html'); ?>
       ```

2. **Load Data Asynchronously**

   * **Why**: Prevents server-side template delays from blocking the initial render.
   * **How**:

     * Show a “Loading…” placeholder in HTML.
     * Call `google.script.run.withSuccessHandler(renderData).fetchData();` in client JS.

3. **Always Use HTTPS for External Resources**

   * **Why**: Mixed-content (HTTP) blocks in IFRAME sandbox mode.
   * **How**:

     * Ensure every `<script>` or `<link>` URL starts with `https://`.

4. **Include the HTML5 Doctype**

   * **Why**: Forces standards-mode rendering; avoids quirks.
   * **How**:

     ```html
     <!DOCTYPE html>
     <html>…</html>
     ```

5. **Place JavaScript at the End of `<body>`**

   * **Why**: Prioritizes HTML/CSS rendering; speeds up perceived load time.
   * **How**:

     * Move `<script>` includes to just before `</body>`.

6. **Leverage jQuery if Needed**

   * **Why**: Simplifies DOM manipulation and event handling.
   * **How**:

     ```html
     <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.6.0/jquery.min.js"></script>
     <script>
       $(function() {
         $('#btn').click(() => alert('Clicked'));
       });
     </script>
     ```

---

## 6. Stylesheet & Script File Rules

* **`.css.html`**: Entire file wrapped in `<style>…</style>`.
* **`.js.html`**: Entire file wrapped in `<script>…</script>`.
* **No ES Module Syntax**: Use IIFEs or global scope.

---

## 7. Testing Policy

* Write or update unit tests under `/tests`.
* Confirm all tests pass before merging.
* Use clear, focused test cases for each new function.

---

## 8. Commit & PR Guidelines

* **Commit Messages**: Follow Conventional Commits:

  ```
  <type>(<scope>): <short description>
  ```

  * types: `feat`, `fix`, `docs`, `style`, `refactor`, `test`, `chore`.
* **Pull Requests**:

  * Keep them small and topic-focused.
  * Provide a descriptive title and summary.

---

## 9. Prohibited Actions

1. **No sensitive data** in any committed files.
2. **Never push directly** to `main`.
3. **Do not break** the file structure—always follow `/src` conventions.
4. **Avoid giant PRs**; split large changes into digestible pieces.

