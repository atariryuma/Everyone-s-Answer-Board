# 🤖 Agent's Guidebook for "Everyone's Answer Board"

This document provides the essential guidelines for all AI agents (and human developers) contributing to this project. Its purpose is to maintain code quality and consistency and to promote efficient, scalable development.

## 1\. Agent's Role and Objective

Your role is not just a code generator, but a **full-stack development partner** responsible for building a robust, maintainable, and user-centric educational tool.

**Objective:**

  * To implement features according to the `README.md`, following the established architecture.
  * To strictly adhere to the file structure and separation of concerns to ensure project scalability.
  * To propose and implement UI/UX enhancements that provide a seamless and high-quality experience for both teachers and students.

## 2\. Core Principles

1.  **Requirements are King**: All implementation must be based on `README.md`.
2.  **Structure is Paramount**: Adherence to the file and directory structure is mandatory for maintaining project sanity.
3.  **Security First**: Exercise extreme caution with sensitive data.
4.  **Test Everything**: All new logic must be accompanied by tests.
5.  **User-Centered Design**: Always prioritize the user experience for teachers and students.

## 3\. Development Workflow

1.  **Understand Requirements**: Analyze the task and relevant sections of `README.md`.
2.  **Identify Target Files**: Based on the architecture, determine which files in the `/src` directory need to be created or modified.
3.  **Implement (Coding)**: Write code following the strict file structure and coding guidelines outlined below.
4.  **Test**: Add or update tests in the `/tests` directory.
5.  **Self-Review**: Ensure your changes conform to all guidelines.
6.  **Commit**: Use the Conventional Commits format.
7.  **Pull Request (PR)**: Create a focused PR with a clear description.

## 4\. Code Architecture & File Structure

This project adopts a **full-stack, separation of concerns** architecture optimized for local development with `clasp` and maintainability in the GAS online editor. You **must** follow this structure precisely.

### **【System-Wide Rule】File Naming and Directory Structure**

All files must be placed within the `/src` directory following the structure below. The GAS online editor will mimic this hierarchy by treating slashes (`/`) in filenames as folders.

```plaintext
/src
├── 📁 server/        # SERVER-SIDE LOGIC (.gs)
│   ├── main.gs       # Entry points (doGet) and top-level routing.
│   ├── database.gs   # Database (Google Sheets) operations.
│   └── services/     # Directory for business logic per feature.
│       └── reactionService.gs
│
├── 📁 client/        # FRONT-END SOURCES
│   ├── 📁 views/        # Main HTML templates for each page.
│   │   ├── AdminPanel.html
│   │   ├── Page.html
│   │   └── Registration.html
│   │
│   ├── 📁 styles/       # CSS files (wrapped in .html).
│   │   ├── main.css.html
│   │   └── Page.css.html
│   │
│   ├── 📁 scripts/      # Client-side JavaScript (wrapped in .html).
│   │   ├── main.js.html
│   │   └── Page.js.html
│   │
│   └── 📁 components/   # Reusable UI component snippets.
│       ├── Header.html
│       └── ConfirmationModal.html
│
└── 📄 appsscript.json # Project manifest (do not modify).
```

### **【Implementation Guide】How to Write Code**

#### **1. Server-Side Logic (in `/src/server/`)**

  * **`main.gs`**: This is the primary entry point. It contains the `doGet` function for routing and the `renderPage` and `include` helper functions for building the front-end. **Do not add business logic here.**
    ```javascript
    // /src/server/main.gs
    function doGet(e) {
      if (isAdmin(e)) return renderPage('AdminPanel');
      if (!isRegistered(e)) return renderPage('Registration');
      return renderPage('Page');
    }

    function renderPage(viewName, data = {}) {
      const template = HtmlService.createTemplateFromFile(`client/views/${viewName}`);
      template.data = data;
      template.include = include; // Make helper available to templates
      return template.evaluate().setTitle('みんなの回答ボード').addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }

    function include(path) {
      return HtmlService.createHtmlOutputFromFile('client/' + path).getContent();
    }
    ```
  * **Other `.gs` files**: All other business logic (e.g., database access, reaction processing) must be in separate files like `database.gs` or within the `services/` directory.

#### **2. Front-End Views (in `/src/client/views/`)**

  * These files are **HTML skeletons only**.
  * They must use the `include()` helper function via scriptlets `<?! ... ?>` to load all CSS, JavaScript, and components.
  * **Example (`/src/client/views/Page.html`):**
    ```html
    <!DOCTYPE html>
    <html lang="ja">
      <head>
        <base target="_top">
        <?! include('styles/main.css.html'); ?>
        <?! include('styles/Page.css.html'); ?>
      </head>
      <body>
        <?! include('components/Header.html'); ?>
        <main id="answers"></main>
        <?! include('scripts/main.js.html'); ?>
        <?! include('scripts/Page.js.html'); ?>
      </body>
    </html>
    ```

#### **3. Stylesheets and Scripts (in `/src/client/styles/` & `/src/client/scripts/`)**

  * **`.css.html` files**: The entire content **must** be wrapped in `<style>` tags.
  * **`.js.html` files**: The entire content **must** be wrapped in `<script>` tags.
  * **No `import`/`export`**: Since these are not ES modules, you cannot use import/export syntax. Use the global scope or IIFE `(function(){ ... })();` to avoid conflicts.

## 5\. Testing Policy

  * Add unit tests to the `/tests` directory for all new server-side business logic.
  * Ensure all existing tests pass (`npm run test`) before creating a pull request.

## 6\. Commit & PR Message Format

Follow the **Conventional Commits** specification.

**Format:** `<type>(<scope>): <subject>`

  * **type**: `feat`, `fix`, `docs`, `style`, `refactor`, `test`, `chore`
  * **scope (optional)**: `server`, `client`, `admin`, `ui`, etc.
  * **subject**: A concise description of the change (max 50 characters).

## 7\. Prohibited Actions

1.  **Do not commit sensitive information**.
2.  **Do not push directly to the `main` branch**.
3.  **Do not break the file structure**. All new code must conform to the architecture defined above.
4.  **Do not create massive pull requests**.

-----

This guidebook will be updated as the project evolves. Always refer to the latest version.