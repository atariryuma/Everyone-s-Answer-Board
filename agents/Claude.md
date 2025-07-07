Of course. Here is a version of the coding AI's guidebook written in English, incorporating the file structure rules we discussed. You can copy and paste this directly into your project's `AGENTS.md` file.

-----

# 🤖 Agent's Guidebook for "Everyone's Answer Board"

This document provides the essential guidelines for all AI agents (and human developers) contributing to this project. Its purpose is to maintain code quality and consistency, and to promote efficient development.

## 1\. Agent's Role and Objective

Your role is not just a code generator, but a **product development partner with the perspective of a user in an educational setting.**

**Objective:**

  * To implement robust and maintainable features based on the requirements in `README.md`.
  * To pursue a high-quality user experience by always considering non-functional requirements such as performance, security, and accessibility.
  * To propose and implement an intuitive UI/UX for teachers and students who may not be familiar with IT, by envisioning how the tool will be used in the classroom.

## 2\. Core Principles

1.  **The requirements document is the single source of truth**: All implementations must be based on the requirements defined in `README.md`. If anything is unclear or contradictory, refer to the requirements first.
2.  **Simplicity is best**: Prioritize simple, readable code over complex solutions.
3.  **Security first**: Exercise the utmost caution when handling sensitive information, such as service account keys and personal data.
4.  **No code without tests**: All new features and bug fixes must be accompanied by tests.
5.  **User-centered design**: When developing features, always consider the perspectives of teachers and students and prioritize their convenience.

## 3\. Development Workflow

1.  **Understand Requirements**: Read the issue or task description and review the relevant sections of `README.md`.
2.  **Design**: Before implementation, mentally map out the structure of classes, functions, and the data flow.
3.  **Implement (Coding)**: Write code in the appropriate files within the `src` directory, following the style guide.
4.  **Test**: Create and run Jest tests for your new features in the `tests` directory.
5.  **Self-Review**: Before committing, ensure your code adheres to the core principles and guidelines.
6.  **Commit**: Commit your changes following the "Commit Message Format" described below.
7.  **Pull Request (PR)**: Create a PR with a clear description of the changes and the reasoning behind them.

## 4\. Code Style and File Structure Guide

### **Language and Conventions**

  * **Language**: Google Apps Script (GAS) uses the modern V8 runtime, which supports modern JavaScript.
  * **Style**: Adhere to the **Google JavaScript Style Guide**.
  * **Naming Conventions**:
      * Variables/Functions: `camelCase` (e.g., `getUserInfo`)
      * Classes: `PascalCase` (e.g., `StudyQuestApp`)
      * Constants: `UPPER_SNAKE_CASE` (e.g., `CACHE_TTL`)

### **【CRITICAL】File Organization and Separation of Concerns**

This project prioritizes maintainability within the Google Apps Script online editor. You **must** adhere to the following file structure rules.

#### **File Naming Convention**

All front-end files must be named using the format **`Role_Purpose.html`**. This ensures related files are grouped together alphabetically in the online editor.

| Role | Prefix | Description | Example |
| :--- | :--- | :--- | :--- |
| **View** | `View_` | The main HTML file that provides the page structure. | `View_Board.html` |
| **Style** | `Style_` | A file containing all CSS for the corresponding View. | `Style_Board.html` |
| **Script**| `Script_`| A file containing all client-side JS for the corresponding View.| `Script_Board.html`|
| **Component**| `Component_`| A reusable HTML snippet used across multiple Views. | `Component_Header.html`|

#### **Implementation Rules**

1.  **HTML (View\_\*.html)**:

      * This file should only contain the structural skeleton of the page.
      * CSS and JavaScript **must** be included from external files using GAS scriptlets `<?! ... ?>`. **Do not** write inline `<style>` or `<script>` tags directly in this file.
      * **Example (`View_Board.html`):**
        ```html
        <!DOCTYPE html>
        <html>
          <head>
            <base target="_top">
            <?! include('Style_Board'); ?>
          </head>
          <body>
            <?! include('Component_Header'); ?>
            <main id="answers"></main>
            <?! include('Script_Board'); ?>
          </body>
        </html>
        ```

2.  **CSS (Style\_\*.html)**:

      * The entire content of this file must be wrapped in a single `<style>` tag.

3.  **JavaScript (Script\_\*.html)**:

      * The entire content of this file must be wrapped in a single `<script>` tag.
      * **You cannot use `import`/`export` syntax.** All scripts are ultimately included in a single HTML scope. To avoid polluting the global scope, it is recommended to wrap your code in an **IIFE `(function(){ ... })();`**.

4.  **Server-Side (`Code.gs`)**:

      * You must use the `include(filename)` helper function to load front-end files.
        ```javascript
        function include(filename) {
          return HtmlService.createHtmlOutputFromFile(filename).getContent();
        }
        ```

### **Comments**

  * Leave comments for complex logic or code that may be difficult to understand.
  * Use **JSDoc** format for all public functions.

## 5\. Testing Policy

  * Add **unit tests** for all new business logic.
  * Tests should be created in the `/tests` directory with the `*.test.js` naming convention.
  * Ensure all existing tests pass (`npm run test`) before creating a pull request.

## 6\. Commit & PR Message Format

Follow the **Conventional Commits** specification to ensure commit messages are clear and descriptive.

**Format:** `<type>(<scope>): <subject>`

  * **type**: `feat`, `fix`, `docs`, `style`, `refactor`, `test`, `chore`
  * **scope (optional)**: `auth`, `admin`, `ui`, etc.
  * **subject**: A concise description of the change (max 50 characters).

## 7\. Prohibited Actions

1.  **Do not commit sensitive information**: Never include API keys, passwords, or service account JSON files in the codebase.
2.  **Do not push directly to the `main` branch**: All changes must go through a pull request.
3.  **Do not break formatting**: Do not disable formatters like `prettier` or `eslint`.
4.  **Do not create massive pull requests**: Break down work into small, logical PRs for each feature or fix.

-----

This guidebook will be updated as the project evolves. Always refer to the latest version.