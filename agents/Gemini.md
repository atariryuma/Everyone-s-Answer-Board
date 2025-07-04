# Developer Documentation: Building Modern Web Applications

This document provides a comprehensive guide for developers aiming to build robust, user-friendly, and scalable web applications. It covers core design principles, a recommended technology stack, and best practices for coding and architecture, using the provided project as a reference.

## 1. Core Philosophy & Design Principles

The foundation of a successful application lies in a strong philosophy that prioritizes the user.

  * **User Safety & Trust**: Create a secure environment where users feel safe to interact and share information.
  * **Inclusivity**: Design for a diverse audience, ensuring accessibility and ease of use for everyone.
  * **Empathetic Interaction**: Implement features that allow for clear, positive, and nuanced communication.
  * **Intuitive UI/UX**:
      * **Simplicity**: Avoid unnecessary complexity. The user interface should be clean and straightforward.
      * **Visual Hierarchy**: Use modern design techniques like "glassmorphism" to create a sense of depth and guide the user's focus. The `.glass-panel` class in the project is a good example of this.
      * **Accessibility**: Ensure high contrast, readable fonts, and keyboard navigability to support all users.

## 2. Color Palette & Usage

A consistent color palette is key to a professional look and feel. This palette is defined using CSS variables for easy theming and maintenance.

**Copy-paste this CSS to use the color palette in your project:**

```css
:root {
  --color-primary: #8be9fd;
  --color-background: #1a1b26;
  --color-surface: rgba(26, 27, 38, 0.7);
  --color-text: #c0caf5;
  --color-border: rgba(255, 255, 255, 0.1);
  --color-accent: #facc15;
  --color-success: #10b981;
  --color-error: #ef4444;
  --color-warning: #f59e0b;
  --color-info: #3b82f6;
}
```

### Color Usage Guide

| Variable              | Hex/RGBA               | Usage Example & Description                                                                                                                                     |
| :-------------------- | :--------------------- | :-------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `--color-primary`     | `#8be9fd` (Cyan)        | **Primary Actions & Highlights:** Used for main buttons, active states, and important interactive elements to draw user attention. (e.g., "Submit" button) |
| `--color-background`  | `#1a1b26` (Dark Blue)   | **Main App Background:** Provides a dark, modern backdrop that helps content and interactive elements stand out clearly.                                |
| `--color-surface`     | `rgba(26,27,38,0.7)`   | **Panel & Card Backgrounds:** Used for "glassmorphism" panels (`.glass-panel`) to create a translucent, layered effect.                            |
| `--color-text`        | `#c0caf5` (Light Gray)  | **Primary Text Color:** Ensures readability against the dark background for all main text content.                                                   |
| `--color-border`      | `rgba(255,255,255,0.1)`| **Borders:** Defines the edges of panels and other UI elements, enhancing the "glass" effect.                                                     |
| `--color-accent`      | `#facc15` (Yellow)     | **Accent & Attention:** Used for secondary highlights, titles, or important icons to provide visual contrast. (e.g., Main Title)               |
| `--color-success`     | `#10b981` (Green)      | **Success Feedback:** For success messages, confirmation indicators, and positive actions. (e.g., "Saved successfully" message)                  |
| `--color-error`       | `#ef4444` (Red)         | **Error Feedback:** For error messages, warnings about destructive actions, and validation failures. (e.g., "Invalid URL" error)                |
| `--color-warning`     | `#f59e0b` (Yellow)      | **Warnings:** Used for non-critical warnings or to draw attention to important information that requires user consideration.                           |
| `--color-info`        | `#3b82f6` (Blue)        | **Informational Messages:** For neutral, informational messages, tips, and guidance within the UI.                                                   |

## 3. Architecture and Technology Stack

This application is built on the Google Workspace platform, making it highly integrated and scalable.

  * **Backend**: **Google Apps Script (GAS)** serves as the serverless backend, handling all business logic, data processing, and API integrations. It leverages the **V8 runtime** for modern JavaScript features and improved performance.
  * **Frontend**:
      * **HTML/CSS/JavaScript**: Standard web technologies are used to build the user interface.
      * **Tailwind CSS**: A utility-first CSS framework is used for rapid and consistent styling. For simplicity in a GAS environment, the **Tailwind CSS CDN** is used. Include the following script tag in the `<head>` of your HTML files:
        ```html
        <script src="https://cdn.tailwindcss.com"></script>
        ```
      * **Client-Server Communication**: The frontend communicates with the GAS backend via the `google.script.run` asynchronous API, further optimized by `ClientOptimizer.html`.
  * **Data Storage**: **Google Sheets** is used as the primary data store for application data, user information, and configurations. It is accessed efficiently via **Google Sheets API v4** direct calls.
  * **Authentication**: Implements a **Service Account Model** using JWT for secure and efficient access to Google APIs, resolving previous 403 errors and simplifying deployment to a **single GAS project**.
  * **Development Tools**:
      * **`clasp`**: The official command-line tool for managing GAS projects locally.
      * **`jest`**: A JavaScript testing framework for ensuring the reliability of client-side logic. Backend logic is tested via `src/test.gs`.

## 4. Coding Standards and Best Practices

### 4.1. The Manifest File (`appsscript.json`)

The manifest file is a critical JSON file that configures your Apps Script project.

  * **`timeZone`**: Set the script's timezone (e.g., `"Asia/Tokyo"`).
  * **`oauthScopes`**: List the *minimum* required OAuth scopes for your script to function. Current scopes include:
    * `https://www.googleapis.com/auth/script.external_request`
    * `https://www.googleapis.com/auth/forms`
    * `https://www.googleapis.com/auth/spreadsheets`
    * `https://www.googleapis.com/auth/drive`
    * `https://www.googleapis.com/auth/userinfo.email`
    * `https://www.googleapis.com/auth/script.scriptapp`
    * `https://www.googleapis.com/auth/script.container.ui`
  * **`webapp`**: Configure the web app deployment settings.
      * `executeAs`: Defines whether the script runs as the user accessing it (`USER_ACCESSING`) or the developer who deployed it (`USER_DEPLOYING`).
      * `access`: Controls who can access the app (`MYSELF`, `DOMAIN`, or `ANYONE_ANONYMOUS`).
  * **`runtimeVersion`**: Use `V8` for the modern JavaScript runtime.
  * **`exceptionLogging`**: Set the destination for logged exceptions, with `STACKDRIVER` being a common choice.

### 4.2. Server-Side Script (`.gs` Files)

  * **Performance**:
      * **Batch Operations**: Minimize calls to services like `SpreadsheetApp`. Instead of reading or writing cell by cell in a loop, read a whole range of data into an array with `getValues()`, manipulate the array in your script, and write it back with `setValues()`.
      * **Cache Service**: Use `CacheService` (`cache.gs`) to cache frequently accessed, infrequently changing data to avoid redundant service calls.
      * **Advanced Optimizations**: Leverage `optimizer.gs` for time-bounded batching, exponential backoff, Sheets API rate limiting, memory optimization, and time-sliced parallel processing.
  * **Organization**:
      * **Modularity**: Separate concerns by splitting code into different `.gs` files (e.g., `main.gs`, `core.gs`, `database.gs`, `auth.gs`, `cache.gs`, `config.gs`, `url.gs`, `monitor.gs`, `optimizer.gs`, `stability.gs`, `test.gs`).
  * **Security**:
      * **Secrets Management**: Store API keys and other secrets in `PropertiesService`, not in the code itself.

### 4.3. Client-Side HTML (`.html` Files)

  * **Separation of Concerns**: Keep HTML structure, CSS styling, and JavaScript logic in separate files to make your project easier to read and maintain.
  * **Asynchronous Loading**: Load data dynamically with `google.script.run` after the initial page load to keep the UI responsive. The `ClientOptimizer.html` provides a `GASOptimizer` class for enhanced client-side call management (concurrency, caching, error handling).
  * **Use of Scriptlets (`<?...?>`)**: Use scriptlets sparingly for simple, one-time server-side tasks. They are executed before the page is served and can slow down the initial load if overused.
      * `<?= ... ?>` (Printing Scriptlet): Outputs data into the HTML with contextual escaping.
      * `<? ... ?>` (Standard Scriptlet): Executes server-side code like loops or conditionals.
  * **Security**: Trust the contextual escaping of printing scriptlets to prevent XSS. Sanitize any user data you manually place in the HTML. For all external links, use `rel="noopener noreferrer"`.

## 5. AI/ML Integration (Future Outlook)

Integrating with large language models can significantly enhance the application's capabilities.

### Potential Enhancements

  * **Content Summarization and Keyword Extraction**: To help users quickly grasp the main points of numerous responses.
  * **AI-Generated Feedback**: To provide users with constructive, personalized feedback.
  * **Data Analysis and Insights**: To analyze user-generated content and extract valuable trends and patterns.
  * **Content Moderation**: To automatically identify and filter inappropriate content, ensuring a safe user environment.

### Integration Strategy

  * The backend (GAS) can use the `UrlFetchApp` service to make HTTP requests to external AI model APIs.
  * Careful consideration must be given to API key management, rate limits, cost control, data privacy, and robust error handling.

## 6. Current Project Status and Development Context

### 6.1. Project Overview (StudyQuest - みんなの回答ボード)

This application provides an interactive answer board for educational settings, leveraging Google Sheets as a data store. Its core purpose is to facilitate real-time answer sharing from Google Forms to a dynamic web board, enabling student reactions and teacher highlights. The project has recently undergone a significant architectural shift to a **Service Account Model** within a **single Google Apps Script (GAS) project**, which has resolved previous 403 errors and streamlined deployment.

### 6.2. Key Features Implemented

*   **Real-time Answer Sharing**: Displays student responses from Google Forms on a web board.
*   **Reaction Features**: Students can react with "Understand!", "Like!", and "Curious!".
*   **Highlighting**: Teachers can highlight important answers.
*   **Admin Panel (`AdminPanel.html`)**: Comprehensive management interface for:
    *   Board publication/unpublication.
    *   Switching between different sheets.
    *   Configuring display modes (anonymous/named).
    *   Sorting options (score, newest, oldest, likes, random).
    *   Creating new boards and utilizing existing spreadsheets.
    *   Displaying the database's `isActive` status with a direct link to the associated spreadsheet.
*   **User Registration (`Registration.html`)**: Streamlined quick-start setup for new users.
*   **Service Account Setup (`SetupPage.html`)**: Dedicated page for initial service account and central database configuration.
*   **Robust Security**: Implements JWT + Google OAuth2 for authentication, domain-based access control, XSS prevention, and robust error handling.
*   **Performance Optimizations**: Includes advanced caching, batch processing, and efficient API calls.
*   **Stability Enhancements**: Incorporates circuit breakers, retry mechanisms, and health monitoring.

### 6.3. Recent Improvements and Bug Fixes

The project has seen several critical enhancements and bug fixes:

*   **New Service Account Architecture**: Full transition to a service account model within a single GAS project, eliminating 403 errors and simplifying setup.
*   **Comprehensive Performance Optimizations**: Integration of `optimizer.gs` and `cache.gs` for multi-level caching, time-bounded batching, Sheets API rate limiting, memory optimization, and time-sliced parallel processing.
*   **Enhanced Stability System**: Implementation of `stability.gs` for circuit breakers, resilient execution, health monitoring, failsafe data access, data integrity checks, and auto-recovery.
*   **Modular Codebase Refinement**: Further separation of concerns across `.gs` files, with core logic consolidated in `core.gs` and deprecated files (`DataProcessor.gs`, `ReactionManager.gs`) removed.
*   **Optimized Client-Side Communication**: `ClientOptimizer.html` provides a robust `GASOptimizer` class for managing `google.script.run` calls, including concurrency limits, caching, and enhanced error handling.
*   **Improved Admin Panel (`AdminPanel.html`)**: Enhanced UI/UX, better status indicators, and more intuitive setup flow.
*   **Refined Registration Flow (`Registration.html`)**: Clearer steps for new user registration and quick setup, with immediate feedback.
*   **Corrected Highlight Logic**: `processHighlightToggle` in `core.gs` now correctly uses `TRUE`/`FALSE` strings for boolean values in Google Sheets.
*   **Updated `getAppConfig`**: Now provides more comprehensive status information, including `spreadsheetUrl` and `systemStatus` details.
*   **Frontend Display Logic**: `Page.html` now includes a `StudyQuestApp` class for managing client-side state, rendering, and interactions, incorporating virtual scrolling and performance tweaks.

### 6.4. Current Known Issues and Future Work

*   **Client-Side Function Call Mismatch (`TypeError: ...[funcName] is not a function`)**: This error, observed in `src/Page.html` for functions like `checkAdmin`, `getAvailableSheets`, and `getPublishedSheetData`, is typically a deployment/caching issue. The server-side functions are correctly defined and exposed. **To resolve this, ensure the Apps Script project is properly redeployed and the browser's cache is cleared after any code changes.**
*   **Future Enhancements**:
    *   Further optimize client-side rendering performance.
    *   Implement more advanced analytics for board usage.
    *   Explore AI/ML integrations for content summarization or feedback generation.

### 6.5. Coding Practices and Conventions

*   **Modularity**: Code is strictly organized into logical `.gs` files, with `core.gs` serving as the central hub for optimized business logic.
*   **Performance-First**: Emphasis on batch operations, strategic caching (both `CacheService` and in-memory `Map`), and efficient Google Sheets API calls.
*   **Robust Error Handling**: All functions include comprehensive error handling with user-friendly messages for the frontend and detailed logs for debugging.
*   **Security**: Adherence to Google Apps Script security best practices, including proper OAuth scopes, input sanitization, and secure property management.
*   **Frontend Development**: Standard HTML/CSS/JavaScript practices are followed, with Tailwind CSS CDN for styling. `google.script.run` is the primary method for asynchronous client-server communication.
*   **Testing**: Unit tests are in place (`src/test.gs`) to ensure code reliability and prevent regressions. Developers should maintain high test coverage.
