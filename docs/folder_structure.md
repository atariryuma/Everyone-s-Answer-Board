# Folder Structure for Google Apps Script Project

This document describes the recommended folder structure for the "Everyone's Answer Board" project. It follows the guidelines from `agents/AGENTS.md` and `REFACTORING_GUIDE.md` so that large HTML files can be split into smaller pieces and managed easily within the Google Apps Script online editor.

## Top Level Layout

```
/src
├── server/               # Apps Script server‑side code
│   ├── main.gs           # doGet and include helpers
│   ├── database.gs       # Spreadsheet I/O
│   └── services/         # Business logic
│
├── client/               # Front‑end resources
│   ├── views/            # HTML templates
│   ├── styles/           # `.css.html` files (wrapped in <style>)
│   ├── scripts/          # `.js.html` files (wrapped in <script>)
│   └── components/       # Reusable HTML snippets
│
└── appsscript.json       # GAS manifest (do not edit)
```

### Example View File

`src/client/views/AdminPanel.html` should include CSS, scripts and components via the `include` helper:

```html
<!DOCTYPE html>
<html lang="ja">
  <head>
    <base target="_top">
    <?!= include('styles/main.css.html'); ?>
    <?!= include('styles/AdminPanel.css.html'); ?>
  </head>
  <body>
    <?!= include('components/Header.html'); ?>

    <main id="app"></main>

    <?!= include('scripts/main.js.html'); ?>
    <?!= include('scripts/AdminPanel.js.html'); ?>
  </body>
</html>
```

### Include Helper

Define the following in `src/server/main.gs` so nested includes work:

```javascript
function include(path) {
  const template = HtmlService.createTemplateFromFile('client/' + path);
  template.include = include;
  return template.evaluate().getContent();
}
```

### Naming Conventions

- CSS snippets use the extension `.css.html` and must be wrapped in a `<style>` tag.
- JavaScript snippets use `.js.html` and must be wrapped in a `<script>` tag.
- Component files contain reusable HTML fragments.

With this structure, each part of the UI—HTML, CSS, and JavaScript—stays modular and easier to maintain, even when edited from the GAS online editor.
