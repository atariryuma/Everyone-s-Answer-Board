# Gemini Agent Documentation

This document outlines how the Gemini agent interacts with the "Everyone's Answer Board" project.

## Project Overview

This is a Google Apps Script project managed with `clasp`. The main application logic is in the `src` directory, and tests are in the `tests` directory.

## Key Technologies

- **Google Apps Script:** The core application logic.
- **clasp:** The command-line tool for managing Google Apps Script projects.
- **Jest:** The testing framework.
- **Node.js:** For utility scripts.

## How to Work with this Project

1.  **Understanding the Code:** The main logic is in `src/Code.gs`. The frontend is likely composed of the `.html` files in `src`.
2.  **Running Tests:** Use the `npm test` command to run the Jest test suite.
3.  **Deployment:** Use `clasp push` to deploy changes to the Google Apps Script project.
