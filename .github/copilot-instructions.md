# GitHub Copilot Instructions

## Repository context
- **Project**: DB Auto Org Chart – Flask backend with static front-end (vanilla JS + D3).
- **Primary entrypoints**:
  - Backend: `app.py`
  - Front-end bundles: `static/app.js`, `static/configureme.js`, `static/reports.js`, etc.
  - Templates: `templates/*.html`
- **Data**: Cached JSON files under the `data/` directory (e.g., `employee_data.json`, report caches).

## Coding conventions
- Prefer modern JavaScript (`const`/`let`, template literals, async/await) without introducing frameworks.
- Maintain separation of concerns: keep inline scripts/styles out of templates; use the existing static files.
- Follow existing logging patterns (`logger.info`, `logger.error`) and error handling style in `app.py`.
- Keep CSS variables and shared styles in `static/styles.css` unless a page-specific stylesheet exists.
- Use i18n keys for user-facing text. Strings live in `static/locales/en-US.json`.

## Workflow expectations
1. **Plan**: Review related files before making changes; reuse helpers (e.g., `fetch_all_employees`, report loaders).
2. **Implement**: Apply minimal diffs using the preferred tools (`apply_patch`, `insert_edit_into_file`). Avoid rewriting unaffected code.
3. **Validate**: Run targeted checks when possible (e.g., linting, unit tests). If tooling is absent, note that in your summary.
4. **Document**: Update README or inline comments when behavior changes. Internationalized text requires locale updates.

## Feature guidelines
- **Reports**: When adding a report, create matching API routes, cache loaders, export endpoints, front-end configs, template options, and locale strings.
- **Translations**: Guard against untranslated flashes by toggling the `i18n-loading` class via JS once translations load.
- **Graph API**: Reuse existing helpers; ensure required permissions are documented (`User.Read.All`, `LicenseAssignment.Read.All`).
- **Caching**: Persist report data in `data/*.json` with graceful fallbacks if files are missing or unreadable.

## Pull request readiness checklist
- [ ] All relevant caches/readers updated.
- [ ] Locale strings added/updated for new UI text.
- [ ] Front-end selectors and summaries reflect new features.
- [ ] Backend routes include authentication decorators and error handling.
- [ ] Tests or manual validation steps recorded in the summary.

## Additional notes
- Do not introduce new dependencies without updating `requirements.txt` (backend) or documenting browser requirements.
- Keep responses concise and focused—summaries should mention actions, validation, and outstanding work.
- If unable to run automated checks (common in this repo), state that explicitly during the handoff.
