# GitHub Pages Integration

This `site/` folder is designed to be copied into `OthmaneBlial.github.io` as a static bundle.

## Recommended copy set

- `index.html`
- `docs.html`
- `styles.css`
- `app.js`
- `assets/`
- `docs/`
- `examples/`
- `generated/`
- `rendered/`
- `README.md`
- `CONTRIBUTING.md`
- `SUPPORT.md`
- `SECURITY.md`
- `CODE_OF_CONDUCT.md`

## Notes

- The docs viewer lives in `docs.html` and reads local markdown files with `fetch(...)`, so the final site should be served over GitHub Pages or another HTTP server rather than opened directly via `file://`.
- Relative links assume these files stay together in one folder tree.
- If you place the microsite in a subfolder of `OthmaneBlial.github.io`, keep the copied structure intact so `assets/`, `docs/`, `examples/`, `generated/`, and `rendered/` remain siblings of `index.html`.
