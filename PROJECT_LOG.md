# UQ Slide Converter — Project Log

## Version History

### v0.9.3 — Session 9 (12 April 2026)

**Azure Deployment (p2001)**
- Created Azure Container Registry (`bschoollearningtools`) on Basic tier, Australia East
- Created Web App `p2001` on existing B1 App Service Plan (`ASP-bschoollearningtools-83d7`)
- Set up GitHub Actions CI/CD pipeline (`.github/workflows/deploy-azure.yml`):
  - Triggers on push to `main` branch or manual dispatch
  - Builds Docker image → pushes to ACR → deploys to Azure Web App
  - Requires three GitHub secrets: `ACR_USERNAME`, `ACR_PASSWORD`, `AZURE_CREDENTIALS`
- Configured environment variables: `WEBSITES_PORT=8501`, `ANTHROPIC_API_KEY`, `GOOGLE_SHEETS_WEBHOOK_URL`
- Container registry credentials configured via `DOCKER_REGISTRY_SERVER_*` env vars and Azure CLI
- App live at: https://p2001-hmfcaaeebsgpg2g0.australiaeast-01.azurewebsites.net/
- Render instance (p2000) remains live in parallel
- B1 plan provides 1.75GB RAM (vs Render free tier 512MB) — 8MB memory guard can be raised/removed

**Deployment notes:**
- First deploy required manual container config fix via Azure CLI (`az webapp config container set`) — portal Deployment Center didn't propagate ACR credentials correctly on initial setup
- Subsequent deploys are automatic via GitHub Actions on push to `main`

**Feedback Report Download**
- Added "Download feedback report" button alongside the PPTX download
- Generates a Markdown (.md) file with full slide-by-slide details: handler, confidence, classification method, verification status, and all flagged issues
- Filename pattern: `*_FEEDBACK.md`

**Google Sheets Webhook — Further Fix**
- Previous `urllib` redirect handler still not working on Render or Azure
- Upgraded to use `requests` library (added to `requirements.txt`) with manual redirect following (`allow_redirects=False` + re-POST to redirect URL)
- Added comprehensive `INFO`-level logging throughout the webhook flow — check Azure Log stream after a conversion to see exactly where it fails/succeeds
- Falls back to `urllib` approach if `requests` is not installed

**App Version**
- Bumped `APP_VERSION` to 0.9.3

**Files added/modified:**
- `.github/workflows/deploy-azure.yml` — GitHub Actions CI/CD workflow
- `deploy-azure.sh` — CLI-based deployment script (alternative to portal+Actions approach)
- `app.py` — feedback download button, version bump
- `utils/cost_logger.py` — webhook rewrite with `requests` + better logging
- `requirements.txt` — added `requests>=2.31.0`
- `PROJECT_LOG.md` — recreated (was lost during repo copy-over)

**Pending / Known Issues:**
- Google Sheets webhook: improved logging deployed but not yet confirmed working — check Azure Log stream after next conversion
- ToC only generates with 3+ Section Divider slides (by design) — most NFS decks don't have these
- References slide only generates with 2+ collected references (by design)
- 8MB memory guard still active — can be raised/removed now that Azure B1 has 1.75GB RAM
- Slide 5 (phantom bullet points behind purple box), Slide 12 (missing pie image), Slide 25 (title/text overlap) — need source file investigation to fix underlying handler logic

---

### v0.9.2 — Session 8 (12 April 2026)

**Table of Contents Generator**
- Created `utils/toc.py` — auto-generates a Contents slide from detected Section Divider slides
- Uses Contents 2 template layout (index 4) with numbered table (01, 02, ...)
- Smart swap logic for misplaced section titles/labels in source decks
- Minimum 3 sections required; skips if ToC already exists in first 5 slides
- Inserts after Acknowledgement of Country slide

**Google Sheets Webhook Fix**
- Root cause: Python's `urllib` converts POST→GET on 302 redirect (Google Apps Script always returns 302)
- Fix: Custom `PostRedirectHandler` in `utils/cost_logger.py` that re-POSTs to redirect URL

**Slide Visualiser Redesign**
- Replaced single-slide scroll-through with expandable panels for all slide pairs
- Added filter: All slides / Issues only / Passed only
- Auto-expands slides with verification issues

**Memory Guard**
- Added 8MB file size threshold — files above this skip LibreOffice rendering and AI classification
- Falls back to heuristic-only conversion with user-facing message
- Addresses Render free tier 512MB RAM limit (can be raised/removed on Azure B1 with 1.75GB)

**Diagram Preservation Fix**
- Fixed group shapes being lost on diagram-dominant slides
- If group shapes exist but <100 chars of non-group text, groups are preserved even for TEXT_HANDLERS

**Handler Cleanup Fix**
- Fixed `_cleanup_empty_placeholders()` crash with list-type placeholder values
- Added `if not isinstance(ph_idx, int): continue` guard

**Test Deck**
- Created `generate_test_deck.py` — generates 34 "before" slides, one per handler
- Output: `test_decks/handler_test_before.pptx` and `test_decks/handler_test_BRANDED.pptx`

---

### v0.9.1 — Sessions 1–7 (prior)

**Core Architecture**
- Streamlit web app (`app.py`) for file upload and conversion
- `converter.py` — main conversion pipeline with handler dispatch
- Template-based slide creation using python-pptx
- CRITICAL RULE: ALL formatting inherited from slide master/theme via paragraph levels; NEVER set font names, sizes, or colours directly
- Template level hierarchy: Level 0 = bold purple heading, Level 1 = body text, Level 2 = bullet, Level 3 = sub-bullet

**Handler System**
- Base class `SlideHandler` with `detect()`, `extract_content()`, `fill_slide()` methods
- 34+ specialised handlers in `handlers/` directory
- Detection confidence scores: 0.0–1.0 (confident >=0.7, flagged 0.35–0.69, skipped <0.35)
- Dual detection: layout name matching + heuristic content analysis

**AI Integration**
- Anthropic Claude API for slide classification and content verification
- Cost logging to Google Sheets via Apps Script webhook
- Three-tier classification pipeline: layout match → heuristic → AI fallback

**Deployment**
- Docker container with LibreOffice for slide rendering
- Render deployment (`render.yaml`)
- Azure deployment (added in v0.9.3)

**Template**
- UQ Business School branded PPTX template with 20+ slide layouts
- Acknowledgement of Country auto-insertion
- References slide auto-generation
