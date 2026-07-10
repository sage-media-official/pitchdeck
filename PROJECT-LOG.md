# Sage Portfolio — Project Log

Everything built in this workspace, what state it's in, and where things stand.
Last updated: **9 July 2026**

---

## 1. Deliverables built so far

| # | Deliverable | Files | Status |
|---|------------|-------|--------|
| 1 | **Echon annual marketing proposal** (web + deck) | `echon-proposal.html`, `Echon-Annual-Marketing-Plan-SageMedia.pptx` / `.pdf`, `slides/build_echon_deck.py` | ✅ Done. Agency fee = 20% carved inside ₹3.99 Cr budget. Contact: business@sagemedia.in |
| 2 | **Bucketlist 3-month social media plan** | `bucketlist-social-plan.html` | ✅ Done. 20 posts/month (10 reels + 10 carousels), competitor comparison, print CSS fixed |
| 3 | **Sage design sample deck** (look & feel from Behance refs) | `Sage-Proposal-Design-Sample.pptx` / `.pdf`, `slides/design_sample.py` | ✅ Done |
| 4 | **Mach Five GTM & growth plan** (health drink) | `MachFive-GTM-Growth-Plan-SageMedia.pptx` / `.pdf`, `slides/build_machfive_deck.py` | ✅ Done. Market size, CAGR, spend plan, breakeven ~M6, milestone triggers |
| 5 | **Pitch Builder platform** (the big one) | `pitch-builder.html` | ✅ Working — see section 2 |
| 6 | **Bandhani Ethnic pitch** (research JSON) | `bandhani-ethnic-pitch.json` | ✅ Done — Import into Pitch Builder → Instant preview |
| 7 | **Vercel deploy bundle** | `build/index.html`, `build/gi-regular.woff2`, `build/gi-bold.woff2`, `build/vercel.json` | ⚠️ Partially deployed — see section 3 |

**Brand system used everywhere:** cream `#FDFBEE` + sage green `#517354/#37503B`, Glacial Indifference (display, real OTFs in `fonts/`), Inter (body), Space Grotesk (numbers).

---

## 2. The Pitch Builder (`pitch-builder.html`)

One self-contained HTML file. Open it in Chrome. What it does:

- **Minimal input**: brand basics + tap-select scope chips + ONE brief box (type or 🎤 dictate — voice needs Chrome over https/localhost, not file://).
- **🚀 Create pitch deck**: runs its own live research in the background with a timer (~30–90s). Uses the Claude API directly from the browser — needs an Anthropic API key (🔑 button, saved only in that browser's localStorage). Searches the web, sizes the market, maps 5 real competitors, audits site/socials/sentiment, then renders the deck.
- **⚡ Instant preview**: renders immediately with placeholders (or with imported research JSON).
- **20-section dense pitch**: cover → brief → snapshot → market forecast → tailwinds → channels → personas → competitor table → gaps → advantage → positioning → audit + sentiment gauge → GTM → channel mix → content pillars → funnel → 12-month roadmap → KPIs → **editable pricing table (left blank on purpose)** → next steps/CTA.
- **🎲 Shuffle design**: 20 industry-researched design systems (Editorial Cream, Fintech Navy, Luxe Ivory, Midnight Venture dark, Sage Heritage, etc.) — every generation can look different. Builder chrome itself stays locked to sage green.
- **✨ Magic edit**: toggle it, click any element on the pitch, leave a comment (or edit text inline). "Copy for Claude" bundles all change requests to paste back to me.
- **Export / Import**: full brief + research as JSON (that's how `bandhani-ethnic-pitch.json` works).
- **⎙ Print / PDF**: fixed on 9 July — see section 4.

---

## 3. Vercel integration — what happened and current state

Goal: host the Pitch Builder on Vercel so voice dictation and the one-click research button work over https.

What was done:
1. Built a slim deploy bundle in `build/`: the same app but with the Glacial Indifference fonts as external `.woff2` files (34KB) instead of embedded base64 (halves the HTML). Committed to the repo.
2. A Vercel project **`sage-pitch-builder`** was created (team: sage-media's projects). Production URL pattern: `https://sage-pitch-builder-sage-media-s-projects.vercel.app`.
3. Several deploy attempts went out through the Vercel integration. **The last completed deploy did NOT contain the full app** — the tool requires pasting the whole file inline and the 220KB+ app kept getting truncated/corrupted in transit. A gzip-packed loader approach (~30KB) was mid-flight when work was stopped.
4. A `gh-pages` branch with the static site was also pushed to this repo as a fallback, but GitHub Pages could not be enabled from this environment (API blocked by proxy).

**Bottom line: there is no verified working live link yet.**

### Fastest reliable way to finish it (recommended)
Skip the copy-paste deploys entirely — connect the repo to Vercel once, in the dashboard:
1. Go to **vercel.com → Add New → Project → Import Git Repository** → pick `sage-media-official/Sage-portfolio`.
2. Set **Root Directory = `build`**, framework = "Other" (static). Deploy.
3. Every future `git push` to the repo auto-deploys. The files in `build/` are already correct and committed.

That gives a permanent link like `https://sage-pitch-builder.vercel.app` in ~2 minutes, with none of the corruption risk.

---

## 4. Print/PDF fixes (9 July — in response to "slide 16 text not visible, alignment off")

Root causes found:
- **Section 16 (Funnel)**: the funnel steps are drawn with CSS `clip-path` shapes — most PDF engines drop them, leaving white text on a white page.
- **Counters/bars printed as "0"**: numbers animate only when scrolled into view; printing before scrolling froze them at zero.
- Sections splitting awkwardly across pages.

Fixes applied to `pitch-builder.html` (committed source; deploy bundle rebuild pending):
- `print-color-adjust: exact` everywhere so backgrounds actually print.
- Funnel prints as solid rounded bars (clip-path disabled in print), white text forced.
- New `finishAnims()` — snaps every counter, bar chart and gauge to its final value before printing; wired to the Print button and the browser's `beforeprint` event.
- `break-inside: avoid` on every card/table/chart/grid; ghost numbers and decorative blobs hidden in print; tighter page margins.

---

## 5. Repo layout

```
pitch-builder.html          ← the Pitch Builder app (source of truth)
build/                      ← slim Vercel bundle (index.html + woff2 fonts + vercel.json)
fonts/                      ← real Glacial Indifference OTFs
slides/                     ← python-pptx deck generators + font embedder
bandhani-ethnic-pitch.json  ← importable research for the Bandhani pitch
echon-proposal.html, bucketlist-social-plan.html, *.pptx, *.pdf
PROJECT-LOG.md              ← this file
```

Branches: work is on `claude/proposal-visual-design-n2i0qa`; `gh-pages` holds the static-site fallback.

---

## 6. Open items

- [ ] Connect repo to Vercel via dashboard (section 3) → get the permanent live link.
- [ ] Rebuild `build/index.html` from the print-fixed `pitch-builder.html` (one command, ask me).
- [ ] Verify the downloaded PDF end-to-end after the rebuild (esp. section 16).
