# Kitchen Roadmap - Webapp

Miro/Asana-grade roadmap voor de Keukengerij launch (EU10, dropship, 10 producten, 90 dagen).

## Wat het is

- **128 taken** verdeeld over **6 fasen** (F0-F5)
- **Filters per eigenaar**: Ali / Kheibar / Farshad / Designer / Hermes
- **Inline edit**: owner aanpassen, taken toevoegen, taken verwijderen
- **Status tracking**: per taak checkbox + comment
- **Multi-user sync**: state wordt opgeslagen in GitHub repo (gratis, geen Supabase)
- **Auto-save**: 1 seconde na elke wijziging
- **Miro/Asana-grade UI**: clean, snel, geen externe libraries

## Tech stack

- **Frontend:** Vanilla HTML/CSS/JS (geen React, geen Tailwind, geen build)
- **Backend:** Vercel serverless function (`api/state.js`) → GitHub Contents API
- **Storage:** `data/state.json` in de GitHub repo (versie-controlled via commits)
- **Hosting:** Vercel (gratis hobby tier)

## Setup

### 1. GitHub Personal Access Token

Maak een PAT: https://github.com/settings/tokens/new
- Scope: `repo` (full control of private repositories)
- Note: `kitchen-roadmap-vercel-function`

### 2. Vercel Environment Variable

```bash
echo -n "$GITHUB_TOKEN_VALUE" | vercel env add GITHUB_TOKEN production --scope agentle --yes
```

### 3. Deploy

```bash
vercel deploy . --scope agentle --prod --yes
```

### 4. Verify

Open https://kitchen-roadmap.vercel.app — eerste load haalt state uit GitHub.

## Companion docs

- `data/state.json` — alle taken + state (auto-saved)
- `kitchen-launch-proposal.pptx` — proposal voor Kheibar/Farshad
- `../plans/kitchen-launch-werkdocument.md` — source of truth voor beslissingen