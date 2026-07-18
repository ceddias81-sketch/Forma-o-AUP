# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project overview

Two-file HTML e-learning on **Antimicrobial Stewardship (AMS)** for the Serviço de Atendimento Urgente Pediátrico (AUP) at Hospital da Luz Lisboa. Author: Carlos Dias, Enfermeiro especialista, Liaison Officer EU-JAMRAI II (DGS).

| File | Purpose |
|---|---|
| `index.html` | 20-slide presentation (all CSS/JS/logo inlined) |
| `questionario.html` | Evaluation form — Likert survey + MCQ quiz + Formspree submission |

No build step, no dependencies, no package manager. Both files must sit in the **same directory** — `questionario.html` links back via `<a href="index.html">`. Open either file directly in a browser or serve via GitHub Pages.

**Known gap:** `index.html` has no link forward to `questionario.html`. Adding that link is a future task.

## Shared design tokens

Both files share an **identical** `:root` CSS block. If you change any token, update it in **both** files manually:

```css
:root {
  --ink:#102A33; --ink-2:#0A1E24;
  --paper:#F5F1E8; --paper-2:#FBF8F1;
  --teal:#14857B; --teal-d:#0E5F58; --teal-l:#3FB0A4;
  --mist:#D6E7E2; --coral:#E2513F; --coral-d:#C23A2A;
  --amber:#F2B33C; --sky:#8FC9E8;
}
```

Verify sync: `grep -o '\-\-teal:[^;]*' index.html questionario.html` — both lines must match.

## index.html — slide engine

**Structure:** `<div class="stage" id="stage">` → `<section class="scene">` (one per slide).  
CSS classes `is-active`, `is-before`, `is-after` drive scale/opacity transitions.

**Scene variants:** `.scene--dark` · `.scene--ink2` · `.scene--teal`  
**Entrance animations:** CSS `--d` (ms delay) on `.r` (rise) / `.pop` / `.fade` children.

**Interactive elements:**
- `.tap` cards — click/tap to flip front→back (case studies, myths)
- `#colony` — 200 `<span class="b">` dots classed `.die` / `.live` / `.grow`
- `#quiz` — 5-question MCQ driven by `quizData` JS array; `a` property = 0-based index of correct option

**Navigation:** `#prev` / `#next` buttons, `#dots`, keyboard (←→ Space PageUp PageDown Home End), swipe, left/right click zones. `go()` updates `#progbar`, `#cur`, `#tot`.

**Slides:** numbered in HTML comments `<!-- 1 ABERTURA -->` … `<!-- 20 ENCERRAMENTO -->`. JS counts `.scene` elements automatically — no numbering needed in code.

**Logo:** `HOSPITAL_LOGO` constant near bottom of `<body>` holds a `data:image/jpeg;base64,...` string (real Hospital da Luz Lisboa logo, JPEG format, not PNG). To swap: replace the base64 string, or set `HOSPITAL_LOGO = ""` and place `hospital-da-luz-lisboa.jpg` alongside `index.html`.

## questionario.html — evaluation form

**Sections:**
1. **Dados do Formando** — category (Médico/Enfermeiro/Outro), anonymous
2. **Avaliação da Formação** — 8 Likert 1–5 items + global rating
3. **Avaliação de Conhecimentos** — 5 MCQ with immediate per-question feedback; correct answers encoded in `data-correct` attribute (answer key: Q1=b, Q2=c, Q3=c, Q4=c, Q5=c)
4. **Comentários e Sugestões** — 3 free-text textareas

After submit: "Obrigado" screen showing quiz score + print/PDF button + `<a href="index.html">` back-link.

**Formspree:** form POSTs to `https://formspree.io/f/${FORMSPREE_ID}`. Current endpoint: `maqzlrky` (active). To change endpoint: edit `const FORMSPREE_ID = 'maqzlrky'` in the JS block. The form collects: category, all Likert responses, global rating, quiz answers + score, and free-text comments.

## Definition of done

Run these before committing any change:

```bash
# Slide count must stay at 20
grep '<section class="scene' index.html | wc -l

# Quiz items in questionario.html must stay at 5
grep -c 'class="quiz-item"' questionario.html

# Design tokens must be identical in both files
grep -o '\-\-teal:[^;]*' index.html questionario.html

# Formspree endpoint still active
grep 'FORMSPREE_ID' questionario.html

# Logo is JPEG, not PNG
grep 'data:image/jpeg' index.html
```

## Language

All content is in **European Portuguese** (pt-PT). Keep new content in the same language and register. Do not translate slide text or UI labels without explicit instruction.

## Open questions

1. Should `index.html` end with a "Avançar para o Questionário →" button linking to `questionario.html`? If yes, which slide (after 20 ENCERRAMENTO, or as slide 21)?
2. Are there other files planned (e.g., a certificate page, a facilitator guide)?
3. Should the design tokens be extracted to a shared `tokens.css` file, or must the project remain two standalone files?
