# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project overview

Single-file HTML e-learning presentation on **Antimicrobial Stewardship (AMS)** for the Serviço de Atendimento Urgente Pediátrico (AUP) at Hospital da Luz Lisboa. Author: Carlos Dias, Enfermeiro especialista, Liaison Officer EU-JAMRAI II (DGS).

The entire application lives in `index.html` — no build step, no dependencies, no package manager. Open the file directly in a browser.

## Architecture

All CSS, SVG icons, JavaScript, and the Hospital da Luz Lisboa logo (base64 data-URI) are inlined in `index.html`.

**Slide engine** (`<div class="stage">` → `<section class="scene">`):
- Each `<section class="scene">` is one slide. CSS classes `is-active`, `is-before`, `is-after` drive the scale/opacity transitions.
- Scene variants: `.scene--dark` (dark ink), `.scene--ink2` (deep dark gradient), `.scene--teal` (teal gradient).
- Entrance animations use CSS custom property `--d` (delay in ms) on elements with class `.r` (rise), `.pop`, or `.fade`.

**Interactive elements:**
- `.tap` cards — click/tap to flip front→back (case studies, myths).
- Colony visualization (`#colony`) — animated bacterial resistance metaphor built with 200 `<span class="b">` dots assigned `.die`, `.live`, or `.grow` classes.
- Quiz (`#quiz`) — 5-question self-assessment driven by the `quizData` array in JS.

**Navigation:**
- Arrow buttons (`#prev`, `#next`), dot indicators (`#dots`), keyboard (←/→/Space/PageUp/PageDown/Home/End), swipe, and click-zone (left third / right third of stage).
- Progress bar (`#progbar`) and slide counter (`#cur` / `#tot`) update on every `go()` call.

## Editing slides

Slides are numbered in HTML comments (`<!-- 1 ABERTURA -->` … `<!-- 20 ENCERRAMENTO -->`). To add a slide, insert a new `<section class="scene">` inside `<div class="stage">` — the JS counts `.scene` elements automatically.

## Logo

The `HOSPITAL_LOGO` constant near the bottom of `<body>` holds the base64 PNG. To replace the logo, swap the base64 string or set `HOSPITAL_LOGO = ""` and place `hospital-da-luz-lisboa.png` alongside `index.html`.

## Quiz

Edit the `quizData` array in the JS to change questions, options, or correct answers (`a` index).

## Language

All content is in **European Portuguese** (pt-PT). Keep new content in the same language and register.
