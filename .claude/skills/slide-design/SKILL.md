---
name: slide-design
description: "Design visually distinctive presentation slides as an HTML/CSS mockup file.
  Use when the user wants to design, plan, or visualize slide aesthetics and layouts before
  creating a .pptx file. Output is an HTML file that ppt-creator can convert to PowerPoint.
  Trigger when the user mentions slide design, presentation design, 슬라이드 디자인, or asks
  to design a deck before making the actual PowerPoint file."
license: MIT
---

# Slide Design Skill

This skill creates visually distinctive slide designs as HTML/CSS mockups, combining production-grade frontend design aesthetics with presentation-specific visual principles.

**Pipeline**: This skill outputs an HTML file. Pass it to `ppt-creator` to generate the actual `.pptx` file.

---

## Design Thinking

Before writing any HTML, understand the context and commit to a BOLD aesthetic direction:

- **Purpose**: What problem does this presentation solve? Who is the audience?
- **Tone**: Pick an extreme — executive/refined, bold/modern, editorial/magazine, minimal/clean, playful, retro-futuristic, brutalist, art deco, etc. Commit fully.
- **Constraints**: Number of slides, content type, formal vs. casual.
- **Differentiation**: What makes this deck UNFORGETTABLE? What's the one thing someone will remember?

**CRITICAL**: Choose a clear conceptual direction and execute it with precision. Bold maximalism and refined minimalism both work — the key is intentionality, not intensity.

---

## Frontend Aesthetics Guidelines

Focus on:

- **Typography**: Choose fonts that are beautiful, unique, and interesting. The primary font is **Pretendard** — use it consistently for all text. Vary weight (100–900) and size to create hierarchy. Avoid generic fallback rendering by always loading via CDN.
- **Color & Theme**: Commit to a cohesive aesthetic. Dominant colors with sharp accents outperform timid, evenly-distributed palettes. One color should dominate (60–70% visual weight).
- **Motion**: Use CSS animations for entrance effects and micro-interactions. Staggered reveals (animation-delay) on page load create delight without clutter.
- **Spatial Composition**: Unexpected layouts. Asymmetry. Overlap. Grid-breaking elements. Generous negative space OR controlled density — never undefined middle ground.
- **Backgrounds & Visual Details**: Create atmosphere and depth. Add gradient meshes, geometric patterns, layered transparencies, dramatic shadows, grain overlays, or decorative borders that match the overall aesthetic.

NEVER use generic AI aesthetics: predictable layouts, clichéd color schemes (purple gradient on white), cookie-cutter component patterns, or designs that lack context-specific character. No design should look the same as another.

---

## Slide-Specific Color Palettes

Pick colors that match THIS topic — don't default to generic blue. Use as inspiration:

| Theme | Primary | Secondary | Accent |
|-------|---------|-----------|--------|
| **Midnight Executive** | `1E2761` (navy) | `CADCFC` (ice blue) | `FFFFFF` (white) |
| **Forest & Moss** | `2C5F2D` (forest) | `97BC62` (moss) | `F5F5F5` (cream) |
| **Coral Energy** | `F96167` (coral) | `F9E795` (gold) | `2F3C7E` (navy) |
| **Warm Terracotta** | `B85042` (terracotta) | `E7E8D1` (sand) | `A7BEAE` (sage) |
| **Ocean Gradient** | `065A82` (deep blue) | `1C7293` (teal) | `21295C` (midnight) |
| **Charcoal Minimal** | `36454F` (charcoal) | `F2F2F2` (off-white) | `212121` (black) |
| **Teal Trust** | `028090` (teal) | `00A896` (seafoam) | `02C39A` (mint) |
| **Berry & Cream** | `6D2E46` (berry) | `A26769` (dusty rose) | `ECE2D0` (cream) |
| **Sage Calm** | `84B59F` (sage) | `69A297` (eucalyptus) | `50808E` (slate) |
| **Cherry Bold** | `990011` (cherry) | `FCF6F5` (off-white) | `2F3C7E` (navy) |

- **Dominance over equality**: One color at 60–70% visual weight.
- **Dark/light contrast**: Dark backgrounds for title + conclusion, light for content ("sandwich") — or fully dark for premium feel.
- **Pick a bold, content-informed palette**: If swapping your colors into a different presentation would still "work," you haven't made specific enough choices.

---

## Typography (Slide-Specific)

All slides use **Pretendard** as the primary font. Use weight variation for hierarchy.

| Element | Size | Weight |
|---------|------|--------|
| Slide title | 36–44pt (48–58px) | 700–800 |
| Section header | 20–24pt (26–32px) | 600–700 |
| Body text | 14–16pt (18–21px) | 400 |
| Captions / labels | 10–12pt (13–16px) | 300–400, muted color |

- **Commit to a visual motif**: Pick ONE distinctive element and repeat it — rounded image frames, icons in colored circles, thick single-side borders. Carry it across every slide.
- **Left-align body text**: Center only titles. Never center paragraphs or lists.
- **Size contrast**: Titles need 36pt+ to stand out from 14–16pt body.

---

## Layout Options (Per Slide)

**Every slide needs a visual element** — image, shape, icon, or chart. Text-only slides are forgettable.

Choose from:
- **Two-column**: text left, illustration or visual on right
- **Icon + text rows**: icon in colored circle, bold header, description below
- **2×2 or 2×3 grid**: image on one side, grid of content blocks on the other
- **Half-bleed shape/color**: full left or right side colored panel with content on the other side
- **Big stat callout**: large number (60–72px) with small label below
- **Comparison columns**: before/after, pros/cons, side-by-side options
- **Timeline or process flow**: numbered steps, arrows, or connected circles
- **Quote / testimonial**: large italic text with attribution and accent decoration

---

## Spacing Rules

- **Slide padding**: minimum 40px on all sides (equivalent to 0.5" margins)
- **Between content blocks**: 24–40px consistently
- **Leave breathing room** — don't fill every inch
- **Choose one gap size** (24px or 40px) and use it consistently within a slide

---

## Design Anti-Patterns to Avoid

- **NEVER use accent lines under titles** — hallmark of AI-generated slides; use whitespace or background color instead
- **Never repeat the same layout** — vary columns, cards, and callouts across slides
- **Never create text-only slides** — add shapes, colored panels, icons, or visual accents
- **Never give all colors equal weight** — one must dominate
- **Never default to blue** — pick colors specific to the topic
- **Never mix spacing randomly** — choose consistent gaps
- **Never use low-contrast elements** — text AND icons need strong contrast against the background

---

## HTML Mockup Workflow

### 1. Understand Context
Ask (or infer from context):
- Presentation topic and purpose
- Target audience (executive, technical, general)
- Number of slides and rough content outline
- Preferred tone/aesthetic (if user has a preference)

### 2. Commit to Aesthetic Direction
Pick a bold, specific direction. State it explicitly before writing code:
> "I'm going for a **dark executive** feel with deep navy (#1E2761), ice blue accents, and strong geometric shapes. Pretendard 800 for titles, 400 for body."

### 3. Build the HTML File

Create a single HTML file with all slides as scrollable `960px × 540px` divs:

```html
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Slide Design Mockup</title>
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css">
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { background: #1a1a1a; padding: 40px; display: flex; flex-direction: column; gap: 32px; }
    .slide {
      width: 960px;
      height: 540px;
      font-family: 'Pretendard', sans-serif;
      overflow: hidden;
      position: relative;
    }
    /* slide-specific styles below */
  </style>
</head>
<body>
  <!-- Slide 1: Title -->
  <div class="slide slide-title"> ... </div>
  <!-- Slide 2: Section Divider -->
  <div class="slide slide-section"> ... </div>
  <!-- Slide 3+: Content slides -->
  <div class="slide slide-content"> ... </div>
</body>
</html>
```

### 4. Required Slide Types

Cover at minimum:

| Type | Description |
|------|-------------|
| **Title** | Bold opening slide — presentation title, subtitle, presenter/date |
| **Section Divider** | Visual break between major sections — large section number or title |
| **Content (varied)** | 2+ content layouts (two-column, icon grid, stat callout, etc.) |
| **Highlight/Callout** | Big number, key quote, or single powerful message |

Use different layouts for each content slide — never repeat consecutively.

### 5. Output

Save as `slide-design.html` (or a descriptive name) in the working directory.

---

## Handoff to ppt-creator

After completing the HTML mockup, summarize the design spec for ppt-creator:

```
Design Spec Summary
-------------------
HTML file: slide-design.html
Color palette:
  - Primary: #1E2761 (navy)
  - Secondary: #CADCFC (ice blue)
  - Accent: #FFFFFF (white)
  - Background: #F8F9FA (light content slides)
Font: Pretendard
  - Titles: 800 weight, 44px
  - Headers: 700 weight, 24px
  - Body: 400 weight, 16px
Layouts used:
  - Slide 1: Full-bleed dark background, centered title
  - Slide 2: Section divider with large number
  - Slide 3: Two-column (text left, shape right)
  - Slide 4: 2×2 icon grid
  - Slide 5: Big stat callout
Visual motif: Thick left-side color bar on content slides
```

The ppt-creator skill will read this HTML and spec to recreate the design as a `.pptx` file.
