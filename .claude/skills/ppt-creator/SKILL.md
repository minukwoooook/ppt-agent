---
name: ppt-creator
description: "Create a .pptx PowerPoint file from an HTML slide design mockup or from scratch.
  Use when the user needs to generate, build, or produce an actual PowerPoint presentation file.
  Reads the HTML output from slide-design and recreates the design using PptxGenJS with Pretendard font.
  Trigger whenever the user mentions making, generating, or creating a .pptx, PowerPoint, or presentation file."
license: MIT
---

# PPT Creator Skill

This skill creates production-ready `.pptx` files using PptxGenJS, applying strong visual design principles and performing thorough QA.

**Pipeline**: This skill is the second step after `slide-design`. If an HTML mockup exists, read it to extract the design and recreate it faithfully in PptxGenJS. If working from scratch, apply the design principles below independently.

**Font**: All text in the PPTX must use `fontFace: "Pretendard"`. Note: Pretendard must be installed on the viewer's machine for correct rendering. Direct users to install it from [github.com/orioncactus/pretendard](https://github.com/orioncactus/pretendard) if needed.

---

## Quick Reference

| Task | Command |
|------|---------|
| Extract text from PPTX | `python -m markitdown presentation.pptx` |
| Visual thumbnail grid | `python scripts/thumbnail.py presentation.pptx` |
| Edit from template | See [editing.md](editing.md) |
| Create from scratch | See [pptxgenjs.md](pptxgenjs.md) |

---

## Reading Existing Content

```bash
# Text extraction
python -m markitdown presentation.pptx

# Visual overview
python scripts/thumbnail.py presentation.pptx

# Raw XML (for template editing)
python scripts/office/unpack.py presentation.pptx unpacked/
```

---

## Choosing a Creation Method

| Method | When to Use |
|--------|-------------|
| **PptxGenJS (from scratch)** | No template; or slide-design HTML mockup provided |
| **Template editing** | User provides an existing `.pptx` template to modify |

Read [pptxgenjs.md](pptxgenjs.md) in full before writing any PptxGenJS code.
Read [editing.md](editing.md) in full before modifying a template.

---

## HTML → PPTX Conversion Workflow

When a `slide-design.html` (or equivalent) file exists:

1. **Read the HTML file** to analyze the design:
   - Background colors and slide structure
   - Color palette (extract hex values — strip any `#` prefix for PptxGenJS)
   - Font sizes and weights (convert px to pt: divide by 1.333)
   - Layout structure per slide (positions, widths, element arrangement)
   - Visual motifs (accent bars, icon circles, decorative shapes)

2. **Map HTML → PptxGenJS coordinates**:
   - HTML slide: 960px × 540px
   - PPTX `LAYOUT_16x9`: 10" × 5.625"
   - Scale factor: `960px = 10"` → divide px by 96 to get inches
   - Example: `left: 40px` → `x: 0.42`, `width: 400px` → `w: 4.17`

3. **Recreate each slide in PptxGenJS**:
   - Match background color: `slide.background = { color: "1E2761" }`
   - Match shapes and colored panels with `slide.addShape()`
   - Match text with `slide.addText()` — always include `fontFace: "Pretendard"`
   - Match layout position and proportions

4. **Before running**: Re-read pptxgenjs.md's Common Pitfalls section.

5. **Run the script** and proceed to QA.

### px to pt conversion reference

| px | pt (approx) |
|----|-------------|
| 13px | 10pt |
| 16px | 12pt |
| 18px | 13.5pt |
| 21px | 16pt |
| 24px | 18pt |
| 32px | 24pt |
| 48px | 36pt |
| 58px | 44pt |

---

## Design Principles (for scratch creation)

If no HTML mockup exists, apply these principles before writing any code:

**Before coding, decide:**
- Color palette: primary (60–70% weight), secondary, accent — all as 6-char hex, no `#`
- Visual motif: one repeated element across all slides (e.g., thick left accent bar, icon circles)
- Layout variety: no two adjacent slides should use the same layout

**Font: Pretendard throughout**

| Element | fontSize | bold |
|---------|----------|------|
| Slide title | 36–44 | true |
| Section header | 20–24 | true |
| Body text | 14–16 | false |
| Caption/label | 10–12 | false |

**Every slide needs a visual element** — colored shape, icon, or chart. Never text-only slides.

**Anti-patterns**:
- NEVER accent lines under titles
- NEVER repeat the same layout consecutively
- NEVER use `#` in hex colors — corrupts the file
- NEVER reuse option objects across `addShape`/`addText` calls

---

## QA (Required)

**Assume there are problems. Your job is to find them.**

Your first render is almost never correct. Approach QA as a bug hunt, not a confirmation step.

### Content QA

```bash
python -m markitdown output.pptx
```

Check for missing content, typos, wrong order.

Check for leftover placeholder text:

```bash
python -m markitdown output.pptx | grep -iE "xxxx|lorem|ipsum|this.*(page|slide).*layout"
```

### Visual QA

**⚠️ USE SUBAGENTS** — even for 2–3 slides. You've been staring at the code and will see what you expect. Subagents have fresh eyes.

Convert slides to images (see Converting to Images below), then use this prompt:

```
Visually inspect these slides. Assume there are issues — find them.

Look for:
- Overlapping elements (text through shapes, lines through words, stacked elements)
- Text overflow or cut off at edges/box boundaries
- Decorative lines positioned for single-line text but title wrapped to two lines
- Source citations or footers colliding with content above
- Elements too close (< 0.3" gaps) or cards/sections nearly touching
- Uneven gaps (large empty area in one place, cramped in another)
- Insufficient margin from slide edges (< 0.5")
- Columns or similar elements not aligned consistently
- Low-contrast text (e.g., light gray text on cream-colored background)
- Low-contrast icons (e.g., dark icons on dark backgrounds without a contrasting circle)
- Text boxes too narrow causing excessive wrapping
- Leftover placeholder content

For each slide, list issues or areas of concern, even if minor.

Read and analyze these images:
1. /path/to/slide-01.jpg (Expected: [brief description])
2. /path/to/slide-02.jpg (Expected: [brief description])

Report ALL issues found, including minor ones.
```

### Verification Loop

1. Generate slides → Convert to images → Inspect
2. **List issues found** (if none found, look again more critically)
3. Fix issues
4. **Re-verify affected slides** — one fix often creates another problem
5. Repeat until a full pass reveals no new issues

**Do not declare success until you've completed at least one fix-and-verify cycle.**

---

## Converting to Images

```bash
python scripts/office/soffice.py --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf slide
```

Creates `slide-01.jpg`, `slide-02.jpg`, etc.

To re-render specific slides after fixes:

```bash
pdftoppm -jpeg -r 150 -f N -l N output.pdf slide-fixed
```

---

## Dependencies

- `pip install "markitdown[pptx]"` — text extraction
- `pip install Pillow` — thumbnail grids
- `npm install -g pptxgenjs` — creating from scratch
- LibreOffice (`soffice`) — PDF conversion (auto-configured via `scripts/office/soffice.py`)
- Poppler (`pdftoppm`) — PDF to images
