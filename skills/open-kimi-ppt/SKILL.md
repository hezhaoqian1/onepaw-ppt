---
name: open-kimi-ppt
description: Create, edit, replicate, read, and export presentations. For every PPT task, the default deliverables are BOTH (1) a self-contained PPTD project folder containing the .pptd manifest plus pages/media dependencies and (2) a locally generated .pptx that uses system fonts (so it looks the same in PowerPoint and WPS on any computer) with fade slide transitions. Use for any presentation, PowerPoint, PPT/PPTX, slide deck, PPTD, infographic, or poster task unless the user explicitly requests another format. Deliver with normal local file/folder links using absolute paths.
---

# Definition
open-kimi-ppt is a presentation creation and export skill built around Moonshot AI's PPTD format and browser-side PPTX writer. It defines a YAML-format intermediate DSL (`.pptd`) that abstracts OOXML and keeps each page self-contained.

**The default output is not PPTD-only.** Unless the user explicitly opts out, always produce both:

1. the complete editable PPTD project directory (`.pptd` + `pages/` + `media/` and other referenced dependencies);
2. the matching locally generated `.pptx`, with fade slide transitions applied by default and only fonts that every Windows/Office computer has (editor web fonts are replaced by system fonts on export).

Existing PPTX files may also be converted into PPTD for editing, after which both outputs are delivered again.

## The pptd format
The .pptd format is a simplified abstraction layer over OOXML that follows basic YAML syntax. This abstraction preserves the core content of OOXML (theme, page layout, element positions and definitions, etc.) while removing complex nesting logic such as Masters; every page is self-contained — what you see is what you get. Read reference/pptd.md for the complete definition of this DSL.

## PPT production workflow

### step0. Check local prerequisites
Before generating, run the environment check once per session (a few seconds, changes nothing):

```bash
python3 ~/.agents/skills/open-kimi-ppt/scripts/doctor.py
```

It checks Python, the Node runtime for PPTX export (an installed Node.js 18+, or the built-in runtime of an Electron desktop agent such as OnePaw, so Node.js does not have to be installed there), the headless Chrome/Edge used for image QA, and the bundled editor/WASM files, then exports a tiny test deck to prove the toolchain works.

1. `PPTX export: OK` -> continue normally. Do not run your own `node --version` / `npm` checks: Node may be off your shell's PATH even though the scripts can use it.
2. PPTX export failed -> tell the user exactly what doctor reports (usually: install Node.js 18+ from https://nodejs.org) and continue with PPTD-only output only if they agree.
3. Image QA failed (no Chrome or Edge) -> continue, use the structural review in step4 instead of image QA, and tell the user.

Nothing has to be installed with pip or npm: PyYAML is bundled, and the browser is driven directly (no agent-browser, Pillow or websocket-client). PPTX export, image QA and manual editing work offline; network is only needed for remote images referenced by the deck.

### step1. Read the context thoroughly
Read **all files uploaded by the user**, the provided URLs, and the pptd format guide `reference/pptd.md` to fully understand the user's requirements.

### step2. Understand the user's requirements
Understand the user's requirements based on the context:
1. First determine the purpose of the request
  - Create a PPT: create a new presentation (from scratch, or from an existing pptx template)
  - Edit a PPT: edit the user's uploaded PPT (local modifications, single-page beautification, etc.)
  - Replicate a PPT: replicate a presentation from a non-pptx format (images, PDF, etc.) into pptd format

2. Then determine the design direction
  - Self-directed design: no preference, or only simple style constraints given; you need to fill in or create the design
  - Design system: a preset design system from the skill (`reference/design_system/`) is specified, or the user provides a complete and detailed design scheme covering all color, font, layout, and component specifications
  - Use a template: a template is provided and must be used
  - Style transfer: a style reference source is provided (images, web pages, etc.)

3. Then determine the input type
  - Topic only: only a PPT topic direction or content requirements for the presentation are given, with no concrete content
  - Full document: the user provides a complete document (paper, research report, press release, etc.)
  - Outline: the user provides a page-by-page outline, speech script, or similar content
  * When the "user input type" is [Full document] or [Outline] and it is not specified whether expansion is allowed: since a page-by-page outline, speech script, or user document can hardly support the full content of a presentation, prefer using search to expand with more relevant material, cases, etc., unless the user explicitly says not to expand

4. Finally determine the exact page count
  - If the user requests a specific page count, the user's requirement takes priority
  - Page-by-page outline/script provided: match the number of pages in the outline/script
  - When a complete and relatively structured document is provided: ask the user how much document content one page should cover, and give an estimated total page count; when only a topic is provided: suggest a recommended page count and confirm with the user

#### Clarification and follow-up questions
When any of the following situations arise, resolve them by asking the user (use the agent's ask/clarification tool when available)
1. Requirements are ambiguous
- The user's intent is unclear or hard to understand
- The files/URLs provided by the user are inaccessible
2. Conflicting intents
- The user's intents contradict each other. For example:
  * A design system is selected while also requesting a style that is completely inconsistent with that design system (e.g., using a McKinsey style while requiring large areas of whitespace on pages) / using a template / referencing an image style
  * Requesting both "make 10 pages" and "deliver 30+ pages of output"
3. Unable to determine the user's requirements on your own
- When the purpose, design direction, input type, page count, etc. are hard to determine by yourself

### step3. Generate the presentation based on the user's requirements

Before generating, first read `reference/pptd.md` to understand the pptd format definition and constraints.

> **Fonts: use only fonts that every Windows/Office computer has.** The PPTX does not embed fonts, so the audience's computer shows it with its own fonts. Use `Microsoft YaHei` for Chinese and `Arial` (or `Segoe UI`) for Latin text by default; `DengXian`, `SimHei`, `SimSun`, `KaiTi`, `Georgia`, `Times New Roman`, `Impact` and `Consolas` are also safe (see `reference/fonts.md`). Editor web fonts such as MiSans, Liter or 思源宋体 are replaced by the closest system font in the PPTX and in the QA images, so prefer choosing the system font yourself.
>
> **Line breaks:** put each line in its own `<p>` paragraph (or its own line of a `|` block scalar). `<br/>` is converted into a real line break by the default exporter, but the `--browser` fallback drops it.

#### Replicating a PPT
- Analyze the images to estimate element positions, fonts and sizes, etc., and **replicate 1:1 as closely as possible**.
- For parts that are difficult to make out, use methods such as grid lines and close-up views to improve understanding.
- Replicate simple content in the image with elements; icons may be approximated with icons provided by Font Awesome. For content that cannot be approximated with icons or shapes, such as photos and avatars, use tools such as bash or python to crop and split the original image, then add the resulting image elements to the presentation

#### Editing a PPT
- Convert the user's uploaded pptx file to .pptd format
- Review the converted pages (structure and key visual details). Read a few key pages individually afterwards.
- Locate the pages to edit, and be careful not to affect parts outside the intended scope.
> Conversion from pptx to pptd is not perfectly lossless. If the user later reports format errors, garbled content, etc., compare against the original pptx and repair the pptd with reference to the comparison

#### Generating a PPT
When generating a PPT, adopt different production approaches for different user [design directions]
##### Self-directed design
1. Read the design guide `reference/slides_categories.md`, and read the scenario document corresponding to the user's query
2. Produce the presentation based on the above

#### Generating content in other formats
- When the user explicitly asks for an infographic, poster, or a highly visual single-page design, read `reference/general-poster.md` and implement it as a single-page or few-page editable PPTD; when the user only asks for an image, still build it with PPTD first, then output the image via screenshot or rendering. Do not load this reference file for ordinary PPT requests.

##### Design system
1. Read the general constraints section of the `reference/slides_categories.md` guide, and read the scenario document corresponding to the user's query as the design foundation
2. Read the specified design system as the presentation style: either the user-provided design scheme, or the matching preset under `reference/design_system/` (search by name / path the user specified; prefer the folder's `design.md` when present). It is strictly forbidden to reference or mix in other design styles
3. Produce the presentation with reference to the above
4. Do not auto-pick a preset during self-directed design; only use `reference/design_system/` when a preset is explicitly specified

##### Using a template
1. Convert the user's uploaded pptx file into pptd form
2. Review the converted pages to understand the template's visual style (color scheme, font style, element characteristics, layout characteristics, content density, etc.)
3. Identify page types; focus on reading special pages such as the cover, summary pages, and section dividers (single-page screenshots, .page files), extracting their page layouts, content structures, reusable components (icons, shapes, smartart, reusable body layout schemes, etc.), and element styles (e.g., whitespace/line/card separators, square/rounded corners, etc.)
4. Produce the presentation using the template

##### Style transfer
1. Analyze the reference file's visual style (color scheme, font style, element characteristics, layout characteristics, content density, etc.), page layouts, content structures, reusable components (icons, shapes, smartart, reusable body layout schemes, etc.), and element styles (e.g., whitespace/line/card separators, square/rounded corners, etc.).
- If the user provides a style reference URL, do not only read the text content; refer to and learn from the page's visual effect more to help understand the style
2. Produce the presentation using the reference file's style characteristics. You are encouraged to reuse illustrations, fonts, font-size hierarchies, elements, etc. from the original pdf/url

##### Images and Visual Materials
1. Images are an effective way to enrich a presentation's visual impact. Appropriate images should be used not only on covers and section dividers, but also on body pages to enrich the page, aid understanding, or support decision-making
2. Images are used to show concrete subjects, explain content, provide evidence, or establish a scene. Logos, icons, decorative textures, and very small thumbnails do not count as substantive imagery.
3. When a page involves products, people, places, buildings, events, cases, interfaces, experimental subjects, or spatial environments, prioritize corresponding real images or screenshots. If real images and screenshots cannot be obtained, generated images may be used instead.
4. Image priority: images provided by the user; images from official websites, official reports, and credible sources; searched images that are directly relevant to the content; images generated for conceptual expression or atmosphere.
5. After deciding which images are needed, complete image search, generation, and downloading in a batch before designing pages around their proportions. Save images in the `media` directory, keep them clear, and never stretch or distort them.
6. Analytical, technical, and academic PPTs should use corresponding evidence images when products, experiments, interfaces, cases, or on-site materials are available. Do not reduce every page to text, color blocks, and shapes.
7. Do not add irrelevant images merely to meet a quantity target. Every image must be directly relevant to the page's conclusion or communication goal.

##### Content Guidelines
1. Language style: unless the user explicitly requests otherwise, strictly avoid overly abstract expressions and uncommon metaphors
- Do not overuse metaphors, slogans, or abstract jargon such as distribution, an N-step argument, everything at a glance, a closed loop, hands-on practice, verification, deconstruction, second-class citizens, poison pills, or wall clocks
- Do not use common AI phrasing such as “not X, but Y,” “X is Y,” “why / based on what / how,” “key takeaway,” or “N battlefronts / paths”
- Do not use overly colloquial expressions such as “where should the ammunition go,” “the Nth thing,” “can't pick the right one,” or “cannot be used as X”

### step4. PPT validation
1. Validate the generated pptd against the format definition in `reference/pptd.md` (required fields, types, bounds, theme tokens, resource paths, etc.) and repair issues over multiple rounds
2. Visual review with exported page images — **required before PPTX export when the model supports image input (multimodal)**:
   - Run `scripts/export_images.py`. It loads the deck into the **local neo-ppt editor** (same offline mirror as `serve`) in a headless Chrome/Edge, chooses 导出 → 图片, unzips the images, and stitches all pages into one overview image. Text is rendered with the same system fonts as the exported PPTX. It **does not** call `www.kimi.com` and needs no extra installs. It usually takes 10–30 s; when your shell tool accepts a timeout, give image and PPTX export commands at least 300 s.

     ```bash
     python3 ~/.agents/skills/open-kimi-ppt/scripts/export_images.py \
       /abs/path/project/deck.pptd \
       --output /abs/path/project/.qa-images
     ```

     The script prints a JSON summary mapping each stitched label (`P1`…`Pn`, 1-based page order) to its `.page` file.
   - Read the stitched overview image (`.qa-images/overview.jpg`) and check every page against this list:
     1. 图片是否清晰、不变形（无拉伸、压缩、模糊）
     2. 文字是否压在关键画面（人脸、产品主体、Logo 等）上
     3. 元素坐标是否超出页面边界
     4. 边界与配色对比是否足够（文字与背景、相邻色块之间）
     5. 排版是否统一（对齐、间距、字号层级、页边距）
     6. 文字是否可能溢出文本框（文本过长、行距过密、字号过大）
     7. 内容是否被上层元素遮挡
   - For any suspicious page, read its full-resolution image (`.qa-images/pages/<n>.jpeg`) to confirm the problem before editing.
   - Fix issues in the corresponding `.page` file, then re-run `scripts/export_images.py --force` and review the new overview; repeat until every page passes.
   - Do not export the PPTX until the visual review passes. `.qa-images/` is an intermediate QA artifact and may be deleted after delivery.
3. When the model cannot read images, fall back to a structural review of the generated pages (bounds, overflow-prone long text, contrast, hierarchy, layout density) over multiple rounds, and state that image-based visual QA was skipped.

### step5. PPT output and delivery
1. Always produce a self-contained project directory. Keep the `.pptd` manifest and every referenced dependency together; never deliver a standalone manifest without its referenced files. Use this layout unless an existing project already has a valid equivalent structure:

   ```text
   deck/
     deck.pptd
     pages/
       *.page
     media/
       *                # when the deck has local media
     deck.pptx          # generated by default
   ```

2. Generate the `.pptx` by default after PPTD validation, even when the user only asks to create or edit a presentation. Skip PPTX export only when the user explicitly requests PPTD-only output or the environment cannot run the exporter; in the latter case, report the exact blocker and still deliver the complete PPTD project.
3. Deliver with normal clickable local links using absolute paths. In the final response, link all of the following:
   - the project directory;
   - the `.pptd` manifest;
   - the `pages/` directory and `media/` directory when present;
   - the generated `.pptx` file.
4. PPTX conversion: use `scripts/export_pptx.py`. **Default path is local patched official WASM** (`scripts/local-export/export-pptd.mjs --no-sign`): no cookie, no signature API, no browser UI. Optional `--browser` uses the **local** neo-ppt mirror (same as `serve`), still offline — not `www.kimi.com`.
5. Default PPTX options:
   - page transition: `fade` (淡入淡出), written to every slide after export;
   - fonts: not embedded (PowerPoint ignores the editor's embedded font files); editor web fonts are replaced by system fonts and listed under `fontsReplaced` in the JSON summary, and any other non-standard font is reported under `warnings`. `--keep-fonts` keeps the web fonts, only for audiences that have them installed;
   - line breaks: `<br/>` and newlines become real `<a:br/>` line breaks (`lineBreaks` in the summary);
   - override with `--transition none` or force browser UI with `--browser`.
6. Export command:

   ```bash
   # preferred: local patched WASM (offline)
   python3 ~/.agents/skills/open-kimi-ppt/scripts/export_pptx.py \
     /abs/path/project/deck.pptd \
     --output /abs/path/project/deck.pptx

   # optional: local neo-ppt browser UI (also offline)
   python3 ~/.agents/skills/open-kimi-ppt/scripts/export_pptx.py \
     /abs/path/project/deck.pptd \
     --output /abs/path/project/deck.pptx --browser
   ```

   A project directory may be passed instead of the manifest only when it contains exactly one `.pptd` file.
   Existing output files are not overwritten unless `--force` is passed.
7. Offline model (本地编辑器 + 本地导出):
   - **PPTX export (default)**: local patched WASM via `scripts/local-export/export-pptd.mjs --no-sign`. Canonical binary: package `editor/neo-ppt/assets/pptd_wasm_bg-DPPWdROu.wasm` (skill install copies it into `scripts/local-export/pptd_wasm_bg.wasm`). Requires **Node.js 18+**, or the Electron runtime of the desktop agent that runs the skill (found automatically).
   - **Image QA / visual review**: `scripts/export_images.py` drives the **local** neo-ppt editor in a headless Chrome or Edge over the DevTools protocol (`scripts/cdp.py`, standard library only; no `www.kimi.com`).
   - **Manual edit / preview**: `npx open-kimi-ppt-skill serve` serves the same offline neo-ppt mirror.
   - Browser PPTX path (`--browser`, also the automatic fallback when no Node runtime is found) uses the same local editor host and headless browser. PyYAML is bundled in `scripts/_vendor/`; nothing is installed at run time.
   - Local PNG/JPEG/GIF/SVG files inside the PPTD project are resolved by the local exporter / injected as data URLs for the editor host.
   - Do not claim PowerPoint/WPS/Keynote playback compatibility solely because ZIP validation succeeds.
8. After export, verify that the output exists and report the generated path. Confirm that every slide has exactly one root-level fade transition in valid CT_Slide order (`cSld`, optional `clrMapOvr`, `transition`, optional `timing/extLst`) and that the PPTX ZIP passes integrity checks. A byte-string search for `<p:fade>` is insufficient because Office ignores transitions nested inside `cSld`. Read the JSON summary: report `fontsReplaced` and any `warnings` to the user. For higher-risk decks, additionally inspect representative rendered/opened pages as appropriate.
9. When the user wants to open, edit, save, or export a PPTD project manually, start the local browser editor with `npx open-kimi-ppt-skill serve`. Ask the user to open `http://127.0.0.1:55173/` and authorize the complete PPTD project directory. Use a Chromium-based browser for writable access; folder-upload fallback is read-only. The host serves the **offline neo-ppt mirror** plus local patched WASM export — no `www.kimi.com`. Remote images or fonts referenced by the deck may still be fetched from their hosts.
10. After completing and delivering any presentation, always end the final response with a concise optional next step telling the user that they can run `npx open-kimi-ppt-skill serve` to view or edit the PPTD project, configure slide transition animations, and export PPTX manually. Keep this reminder in addition to, not instead of, the required project and file links.
11. Element animations (`page.animations` in PPTD — entrance / emphasis / exit / motion-path; see `reference/pptd.md` §6): use them only when the user explicitly requests animations, or when the deck is clearly intended for live presentation / slideshow playback and animation provides a clear benefit for staged disclosure, process demonstration, causal explanation, pacing, visual impact, or brand storytelling. By default, do not add element animations to reading-oriented, self-study, print, or primarily send-and-browse decks. Prefer 1–3 animation groups per page and simple effects such as fade, fly, and zoom. This is separate from the default PPTX slide-level fade page transition written by `export_pptx.py`.
12. Speaker notes (`notes` on each `.page`): use them only when the user explicitly requests them; otherwise, do not add them.
13. Parallel tool calls: during PPT production, make tool calls in parallel whenever possible; in each round, write multiple page files in parallel to reduce the number of steps.
