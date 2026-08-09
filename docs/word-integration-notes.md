# Word Integration — Limitations, Workarounds & Extra Effort

Living log for the team presentation. Updated as we build. Each entry: what
Word/Office.js/Tiptap does out of the box, what breaks, what we had to add,
and what it costs.

> **At-a-glance coverage matrix:** [`word-integration-coverage.html`](./word-integration-coverage.html)
> — a standalone, self-contained page (open it directly in a browser) summarizing
> what the OOXML→JSON import supports vs. the gaps, with feasibility. Handy for the deck.

---

## 1. Images: default HTML export is broken

**Limitation:** `body.getHtml()` does not embed images. It references them
via `~WRS{GUID}_files/imageNNN.ext` — a path that only exists inside Word's
internal export mechanism, not something our app can fetch.

**Workaround:** Separately read `body.inlinePictures`, pull each image out
as base64 via `getBase64ImageSrc()`, then rewrite the `<img src>` in the
captured HTML to a `data:image/...;base64,...` URI before handing it to the
editor.

**Extra code:** `inlineImageSources()` + capture/merge logic in
[`src/adapters/officeAdapter.ts`](../src/adapters/officeAdapter.ts) —
~25 lines, plus an extra `Word.run()` round trip on every save.

**Cost:** none to the user; transparent. Small perf cost (extra API round
trip per save, scales with image count).

---

## 2. Floating/anchored images: genuinely unsupported, no workaround exists

> **Scope note (added after §10/§14):** this finding is specifically about
> the live Word add-in path (`Word.Shape` via Office.js), which we've since
> retired as an authoring surface (§10). It does **not** describe the
> paste/import path's floating-image situation — see §14, where the same
> real-world problem turns out to be fixable once content only ever arrives
> via HTML rather than a live Office.js connection.

**Limitation:** Word has two kinds of pictures:
- **Inline** pictures → `Word.InlinePicture`, supports `getBase64ImageSrc()`.
- **Floating/anchored** pictures (text wrap = Square/Tight/Through/Behind/
  Front, incl. any image a user has dragged to "float" over text) →
  `Word.Shape`. **`Word.Shape` has no method to extract image bytes at all.**
  This was confirmed against `@types/office-js` — it is a hard gap in the
  Office.js API, not a bug in our code. There is no code-level fix.

**Product decision (mitigation, not a fix):** validate on save and warn the
user instead of silently losing the image. We detect floating pictures via
`context.document.body.shapes`, and use `shape.textWrap.type` to tell inline
from floating (note: `body.shapes` contains *all* picture shapes, inline
included — wrap type, not collection membership, is the actual signal; this
was itself a bug we shipped and fixed once, see commit history on
`officeAdapter.ts`). The user still can save; they just get told which
image(s) won't make it across, identified by alt text / shape name (Word's
JS API exposes no page number for a shape, so that's the best available
pointer), and how to fix it (Picture Format → Wrap Text → In Line with
Text).

**Extra code:** `getContentWarnings()` in
[`src/adapters/officeAdapter.ts`](../src/adapters/officeAdapter.ts), a new
`getContentWarnings()` method on the `DocumentAdapter` interface (no-op on
the web/Tiptap side), plus warning UI in
[`PublisherPanel.tsx`](../src/features/publisher-panel/PublisherPanel.tsx).
~40 lines total.

**Cost / requires from Word authors:** authors must know to use inline
images, or manually convert floating ones before/after our warning fires.
This is a genuine authoring constraint to flag to the team — not something
more engineering effort resolves.

**Platform requirement cost:** detecting this at all requires the
`WordApiDesktop 1.2` requirement set (see `manifest.xml`), which is
**desktop-only — no Word Online / browser support**. Authors using Word on
the web won't get this warning; their floating images will simply
disappear silently with no way for us to detect or warn about it there.

---

## 3. Heading color / font-size / font-family lost in the JSON representation

**Limitation:** Word's HTML export expresses *style-based* formatting (as
opposed to manual/direct formatting) via a `<style>` stylesheet block
(`h1 { color: ... }`, `.MsoNormal { ... }`, `mso-style-link`), not inline
`style=""` attributes. Tiptap's parser (`generateJSON`) only ever reads
`element.style.propertyName` directly off the element — it does not run a
CSS cascade, so stylesheet-only rules are invisible to it. Net effect: any
heading whose color/size/font came from Word's built-in "Heading 1" style
(the common case) silently lost that styling once converted to our JSON
representation, even though it displayed fine in the raw-HTML preview.

**Workaround:** two-part fix, both required:
1. Run captured HTML through `juice` (CSS-inlining library) before parsing,
   so cascade-resolved styles become inline `style=""` attributes Tiptap
   can see.
2. Extend Tiptap's `Color`, `FontSize`, `FontFamily` extensions' `types`
   option (which node/mark types the attribute applies to) to include
   `heading` and `paragraph` — by default these only apply to `textStyle`
   (i.e. manually-selected inline text formatting), not block-level nodes.

**Extra code:** new dependency (`juice`, npm), preprocessing step in
[`htmlJsonConversion.ts`](../src/core/tiptap-utils/htmlJsonConversion.ts),
extension config changes in
[`editorExtensions.ts`](../src/core/tiptap-utils/editorExtensions.ts).
~10 lines of app code + 1 new dependency.

**Body font-size dropped on the *OOXML* path — the same class of gap, a
different cascade (2026-08-09).** The paste path gets body size for free
(Word inlines the resolved `font-size:11pt` onto every span, juice-visible),
but the OOXML converter read size only from a run's own `<w:rPr><w:sz>` — and
body runs usually have no `rPr` at all. The 11pt lives in `docDefaults`
(`<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/>` — half-points, so
÷2 = 11pt), Word's document-wide default. So OOXML body text fell back to the
editor's 16px base and read visibly larger than both the source and the paste
import. Fix: [`styles.ts`](../src/core/tiptap-utils/ooxml/styles.ts) now
resolves the `docDefaults` size and the converter seeds each non-heading
paragraph's runs with it (a run's own `<w:sz>` still overrides), emitting it
as a `fontSize` on the run's `textStyle` mark. Headings are left unseeded so
they keep their heading-level + CSS sizing. Verified: sample-1 body runs →
`11pt`, headings → none. List items also hoist that size onto the `<li>` node
itself (not just their inner text), because the list marker (1., I., •)
inherits its size from the `<li>` — otherwise the number/bullet renders at
the editor's base size while the item text sits at 11pt, the same
marker-inheritance gap the paste path handles in wordListReconstruction. Not
handled: per-*paragraph-style* size overrides
(e.g. a custom "Caption" style at 9pt) — only run-level and docDefaults are
resolved, since these documents put body size in docDefaults; a styled
paragraph relying on its style's own `sz` would fall back to the default.

---

## 4. List indentation lost entirely — no Tiptap extension exists for it

**Limitation:** Word doesn't export numbered/bulleted lists as real
`<ol>/<ul>/<li>` — it exports plain `<p>` tags containing literal text
("1.", "·") plus `margin-left`/`text-indent` CSS to fake the hanging
indent. Tiptap has no built-in extension that captures `margin-left`/
`text-indent` at all (verified — no such attribute exists on the stock
paragraph/heading nodes).

**Original scope decision (superseded, see below):** we initially decided
**not** to attempt semantic list reconstruction (inferring real `<ol>/<li>`
structure from the fake numbering) — out of scope for the POC, visual
indentation only.

**Visual-only workaround:** wrote a custom Tiptap extension from scratch
(`Indent`) that adds `marginLeft`/`textIndent` as global attributes on
paragraph/heading nodes, parsing and re-rendering the inline CSS.

**Extra code:** new file
[`indentExtension.ts`](../src/core/tiptap-utils/indentExtension.ts),
~30 lines, entirely custom (no library covers this).

**Confirmed (2026-08-05):** checked all four of our own captured samples —
zero instances of real `<ol>/<li>`, and zero instances of Word's `mso-list`
CSS metadata (`mso-list: l0 level1 lfo1`) either, the semantic hint some
Word export paths *do* carry and which a reconstruction feature would
ideally key off instead of pure text-pattern matching. Not present via the
`body.getHtml()`/Office.js capture path we use.

**Reversed (2026-08-06) — semantic reconstruction built after all.** The
visual-only gap turned out to be more than a fidelity nitpick: pressing
Enter at the end of one of these fake list "items" didn't continue the
numbering/bullet, because there was no list node there to continue — just a
plain paragraph that happens to start with "1.". That's directly visible in
the "paste from Word, keep editing" pitch this project is now built around,
so it was worth the investment.

**How it works:** [`wordListReconstruction.ts`](../src/core/tiptap-utils/wordListReconstruction.ts)
walks the DOM after `juice()` has resolved styles, detects Word's
`MsoListParagraph*` paragraph shape (marker character/number + a run of
non-breaking-space padding simulating a tab, per Word's own convention),
and rebuilds real nested `<ol>`/`<ul>`/`<li>` from consecutive runs of them —
inferring nesting depth from Word's `mso-list: lN levelM lfoK` marker when
present (the authoritative level signal, 1-indexed — see the 2026-08-08 note
below), falling back to `margin-left` (Word indents ~0.5in per level) when
it isn't, and bullet-vs-ordered from the marker token itself (digit/letter/roman +
punctuation → ordered; a bare symbol, or `Symbol`/`Wingdings` font-family →
bullet). Inline formatting inside each item (bold, color, links, …) is
preserved via `Range.extractContents()` rather than rebuilt from scratch.
Verified against real captured markup from `sample-1.html`/`sample-4.html`
in an isolated jsdom test — correct nesting, correct ordered/bullet
detection, formatting preserved, non-list paragraphs left untouched.

**Extra code:** new file
[`wordListReconstruction.ts`](../src/core/tiptap-utils/wordListReconstruction.ts)
(~180 lines), plus a small shared
[`wordHtmlPreprocessing.ts`](../src/core/tiptap-utils/wordHtmlPreprocessing.ts)
so the same fix applies uniformly on save/load *and* on paste — no new
dependency.

**Still a heuristic, not a documented contract** — same caveat as the
video-embed and icon-MIME fixes elsewhere in this doc: Word gives no
semantic "this is a list" marker to key off (confirmed above), so this is
pattern-matching on export shape, not a guarantee. Known simplifications:
nesting depth is inferred from indent amount in fixed ~0.5in steps (an
unusual per-document indent scale could misjudge levels), and a marker-type
change at the *same* indent level (bullets switching to numbers without an
indent change) isn't specially handled — accepted as edge cases for a
first pass, consistent with "further iterations" being expected rather than
chasing full Word parity in one pass.

**Nesting collapsed on *paste* — `mso-list` level detection added
(2026-08-08).** The margin-left-only depth inference (above) worked on our
Save-As-HTML samples but flattened nested lists pasted from a live Word
**clipboard** copy — sub-items ("Simple Tables"/"Complex Tables") collapsed
into their parent's numbering. Root cause, reproduced end-to-end in an
isolated jsdom + full-schema test: a clipboard copy encodes nesting in the
`mso-list: lN levelM lfoK` marker (the same metadata the 2026-08-05 note
found *absent* from the `body.getHtml()` capture path) and frequently leaves
`margin-left` **uniform** across every level — so margin-based inference
reads it as flat. Save-As-HTML is the opposite: it strips `mso-list` and
varies `margin-left` per level. Fix: read `levelM` from the raw `style`
attribute (the CSSOM drops the unknown `mso-list` property, but
`getAttribute('style')` keeps it verbatim — the same technique CKEditor /
TinyMCE use), preferring it and falling back to `margin-left`. Both capture
paths now nest correctly; verified `juice()` preserves the property through
its round-trip. This narrows the earlier "unusual indent scale could misjudge
levels" caveat: it only applies to Save-As-HTML now, since the clipboard path
carries an exact level.

**Tab behavior on lists (2026-08-08).** Tiptap ships no indent extension, so
Tab handling is entirely ours ([`indentExtension.ts`](../src/core/tiptap-utils/indentExtension.ts)).
Two Word-parity gaps surfaced and were fixed: (1) Tab moved focus *out of the
editor* on a paragraph selection or a list's first item — the built-in
`sinkListItem` returns false when it can't nest (first item has no preceding
sibling), and that false let the bare Tab key hit the browser's focus
navigation. We now handle Tab ourselves and always swallow it in editable
content. (2) Selecting a whole list and pressing Tab did nothing (same
first-item `sinkListItem` refusal). Word shifts the whole list a level;
ProseMirror can't nest a first item, so we fall back to a block `margin-left`
shift on the list node itself (`bulletList`/`orderedList` added to the Indent
extension's types), matching how Tab moves a whole selected paragraph.
Collapsed cursor → first-line `text-indent`; range selection → whole-block
`margin-left`; list → nest, else block-shift.

**Numbering format was being dropped — now preserved per level
(2026-08-09).** The same nested list rendered *three different* ways: the
Word doc used upper-roman (I, II) at level 2, the editor showed lower-alpha
(a, b), and the Post Feed showed decimal (1, 2). Cause: both pipelines
resolved the real format but immediately collapsed it to an `ordered`
boolean and emitted a bare `<ol>`, so every surface fell back to CSS — and
the editor's list-style cycle (App.css) and the Post Feed's (which had no
nested rule at all) didn't even agree with each other. Fix: carry the real
format through as an inline `list-style-type` on the `<ol>` via a small
[`listStyleExtension.ts`](../src/core/tiptap-utils/listStyleExtension.ts)
(a `listStyleType` node attribute on `orderedList`). The OOXML path maps
numbering.xml's `numFmt` → CSS (`upperRoman`→`upper-roman`, `lowerLetter`→
`lower-alpha`, …); the paste path infers it from the marker token, which is
unambiguous because Word starts every ordered list at the format's first
value (1 / a / A / i / I). Both now render the document's own markers, and
the CSS cycles (editor + Post Feed, now matched) remain only as the fallback
for lists authored in the editor. Verified end-to-end: sample-2 paste →
nested `upper-roman`, sample-1 OOXML → level-1 `decimal`, level-2
`upper-roman`. Fixed a latent OOXML quirk along the way — a child list took
its *parent paragraph's* format; it now takes the format of the items it
actually contains.

---

## 5. Base64 images silently dropped by default

**Limitation:** `@tiptap/extension-image` defaults to `allowBase64: false`
— any base64-embedded `<img>` (which is exactly what our inline-image
workaround in §1 produces) is silently stripped during JSON parsing. Showed
up as: image visible in the raw-HTML preview, missing from the JSON-derived
render.

**Fix:** one-line config change, `Image.configure({ allowBase64: true })`.

**Cost:** none — just a config flag, but non-obvious/easy to miss (default
behavior fails silently, no console warning).

---

## 6. Article excerpt showing literal `&nbsp;`

**Limitation:** naive regex-based HTML tag stripping (`<[^>]+>`) doesn't
decode HTML entities, so excerpts generated this way showed literal
`&nbsp;` text instead of a space.

**Fix:** switched excerpt generation to walk the already-parsed Tiptap JSON
(`tiptapJsonToPlainText`) instead of regex-stripping raw HTML — entities
are already decoded to real characters by the time HTML becomes JSON.

**Cost:** none, small refactor.

---

## 7. Flagged-term highlighting: live on web, one-shot scan on Word

**What we built:** a hardcoded compliance-style term list
([`core/flaggedTerms.ts`](../src/core/flaggedTerms.ts): "guarantee",
"insider", "confidential", "stupid") flagged wherever it appears, with a
comment explaining why. Demonstrates that the panel isn't limited to plain
buttons/text — it can annotate content inline with real styling, in both
hosts, off a single shared rule list.

**Web (Tiptap):** a custom ProseMirror decoration plugin
([`flaggedTermsExtension.ts`](../src/core/tiptap-utils/flaggedTermsExtension.ts))
highlights matches live, on every keystroke, with a wavy red underline +
yellow background + a small ⚠ indicator carrying the comment as a tooltip.
Decorations are render-only — nothing is written into the document JSON, so
flags never get saved/exported. ~55 lines, no new dependency (`@tiptap/pm`,
ProseMirror's package, ships as part of Tiptap already).

**Word:** no decoration-layer concept exists in Word's object model, so
there's no way to get the same "live, non-destructive" behavior — any
highlighting or commenting is a real, saved edit to the document. Building
it as a live-as-you-type feature would mean re-running search + re-applying
highlight color + re-inserting comments on every keystroke, which would
spam duplicate comments and thrash formatting. Instead this is a one-shot
"Scan for flagged terms" button: it calls `body.search()` per term, sets
`range.font.highlightColor` (native Word highlight, `WordApi 1.1`), and adds
a native Word review comment via `range.insertComment()` (`WordApi 1.4` —
manifest bumped from 1.3 to 1.4 for this). ~20 lines in
[`officeAdapter.ts`](../src/adapters/officeAdapter.ts).

**Known gap:** clicking "Scan for flagged terms" more than once inserts
duplicate comments — there's no dedup check against existing comments. Fine
for a POC demo, would need addressing for real use.

**Cost / takeaway for the deck:** same rule list drives both hosts, but the
*mechanism* is fundamentally different per host — Tiptap gets a cheap,
non-destructive live overlay; Word gets a heavier, explicit, saved action
using its native highlight + comment features (which arguably reads more
"native" to a Word user, at the cost of not being live).

---

## 8. Video embeds: Word exports a linked thumbnail, not a video

**Limitation:** "Insert Online Video" in Word doesn't export a real embed at
all — no `<iframe>`, no video element. It exports a clickable static
thumbnail: `<a href="https://www.youtube.com/embed/{id}?feature=oembed">`
wrapping a base64 poster `<img>` (confirmed against a real captured
`sample-3.html`/`sample-3.json` pair). Before this fix, that pattern parsed
into a plain `image` node — the link, and the fact that it was ever a video,
were both silently dropped; only a static 200-300KB thumbnail survived.

**Workaround:** preprocess the captured HTML before parsing: detect the
`youtube(-nocookie)?.com/embed/{id}` link pattern and rewrite it to
`<div data-youtube-video><iframe src="https://www.youtube.com/embed/{id}">`
— the exact shape Tiptap's already-installed `Youtube` extension's
`parseHTML` matches (verified against its source). That produces a real
`youtube` node instead of a flattened image, which plays as a genuine iframe
both in the live editor and in Post Feed's News format (same JSON → HTML
render path, confirmed `dangerouslySetInnerHTML` doesn't strip iframes).

**Extra code:** new file
[`wordVideoEmbed.ts`](../src/core/tiptap-utils/wordVideoEmbed.ts)
(`resolveWordVideoEmbeds()`, ~30 lines), wired into
[`htmlJsonConversion.ts`](../src/core/tiptap-utils/htmlJsonConversion.ts)
ahead of the existing `juice()` step. No new dependency — reuses the
`Youtube` extension that was already in `EDITOR_EXTENSIONS` but unused.

**Known caveat:** there's no semantic "this is a video" marker anywhere in
Word's export — detection is a pattern-match on the embed link's `/embed/
{id}` shape, not a documented contract. Word generates that shape
deterministically for YouTube (it resolves the thumbnail via YouTube's
oEmbed API), so it's reliable in practice, but it is a heuristic, not an
official API guarantee. Scoped to YouTube only for now — same "visual/
functional fidelity over full generality" tradeoff as the list-indentation
scope decision (§4).

**Open question — enterprise/custom video providers (e.g. Brightcove):**
real deployments may embed video via a same-origin/custom player URL
instead of a YouTube link — inserted through Word's "From a Video Embed
Code" flow rather than oEmbed autodetect. The swap-the-link-for-an-iframe
*mechanism* should generalize (just point the regex/iframe `src` at the
provider's own URL shape instead of youtube.com), but it's **unverified**:
we don't yet know whether Word's export preserves a usable link/URL for
that flow the same way, or whether it rasterizes to a flat, unlinked image
with nothing recoverable. Needs a real captured sample to confirm before
building — blocked on internal network access, to be tested later.

**OOXML path (add-in) is strictly better than the paste heuristic
(2026-08-09).** On the `getOoxml()` route the video is *not* a guessed
pattern — Word stores it explicitly as `<wp15:webVideoPr embeddedHtml="…">`
inside the `<w:drawing>`, and `embeddedHtml` holds the **literal embed
`<iframe>`** (exact `src`, dimensions, sandbox attrs), alongside the poster
`<a:blip>` and a title. So there's a real "this is a video" marker and the
URL is authoritative, not inferred. [`convertDrawing`](../src/core/tiptap-utils/ooxml/ooxmlToTiptapJson.ts)
now checks for `webVideoPr` first, pulls `src` out of `embeddedHtml`, and
emits the *same* `youtube` node the paste path produces (so both render
identically) — verified against a real captured `sample-3.xml`
(→ `youtube` node, `embed/1V1yblnnLy8`, 480×360), and the node round-trips
through the editor schema to a live `<iframe>`. This also de-risks the
enterprise-player open question above on the OOXML route specifically:
`embeddedHtml` is a full iframe for *any* provider, so a non-YouTube embed is
recoverable there (we currently still scope the emitted node to YouTube and
fall back to the poster image otherwise — generalizing is now a
node-support question, not a capture one).

---

## 9. Icons render as broken images; Shapes can't be captured at all

Two more "Insert" gallery items tested against a real captured
`sample-4.html`, alongside the "Online Picture" case which already works
today with no changes needed (it's just a normal inline image).

**Icons — fixed.** Word's "Insert Icons" pictures came through as clean
inlined base64 (they're genuine `InlinePicture` objects, so the existing
capture pipeline already reached them) — but they rendered as broken images.
Root cause: `getBase64ImageSrc()` returns raw base64 with no format
indicator, so we always labeled it `image/png`. Decoding an icon's bytes
showed it's actually SVG markup (`<svg overflow="hidden" viewBox=...>`), and
a browser's `<img>` decoder trusts the data URI's *declared* MIME type over
the actual bytes — labeling SVG content as PNG means it fails to decode.

**Fix:** `isSvgBase64()` in
[`officeAdapter.ts`](../src/adapters/officeAdapter.ts) decodes a short
prefix of each image's base64 (just enough to see past any XML prologue,
cheap regardless of the full image size) and checks for an `<svg`/`<?xml`
opening tag, picking `image/svg+xml` instead of `image/png` when it matches.
~15 lines, no new dependency (`atob` is a browser/WebView2 built-in).

**Shapes — genuine caveat, not fixable, not a dealbreaker.** Word's HTML
export has zero vector/shape semantics: it fabricates an HTML
`<table cellpadding="0" cellspacing="0" align="left">` mockup, slicing the
shape into raster image fragments positioned by table cells to visually
approximate the original — no `<v:shape>`/VML, no shape-type hint of any
kind survives. Worse, even that raster fallback is incomplete: one of the
shape's own image fragments in the sample still has the broken, unfixable
`~WRS{GUID}_files/...` path (§2's limitation) — confirming it's a genuine
`Word.Shape`, not an `InlinePicture`, so there's no byte-extraction path
even in principle, not just a parsing gap.

**Workarounds considered:**
- **Author-side (recommended, no code):** ask authors to right-click the
  shape → *Save as Picture*, then re-insert it as a picture. That converts
  it to a true `InlinePicture`, which the existing pipeline already handles
  fully. Cheapest fix by far — it's a one-time authoring habit, not
  engineering work.
- **Detection/warning (not implemented):** we can't detect "this table is
  secretly a shape" from the HTML — the fallback table has no signature
  distinguishing it from a real content table someone typed. But we don't
  need HTML for that: `context.document.body.shapes` (already loaded by the
  floating-image warning in §2) returns every shape regardless of type.
  Extending the existing `shape.type === Word.ShapeType.picture` filter to
  also include `geometricShape`/`group`/`textBox`/`canvas` would let
  `getContentWarnings()` flag shapes the same way it already flags floating
  images — a few lines, reusing code that's already there. Worth doing if
  shapes turn out to be common in real documents; skipped for now since it's
  a caveat, not a blocker.
- **Reconstruction (ruled out):** no path exists to recover real shape
  geometry/fill/type from the HTML export — same category decision as
  ruling out semantic list reconstruction (§4).

---

## 10. Direction change: paste-from-Word replaces the live add-in

**Decision:** we're dropping the live Word task-pane add-in (`Word.run()`,
the manifest, `WordApi`/`WordApiDesktop` requirement sets) as the authoring
surface. Every unfixable limitation in this document (§2 floating images,
§9b shapes) — and every workaround that needed live Office.js access
(§7 flagged-term scanning) — exists specifically because the add-in edits an
*open* Word document in real time. None of that risk applies to a one-time
import: someone pastes Word content into Tiptap once, then edits entirely in
Tiptap from then on. Tiptap becomes the only editor; Word becomes an input
format, not a live-connected host.

**Gap this exposed:** the HTML→JSON preprocessing built for the add-in
(§3 juice for heading styles, §8 video embeds) was only ever wired into
`htmlToTiptapJson()` — the explicit "Save article as HTML/JSON" and
"load a saved article" paths. Pasting Word content directly into the live
Tiptap editor bypassed all of it: Tiptap's default paste handling parses
clipboard HTML straight into the schema with no preprocessing, so pasted
headings would lose their color/font and pasted videos would flatten to a
plain image, exactly like the add-in did before those fixes.

**Fix:** a new extension,
[`wordPasteExtension.ts`](../src/core/tiptap-utils/wordPasteExtension.ts),
using Tiptap's `transformPastedHTML` hook (any extension can define one;
Tiptap chains them before parsing pasted HTML into the document). It runs
the same `resolveWordVideoEmbeds()` + `juice()` steps already used for
saved articles, so paste and import are now on equal footing. ~10 lines, no
new dependency. Everything schema-level (indentation, flagged-term
highlighting, base64 images) already applied to pasted content automatically
— those aren't tied to `htmlToTiptapJson()` specifically, they're just part
of the editor's schema, active regardless of how content enters it.

**Known gap — genuine non-picture Shapes (and comment anchors) can silently
kill an entire paste — confirmed, unfixable.** Tested live: copying a
selection that included a *drawn* Shape (an arrow — a real vector shape, not
a picture) and a comment anchor caused the *whole* paste to lose all
formatting — no headings, no bold/italic, no table structure, nothing —
with zero console errors. Traced the cause in `prosemirror-view`'s own
source (`parseFromClipboard`, `dist/index.js`): it only calls our
`transformPastedHTML` hook when the clipboard actually has an `html` string;
if `html` is empty, it takes the plain-text path entirely instead
(`asText = !!text && (plainText || inCode || !html)`). So the browser
received *no* `text/html` at all for that paste — Word failed to populate it.

**Correction, since this was originally stated too broadly:** this was
first attributed to "floating shapes" generally, by analogy to
officeAdapter's `Word.Shape` gap (§2/§9b) — but that conflated genuine
drawn Shapes with floating *pictures*, which are a different thing. §14
later tested floating pictures (including an embedded chart) specifically
on the clipboard path and found they populate `text/html` completely fine
— align info, real image bytes, all of it. So this gap is narrower than
first described: it's specific to actual vector Shapes (arrows, text boxes,
WordArt) and comment anchors, not floating/anchored pictures in general.
Still no code-level fix possible for the narrower case — if the clipboard
never had HTML, there's nothing to transform. Practical mitigation for
authors: paste in smaller sections, and avoid including drawn shapes or
comments in a copied selection when formatting matters.

**Correction — "images on paste" was never actually tested when this was
first written, and turned out not to be a gap at all.** Originally
speculated here as "likely unfixable" by analogy to Save-As-HTML's `file://`
references (§1) — but that assumption was never verified against a real
paste, and it was wrong. Since verified extensively (§14, §15): Word's
*clipboard* HTML embeds images as base64 directly (unlike Save-As-HTML,
which externalizes them to a `_files` folder) — they come through
`allowBase64: true` with no extra work needed. Leaving this note in place
as a reminder of the broader pattern: Save-As-HTML and live clipboard paste
are two different Word export paths that can behave differently, so an
assumption verified against one isn't automatically safe to apply to the
other — several fixes in this doc (§14 especially) exist because that
distinction wasn't always obvious upfront.

---

## 11. Centered/right-aligned standalone images lost their alignment

**Limitation:** Word expresses a centered or right-aligned standalone image
by putting `align="center"`/`style="text-align:center"` on the *paragraph*
wrapping the image, not on the `<img>` itself. Our `Image` node is
block-level (`inline: false`, the extension's default — we never override
it), so it can't legally sit inside a `paragraph` node (a paragraph's
content model only allows inline content). Tiptap's parser hoists the image
out as its own sibling block, and the wrapping paragraph — the only thing
carrying the alignment — has nowhere to survive. Net effect: an image the
author deliberately centered in Word always renders left-aligned once
converted, with the alignment silently dropped, not just visually
approximated.

**Fix:** two-part, same pattern as most fixes in this doc:
1. [`wordImageAlignment.ts`](../src/core/tiptap-utils/wordImageAlignment.ts)
   — preprocessing step that finds paragraphs whose only real content is a
   single image, and transfers the paragraph's alignment onto the `<img>`
   itself (as a `data-align` attribute) before parsing, since the paragraph
   won't survive to carry it.
2. [`imageAlignmentExtension.ts`](../src/core/tiptap-utils/imageAlignmentExtension.ts)
   — gives the `Image` node its own `align` attribute (reads/writes
   `data-align`; see §13c for why this renders as a bare attribute rather
   than an inline `style`, and where the actual CSS lives).

**Extra code:** ~75 lines total across both files, no new dependency.
Verified against real captured markup from `sample-2.html` (the "Web Access
Symbol" image) in an isolated jsdom test — alignment correctly transfers
onto the image; a plain paragraph of body text sitting next to it is
correctly left untouched.

**Scope:** deliberately narrow — only fires when the image is the
paragraph's *only* meaningful content (Word's own convention for a
centered/right-aligned standalone image). An image sitting inline alongside
real text is left alone on purpose: "align this whole paragraph" and "align
this one inline image within a paragraph of text" aren't the same thing,
and Word doesn't give us enough signal to safely guess the latter.

---

## 12. Editing-time image gaps: no align control, no room to type next to one

**Limitation:** §11 fixed alignment surviving *import* — but there was no
way to *set* it while actively editing, and separately, block-level images
(ours are, by the `@tiptap/extension-image` default) have no text content
of their own. If there's no paragraph immediately before/after an image,
there's nowhere for a typed character — including a plain space — to go.
Tiptap's bundled `Gapcursor` (on by default via StarterKit) lets you
*navigate* a cursor next to an atom node like an image, but a gap-cursor
position isn't inside any text-containing node, so typing does nothing
there. Word-imported images routinely land with no such gap on either side.

**Fix, two parts:**
1. [`imageAlignmentExtension.ts`](../src/core/tiptap-utils/imageAlignmentExtension.ts)
   gained a `setImageAlign` command (§11 only had the attribute/rendering
   side). The toolbar's existing Align Left/Center/Right buttons now double
   as image alignment: they check `editor.isActive('image')` and call
   `setImageAlign` instead of `setTextAlign` when an image is selected — no
   separate UI needed, since text-align and image-align are mutually
   exclusive by definition (an image doesn't have the paragraph a text
   alignment would apply to).
2. New [`imageSpacingExtension.ts`](../src/core/tiptap-utils/imageSpacingExtension.ts)
   — a ProseMirror `appendTransaction` plugin that runs on every document
   change: any top-level image missing a text-block sibling on either side
   gets an empty paragraph inserted there automatically, so a place to type
   always exists.

**Extra code:** ~70 lines across both files, no new dependency (Gapcursor
was already active; this builds on top of it rather than needing it
configured differently).

**Scope:** `imageSpacingExtension.ts` only inspects *top-level* document
children — the case that actually occurs (Word's standalone images are
direct body children). An image nested inside a list item or table cell
wouldn't get this treatment, but that's not a shape Word's export produces
for a standalone image.

---

## 13. Image resize, and a resize-vs-alignment CSS conflict it created

**Feature:** users can now drag-resize an inline image directly in the
editor. This is a built-in `@tiptap/extension-image` capability
(`resize: {...}`), just off by default — no third-party package or custom
node view needed:

```ts
Image.configure({
  allowBase64: true,
  resize: {
    enabled: true,
    directions: ['top-left', 'top-right', 'bottom-left', 'bottom-right'],
    minWidth: 40,
    minHeight: 40,
    alwaysPreserveAspectRatio: true,
  },
}),
```

**Gap found immediately after enabling it:** the resize handles were
present in the DOM (confirmed by reading `@tiptap/core`'s actual
`ResizableNodeView` source) but invisible and unclickable — Tiptap wires up
the drag behavior fully but ships the handles as bare, 0-size, unstyled
`<div>`s by design, expecting the consumer to style them. The simplified
`Image.configure({ resize: ... })` surface doesn't even expose a
`className`/`createCustomHandle` passthrough for this — but every handle
still gets an unconditional `data-resize-handle="<direction>"` attribute
(plus `data-resize-wrapper`/`data-resize-container` on its ancestors)
regardless, which is what the fix targets. Pure CSS in
[`App.css`](../src/App.css) — round red dot, hidden until the wrapper is
hovered, direction-aware resize cursor — no source/config change needed.

**Second gap, found right after that:** enabling resize silently broke §12's
image-alignment buttons — clicking Left/Center/Right on a selected image did
nothing. Root cause: our alignment CSS applies `margin-left/right: auto` to
the `<img>` itself, which only shifts the image if its containing block is
*wider* than the image. That's true for static rendering (Post Feed, raw
HTML export), but `ResizableNodeView`'s wrapper (`[data-resize-wrapper]`) is
a flex item that shrinks to exactly the image's own size once resize is on —
zero extra width, so the auto-margin trick has nothing to distribute and all
three alignments render identically.

**Fix:** move alignment to the *outer* `[data-resize-container]` element's
`justify-content` instead, forcing it to `width: 100%` and using CSS
`:has()` to key off the image's own `data-align` attribute — deliberately
left the `[data-resize-wrapper]` itself untouched, since widening it too
would throw off where the resize handles anchor (they position relative to
the wrapper's own edges, not the container's).

```css
[data-resize-container] {
  width: 100%;
}

[data-resize-container]:has(img[data-align='center']) {
  justify-content: center;
}

[data-resize-container]:has(img[data-align='right']) {
  justify-content: flex-end;
}

[data-resize-container]:has(img[data-align='left']) {
  justify-content: flex-start;
}
```

**Extra code:** ~50 lines of CSS total (handle styling + this), no new
dependency, no change to `imageAlignmentExtension.ts`/`setImageAlign` itself
— the attribute and command from §12 were already correct, only the
*rendering* of that attribute needed a second rule once resize was in play.

**Scope/caveat:** this is a good concrete example of two independently-fine
features interacting in a non-obvious way once combined — worth calling out
in the deck as a real "gotcha" we hit and root-caused via reading Tiptap's
actual compiled source rather than guessing. `:has()` requires a reasonably
modern browser (Chrome/Edge 105+, Safari 15.4+, Firefox 121+) — fine for an
internal tool, worth a footnote if the audience asks about browser support.

**13c — a third bug in the same chain: resizing an aligned image snapped
it back to full size the instant the handle was released.** Drag itself
worked (visibly grows/shrinks) but the moment `mouseup` fired, the image
reverted to its natural size. Root cause, confirmed by reading
`@tiptap/extension-image`'s actual compiled source: on every attribute
update, its resize node view resyncs the `<img>` element's HTML attributes
from scratch — and for the `style` attribute specifically, that's a
wholesale `el.setAttribute('style', ...)`, not a merge. `width`/`height`
are special-cased and skipped in that resync (they're applied imperatively
via `el.style.width/height` during drag, deliberately kept out of the
attribute system), but nothing else is. Since §11/§12's alignment fix
rendered `align` as an inline `style` (`margin-left/right: auto`), *any*
node attribute update — including the drag's own `onCommit`, which persists
the final width/height as node attrs to make the resize stick — triggered
that resync and clobbered the very `style` attribute holding the
just-imperatively-set width/height. Resize and alignment weren't just
visually conflicting (§13b) — combined, they were mutually destructive.

**Fix:** [`imageAlignmentExtension.ts`](../src/core/tiptap-utils/imageAlignmentExtension.ts)
no longer renders a `style` attribute at all — only `data-align`. Since that
key isn't a wholesale-overwrite target the way `style` is, it survives the
resize resync untouched, and so does whatever `el.style.width/height` the
drag set. The actual centering CSS moved out of the attribute system
entirely, into two stylesheets covering the two shapes it needs to handle:
[`App.css`](../src/App.css)'s `[data-resize-container]:has(...)`
(§13b — the live editor, always resize-wrapped) and a new
[`PostFeedTab.css`](../src/features/post-feed/PostFeedTab.css) rule,
`.feed-post-body img[data-align=...]`, for the plain exported-HTML case
(News format) where there's no resize wrapper and ordinary margin:auto on
the `<img>` itself works fine.

**Extra code:** ~25 lines of CSS moved/added, ~15 lines removed from
`imageAlignmentExtension.ts` (net roughly neutral). No new dependency.

**Scope/caveat:** any *other* consumer of `tiptapJsonToHtml`'s raw output
(if one gets added later, beyond News format) will need this same CSS rule
included to keep image alignment visible — it's no longer self-contained
in the HTML string the way the old inline `style` was. Worth a line in the
deck: portability was traded for correctness here, and it's a trade an
inline-style approach can't avoid once resize is in the picture at all —
this isn't specific to our extension, it's how Tiptap's own resize node
view resyncs attributes.

---

## 14. Floating/wrapped images: revisited for the paste/import path — turns out fixable

**Background:** §2 concluded floating images were a hard, unfixable gap —
but that finding was about `Word.Shape` via the live Office.js add-in,
which no longer exists as an authoring path (§10). Once everything only
ever arrives as HTML (paste or a saved file), the question is different:
does Word's HTML actually carry the wrap/position info at all? Answered by
inspecting a real "Save As Web Page" export containing two genuinely
floating images (`samples/saved-from-word/Sample-floating.htm`) — one a
native embedded Excel chart (`Insert → Chart`, wrap = Square, sharing its
paragraph with body text), one a plain picture with wrap = Square as its
paragraph's only content.

**What the real markup showed:** Word wraps every floating image in a VML
block (`<v:shape>…</v:shape>`, plus an `<o:OLEObject>` for the chart case)
guarded by IE-only conditional-comment syntax (`<!--[if gte vml 1]>…
<![endif]-->`). Any non-IE parser — every real browser today, and our own
DOM-based preprocessing — treats that whole block as one long HTML comment
and never sees it at all. What *does* survive, wrapped in a
"downlevel-revealed" conditional comment (`<![if !vml]>…<![endif]>`, no
leading `--`, parsed as two separate bogus-comment nodes with a live
element in between) is a plain fallback `<img>` — critically, carrying a
legacy `align="left"`/`"right"` attribute, the same attribute a browser's
own default stylesheet maps straight to `float: left`/`float: right`. So
the wrap direction isn't lost in Word's export at all; it's sitting in
plain sight on the fallback image, just not read into our schema before
now.

The chart case additionally confirms something worth its own callout:
**Word does export a raster snapshot for native embedded charts** (a
`.gif` here) alongside the VML, going through the exact same `<img>`
fallback path as a real picture — meaning charts are a §1-style fix
(normalize/inline the src), not a §2-style dead end. Directly relevant
since Tiptap Pro's own [feature support matrix](https://tiptap.dev/docs/conversion/getting-started/feature-support-matrix)
lists "Shapes, SmartArt, WordArt" — charts included — as **not converted,
not extracted at all** during DOCX import. Their native OOXML parser drops
charts outright; our HTML-based approach recovers at least the visual
snapshot.

**Fix:** [`wordImageAlignment.ts`](../src/core/tiptap-utils/wordImageAlignment.ts)
extended (not replaced — §11's non-wrapping paragraph-alignment case is
still handled by the same file) to also read the fallback `<img>`'s own
`align` attribute, writing `data-align="float-left"`/`"float-right"` (a
value distinct from §11's plain `left`/`center`/`right`, since one means
"position" and the other means "wrap text around this"). Our `Image` node
is block-level and can't sit as literal inline content next to a
paragraph's text the way Word's DOM does, so when the image shares its
paragraph with other real content, it gets hoisted out to a sibling
positioned immediately before that paragraph — the text then becomes an
ordinary sibling paragraph, and CSS float still pulls it around the
hoisted image on render, since float affects subsequent flow content
regardless of "logical" paragraph ownership. (Comment nodes — the VML
block and the two bogus-comment markers — are deliberately excluded from
the "does this paragraph have other content" check; `Node.textContent` on
a Comment still returns its full text, which would otherwise make every
floating image look non-empty even when it's genuinely alone in its
paragraph.) `imageAlignmentExtension.ts` needed no changes — it already
renders whatever string `data-align` holds generically. New CSS in
[`App.css`](../src/App.css) (`float` on `[data-resize-container]`, width
reset to `auto` since a 100%-wide floated box leaves nothing to wrap into)
and [`PostFeedTab.css`](../src/features/post-feed/PostFeedTab.css) (plain
`float` on the `<img>`, no resize wrapper to work around there).

**Verified** against the real captured markup in an isolated jsdom test —
both the shared-paragraph chart and the alone-in-its-paragraph picture
resolve correctly (hoist only happens for the former; both end up with the
right `data-align`), before wiring into the app. Typecheck and dev-server
checks clean.

**Extra code:** ~45 lines net across `wordImageAlignment.ts` (rewritten
with the same file, not a new one) and CSS. No new dependency.

**Update — live-paste caveat closed.** This was originally verified only
against a "Save As Web Page" export, not an actual clipboard copy-paste,
with an open question about whether the two paths agree. Since confirmed
live: pasting a real floating/wrapped image directly into the editor
produces correctly-floated output, same as the Save-As-HTML test predicted.
The VML/downlevel-comment convention this fix relies on is shared by both
Word export paths.

---

## 15. Pasted image quality: Word's HTML gives a legacy low-res fallback; RTF has the real thing

**Limitation, found while checking §14's chart example up close:** the
chart image floated correctly after §14, but looked visibly blurry.
Decoded the actual embedded file directly (its base64 payload's own magic
bytes, independent of whatever the browser's data URI claimed): it's a
`GIF89a`, natively **430×216px**, 256-color-palette — a decades-old,
low-fidelity fallback format. Worse, our `Image` node was displaying it
*larger* than that (rendering it at whatever taller `height` Word's
fallback `<img>` tag declared, reflecting how big the chart looked in the
original document layout, not the asset's real resolution) — compounding
an already-soft source with unnecessary upscaling. Root cause: Word's
`<v:imagedata>` inside the VML block references a real vector EMF, but (per
§14) no modern browser or clipboard consumer ever sees that block at all —
every fallback path, on both Save-As-HTML and live clipboard paste,
downgrades to this same old GIF. Cross-checked against TinyMCE's own
PowerPaste docs — they describe the identical ceiling ("image quality
depends on what the source application provides in its clipboard
representation") and specifically flag charts as a weak spot, so this
isn't something our HTML-reading approach was doing wrong relative to a
commercial equivalent.

**The fix — read `text/rtf` too, not just `text/html`:** Word's clipboard
populates multiple independent formats simultaneously (confirmed directly
via a `clipboardData.items` listener: `text/plain`, `text/html`,
`text/rtf` all present at once for the same copy). `text/rtf` represents
an embedded picture completely differently — inline as a `{\pict ...}`
group containing the picture's *actual file bytes*, spelled out as hex
digits, tagged with its real format (`\pngblip`/`\jpegblip`). Decoding that
hex directly for the same chart produced a **102KB JPEG** — genuinely valid
(verified its magic bytes: `FF D8 FF E0`), enormously higher-fidelity than
the ~460-byte GIF the HTML path offered. This is the same class of fix
CKEditor ships for the same reason
([ckeditor5#18363](https://github.com/ckeditor/ckeditor5/pull/18363)).

**Implementation:**
1. [`wordRtfImage.ts`](../src/core/tiptap-utils/wordRtfImage.ts) — a small
   RTF `\pict` parser. Skips past a tag's dimension/cache-id control words
   (including a nested `{\*\blipuid ...}` destination group, confirmed
   necessary against real captured RTF — a naive "skip words then read hex"
   scan undercounts without it), reads the contiguous hex run that follows,
   and hex-decodes it back into real image bytes. Only `\pngblip`/
   `\jpegblip` are extracted — `\wmetafile`/`\emfblip` are vector formats
   with no browser-native renderer.
2. [`wordPasteExtension.ts`](../src/core/tiptap-utils/wordPasteExtension.ts)
   gained a second ProseMirror plugin, `handlePaste`. Confirmed via reading
   `prosemirror-view`'s own `doPaste` source that `handlePaste` runs
   *after* `transformPastedHTML` has already parsed the HTML into a slice —
   so every other Word fix (lists, alignment, spacing) has already applied
   by the time this runs; it only ever touches `src`. When it finds usable
   RTF pictures, it rebuilds the pasted slice with the image's `src`
   swapped to the best RTF-decoded picture (as a base64 data URL — kept
   portable through localStorage save and Post Feed rendering, deliberately
   not a `blob:` URL, which wouldn't survive a reload).
3. **Multi-image correlation:** RTF has no shared ID linking a `\pict`
   group back to a specific HTML `<img>`, so with more than one image in a
   paste, matching relies on two heuristics confirmed against real
   captures: (a) real full-fidelity renditions run far bigger than the
   small cached icon/thumbnail pictures Word embeds alongside them (a
   ~277-byte "icon" showed up consistently regardless of what was actually
   copied), so the *N* largest RTF pictures are taken as the real
   candidates; (b) both clipboard formats serialize the same selection in
   the same reading order, so those survivors are re-sorted back into
   document order and paired 1:1, in order, with the image nodes in the
   pasted slice (also walked in document order). Verified standalone
   against a synthetic interleaved-candidates scenario before wiring in.
   Capped at 5 images per paste (`MAX_IMAGES_TO_UPGRADE`) — a safety valve,
   not a hard technical limit, since more images means more that could go
   wrong in that correlation. If there are more images than usable RTF
   pictures, or the image count exceeds the cap, every image is left
   exactly as the HTML fragment already resolved it — this only ever makes
   quality *better*, never worse, by design (falls through to the
   unmodified paste on any ambiguity or error).

**Extra code:** ~90 lines net (`wordRtfImage.ts` new; `wordPasteExtension.ts`
extended) — no new dependency.

**Verified:** RTF hex-parsing logic tested against real captured payload
fragments (magic bytes, byte counts, and exact trailing bytes all
confirmed correct via an isolated Node script) before wiring in. End-to-end
tested live: single-image paste (confirmed visually sharper — "yeah its
crisp") and multi-image paste (confirmed the safety gate correctly declines
when correlation would be ambiguous, e.g. more images than confidently
matchable RTF pictures).

**Caveat:** the size-based correlation is a heuristic, not a structural
guarantee of the RTF format — it assumes each real image contributes one
dominant large-byte picture, which held in every capture tested here but
isn't something the format enforces. Fine for this POC's use (a handful of
images per article); would need a sturdier correlation approach (e.g.
matching by shape ID within the RTF's own object structure, not yet
attempted) before relying on it for arbitrary documents at scale.

---

## 16. OOXML fidelity pass — style resolution, character/paragraph formatting, tables, TOC (2026-08-09)

Stress-tested the OOXML→JSON path with a document exercising most of Word's
formatting surface. Seven gaps surfaced; six were fixed (the seventh, truly
inline mid-sentence images, is left as a known limitation — our `image` node
is block-level, so an inline image is hoisted out of its sentence; fixing it
needs an inline-image schema node).

**The one fundamental loophole behind most of them: we weren't resolving
styles.** The converter only pulled *heading level* and *docDefaults font
size* out of `styles.xml`. Everything else a paragraph/character **style**
contributes — numbering, indentation, spacing, italics, color, size — was
invisible unless written directly on the run/paragraph. Real Word documents
lean heavily on styles, so this alone broke lists, the built-in paragraph
styles, and most indentation/spacing.

**Fix — a proper style-props resolver** ([`styles.ts`](../src/core/tiptap-utils/ooxml/styles.ts)):
`parseStyles` now reads each style's own `pPr`/`rPr` plus `docDefaults`, and
`styleProps(styleId)` returns the *effective* paragraph + run properties
merged along the `basedOn` chain over docDefaults. The converter computes each
paragraph's effective props (`style ∪ direct`) and each run's
(`style ∪ paragraph-mark ∪ direct`), so style-only formatting finally reaches
the output. This is the backbone the rest build on.

What that unlocked, verified end-to-end (convert → real editor schema →
serialize):

- **Lists (issue 5)** — a `ListBullet`/`ListNumber` item carries its `<w:numPr>`
  on the *style*, not the paragraph, so membership was missed and items fell
  through to plain paragraphs. Membership is now read from effective para
  props, so style-based lists reconstruct correctly.
- **Built-in styles (issue 2)** — Quote / Intense Quote / Caption / No Spacing
  now render their style's italics, color, size, indent, alignment and
  spacing (e.g. Intense Quote → centered, indented both sides, blue italic).
- **Indentation + spacing (issue 4)** — `<w:ind>` → `marginLeft` / `marginRight`
  / `textIndent` (first-line positive, hanging negative), `<w:spacing>` →
  `line-height` + top/bottom margins. Spacing is emitted **only when it
  differs from docDefaults**, so it doesn't fight the CSS baseline or
  re-loosen tight list items. Added `marginRight` to the Indent extension and
  `marginTop`/`marginBottom` to ParagraphStyle.
- **Character formatting (issue 3)** — `<w:highlight>` was never read at all
  (now mapped, named color → CSS); underline collapsed to a boolean (now
  keeps `dotted`/`wavy`/`double` via `text-decoration-style`); `<w:dstrike>`
  (double strike) and `<w:smallCaps>` (font-variant) added. New global-
  attribute extensions in [`ooxmlFormattingExtensions.ts`](../src/core/tiptap-utils/ooxmlFormattingExtensions.ts)
  give the underline/strike/textStyle marks somewhere to carry these.
- **Table shading (issue 6)** — `<w:shd w:fill>` was dropped, which not only
  lost banner/cell colors but hid **white-on-fill header text** (white text on
  the default white cell → invisible). Added a `backgroundColor` attribute to
  table cells and read the fill into it.
- **TOC (issue 1)** — a dedicated `tocEntry` node ([`tocEntryExtension.ts`](../src/core/tiptap-utils/tocEntryExtension.ts))
  lays out entry text, a dotted leader (a CSS dotted border, *not* real tab
  stops — appearance only, per the ask), and the right-aligned page number
  read from the entry's trailing numeric run. The TOC field itself is not
  reproduced (no live page recalculation).

**Known limitation deliberately left:** inline mid-sentence images (block-only
image node), Word field semantics beyond TOC layout (`PAGEREF`, cross-refs),
per-side paragraph borders (Intense Quote's top/bottom rules render as no box
rather than an approximated one), and `<w:caps>` all-caps.

**Title/Subtitle styles — lost on the paste path when the `<style>` block
didn't survive the clipboard (2026-08-09).** Unlike Heading 1–6 (which Word
exports as real `<h1>`–`<h6>`), the Title and Subtitle styles export as
`<p class=MsoTitle>` with their size living *only* in the stylesheet's
`p.MsoTitle { font-size: 28pt }` rule. That size reaches the output solely via
juice inlining that rule — so a pasted payload without (or with a stripped)
`<style>` block loses it and the document title collapses to body size
(confirmed: stripping the block drops the inlined size to `""`). The class on
the element always survives, so [`wordTitleReconstruction.ts`](../src/core/tiptap-utils/wordTitleReconstruction.ts)
promotes `MsoTitle`→`<h1>` / `MsoSubtitle`→`<h2>` by class (after juice, so an
inlined size still wins where present). The OOXML path maps `Title`/`Subtitle`
to heading levels in the same spirit ([`styles.ts`](../src/core/tiptap-utils/ooxml/styles.ts)),
so both import routes render the document title the same way rather than one
as a heading and the other as a large plain paragraph.

---

## Running summary (for the deck)

| # | Issue | Root cause | Fixable in our code? | Extra effort |
|---|---|---|---|---|
| 1 | Images missing from exported HTML | Word export references internal-only file paths | Yes | ~25 lines |
| 2 | Floating images can't be captured at all | `Word.Shape` has no byte-extraction API — Office.js gap | **No — warn only** | ~40 lines + desktop-only requirement |
| 3 | Heading color/font lost | Style-based (not inline) formatting; Tiptap doesn't cascade CSS | Yes | ~10 lines + 1 dependency |
| 4 | Lists were visual-only, not real `<ol>/<li>` — couldn't continue on Enter | No official Tiptap extension for margin/indent; Word gives no semantic list marker to reconstruct from | Yes — reversed course and built real reconstruction (2026-08-06) | ~30 lines (visual) + ~180 lines (semantic reconstruction, heuristic-based) |
| 5 | Base64 images dropped | Library default (`allowBase64: false`) | Yes | 1 line |
| 6 | `&nbsp;` in excerpts | Naive regex strip, no entity decoding | Yes | refactor, no new code |
| 7 | Flagged-term highlighting | Word has no decoration/overlay concept — any highlight is a real saved edit | Yes, but different mechanism per host | ~55 lines web (live) + ~20 lines Word (one-shot scan, native highlight+comment) |
| 8 | Video embeds flattened to a static thumbnail | Word's HTML export has no real embed, just a linked poster image | Yes for YouTube (pattern-match); **unverified for enterprise/custom players** | ~30 lines, no new dependency |
| 9a | Icons render as broken images | Base64 payload is SVG, mislabeled `image/png` | Yes | ~15 lines |
| 9b | Shapes can't be captured | Word's export has no vector/shape semantics — `Word.Shape` has no byte-extraction API | **No — same as §2** | Author workaround only (Save as Picture); optional warning reuses §2's code |
| 10a | Paste-from-Word bypassed all fixes above | Tiptap's default paste path skips our HTML preprocessing entirely | Yes | ~10 lines, reuses existing code |
| 10b | Genuine non-picture Shape (or comment anchor) in selection kills entire paste's formatting — narrower than first stated, see §14 | Word fails to populate clipboard `text/html` at all for these — confirmed in prosemirror-view source; floating *pictures* were later confirmed NOT to have this problem (§14) | **No** | Author workaround only (paste smaller sections, avoid drawn shapes/comments) |
| 11 | Centered/right-aligned standalone images lost their alignment | Alignment lives on the wrapping paragraph, not the image; our Image node is block-level so it can't inherit it | Yes | ~75 lines, no new dependency |
| 12 | Couldn't set image alignment while editing; couldn't type a space next to an image | No editing-time command existed for §11's attribute; block-level image has no text content, so no paragraph nearby means nowhere to type | Yes | ~70 lines, no new dependency |
| 13a | Resize handles present but invisible/unclickable | `ResizableNodeView` ships handles unstyled by design; simplified `Image.configure({resize})` surface has no className passthrough | Yes | ~20 lines CSS |
| 13b | Enabling resize silently broke image alignment | Margin-auto alignment needs a wider containing block; resize wrapper is a flex item that shrinks to the image's exact size | Yes | ~30 lines CSS |
| 13c | Resizing an aligned image snapped back to full size on mouse-up | Resize node view wholesale-overwrites the `style` attribute on every attr update; alignment's inline `style` collided with resize's own imperative width/height | Yes | moved ~25 lines CSS out of the attribute system, net neutral |
| 14 | Floating/wrapped images (revisited, paste/import path): §2 called this a dead end — turns out it isn't | §2 was about `Word.Shape`/Office.js, now retired; via HTML the wrap direction survives as a plain `align=` on Word's own `<img>` fallback | Yes — reversed course again (like §4) | ~45 lines, verified via real captured markup and live paste |
| 15 | Pasted images (esp. charts/OLE objects) looked blurry — worse than the source | HTML's `<img>` fallback is a legacy low-res format (confirmed: 430×216px GIF for one real chart); `text/rtf`, read separately, carries the real embedded bytes (confirmed: a 102KB JPEG for the same chart) | Yes | ~90 lines, no new dependency; heuristic-based multi-image correlation, capped at 5 |
| 16 | OOXML→JSON dropped style-based formatting wholesale — lists broke, built-in styles (Quote/Caption…) flattened, indent/spacing/highlight/underline-style/small-caps/table-shading/TOC all lost | We resolved only heading level + default size from `styles.xml`; anything a paragraph/character **style** contributed (numbering, indent, italics, …) was invisible | Yes — built an effective style-props resolver + a handful of small attribute extensions | style resolver ~180 lines; ~5 small extensions; TOC node. Inline mid-sentence image left as known limit (needs inline-image schema) |

**Direction change (§10):** we're dropping the live Word add-in as the
authoring surface — every unfixable limitation above (§2, §9b) and the
heaviest workaround (§7's live scan) exist *because* it edits an open Word
document live. Tiptap becomes the only editor; Word becomes an import
format via copy-paste, not a connected host. That's a net simplification:
no manifest, no `WordApi`/`WordApiDesktop` requirement-set management, no
Office.js runtime risk — those go away entirely, not just get mitigated.

**Bottom line for the team:** with the add-in gone, genuine drawn Shapes
(and comment anchors) follow us into the paste world too, and worse than
expected — a Shape anywhere in a copied selection can silently blank out
formatting for the *whole* paste, confirmed live (§10b), not just fail to
render itself (§2/§9b). Practical guidance for authors: paste in smaller
sections and keep drawn shapes/comments out of the selection when
formatting matters. **This is specifically about genuine vector Shapes —
floating/anchored pictures are a different case, and are fully handled**
(§14 for correct positioning, §15 for full-fidelity image quality via
`text/rtf`, both confirmed against live paste, not just Save-As-HTML).
Everything else in this doc costs real but modest extra code (roughly 250
lines total across the POC) plus one new dependency (`juice`) — and now
applies uniformly whether content arrives via paste or a loaded saved
article. Video embeds work end-to-end for YouTube; an enterprise player
(e.g. Brightcove) is designed but not yet verified — pending access to test
internally.
