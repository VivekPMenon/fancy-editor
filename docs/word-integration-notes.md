# Word Integration — Limitations, Workarounds & Extra Effort

Living log for the team presentation. Updated as we build. Each entry: what
Word/Office.js/Tiptap does out of the box, what breaks, what we had to add,
and what it costs.

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

---

## 4. List indentation lost entirely — no Tiptap extension exists for it

**Limitation:** Word doesn't export numbered/bulleted lists as real
`<ol>/<ul>/<li>` — it exports plain `<p>` tags containing literal text
("1.", "·") plus `margin-left`/`text-indent` CSS to fake the hanging
indent. Tiptap has no built-in extension that captures `margin-left`/
`text-indent` at all (verified — no such attribute exists on the stock
paragraph/heading nodes).

**Scope decision:** we are **not** attempting semantic list reconstruction
(inferring real `<ol>/<li>` structure from the fake numbering) — out of
scope for this POC. We only preserve the *visual* indentation.

**Workaround:** wrote a custom Tiptap extension from scratch
(`Indent`) that adds `marginLeft`/`textIndent` as global attributes on
paragraph/heading nodes, parsing and re-rendering the inline CSS.

**Extra code:** new file
[`indentExtension.ts`](../src/core/tiptap-utils/indentExtension.ts),
~30 lines, entirely custom (no library covers this).

**Known gap:** because we scoped out semantic reconstruction, content
edited in the web/Tiptap app will never have "real" numbered lists coming
from a Word doc — just correctly-indented paragraphs that visually look
like a list but aren't structurally one (e.g. can't be re-numbered
automatically, screen readers won't announce them as a list).

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

## Running summary (for the deck)

| # | Issue | Root cause | Fixable in our code? | Extra effort |
|---|---|---|---|---|
| 1 | Images missing from exported HTML | Word export references internal-only file paths | Yes | ~25 lines |
| 2 | Floating images can't be captured at all | `Word.Shape` has no byte-extraction API — Office.js gap | **No — warn only** | ~40 lines + desktop-only requirement |
| 3 | Heading color/font lost | Style-based (not inline) formatting; Tiptap doesn't cascade CSS | Yes | ~10 lines + 1 dependency |
| 4 | List indentation lost | No official Tiptap extension for margin/indent | Yes (custom extension) | ~30 lines, fully custom |
| 5 | Base64 images dropped | Library default (`allowBase64: false`) | Yes | 1 line |
| 6 | `&nbsp;` in excerpts | Naive regex strip, no entity decoding | Yes | refactor, no new code |

**Bottom line for the team:** every gap so far has a workaround *except*
floating/anchored images, which is a genuine Office.js platform limitation
with no API-level fix — the only mitigation is pre-save detection +
warning (desktop Word only) and an authoring-guidelines ask (always use
inline images). Everything else costs real but modest extra code
(roughly 100 lines total across the POC so far) plus one new dependency
(`juice`).
