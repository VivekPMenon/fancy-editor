// Normalizes pasted Word content to the editor's own typography — the paste
// equivalent of what the OOXML path does in convertRun/convertParagraph.
// Word's HTML (after juice inlines its stylesheet) carries an explicit
// font-family, font-size, line-height, and paragraph margin on essentially
// every run/paragraph — all seeded from the MsoNormal body defaults — so
// pasted content would render in Word's typography and vertical rhythm, out of
// step with text the user types fresh (which adopts the editor's Segoe UI /
// 15px / 1.6-ish spacing). We strip the *inherited body defaults* and keep only
// values that deliberately DIFFER from them:
//
//   1. font-family — stripped everywhere. Everything adopts the editor base
//      font. The OOXML path never captures font-family at all (we don't read
//      <w:rFonts>), so this keeps the two import routes consistent.
//
//   2. font-size — strip the *body-default* size (the dominant inline size,
//      i.e. Word's MsoNormal size), plus any heading's inline size (headings
//      size via the tag + our CSS). Deliberately-different sizes (callouts,
//      captions) are kept.
//
//   3. line-height — same treatment: Word inlines `line-height: 115%` from
//      MsoNormal onto every paragraph, tighter than the editor's own line
//      spacing, so pasted paragraphs read cramped next to typed ones. Strip the
//      body-default line-height (and any heading's); keep a paragraph that was
//      deliberately set to a different spacing.
//
//   4. paragraph vertical margin (margin-top / margin-bottom) — Word's
//      `margin: 6pt 0in` spacing-before/after, redundant with (and conflicting
//      with) the editor's own `margin: 0 0 10px` block rhythm. Strip the
//      body-default top/bottom margin (and all of it on headings, where our CSS
//      owns heading spacing). Left/right margin is indent (§indentExtension) and
//      is left untouched.
//
// Runs last in the paste chain, so it also cleans the values the title / list /
// column reconstruction steps carry onto their own elements.

// An element's own directly-contained text length (not descendants') — used to
// weight the dominant-value vote so a parent doesn't double-count text that
// really belongs to its differently-styled child spans.
function directTextLength(el: Element): number {
  let length = 0;
  el.childNodes.forEach((node) => {
    if (node.nodeType === 3) {
      length += (node.textContent ?? '').trim().length;
    }
  });
  return length;
}

function isHeadingTag(el: Element): boolean {
  return /^H[1-6]$/.test(el.tagName);
}

// The value of a given inline style property carrying the most actual text —
// that's the MsoNormal body default (a big body of copy outweighs the odd
// callout/caption). Returns null when nothing carries the property.
function dominantValue(container: HTMLElement, read: (el: HTMLElement) => string): string | null {
  const weightByValue = new Map<string, number>();
  container.querySelectorAll<HTMLElement>('*').forEach((el) => {
    const value = read(el);
    if (value) {
      weightByValue.set(value, (weightByValue.get(value) ?? 0) + directTextLength(el));
    }
  });
  let dominant: string | null = null;
  let bestWeight = 0;
  for (const [value, weight] of weightByValue) {
    if (weight > bestWeight) {
      bestWeight = weight;
      dominant = value;
    }
  }
  return dominant;
}

export function normalizeWordTypography(html: string): string {
  if (!/font-size|font-family|line-height|margin/i.test(html)) {
    return html;
  }

  const container = document.createElement('div');
  container.innerHTML = html;

  const bodySize = dominantValue(container, (el) => el.style.fontSize);
  const bodyLineHeight = dominantValue(container, (el) => el.style.lineHeight);
  const bodyMarginTop = dominantValue(container, (el) => el.style.marginTop);
  const bodyMarginBottom = dominantValue(container, (el) => el.style.marginBottom);

  container.querySelectorAll<HTMLElement>('*').forEach((el) => {
    const heading = isHeadingTag(el);

    el.style.removeProperty('font-family');

    if (el.style.fontSize && (heading || el.style.fontSize === bodySize)) {
      el.style.removeProperty('font-size');
    }
    if (el.style.lineHeight && (heading || el.style.lineHeight === bodyLineHeight)) {
      el.style.removeProperty('line-height');
    }
    // Vertical margin only — left/right is indent, owned by indentExtension.
    // (Reading style.marginTop/Bottom expands Word's `margin: 6pt 0in`
    // shorthand via the CSSOM; removing the longhand re-serializes the rest.)
    if (el.style.marginTop && (heading || el.style.marginTop === bodyMarginTop)) {
      el.style.removeProperty('margin-top');
    }
    if (el.style.marginBottom && (heading || el.style.marginBottom === bodyMarginBottom)) {
      el.style.removeProperty('margin-bottom');
    }
  });

  return container.innerHTML;
}
