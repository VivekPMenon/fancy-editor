// Word's Title (and Subtitle) styles export as an ordinary
// `<p class=MsoTitle>` — NOT a heading tag the way Heading 1–6 do (those come
// through as real `<h1>`–`<h6>`). Their large size lives only in the
// stylesheet's `p.MsoTitle { font-size: 28pt }` rule, so it survives *only*
// when juice can inline that rule. A pasted clipboard payload that arrives
// without (or with a stripped) `<style>` block therefore loses the size
// entirely and the document title collapses to body text — confirmed against
// a real paste.
//
// The class on the element itself always survives, so we key off that:
// promote a Title paragraph to a real `<h1>` (and Subtitle to `<h2>`), which
// then renders via the editor's heading styles regardless of whether the
// size was inlined. Any inline style juice *did* manage to add (e.g. the
// real font-size when the stylesheet was present) is carried over, so it
// still wins where available. Mirrors the OOXML path, where Title/Subtitle
// are now mapped to heading levels too (styles.ts).

const TITLE_CLASS = /\bMsoTitle\b/;
const SUBTITLE_CLASS = /\bMsoSubtitle\b/;

export function reconstructWordTitles(html: string): string {
  if (!html.includes('MsoTitle') && !html.includes('MsoSubtitle')) {
    return html;
  }

  const container = document.createElement('div');
  container.innerHTML = html;

  container.querySelectorAll('p').forEach((p) => {
    const tag = TITLE_CLASS.test(p.className) ? 'h1' : SUBTITLE_CLASS.test(p.className) ? 'h2' : null;
    if (!tag) {
      return;
    }
    const heading = document.createElement(tag);
    // Carry over only the *appearance* the Title/Subtitle style contributed
    // (size, font, color, tracking) — NOT its margins. Word's Title uses a
    // tight 4pt bottom margin; copying that would override the editor's
    // heading spacing and leave the title jammed against the body text. With
    // margins dropped, the promoted heading gets its normal CSS spacing.
    for (const prop of ['font-size', 'font-family', 'color', 'letter-spacing']) {
      const value = p.style.getPropertyValue(prop);
      if (value) {
        heading.style.setProperty(prop, value);
      }
    }
    while (p.firstChild) {
      heading.appendChild(p.firstChild);
    }
    p.replaceWith(heading);
  });

  return container.innerHTML;
}
