// Word expresses an image's positioning in two genuinely different ways,
// and both live outside the image's own node once we parse — our Image
// node is block-level (inline: false), so it can't legally sit inside a
// paragraph node's content, and Tiptap's parser hoists it out, discarding
// whatever carried the positioning:
//
// 1. Simple, non-wrapping alignment ("Center"/"Right" picture position, no
//    text wrap) — this is on the wrapping *paragraph* (`align=`/
//    `text-align:`), and only appears when the image is that paragraph's
//    sole content.
// 2. Real floating/anchored images (Wrap Text = Square/Tight/Through/etc.)
//    — confirmed against a real "Save As Web Page" export
//    (samples/saved-from-word/Sample-floating.htm): the positioning signal
//    here is a legacy `align="left"`/`"right"` attribute on the <img> tag
//    itself (Word's non-VML fallback, which is what every modern browser —
//    and our own preprocessing — actually sees; the VML block above it is
//    swallowed as one long HTML comment by any non-IE parser). Crucially,
//    a wrapped image commonly sits *inline alongside real paragraph text*,
//    not alone — that's the whole point of "wrap".
//
// Both get resolved onto a `data-align` attribute for imageAlignmentExtension
// to render as CSS (`left`/`center`/`right` → justify-content/margin;
// `float-left`/`float-right` → real CSS float). Case 2 additionally hoists
// the <img> out to a sibling positioned immediately before its paragraph
// when there's other text in that paragraph, since a block-level image
// can't stay inline — the paragraph's text becomes an ordinary sibling
// afterward, and CSS float still wraps it around the hoisted image exactly
// as Word renders it, even though they're no longer literally one node.
export function resolveWordImageAlignment(html: string): string {
  if (!html.includes('<img')) {
    return html;
  }

  const container = document.createElement('div');
  container.innerHTML = html;

  container.querySelectorAll('img').forEach((img) => {
    const p = img.closest('p');
    if (!p) {
      return;
    }

    const imgAlign = img.getAttribute('align');
    if (imgAlign === 'left' || imgAlign === 'right') {
      img.setAttribute('data-align', `float-${imgAlign}`);
      if (paragraphHasOtherContent(p, img)) {
        p.parentElement?.insertBefore(img, p);
      }
      return;
    }

    // Simple case: a lone image in its own paragraph, alignment expressed
    // on the wrapping <p> itself.
    if (paragraphHasOtherContent(p, img)) {
      return;
    }
    const align = p.getAttribute('align') || (p as HTMLElement).style.textAlign;
    if (align === 'center' || align === 'right' || align === 'left') {
      img.setAttribute('data-align', align);
    }
  });

  return container.innerHTML;
}

function paragraphHasOtherContent(p: Element, img: Element): boolean {
  // Comment nodes are excluded deliberately — Word's VML fallback block
  // (and the downlevel-revealed `<![if !vml]>`/`<![endif]>` markers
  // wrapping the plain <img>) parse as Comment nodes in any non-IE parser,
  // but Node.textContent still returns a Comment's full text data, which
  // would otherwise make every floating image look like it shares its
  // paragraph with "other content" even when it's genuinely alone there.
  return (
    Array.from(p.childNodes)
      .filter((node) => node !== img && !img.contains(node) && node.nodeType !== 8)
      .map((node) => node.textContent ?? '')
      .join('')
      .replace(/[\s ]/g, '').length > 0
  );
}
