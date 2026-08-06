// Word centers/right-aligns a standalone image by putting `align`/
// `text-align` on the *paragraph* wrapping it, not the image itself — but
// our Image node is block-level (inline: false), so it can't legally sit
// inside a paragraph node. Tiptap's parser hoists the image out as its own
// sibling block, and the wrapping paragraph — the only thing carrying the
// alignment — has nowhere to go. This runs before parsing: find paragraphs
// whose only real content is a single image, and transfer the alignment
// onto the <img> itself via a data-align attribute, so imageAlignmentExtension
// has something to read once the paragraph is gone.
//
// Deliberately narrow: only acts when the image is the paragraph's *only*
// meaningful content (Word's own convention for a centered standalone
// image) — an image sitting inline alongside real text is left alone, since
// "align the whole paragraph" and "align this one inline image" aren't the
// same thing.
export function resolveWordImageAlignment(html: string): string {
  if (!html.includes('<img')) {
    return html;
  }

  const container = document.createElement('div');
  container.innerHTML = html;

  container.querySelectorAll('p').forEach((p) => {
    const img = p.querySelector('img');
    if (!img) {
      return;
    }

    const otherText = Array.from(p.childNodes)
      .filter((node) => node !== img && !img.contains(node))
      .map((node) => node.textContent ?? '')
      .join('')
      .replace(/[\s ]/g, '');
    if (otherText.length > 0) {
      return;
    }

    const align = p.getAttribute('align') || (p as HTMLElement).style.textAlign;
    if (align === 'center' || align === 'right' || align === 'left') {
      img.setAttribute('data-align', align);
    }
  });

  return container.innerHTML;
}
