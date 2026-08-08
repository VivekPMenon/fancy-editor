// Word's HTML export marks an explicit column break (Insert -> Break ->
// Column, and the manual flow between newspaper columns) as a
// `<br style='mso-column-break-before:always'>` *inside* the paragraph —
// the HTML-side equivalent of OOXML's `<w:br w:type="column"/>`. Confirmed
// against a real captured paste.
//
// Word does NOT put the column *count* anywhere inline (it lives in an
// mso-specific @page rule that juice() doesn't apply to elements), so we
// infer it from the number of breaks: N breaks means N+1 content chunks,
// i.e. N+1 columns. Each such paragraph is split at its breaks into
// separate paragraphs and wrapped in a `<div data-column-layout
// data-columns=N>`, which the ColumnLayout node (columnLayoutExtension.ts)
// already knows how to parse — so no schema work is needed here, just this
// preprocessing pass.
//
// Only explicit column breaks are handled. A section Word auto-balanced
// into columns with no manual break leaves no marker to detect, so it stays
// single-column — same limitation as the OOXML path, and unavoidable
// without a signal to reconstruct from.

const COLUMN_BREAK_RE = /mso-column-break-before\s*:\s*always/i;

// Word puts the paragraph's font-size (and font-family/color) at the
// paragraph level — after juice() inlines the MsoNormal class rule, it's on
// the <p>'s own style. Splitting into fresh <p> elements would drop it
// (same issue the list reconstruction hit), so copy those text-appearance
// properties onto each new column paragraph. Our editor configures
// FontSize/Color/FontFamily for the 'paragraph' type, so paragraph-level
// style is picked up directly.
function copyParagraphTextStyle(source: HTMLElement, target: HTMLElement): void {
  if (source.style.fontSize) target.style.fontSize = source.style.fontSize;
  if (source.style.fontFamily) target.style.fontFamily = source.style.fontFamily;
  if (source.style.color) target.style.color = source.style.color;
}

function isColumnBreak(node: Node): boolean {
  return (
    node.nodeType === 1 &&
    (node as Element).tagName === 'BR' &&
    COLUMN_BREAK_RE.test((node as Element).getAttribute('style') || '')
  );
}

export function reconstructWordColumns(html: string): string {
  if (!/mso-column-break-before/i.test(html)) {
    return html;
  }

  const container = document.createElement('div');
  container.innerHTML = html;

  container.querySelectorAll('p').forEach((p) => {
    const hasColumnBreak = Array.from(p.childNodes).some(isColumnBreak);
    if (!hasColumnBreak) {
      return;
    }

    // Split the paragraph's children into segments at each column-break
    // <br>, preserving whatever inline content (spans, formatting) each
    // chunk carries. Collect references first, then move them — appendChild
    // relocates a live node, but the segment arrays already hold the
    // references so moving them one by one stays correct.
    const segments: Node[][] = [[]];
    Array.from(p.childNodes).forEach((node) => {
      if (isColumnBreak(node)) {
        segments.push([]);
      } else {
        segments[segments.length - 1].push(node);
      }
    });

    const layout = document.createElement('div');
    layout.setAttribute('data-column-layout', '');
    layout.setAttribute('data-columns', String(segments.length));

    segments.forEach((segmentNodes) => {
      const para = document.createElement('p');
      copyParagraphTextStyle(p, para);
      segmentNodes.forEach((node) => para.appendChild(node));
      layout.appendChild(para);
    });

    p.replaceWith(layout);
  });

  return container.innerHTML;
}
