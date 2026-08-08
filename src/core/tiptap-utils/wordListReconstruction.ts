// Semantic list reconstruction — deliberately scoped OUT earlier in this
// project (see docs/word-integration-notes.md §4) in favor of preserving
// only visual indentation. Revisited because pressing Enter inside a fake
// list doesn't continue the numbering/bullet — a real problem for the
// "paste from Word, keep editing" pitch, not just a fidelity nitpick.
//
// Word's HTML export never emits real <ol>/<ul>/<li> (confirmed via
// Microsoft's own Q&A: "this is not a library bug — it's how Word works by
// design"). Instead every list item is a paragraph carrying an
// MsoListParagraph* class, a hanging indent (negative text-indent), and a
// manually "typed" marker (a bullet character or "1.") followed by a run of
// non-breaking spaces standing in for a tab. This walks the DOM, detects
// that shape, and rebuilds real list nodes from it — a heuristic on export
// shape, not a documented contract, same caveat as the video-embed and
// icon-MIME fixes elsewhere in this file set.
//
// Must run AFTER juice() — margin-left/text-indent are frequently defined
// via the MsoListParagraph* class's stylesheet rule rather than inline, and
// this reads element.style directly.

const LIST_PARAGRAPH_CLASS = /\bMsoListParagraph/;

// Marker + tab-padding shape: an optional bullet/number/letter/roman token,
// then a run of 2+ non-breaking-space/whitespace characters simulating a
// tab stop before the real content starts.
const MARKER_PATTERN =
  /^[\s ]*([•·▪◦§o]|[0-9]+(?:\.[0-9]+)*[.)]|[a-zA-Z]{1,3}[.)]|[ivxlcdmIVXLCDM]{1,6}[.)])[\s ]{2,}/;

interface ListParagraphInfo {
  element: Element;
  level: number;
  ordered: boolean;
  contentFragment: DocumentFragment;
  // The paragraph's own font-size/family/color, applied to the <li> so the
  // list marker (which inherits its font from the <li>, not the item's
  // inner content) matches the item text — same point-size but a different
  // typeface on the marker still reads as a size mismatch, so family
  // matters as much as size. The inner content also keeps all of this via
  // preserveParagraphTextStyle's span, so the text stays correct even if
  // the <li>-level styling ever fails to round-trip.
  fontSize?: string;
  fontFamily?: string;
  color?: string;
}

function inchesFromCss(value: string): number {
  const match = /^(-?[\d.]+)(in|pt|cm|px)?$/.exec(value.trim());
  if (!match) {
    return 0;
  }
  const num = parseFloat(match[1]);
  switch (match[2]) {
    case 'pt':
      return num / 72;
    case 'cm':
      return num / 2.54;
    case 'px':
      return num / 96;
    default:
      return num; // 'in', or unitless (Word's own convention is inches).
  }
}

// Word indents ~0.5in per nesting level — round the resolved margin-left
// into a level number rather than trying to track exact measurements.
function levelFromIndent(el: HTMLElement): number {
  if (!el.style.marginLeft) {
    return 0;
  }
  return Math.max(0, Math.round(inchesFromCss(el.style.marginLeft) / 0.5));
}

function isBulletToken(token: string, fontFamilyHint: string): boolean {
  if (/symbol|wingdings/i.test(fontFamilyHint)) {
    return true;
  }
  if (/^[0-9]+(?:\.[0-9]+)*[.)]$/.test(token)) {
    return false;
  }
  if (/^[a-zA-Z]{1,3}[.)]$/.test(token) || /^[ivxlcdmIVXLCDM]{1,6}[.)]$/.test(token)) {
    return false;
  }
  return true; // A bare symbol (·, o, §, …) with no trailing punctuation.
}

// Finds the marker+padding boundary in a paragraph's flattened text, then
// physically cuts the DOM there via Range so the remaining content
// (preserving whatever inline formatting it carries) can be moved straight
// into a new <li> — much simpler than trying to reconstruct formatting from
// scratch.
function analyzeListParagraph(p: Element): ListParagraphInfo | null {
  if (!(p instanceof HTMLElement)) {
    return null;
  }
  const isCandidate = LIST_PARAGRAPH_CLASS.test(p.className) || parseFloat(p.style.textIndent || '0') < 0;
  if (!isCandidate) {
    return null;
  }

  const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
  const boundaries: { textNode: Text; start: number; end: number }[] = [];
  let combined = '';
  let current = walker.nextNode();
  while (current) {
    const textNode = current as Text;
    const start = combined.length;
    combined += textNode.data;
    boundaries.push({ textNode, start, end: combined.length });
    current = walker.nextNode();
  }

  const match = MARKER_PATTERN.exec(combined);
  if (!match) {
    return null;
  }

  const splitOffset = match[0].length;
  const marker = match[1];
  const target = boundaries.find((b) => splitOffset >= b.start && splitOffset <= b.end);
  if (!target || !p.lastChild) {
    return null;
  }

  const markerFontFamily =
    (target.textNode.parentElement?.closest<HTMLElement>('[style*="font-family"]'))?.style.fontFamily ?? '';

  const range = document.createRange();
  range.setStart(target.textNode, splitOffset - target.start);
  range.setEndAfter(p.lastChild);
  const contentFragment = range.extractContents();
  removeEmptySpans(contentFragment);
  preserveParagraphTextStyle(p, contentFragment);

  return {
    element: p,
    level: levelFromIndent(p),
    ordered: !isBulletToken(marker, markerFontFamily),
    contentFragment,
    fontSize: p.style.fontSize || undefined,
    fontFamily: p.style.fontFamily || undefined,
    color: p.style.color || undefined,
  };
}

// Word puts a list paragraph's font-size (and font-family/color) at the
// paragraph level — after juice() inlines the MsoListParagraph* class rule,
// it lands on the <p>'s own style. But only the paragraph's *inner* content
// gets hoisted into the reconstructed <li>; the <p> itself is discarded, so
// that paragraph-level styling would be lost (non-list paragraphs keep it
// because their <p> survives as a real paragraph node). This re-applies it
// by wrapping the extracted content in a span carrying those properties.
// Any inner span with its own font-size still wins for its own text — the
// wrapper is only a fallback for content that had none of its own.
function preserveParagraphTextStyle(p: HTMLElement, fragment: DocumentFragment): void {
  const props: string[] = [];
  if (p.style.fontSize) props.push(`font-size:${p.style.fontSize}`);
  if (p.style.fontFamily) props.push(`font-family:${p.style.fontFamily}`);
  if (p.style.color) props.push(`color:${p.style.color}`);
  if (props.length === 0 || fragment.childNodes.length === 0) {
    return;
  }
  const wrapper = document.createElement('span');
  wrapper.setAttribute('style', props.join(';'));
  while (fragment.firstChild) {
    wrapper.appendChild(fragment.firstChild);
  }
  fragment.appendChild(wrapper);
}

// extractContents() clones whatever ancestor spans structurally wrapped the
// marker text (font-family:Symbol, the padding span, …) even though the
// range's start point left them with no actual text — harmless empty shells
// for parsing, but worth tidying out of the generated markup. These are
// often nested (an empty span inside another empty span), so this repeats
// until nothing empty is left rather than a single top-down pass, which
// would strip the innermost one and leave its now-empty parent behind.
function removeEmptySpans(fragment: DocumentFragment): void {
  let removedAny = true;
  while (removedAny) {
    removedAny = false;
    fragment.querySelectorAll('span').forEach((span) => {
      if (!span.textContent?.trim() && span.children.length === 0) {
        span.remove();
        removedAny = true;
      }
    });
  }
}

function buildListTree(runInfos: ListParagraphInfo[]): HTMLElement {
  const root = document.createElement(runInfos[0].ordered ? 'ol' : 'ul');
  const stack: { level: number; listEl: HTMLElement }[] = [{ level: runInfos[0].level, listEl: root }];
  const lastLiByLevel = new Map<number, HTMLElement>();

  for (const info of runInfos) {
    while (stack.length > 1 && info.level < stack[stack.length - 1].level) {
      stack.pop();
    }

    if (info.level > stack[stack.length - 1].level) {
      const parentLi = lastLiByLevel.get(stack[stack.length - 1].level);
      const newList = document.createElement(info.ordered ? 'ol' : 'ul');
      (parentLi ?? stack[stack.length - 1].listEl).appendChild(newList);
      stack.push({ level: info.level, listEl: newList });
    }

    const li = document.createElement('li');
    if (info.fontSize) {
      li.style.fontSize = info.fontSize;
    }
    if (info.fontFamily) {
      li.style.fontFamily = info.fontFamily;
    }
    if (info.color) {
      li.style.color = info.color;
    }
    li.appendChild(info.contentFragment);
    stack[stack.length - 1].listEl.appendChild(li);
    lastLiByLevel.set(stack[stack.length - 1].level, li);
  }

  return root;
}

export function reconstructWordLists(html: string): string {
  if (!html.includes('MsoListParagraph') && !html.includes('text-indent:-')) {
    return html;
  }

  const container = document.createElement('div');
  container.innerHTML = html;

  // Group candidates by immediate parent first, so a run of list paragraphs
  // only merges with siblings under the same parent — keeps unrelated lists
  // elsewhere in the document from being merged together.
  const candidateParents = new Set<Element>();
  container.querySelectorAll('p').forEach((p) => {
    const isCandidate = LIST_PARAGRAPH_CLASS.test(p.className) || parseFloat((p as HTMLElement).style.textIndent || '0') < 0;
    if (isCandidate && p.parentElement) {
      candidateParents.add(p.parentElement);
    }
  });

  candidateParents.forEach((parent) => {
    let children = Array.from(parent.children);
    let i = 0;
    while (i < children.length) {
      const info = analyzeListParagraph(children[i]);
      if (!info) {
        i += 1;
        continue;
      }

      const runInfos: ListParagraphInfo[] = [info];
      let j = i + 1;
      while (j < children.length) {
        const nextInfo = analyzeListParagraph(children[j]);
        if (!nextInfo) {
          break;
        }
        runInfos.push(nextInfo);
        j += 1;
      }

      const listTree = buildListTree(runInfos);
      parent.insertBefore(listTree, children[i]);
      runInfos.forEach((r) => r.element.remove());

      children = Array.from(parent.children);
      i = children.indexOf(listTree) + 1;
    }
  });

  return container.innerHTML;
}
