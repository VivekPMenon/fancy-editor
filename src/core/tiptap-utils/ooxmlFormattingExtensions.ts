import { Extension } from '@tiptap/core';

// A grab-bag of small global-attribute extensions that give the OOXML import
// somewhere to land formatting Tiptap's stock marks/nodes don't model:
//  - underline *style* (dotted / wavy / double) on the underline mark
//  - double strikethrough on the strike mark
//  - small-caps (font-variant) on the textStyle mark
//  - cell background shading on table cells
// Each renders as an inline style so it round-trips through generateHTML into
// the Post Feed unchanged, same approach as listStyleExtension.

// text-decoration-style on the <u> produced by the underline mark. A plain
// single underline leaves this null and renders as the default solid line.
export const UnderlineStyle = Extension.create({
  name: 'underlineStyle',
  addGlobalAttributes() {
    return [
      {
        types: ['underline'],
        attributes: {
          underlineStyle: {
            default: null,
            parseHTML: (element) => element.style.textDecorationStyle || null,
            renderHTML: (attributes) =>
              attributes.underlineStyle ? { style: `text-decoration-style: ${attributes.underlineStyle}` } : {},
          },
        },
      },
    ];
  },
});

// Word's double strikethrough (<w:dstrike>) — a double line-through, rendered
// by pairing the strike mark's <s> with text-decoration-style: double.
export const StrikeStyle = Extension.create({
  name: 'strikeStyle',
  addGlobalAttributes() {
    return [
      {
        types: ['strike'],
        attributes: {
          strikeStyle: {
            default: null,
            parseHTML: (element) => (element.style.textDecorationStyle === 'double' ? 'double' : null),
            renderHTML: (attributes) =>
              attributes.strikeStyle ? { style: `text-decoration-style: ${attributes.strikeStyle}` } : {},
          },
        },
      },
    ];
  },
});

// font-variant on the textStyle mark's <span>, for Word's small caps.
export const FontVariant = Extension.create({
  name: 'fontVariant',
  addGlobalAttributes() {
    return [
      {
        types: ['textStyle'],
        attributes: {
          fontVariant: {
            default: null,
            parseHTML: (element) => element.style.fontVariant || null,
            renderHTML: (attributes) =>
              attributes.fontVariant ? { style: `font-variant: ${attributes.fontVariant}` } : {},
          },
        },
      },
    ];
  },
});

// Cell background shading (Word's <w:shd w:fill>) on table cells/headers — the
// stock TableKit cell has no fill attribute, so a shaded header (and, worse, a
// header with white text on a dropped dark fill) came through blank.
export const TableCellBackground = Extension.create({
  name: 'tableCellBackground',
  addGlobalAttributes() {
    return [
      {
        types: ['tableCell', 'tableHeader'],
        attributes: {
          backgroundColor: {
            default: null,
            parseHTML: (element) => element.style.backgroundColor || null,
            renderHTML: (attributes) =>
              attributes.backgroundColor ? { style: `background-color: ${attributes.backgroundColor}` } : {},
          },
        },
      },
    ];
  },
});
