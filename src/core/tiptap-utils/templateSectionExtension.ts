import { Node, mergeAttributes } from '@tiptap/core';
import { ReactNodeViewRenderer } from '@tiptap/react';
import { TemplateSectionView } from '../../features/fancy-editor/TemplateSectionView';
import { getTemplateSectionConfig, type TemplateSectionType } from './templateSectionTypes';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    templateSection: {
      insertTemplateSection: (sectionType: TemplateSectionType) => ReturnType;
    };
  }
}

// A structured article section — one Tiptap node parameterized by
// `sectionType` (see templateSectionTypes.ts) rather than 15 near-duplicate
// node types, following the same pattern as columnLayoutExtension (varies by
// `columns`) and tocEntryExtension (varies by `level`) elsewhere in this app.
// `selection` holds whatever the header control resolved to (a ticker, a
// country, a sector, …) as a plain string — structured metadata, same role
// `symbol` plays on TickerCard. The body is real editable document content
// (content: 'block+'), not an attribute, so normal rich-text editing works
// inside it — that's the one way this differs from TickerCard, which is atom
// (a card has no prose body to edit).
export const TemplateSection = Node.create({
  name: 'templateSection',
  group: 'block',
  content: 'block+',
  defining: true,
  draggable: true,

  addAttributes() {
    return {
      sectionType: {
        default: 'richText',
        parseHTML: (element) => element.getAttribute('data-section-type') ?? 'richText',
        renderHTML: (attributes) => ({ 'data-section-type': attributes.sectionType }),
      },
      selection: {
        default: '',
        parseHTML: (element) => element.getAttribute('data-selection') ?? '',
        renderHTML: (attributes) => ({ 'data-selection': attributes.selection }),
      },
    };
  },

  parseHTML() {
    // contentElement tells ProseMirror to pull the node's content from just
    // the `.template-section-static-body` child — without it, the static
    // header paragraph rendered alongside the content hole below would get
    // parsed back in as body content too.
    return [{ tag: 'div[data-template-section]', contentElement: '.template-section-static-body' }];
  },

  // Static fallback for contexts that don't run the React node view (Post
  // Feed's News format, a plain HTML export) — a labeled header plus the
  // real body content, not the interactive controls.
  //
  // The content hole (0) must be the ONLY child of its parent node in a
  // ProseMirror DOMOutputSpec — it can't sit next to the static header
  // paragraph at the same level (that's what threw "Content hole must be the
  // only child of its parent node" from editor.getHTML(), breaking Publish/
  // Save as Draft the moment a templateSection existed in the doc). Wrapping
  // the hole in its own div fixes it: the header and that wrapper are
  // siblings, and the hole is the sole child of the wrapper.
  renderHTML({ HTMLAttributes, node }) {
    const config = getTemplateSectionConfig(node.attrs.sectionType);
    const headerText = node.attrs.selection ? `${config.label}: ${node.attrs.selection}` : config.label;
    return [
      'div',
      mergeAttributes(HTMLAttributes, { 'data-template-section': '', class: 'template-section-static' }),
      ['p', { class: 'template-section-static-header' }, headerText],
      ['div', { class: 'template-section-static-body' }, 0],
    ];
  },

  addNodeView() {
    return ReactNodeViewRenderer(TemplateSectionView);
  },

  addCommands() {
    return {
      insertTemplateSection:
        (sectionType: TemplateSectionType) =>
        ({ commands }) =>
          commands.insertContent({
            type: this.name,
            attrs: { sectionType, selection: '' },
            content: [{ type: 'paragraph' }],
          }),
    };
  },
});
