import { Node, mergeAttributes } from '@tiptap/core';
import { ReactNodeViewRenderer } from '@tiptap/react';
import { TickerCardView } from '../../features/fancy-editor/TickerCardView';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    tickerCard: {
      insertTickerCard: (symbol: string) => ReturnType;
    };
  }
}

// The "this is impossible in Word" node: a live, interactive React component
// embedded directly in the document flow, not a static image. Renders via
// ReactNodeViewRenderer instead of the usual renderHTML-only path, but still
// defines parseHTML/renderHTML so it round-trips through our HTML<->JSON
// pipeline (paste, save, Post Feed) as a plain, inert card — only the live
// editor gets the interactive version.
export const TickerCard = Node.create({
  name: 'tickerCard',
  group: 'block',
  atom: true,
  draggable: true,

  addAttributes() {
    return {
      symbol: {
        default: 'AAPL',
        parseHTML: (element) => element.getAttribute('data-symbol') ?? 'AAPL',
        renderHTML: (attributes) => ({ 'data-symbol': attributes.symbol }),
      },
    };
  },

  parseHTML() {
    return [{ tag: 'div[data-ticker-card]' }];
  },

  renderHTML({ HTMLAttributes, node }) {
    // Static fallback for contexts that don't run the React node view (Post
    // Feed's News format, a plain HTML export) — a simple, styled snapshot,
    // not the interactive card.
    return [
      'div',
      mergeAttributes(HTMLAttributes, { 'data-ticker-card': '', class: 'ticker-card-static' }),
      `📈 ${node.attrs.symbol}`,
    ];
  },

  addNodeView() {
    return ReactNodeViewRenderer(TickerCardView);
  },

  addCommands() {
    return {
      insertTickerCard:
        (symbol: string) =>
        ({ commands }) =>
          commands.insertContent({ type: this.name, attrs: { symbol } }),
    };
  },
});
