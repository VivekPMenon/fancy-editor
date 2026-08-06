import Mention from '@tiptap/extension-mention';
import { ReactRenderer } from '@tiptap/react';
import { MENTIONABLE_PEOPLE, type MentionablePerson } from '../mentionablePeople';
import { MentionList } from '../../features/fancy-editor/MentionList';
import type { MentionListRef } from '../../features/fancy-editor/MentionList';

// Official @tiptap/extension-mention, wired to our own hardcoded people list
// (reusing the mock post authors) and a custom dropdown — same
// Suggestion-based mount pattern as the slash command palette.
export const MentionExtension = Mention.configure({
  HTMLAttributes: { class: 'mention-chip' },
  renderText({ node }) {
    return `@${node.attrs.label ?? node.attrs.id}`;
  },
  renderHTML({ node }) {
    return ['span', { class: 'mention-chip', 'data-mention-id': node.attrs.id }, `@${node.attrs.label ?? node.attrs.id}`];
  },
  suggestion: {
    items: ({ query }) =>
      MENTIONABLE_PEOPLE.filter((person) => person.name.toLowerCase().includes(query.toLowerCase())).slice(0, 5),
    command: ({ editor, range, props }) => {
      // Mention's own types pin the suggestion's `props` to MentionNodeAttrs
      // (id/label), but at runtime it's whatever MentionList's `command`
      // prop was called with — our own MentionablePerson (id/name/role).
      const person = props as unknown as MentionablePerson;
      editor
        .chain()
        .focus()
        .insertContentAt(range, [
          { type: 'mention', attrs: { id: person.id, label: person.name } },
          { type: 'text', text: ' ' },
        ])
        .run();
    },
    render: () => {
      let component: ReactRenderer<MentionListRef>;
      let unmount: (() => void) | undefined;

      return {
        onStart: (props) => {
          component = new ReactRenderer(MentionList, {
            props: { items: props.items, command: props.command },
            editor: props.editor,
          });
          unmount = props.mount(component.element);
        },
        onUpdate: (props) => {
          component.updateProps({ items: props.items, command: props.command });
        },
        onKeyDown: (props) => {
          if (props.event.key === 'Escape') {
            unmount?.();
            return true;
          }
          return component.ref?.onKeyDown(props) ?? false;
        },
        onExit: () => {
          unmount?.();
          component.destroy();
        },
      };
    },
  },
});
