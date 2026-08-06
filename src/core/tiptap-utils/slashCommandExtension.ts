import { Extension } from '@tiptap/core';
import type { Editor, Range } from '@tiptap/core';
import Suggestion from '@tiptap/suggestion';
import { ReactRenderer } from '@tiptap/react';
import {
  TextHeader1Regular,
  TextHeader2Regular,
  TextHeader3Regular,
  TextBulletListRegular,
  TextNumberListLtrRegular,
  TableFilled,
  LineHorizontal1Regular,
  ArrowTrendingRegular,
  type FluentIcon,
} from '@fluentui/react-icons';
import { SlashCommandList } from '../../features/fancy-editor/SlashCommandList';
import type { SlashCommandListRef } from '../../features/fancy-editor/SlashCommandList';

export interface SlashCommandItem {
  title: string;
  description: string;
  icon: FluentIcon;
  action: (editor: Editor, range: Range) => void;
}

const SLASH_COMMAND_ITEMS: SlashCommandItem[] = [
  {
    title: 'Heading 1',
    description: 'Big section heading',
    icon: TextHeader1Regular,
    action: (editor, range) => editor.chain().focus().deleteRange(range).setHeading({ level: 1 }).run(),
  },
  {
    title: 'Heading 2',
    description: 'Medium section heading',
    icon: TextHeader2Regular,
    action: (editor, range) => editor.chain().focus().deleteRange(range).setHeading({ level: 2 }).run(),
  },
  {
    title: 'Heading 3',
    description: 'Small section heading',
    icon: TextHeader3Regular,
    action: (editor, range) => editor.chain().focus().deleteRange(range).setHeading({ level: 3 }).run(),
  },
  {
    title: 'Bullet List',
    description: 'Simple bulleted list',
    icon: TextBulletListRegular,
    action: (editor, range) => editor.chain().focus().deleteRange(range).toggleBulletList().run(),
  },
  {
    title: 'Numbered List',
    description: 'List with numbering',
    icon: TextNumberListLtrRegular,
    action: (editor, range) => editor.chain().focus().deleteRange(range).toggleOrderedList().run(),
  },
  {
    title: 'Table',
    description: 'Insert a 3×3 table',
    icon: TableFilled,
    action: (editor, range) =>
      editor.chain().focus().deleteRange(range).insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run(),
  },
  {
    title: 'Divider',
    description: 'Horizontal rule',
    icon: LineHorizontal1Regular,
    action: (editor, range) => editor.chain().focus().deleteRange(range).setHorizontalRule().run(),
  },
  {
    title: 'Ticker Card',
    description: 'Live stock card — try switching the symbol',
    icon: ArrowTrendingRegular,
    action: (editor, range) => editor.chain().focus().deleteRange(range).insertTickerCard('AAPL').run(),
  },
];

// Notion-style "/" command palette — proves Tiptap can support entirely new
// interaction paradigms Word has no equivalent for, not just parity with its
// ribbon. Built on Tiptap's own Suggestion utility (the same primitive the
// official Mention extension uses).
export const SlashCommand = Extension.create({
  name: 'slashCommand',
  addProseMirrorPlugins() {
    const editor = this.editor;
    return [
      Suggestion<SlashCommandItem, SlashCommandItem>({
        editor,
        char: '/',
        startOfLine: false,
        items: ({ query }) =>
          SLASH_COMMAND_ITEMS.filter((item) => item.title.toLowerCase().includes(query.toLowerCase())).slice(0, 8),
        command: ({ editor: cmdEditor, range, props }) => {
          props.action(cmdEditor, range);
        },
        render: () => {
          let component: ReactRenderer<SlashCommandListRef>;
          let unmount: (() => void) | undefined;

          return {
            onStart: (props) => {
              component = new ReactRenderer(SlashCommandList, {
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
      }),
    ];
  },
});
