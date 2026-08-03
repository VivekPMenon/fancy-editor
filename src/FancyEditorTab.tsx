import { useMemo } from 'react';
import { useEditor, EditorContent } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import { TableKit } from '@tiptap/extension-table';
import { TextAlign } from '@tiptap/extension-text-align';
import { TextStyle } from '@tiptap/extension-text-style';
import { Color } from '@tiptap/extension-color';
import { FontFamily } from '@tiptap/extension-font-family';
import { Youtube } from '@tiptap/extension-youtube';
import { PublisherPanel } from './core/PublisherPanel';
import { createTiptapAdapter } from './adapters/tiptapAdapter';
import { EditorToolbar } from './EditorToolbar';

export function FancyEditorTab() {
  const editor = useEditor({
    extensions: [
      StarterKit.configure({ link: { openOnClick: false } }),
      TableKit.configure({ table: { resizable: true } }),
      TextAlign.configure({ types: ['heading', 'paragraph'] }),
      TextStyle,
      Color,
      FontFamily,
      Youtube.configure({ width: 480, height: 270 }),
    ],
    content: '<p>Start writing your article here…</p>',
  });

  // Memoized so the adapter's identity stays stable across re-renders;
  // PublisherPanel resubscribes to onContentChange whenever it changes.
  const adapter = useMemo(() => (editor ? createTiptapAdapter(editor) : null), [editor]);

  if (!editor || !adapter) {
    return null;
  }

  return (
    <div className="app-layout">
      <div className="tiptap-editor-column">
        <EditorToolbar editor={editor} />
        <EditorContent className="tiptap-editor" editor={editor} />
      </div>
      <PublisherPanel adapter={adapter} hostLabel="Web" />
    </div>
  );
}
