import { useMemo } from 'react';
import { useEditor, EditorContent } from '@tiptap/react';
import { EDITOR_EXTENSIONS } from '../../core/editorExtensions';
import { PublisherPanel } from '../publisher-panel/PublisherPanel';
import { createTiptapAdapter } from '../../adapters/tiptapAdapter';
import { EditorToolbar } from '../editor-toolbar/EditorToolbar';
import { ArticlesPanel } from '../articles/ArticlesPanel';

const BLANK_CONTENT = '<p>Start writing your article here…</p>';

export function FancyEditorTab() {
  const editor = useEditor({
    extensions: EDITOR_EXTENSIONS,
    content: BLANK_CONTENT,
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
      <div className="app-right-rail">
        <PublisherPanel adapter={adapter} hostLabel="Web" />
        <ArticlesPanel
          onSelectArticle={(json) => editor.chain().focus().setContent(json).run()}
          onCreateNew={() => editor.chain().focus().setContent(BLANK_CONTENT).run()}
        />
      </div>
    </div>
  );
}
