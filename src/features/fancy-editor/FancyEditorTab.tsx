import { useMemo } from 'react';
import { useEditor, EditorContent } from '@tiptap/react';
import { EDITOR_EXTENSIONS } from '../../core/tiptap-utils/editorExtensions';
import { PublisherPlugin } from '../publisher-panel/PublisherPlugin';
import { createTiptapAdapter } from '../../adapters/tiptapAdapter';
import { EditorToolbar } from '../editor-toolbar/EditorToolbar';
import { EditorContextMenu } from '../editor-toolbar/EditorContextMenu';
import { BLANK_ARTICLE_HTML } from '../articles/ArticlesPanel';

export function FancyEditorTab() {
  const editor = useEditor({
    extensions: EDITOR_EXTENSIONS,
    content: BLANK_ARTICLE_HTML,
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
        <EditorContextMenu editor={editor} />
      </div>
      <div className="app-right-rail">
        <PublisherPlugin adapter={adapter} hostLabel="Web" />
      </div>
    </div>
  );
}
