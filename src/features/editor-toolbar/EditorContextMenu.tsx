import { useEffect, useState } from 'react';
import { createPortal } from 'react-dom';
import type { Editor } from '@tiptap/react';
import { FluentProvider, Tooltip } from '@fluentui/react-components';
import {
  TextBoldRegular,
  TextItalicRegular,
  TextUnderlineRegular,
  CutRegular,
  CopyRegular,
  ClipboardPasteRegular,
  LinkRegular,
  TextFontRegular,
  TextParagraphRegular,
  TableInsertRowRegular,
  TableInsertColumnRegular,
  TableDeleteRowRegular,
  TableDeleteColumnRegular,
} from '@fluentui/react-icons';
import { ubsFluentTheme } from './ubsFluentTheme';
import './EditorContextMenu.css';

const NOT_IMPLEMENTED = 'Not implemented in this POC yet';

interface MenuPosition {
  x: number;
  y: number;
}

export function EditorContextMenu({ editor }: { editor: Editor }) {
  const [position, setPosition] = useState<MenuPosition | null>(null);

  useEffect(() => {
    const dom = editor.view.dom;

    function handleContextMenu(event: MouseEvent) {
      event.preventDefault();

      // Right-clicking outside the current selection moves the cursor there
      // first, matching native behavior — a right-click inside an existing
      // selection leaves it alone so Cut/Copy still act on the whole thing.
      const coords = editor.view.posAtCoords({ left: event.clientX, top: event.clientY });
      if (coords) {
        const { from, to } = editor.state.selection;
        if (coords.pos < from || coords.pos > to) {
          editor.chain().focus().setTextSelection(coords.pos).run();
        } else {
          editor.chain().focus().run();
        }
      }

      setPosition({ x: event.clientX, y: event.clientY });
    }

    dom.addEventListener('contextmenu', handleContextMenu);
    return () => dom.removeEventListener('contextmenu', handleContextMenu);
  }, [editor]);

  useEffect(() => {
    if (!position) {
      return;
    }
    function close() {
      setPosition(null);
    }
    // Capture phase + a microtask delay isn't needed here: this listener is
    // attached on the render *after* the menu opens, so the contextmenu
    // click that opened it has already finished bubbling.
    window.addEventListener('mousedown', close);
    window.addEventListener('scroll', close, true);
    window.addEventListener('keydown', closeOnEscape);
    return () => {
      window.removeEventListener('mousedown', close);
      window.removeEventListener('scroll', close, true);
      window.removeEventListener('keydown', closeOnEscape);
    };

    function closeOnEscape(event: KeyboardEvent) {
      if (event.key === 'Escape') {
        close();
      }
    }
  }, [position]);

  if (!position) {
    return null;
  }

  function run(action: () => void) {
    action();
    setPosition(null);
  }

  async function handlePaste() {
    try {
      // execCommand('paste') is blocked by browsers for scripted calls —
      // the async Clipboard API is the only way to read clipboard content
      // programmatically. It only reliably gets plain text without extra
      // permission/support complexity, so rich Word formatting pasted this
      // way (vs. a real Ctrl+V, which goes through WordPaste's HTML
      // preprocessing) won't carry over — plain text only for now.
      const text = await navigator.clipboard.readText();
      if (text) {
        editor.chain().focus().insertContent(text).run();
      }
    } catch {
      // Clipboard permission denied or unavailable — silently no-op rather
      // than break the menu.
    }
  }

  function handleLink() {
    if (editor.isActive('link')) {
      editor.chain().focus().unsetLink().run();
      return;
    }
    const url = window.prompt('Enter a URL');
    if (!url) {
      return;
    }
    editor.chain().focus().extendMarkRange('link').setLink({ href: url }).run();
  }

  return createPortal(
    <FluentProvider theme={ubsFluentTheme} className="editor-context-menu-fluent-root">
      <div className="editor-context-menu" style={{ left: position.x, top: position.y }} onContextMenu={(e) => e.preventDefault()}>
        <div className="editor-context-menu-quick-format">
          <button
            type="button"
            className={editor.isActive('bold') ? 'active' : ''}
            title="Bold"
            onClick={() => run(() => editor.chain().focus().toggleBold().run())}
          >
            <TextBoldRegular />
          </button>
          <button
            type="button"
            className={editor.isActive('italic') ? 'active' : ''}
            title="Italic"
            onClick={() => run(() => editor.chain().focus().toggleItalic().run())}
          >
            <TextItalicRegular />
          </button>
          <button
            type="button"
            className={editor.isActive('underline') ? 'active' : ''}
            title="Underline"
            onClick={() => run(() => editor.chain().focus().toggleUnderline().run())}
          >
            <TextUnderlineRegular />
          </button>
        </div>

        <div className="editor-context-menu-divider" />

        <button type="button" className="editor-context-menu-item" onClick={() => run(() => document.execCommand('cut'))}>
          <CutRegular /> Cut
        </button>
        <button type="button" className="editor-context-menu-item" onClick={() => run(() => document.execCommand('copy'))}>
          <CopyRegular /> Copy
        </button>
        <button type="button" className="editor-context-menu-item" onClick={() => run(handlePaste)}>
          <ClipboardPasteRegular /> Paste
        </button>

        {editor.isActive('table') && (
          <>
            <div className="editor-context-menu-divider" />
            <button
              type="button"
              className="editor-context-menu-item"
              onClick={() => run(() => editor.chain().focus().addRowAfter().run())}
            >
              <TableInsertRowRegular /> Insert Row
            </button>
            <button
              type="button"
              className="editor-context-menu-item"
              onClick={() => run(() => editor.chain().focus().addColumnAfter().run())}
            >
              <TableInsertColumnRegular /> Insert Column
            </button>
            <button
              type="button"
              className="editor-context-menu-item"
              onClick={() => run(() => editor.chain().focus().deleteRow().run())}
            >
              <TableDeleteRowRegular /> Delete Row
            </button>
            <button
              type="button"
              className="editor-context-menu-item"
              onClick={() => run(() => editor.chain().focus().deleteColumn().run())}
            >
              <TableDeleteColumnRegular /> Delete Column
            </button>
          </>
        )}

        <div className="editor-context-menu-divider" />

        <button type="button" className="editor-context-menu-item" onClick={() => run(handleLink)}>
          <LinkRegular /> {editor.isActive('link') ? 'Remove Link' : 'Link…'}
        </button>
        <Tooltip content={NOT_IMPLEMENTED} relationship="label" withArrow>
          <button type="button" className="editor-context-menu-item" disabled>
            <TextFontRegular /> Font…
          </button>
        </Tooltip>
        <Tooltip content={NOT_IMPLEMENTED} relationship="label" withArrow>
          <button type="button" className="editor-context-menu-item" disabled>
            <TextParagraphRegular /> Paragraph…
          </button>
        </Tooltip>
      </div>
    </FluentProvider>,
    document.body,
  );
}
