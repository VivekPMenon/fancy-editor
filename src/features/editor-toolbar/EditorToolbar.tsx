import { useEffect, useState, type ReactNode } from 'react';
import type { Editor } from '@tiptap/react';
import { CIcon } from '@coreui/icons-react';
import {
  cilActionUndo,
  cilActionRedo,
  // cilBold,
  // cilItalic,
  // cilUnderline,
  // cilStrikethrough,
  // cilFont,
  cilAlignLeft,
  cilAlignCenter,
  cilAlignRight,
  cilMenu,
  cilList,
  cilListNumbered,
  cilDoubleQuoteSansRight,
  cilGrid,
  cilLink,
  cilVideo,
  cilPlus,
  cilTrash,
} from '@coreui/icons';
import './EditorToolbar.css';

interface EditorToolbarProps {
  editor: Editor;
}

const FONT_FAMILIES = [
  { label: 'Default font', value: '' },
  { label: 'Arial', value: 'Arial' },
  { label: 'Georgia', value: 'Georgia' },
  { label: 'Times New Roman', value: "'Times New Roman', serif" },
  { label: 'Courier New', value: "'Courier New', monospace" },
  { label: 'Verdana', value: 'Verdana' },
];

const FONT_SIZES = ['8', '9', '10', '10.5', '11', '12', '14', '16', '18', '20', '22', '24', '26', '28', '36', '48', '72'];

// Cosmetic only — this POC implements the Home tab; the rest are inert, same
// as a real ribbon has tabs you're not currently on.
const RIBBON_TABS = ['File', 'Home', 'Insert', 'Draw', 'Design', 'Layout', 'References', 'Review', 'View'];

function ToolbarGroup({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div className="editor-toolbar-group">
      <div className="editor-toolbar-group-body">{children}</div>
      <span className="editor-toolbar-group-label">{label}</span>
    </div>
  );
}

function ToolbarButton({
  icon,
  label,
  title,
  active,
  onClick,
}: {
  icon: string[];
  label?: string;
  title: string;
  active?: boolean;
  onClick: () => void;
}) {
  return (
    <button type="button" title={title} className={active ? 'active' : ''} onClick={onClick}>
      <CIcon icon={icon} size="sm" />
      {label && <span className="editor-toolbar-button-label">{label}</span>}
    </button>
  );
}

export function EditorToolbar({ editor }: EditorToolbarProps) {
  // Tiptap doesn't re-render React on its own; re-render on every editor
  // transaction so active-state (bold/italic/heading/…) buttons stay in sync.
  const [, forceRender] = useState(0);

  useEffect(() => {
    const rerender = () => forceRender((n) => n + 1);
    editor.on('transaction', rerender);
    return () => {
      editor.off('transaction', rerender);
    };
  }, [editor]);

  const headingValue = ([1, 2, 3] as const).find((level) => editor.isActive('heading', { level }));

  function handleFontFamilyChange(value: string) {
    if (!value) {
      editor.chain().focus().unsetFontFamily().run();
    } else {
      editor.chain().focus().setFontFamily(value).run();
    }
  }

  function handleFontSizeChange(value: string) {
    if (!value) {
      editor.chain().focus().unsetFontSize().run();
    } else {
      editor.chain().focus().setFontSize(`${value}px`).run();
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

  function handleVideo() {
    const url = window.prompt('Enter a YouTube video URL');
    if (!url) {
      return;
    }
    editor.chain().focus().setYoutubeVideo({ src: url }).run();
  }

  // const currentColor = editor.getAttributes('textStyle').color || '#000000';

  return (
    <div className="editor-toolbar-shell">
      <div className="editor-toolbar-tabs">
        {RIBBON_TABS.map((tab) => (
          <button
            key={tab}
            type="button"
            className={tab === 'Home' ? 'active' : ''}
            disabled={tab !== 'Home'}
            title={tab === 'Home' ? undefined : 'Not implemented in this POC'}
          >
            {tab}
          </button>
        ))}
      </div>

      <div className="editor-toolbar">
        <ToolbarGroup label="History">
          <div className="editor-toolbar-icon-grid">
            <ToolbarButton icon={cilActionUndo} title="Undo" onClick={() => editor.chain().focus().undo().run()} />
            <ToolbarButton icon={cilActionRedo} title="Redo" onClick={() => editor.chain().focus().redo().run()} />
          </div>
        </ToolbarGroup>

        <ToolbarGroup label="Font">
          <div className="editor-toolbar-font-row">
            <select
              className="editor-toolbar-select editor-toolbar-select-family"
              title="Font family"
              value={editor.getAttributes('textStyle').fontFamily ?? ''}
              onChange={(event) => handleFontFamilyChange(event.target.value)}
            >
              {FONT_FAMILIES.map((font) => (
                <option key={font.label} value={font.value}>
                  {font.label}
                </option>
              ))}
            </select>
            <select
              className="editor-toolbar-select editor-toolbar-select-size"
              title="Font size"
              value={editor.getAttributes('textStyle').fontSize?.replace('px', '') ?? '18'}
              onChange={(event) => handleFontSizeChange(event.target.value)}
            >
              <option value="">--</option>
              {FONT_SIZES.map((size) => (
                <option key={size} value={size}>
                  {size}
                </option>
              ))}
            </select>
          </div>
          {/* <div className="editor-toolbar-icon-grid">
            <ToolbarButton
              icon={cilBold}
              title="Bold"
              active={editor.isActive('bold')}
              onClick={() => editor.chain().focus().toggleBold().run()}
            />
            <ToolbarButton
              icon={cilItalic}
              title="Italic"
              active={editor.isActive('italic')}
              onClick={() => editor.chain().focus().toggleItalic().run()}
            />
            <ToolbarButton
              icon={cilUnderline}
              title="Underline"
              active={editor.isActive('underline')}
              onClick={() => editor.chain().focus().toggleUnderline().run()}
            />
            <ToolbarButton
              icon={cilStrikethrough}
              title="Strikethrough"
              active={editor.isActive('strike')}
              onClick={() => editor.chain().focus().toggleStrike().run()}
            />
            <label className="editor-toolbar-color-button" title="Font color">
              <CIcon icon={cilFont} size="sm" />
              <span className="editor-toolbar-color-swatch" style={{ backgroundColor: currentColor }} />
              <input
                type="color"
                value={currentColor}
                onChange={(event) => editor.chain().focus().setColor(event.target.value).run()}
              />
            </label>
          </div> */}
        </ToolbarGroup>

        <ToolbarGroup label="Paragraph">
          <div className="editor-toolbar-icon-grid">
            <ToolbarButton
              icon={cilList}
              title="Bullet list"
              active={editor.isActive('bulletList')}
              onClick={() => editor.chain().focus().toggleBulletList().run()}
            />
            <ToolbarButton
              icon={cilListNumbered}
              title="Numbered list"
              active={editor.isActive('orderedList')}
              onClick={() => editor.chain().focus().toggleOrderedList().run()}
            />
            <ToolbarButton
              icon={cilDoubleQuoteSansRight}
              title="Blockquote"
              active={editor.isActive('blockquote')}
              onClick={() => editor.chain().focus().toggleBlockquote().run()}
            />
            <ToolbarButton
              icon={cilAlignLeft}
              title="Align left"
              active={editor.isActive({ textAlign: 'left' })}
              onClick={() => editor.chain().focus().setTextAlign('left').run()}
            />
            <ToolbarButton
              icon={cilAlignCenter}
              title="Align center"
              active={editor.isActive({ textAlign: 'center' })}
              onClick={() => editor.chain().focus().setTextAlign('center').run()}
            />
            <ToolbarButton
              icon={cilAlignRight}
              title="Align right"
              active={editor.isActive({ textAlign: 'right' })}
              onClick={() => editor.chain().focus().setTextAlign('right').run()}
            />
            <ToolbarButton
              icon={cilMenu}
              title="Justify"
              active={editor.isActive({ textAlign: 'justify' })}
              onClick={() => editor.chain().focus().setTextAlign('justify').run()}
            />
          </div>
        </ToolbarGroup>

        <ToolbarGroup label="Styles">
          <div className="editor-toolbar-styles-gallery">
            <button
              type="button"
              className={`editor-toolbar-style-preview ${!headingValue ? 'active' : ''}`}
              title="Normal"
              onClick={() => editor.chain().focus().setParagraph().run()}
            >
              Normal
            </button>
            <button
              type="button"
              className={`editor-toolbar-style-preview editor-toolbar-style-h1 ${headingValue === 1 ? 'active' : ''}`}
              title="Heading 1"
              onClick={() => editor.chain().focus().setHeading({ level: 1 }).run()}
            >
              Heading 1
            </button>
            <button
              type="button"
              className={`editor-toolbar-style-preview editor-toolbar-style-h2 ${headingValue === 2 ? 'active' : ''}`}
              title="Heading 2"
              onClick={() => editor.chain().focus().setHeading({ level: 2 }).run()}
            >
              Heading 2
            </button>
            <button
              type="button"
              className={`editor-toolbar-style-preview editor-toolbar-style-h3 ${headingValue === 3 ? 'active' : ''}`}
              title="Heading 3"
              onClick={() => editor.chain().focus().setHeading({ level: 3 }).run()}
            >
              Heading 3
            </button>
          </div>
        </ToolbarGroup>

        <ToolbarGroup label="Insert">
          <div className="editor-toolbar-icon-grid">
            <ToolbarButton
              icon={cilGrid}
              title="Insert table"
              onClick={() => editor.chain().focus().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run()}
            />
            <ToolbarButton icon={cilLink} title="Insert link" active={editor.isActive('link')} onClick={handleLink} />
            <ToolbarButton icon={cilVideo} title="Insert video" onClick={handleVideo} />
          </div>
        </ToolbarGroup>

        {editor.isActive('table') && (
          <ToolbarGroup label="Table">
            <ToolbarButton
              icon={cilPlus}
              label="Row"
              title="Add row"
              onClick={() => editor.chain().focus().addRowAfter().run()}
            />
            <ToolbarButton
              icon={cilPlus}
              label="Col"
              title="Add column"
              onClick={() => editor.chain().focus().addColumnAfter().run()}
            />
            <ToolbarButton
              icon={cilTrash}
              label="Table"
              title="Delete table"
              onClick={() => editor.chain().focus().deleteTable().run()}
            />
          </ToolbarGroup>
        )}
      </div>
    </div>
  );
}
