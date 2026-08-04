import { useEffect, useState, type ReactNode } from 'react';
import type { Editor } from '@tiptap/react';
import { FluentProvider, Toolbar, ToolbarButton, Tooltip } from '@fluentui/react-components';
import {
  ArrowUndoRegular,
  ArrowRedoRegular,
  CutRegular,
  CopyRegular,
  FontIncreaseRegular,
  FontDecreaseRegular,
  TextCaseUppercaseRegular,
  EraserRegular,
  TextBoldRegular,
  TextItalicRegular,
  TextUnderlineRegular,
  TextStrikethroughRegular,
  TextSubscriptRegular,
  TextSuperscriptRegular,
  TextEffectsRegular,
  HighlightRegular,
  TextColorRegular,
  TextFieldRegular,
  TextBulletListRegular,
  TextNumberListLtrRegular,
  TextBulletListTreeRegular,
  TextIndentDecreaseRegular,
  TextIndentIncreaseRegular,
  ArrowSortRegular,
  TextParagraphRegular,
  TextAlignLeftRegular,
  TextAlignCenterRegular,
  TextAlignRightRegular,
  TextAlignJustifyRegular,
  TextLineSpacingRegular,
  PaintBucketRegular,
  BorderAllRegular,
  AddRegular,
  DeleteRegular,
  ChevronDown12Regular,
  type FluentIcon,
  Table32Light,
  Video32Light,
  Link32Light,
} from '@fluentui/react-icons';
import { ubsFluentTheme } from './ubsFluentTheme';
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
// as a real ribbon has tabs you're not currently on. Tab strip JSX below is
// currently commented out; kept paired with it for easy re-enabling.
// const RIBBON_TABS = ['File', 'Home', 'Insert', 'Draw', 'Design', 'Layout', 'References', 'Review', 'View'];

const NOT_IMPLEMENTED = 'Not implemented in this POC yet';

function RibbonGroup({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div className="editor-toolbar-group">
      <div className="editor-toolbar-group-body">{children}</div>
      <span className="editor-toolbar-group-label">{label}</span>
    </div>
  );
}

// A single icon, or an icon with a small decorative chevron for controls
// that look like a split-button/dropdown in Word (clicking anywhere still
// performs one action here — see the toolbar-parity note in the response
// this shipped with for which controls are fully wired vs visual-only).
function RibbonIcon({ icon: Icon, dropdown }: { icon: FluentIcon; dropdown?: boolean }) {
  if (!dropdown) {
    return <Icon />;
  }
  return (
    <span className="editor-toolbar-split-icon">
      <Icon />
      <ChevronDown12Regular className="editor-toolbar-split-chevron" />
    </span>
  );
}

function RibbonButton({
  icon,
  dropdown,
  title,
  disabled,
  onClick,
}: {
  icon: FluentIcon;
  dropdown?: boolean;
  title: string;
  disabled?: boolean;
  onClick?: () => void;
}) {
  return (
    <Tooltip content={disabled ? `${title} — ${NOT_IMPLEMENTED}` : title} relationship="label" withArrow>
      <ToolbarButton
        icon={<RibbonIcon icon={icon} dropdown={dropdown} />}
        onClick={onClick}
        appearance="subtle"
      />
    </Tooltip>
  );
}

function RibbonToggle({
  icon,
  dropdown,
  title,
  checked,
  onClick,
}: {
  icon: FluentIcon;
  dropdown?: boolean;
  title: string;
  checked: boolean;
  onClick: () => void;
}) {
  return (
    <Tooltip content={title} relationship="label" withArrow>
      {/* ToolbarToggleButton's public props don't include `checked` (it's
          meant to be driven via Toolbar's checkedValues/onCheckedValueChange),
          but our source of truth is editor.isActive(), recomputed every
          render — so a plain ToolbarButton with a conditional class is a
          more direct fit than fighting that controlled-group API. */}
      <ToolbarButton
        icon={<RibbonIcon icon={icon} dropdown={dropdown} />}
        onClick={onClick}
        appearance="subtle"
        className={checked ? 'editor-toolbar-active' : undefined}
      />
    </Tooltip>
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

  const headingValue = ([1, 2, 3, 4] as const).find((level) => editor.isActive('heading', { level }));
  const currentFontSize = Number(editor.getAttributes('textStyle').fontSize?.replace('px', '')) || 18;
  const currentColor = editor.getAttributes('textStyle').color || '#000000';

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

  function handleGrowFont() {
    editor.chain().focus().setFontSize(`${currentFontSize + 2}px`).run();
  }

  function handleShrinkFont() {
    editor.chain().focus().setFontSize(`${Math.max(8, currentFontSize - 2)}px`).run();
  }

  function handleClearFormatting() {
    editor.chain().focus().unsetAllMarks().clearNodes().run();
  }

  function handleCut() {
    editor.view.dom.focus();
    document.execCommand('cut');
  }

  function handleCopy() {
    editor.view.dom.focus();
    document.execCommand('copy');
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

  return (
    <FluentProvider theme={ubsFluentTheme} className="editor-toolbar-fluent-root">
      <div className="editor-toolbar-shell">
        {/* <div className="editor-toolbar-tabs">
          {RIBBON_TABS.map((tab) => (
            <button
              key={tab}
              type="button"
              className={tab === 'Home' ? 'active' : ''}
              disabled={tab !== 'Home'}
              title={tab === 'Home' ? undefined : NOT_IMPLEMENTED}
            >
              {tab}
            </button>
          ))}
        </div> */}

        <div className="editor-toolbar">
          <RibbonGroup label="History">
            <Toolbar size="small" className="editor-toolbar-row">
              <RibbonButton icon={ArrowUndoRegular} title="Undo" onClick={() => editor.chain().focus().undo().run()} />
              <RibbonButton icon={ArrowRedoRegular} title="Redo" onClick={() => editor.chain().focus().redo().run()} />
            </Toolbar>
            <Toolbar size="small" className="editor-toolbar-row">
              <RibbonButton icon={CutRegular} title="Cut" onClick={handleCut} />
              <RibbonButton icon={CopyRegular} title="Copy" onClick={handleCopy} />
            </Toolbar>
          </RibbonGroup>

          <RibbonGroup label="Font">
            <div className="editor-toolbar-row">
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
              <Toolbar size="small" className="editor-toolbar-row">
                <RibbonButton icon={FontIncreaseRegular} title="Grow font" onClick={handleGrowFont} />
                <RibbonButton icon={FontDecreaseRegular} title="Shrink font" onClick={handleShrinkFont} />
                <RibbonButton icon={TextCaseUppercaseRegular} dropdown title="Change case" disabled />
                <RibbonButton icon={EraserRegular} title="Clear formatting" onClick={handleClearFormatting} />
              </Toolbar>
            </div>
            <Toolbar size="small" className="editor-toolbar-row">
              <RibbonToggle
                icon={TextBoldRegular}
                title="Bold"
                checked={editor.isActive('bold')}
                onClick={() => editor.chain().focus().toggleBold().run()}
              />
              <RibbonToggle
                icon={TextItalicRegular}
                title="Italic"
                checked={editor.isActive('italic')}
                onClick={() => editor.chain().focus().toggleItalic().run()}
              />
              <RibbonToggle
                icon={TextUnderlineRegular}
                dropdown
                title="Underline"
                checked={editor.isActive('underline')}
                onClick={() => editor.chain().focus().toggleUnderline().run()}
              />
              <RibbonToggle
                icon={TextStrikethroughRegular}
                title="Strikethrough"
                checked={editor.isActive('strike')}
                onClick={() => editor.chain().focus().toggleStrike().run()}
              />
              <RibbonButton icon={TextSubscriptRegular} title="Subscript" disabled />
              <RibbonButton icon={TextSuperscriptRegular} title="Superscript" disabled />
              <RibbonButton icon={TextEffectsRegular} dropdown title="Text effects" disabled />
              <RibbonButton icon={HighlightRegular} dropdown title="Text highlight color" disabled />
              <label className="editor-toolbar-color-button" title="Font color">
                <TextColorRegular />
                <span className="editor-toolbar-color-swatch" style={{ backgroundColor: currentColor }} />
                <input
                  type="color"
                  value={currentColor}
                  onChange={(event) => editor.chain().focus().setColor(event.target.value).run()}
                />
              </label>
              <RibbonButton icon={TextFieldRegular} title="Enclose characters" disabled />
            </Toolbar>
          </RibbonGroup>

          <RibbonGroup label="Paragraph">
            <Toolbar size="small" className="editor-toolbar-row">
              <RibbonToggle
                icon={TextBulletListRegular}
                dropdown
                title="Bullets"
                checked={editor.isActive('bulletList')}
                onClick={() => editor.chain().focus().toggleBulletList().run()}
              />
              <RibbonToggle
                icon={TextNumberListLtrRegular}
                dropdown
                title="Numbering"
                checked={editor.isActive('orderedList')}
                onClick={() => editor.chain().focus().toggleOrderedList().run()}
              />
              <RibbonButton icon={TextBulletListTreeRegular} dropdown title="Multilevel list" disabled />
              <RibbonButton
                icon={TextIndentDecreaseRegular}
                title="Decrease indent"
                onClick={() => editor.chain().focus().outdent().run()}
              />
              <RibbonButton
                icon={TextIndentIncreaseRegular}
                title="Increase indent"
                onClick={() => editor.chain().focus().indent().run()}
              />
              <RibbonButton icon={ArrowSortRegular} title="Sort" disabled />
              <RibbonButton icon={TextParagraphRegular} title="Show/hide formatting marks" disabled />
            </Toolbar>
            <Toolbar size="small" className="editor-toolbar-row">
              <RibbonToggle
                icon={TextAlignLeftRegular}
                title="Align left"
                checked={editor.isActive({ textAlign: 'left' })}
                onClick={() => editor.chain().focus().setTextAlign('left').run()}
              />
              <RibbonToggle
                icon={TextAlignCenterRegular}
                title="Align center"
                checked={editor.isActive({ textAlign: 'center' })}
                onClick={() => editor.chain().focus().setTextAlign('center').run()}
              />
              <RibbonToggle
                icon={TextAlignRightRegular}
                title="Align right"
                checked={editor.isActive({ textAlign: 'right' })}
                onClick={() => editor.chain().focus().setTextAlign('right').run()}
              />
              <RibbonToggle
                icon={TextAlignJustifyRegular}
                title="Justify"
                checked={editor.isActive({ textAlign: 'justify' })}
                onClick={() => editor.chain().focus().setTextAlign('justify').run()}
              />
              <RibbonButton icon={TextLineSpacingRegular} dropdown title="Line and paragraph spacing" disabled />
              <RibbonButton icon={PaintBucketRegular} dropdown title="Shading" disabled />
              <RibbonButton icon={BorderAllRegular} dropdown title="Borders" disabled />
            </Toolbar>
          </RibbonGroup>

          <RibbonGroup label="Styles">
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
                className="editor-toolbar-style-preview"
                title="No Spacing — same as Normal in this POC (paragraph spacing isn't tracked)"
                onClick={() => editor.chain().focus().setParagraph().run()}
              >
                No Spacing
              </button>
              <button
                type="button"
                className={`editor-toolbar-style-preview editor-toolbar-style-h1 ${headingValue === 1 ? 'active' : ''}`}
                title="Heading 1"
                onClick={() => editor.chain().focus().setHeading({ level: 1 }).run()}
              >
                Heading
              </button>
              <button
                type="button"
                className={`editor-toolbar-style-preview editor-toolbar-style-h2 ${headingValue === 2 ? 'active' : ''}`}
                title="Heading 2"
                onClick={() => editor.chain().focus().setHeading({ level: 2 }).run()}
              >
                Heading 2
              </button>
              <button type="button" className="editor-toolbar-style-more" title={NOT_IMPLEMENTED} disabled>
                <ChevronDown12Regular />
              </button>
            </div>
          </RibbonGroup>

          <RibbonGroup label="Insert">
            <Toolbar size="small" className="editor-toolbar-row">
              <RibbonButton
                icon={Table32Light}
                title="Insert table"
                onClick={() => editor.chain().focus().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run()}
              />
              <RibbonButton icon={Link32Light} title="Insert link" onClick={handleLink} />
              <RibbonButton icon={Video32Light} title="Insert video" onClick={handleVideo} />
            </Toolbar>
          </RibbonGroup>

          {editor.isActive('table') && (
            <RibbonGroup label="Table">
              <Toolbar size="small" className="editor-toolbar-row">
                <Tooltip content="Add row" relationship="label" withArrow>
                  <ToolbarButton
                    icon={<AddRegular />}
                    appearance="subtle"
                    className="editor-toolbar-labeled-button"
                    onClick={() => editor.chain().focus().addRowAfter().run()}
                  >
                    Row
                  </ToolbarButton>
                </Tooltip>
                <Tooltip content="Add column" relationship="label" withArrow>
                  <ToolbarButton
                    icon={<AddRegular />}
                    appearance="subtle"
                    className="editor-toolbar-labeled-button"
                    onClick={() => editor.chain().focus().addColumnAfter().run()}
                  >
                    Col
                  </ToolbarButton>
                </Tooltip>
                <Tooltip content="Delete table" relationship="label" withArrow>
                  <ToolbarButton
                    icon={<DeleteRegular />}
                    appearance="subtle"
                    className="editor-toolbar-labeled-button"
                    onClick={() => editor.chain().focus().deleteTable().run()}
                  >
                    Table
                  </ToolbarButton>
                </Tooltip>
              </Toolbar>
            </RibbonGroup>
          )}
        </div>
      </div>
    </FluentProvider>
  );
}
