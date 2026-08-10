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
  ChevronDown12Regular,
  type FluentIcon,
  Video32Light,
  Link32Light,
} from '@fluentui/react-icons';
import { ubsFluentTheme } from './ubsFluentTheme';
import { TableInsertPopover } from './TableInsertPopover';
import { RibbonOptionsPopover } from './RibbonOptionsPopover';
import { changeSelectionCase, sortSelection } from './textTransformCommands';
import './EditorToolbar.css';

interface EditorToolbarProps {
  editor: Editor;
}

// Deliberately a short, curated list — no need for a big font menu. The empty
// value means "no font-family mark", i.e. the editor's CSS default; that
// default IS Segoe UI (see App.css .tiptap-editor), so it's labelled as such
// so the applied default is clear rather than a vague "Default font".
const FONT_FAMILIES = [
  { label: 'Segoe UI (default)', value: '' },
  { label: 'Calibri', value: 'Calibri' },
  { label: 'Arial', value: 'Arial' },
  { label: 'Georgia', value: 'Georgia' },
  { label: 'Times New Roman', value: "'Times New Roman', serif" },
];

const FONT_SIZES = ['8', '9', '10', '10.5', '11', '12', '14', '15', '16', '18', '20', '22', '24', '26', '28', '36', '48', '72'];

// The editor's base font size (App.css .tiptap-editor) — shown in the toolbar
// when a run has no size of its own, so the applied default is clear.
const EDITOR_DEFAULT_FONT_SIZE = 15;

// Imported OOXML sizes are in points ("11pt"); the toolbar works in px. Strip
// either unit so an imported size still lines up with the dropdown / stepper.
function fontSizeNumber(value: string | undefined): number | undefined {
  if (!value) {
    return undefined;
  }
  const n = Number(value.replace(/px|pt/i, ''));
  return Number.isFinite(n) ? n : undefined;
}

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

  const [showMarks, setShowMarks] = useState(false);

  const headingValue = ([1, 2, 3, 4] as const).find((level) => editor.isActive('heading', { level }));
  const activeBlockType = headingValue ? 'heading' : 'paragraph';
  const currentFontSize = fontSizeNumber(editor.getAttributes('textStyle').fontSize) ?? EDITOR_DEFAULT_FONT_SIZE;
  const currentColor = editor.getAttributes('textStyle').color || '#000000';
  const currentHighlight = editor.getAttributes('highlight').color || '#ffff00';
  const currentShading = editor.getAttributes(activeBlockType).shading || '#ffffff';

  function handleToggleShowMarks() {
    const next = !showMarks;
    setShowMarks(next);
    editor.view.dom.classList.toggle('show-formatting-marks', next);
  }

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

  // Align left/center/right double as image alignment when an image is
  // selected — text alignment doesn't apply to a block-level image (there's
  // no paragraph there to carry it), so the same three buttons drive
  // whichever is actually selected instead of needing separate controls.
  function isAlignActive(align: 'left' | 'center' | 'right') {
    if (editor.isActive('image')) {
      return editor.getAttributes('image').align === align;
    }
    return editor.isActive({ textAlign: align });
  }

  function handleAlign(align: 'left' | 'center' | 'right') {
    if (editor.isActive('image')) {
      editor.chain().focus().setImageAlign(align).run();
    } else {
      editor.chain().focus().setTextAlign(align).run();
    }
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
                value={String(fontSizeNumber(editor.getAttributes('textStyle').fontSize) ?? EDITOR_DEFAULT_FONT_SIZE)}
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
                <RibbonOptionsPopover
                  icon={TextCaseUppercaseRegular}
                  title="Change case"
                  options={[
                    { label: 'Sentence case.', onClick: () => changeSelectionCase(editor, 'sentence') },
                    { label: 'lowercase', onClick: () => changeSelectionCase(editor, 'lower') },
                    { label: 'UPPERCASE', onClick: () => changeSelectionCase(editor, 'upper') },
                    { label: 'Capitalize Each Word', onClick: () => changeSelectionCase(editor, 'title') },
                  ]}
                />
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
              <RibbonToggle
                icon={TextSubscriptRegular}
                title="Subscript"
                checked={editor.isActive('subscript')}
                onClick={() => editor.chain().focus().toggleSubscript().run()}
              />
              <RibbonToggle
                icon={TextSuperscriptRegular}
                title="Superscript"
                checked={editor.isActive('superscript')}
                onClick={() => editor.chain().focus().toggleSuperscript().run()}
              />
              {/* WordArt-style outline/shadow/glow/reflection — a large preset
                  gallery with little payoff for a business-document editor;
                  left out of scope rather than half-built. */}
              <RibbonButton icon={TextEffectsRegular} dropdown title="Text effects" disabled />
              <label className="editor-toolbar-color-button" title="Text highlight color">
                <HighlightRegular />
                <span className="editor-toolbar-color-swatch" style={{ backgroundColor: currentHighlight }} />
                <input
                  type="color"
                  value={currentHighlight}
                  onChange={(event) => editor.chain().focus().setHighlight({ color: event.target.value }).run()}
                />
              </label>
              <label className="editor-toolbar-color-button" title="Font color">
                <TextColorRegular />
                <span className="editor-toolbar-color-swatch" style={{ backgroundColor: currentColor }} />
                <input
                  type="color"
                  value={currentColor}
                  onChange={(event) => editor.chain().focus().setColor(event.target.value).run()}
                />
              </label>
              <RibbonToggle
                icon={TextFieldRegular}
                title="Enclose characters"
                checked={editor.isActive('encloseCharacters')}
                onClick={() => editor.chain().focus().toggleEncloseCharacters().run()}
              />
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
              {/* Word's true multilevel list assigns a distinct numbering
                  style per depth (1, a, i…) — simplified here to the same
                  ordered-list toggle as Numbering, with depth-based
                  numbering styles applied globally via CSS (nesting itself
                  already works via Tab/Shift-Tab, see indentExtension.ts). */}
              <RibbonToggle
                icon={TextBulletListTreeRegular}
                dropdown
                title="Multilevel list"
                checked={editor.isActive('orderedList')}
                onClick={() => editor.chain().focus().toggleOrderedList().run()}
              />
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
              <RibbonOptionsPopover
                icon={ArrowSortRegular}
                title="Sort"
                options={[
                  { label: 'Ascending (A to Z)', onClick: () => sortSelection(editor, 'asc') },
                  { label: 'Descending (Z to A)', onClick: () => sortSelection(editor, 'desc') },
                ]}
              />
              <RibbonToggle
                icon={TextParagraphRegular}
                title="Show/hide formatting marks"
                checked={showMarks}
                onClick={handleToggleShowMarks}
              />
            </Toolbar>
            <Toolbar size="small" className="editor-toolbar-row">
              <RibbonToggle
                icon={TextAlignLeftRegular}
                title="Align left"
                checked={isAlignActive('left')}
                onClick={() => handleAlign('left')}
              />
              <RibbonToggle
                icon={TextAlignCenterRegular}
                title="Align center"
                checked={isAlignActive('center')}
                onClick={() => handleAlign('center')}
              />
              <RibbonToggle
                icon={TextAlignRightRegular}
                title="Align right"
                checked={isAlignActive('right')}
                onClick={() => handleAlign('right')}
              />
              <RibbonToggle
                icon={TextAlignJustifyRegular}
                title="Justify"
                checked={editor.isActive({ textAlign: 'justify' })}
                onClick={() => editor.chain().focus().setTextAlign('justify').run()}
              />
              <RibbonOptionsPopover
                icon={TextLineSpacingRegular}
                title="Line and paragraph spacing"
                options={[
                  { label: '1.0', onClick: () => editor.chain().focus().setLineHeight('1').run() },
                  { label: '1.15', onClick: () => editor.chain().focus().setLineHeight('1.15').run() },
                  { label: '1.5', onClick: () => editor.chain().focus().setLineHeight('1.5').run() },
                  { label: '2.0', onClick: () => editor.chain().focus().setLineHeight('2').run() },
                  { label: 'Remove custom spacing', onClick: () => editor.chain().focus().unsetLineHeight().run() },
                ]}
              />
              <label className="editor-toolbar-color-button" title="Shading">
                <PaintBucketRegular />
                <span className="editor-toolbar-color-swatch" style={{ backgroundColor: currentShading }} />
                <input
                  type="color"
                  value={currentShading}
                  onChange={(event) => editor.chain().focus().setShading(event.target.value).run()}
                />
              </label>
              {/* Simplified to a single uniform border toggle rather than
                  Word's full per-side style/width/color control. */}
              <RibbonToggle
                icon={BorderAllRegular}
                dropdown
                title="Borders"
                checked={!!editor.getAttributes(activeBlockType).border}
                onClick={() => editor.chain().focus().toggleParagraphBorder().run()}
              />
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
              <TableInsertPopover editor={editor} />
              <RibbonButton icon={Link32Light} title="Insert link" onClick={handleLink} />
              <RibbonButton icon={Video32Light} title="Insert video" onClick={handleVideo} />
            </Toolbar>
          </RibbonGroup>

        </div>
      </div>
    </FluentProvider>
  );
}
