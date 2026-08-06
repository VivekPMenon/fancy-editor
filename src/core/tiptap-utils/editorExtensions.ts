import StarterKit from '@tiptap/starter-kit';
import { TableKit } from '@tiptap/extension-table';
import { TextAlign } from '@tiptap/extension-text-align';
import { TextStyle, FontSize } from '@tiptap/extension-text-style';
import { Color } from '@tiptap/extension-color';
import { FontFamily } from '@tiptap/extension-font-family';
import { Youtube } from '@tiptap/extension-youtube';
import { Image } from '@tiptap/extension-image';
import { Subscript } from '@tiptap/extension-subscript';
import { Superscript } from '@tiptap/extension-superscript';
import { Highlight } from '@tiptap/extension-highlight';
import { Indent } from './indentExtension';
import { FlaggedTerms } from './flaggedTermsExtension';
import { WordPaste } from './wordPasteExtension';
import { ParagraphStyle } from './paragraphStyleExtension';
import { EncloseCharacters } from './encloseCharactersExtension';
import { SlashCommand } from './slashCommandExtension';
import { MentionExtension } from './mentionExtension';
import { TickerCard } from './tickerCardExtension';
import { ImageAlignment } from './imageAlignmentExtension';
import { ImageSpacing } from './imageSpacingExtension';

// Single source of truth for the editor's schema, shared by the live Tiptap
// editor (FancyEditorTab) and the HTML<->JSON conversion utilities. They
// have to stay in sync: generateJSON/generateHTML only understand node/mark
// types covered by this list, so a mismatch here silently drops content.
export const EDITOR_EXTENSIONS = [
  StarterKit.configure({ link: { openOnClick: false } }),
  TableKit.configure({ table: { resizable: true } }),
  TextAlign.configure({ types: ['heading', 'paragraph'] }),
  TextStyle,
  // Default types is ['textStyle'] (inline spans only) — extended to
  // heading/paragraph so block-level formatting (e.g. Word's `h1 { color;
  // font-size; font-family }` stylesheet rules, resolved onto the element by
  // juice() in htmlToTiptapJson) has a node attribute to land in, not just
  // character-run marks.
  FontSize.configure({ types: ['textStyle', 'heading', 'paragraph'] }),
  Color.configure({ types: ['textStyle', 'heading', 'paragraph'] }),
  FontFamily.configure({ types: ['textStyle', 'heading', 'paragraph'] }),
  Youtube.configure({ width: 480, height: 270 }),
  Image.configure({
    // allowBase64 defaults to false, which silently drops <img src="data:...">
    // tags during generateJSON parsing — exactly what Word's captured images
    // are, once inlineImageSources() patches them in.
    allowBase64: true,
    // Built into the official extension itself (just disabled by default) —
    // no third-party package or custom node view needed. Locks aspect ratio
    // so corner-dragging can't produce a squashed/stretched image.
    resize: {
      enabled: true,
      directions: ['top-left', 'top-right', 'bottom-left', 'bottom-right'],
      minWidth: 40,
      minHeight: 40,
      alwaysPreserveAspectRatio: true,
    },
  }),
  ImageAlignment,
  ImageSpacing,
  Indent.configure({ types: ['paragraph', 'heading'] }),
  FlaggedTerms,
  WordPaste,
  Subscript,
  Superscript,
  // multicolor: true — without it, Highlight only supports a single default
  // yellow, ignoring whatever color the toolbar picker sends.
  Highlight.configure({ multicolor: true }),
  ParagraphStyle.configure({ types: ['paragraph', 'heading'] }),
  EncloseCharacters,
  SlashCommand,
  MentionExtension,
  TickerCard,
];
