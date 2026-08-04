import StarterKit from '@tiptap/starter-kit';
import { TableKit } from '@tiptap/extension-table';
import { TextAlign } from '@tiptap/extension-text-align';
import { TextStyle, FontSize } from '@tiptap/extension-text-style';
import { Color } from '@tiptap/extension-color';
import { FontFamily } from '@tiptap/extension-font-family';
import { Youtube } from '@tiptap/extension-youtube';
import { Image } from '@tiptap/extension-image';
import { Indent } from './indentExtension';
import { FlaggedTerms } from './flaggedTermsExtension';
import { WordPaste } from './wordPasteExtension';

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
  // allowBase64 defaults to false, which silently drops <img src="data:...">
  // tags during generateJSON parsing — exactly what Word's captured images
  // are, once inlineImageSources() patches them in.
  Image.configure({ allowBase64: true }),
  Indent.configure({ types: ['paragraph', 'heading'] }),
  FlaggedTerms,
  WordPaste,
];
