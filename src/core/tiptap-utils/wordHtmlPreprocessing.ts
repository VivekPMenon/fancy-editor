import juice from 'juice';
import { resolveWordVideoEmbeds } from './wordVideoEmbed';
import { reconstructWordLists } from './wordListReconstruction';

// Shared by htmlJsonConversion.ts (save/load) and wordPasteExtension.ts
// (live paste) — kept in its own file, independent of editorExtensions.ts,
// specifically to avoid a circular import (wordPasteExtension.ts is itself
// one of EDITOR_EXTENSIONS).
//
// Order matters: video-embed detection and list reconstruction both read
// resolved inline styles/hrefs, so juice() (which inlines stylesheet rules
// Word expressed via classes) runs in between.
export function preprocessWordHtml(html: string): string {
  return reconstructWordLists(juice(resolveWordVideoEmbeds(html)));
}
