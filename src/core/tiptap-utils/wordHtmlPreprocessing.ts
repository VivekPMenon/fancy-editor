import juice from 'juice';
import { resolveWordVideoEmbeds } from './wordVideoEmbed';
import { reconstructWordLists } from './wordListReconstruction';
import { resolveWordImageAlignment } from './wordImageAlignment';
import { reconstructWordColumns } from './wordColumnReconstruction';

// Shared by htmlJsonConversion.ts (save/load) and wordPasteExtension.ts
// (live paste) — kept in its own file, independent of editorExtensions.ts,
// specifically to avoid a circular import (wordPasteExtension.ts is itself
// one of EDITOR_EXTENSIONS).
//
// Order matters: video-embed detection, image alignment, and list
// reconstruction all read resolved inline styles/hrefs, so juice() (which
// inlines stylesheet rules Word expressed via classes) runs first among
// them. Column reconstruction reads the mso-column-break style Word already
// puts inline on the <br>, so it's order-independent of the others — it
// just needs to run within the pass.
export function preprocessWordHtml(html: string): string {
  const resolved = juice(resolveWordVideoEmbeds(html));
  return reconstructWordColumns(reconstructWordLists(resolveWordImageAlignment(resolved)));
}
