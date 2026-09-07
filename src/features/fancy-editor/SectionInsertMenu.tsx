import { useState } from 'react';
import { FloatingMenu } from '@tiptap/react/menus';
import type { Editor } from '@tiptap/core';
import { AddCircleRegular } from '@fluentui/react-icons';
import { TemplateSectionPicker } from './TemplateSectionPicker';
import type { TemplateSectionType } from '../../core/tiptap-utils/templateSectionTypes';
import './SectionInsertMenu.css';

interface SectionInsertMenuProps {
  editor: Editor;
}

// The "+" gutter affordance from the reference UI — floats near the caret's
// block via Tiptap's own FloatingMenu (so positioning/scroll/resize is
// handled by floating-ui, not hand-rolled coordsAtPos math) and opens the
// same tile-grid picker used from the slash menu. This is purely a second
// trigger for insertTemplateSection, not a second insertion path.
export function SectionInsertMenu({ editor }: SectionInsertMenuProps) {
  const [open, setOpen] = useState(false);

  function handleSelect(type: TemplateSectionType) {
    editor.chain().focus().insertTemplateSection(type).run();
    setOpen(false);
  }

  return (
    <FloatingMenu
      editor={editor}
      shouldShow={({ state }) => {
        const { $from, empty } = state.selection;
        // Only when the caret sits at the very start of a block — matches
        // where the reference UI's "+" appears, and avoids popping up
        // mid-selection or mid-word.
        return empty && $from.parentOffset === 0;
      }}
      options={{ placement: 'left-start', offset: 8 }}
    >
      <div className="section-insert-menu">
        <button
          type="button"
          className="section-insert-menu-trigger"
          contentEditable={false}
          title="Insert a section"
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => setOpen((o) => !o)}
        >
          <AddCircleRegular />
        </button>
        {open && (
          <div className="section-insert-menu-popover">
            <TemplateSectionPicker onSelect={handleSelect} />
          </div>
        )}
      </div>
    </FloatingMenu>
  );
}
