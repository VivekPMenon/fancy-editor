import { TEMPLATE_SECTION_TYPES, type TemplateSectionType } from '../../core/tiptap-utils/templateSectionTypes';
import './TemplateSectionPicker.css';

interface TemplateSectionPickerProps {
  onSelect: (type: TemplateSectionType) => void;
}

// The grid-of-tiles template picker — one tile per templateSectionTypes.ts
// entry, so adding a 16th template needs no change here.
export function TemplateSectionPicker({ onSelect }: TemplateSectionPickerProps) {
  return (
    <div className="template-section-picker" role="menu">
      {TEMPLATE_SECTION_TYPES.map((config) => {
        const Icon = config.icon;
        return (
          <button
            key={config.type}
            type="button"
            className="template-section-picker-tile"
            role="menuitem"
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => onSelect(config.type)}
          >
            <Icon className="template-section-picker-tile-icon" />
            <span>{config.label}</span>
          </button>
        );
      })}
    </div>
  );
}
