import { NodeViewWrapper, NodeViewContent, type ReactNodeViewProps } from '@tiptap/react';
import { DismissRegular } from '@fluentui/react-icons';
import { getTemplateSectionConfig } from '../../core/tiptap-utils/templateSectionTypes';
import { SectionValuePicker } from './SectionValuePicker';
import './TemplateSectionView.css';

// The live editing view for a templateSection node: a header (icon + label +
// the type-specific control — ticker lookup, country/sector/etc. select, or
// nothing for Rich Text) above a real editable body (NodeViewContent, so
// normal rich-text editing — bold, links, even nested lists — works inside
// it exactly like any other block).
export function TemplateSectionView({ node, updateAttributes, deleteNode, selected }: ReactNodeViewProps) {
  const config = getTemplateSectionConfig(node.attrs.sectionType as string);
  const selection = (node.attrs.selection as string) ?? '';
  const Icon = config.icon;

  return (
    <NodeViewWrapper className={`template-section-view ${selected ? 'selected' : ''}`} data-drag-handle>
      <div className="template-section-view-header">
        <Icon className="template-section-view-icon" />
        <span className="template-section-view-label">{config.label}</span>

        {config.control !== 'none' && (
          <SectionValuePicker
            value={selection}
            options={config.options ?? []}
            placeholder={config.controlPlaceholder}
            onChange={(value) => updateAttributes({ selection: value })}
          />
        )}

        <button
          type="button"
          className="template-section-view-remove"
          contentEditable={false}
          title="Remove section"
          onClick={() => deleteNode()}
        >
          <DismissRegular />
        </button>
      </div>

      {/* No placeholder-text styling here — this app doesn't use Tiptap's
          Placeholder extension anywhere else either, so an empty body reads
          the same as any other empty paragraph in the document. */}
      <NodeViewContent className="template-section-view-body" />
    </NodeViewWrapper>
  );
}
