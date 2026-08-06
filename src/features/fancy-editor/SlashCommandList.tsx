import { forwardRef, useEffect, useImperativeHandle, useState } from 'react';
import type { SlashCommandItem } from '../../core/tiptap-utils/slashCommandExtension';
import './SlashCommandList.css';

interface SlashCommandListProps {
  items: SlashCommandItem[];
  command: (item: SlashCommandItem) => void;
}

export interface SlashCommandListRef {
  onKeyDown: (props: { event: KeyboardEvent }) => boolean;
}

// The dropdown rendered by slashCommandExtension's Suggestion plugin —
// mounted imperatively via ReactRenderer, so keyboard nav is exposed through
// a ref rather than normal DOM event handlers (the actual keydown happens on
// the editor, not this popup).
export const SlashCommandList = forwardRef<SlashCommandListRef, SlashCommandListProps>(({ items, command }, ref) => {
  const [selectedIndex, setSelectedIndex] = useState(0);

  useEffect(() => setSelectedIndex(0), [items]);

  function selectItem(index: number) {
    const item = items[index];
    if (item) {
      command(item);
    }
  }

  useImperativeHandle(
    ref,
    () => ({
      onKeyDown({ event }) {
        if (event.key === 'ArrowUp') {
          setSelectedIndex((index) => (index + items.length - 1) % items.length);
          return true;
        }
        if (event.key === 'ArrowDown') {
          setSelectedIndex((index) => (index + 1) % items.length);
          return true;
        }
        if (event.key === 'Enter') {
          selectItem(selectedIndex);
          return true;
        }
        return false;
      },
    }),
    [items, selectedIndex],
  );

  if (items.length === 0) {
    return (
      <div className="slash-command-menu">
        <p className="slash-command-empty">No matching commands</p>
      </div>
    );
  }

  return (
    <div className="slash-command-menu">
      {items.map((item, index) => (
        <button
          key={item.title}
          type="button"
          className={`slash-command-item ${index === selectedIndex ? 'active' : ''}`}
          onClick={() => selectItem(index)}
          onMouseEnter={() => setSelectedIndex(index)}
        >
          <item.icon />
          <span className="slash-command-item-text">
            <span className="slash-command-item-title">{item.title}</span>
            <span className="slash-command-item-description">{item.description}</span>
          </span>
        </button>
      ))}
    </div>
  );
});

SlashCommandList.displayName = 'SlashCommandList';
