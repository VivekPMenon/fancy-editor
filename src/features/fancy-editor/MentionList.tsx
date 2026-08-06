import { forwardRef, useEffect, useImperativeHandle, useState } from 'react';
import type { MentionablePerson } from '../../core/mentionablePeople';
import './MentionList.css';

interface MentionListProps {
  items: MentionablePerson[];
  command: (item: MentionablePerson) => void;
}

export interface MentionListRef {
  onKeyDown: (props: { event: KeyboardEvent }) => boolean;
}

// Same forwardRef + imperative-keyboard-nav pattern as SlashCommandList —
// both are mounted the same way, via @tiptap/suggestion's ReactRenderer.
export const MentionList = forwardRef<MentionListRef, MentionListProps>(({ items, command }, ref) => {
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
      <div className="mention-list">
        <p className="mention-list-empty">No matches</p>
      </div>
    );
  }

  return (
    <div className="mention-list">
      {items.map((item, index) => (
        <button
          key={item.id}
          type="button"
          className={`mention-list-item ${index === selectedIndex ? 'active' : ''}`}
          onClick={() => selectItem(index)}
          onMouseEnter={() => setSelectedIndex(index)}
        >
          <span className="mention-list-item-avatar">{item.name.charAt(0)}</span>
          <span className="mention-list-item-text">
            <span className="mention-list-item-name">{item.name}</span>
            <span className="mention-list-item-role">{item.role}</span>
          </span>
        </button>
      ))}
    </div>
  );
});

MentionList.displayName = 'MentionList';
