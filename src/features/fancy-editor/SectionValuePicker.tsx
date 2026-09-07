import { useState } from 'react';
import './SectionValuePicker.css';

interface SectionValuePickerProps {
  value: string;
  options: string[];
  placeholder?: string;
  onChange: (value: string) => void;
}

const MAX_SUGGESTIONS = 8;

// One typeahead control for both 'lookup' (large lists — tickers) and
// 'select' (short fixed lists — sectors, countries, …) section kinds. A
// native <select> would be fine for the short lists but not for tickers, so
// this filter-as-you-type box covers both without two separate controls.
export function SectionValuePicker({ value, options, placeholder, onChange }: SectionValuePickerProps) {
  const [query, setQuery] = useState(value);
  const [open, setOpen] = useState(false);

  const suggestions = options
    .filter((option) => option.toLowerCase().includes(query.trim().toLowerCase()))
    .slice(0, MAX_SUGGESTIONS);

  function handleSelect(option: string) {
    onChange(option);
    setQuery(option);
    setOpen(false);
  }

  return (
    <div className="section-value-picker">
      <input
        type="text"
        className="section-value-picker-input"
        value={query}
        placeholder={placeholder}
        onChange={(event) => {
          setQuery(event.target.value);
          setOpen(true);
        }}
        onFocus={() => setOpen(true)}
        onBlur={() => {
          // Let a suggestion's onMouseDown fire before the list disappears.
          setTimeout(() => setOpen(false), 120);
        }}
        contentEditable={false}
      />
      {open && suggestions.length > 0 && (
        <ul className="section-value-picker-suggestions">
          {suggestions.map((option) => (
            <li key={option}>
              <button type="button" onMouseDown={(e) => e.preventDefault()} onClick={() => handleSelect(option)}>
                {option}
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
