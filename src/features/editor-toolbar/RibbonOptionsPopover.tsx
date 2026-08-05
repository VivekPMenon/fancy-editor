import { useState } from 'react';
import { Popover, PopoverTrigger, PopoverSurface, ToolbarButton, Tooltip } from '@fluentui/react-components';
import type { FluentIcon } from '@fluentui/react-icons';

interface RibbonOption {
  label: string;
  onClick: () => void;
}

// A small "click icon, pick one of a few options" dropdown — reused for
// Change Case, Sort, and Line Spacing. Same Popover approach as
// TableInsertPopover, just with a plain option list instead of custom
// content.
export function RibbonOptionsPopover({
  icon: Icon,
  title,
  options,
}: {
  icon: FluentIcon;
  title: string;
  options: RibbonOption[];
}) {
  const [open, setOpen] = useState(false);

  return (
    <Popover open={open} onOpenChange={(_, data) => setOpen(data.open)} positioning="below-start">
      <PopoverTrigger disableButtonEnhancement>
        <Tooltip content={title} relationship="label" withArrow>
          <ToolbarButton icon={<Icon />} appearance="subtle" />
        </Tooltip>
      </PopoverTrigger>
      <PopoverSurface className="ribbon-options-popover">
        {options.map((option) => (
          <button
            key={option.label}
            type="button"
            className="ribbon-options-item"
            onClick={() => {
              option.onClick();
              setOpen(false);
            }}
          >
            {option.label}
          </button>
        ))}
      </PopoverSurface>
    </Popover>
  );
}
