import { useState } from 'react';
import type { Editor } from '@tiptap/react';
import { Popover, PopoverTrigger, PopoverSurface, ToolbarButton, Tooltip } from '@fluentui/react-components';
import {
  Table32Light,
  TableSimpleRegular,
  DrawTextRegular,
  TableSwitchRegular,
  TableRegular,
  TableStackAboveRegular,
} from '@fluentui/react-icons';

const GRID_MAX_ROWS = 6;
const GRID_MAX_COLS = 8;
const NOT_IMPLEMENTED = 'Not implemented in this POC yet';

interface HoverCell {
  row: number;
  col: number;
}

// Word's "hover a grid, click to insert" table picker — the actual point of
// this component is demonstrating a popover triggered off a toolbar icon,
// not table-insertion itself (that part's a one-line Tiptap command).
export function TableInsertPopover({ editor }: { editor: Editor }) {
  const [open, setOpen] = useState(false);
  const [hover, setHover] = useState<HoverCell | null>(null);

  function close() {
    setOpen(false);
    setHover(null);
  }

  function insertTable(rows: number, cols: number) {
    editor.chain().focus().insertTable({ rows, cols, withHeaderRow: true }).run();
    close();
  }

  function handleInsertTableDialog() {
    const rowsInput = window.prompt('Number of rows', '3');
    if (!rowsInput) {
      return;
    }
    const colsInput = window.prompt('Number of columns', '3');
    if (!colsInput) {
      return;
    }
    const rows = Math.max(1, Math.round(Number(rowsInput)) || 3);
    const cols = Math.max(1, Math.round(Number(colsInput)) || 3);
    insertTable(rows, cols);
  }

  // Builds table HTML directly and feeds it through the normal HTML->schema
  // parser instead of manually walking positions after insertTable() — far
  // simpler than hand-computing where each cell landed.
  function handleConvertTextToTable() {
    const { from, to } = editor.state.selection;
    if (from === to) {
      window.alert('Select some text first — one row per line, cells separated by tabs or commas.');
      return;
    }
    const text = editor.state.doc.textBetween(from, to, '\n');
    const rows = text
      .split('\n')
      .map((line) => line.split(/\t|,/).map((cell) => cell.trim()))
      .filter((row) => row.some(Boolean));
    if (rows.length === 0) {
      return;
    }
    const cols = Math.max(...rows.map((row) => row.length));
    const html = `<table><tbody>${rows
      .map((row) => `<tr>${Array.from({ length: cols }, (_, i) => `<td>${row[i] ?? ''}</td>`).join('')}</tr>`)
      .join('')}</tbody></table>`;
    editor.chain().focus().deleteSelection().insertContent(html).run();
    close();
  }

  return (
    <Popover
      open={open}
      onOpenChange={(_, data) => {
        setOpen(data.open);
        if (!data.open) {
          setHover(null);
        }
      }}
      positioning="below-start"
    >
      <PopoverTrigger disableButtonEnhancement>
        <Tooltip content="Insert table" relationship="label" withArrow>
          <ToolbarButton icon={<Table32Light />} appearance="subtle" />
        </Tooltip>
      </PopoverTrigger>
      <PopoverSurface className="table-insert-popover">
        <p className="table-insert-popover-label">{hover ? `${hover.col + 1} x ${hover.row + 1} Table` : 'Insert Table'}</p>
        <div
          className="table-insert-grid"
          style={{ gridTemplateColumns: `repeat(${GRID_MAX_COLS}, 1fr)` }}
          onMouseLeave={() => setHover(null)}
        >
          {Array.from({ length: GRID_MAX_ROWS }).map((_, row) =>
            Array.from({ length: GRID_MAX_COLS }).map((_, col) => (
              <button
                key={`${row}-${col}`}
                type="button"
                className={`table-insert-cell ${hover && row <= hover.row && col <= hover.col ? 'active' : ''}`}
                onMouseEnter={() => setHover({ row, col })}
                onClick={() => insertTable(row + 1, col + 1)}
              />
            )),
          )}
        </div>

        <div className="table-insert-divider" />

        <button type="button" className="table-insert-menu-item" onClick={handleInsertTableDialog}>
          <TableSimpleRegular /> Insert Table…
        </button>
        <Tooltip content={NOT_IMPLEMENTED} relationship="label" withArrow>
          <button type="button" className="table-insert-menu-item" disabled>
            <DrawTextRegular /> Draw Table
          </button>
        </Tooltip>
        <button type="button" className="table-insert-menu-item" onClick={handleConvertTextToTable}>
          <TableSwitchRegular /> Convert Text to Table…
        </button>
        <Tooltip content={NOT_IMPLEMENTED} relationship="label" withArrow>
          <button type="button" className="table-insert-menu-item" disabled>
            <TableRegular /> Excel Spreadsheet
          </button>
        </Tooltip>
        <div className="table-insert-divider" />
        <p className="table-insert-popover-label">Quick Tables</p>
        <button
          type="button"
          className="table-insert-menu-item"
          onClick={() => insertTable(3, 3)}
        >
          <TableStackAboveRegular /> Simple 3×3
        </button>
        <button
          type="button"
          className="table-insert-menu-item"
          onClick={() => insertTable(5, 2)}
        >
          <TableStackAboveRegular /> Two Columns, 5 Rows
        </button>
      </PopoverSurface>
    </Popover>
  );
}
