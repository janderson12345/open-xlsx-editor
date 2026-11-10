// Types for the agent API (agent.js).

/** A cell as an agent reads it. */
export interface AgentCell {
  /** "B2" (the range's sheet is in AgentRange.range). */
  ref: string;
  /** The value, or a formula's result: a number, string, boolean, an error's text ("#DIV/0!"), or null. */
  value: number | string | boolean | null;
  /** What the cell shows, its number format applied. */
  text: string;
  /** The formula, starting with "=", or null for a constant. */
  formula: string | null;
}
export interface AgentRange { range: string; rows: AgentCell[][]; }
export interface AgentSheet { name: string; kind: 'worksheet' | 'chartsheet'; hidden: boolean; used: string | null; tables: string[]; }
export interface AgentName { name: string; refersTo: string | null; table: boolean; columns?: string[]; }
export interface AgentNote { ref: string; kind: 'note' | 'thread'; author: string | null; text: string; replies: { author: string; text: string }[]; }
type OnSheet = { sheet?: string | number };

/** A workbook to read and edit without a browser. Ranges are what Excel's Name Box takes: "B2:D9", "Data!A:A", a
 *  defined name, a table ("Fruit") or a table reference ("Fruit[Qty]"). Without a sheet name they are on `sheet`, else
 *  the active sheet. */
export interface AgentWorkbook {
  sheets(): AgentSheet[];
  names(): AgentName[];
  /** A range's cells (the sheet's used range when none is given). */
  read(range?: string | null, options?: OnSheet): AgentRange;
  /** A range as a Markdown table, column letters across and row numbers down; `formulas` shows formulas. */
  markdown(range?: string | null, options?: OnSheet & { formulas?: boolean }): string;
  /** Cells matching `query` (Excel's * and ? wildcards, ~ to escape). */
  find(query: string, options?: OnSheet & { matchCase?: boolean; whole?: boolean; lookIn?: 'formulas' | 'values'; within?: 'workbook' | 'sheet' }): { ref: string; text: string }[];
  notes(options?: OnSheet): AgentNote[];
  /** Enter text as a person types it ({ "B2": "12", "C2": "=B2*2" }); data validation refuses unless `force`. */
  write(entries: Record<string, string | number | boolean>, options?: OnSheet & { force?: boolean }): { written: string[]; refused: { ref: string; error: string }[] };
  /** A block of values from `topLeft`, row by row, taken as they are (a string is never a formula). Returns the range. */
  setValues(topLeft: string, rows: (number | string | boolean | null)[][], options?: OnSheet): string | null;
  /** `at` is the 1-based row number; refusals throw with Excel's words. */
  insertRows(at: number, count?: number, options?: OnSheet): void;
  deleteRows(at: number, count?: number, options?: OnSheet): void;
  /** `at` is a column letter or 1-based number. */
  insertColumns(at: string | number, count?: number, options?: OnSheet): void;
  deleteColumns(at: string | number, count?: number, options?: OnSheet): void;
  /** A new worksheet at the end (or before `before`); returns its name. */
  addSheet(name?: string | null, options?: { before?: string | number }): string;
  renameSheet(sheet: string | number, name: string): void;
  deleteSheet(sheet: string | number): void;
  /** Work out every formula again (Ctrl+Alt+F9). */
  calculate(): void;
  save(): Uint8Array;
}

/** Opens an .xlsx or .xlsm for an agent. In Node, set up a DOMParser and XMLSerializer first. */
export function openWorkbook(bytes: Uint8Array | ArrayBuffer): AgentWorkbook;
