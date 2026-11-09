// Types for the public API (hand-written; the code is plain JavaScript).

export interface Cell {
  /** n number, s string, b boolean, e error, str formula text, d ISO date. */
  t: 'n' | 's' | 'b' | 'e' | 'str' | 'd';
  /** The value; null for a cell that only carries a style. */
  v: number | string | boolean | null;
  /** Index into the workbook's cellXfs. */
  s: number;
  /** Formula text (without "="), '' for a shared formula's later cells, null for none. */
  f: string | null;
}

export interface Range { r0: number; c0: number; r1: number; c1: number; }

export interface SheetModel {
  /** Cells by 0-based row, then 0-based column. */
  cells: Map<number, Map<number, Cell>>;
  rows: Map<number, { ht: number | null; hidden: boolean; customHeight: boolean }>;
  cols: { min: number; max: number; width: number | null; hidden: boolean; style: number }[];
  merges: Range[];
  format: { defaultRowHeight: number; defaultColWidth: number | null; baseColWidth: number | null };
  showGridLines: boolean;
  maxRow: number;
  maxCol: number;
}

export interface SheetInfo {
  name: string;
  sheetId: number;
  state: 'visible' | 'hidden' | 'veryHidden' | string;
  rId: string | null;
  path: string | null;
  kind: 'worksheet' | 'chartsheet' | 'dialogsheet' | string | null;
}

export declare class Workbook {
  static fromBytes(bytes: Uint8Array | ArrayBuffer): Workbook;
  readonly sheets: SheetInfo[];
  readonly sharedStrings: string[];
  readonly date1904: boolean;
  readonly activeSheet: number;
  sheet(index: number): SheetModel;
  /** Set a cell ("B2" or { r, c }, 0-based) to a number, boolean, string (written inline) or null. */
  setCell(index: number, ref: string | { r: number; c: number }, value: number | boolean | string | null): void;
  toBytes(): Uint8Array;
}

export interface GridHandle {
  el: HTMLElement;
  window(): Range;
  scrollToCell(r: number, c: number): void;
}

/** The selection: the active cell (ar, ac) and the range r0..r1, c0..c1, 0-based. */
export interface Selection { ar: number; ac: number; r0: number; c0: number; r1: number; c1: number; }

export type CellRef = string | { r: number; c: number };

export interface XlsxEditor {
  open(bytes: Uint8Array | ArrayBuffer, fileName?: string | null): XlsxEditor;
  save(): Uint8Array;
  /** Select a cell. */
  select(ref: CellRef): void;
  /** Enter text into a cell as if typed ("12", "50%", "TRUE", "'007"); false if refused. */
  enter(ref: CellRef, text: string): boolean;
  /** Set a cell's value; false if refused (read-only, or a formula cell). */
  setCell(ref: CellRef, value: number | boolean | string | null): boolean;
  undo(): boolean;
  redo(): boolean;
  /** Run a named command: select, extendTo, move, page, home, end, enter, setValue, clear, undo, redo. */
  command(name: string, ...args: unknown[]): unknown;
  readonly selection: Selection | null;
  readonly readOnly: boolean;
  setReadOnly(flag: boolean): void;
  /** Build the printed pages of the sheet shown; returns the page count. */
  preparePrint(): number;
  readonly workbook: Workbook | null;
  readonly grid: GridHandle | null;
  readonly sheetIndex: number;
  readonly fileName: string | null;
  setTheme(theme: 'light' | 'dark' | null): void;
  destroy(): void;
}

export declare function mountXlsxEditor(
  host: HTMLElement,
  options?: { bytes?: Uint8Array | ArrayBuffer; fileName?: string; theme?: 'light' | 'dark'; readOnly?: boolean; onChange?: () => void },
): XlsxEditor;
