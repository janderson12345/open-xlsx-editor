// Types for the public API (hand-written; the code is plain JavaScript).
import type { AgentWorkbook } from './agent.js';
export type { AgentWorkbook, AgentCell, AgentRange, AgentSheet, AgentName, AgentNote } from './agent.js';
export { openWorkbook } from './agent.js';
/** Why an encrypted workbook didn't open: 'needed' (no password), 'wrong', 'unsupported' (an older encryption), 'damaged'. */
export declare class PasswordError extends Error { reason: 'needed' | 'wrong' | 'unsupported' | 'damaged'; }

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
  /** The row and column headings shown (View > Headings). */
  showHeadings: boolean;
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
  /** One saved with a password needs `password` (a PasswordError otherwise). */
  static fromBytes(bytes: Uint8Array | ArrayBuffer, options?: { password?: string }): Workbook;
  setPassword(password: string | null): void;
  readonly encrypted: boolean;
  readonly sheets: SheetInfo[];
  readonly sharedStrings: string[];
  readonly date1904: boolean;
  readonly activeSheet: number;
  sheet(index: number): SheetModel;
  /** Whether a sheet can be shown: a worksheet whose tab isn't hidden. */
  canShow(index: number): boolean;
  /** A sheet's tab colour as '#RRGGBB', or null. */
  tabColor(index: number): string | null;
  /** Sheets changed as Excel changes them; a refusal throws an Error with Excel's message. */
  addSheet(at?: number, name?: string): number;
  renameSheet(index: number, name: string): void;
  moveSheet(from: number, to: number): void;
  setSheetHidden(index: number, hidden: boolean): void;
  setTabColor(index: number, color: { rgb: string } | { theme: number; tint?: number } | null): void;
  deleteSheet(index: number): void;
  /** Set a cell ("B2" or { r, c }, 0-based) to a number, boolean, string (written inline) or null. */
  setCell(index: number, ref: string | { r: number; c: number }, value: number | boolean | string | null): { circular: unknown[] };
  /** Give a cell a formula as the file has it ("SUM(A1:A3)", without "="); the formulas that read it are worked out again. */
  setFormula(index: number, ref: string | { r: number; c: number }, formula: string): { circular: unknown[] };
  /** Work out every formula again. */
  calculate(): { circular: unknown[] };
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
  /** A workbook saved with a password opens with `password`; without it the editor asks, and a promise is returned
   *  (the editor once open, or null if cancelled or the password was wrong). */
  open(bytes: Uint8Array | ArrayBuffer, fileName?: string | null, options?: { password?: string }): XlsxEditor | Promise<XlsxEditor | null>;
  /** Save with this password from now on (Encrypt with Password), or without one (null). */
  setPassword(password: string | null): void;
  /** Whether the workbook is saved with a password. */
  readonly encrypted: boolean;
  save(): Uint8Array;
  /** Select a cell. */
  select(ref: CellRef): void;
  /** Enter text into a cell as if typed ("12", "50%", "TRUE", "'007", "=SUM(A1:A3)"); false if refused. */
  enter(ref: CellRef, text: string): boolean;
  /** Set a cell's value; false if refused (read-only, or part of an array formula). */
  setCell(ref: CellRef, value: number | boolean | string | null): boolean;
  undo(): boolean;
  redo(): boolean;
  /** Show a sheet by index or name, with the selection it had; false if it can't be shown (hidden, or a chart sheet). */
  showSheet(sheet: number | string): boolean;
  /** A new sheet before the one shown (or after it), then shown; false if read-only. */
  insertSheet(options?: { after?: boolean }): boolean;
  /** Rename a sheet; false if refused (Excel's message is shown). */
  renameSheet(sheet: number | string, name: string): boolean;
  /** Move a sheet to index `to`. */
  moveSheet(sheet: number | string, to: number): boolean;
  /** Hide a sheet's tab, or show it again; false if refused (the last sheet shown). */
  hideSheet(sheet: number | string, hidden?: boolean): boolean;
  /** A sheet's tab colour; null for none. */
  setTabColor(sheet: number | string, color: { rgb: string } | { theme: number; tint?: number } | null): boolean;
  /** Delete a sheet without asking; false if refused (the last sheet shown). */
  deleteSheet(sheet: number | string): boolean;
  /** Run a named command: select, extendTo, move, page, home, end, enter, setValue, clear, undo, redo, showSheet, nextSheet,
   *  insertSheet, renameSheet, moveSheet, hideSheet, tabColor, deleteSheet; on the selection, as the Home tab does:
   *  format(change), toggleFont('b' | 'i' | 'u' | 'strike'), stepFont(1 | -1), toggleWrap, stepIndent(1 | -1),
   *  align('horizontal' | 'vertical', value), borders(preset), stepDecimal(1 | -1), clearFormats,
   *  moreColor('font' | 'fill', 'RRGGBB'); Draw Borders over a range: drawBorders('draw' | 'grid' | 'erase',
   *  { r0, c0, r1, c1 }, { side?, line?: { style, color? } }); recentColor('RRGGBB'). */
  command(name: string, ...args: unknown[]): unknown;
  /** The workbook open here for an agent: each edit one step of the undo history, drawn as it is made. */
  agent(): AgentWorkbook;
  readonly selection: Selection | null;
  /** Call `fn` each time the selection is shown anew (moved, extended, another sheet shown); returns a stop function. */
  onSelect(fn: () => void): () => void;
  readonly readOnly: boolean;
  setReadOnly(flag: boolean): void;
  /** Turn ribbon features on or off: a tab by name ({ home: false }) or a group by its label ({ font: false }). */
  setFeatures(features: Record<string, boolean>): void;
  readonly features: Record<string, boolean>;
  /** Switch the ribbon to the one-line simplified layout ('simple') or the full one ('classic'). */
  setRibbon(mode: 'simple' | 'classic'): void;
  /** The ribbon's layout now. */
  readonly ribbon: 'simple' | 'classic' | null;
  /** Build the printed pages of the sheet shown; returns the page count. `selection`: Excel's Print Selection, the
   *  selection's ranges each on pages of their own instead of the print area. */
  preparePrint(options?: { selection?: boolean }): number;
  readonly workbook: Workbook | null;
  readonly grid: GridHandle | null;
  readonly sheetIndex: number;
  readonly fileName: string | null;
  setTheme(theme: 'light' | 'dark' | null): void;
  /** The sheet shown as a CSV or text file, as Excel's Save As writes it (docs/CSV.md). */
  saveCsv(options?: { utf8?: boolean; kind?: 'csv-utf8' | 'csv' | 'txt' | 'unicode'; sep?: string; encoding?: 'utf-8' | 'utf-16le' | 'windows-1252'; bom?: boolean }): Uint8Array;
  /** A text file through the Text Import box: into a new workbook ('new') or at the active cell ('sheet'); `ask`
   *  false takes `settings` without the box. Resolves to null if cancelled. */
  importText(bytes: Uint8Array | ArrayBuffer, fileName?: string, options?: { into?: 'new' | 'sheet'; settings?: Partial<TextImportSettings>; ask?: boolean }): Promise<unknown>;
  /** Save As (F12): the box (or `choice` as it would answer), the file given to onSave or downloaded. */
  saveAs(choice?: SaveAsChoice): Promise<SavedFile | null>;
  /** The page's country settings (docs/LOCALE.md): a tag ('de-DE') or settings over one. */
  setLocale(locale: string | Partial<LocaleSettings>): LocaleSettings;
  readonly locale: LocaleSettings;
  destroy(): void;
}

export interface LocaleSettings {
  tag: string; decimal: string; group: string; dateOrder: 'MDY' | 'DMY' | 'YMD'; dateSep: string;
  /** Built-in format 14's code in this locale ('dd.mm.yyyy'). */
  shortDate: string; hour12: boolean; listSep: string; arrayCol: string; arrayRow: string;
}

export interface TextImportSettings {
  encoding: 'utf-8' | 'utf-16le' | 'utf-16be' | 'windows-1252'; kind: 'delimited' | 'fixed';
  /** The delimiter characters. */
  seps: string; quote: '"' | "'" | ''; consecutive: boolean;
  /** Fixed width: where each column after the first starts, in characters. */
  breaks: number[]; startRow: number;
  columns: ('general' | 'text' | 'MDY' | 'DMY' | 'YMD' | 'skip')[];
  locale: string; decimal: string; group: string;
}

export interface SaveAsChoice {
  fileName: string; type: 'xlsx' | 'xlsm' | 'csv' | 'txt' | 'pdf';
  sep?: string | null; encoding?: 'utf-8' | 'utf-16le' | 'windows-1252' | null; bom?: boolean | null;
}

export interface SavedFile { fileName: string; type: string; bytes: Uint8Array; mime: string }

export declare function mountXlsxEditor(
  host: HTMLElement,
  options?: {
    bytes?: Uint8Array | ArrayBuffer; fileName?: string; theme?: 'light' | 'dark';
    /** Show the workbook only: no ribbon, nothing editable. */
    readOnly?: boolean;
    /** Ribbon tabs or groups turned off: { home: false }, { font: false, number: false }. */
    features?: Record<string, boolean>;
    /** false hides the ribbon but keeps editing. */
    toolbar?: boolean;
    /** 'simple' starts with the one-line simplified ribbon, 'classic' with the full one; without it, the user's last
     *  choice in this browser. The ribbon's own button still switches. */
    ribbon?: 'simple' | 'classic';
    /** false hides the status bar's figures for the selection; an object ({ minimum: true, maximum: true,
     *  numericalCount: true, average, count, sum }) chooses them, as Excel's Customize Status Bar does. */
    statusBar?: boolean | Record<string, boolean>;
    /** Custom lists the fill handle goes round, as Excel's Edit Custom Lists: [['North', 'East', 'South', 'West']]. */
    fillLists?: string[][];
    /** false leaves formulas saved without a result empty on opening (by default they are worked out, as Excel does). */
    calculateOnOpen?: boolean;
    /** Excel's Error Checking Options: false turns the green error triangles off; { background, rules } chooses
     *  them, `rules` naming those on ('evalError', 'twoDigitTextYear', 'numberStoredAsText', 'formula',
     *  'formulaRange', 'unlockedFormula'). */
    errorChecking?: boolean | { background?: boolean; rules?: string[] };
    /** open-metric-fonts' metric twins of 151 Microsoft fonts (Aptos Narrow's, Tahoma's, Verdana's, Garamond's...),
     *  drawn where the computer lacks the font, from the site's own copy: the folder's URL (README, Fonts). */
    fonts?: { omf?: string };
    /** The name notes and comments are written under. */
    author?: string;
    /** The people @ offers in a comment, besides the workbook's own (its persons part), or a function giving them when
     *  asked; userId (a directory's, such as an email) is written for the person mentioned, as Excel keeps it. */
    people?: Array<{ name: string; userId?: string; providerId?: string }> | (() => Array<{ name: string; userId?: string; providerId?: string }>);
    /** A hyperlink to a web address clicked: the host page opens it (otherwise the editor opens a new tab). */
    onLink?: (url: string) => void;
    /** Country settings for showing and typing numbers and dates: a tag or settings over one (navigator.language
     *  without it). */
    locale?: string | Partial<LocaleSettings>;
    /** Save As's file; without it, the browser downloads it. */
    onSave?: (file: SavedFile) => void;
    /** A PDF export for Save As to offer: open-xlsx-editor/pdf's exportPdf. */
    exportPdf?: (editor: XlsxEditor) => Promise<Uint8Array>;
    onChange?: () => void;
  },
): XlsxEditor;

/** The open-metric-fonts release the editor's map of faces is made for (README, Fonts). */
export declare const OMF_VERSION: string;
