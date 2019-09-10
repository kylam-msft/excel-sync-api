interface ExcelScript {
  workbook: Workbook;
}

interface Workbook {
  getWorksheets(): Worksheet [];
  getWorksheet(id: string): Worksheet;
  addWorksheet(id: string): Worksheet;
}

interface Worksheet {
  getName(): string;
  getTable(id: string): Table;
  addTable(address: string | Range, hasHeaders?: boolean): Table;
  delete(): void;
  getRange(address: string): Range
  setPosition(position: number): void
}

interface Table {
  getDataBodyRange(): Range;
  getRows(): TableRow[]
  getRowAt(index: number): TableRow
  getColumn(id: string): TableColumn
  getHeaderRowRange(): Range;
}

interface TableRow {
  getValues(): string[][]
  getRange(): Range;
}

interface TableColumn {
  getValues(): string[][]
  getDataBodyRange(): Range
}

interface Range {
  getValues(): string[][];
  setValues(values: string[][]): void;
  getAddress(): string;
  getAbsoluteResizedRange(rowCount: number, columnCount: number): Range;
  getResizedRange(rowCount: number, columnCount: number): Range;
  getCell(rowIndex: number, columnIndex: number): Range;
  getColumnCount(): number;
  getFormat(): Format
}

interface Format {
  autofitColumns(): string[][];
  getFill(): Fill;
}

interface Fill {
  getColor(): string;
  setColor(color: string): void;
}