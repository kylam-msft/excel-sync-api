import * as monaco from "monaco-editor";

monaco.languages.typescript.typescriptDefaults.setCompilerOptions({
  target: monaco.languages.typescript.ScriptTarget.ES2016,
  allowNonTsExtensions: true,
});

monaco.languages.typescript.typescriptDefaults.addExtraLib(`
  declare namespace ExcelScript {
    const workbook: Workbook;

    interface Workbook {
      getSelectedRange(): Range;
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
      setValues(values: string[][]): void
      getAddress(): string;
      getAbsoluteResizedRange(rowCount: number, columnCount: number): Range;
      getResizedRange(rowCount: number, columnCount: number): Range;
      getCell(rowIndex: number, columnIndex: number): Range;
      getColumnCount(): number;
      getFormat(): Format;
    }

    interface Format {
      autofitColumns(): string[][];
      getFill(): Fill;
      getFont(): Font;
    }

    interface Fill {
      getColor(): string;
      setColor(color: string): void;
    }

    interface Font {
      getBold(): boolean;
      setBold(bold: boolean): void;
      getColor(): string;
      setColor(color: string): void;
      getName(): string;
      setName(name: string): void;
      getSize(): number;
      setSize(size: number): void;
    }
  }
`);

const editor: monaco.editor.IStandaloneCodeEditor = monaco.editor.create(document.getElementById("editor"), {
  language: "typescript",
  lineNumbers: "off",
  selectOnLineNumbers: false,
  minimap: {
    enabled: false,
  },
});

export default editor;
