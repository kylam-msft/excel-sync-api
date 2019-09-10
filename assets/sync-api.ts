/* tslint:disable */
/// <reference path="../office.experimental.d.ts" />

(ExcelOp.Workbook.prototype as any).getWorksheets = function() {
  return this.worksheets.retrieve("name").then(sheets => {
    const promises = [];
    for (const sheet of sheets) {
      promises.push(this.worksheets.getItem(sheet.name));
    }
    return Promise.all(promises);
  });
};

(ExcelOp.Workbook.prototype as any).getWorksheet = function(id: string) {
  return this.worksheets.getItem(id);
};

(ExcelOp.Workbook.prototype as any).addWorksheet = function(id: string) {
  return this.worksheets.add(id);
};

// const originalGetRange = ExcelOp.Worksheet.prototype.getRange;
// (ExcelOp.Worksheet.prototype as any).getRange = function(address: string) {
//   return originalGetRange.call(this, address);
// };

(ExcelOp.Worksheet.prototype as any).getName = function() {
  return this.retrieve("name").then(({ name }) => name);
};

(ExcelOp.Worksheet.prototype as any).getTable = function(id: string) {
  return this.tables.getItem(id);
};

(ExcelOp.Worksheet.prototype as any).addTable = function(address: string | ExcelOp.Range, hasHeaders?: boolean) {
  return this.tables.add(address, hasHeaders);
};

(ExcelOp.Worksheet.prototype as any).setPosition = function(position: number) {
  return this.update({ position });
};

(ExcelOp.Worksheet.prototype as any).getName = function() {
  return this.retrieve("name").then(({ name }) => name);
};

(ExcelOp.Table.prototype as any).getRows = function() {
  return this.rows.getCount().then(count => {
    const promises = [];
    for (let i = 0; i < count; i++) {
      promises.push(this.rows.getItemAt(i));
    }
    return Promise.all(promises);
  });
};

(ExcelOp.Table.prototype as any).getRowAt = function(index: number) {
  return this.rows.getItemAt(index);
};

(ExcelOp.Table.prototype as any).getColumn = function(id: string) {
  return this.columns.getItem(id);
};

(ExcelOp.TableRow.prototype as any).getValues = function() {
  return this.retrieve("values").then(({ values }) => values);
};

(ExcelOp.TableColumn.prototype as any).getValues = function() {
  return this.retrieve("values").then(({ values }) => values);
};

(ExcelOp.Range.prototype as any).getValues = function() {
  return this.retrieve("values").then(({ values }) => values);
};

(ExcelOp.Range.prototype as any).setValues = function(values) {
  return this.update({ values });
};

(ExcelOp.Range.prototype as any).getAddress = function() {
  return this.retrieve("address").then(({ address }) => address);
};

(ExcelOp.Range.prototype as any).getColumnCount = function() {
  return this.retrieve("columnCount").then(({ columnCount }) => columnCount);
};

(ExcelOp.Range.prototype as any).getFormat = function() {
  return this.format;
};

(ExcelOp.RangeFormat.prototype as any).getFill = function() {
  return this.fill;
};

(ExcelOp.RangeFill.prototype as any).getColor = function() {
  return this.retrieve("color").then(({ color }) => color);
};

(ExcelOp.RangeFill.prototype as any).setColor = function(color: string) {
  return this.update({ color });
};

namespace ExcelScript {
  export const workbook = ExcelOp.workbook;
}

// namespace ExcelScript {
//   export class Workbook extends Excel.Workbook {
//     getWorksheet(id: string) {
//       return Excel.run(async (context) => {
//          return context.workbook.worksheets.getItem(id);
//       });
//     }
//   }

//   export class Worksheet extends Excel.Worksheet {
//     getRange(address: string) {
//       return Excel.run(async () => {
//         return this.getRange(address);
//       });
//     }
//   }

//   export class Range extends Excel.Range {
//     getValues() {
//       return Excel.run(context => {
//          this.load('values');
//          return context.sync().then(() => this.values);
//       });
//     }

//     setValues(values) {
//       return Excel.run(async () => {
//          this.values = values;
//       });
//     }
//   }

// export const workbook = (OfficeExtension as any).BatchApiHelper.createRootServiceObject(Excel.Workbook, new Excel.RequestContext());
// }
