/* tslint:disable */
/// <reference path="../office.experimental.d.ts" />
ExcelOp.Workbook.prototype.getWorksheets = function () {
    var _this = this;
    return this.worksheets.retrieve("name").then(function (sheets) {
        var promises = [];
        for (var _i = 0, sheets_1 = sheets; _i < sheets_1.length; _i++) {
            var sheet = sheets_1[_i];
            promises.push(_this.worksheets.getItem(sheet.name));
        }
        return Promise.all(promises);
    });
};
ExcelOp.Workbook.prototype.getWorksheet = function (id) {
    return this.worksheets.getItem(id);
};
ExcelOp.Workbook.prototype.addWorksheet = function (id) {
    return this.worksheets.add(id);
};
// const originalGetRange = ExcelOp.Worksheet.prototype.getRange;
// (ExcelOp.Worksheet.prototype as any).getRange = function(address: string) {
//   return originalGetRange.call(this, address);
// };
ExcelOp.Worksheet.prototype.getName = function () {
    return this.retrieve("name").then(function (_a) {
        var name = _a.name;
        return name;
    });
};
ExcelOp.Worksheet.prototype.getTable = function (id) {
    return this.tables.getItem(id);
};
ExcelOp.Worksheet.prototype.addTable = function (address, hasHeaders) {
    return this.tables.add(address, hasHeaders);
};
ExcelOp.Worksheet.prototype.setPosition = function (position) {
    return this.update({ position: position });
};
ExcelOp.Worksheet.prototype.getName = function () {
    return this.retrieve("name").then(function (_a) {
        var name = _a.name;
        return name;
    });
};
ExcelOp.Table.prototype.getRows = function () {
    var _this = this;
    return this.rows.getCount().then(function (count) {
        var promises = [];
        for (var i = 0; i < count; i++) {
            promises.push(_this.rows.getItemAt(i));
        }
        return Promise.all(promises);
    });
};
ExcelOp.Table.prototype.getRowAt = function (index) {
    return this.rows.getItemAt(index);
};
ExcelOp.Table.prototype.getColumn = function (id) {
    return this.columns.getItem(id);
};
ExcelOp.TableRow.prototype.getValues = function () {
    return this.retrieve("values").then(function (_a) {
        var values = _a.values;
        return values;
    });
};
ExcelOp.TableColumn.prototype.getValues = function () {
    return this.retrieve("values").then(function (_a) {
        var values = _a.values;
        return values;
    });
};
ExcelOp.Range.prototype.getValues = function () {
    return this.retrieve("values").then(function (_a) {
        var values = _a.values;
        return values;
    });
};
ExcelOp.Range.prototype.setValues = function (values) {
    return this.update({ values: values });
};
ExcelOp.Range.prototype.getAddress = function () {
    return this.retrieve("address").then(function (_a) {
        var address = _a.address;
        return address;
    });
};
ExcelOp.Range.prototype.getColumnCount = function () {
    return this.retrieve("columnCount").then(function (_a) {
        var columnCount = _a.columnCount;
        return columnCount;
    });
};
ExcelOp.Range.prototype.getFormat = function () {
    return this.format;
};
ExcelOp.RangeFormat.prototype.getFill = function () {
    return this.fill;
};
ExcelOp.RangeFormat.prototype.getFont = function () {
    return this.font;
};
ExcelOp.RangeFill.prototype.getColor = function () {
    return this.retrieve("color").then(function (_a) {
        var color = _a.color;
        return color;
    });
};
ExcelOp.RangeFill.prototype.setColor = function (color) {
    return this.update({ color: color });
};
ExcelOp.RangeFont.prototype.getBold = function () {
    return this.retrieve("bold").then(function (_a) {
        var bold = _a.bold;
        return bold;
    });
};
ExcelOp.RangeFont.prototype.setBold = function (bold) {
    return this.update({ bold: bold });
};
ExcelOp.RangeFont.prototype.getColor = function () {
    return this.retrieve("color").then(function (_a) {
        var color = _a.color;
        return color;
    });
};
ExcelOp.RangeFont.prototype.setColor = function (color) {
    return this.update({ color: color });
};
ExcelOp.RangeFont.prototype.getName = function () {
    return this.retrieve("name").then(function (_a) {
        var name = _a.name;
        return name;
    });
};
ExcelOp.RangeFont.prototype.setName = function (name) {
    return this.update({ name: name });
};
ExcelOp.RangeFont.prototype.getSize = function () {
    return this.retrieve("size").then(function (_a) {
        var size = _a.size;
        return size;
    });
};
ExcelOp.RangeFont.prototype.setSize = function (size) {
    return this.update({ size: size });
};
var ExcelScript;
(function (ExcelScript) {
    ExcelScript.workbook = ExcelOp.workbook;
})(ExcelScript || (ExcelScript = {}));
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
