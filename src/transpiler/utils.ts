import * as ts from "typescript";

/**
 * Generates the transformed version of the test file.
 *
 * @param fileName The name of the file we're generating.
 */
export const generate = (fileName: string, actual: ts.SourceFile): ts.SourceFile => {
  const resultFile = ts.createSourceFile(
    fileName + "ts",
    "",
    ts.ScriptTarget.Latest,
    /*setParentNodes*/ false,
    ts.ScriptKind.TS
  );
  const printer = ts.createPrinter({
    newLine: ts.NewLineKind.LineFeed,
  });
  const result = printer.printNode(ts.EmitHint.Unspecified, actual, resultFile);

  actual = ts.createSourceFile(fileName, result, ts.ScriptTarget.Latest, true);
  return actual;
};
