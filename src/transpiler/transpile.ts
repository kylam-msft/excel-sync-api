import * as $ from "jquery";
import * as ts from "typescript";
import { CompilerHost } from "./compiler-host";
import { log } from "../util/log";

const SOURCE_FILE_NAME: string = "file.ts";

const compilerHost = new CompilerHost();

function getCompilerOptions(options: ts.CompilerOptions): ts.CompilerOptions {
  const result = (ts as any).clone(options) as ts.CompilerOptions;

  result.isolatedModules = true;

  /* arrange for an external source map */
  if (result.sourceMap === undefined) {
    result.sourceMap = result.inlineSourceMap;
  }

  if (result.sourceMap === undefined) {
    result.sourceMap = true;
  }

  result.inlineSourceMap = false;

  /* these options are incompatible with isolatedModules */
  result.declaration = false;
  result.noEmitOnError = false;
  result.out = undefined;
  result.outFile = undefined;

  /* without this we get a 'lib.d.ts not found' error */
  result.noLib = true;
  result.lib = undefined; // incompatible with noLib

  /* without this we get a 'cannot find type-reference' error */
  result.types = [];

  /* without this we get 'cannot overwrite existing file' when transpiling js files */
  result.suppressOutputPathCheck = true;

  /* with this we don't get any files */
  result.noEmit = false;

  return result;
}

export default function transpile(code: string, options: ts.CompilerOptions = {}): any {
  const start = window.performance.now();
  let sourceFile = ts.createSourceFile(SOURCE_FILE_NAME, code, ts.ScriptTarget.Latest, true);

  sourceFile = ts.transform(sourceFile, [addAwaitToAllCallExpressions as ts.TransformerFactory<ts.Node>])
    .transformed[0] as ts.SourceFile;

  addAsyncToAllFunctions(sourceFile);

  if ($("#transpile").prop("checked")) {
    log("TRANSPILED RESULT:");
    log(printSourceFile(sourceFile));
    const end = window.performance.now();
    const time = end - start;
    log(`Transpile time: ${time.toFixed(4)} ms`);
  }

  wrapCodeInMain(sourceFile);

  const result = printSourceFile(sourceFile);

  const outputs: any[] = [];

  const compilerOptions = getCompilerOptions(options);

  compilerHost.addFile(SOURCE_FILE_NAME, result, compilerOptions.target);

  compilerHost.writeFile = (name: string, text: string, writeByteOrderMark: boolean) => {
    outputs.push({ name, text, writeByteOrderMark });
  };

  const program: ts.Program = ts.createProgram([SOURCE_FILE_NAME], compilerOptions, compilerHost);

  const emitResult: ts.EmitResult = program.emit();
  // console.log(JSON.stringify(outputs));
  /* tslint:disable-next-line */
  eval(outputs[1].text);
}

function wrapCodeInMain(sourceFile: ts.SourceFile): void {
  const body = ts.createBlock(sourceFile.statements);
  const identifier = ts.createIdentifier("main");
  const modifier = ts.createModifier(ts.SyntaxKind.AsyncKeyword);
  const functionExpression = ts.createFunctionExpression([modifier], undefined, identifier, undefined, [], null, body);
  const parenthesizedExpression = ts.createParen(functionExpression);
  const callExpression = ts.createCall(parenthesizedExpression, [], []);
  const statement = ts.createExpressionStatement(callExpression);
  sourceFile.statements = ts.createNodeArray([statement]);
}

function addAwaitToAllCallExpressions<T extends ts.Node>(context: ts.TransformationContext) {
  return (rootNode: T) => {
    const visit = (node: ts.Node): ts.Node => {
      if (node && ts.isCallExpression(node)) {
        node = ts.visitEachChild(node, visit, context);
        return ts.createAwait(node as ts.Expression);
      }

      return ts.visitEachChild(node, visit, context);
    };

    return ts.visitNode(rootNode, visit);
  };
}

function addAsyncToAllFunctions(sourceFile: ts.SourceFile): void {
  const addAsync = (node: ts.Node): void => {
    if (!ts.isFunctionDeclaration(node)) {
      return;
    }

    if (
      node.modifiers !== undefined &&
      node.modifiers.filter(modifier => modifier.kind === ts.SyntaxKind.AsyncKeyword).length > 0
    ) {
      return;
    }

    node.modifiers = !node.modifiers ? ts.createNodeArray() : node.modifiers;
    node.modifiers = ts.createNodeArray(
      node.modifiers.concat(ts.createModifiersFromModifierFlags(ts.ModifierFlags.Async))
    );
  };

  ts.forEachChild(sourceFile, addAsync);
}

function printSourceFile(sourceFile: ts.SourceFile): string {
  const printer = ts.createPrinter({
    newLine: ts.NewLineKind.LineFeed,
  });
  return printer.printNode(ts.EmitHint.Unspecified, sourceFile, sourceFile);
}
