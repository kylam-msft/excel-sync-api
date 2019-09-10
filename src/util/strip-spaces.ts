export function stripSpaces(text: string) {
  const lines: string[] = text.split("\n");

  // Replace each tab with 4 spaces.
  // this doesn't even work.
  /*
  for (let i: number = 0; i < lines.length; i++) {
      lines[i].replace("\t", "    ");
  }
  */

  let isZeroLengthLine: boolean = true;
  let arrayPosition: number = 0;

  // Remove zero length lines from the beginning of the snippet.
  do {
      const currentLine: string = lines[arrayPosition];
      if (currentLine.trim() === "") {
          lines.splice(arrayPosition, 1);
      } else {
          isZeroLengthLine = false;
      }
  } while (isZeroLengthLine || arrayPosition === lines.length);

  arrayPosition = lines.length - 1;
  isZeroLengthLine = true;

  // Remove zero length lines from the end of the snippet.
  do {
      const currentLine: string = lines[arrayPosition];
      if (currentLine.trim() === "") {
          lines.splice(arrayPosition, 1);
          arrayPosition--;
      } else {
          isZeroLengthLine = false;
      }
  } while (isZeroLengthLine);

  // Get smallest indent for align left.
  let shortestIndentSize: number = 1024;
  lines.forEach(currentLine => {
      if (currentLine.trim() !== "") {
          const spaces: number = currentLine.search(/\S/);
          if (spaces < shortestIndentSize) {
              shortestIndentSize = spaces;
          }
      }
  });

  // Align left
  for (let i: number = 0; i < lines.length; i++) {
      if (lines[i].length >= shortestIndentSize) {
          lines[i] = lines[i].substring(shortestIndentSize);
      }
  }

  // Convert the array back into a string and return it.
  let finalSetOfLines: string = "";
  for (let i: number = 0; i < lines.length; i++) {
      if (i < lines.length - 1) {
          finalSetOfLines += lines[i] + "\n";
      } else {
          finalSetOfLines += lines[i];
      }
  }
  return finalSetOfLines;
}
