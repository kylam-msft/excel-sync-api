import * as $ from "jquery";
import { generateLogString } from "./logger-utils";

export function log(data: any, severity: ConsoleLogTypes = "info", $outputDiv: JQuery = $("#output")) {
  if (!Array.isArray(data)) {
    data = [data];
  }

  const generated = generateLogString(data, severity);
  severity = generated.severity;
  const message = generated.message;

  const color =
    ({
      log: "black",
      info: "blue",
      warn: "yellow",
      error: "red",
    } as { [key: string]: string })[severity] || "purple";

  const $log = $("<pre class='log-entry'>")
    .text(message)
    .css("color", color);
  $log.appendTo($outputDiv);
  $log[0].scrollIntoView();
}
