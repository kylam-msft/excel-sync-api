import * as $ from "jquery";
import editor from "../editor/editor";
import transpile from "../transpiler/transpile";
import { SAMPLES } from "../util/constants";

initUI();

/* tslint:disable */
(function() {
  const old = console.log;
  console.log = function(message) {
    let formattedMessage = message;
    if (typeof message == "object") {
      formattedMessage = JSON && JSON.stringify ? JSON.stringify(message) : message;
    }

    const $log = $("<pre class='log-entry'>").text(formattedMessage);
    $log.appendTo($("#output"));
    $log[0].scrollIntoView();
    old(message);
  };
})();

function run() {
  const code = editor.getValue();
  transpile(code);
}

function clearLogs() {
  $("#output, #footer").empty();
}

function initUI() {
  const $sampleList = $("#sample-list");
  ["BASIC_SAMPLE", "INVENTORY_MAKER"].forEach(name => {
    $sampleList.append('<option id="' + name + '">' + name + "</option>");
  });
  $sampleList.val("BASIC_SAMPLE");
  editor.setValue(SAMPLES["BASIC_SAMPLE"]);
  $sampleList.change(() => {
    const sampleId = $sampleList.val();
    if (sampleId !== "") {
      const code = (SAMPLES as any)[sampleId as string];
      editor.setValue(code);
    }
  });

  $("#run").click(run);
  $("#clear").click(clearLogs);
}
