// Copied from https://github.com/OfficeDev/script-lab/blob/master/src/client/app/helpers/standalone-log-helper.ts

// Moved those to interfaces.d.ts:
//     export type ConsoleLogTypes = "log" | "info" | "warn" | "error";

let logColor = "black";

export function getLogColor() {
  return logColor;
}

export function setLogColor(color: "black" | "blue" | "green" | "red" | "yellow") {
  logColor = color;
}

export function generateLogString(
  args: any[],
  severityType: ConsoleLogTypes
): { severity: ConsoleLogTypes; message: string } {
  let message: string = "";
  let isSuccessfulMsg: boolean = true;
  args.forEach((element, index, array) => {
    try {
      message += stringifyPlusPlus(element);
    } catch (e) {
      isSuccessfulMsg = false;
      message += "<Unable to log>";
    }
    message += "\n";
  });
  if (message.length > 0) {
    message = message.trim();
  }

  return {
    message,
    severity: isSuccessfulMsg ? severityType : "error",
  };
}

export function stringifyPlusPlus(object: any): string {
  if (object === null) {
    return "null";
  }

  if (typeof object === "undefined") {
    return "undefined";
  }

  // Don't JSON.stringify strings, because we don't want quotes in the output
  if (typeof object === "string") {
    return object;
  }

  if (object instanceof Error) {
    try {
      return "ERROR: " + "\n" + object.message;
    } catch (e) {
      return stringifyPlusPlus(object.toString());
    }
  }

  if (object instanceof ErrorEvent) {
    try {
      return "ERROR: " + "\n" + jsonStringify(object.message);
    } catch (e) {
      return stringifyPlusPlus(object.toString());
    }
  }

  if (object.toString() !== "[object Object]") {
    return object.toString();
  }

  // Otherwise, stringify the object
  return jsonStringify(object);
}

function jsonStringify(object: any): string {
  return JSON.stringify(
    object,
    (key, value) => {
      if (value && typeof value === "object" && !Array.isArray(value)) {
        return getStringifiableSnapshot(value);
      }
      return value;
    },
    4
  );

  // tslint:disable-next-line:no-shadowed-variable
  function getStringifiableSnapshot(object: any) {
    const snapshot: any = {};

    try {
      let current = object;

      do {
        Object.keys(current).forEach(tryAddName);
        current = Object.getPrototypeOf(current);
      } while (current);

      return snapshot;
    } catch (e) {
      return object;
    }

    function tryAddName(name: string) {
      const hasOwnProperty = Object.prototype.hasOwnProperty;
      if (name.indexOf(" ") < 0 && !hasOwnProperty.call(snapshot, name)) {
        Object.defineProperty(snapshot, name, {
          configurable: true,
          enumerable: true,
          get: () => object[name],
        });
      }
    }
  }
}
