import writeXlsxFile, { type Sheet } from "write-excel-file/node";
import dataValidation, { type DataValidationSheetOptions } from "../src/index.js";

// A dropdown whose options come from a range of cells in a DIFFERENT sheet.
// Tests whether `valuesRange` accepts a cross-sheet reference like `Lists!$A$2:$A$10`.
const mainSheet: Sheet<any> & DataValidationSheetOptions = {
  sheet: "Main",
  data: [
    [
      { value: "Item", fontWeight: "bold" },
      { value: "Status", fontWeight: "bold" },
    ],
    ["Order #1", ""],
    ["Order #2", ""],
    ["Order #3", ""],
  ],
  dataValidation: [
    {
      cellRange: { from: { row: 2, column: 2 }, to: { row: 4, column: 2 } },
      validation: {
        type: "list",
        valuesRange: "Lists!$A$2:$A$5",
      },
    },
  ],
};

const listsSheet: Sheet<any> = {
  sheet: "Lists",
  data: [
    [{ value: "Status options", fontWeight: "bold" }],
    ["Pending"],
    ["In progress"],
    ["Completed"],
    ["Cancelled"],
  ],
};

const output = new URL("./list-range-cross-sheet.xlsx", import.meta.url).pathname;

await writeXlsxFile([mainSheet, listsSheet], { features: [dataValidation] }).toFile(output);

console.log(`Wrote ${output}`);
