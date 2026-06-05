import writeXlsxFile, { type SheetData, type SheetOptions } from "write-excel-file/node";
import dataValidation, { type DataValidationSheetOptions } from "../src/index.js";

// A dropdown whose options come from a range of cells in the same sheet.
// Useful when values contain commas/quotes (which inline lists can't handle)
// or when the set of options is long enough to bust the 255-char inline limit.
const data: SheetData = [
  [
    { value: "Pick", fontWeight: "bold" },
    { value: "Options", fontWeight: "bold" },
  ],
  ["", "apple"],
  ["", "banana"],
  ["", "cherry"],
];

const sheetOptions: SheetOptions<any> & DataValidationSheetOptions = {
  sheet: "List (range)",
  dataValidation: [
    {
      cellRange: { from: { row: 2, column: 1 }, to: { row: 4, column: 1 } },
      validation: {
        type: "list",
        valuesRange: "$B$2:$B$4",
      },
    },
  ],
};

const output = new URL("./list-range.xlsx", import.meta.url).pathname;

await writeXlsxFile(data, sheetOptions, { features: [dataValidation] }).toFile(output);

console.log(`Wrote ${output}`);
