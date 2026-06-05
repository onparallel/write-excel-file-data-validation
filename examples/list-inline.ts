import writeXlsxFile, { type SheetData, type SheetOptions } from "write-excel-file/node";
import dataValidation, { type DataValidationSheetOptions } from "../src/index.js";

// A dropdown of literal values defined inline.
const data: SheetData = [
  [{ value: "Status", fontWeight: "bold" }],
  ["open"],
  ["closed"],
  ["pending"],
];

const sheetOptions: SheetOptions<any> & DataValidationSheetOptions = {
  sheet: "List (inline)",
  dataValidation: [
    {
      cellRange: { from: { row: 2, column: 1 }, to: { row: 4, column: 1 } },
      validation: {
        type: "list",
        values: ["open", "closed", "pending"],
        error: "Pick one of the allowed statuses",
      },
    },
  ],
};

const output = new URL("./list-inline.xlsx", import.meta.url).pathname;

await writeXlsxFile(data, sheetOptions, { features: [dataValidation] }).toFile(output);

console.log(`Wrote ${output}`);
