import writeXlsxFile, { type SheetData, type SheetOptions } from "write-excel-file/node";
import dataValidation, { type DataValidationSheetOptions } from "../src/index.js";

// Showcase the full set of metadata: an input popup on cell selection
// and a `warning` style error popup that lets the user override.
const data: SheetData = [[{ value: "Age", fontWeight: "bold" }], [25]];

const sheetOptions: SheetOptions<any> & DataValidationSheetOptions = {
  sheet: "Messages",
  dataValidation: [
    {
      cellRange: { from: { row: 2, column: 1 }, to: { row: 2, column: 1 } },
      validation: {
        type: "integer",
        operator: ">=",
        value: 0,
        inputTitle: "Enter age",
        input: "Must be a non-negative integer",
        errorTitle: "Invalid age",
        error: "Age must be 0 or greater",
        errorStyle: "warning",
      },
    },
  ],
};

const output = new URL("./input-and-error-messages.xlsx", import.meta.url).pathname;

await writeXlsxFile(data, sheetOptions, { features: [dataValidation] }).toFile(output);

console.log(`Wrote ${output}`);
