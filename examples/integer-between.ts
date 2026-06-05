import writeXlsxFile, { type SheetData, type SheetOptions } from "write-excel-file/node";
import dataValidation, { type DataValidationSheetOptions } from "../src/index.js";

// An integer between 1 and 10 (inclusive). Excel will block any other input.
const data: SheetData = [[{ value: "Quantity (1-10)", fontWeight: "bold" }], [5], [3], [10]];

const sheetOptions: SheetOptions<any> & DataValidationSheetOptions = {
  sheet: "Integer between",
  dataValidation: [
    {
      cellRange: { from: { row: 2, column: 1 }, to: { row: 4, column: 1 } },
      validation: {
        type: "integer",
        operator: "...",
        value: 1,
        value2: 10,
        errorTitle: "Out of range",
        error: "Quantity must be between 1 and 10",
        errorStyle: "stop",
      },
    },
  ],
};

const output = new URL("./integer-between.xlsx", import.meta.url).pathname;

await writeXlsxFile(data, sheetOptions, { features: [dataValidation] }).toFile(output);

console.log(`Wrote ${output}`);
