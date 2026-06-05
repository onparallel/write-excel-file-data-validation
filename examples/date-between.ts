import writeXlsxFile, { type SheetData, type SheetOptions } from "write-excel-file/node";
import dataValidation, { type DataValidationSheetOptions } from "../src/index.js";

// A date that must fall within 2024. `Date` values are converted to Excel serials.
const data: SheetData = [
  [{ value: "Date in 2024", fontWeight: "bold" }],
  [{ value: new Date(Date.UTC(2024, 5, 1)), type: Date, format: "mm/dd/yyyy" }],
];

const sheetOptions: SheetOptions<any> & DataValidationSheetOptions = {
  sheet: "Date between",
  dataValidation: [
    {
      cellRange: { from: { row: 2, column: 1 }, to: { row: 2, column: 1 } },
      validation: {
        type: "date",
        operator: "...",
        value: new Date(Date.UTC(2024, 0, 1)),
        value2: new Date(Date.UTC(2024, 11, 31)),
        error: "Date must be in 2024",
      },
    },
  ],
};

const output = new URL("./date-between.xlsx", import.meta.url).pathname;

await writeXlsxFile(data, sheetOptions, { features: [dataValidation] }).toFile(output);

console.log(`Wrote ${output}`);
