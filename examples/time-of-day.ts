import writeXlsxFile, { type SheetData, type SheetOptions } from "write-excel-file/node";
import dataValidation, { type DataValidationSheetOptions } from "../src/index.js";

// A time-of-day validation: only times between 09:00 and 17:00 are accepted.
// `Date` values are reduced to their fractional-day component so 09:00 -> 0.375.
const data: SheetData = [
  [{ value: "Working hours (09:00-17:00)", fontWeight: "bold" }],
  [{ value: 0.5, format: "h:mm AM/PM" }],
];

const sheetOptions: SheetOptions<any> & DataValidationSheetOptions = {
  sheet: "Time of day",
  dataValidation: [
    {
      cellRange: { from: { row: 2, column: 1 }, to: { row: 2, column: 1 } },
      validation: {
        type: "time",
        operator: "...",
        value: new Date(Date.UTC(1970, 0, 1, 9, 0, 0)),
        value2: new Date(Date.UTC(1970, 0, 1, 17, 0, 0)),
        error: "Time must be between 09:00 and 17:00",
      },
    },
  ],
};

const output = new URL("./time-of-day.xlsx", import.meta.url).pathname;

await writeXlsxFile(data, sheetOptions, { features: [dataValidation] }).toFile(output);

console.log(`Wrote ${output}`);
