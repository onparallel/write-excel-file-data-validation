import { describe, it, expect } from "vitest";

import dataValidation from "./dataValidation.js";
import type { DataValidationSheetOptions } from "./types.js";

const transformSheetXml =
  dataValidation.files!.transform!["xl/worksheets/sheet{id}.xml"]!.transform!;

function transform(sheetOptions: DataValidationSheetOptions): string {
  return transformSheetXml("<worksheet></worksheet>", sheetOptions as never, {
    sheetIndex: 0,
    sheetId: "1",
  });
}

describe("dataValidation", () => {
  it("should not change XML when no data validation is set", () => {
    const xml = "<worksheet></worksheet>";
    expect(transformSheetXml(xml, {} as never, { sheetIndex: 0, sheetId: "1" })).toBe(xml);
  });

  it("should not change XML when data validation is an empty array", () => {
    const xml = "<worksheet></worksheet>";
    expect(
      transformSheetXml(xml, { dataValidation: [] } as never, { sheetIndex: 0, sheetId: "1" }),
    ).toBe(xml);
  });

  it("should generate a list validation with inline values", () => {
    expect(
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 2, column: 1 }, to: { row: 5, column: 1 } },
            validation: { type: "list", values: ["open", "closed", "pending"] },
          },
        ],
      }),
    ).toBe(
      "<worksheet>" +
        '<dataValidations count="1">' +
        '<dataValidation type="list" allowBlank="1" showErrorMessage="1" showInputMessage="1" sqref="A2:A5">' +
        '<formula1>"open,closed,pending"</formula1>' +
        "</dataValidation>" +
        "</dataValidations>" +
        "</worksheet>",
    );
  });

  it("should generate a list validation with a values range", () => {
    expect(
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
            validation: { type: "list", valuesRange: "$E$4:$G$4" },
          },
        ],
      }),
    ).toContain("<formula1>$E$4:$G$4</formula1>");
  });

  it("should throw when a list value contains a comma", () => {
    expect(() =>
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
            validation: { type: "list", values: ["a,b", "c"] },
          },
        ],
      }),
    ).toThrow("cannot contain commas");
  });

  it("should throw when a list value contains a double-quote", () => {
    expect(() =>
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
            validation: { type: "list", values: ['say "hi"'] },
          },
        ],
      }),
    ).toThrow("cannot contain double-quote");
  });

  it("should throw when a list has an empty values array", () => {
    expect(() =>
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
            validation: { type: "list", values: [] },
          },
        ],
      }),
    ).toThrow("at least one value");
  });

  it("should throw when an inline list of values exceeds 255 characters", () => {
    const longValues: string[] = [];
    for (let i = 0; i < 100; i++) {
      longValues.push("option" + i);
    }
    expect(() =>
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
            validation: { type: "list", values: longValues },
          },
        ],
      }),
    ).toThrow("exceeds the maximum length");
  });

  it("should generate an integer between validation with two formulas", () => {
    expect(
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 5, column: 2 }, to: { row: 5, column: 2 } },
            validation: { type: "integer", operator: "...", value: 1, value2: 10 },
          },
        ],
      }),
    ).toBe(
      "<worksheet>" +
        '<dataValidations count="1">' +
        '<dataValidation type="whole" operator="between" allowBlank="1" showErrorMessage="1" showInputMessage="1" sqref="B5:B5">' +
        "<formula1>1</formula1>" +
        "<formula2>10</formula2>" +
        "</dataValidation>" +
        "</dataValidations>" +
        "</worksheet>",
    );
  });

  it("should generate a decimal greater-than validation", () => {
    expect(
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
            validation: { type: "decimal", operator: ">", value: 0.5 },
          },
        ],
      }),
    ).toContain('<dataValidation type="decimal" operator="greaterThan"');
  });

  it("should generate a textLength validation", () => {
    expect(
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
            validation: { type: "textLength", operator: "<=", value: 50 },
          },
        ],
      }),
    ).toContain('<dataValidation type="textLength" operator="lessThanOrEqual"');
  });

  it("should generate a date between validation, converting Dates to Excel serials", () => {
    const from = new Date(Date.UTC(2000, 0, 1));
    const to = new Date(Date.UTC(2000, 0, 2));
    const xml = transform({
      dataValidation: [
        {
          cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
          validation: { type: "date", operator: "...", value: from, value2: to },
        },
      ],
    });
    expect(xml).toContain("<formula1>36526</formula1>");
    expect(xml).toContain("<formula2>36527</formula2>");
  });

  it("should convert a `time` Date to a fractional day, not a full date serial", () => {
    const nineAm = new Date(Date.UTC(1970, 0, 1, 9, 0, 0));
    const xml = transform({
      dataValidation: [
        {
          cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
          validation: { type: "time", operator: ">=", value: nineAm },
        },
      ],
    });
    expect(xml).toContain("<formula1>0.375</formula1>");
    expect(xml).not.toContain("25569");
  });

  it("should generate a custom validation", () => {
    expect(
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 5, column: 6 }, to: { row: 5, column: 6 } },
            validation: { type: "custom", formula: "=AND(F5=50,G5=60)" },
          },
        ],
      }),
    ).toContain("<formula1>=AND(F5=50,G5=60)</formula1>");
  });

  it('should generate an "any" validation that only shows messages', () => {
    const xml = transform({
      dataValidation: [
        {
          cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
          validation: { type: "any", inputTitle: "Tip", input: "Enter anything here" },
        },
      ],
    });
    expect(xml).toContain('promptTitle="Tip"');
    expect(xml).toContain('prompt="Enter anything here"');
    expect(xml).not.toContain("type=");
  });

  it("should include error and input metadata as attributes", () => {
    const xml = transform({
      dataValidation: [
        {
          cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
          validation: {
            type: "integer",
            operator: "...",
            value: 1,
            value2: 10,
            errorTitle: "Invalid",
            error: "Must be 1-10",
            errorStyle: "warning",
            inputTitle: "Range",
            input: "Enter 1-10",
            allowBlank: false,
          },
        },
      ],
    });
    expect(xml).toContain('errorStyle="warning"');
    expect(xml).toContain('errorTitle="Invalid"');
    expect(xml).toContain('error="Must be 1-10"');
    expect(xml).toContain('promptTitle="Range"');
    expect(xml).toContain('prompt="Enter 1-10"');
    expect(xml).toContain('allowBlank="0"');
  });

  it("should respect explicit allowBlank=false, showErrorMessage=false and showInputMessage=false", () => {
    const xml = transform({
      dataValidation: [
        {
          cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
          validation: {
            type: "integer",
            operator: ">",
            value: 0,
            allowBlank: false,
            showErrorMessage: false,
            showInputMessage: false,
          },
        },
      ],
    });
    expect(xml).toContain('allowBlank="0"');
    expect(xml).toContain('showErrorMessage="0"');
    expect(xml).toContain('showInputMessage="0"');
  });

  it('should emit showDropDown="1" only when showDropdown is explicitly disabled', () => {
    expect(
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
            validation: { type: "list", values: ["a", "b"], showDropdown: false },
          },
        ],
      }),
    ).toContain('showDropDown="1"');

    expect(
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
            validation: { type: "list", values: ["a", "b"] },
          },
        ],
      }),
    ).not.toContain("showDropDown");
  });

  it("should escape XML special characters in error and input messages", () => {
    const xml = transform({
      dataValidation: [
        {
          cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
          validation: {
            type: "any",
            error: 'A & B < C > "D"',
            errorTitle: "<oops>",
          },
        },
      ],
    });
    expect(xml).toContain('error="A &amp; B &lt; C &gt; &quot;D&quot;"');
    expect(xml).toContain('errorTitle="&lt;oops&gt;"');
  });

  it("should generate multiple data validations in a single block", () => {
    const xml = transform({
      dataValidation: [
        {
          cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
          validation: { type: "list", values: ["a", "b"] },
        },
        {
          cellRange: { from: { row: 2, column: 1 }, to: { row: 2, column: 1 } },
          validation: { type: "integer", operator: ">", value: 0 },
        },
      ],
    });
    expect(xml).toContain('<dataValidations count="2">');
  });

  it("should throw on an unknown type", () => {
    expect(() =>
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
            validation: { type: "phoneNumber" as never },
          },
        ],
      }),
    ).toThrow("Unknown data validation type");
  });

  it("should throw on an unknown operator", () => {
    expect(() =>
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
            validation: { type: "integer", operator: "??" as never, value: 1 },
          },
        ],
      }),
    ).toThrow("Unknown data validation operator");
  });

  it("should throw on an unknown errorStyle", () => {
    expect(() =>
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
            validation: { type: "any", errorStyle: "foo" as never },
          },
        ],
      }),
    ).toThrow("Unknown data validation `errorStyle`");
  });

  it("should throw when a list validation lacks both `values` and `valuesRange`", () => {
    expect(() =>
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
            validation: { type: "list" } as never,
          },
        ],
      }),
    ).toThrow("list");
  });

  it("should throw when a between operator lacks `value2`", () => {
    expect(() =>
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
            validation: { type: "integer", operator: "...", value: 1 } as never,
          },
        ],
      }),
    ).toThrow("value2");
  });

  it("should throw when a custom validation lacks `formula`", () => {
    expect(() =>
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
            validation: { type: "custom" } as never,
          },
        ],
      }),
    ).toThrow("formula");
  });

  it("should throw when a numeric `value` is NaN or Infinity", () => {
    expect(() =>
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
            validation: { type: "integer", operator: ">", value: NaN },
          },
        ],
      }),
    ).toThrow("finite number");

    expect(() =>
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
            validation: { type: "decimal", operator: ">", value: Infinity },
          },
        ],
      }),
    ).toThrow("finite number");
  });

  it("should throw when a date/time `value` is an invalid Date", () => {
    expect(() =>
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
            validation: { type: "date", operator: ">", value: new Date("not-a-date") },
          },
        ],
      }),
    ).toThrow("finite number");
  });

  it("should throw when `cellRange` is missing", () => {
    expect(() =>
      transform({
        dataValidation: [
          {
            validation: { type: "any" },
          } as never,
        ],
      }),
    ).toThrow("cellRange");
  });

  it("should throw when `cellRange` row or column is less than 1", () => {
    expect(() =>
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 0, column: 1 }, to: { row: 1, column: 1 } },
            validation: { type: "any" },
          },
        ],
      }),
    ).toThrow(">= 1");
  });

  it("should throw when `cellRange` `to` is less than `from`", () => {
    expect(() =>
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 5, column: 1 }, to: { row: 2, column: 1 } },
            validation: { type: "any" },
          },
        ],
      }),
    ).toThrow("greater than or equal to");
  });

  it("should throw when an error title is longer than 32 characters", () => {
    expect(() =>
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
            validation: { type: "any", errorTitle: "a".repeat(33) },
          },
        ],
      }),
    ).toThrow("longer than 32 characters");
  });

  it("should throw when an error message is longer than 255 characters", () => {
    expect(() =>
      transform({
        dataValidation: [
          {
            cellRange: { from: { row: 1, column: 1 }, to: { row: 1, column: 1 } },
            validation: { type: "any", error: "a".repeat(256) },
          },
        ],
      }),
    ).toThrow("longer than 255 characters");
  });
});
