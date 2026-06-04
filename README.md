# @onparallel/write-excel-file-data-validation

Data validation custom feature for [`write-excel-file`](https://www.npmjs.com/package/write-excel-file).

Adds support for Excel's "data validation" rules — dropdown lists, numeric/date/time ranges, text length limits, and custom formulas — without requiring a fork of the base package.

## Install

```sh
npm install write-excel-file @onparallel/write-excel-file-data-validation
```

`write-excel-file` is a peer dependency (`^4.0.0`). The package is published publicly under the `@onparallel` scope — no authentication required to install.

## Usage

Register the feature when calling `writeXlsxFile()`. `write-excel-file/node`'s built-in `SheetOptions` does not know about `dataValidation`, so intersect it with the `DataValidationSheetOptions` type exported by this package:

```ts
import writeXlsxFile, { type SheetOptions } from 'write-excel-file/node'
import dataValidation, {
	type DataValidationSheetOptions
} from '@onparallel/write-excel-file-data-validation'

const sheetOptions: SheetOptions<any> & DataValidationSheetOptions = {
	sheet: 'Sheet1',
	dataValidation: [
		{
			cellRange: {
				from: { row: 2, column: 1 },
				to: { row: 10, column: 1 }
			},
			validation: {
				type: 'list',
				values: ['open', 'closed', 'pending'],
				error: 'Pick one'
			}
		}
	]
}

await writeXlsxFile(data, sheetOptions, { features: [dataValidation] }).toFile('out.xlsx')
```

If you use the intersection in many places, alias it locally: `type MySheetOptions = SheetOptions<any> & DataValidationSheetOptions`.

Why not module augmentation? `write-excel-file/node` re-exports its types with `export type { ... }`, which TypeScript does not allow downstream packages to augment.

## Supported validation types

| `type`         | OOXML `type` | Required fields                                                                   |
| -------------- | ------------ | --------------------------------------------------------------------------------- |
| `'list'`       | `list`       | `values: string[]` **or** `valuesRange: string` (e.g. `'$E$4:$G$4'`)              |
| `'integer'`    | `whole`      | `operator`, `value` (and `value2` for `'...'` / `'!...'`)                         |
| `'decimal'`    | `decimal`    | same as `integer`                                                                 |
| `'date'`       | `date`       | `operator`, `value` (and `value2` for between operators); values: `Date`/`number` |
| `'time'`       | `time`       | same as `date`; `Date` values are reduced to fractional time-of-day               |
| `'textLength'` | `textLength` | same as `integer`                                                                 |
| `'custom'`     | `custom`     | `formula: string`                                                                 |
| `'any'`        | (no `type`)  | Only metadata (used to attach a tooltip/error message without restricting input)  |

Operators (mirror the symbols used by `conditionalFormatting`):

| Symbol | OOXML                |
| ------ | -------------------- |
| `<`    | `lessThan`           |
| `>`    | `greaterThan`        |
| `<=`   | `lessThanOrEqual`    |
| `>=`   | `greaterThanOrEqual` |
| `=`    | `equal`              |
| `!=`   | `notEqual`           |
| `...`  | `between`            |
| `!...` | `notBetween`         |

## Common options (any type)

| Field              | Default  | OOXML attribute                                                                  |
| ------------------ | -------- | -------------------------------------------------------------------------------- |
| `error`            | —        | `error`                                                                          |
| `errorTitle`       | —        | `errorTitle`                                                                     |
| `errorStyle`       | `'stop'` | `errorStyle`                                                                     |
| `input`            | —        | `prompt`                                                                         |
| `inputTitle`       | —        | `promptTitle`                                                                    |
| `allowBlank`       | `true`   | `allowBlank`                                                                     |
| `showErrorMessage` | `true`   | `showErrorMessage`                                                               |
| `showInputMessage` | `true`   | `showInputMessage`                                                               |
| `showDropdown`     | `true`   | `showDropDown` (inverted: OOXML `"1"` means hide) — only meaningful for `'list'` |

Maximum lengths enforced (Excel limits): titles ≤ 32 characters, messages ≤ 255 characters, inline list formulas ≤ 255 characters.

## Limitations

- **Sheet-level only.** This feature reads `dataValidation` from `sheetOptions`. Per-cell validation (attaching a `validation` property to a cell object) is **not** supported because `write-excel-file`'s feature API does not expose sheet data to the transform hook. If you need per-cell validation, encode it as a sheet-level rule with a single-cell range:

  ```ts
  { cellRange: { from: { row: 5, column: 3 }, to: { row: 5, column: 3 } }, validation: { ... } }
  ```

- **Inline list commas and quotes.** Excel uses `,` as the list separator inside `<formula1>"…"</formula1>` and `"` as the surrounding delimiter — neither has a working escape. Values containing `,` or `"` are rejected; use `valuesRange` to reference a range of cells instead.

## Examples

The [`examples/`](./examples) folder contains runnable TypeScript scripts — one per validation type — that emit `.xlsx` files next to themselves. Open the generated files in Excel/Numbers/LibreOffice to confirm each rule works.

```sh
npm install
npx tsx examples/list-inline.ts          # run a single example
npm run examples                          # run them all
```

See [`examples/README.md`](./examples/README.md) for the full list.

## Development

```sh
npm install
npm test          # vitest
npm run typecheck # tsc --noEmit
npm run build     # emit dist/ via tsc
```

## License

[MIT](./LICENSE)
