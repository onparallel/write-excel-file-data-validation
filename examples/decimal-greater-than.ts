import writeXlsxFile, { type SheetData, type SheetOptions } from 'write-excel-file/node'
import dataValidation, { type DataValidationSheetOptions } from '../src/index.js'

// A decimal strictly greater than 0.5. Demonstrates the single-value operator form.
const data: SheetData = [[{ value: 'Ratio (> 0.5)', fontWeight: 'bold' }], [0.75], [0.9]]

const sheetOptions: SheetOptions<any> & DataValidationSheetOptions = {
	sheet: 'Decimal',
	dataValidation: [
		{
			cellRange: { from: { row: 2, column: 1 }, to: { row: 3, column: 1 } },
			validation: {
				type: 'decimal',
				operator: '>',
				value: 0.5,
				error: 'Value must be greater than 0.5'
			}
		}
	]
}

const output = new URL('./decimal-greater-than.xlsx', import.meta.url).pathname

await writeXlsxFile(data, sheetOptions, { features: [dataValidation] }).toFile(output)

console.log(`Wrote ${output}`)
