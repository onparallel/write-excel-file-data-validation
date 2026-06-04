import writeXlsxFile, { type SheetData, type SheetOptions } from 'write-excel-file/node'
import dataValidation, { type DataValidationSheetOptions } from '../src/index.js'

// Reject text longer than 50 characters.
const data: SheetData = [
	[{ value: 'Short description (<=50 chars)', fontWeight: 'bold' }],
	['Hello world']
]

const sheetOptions: SheetOptions<any> & DataValidationSheetOptions = {
	sheet: 'Text length',
	dataValidation: [
		{
			cellRange: { from: { row: 2, column: 1 }, to: { row: 2, column: 1 } },
			validation: {
				type: 'textLength',
				operator: '<=',
				value: 50,
				error: 'Maximum 50 characters'
			}
		}
	]
}

const output = new URL('./text-length.xlsx', import.meta.url).pathname

await writeXlsxFile(data, sheetOptions, { features: [dataValidation] }).toFile(output)

console.log(`Wrote ${output}`)
