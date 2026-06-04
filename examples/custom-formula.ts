import writeXlsxFile, { type SheetData, type SheetOptions } from 'write-excel-file/node'
import dataValidation, { type DataValidationSheetOptions } from '../src/index.js'

// Only accept even numbers, using a custom boolean formula.
const data: SheetData = [[{ value: 'Even number', fontWeight: 'bold' }], [4], [2], [6]]

const sheetOptions: SheetOptions<any> & DataValidationSheetOptions = {
	sheet: 'Custom formula',
	dataValidation: [
		{
			cellRange: { from: { row: 2, column: 1 }, to: { row: 4, column: 1 } },
			validation: {
				type: 'custom',
				formula: '=MOD(A2,2)=0',
				error: 'Must be an even number'
			}
		}
	]
}

const output = new URL('./custom-formula.xlsx', import.meta.url).pathname

await writeXlsxFile(data, sheetOptions, { features: [dataValidation] }).toFile(output)

console.log(`Wrote ${output}`)
