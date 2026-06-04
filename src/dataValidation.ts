import {
	insertElementMarkupAccordingToOrderOfSiblings,
	getOrderOfSiblings,
	sanitizeAttributeValue,
	sanitizeTextContent
} from 'write-excel-file/utility'
import type { Feature } from 'write-excel-file/node'

import getCellCoordinate from './getCellCoordinate.js'
import convertDateToExcelSerial from './convertDateToExcelSerial.js'
import type {
	DataValidation,
	DataValidationOperator,
	DataValidationRule,
	DataValidationSheetOptions
} from './types.js'

const MAX_TITLE_LENGTH = 32
const MAX_MESSAGE_LENGTH = 255

// `Feature<any>` rather than `Feature<unknown>` so the feature is assignable to
// `Feature<FileContent>` for any `FileContent` (Buffer/Blob/etc) chosen by the caller.
// This feature only transforms worksheet XML and never reads or writes file content,
// so the `FileContent` parameter is irrelevant here.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const dataValidation: Feature<any> = {
	files: {
		transform: {
			'xl/worksheets/sheet{id}.xml': {
				// `<dataValidations/>` must be placed between `<conditionalFormatting/>` and `<hyperlinks/>`
				// inside `<worksheet/>`. We use `transform()` (not `insert()`) so the inserter can place
				// the markup at the canonical sibling position even when other features have already
				// appended siblings (such as `<drawing/>`).
				transform: (xml, sheetOptions) => {
					const { dataValidation: rules } = sheetOptions as DataValidationSheetOptions
					if (rules && rules.length > 0) {
						const dataValidationsXml = getDataValidationsXml(rules)
						const order = getOrderOfSiblings('xl/worksheets/sheet{id}.xml', 'worksheet')
						if (!order) {
							throw new Error(
								'write-excel-file did not return an order of siblings for `worksheet`'
							)
						}
						return insertElementMarkupAccordingToOrderOfSiblings(
							xml,
							dataValidationsXml,
							order,
							'worksheet'
						)
					}
					return xml
				}
			}
		}
	}
}

export default dataValidation

function getDataValidationsXml(rules: DataValidationRule[]): string {
	let xml = `<dataValidations count="${rules.length}">`
	for (const rule of rules) {
		xml += getDataValidationXml(rule)
	}
	xml += '</dataValidations>'
	return xml
}

function getDataValidationXml(rule: DataValidationRule): string {
	const { cellRange, validation } = rule

	if (!cellRange) {
		throw new Error('A data validation rule must specify a `cellRange`')
	}

	if (!validation) {
		throw new Error('A data validation rule must specify a `validation`')
	}

	const { from, to } = cellRange

	if (!from || !to) {
		throw new Error('A data validation `cellRange` must specify both `from` and `to`')
	}

	if (from.row < 1 || from.column < 1 || to.row < 1 || to.column < 1) {
		throw new Error(
			`Data validation \`cellRange\` row and column indexes must be >= 1 (1-based), got: ${JSON.stringify(cellRange)}`
		)
	}

	if (to.row < from.row || to.column < from.column) {
		throw new Error(
			`Data validation \`cellRange\` \`to\` must be greater than or equal to \`from\`, got: ${JSON.stringify(cellRange)}`
		)
	}

	const sqref =
		getCellCoordinate(from.row - 1, from.column - 1) +
		':' +
		getCellCoordinate(to.row - 1, to.column - 1)

	const attributes: string[] = []

	const xlsxType = getXlsxType(validation.type)
	if (xlsxType) {
		attributes.push(`type="${xlsxType}"`)
	}

	if (hasOperator(validation.type)) {
		attributes.push(
			`operator="${getXlsxOperatorName((validation as { operator: DataValidationOperator }).operator)}"`
		)
	}

	const {
		error,
		errorTitle,
		errorStyle,
		input,
		inputTitle,
		allowBlank,
		showErrorMessage,
		showInputMessage
	} = validation

	// OOXML defaults `allowBlank` to false; we default it to true (matching xlsxwriter).
	attributes.push(`allowBlank="${allowBlank === false ? '0' : '1'}"`)

	// `showDropDown` in OOXML is inverted: setting it to `1` HIDES the dropdown.
	if (validation.type === 'list' && validation.showDropdown === false) {
		attributes.push('showDropDown="1"')
	}

	// OOXML defaults `showErrorMessage`/`showInputMessage` to false; without them Excel does
	// not reject invalid input. Default both to true.
	attributes.push(`showErrorMessage="${showErrorMessage === false ? '0' : '1'}"`)
	attributes.push(`showInputMessage="${showInputMessage === false ? '0' : '1'}"`)

	if (errorStyle !== undefined) {
		validateErrorStyle(errorStyle)
		// OOXML default for `errorStyle` is `stop`. Skip the attribute in that case.
		if (errorStyle !== 'stop') {
			attributes.push(`errorStyle="${errorStyle}"`)
		}
	}

	if (errorTitle !== undefined) {
		validateTitleLength(errorTitle, 'errorTitle')
		attributes.push(`errorTitle="${sanitizeAttributeValue(errorTitle)}"`)
	}

	if (error !== undefined) {
		validateMessageLength(error, 'error')
		attributes.push(`error="${sanitizeAttributeValue(error)}"`)
	}

	if (inputTitle !== undefined) {
		validateTitleLength(inputTitle, 'inputTitle')
		attributes.push(`promptTitle="${sanitizeAttributeValue(inputTitle)}"`)
	}

	if (input !== undefined) {
		validateMessageLength(input, 'input')
		attributes.push(`prompt="${sanitizeAttributeValue(input)}"`)
	}

	attributes.push(`sqref="${sqref}"`)

	const [formula1, formula2] = getFormulas(validation)

	let xml = `<dataValidation ${attributes.join(' ')}>`
	if (formula1 !== undefined) {
		xml += `<formula1>${sanitizeTextContent(formula1)}</formula1>`
	}
	if (formula2 !== undefined) {
		xml += `<formula2>${sanitizeTextContent(formula2)}</formula2>`
	}
	xml += '</dataValidation>'

	return xml
}

function getXlsxType(type: DataValidation['type']): string | undefined {
	switch (type) {
		case 'list':
			return 'list'
		case 'integer':
			return 'whole'
		case 'decimal':
			return 'decimal'
		case 'date':
			return 'date'
		case 'time':
			return 'time'
		case 'textLength':
			return 'textLength'
		case 'custom':
			return 'custom'
		case 'any':
			return undefined
		default:
			throw new Error(`Unknown data validation type: ${type as string}`)
	}
}

function hasOperator(type: DataValidation['type']): boolean {
	switch (type) {
		case 'integer':
		case 'decimal':
		case 'date':
		case 'time':
		case 'textLength':
			return true
		default:
			return false
	}
}

function getXlsxOperatorName(operator: DataValidationOperator): string {
	switch (operator) {
		case '<':
			return 'lessThan'
		case '>':
			return 'greaterThan'
		case '<=':
			return 'lessThanOrEqual'
		case '>=':
			return 'greaterThanOrEqual'
		case '=':
			return 'equal'
		case '!=':
			return 'notEqual'
		case '...':
			return 'between'
		case '!...':
			return 'notBetween'
		default:
			throw new Error(`Unknown data validation operator: ${operator as string}`)
	}
}

function isBetweenOperator(operator: DataValidationOperator): boolean {
	return operator === '...' || operator === '!...'
}

type Formulas = [string | undefined, string | undefined] | [string] | []

function getFormulas(validation: DataValidation): Formulas {
	switch (validation.type) {
		case 'list': {
			if ('values' in validation) {
				return [getListFormulaFromValues(validation.values)]
			}
			if ('valuesRange' in validation) {
				return [validation.valuesRange]
			}
			throw new Error('A `list` data validation must specify either `values` or `valuesRange`')
		}

		case 'custom': {
			if (validation.formula === undefined) {
				throw new Error('A `custom` data validation must specify a `formula`')
			}
			return [validation.formula]
		}

		case 'any':
			return []

		case 'integer':
		case 'decimal':
		case 'textLength': {
			if (validation.value === undefined) {
				throw new Error(`A \`${validation.type}\` data validation must specify a \`value\``)
			}
			if (isBetweenOperator(validation.operator)) {
				if ((validation as { value2?: number }).value2 === undefined) {
					throw new Error(
						`A \`${validation.type}\` data validation with operator \`${validation.operator}\` must specify a \`value2\``
					)
				}
				return [
					formatNumericValue(validation.value, validation.type),
					formatNumericValue((validation as { value2: number }).value2, validation.type)
				]
			}
			return [formatNumericValue(validation.value, validation.type)]
		}

		case 'date':
		case 'time': {
			if (validation.value === undefined) {
				throw new Error(`A \`${validation.type}\` data validation must specify a \`value\``)
			}
			if (isBetweenOperator(validation.operator)) {
				const value2 = (validation as { value2?: Date | number }).value2
				if (value2 === undefined) {
					throw new Error(
						`A \`${validation.type}\` data validation with operator \`${validation.operator}\` must specify a \`value2\``
					)
				}
				return [
					formatDateOrTimeValue(validation.value, validation.type),
					formatDateOrTimeValue(value2, validation.type)
				]
			}
			return [formatDateOrTimeValue(validation.value, validation.type)]
		}
	}
}

function getListFormulaFromValues(values: string[]): string {
	if (!Array.isArray(values)) {
		throw new Error(
			`A \`list\` data validation \`values\` must be an array, got: ${String(values)}`
		)
	}
	if (values.length === 0) {
		throw new Error('A `list` data validation `values` must contain at least one value')
	}
	for (const value of values) {
		if (typeof value !== 'string') {
			throw new Error(
				`A \`list\` data validation \`values\` must be strings, got: ${String(value)}`
			)
		}
		if (value.indexOf(',') !== -1) {
			throw new Error(
				`A \`list\` data validation \`values\` cannot contain commas (used by Excel as the separator inside the inline list formula). Use \`valuesRange\` to reference a range of cells instead. Got: "${value}"`
			)
		}
		if (value.indexOf('"') !== -1) {
			throw new Error(
				`A \`list\` data validation \`values\` cannot contain double-quote characters. Use \`valuesRange\` to reference a range of cells instead. Got: "${value}"`
			)
		}
	}
	const serialized = values.join(',')
	if (serialized.length + 2 > MAX_MESSAGE_LENGTH) {
		throw new Error(
			`A \`list\` data validation \`values\` serialized to "${serialized}" exceeds the maximum length of ${MAX_MESSAGE_LENGTH} characters. Use \`valuesRange\` instead to reference a range of cells.`
		)
	}
	return `"${serialized}"`
}

function formatNumericValue(value: number, type: string): string {
	if (typeof value !== 'number' || !Number.isFinite(value)) {
		throw new Error(
			`A \`${type}\` data validation \`value\` must be a finite number, got: ${String(value)}`
		)
	}
	return String(value)
}

function formatDateOrTimeValue(value: Date | number, type: 'date' | 'time'): string {
	let serial: number
	if (value instanceof Date) {
		serial = convertDateToExcelSerial(value)
	} else if (typeof value === 'number') {
		serial = value
	} else {
		throw new Error(
			`A \`${type}\` data validation \`value\` must be a Date or a number, got: ${String(value)}`
		)
	}
	if (!Number.isFinite(serial)) {
		throw new Error(
			`A \`${type}\` data validation \`value\` is not a finite number (possibly an invalid Date), got: ${String(value)}`
		)
	}
	// For `time`, Excel expects a fractional value in [0, 1) representing the time of day.
	// A `Date` carries both a date and a time, so reduce the serial to its fractional part.
	if (type === 'time') {
		serial = serial - Math.floor(serial)
	}
	return String(serial)
}

function validateErrorStyle(errorStyle: string): void {
	if (errorStyle !== 'stop' && errorStyle !== 'warning' && errorStyle !== 'information') {
		throw new Error(
			`Unknown data validation \`errorStyle\`: ${errorStyle}. Expected \`stop\`, \`warning\` or \`information\``
		)
	}
}

function validateTitleLength(title: string, fieldName: string): void {
	if (typeof title !== 'string') {
		throw new Error(`Data validation \`${fieldName}\` must be a string`)
	}
	if (title.length > MAX_TITLE_LENGTH) {
		throw new Error(
			`Data validation \`${fieldName}\` is longer than ${MAX_TITLE_LENGTH} characters: "${title}"`
		)
	}
}

function validateMessageLength(message: string, fieldName: string): void {
	if (typeof message !== 'string') {
		throw new Error(`Data validation \`${fieldName}\` must be a string`)
	}
	if (message.length > MAX_MESSAGE_LENGTH) {
		throw new Error(
			`Data validation \`${fieldName}\` is longer than ${MAX_MESSAGE_LENGTH} characters: "${message}"`
		)
	}
}
