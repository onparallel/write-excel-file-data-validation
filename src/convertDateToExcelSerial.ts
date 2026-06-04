// Excel serial date = count of days since 1900-01-01 (with the 1900-leap-year quirk).
// 70 * 365 + 19 = days between 1900-01-01 (Excel epoch) and 1970-01-01 (Unix epoch),
// accounting for 19 leap days in that range.
const DAYS_BEFORE_UNIX_EPOCH = 70 * 365 + 19
const MS_PER_DAY = 24 * 60 * 60 * 1000

export default function convertDateToExcelSerial(date: Date): number {
	return date.getTime() / MS_PER_DAY + DAYS_BEFORE_UNIX_EPOCH
}
