const LETTERS_COUNT = 26;

function getColumnLetter(columnIndex: number): string {
  const prefix = Math.floor(columnIndex / LETTERS_COUNT);
  const letter = String.fromCharCode(97 + (columnIndex % LETTERS_COUNT)).toUpperCase();
  if (prefix === 0) {
    return letter;
  }
  return getColumnLetter(prefix - 1) + letter;
}

export default function getCellCoordinate(rowIndex: number, columnIndex: number): string {
  return `${getColumnLetter(columnIndex)}${rowIndex + 1}`;
}
