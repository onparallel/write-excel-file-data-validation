# Examples

Runnable TypeScript scripts. Each one writes its own `.xlsx` next to the script.

| Example                                                        | Demonstrates                                                   |
| -------------------------------------------------------------- | -------------------------------------------------------------- |
| [`list-inline.ts`](./list-inline.ts)                           | A dropdown with literal values defined inline.                 |
| [`list-range.ts`](./list-range.ts)                             | A dropdown sourcing options from a range of cells.             |
| [`integer-between.ts`](./integer-between.ts)                   | An integer in `[1, 10]` with `stop`-style error popup.         |
| [`decimal-greater-than.ts`](./decimal-greater-than.ts)         | A decimal strictly greater than `0.5`.                         |
| [`date-between.ts`](./date-between.ts)                         | A date constrained to 2024 (Date → Excel serial).              |
| [`time-of-day.ts`](./time-of-day.ts)                           | A time-of-day between 09:00 and 17:00 (Date → fractional day). |
| [`text-length.ts`](./text-length.ts)                           | Text length ≤ 50 characters.                                   |
| [`custom-formula.ts`](./custom-formula.ts)                     | A boolean formula (`=MOD(A2,2)=0` — even numbers only).        |
| [`input-and-error-messages.ts`](./input-and-error-messages.ts) | Input tooltip + `warning`-style error popup.                   |

## Run

```sh
# Single example
npx tsx examples/list-inline.ts

# All examples
npm run examples
```
