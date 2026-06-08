# Changelog

## [1.0.0](https://github.com/onparallel/write-excel-file-data-validation/compare/v0.1.0...v1.0.0) (2026-06-08)


### ⚠ BREAKING CHANGES

* peerDependency `write-excel-file` bumped from `^4.0.0` to `^4.1.1`. Drops the bundled `getCellCoordinate`/`convertDateToExcelSerial` helpers in favor of the equivalents exported by `write-excel-file/utility` in 4.1.1.

### Code Refactoring

* use getCellAddress/convertDateToSerialNumber from write-excel-file/utility ([d331abf](https://github.com/onparallel/write-excel-file-data-validation/commit/d331abf83804f12ffaaacd596371fca094e73f1d))
