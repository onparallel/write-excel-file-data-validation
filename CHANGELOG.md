# Changelog

## [1.0.2](https://github.com/onparallel/write-excel-file-data-validation/compare/v1.0.1...v1.0.2) (2026-10-09)


### Bug Fixes

* accept every Node.js 22 version in "engines" again ([4aa6cf8](https://github.com/onparallel/write-excel-file-data-validation/commit/4aa6cf860ef90c0766b86a4b0eb2a0ee2a40a6ba))

## [1.0.1](https://github.com/onparallel/write-excel-file-data-validation/compare/v1.0.0...v1.0.1) (2026-10-09)


### Bug Fixes

* allow require() from CommonJS ([9f0fe3b](https://github.com/onparallel/write-excel-file-data-validation/commit/9f0fe3b1baaeea9e888198379615f5b5f855e004)), closes [#20](https://github.com/onparallel/write-excel-file-data-validation/issues/20)

## [1.0.0](https://github.com/onparallel/write-excel-file-data-validation/compare/v0.1.0...v1.0.0) (2026-06-08)


### ⚠ BREAKING CHANGES

* peerDependency `write-excel-file` bumped from `^4.0.0` to `^4.1.1`. Drops the bundled `getCellCoordinate`/`convertDateToExcelSerial` helpers in favor of the equivalents exported by `write-excel-file/utility` in 4.1.1.

### Code Refactoring

* use getCellAddress/convertDateToSerialNumber from write-excel-file/utility ([d331abf](https://github.com/onparallel/write-excel-file-data-validation/commit/d331abf83804f12ffaaacd596371fca094e73f1d))
