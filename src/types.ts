export type DataValidationOperatorOnOneValue = "<" | "<=" | ">" | ">=" | "=" | "!=";

export type DataValidationOperatorOnTwoValues = "..." | "!...";

export type DataValidationOperator =
  | DataValidationOperatorOnOneValue
  | DataValidationOperatorOnTwoValues;

export type DataValidationErrorStyle = "stop" | "warning" | "information";

interface DataValidationCommonProperties {
  error?: string;
  errorTitle?: string;
  errorStyle?: DataValidationErrorStyle;
  input?: string;
  inputTitle?: string;
  allowBlank?: boolean;
  showErrorMessage?: boolean;
  showInputMessage?: boolean;
}

interface DataValidationListWithValues extends DataValidationCommonProperties {
  type: "list";
  values: string[];
  showDropdown?: boolean;
}

interface DataValidationListWithRange extends DataValidationCommonProperties {
  type: "list";
  valuesRange: string;
  showDropdown?: boolean;
}

interface DataValidationNumericOnOneValue extends DataValidationCommonProperties {
  type: "integer" | "decimal" | "textLength";
  operator: DataValidationOperatorOnOneValue;
  value: number;
}

interface DataValidationNumericOnTwoValues extends DataValidationCommonProperties {
  type: "integer" | "decimal" | "textLength";
  operator: DataValidationOperatorOnTwoValues;
  value: number;
  value2: number;
}

interface DataValidationDateOnOneValue extends DataValidationCommonProperties {
  type: "date" | "time";
  operator: DataValidationOperatorOnOneValue;
  value: Date | number;
}

interface DataValidationDateOnTwoValues extends DataValidationCommonProperties {
  type: "date" | "time";
  operator: DataValidationOperatorOnTwoValues;
  value: Date | number;
  value2: Date | number;
}

interface DataValidationCustom extends DataValidationCommonProperties {
  type: "custom";
  formula: string;
}

interface DataValidationAny extends DataValidationCommonProperties {
  type: "any";
}

export type DataValidation =
  | DataValidationListWithValues
  | DataValidationListWithRange
  | DataValidationNumericOnOneValue
  | DataValidationNumericOnTwoValues
  | DataValidationDateOnOneValue
  | DataValidationDateOnTwoValues
  | DataValidationCustom
  | DataValidationAny;

export interface DataValidationRule {
  cellRange: {
    from: { row: number; column: number };
    to: { row: number; column: number };
  };
  validation: DataValidation;
}

export interface DataValidationSheetOptions {
  dataValidation?: DataValidationRule[];
}
