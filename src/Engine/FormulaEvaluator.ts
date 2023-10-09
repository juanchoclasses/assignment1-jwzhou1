// FormulaEvaluator class is responsible for evaluating mathematical formulas
// expressed as arrays of tokens. It provides methods for parsing and calculating
// the result of these formulas while handling various error conditions.
// It relies on the SheetMemory class for accessing cell values and error messages.
// The class maintains state, error flags, and the last calculated result.
// It can be used to evaluate mathematical expressions in spreadsheet cells.

import SheetMemory from "./SheetMemory";
import Cell from "./Cell";
import { ErrorMessages } from "./GlobalDefinitions";

export class FormulaEvaluator {
  private _errorOccured: boolean = false;
  private _errorMessage: string = "";
  private _currentFormula: FormulaType = [];
  private _lastResult: number = 0;
  private _sheetMemory: SheetMemory;
  private _result: number = 0;

  constructor(memory: SheetMemory) {
    this._sheetMemory = memory;
  }

  // The evaluate method is the entry point for formula evaluation.
  // It sets initial values, handles empty formula cases, and calls
  // the expression method to evaluate the formula.
  evaluate(formula: FormulaType) {
    this._currentFormula = [...formula];
    this._lastResult = 0;

    if (formula.length === 0) {
      this._errorMessage = ErrorMessages.emptyFormula;
      this._result = 0;
      return;
    }

    this._errorOccured = false;
    this._errorMessage = "";

    const resultValue = this.expression();
    this._result = resultValue;

    if (this._currentFormula.length > 0 && !this._errorOccured) {
      this._errorOccured = true;
      this._errorMessage = ErrorMessages.invalidFormula;
    }

    if (this._errorOccured) {
      this._result = this._lastResult;
    }
  }

  // Getters to retrieve the error message.
  public get error(): string {
    return this._errorMessage;
  }

  // Getters to retrieve the final result.
  public get result(): number {
    return this._result;
  }

  // The expression method evaluates expressions containing addition and subtraction.
  // It calls the term method to evaluate terms containing multiplication and division.
  private expression(): number {
    if (this._errorOccured) {
      return this._lastResult;
    }
    let result = this.term();
    while (
      this._currentFormula.length > 0 &&
      (this._currentFormula[0] === "+" || this._currentFormula[0] === "-")
    ) {
      const operator = this._currentFormula.shift();
      const term = this.term();
      if (operator === "+") {
        result += term;
      } else if (operator === "-") {
        result -= term;
      }
    }
    this._lastResult = result;
    return result;
  }

  // The allowedOperator method checks if the current token is an allowed operator.
  // Allowed operators are multiplication, division, and change of sign.
  private allowedOperator(): boolean {
    return (
      this._currentFormula[0] === "*" ||
      this._currentFormula[0] === "/" ||
      this._currentFormula[0] === "+/-"
    );
  }

   // The term method evaluates terms containing multiplication, division, and sign changes.
    // It calls the factor method to evaluate factors containing numbers, cell references,
  private term(): number {
    if (this._errorOccured) {
      return this._lastResult;
    }
    let result = this.factor();
    while (this._currentFormula.length > 0 && this.allowedOperator()) {
      const operator = this._currentFormula.shift();
      if (operator === "+/-") {
        if (result !== 0) {
          result = -result;
        }
      } else {
        const factor = this.factor();
        if (operator === "*") {
          result *= factor;
        } else if (operator === "/") {
          if (factor === 0) {
            this._errorOccured = true;
            this._errorMessage = ErrorMessages.divideByZero;
            this._lastResult = Infinity;
            return Infinity;
          } else {
            result /= factor;
          }
        }
      }
    }
    this._lastResult = result;
    return result;
  }
  
  // The factor method evaluates factors, which can be numbers, parentheses, or cell references.
  // It calls the expression method to evaluate expressions in parentheses.
  private factor(): number {
    if (this._errorOccured) {
      return this._lastResult;
    }
    let result = 0;
    if (this._currentFormula.length === 0) {
      this._errorOccured = true;
      this._errorMessage = ErrorMessages.partial;
      return result;
    }
    let token = this._currentFormula.shift();
    if (token === "") {
      this._errorOccured = true;
      this._errorMessage = ErrorMessages.invalidFormula;
      return result;
    }
    if (this.isNumber(token)) {
      result = Number(token);
      this._lastResult = result;
    } else if (token === "(") {
      result = this.expression();
      if (
        this._currentFormula.length === 0 ||
        this._currentFormula.shift() !== ")"
      ) {
        this._errorOccured = true;
        this._errorMessage = ErrorMessages.missingParentheses;
        this._lastResult = result;
      }
    } else if (this.isCellReference(token)) {
      [result, this._errorMessage] = this.getCellValue(token);
      if (this._errorMessage !== "") {
        this._errorOccured = true;
        this._lastResult = result;
      }
    } else {
      this._errorOccured = true;
      this._errorMessage = ErrorMessages.invalidFormula;
    }
    return result;
  }

  /**
   * 
   * @param token 
   * @returns true if the toke can be parsed to a number
   */
  isNumber(token: TokenType): boolean {
    return !isNaN(Number(token));
  }

   /**
   * 
   * @param token
   * @returns true if the token is a cell reference
   * 
   */
  isCellReference(token: TokenType): boolean {
    return Cell.isValidCellLabel(token);
  }

   /**
   * 
   * @param token
   * @returns [value, ""] if the cell formula is not empty and has no error
   * @returns [0, error] if the cell has an error
   * @returns [0, ErrorMessages.invalidCell] if the cell formula is empty
   * 
   */
   getCellValue(token: TokenType): [number, string] {
    // Retrieve the cell and its formula and error message
    let cell = this._sheetMemory.getCellByLabel(token);
    let cellFormula = cell.getFormula();
    let cellError = cell.getError();
    // Check if the cell has an error, return 0 with the error message
    if (cellError !== "" && cellError !== ErrorMessages.emptyFormula) {
      return [0, cellError];
    }
    // Check if the cell formula is empty, return 0 with an appropriate error message
    if (cellFormula.length === 0) {
      return [0, ErrorMessages.invalidCell];
    }
    // Retrieve the cell value and return it along with an empty error message
    let cellValue = cell.getValue();
    return [cellValue, ""];
  }
  
}

export default FormulaEvaluator;
