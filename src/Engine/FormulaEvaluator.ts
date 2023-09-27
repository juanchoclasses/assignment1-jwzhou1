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

  public get error(): string {
    return this._errorMessage;
  }

  public get result(): number {
    return this._result;
  }

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

  private allowedOperator(): boolean {
    return (
      this._currentFormula[0] === "*" ||
      this._currentFormula[0] === "/" ||
      this._currentFormula[0] === "+/-"
    );
  }

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

  isNumber(token: TokenType): boolean {
    return !isNaN(Number(token));
  }

  isCellReference(token: TokenType): boolean {
    return Cell.isValidCellLabel(token);
  }

  getCellValue(token: TokenType): [number, string] {
    if (token === "") {
      return [0, ErrorMessages.invalidCell];
    } else {
      let cell = this._sheetMemory.getCellByLabel(token);
      let cellValue = cell.getValue();
      let cellError = cell.getError();
      if (cellError !== "") {
        return [0, cellError];
      } else {
        return [cellValue, ""];
      }
    }
  }
}

export default FormulaEvaluator;
