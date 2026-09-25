// Excel functions HyperFormula lacks. Registered once, before any HyperFormula instance is built.
import { FunctionPlugin, FunctionArgumentType, HyperFormula, SimpleRangeValue, CellError, ErrorType } from 'hyperformula'
type InterpreterValue = any

function flatten(args: unknown[]): unknown[] {
  const out: unknown[] = []
  for (const a of args) {
    if (a instanceof SimpleRangeValue) for (const row of a.data) for (const v of row) out.push(v)
    else out.push(a)
  }
  return out
}
function toText(v: unknown): string {
  if (v === null || v === undefined || typeof v === 'symbol') return ''
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  if (typeof v === 'object' && v !== null && 'val' in (v as object)) return String((v as { val: unknown }).val) // rich numbers
  return String(v)
}

class ExtraTextPlugin extends FunctionPlugin {
  static implementedFunctions = {
    CONCAT: { method: 'concat', parameters: [{ argumentType: FunctionArgumentType.ANY }], repeatLastArgs: 1 },
    TEXTJOIN: {
      method: 'textjoin',
      parameters: [
        { argumentType: FunctionArgumentType.STRING },
        { argumentType: FunctionArgumentType.BOOLEAN },
        { argumentType: FunctionArgumentType.ANY }
      ],
      repeatLastArgs: 1
    }
  }
  concat(ast: any, state: any): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('CONCAT'), (...args: unknown[]) => {
      const vals = flatten(args)
      const err = vals.find((v) => v instanceof CellError)
      if (err) return err as CellError
      return vals.map(toText).join('')
    })
  }
  textjoin(ast: any, state: any): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('TEXTJOIN'), (delim: string, ignoreEmpty: boolean, ...rest: unknown[]) => {
      const vals = flatten(rest)
      const err = vals.find((v) => v instanceof CellError)
      if (err) return err as CellError
      const parts = vals.map(toText).filter((s) => !(ignoreEmpty && s === ''))
      const out = parts.join(delim)
      return out.length > 32767 ? new CellError(ErrorType.VALUE) : out
    })
  }
}

let registered = false
export function registerExtraFunctions(): void {
  if (registered) return
  registered = true
  HyperFormula.registerFunctionPlugin(ExtraTextPlugin, { enGB: { CONCAT: 'CONCAT', TEXTJOIN: 'TEXTJOIN' }, enUS: { CONCAT: 'CONCAT', TEXTJOIN: 'TEXTJOIN' } } as any)
}

/** Excel treats bare TRUE/FALSE as literals; HyperFormula only knows TRUE()/FALSE(). */
export const BOOLEAN_NAMES = [
  { name: 'TRUE', expression: '=TRUE()' },
  { name: 'FALSE', expression: '=FALSE()' }
]
