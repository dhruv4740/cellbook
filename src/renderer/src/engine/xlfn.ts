// Excel stores functions newer than Excel 2007 with an "_xlfn." prefix in the file (FILTER/SORT
// with "_xlfn._xlws."); without it Excel shows #NAME? when opening the file. The app keeps plain
// names in memory, strips the prefixes on load and adds them back on save.
const FUTURE = `ACOT ACOTH AGGREGATE ARABIC BASE BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV BITAND
BITLSHIFT BITOR BITRSHIFT BITXOR CEILING.MATH CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT
CHISQ.TEST COMBINA CONCAT CONFIDENCE.NORM CONFIDENCE.T COT COTH COVARIANCE.P COVARIANCE.S CSC CSCH DAYS DECIMAL
ERF.PRECISE ERFC.PRECISE EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FILTERXML FLOOR.MATH FLOOR.PRECISE
FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR FORMULATEXT GAMMA
GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAUSS HYPGEOM.DIST IFNA IFS IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH
IMTAN ISFORMULA ISOWEEKNUM LOGNORM.DIST LOGNORM.INV MAXIFS MINIFS MODE.MULT MODE.SNGL MUNIT NEGBINOM.DIST
NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC PERCENTRANK.EXC
PERCENTRANK.INC PERMUTATIONA PHI POISSON.DIST QUARTILE.EXC QUARTILE.INC QUERYSTRING RANK.AVG RANK.EQ RRI SEC SECH
SHEET SHEETS SKEW.P STDEV.P STDEV.S SWITCH T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TEXTJOIN UNICHAR
UNICODE VAR.P VAR.S WEBSERVICE WEIBULL.DIST XOR Z.TEST ANCHORARRAY RANDARRAY SEQUENCE SINGLE SORTBY UNIQUE
XLOOKUP XMATCH LET LAMBDA TEXTBEFORE TEXTAFTER TEXTSPLIT VSTACK HSTACK TOCOL TOROW CHOOSECOLS CHOOSEROWS DROP
TAKE EXPAND WRAPROWS WRAPCOLS ARRAYTOTEXT VALUETOTEXT ISOMITTED BYROW BYCOL MAP REDUCE SCAN MAKEARRAY`.split(/\s+/)
const WS = new Set(['FILTER', 'SORT'])
const FUTURE_SET = new Set(FUTURE)

function mapOutsideStrings(f: string, fn: (part: string) => string): string {
  return f.split(/("(?:[^"]|"")*")/g).map((p, i) => (i % 2 ? p : fn(p))).join('')
}

/** "_xlfn.XLOOKUP(" -> "XLOOKUP(", "_xlfn._xlws.FILTER(" -> "FILTER(", LET's "_xlpm.x" -> "x". */
export function stripFuturePrefixes(formula: string): string {
  if (formula.indexOf('_xl') === -1) return formula
  return mapOutsideStrings(formula, (p) => p.replace(/_xlfn\.(?:_xlws\.)?/gi, '').replace(/_xlpm\./gi, ''))
}

/** Adds the prefixes Excel expects in the file. */
export function addFuturePrefixes(formula: string): string {
  const out = mapOutsideStrings(formula, (p) =>
    p.replace(/(?<![A-Za-z0-9_.])([A-Za-z][A-Za-z0-9.]*)\s*\(/g, (m, name: string) => {
      const up = name.toUpperCase()
      if (WS.has(up)) return `_xlfn._xlws.${name}(`
      if (FUTURE_SET.has(up)) return `_xlfn.${name}(`
      return m
    })
  )
  return /\b(?:LET|LAMBDA)\s*\(/i.test(out) ? addParamPrefixes(out) : out
}

/** Top-level comma-separated arguments of the call whose "(" is at `open`. */
function callArgs(f: string, open: number): string[] {
  const args: string[] = []
  let depth = 0, start = open + 1, inStr = false
  for (let i = open + 1; i < f.length; i++) {
    const ch = f[i]
    if (ch === '"') inStr = !inStr
    else if (inStr) continue
    else if (ch === '(' || ch === '{') depth++
    else if ((ch === ')' || ch === '}') && depth > 0) depth--
    else if (ch === ')') { args.push(f.slice(start, i)); break }
    else if (ch === ',' && depth === 0) { args.push(f.slice(start, i)); start = i + 1 }
  }
  return args.map((a) => a.trim())
}

/** Excel stores LET/LAMBDA parameter names as "_xlpm.name"; without it Excel refuses to open the file. */
function addParamPrefixes(f: string): string {
  const names = new Set<string>()
  const re = /\b(LET|LAMBDA)\s*\(/gi
  let m: RegExpExecArray | null
  while ((m = re.exec(f))) {
    const args = callArgs(f, m.index + m[0].length - 1)
    const params = m[1].toUpperCase() === 'LET' ? args.filter((_, i) => i % 2 === 0 && i < args.length - 1) : args.slice(0, -1)
    for (const p of params) if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(p) && !p.startsWith('_xlpm.')) names.add(p.toLowerCase())
  }
  if (!names.size) return f
  return mapOutsideStrings(f, (p) =>
    p.replace(/(?<![A-Za-z0-9_.!$'\]])([A-Za-z_][A-Za-z0-9_.]*)(?![A-Za-z0-9_.(!])/g, (w, name: string) => (names.has(name.toLowerCase()) ? `_xlpm.${name}` : w))
  )
}
