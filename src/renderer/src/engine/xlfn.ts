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
  return mapOutsideStrings(formula, (p) =>
    p.replace(/(^|[^A-Za-z0-9_.])([A-Za-z][A-Za-z0-9.]*)\s*\(/g, (m, pre: string, name: string) => {
      const up = name.toUpperCase()
      if (WS.has(up)) return `${pre}_xlfn._xlws.${name}(`
      if (FUTURE_SET.has(up)) return `${pre}_xlfn.${name}(`
      return m
    })
  )
}
