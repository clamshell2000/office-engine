/** Small number-format engine for TEXT(): 0, 0.00, #,##0, 0%, literals, sections, dates/times. */
import { formatDate, isDateFormat, splitSections } from './dates'
import { ERR, FormulaError, type Value } from './types'
import { numToText, parseNumberText } from './values'

/** Round half away from zero to d decimals, removing binary noise first. */
export function roundHalf(x: number, d: number): number {
  if (d > 15) return x
  const f = Math.pow(10, Math.abs(d))
  const ax = Math.abs(x)
  let y = d >= 0 ? ax * f : ax / f
  y = Math.floor(Number(y.toPrecision(15)) + 0.5)
  const r = d >= 0 ? y / f : y * f
  return x < 0 ? -r : r
}

/** Round a nonnegative value's 15 significant decimal digits without returning to binary. */
export function fixedDecimal(v: number, decimals: number): string {
  if (!Number.isFinite(v)) return String(v)
  const [mantissa, exponent] = v.toExponential(14).split('e')
  let coefficient = BigInt(mantissa.replace('.', ''))
  const shift = Number(exponent) - 14 + decimals
  let digits: string
  if (shift >= 0) digits = coefficient.toString() + '0'.repeat(shift)
  else {
    const divisor = 10n ** BigInt(-shift)
    coefficient = (coefficient + divisor / 2n) / divisor
    digits = coefficient.toString()
  }
  digits = digits.padStart(decimals + 1, '0')
  return decimals ? digits.slice(0, -decimals) + '.' + digits.slice(-decimals) : digits
}

/**
 * Index of the format section Excel applies to n: the second for negatives, the third for zero,
 * otherwise the first. An empty second section falls back to the first, but an empty third section
 * means zero shows nothing, which is returned as null.
 */
export function pickSection(n: number, count: number, isEmpty: (i: number) => boolean): number | null {
  if (n < 0 && count > 1 && !isEmpty(1)) return 1
  if (n === 0 && count > 2) return isEmpty(2) ? null : 2
  return 0
}

function formatNumber(x: number, sec: string, signed: boolean): string {
  let prefix = ''
  let suffix = ''
  let pattern = ''
  let pct = 0
  const lit = (s: string) => {
    if (pattern) suffix += s
    else prefix += s
  }
  for (let i = 0; i < sec.length; i++) {
    const c = sec[i]
    if (c === '"') {
      const j = sec.indexOf('"', i + 1)
      const end = j < 0 ? sec.length : j
      lit(sec.slice(i + 1, end))
      i = end
    } else if (c === '\\') lit(sec[++i] ?? '')
    else if (c === '_') {
      lit(' ')
      i++
    } else if (c === '*') i++
    else if (c === '[') {
      const j = sec.indexOf(']', i)
      i = j < 0 ? sec.length : j
    } else if ('0#?'.includes(c) || ((c === '.' || c === ',') && !suffix)) pattern += c
    else {
      if (c === '%') pct++
      lit(c)
    }
  }
  let p = pattern
  let v = x * Math.pow(100, pct)
  while (p.endsWith(',')) {
    p = p.slice(0, -1)
    v /= 1000
  }
  const dot = p.indexOf('.')
  const ip = dot < 0 ? p : p.slice(0, dot)
  const dp = dot < 0 ? '' : p.slice(dot + 1)
  const decs = (dp.match(/[0#?]/g) ?? []).length
  const minDec = (dp.match(/[0?]/g) ?? []).length
  const minInt = (ip.match(/[0?]/g) ?? []).length
  // the digits as a cell shows them (15 significant ones, no way back through a double: 1E+21 is not 999999999999999868928)
  const fixed = fixedDecimal(Math.abs(v), decs)
  let [is, ds = ''] = fixed.split('.')
  if (is === '0' && minInt === 0) is = ''
  is = is.padStart(minInt, '0')
  if (ip.includes(',')) is = is.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
  while (ds.length > minDec && ds.endsWith('0')) ds = ds.slice(0, -1)
  const body = p === '' ? '' : is + (dot >= 0 ? `.${ds}` : '')
  const neg = signed && v < 0 && /[1-9]/.test(fixed)
  return (neg ? '-' : '') + prefix + body + suffix
}

/** TEXT(value, format). Throws FormulaError on invalid input. */
export function formatText(v: Value, fmt: string): string {
  if (v instanceof FormulaError) throw v
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  let n: number
  const sections = splitSections(fmt)
  if (typeof v === 'string') {
    const p = parseNumberText(v)
    if (p === null) {
      const textSec = sections.length > 3 ? sections[3] : sections.find((s) => s.includes('@'))
      return textSec === undefined ? v : textSec.replace(/"([^"]*)"/g, '$1').replace(/@/g, v)
    }
    n = p
  } else n = v ?? 0
  const idx = pickSection(n, sections.length, (i) => sections[i] === '')
  if (idx === null) return ''
  const sec = sections[idx]
  const signed = idx !== 1
  if (/^\s*general\s*$/i.test(sec) || sec === '') return numToText(signed ? n : Math.abs(n))
  if (isDateFormat(sec)) {
    if (n < 0) throw ERR.VALUE
    return formatDate(n, sec)
  }
  return formatNumber(signed ? n : Math.abs(n), sec, signed)
}
