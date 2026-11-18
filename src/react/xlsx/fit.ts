import { formatGeneral } from '../../core/xlsx/numfmt'

/** px size of a CSS font size ('11pt', px number), `base` when unset */
export const fontPx = (v: string | number | undefined, base: number) =>
  typeof v === 'string' && v.endsWith('pt') ? (parseFloat(v) * 4) / 3 : typeof v === 'number' ? v : base

/** a digit's width in the fonts of a sheet (Calibri 11 pt: 7 px of its 14.7) */
const DIGIT = 0.48

/** a number as the General format writes it: plain digits, a point, an exponent */
const PLAIN = /^-?(?:\d+\.?\d*|\.\d+)(?:E[+-]?\d+)?$/i

/**
 * Excel-like: a General number gives up decimals to fit the column, then goes scientific when its integer part does
 * not fit, and shows ### only when not even that fits. A formatted number (1,234.50 / 12%) cannot change: ###.
 */
export function fitNumber(text: string, width: number, px: number): string {
  const room = Math.floor(width / (px * DIGIT) + 1e-9)
  if (text.length <= room) return text
  if (room < 1 || !PLAIN.test(text)) return '#'.repeat(Math.max(1, room))
  return formatGeneral(Number(text), room)
}
