import { fontPx } from './fit'

type Css = Record<string, string | number | undefined>
type Base = { fontFamily?: string; fontSize?: string | number }

let ctx: CanvasRenderingContext2D | null | undefined
const canvas = () => (ctx === undefined ? (ctx = typeof document === 'undefined' ? null : document.createElement('canvas').getContext('2d')) : ctx)

/** the CSS font of a cell: its style over the workbook's base font */
export function fontOf(css: Css, base: Base, basePx: number): string {
  const px = fontPx(css.fontSize as string | number | undefined, basePx)
  return `${css.fontStyle === 'italic' ? 'italic ' : ''}${css.fontWeight === 700 ? 'bold ' : ''}${px}px ${css.fontFamily ?? base.fontFamily ?? 'sans-serif'}`
}

/** width in px of the widest line of `text` as a cell with this style draws it */
export function textWidth(text: string, css: Css, base: Base, basePx: number): number {
  const c = canvas()
  if (!c) return text.length * fontPx(css.fontSize as string | number | undefined, basePx) * 0.56
  c.font = fontOf(css, base, basePx)
  let w = 0
  for (const line of text.split('\n')) w = Math.max(w, c.measureText(line).width)
  return w
}

/** height in px of `text` wrapped in a cell `width` px wide (the line height and bottom padding of `.de-cell`) */
export function wrapHeight(text: string, css: Css, base: Base, basePx: number, width: number): number {
  const el = document.createElement('div')
  Object.assign(el.style, {
    position: 'fixed', left: '-9999px', top: '0', visibility: 'hidden', width: `${Math.max(1, width)}px`, whiteSpace: 'pre-wrap', lineHeight: '1.2',
    font: fontOf(css, base, basePx), fontSize: `${fontPx(css.fontSize as string | number | undefined, basePx)}px`,
  })
  el.style.lineHeight = '1.2'
  el.textContent = text
  document.body.appendChild(el)
  const h = el.getBoundingClientRect().height
  el.remove()
  return Math.ceil(h) + 1
}
