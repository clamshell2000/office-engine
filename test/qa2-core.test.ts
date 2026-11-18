/**
 * Second round of the QA pass of 18/11/2026 (findings rated medium and low; ids as in its report): recalculation,
 * what is typed, sorting, formats of whole columns, collaboration.
 */
import { describe, expect, test } from 'bun:test'
import { at, book, openBook, rows, type Opened } from './xlsx-builder'

const recalculated = async (o: Opened) => {
  while (!o.events.some((e) => e.event === 'xlsx:progress' && e.data.recalc)) await new Promise((r) => setTimeout(r, 5))
}
const num = async (o: Opened, sheet: number, ref: string) => Number(((await o.engine.call('xlsx.cell', [sheet, at(ref).r, at(ref).c])) as any)?.value)

describe('P1-013: a sheet where every row reads the whole table', () => {
  const N = 6000
  test('running totals, COUNTIF and SUMIFS down a column follow one edit, with the right values', async () => {
    const grid: any[][] = []
    for (let r = 1; r <= N; r++) {
      grid.push([r % 50, r % 7, r, { f: `SUM($C$1:C${r})`, v: 0 }, { f: `COUNTIF(A:A,A${r})`, v: 0 }, { f: `SUMIFS(C:C,A:A,A${r},B:B,B${r})`, v: 0 }, { f: `VLOOKUP(${(r % 40) + 1},$A$1:$C$60,3,FALSE)`, v: 0 }])
    }
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows(grid) }] }))
    await o.set(0, 'Z1', '1')
    const t0 = performance.now()
    await o.set(0, 'C1', '1000')
    await o.set(0, 'A2', '49')
    const ms = performance.now() - t0
    // by hand
    const A = grid.map((g) => g[0] as number)
    const B = grid.map((g) => g[1] as number)
    const C = grid.map((g) => g[2] as number)
    C[0] = 1000
    A[1] = 49
    const sumifs = (r: number) => C.reduce((s, c, i) => (A[i] === A[r] && B[i] === B[r] ? s + c : s), 0)
    for (const r of [0, 1, 2, 48, 49, 349, N - 1]) {
      expect(await num(o, 0, `D${r + 1}`)).toBe(C.slice(0, r + 1).reduce((x, y) => x + y, 0))
      expect(await num(o, 0, `E${r + 1}`)).toBe(A.filter((a) => a === A[r]).length)
      expect(await num(o, 0, `F${r + 1}`)).toBe(sumifs(r))
      expect(await num(o, 0, `G${r + 1}`)).toBe(C[A.slice(0, 60).indexOf(((r + 1) % 40) + 1)])
    }
    // the cells were read twenty million times per column: seconds, not minutes (the bound is for a slow machine)
    expect(ms).toBeLessThan(60_000)
  }, 120_000)

  test('a block kept while computing is read again once a formula in it has its value', async () => {
    // B reads A whole, twice over (kept from the second time); A2 is a formula still to compute when B1 reads it
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[1, { f: 'SUM($A$1:$A$3)', v: 0 }], [{ f: 'A1*10', v: 0 }, { f: 'SUM($A$1:$A$3)', v: 0 }], [3, { f: 'SUM($A$1:$A$3)+C3', v: 0 }, { f: 'A2+1', v: 0 }]]) }] }))
    await o.set(0, 'A1', '2')
    expect([await o.shown(0, 0, 1), await o.shown(0, 1, 1), await o.shown(0, 2, 1)]).toEqual(['25', '25', '46'])
    await o.set(0, 'A3', '=B1')
    // a circular reference through a kept block: an error, the same however often it is computed
    const first = [await o.shown(0, 0, 1), await o.shown(0, 1, 1), await o.shown(0, 2, 0)]
    await o.set(0, 'A1', '2')
    expect([await o.shown(0, 0, 1), await o.shown(0, 1, 1), await o.shown(0, 2, 0)]).toEqual(first)
    expect(first).toContain('#REF!')
    await o.set(0, 'A3', '4')
    expect([await o.shown(0, 0, 1), await o.shown(0, 1, 1), await o.shown(0, 2, 1)]).toEqual(['26', '26', '47'])
  })
})

describe('P1-025: TODAY, NOW and RAND are of the moment the file is read', () => {
  test('what the file holds for them, and for what reads them, is computed again; the file is not changed by it', async () => {
    const src = await book({
      sheets: [
        { name: 'Due', rows: rows([[{ f: 'TODAY()', v: 45000 }, 44990, { f: 'A1-B1', v: 10 }, { f: 'NOW()', v: 45000.5 }, { f: 'RAND()', v: 0.123 }, { f: 'B1*2', v: 1 }]]) },
        { name: 'Other', rows: rows([[{ f: 'Due!C1+1', v: 11 }]]) },
      ],
    })
    const o = await openBook(src)
    await recalculated(o)
    const today = await num(o, 0, 'A1')
    expect(today).toBeGreaterThan(46000)
    expect(Number.isInteger(today)).toBe(true)
    expect(await num(o, 0, 'C1')).toBe(today - 44990)
    expect(await num(o, 0, 'D1')).toBeGreaterThanOrEqual(today)
    expect(await num(o, 0, 'E1')).not.toBe(0.123)
    // what does not read the clock keeps the value of the file, right or not: nothing asked to compute it
    expect(await num(o, 0, 'F1')).toBe(1)
    await o.load(1)
    expect(await num(o, 1, 'A1')).toBe(today - 44990 + 1)
    expect(await o.engine.call('dirty', [])).toBe(false)
  })
})
