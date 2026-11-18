import { describe, expect, test } from 'bun:test'
import { parseXml } from '../src/core/xml'
import { parseSheetStyles } from '../src/core/xlsx/styles'

const theme = { colors: {}, major: 'Calibri', minor: 'Calibri', fillStyles: [], bgFillStyles: [] } as never

const styles = (fills: string, borders = '<border><left/><right/><top/><bottom/><diagonal/></border>', xf = '') =>
  parseSheetStyles(
    parseXml(
      `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill>${fills}</fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>${borders}</borders><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="1" borderId="1">${xf}</xf></cellXfs></styleSheet>`,
    ),
    theme,
  ).xfs[1].css

describe('P2-037 cell styles', () => {
  test('a pattern fill is its colors mixed by the pattern, not the solid foreground', () => {
    const css = styles('<fill><patternFill patternType="gray125"><fgColor rgb="FFFF0000"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>')
    // 12.5 % red over white
    expect(css.background).toBe('#ffdfdf')
    expect(styles('<fill><patternFill patternType="darkGray"><fgColor rgb="FF000000"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>').background).toBe('#404040')
    // a solid fill stays as it is; a pattern without a background color mixes with white
    expect(styles('<fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill>').background).toBe('#00ff00')
    expect(styles('<fill><patternFill patternType="lightGray"><fgColor rgb="FF0000FF"/></patternFill></fill>').background).toBe('#bfbfff')
  })

  test('diagonal borders are drawn as lines across the cell', () => {
    const css = styles(
      '<fill><patternFill patternType="none"/></fill>',
      '<border diagonalUp="1" diagonalDown="1"><left/><right/><top/><bottom/><diagonal style="thin"><color rgb="FFFF0000"/></diagonal></border>',
    )
    expect(String(css.backgroundImage)).toContain('to top right')
    expect(String(css.backgroundImage)).toContain('to bottom right')
    expect(String(css.backgroundImage)).toContain('#ff0000')
    const up = styles('<fill><patternFill patternType="none"/></fill>', '<border diagonalUp="1"><left/><right/><top/><bottom/><diagonal style="thin"/></border>')
    // "up" runs from the bottom left to the top right corner: the 50 % line of "to bottom right"
    expect(String(up.backgroundImage)).toContain('to bottom right')
    expect(String(up.backgroundImage)).not.toContain('to top right')
  })

  test('shrink to fit is carried to the grid', () => {
    const css = styles('<fill><patternFill patternType="none"/></fill>', undefined, '<alignment shrinkToFit="1"/>')
    expect(css.shrinkToFit).toBe(1)
  })
})
