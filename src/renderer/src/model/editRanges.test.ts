import { unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { WorkbookStore } from './store'
import { defaultAllow, inEditRanges, parseSqref, protectedRangesXml } from './protection'
import { readXlsx, writeXlsx } from '../io/xlsx'

describe('Allow Edit Ranges', () => {
  it('parses sqref lists and tests membership', () => {
    expect(parseSqref('A1:B2 D4 junk')).toEqual([{ r1: 0, c1: 0, r2: 1, c2: 1 }, { r1: 3, c1: 3, r2: 3, c2: 3 }])
    const list = [{ title: 'Inputs', ref: 'B2:C3' }]
    expect(inEditRanges(list, 1, 1)).toBe(true)
    expect(inEditRanges(list, 3, 1)).toBe(false)
    expect(inEditRanges(undefined, 0, 0)).toBe(false)
    expect(protectedRangesXml([{ title: 'A & "B"', ref: 'a1:b2' }])).toBe('<protectedRanges><protectedRange sqref="A1:B2" name="A &amp; &quot;B&quot;"/></protectedRanges>')
    expect(protectedRangesXml([{ title: 'x', ref: 'nope' }])).toBe('')
  })

  it('lets a protected sheet be edited inside a range only', () => {
    const s = new WorkbookStore()
    const sh = s.active
    s.updateLayout(sh, 'Allow Edit Ranges', (l) => { l.editRanges = [{ title: 'Inputs', ref: 'B2' }] })
    s.setSheetProtection(sh, { allow: defaultAllow() })
    s.onProtectedEdit = () => {}
    s.setRaw(sh, 1, 1, '5')
    s.setRaw(sh, 0, 0, '6')
    expect(s.getRaw(sh, 1, 1)).toBe('5')
    expect(s.getRaw(sh, 0, 0)).toBe('')
  })

  it('round-trips through xlsx, with protectedRanges between sheetProtection and the next element, also when unprotected', async () => {
    const s = new WorkbookStore()
    const sh = s.active
    s.setRaw(sh, 0, 0, '1')
    s.updateLayout(sh, 'Allow Edit Ranges', (l) => { l.editRanges = [{ title: 'Inputs', ref: 'B2:C3 E5' }, { title: 'Notes', ref: 'A10' }] })
    const unprotected = await readXlsx(await writeXlsx(s.toData()))
    expect(unprotected.sheets[0].layout?.editRanges).toEqual([{ title: 'Inputs', ref: 'B2:C3 E5' }, { title: 'Notes', ref: 'A10' }])
    expect(unprotected.sheets[0].layout?.protection).toBeUndefined()

    s.setSheetProtection(sh, { allow: defaultAllow() })
    const bytes = await writeXlsx(s.toData())
    const xml = new TextDecoder().decode(unzipSync(bytes)['xl/worksheets/sheet1.xml'])
    expect(xml.indexOf('<sheetProtection')).toBeGreaterThan(xml.indexOf('</sheetData>'))
    expect(xml.indexOf('<protectedRanges>')).toBeGreaterThan(xml.indexOf('<sheetProtection'))
    const back = await readXlsx(bytes)
    expect(back.sheets[0].layout?.editRanges).toHaveLength(2)
    expect(back.sheets[0].layout?.protection).toBeDefined()
  })
})
