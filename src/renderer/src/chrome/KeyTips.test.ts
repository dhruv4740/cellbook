import { describe, expect, it } from 'vitest'
import { TAB_KEYTIPS } from './Ribbon'
import { matchKeyTipBuffer } from './KeyTips'

describe('Ribbon KeyTip sequence table', () => {
  it('matches Excel 2021\'s real tab access-key letters', () => {
    expect(TAB_KEYTIPS).toEqual({
      File: 'F', Home: 'H', Insert: 'N', Draw: 'JI', 'Page Layout': 'P', Formulas: 'M',
      Data: 'A', Review: 'R', View: 'W', Help: 'Y',
      'Chart Design': 'JC', Format: 'JA'
    })
  })

  it('has exactly one code per visible tab, all non-empty', () => {
    const tabs = ['File', 'Home', 'Insert', 'Draw', 'Page Layout', 'Formulas', 'Data', 'Review', 'View', 'Help', 'Chart Design', 'Format']
    for (const t of tabs) {
      expect(TAB_KEYTIPS[t]).toBeTruthy()
    }
    expect(Object.keys(TAB_KEYTIPS).sort()).toEqual([...tabs].sort())
  })

  it('has no code that is a strict prefix of another (so typing is unambiguous once complete)', () => {
    const codes = Object.values(TAB_KEYTIPS)
    for (const a of codes) {
      for (const b of codes) {
        if (a === b) continue
        expect(b.startsWith(a)).toBe(false)
      }
    }
  })
})

describe('matchKeyTipBuffer (the KeyTip sequence resolver)', () => {
  const codes = ['V', 'X', 'C', 'FP', 'FF', 'FS', 'FG', 'FK', 'B', 'AT', 'AM', 'AB']

  it('resolves a single unambiguous letter immediately', () => {
    expect(matchKeyTipBuffer(codes, 'V')).toEqual({ kind: 'exact', code: 'V' })
    expect(matchKeyTipBuffer(codes, 'B')).toEqual({ kind: 'exact', code: 'B' })
  })

  it('waits for more letters while the buffer is an ambiguous prefix ("F" of FP/FF/FS/FG/FK)', () => {
    expect(matchKeyTipBuffer(codes, 'F')).toEqual({ kind: 'prefix' })
  })

  it('resolves the second letter once it disambiguates a multi-letter code', () => {
    expect(matchKeyTipBuffer(codes, 'FP')).toEqual({ kind: 'exact', code: 'FP' })
    expect(matchKeyTipBuffer(codes, 'FS')).toEqual({ kind: 'exact', code: 'FS' })
  })

  it('is case-insensitive (typed letters are lowercase, codes are uppercase)', () => {
    expect(matchKeyTipBuffer(codes, 'fp')).toEqual({ kind: 'exact', code: 'FP' })
    expect(matchKeyTipBuffer(codes, 'a')).toEqual({ kind: 'prefix' }) // AT/AM/AB
  })

  it('reports "none" for a letter that starts no code at all', () => {
    expect(matchKeyTipBuffer(codes, 'Z')).toEqual({ kind: 'none' })
  })

  it("an exact match that is also a prefix of a longer code stays ambiguous (Excel's own rule)", () => {
    // e.g. Home tab has both keytip="AT" (Top Align) with no shorter/longer collision in this
    // suite's fixture, but a code that is simultaneously a full match AND a strict prefix of
    // another code must NOT fire early - construct that case explicitly:
    const withCollision = ['A', 'AT']
    expect(matchKeyTipBuffer(withCollision, 'A')).toEqual({ kind: 'prefix' })
    expect(matchKeyTipBuffer(withCollision, 'AT')).toEqual({ kind: 'exact', code: 'AT' })
  })

  it("Alt,H,O,I resolves through the real Home > Format menu's access keys", () => {
    // Mirrors HomeTab.tsx's formatMenu access keys (Row Height/AutoFit Row/Column Width/
    // AutoFit Column/Default Width/...), Excel 2021's real Alt,H,O,<letter> mnemonics.
    const formatMenu = ['H', 'A', 'W', 'I', 'D', 'U', 'R', 'M', 'T', 'P', 'L', 'E']
    expect(matchKeyTipBuffer(formatMenu, 'I')).toEqual({ kind: 'exact', code: 'I' }) // AutoFit Column Width
    expect(matchKeyTipBuffer(formatMenu, 'A')).toEqual({ kind: 'exact', code: 'A' }) // AutoFit Row Height
    expect(matchKeyTipBuffer(formatMenu, 'R')).toEqual({ kind: 'exact', code: 'R' }) // Rename Sheet
  })
})
