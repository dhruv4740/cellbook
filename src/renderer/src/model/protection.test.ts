import { describe, expect, it } from 'vitest'
import { checkPassword, defaultAllow, hashPassword, legacyHash, parseSheetProtection, sheetProtectionXml } from './protection'

describe('protection', () => {
  it('verifies a hash written by real Excel (password "abc")', () => {
    const h = { hashValue: 'BpGP+PwU5OFO39q/xMHSUT+vmyfyXiylD6TyZTbrVvABFqlxRU1viqDtEWpcCBj/V1+UdD/kJKpz+yZWtf1dKQ==', saltValue: '/USW8CwA2sAHLOw/DpDWDg==', spinCount: 100000 }
    expect(checkPassword(h, 'abc')).toBe(true)
    expect(checkPassword(h, 'abd')).toBe(false)
  })
  it('hashes and verifies its own passwords', () => {
    const h = hashPassword('Pässword1')
    expect(checkPassword(h, 'Pässword1')).toBe(true)
    expect(checkPassword(h, 'pässword1')).toBe(false)
    expect(hashPassword('')).toBeUndefined()
  })
  it('legacy 16-bit hash', () => {
    expect(legacyHash('password')).toBe('83AF')
    expect(checkPassword({ legacy: '83AF' }, 'password')).toBe(true)
    expect(checkPassword({ legacy: '83AF' }, 'Password')).toBe(false)
  })
  it('round-trips the options through the sheetProtection attributes', () => {
    const p = { allow: { ...defaultAllow(), formatCells: true, sort: true, insertRows: true, selectLockedCells: false } }
    const xml = sheetProtectionXml(p)
    expect(xml).toBe('<sheetProtection sheet="1" objects="1" scenarios="1" selectLockedCells="1" formatCells="0" insertRows="0" sort="0"/>')
    const attrs = Object.fromEntries([...xml.matchAll(/(\w+)="([^"]*)"/g)].map((m) => [m[1], m[2]]))
    expect(parseSheetProtection(attrs)).toEqual({ allow: p.allow })
  })
})
