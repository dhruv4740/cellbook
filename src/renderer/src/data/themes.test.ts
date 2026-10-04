import { describe, expect, it } from 'vitest'
import { applyThemeToXml, COLOR_SCHEMES, FONT_SCHEMES } from './themes'
import { store } from '../model/store'
import { normalFont } from '../model/types'

const XML = '<a:theme><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>'

describe('themes', () => {
  it('rewrites colour and font scheme in theme1.xml', () => {
    const out = applyThemeToXml(XML, { colors: COLOR_SCHEMES[1], fonts: FONT_SCHEMES[1] })
    expect(out).toContain('<a:accent1><a:srgbClr val="4F81BD"/></a:accent1>')
    expect(out).toContain('<a:dk1><a:srgbClr val="000000"/></a:dk1>')
    expect(out).toContain('<a:clrScheme name="Office 2007 - 2010">')
    expect(out).toContain('<a:majorFont><a:latin typeface="Cambria"/>')
    expect(out).toContain('<a:minorFont><a:latin typeface="Calibri"/>')
  })
  it('a font scheme becomes the Normal font and is exported', () => {
    store.setTheme({ fonts: FONT_SCHEMES[2] })
    expect(normalFont.name).toBe('Arial')
    expect(store.toData().defaultFont?.name).toBe('Arial')
    expect(store.toData().theme?.fonts?.name).toBe('Arial')
    store.setTheme({ fonts: FONT_SCHEMES[0] })
    expect(store.toData().defaultFont).toBeUndefined()
  })
})
