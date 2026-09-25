// Excel's KeyTips: tap Alt (or F10) to badge every ribbon tab + QAT command with its access-key
// letter(s); typing them switches tabs / runs commands, exactly like real Excel's "Alt, H, B, A"
// style sequences. Esc steps back one level; Alt/Esc/click outside exits KeyTips entirely.
//
// Implementation: every command that wants a KeyTip renders a small hidden badge
// (<span data-keytip="X">) right on itself (see controls.tsx: KeyTipBadge). This component only
// toggles a body-level CSS class that reveals those badges, and listens for the matching keys.
import { useEffect, useRef, useState } from 'react'
import { getUi, setUi } from './ui'
import * as cmds from '../grid/commands'

type Level = 'off' | 'tabs' | 'commands' | 'menu'

// Excel's "access keys for previous versions" compatibility mode: Alt then a legacy Excel-2003
// menu letter (none of which collide with a modern ribbon tab's own KeyTip letter) silently
// executes the equivalent ribbon command with no visible menu, exactly like real Excel.
const LEGACY_SEQUENCES: Record<string, () => void> = {
  ES: () => setUi({ dialog: { kind: 'pasteSpecial' } }), // Edit > Paste Special
  ED: () => cmds.deleteCellsDialog(), // Edit > Delete
  OCA: () => cmds.autoFitColumns(), // Format > Column > AutoFit Selection
  IR: () => cmds.insertSheetRows(), // Insert > Rows
  IC: () => cmds.insertSheetColumns() // Insert > Columns
}

/** Pure prefix-matching used to resolve a typed buffer against a set of KeyTip codes (ribbon
 *  badges, an open menu's access keys, or the legacy Excel-2003 sequence table): case-insensitive,
 *  returns the single exact code once the buffer stops being an ambiguous prefix of more than one.
 *  Exported (and covered by KeyTips.test.ts) since the rest of the engine needs a live DOM. */
export type KeyTipMatch =
  | { kind: 'none' } // buffer doesn't prefix any code: ignore the keystroke
  | { kind: 'prefix' } // buffer is a valid prefix of 1+ codes: keep buffering
  | { kind: 'exact'; code: string } // buffer exactly matches exactly one code: run it
export function matchKeyTipBuffer(codes: string[], buf: string): KeyTipMatch {
  const b = buf.toUpperCase()
  const matches = codes.filter((c) => c.toUpperCase().startsWith(b))
  if (!matches.length) return { kind: 'none' }
  const exact = matches.find((c) => c.toUpperCase() === b)
  if (exact && matches.length === 1) return { kind: 'exact', code: exact }
  return { kind: 'prefix' }
}

/** True while the KeyTips overlay is showing (exported so Grid/menus can suppress their own
 *  key handling while it's active, if they choose to check it). */
let active: Level = 'off'
const listeners = new Set<() => void>()
function setActive(v: Level): void {
  active = v
  document.body.classList.toggle('xl-keytips-active', v !== 'off')
  document.body.classList.toggle('xl-keytips-tabs', v === 'tabs')
  document.body.classList.toggle('xl-keytips-commands', v === 'commands')
  listeners.forEach((l) => l())
}
export function keyTipsActive(): boolean {
  return active !== 'off'
}

function visibleBadges(scope: 'tabs' | 'commands'): HTMLElement[] {
  const root = scope === 'tabs' ? document.querySelector('.xl-tabstrip, .xl-qat')?.parentElement ?? document : document.querySelector('.xl-ribbon-body') ?? document
  const sel = scope === 'tabs' ? '.xl-keytip-tab, .xl-qat .xl-keytip-badge' : '.xl-ribbon-body [data-keytip]'
  return Array.from((scope === 'tabs' ? document : root).querySelectorAll<HTMLElement>(sel)).filter((el) => {
    const code = el.getAttribute('data-keytip')
    if (!code) return false
    // Skip the RibbonScaleProvider's hidden off-screen measurement clone of each group (it
    // renders the same buttons/badges again to measure their natural width) and anything else
    // marked aria-hidden.
    if (el.closest('.xl-group-measure, [aria-hidden="true"]')) return false
    // Skip buttons that are disabled.
    const btn = el.closest('button') as HTMLButtonElement | null
    return !btn || !btn.disabled
  })
}

export default function KeyTips() {
  const [, bump] = useState(0)
  const buffer = useRef('')
  const altDown = useRef(false)
  const altMoved = useRef(false)

  useEffect(() => {
    const l = () => bump((n) => n + 1)
    listeners.add(l)
    return () => { listeners.delete(l) }
  }, [])

  useEffect(() => {
    function exit(): void {
      buffer.current = ''
      setActive('off')
    }
    function enterTabs(): void {
      buffer.current = ''
      setActive('tabs')
    }
    function enterMenu(): void {
      buffer.current = ''
      setActive('menu')
    }
    function stepBack(): void {
      if (active === 'menu' || active === 'commands') enterTabs()
      else exit()
    }

    /** Excel opens the dropdown/gallery instead of running the default action for a command
     *  that carries one (Paste, Insert, Delete, Format, Borders, Sort & Filter, ...), then reads
     *  the next keystroke against that open menu's own access-key letters (see tryMatchMenu). A
     *  plain leaf command (Bold, Copy, ...) just runs and KeyTips exits, matching real Excel. */
    function clickCommand(exact: HTMLElement): void {
      const btn = (exact.closest('button') as HTMLElement) || exact
      const splitWrap = btn.closest('.xl-split, .xl-stacksplit') as HTMLElement | null
      const splitCaret = splitWrap?.querySelector('[class*="caret"]') as HTMLElement | null
      const inlineCaret = btn.querySelector('.xl-drop-caret')
      if (splitCaret) { splitCaret.click(); enterMenu(); return }
      if (inlineCaret) { btn.click(); enterMenu(); return }
      btn.click()
      exit()
    }

    function tryMatch(scope: 'tabs' | 'commands'): void {
      const badges = visibleBadges(scope)
      const codes = badges.map((b) => b.getAttribute('data-keytip') || '')
      const buf = buffer.current
      const result = matchKeyTipBuffer(codes, buf)
      if (result.kind === 'none') {
        if (scope === 'tabs') {
          const legacy = matchKeyTipBuffer(Object.keys(LEGACY_SEQUENCES), buf)
          if (legacy.kind === 'exact') { LEGACY_SEQUENCES[legacy.code](); exit(); return }
          if (legacy.kind === 'prefix') return // valid legacy prefix: keep buffering, no visible menu (matches real Excel)
        }
        buffer.current = ''
        return // no such sequence: ignore keystroke, keep level
      }
      if (result.kind === 'exact') {
        buffer.current = ''
        const exact = badges.find((b) => (b.getAttribute('data-keytip') || '').toUpperCase() === result.code.toUpperCase())!
        if (scope === 'tabs') {
          // Tab strip element itself carries the click handler (badge is a child span).
          const tabEl = (exact.closest('.xl-tab') as HTMLElement) || exact
          tabEl.click()
          if (getUi().backstage) { exit(); return } // File tab: backstage has its own UI, not KeyTips
          setUi({ ribbonCollapsed: false })
          setActive('commands')
        } else {
          clickCommand(exact)
        }
      }
      // else: still a valid prefix of ≥1 code — wait for more letters.
    }

    /** Matches a typed letter against an *open* dropdown/context menu's own access-key
     *  underlines (Menu.tsx's `&` markup), e.g. Alt,H,O,I -> Format menu's "AutoFit Column
     *  &Width". Mirrors Excel forwarding KeyTip keystrokes into an open menu. */
    function tryMatchMenu(): void {
      if (!buffer.current) return
      const items = Array.from(document.querySelectorAll<HTMLElement>('.xl-menu-layer .xl-menu-item'))
        .filter((el) => !el.classList.contains('disabled'))
        .map((el) => ({ el, key: el.querySelector('.xl-menu-label u')?.textContent || '' }))
        .filter((x) => x.key)
      const result = matchKeyTipBuffer(items.map((x) => x.key), buffer.current)
      if (result.kind === 'none') { buffer.current = ''; return }
      if (result.kind === 'exact') {
        buffer.current = ''
        items.find((x) => x.key.toUpperCase() === result.code.toUpperCase())!.el.click()
        exit()
      }
    }

    function onKeyDown(e: KeyboardEvent): void {
      if (e.key === 'Alt') { altDown.current = true; altMoved.current = false; return }
      if (e.key === 'F10' && !e.shiftKey && !e.ctrlKey) {
        e.preventDefault()
        active === 'off' ? enterTabs() : exit()
        return
      }
      // Any other key while Alt is held (Alt+F1, Alt+=, Alt+Enter...) means this was a chord,
      // so releasing Alt must not toggle KeyTips.
      altMoved.current = true
      if (active === 'off') return
      if (e.key === 'Escape') { e.preventDefault(); stepBack(); return }
      if (e.ctrlKey || e.metaKey) { exit(); return }
      if (e.key.length !== 1 || !/[A-Za-z0-9]/.test(e.key)) return
      e.preventDefault()
      buffer.current += e.key.toUpperCase()
      if (active === 'menu') tryMatchMenu()
      else tryMatch(active === 'tabs' ? 'tabs' : 'commands')
      bump((n) => n + 1)
    }
    function onKeyUp(e: KeyboardEvent): void {
      if (e.key !== 'Alt') return
      const wasDown = altDown.current
      altDown.current = false
      if (wasDown && !altMoved.current) {
        // A bare tap of Alt: toggle KeyTips at the tab level.
        active === 'off' ? enterTabs() : exit()
      }
      altMoved.current = false
    }
    function onMouseDown(e: MouseEvent): void {
      if (active === 'off') return
      // Clicking a badge's own button is handled by tryMatch; any other click exits KeyTips.
      if (!(e.target as HTMLElement).closest('[data-keytip]')) exit()
    }
    window.addEventListener('keydown', onKeyDown, true)
    window.addEventListener('keyup', onKeyUp, true)
    window.addEventListener('mousedown', onMouseDown, true)
    return () => {
      window.removeEventListener('keydown', onKeyDown, true)
      window.removeEventListener('keyup', onKeyUp, true)
      window.removeEventListener('mousedown', onMouseDown, true)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  return null
}
