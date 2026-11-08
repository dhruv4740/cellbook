// Renders a ribbon tab from data: groups of big buttons and stacked small buttons, laid out like Excel's own
// (a big button, or a column of up to three small ones, in the order they appear in `items`).
import { ReactNode } from 'react'
import { Group, GroupSep, BigButton, SmallButton, SplitButton, DialogLauncher } from './controls'
import { MenuItem } from '../../components/Menu'
import { notAvailable } from '../ui'

export interface GBtn {
  label: string
  icon: ReactNode
  size?: 'big' | 'small'
  desc?: string
  shortcut?: string
  menu?: MenuItem[]
  onClick?: () => void
  active?: boolean
  disabled?: boolean
  wide?: boolean
  /** Fixed width in px for a big button (Excel sizes them to their label). */
  width?: number
  /** Big buttons: draw the icon only (the label still names the tooltip), like Excel's Draw tab pens. */
  noLabel?: boolean
  /** Small buttons only: show just the icon (Excel does this for Cut/Copy, AutoSum, Flash Fill...). */
  iconOnly?: boolean
  /** With both menu and onClick, the button is a split button unless this is set (menu only). */
  menuOnly?: boolean
  /** Excel KeyTip access-key sequence for this command, e.g. 'FN', shown when KeyTips are active. */
  keytip?: string
  /** A custom control (checkbox, combo box, spinner) instead of a button. */
  custom?: ReactNode
}
export interface GGroup {
  label: string
  launcher?: boolean
  onLauncher?: () => void
  rows?: GBtn[][] // small buttons laid out in up to 3 stacked rows
  big?: GBtn[] // large buttons
  /** Buttons in display order: a GBtn is a big button, a GBtn[] is a column of up to three small ones. */
  items?: (GBtn | GBtn[])[]
}

function Big({ b }: { b: GBtn }) {
  if (b.custom) return <>{b.custom}</>
  const tip = { title: b.label, desc: b.desc, shortcut: b.shortcut }
  if (b.menu && b.onClick && !b.menuOnly) return <SplitButton icon={b.icon} label={b.label} menu={b.menu} onClick={b.onClick} active={b.active} tip={tip} keytip={b.keytip} />
  return (
    <BigButton
      icon={b.icon} label={b.noLabel ? undefined : b.label} disabled={b.disabled} menu={b.menu} onClick={b.onClick ?? (() => notAvailable(b.label))} active={b.active}
      wide={b.wide ?? b.label.length > 7} tip={tip} keytip={b.keytip} width={b.width}
    />
  )
}

function Small({ b }: { b: GBtn }) {
  if (b.custom) return <>{b.custom}</>
  return (
    <SmallButton
      icon={b.icon} label={b.iconOnly ? undefined : b.label} disabled={b.disabled} menu={b.menu} split={!!(b.menu && b.onClick && !b.menuOnly)}
      onClick={b.onClick ?? (() => notAvailable(b.label))} active={b.active} tip={{ title: b.label, desc: b.desc, shortcut: b.shortcut }} keytip={b.keytip}
    />
  )
}

export function GenericTab({ groups }: { groups: GGroup[] }) {
  return (
    <>
      {groups.map((g, gi) => {
        const first = g.items?.[0]
        const firstBtn = Array.isArray(first) ? first[0] : first
        return (
          <div key={gi} style={{ display: 'contents' }}>
            <Group label={g.label} icon={firstBtn?.icon ?? g.big?.[0]?.icon ?? g.rows?.[0]?.[0]?.icon}>
              {g.items?.map((it, i) =>
                Array.isArray(it) ? (
                  <div className="xl-col xl-stack" key={i}>
                    {it.map((b, bi) => <div className="xl-row" key={bi}><Small b={b} /></div>)}
                  </div>
                ) : (
                  <Big key={i} b={it} />
                )
              )}
              {g.big?.map((b, i) => <Big key={i} b={b} />)}
              {g.rows && (
                <div className="xl-col xl-stack">
                  {g.rows.map((row, ri) => (
                    <div className="xl-row" key={ri}>
                      {row.map((b, bi) => <Small key={bi} b={b} />)}
                    </div>
                  ))}
                </div>
              )}
              {g.launcher && <DialogLauncher tip={{ title: g.label }} onClick={g.onLauncher ?? (() => notAvailable(g.label))} />}
            </Group>
            <GroupSep />
          </div>
        )
      })}
    </>
  )
}
