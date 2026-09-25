// Renders a static ribbon tab from data: every group/button visible and clickable,
// but with no store wiring beyond an optional onClick. Used for tabs whose commands
// the current store doesn't support (Insert, Page Layout galleries, Review, Automate, Help, ...).
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
  /** With both menu and onClick, the button is a split button unless this is set (menu only). */
  menuOnly?: boolean
  /** Excel KeyTip access-key sequence for this command, e.g. 'FN', shown when KeyTips are active. */
  keytip?: string
}
export interface GGroup {
  label: string
  launcher?: boolean
  onLauncher?: () => void
  rows?: GBtn[][] // small buttons laid out in up to 3 stacked rows
  big?: GBtn[] // large buttons
}

export function GenericTab({ groups }: { groups: GGroup[] }) {
  return (
    <>
      {groups.map((g, gi) => (
        <div key={gi} style={{ display: 'contents' }}>
          <Group label={g.label} icon={g.big?.[0]?.icon ?? g.rows?.[0]?.[0]?.icon}>
            {g.big?.map((b, i) => b.menu && b.onClick && !b.menuOnly ? (
              <SplitButton key={i} icon={b.icon} label={b.label} menu={b.menu} onClick={b.onClick} active={b.active} tip={{ title: b.label, desc: b.desc, shortcut: b.shortcut }} keytip={b.keytip} />
            ) : (
              <BigButton key={i} icon={b.icon} label={b.label} disabled={b.disabled} menu={b.menu} onClick={b.onClick ?? (() => notAvailable(b.label))} active={b.active} wide={b.wide ?? b.label.length > 7} tip={{ title: b.label, desc: b.desc, shortcut: b.shortcut }} keytip={b.keytip} />
            ))}
            {g.rows && (
              <div className="xl-col xl-stack">
                {g.rows.map((row, ri) => (
                  <div className="xl-row" key={ri}>
                    {row.map((b, bi) => (
                      <SmallButton key={bi} icon={b.icon} label={b.label} disabled={b.disabled} menu={b.menu} split={!!(b.menu && b.onClick && !b.menuOnly)} onClick={b.onClick ?? (() => notAvailable(b.label))} active={b.active} tip={{ title: b.label, desc: b.desc, shortcut: b.shortcut }} keytip={b.keytip} />
                    ))}
                  </div>
                ))}
              </div>
            )}
            {g.launcher && <DialogLauncher tip={{ title: g.label }} onClick={g.onLauncher ?? (() => notAvailable(g.label))} />}
          </Group>
          <GroupSep />
        </div>
      ))}
    </>
  )
}
