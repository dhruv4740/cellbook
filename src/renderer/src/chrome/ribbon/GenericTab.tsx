// Renders a static ribbon tab from data: every group/button visible and clickable,
// but with no store wiring beyond an optional onClick. Used for tabs whose commands
// the current store doesn't support (Insert, Page Layout galleries, Review, Automate, Help, ...).
import { ReactNode } from 'react'
import { Group, GroupSep, BigButton, SmallButton, DialogLauncher } from './controls'
import { MenuItem } from '../../components/Menu'

export interface GBtn {
  label: string
  icon: ReactNode
  size?: 'big' | 'small'
  desc?: string
  shortcut?: string
  menu?: MenuItem[]
  onClick?: () => void
  active?: boolean
  wide?: boolean
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
          <Group label={g.label}>
            {g.big?.map((b, i) => (
              <BigButton key={i} icon={b.icon} label={b.label} menu={b.menu} onClick={b.onClick ?? (() => {})} active={b.active} wide={b.wide ?? b.label.length > 7} tip={{ title: b.label, desc: b.desc, shortcut: b.shortcut }} />
            ))}
            {g.rows && (
              <div className="xl-col xl-stack">
                {g.rows.map((row, ri) => (
                  <div className="xl-row" key={ri}>
                    {row.map((b, bi) => (
                      <SmallButton key={bi} icon={b.icon} label={b.label} menu={b.menu} onClick={b.onClick ?? (() => {})} active={b.active} tip={{ title: b.label, desc: b.desc, shortcut: b.shortcut }} />
                    ))}
                  </div>
                ))}
              </div>
            )}
            {g.launcher && <DialogLauncher tip={{ title: g.label }} onClick={g.onLauncher ?? (() => {})} />}
          </Group>
          <GroupSep />
        </div>
      ))}
    </>
  )
}
