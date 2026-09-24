// The ribbon: tab strip + tab content, collapsible, with tooltips throughout.
import { useEffect, useRef, useState } from 'react'
import { useUi, setUi } from './ui'
import HomeTab from './ribbon/HomeTab'
import { InsertTab, PageLayoutTab, FormulasTab, DataTab, ReviewTab, ViewTab, AutomateTab, HelpTab } from './ribbon/OtherTabs'
import { ChevronUp20Regular, ChevronDown20Regular, CommentAdd20Regular, Share20Regular } from '@fluentui/react-icons'
import './Ribbon.css'

const TABS = ['File', 'Home', 'Insert', 'Page Layout', 'Formulas', 'Data', 'Review', 'View', 'Automate', 'Help']

export default function Ribbon() {
  const ui = useUi()
  const [displayMenu, setDisplayMenu] = useState(false)
  const lastClick = useRef(0)

  useEffect(() => {
    const key = (e: KeyboardEvent) => {
      if (e.ctrlKey && e.key === 'F1') { e.preventDefault(); setUi({ ribbonCollapsed: !ui.ribbonCollapsed }) }
    }
    window.addEventListener('keydown', key)
    return () => window.removeEventListener('keydown', key)
  }, [ui.ribbonCollapsed])

  const selectTab = (t: string) => {
    if (t === 'File') { setUi({ backstage: true }); return }
    const now = Date.now()
    if (ui.ribbonTab === t && now - lastClick.current < 400) setUi({ ribbonCollapsed: !ui.ribbonCollapsed })
    lastClick.current = now
    setUi({ ribbonTab: t, ribbonCollapsed: t === ui.ribbonTab ? ui.ribbonCollapsed : false })
  }

  return (
    <div className="xl-ribbon">
      <div className="xl-tabstrip">
        <div className="xl-tabstrip-left">
          {TABS.map((t) => (
            <div
              key={t}
              className={'xl-tab' + (t === 'File' ? ' file-tab' : '') + (ui.ribbonTab === t && !ui.backstage ? ' active' : '')}
              onClick={() => selectTab(t)}
            >
              {t}
            </div>
          ))}
        </div>
        <div className="xl-tabstrip-right">
          <button className="xl-tab-action" title="Comments"><CommentAdd20Regular /><span>Comments</span></button>
          <button className="xl-tab-action share" title="Share"><Share20Regular /><span>Share</span></button>
        </div>
      </div>
      {!ui.ribbonCollapsed && (
        <div className="xl-ribbon-body">
          <div className="xl-ribbon-groups">
            {ui.ribbonTab === 'Home' && <HomeTab />}
            {ui.ribbonTab === 'Insert' && <InsertTab />}
            {ui.ribbonTab === 'Page Layout' && <PageLayoutTab />}
            {ui.ribbonTab === 'Formulas' && <FormulasTab />}
            {ui.ribbonTab === 'Data' && <DataTab />}
            {ui.ribbonTab === 'Review' && <ReviewTab />}
            {ui.ribbonTab === 'View' && <ViewTab />}
            {ui.ribbonTab === 'Automate' && <AutomateTab />}
            {ui.ribbonTab === 'Help' && <HelpTab />}
          </div>
          <div className="xl-ribbon-collapse">
            <button
              className="xl-ribbon-display"
              title="Ribbon Display Options"
              onClick={() => setDisplayMenu((v) => !v)}
            >
              {ui.ribbonCollapsed ? <ChevronDown20Regular /> : <ChevronUp20Regular />}
            </button>
          </div>
        </div>
      )}
      {displayMenu && (
        <div className="xl-menu-layer" onMouseDown={() => setDisplayMenu(false)}>
          <div className="xl-menu" style={{ position: 'fixed', right: 8, top: 62 }} onMouseDown={(e) => e.stopPropagation()}>
            <div className="xl-menu-item" onClick={() => { setUi({ ribbonCollapsed: false }); setDisplayMenu(false) }}><span className="xl-menu-label">Show Tabs and Commands</span></div>
            <div className="xl-menu-item" onClick={() => { setUi({ ribbonCollapsed: true }); setDisplayMenu(false) }}><span className="xl-menu-label">Show Tabs</span></div>
          </div>
        </div>
      )}
    </div>
  )
}
