// The ribbon: tab strip + tab content, collapsible, with tooltips throughout.
import { useEffect, useRef, useState } from 'react'
import { useUi, setUi } from './ui'
import HomeTab from './ribbon/HomeTab'
import { InsertTab, PageLayoutTab, FormulasTab, DataTab, ReviewTab, ViewTab, DrawTab, HelpTab } from './ribbon/OtherTabs'
import { RibbonScaleProvider } from './ribbon/controls'
import { ChartDesignTab, ChartFormatTab } from '../charts/ChartRibbonTabs'
import { ChevronUp20Regular, ChevronDown20Regular, CommentAdd20Regular, Share20Regular } from '@fluentui/react-icons'
import KeyTips from './KeyTips'
import './Ribbon.css'

const TABS = ['File', 'Home', 'Insert', 'Draw', 'Page Layout', 'Formulas', 'Data', 'Review', 'View', 'Help']
// Contextual tabs (Chart Tools), shown only while a chart is selected.
const CHART_TABS = ['Chart Design', 'Format']
/** Excel 2021's real tab KeyTips (Alt/F10, then the letter(s)). */
export const TAB_KEYTIPS: Record<string, string> = {
  File: 'F', Home: 'H', Insert: 'N', Draw: 'JI', 'Page Layout': 'P', Formulas: 'M',
  Data: 'A', Review: 'R', View: 'W', Help: 'Y',
  // Contextual tabs, shown while a chart is selected.
  'Chart Design': 'JC', Format: 'JA'
}

export default function Ribbon() {
  const ui = useUi()
  const [displayMenu, setDisplayMenu] = useState(false)
  const lastClick = useRef(0)
  const groupsRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const key = (e: KeyboardEvent) => {
      if (e.ctrlKey && e.key === 'F1') { e.preventDefault(); setUi({ ribbonCollapsed: !ui.ribbonCollapsed }) }
    }
    window.addEventListener('keydown', key)
    return () => window.removeEventListener('keydown', key)
  }, [ui.ribbonCollapsed])

  // Fall back to Home once the chart is deselected and a chart-only tab was active.
  useEffect(() => {
    if (!ui.selectedChart && CHART_TABS.includes(ui.ribbonTab)) setUi({ ribbonTab: 'Home' })
  }, [ui.selectedChart, ui.ribbonTab])

  const selectTab = (t: string) => {
    if (t === 'File') { setUi({ backstage: true }); return }
    const now = Date.now()
    if (ui.ribbonTab === t && now - lastClick.current < 400) setUi({ ribbonCollapsed: !ui.ribbonCollapsed })
    lastClick.current = now
    setUi({ ribbonTab: t, ribbonCollapsed: t === ui.ribbonTab ? ui.ribbonCollapsed : false })
  }

  return (
    <div className="xl-ribbon">
      <KeyTips />
      <div className="xl-tabstrip">
        <div className="xl-tabstrip-left">
          {TABS.map((t) => (
            <div
              key={t}
              className={'xl-tab' + (t === 'File' ? ' file-tab' : '') + (ui.ribbonTab === t && !ui.backstage ? ' active' : '')}
              onClick={() => selectTab(t)}
              style={{ position: 'relative' }}
            >
              {t}
              <span className="xl-keytip-badge xl-keytip-tab" data-keytip={TAB_KEYTIPS[t]}>{TAB_KEYTIPS[t]}</span>
            </div>
          ))}
          {ui.selectedChart && CHART_TABS.map((t) => (
            <div
              key={t}
              className={'xl-tab' + (ui.ribbonTab === t && !ui.backstage ? ' active' : '')}
              style={{ color: '#217346', fontWeight: 600, position: 'relative' }}
              onClick={() => selectTab(t)}
            >
              {t}
              <span className="xl-keytip-badge xl-keytip-tab" data-keytip={TAB_KEYTIPS[t]}>{TAB_KEYTIPS[t]}</span>
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
          <div className="xl-ribbon-groups" ref={groupsRef}>
            <RibbonScaleProvider key={ui.ribbonTab} containerRef={groupsRef}>
              {ui.ribbonTab === 'Home' && <HomeTab />}
              {ui.ribbonTab === 'Insert' && <InsertTab />}
              {ui.ribbonTab === 'Draw' && <DrawTab />}
              {ui.ribbonTab === 'Page Layout' && <PageLayoutTab />}
              {ui.ribbonTab === 'Formulas' && <FormulasTab />}
              {ui.ribbonTab === 'Data' && <DataTab />}
              {ui.ribbonTab === 'Review' && <ReviewTab />}
              {ui.ribbonTab === 'View' && <ViewTab />}
              {ui.ribbonTab === 'Help' && <HelpTab />}
              {ui.ribbonTab === 'Chart Design' && <ChartDesignTab />}
              {ui.ribbonTab === 'Format' && ui.selectedChart && <ChartFormatTab />}
            </RibbonScaleProvider>
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
        <div className="xl-menu-layer" style={{ pointerEvents: 'auto' }} onMouseDown={() => setDisplayMenu(false)}>
          <div className="xl-menu" style={{ position: 'fixed', right: 8, top: 62 }} onMouseDown={(e) => e.stopPropagation()}>
            <div className="xl-menu-item" onClick={() => { setUi({ ribbonCollapsed: false }); setDisplayMenu(false) }}><span className="xl-menu-label">Show Tabs and Commands</span></div>
            <div className="xl-menu-item" onClick={() => { setUi({ ribbonCollapsed: true }); setDisplayMenu(false) }}><span className="xl-menu-label">Show Tabs</span></div>
          </div>
        </div>
      )}
    </div>
  )
}
