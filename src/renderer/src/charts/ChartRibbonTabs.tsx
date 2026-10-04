// Contextual "Chart Design" and "Format" ribbon tabs, shown only while a chart is selected.
import { useStore } from '../model/store'
import { useUi, setUi } from '../chrome/ui'
import { GenericTab, GGroup } from '../chrome/ribbon/GenericTab'
import { MenuItem } from '../components/Menu'
import { chartsOf, updateChart, setChartRange } from './chartsStore'
import { COLUMN_BAR_GALLERY, LINE_AREA_GALLERY, PIE_DOUGHNUT_GALLERY, SCATTER_BUBBLE_GALLERY, GalleryEntry } from './gallery'
import {
  AddSquare20Regular, TableSimple20Regular, ColorBackground20Regular, PaintBrush20Regular,
  ArrowSwap20Regular, TableSettings20Regular, ChartMultiple20Regular, ArrowMove20Regular,
  Shapes20Regular, TextEffects20Regular, AccessibilityCheckmark20Regular, TableFreezeRow20Regular, Ruler20Regular
} from '@fluentui/react-icons'

function allChartTypeMenu(onPick: (e: GalleryEntry) => void): MenuItem[] {
  const section = (name: string, entries: GalleryEntry[]): MenuItem => ({
    label: name,
    submenu: entries.map((e) => ({ label: e.label, onClick: () => onPick(e) }))
  })
  return [
    section('Column or Bar', COLUMN_BAR_GALLERY),
    section('Line or Area', LINE_AREA_GALLERY),
    section('Pie or Doughnut', PIE_DOUGHNUT_GALLERY),
    section('Scatter (X, Y) or Bubble', SCATTER_BUBBLE_GALLERY)
  ]
}

export function ChartDesignTab() {
  const store = useStore()
  const ui = useUi()
  const sheet = store.active
  const spec = chartsOf(sheet).find((c) => c.id === ui.selectedChart)
  if (!spec) return <GenericTab groups={[]} />

  const addElementMenu: MenuItem[] = (['axes', 'axisTitles', 'title', 'dataLabels', 'gridlines', 'legend'] as const).map((k) => ({
    label: k === 'axisTitles' ? 'Axis Titles' : k === 'dataLabels' ? 'Data Labels' : k[0].toUpperCase() + k.slice(1),
    checked: spec.elements[k],
    onClick: () => updateChart(store, sheet, spec.id, { elements: { ...spec.elements, [k]: !spec.elements[k] } }, 'Chart Elements')
  }))

  const quickLayoutMenu: MenuItem[] = [1, 2, 3, 4].map((n) => ({ label: 'Layout ' + n, onClick: () => {} }))
  const changeColorsMenu: MenuItem[] = Array.from({ length: 7 }, (_, i) => ({
    label: i === 0 ? 'Colorful Palette' : 'Monochromatic ' + i,
    checked: (spec.colorSetId ?? 0) === i,
    onClick: () => updateChart(store, sheet, spec.id, { colorSetId: i }, 'Change Colors')
  }))
  const chartStylesMenu: MenuItem[] = Array.from({ length: 8 }, (_, i) => ({
    label: 'Style ' + (i + 1),
    checked: (spec.styleId ?? 0) === i,
    onClick: () => updateChart(store, sheet, spec.id, { styleId: i }, 'Chart Style')
  }))

  const groups: GGroup[] = [
    { label: 'Chart Layouts', big: [
      { label: 'Add Chart Element', icon: <AddSquare20Regular fontSize={28} />, menu: addElementMenu },
      { label: 'Quick Layout', icon: <TableSimple20Regular fontSize={28} />, menu: quickLayoutMenu }
    ] },
    { label: 'Chart Styles', big: [
      { label: 'Change Colors', icon: <ColorBackground20Regular fontSize={28} />, menu: changeColorsMenu },
      { label: 'Chart Styles', icon: <PaintBrush20Regular fontSize={28} />, menu: chartStylesMenu }
    ] },
    { label: 'Data', big: [
      { label: 'Switch Row/Column', icon: <ArrowSwap20Regular fontSize={28} />, onClick: () => updateChart(store, sheet, spec.id, { seriesInColumns: !spec.seriesInColumns }, 'Switch Row/Column') },
      { label: 'Select Data', icon: <TableSettings20Regular fontSize={28} />, onClick: () => setChartRange(store, sheet, spec.id, spec.dataRange, spec.seriesInColumns) }
    ] },
    { label: 'Type', big: [
      { label: 'Change Chart Type', icon: <ChartMultiple20Regular fontSize={28} />, menu: [
        ...allChartTypeMenu((e) => updateChart(store, sheet, spec.id, { type: e.type, subtype: e.subtype }, 'Change Chart Type')),
        { label: 'Combo...', onClick: () => setUi({ dialog: { kind: 'comboChart', props: { chartId: spec.id } } }) }
      ] }
    ] },
    { label: 'Location', big: [
      { label: 'Move Chart', icon: <ArrowMove20Regular fontSize={28} /> }
    ] }
  ]
  return <GenericTab groups={groups} />
}

export function ChartFormatTab() {
  const store = useStore()
  const ui = useUi()
  const sheet = store.active
  const spec = chartsOf(sheet).find((c) => c.id === ui.selectedChart)
  if (!spec) return <GenericTab groups={[]} />

  const groups: GGroup[] = [
    { label: 'Current Selection', big: [
      { label: 'Chart Area', icon: <TableFreezeRow20Regular fontSize={28} /> },
      { label: 'Format Selection', icon: <Ruler20Regular fontSize={28} /> }
    ] },
    { label: 'Insert Shapes', big: [{ label: 'Shapes', icon: <Shapes20Regular fontSize={28} /> }] },
    { label: 'Shape Styles', big: [{ label: 'Shape Styles', icon: <PaintBrush20Regular fontSize={28} /> }] },
    { label: 'WordArt Styles', big: [{ label: 'WordArt', icon: <TextEffects20Regular fontSize={28} /> }] },
    { label: 'Accessibility', big: [{ label: 'Check Accessibility', icon: <AccessibilityCheckmark20Regular fontSize={28} /> }] },
    { label: 'Arrange', big: [{ label: 'Arrange', icon: <ArrowMove20Regular fontSize={28} /> }] },
    { label: 'Size', rows: [
      [{ label: 'Height', icon: <Ruler20Regular />, onClick: () => {} }],
      [{ label: 'Width', icon: <Ruler20Regular />, onClick: () => {} }]
    ] }
  ]
  return <GenericTab groups={groups} />
}

// Kept here so setUi's selectedChart flag stays a one-line change everywhere it's needed.
export function clearChartSelection(): void {
  setUi({ selectedChart: null })
}
