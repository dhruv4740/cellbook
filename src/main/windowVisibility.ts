// View > Hide / Unhide window: which windows may be hidden and which are listed by Unhide.
export interface WinVis { id: number; visible: boolean; title: string }

/** The last visible window can't be hidden: nothing would be left to run Unhide from. */
export function canHideWindow(wins: WinVis[], id: number): boolean {
  return wins.some((w) => w.id === id && w.visible) && wins.some((w) => w.id !== id && w.visible)
}

export function hiddenWindows(wins: WinVis[]): { id: number; title: string }[] {
  return wins.filter((w) => !w.visible).map((w) => ({ id: w.id, title: w.title }))
}
