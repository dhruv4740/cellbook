// Notes (yellow sticky, single author note) and Comments (threaded discussion card), Excel 365 style.
// Rendered as small non-modal floating boxes anchored to their cell; all writes go through the store
// so they're undoable.
import { useEffect, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { store, SheetState, DEFAULT_AUTHOR } from '../model/store'
import { a1 } from '../model/types'
import { setUi } from '../chrome/ui'
import './Annotations.css'

function fmtTime(iso: string): string {
  const d = new Date(iso)
  return d.toLocaleString('en-US', { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' })
}

/** Editable yellow note box, positioned just right of the cell it's shown for `rect`. */
export function NoteBox({ sheet, r, c, rect, onClose }: { sheet: SheetState; r: number; c: number; rect: { x: number; y: number; w: number; h: number }; onClose: () => void }): React.ReactElement {
  const existing = sheet.layout.notes[r + ',' + c]
  const [text, setText] = useState(existing?.text ?? '')
  const ref = useRef<HTMLTextAreaElement>(null)
  useEffect(() => { ref.current?.focus() }, [])

  function commit(): void {
    // Excel keeps what was typed when you click away or press Esc; an untouched new note is dropped.
    if (!existing && !text.trim()) { onClose(); return }
    if (!existing || existing.text !== text) store.setNote(sheet, r, c, text, existing?.author ?? DEFAULT_AUTHOR)
    onClose()
  }
  const commitRef = useRef(commit)
  commitRef.current = commit
  useEffect(() => {
    const key = (e: KeyboardEvent) => { if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); commitRef.current() } }
    window.addEventListener('keydown', key, true)
    return () => window.removeEventListener('keydown', key, true)
  }, [])
  return createPortal(
    <div className="xl-note-scrim" onMouseDown={commit}>
      <div className="xl-note-box" style={{ left: rect.x + rect.w + 4, top: Math.max(4, rect.y - 4) }} onMouseDown={(e) => e.stopPropagation()}>
        <div className="xl-note-author">{existing?.author ?? DEFAULT_AUTHOR}:</div>
        <textarea ref={ref} value={text} onChange={(e) => setText(e.target.value)} />
      </div>
    </div>,
    document.body
  )
}

/** Read-only hover tooltip for a note (Excel shows the note without opening the editor on hover). */
export function NoteTooltip({ sheet, r, c, rect }: { sheet: SheetState; r: number; c: number; rect: { x: number; y: number; w: number; h: number } }): React.ReactElement | null {
  const note = sheet.layout.notes[r + ',' + c]
  if (!note) return null
  return createPortal(
    <div className="xl-note-tip" style={{ left: rect.x + rect.w + 4, top: rect.y }}>
      <div className="xl-note-author">{note.author}:</div>
      <div className="xl-note-text">{note.text}</div>
    </div>,
    document.body
  )
}

/** Threaded comment card: original comment, replies, a reply box, Resolve/Delete. */
export function CommentBox({ sheet, r, c, rect, onClose }: { sheet: SheetState; r: number; c: number; rect: { x: number; y: number; w: number; h: number }; onClose: () => void }): React.ReactElement {
  const thread = sheet.layout.comments[r + ',' + c]
  const [draft, setDraft] = useState('')
  const [replyDraft, setReplyDraft] = useState('')

  function postNew(): void {
    store.addComment(sheet, r, c, draft)
    setDraft('')
  }
  function postReply(): void {
    store.replyToComment(sheet, r, c, replyDraft)
    setReplyDraft('')
  }
  const keys = useRef({ draft, replyDraft, thread, onClose, postNew, postReply })
  keys.current = { draft, replyDraft, thread, onClose, postNew, postReply }
  useEffect(() => {
    const key = (e: KeyboardEvent) => {
      const k = keys.current
      if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); k.onClose() }
      else if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
        e.preventDefault(); e.stopPropagation()
        if (!k.thread && k.draft.trim()) k.postNew()
        else if (k.thread && k.replyDraft.trim()) k.postReply()
      }
    }
    window.addEventListener('keydown', key, true)
    return () => window.removeEventListener('keydown', key, true)
  }, [])

  return createPortal(
    <div className="xl-note-scrim" onMouseDown={onClose}>
      <div className="xl-comment-box" style={{ left: rect.x + rect.w + 4, top: Math.max(4, rect.y - 4) }} onMouseDown={(e) => e.stopPropagation()}>
        <div className="xl-comment-head">
          <span>Comments</span>
          <button className="xl-comment-x" onClick={onClose}>✕</button>
        </div>
        {!thread ? (
          <div className="xl-comment-new">
            <div className="xl-comment-cellref">{a1(r, c)}</div>
            <textarea autoFocus placeholder="Type @ to mention someone or type your comment here."
              value={draft} onChange={(e) => setDraft(e.target.value)} />
            <div className="xl-comment-actions">
              <button className="xl-btn" onClick={onClose}>Cancel</button>
              <button className="xl-btn primary" disabled={!draft.trim()} onClick={postNew}>Post</button>
            </div>
          </div>
        ) : (
          <>
            <div className={'xl-comment-thread' + (thread.resolved ? ' resolved' : '')}>
              <div className="xl-comment-item">
                <div className="xl-comment-avatar">{thread.author.slice(0, 1).toUpperCase()}</div>
                <div className="xl-comment-body">
                  <div className="xl-comment-meta"><b>{thread.author}</b><span>{fmtTime(thread.time)}</span></div>
                  <div className="xl-comment-text">{thread.text}</div>
                </div>
              </div>
              {thread.replies.map((rep) => (
                <div className="xl-comment-item reply" key={rep.id}>
                  <div className="xl-comment-avatar">{rep.author.slice(0, 1).toUpperCase()}</div>
                  <div className="xl-comment-body">
                    <div className="xl-comment-meta"><b>{rep.author}</b><span>{fmtTime(rep.time)}</span></div>
                    <div className="xl-comment-text">{rep.text}</div>
                  </div>
                </div>
              ))}
            </div>
            {!thread.resolved && (
              <div className="xl-comment-new">
                <textarea placeholder="Reply..." value={replyDraft} onChange={(e) => setReplyDraft(e.target.value)} />
                <div className="xl-comment-actions">
                  <button className="xl-btn" onClick={() => store.resolveComment(sheet, r, c, true)}>Resolve</button>
                  <button className="xl-btn" onClick={() => { store.deleteComment(sheet, r, c); onClose() }}>Delete Thread</button>
                  <button className="xl-btn primary" disabled={!replyDraft.trim()} onClick={postReply}>Post</button>
                </div>
              </div>
            )}
            {thread.resolved && (
              <div className="xl-comment-actions" style={{ padding: '8px 10px' }}>
                <button className="xl-btn" onClick={() => store.resolveComment(sheet, r, c, false)}>Reopen</button>
                <button className="xl-btn" onClick={() => { store.deleteComment(sheet, r, c); onClose() }}>Delete Thread</button>
              </div>
            )}
          </>
        )}
      </div>
    </div>,
    document.body
  )
}

export function closeNoteEditor(): void { setUi({ noteEditor: null }) }
export function closeCommentPanel(): void { setUi({ commentPanel: null }) }
